import os
import sys
import traceback
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse
from uuid import uuid4
from dotenv import load_dotenv

from microsoft_agents.hosting.core import (
    AgentApplication,
    TurnState,
    TurnContext,
    MemoryStorage,
)
from microsoft_agents.activity import (
    load_configuration_from_env,
    ActivityTypes,
)
from microsoft_agents.hosting.aiohttp import CloudAdapter
from microsoft_agents.authentication.msal import MsalConnectionManager
from anthropic import Anthropic
from microsoft_agents_a365.observability.core import config as observability_config
from microsoft_agents_a365.observability.core.agent_details import AgentDetails
from microsoft_agents_a365.observability.core.execution_type import ExecutionType
from microsoft_agents_a365.observability.core.inference_call_details import InferenceCallDetails
from microsoft_agents_a365.observability.core.inference_operation_type import InferenceOperationType
from microsoft_agents_a365.observability.core.inference_scope import InferenceScope
from microsoft_agents_a365.observability.core.invoke_agent_details import InvokeAgentDetails
from microsoft_agents_a365.observability.core.invoke_agent_scope import InvokeAgentScope
from microsoft_agents_a365.observability.core.middleware.baggage_builder import BaggageBuilder
from microsoft_agents_a365.observability.core.models.caller_details import CallerDetails
from microsoft_agents_a365.observability.core.request import Request
from microsoft_agents_a365.observability.core.source_metadata import SourceMetadata
from microsoft_agents_a365.observability.core.tenant_details import TenantDetails
from opentelemetry import trace
from opentelemetry.sdk.trace import TracerProvider
from opentelemetry.sdk.trace.export import ConsoleSpanExporter, SimpleSpanProcessor
from openai import AzureOpenAI

from config import Config

load_dotenv()

# Load configuration
config = Config(os.environ)
agents_sdk_config = load_configuration_from_env(os.environ)

client = None
azure_openai_client = None

if config.llm_provider == "anthropic":
    client = Anthropic(
        api_key=config.anthropic_api_key,
        base_url=config.anthropic_base_url,
    )
else:
    azure_openai_client = AzureOpenAI(
        api_key=config.azure_openai_api_key,
        api_version=config.azure_openai_api_version,
        azure_endpoint=config.azure_openai_endpoint,
    )


def _observability_token_resolver(_agent_id: str, _tenant_id: str) -> str | None:
    return config.a365_observability_token


observability_config.configure(
    service_name=config.a365_service_name,
    service_namespace=config.a365_service_namespace,
    token_resolver=_observability_token_resolver,
    cluster_category=config.a365_cluster_category,
)


def _configure_console_span_exporter() -> None:
    if not config.a365_enable_console_span_exporter:
        return

    tracer_provider = trace.get_tracer_provider()
    if not isinstance(tracer_provider, TracerProvider):
        return

    if getattr(tracer_provider, "_grilling_console_exporter_configured", False):
        return

    tracer_provider.add_span_processor(
        SimpleSpanProcessor(ConsoleSpanExporter(out=sys.stdout))
    )
    tracer_provider._grilling_console_exporter_configured = True
    print("OpenTelemetry console span exporter enabled.", flush=True)


_configure_console_span_exporter()
otel_tracer = trace.get_tracer("grilling-agent")

LOG_FILE_PATH = Path(__file__).resolve().parents[1] / "logs" / "prompt-history.log"


def _log_prompt_event(label: str, content: str) -> None:
    print(f"\n[{label}]\n{content}\n", flush=True)
    LOG_FILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    with LOG_FILE_PATH.open("a", encoding="utf-8") as log_file:
        log_file.write(
            f"[{datetime.now(timezone.utc).isoformat()}] {label}\n{content}\n\n"
        )

system_prompt = """You are a Grilling Chef Companion Agent.

Your role is to provide expert, practical, and safety-aware guidance on grilling and outdoor cooking. You act like a knowledgeable grilling chef and host, helping users plan meals, cook food safely, and adapt grilling advice to real-world constraints.

CORE RESPONSIBILITIES
- Recommend grilling recipes, techniques, and meal plans.
- Adjust advice based on:
    - Number of people being served
    - Dietary restrictions and allergies (for example: nut allergies, gluten intolerance, vegetarian/vegan)
    - Nutrition goals (for example: high-protein, low-sodium, low-fat, calorie-conscious)
    - Available equipment (gas grill, charcoal grill, smoker, electric grill)
    - Skill level (beginner to advanced)
- Provide clear, step-by-step cooking guidance when requested.

SAFETY AND HEALTH
- Always promote safe food-handling and grilling practices.
- Call out safe internal cooking temperatures for meats, poultry, and seafood.
- Warn about cross-contamination risks when allergies are mentioned.
- Avoid giving medical advice; provide general nutrition and food safety guidance only.

COMMUNICATION STYLE
- Be friendly, confident, and practical.
- Use clear instructions and simple language.
- Ask follow-up questions when needed to personalize advice (for example: number of guests, allergies, grill type).
- Prioritize clarity over culinary jargon.

BOUNDARIES
- Do not provide medical diagnoses or treatment advice.
- Do not ignore allergy or safety constraints.
- If information is missing, ask clarifying questions before giving recommendations.

EXAMPLES OF SUPPORTED REQUESTS
- "I'm grilling for 8 people-2 are vegetarian-what should I make?"
- "Give me a healthy grilling plan for a family barbecue."
- "How do I safely grill chicken for someone with a nut allergy?"
- "What's a good low-carb grilling option for dinner?"
- "I only have a small gas grill-what works best?"

GOAL
Help users enjoy grilling confidently by combining flavor, safety, nutrition awareness, and practical planning."""

# Define storage and application
storage = MemoryStorage()
connection_manager = MsalConnectionManager(**agents_sdk_config)
adapter = CloudAdapter(connection_manager=connection_manager)

agent_app = AgentApplication[TurnState](
    storage=storage, 
    adapter=adapter, 
    **agents_sdk_config
)

@agent_app.conversation_update("membersAdded")
async def on_members_added(context: TurnContext, _state: TurnState):
    # Keep installation/update flow side-effect free in local playground.
    return

# Listen for ANY message to be received. MUST BE AFTER ANY OTHER MESSAGE HANDLERS
@agent_app.activity(ActivityTypes.message)
async def on_message(context: TurnContext, _state: TurnState):
    if not context.activity.text:
        return

    activity = context.activity
    conversation_id = getattr(getattr(activity, "conversation", None), "id", None) or str(uuid4())
    tenant_id = (
        getattr(getattr(activity, "conversation", None), "tenant_id", None)
        or config.tenant_id
    )
    correlation_id = str(uuid4())

    endpoint = urlparse(os.getenv("BOT_ENDPOINT", f"http://localhost:{config.PORT}/api/messages"))
    agent_details = AgentDetails(
        agent_id=config.agent_id,
        agent_name=config.agent_name,
        agent_auid=config.agent_auid,
        agent_upn=config.agent_upn,
        agent_blueprint_id=config.agent_blueprint_id,
        tenant_id=tenant_id,
        conversation_id=conversation_id,
    )
    tenant_details = TenantDetails(tenant_id=tenant_id)
    request = Request(
        content=activity.text,
        execution_type=ExecutionType.HUMAN_TO_AGENT,
        session_id=conversation_id,
        source_metadata=SourceMetadata(
            id=conversation_id,
            name=getattr(activity, "channel_id", "unknown"),
        ),
    )

    from_user = getattr(activity, "from_property", None) or getattr(activity, "from", None)
    caller_id = getattr(from_user, "aad_object_id", None) or getattr(from_user, "id", None)
    caller_name = getattr(from_user, "name", None)
    caller_upn = getattr(from_user, "email", None) or caller_id
    caller_details = CallerDetails(
        caller_id=caller_id,
        caller_name=caller_name,
        caller_upn=caller_upn,
        caller_user_id=caller_id,
        tenant_id=tenant_id,
    )

    invoke_details = InvokeAgentDetails(
        details=agent_details,
        endpoint=endpoint,
        session_id=conversation_id,
    )

    with otel_tracer.start_as_current_span("invoke_agent") as invoke_span:
        invoke_span.set_attribute("gen_ai.system", "az.ai.agent365")
        invoke_span.set_attribute("gen_ai.operation.name", "invoke_agent")
        invoke_span.set_attribute("tenant.id", tenant_id)
        invoke_span.set_attribute("correlation.id", correlation_id)
        invoke_span.set_attribute("gen_ai.conversation.id", conversation_id)
        invoke_span.set_attribute("gen_ai.agent.id", config.agent_id)
        invoke_span.set_attribute("gen_ai.agent.name", config.agent_name)
        if caller_id:
            invoke_span.set_attribute("gen_ai.caller.id", caller_id)
        if caller_name:
            invoke_span.set_attribute("gen_ai.caller.name", caller_name)

        with (
            BaggageBuilder()
            .tenant_id(tenant_id)
            .agent_id(config.agent_id)
            .correlation_id(correlation_id)
            .build()
        ):
            with InvokeAgentScope.start(invoke_details, tenant_details, request, caller_details=caller_details) as invoke_scope:
                invoke_scope.record_input_messages([activity.text])

                try:
                    with InferenceScope.start(
                        InferenceCallDetails(
                            operationName=InferenceOperationType.CHAT,
                            model=config.anthropic_model,
                            providerName="anthropic",
                        ),
                        agent_details,
                        tenant_details,
                        request,
                    ) as inference_scope:
                        inference_scope.record_input_messages([activity.text])
                        _log_prompt_event("system prompt", system_prompt)
                        _log_prompt_event("user prompt", activity.text)

                        with otel_tracer.start_as_current_span("inference") as inference_span:
                            inference_span.set_attribute("gen_ai.system", "az.ai.agent365")
                            inference_span.set_attribute("gen_ai.operation.name", "inference")
                            if config.llm_provider == "anthropic":
                                inference_span.set_attribute("gen_ai.request.model", config.anthropic_model)
                                inference_span.set_attribute("gen_ai.provider.name", "anthropic")

                                result = client.messages.create(
                                    model=config.anthropic_model,
                                    max_tokens=800,
                                    system=system_prompt,
                                    messages=[
                                        {
                                            "role": "user",
                                            "content": activity.text,
                                        },
                                    ],
                                )

                                text_chunks = [block.text for block in result.content if hasattr(block, "text")]
                                answer = "\n".join(text_chunks).strip() or "I couldn't generate a response."

                                usage = getattr(result, "usage", None)
                                if usage is not None:
                                    if getattr(usage, "input_tokens", None) is not None:
                                        inference_scope.record_input_tokens(usage.input_tokens)
                                        inference_span.set_attribute("gen_ai.usage.input_tokens", usage.input_tokens)
                                    if getattr(usage, "output_tokens", None) is not None:
                                        inference_scope.record_output_tokens(usage.output_tokens)
                                        inference_span.set_attribute("gen_ai.usage.output_tokens", usage.output_tokens)

                                if getattr(result, "stop_reason", None):
                                    stop_reason = str(result.stop_reason)
                                    inference_scope.record_finish_reasons([stop_reason])
                                    inference_span.set_attribute("gen_ai.response.finish_reason", stop_reason)
                            else:
                                inference_span.set_attribute(
                                    "gen_ai.request.model", config.azure_openai_deployment_name
                                )
                                inference_span.set_attribute("gen_ai.provider.name", "azure_openai")

                                result = azure_openai_client.chat.completions.create(
                                    model=config.azure_openai_deployment_name,
                                    messages=[
                                        {"role": "system", "content": system_prompt},
                                        {"role": "user", "content": activity.text},
                                    ],
                                    max_tokens=800,
                                )

                                answer = (
                                    (result.choices[0].message.content or "").strip()
                                    if result.choices and result.choices[0].message
                                    else ""
                                ) or "I couldn't generate a response."

                                usage = getattr(result, "usage", None)
                                if usage is not None:
                                    if getattr(usage, "prompt_tokens", None) is not None:
                                        inference_scope.record_input_tokens(usage.prompt_tokens)
                                        inference_span.set_attribute("gen_ai.usage.input_tokens", usage.prompt_tokens)
                                    if getattr(usage, "completion_tokens", None) is not None:
                                        inference_scope.record_output_tokens(usage.completion_tokens)
                                        inference_span.set_attribute(
                                            "gen_ai.usage.output_tokens", usage.completion_tokens
                                        )

                                if result.choices and getattr(result.choices[0], "finish_reason", None):
                                    stop_reason = str(result.choices[0].finish_reason)
                                    inference_scope.record_finish_reasons([stop_reason])
                                    inference_span.set_attribute("gen_ai.response.finish_reason", stop_reason)

                            inference_scope.record_output_messages([answer])
                            _log_prompt_event("assistant response", answer)
                except Exception as error:
                    invoke_scope.record_error(error)
                    raise

                invoke_scope.record_output_messages([answer])
                invoke_scope.record_response(answer)

    await context.send_activity(answer)

@agent_app.error
async def on_error(context: TurnContext, error: Exception):
    # This check writes out errors to console log .vs. app insights.
    # NOTE: In production environment, you should consider logging this to Azure
    #       application insights.
    print(f"\n [on_turn_error] unhandled error: {error}", file=sys.stderr)
    traceback.print_exc()

    # Avoid sending a secondary activity on error in local playground,
    # because connector failures can cascade into repeated 400/500 responses.
    return
