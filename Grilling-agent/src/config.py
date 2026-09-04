"""
Copyright (c) Microsoft Corporation. All rights reserved.
Licensed under the MIT License.
"""

class Config:
    """Agent Configuration"""

    def __init__(self, env):
        self.PORT = 3978
        # Anthropic configuration.
        self.anthropic_api_key = env.get("ANTHROPIC_API_KEY")
        self.anthropic_model = env.get("ANTHROPIC_MODEL")
        raw_anthropic_base_url = env.get("ANTHROPIC_BASE_URL", "https://api.anthropic.com")
        self.anthropic_base_url = raw_anthropic_base_url.strip().strip("\"'")

        # Azure OpenAI configuration.
        self.azure_openai_api_key = env.get("AZURE_OPENAI_API_KEY")
        raw_azure_openai_endpoint = env.get("AZURE_OPENAI_ENDPOINT", "")
        self.azure_openai_endpoint = raw_azure_openai_endpoint.strip().strip("\"'")
        self.azure_openai_deployment_name = env.get("AZURE_OPENAI_DEPLOYMENT_NAME")
        self.azure_openai_api_version = env.get("AZURE_OPENAI_API_VERSION", "2024-10-21")

        # Compatibility mode: some deployments store Anthropic settings in
        # AZURE_OPENAI_* names while still targeting api.anthropic.com.
        if (
            (not self.anthropic_api_key or not self.anthropic_model)
            and self.azure_openai_endpoint
            and "anthropic.com" in self.azure_openai_endpoint.lower()
            and self.azure_openai_api_key
            and self.azure_openai_deployment_name
        ):
            self.anthropic_api_key = self.azure_openai_api_key
            self.anthropic_model = self.azure_openai_deployment_name
            self.anthropic_base_url = self.azure_openai_endpoint

        # Provider selection: prefer Anthropic when available, otherwise use Azure OpenAI.
        if self.anthropic_api_key and self.anthropic_model:
            self.llm_provider = "anthropic"
        elif (
            self.azure_openai_api_key
            and self.azure_openai_endpoint
            and self.azure_openai_deployment_name
        ):
            self.llm_provider = "azure_openai"
        else:
            raise ValueError(
                "Missing LLM configuration. Set ANTHROPIC_API_KEY + ANTHROPIC_MODEL "
                "or AZURE_OPENAI_API_KEY + AZURE_OPENAI_ENDPOINT + AZURE_OPENAI_DEPLOYMENT_NAME."
            )

        # Agent 365 observability settings
        self.enable_a365_observability_exporter = (env.get("ENABLE_A365_OBSERVABILITY_EXPORTER", "false").lower() == "true")
        self.a365_enable_console_span_exporter = (env.get("A365_ENABLE_CONSOLE_SPAN_EXPORTER", "true").lower() == "true")
        self.a365_service_name = env.get("A365_OBSERVABILITY_SERVICE_NAME", "grilling-agent")
        self.a365_service_namespace = env.get("A365_OBSERVABILITY_SERVICE_NAMESPACE", "m365.agents")
        self.a365_cluster_category = env.get("A365_OBSERVABILITY_CLUSTER_CATEGORY", "prod")
        self.a365_observability_token = env.get("A365_OBSERVABILITY_TOKEN")

        # Details used by manual scopes
        self.agent_id = env.get("AGENT_ID") or env.get("BOT_ID") or "grilling-agent"
        self.agent_name = env.get("AGENT_NAME", "Grilling Chef Companion")
        self.agent_upn = env.get("AGENT_UPN") or self.agent_name
        self.agent_auid = env.get("AGENT_AUID") or self.agent_id
        self.agent_blueprint_id = env.get("AGENT_BLUEPRINT_ID") or env.get("TEAMS_APP_ID") or self.agent_id
        self.tenant_id = (
            env.get("TENANT_ID")
            or env.get("CONNECTIONS__SERVICE_CONNECTION__SETTINGS__TENANTID")
            or "unknown-tenant"
        )
