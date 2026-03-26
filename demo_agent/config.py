"""
D365 Demo Copilot — Configuration

Central configuration loaded from environment variables and .env file.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from pathlib import Path

from dotenv import load_dotenv


# Load .env from project root
_env_path = Path(__file__).parent.parent / ".env"
if _env_path.exists():
    load_dotenv(_env_path)


@dataclass
class DemoConfig:
    """Configuration for the Demo Copilot agent."""

    # ---- D365 ----
    d365_base_url: str = field(
        default_factory=lambda: os.getenv(
            "D365_BASE_URL",
            "https://projectopscoreagentimplemented.crm.dynamics.com",
        )
    )
    d365_fo_base_url: str = field(
        default_factory=lambda: os.getenv("D365_FO_BASE_URL", "")
    )

    # ---- Dataverse OAuth (App Registration / Client Credentials) ----
    dataverse_tenant_id: str = field(
        default_factory=lambda: os.getenv("DATAVERSE_TENANT_ID", "")
    )
    dataverse_client_id: str = field(
        default_factory=lambda: os.getenv("DATAVERSE_CLIENT_ID", "")
    )
    dataverse_client_secret: str = field(
        default_factory=lambda: os.getenv("DATAVERSE_CLIENT_SECRET", "")
    )

    # ---- Microsoft Learn MCP ----
    ms_learn_mcp_url: str = field(
        default_factory=lambda: os.getenv(
            "MS_LEARN_MCP_URL", "https://learn.microsoft.com/api/mcp"
        )
    )

    # ---- LLM / Azure OpenAI ----
    azure_openai_endpoint: str = field(
        default_factory=lambda: os.getenv("AZURE_OPENAI_ENDPOINT", "")
    )
    azure_openai_api_key: str = field(
        default_factory=lambda: os.getenv("AZURE_OPENAI_API_KEY", "")
    )
    azure_openai_api_version: str = field(
        default_factory=lambda: os.getenv("AZURE_OPENAI_API_VERSION", "2024-10-21")
    )
    llm_model: str = field(
        default_factory=lambda: os.getenv("LLM_MODEL", "gpt-4o")
    )

    # ---- OpenAI (alternative to Azure OpenAI) ----
    openai_api_key: str = field(
        default_factory=lambda: os.getenv("OPENAI_API_KEY", "")
    )

    # ---- GitHub Models (alternative — uses OpenAI-compatible endpoint) ----
    github_token: str = field(
        default_factory=lambda: os.getenv("GITHUB_TOKEN", "")
    )
    github_models_base_url: str = field(
        default_factory=lambda: os.getenv(
            "GITHUB_MODELS_BASE_URL", "https://models.github.ai/inference"
        )
    )
    github_copilot_model: str = field(
        default_factory=lambda: os.getenv("GITHUB_COPILOT_MODEL", "openai/gpt-5")
    )

    # ---- Anthropic (alternative) ----
    anthropic_api_key: str = field(
        default_factory=lambda: os.getenv("ANTHROPIC_API_KEY", "")
    )

    # ---- Browser ----
    headless: bool = field(
        default_factory=lambda: os.getenv("BROWSER_HEADLESS", "false").lower() == "true"
    )
    slow_mo: int = field(
        default_factory=lambda: int(os.getenv("BROWSER_SLOW_MO", "50"))
    )
    auth_state_path: str = field(
        default_factory=lambda: os.getenv("AUTH_STATE_PATH", "auth_state.json")
    )

    # ---- Voice Narration (TTS) ----
    voice_enabled: bool = field(
        default_factory=lambda: os.getenv("VOICE_ENABLED", "false").lower() == "true"
    )
    voice_provider: str = field(
        default_factory=lambda: os.getenv("VOICE_PROVIDER", "auto")  # auto, edge, openai
    )
    voice_model: str = field(
        default_factory=lambda: os.getenv("VOICE_MODEL", "tts-1")
    )
    voice_name: str = field(
        default_factory=lambda: os.getenv("VOICE_NAME", "nova")
    )
    voice_speed: float = field(
        default_factory=lambda: float(os.getenv("VOICE_SPEED", "1.0"))
    )

    # ---- Demo defaults ----
    default_caption_speed: int = 25   # ms per character for typewriter effect
    max_demo_steps: int = 20
    default_action_delay_ms: int = 1000

    @property
    def use_dataverse_auth(self) -> bool:
        """Whether Dataverse app-registration credentials are configured."""
        return bool(
            self.dataverse_tenant_id
            and self.dataverse_client_id
            and self.dataverse_client_secret
        )

    @property
    def use_azure_openai(self) -> bool:
        """Whether Azure OpenAI credentials are configured."""
        return bool(self.azure_openai_endpoint and self.azure_openai_api_key)

    @property
    def use_github_copilot(self) -> bool:
        """Whether GitHub Copilot credentials are configured."""
        return bool(self.github_token)

    @property
    def use_openai(self) -> bool:
        """Whether standard OpenAI credentials are configured."""
        return bool(self.openai_api_key)

    def validate(self) -> list[str]:
        """Validate configuration and return any errors."""
        errors = []
        if (
            not self.use_azure_openai
            and not self.use_openai
            and not self.use_github_copilot
            and not self.anthropic_api_key
        ):
            errors.append(
                "No LLM credentials configured. Set AZURE_OPENAI_ENDPOINT + AZURE_OPENAI_API_KEY, "
                "or OPENAI_API_KEY, or GITHUB_TOKEN, or ANTHROPIC_API_KEY in .env"
            )
        return errors
