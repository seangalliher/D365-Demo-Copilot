"""
D365 Demo Copilot — Dataverse OAuth Authentication

Acquires and caches OAuth 2.0 tokens for the Dataverse MCP endpoint
using an Azure AD app registration (client credentials grant).

Usage:
    auth = DataverseAuth(tenant_id, client_id, client_secret, d365_base_url)
    token = auth.get_token()          # cached; refreshes automatically
    headers = auth.get_auth_headers()  # {"Authorization": "Bearer <token>"}
"""

from __future__ import annotations

import logging
import time
from typing import Optional

import msal

logger = logging.getLogger("demo_agent.auth")


class DataverseAuth:
    """
    Acquires OAuth 2.0 access tokens for Dataverse via MSAL client credentials.

    The token is cached in-memory and refreshed automatically when it nears
    expiry (60-second buffer before ``expires_in``).
    """

    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        resource_url: str,
    ):
        """
        Args:
            tenant_id:     Azure AD tenant ID (GUID or domain).
            client_id:     App registration Application (client) ID.
            client_secret: App registration client secret value.
            resource_url:  Dataverse org URL, e.g.
                           ``https://myorg.crm.dynamics.com``.
                           The ``/.default`` scope is appended automatically.
        """
        self._tenant_id = tenant_id
        self._client_id = client_id
        self._client_secret = client_secret

        # Normalize: strip trailing slash, append /.default for scope
        base = resource_url.rstrip("/")
        self._scope = [f"{base}/.default"]

        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self._app = msal.ConfidentialClientApplication(
            client_id=client_id,
            client_credential=client_secret,
            authority=authority,
        )

        # In-memory cache
        self._cached_token: Optional[str] = None
        self._expires_at: float = 0.0

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def get_token(self) -> str:
        """
        Return a valid access token, refreshing from Azure AD if needed.

        Raises ``RuntimeError`` if token acquisition fails.
        """
        # Return cached token if still valid (with 60-second buffer)
        if self._cached_token and time.time() < (self._expires_at - 60):
            return self._cached_token

        result = self._app.acquire_token_for_client(scopes=self._scope)

        if result and isinstance(result, dict) and "access_token" in result:
            token: str = result["access_token"]
            expires_in: int = int(result.get("expires_in", 3600))
            self._cached_token = token
            self._expires_at = time.time() + expires_in
            logger.info("Acquired Dataverse token (expires in %ds)", expires_in)
            return token

        error = result.get("error", "unknown") if isinstance(result, dict) else "unknown"
        desc = result.get("error_description", "No description") if isinstance(result, dict) else str(result)
        raise RuntimeError(
            f"Failed to acquire Dataverse token: {error} — {desc}"
        )

    def get_auth_headers(self) -> dict[str, str]:
        """Return HTTP headers with the current Bearer token."""
        return {"Authorization": f"Bearer {self.get_token()}"}

    @property
    def configured(self) -> bool:
        """True if all required credentials are present (non-empty)."""
        return bool(
            self._tenant_id and self._client_id and self._client_secret
        )
