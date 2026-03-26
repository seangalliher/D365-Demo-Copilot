"""
D365 Demo Copilot — MCP Client

Low-level MCP client that connects to a single MCP server over
Streamable HTTP or SSE transport.  Wraps the official `mcp` Python SDK.

Usage:
    config = MCPServerConfig(name="Dataverse", url="https://…/api/mcp")
    async with MCPClient(config) as client:
        tools = await client.list_tools()
        result = await client.call_tool("create_record", {"table": "…", …})
"""

from __future__ import annotations

import asyncio
import logging
from contextlib import asynccontextmanager
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, AsyncGenerator, Optional

from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client
from mcp.client.sse import sse_client

logger = logging.getLogger("demo_agent.mcp.client")


class TransportType(str, Enum):
    """MCP transport protocol."""
    STREAMABLE_HTTP = "streamable_http"
    SSE = "sse"
    AUTO = "auto"  # Try streamable HTTP first, fall back to SSE


@dataclass
class MCPServerConfig:
    """Configuration for a single MCP server connection."""
    name: str
    url: str
    transport: TransportType = TransportType.AUTO
    headers: dict[str, str] = field(default_factory=dict)
    timeout: float = 30.0
    sse_read_timeout: float = 300.0
    enabled: bool = True

    def to_dict(self) -> dict:
        """Serialize to dict for JSON persistence."""
        return {
            "name": self.name,
            "url": self.url,
            "transport": self.transport.value,
            "headers": self.headers,
            "timeout": self.timeout,
            "sse_read_timeout": self.sse_read_timeout,
            "enabled": self.enabled,
        }

    @classmethod
    def from_dict(cls, data: dict) -> MCPServerConfig:
        """Deserialize from dict."""
        return cls(
            name=data["name"],
            url=data["url"],
            transport=TransportType(data.get("transport", "auto")),
            headers=data.get("headers", {}),
            timeout=data.get("timeout", 30.0),
            sse_read_timeout=data.get("sse_read_timeout", 300.0),
            enabled=data.get("enabled", True),
        )


class MCPClient:
    """
    MCP client for a single server.

    Manages connection lifecycle, tool discovery, and tool invocation.
    Use as an async context manager or call connect()/disconnect() manually.
    """

    def __init__(self, config: MCPServerConfig):
        self.config = config
        self._session: Optional[ClientSession] = None
        self._transport_ctx = None
        self._session_ctx = None
        self._connected = False
        self._tools: list[dict] = []
        self._server_info: dict = {}

    @property
    def connected(self) -> bool:
        return self._connected

    @property
    def tools(self) -> list[dict]:
        """Cached list of tools from last list_tools() call."""
        return self._tools

    @property
    def server_info(self) -> dict:
        """Server info from initialization."""
        return self._server_info

    async def connect(self) -> None:
        """
        Connect to the MCP server.

        Tries Streamable HTTP first (MCP 2025+), falls back to SSE if needed.
        """
        if self._connected:
            logger.warning("Already connected to %s", self.config.name)
            return

        transport = self.config.transport
        last_error = None

        # Determine which transports to try
        if transport == TransportType.AUTO:
            attempts = [TransportType.STREAMABLE_HTTP, TransportType.SSE]
        else:
            attempts = [transport]

        for attempt in attempts:
            try:
                logger.info(
                    "Connecting to %s via %s at %s",
                    self.config.name, attempt.value, self.config.url,
                )

                if attempt == TransportType.STREAMABLE_HTTP:
                    self._transport_ctx = streamablehttp_client(
                        url=self.config.url,
                        headers=self.config.headers or None,
                        timeout=self.config.timeout,
                        sse_read_timeout=self.config.sse_read_timeout,
                    )
                else:
                    self._transport_ctx = sse_client(
                        url=self.config.url,
                        headers=self.config.headers or None,
                        timeout=self.config.timeout,
                        sse_read_timeout=self.config.sse_read_timeout,
                    )

                transport_result = await self._transport_ctx.__aenter__()

                # streamablehttp returns (read, write, get_session_id)
                # sse returns (read, write)
                if len(transport_result) == 3:
                    read_stream, write_stream, _ = transport_result
                else:
                    read_stream, write_stream = transport_result

                self._session_ctx = ClientSession(read_stream, write_stream)
                self._session = await self._session_ctx.__aenter__()

                # Initialize the MCP session
                init_result = await self._session.initialize()
                self._server_info = {
                    "name": getattr(init_result, "server_info", {}).get("name", self.config.name)
                    if isinstance(getattr(init_result, "server_info", None), dict)
                    else self.config.name,
                    "protocol_version": getattr(init_result, "protocol_version", "unknown"),
                }

                # Discover tools
                await self.refresh_tools()

                self._connected = True
                logger.info(
                    "Connected to %s — %d tools available",
                    self.config.name, len(self._tools),
                )
                return

            except Exception as e:
                last_error = e
                logger.warning(
                    "Failed to connect via %s: %s", attempt.value, e,
                )
                # Clean up the failed attempt
                await self._cleanup_transport()
                continue

        raise ConnectionError(
            f"Failed to connect to {self.config.name} at {self.config.url}: {last_error}"
        )

    async def disconnect(self) -> None:
        """Disconnect from the MCP server."""
        if not self._connected:
            return

        logger.info("Disconnecting from %s", self.config.name)
        await self._cleanup_transport()
        self._connected = False
        self._tools = []
        self._server_info = {}

    async def _cleanup_transport(self):
        """Clean up session and transport context managers."""
        if self._session_ctx:
            try:
                await self._session_ctx.__aexit__(None, None, None)
            except Exception:
                pass
            self._session_ctx = None
            self._session = None

        if self._transport_ctx:
            try:
                await self._transport_ctx.__aexit__(None, None, None)
            except Exception:
                pass
            self._transport_ctx = None

    async def refresh_tools(self) -> list[dict]:
        """Refresh the cached tool list from the server."""
        if not self._session:
            raise RuntimeError("Not connected")

        result = await self._session.list_tools()
        self._tools = [
            {
                "name": tool.name,
                "description": getattr(tool, "description", ""),
                "input_schema": (
                    tool.inputSchema
                    if hasattr(tool, "inputSchema")
                    else getattr(tool, "input_schema", {})
                ),
            }
            for tool in result.tools
        ]
        return self._tools

    async def call_tool(
        self, tool_name: str, arguments: dict[str, Any] | None = None
    ) -> dict:
        """
        Call a tool on the MCP server.

        Args:
            tool_name: Name of the tool to call.
            arguments: Tool arguments as a dict.

        Returns:
            Dict with 'content' (list of content blocks) and 'isError' flag.
        """
        if not self._session:
            raise RuntimeError("Not connected")

        logger.info("Calling tool %s on %s", tool_name, self.config.name)
        result = await self._session.call_tool(tool_name, arguments or {})

        # Serialize content blocks to plain dicts
        content = []
        for block in result.content:
            if hasattr(block, "text"):
                content.append({"type": "text", "text": block.text})
            elif hasattr(block, "data"):
                content.append({"type": "resource", "data": str(block.data)})
            else:
                content.append({"type": "unknown", "value": str(block)})

        return {
            "content": content,
            "isError": getattr(result, "isError", False),
        }

    async def list_tools(self) -> list[dict]:
        """Get cached tools (refreshes if empty)."""
        if not self._tools:
            await self.refresh_tools()
        return self._tools

    # ---- Context manager ----

    async def __aenter__(self):
        await self.connect()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.disconnect()
        return False
