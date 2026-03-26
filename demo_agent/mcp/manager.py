"""
D365 Demo Copilot — MCP Manager

High-level manager for multiple MCP server connections.
Handles lifecycle, tool routing, and server configuration
persistence.

Usage:
    manager = MCPManager()
    manager.add_server(MCPServerConfig(name="Dataverse", url="..."))
    async with manager:
        tools = await manager.list_all_tools()
        result = await manager.call_tool("dataverse", "create_record", {...})
"""

from __future__ import annotations

import asyncio
import json
import logging
from pathlib import Path
from typing import Any, Optional

from .client import MCPClient, MCPServerConfig

logger = logging.getLogger("demo_agent.mcp.manager")


class MCPManager:
    """
    Manages multiple MCP server connections.

    Provides a unified interface for discovering and calling tools
    across all connected MCP servers.
    """

    def __init__(self):
        self._servers: dict[str, MCPServerConfig] = {}
        self._clients: dict[str, MCPClient] = {}
        self._connected = False

    @property
    def connected(self) -> bool:
        return self._connected

    @property
    def server_names(self) -> list[str]:
        """Names of all registered servers."""
        return list(self._servers.keys())

    def add_server(self, config: MCPServerConfig) -> None:
        """Register a server configuration."""
        if self._connected:
            raise RuntimeError(
                "Cannot add servers while connected. Disconnect first."
            )
        self._servers[config.name] = config
        logger.info("Registered MCP server: %s (%s)", config.name, config.url)

    def remove_server(self, name: str) -> None:
        """Remove a server configuration."""
        if self._connected:
            raise RuntimeError(
                "Cannot remove servers while connected. Disconnect first."
            )
        self._servers.pop(name, None)

    async def connect(self) -> None:
        """Connect to all registered and enabled servers."""
        if self._connected:
            logger.warning("Already connected")
            return

        tasks = []
        for name, config in self._servers.items():
            if not config.enabled:
                logger.info("Skipping disabled server: %s", name)
                continue
            client = MCPClient(config)
            self._clients[name] = client
            tasks.append(self._connect_one(name, client))

        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Log any connection failures
        for name, result in zip(self._clients.keys(), results):
            if isinstance(result, Exception):
                logger.error("Failed to connect to %s: %s", name, result)
                del self._clients[name]

        self._connected = True
        logger.info(
            "MCPManager connected — %d/%d servers active",
            len(self._clients),
            len(self._servers),
        )

    @staticmethod
    async def _connect_one(name: str, client: MCPClient) -> None:
        """Connect a single client (for use with gather)."""
        await client.connect()

    async def disconnect(self) -> None:
        """Disconnect from all servers."""
        if not self._connected:
            return

        tasks = [client.disconnect() for client in self._clients.values()]
        await asyncio.gather(*tasks, return_exceptions=True)

        self._clients.clear()
        self._connected = False
        logger.info("MCPManager disconnected")

    async def list_all_tools(self) -> dict[str, list[dict]]:
        """
        List tools from all connected servers.

        Returns:
            Dict mapping server name → list of tool descriptors.
        """
        result: dict[str, list[dict]] = {}
        for name, client in self._clients.items():
            try:
                result[name] = await client.list_tools()
            except Exception as e:
                logger.error("Failed to list tools from %s: %s", name, e)
                result[name] = []
        return result

    async def call_tool(
        self,
        server_name: str,
        tool_name: str,
        arguments: dict[str, Any] | None = None,
    ) -> dict:
        """
        Call a tool on a specific server.

        Args:
            server_name: Name of the MCP server.
            tool_name: Name of the tool to call.
            arguments: Tool arguments.

        Returns:
            Dict with 'content' and 'isError' keys.

        Raises:
            KeyError: If server_name is not connected.
            RuntimeError: If not connected.
        """
        if not self._connected:
            raise RuntimeError("MCPManager is not connected")

        client = self._clients.get(server_name)
        if client is None:
            raise KeyError(
                f"Server '{server_name}' not found. "
                f"Connected servers: {list(self._clients.keys())}"
            )

        return await client.call_tool(tool_name, arguments)

    def find_tool_server(self, tool_name: str) -> Optional[str]:
        """
        Find which server provides a given tool.

        Returns the server name, or None if not found.
        """
        for name, client in self._clients.items():
            for tool in client.tools:
                if tool["name"] == tool_name:
                    return name
        return None

    async def call_tool_auto(
        self,
        tool_name: str,
        arguments: dict[str, Any] | None = None,
    ) -> dict:
        """
        Call a tool, automatically routing to the correct server.

        Searches all connected servers for the named tool and calls it.

        Raises:
            ValueError: If no server provides the tool.
        """
        server_name = self.find_tool_server(tool_name)
        if server_name is None:
            raise ValueError(
                f"Tool '{tool_name}' not found on any connected server"
            )
        return await self.call_tool(server_name, tool_name, arguments)

    # ---- Persistence ----

    def save_config(self, path: str | Path) -> None:
        """Save all server configurations to a JSON file."""
        path = Path(path)
        data = {
            name: config.to_dict()
            for name, config in self._servers.items()
        }
        path.write_text(json.dumps(data, indent=2), encoding="utf-8")
        logger.info("Saved MCP config to %s", path)

    def load_config(self, path: str | Path) -> None:
        """Load server configurations from a JSON file."""
        path = Path(path)
        if not path.exists():
            logger.warning("Config file not found: %s", path)
            return

        data = json.loads(path.read_text(encoding="utf-8"))
        for name, cfg_dict in data.items():
            config = MCPServerConfig.from_dict(cfg_dict)
            self._servers[config.name] = config
        logger.info("Loaded %d MCP configs from %s", len(data), path)

    # ---- Context manager ----

    async def __aenter__(self):
        await self.connect()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.disconnect()
        return False
