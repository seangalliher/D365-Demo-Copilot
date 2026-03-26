"""
D365 Demo Copilot — MCP Client Module

Provides an MCP (Model Context Protocol) client that connects to
HTTP/SSE MCP servers like the Dataverse MCP Server. Supports:

  - Connecting to Streamable HTTP and SSE MCP transports
  - Listing and calling tools (e.g., create_record, read_query)
  - Managing multiple server connections
  - Settings persistence for server configurations
"""

from .client import MCPClient, MCPServerConfig
from .manager import MCPManager

__all__ = ["MCPClient", "MCPServerConfig", "MCPManager"]
