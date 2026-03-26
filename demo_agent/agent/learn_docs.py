"""
D365 Demo Copilot — Microsoft Learn Documentation Integration

Queries the Microsoft Learn MCP server (learn.microsoft.com/api/mcp)
to fetch relevant documentation snippets that ground demo narration
and planning in official Microsoft content.

This is a lightweight wrapper — it uses the same MCPClient infrastructure
as the Dataverse MCP connection but connects to the public MS Docs endpoint.
"""

from __future__ import annotations

import asyncio
import logging
from typing import Optional

import httpx

from ..mcp.client import MCPClient, MCPServerConfig, TransportType

logger = logging.getLogger("demo_agent.learn_docs")

# Default MS Learn MCP endpoint
MS_LEARN_MCP_URL = "https://learn.microsoft.com/api/mcp"

# Keywords that map demo topics → MS Learn search queries
TOPIC_SEARCH_MAP: dict[str, list[str]] = {
    # CE / Project Operations
    "time entry": [
        "Dynamics 365 Project Operations time entry",
        "create time entries project operations",
    ],
    "expense": [
        "Dynamics 365 Project Operations expense management",
        "create expense entries project operations",
    ],
    "project": [
        "Dynamics 365 Project Operations create project",
        "project management project operations",
    ],
    "resource": [
        "Dynamics 365 Project Operations resource management",
        "bookable resources project operations",
    ],
    "approval": [
        "Dynamics 365 Project Operations approvals",
        "approve time expense project operations",
    ],
    "pricing": [
        "Dynamics 365 Project Operations pricing",
        "cost bill rate price lists",
    ],
    "contract": [
        "Dynamics 365 Project Operations contracts",
        "project contracts invoicing",
    ],
    "subcontract": [
        "Dynamics 365 Project Operations subcontracting",
        "vendor subcontracts project operations",
    ],
    "invoice": [
        "Dynamics 365 Project Operations invoicing",
        "create project invoice",
    ],
    "schedule": [
        "Dynamics 365 Project Operations scheduling",
        "work breakdown structure project",
    ],
    # F&O / Finance
    "purchase order": [
        "Dynamics 365 Finance purchase orders",
        "create purchase order supply chain management",
    ],
    "vendor": [
        "Dynamics 365 Finance vendor management",
        "vendor master data supply chain",
    ],
    "general ledger": [
        "Dynamics 365 Finance general ledger",
        "journal entries general ledger",
    ],
    "journal": [
        "Dynamics 365 Finance general journal",
        "post journal entries ledger",
    ],
    "chart of accounts": [
        "Dynamics 365 Finance chart of accounts",
        "main accounts financial dimensions",
    ],
    "sales order": [
        "Dynamics 365 Supply Chain Management sales orders",
        "create sales order SCM",
    ],
    "inventory": [
        "Dynamics 365 Supply Chain Management inventory",
        "inventory management warehousing",
    ],
    "production": [
        "Dynamics 365 Supply Chain Management production orders",
        "manufacturing production control",
    ],
    "fixed asset": [
        "Dynamics 365 Finance fixed assets",
        "fixed asset acquisition depreciation",
    ],
    "budget": [
        "Dynamics 365 Finance budgeting",
        "budget register entries",
    ],
    "bank": [
        "Dynamics 365 Finance bank reconciliation",
        "cash bank management",
    ],
    "intercompany": [
        "Dynamics 365 Finance intercompany accounting",
        "intercompany transactions project operations",
    ],
}


class LearnDocsClient:
    """
    Client for querying Microsoft Learn documentation via the MCP server.

    Provides search and fetch capabilities to enrich demo plans
    with official Microsoft documentation content.
    """

    def __init__(self, url: str = MS_LEARN_MCP_URL):
        self.url = url
        self._mcp_client: Optional[MCPClient] = None
        self._connected = False
        self._available_tools: list[str] = []

    @property
    def connected(self) -> bool:
        return self._connected

    async def connect(self) -> bool:
        """
        Connect to the MS Learn MCP server.

        Returns True if connected successfully, False otherwise.
        The agent can still function without this — it just won't
        have doc enrichment.
        """
        try:
            # Probe the endpoint first to avoid event loop issues
            async with httpx.AsyncClient(timeout=10) as http:
                resp = await http.get(self.url, follow_redirects=True)
                logger.info(
                    "MS Learn MCP probe: %d %s",
                    resp.status_code,
                    resp.reason_phrase,
                )
                # The MCP endpoint may return various status codes for GET;
                # we just need it to be reachable (not a connection error)

            config = MCPServerConfig(
                name="Microsoft Learn",
                url=self.url,
                transport=TransportType.AUTO,
                timeout=15.0,
                sse_read_timeout=30.0,
            )
            self._mcp_client = MCPClient(config)
            await self._mcp_client.connect()
            self._connected = True

            # Cache available tool names
            tools = await self._mcp_client.list_tools()
            self._available_tools = [t["name"] for t in tools]
            logger.info(
                "MS Learn MCP connected — tools: %s",
                ", ".join(self._available_tools),
            )
            return True

        except Exception as e:
            logger.warning("MS Learn MCP connection failed: %s", e)
            self._connected = False
            return False

    async def disconnect(self):
        """Disconnect from the MS Learn MCP server."""
        if self._mcp_client:
            try:
                await self._mcp_client.disconnect()
            except Exception:
                pass
        self._connected = False
        self._mcp_client = None

    async def search_docs(self, query: str, max_results: int = 5) -> str:
        """
        Search Microsoft Learn documentation.

        Args:
            query: Search query string
            max_results: Maximum number of results

        Returns:
            Formatted string with search results, or empty string if unavailable
        """
        if not self._connected or not self._mcp_client:
            return ""

        # Try the documented tool names
        tool_name = None
        for candidate in [
            "microsoft_docs_search",
            "search",
            "docs_search",
            "microsoft_learn_search",
        ]:
            if candidate in self._available_tools:
                tool_name = candidate
                break

        if not tool_name:
            logger.warning(
                "No search tool found among: %s", self._available_tools
            )
            return ""

        try:
            result = await self._mcp_client.call_tool(
                tool_name,
                {"query": query},
            )

            if result.get("isError"):
                logger.warning("MS Learn search error: %s", result)
                return ""

            # Extract text content from result
            texts = []
            for block in result.get("content", []):
                if block.get("type") == "text" and block.get("text"):
                    texts.append(block["text"])

            return "\n\n".join(texts[:max_results])

        except Exception as e:
            logger.warning("MS Learn search failed for '%s': %s", query, e)
            return ""

    async def fetch_doc(self, url: str) -> str:
        """
        Fetch a specific Microsoft Learn documentation page.

        Args:
            url: Full URL of the MS Learn page

        Returns:
            Markdown content of the page, or empty string if unavailable
        """
        if not self._connected or not self._mcp_client:
            return ""

        tool_name = None
        for candidate in [
            "microsoft_docs_fetch",
            "fetch",
            "docs_fetch",
            "microsoft_learn_fetch",
        ]:
            if candidate in self._available_tools:
                tool_name = candidate
                break

        if not tool_name:
            return ""

        try:
            result = await self._mcp_client.call_tool(
                tool_name,
                {"url": url},
            )

            if result.get("isError"):
                return ""

            texts = []
            for block in result.get("content", []):
                if block.get("type") == "text" and block.get("text"):
                    texts.append(block["text"])

            return "\n\n".join(texts)

        except Exception as e:
            logger.warning("MS Learn fetch failed for '%s': %s", url, e)
            return ""

    async def search_code_samples(
        self, query: str, language: str = "csharp"
    ) -> str:
        """
        Search for code samples in Microsoft Learn.

        Args:
            query: Search query
            language: Programming language filter

        Returns:
            Formatted string with code samples
        """
        if not self._connected or not self._mcp_client:
            return ""

        tool_name = None
        for candidate in [
            "microsoft_code_sample_search",
            "code_sample_search",
            "search_code_samples",
        ]:
            if candidate in self._available_tools:
                tool_name = candidate
                break

        if not tool_name:
            return ""

        try:
            result = await self._mcp_client.call_tool(
                tool_name,
                {"query": query, "language": language},
            )

            if result.get("isError"):
                return ""

            texts = []
            for block in result.get("content", []):
                if block.get("type") == "text" and block.get("text"):
                    texts.append(block["text"])

            return "\n\n".join(texts)

        except Exception as e:
            logger.warning("MS Learn code search failed: %s", e)
            return ""

    async def get_docs_for_request(self, customer_request: str) -> str:
        """
        Analyze a customer request and fetch relevant MS Learn documentation.

        This is the main entry point for the planner — it maps the request
        to relevant search queries and returns a formatted documentation
        context string.

        Args:
            customer_request: The customer's demo request text

        Returns:
            Formatted documentation context string for inclusion in the
            planner prompt, or empty string if no docs are available.
        """
        if not self._connected:
            return ""

        request_lower = customer_request.lower()

        # Find matching topics
        queries: list[str] = []
        for topic, topic_queries in TOPIC_SEARCH_MAP.items():
            if topic in request_lower:
                queries.extend(topic_queries[:1])  # Take first query per topic

        # Fallback: use the request directly
        if not queries:
            queries = [f"Dynamics 365 {customer_request}"]

        # Limit to 3 queries to avoid excessive API calls
        queries = queries[:3]

        # Run searches in parallel
        tasks = [self.search_docs(q, max_results=3) for q in queries]
        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Combine results
        doc_snippets: list[str] = []
        for i, result in enumerate(results):
            if isinstance(result, str) and result.strip():
                doc_snippets.append(result)
            elif isinstance(result, Exception):
                logger.warning(
                    "Doc search failed for '%s': %s", queries[i], result
                )

        if not doc_snippets:
            return ""

        # Format for inclusion in planner prompt
        combined = "\n\n---\n\n".join(doc_snippets)
        return (
            f"## Microsoft Learn Documentation Context\n\n"
            f"The following documentation snippets from learn.microsoft.com "
            f"are relevant to this demo request. Use them to ensure accuracy "
            f"in narration and navigation paths:\n\n{combined}"
        )

    # ---- Context manager ----

    async def __aenter__(self):
        await self.connect()
        return self

    async def __aexit__(self, exc_type, exc_val, exc_tb):
        await self.disconnect()
        return False
