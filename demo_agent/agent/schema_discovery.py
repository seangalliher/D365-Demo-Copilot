"""
D365 Demo Copilot — Schema Discovery via Dataverse MCP

Queries the Dataverse MCP server to discover real table schemas,
field names, and form metadata. This information is fed to the
LLM planner so it generates plans with accurate field references
instead of guessing selectors.
"""

from __future__ import annotations

import logging
from typing import Any, Optional

from ..mcp.client import MCPClient, MCPServerConfig, TransportType

logger = logging.getLogger("demo_agent.schema_discovery")

# Map of demo-relevant entities to their collection names and display names
DEMO_ENTITIES: dict[str, dict[str, str]] = {
    "msdyn_timeentry": {
        "collection": "msdyn_timeentries",
        "display": "Time Entry",
    },
    "msdyn_expense": {
        "collection": "msdyn_expenses",
        "display": "Expense",
    },
    "msdyn_project": {
        "collection": "msdyn_projects",
        "display": "Project",
    },
    "msdyn_projecttask": {
        "collection": "msdyn_projecttasks",
        "display": "Project Task",
    },
    "msdyn_projectteam": {
        "collection": "msdyn_projectteams",
        "display": "Project Team Member",
    },
    "bookableresource": {
        "collection": "bookableresources",
        "display": "Bookable Resource",
    },
    "msdyn_approval": {
        "collection": "msdyn_approvals",
        "display": "Approval",
    },
    "msdyn_actual": {
        "collection": "msdyn_actuals",
        "display": "Actual",
    },
}

# ── Built-in field definitions (fallback when MCP is not available) ───────
# Sourced from D365 Project Operations entity reference.  These give the
# planner enough information to generate correct selectors without MCP.
BUILTIN_SCHEMAS: dict[str, dict] = {
    "msdyn_timeentry": {
        "columns": [
            {"logicalName": "msdyn_date", "displayName": "Date", "type": "DateTime", "isRequired": True},
            {"logicalName": "msdyn_duration", "displayName": "Duration", "type": "Integer", "isRequired": True,
             "note": "Duration in minutes. D365 renders as a duration combo-box."},
            {"logicalName": "msdyn_start", "displayName": "Start", "type": "DateTime", "isRequired": False},
            {"logicalName": "msdyn_project", "displayName": "Project", "type": "Lookup", "isRequired": False,
             "target": "msdyn_project"},
            {"logicalName": "msdyn_projecttask", "displayName": "Project Task", "type": "Lookup", "isRequired": False,
             "target": "msdyn_projecttask"},
            {"logicalName": "msdyn_bookableresource", "displayName": "Bookable Resource", "type": "Lookup",
             "isRequired": False, "target": "bookableresource"},
            {"logicalName": "msdyn_description", "displayName": "Description", "type": "String", "isRequired": False,
             "note": "Free-text memo. On the form this is labeled 'External Comments' or 'Internal Comments'."},
            {"logicalName": "msdyn_type", "displayName": "Type", "type": "Picklist", "isRequired": True,
             "options": {"Work": 192350000, "Absence": 192350001, "Vacation": 192350002, "Overtime": 192350004}},
            {"logicalName": "msdyn_entrystatus", "displayName": "Entry Status", "type": "Picklist",
             "isRequired": False,
             "options": {"Draft": 192350000, "Returned": 192350001, "Approved": 192350002, "Submitted": 192350003}},
            {"logicalName": "msdyn_resourcecategory", "displayName": "Role", "type": "Lookup", "isRequired": False,
             "target": "bookableresourcecategory"},
            {"logicalName": "msdyn_externalDescription", "displayName": "External Comments", "type": "String",
             "isRequired": False},
        ],
    },
    "msdyn_expense": {
        "columns": [
            {"logicalName": "msdyn_name", "displayName": "Name", "type": "String", "isRequired": True},
            {"logicalName": "msdyn_amount", "displayName": "Amount", "type": "Money", "isRequired": True},
            {"logicalName": "msdyn_price", "displayName": "Price", "type": "Money", "isRequired": True},
            {"logicalName": "msdyn_quantity", "displayName": "Quantity", "type": "Decimal", "isRequired": True},
            {"logicalName": "msdyn_expensecategory", "displayName": "Expense Category", "type": "Lookup",
             "isRequired": True, "target": "msdyn_expensecategory"},
            {"logicalName": "msdyn_transactiondate", "displayName": "Transaction Date", "type": "DateTime",
             "isRequired": True},
            {"logicalName": "msdyn_project", "displayName": "Project", "type": "Lookup", "isRequired": False,
             "target": "msdyn_project"},
            {"logicalName": "msdyn_projecttask", "displayName": "Project Task", "type": "Lookup",
             "isRequired": False, "target": "msdyn_projecttask"},
            {"logicalName": "msdyn_description", "displayName": "Description", "type": "String",
             "isRequired": False},
            {"logicalName": "transactioncurrencyid", "displayName": "Currency", "type": "Lookup",
             "isRequired": True, "target": "transactioncurrency"},
            {"logicalName": "msdyn_salestaxamount", "displayName": "Sales Tax", "type": "Money",
             "isRequired": False},
            {"logicalName": "msdyn_expensestatus", "displayName": "Expense Status", "type": "Picklist",
             "isRequired": False,
             "options": {"Draft": 192350000, "Submitted": 192350001, "Approved": 192350002,
                         "Rejected": 192350003, "Posted": 192350004}},
        ],
    },
    "msdyn_project": {
        "columns": [
            {"logicalName": "msdyn_subject", "displayName": "Project Name", "type": "String", "isRequired": True},
            {"logicalName": "msdyn_projectmanager", "displayName": "Project Manager", "type": "Lookup",
             "isRequired": True, "target": "systemuser"},
            {"logicalName": "msdyn_contractorganizationalunitid", "displayName": "Contracting Unit",
             "type": "Lookup", "isRequired": True, "target": "msdyn_organizationalunit"},
            {"logicalName": "msdyn_customer", "displayName": "Customer", "type": "Lookup",
             "isRequired": False, "target": "account"},
            {"logicalName": "msdyn_description", "displayName": "Description", "type": "String",
             "isRequired": False},
            {"logicalName": "msdyn_schedulemode", "displayName": "Schedule Mode", "type": "Picklist",
             "isRequired": True,
             "options": {"Fixed Duration": 192350000, "Fixed Effort": 192350001, "Fixed Units": 192350002}},
            {"logicalName": "msdyn_scheduledstart", "displayName": "Start Date", "type": "DateTime",
             "isRequired": False},
            {"logicalName": "msdyn_finish", "displayName": "Finish Date", "type": "DateTime", "isRequired": False},
            {"logicalName": "msdyn_workhourtemplate", "displayName": "Work Hour Template", "type": "Lookup",
             "isRequired": False},
            {"logicalName": "transactioncurrencyid", "displayName": "Currency", "type": "Lookup",
             "isRequired": True, "target": "transactioncurrency"},
        ],
    },
    "msdyn_projecttask": {
        "columns": [
            {"logicalName": "msdyn_subject", "displayName": "Name", "type": "String", "isRequired": True},
            {"logicalName": "msdyn_project", "displayName": "Project", "type": "Lookup",
             "isRequired": True, "target": "msdyn_project"},
            {"logicalName": "msdyn_parenttask", "displayName": "Parent Task", "type": "Lookup",
             "isRequired": False, "target": "msdyn_projecttask"},
            {"logicalName": "msdyn_start", "displayName": "Start", "type": "DateTime", "isRequired": False},
            {"logicalName": "msdyn_finish", "displayName": "Finish", "type": "DateTime", "isRequired": False},
            {"logicalName": "msdyn_effort", "displayName": "Effort", "type": "Decimal", "isRequired": False},
            {"logicalName": "msdyn_duration", "displayName": "Duration", "type": "Integer", "isRequired": False},
            {"logicalName": "msdyn_ismilestone", "displayName": "Is Milestone", "type": "Boolean",
             "isRequired": False},
        ],
    },
    "msdyn_projectteam": {
        "columns": [
            {"logicalName": "msdyn_project", "displayName": "Project", "type": "Lookup",
             "isRequired": True, "target": "msdyn_project"},
            {"logicalName": "msdyn_bookableresourceid", "displayName": "Resource", "type": "Lookup",
             "isRequired": False, "target": "bookableresource"},
            {"logicalName": "msdyn_resourcecategory", "displayName": "Role", "type": "Lookup",
             "isRequired": False, "target": "bookableresourcecategory"},
            {"logicalName": "msdyn_allocationmethod", "displayName": "Allocation Method", "type": "Picklist",
             "isRequired": False},
            {"logicalName": "msdyn_billingtype", "displayName": "Billing Type", "type": "Picklist",
             "isRequired": False},
        ],
    },
    "bookableresource": {
        "columns": [
            {"logicalName": "name", "displayName": "Name", "type": "String", "isRequired": True},
            {"logicalName": "resourcetype", "displayName": "Resource Type", "type": "Picklist", "isRequired": True,
             "options": {"Generic": 1, "Contact": 2, "User": 3, "Equipment": 4}},
            {"logicalName": "msdyn_organizationalunit", "displayName": "Organizational Unit", "type": "Lookup",
             "isRequired": False, "target": "msdyn_organizationalunit"},
            {"logicalName": "msdyn_targetutilization", "displayName": "Target Utilization", "type": "Integer",
             "isRequired": False},
        ],
    },
    "msdyn_approval": {
        "columns": [
            {"logicalName": "subject", "displayName": "Subject", "type": "String", "isRequired": False},
            {"logicalName": "msdyn_approvalstatus", "displayName": "Approval Status", "type": "Picklist",
             "isRequired": False,
             "options": {"Saved": 192350000, "Pending Approval": 192350005, "Rejected": 192350004,
                         "Approved": 192350003, "Recalled": 192350009}},
        ],
    },
}

# ── F&O Built-in schemas (for D365 Finance & Supply Chain Management) ─────
# These use data-dyn-controlname patterns instead of data-id patterns.
# Not in Dataverse — these are X++ entities with separate selector conventions.
FO_BUILTIN_SCHEMAS: dict[str, dict] = {
    "LedgerJournalTable": {
        "platform": "fo",
        "menu_item": "LedgerJournalTable",
        "display": "General Journal",
        "columns": [
            {"controlName": "LedgerJournalTable_JournalNum", "displayName": "Journal batch number", "type": "String", "isRequired": True},
            {"controlName": "LedgerJournalTable_JournalName", "displayName": "Name", "type": "Lookup", "isRequired": True},
            {"controlName": "LedgerJournalTable_Description", "displayName": "Description", "type": "String", "isRequired": False},
            {"controlName": "LedgerJournalTrans_AccountNum", "displayName": "Main account", "type": "SegmentedEntry", "isRequired": True},
            {"controlName": "LedgerJournalTrans_Txt", "displayName": "Description (line)", "type": "String", "isRequired": False},
            {"controlName": "LedgerJournalTrans_AmountCurDebit", "displayName": "Debit", "type": "Real", "isRequired": False},
            {"controlName": "LedgerJournalTrans_AmountCurCredit", "displayName": "Credit", "type": "Real", "isRequired": False},
            {"controlName": "LedgerJournalTrans_OffsetAccountNum", "displayName": "Offset account", "type": "SegmentedEntry", "isRequired": False},
        ],
    },
    "PurchTable": {
        "platform": "fo",
        "menu_item": "PurchTableListPage",
        "display": "Purchase Order",
        "columns": [
            {"controlName": "PurchTable_PurchId", "displayName": "Purchase order", "type": "String", "isRequired": True, "note": "Auto-generated"},
            {"controlName": "PurchTable_OrderAccount", "displayName": "Vendor account", "type": "Lookup", "isRequired": True},
            {"controlName": "PurchTable_PurchName", "displayName": "Vendor name", "type": "String", "isRequired": False},
            {"controlName": "PurchLine_ItemId", "displayName": "Item number", "type": "Lookup", "isRequired": True},
            {"controlName": "PurchLine_PurchQty", "displayName": "Quantity", "type": "Real", "isRequired": True},
            {"controlName": "PurchLine_PurchPrice", "displayName": "Unit price", "type": "Real", "isRequired": False},
            {"controlName": "PurchTable_DeliveryDate", "displayName": "Delivery date", "type": "Date", "isRequired": False},
            {"controlName": "PurchTable_CurrencyCode", "displayName": "Currency", "type": "Lookup", "isRequired": True},
        ],
    },
    "SalesTable": {
        "platform": "fo",
        "menu_item": "SalesTableListPage",
        "display": "Sales Order",
        "columns": [
            {"controlName": "SalesTable_SalesId", "displayName": "Sales order", "type": "String", "isRequired": True, "note": "Auto-generated"},
            {"controlName": "SalesTable_CustAccount", "displayName": "Customer account", "type": "Lookup", "isRequired": True},
            {"controlName": "SalesLine_ItemId", "displayName": "Item number", "type": "Lookup", "isRequired": True},
            {"controlName": "SalesLine_SalesQty", "displayName": "Quantity", "type": "Real", "isRequired": True},
            {"controlName": "SalesLine_SalesPrice", "displayName": "Unit price", "type": "Real", "isRequired": False},
            {"controlName": "SalesTable_DeliveryDate", "displayName": "Delivery date", "type": "Date", "isRequired": False},
            {"controlName": "SalesTable_CurrencyCode", "displayName": "Currency", "type": "Lookup", "isRequired": True},
        ],
    },
    "VendTable": {
        "platform": "fo",
        "menu_item": "VendTableListPage",
        "display": "Vendor",
        "columns": [
            {"controlName": "VendTable_AccountNum", "displayName": "Vendor account", "type": "String", "isRequired": True},
            {"controlName": "DirPartyTable_Name", "displayName": "Name", "type": "String", "isRequired": True},
            {"controlName": "VendTable_VendGroup", "displayName": "Vendor group", "type": "Lookup", "isRequired": True},
            {"controlName": "VendTable_Currency", "displayName": "Currency", "type": "Lookup", "isRequired": True},
            {"controlName": "VendTable_PaymTermId", "displayName": "Payment terms", "type": "Lookup", "isRequired": False},
        ],
    },
    "CustTable": {
        "platform": "fo",
        "menu_item": "CustTableListPage",
        "display": "Customer",
        "columns": [
            {"controlName": "CustTable_AccountNum", "displayName": "Customer account", "type": "String", "isRequired": True},
            {"controlName": "DirPartyTable_Name", "displayName": "Name", "type": "String", "isRequired": True},
            {"controlName": "CustTable_CustGroup", "displayName": "Customer group", "type": "Lookup", "isRequired": True},
            {"controlName": "CustTable_Currency", "displayName": "Currency", "type": "Lookup", "isRequired": True},
            {"controlName": "CustTable_PaymTermId", "displayName": "Payment terms", "type": "Lookup", "isRequired": False},
        ],
    },
    "ProdTable": {
        "platform": "fo",
        "menu_item": "ProdTableListPage",
        "display": "Production Order",
        "columns": [
            {"controlName": "ProdTable_ProdId", "displayName": "Production number", "type": "String", "isRequired": True, "note": "Auto-generated"},
            {"controlName": "ProdTable_ItemId", "displayName": "Item number", "type": "Lookup", "isRequired": True},
            {"controlName": "ProdTable_QtySched", "displayName": "Quantity", "type": "Real", "isRequired": True},
            {"controlName": "ProdTable_SchedDate", "displayName": "Scheduled date", "type": "Date", "isRequired": False},
        ],
    },
    "AssetTable": {
        "platform": "fo",
        "menu_item": "AssetTable",
        "display": "Fixed Asset",
        "columns": [
            {"controlName": "AssetTable_AssetId", "displayName": "Fixed asset number", "type": "String", "isRequired": True},
            {"controlName": "AssetTable_Name", "displayName": "Name", "type": "String", "isRequired": True},
            {"controlName": "AssetTable_AssetGroup", "displayName": "Fixed asset group", "type": "Lookup", "isRequired": True},
            {"controlName": "AssetTable_Location", "displayName": "Location", "type": "Lookup", "isRequired": False},
        ],
    },
}


class SchemaDiscovery:
    """
    Discovers Dataverse table schemas via MCP to provide
    accurate field information for demo planning.
    """

    def __init__(self, mcp_url: str, auth_headers: Optional[dict[str, str]] = None):
        self._mcp_url = mcp_url
        self._auth_headers = auth_headers or {}
        self._client: Optional[MCPClient] = None
        self._schema_cache: dict[str, dict] = {}

    async def connect(self) -> bool:
        """Connect to the Dataverse MCP server.

        Returns True if live MCP is available, False if falling back
        to built-in schemas.  Never raises — MCP is always optional.

        When ``auth_headers`` are provided (from DataverseAuth), they
        are passed to both the HTTP probe and the MCP transport.
        Without them, a 401 falls back gracefully to built-in schemas.
        """
        probe_headers: dict[str, str] = {"Content-Type": "application/json"}
        probe_headers.update(self._auth_headers)

        # Quick HTTP probe to check if the endpoint is reachable and
        # doesn't immediately reject with 401/403.
        try:
            import httpx
            async with httpx.AsyncClient(timeout=10.0) as http:
                resp = await http.post(
                    self._mcp_url,
                    json={"jsonrpc": "2.0", "method": "initialize", "id": 1, "params": {}},
                    headers=probe_headers,
                )
                if resp.status_code in (401, 403):
                    logger.info(
                        "Dataverse MCP returned %d — using built-in schemas",
                        resp.status_code,
                    )
                    return False
        except Exception as e:
            logger.info("Dataverse MCP probe failed (%s) — using built-in schemas", e)
            return False

        # Endpoint is reachable and doesn't reject auth — try full MCP
        try:
            config = MCPServerConfig(
                name="Dataverse",
                url=self._mcp_url,
                transport=TransportType.AUTO,
                headers=dict(self._auth_headers),
                timeout=30.0,
            )
            self._client = MCPClient(config)
            await self._client.connect()
            logger.info("Schema discovery connected to Dataverse MCP")
            return True
        except BaseException as e:
            logger.warning(
                "MCP connection failed after probe (%s) — using built-in schemas",
                type(e).__name__,
            )
            self._client = None
            return False

    async def disconnect(self):
        """Disconnect from the Dataverse MCP server."""
        if self._client:
            await self._client.disconnect()
            self._client = None

    async def describe_table(self, table_name: str) -> dict | None:
        """
        Get the schema for a Dataverse table.

        Tries MCP first, then falls back to built-in schemas.
        Returns parsed schema dict or None on failure.
        """
        if table_name in self._schema_cache:
            return self._schema_cache[table_name]

        # Try live MCP if available
        if self._client:
            try:
                result = await self._client.call_tool(
                    "describe_table",
                    {"table": table_name},
                )
                if not result.get("isError"):
                    text = ""
                    for block in result.get("content", []):
                        if block.get("type") == "text":
                            text += block["text"]
                    if text:
                        import json
                        try:
                            schema = json.loads(text)
                        except json.JSONDecodeError:
                            schema = {"raw_description": text}
                        self._schema_cache[table_name] = schema
                        logger.info(
                            "Fetched live schema for %s — %d fields",
                            table_name,
                            len(schema.get("columns", schema.get("attributes", [])))
                            if isinstance(schema, dict) else 0,
                        )
                        return schema
            except Exception as e:
                logger.debug("MCP describe_table failed for %s: %s", table_name, e)

        # Fallback to built-in schema knowledge
        if table_name in BUILTIN_SCHEMAS:
            schema = BUILTIN_SCHEMAS[table_name]
            self._schema_cache[table_name] = schema
            logger.info(
                "Using built-in schema for %s — %d fields",
                table_name, len(schema.get("columns", [])),
            )
            return schema

        # Fallback to F&O built-in schemas
        if table_name in FO_BUILTIN_SCHEMAS:
            schema = FO_BUILTIN_SCHEMAS[table_name]
            self._schema_cache[table_name] = schema
            logger.info(
                "Using F&O built-in schema for %s — %d fields",
                table_name, len(schema.get("columns", [])),
            )
            return schema

        logger.warning("No schema available for %s", table_name)
        return None

    async def get_entity_schemas_for_request(
        self, request: str
    ) -> dict[str, Any]:
        """
        Determine which entities are relevant to a demo request
        and fetch their schemas.

        Args:
            request: The customer's demo request text

        Returns:
            Dict mapping entity names to their schema info
        """
        request_lower = request.lower()

        # Determine relevant entities based on request keywords
        relevant: list[str] = []

        keyword_map = {
            "msdyn_timeentry": [
                "time", "timesheet", "hours", "weekly", "time entry",
                "duration", "work hours",
            ],
            "msdyn_expense": [
                "expense", "receipt", "travel", "mileage", "per diem",
                "reimbursement",
            ],
            "msdyn_project": [
                "project", "wbs", "schedule", "task", "milestone",
                "work breakdown",
            ],
            "msdyn_projecttask": [
                "task", "wbs", "work breakdown", "milestone", "schedule",
            ],
            "msdyn_projectteam": [
                "team", "resource", "staff", "assign", "member",
                "booking",
            ],
            "bookableresource": [
                "resource", "booking", "utilization", "capacity",
                "consultant",
            ],
            "msdyn_approval": [
                "approv", "submit", "review", "reject", "workflow",
            ],
            "msdyn_actual": [
                "actual", "cost", "revenue", "financial", "billing",
                "invoice",
            ],
        }

        # F&O entity keywords
        fo_keyword_map = {
            "LedgerJournalTable": [
                "journal", "general ledger", "ledger", "gl entry",
                "journal entry", "posting",
            ],
            "PurchTable": [
                "purchase order", "procurement", "po", "procure",
                "buy", "purchasing",
            ],
            "SalesTable": [
                "sales order", "sell", "sales", "order to cash",
            ],
            "VendTable": [
                "vendor", "supplier", "ap", "accounts payable",
            ],
            "CustTable": [
                "customer", "client", "ar", "accounts receivable",
            ],
            "ProdTable": [
                "production", "manufacturing", "bom", "production order",
                "shop floor",
            ],
            "AssetTable": [
                "fixed asset", "asset", "depreciation", "capital",
            ],
        }

        for entity, keywords in keyword_map.items():
            if any(kw in request_lower for kw in keywords):
                if entity not in relevant:
                    relevant.append(entity)

        # Check F&O entities
        for entity, keywords in fo_keyword_map.items():
            if any(kw in request_lower for kw in keywords):
                if entity not in relevant:
                    relevant.append(entity)

        # Always include the primary entity and approval if submitting
        if not relevant:
            relevant = ["msdyn_timeentry", "msdyn_project"]

        # Fetch schemas
        schemas: dict[str, Any] = {}
        for entity in relevant:
            schema = await self.describe_table(entity)
            if schema:
                schemas[entity] = schema

        logger.info(
            "Discovered schemas for %d/%d relevant entities: %s",
            len(schemas), len(relevant), list(schemas.keys()),
        )
        return schemas

    def format_schemas_for_prompt(self, schemas: dict[str, Any]) -> str:
        """
        Format discovered schemas into a concise prompt section
        the LLM can use when generating plans.

        Focuses on field names, display names, types, and whether
        they're required — the info needed to write correct selectors.
        Handles both CE (data-id) and F&O (data-dyn-controlname) patterns.
        """
        if not schemas:
            return ""

        # Separate CE vs F&O schemas
        ce_schemas = {k: v for k, v in schemas.items() if not v.get("platform") == "fo"}
        fo_schemas = {k: v for k, v in schemas.items() if v.get("platform") == "fo"}

        lines = [
            "## Discovered Field Schemas",
            "",
        ]

        # CE schemas section
        if ce_schemas:
            lines.extend([
                "### CE / Dataverse Field Schemas",
                "",
                "Use these REAL field names when generating selectors. "
                "D365 CE form fields use `data-id` attributes based on the logical name.",
                "",
                "#### Selector Patterns for D365 CE Form Fields",
                "- Text/number input: `input[data-id=\"{logical_name}.fieldControl-text-box-text\"]`",
                "- Lookup input: `input[data-id=\"{logical_name}.fieldControl-LookupResultsDropdown_{logical_name}_textInputBox_with_filter_new\"]`",
                "- Option set: `select[data-id=\"{logical_name}.fieldControl-option-set-select\"]`",
                "- Date picker: `input[data-id=\"{logical_name}.fieldControl-date-time-input\"]`",
                "- Duration: `input[data-id=\"{logical_name}.fieldControl-duration-combobox-text\"]`",
                "",
            ])

            for entity_name, schema in ce_schemas.items():
                display = DEMO_ENTITIES.get(entity_name, {}).get("display", entity_name)
                lines.append(f"#### {display} (`{entity_name}`)")
                lines.append("")
                lines.append("| Logical Name | Display Name | Type | Required | Notes |")
                lines.append("|---|---|---|---|---|")

                columns = (
                    schema.get("columns")
                    or schema.get("attributes")
                    or schema.get("fields")
                    or []
                )

                if isinstance(columns, list):
                    for col in columns:
                        if isinstance(col, dict):
                            logical = col.get("logicalName", col.get("logical_name", col.get("name", "?")))
                            display_name = col.get("displayName", col.get("display_name", col.get("label", logical)))
                            col_type = col.get("type", col.get("attributeType", col.get("dataType", "?")))
                            required = col.get("isRequired", col.get("required", col.get("requiredLevel", "?")))
                            note = col.get("note", "")
                            options = col.get("options")
                            if options and isinstance(options, dict):
                                opts_str = ", ".join(f"{k}={v}" for k, v in options.items())
                                note = f"Options: {opts_str}. {note}".strip()
                            target = col.get("target")
                            if target:
                                note = f"Lookup→{target}. {note}".strip()
                            lines.append(
                                f"| `{logical}` | {display_name} | {col_type} | {required} | {note} |"
                            )
                elif isinstance(schema, dict) and "raw_description" in schema:
                    lines.append(schema["raw_description"][:2000])

                lines.append("")

        # F&O schemas section
        if fo_schemas:
            lines.extend([
                "### F&O / FinOps Field Schemas",
                "",
                "Use `data-dyn-controlname` selectors for D365 Finance & SCM forms.",
                "",
                "#### Selector Patterns for D365 F&O Form Fields",
                "- Input field: `input[data-dyn-controlname=\"{controlName}\"]`",
                "- Dropdown: `select[data-dyn-controlname=\"{controlName}\"]`",
                "- Lookup: `input[data-dyn-controlname=\"{controlName}\"]` then click adjacent lookup button",
                "- Tab/FastTab: `button[data-dyn-controlname=\"{TabName}\"]`",
                "",
            ])

            for entity_name, schema in fo_schemas.items():
                display = schema.get("display", entity_name)
                menu_item = schema.get("menu_item", "")
                lines.append(f"#### {display} (`{entity_name}`, menu item: `{menu_item}`)")
                lines.append("")
                lines.append("| Control Name | Display Name | Type | Required | Notes |")
                lines.append("|---|---|---|---|---|")

                columns = schema.get("columns", [])
                for col in columns:
                    if isinstance(col, dict):
                        ctrl = col.get("controlName", "?")
                        display_name = col.get("displayName", ctrl)
                        col_type = col.get("type", "?")
                        required = col.get("isRequired", False)
                        note = col.get("note", "")
                        lines.append(
                            f"| `{ctrl}` | {display_name} | {col_type} | {required} | {note} |"
                        )

                lines.append("")

        return "\n".join(lines)


class PageIntrospector:
    """
    Introspects the live D365 page DOM to discover actual
    form fields, their selectors, and current values.

    This runs in the browser via Playwright and gives the
    executor real selector information.
    """

    # JavaScript to extract all visible form fields from a D365 page
    INTROSPECT_FIELDS_JS = """
    () => {
        const fields = [];

        // Method 1: Find all inputs with data-id (D365 standard)
        document.querySelectorAll('input[data-id], select[data-id], textarea[data-id]').forEach(el => {
            const dataId = el.getAttribute('data-id') || '';
            const ariaLabel = el.getAttribute('aria-label') || '';
            const type = el.tagName.toLowerCase();
            const visible = el.offsetParent !== null;
            if (visible && (dataId || ariaLabel)) {
                fields.push({
                    selector_data_id: dataId ? `[data-id="${dataId}"]` : null,
                    selector_aria: ariaLabel ? `[aria-label="${ariaLabel}"]` : null,
                    data_id: dataId,
                    aria_label: ariaLabel,
                    tag: type,
                    input_type: el.getAttribute('type') || '',
                    value: el.value || '',
                    placeholder: el.getAttribute('placeholder') || '',
                    required: el.hasAttribute('required') || el.getAttribute('aria-required') === 'true',
                    readonly: el.hasAttribute('readonly') || el.getAttribute('aria-readonly') === 'true',
                    disabled: el.hasAttribute('disabled'),
                });
            }
        });

        // Method 2: Find labeled sections/fields from form sections
        document.querySelectorAll('section[data-id] label').forEach(label => {
            const text = label.textContent?.trim();
            const forEl = label.getAttribute('for');
            if (text && forEl) {
                const input = document.getElementById(forEl);
                if (input && input.offsetParent !== null) {
                    const dataId = input.getAttribute('data-id') || '';
                    if (!fields.some(f => f.data_id === dataId)) {
                        fields.push({
                            selector_data_id: dataId ? `[data-id="${dataId}"]` : null,
                            selector_aria: `[aria-label="${text}"]`,
                            data_id: dataId,
                            aria_label: text,
                            tag: input.tagName.toLowerCase(),
                            input_type: input.getAttribute('type') || '',
                            value: input.value || '',
                            placeholder: input.getAttribute('placeholder') || '',
                            required: input.hasAttribute('required'),
                            readonly: input.hasAttribute('readonly'),
                            disabled: input.hasAttribute('disabled'),
                        });
                    }
                }
            }
        });

        // Method 3: Find command bar buttons
        const buttons = [];
        document.querySelectorAll('button[data-id], button[aria-label]').forEach(btn => {
            const visible = btn.offsetParent !== null;
            if (visible) {
                buttons.push({
                    data_id: btn.getAttribute('data-id') || '',
                    aria_label: btn.getAttribute('aria-label') || '',
                    text: btn.textContent?.trim()?.substring(0, 60) || '',
                });
            }
        });

        return { fields, buttons, url: window.location.href };
    }
    """

    @staticmethod
    async def discover_page_fields(page) -> dict:
        """
        Run DOM introspection on the current D365 page.

        Returns dict with:
        - fields: list of form field descriptors
        - buttons: list of visible button descriptors
        - url: current page URL
        """
        try:
            result = await page.evaluate(PageIntrospector.INTROSPECT_FIELDS_JS)
            field_count = len(result.get("fields", []))
            btn_count = len(result.get("buttons", []))
            logger.info(
                "Page introspection: %d fields, %d buttons on %s",
                field_count, btn_count, result.get("url", "?"),
            )
            return result
        except Exception as e:
            logger.warning("Page introspection failed: %s", e)
            return {"fields": [], "buttons": [], "url": ""}

    @staticmethod
    def format_fields_for_prompt(page_info: dict) -> str:
        """Format discovered page fields into prompt-friendly text."""
        lines = ["## Live Page Fields (discovered from DOM)"]
        lines.append(f"Current URL: {page_info.get('url', 'unknown')}")
        lines.append("")

        fields = page_info.get("fields", [])
        if fields:
            lines.append("### Form Fields")
            lines.append("| Selector | Label | Type | Required | Current Value |")
            lines.append("|---|---|---|---|---|")
            for f in fields[:40]:  # Cap at 40 fields
                sel = f.get("selector_data_id") or f.get("selector_aria", "?")
                label = f.get("aria_label", "")
                tag = f.get("tag", "input")
                req = "Yes" if f.get("required") else ""
                val = (f.get("value", "") or "")[:30]
                lines.append(f"| `{sel}` | {label} | {tag} | {req} | {val} |")
            lines.append("")

        buttons = page_info.get("buttons", [])
        if buttons:
            lines.append("### Visible Buttons")
            for b in buttons[:20]:
                did = b.get("data_id", "")
                aria = b.get("aria_label", "")
                text = b.get("text", "")
                sel = f'button[data-id="{did}"]' if did else f'button[aria-label="{aria}"]'
                lines.append(f"- `{sel}` — {text or aria}")
            lines.append("")

        if not fields and not buttons:
            lines.append("(No form fields discovered on current page)")

        return "\n".join(lines)
