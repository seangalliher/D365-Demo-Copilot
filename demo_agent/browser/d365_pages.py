"""
D365 Demo Copilot — D365-specific Page Helpers

Provides high-level navigation helpers specific to Dynamics 365
Model-Driven Apps and Project Operations screens.
"""

from __future__ import annotations

import logging
from typing import Optional

from playwright.async_api import Page

logger = logging.getLogger("demo_agent.browser.d365")


# Common D365 selectors
class D365Selectors:
    """Commonly used CSS/aria selectors for D365 Model-Driven App elements."""

    # Navigation
    SITEMAP_BUTTON = 'button[data-id="navbutton"]'
    AREA_SWITCHER = 'button[data-id="sitemap-areaSwitcher-expand-btn"]'
    NAV_GROUP = 'ul[data-id="sitemap-body"]'

    # Sitemap entity links
    NAV_TIME_ENTRIES = 'li[data-id="sitemap-entity-msdyn_TimeEntrySubArea"]'
    NAV_EXPENSES = 'li[data-id="sitemap-entity-msdyn_ExpensesSubArea"]'
    NAV_PROJECTS = 'li[data-id="sitemap-entity-msdyn_ProjectSubArea"]'
    NAV_APPROVALS = 'li[data-id="sitemap-entity-msdyn_ProjectApprovalSubArea"]'
    NAV_RESOURCES = 'li[data-id="sitemap-entity-msdyn_ResourceSubArea"]'
    NAV_DASHBOARDS = 'li[data-id="sitemap-entity-msdyn_PracticeManagementSubArea"]'
    NAV_SCHEDULE_BOARD = 'li[data-id="sitemap-entity-msdyn_ScheduleBoardSettingsSubArea"]'
    NAV_RESOURCE_UTILIZATION = 'li[data-id="sitemap-entity-msdyn_ResourceUtilizationSubArea"]'
    NAV_CONTRACT_WORKERS = 'li[data-id="sitemap-entity-msdyn_ContractWorkersSubArea"]'
    NAV_ROLES = 'li[data-id="sitemap-entity-msdyn_BookableResourceBookingSubArea"]'
    NAV_VENDORS = 'li[data-id="sitemap-entity-msdyn_VendorsSubArea"]'

    # Forms
    FORM_SAVE = 'button[data-id="edit-form-save-btn"]'
    FORM_SAVE_CLOSE = 'button[data-id="edit-form-save-and-close-btn"]'
    FORM_NEW = 'button[data-id="edit-form-new-btn"]'
    FORM_DELETE = 'button[data-id="edit-form-delete-btn"]'
    FORM_HEADER = 'h1[data-id="header_title"]'
    COMMAND_BAR = 'ul[data-id="CommandBar"]'
    QUICK_CREATE = 'button[data-id="quickCreateLauncher"]'

    # Grid / List views
    GRID_CONTAINER = 'div[data-id="data-set-body"]'
    GRID_ROW = 'div[data-id="cell-0-1"]'  # First data cell
    VIEW_SELECTOR = 'span[data-id="view-selector"]'
    SEARCH_BOX = 'input[data-id="quickFind_text_1"]'

    # Tabs & Sections
    TAB_LIST = 'ul[role="tablist"]'
    SECTION = 'section[data-id]'

    # Notifications
    NOTIFICATION_BAR = 'div[data-id="notificationWrapper"]'
    BUSINESS_PROCESS_FLOW = 'div[data-id="BPFContainer"]'

    # Loading
    LOADING_SPINNER = '#progressContainer'
    FORM_LOADING = 'div.ms-Spinner'

    # Mapping from display names to sitemap data-ids
    NAV_MAP = {
        "time entries": 'li[data-id="sitemap-entity-msdyn_TimeEntrySubArea"]',
        "expenses": 'li[data-id="sitemap-entity-msdyn_ExpensesSubArea"]',
        "projects": 'li[data-id="sitemap-entity-msdyn_ProjectSubArea"]',
        "approvals": 'li[data-id="sitemap-entity-msdyn_ProjectApprovalSubArea"]',
        "resources": 'li[data-id="sitemap-entity-msdyn_ResourceSubArea"]',
        "dashboards": 'li[data-id="sitemap-entity-msdyn_PracticeManagementSubArea"]',
        "schedule board": 'li[data-id="sitemap-entity-msdyn_ScheduleBoardSettingsSubArea"]',
        "resource utilization": 'li[data-id="sitemap-entity-msdyn_ResourceUtilizationSubArea"]',
        "contract workers": 'li[data-id="sitemap-entity-msdyn_ContractWorkersSubArea"]',
        "roles": 'li[data-id="sitemap-entity-msdyn_BookableResourceBookingSubArea"]',
        "vendors": 'li[data-id="sitemap-entity-msdyn_VendorsSubArea"]',
        "project reports": 'li[data-id="sitemap-entity-msdyn_ProjectReportsSubArea"]',
    }


class D365FOSelectors:
    """Commonly used CSS selectors for D365 Finance & Operations (FinOps) apps."""

    # Navigation
    NAV_SEARCH = 'input[aria-label="Navigation search"], input[name="NavigationSearchBox"]'
    NAV_PANE = 'div.navigation-groups'
    COMPANY_PICKER = 'input[aria-label="Company"]'

    # Action Pane
    NEW_BUTTON = 'button[data-dyn-controlname="SystemDefinedNewButton"]'
    SAVE_BUTTON = 'button[data-dyn-controlname="SystemDefinedSaveButton"]'
    DELETE_BUTTON = 'button[data-dyn-controlname="SystemDefinedDeleteButton"]'
    CLOSE_BUTTON = 'button[data-dyn-controlname="SystemDefinedCloseButton"]'
    FILTER_PANE = 'button[data-dyn-controlname="FilterPaneButton"]'

    # Grid
    GRID_ROWS = 'div[data-dyn-role="Grid"] tr[data-dyn-row-id]'
    GRID_HEADER = 'div[data-dyn-role="Grid"] th'

    # Loading
    LOADING_INDICATOR = 'div.busy-indicator-container'
    MESSAGE_BAR = 'div[data-dyn-controlname="MessageBar"]'

    # F&O menu item URL map
    MENU_ITEMS = {
        "general journal": "LedgerJournalTable",
        "general journals": "LedgerJournalTable",
        "chart of accounts": "MainAccountListPage",
        "main accounts": "MainAccountListPage",
        "vendors": "VendTableListPage",
        "vendor": "VendTableListPage",
        "purchase orders": "PurchTableListPage",
        "purchase order": "PurchTableListPage",
        "sales orders": "SalesTableListPage",
        "sales order": "SalesTableListPage",
        "customers": "CustTableListPage",
        "customer": "CustTableListPage",
        "released products": "InventTableListPage",
        "inventory": "InventTableListPage",
        "production orders": "ProdTableListPage",
        "production order": "ProdTableListPage",
        "fixed assets": "AssetTable",
        "fixed asset": "AssetTable",
        "budget entries": "BudgetRegisterEntryListPage",
        "bank accounts": "BankAccountTableListPage",
        "expense reports": "TrvExpenses",
        "projects": "ProjProjectsListPage",
        "workers": "HcmWorkerListPage",
        "leave requests": "HcmLeaveRequestListPage",
    }


class D365Navigator:
    """
    High-level navigation helpers for D365 Model-Driven Apps.

    Wraps common navigation patterns like opening areas, switching views,
    creating records, and interacting with forms.
    """

    def __init__(self, page: Page, base_url: str):
        self._page = page
        self.base_url = base_url.rstrip("/")

    # ---- Area Navigation ----

    async def open_area(self, area_name: str):
        """
        Open a sitemap area by its display name.
        Examples: 'Project', 'Service', 'Settings'
        """
        logger.info("Opening area: %s", area_name)
        # Click area switcher
        try:
            await self._page.click(D365Selectors.AREA_SWITCHER, timeout=5000)
            await self._page.click(f'text="{area_name}"', timeout=5000)
        except Exception:
            logger.warning("Could not open area via switcher, trying direct nav")

    async def open_entity_list(self, entity_name: str):
        """
        Open an entity list view from the left navigation.
        Examples: 'Projects', 'Time Entries', 'Expenses', 'Resources'
        """
        logger.info("Opening entity list: %s", entity_name)
        # Try the known sitemap data-id selector first
        nav_sel = D365Selectors.NAV_MAP.get(entity_name.lower())
        if nav_sel:
            try:
                await self._page.click(nav_sel, timeout=5000)
                await self._wait_load()
                return
            except Exception:
                logger.warning("Nav map click failed for '%s', trying fallbacks", entity_name)

        try:
            # Try clicking by aria-label
            await self._page.click(
                f'li[aria-label="{entity_name}"]',
                timeout=5000,
            )
        except Exception:
            # Try clicking by data-id partial match
            try:
                await self._page.click(
                    f'li[data-id*="sitemap-entity"] >> text="{entity_name}"',
                    timeout=5000,
                )
            except Exception:
                # Fall back to text match
                await self._page.click(f'text="{entity_name}"', timeout=5000)
        await self._wait_load()

    # ---- Record Operations ----

    async def open_record_by_name(self, name: str):
        """Open a record from a list view by clicking its name."""
        logger.info("Opening record: %s", name)
        await self._page.click(f'a:text-is("{name}"), span:text-is("{name}")', timeout=10000)
        await self._wait_load()

    async def click_new_record(self):
        """Click the New button to create a record."""
        await self._page.click(
            'button[aria-label="New"], button:text-is("New")',
            timeout=5000,
        )
        await self._wait_load()

    async def save_record(self):
        """Save the current record."""
        await self._page.click(D365Selectors.FORM_SAVE, timeout=5000)
        await self._wait_load()

    async def save_and_close(self):
        """Save and close the current record."""
        await self._page.click(D365Selectors.FORM_SAVE_CLOSE, timeout=5000)
        await self._wait_load()

    # ---- Form Interaction ----

    async def set_field(self, field_label: str, value: str):
        """
        Set a form field value by its label text.
        Works with text, number, and lookup fields.
        """
        logger.info("Setting field '%s' = '%s'", field_label, value)
        # Find the field container by label
        field_sel = f'div[data-id="{field_label}"] input, input[aria-label="{field_label}"]'
        try:
            await self._page.fill(field_sel, value, timeout=5000)
        except Exception:
            # Try alternative selectors
            await self._page.fill(
                f'input[aria-label*="{field_label}"]', value, timeout=5000
            )

    async def set_lookup(self, field_label: str, search_text: str):
        """
        Set a lookup field by searching and selecting.
        """
        logger.info("Setting lookup '%s' = '%s'", field_label, search_text)
        field_sel = f'input[aria-label="{field_label}"]'
        await self._page.fill(field_sel, search_text)
        await self._page.keyboard.press("Enter")
        # Wait for lookup results and click first match
        await self._page.wait_for_timeout(2000)
        try:
            await self._page.click(
                f'ul[aria-label*="Lookup results"] li:first-child',
                timeout=5000,
            )
        except Exception:
            logger.warning("Could not auto-select lookup result for '%s'", field_label)

    async def set_option_set(self, field_label: str, option_text: str):
        """Set an option set (dropdown) field by visible text."""
        logger.info("Setting option '%s' = '%s'", field_label, option_text)
        field_sel = f'select[aria-label="{field_label}"]'
        try:
            await self._page.select_option(field_sel, label=option_text, timeout=5000)
        except Exception:
            # Try clicking approach for custom dropdowns
            await self._page.click(f'div[aria-label="{field_label}"]', timeout=5000)
            await self._page.click(f'li:text-is("{option_text}")', timeout=5000)

    async def click_tab(self, tab_name: str):
        """Click a form tab by its name."""
        await self._page.click(
            f'li[aria-label="{tab_name}"], button[aria-label="{tab_name}"]',
            timeout=5000,
        )
        await self._wait_load()

    # ---- Command Bar ----

    async def click_command(self, command_name: str):
        """Click a command bar button by its label."""
        logger.info("Clicking command: %s", command_name)
        await self._page.click(
            f'button[aria-label="{command_name}"], '
            f'button:text-is("{command_name}")',
            timeout=5000,
        )
        await self._wait_load()

    # ---- View Switching ----

    async def switch_view(self, view_name: str):
        """Switch to a different list view."""
        logger.info("Switching to view: %s", view_name)
        await self._page.click(D365Selectors.VIEW_SELECTOR, timeout=5000)
        await self._page.click(f'text="{view_name}"', timeout=5000)
        await self._wait_load()

    # ---- Search / Filter ----

    async def quick_find(self, search_text: str):
        """Use the Quick Find search box in a list view."""
        logger.info("Quick Find: %s", search_text)
        await self._page.fill(D365Selectors.SEARCH_BOX, search_text, timeout=5000)
        await self._page.keyboard.press("Enter")
        await self._wait_load()

    # ---- Helpers ----

    async def _wait_load(self, timeout: int = 10_000):
        """Wait for D365 to finish loading."""
        try:
            await self._page.wait_for_selector(
                D365Selectors.LOADING_SPINNER,
                state="hidden",
                timeout=timeout,
            )
        except Exception:
            pass
        await self._page.wait_for_timeout(500)

    async def get_form_title(self) -> str:
        """Get the current form's header title."""
        try:
            return await self._page.text_content(D365Selectors.FORM_HEADER, timeout=5000) or ""
        except Exception:
            return ""

    async def get_field_value(self, field_label: str) -> str:
        """Get the current value of a form field."""
        try:
            field_sel = f'input[aria-label="{field_label}"]'
            return await self._page.input_value(field_sel, timeout=5000)
        except Exception:
            return ""

    def entity_url(self, entity_logical_name: str, record_id: Optional[str] = None) -> str:
        """Build a URL for a D365 entity page."""
        if record_id:
            return (
                f"{self.base_url}/main.aspx?"
                f"etn={entity_logical_name}&id={record_id}&pagetype=entityrecord"
            )
        return f"{self.base_url}/main.aspx?etn={entity_logical_name}&pagetype=entitylist"
