"""
D365 Demo Copilot — Browser Controller

Manages a Playwright browser instance for navigating Dynamics 365.
Provides high-level methods for common D365 interactions like
navigating to areas, clicking controls, filling fields, and
waiting for page loads.
"""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path
from typing import Optional

from playwright.async_api import (
    Browser,
    BrowserContext,
    Page,
    Playwright,
    async_playwright,
)

logger = logging.getLogger("demo_agent.browser")


class BrowserController:
    """
    Controls a Chromium browser instance via Playwright.

    Provides methods for:
    - Launching and connecting to a browser
    - Navigating D365 pages
    - Interacting with elements (click, fill, select)
    - Taking screenshots
    - Waiting for D365-specific load states
    """

    def __init__(self, base_url: str, headless: bool = False, slow_mo: int = 50):
        self.base_url = base_url.rstrip("/")
        self.headless = headless
        self.slow_mo = slow_mo
        self._playwright: Optional[Playwright] = None
        self._browser: Optional[Browser] = None
        self._context: Optional[BrowserContext] = None
        self._page: Optional[Page] = None

    @property
    def page(self) -> Page:
        """The active browser page."""
        if not self._page:
            raise RuntimeError("Browser not started. Call start() first.")
        return self._page

    async def start(self, storage_state: Optional[str] = None) -> Page:
        """
        Launch the browser and navigate to D365.

        Args:
            storage_state: Path to saved auth state (cookies/localStorage).
                          Use save_auth_state() to create this after manual login.

        Returns:
            The active Page instance.
        """
        logger.info("Starting browser (headless=%s, slow_mo=%d)", self.headless, self.slow_mo)

        self._playwright = await async_playwright().start()
        self._browser = await self._playwright.chromium.launch(
            headless=self.headless,
            slow_mo=self.slow_mo,
            args=[
                "--start-maximized",
                "--disable-blink-features=AutomationControlled",
                "--remote-debugging-port=9222",
            ],
        )

        context_options = {
            "no_viewport": True,  # Use full window size — adapts to any screen
            "ignore_https_errors": True,
        }

        if storage_state and Path(storage_state).exists():
            context_options["storage_state"] = storage_state
            logger.info("Loaded auth state from %s", storage_state)

        self._context = await self._browser.new_context(**context_options)
        self._page = await self._context.new_page()

        # Set default timeouts appropriate for D365 (pages can be slow)
        self._page.set_default_timeout(30_000)
        self._page.set_default_navigation_timeout(60_000)

        return self._page

    async def stop(self):
        """Close the browser and clean up."""
        if self._browser:
            await self._browser.close()
        if self._playwright:
            await self._playwright.stop()
        self._page = None
        self._context = None
        self._browser = None
        self._playwright = None
        logger.info("Browser stopped")

    async def save_auth_state(self, path: str = "auth_state.json"):
        """
        Save the current authentication state (cookies, localStorage)
        so future sessions can skip login.
        """
        if self._context:
            await self._context.storage_state(path=path)
            logger.info("Auth state saved to %s", path)

    # ---- Navigation ----

    async def navigate(self, url: str):
        """Navigate to a URL (relative or absolute)."""
        if url.startswith("/"):
            url = f"{self.base_url}{url}"
        elif not url.startswith("http"):
            url = f"{self.base_url}/{url}"

        logger.info("Navigating to %s", url)
        await self.page.goto(url, wait_until="domcontentloaded")
        await self._wait_for_d365_load()

    async def navigate_to_area(self, area_name: str):
        """
        Navigate to a D365 sitemap area by name.
        Examples: 'Projects', 'Time Entries', 'Expenses', 'Resources'
        """
        logger.info("Navigating to area: %s", area_name)
        # Click the sitemap/area switcher
        try:
            # Try the main nav area selector
            await self.page.click(f'button[aria-label="{area_name}"]', timeout=5000)
        except Exception:
            # Try clicking the area in the left nav
            try:
                await self.page.click(
                    f'li[data-text="{area_name}"], '
                    f'span:text-is("{area_name}")',
                    timeout=5000,
                )
            except Exception:
                # Fall back to searching the page
                await self.page.click(f'text="{area_name}"', timeout=5000)

        await self._wait_for_d365_load()

    async def navigate_to_record(self, entity: str, record_id: str):
        """Navigate directly to a specific record."""
        url = f"{self.base_url}/main.aspx?etn={entity}&id={record_id}&pagetype=entityrecord"
        await self.navigate(url)

    async def navigate_to_list(self, entity: str):
        """Navigate to an entity list view."""
        url = f"{self.base_url}/main.aspx?etn={entity}&pagetype=entitylist"
        await self.navigate(url)

    # ---- Element Interaction ----

    async def click(self, selector: str, force: bool = False):
        """Click an element, with D365-aware waiting."""
        logger.info("Clicking: %s", selector)
        await self.page.click(selector, force=force)
        await asyncio.sleep(0.3)  # Brief pause for D365 reactivity

    async def fill(self, selector: str, value: str, clear_first: bool = True):
        """Fill a text field with a value."""
        logger.info("Filling '%s' into %s", value, selector)
        if clear_first:
            await self.page.fill(selector, "")
        await self.page.fill(selector, value)

    async def type_slowly(self, selector: str, value: str, delay: int = 80):
        """
        Type text character-by-character for visual effect during demos.
        This creates a natural typing appearance.
        """
        logger.info("Typing slowly '%s' into %s", value, selector)
        await self.page.click(selector)
        await self.page.fill(selector, "")
        await self.page.type(selector, value, delay=delay)

    async def select_option(self, selector: str, value: str):
        """Select an option from a dropdown."""
        logger.info("Selecting '%s' in %s", value, selector)
        await self.page.select_option(selector, value)

    async def hover(self, selector: str):
        """Hover over an element."""
        await self.page.hover(selector)

    # ---- Waiting ----

    async def wait_for(self, selector: str, timeout: int = 30_000):
        """Wait for an element to be visible."""
        await self.page.wait_for_selector(selector, state="visible", timeout=timeout)

    async def wait_for_text(self, text: str, timeout: int = 30_000):
        """Wait for specific text to appear on the page."""
        await self.page.wait_for_selector(f'text="{text}"', state="visible", timeout=timeout)

    async def _wait_for_d365_load(self, timeout: int = 15_000):
        """
        Wait for D365 Model-Driven App to finish loading.
        Checks for common loading indicators to disappear.
        """
        try:
            # Wait for the main loading spinner to disappear
            await self.page.wait_for_selector(
                '#progressContainer, .ms-Spinner, [data-id="loading"]',
                state="hidden",
                timeout=timeout,
            )
        except Exception:
            pass  # Element may not exist — that's fine

        # Wait for document ready state
        try:
            await self.page.wait_for_function(
                "document.readyState === 'complete'",
                timeout=10_000,
            )
        except Exception:
            pass

        # Give D365 a moment to settle JS execution
        await asyncio.sleep(1.0)

    # ---- Information ----

    async def get_page_title(self) -> str:
        """Get the current page title."""
        return await self.page.title()

    async def get_element_text(self, selector: str) -> str:
        """Get the text content of an element."""
        return await self.page.text_content(selector) or ""

    async def element_exists(self, selector: str, timeout: int = 3000) -> bool:
        """Check if an element exists on the page."""
        try:
            await self.page.wait_for_selector(selector, timeout=timeout)
            return True
        except Exception:
            return False

    async def get_element_rect(self, selector: str) -> Optional[dict]:
        """Get the bounding rectangle of an element."""
        try:
            box = await self.page.locator(selector).bounding_box()
            return box
        except Exception:
            return None

    async def screenshot(self, path: str = "demo_screenshot.png", full_page: bool = False):
        """Take a screenshot of the current page."""
        await self.page.screenshot(path=path, full_page=full_page)
        logger.info("Screenshot saved to %s", path)

    # ---- JavaScript Execution ----

    async def evaluate(self, expression: str):
        """Execute JavaScript in the page context."""
        return await self.page.evaluate(expression)

    async def evaluate_handle(self, expression: str):
        """Execute JavaScript and return a handle to the result."""
        return await self.page.evaluate_handle(expression)
