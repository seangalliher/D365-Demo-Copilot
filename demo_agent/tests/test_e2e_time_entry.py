"""End-to-end Playwright tests for the D365 Demo Copilot.

These tests require:
- A valid ``auth_state.json`` in the project root (created via ``python -m demo_agent.main``)
- D365 environment accessible at the base URL in ``.env``
- Run with: ``pytest demo_agent/tests/test_e2e_time_entry.py -v -s``

The tests are automatically skipped when ``auth_state.json`` is missing.
"""

from __future__ import annotations

import asyncio
from pathlib import Path

import pytest

from demo_agent.agent.executor import DemoExecutor
from demo_agent.agent.state import DemoState, DemoStatus
from demo_agent.browser.controller import BrowserController
from demo_agent.browser.d365_pages import D365Navigator
from demo_agent.browser.overlay_manager import OverlayManager
from demo_agent.config import DemoConfig
from demo_agent.models.demo_plan import (
    ActionType,
    DemoPlan,
    DemoSection,
    DemoStep,
    StepAction,
    ValueHighlight,
)

# Path to auth state — tests are skipped if this doesn't exist
AUTH_STATE = Path(__file__).parent.parent.parent / "auth_state.json"

pytestmark = pytest.mark.skipif(
    not AUTH_STATE.exists(),
    reason="No auth_state.json — run 'python -m demo_agent.main' to create one",
)


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture
def config():
    """Load DemoConfig from .env (uses real credentials)."""
    return DemoConfig()


@pytest.fixture
async def demo_env(config):
    """Set up browser, overlay, navigator, state, and executor for E2E testing.

    Yields a dict with all components. Tears down the browser on exit.
    """
    browser = BrowserController(
        base_url=config.d365_base_url,
        headless=config.headless,
        slow_mo=config.slow_mo,
    )
    page = await browser.start(storage_state=str(AUTH_STATE))

    overlay = OverlayManager(page)
    d365_nav = D365Navigator(page, config.d365_base_url)
    state = DemoState()

    executor = DemoExecutor(
        browser=browser,
        overlay=overlay,
        d365_nav=d365_nav,
        state=state,
    )

    yield {
        "browser": browser,
        "page": page,
        "overlay": overlay,
        "d365_nav": d365_nav,
        "state": state,
        "executor": executor,
        "config": config,
    }

    await browser.stop()


# ---------------------------------------------------------------------------
# Hardcoded Time Entry Demo Plan
# ---------------------------------------------------------------------------

def _time_entry_plan(base_url: str) -> DemoPlan:
    """Build a known-good DemoPlan for time entry (no LLM needed)."""
    return DemoPlan(
        id="e2e_time_entry",
        title="Time Entry Demo",
        subtitle="End-to-end test plan",
        customer_request="Show me how to create a time entry",
        d365_base_url=base_url,
        sections=[
            DemoSection(
                id="sec_nav",
                title="Navigate to Time Entries",
                description="Open the Time Entries list",
                steps=[
                    DemoStep(
                        id="step_nav",
                        title="Open Time Entries",
                        tell_before="Let's navigate to the Time Entries area.",
                        actions=[
                            StepAction(
                                action_type=ActionType.NAVIGATE,
                                value="/main.aspx?pagetype=entitylist&etn=msdyn_timeentry",
                                description="Navigate to Time Entries list",
                                delay_before_ms=0,
                                delay_after_ms=500,
                            ),
                        ],
                        tell_after="We can see the Time Entries list view.",
                    ),
                ],
            ),
            DemoSection(
                id="sec_create",
                title="Create a Time Entry",
                description="Create and save a new time entry",
                steps=[
                    DemoStep(
                        id="step_new",
                        title="Click New",
                        tell_before="Click New to create a time entry.",
                        actions=[
                            StepAction(
                                action_type=ActionType.CLICK,
                                selector='button[data-id="msdyn_timeentry|NoRelationship|HomePageGrid|Mscrm.HomepageGrid.msdyn_timeentry.NewRecord"],'
                                         'button[aria-label="New"],'
                                         'button[data-id="edit-form-new-btn"]',
                                description="Click New button",
                                delay_before_ms=0,
                                delay_after_ms=1000,
                            ),
                        ],
                        tell_after="A new time entry form is now open.",
                    ),
                    DemoStep(
                        id="step_fill_duration",
                        title="Enter duration",
                        tell_before="Enter the number of hours worked.",
                        actions=[
                            StepAction(
                                action_type=ActionType.FILL,
                                selector='input[data-id="msdyn_duration-integer-text-input"],'
                                         'input[aria-label="Duration"]',
                                value="60",
                                description="Enter 60 minutes",
                                delay_before_ms=0,
                                delay_after_ms=500,
                            ),
                        ],
                        tell_after="Duration has been entered.",
                        value_highlight=ValueHighlight(
                            title="Productivity",
                            description="Quick time entry reduces admin overhead",
                            metric_value="2 min",
                            metric_label="average entry time",
                        ),
                    ),
                ],
            ),
        ],
    )


# ---------------------------------------------------------------------------
# E2E Tests
# ---------------------------------------------------------------------------

class TestE2EOverlay:
    """Tests for overlay injection and DOM presence."""

    @pytest.mark.asyncio
    async def test_overlay_injection(self, demo_env):
        """Verify the overlay injects successfully and DOM elements exist."""
        overlay = demo_env["overlay"]
        page = demo_env["page"]

        # Navigate to D365 first
        config = demo_env["config"]
        await demo_env["browser"].navigate(config.d365_base_url)

        # Inject overlay
        await overlay.inject()

        # Check JS API exists
        has_api = await page.evaluate(
            "typeof window.DemoCopilot !== 'undefined'"
        )
        assert has_api, "window.DemoCopilot should be defined after injection"

        # Check DOM root exists
        has_dom = await page.evaluate(
            "!!document.getElementById('demo-copilot-root')"
        )
        assert has_dom, "demo-copilot-root element should exist after injection"

    @pytest.mark.asyncio
    async def test_caption_display(self, demo_env):
        """Verify captions render in the overlay DOM."""
        overlay = demo_env["overlay"]
        page = demo_env["page"]

        config = demo_env["config"]
        await demo_env["browser"].navigate(config.d365_base_url)
        await overlay.inject()

        # Show a caption
        await overlay.show_caption("Test caption text", phase="tell", position="top")
        await asyncio.sleep(0.5)

        # Check that caption bar is visible
        caption_visible = await page.evaluate(
            "document.getElementById('demo-caption-bar')?.classList.contains('visible') ?? false"
        )
        assert caption_visible, "Caption bar should be visible after showCaption"

        # Check caption text content
        caption_text = await page.evaluate(
            "document.getElementById('demo-caption-text')?.textContent ?? ''"
        )
        assert "Test caption text" in caption_text


class TestE2EKeyboard:
    """Tests for Ctrl+Space keyboard shortcut behavior."""

    @pytest.mark.asyncio
    async def test_ctrl_space_advance(self, demo_env):
        """Verify Ctrl+Space fires the advance action via keyboard dispatch."""
        page = demo_env["page"]
        overlay = demo_env["overlay"]

        config = demo_env["config"]
        await demo_env["browser"].navigate(config.d365_base_url)
        await overlay.inject()

        # Set up a listener to capture the action
        await page.evaluate("""
            window.__testActions = [];
            window.__demoCopilotAction = (action) => {
                window.__testActions.push(action);
            };
        """)

        # Press Ctrl+Space (overlay is active after inject → init())
        await page.keyboard.down("Control")
        await page.keyboard.press("Space")
        await page.keyboard.up("Control")
        await asyncio.sleep(0.3)

        # Check that advance was called
        actions = await page.evaluate("window.__testActions")
        assert "advance" in actions, (
            f"Ctrl+Space should trigger 'advance' action, got: {actions}"
        )

    @pytest.mark.asyncio
    async def test_plain_space_does_not_advance(self, demo_env):
        """Verify plain Space does NOT fire advance (only Ctrl+Space does)."""
        page = demo_env["page"]
        overlay = demo_env["overlay"]

        config = demo_env["config"]
        await demo_env["browser"].navigate(config.d365_base_url)
        await overlay.inject()

        await page.evaluate("""
            window.__testActions = [];
            window.__demoCopilotAction = (action) => {
                window.__testActions.push(action);
            };
        """)

        # Press Space without Ctrl
        await page.keyboard.press("Space")
        await asyncio.sleep(0.3)

        actions = await page.evaluate("window.__testActions")
        assert "advance" not in actions, (
            f"Plain Space should NOT trigger advance, got: {actions}"
        )

    @pytest.mark.asyncio
    async def test_pause_resume_cycle(self, demo_env):
        """Verify Escape pauses and Ctrl+Space resumes."""
        page = demo_env["page"]
        overlay = demo_env["overlay"]

        config = demo_env["config"]
        await demo_env["browser"].navigate(config.d365_base_url)
        await overlay.inject()

        # Press Escape to pause
        await page.keyboard.press("Escape")
        await asyncio.sleep(0.3)

        is_paused = await page.evaluate("window.DemoCopilot.isPaused()")
        assert is_paused, "Escape should pause the overlay"

        # Press Ctrl+Space to resume
        await page.keyboard.down("Control")
        await page.keyboard.press("Space")
        await page.keyboard.up("Control")
        await asyncio.sleep(0.3)

        is_paused_after = await page.evaluate("window.DemoCopilot.isPaused()")
        assert not is_paused_after, "Ctrl+Space should resume after pause"


class TestE2ETimeEntryDemo:
    """Full end-to-end demo execution test."""

    @pytest.mark.asyncio
    async def test_time_entry_demo_e2e(self, demo_env):
        """Execute a time entry demo plan and verify all steps complete."""
        state = demo_env["state"]
        executor = demo_env["executor"]
        config = demo_env["config"]

        # Navigate to D365 first to ensure we have a page
        await demo_env["browser"].navigate(config.d365_base_url)

        # Build a hardcoded plan (no LLM dependency)
        plan = _time_entry_plan(config.d365_base_url)

        # Auto-play mode: no waiting for Ctrl+Space between steps
        state.step_mode = False

        # Execute the full demo
        await executor.execute(plan)

        # Verify completion
        assert state.status == DemoStatus.COMPLETED, (
            f"Expected COMPLETED, got {state.status.value}"
        )
        assert state.error_message is None, (
            f"Demo should not have errors: {state.error_message}"
        )
        assert len(state.history) == plan.total_steps, (
            f"Expected {plan.total_steps} history entries, got {len(state.history)}"
        )

        # Verify no steps were skipped
        skipped = [h for h in state.history if h.skipped]
        assert len(skipped) == 0, (
            f"No steps should be skipped, but {len(skipped)} were: "
            f"{[h.step_id for h in skipped]}"
        )
