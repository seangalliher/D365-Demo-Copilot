"""
D365 Demo Copilot — Script Recorder

Captures per-step screenshots and narration metadata during demo
execution. Used by ScriptGenerator to produce a PDF demo script
after the demo completes.

Screenshots are clipped to the D365 content area, excluding the
sidecar chat panel and demo overlays (captions, spotlights, etc.).
"""

from __future__ import annotations

import asyncio
import logging
import time
from dataclasses import dataclass, field
from typing import Optional

from playwright.async_api import Page

from ..browser.overlay_manager import OverlayManager
from ..models.demo_plan import DemoSection, DemoStep

logger = logging.getLogger("demo_agent.script_recorder")


@dataclass
class StepCapture:
    """Data captured for one step during demo execution."""

    step_id: str
    step_title: str
    section_id: str
    section_title: str
    section_description: str
    tell_before: str
    tell_after: str
    screenshot_png: bytes  # PNG bytes (clipped to D365 content, no sidecar)
    value_highlight: Optional[dict] = None  # {title, description, metric_value, metric_label}
    timestamp: float = 0.0
    step_number: int = 0  # 1-based global step number
    skipped: bool = False


class ScriptRecorder:
    """
    Captures per-step screenshots and narration metadata during demo
    execution.

    Screenshots are clipped to the D365 content area, excluding the
    sidecar chat panel and demo overlays.
    """

    def __init__(self, page: Page, overlay: OverlayManager):
        self._page = page
        self._overlay = overlay
        self._captures: list[StepCapture] = []

    @property
    def captures(self) -> list[StepCapture]:
        """Return all step captures."""
        return list(self._captures)

    async def capture_step(
        self,
        step: DemoStep,
        section: DemoSection,
        global_step_index: int,
        skipped: bool = False,
    ) -> None:
        """Capture screenshot and metadata after the SHOW phase.

        For skipped steps, only metadata is recorded (empty screenshot).
        For normal steps:
          1. Clear demo overlays (captions, spotlights, status, tooltips)
          2. Wait briefly for CSS transitions to complete
          3. Take clipped screenshot excluding sidecar panel
        """
        screenshot_png = b""

        if not skipped:
            try:
                # Clear overlays so screenshot shows clean D365 state
                await self._overlay.clear_all()
                await asyncio.sleep(0.3)  # Wait for CSS transitions

                screenshot_png = await self._take_content_screenshot()
                logger.info(
                    "[SCRIPT] Captured step %d: %s (%.1f KB)",
                    global_step_index + 1,
                    step.title,
                    len(screenshot_png) / 1024,
                )
            except Exception as e:
                logger.warning(
                    "[SCRIPT] Screenshot failed for step %s: %s", step.id, e
                )

        vh_dict = None
        if step.value_highlight:
            vh_dict = {
                "title": step.value_highlight.title,
                "description": step.value_highlight.description,
                "metric_value": step.value_highlight.metric_value,
                "metric_label": step.value_highlight.metric_label,
            }

        capture = StepCapture(
            step_id=step.id,
            step_title=step.title,
            section_id=section.id,
            section_title=section.title,
            section_description=section.description,
            tell_before=step.tell_before,
            tell_after=step.tell_after,
            screenshot_png=screenshot_png,
            value_highlight=vh_dict,
            timestamp=time.time(),
            step_number=global_step_index + 1,
            skipped=skipped,
        )
        self._captures.append(capture)

    async def _take_content_screenshot(self) -> bytes:
        """Take a PNG screenshot of just the D365 content area.

        Uses Playwright's clip parameter to capture only the content
        left of the sidecar panel.
        """
        dims = await self._page.evaluate(
            """() => ({
                viewWidth: window.innerWidth,
                viewHeight: window.innerHeight,
                sidecarWidth: (() => {
                    const host = document.getElementById('demo-chat-panel-host');
                    if (!host || host.classList.contains('collapsed')) return 40;
                    return host.offsetWidth;
                })()
            })"""
        )

        content_width = dims["viewWidth"] - dims["sidecarWidth"]
        clip = {
            "x": 0,
            "y": 0,
            "width": max(content_width, 100),
            "height": dims["viewHeight"],
        }
        return await self._page.screenshot(clip=clip)

    def reset(self) -> None:
        """Clear all captures for a new demo run."""
        self._captures.clear()
