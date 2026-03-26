"""
D365 Demo Copilot — Overlay Manager

Injects and controls the visual overlay system in the browser page.
Acts as a bridge between the Python agent and the JavaScript DemoCopilot
API running in the browser context.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Optional

from playwright.async_api import Page

logger = logging.getLogger("demo_agent.overlay")

# Paths to overlay assets
OVERLAY_DIR = Path(__file__).parent.parent / "overlay"
CSS_PATH = OVERLAY_DIR / "demo-overlay.css"
JS_PATH = OVERLAY_DIR / "demo-overlay.js"


class OverlayManager:
    """
    Manages the visual overlay system injected into the D365 page.

    Provides Python-friendly wrappers around the browser-side
    window.DemoCopilot JavaScript API for:
    - Spotlight (highlight elements)
    - Captions (subtitle-style text)
    - Business Value Cards
    - Progress indicators
    - Click ripple effects
    - Tooltips
    - Pause/resume overlay
    - Title slides
    """

    def __init__(self, page: Page):
        self._page = page
        self._injected = False
        self._init_script_added = False

    async def inject(self, force: bool = False):
        """
        Inject the overlay CSS and JavaScript into the current page.
        Verifies both the JS API and DOM elements exist.

        Args:
            force: If True, always re-inject even if overlay appears present.
        """
        if not force and self._injected:
            # Check if overlay JS AND DOM both still exist
            both_exist = await self._page.evaluate(
                "typeof window.DemoCopilot !== 'undefined' && !!document.getElementById('demo-copilot-root')"
            )
            if both_exist:
                return

        logger.info("Injecting demo overlay into page (force=%s)", force)

        # Read assets
        css_content = CSS_PATH.read_text(encoding="utf-8")
        js_content = JS_PATH.read_text(encoding="utf-8")

        # Clean up any stale remnants first
        await self._page.evaluate(
            """() => {
                const oldStyle = document.getElementById('demo-copilot-styles');
                if (oldStyle) oldStyle.remove();
                const oldRoot = document.getElementById('demo-copilot-root');
                if (oldRoot) oldRoot.remove();
                delete window.DemoCopilot;
            }"""
        )

        # Inject CSS
        await self._page.evaluate(
            """(css) => {
                const style = document.createElement('style');
                style.id = 'demo-copilot-styles';
                style.textContent = css;
                document.head.appendChild(style);
            }""",
            css_content,
        )

        # Inject JS (includes IIFE that calls init() → buildDOM())
        await self._page.evaluate(js_content)

        # Verify injection succeeded
        ok = await self._page.evaluate(
            "typeof window.DemoCopilot !== 'undefined' && !!document.getElementById('demo-copilot-root')"
        )
        if not ok:
            logger.warning("Overlay injection verification failed — DOM may not be ready")
        else:
            logger.info("Overlay injected and verified successfully")

        self._injected = True

        # Register init script so overlay auto-injects on future navigations
        if not self._init_script_added:
            combined_script = (
                f"(function() {{\n"
                f"  // Auto-inject overlay CSS\n"
                f"  if (!document.getElementById('demo-copilot-styles')) {{\n"
                f"    var s = document.createElement('style');\n"
                f"    s.id = 'demo-copilot-styles';\n"
                f"    s.textContent = {repr(css_content)};\n"
                f"    document.head.appendChild(s);\n"
                f"  }}\n"
                f"}})();\n"
                + js_content
            )
            await self._page.add_init_script(script=combined_script)
            self._init_script_added = True
            logger.info("Registered overlay init script for future navigations")

    async def ensure_injected(self):
        """Ensure overlay is injected AND DOM elements exist. Re-inject if not."""
        try:
            both_exist = await self._page.evaluate(
                "typeof window.DemoCopilot !== 'undefined' && !!document.getElementById('demo-copilot-root')"
            )
            if not both_exist:
                logger.info("Overlay missing (JS=%s, DOM=%s) — re-injecting",
                    await self._page.evaluate("typeof window.DemoCopilot !== 'undefined'"),
                    await self._page.evaluate("!!document.getElementById('demo-copilot-root')"),
                )
                self._injected = False
                await self.inject(force=True)
        except Exception as e:
            logger.warning("ensure_injected check failed (%s) — re-injecting", e)
            self._injected = False
            await self.inject(force=True)

    # ---- Spotlight ----

    async def spotlight_on(self, selector: str, padding: int = 12):
        """
        Highlight an element with the spotlight effect.
        Dims the rest of the page and draws a glowing ring around the target.
        """
        await self.ensure_injected()
        logger.info("Spotlight ON: %s", selector)
        await self._page.evaluate(
            "(args) => window.DemoCopilot.spotlightOn(args.sel, args.pad)",
            {"sel": selector, "pad": padding},
        )

    async def spotlight_off(self):
        """Remove the spotlight effect."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.spotlightOff()")

    # ---- Captions ----

    async def show_caption(self, text: str, phase: str = "tell", position: str = "auto"):
        """
        Show a caption on screen (instant).

        Args:
            text: Caption text. Supports <span class="highlight"> for emphasis.
            phase: One of 'tell', 'show', or 'value'.
            position: 'auto' (based on spotlight target), 'top', or 'bottom'.
        """
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.showCaption(args.text, args.phase, args.position)",
            {"text": text, "phase": phase, "position": position},
        )

    async def show_caption_animated(self, text: str, phase: str = "tell", speed: int = 25, position: str = "auto"):
        """
        Show a caption with typewriter animation.

        Args:
            text: Caption text.
            phase: 'tell', 'show', or 'value'.
            speed: Milliseconds per character.
            position: 'auto' (based on spotlight target), 'top', or 'bottom'.
        """
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.showCaptionAnimated(args.text, args.phase, args.speed, args.position)",
            {"text": text, "phase": phase, "speed": speed, "position": position},
        )

    async def hide_caption(self):
        """Hide the caption bar."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideCaption()")

    # ---- Business Value Card ----

    async def show_value_card(
        self,
        title: str,
        text: str,
        metric_value: Optional[str] = None,
        metric_label: Optional[str] = None,
        position: str = "top-right",
    ):
        """
        Show a business value callout card.

        Args:
            title: Card title, e.g. "Time Savings"
            text: Value description
            metric_value: Optional metric, e.g. "40%"
            metric_label: Optional metric label, e.g. "reduction in manual entry"
            position: Card position on screen
        """
        await self.ensure_injected()
        metric = None
        if metric_value and metric_label:
            metric = {"value": metric_value, "label": metric_label}

        await self._page.evaluate(
            "(args) => window.DemoCopilot.showValueCard(args.title, args.text, args.metric, args.position)",
            {"title": title, "text": text, "metric": metric, "position": position},
        )

    async def hide_value_card(self):
        """Hide the business value card."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideValueCard()")

    # ---- Progress ----

    async def init_progress(self, total_steps: int):
        """Initialize the progress bar and step indicator."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(n) => window.DemoCopilot.initProgress(n)", total_steps
        )

    async def update_progress(self, step_index: int):
        """Update the progress to the given step index (0-based)."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(i) => window.DemoCopilot.updateProgress(i)", step_index
        )

    async def hide_progress(self):
        """Hide the progress indicators."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideProgress()")

    # ---- Click Ripple ----

    async def click_ripple(self, x: float, y: float):
        """Show a ripple effect at specific coordinates."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.clickRipple(args.x, args.y)",
            {"x": x, "y": y},
        )

    async def click_ripple_on(self, selector: str):
        """Show a ripple effect centered on an element."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(sel) => window.DemoCopilot.clickRippleOnElement(sel)", selector
        )

    # ---- Tooltip ----

    async def show_tooltip(self, selector: str, text: str, position: str = "above"):
        """Show a tooltip annotation on an element."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.showTooltip(args.sel, args.text, args.pos)",
            {"sel": selector, "text": text, "pos": position},
        )

    async def hide_tooltip(self):
        """Hide the active tooltip."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideTooltip()")

    # ---- Pause / Resume ----

    async def pause(self):
        """Show the pause overlay."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.pause()")

    async def resume(self):
        """Hide the pause overlay and resume."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.resume()")

    async def is_paused(self) -> bool:
        """Check if the demo is currently paused."""
        await self.ensure_injected()
        return await self._page.evaluate("window.DemoCopilot.isPaused()")

    # ---- Title Slide ----

    async def show_title_slide(
        self, heading: str, subheading: str = "", meta: str = ""
    ):
        """Show a full-screen title slide."""
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.showTitleSlide(args.heading, args.sub, args.meta)",
            {"heading": heading, "sub": subheading, "meta": meta},
        )

    async def hide_title_slide(self):
        """Hide the title slide."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideTitleSlide()")

    # ---- Agent Status Indicator ----

    async def show_status(self, text: str, mode: str = "working"):
        """Show the agent status pill.

        Args:
            text: Status message (e.g. 'Working...', 'Loading page...').
            mode: Visual style — 'working', 'loading', 'waiting', 'narrating'.
        """
        await self.ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilot.showStatus(args.text, args.mode)",
            {"text": text, "mode": mode},
        )

    async def hide_status(self):
        """Hide the agent status pill."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.hideStatus()")

    # ---- Lifecycle ----

    async def clear_all(self):
        """Clear all overlays (spotlight, caption, value card, tooltip)."""
        await self.ensure_injected()
        await self._page.evaluate("window.DemoCopilot.clearAll()")

    async def destroy(self):
        """Remove all overlay DOM elements from the page."""
        try:
            await self._page.evaluate("window.DemoCopilot.destroy()")
        except Exception:
            pass
        self._injected = False
