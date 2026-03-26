"""
D365 Demo Copilot — Chat Panel Manager

Manages the sidecar chat panel injected into the browser page.
Bridges browser-side chat UI with the Python agent via
Playwright's page.expose_function().

Architecture:
  Browser (chat-panel.js)  ←→  ChatPanelManager (Python)
    - User types message   →  __demoCopilotSend(text)  →  on_user_message callback
    - Quick action clicked  →  __demoCopilotAction(act) →  on_action callback
    - Python calls         →  page.evaluate(...)        →  window.DemoCopilotChat.*
"""

from __future__ import annotations

import asyncio
import logging
from pathlib import Path
from typing import Callable, Optional, Awaitable

from playwright.async_api import Page

logger = logging.getLogger("demo_agent.chat_panel")

# Paths to chat panel assets
OVERLAY_DIR = Path(__file__).parent.parent / "overlay"
CHAT_CSS_PATH = OVERLAY_DIR / "chat-panel.css"
CHAT_JS_PATH = OVERLAY_DIR / "chat-panel.js"

# Chat panel width in pixels (must match CSS)
CHAT_PANEL_WIDTH = 400


class ChatPanelManager:
    """
    Manages the sidecar chat panel in the browser.

    Provides Python-friendly wrappers around window.DemoCopilotChat
    and handles the browser → Python communication bridge via
    Playwright's page.expose_function().
    """

    def __init__(
        self,
        page: Page,
        on_user_message: Optional[Callable[[str], Awaitable[None]]] = None,
        on_action: Optional[Callable[[str], Awaitable[None]]] = None,
    ):
        self._page = page
        self._injected = False
        self._functions_exposed = False
        self._on_user_message = on_user_message
        self._on_action = on_action
        self._message_queue: asyncio.Queue[str] = asyncio.Queue()
        self._action_queue: asyncio.Queue[str] = asyncio.Queue()
        self._voice_enabled: bool = False  # Tracked so state survives re-injection
        # Demo state tracked for re-injection persistence
        self._ui_mode: str = "welcome"
        self._demo_running: bool = False
        self._tracker_plan: dict | None = None
        self._tracker_step_states: dict[int, str] = {}  # idx -> 'active'/'done'
        self._completion_message: str | None = None
        self._post_demo_actions: list[dict] = []

    async def inject(self, force: bool = False):
        """
        Inject the chat panel inside a Shadow DOM root for full CSS isolation.
        Also exposes the bridge functions for browser → Python communication.

        Architecture:
          - Host element (#demo-chat-panel-host) in light DOM
          - Shadow root contains all chat HTML + CSS (D365 CSS can't reach it)
          - Toggle button in light DOM with inline styles
        """
        if not force and self._injected:
            exists = await self._page.evaluate(
                "typeof window.DemoCopilotChat !== 'undefined' && !!document.getElementById('demo-chat-panel-host')"
            )
            if exists:
                return

        logger.info("Injecting chat panel with Shadow DOM isolation (force=%s)", force)

        css_content = CHAT_CSS_PATH.read_text(encoding="utf-8")
        js_content = CHAT_JS_PATH.read_text(encoding="utf-8")

        # Clean up stale remnants (both shadow and legacy non-shadow versions)
        await self._page.evaluate(
            """() => {
                const oldHost = document.getElementById('demo-chat-panel-host');
                if (oldHost) oldHost.remove();
                const oldPanel = document.getElementById('demo-chat-panel');
                if (oldPanel) oldPanel.remove();
                const oldToggle = document.getElementById('demo-chat-toggle');
                if (oldToggle) oldToggle.remove();
                const oldStyles = document.getElementById('demo-chat-panel-styles');
                if (oldStyles) oldStyles.remove();
                delete window.DemoCopilotChat;
            }"""
        )

        # Set CSS as a global so the JS IIFE can pick it up and inject into shadow root
        await self._page.evaluate(
            "(css) => { window.__demoCopilotCSS = css; }",
            css_content,
        )

        # Inject JS — it creates the shadow DOM host, attaches shadow root,
        # injects CSS + HTML inside, and exposes window.DemoCopilotChat API
        await self._page.evaluate(js_content)

        # Expose bridge functions (only once per page context)
        if not self._functions_exposed:
            await self._page.expose_function(
                "__demoCopilotSend", self._handle_user_message
            )
            await self._page.expose_function(
                "__demoCopilotAction", self._handle_action
            )
            self._functions_exposed = True
            logger.info("Exposed bridge functions: __demoCopilotSend, __demoCopilotAction")

        # Verify — check for host element in light DOM and API on window
        ok = await self._page.evaluate(
            "typeof window.DemoCopilotChat !== 'undefined' && !!document.getElementById('demo-chat-panel-host')"
        )
        if ok:
            logger.info("Chat panel injected and verified (Shadow DOM)")
        else:
            logger.warning("Chat panel injection verification failed")

        self._injected = True

    async def _handle_user_message(self, text: str):
        """Called when user sends a message from the browser chat."""
        logger.info("User message from chat: %s", text[:80])
        await self._message_queue.put(text)
        if self._on_user_message:
            await self._on_user_message(text)

    async def _handle_action(self, action: str):
        """Called when user clicks a quick action or plan button."""
        logger.info("Action from chat: %s", action)
        await self._action_queue.put(action)
        if self._on_action:
            await self._on_action(action)

    async def wait_for_message(self, timeout: Optional[float] = None) -> str:
        """
        Wait for the next user message from the chat panel.

        Args:
            timeout: Maximum seconds to wait. None = wait forever.

        Returns:
            The user's message text.

        Raises:
            asyncio.TimeoutError: If timeout expires.
        """
        if timeout:
            return await asyncio.wait_for(self._message_queue.get(), timeout)
        return await self._message_queue.get()

    async def wait_for_action(self, timeout: Optional[float] = None) -> str:
        """
        Wait for the next action from the chat panel.

        Returns:
            The action string (e.g., 'start_demo', 'pause', 'quit').
        """
        if timeout:
            return await asyncio.wait_for(self._action_queue.get(), timeout)
        return await self._action_queue.get()

    async def wait_for_message_or_action(
        self, timeout: Optional[float] = None
    ) -> tuple[str, str]:
        """
        Wait for either a user message or an action.

        Returns:
            Tuple of (type, value) where type is 'message' or 'action'.
        """
        msg_task = asyncio.create_task(self._message_queue.get())
        act_task = asyncio.create_task(self._action_queue.get())

        tasks = [msg_task, act_task]
        try:
            done, pending = await asyncio.wait(
                tasks,
                timeout=timeout,
                return_when=asyncio.FIRST_COMPLETED,
            )

            for p in pending:
                p.cancel()

            if not done:
                raise asyncio.TimeoutError()

            result = done.pop()
            if result is msg_task:
                return ("message", result.result())
            else:
                return ("action", result.result())
        except Exception:
            for t in tasks:
                t.cancel()
            raise

    # ---- Chat UI Methods ----

    async def add_message(self, role: str, text: str):
        """
        Add a message to the chat.

        Args:
            role: 'user', 'assistant', or 'system'
            text: Message text
        """
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.addMessage(args.role, args.text)",
            {"role": role, "text": text},
        )

    async def add_message_html(self, role: str, html: str):
        """Add a message with raw HTML content."""
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.addMessageHtml(args.role, args.html)",
            {"role": role, "html": html},
        )

    async def show_typing(self):
        """Show the typing indicator."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.showTyping()")

    async def hide_typing(self):
        """Hide the typing indicator."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.hideTyping()")

    async def show_plan(self, plan_dict: dict):
        """
        Show a plan summary card in the chat.

        Args:
            plan_dict: Plan data with 'title' and 'sections' containing 'steps'.
        """
        await self._ensure_injected()
        await self._page.evaluate(
            "(plan) => window.DemoCopilotChat.showPlan(plan)", plan_dict
        )

    async def update_plan_step(self, step_index: int, status: str):
        """
        Update a step's status in the plan card.

        Args:
            step_index: 0-based step index
            status: 'active', 'done', or ''
        """
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.updatePlanStep(args.idx, args.status)",
            {"idx": step_index, "status": status},
        )

    async def show_progress(self, step: int, total: int, label: str = ""):
        """Show a progress indicator in the chat."""
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.showProgress(args.step, args.total, args.label)",
            {"step": step, "total": total, "label": label},
        )

    async def update_progress(self, step: int, total: int, label: str = ""):
        """Update the existing progress indicator."""
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.updateProgress(args.step, args.total, args.label)",
            {"step": step, "total": total, "label": label},
        )

    async def hide_progress(self):
        """Hide the progress indicator."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.hideProgress()")

    async def set_status(self, status: str, status_type: str = ""):
        """
        Set the header status text.

        Args:
            status: Status text (e.g., 'Ready', 'Planning...', 'Running demo')
            status_type: CSS class — '', 'busy', or 'error'
        """
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.setStatus(args.status, args.type)",
            {"status": status, "type": status_type},
        )

    async def disable_input(self):
        """Disable the chat input."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.disable()")

    async def enable_input(self):
        """Enable the chat input."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.enable()")

    async def show_quick_actions(self, actions: list[dict]):
        """
        Show quick action buttons.

        Args:
            actions: List of dicts with 'label', 'action', and optional 'danger' keys.
        """
        await self._ensure_injected()
        await self._page.evaluate(
            "(actions) => window.DemoCopilotChat.showQuickActions(actions)", actions
        )

    async def hide_quick_actions(self):
        """Hide quick action buttons."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.hideQuickActions()")

    async def clear(self):
        """Clear all chat messages."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.clear()")

    async def show_welcome(self):
        """Show the welcome screen."""
        self._ui_mode = "welcome"
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.showWelcome()")

    async def hide_welcome(self):
        """Hide the welcome screen."""
        if self._ui_mode == "welcome":
            self._ui_mode = "messages"
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.hideWelcome()")

    async def collapse(self):
        """Collapse the chat panel."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.collapse()")

    async def expand(self):
        """Expand the chat panel."""
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.expand()")

    async def trigger_download(self, filename: str, b64_data: str, mime_type: str):
        """Trigger a file download in the browser via Blob URL.

        Sends base64-encoded data to the browser, converts to a Blob,
        creates a temporary <a> element with download attribute, and clicks it.
        """
        await self._page.evaluate(
            """(args) => {
                const byteChars = atob(args.b64);
                const byteArray = new Uint8Array(byteChars.length);
                for (let i = 0; i < byteChars.length; i++) {
                    byteArray[i] = byteChars.charCodeAt(i);
                }
                const blob = new Blob([byteArray], { type: args.mime });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = args.filename;
                document.body.appendChild(a);
                a.click();
                setTimeout(() => {
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                }, 5000);
            }""",
            {"b64": b64_data, "filename": filename, "mime": mime_type},
        )

    async def set_voice_enabled(self, enabled: bool):
        """Update the voice toggle button state in the sidecar UI."""
        self._voice_enabled = enabled  # Remember for re-injection
        await self._ensure_injected()
        await self._page.evaluate(
            "(enabled) => window.DemoCopilotChat.setVoiceEnabled(enabled)",
            enabled,
        )

    async def show_step_tracker(self, plan_dict: dict):
        """Show the persistent step tracker panel with the demo plan steps."""
        self._ui_mode = "demo_running"
        self._demo_running = True
        self._tracker_plan = plan_dict
        self._tracker_step_states = {}
        self._completion_message = None
        self._post_demo_actions = []
        await self._ensure_injected()
        await self._page.evaluate(
            "(plan) => window.DemoCopilotChat.showStepTracker(plan)", plan_dict
        )

    async def update_tracker_step(self, step_index: int, status: str):
        """Update a step's status in the persistent step tracker.

        Args:
            step_index: 0-based step index
            status: 'active', 'done', or ''
        """
        self._tracker_step_states[step_index] = status
        await self._ensure_injected()
        await self._page.evaluate(
            "(args) => window.DemoCopilotChat.updateTrackerStep(args.idx, args.status)",
            {"idx": step_index, "status": status},
        )

    async def hide_step_tracker(self):
        """Hide the persistent step tracker panel."""
        self._ui_mode = "messages"
        self._demo_running = False
        self._tracker_plan = None
        self._tracker_step_states = {}
        self._completion_message = None
        self._post_demo_actions = []
        await self._ensure_injected()
        await self._page.evaluate("window.DemoCopilotChat.hideStepTracker()")

    async def set_demo_completed(
        self,
        completion_message: str,
        post_demo_actions: Optional[list[dict]] = None,
    ):
        """Persist the post-demo UI so it can be restored after re-injection."""
        self._ui_mode = "demo_complete"
        self._demo_running = False
        self._completion_message = completion_message
        self._post_demo_actions = list(post_demo_actions or [])

    async def reset_demo_ui_for_new_request(self):
        """Clear persisted demo UI so a new request starts from a clean state."""
        self._ui_mode = "messages"
        self._demo_running = False
        self._tracker_plan = None
        self._tracker_step_states = {}
        self._completion_message = None
        self._post_demo_actions = []
        await self._ensure_injected()
        await self._page.evaluate(
            """() => {
                window.DemoCopilotChat.hideStepTracker();
                window.DemoCopilotChat.hideQuickActions();
                window.DemoCopilotChat.hideWelcome();
            }"""
        )

    # ---- Internal ----

    async def _ensure_injected(self):
        """Ensure chat panel is injected, re-inject if needed."""
        try:
            exists = await self._page.evaluate(
                "typeof window.DemoCopilotChat !== 'undefined' && !!document.getElementById('demo-chat-panel-host')"
            )
            if not exists:
                logger.info("Chat panel missing — re-injecting")
                self._injected = False
                await self.inject(force=True)
                await self._restore_ui_state()
        except Exception as e:
            logger.warning("Chat panel check failed (%s) — re-injecting", e)
            self._injected = False
            await self.inject(force=True)
            await self._restore_ui_state()

    async def _restore_ui_state(self):
        """Restore persisted UI state after a re-injection."""
        try:
            await self._page.evaluate(
                "(enabled) => window.DemoCopilotChat.setVoiceEnabled(enabled)",
                self._voice_enabled,
            )
        except Exception:
            pass

        if self._ui_mode == "welcome":
            return

        # Restore demo-running or completed-demo state: hide welcome and rebuild tracker.
        if self._ui_mode in {"demo_running", "demo_complete"} and self._tracker_plan:
            try:
                await self._page.evaluate(
                    "window.DemoCopilotChat.hideWelcome()"
                )
                await self._page.evaluate(
                    "(plan) => window.DemoCopilotChat.showStepTracker(plan)",
                    self._tracker_plan,
                )
                # Restore each step's status (done/active markers)
                for idx, status in self._tracker_step_states.items():
                    await self._page.evaluate(
                        "(args) => window.DemoCopilotChat.updateTrackerStep(args.idx, args.status)",
                        {"idx": idx, "status": status},
                    )
                if self._ui_mode == "demo_complete":
                    await self._page.evaluate(
                        "(message) => window.DemoCopilotChat.addMessage('assistant', message)",
                        self._completion_message or "",
                    )
                    if self._post_demo_actions:
                        await self._page.evaluate(
                            "(actions) => window.DemoCopilotChat.showQuickActions(actions)",
                            self._post_demo_actions,
                        )
                    await self._page.evaluate(
                        "window.DemoCopilotChat.setStatus('Completed', '')"
                    )
                    logger.info("Restored completed demo state after re-injection")
                else:
                    logger.info("Restored step tracker state after re-injection")
            except Exception as e:
                logger.warning("Failed to restore persisted demo state: %s", e)

    def start_navigation_watcher(self):
        """Register a listener that re-injects the chat panel after page navigations.

        D365 is an SPA that can trigger late client-side navigations (e.g. after
        SSO redirects or slow module loads). These destroy the injected DOM.
        This watcher automatically re-injects the chat panel when that happens.
        """
        if getattr(self, "_nav_watcher_registered", False):
            return
        self._nav_watcher_registered = True
        self._reinject_lock = asyncio.Lock()

        async def _on_frame_navigated(frame):
            """Re-inject chat panel when the main frame navigates."""
            if frame != self._page.main_frame:
                return  # Ignore child iframes
            logger.info("Main frame navigated — scheduling chat panel re-injection")
            # Small delay to let the new DOM settle
            await asyncio.sleep(1.5)
            async with self._reinject_lock:
                try:
                    await self._ensure_injected()
                except Exception as e:
                    logger.warning("Re-injection after navigation failed: %s", e)

        self._page.on("framenavigated", lambda frame: asyncio.ensure_future(_on_frame_navigated(frame)))
        logger.info("Navigation watcher registered — chat panel will auto-reinject")

    async def destroy(self):
        """Remove chat panel from the page."""
        try:
            await self._page.evaluate(
                """() => {
                    const host = document.getElementById('demo-chat-panel-host');
                    if (host) host.remove();
                    const toggle = document.getElementById('demo-chat-toggle');
                    if (toggle) toggle.remove();
                    delete window.DemoCopilotChat;
                }"""
            )
        except Exception:
            pass
        self._injected = False
