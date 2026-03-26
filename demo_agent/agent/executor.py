"""
D365 Demo Copilot — Demo Executor

The heart of the demo agent. Orchestrates the full demo lifecycle:
1. Title slide
2. For each section:
   a. Section transition narration
   b. For each step (Tell-Show-Tell):
      i.   TELL: Show narration caption explaining what's next
      ii.  SHOW: Execute browser actions with spotlight/tooltips
      iii. TELL: Show summary caption + business value card
   c. Pause if requested
3. Closing slide

Respects pause/resume commands throughout execution.
"""

from __future__ import annotations

import asyncio
import logging
import re
import time
from typing import Callable, Optional

import requests

from ..browser.controller import BrowserController
from ..browser.d365_pages import D365Navigator
from ..browser.overlay_manager import OverlayManager
from ..models.demo_plan import (
    ActionType,
    DemoPlan,
    DemoSection,
    DemoStep,
    StepAction,
)
from .schema_discovery import PageIntrospector
from .script_recorder import ScriptRecorder
from .state import DemoState, DemoStatus
from .voice import VoiceNarrator

logger = logging.getLogger("demo_agent.executor")


class PageUnhealthyError(RuntimeError):
    """Raised when an action fails because the page is in a broken state."""

    def __init__(self, health: dict):
        self.health = health
        super().__init__(health.get("message", "Page unhealthy"))


class DemoExecutor:
    """
    Executes a DemoPlan in the browser with visual overlays.

    The executor drives a Playwright browser through the D365 UI,
    injecting spotlight effects, captions, and business value cards
    at each step following the Tell-Show-Tell pattern.
    """

    def __init__(
        self,
        browser: BrowserController,
        overlay: OverlayManager,
        d365_nav: D365Navigator,
        state: DemoState,
        on_status_change: Optional[Callable[[DemoState], None]] = None,
        voice: Optional[VoiceNarrator] = None,
        script_recorder: Optional[ScriptRecorder] = None,
        dataverse_api_url: Optional[str] = None,
        auth_headers: Optional[dict[str, str]] = None,
    ):
        self.browser = browser
        self.overlay = overlay
        self.d365_nav = d365_nav
        self.state = state
        self._on_status_change = on_status_change
        self._voice = voice
        self._script_recorder = script_recorder
        self._dataverse_api_url = dataverse_api_url  # e.g. https://org.crm.dynamics.com/api/data/v9.2
        self._auth_headers = auth_headers or {}
        self._pending_field_updates: list[dict] = []  # tracks fill actions for post-step verification
        self._preflight_data: dict = {}  # stores projectName, taskName, etc. from preflight

    def set_preflight_data(self, data: dict):
        """Store verified preflight data so lookup methods can use it directly."""
        self._preflight_data = data or {}
        logger.info("[PREFLIGHT DATA] Stored: project=%s, task=%s",
                    data.get("projectName"), data.get("taskName"))

    def _notify(self):
        """Notify listeners of state change."""
        if self._on_status_change:
            self._on_status_change(self.state)

    async def _interruptible_sleep(self, seconds: float) -> bool:
        """Sleep that can be cut short by Ctrl+Space (advance event).

        The advance event is **not** cleared — it stays set so subsequent
        interruptible sleeps also resolve instantly until the event is
        explicitly cleared (at the next step boundary or by
        ``wait_for_step_trigger``).

        Returns True if the sleep was interrupted (advance signalled).
        """
        if seconds <= 0:
            return False
        if self.state._advance_event.is_set():
            return True
        try:
            await asyncio.wait_for(
                self.state._advance_event.wait(),
                timeout=seconds,
            )
            return True
        except asyncio.TimeoutError:
            return False

    async def execute(self, plan: DemoPlan):
        """
        Execute a complete demo plan.

        This is the main entry point. It runs the full demo lifecycle
        including title slide, all sections/steps, and closing.
        """
        logger.info("Starting demo: %s (%d steps)", plan.title, plan.total_steps)

        # Initialize state
        self.state.start(plan.total_steps, len(plan.sections))
        self._notify()

        try:
            # Ensure overlay is injected
            await self.overlay.inject()

            # Initialize progress
            await self.overlay.init_progress(plan.total_steps)

            # ---- Title Slide ----
            await self._show_title_slide(plan)

            # ---- Execute Sections ----
            global_step_idx = 0
            for section_idx, section in enumerate(plan.sections):
                self.state.current_section_index = section_idx
                self._notify()

                # Section transition
                if section_idx > 0:
                    await self._section_transition(
                        plan.sections[section_idx - 1], section
                    )

                # Execute each step in the section
                for local_step_idx, step in enumerate(section.steps):
                    self.state.current_step_index = global_step_idx
                    self.state.current_local_step_index = local_step_idx
                    self._notify()

                    # Clear any lingering advance signal from previous step
                    self.state._advance_event.clear()

                    step_start = time.time()

                    # Re-inject overlay if page navigated
                    await self.overlay.ensure_injected()
                    await self.overlay.update_progress(global_step_idx)

                    # ---- Page health gate ----
                    health = await self._check_page_health()
                    if not health["healthy"]:
                        recovered = await self._try_recover_page(
                            step, plan.d365_base_url,
                        )
                        if not recovered:
                            logger.error(
                                "Cannot recover page for step %s — skipping",
                                step.id,
                            )
                            await self.overlay.show_caption(
                                "⚠ Page error — skipping this step",
                                phase="show",
                                position="bottom",
                            )
                            await self._interruptible_sleep(2.0)
                            await self.overlay.hide_caption()
                            self.state.record_step(
                                step.id, section.id, step_start, skipped=True,
                            )
                            if self._script_recorder:
                                await self._script_recorder.capture_step(
                                    step=step, section=section,
                                    global_step_index=global_step_idx, skipped=True,
                                )
                            await self.overlay.clear_all()
                            await self._interruptible_sleep(0.5)
                            global_step_idx += 1
                            continue
                        # After recovery the overlay may be gone — re-inject
                        await self.overlay.ensure_injected()
                        await self.overlay.update_progress(global_step_idx)

                    # ---- TELL (Before) ----
                    await self._phase_tell_before(step)

                    # ---- Wait: user acts OR presses CTRL+SPACE ----
                    has_watchers = await self._setup_action_watchers(step)

                    # Spotlight the first target so user sees where to act
                    first_sel = self._first_actionable_selector(step)
                    if first_sel:
                        try:
                            await self.overlay.spotlight_on(first_sel)
                        except Exception:
                            pass

                    try:
                        await self.overlay.show_status("✉ Your turn — CTRL+SPACE to auto-fill", "waiting")
                    except Exception:
                        pass

                    if has_watchers:
                        await self.overlay.show_caption(
                            "\U0001f446 Do it yourself, or press "
                            "<span class='highlight'>CTRL+SPACE</span> to auto-fill",
                            phase="show",
                            position="bottom",
                        )
                    else:
                        await self.overlay.show_caption(
                            "\u23ce Press <span class='highlight'>CTRL+SPACE</span> "
                            "to continue",
                            phase="tell",
                            position="bottom",
                        )

                    trigger = await self.state.wait_for_step_trigger()
                    await self._teardown_action_watchers()
                    await self.overlay.hide_caption()

                    try:
                        await self.overlay.hide_status()
                    except Exception:
                        pass

                    if first_sel:
                        try:
                            await self.overlay.spotlight_off()
                        except Exception:
                            pass

                    # ---- SHOW (conditional) ----
                    try:
                        if trigger == "advance":
                            # Ctrl+Space pressed — agent executes all actions
                            await self._phase_show(step)
                        else:
                            # User performed field actions manually
                            logger.info("[USER] Manual completion: %s", step.id)
                            self.state.set_phase("show")
                            self._notify()
                            # Execute remaining structural actions (Save, etc.)
                            remaining = [
                                a for a in step.actions
                                if a.action_type not in (
                                    ActionType.FILL, ActionType.SELECT,
                                )
                            ]
                            if remaining:
                                await asyncio.sleep(0.5)  # let D365 settle
                                for action in remaining:
                                    await self._execute_action(action)
                            await self.overlay.show_caption(
                                "\u2713 Got it \u2014 moving on\u2026",
                                phase="show",
                                position="bottom",
                            )
                            await self._interruptible_sleep(1.0)
                            await self.overlay.hide_caption()
                    except PageUnhealthyError:
                        # Page broke during SHOW — attempt recovery and skip step
                        logger.warning(
                            "[RECOVERY] Page broke during SHOW for step %s",
                            step.id,
                        )
                        recovered = await self._try_recover_page(
                            step, plan.d365_base_url,
                        )
                        if not recovered:
                            await self.overlay.show_caption(
                                "⚠ Page error — skipping this step",
                                phase="show",
                                position="bottom",
                            )
                            await self._interruptible_sleep(2.0)
                            await self.overlay.hide_caption()
                        self.state.record_step(
                            step.id, section.id, step_start, skipped=True,
                        )
                        if self._script_recorder:
                            await self._script_recorder.capture_step(
                                step=step, section=section,
                                global_step_index=global_step_idx, skipped=True,
                            )
                        await self.overlay.clear_all()
                        await self._interruptible_sleep(0.5)
                        global_step_idx += 1
                        continue

                    # ---- Capture for demo script (after SHOW, before TELL AFTER) ----
                    if self._script_recorder:
                        await self._script_recorder.capture_step(
                            step=step,
                            section=section,
                            global_step_index=global_step_idx,
                        )

                    # ---- TELL (After) ----
                    await self._phase_tell_after(step)

                    # ---- Business Value ----
                    if step.value_highlight:
                        await self._phase_value(step)

                    # Record completion
                    self.state.record_step(
                        step.id, section.id, step_start
                    )

                    # Pause after step if flagged (disabled — agent runs continuously)
                    # Users can pause via the sidecar chat panel if needed.
                    # if step.pause_after:
                    #     await self._pause_for_user()

                    # Clear overlays between steps
                    await self.overlay.clear_all()
                    await self._interruptible_sleep(0.5)

                    global_step_idx += 1

                # Section complete
                self.state.advance_section()

            # ---- Closing ----
            await self._show_closing(plan)

            self.state.complete()
            self._notify()
            logger.info(
                "Demo completed in %s (%d steps)",
                self.state.elapsed_display,
                len(self.state.history),
            )

        except asyncio.CancelledError:
            logger.info("Demo cancelled")
            self.state.set_error("Demo was cancelled")
            raise
        except Exception as e:
            logger.error("Demo execution error: %s", e, exc_info=True)
            self.state.set_error(str(e))
            self._notify()
            raise

    # ---- Phase Handlers ----

    async def _phase_tell_before(self, step: DemoStep):
        """TELL phase 1: Explain what we're about to show."""
        self.state.set_phase("tell_before")
        self._notify()

        await self.state.wait_if_paused()
        logger.info("[TELL] %s: %s", step.id, step.tell_before[:80])

        try:
            await self.overlay.show_status("Narrating…", "narrating")
        except Exception:
            pass

        # Start voice narration in background (plays concurrently with caption)
        if self._voice:
            await self._voice.speak_async(step.tell_before)

        await self.overlay.show_caption_animated(
            step.tell_before,
            phase="tell",
            speed=step.caption_speed,
            position="top",  # TELL captions at top so audience reads while seeing the page
        )

        # Hold the caption for reading time (skippable via Ctrl+Space)
        # If voice is playing, wait for it to finish instead of the fixed timer
        if self._voice and self._voice.enabled:
            await self._voice.wait_for_completion()
        else:
            await self._interruptible_sleep(max(2.0, len(step.tell_before) * 0.04))
        await self.state.wait_if_paused()

        try:
            await self.overlay.hide_status()
        except Exception:
            pass

    async def _phase_show(self, step: DemoStep):
        """SHOW phase: Execute browser actions with visual feedback.

        Actions always execute to completion — Ctrl+Space only skips cosmetic
        delays (delay_before / delay_after), never aborts an active action.

        Tracks action failures in ``self._last_step_failures`` so the
        TELL AFTER phase can report honestly.
        """
        self.state.set_phase("show")
        self._notify()
        self._last_step_failures: list[str] = []
        self._pending_field_updates: list[dict] = []  # Reset for this step

        # Clear advance so the Ctrl+Space that triggered SHOW doesn't
        # immediately skip the cosmetic delays inside the action loop.
        self.state._advance_event.clear()

        await self.overlay.show_caption(
            f"👁 <span class='highlight'>{step.title}</span>",
            phase="show",
            position="auto",  # Adapts based on spotlight target location
        )

        for action in step.actions:
            await self.state.wait_if_paused()

            # Delay before action (skippable via Ctrl+Space)
            if action.delay_before_ms > 0:
                await self._interruptible_sleep(action.delay_before_ms / 1000)

            # Execute the action — always runs to completion
            success = await self._execute_action(action)
            if not success:
                desc = action.description or action.selector or action.action_type.value
                self._last_step_failures.append(desc)

            # Re-show SHOW caption after navigation (DOM may have been rebuilt)
            if action.action_type == ActionType.NAVIGATE:
                await self.overlay.show_caption(
                    f"👁 <span class='highlight'>{step.title}</span>",
                    phase="show",
                    position="auto",
                )

            # Delay after action (skippable via Ctrl+Space)
            if action.delay_after_ms > 0:
                await self._interruptible_sleep(action.delay_after_ms / 1000)

        # ---- Post-step: verify fill values and apply MCP fallback if needed ----
        if self._pending_field_updates and self._dataverse_api_url and self._auth_headers:
            await self._apply_mcp_fallback_for_unsaved_fields(step)

    # ---- Dataverse Web API Fallback ----

    async def _apply_mcp_fallback_for_unsaved_fields(self, step: DemoStep):
        """When Playwright fill didn't persist values, update via Dataverse Web API and refresh."""
        if not self._pending_field_updates:
            return

        # Extract the record's entity type and GUID from the current page URL
        record_info = await self._get_current_record_info()
        if not record_info:
            logger.warning("[MCP FALLBACK] Could not determine current record — skipping")
            self._pending_field_updates.clear()
            return

        entity_collection = record_info["collection"]
        record_id = record_info["id"]
        logger.info("[MCP FALLBACK] Applying %d unsaved fields to %s/%s",
                    len(self._pending_field_updates), entity_collection, record_id)

        # Build the PATCH payload — resolve lookup values via Dataverse queries
        patch_data = {}
        for field_update in self._pending_field_updates:
            field_payload = await self._resolve_field_for_api(
                field_update, entity_collection
            )
            if field_payload:
                patch_data.update(field_payload)

        if patch_data:
            try:
                await self.overlay.show_status("Saving values via API…", "working")
            except Exception:
                pass

            success = await self._patch_record(entity_collection, record_id, patch_data)
            if success:
                logger.info("[MCP FALLBACK] Updated record successfully — refreshing form")
                # Refresh the D365 form to reflect API changes
                await self.browser.page.evaluate("""() => {
                    // D365 UCI form refresh
                    if (typeof Xrm !== 'undefined' && Xrm.Page && Xrm.Page.data) {
                        Xrm.Page.data.refresh(false);
                    } else {
                        location.reload();
                    }
                }""")
                await asyncio.sleep(2.0)
                await self._wait_for_d365_form_load()
                # Re-inject overlay after refresh
                await self.overlay.inject(force=True)
            else:
                logger.warning("[MCP FALLBACK] API update failed")
                for fu in self._pending_field_updates:
                    desc = f"Save {fu.get('value')} to {fu.get('selector', 'field')}"
                    self._last_step_failures.append(desc)

            try:
                await self.overlay.hide_status()
            except Exception:
                pass

        self._pending_field_updates.clear()

    async def _get_current_record_info(self) -> Optional[dict]:
        """Extract entity collection name and record GUID from the current D365 form URL."""
        url = self.browser.page.url
        # D365 UCI URL patterns:
        # .../main.aspx?etn=msdyn_timeentry&id={GUID}&pagetype=entityrecord
        # .../main.aspx?pagetype=entityrecord&etn=msdyn_timeentry&id={GUID}
        etn_match = re.search(r'etn=([a-zA-Z_]+)', url)
        id_match = re.search(r'id=(%7[Bb])?([0-9a-fA-F\-]{36})(%7[Dd])?', url)

        if not etn_match or not id_match:
            # Try to get it from Xrm context
            info = await self.browser.page.evaluate("""() => {
                try {
                    if (typeof Xrm !== 'undefined' && Xrm.Page && Xrm.Page.data) {
                        const entity = Xrm.Page.data.entity;
                        return {
                            entityName: entity.getEntityName(),
                            id: entity.getId().replace(/[{}]/g, '')
                        };
                    }
                } catch(e) {}
                return null;
            }""")
            if info:
                collection = self._entity_to_collection(info["entityName"])
                return {"collection": collection, "id": info["id"]}
            return None

        entity_name = etn_match.group(1)
        record_id = id_match.group(2)
        collection = self._entity_to_collection(entity_name)
        return {"collection": collection, "id": record_id}

    @staticmethod
    def _entity_to_collection(entity_name: str) -> str:
        """Convert D365 entity logical name to Web API collection name."""
        # Common irregular pluralizations
        irregulars = {
            "bookableresource": "bookableresources",
            "bookableresourcecategory": "bookableresourcecategories",
            "pricelevel": "pricelevels",
            "salesorder": "salesorders",
            "salesorderdetail": "salesorderdetails",
            "transactioncurrency": "transactioncurrencies",
            "systemuser": "systemusers",
            "account": "accounts",
        }
        if entity_name in irregulars:
            return irregulars[entity_name]
        # Standard D365: logical name + "s" (with "y" → "ies" rule)
        if entity_name.endswith("y") and not entity_name.endswith("ay"):
            return entity_name[:-1] + "ies"
        if entity_name.endswith("s"):
            return entity_name + "es"
        return entity_name + "s"

    async def _resolve_field_for_api(
        self, field_update: dict, entity_collection: str
    ) -> Optional[dict]:
        """Resolve a failed fill action into a Dataverse Web API PATCH payload field."""
        selector = field_update.get("selector", "")
        value = field_update.get("value", "")
        field_type = field_update.get("field_type", "")

        # Extract the logical name from the selector
        logical_name = None
        m = re.search(r'data-id=["\']([^."\']+)', selector)
        if m:
            logical_name = m.group(1).lower()
        if not logical_name:
            # Try from aria-label
            m = re.search(r'aria-label=["\']([^"\']+)', selector)
            if m:
                label = m.group(1).lower().replace(" ", "_")
                logical_name = f"msdyn_{label}"

        if not logical_name:
            logger.warning("[MCP FALLBACK] Cannot determine field name from selector: %s", selector)
            return None

        if field_type == "lookup":
            # For lookups, query Dataverse to find the target record GUID
            return await self._resolve_lookup_for_api(logical_name, value)

        # For simple fields (duration, text, etc.)
        # Duration: D365 stores in minutes
        if "duration" in logical_name:
            # Convert hours to minutes
            hours_match = re.match(r'^(\d+(?:\.\d+)?)\s*h?$', value.strip())
            if hours_match:
                minutes = int(float(hours_match.group(1)) * 60)
                return {logical_name: minutes}
            return {logical_name: value}

        return {logical_name: value}

    async def _resolve_lookup_for_api(
        self, logical_name: str, search_value: str
    ) -> Optional[dict]:
        """Resolve a lookup field value to a Dataverse bind reference.

        Queries Dataverse to find the target record and returns the
        OData bind syntax: {"fieldname@odata.bind": "/collection(GUID)"}
        """
        # Map common lookup fields to their target entity/collection
        lookup_targets = {
            "msdyn_project": ("msdyn_projects", "msdyn_subject"),
            "msdyn_projecttask": ("msdyn_projecttasks", "msdyn_subject"),
            "msdyn_bookableresource": ("bookableresources", "name"),
            "msdyn_resourcecategory": ("bookableresourcecategories", "name"),
            "msdyn_expensecategory": ("msdyn_expensecategories", "msdyn_name"),
            "msdyn_transactioncategory": ("msdyn_transactioncategories", "msdyn_name"),
            "msdyn_contractorganizationalunitid": ("msdyn_organizationalunits", "msdyn_name"),
            "msdyn_organizationalunit": ("msdyn_organizationalunits", "msdyn_name"),
        }

        target = lookup_targets.get(logical_name)
        if not target:
            logger.warning("[MCP FALLBACK] Unknown lookup target for: %s", logical_name)
            return None

        collection, name_field = target
        # Query Dataverse for the record
        filter_value = search_value.replace("'", "''")
        query_url = (
            f"{self._dataverse_api_url}/{collection}"
            f"?$filter=contains({name_field},'{filter_value}')"
            f"&$select={name_field}&$top=1"
        )
        try:
            headers = {**self._auth_headers, "Accept": "application/json", "OData-MaxVersion": "4.0", "OData-Version": "4.0"}
            resp = requests.get(query_url, headers=headers, timeout=10)
            resp.raise_for_status()
            data = resp.json()
            records = data.get("value", [])
            if records:
                record_id = records[0].get(f"{collection.rstrip('s')}id") or records[0].get(f"{logical_name}id")
                # Try generic ID extraction
                if not record_id:
                    for key, val in records[0].items():
                        if key.endswith("id") and isinstance(val, str) and len(val) == 36:
                            record_id = val
                            break
                if record_id:
                    bind_key = f"{logical_name}@odata.bind"
                    return {bind_key: f"/{collection}({record_id})"}
                logger.warning("[MCP FALLBACK] Found record but no ID for %s", logical_name)
            else:
                logger.warning("[MCP FALLBACK] No records found for %s='%s'", name_field, search_value)
        except Exception as exc:
            logger.warning("[MCP FALLBACK] Lookup query failed for %s: %s", logical_name, exc)
        return None

    async def _patch_record(
        self, entity_collection: str, record_id: str, data: dict
    ) -> bool:
        """PATCH a Dataverse record via Web API."""
        url = f"{self._dataverse_api_url}/{entity_collection}({record_id})"
        headers = {
            **self._auth_headers,
            "Content-Type": "application/json",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "If-Match": "*",
        }
        try:
            import json
            logger.info("[MCP FALLBACK] PATCH %s — %s", url, json.dumps(data, indent=2))
            resp = requests.patch(url, json=data, headers=headers, timeout=15)
            if resp.status_code in (200, 204):
                logger.info("[MCP FALLBACK] PATCH succeeded (HTTP %d)", resp.status_code)
                return True
            logger.warning("[MCP FALLBACK] PATCH failed: HTTP %d — %s", resp.status_code, resp.text[:300])
            return False
        except Exception as exc:
            logger.warning("[MCP FALLBACK] PATCH exception: %s", exc)
            return False

    async def _post_record(
        self, entity_collection: str, data: dict
    ) -> str | None:
        """POST (create) a Dataverse record via Web API.

        Returns the GUID of the created record, or None on failure.
        """
        if not self._dataverse_api_url or not self._auth_headers:
            logger.warning("[SAMPLE DATA] No Dataverse API URL / auth — cannot create records")
            return None
        url = f"{self._dataverse_api_url}/{entity_collection}"
        headers = {
            **self._auth_headers,
            "Content-Type": "application/json",
            "Accept": "application/json",
            "OData-MaxVersion": "4.0",
            "OData-Version": "4.0",
            "Prefer": "return=representation",
        }
        try:
            import json as _json
            logger.info("[SAMPLE DATA] POST %s — %s", url, _json.dumps(data, indent=2))
            resp = requests.post(url, json=data, headers=headers, timeout=20)
            if resp.status_code in (200, 201):
                body = resp.json()
                # Extract the primary key — Dataverse returns it in the response
                record_id = None
                for key, val in body.items():
                    if key.endswith("id") and isinstance(val, str) and len(val) == 36:
                        record_id = val
                        break
                logger.info("[SAMPLE DATA] Created %s — ID: %s", entity_collection, record_id)
                return record_id
            logger.warning(
                "[SAMPLE DATA] POST failed: HTTP %d — %s",
                resp.status_code, resp.text[:500],
            )
            return None
        except Exception as exc:
            logger.warning("[SAMPLE DATA] POST exception: %s", exc)
            return None

    async def _phase_tell_after(self, step: DemoStep):
        """TELL phase 2: Summarize what was demonstrated.

        If critical actions failed during SHOW, append a warning rather than
        claiming success.
        """
        self.state.set_phase("tell_after")
        self._notify()

        await self.state.wait_if_paused()

        # Build narration — add failure warning if any actions failed
        failures = getattr(self, "_last_step_failures", [])
        narration = step.tell_after
        if failures:
            warning = " ⚠ <em>Note: some actions could not be completed automatically.</em>"
            narration = narration + warning
            logger.warning("[TELL AFTER] %s — %d action(s) failed: %s",
                           step.id, len(failures), "; ".join(failures))
        else:
            logger.info("[TELL AFTER] %s: %s", step.id, step.tell_after[:80])

        await self.overlay.spotlight_off()
        try:
            await self.overlay.show_status("Summarizing…", "narrating")
        except Exception:
            pass

        # Start voice narration in background (plays concurrently with caption)
        if self._voice:
            await self._voice.speak_async(narration)

        await self.overlay.show_caption_animated(
            narration,
            phase="tell",
            speed=step.caption_speed,
            position="bottom",  # Summary captions at bottom
        )

        # Hold for reading time (skippable via Ctrl+Space)
        # If voice is playing, wait for it to finish instead of the fixed timer
        if self._voice and self._voice.enabled:
            await self._voice.wait_for_completion()
        else:
            await self._interruptible_sleep(max(2.0, len(narration) * 0.04))
        await self.state.wait_if_paused()

        try:
            await self.overlay.hide_status()
        except Exception:
            pass

    async def _phase_value(self, step: DemoStep):
        """Business Value phase: Show value callout card."""
        self.state.set_phase("value")
        self._notify()

        # Clear any lingering advance signal so the value card isn't
        # instantly dismissed by a Ctrl+Space press from a prior phase.
        self.state._advance_event.clear()

        vh = step.value_highlight
        await self.overlay.show_caption(
            f"💡 {vh.description}",
            phase="value",
            position="bottom",  # Value captions at bottom
        )

        await self.overlay.show_value_card(
            title=vh.title,
            text=vh.description,
            metric_value=vh.metric_value,
            metric_label=vh.metric_label,
            position=vh.position,
        )

        # Narrate the value statement
        if self._voice:
            narration = f"{vh.title}. {vh.description}"
            if vh.metric_value and vh.metric_label:
                narration += f" — {vh.metric_value} {vh.metric_label}."
            await self._voice.speak_async(narration)

        # Hold the value card for reading (skippable via Ctrl+Space).
        # Base 8s, extended for longer text / metric animation.
        hold = max(8.0, len(vh.description) * 0.06)
        if vh.metric_value:
            hold = max(hold, 10.0)  # Extra time for animated metric
        # If voice is playing, wait for it to finish instead of the fixed timer
        if self._voice and self._voice.enabled:
            await self._voice.wait_for_completion()
        else:
            await self._interruptible_sleep(hold)
        await self.state.wait_if_paused()

        await self.overlay.hide_value_card()

    # ---- Action Watching (user-completion detection) ----

    _WATCHER_JS = """(conditions) => {
        if (window.__demoWatcherCleanup) window.__demoWatcherCleanup();
        if (!conditions || conditions.length === 0) return;
        const completed = new Set();
        const total = conditions.length;
        let intervalId = null;
        function val(sel) {
            try {
                const el = document.querySelector(sel);
                if (!el) return '';
                if (typeof el.value === 'string' && el.value !== '') return el.value;
                if (el.isContentEditable) return (el.textContent || '').trim();
                const n = el.querySelector('input,textarea,select,[contenteditable="true"]');
                if (n) {
                    if (typeof n.value === 'string' && n.value !== '') return n.value;
                    if (n.isContentEditable) return (n.textContent || '').trim();
                }
                return '';
            } catch (e) { return ''; }
        }
        const snap = conditions.map(c => val(c.selector));
        function check() {
            for (let i = 0; i < conditions.length; i++) {
                if (completed.has(i)) continue;
                const cur = val(conditions[i].selector);
                if (cur.length > 0 && cur !== snap[i]) completed.add(i);
            }
            if (completed.size >= total) {
                cleanup();
                if (typeof window.__demoCopilotAction === 'function')
                    window.__demoCopilotAction('user_acted');
            }
        }
        intervalId = setInterval(check, 400);
        function cleanup() {
            if (intervalId) { clearInterval(intervalId); intervalId = null; }
            delete window.__demoWatcherCleanup;
        }
        window.__demoWatcherCleanup = cleanup;
    }"""

    _PW_MARKERS = (":has-text(", ":text(", ">>", ":visible", ":nth-match(")

    async def _setup_action_watchers(self, step: DemoStep) -> bool:
        """Inject JS that detects when the user fills in this step's fields.

        Only watches FILL / SELECT actions whose selectors are standard CSS
        (Playwright-specific pseudo-selectors are skipped).

        Returns True if at least one watchable condition was set up.
        """
        conditions: list[dict] = []
        for action in step.actions:
            if action.action_type not in (ActionType.FILL, ActionType.SELECT):
                continue
            sel = action.selector
            if not sel or any(m in sel for m in self._PW_MARKERS):
                continue
            conditions.append({"type": "fill", "selector": sel})

        if not conditions:
            return False

        try:
            await self.browser.page.evaluate(self._WATCHER_JS, conditions)
        except Exception as e:
            logger.warning("Failed to set up action watchers: %s", e)
            return False
        return True

    async def _teardown_action_watchers(self):
        """Remove any active JS action watchers from the page."""
        try:
            await self.browser.page.evaluate(
                "() => { if (typeof window.__demoWatcherCleanup === 'function') window.__demoWatcherCleanup(); }"
            )
        except Exception:
            pass

    @staticmethod
    def _first_actionable_selector(step: DemoStep) -> str | None:
        """Return the CSS selector of the first user-facing action in the step."""
        for action in step.actions:
            if action.selector and action.action_type in (
                ActionType.FILL, ActionType.SELECT, ActionType.CLICK, ActionType.HOVER,
            ):
                return action.selector
        return None

    # ---- Page Health & Recovery ----

    _PAGE_HEALTH_JS = """() => {
        const result = { healthy: true, error_type: null, message: '' };

        // 1. Error dialogs (D365 modal errors, blocking dialogs)
        const dlgSels = [
            'div[data-id="errorDialog"]',
            'div[role="alertdialog"]',
            'div[role="dialog"][aria-label*="Error"]',
            'div[role="dialog"][aria-label*="error"]',
            'div.ms-Dialog--blocking',
        ];
        for (const sel of dlgSels) {
            const el = document.querySelector(sel);
            if (el && el.offsetParent !== null) {
                const text = (el.innerText || '').substring(0, 300);
                if (/error|something went wrong|not found|access denied|can't|cannot|failed|unexpected/i.test(text) || text.length > 10) {
                    result.healthy = false;
                    result.error_type = 'dialog';
                    result.message = text.replace(/\\s+/g, ' ').trim().substring(0, 200);
                    return result;
                }
            }
        }

        // 2. D365 error notification banners
        const notif = document.querySelector('div[data-id="notificationWrapper"]');
        if (notif) {
            const errEl = notif.querySelector('[class*="error" i], [data-id*="error" i]');
            if (errEl && errEl.offsetParent !== null) {
                result.healthy = false;
                result.error_type = 'notification';
                result.message = (errEl.innerText || '').substring(0, 200);
                return result;
            }
        }

        // 3. Login redirect (session expired)
        const href = window.location.href.toLowerCase();
        if (href.includes('login.microsoftonline.com') ||
            href.includes('/adfs/') ||
            href.includes('login.live.com') ||
            href.includes('aadcdn.msauth.net')) {
            result.healthy = false;
            result.error_type = 'login_redirect';
            result.message = 'Session expired — redirected to login page';
            return result;
        }

        // 4. Generic browser / IIS error pages
        const bodyText = (document.body ? document.body.innerText : '').substring(0, 600);
        if (/something went wrong|page cannot be displayed|this page isn.t working|HTTP ERROR|ERR_|Server Error|404 Not Found|403 Forbidden|500 Internal/i.test(bodyText)) {
            result.healthy = false;
            result.error_type = 'error_page';
            result.message = bodyText.replace(/\\s+/g, ' ').trim().substring(0, 200);
            return result;
        }

        // 5. Blank page (no D365 chrome loaded after readyState=complete)
        if (document.readyState === 'complete') {
            const hasShell = document.querySelector(
                '#shell-container, #CRMAppModuleContainer, #ApplicationShell, ' +
                'div[data-id="mainContent"], div[data-dyn-role="Workspace"]'
            );
            const hasContent = document.querySelector(
                '[data-id="form-body"], [data-id="data-set-body"], ' +
                'div[data-id="CommandBar"], div[data-id="sitemap-body"], ' +
                'div[data-dyn-role], div.workspace-container'
            );
            if (!hasShell && !hasContent) {
                result.healthy = false;
                result.error_type = 'blank_page';
                result.message = 'No D365 content detected on page';
                return result;
            }
        }

        return result;
    }"""

    async def _check_page_health(self) -> dict:
        """Evaluate the page for D365 error states.

        Returns ``{"healthy": True/False, "error_type": str|None, "message": str}``.
        """
        try:
            return await self.browser.page.evaluate(self._PAGE_HEALTH_JS)
        except Exception as e:
            return {"healthy": False, "error_type": "evaluation_error", "message": str(e)[:200]}

    async def _try_recover_page(self, step: DemoStep, base_url: str) -> bool:
        """Attempt to recover from a bad page state.

        Strategy (in order):
        1. Dismiss visible error dialogs and re-check.
        2. If session expired, pause (user must re-authenticate).
        3. Navigate to the step's own NAVIGATE target (if any).
        4. Navigate to the D365 home page.
        5. Browser back as a last resort.

        Returns True if the page is healthy after recovery.
        """
        health = await self._check_page_health()
        if health["healthy"]:
            return True

        error_type = health.get("error_type", "unknown")
        logger.warning(
            "[RECOVERY] Page unhealthy (%s): %s", error_type, health["message"],
        )

        # Show recovery notice
        try:
            await self.overlay.ensure_injected()
            await self.overlay.show_caption(
                "⚙ Recovering — navigating to the correct page…",
                phase="show",
                position="bottom",
            )
        except Exception:
            pass

        # --- 1. Dismiss error dialogs ---
        if error_type == "dialog":
            dismiss_sels = [
                'button[data-id="errorOkButton"]',
                'button[data-id="cancelButton"]',
                'button[data-id="okButton"]',
                'button:has-text("OK")',
                'button:has-text("Close")',
                'button[aria-label="Close"]',
            ]
            for sel in dismiss_sels:
                try:
                    await self.browser.page.click(sel, timeout=2000)
                    await asyncio.sleep(1.0)
                    recheck = await self._check_page_health()
                    if recheck["healthy"]:
                        logger.info("[RECOVERY] Dismissed error dialog")
                        return True
                except Exception:
                    continue

        # --- 2. Session expired → pause for user ---
        if error_type == "login_redirect":
            logger.error("[RECOVERY] Session expired — user must re-authenticate")
            try:
                await self.overlay.show_caption(
                    "⚠ Session expired. Please log in again, then press "
                    "<span class='highlight'>CTRL+SPACE</span> to resume.",
                    phase="show",
                    position="bottom",
                )
            except Exception:
                pass
            self.state.pause()
            self._notify()
            await self.state.wait_if_paused()
            # After user resumes, re-check
            recheck = await self._check_page_health()
            return recheck.get("healthy", False)

        # --- 3. Navigate to the step's own target URL ---
        target_url = self._find_step_target_url(step, base_url)
        if target_url:
            ok = await self._nav_recovery(target_url)
            if ok:
                return True

        # --- 4. Navigate to D365 home page ---
        ok = await self._nav_recovery(base_url)
        if ok:
            return True

        # --- 5. Browser back ---
        try:
            logger.info("[RECOVERY] Trying browser back")
            await self.browser.page.go_back(wait_until="commit", timeout=15000)
            await asyncio.sleep(2.0)
            await self.overlay.inject(force=True)
            recheck = await self._check_page_health()
            if recheck.get("healthy"):
                logger.info("[RECOVERY] Recovered via browser back")
                return True
        except Exception:
            pass

        logger.error("[RECOVERY] All recovery attempts failed")
        return False

    async def _nav_recovery(self, url: str) -> bool:
        """Navigate to *url* and return True if the page is healthy afterwards."""
        try:
            logger.info("[RECOVERY] Navigating to: %s", url)
            await self.browser.page.goto(url, wait_until="commit", timeout=30000)
            await asyncio.sleep(2.0)
            await self.overlay.inject(force=True)
            await self._wait_for_d365_form_load()
            recheck = await self._check_page_health()
            if recheck.get("healthy"):
                logger.info("[RECOVERY] Page healthy after navigating to %s", url)
                return True
        except Exception as e:
            logger.warning("[RECOVERY] Navigation to %s failed: %s", url, e)
        return False

    @staticmethod
    def _find_step_target_url(step: DemoStep, base_url: str) -> str | None:
        """Extract the navigation URL from a step's NAVIGATE actions, if any."""
        for action in step.actions:
            if action.action_type == ActionType.NAVIGATE and action.value:
                url = action.value
                if not url.startswith("http"):
                    url = f"{base_url.rstrip('/')}/{url.lstrip('/')}"
                return url
        return None

    # ---- Action Execution ----

    async def _execute_action(self, action: StepAction) -> bool:
        """Execute a single browser action with visual feedback.

        Returns True if the action succeeded, False if it was skipped due to error.
        Raises PageUnhealthyError if the entire page is broken.
        """
        logger.info("[ACTION] %s: %s — %s", action.action_type.value, action.selector, action.description)

        try:
            # Spotlight the target element before acting
            if action.selector and action.action_type not in (
                ActionType.NAVIGATE, ActionType.WAIT, ActionType.SCREENSHOT, ActionType.CUSTOM_JS
            ):
                try:
                    await self.overlay.ensure_injected()
                    await self.overlay.spotlight_on(action.selector)
                    await self._interruptible_sleep(0.8)  # Let user see the spotlight
                except Exception:
                    pass  # Spotlight is best-effort

            # Show tooltip on every actionable field — use explicit tooltip
            # or auto-generate from the action description so the audience
            # always sees what the agent is doing.
            tooltip_text = action.tooltip
            if not tooltip_text and action.description and action.action_type in (
                ActionType.FILL, ActionType.SELECT, ActionType.CLICK,
                ActionType.HOVER, ActionType.SPOTLIGHT,
            ):
                tooltip_text = action.description
            if tooltip_text and action.selector:
                try:
                    await self.overlay.show_tooltip(action.selector, tooltip_text)
                    await self._interruptible_sleep(1.5)
                except Exception:
                    pass

            # Turn off spotlight before executing — prevents overlay from
            # blocking Playwright interactions even when pointer-events: none
            # isn't fully respected (e.g. some Chromium compositing quirks).
            try:
                await self.overlay.spotlight_off()
            except Exception:
                pass

            # Show working status
            try:
                await self.overlay.show_status(
                    f"Performing: {action.description[:40]}" if action.description else "Working…",
                    "working",
                )
            except Exception:
                pass

            # Perform the action
            match action.action_type:
                case ActionType.NAVIGATE:
                    url = action.value
                    # Fix relative URLs
                    if url and not url.startswith("http"):
                        url = f"{self.browser.base_url}/{url.lstrip('/')}"
                    # Detect placeholder GUIDs (GUID_HERE, {GUID}, 00000000-..., etc.)
                    # and replace with the actual record ID from the current form
                    if url and re.search(
                        r'id=(GUID_HERE|%7[Bb]?GUID%7[Dd]?|\{?GUID\}?|00000000-0000-0000-0000-000000000000)',
                        url, re.IGNORECASE,
                    ):
                        real_info = await self._get_current_record_info()
                        if real_info and real_info.get("id"):
                            url = re.sub(
                                r'id=(GUID_HERE|%7[Bb]?GUID%7[Dd]?|\{?GUID\}?|00000000-0000-0000-0000-000000000000)',
                                f'id={real_info["id"]}',
                                url,
                            )
                            logger.info(
                                "[NAVIGATE] Resolved placeholder GUID → %s",
                                real_info["id"],
                            )
                        else:
                            logger.warning(
                                "[NAVIGATE] URL has placeholder GUID but could not "
                                "resolve current record — skipping navigation"
                            )
                            return True  # Skip gracefully
                    try:
                        await self.overlay.show_status("Loading page…", "loading")
                    except Exception:
                        pass
                    await self.browser.navigate(url)
                    # Wait for D365 page to fully render (loads slowly via SPA)
                    await asyncio.sleep(3.0)
                    # Force re-inject overlay (navigation destroys DOM)
                    await self.overlay.inject(force=True)
                    # Wait for D365 framework to finish rendering forms
                    await self._wait_for_d365_form_load()
                    await asyncio.sleep(0.5)

                case ActionType.CLICK:
                    if action.selector:
                        # Try the selector directly
                        clicked = await self._try_click(action.selector)
                        if not clicked:
                            logger.warning("Primary selector failed, trying alternatives for: %s", action.selector)
                            # Try well-known D365 fallbacks first
                            fallbacks = self.D365_SELECTOR_FALLBACKS.get(action.selector, [])
                            for fb in fallbacks:
                                clicked = await self._try_click(fb)
                                if clicked:
                                    logger.info("Fallback selector worked: %s", fb)
                                    break
                            # Then try the generic alternative selector
                            if not clicked:
                                alt = self._alternative_selector(action.selector)
                                if alt:
                                    clicked = await self._try_click(alt)
                            # Last resort for save buttons — use Ctrl+S keyboard shortcut
                            if not clicked and "save" in (action.selector or "").lower():
                                logger.info("Save button not found — using Ctrl+S keyboard shortcut")
                                await self.browser.page.keyboard.press("Control+s")
                                await asyncio.sleep(2.0)
                                clicked = True
                            # Last resort for submit buttons — find via JS and click
                            if not clicked and "submit" in (action.selector or "").lower():
                                logger.info("Submit button not found via selectors — trying JS command bar search")
                                js_submit = await self.browser.page.evaluate("""() => {
                                    // Search for any visible button with "Submit" text in the command bar
                                    const buttons = document.querySelectorAll(
                                        'button, li[role="menuitem"], span[role="button"]'
                                    );
                                    for (const btn of buttons) {
                                        const text = btn.textContent?.trim() || '';
                                        const label = btn.getAttribute('aria-label') || '';
                                        const dataId = btn.getAttribute('data-id') || '';
                                        if (
                                            (text.toLowerCase().includes('submit') ||
                                             label.toLowerCase().includes('submit') ||
                                             dataId.toLowerCase().includes('submit')) &&
                                            btn.offsetParent !== null
                                        ) {
                                            btn.click();
                                            return text || label || dataId;
                                        }
                                    }
                                    return false;
                                }""")
                                if js_submit:
                                    logger.info("JS Submit clicked: %s", js_submit)
                                    await asyncio.sleep(2.0)
                                    clicked = True
                            # If a lookup dropdown click fails, check if the
                            # lookup was already set via Xrm SDK during the
                            # fill step — if so, skip gracefully.
                            if not clicked and "LookupResultsDropdown" in (action.selector or ""):
                                m_lk = re.search(r'data-id=["\']([^."\']+)', action.selector or "")
                                if m_lk:
                                    lk_field = m_lk.group(1)
                                    resolved = await self.browser.page.evaluate(f"""() => {{
                                        try {{
                                            const attr = Xrm.Page.getAttribute('{lk_field}');
                                            if (attr) {{
                                                const val = attr.getValue();
                                                return val && val.length > 0 ? val[0].name : null;
                                            }}
                                        }} catch(e) {{}}
                                        return null;
                                    }}""")
                                    if resolved:
                                        logger.info("[CLICK] Lookup '%s' already resolved to '%s' — skipping dropdown click", lk_field, resolved)
                                        clicked = True

                            if not clicked:
                                raise Exception(f"Could not find element: {action.selector}")

                case ActionType.FILL:
                    if action.selector and action.value:
                        fill_value = action.value
                        # Pre-normalize duration values before fill attempt
                        if "duration" in (action.selector or "").lower():
                            stripped = fill_value.strip()
                            if re.match(r'^\d+(\.\d+)?$', stripped):
                                fill_value = f"{stripped}h 0m"
                                logger.info("Pre-normalized duration: '%s' → '%s'", stripped, fill_value)
                        await self._try_fill(action.selector, fill_value)

                case ActionType.SELECT:
                    if action.selector and action.value:
                        await self._try_select_option(action.selector, action.value)

                case ActionType.HOVER:
                    if action.selector:
                        await self._try_hover(action.selector)

                case ActionType.SCROLL:
                    if action.selector:
                        await self.browser.page.evaluate(
                            f'document.querySelector("{action.selector}")?.scrollIntoView({{behavior: "smooth"}})'
                        )
                    else:
                        await self.browser.page.evaluate(
                            f'window.scrollBy(0, {action.value or 300})'
                        )

                case ActionType.WAIT:
                    if action.selector:
                        await self.browser.wait_for(action.selector)
                    elif action.value:
                        await asyncio.sleep(float(action.value) / 1000)

                case ActionType.SCREENSHOT:
                    path = action.value or f"demo_step_{self.state.current_step_index}.png"
                    await self.browser.screenshot(path)

                case ActionType.SPOTLIGHT:
                    if action.selector:
                        await self.overlay.spotlight_on(action.selector)
                        await self._interruptible_sleep(2.0)

                case ActionType.CUSTOM_JS:
                    if action.value:
                        await self.browser.evaluate(action.value)

            # Clear tooltip after action
            await self.overlay.hide_tooltip()

            # ---- Post-save validation ----
            # After a save/submit click, wait for D365 to process and check
            # for form-level errors (notifications, validation messages, etc.)
            if action.action_type == ActionType.CLICK and self._is_save_action(action):
                await asyncio.sleep(2.5)  # Give D365 time to process the save
                save_error = await self._check_d365_save_errors()
                if save_error:
                    logger.warning("[SAVE CHECK] D365 error after save: %s", save_error)
                    await self.overlay.show_caption(
                        f"⚠ {save_error[:120]}",
                        phase="show",
                        position="bottom",
                    )
                    await self._interruptible_sleep(3.0)
                    await self.overlay.hide_caption()
                    return False  # Record as failed action

            return True  # Action succeeded

        except Exception as e:
            logger.warning("Action failed (%s on %s): %s", action.action_type.value, action.selector, e)

            # Check if the page itself is broken (error screen, etc.)
            health = await self._check_page_health()
            if not health["healthy"]:
                # Re-raise so the step loop can trigger full recovery
                raise PageUnhealthyError(health) from e

            # Page is OK — just this element was missing. Show soft error.
            await self.overlay.show_caption(
                f"⚠ Skipping action: {action.description}",
                phase="show",
            )
            await self._interruptible_sleep(2.0)
            return False  # Action failed
        finally:
            # Always clear spotlight and status to prevent persistent dim overlay
            try:
                await self.overlay.spotlight_off()
            except Exception:
                pass
            try:
                await self.overlay.hide_status()
            except Exception:
                pass

    # ---- Post-Save Validation Helpers ----

    @staticmethod
    def _is_save_action(action: StepAction) -> bool:
        """Return True if the action looks like a save/submit operation."""
        sel = (action.selector or "").lower()
        desc = (action.description or "").lower()
        return (
            "save" in sel
            or "save" in desc
            or "submit" in sel
            or "submit" in desc
            or "ctrl+s" in desc
        )

    _SAVE_ERROR_JS = """() => {
        const errors = [];

        // 1. Xrm form-level notifications (most reliable)
        try {
            if (typeof Xrm !== 'undefined' && Xrm.Page && Xrm.Page.ui) {
                const notifs = Xrm.Page.ui.getFormNotifications
                    ? Xrm.Page.ui.getFormNotifications()
                    : [];
                for (const n of notifs) {
                    if (n.type === 'ERROR' || n.type === 'WARNING') {
                        errors.push(n.message || n.text || 'Form error');
                    }
                }
            }
        } catch(e) {}

        // 2. Visible notification bars / alert banners
        const barSels = [
            'div[data-id="notificationWrapper"] div[data-id*="notification"]',
            'div[data-id="notificationWrapper"] span',
            'div[role="alert"]',
            'div[class*="MessageBar--error"]',
            'div[data-id*="error_message"]',
            'div[data-id*="warningNotification"]',
        ];
        for (const sel of barSels) {
            for (const el of document.querySelectorAll(sel)) {
                if (el.offsetParent === null) continue;
                const text = (el.innerText || '').trim();
                if (text.length > 5 &&
                    /error|fail|cannot|invalid|missing|required|not (?:a |allowed)|denied|does not/i.test(text)) {
                    errors.push(text.substring(0, 200));
                }
            }
        }

        // 3. Inline field validation errors
        const fieldErrSels = [
            'span[data-id*="error_message_text"]',
            'div[class*="errorMessage"]',
            'span[class*="field-error"]',
        ];
        for (const sel of fieldErrSels) {
            for (const el of document.querySelectorAll(sel)) {
                if (el.offsetParent === null) continue;
                const text = (el.innerText || '').trim();
                if (text.length > 3) errors.push(text.substring(0, 200));
            }
        }

        // 4. Blocking error dialogs
        const dlgSels = [
            'div[data-id="errorDialog"]',
            'div[role="alertdialog"]',
            'div[role="dialog"][aria-label*="Error"]',
            'div[role="dialog"][aria-label*="error"]',
        ];
        for (const sel of dlgSels) {
            const el = document.querySelector(sel);
            if (el && el.offsetParent !== null) {
                const text = (el.innerText || '').substring(0, 300).trim();
                if (text.length > 5) errors.push(text.replace(/\\s+/g, ' ').substring(0, 200));
            }
        }

        // De-duplicate and join
        const unique = [...new Set(errors)];
        return unique.length > 0 ? unique.join(' | ').substring(0, 500) : null;
    }"""

    async def _check_d365_save_errors(self) -> str | None:
        """Check for D365 form errors/notifications after a save operation.

        Returns an error message string if errors are found, ``None`` if clean.
        """
        try:
            return await self.browser.page.evaluate(self._SAVE_ERROR_JS)
        except Exception:
            return None

    # Logical names that are lookup fields — used to route through _fill_d365_lookup
    # even when the broad-scan input doesn't have "LookupResultsDropdown" in data-id.
    _KNOWN_LOOKUP_FIELDS: set[str] = {
        "msdyn_project", "msdyn_projecttask", "msdyn_bookableresource",
        "msdyn_resourcecategory", "msdyn_expensecategory", "msdyn_customer",
        "msdyn_contractorganizationalunitid", "msdyn_organizationalunit",
        "bookableresourcecategory", "msdyn_bookableresourceid",
        "msdyn_transactioncategory", "msdyn_projectmanager",
        "msdyn_contractid", "msdyn_vendor",
    }

    # Common D365 option-set label → integer value mappings.
    _OPTIONSET_VALUE_MAP: dict[str, dict[str, int]] = {
        "msdyn_type": {
            "work": 192350000, "absence": 192350001,
            "vacation": 192350002, "overtime": 192350004,
        },
        "msdyn_entrystatus": {
            "draft": 192350000, "returned": 192350001,
            "approved": 192350002, "submitted": 192350003,
        },
        "msdyn_billingstatus": {
            "unbilled sales created": 192350000, "customer invoice created": 192350001,
            "customer invoice posted": 192350002, "canceled": 192350003,
            "ready to invoice": 192350004,
        },
        "msdyn_billingtype": {
            "non chargeable": 192350000, "chargeable": 192350001,
            "complimentary": 192350002, "not available": 192350003,
        },
    }

    # Well-known D365 selector fallbacks for common buttons that vary by page
    D365_SELECTOR_FALLBACKS: dict[str, list[str]] = {
        'button[data-id="edit-form-new-btn"]': [
            'button[aria-label*="New"]',
            'button[data-id="quickCreateLauncher"]',
            'button:has-text("New")',
        ],
        'button[data-id="edit-form-save-btn"]': [
            'button[aria-label="Save"]',
            'button[aria-label*="Save (CTRL+S)"]',
            'button[data-id*="save"]',
            'button:has-text("Save")',
        ],
        'button[data-id="edit-form-save-and-close-btn"]': [
            'button[aria-label="Save & Close"]',
            'button[aria-label*="Save and Close"]',
            'button[data-id="edit-form-save-btn"]',
            'button[aria-label="Save"]',
            'button[aria-label*="Save (CTRL+S)"]',
            'button[data-id*="save"]',
            'button:has-text("Save & Close")',
            'button:has-text("Save")',
        ],
        'button[data-id="edit-form-delete-btn"]': [
            'button[aria-label="Delete"]',
            'button:has-text("Delete")',
        ],
        # Submit button — D365 command bar varies by form
        'button[data-id="msdyn_Submit"]': [
            'button[aria-label="Submit"]',
            'button[aria-label*="Submit"]',
            'button[data-id*="Submit"]',
            'button[data-id*="submit"]',
            'button:has-text("Submit")',
            'li[data-id*="Submit"] button',
            'button[command="msdyn_Submit"]',
            'span[data-id*="Submit"]',
        ],
        # F&O button fallbacks
        'button[data-dyn-controlname="SystemDefinedNewButton"]': [
            'button[aria-label*="New"]',
            'button:has-text("New")',
        ],
        'button[data-dyn-controlname="SystemDefinedSaveButton"]': [
            'button[aria-label*="Save"]',
            'button:has-text("Save")',
        ],
    }

    # ---- D365 Wait / Settle Helpers ----

    async def _wait_for_d365_form_load(self, timeout: int = 10000):
        """Wait for a D365 form to finish loading.

        CE forms: waits for the form body or save button to appear.
        F&O forms: waits for a form group or action pane to appear.
        """
        # Try CE form indicators
        ce_selectors = [
            'button[data-id="edit-form-save-btn"]',
            'div[data-id="form-body"]',
            'div[data-id="header_overflowButton"]',
        ]
        # Try F&O form indicators
        fo_selectors = [
            'button[data-dyn-controlname="SystemDefinedSaveButton"]',
            'div[data-dyn-controlname]',
            'div.workspace-container',
        ]

        all_selectors = ce_selectors + fo_selectors
        settled = False

        for sel in all_selectors:
            try:
                await self.browser.page.wait_for_selector(
                    sel, state="visible", timeout=timeout // len(all_selectors)
                )
                settled = True
                logger.debug("Form settled — detected: %s", sel)
                break
            except Exception:
                continue

        if not settled:
            # Fallback: just wait a bit more
            logger.debug("No form settlement indicator found — waiting 2s extra")
            await asyncio.sleep(2.0)

    async def _try_select_option(self, selector: str, value: str, timeout: int = 3000):
        """Select an option in a D365 option set / dropdown.

        D365 CE option sets render as custom dropdowns, not standard <select>.
        This method handles both standard HTML selects and D365 custom dropdowns.
        """
        # Attempt 1: Standard HTML <select> element
        try:
            el = await self.browser.page.wait_for_selector(
                selector, state="visible", timeout=timeout
            )
            tag = await el.evaluate("el => el.tagName.toLowerCase()")
            if tag == "select":
                await self.browser.select_option(selector, value)
                logger.info("Selected option '%s' via standard <select>", value)
                return
        except Exception:
            pass

        # Attempt 2: D365 custom option set — click to open, then select option
        try:
            # Click on the option set control to open the dropdown
            el = await self.browser.page.wait_for_selector(
                selector, state="visible", timeout=2000
            )
            await el.click()
            await asyncio.sleep(0.8)  # Wait for dropdown animation

            # Look for the option by text in the dropdown list
            option_selectors = [
                f'div[role="listbox"] div[role="option"]:has-text("{value}")',
                f'ul[role="listbox"] li[role="option"]:has-text("{value}")',
                f'option:has-text("{value}")',
                f'div[data-id*="fieldControl-option-set"] div:has-text("{value}")',
            ]
            for opt_sel in option_selectors:
                try:
                    opt = await self.browser.page.wait_for_selector(
                        opt_sel, state="visible", timeout=2000
                    )
                    await opt.click()
                    logger.info("Selected option '%s' via custom dropdown", value)
                    return
                except Exception:
                    continue
        except Exception:
            pass

        # Attempt 3: Use JavaScript to set the value directly on the DOM
        try:
            await self.browser.page.evaluate(f"""
                (() => {{
                    const el = document.querySelector('{selector}');
                    if (el && el.tagName === 'SELECT') {{
                        const opt = Array.from(el.options).find(o =>
                            o.text.includes('{value}') || o.value === '{value}'
                        );
                        if (opt) {{
                            el.value = opt.value;
                            el.dispatchEvent(new Event('change', {{bubbles: true}}));
                            return;
                        }}
                    }}
                }})()
            """)
            logger.info("Selected option '%s' via JS fallback", value)
            return
        except Exception:
            pass

        # Attempt 4: Use Xrm SDK to set the option set value directly
        logical_name = None
        m = re.search(r'data-id=["\']([^."\']+)\.', selector)
        if m:
            logical_name = m.group(1)
        if logical_name:
            xrm_ok = await self._set_d365_optionset_via_xrm(logical_name, value)
            if xrm_ok:
                return

        logger.warning("Could not select option '%s' in '%s' — skipping", value, selector)

    async def _set_d365_optionset_via_xrm(self, logical_name: str, label: str) -> bool:
        """Set an option set value using Xrm.Page.getAttribute().setValue().

        Looks up the integer value from _OPTIONSET_VALUE_MAP first, then
        falls back to querying the attribute's options metadata.
        """
        # Try known mapping first
        int_value = None
        field_map = self._OPTIONSET_VALUE_MAP.get(logical_name)
        if field_map:
            int_value = field_map.get(label.lower().strip())

        if int_value is not None:
            js = f"""() => {{
                try {{
                    if (typeof Xrm === 'undefined') return false;
                    const attr = Xrm.Page.getAttribute('{logical_name}');
                    if (!attr) return false;
                    attr.setValue({int_value});
                    attr.fireOnChange();
                    return true;
                }} catch (e) {{ return false; }}
            }}"""
            try:
                ok = await self.browser.page.evaluate(js)
                if ok:
                    logger.info("[OPTIONSET XRM] Set '%s' to %d ('%s')",
                                logical_name, int_value, label)
                    return True
            except Exception as exc:
                logger.debug("[OPTIONSET XRM] JS failed: %s", exc)

        # Fallback: query attribute options from Xrm to find the value by label
        safe_label = label.replace("'", "\\'")
        js_search = f"""() => {{
            try {{
                if (typeof Xrm === 'undefined') return {{ ok: false }};
                const attr = Xrm.Page.getAttribute('{logical_name}');
                if (!attr) return {{ ok: false }};
                const opts = attr.getOptions ? attr.getOptions() : [];
                const match = opts.find(o =>
                    (o.text || '').toLowerCase().trim() === '{safe_label}'.toLowerCase().trim()
                );
                if (match) {{
                    attr.setValue(match.value);
                    attr.fireOnChange();
                    return {{ ok: true, value: match.value, text: match.text }};
                }}
                return {{ ok: false, available: opts.map(o => o.text) }};
            }} catch (e) {{ return {{ ok: false, error: e.message }}; }}
        }}"""
        try:
            result = await self.browser.page.evaluate(js_search)
            if result and result.get("ok"):
                logger.info("[OPTIONSET XRM] Set '%s' to %s via option search",
                            logical_name, result.get("value"))
                return True
            else:
                avail = result.get("available", []) if result else []
                logger.warning("[OPTIONSET XRM] '%s' not found in '%s'. Available: %s",
                               label, logical_name, avail)
        except Exception as exc:
            logger.debug("[OPTIONSET XRM] Search failed: %s", exc)

        return False

    async def _try_hover(self, selector: str, timeout: int = 3000):
        """Hover over an element with retry logic for D365 async rendering."""
        for attempt in range(2):
            try:
                await self.browser.page.wait_for_selector(
                    selector, state="visible", timeout=timeout
                )
                await self.browser.hover(selector)
                return
            except Exception:
                if attempt < 1:
                    logger.debug(
                        "Hover attempt %d failed for %s — retrying",
                        attempt + 1, selector,
                    )
                    await asyncio.sleep(0.5)
                else:
                    logger.warning("Hover failed after 2 attempts for %s — skipping", selector)

    # ---- Smart Click / Fill Helpers ----

    async def _try_click(self, selector: str, timeout: int = 3000) -> bool:
        """Try to click a selector, return True on success."""
        try:
            await self.browser.page.wait_for_selector(selector, state="visible", timeout=timeout)
            await self.overlay.click_ripple_on(selector)
            await asyncio.sleep(0.3)
            await self.browser.click(selector)
            return True
        except Exception:
            return False

    async def _try_fill(self, selector: str, value: str, timeout: int = 3000):
        """Try to fill a field with D365 CE smart resolution fallbacks.

        Resolution order:
        1. Exact selector
        2. Alternative selector pattern (aria-label ↔ data-id)
        3. D365 CE DOM introspection — find the real <input> by logical name
           • Lookup fields get special treatment (type → pick from dropdown)
           • Date/Duration fields get click → clear → type → Tab
        4. Broader DOM scan with retry (waits for async form rendering)
        """
        # ---- Attempt 1: exact selector ----
        try:
            el = await self.browser.page.wait_for_selector(
                selector, state="visible", timeout=timeout,
            )
            # Check if this is a lookup — handle specially
            data_id = await el.get_attribute("data-id") or ""
            if "LookupResultsDropdown" in data_id:
                await self._fill_d365_lookup(selector, value)
                return
            # Check if this is a date/duration — click+clear+type+Tab
            if "date-time-input" in data_id or "duration-combobox" in data_id:
                await self._fill_d365_date_or_duration(el, value)
                return
            await self.browser.type_slowly(selector, value, delay=60)
            return
        except Exception:
            pass

        # ---- Attempt 2: alternative selector pattern ----
        alt = self._alternative_selector(selector)
        if alt:
            try:
                await self.browser.page.wait_for_selector(
                    alt, state="visible", timeout=2000,
                )
                await self.browser.type_slowly(alt, value, delay=60)
                return
            except Exception:
                pass

        # ---- Attempt 3: D365 CE smart field resolution (fast JS scan) ----
        logger.info("Resolving D365 CE field for failed selector: %s", selector)
        field_info = await self._find_d365_ce_field(selector)
        if field_info and field_info.get("selector"):
            resolved = field_info["selector"]
            logger.info("D365 CE resolved → %s (lookup=%s, date=%s)",
                        resolved, field_info.get("isLookup"), field_info.get("isDate"))
            try:
                if field_info.get("isLookup"):
                    await self._fill_d365_lookup(resolved, value)
                    return

                el = await self.browser.page.wait_for_selector(
                    resolved, state="visible", timeout=2000,
                )
                if field_info.get("isDate") or field_info.get("isDuration"):
                    await self._fill_d365_date_or_duration(el, value)
                else:
                    await self.browser.type_slowly(resolved, value, delay=60)
                return
            except Exception:
                pass

        # ---- Attempt 4: Wait for async form render then retry DOM scan ----
        # D365 Time Entry and certain PCF controls render fields late
        logger.info("Field not found yet — waiting for async rendering and retrying…")
        await asyncio.sleep(2.0)
        field_info = await self._find_d365_ce_field(selector)
        if field_info and field_info.get("selector"):
            resolved = field_info["selector"]
            logger.info("D365 CE resolved (retry) → %s", resolved)
            try:
                if field_info.get("isLookup"):
                    await self._fill_d365_lookup(resolved, value)
                    return
                el = await self.browser.page.wait_for_selector(
                    resolved, state="visible", timeout=2000,
                )
                if field_info.get("isDate") or field_info.get("isDuration"):
                    await self._fill_d365_date_or_duration(el, value)
                else:
                    await self.browser.type_slowly(resolved, value, delay=60)
                return
            except Exception:
                pass

        # ---- Attempt 5: Broader CSS scan by logical name fragment ----
        m = re.search(r'data-id=["\']([^."\']+)\.', selector)
        logical_name = m.group(1) if m else None
        if logical_name:
            is_known_lookup = logical_name.lower() in self._KNOWN_LOOKUP_FIELDS
            broad_selectors = [
                f'input[data-id*="{logical_name}"]',
                f'textarea[data-id*="{logical_name}"]',
                f'select[data-id*="{logical_name}"]',
                f'div[data-id*="{logical_name}"] input',
                f'div[data-id*="{logical_name}"] textarea',
                f'section[data-id*="{logical_name}"] input',
            ]
            is_duration_field = "duration" in logical_name.lower()
            is_date_field = "date" in logical_name.lower() and "update" not in logical_name.lower()

            # For duration fields, try Xrm SDK first — no DOM needed
            if is_duration_field:
                stripped = value.strip()
                hours: float | None = None
                h_match = re.match(r'^(\d+(?:\.\d+)?)\s*h', stripped, re.IGNORECASE)
                if h_match:
                    hours = float(h_match.group(1))
                    m_match = re.search(r'(\d+)\s*m', stripped, re.IGNORECASE)
                    if m_match:
                        hours += float(m_match.group(1)) / 60.0
                elif re.match(r'^\d+(\.\d+)?$', stripped):
                    hours = float(stripped)
                if hours is not None:
                    minutes = int(round(hours * 60))
                    ok = await self._set_d365_duration_via_xrm(logical_name, minutes)
                    if ok:
                        return

            for bsel in broad_selectors:
                try:
                    el = await self.browser.page.wait_for_selector(
                        bsel, state="visible", timeout=1500,
                    )
                    data_id = await el.get_attribute("data-id") or ""
                    if "LookupResultsDropdown" in data_id or is_known_lookup:
                        await self._fill_d365_lookup(bsel, value)
                        return
                    if ("date-time-input" in data_id or "duration-combobox" in data_id
                            or is_duration_field or is_date_field):
                        await self._fill_d365_date_or_duration(
                            el, value, logical_name_hint=logical_name)
                        return
                    await el.click()
                    await asyncio.sleep(0.2)
                    await el.fill("")
                    await asyncio.sleep(0.1)
                    await el.type(value, delay=60)
                    await self.browser.page.keyboard.press("Tab")
                    logger.info("Filled via broad scan: %s → %s", bsel, data_id)
                    return
                except Exception:
                    continue

        raise Exception(f"Could not find field: {selector}")

    async def _fill_d365_date_or_duration(self, el, value: str, *, logical_name_hint: str | None = None):
        """Fill a D365 CE date or duration field.

        For **duration** fields, the preferred strategy is to use the Xrm SDK
        to set the value directly in minutes, completely bypassing the D365
        duration combobox which misinterprets character-by-character typing
        (e.g. "8h 0m" → 80 minutes → 1.33 hours).

        For **date** fields, the DOM approach (triple-click → type → Tab) is
        kept as-is because Xrm date attributes accept JS Date objects that
        would need timezone handling.
        """
        data_id = (await el.get_attribute("data-id") or "").lower()
        aria_label = (await el.get_attribute("aria-label") or "").lower()
        hint_lower = (logical_name_hint or "").lower()
        is_duration = ("duration" in data_id or "duration" in aria_label
                       or "duration" in hint_lower)

        if is_duration:
            stripped = value.strip()
            hours: float | None = None

            # Parse "8h 0m", "8h", "8.5h 30m", etc.
            h_match = re.match(r'^(\d+(?:\.\d+)?)\s*h', stripped, re.IGNORECASE)
            if h_match:
                hours = float(h_match.group(1))
                # Also add extra minutes if present (e.g. "1h 30m")
                m_match = re.search(r'(\d+)\s*m', stripped, re.IGNORECASE)
                if m_match:
                    hours += float(m_match.group(1)) / 60.0
            # Bare number → treat as hours
            elif re.match(r'^\d+(\.\d+)?$', stripped):
                hours = float(stripped)

            if hours is not None:
                minutes = int(round(hours * 60))
                # Extract logical name from data-id, fall back to caller hint
                logical_name = None
                m_dn = re.search(r'^([^.]+)\.', data_id)
                if m_dn:
                    logical_name = m_dn.group(1)
                if not logical_name and logical_name_hint:
                    logical_name = logical_name_hint
                if logical_name:
                    ok = await self._set_d365_duration_via_xrm(logical_name, minutes)
                    if ok:
                        return
                # Xrm SDK unavailable — fall through to DOM approach
                logger.info("Xrm SDK duration set failed — falling back to DOM typing")

        # ---- DOM fallback (dates and last-resort duration) ----
        await el.click()
        await asyncio.sleep(0.3)
        # Triple-click to select all existing text
        await el.click(click_count=3)
        await asyncio.sleep(0.2)
        await el.type(value, delay=60)
        await asyncio.sleep(0.2)
        await self.browser.page.keyboard.press("Tab")
        await asyncio.sleep(1.0)  # Give D365 time to commit the value
        logger.info("Filled date/duration field with: %s", value)

    async def _set_d365_duration_via_xrm(self, logical_name: str, minutes: int) -> bool:
        """Set a duration field via Xrm.Page.getAttribute().setValue().

        Duration fields in D365 store values in **minutes** as an integer.
        This bypasses the unreliable duration combobox entirely.
        """
        js = f"""() => {{
            try {{
                if (typeof Xrm === 'undefined') return false;
                const attr = Xrm.Page.getAttribute('{logical_name}');
                if (!attr) return false;
                attr.setValue({minutes});
                attr.fireOnChange();
                return true;
            }} catch (e) {{ return false; }}
        }}"""
        try:
            ok = await self.browser.page.evaluate(js)
            if ok:
                logger.info(
                    "[DURATION XRM] Set '%s' to %d minutes (%.1f hours)",
                    logical_name, minutes, minutes / 60,
                )
                return True
        except Exception as exc:
            logger.debug("[DURATION XRM] Failed: %s", exc)
        return False

    # ---- D365 CE Field Resolution ----

    # JavaScript that searches the live DOM for a D365 CE input by logical name / label.
    _D365_FIND_FIELD_JS = """(hint) => {
        const ln = (hint.logicalName || '').toLowerCase();
        const lbl = (hint.label || '').toLowerCase();

        function info(el) {
            const did = el.getAttribute('data-id') || '';
            const aid = (el.getAttribute('aria-label') || '').toLowerCase();
            return {
                selector: did ? '[data-id=\"' + did + '\"]' : null,
                dataId: did,
                isLookup: did.includes('LookupResultsDropdown') || did.includes('lookup'),
                isDate: did.includes('date-time-input') || aid.includes('date'),
                isDuration: did.includes('duration-combobox') || aid.includes('duration'),
            };
        }
        function visible(el) {
            if (!el) return false;
            // Check offsetParent first (fast) — null means hidden, except for <body>/<html>
            if (el.offsetParent !== null) return true;
            // Some D365 controls have offsetParent=null but are actually visible
            const rect = el.getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        }

        // 1. By data-id prefix  (msdyn_date.fieldControl-…)
        if (ln) {
            const prefix = ln + '.fieldControl';
            for (const el of document.querySelectorAll(
                'input[data-id^=\"' + prefix + '\"], ' +
                'textarea[data-id^=\"' + prefix + '\"], ' +
                'select[data-id^=\"' + prefix + '\"]'
            )) { if (visible(el)) return info(el); }

            // Broader: data-id contains logicalName anywhere
            for (const el of document.querySelectorAll(
                'input[data-id*=\"' + ln + '\"], ' +
                'textarea[data-id*=\"' + ln + '\"], ' +
                'select[data-id*=\"' + ln + '\"]'
            )) { if (visible(el)) return info(el); }

            // Search inside D365 sections/divs that wrap controls
            for (const el of document.querySelectorAll(
                'div[data-id*=\"' + ln + '\"] input, ' +
                'div[data-id*=\"' + ln + '\"] textarea, ' +
                'div[data-id*=\"' + ln + '\"] select, ' +
                'section[data-id*=\"' + ln + '\"] input, ' +
                'section[data-id*=\"' + ln + '\"] textarea'
            )) { if (visible(el)) return info(el); }
        }

        // 2. By aria-label containing the label text
        if (lbl) {
            for (const el of document.querySelectorAll('input[aria-label], textarea[aria-label], select[aria-label]')) {
                if (!visible(el)) continue;
                const al = (el.getAttribute('aria-label') || '').toLowerCase();
                if (al === lbl || al.includes(lbl) || lbl.includes(al)) return info(el);
            }
        }

        // 3. Walk all visible inputs on the form body — match by logical name fragment
        if (ln) {
            const nm = ln.replace('msdyn_', '');
            const containers = [
                document.querySelector('[data-id="form-body"]'),
                document.querySelector('[data-id="mainContent"]'),
                document.querySelector('#mainContent'),
                document.body,
            ];
            for (const container of containers) {
                if (!container) continue;
                for (const el of container.querySelectorAll('input, textarea, select')) {
                    if (!visible(el)) continue;
                    const did = (el.getAttribute('data-id') || '').toLowerCase();
                    const al = (el.getAttribute('aria-label') || '').toLowerCase();
                    const nm2 = (el.getAttribute('name') || '').toLowerCase();
                    if (did.includes(nm) || al.includes(nm) || nm2.includes(nm)) return info(el);
                }
                // If we found a container but no match, continue to broader containers
            }
        }

        return null;
    }"""

    # Common human-readable labels → D365 logical names (Time Entry & Expense).
    _D365_LABEL_MAP: dict[str, str] = {
        "date": "msdyn_date",
        "project": "msdyn_project",
        "task": "msdyn_projecttask",
        "project task": "msdyn_projecttask",
        "duration": "msdyn_duration",
        "description": "msdyn_description",
        "internal description": "msdyn_description",
        "external description": "msdyn_externaldescription",
        "role": "msdyn_resourcecategory",
        "resource category": "msdyn_resourcecategory",
        "type": "msdyn_type",
        "entry status": "msdyn_entrystatus",
        "amount": "msdyn_amount",
        "price": "msdyn_price",
        "quantity": "msdyn_quantity",
        "expense category": "msdyn_expensecategory",
        "transaction date": "msdyn_transactiondate",
        "name": "msdyn_name",
        "subject": "msdyn_subject",
        "customer": "msdyn_customer",
        "start": "msdyn_start",
        "end": "msdyn_finish",
        "start date": "msdyn_start",
        "end date": "msdyn_finish",
        "bookable resource": "msdyn_bookableresource",
        "resource": "msdyn_bookableresource",
    }

    async def _find_d365_ce_field(self, failed_selector: str) -> dict | None:
        """Search the live D365 CE DOM for the real <input> matching *failed_selector*.

        Returns ``{selector, dataId, isLookup, isDate, isDuration}`` or ``None``.
        """
        logical_name: str | None = None
        label: str | None = None

        # Pull logical name from data-id="msdyn_xxx.fieldControl-…"
        m = re.search(r'data-id=["\']([^."\']+)\.', failed_selector)
        if m:
            logical_name = m.group(1)

        # Pull label from aria-label="…"
        m = re.search(r'aria-label=["\']([^"\']+)["\']', failed_selector)
        if m:
            label = m.group(1).strip()
            if not logical_name:
                logical_name = self._D365_LABEL_MAP.get(label.lower())

        if not logical_name and not label:
            return None

        try:
            return await self.browser.page.evaluate(
                self._D365_FIND_FIELD_JS,
                {"logicalName": logical_name or "", "label": label or ""},
            )
        except Exception as exc:
            logger.debug("D365 CE field resolution JS failed: %s", exc)
            return None

    async def _fill_d365_lookup(self, selector: str, value: str):
        """Fill a D365 CE lookup field: click → type → wait → pick first result → verify.

        D365 CE lookups are not standard <input> — after typing a search string
        the system shows a dropdown of matching records that must be clicked.
        After selection, waits for D365 to commit the value to the form.
        Falls back to Xrm SDK to search and set the value programmatically.

        For **project** lookups, the value is automatically replaced with a
        project the current user is actually assigned to (team member of),
        preventing D365 validation errors on save.
        """
        logger.info("[LOOKUP] Filling '%s' into %s", value, selector)

        # Extract logical name for Xrm fallback
        logical_name = None
        m = re.search(r'data-id=["\']([^."\']+)\.', selector)
        if m:
            logical_name = m.group(1)
        if not logical_name:
            m = re.search(r'data-id\*="([^"]+)"', selector)
            if m:
                logical_name = m.group(1)

        # ---- For project lookups, use preflight data or query live ----
        if logical_name and logical_name.lower() == "msdyn_project":
            # If we have the exact projectId from preflight, set it directly via Xrm
            pf_project_id = self._preflight_data.get("projectId")
            pf_project = self._preflight_data.get("projectName")
            if pf_project_id and pf_project:
                logger.info(
                    "[LOOKUP] Setting project directly via Xrm: '%s' (ID: %s)",
                    pf_project, pf_project_id,
                )
                ok = await self._set_d365_lookup_by_id(
                    "msdyn_project", "msdyn_project", pf_project_id, pf_project
                )
                if ok:
                    return
                logger.warning("[LOOKUP] Direct project set failed — falling back to search")
                value = pf_project
            elif pf_project:
                logger.info(
                    "[LOOKUP] Using preflight project '%s' (replacing '%s')",
                    pf_project, value,
                )
                value = pf_project
            else:
                assigned_name = await self._get_user_assigned_project_name()
                if assigned_name:
                    logger.info(
                        "[LOOKUP] Replacing generic project search '%s' with "
                        "user-assigned project '%s'", value, assigned_name,
                    )
                    value = assigned_name

        # ---- For project task lookups, use preflight data or query live ----
        if logical_name and logical_name.lower() == "msdyn_projecttask":
            # If we have the exact taskId from preflight, set it directly via Xrm
            pf_task_id = self._preflight_data.get("taskId")
            pf_task = self._preflight_data.get("taskName")
            if pf_task_id and pf_task:
                logger.info(
                    "[LOOKUP] Setting task directly via Xrm: '%s' (ID: %s)",
                    pf_task, pf_task_id,
                )
                ok = await self._set_d365_lookup_by_id(
                    "msdyn_projecttask", "msdyn_projecttask", pf_task_id, pf_task
                )
                if ok:
                    return
                logger.warning("[LOOKUP] Direct task set failed — falling back to search")
                value = pf_task
            elif pf_task:
                logger.info(
                    "[LOOKUP] Using preflight task '%s' (replacing '%s')",
                    pf_task, value,
                )
                value = pf_task
            else:
                assigned_task = await self._get_user_assigned_task_name()
                if assigned_task:
                    logger.info(
                        "[LOOKUP] Replacing generic task search '%s' with "
                        "user-assigned task '%s'", value, assigned_task,
                    )
                    value = assigned_task

        # Find the lookup input element — try multiple approaches
        el = None
        try:
            el = await self.browser.page.wait_for_selector(
                selector, state="visible", timeout=3000,
            )
        except Exception:
            # Try broader selectors if exact one fails
            if logical_name:
                for alt_sel in [
                    f'div[data-id*="{logical_name}"] input',
                    f'input[data-id*="{logical_name}"]',
                ]:
                    try:
                        el = await self.browser.page.wait_for_selector(
                            alt_sel, state="visible", timeout=2000,
                        )
                        logger.info("[LOOKUP] Found via alt selector: %s", alt_sel)
                        break
                    except Exception:
                        continue

        if not el:
            logger.warning("[LOOKUP] Element not found — trying Xrm SDK fallback")
            if logical_name:
                await self._set_d365_lookup_via_xrm(logical_name, value)
            return

        await el.click()
        await asyncio.sleep(0.5)
        # Clear any existing value
        await el.fill("")
        await asyncio.sleep(0.3)
        # Type search text slowly for demo effect
        await el.type(value, delay=80)
        # D365 lookup search is async — wait longer for results
        await asyncio.sleep(4.0)

        # Try to click the first dropdown result
        selected = False
        result_selectors = [
            'ul[aria-label*="Lookup results"] li:first-child',
            'ul[aria-label*="Lookup Results"] li:first-child',
            'div[data-id*="LookupResultsDropdown"] li[role="option"]:first-child',
            'div[data-id*="LookupResultsDropdown"] div[role="option"]:first-child',
            'div[role="presentation"] ul[role="listbox"] li:first-child',
            'ul[role="listbox"] li[role="option"]:first-child',
            'div[role="listbox"] div[role="option"]:first-child',
            'div.ms-BasePicker-text ~ div li:first-child',
            # Flyout / popup result containers
            'div[id*="lookupDialogLookup"] li:first-child',
            'div[class*="lookupFlyout"] li:first-child',
            # Additional D365 UCI patterns
            'li[data-id*="LookupResultsDropdown"][role="option"]:first-child',
            'div[data-id*="resultsContainer"] li:first-child',
            'div[data-id*="lookupDialogLookup_MscrmControls"] li:first-child',
        ]
        for rsel in result_selectors:
            try:
                result_el = await self.browser.page.wait_for_selector(
                    rsel, state="visible", timeout=1500,
                )
                await result_el.click()
                logger.info("[LOOKUP] Selected '%s' via: %s", value, rsel)
                selected = True
                break
            except Exception:
                continue

        # JavaScript fallback — scan the DOM for any visible dropdown option
        if not selected:
            logger.info("[LOOKUP] CSS selectors missed — trying JS click fallback")
            js_clicked = await self.browser.page.evaluate("""() => {
                // Find any visible option/listitem inside lookup dropdowns
                const candidates = document.querySelectorAll(
                    '[role="option"], [role="listbox"] li, ' +
                    '[data-id*="LookupResultsDropdown"] li, ' +
                    '[data-id*="lookupDialogLookup"] li'
                );
                for (const el of candidates) {
                    const rect = el.getBoundingClientRect();
                    if (rect.width > 0 && rect.height > 0 && rect.top >= 0) {
                        el.click();
                        return el.textContent?.trim() || true;
                    }
                }
                return false;
            }""")
            if js_clicked:
                logger.info("[LOOKUP] JS click fallback selected: %s", js_clicked)
                selected = True

        if not selected:
            # Keyboard fallback — arrow-down into the first result and press Enter
            logger.info("[LOOKUP] No dropdown result found — using keyboard fallback")
            await self.browser.page.keyboard.press("ArrowDown")
            await asyncio.sleep(0.5)
            await self.browser.page.keyboard.press("Enter")
            await asyncio.sleep(1.0)

        # Wait for D365 to commit the lookup value to the form
        await asyncio.sleep(1.5)

        # Tab out to force D365 to finalize the lookup binding
        await self.browser.page.keyboard.press("Tab")
        await asyncio.sleep(1.0)

        # Verify the lookup resolved — check if the field shows a resolved tag
        lookup_resolved = await self.browser.page.evaluate("""(sel) => {
            // D365 lookups show a resolved tag with the record name
            const container = document.querySelector(sel)?.closest('[data-id]');
            if (!container) return null;
            // Look for the resolved record pill/tag
            const tag = container.querySelector(
                'span[data-id*="selected_tag_text"], ' +
                'li[data-id*="selected_tag"], ' +
                'span.mectrl_headertext, ' +
                'div[data-id*="LookupResultsDropdown"] span'
            );
            return tag ? tag.textContent?.trim() : null;
        }""", selector)

        if lookup_resolved:
            logger.info("[LOOKUP] Verified — resolved to: '%s'", lookup_resolved)
        else:
            logger.warning("[LOOKUP] DOM verification failed — trying Xrm SDK fallback")
            # Try Xrm SDK as final fallback before tracking for MCP
            if logical_name:
                xrm_ok = await self._set_d365_lookup_via_xrm(logical_name, value)
                if xrm_ok:
                    return
            self._pending_field_updates.append({
                "selector": selector,
                "value": value,
                "field_type": "lookup",
                "resolved": False,
            })

    async def _set_d365_lookup_by_id(
        self, logical_name: str, entity_name: str, record_id: str, record_name: str
    ) -> bool:
        """Set a lookup field directly using a known record ID via Xrm.Page.

        This is the most reliable approach when we already know the exact
        record (e.g. from preflight verification). No search required.
        """
        safe_name = record_name.replace('\\', '\\\\').replace("'", "\\'")
        js = f"""async () => {{
            try {{
                if (typeof Xrm === 'undefined') return {{ ok: false, reason: 'no_xrm' }};
                const formCtx = Xrm.Page || (Xrm.Utility && Xrm.Utility.getPageContext
                    && Xrm.Utility.getPageContext().input);
                if (!formCtx) return {{ ok: false, reason: 'no_form_context' }};
                const attr = formCtx.getAttribute && formCtx.getAttribute('{logical_name}');
                if (!attr) return {{ ok: false, reason: 'no_attribute' }};
                attr.setValue([{{
                    id: '{record_id}',
                    name: '{safe_name}',
                    entityType: '{entity_name}'
                }}]);
                attr.fireOnChange();
                return {{ ok: true }};
            }} catch (e) {{
                return {{ ok: false, reason: e.message || String(e) }};
            }}
        }}"""
        try:
            result = await self.browser.page.evaluate(js)
            if result and result.get("ok"):
                logger.info("[LOOKUP BY ID] Set '%s' to '%s' (ID: %s)",
                            logical_name, record_name, record_id)
                await asyncio.sleep(1.5)  # Let D365 react to the change
                return True
            reason = result.get("reason", "unknown") if result else "null result"
            logger.warning("[LOOKUP BY ID] Failed for '%s': %s", logical_name, reason)
            return False
        except Exception as exc:
            logger.warning("[LOOKUP BY ID] Exception for '%s': %s", logical_name, exc)
            return False

    async def _set_d365_lookup_via_xrm(self, logical_name: str, search_value: str) -> bool:
        """Use Xrm.WebApi to search for a record and Xrm.Page to set the lookup value.

        This bypasses DOM interaction entirely and is the most reliable way
        to set lookup values on D365 UCI forms.
        """
        # Map logical names to their target entity and name field
        lookup_entity_map: dict[str, tuple[str, str]] = {
            "msdyn_project": ("msdyn_project", "msdyn_subject"),
            "msdyn_projecttask": ("msdyn_projecttask", "msdyn_subject"),
            "msdyn_bookableresource": ("bookableresource", "name"),
            "msdyn_bookableresourceid": ("bookableresource", "name"),
            "msdyn_resourcecategory": ("bookableresourcecategory", "name"),
            "msdyn_expensecategory": ("msdyn_expensecategory", "msdyn_name"),
            "msdyn_customer": ("account", "name"),
            "msdyn_contractorganizationalunitid": ("msdyn_organizationalunit", "msdyn_name"),
            "msdyn_organizationalunit": ("msdyn_organizationalunit", "msdyn_name"),
            "bookableresourcecategory": ("bookableresourcecategory", "name"),
            "msdyn_transactioncategory": ("msdyn_transactioncategory", "msdyn_name"),
            "msdyn_projectmanager": ("systemuser", "fullname"),
        }

        target = lookup_entity_map.get(logical_name)
        if not target:
            logger.warning("[LOOKUP XRM] No entity mapping for: %s", logical_name)
            return False

        entity_name, name_field = target
        # Escape for OData single-quoted string values
        safe_value = search_value.replace("'", "''")
        # Build the OData filter as a plain Python string (no JS escaping needed)
        odata_filter = f"?$filter=contains({name_field},'{safe_value}')"

        # For project tasks, scope query to the current project
        if logical_name == "msdyn_projecttask":
            pf_project_id = self._preflight_data.get("projectId")
            if pf_project_id:
                odata_filter += f" and _msdyn_project_value eq {pf_project_id}"

        odata_filter += "&$top=1"
        # Escape for embedding in a JS double-quoted string
        odata_js = odata_filter.replace('\\', '\\\\').replace('"', '\\"')
        safe_search_js = search_value.replace('\\', '\\\\').replace("'", "\\'")

        js = f"""async () => {{
            try {{
                if (typeof Xrm === 'undefined') return {{ ok: false, reason: 'no_xrm' }};
                const result = await Xrm.WebApi.retrieveMultipleRecords(
                    "{entity_name}",
                    "{odata_js}"
                );
                if (!result.entities || result.entities.length === 0) {{
                    return {{ ok: false, reason: 'no_records' }};
                }}
                const record = result.entities[0];
                const recordId = record['{entity_name}id'];
                const recordName = record['{name_field}'] || '{safe_search_js}';
                if (!recordId) return {{ ok: false, reason: 'no_id' }};
                const formCtx = Xrm.Page || (Xrm.Utility && Xrm.Utility.getPageContext && Xrm.Utility.getPageContext().input);
                if (!formCtx) return {{ ok: false, reason: 'no_form_context' }};
                const attr = formCtx.getAttribute && formCtx.getAttribute('{logical_name}');
                if (!attr) return {{ ok: false, reason: 'no_attribute' }};
                attr.setValue([{{
                    id: recordId,
                    name: recordName,
                    entityType: '{entity_name}'
                }}]);
                attr.fireOnChange();
                return {{ ok: true, name: recordName, id: recordId }};
            }} catch (e) {{
                return {{ ok: false, reason: e.message || String(e) }};
            }}
        }}"""

        try:
            result = await self.browser.page.evaluate(js)
            if result and result.get("ok"):
                logger.info("[LOOKUP XRM] Set '%s' to '%s' (ID: %s)",
                            logical_name, result.get("name"), result.get("id"))
                await asyncio.sleep(1.0)  # Let D365 react to the change
                return True
            else:
                reason = result.get("reason", "unknown") if result else "null result"
                logger.warning("[LOOKUP XRM] Failed for '%s': %s", logical_name, reason)
                return False
        except Exception as exc:
            logger.warning("[LOOKUP XRM] Exception for '%s': %s", logical_name, exc)
            return False

    async def _get_user_assigned_project_name(self) -> str | None:
        """Query D365 for an active project the current user is a team member of.

        Uses Xrm.WebApi to:
        1. Find the current user's bookable resource ID
        2. Query ``msdyn_projectteam`` for that resource's team memberships
        3. Retrieve the project name for the first active project found

        Returns the project ``msdyn_subject`` (name) or ``None``.
        """
        js = """async () => {
            try {
                if (typeof Xrm === 'undefined') return { ok: false, reason: 'no_xrm' };
                const userId = Xrm.Utility.getGlobalContext()
                    .userSettings.userId.replace(/[{}]/g, '');

                // 1. Find the bookable resource for the current user
                const resources = await Xrm.WebApi.retrieveMultipleRecords(
                    'bookableresource',
                    '?$filter=_userid_value eq ' + userId +
                    '&$select=bookableresourceid,name&$top=1'
                );
                if (!resources.entities || resources.entities.length === 0)
                    return { ok: false, reason: 'no_bookable_resource_for_user' };
                const resourceId = resources.entities[0].bookableresourceid;

                // 2. Find project team memberships for this resource
                const teams = await Xrm.WebApi.retrieveMultipleRecords(
                    'msdyn_projectteam',
                    '?$filter=_msdyn_bookableresourceid_value eq ' + resourceId +
                    '&$select=_msdyn_project_value&$top=10'
                );
                if (!teams.entities || teams.entities.length === 0)
                    return { ok: false, reason: 'no_team_memberships' };

                // 3. Check each project for Active status (statecode=0)
                for (const tm of teams.entities) {
                    const projId = tm['_msdyn_project_value'];
                    if (!projId) continue;
                    try {
                        const proj = await Xrm.WebApi.retrieveRecord(
                            'msdyn_project', projId,
                            '?$select=msdyn_subject,statecode'
                        );
                        if (proj && proj.msdyn_subject && proj.statecode === 0) {
                            return { ok: true, name: proj.msdyn_subject, id: projId };
                        }
                    } catch (e) { continue; }
                }
                return { ok: false, reason: 'no_active_projects' };
            } catch (e) {
                return { ok: false, reason: e.message || String(e) };
            }
        }"""
        try:
            result = await self.browser.page.evaluate(js)
            if result and result.get("ok"):
                name = result.get("name")
                logger.info(
                    "[LOOKUP] User is assigned to project: '%s' (ID: %s)",
                    name, result.get("id"),
                )
                return name
            reason = result.get("reason", "unknown") if result else "null"
            logger.info("[LOOKUP] Could not determine assigned project: %s", reason)
            return None
        except Exception as exc:
            logger.warning("[LOOKUP] Assigned-project query failed: %s", exc)
            return None

    async def _get_user_assigned_task_name(self) -> str | None:
        """Query D365 for a project task the current user is assigned to.

        Uses Xrm.WebApi to:
        1. Find the current user's bookable resource ID
        2. Find the project currently set on the form (if any)
        3. Query ``msdyn_resourceassignment`` for that resource's task assignments
           on the current project
        4. Retrieve the task name (``msdyn_subject``) for the first assignment found

        Returns the task ``msdyn_subject`` (name) or ``None``.
        """
        js = """async () => {
            try {
                if (typeof Xrm === 'undefined') return { ok: false, reason: 'no_xrm' };
                const userId = Xrm.Utility.getGlobalContext()
                    .userSettings.userId.replace(/[{}]/g, '');

                // 1. Find the bookable resource for the current user
                const resources = await Xrm.WebApi.retrieveMultipleRecords(
                    'bookableresource',
                    '?$filter=_userid_value eq ' + userId +
                    '&$select=bookableresourceid,name&$top=1'
                );
                if (!resources.entities || resources.entities.length === 0)
                    return { ok: false, reason: 'no_bookable_resource_for_user' };
                const resourceId = resources.entities[0].bookableresourceid;

                // 2. Get the project — from form first, then from team membership
                let projectId = null;
                try {
                    const projAttr = Xrm.Page.getAttribute('msdyn_project');
                    if (projAttr) {
                        const projVal = projAttr.getValue();
                        if (projVal && projVal.length > 0) {
                            projectId = projVal[0].id.replace(/[{}]/g, '');
                        }
                    }
                } catch (e) { /* no project on form yet */ }

                // If project not on form, find the user's assigned project
                if (!projectId) {
                    const teams = await Xrm.WebApi.retrieveMultipleRecords(
                        'msdyn_projectteam',
                        '?$filter=_msdyn_bookableresourceid_value eq ' + resourceId +
                        '&$select=_msdyn_project_value&$top=10'
                    );
                    if (teams.entities) {
                        for (const tm of teams.entities) {
                            const pId = tm['_msdyn_project_value'];
                            if (!pId) continue;
                            try {
                                const proj = await Xrm.WebApi.retrieveRecord(
                                    'msdyn_project', pId,
                                    '?$select=msdyn_subject,statecode'
                                );
                                if (proj && proj.statecode === 0) {
                                    projectId = pId;
                                    break;
                                }
                            } catch (e) { continue; }
                        }
                    }
                }

                if (!projectId)
                    return { ok: false, reason: 'no_project_for_task_scope' };

                // 3. Query resource assignments scoped to this project
                const assignments = await Xrm.WebApi.retrieveMultipleRecords(
                    'msdyn_resourceassignment',
                    '?$filter=_msdyn_bookableresourceid_value eq ' + resourceId +
                    ' and _msdyn_projectid_value eq ' + projectId +
                    '&$select=_msdyn_taskid_value&$top=10'
                );
                if (!assignments.entities || assignments.entities.length === 0)
                    return { ok: false, reason: 'no_task_assignments' };

                // 4. Get the task name — verify it belongs to the correct project
                for (const asgn of assignments.entities) {
                    const taskId = asgn['_msdyn_taskid_value'];
                    if (!taskId) continue;
                    try {
                        const task = await Xrm.WebApi.retrieveRecord(
                            'msdyn_projecttask', taskId,
                            '?$select=msdyn_subject,_msdyn_project_value'
                        );
                        if (task && task.msdyn_subject
                            && task['_msdyn_project_value']
                            && task['_msdyn_project_value'].toLowerCase() === projectId.toLowerCase()) {
                            return { ok: true, name: task.msdyn_subject, id: taskId };
                        }
                    } catch (e) { continue; }
                }
                return { ok: false, reason: 'no_tasks_found' };
            } catch (e) {
                return { ok: false, reason: e.message || String(e) };
            }
        }"""
        try:
            result = await self.browser.page.evaluate(js)
            if result and result.get("ok"):
                name = result.get("name")
                logger.info(
                    "[LOOKUP] User is assigned to task: '%s' (ID: %s)",
                    name, result.get("id"),
                )
                return name
            reason = result.get("reason", "unknown") if result else "null"
            logger.info("[LOOKUP] Could not determine assigned task: %s", reason)
            return None
        except Exception as exc:
            logger.warning("[LOOKUP] Assigned-task query failed: %s", exc)
            return None

    async def preflight_check(self) -> dict:
        """Run a pre-flight data check via Xrm.WebApi before starting a demo.

        Verifies:
        - The current user has a bookable resource record
        - The user is a team member on at least one active project
        - The project has at least one task with a resource assignment

        Returns a dict with 'ok' (bool), 'warnings' (list of strings),
        and 'data' (dict with discovered project/task names for the planner).
        """
        js = """async () => {
            const result = { ok: true, warnings: [], data: {} };
            try {
                if (typeof Xrm === 'undefined') {
                    result.ok = false;
                    result.warnings.push('Xrm SDK not available — cannot verify data');
                    return result;
                }
                const userId = Xrm.Utility.getGlobalContext()
                    .userSettings.userId.replace(/[{}]/g, '');
                result.data.userId = userId;

                // 1. Check bookable resource
                const resources = await Xrm.WebApi.retrieveMultipleRecords(
                    'bookableresource',
                    '?$filter=_userid_value eq ' + userId +
                    '&$select=bookableresourceid,name&$top=1'
                );
                if (!resources.entities || resources.entities.length === 0) {
                    result.ok = false;
                    result.warnings.push('No bookable resource found for the current user. Create one in Resource Management → Resources.');
                    return result;
                }
                const resourceId = resources.entities[0].bookableresourceid;
                result.data.resourceName = resources.entities[0].name;

                // 2. Check project team memberships
                const teams = await Xrm.WebApi.retrieveMultipleRecords(
                    'msdyn_projectteam',
                    '?$filter=_msdyn_bookableresourceid_value eq ' + resourceId +
                    '&$select=_msdyn_project_value&$top=10'
                );
                if (!teams.entities || teams.entities.length === 0) {
                    result.ok = false;
                    result.warnings.push('User is not a team member on any project. Add the user to a project team first.');
                    return result;
                }

                // 3+4. Find an active project that HAS task assignments for this user.
                //       Iterate ALL team projects; only select one with verified assignments.
                let projectId = null;
                let projectName = null;
                result.data.resourceId = resourceId;

                for (const tm of teams.entities) {
                    const pId = tm['_msdyn_project_value'];
                    if (!pId) continue;
                    try {
                        const proj = await Xrm.WebApi.retrieveRecord(
                            'msdyn_project', pId,
                            '?$select=msdyn_subject,statecode'
                        );
                        if (!proj || proj.statecode !== 0) continue;

                        // Check assignments on this specific project
                        const assignments = await Xrm.WebApi.retrieveMultipleRecords(
                            'msdyn_resourceassignment',
                            '?$filter=_msdyn_bookableresourceid_value eq ' + resourceId +
                            ' and _msdyn_projectid_value eq ' + pId +
                            '&$select=_msdyn_taskid_value&$top=5'
                        );
                        if (!assignments.entities || assignments.entities.length === 0) continue;

                        // Verify at least one assignment maps to a valid task on this project
                        for (const asgn of assignments.entities) {
                            const taskId = asgn['_msdyn_taskid_value'];
                            if (!taskId) continue;
                            try {
                                const task = await Xrm.WebApi.retrieveRecord(
                                    'msdyn_projecttask', taskId,
                                    '?$select=msdyn_subject,_msdyn_project_value'
                                );
                                if (task && task['_msdyn_project_value'] &&
                                    task['_msdyn_project_value'].toLowerCase() === pId.toLowerCase()) {
                                    projectId = pId;
                                    projectName = proj.msdyn_subject;
                                    result.data.taskName = task.msdyn_subject;
                                    result.data.taskId = taskId;
                                    break;
                                }
                            } catch (e) { continue; }
                        }
                        if (projectId) break;
                    } catch (e) { continue; }
                }

                if (!projectId) {
                    result.ok = false;
                    result.warnings.push(
                        'No active project with task assignments found for the user. ' +
                        'Will create sample data.'
                    );
                    return result;
                }
                result.data.projectName = projectName;
                result.data.projectId = projectId;

                return result;
            } catch (e) {
                result.ok = false;
                result.warnings.push('Preflight check error: ' + (e.message || String(e)));
                return result;
            }
        }"""
        try:
            result = await self.browser.page.evaluate(js)
            if not result:
                return {"ok": False, "warnings": ["Preflight check returned no result"], "data": {}}
            logger.info("[PREFLIGHT] ok=%s, warnings=%d, data=%s",
                        result.get("ok"), len(result.get("warnings", [])),
                        {k: v for k, v in result.get("data", {}).items() if k != "userId"})
            return result
        except Exception as exc:
            logger.warning("[PREFLIGHT] Exception: %s", exc)
            return {"ok": False, "warnings": [f"Preflight check failed: {exc}"], "data": {}}

    async def create_sample_data(self, preflight_result: dict) -> dict:
        """Create missing D365 sample data so the demo can proceed.

        Uses the Dataverse Web API (via ``_post_record``) to provision
        whatever the preflight check found missing.  The method is
        **idempotent**: it only creates what is absent.

        The creation sequence mirrors D365 Project Operations dependencies:
        1. Bookable resource (if missing)
        2. Project (if no active projects for the user)
        3. Project team member record (if user not on a project team)
        4. Project task (if project has no tasks)
        5. Resource assignment (if user is not assigned to any task)

        Returns a new preflight-style dict with ``ok``, ``warnings``,
        ``data``, and an extra ``created`` list describing what was made.
        """
        warnings: list[str] = preflight_result.get("warnings", [])
        data: dict = preflight_result.get("data", {})
        created: list[str] = []
        user_id = data.get("userId")

        if not user_id:
            # Try to get userId via Xrm
            try:
                user_id = await self.browser.page.evaluate(
                    "() => Xrm.Utility.getGlobalContext().userSettings.userId.replace(/[{}]/g, '')"
                )
                data["userId"] = user_id
            except Exception:
                return {
                    "ok": False,
                    "warnings": ["Cannot determine current user ID — sample data creation aborted."],
                    "data": data,
                    "created": [],
                }

        # Helper: query Xrm for a single value
        async def xrm_query(entity: str, odata: str) -> list[dict]:
            js = f"""async () => {{
                try {{
                    const r = await Xrm.WebApi.retrieveMultipleRecords('{entity}', '{odata}');
                    return r.entities || [];
                }} catch (e) {{ return []; }}
            }}"""
            return await self.browser.page.evaluate(js) or []

        # ----------------------------------------------------------------
        # 1. Bookable resource
        # ----------------------------------------------------------------
        resource_id = None
        resources = await xrm_query(
            "bookableresource",
            f"?$filter=_userid_value eq {user_id}&$select=bookableresourceid,name&$top=1",
        )
        if resources:
            resource_id = resources[0].get("bookableresourceid")
            data["resourceName"] = resources[0].get("name", "")
        else:
            # Need to create a bookable resource for the current user
            # First get the user's display name
            user_name = await self.browser.page.evaluate(
                "() => Xrm.Utility.getGlobalContext().userSettings.userName || 'Demo User'"
            )
            # Get the default org unit from project parameters
            params = await xrm_query(
                "msdyn_projectparameter",
                "?$select=_msdyn_defaultorganizationalunit_value&$top=1",
            )
            org_unit_id = (
                params[0].get("_msdyn_defaultorganizationalunit_value")
                if params else None
            )
            resource_payload: dict = {
                "name": user_name,
                "resourcetype": 3,  # User
                "userid@odata.bind": f"/systemusers({user_id})",
            }
            if org_unit_id:
                resource_payload[
                    "msdyn_organizationalunit@odata.bind"
                ] = f"/msdyn_organizationalunits({org_unit_id})"
            resource_id = await self._post_record("bookableresources", resource_payload)
            if resource_id:
                created.append(f"Bookable resource '{user_name}' (ID: {resource_id})")
                data["resourceName"] = user_name
            else:
                return {
                    "ok": False,
                    "warnings": ["Failed to create bookable resource for the current user."],
                    "data": data,
                    "created": created,
                }

        # ----------------------------------------------------------------
        # 2. Active project
        # ----------------------------------------------------------------
        project_id = data.get("projectId")
        project_name = data.get("projectName")
        if not project_id:
            # Check if there is *any* active project in the system we can reuse
            any_projects = await xrm_query(
                "msdyn_project",
                "?$filter=statecode eq 0&$select=msdyn_projectid,msdyn_subject&$top=1",
            )
            if any_projects:
                project_id = any_projects[0].get("msdyn_projectid")
                project_name = any_projects[0].get("msdyn_subject", "")
                logger.info(
                    "[SAMPLE DATA] Found existing active project '%s' to reuse",
                    project_name,
                )
            else:
                # Create a sample project
                # Look up default org unit + currency
                params = await xrm_query(
                    "msdyn_projectparameter",
                    "?$select=_msdyn_defaultorganizationalunit_value&$top=1",
                )
                org_unit_id = (
                    params[0].get("_msdyn_defaultorganizationalunit_value")
                    if params else None
                )
                # Get currency from org unit
                currency_id = None
                if org_unit_id:
                    org_units = await xrm_query(
                        "msdyn_organizationalunit",
                        f"?$filter=msdyn_organizationalunitid eq {org_unit_id}"
                        "&$select=_transactioncurrencyid_value&$top=1",
                    )
                    if org_units:
                        currency_id = org_units[0].get("_transactioncurrencyid_value")

                project_payload: dict = {
                    "msdyn_subject": "Zava Demo Project",
                    "msdyn_description": "Auto-created sample project for demo purposes",
                    "msdyn_schedulemode": 192350000,  # Fixed duration
                }
                if org_unit_id:
                    project_payload[
                        "msdyn_contractorganizationalunitid@odata.bind"
                    ] = f"/msdyn_organizationalunits({org_unit_id})"
                if currency_id:
                    project_payload[
                        "transactioncurrencyid@odata.bind"
                    ] = f"/transactioncurrencies({currency_id})"
                # Set the current user as project manager
                project_payload[
                    "msdyn_projectmanager@odata.bind"
                ] = f"/systemusers({user_id})"

                project_id = await self._post_record("msdyn_projects", project_payload)
                if project_id:
                    project_name = "Zava Demo Project"
                    created.append(f"Project '{project_name}' (ID: {project_id})")
                else:
                    return {
                        "ok": False,
                        "warnings": ["Failed to create a sample project."],
                        "data": data,
                        "created": created,
                    }
            data["projectId"] = project_id
            data["projectName"] = project_name

        # ----------------------------------------------------------------
        # 3. Team membership
        # ----------------------------------------------------------------
        teams = await xrm_query(
            "msdyn_projectteam",
            f"?$filter=_msdyn_bookableresourceid_value eq {resource_id}"
            f" and _msdyn_project_value eq {project_id}"
            "&$select=msdyn_projectteamid&$top=1",
        )
        if not teams:
            team_payload: dict = {
                "msdyn_project@odata.bind": f"/msdyn_projects({project_id})",
                "msdyn_bookableresourceid@odata.bind": f"/bookableresources({resource_id})",
                "msdyn_name": data.get("resourceName", "Demo User"),
            }
            team_id = await self._post_record("msdyn_projectteams", team_payload)
            if team_id:
                created.append(f"Project team member (ID: {team_id})")
            else:
                logger.warning("[SAMPLE DATA] Failed to create project team member — continuing")

        # ----------------------------------------------------------------
        # 4. Project task — find or create one ON this specific project
        # ----------------------------------------------------------------
        task_id = data.get("taskId")  # May have been found by preflight
        task_name = data.get("taskName")

        if not task_id:
            # Look for existing tasks on THIS project
            existing_tasks = await xrm_query(
                "msdyn_projecttask",
                f"?$filter=_msdyn_project_value eq {project_id}"
                "&$select=msdyn_projecttaskid,msdyn_subject&$top=5",
            )
            if existing_tasks:
                task_id = existing_tasks[0].get("msdyn_projecttaskid")
                task_name = existing_tasks[0].get("msdyn_subject", "")
                logger.info(
                    "[SAMPLE DATA] Found existing task '%s' on project '%s'",
                    task_name, project_name,
                )
            else:
                task_payload: dict = {
                    "msdyn_subject": "Engineering Design Review",
                    "msdyn_project@odata.bind": f"/msdyn_projects({project_id})",
                }
                task_id = await self._post_record("msdyn_projecttasks", task_payload)
                if task_id:
                    task_name = "Engineering Design Review"
                    created.append(f"Project task '{task_name}' (ID: {task_id})")
                else:
                    logger.warning("[SAMPLE DATA] Failed to create project task — task field will be empty")

        if task_name:
            data["taskName"] = task_name
        if task_id:
            data["taskId"] = task_id

        # ----------------------------------------------------------------
        # 5. Resource assignment — ensure user is assigned to the task
        # ----------------------------------------------------------------
        if task_id and resource_id:
            # Check if user already has an assignment to THIS specific task
            assignments = await xrm_query(
                "msdyn_resourceassignment",
                f"?$filter=_msdyn_bookableresourceid_value eq {resource_id}"
                f" and _msdyn_taskid_value eq {task_id}"
                f" and _msdyn_projectid_value eq {project_id}"
                "&$select=msdyn_resourceassignmentid&$top=1",
            )
            if not assignments:
                assign_payload: dict = {
                    "msdyn_projectid@odata.bind": f"/msdyn_projects({project_id})",
                    "msdyn_taskid@odata.bind": f"/msdyn_projecttasks({task_id})",
                    "msdyn_bookableresourceid@odata.bind": f"/bookableresources({resource_id})",
                }
                assign_id = await self._post_record(
                    "msdyn_resourceassignments", assign_payload
                )
                if assign_id:
                    created.append(f"Resource assignment on task '{task_name}' (ID: {assign_id})")
                else:
                    logger.warning("[SAMPLE DATA] Failed to create resource assignment")
            else:
                logger.info("[SAMPLE DATA] User already assigned to task '%s'", task_name)

        logger.info(
            "[SAMPLE DATA] Done — created %d records: %s",
            len(created), "; ".join(created) if created else "(none needed)",
        )
        return {
            "ok": True,
            "warnings": [],
            "data": data,
            "created": created,
        }

    @staticmethod
    def _alternative_selector(selector: str) -> str | None:
        """Generate an alternative selector from a failing one."""
        # If selector uses aria-label, try data-id based match
        m = re.search(r'aria-label=["\']([^"\']+)["\']', selector)
        if m:
            label = m.group(1)
            # Try common D365 data-id patterns
            return (
                f'[data-id*="{label}"], '
                f'button:has-text("{label}"), '
                f'span:has-text("{label}")'
            )

        # If selector uses data-id, try aria-label fallback
        m = re.search(r'data-id=["\']([^"\']+)["\']', selector)
        if m:
            data_id = m.group(1)
            return f'[aria-label*="{data_id}"]'

        return None
    # ---- Title & Closing Slides ----

    async def _show_title_slide(self, plan: DemoPlan):
        """Show the opening title slide."""
        await self.overlay.show_title_slide(
            heading=plan.title,
            subheading=plan.subtitle or "",
            meta=f"Dynamics 365 Project Operations  •  {plan.estimated_duration_minutes} min  •  {plan.total_steps} steps",
        )
        await self._interruptible_sleep(4.0)
        await self.state.wait_if_paused()
        await self.overlay.hide_title_slide()
        await self._interruptible_sleep(0.5)

    async def _show_closing(self, plan: DemoPlan):
        """Show the closing slide."""
        closing = plan.closing_text or f"Thank you for joining this demo of {plan.title}."
        await self.overlay.clear_all()
        await self.overlay.show_title_slide(
            heading="Demo Complete",
            subheading=closing,
            meta=(
                f"{plan.total_steps} capabilities demonstrated  •  "
                f"{self.state.elapsed_display} elapsed"
            ),
        )
        await self._interruptible_sleep(5.0)
        await self.overlay.hide_title_slide()

    # ---- Section Transitions ----

    async def _section_transition(self, prev: DemoSection, next_section: DemoSection):
        """Show a transition between sections."""
        self.state.set_phase("transition")
        self._notify()

        text = next_section.transition_text or (
            f"Now let's look at <span class='highlight'>{next_section.title}</span> — "
            f"{next_section.description}"
        )

        await self.overlay.clear_all()

        # Start voice narration in background (plays concurrently with caption)
        if self._voice:
            await self._voice.speak_async(text)

        await self.overlay.show_caption_animated(text, phase="tell", speed=20)

        if self._voice and self._voice.enabled:
            await self._voice.wait_for_completion()
        else:
            await self._interruptible_sleep(3.0)
        await self.state.wait_if_paused()
        await self.overlay.hide_caption()
        await self._interruptible_sleep(0.5)

    # ---- Pause ----

    async def _pause_for_user(self):
        """Pause execution and wait for user to resume."""
        logger.info("Pausing for user...")
        # Stop any active voice narration during pause
        if self._voice:
            await self._voice.stop()
        self.state.pause()
        self._notify()

        await self.overlay.pause()
        await self.state.wait_if_paused()
        await self.overlay.resume()
