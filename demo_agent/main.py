"""
D365 Demo Copilot — Main Entry Point

Sidecar-mode agent that:
1. Opens a browser with D365 + an in-browser chat panel
2. Accepts demo requests through the sidecar chat UI
3. Generates demo plans using an LLM, shown inline in the chat
4. Executes demos with visual overlays, progress shown in the chat
5. Handles pause/resume/skip/quit via chat quick-action buttons

The terminal is log-only — all user interaction happens in the browser.

Usage:
    python -m demo_agent.main
"""

from __future__ import annotations

import asyncio
import base64
import logging
import sys
from pathlib import Path

from rich.console import Console
from rich.logging import RichHandler
from rich.panel import Panel

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[RichHandler(rich_tracebacks=True, show_time=False)],
)
logger = logging.getLogger("demo_agent")

console = Console()


def print_banner():
    """Print the application banner."""
    console.print(
        Panel.fit(
            "[bold blue]D365 Demo Copilot[/bold blue]\n"
            "[dim]AI-powered live demonstrations for Dynamics 365[/dim]\n"
            "[dim]CE/Dataverse • Finance • Supply Chain Management[/dim]\n\n"
            "[green]The sidecar chat panel will open in the browser.[/green]\n"
            "[green]Interact with the demo through the browser panel.[/green]",
            border_style="blue",
            padding=(1, 4),
        )
    )


async def run_login_flow(browser):
    """
    Run an interactive login flow so the user can authenticate to D365.
    Saves the auth state for future sessions.
    """
    from .config import DemoConfig

    config = DemoConfig()

    console.print("\n[yellow]Authentication Required[/yellow]")
    console.print(
        f"A browser will open to [bold]{config.d365_base_url}[/bold].\n"
        "Please log in with your credentials.\n"
    )

    await browser.start()
    await browser.navigate(config.d365_base_url)

    console.print("[dim]Waiting for you to complete login...[/dim]")
    console.print("[dim]Press Enter here once you're logged in and see the D365 home page.[/dim]\n")

    # Wait for user to confirm login
    await asyncio.get_event_loop().run_in_executor(None, input)

    # Save auth state
    await browser.save_auth_state(config.auth_state_path)
    console.print(f"[green]Auth state saved to {config.auth_state_path}[/green]\n")

    return browser


async def create_llm_client(config):
    """Create the appropriate LLM client based on configuration.

    Priority: Azure OpenAI -> GitHub Models -> OpenAI direct.
    """
    if config.use_azure_openai:
        from openai import AsyncAzureOpenAI

        return AsyncAzureOpenAI(
            azure_endpoint=config.azure_openai_endpoint,
            api_key=config.azure_openai_api_key,
            api_version=config.azure_openai_api_version,
        )
    elif config.use_github_copilot:
        from openai import AsyncOpenAI

        logger.info(
            "Using GitHub Models API (model: %s, endpoint: %s)",
            config.github_copilot_model,
            config.github_models_base_url,
        )
        return AsyncOpenAI(
            base_url=config.github_models_base_url,
            api_key=config.github_token,
        )
    elif config.use_openai:
        from openai import AsyncOpenAI

        return AsyncOpenAI(api_key=config.openai_api_key)
    else:
        raise ValueError(
            "No LLM credentials configured. "
            "Set AZURE_OPENAI_ENDPOINT + AZURE_OPENAI_API_KEY, "
            "GITHUB_TOKEN, or OPENAI_API_KEY in .env"
        )


async def main():
    """Main entry point — launches browser with sidecar chat panel."""
    from .agent.executor import DemoExecutor
    from .agent.planner import DemoPlanner
    from .agent.schema_discovery import SchemaDiscovery
    from .agent.learn_docs import LearnDocsClient
    from .agent.script_recorder import ScriptRecorder
    from .agent.script_generator import ScriptGenerator
    from .agent.state import DemoState
    from .agent.voice import VoiceNarrator
    from .auth.dataverse_auth import DataverseAuth
    from .browser.chat_panel import ChatPanelManager
    from .browser.controller import BrowserController
    from .browser.d365_pages import D365Navigator
    from .browser.overlay_manager import OverlayManager
    from .config import DemoConfig

    print_banner()

    # Load config
    config = DemoConfig()
    errors = config.validate()
    if errors:
        for e in errors:
            console.print(f"[red]{e}[/red]")
        console.print(
            "\n[yellow]Create a .env file in the project root with your credentials.[/yellow]"
        )
        return

    # ---- Dataverse OAuth (app registration) ----
    auth_headers: dict[str, str] = {}
    if config.use_dataverse_auth:
        try:
            dv_auth = DataverseAuth(
                tenant_id=config.dataverse_tenant_id,
                client_id=config.dataverse_client_id,
                client_secret=config.dataverse_client_secret,
                resource_url=config.d365_base_url,
            )
            auth_headers = dv_auth.get_auth_headers()
            console.print("[green]Acquired Dataverse OAuth token via app registration[/green]")
        except Exception as exc:
            console.print(
                f"[yellow]Dataverse OAuth failed ({exc}) — will fall back to built-in schemas[/yellow]"
            )
    else:
        console.print(
            "[dim]No Dataverse app-registration credentials — set DATAVERSE_TENANT_ID, "
            "DATAVERSE_CLIENT_ID, DATAVERSE_CLIENT_SECRET in .env for live MCP[/dim]"
        )

    # ---- Connect to Dataverse MCP for schema discovery ----
    mcp_url = f"{config.d365_base_url}/api/mcp"
    schema_discovery = SchemaDiscovery(mcp_url, auth_headers=auth_headers)
    mcp_connected = await schema_discovery.connect()
    if mcp_connected:
        console.print("[green]Connected to Dataverse MCP for live schema discovery[/green]")
    else:
        console.print(
            "[yellow]Dataverse MCP not available — using built-in D365 schemas[/yellow]"
        )

    # ---- Connect to Microsoft Learn MCP for documentation enrichment ----
    learn_docs = LearnDocsClient(url=config.ms_learn_mcp_url)
    learn_connected = await learn_docs.connect()
    if learn_connected:
        console.print("[green]Connected to Microsoft Learn MCP for doc enrichment[/green]")
    else:
        console.print(
            "[yellow]MS Learn MCP not available — demos will use built-in knowledge only[/yellow]"
        )

    # Initialize LLM
    llm_client = await create_llm_client(config)
    active_model = (
        config.github_copilot_model if config.use_github_copilot else config.llm_model
    )
    planner = DemoPlanner(
        llm_client=llm_client,
        model=active_model,
        d365_base_url=config.d365_base_url,
        schema_discovery=schema_discovery,
        learn_docs=learn_docs if learn_connected else None,
    )
    state = DemoState()

    # ---- Launch browser ----
    browser = BrowserController(
        base_url=config.d365_base_url,
        headless=config.headless,
        slow_mo=config.slow_mo,
    )

    auth_path = config.auth_state_path
    if not Path(auth_path).exists():
        await run_login_flow(browser)
    else:
        await browser.start(storage_state=auth_path)
        console.print("[dim]Browser started — navigating to D365...[/dim]")
        try:
            await browser.page.goto(
                config.d365_base_url, wait_until="commit", timeout=90_000
            )
        except Exception as e:
            logger.warning("Initial navigation timeout (%s) — continuing anyway", e)

        # Wait for the page to settle after SSO redirects / slow loads
        for attempt in range(15):
            try:
                await browser.page.wait_for_load_state("domcontentloaded", timeout=5_000)
                # Verify execution context is alive
                await browser.page.evaluate("() => document.readyState")
                break
            except Exception:
                logger.info("Page not stable yet (attempt %d/15) — waiting...", attempt + 1)
                await asyncio.sleep(2)
        else:
            logger.warning("Page did not stabilise after 15 attempts — injecting anyway")

    console.print("[dim]Browser launched. Injecting chat panel...[/dim]")

    # ---- Inject sidecar chat panel ----
    chat = ChatPanelManager(page=browser.page)

    # Retry injection — context can still be destroyed by late redirects
    for inject_attempt in range(3):
        try:
            await chat.inject()
            break
        except Exception as inject_err:
            logger.warning(
                "Chat inject attempt %d failed (%s) — retrying...",
                inject_attempt + 1,
                inject_err,
            )
            await asyncio.sleep(3)
    else:
        logger.error("Failed to inject chat panel after 3 attempts")
        console.print("[red]Could not inject chat panel. Try refreshing D365 and re-running.[/red]")
        return

    # Shrink D365 content to make room for the chat panel
    await browser.page.evaluate(
        """() => {
            function syncMargin() {
                const host = document.getElementById('demo-chat-panel-host');
                if (host && !host.classList.contains('collapsed')) {
                    document.body.style.marginRight = host.offsetWidth + 'px';
                } else if (!host || host.classList.contains('collapsed')) {
                    document.body.style.marginRight = '40px';
                }
            }
            syncMargin();
            document.body.style.transition = 'margin-right 0.3s ease';
            // Re-adjust on resize and at intervals (catches display changes)
            window.__demoChatResizeHandler = syncMargin;
            window.addEventListener('resize', syncMargin);
            setInterval(syncMargin, 3000);
        }"""
    )

    overlay = OverlayManager(browser.page)
    d365_nav = D365Navigator(browser.page, config.d365_base_url)
    current_plan = None
    demo_task = None

    # ---- Voice narrator ----
    voice = VoiceNarrator(
        page=browser.page,
        model=config.voice_model,
        voice=config.voice_name,
        speed=config.voice_speed,
        provider=config.voice_provider,
    )
    voice.enabled = config.voice_enabled

    # Sync initial voice state to the sidecar UI
    try:
        await chat.set_voice_enabled(config.voice_enabled)
    except Exception:
        pass

    if config.voice_enabled:
        console.print(f"[green]Voice narration enabled ({voice.provider} / {config.voice_name})[/green]")
    else:
        console.print(f"[dim]Voice narration disabled — provider: {voice.provider} (toggle in sidecar)[/dim]")

    # ---- Status change callback (updates chat panel) ----
    async def on_status_change_async(s: DemoState):
        """Update chat panel status and progress during demo execution."""
        status_labels = {
            "idle": ("Ready", ""),
            "planning": ("Planning...", "busy"),
            "ready": ("Ready", ""),
            "title_slide": ("Title Slide", "busy"),
            "tell_before": ("Narrating...", "busy"),
            "showing": ("Demonstrating...", "busy"),
            "tell_after": ("Summarizing...", "busy"),
            "value": ("Business Value", "busy"),
            "transitioning": ("Transitioning...", "busy"),
            "paused": ("Paused", "error"),
            "completed": ("Completed", ""),
            "error": ("Error", "error"),
        }
        label, stype = status_labels.get(s.status.value, ("Working...", "busy"))
        try:
            await chat.set_status(label, stype)
            if s.status.value == "completed":
                # Mark all steps done in both plan card and tracker
                if s.total_steps > 0:
                    for i in range(s.total_steps):
                        await chat.update_plan_step(i, "done")
                        await chat.update_tracker_step(i, "done")
                    await chat.update_progress(
                        s.total_steps - 1, s.total_steps, "Completed"
                    )
            elif s.total_steps > 0:
                step_label = f"{s.status.value.replace('_', ' ').title()}"
                await chat.update_progress(
                    s.current_step_index, s.total_steps, step_label
                )
                # Update plan card and step tracker highlights
                if s.current_step_index > 0:
                    await chat.update_plan_step(s.current_step_index - 1, "done")
                    await chat.update_tracker_step(s.current_step_index - 1, "done")
                await chat.update_plan_step(s.current_step_index, "active")
                await chat.update_tracker_step(s.current_step_index, "active")
        except Exception:
            pass  # Don't crash the demo if chat panel update fails

    def on_status_change(s: DemoState):
        """Sync wrapper that schedules the async status update."""
        asyncio.ensure_future(on_status_change_async(s))

    # Start navigation watcher so chat panel auto-reinjects after D365 SPA navigations
    chat.start_navigation_watcher()

    console.print("[green]Sidecar chat panel ready. Interact via the browser.[/green]")
    console.print("[dim]Terminal is log-only. Press Ctrl+C to quit.[/dim]\n")

    # ---- PDF script export state ----
    pending_pdf_b64 = None
    pending_pdf_filename = None
    pending_pdf_path = None

    # ---- Main event loop (browser-driven) ----
    try:
        while True:
            # Wait for user input from the chat panel
            try:
                event_type, event_value = await chat.wait_for_message_or_action()
            except Exception as e:
                logger.error("Error waiting for chat input: %s", e)
                await asyncio.sleep(1)
                continue

            # ---- Handle actions (quick buttons, plan buttons) ----
            if event_type == "action":
                action = event_value
                logger.info("Chat action: %s", action)

                if action == "start_demo" and current_plan:
                    # Start demo execution
                    await chat.disable_input()
                    await chat.set_status("Starting demo...", "busy")

                    # Reset PDF state from any prior demo
                    pending_pdf_b64 = None
                    pending_pdf_filename = None
                    pending_pdf_path = None

                    # Create script recorder for this demo run
                    script_recorder = ScriptRecorder(
                        page=browser.page, overlay=overlay,
                    )

                    executor = DemoExecutor(
                        browser=browser,
                        overlay=overlay,
                        d365_nav=d365_nav,
                        state=state,
                        on_status_change=on_status_change,
                        voice=voice,
                        script_recorder=script_recorder,
                        dataverse_api_url=f"{config.d365_base_url.rstrip('/')}/api/data/v9.2",
                        auth_headers=auth_headers,
                    )

                    # ---- Pre-flight data check ----
                    await chat.set_status("Checking data...", "busy")
                    preflight = await executor.preflight_check()
                    if preflight.get("warnings"):
                        for w in preflight["warnings"]:
                            await chat.add_message("system", f"⚠️ {w}")
                            logger.warning("[PREFLIGHT] %s", w)

                    if not preflight.get("ok") or preflight.get("warnings"):
                        # Attempt to create missing sample data via Dataverse API
                        await chat.add_message(
                            "system",
                            "Creating sample data for the demo...",
                        )
                        await chat.set_status("Creating sample data...", "busy")
                        sample_result = await executor.create_sample_data(preflight)
                        if sample_result.get("created"):
                            for item in sample_result["created"]:
                                await chat.add_message("system", f"✅ Created: {item}")
                        if not sample_result.get("ok"):
                            # Still failing after creation attempt
                            for w in sample_result.get("warnings", []):
                                await chat.add_message("system", f"⚠️ {w}")
                            await chat.add_message(
                                "system",
                                "Pre-flight check failed — could not create required sample data.",
                            )
                            await chat.set_status("Data issues found", "error")
                            await chat.enable_input()
                            continue
                        # Sample data created — update preflight with new data
                        preflight = sample_result

                    # Store preflight data on executor so lookups use verified values
                    pf_data = preflight.get("data", {})
                    executor.set_preflight_data(pf_data)

                    # Log discovered data
                    if pf_data.get("projectName"):
                        await chat.add_message(
                            "system",
                            f"✅ Data check passed — project: {pf_data['projectName']}"
                            + (f", task: {pf_data['taskName']}" if pf_data.get('taskName') else ""),
                        )

                    await chat.show_quick_actions([
                        {"label": "⏸ Pause", "action": "pause"},
                        {"label": "⏭ Skip", "action": "skip"},
                        {"label": "⏹ Stop", "action": "quit", "danger": True},
                    ])

                    # Show the vertical step timeline and hide welcome/plan
                    plan_dict = _plan_to_dict(current_plan)
                    await chat.hide_welcome()
                    await chat.show_step_tracker(plan_dict)

                    await chat.add_message("system", "Demo starting...")
                    await chat.show_progress(0, current_plan.total_steps, "Starting...")

                    demo_task = asyncio.create_task(executor.execute(current_plan))

                    # Wait for demo to finish or be interrupted by actions
                    try:
                        while not demo_task.done():
                            try:
                                act_type, act_val = await chat.wait_for_message_or_action(
                                    timeout=0.5
                                )
                                if act_type == "action":
                                    if act_val == "pause":
                                        state.pause()
                                        await voice.stop()
                                        await overlay.pause()
                                        await chat.add_message("system", "Demo paused")
                                        await chat.show_quick_actions([
                                            {"label": "▶ Resume", "action": "resume"},
                                            {"label": "⏭ Skip", "action": "skip"},
                                            {"label": "⏹ Stop", "action": "quit", "danger": True},
                                        ])
                                    elif act_val == "resume":
                                        state.resume()
                                        await overlay.resume()
                                        await chat.add_message("system", "Demo resumed")
                                        await chat.show_quick_actions([
                                            {"label": "⏸ Pause", "action": "pause"},
                                            {"label": "⏭ Skip", "action": "skip"},
                                            {"label": "⏹ Stop", "action": "quit", "danger": True},
                                        ])
                                    elif act_val == "advance":
                                        if state.is_paused:
                                            state.resume()
                                            await overlay.resume()
                                        state.signal_advance()
                                    elif act_val == "user_acted":
                                        state.signal_user_acted()
                                    elif act_val == "skip":
                                        if state.is_paused:
                                            state.resume()
                                            await overlay.resume()
                                        state.signal_advance()  # Skip also advances
                                        await chat.add_message("system", "Skipping step...")
                                    elif act_val == "quit":
                                        await voice.stop()
                                        demo_task.cancel()
                                        await chat.add_message("system", "Demo stopped by user")
                                        break
                                    elif act_val == "voice_enable":
                                        voice.enabled = True
                                        await chat.add_message("system", "Voice narration enabled")
                                    elif act_val == "voice_disable":
                                        voice.enabled = False
                                        await voice.stop()
                                        await chat.add_message("system", "Voice narration disabled")
                            except asyncio.TimeoutError:
                                continue
                    except Exception as e:
                        logger.error("Error during demo: %s", e)

                    # Demo finished
                    try:
                        await demo_task
                    except asyncio.CancelledError:
                        pass
                    except Exception as e:
                        await chat.add_message("system", f"Demo error: {e}")
                        logger.exception("Demo execution error")

                    # ---- Generate demo script PDF ----
                    if script_recorder and script_recorder.captures:
                        try:
                            await chat.add_message("system", "Generating demo script PDF...")
                            generator = ScriptGenerator()
                            pdf_bytes = generator.generate(
                                plan=current_plan,
                                captures=script_recorder.captures,
                                elapsed_display=state.elapsed_display,
                            )
                            safe_title = "".join(
                                c if c.isalnum() or c in " -_" else "_"
                                for c in current_plan.title
                            ).strip().replace(" ", "_")
                            pending_pdf_filename = f"Demo_Script_{safe_title}.pdf"

                            # Save PDF to disk (most reliable delivery)
                            downloads_dir = Path(__file__).resolve().parent.parent / "downloads"
                            downloads_dir.mkdir(exist_ok=True)
                            pending_pdf_path = downloads_dir / pending_pdf_filename
                            pending_pdf_path.write_bytes(pdf_bytes)

                            # Also keep base64 for browser download fallback
                            pending_pdf_b64 = base64.b64encode(pdf_bytes).decode("ascii")

                            logger.info(
                                "Demo script PDF saved: %s (%.1f KB)",
                                pending_pdf_path,
                                len(pdf_bytes) / 1024,
                            )
                        except Exception as e:
                            logger.error("PDF generation failed: %s", e, exc_info=True)
                            await chat.add_message(
                                "system", f"Could not generate demo script: {e}"
                            )

                    await chat.hide_progress()
                    await chat.hide_quick_actions()
                    await chat.enable_input()
                    await chat.set_status("Completed", "")

                    # Keep step tracker visible with all steps marked done
                    # (don't call hide_step_tracker — leave the timeline visible)

                    # Build summary with download action
                    summary_parts = [
                        f"✅ **Demo complete!**\n\n"
                        f"⏱ Duration: {state.elapsed_display}\n"
                        f"📋 Steps completed: {len(state.history)}\n",
                    ]
                    if pending_pdf_path:
                        summary_parts.append(
                            f"📄 Script saved to: {pending_pdf_path}\n"
                        )

                    summary_parts.append(
                        "\nWhat would you like to demonstrate next?"
                    )
                    completion_message = "".join(summary_parts)
                    await chat.add_message("assistant", completion_message)

                    # Show download button if PDF was generated
                    post_demo_actions = []
                    if pending_pdf_b64:
                        post_demo_actions = [
                            {"label": "\U0001f4e5 Download Script", "action": "download_script"},
                        ]
                        await chat.show_quick_actions(post_demo_actions)

                    await chat.set_demo_completed(
                        completion_message=completion_message,
                        post_demo_actions=post_demo_actions,
                    )

                    current_plan = None
                    demo_task = None

                elif action == "modify_plan" and current_plan:
                    await chat.add_message(
                        "assistant",
                        "What would you like to change about the plan?"
                    )
                    await chat.enable_input()

                elif action == "panel_collapsed":
                    # Restore D365 to full width
                    await browser.page.evaluate(
                        "document.body.style.marginRight = '40px'"
                    )

                elif action == "panel_expanded":
                    # Shrink D365 for chat panel (use host element width)
                    await browser.page.evaluate(
                        """() => {
                            const host = document.getElementById('demo-chat-panel-host');
                            document.body.style.marginRight = host ? host.offsetWidth + 'px' : '400px';
                        }"""
                    )

                elif action == "voice_enable":
                    voice.enabled = True
                    await chat.add_message("system", "Voice narration enabled")
                    logger.info("Voice narration enabled via sidecar toggle")

                elif action == "voice_disable":
                    voice.enabled = False
                    await voice.stop()
                    await chat.add_message("system", "Voice narration disabled")
                    logger.info("Voice narration disabled via sidecar toggle")

                elif action == "download_script" and (pending_pdf_b64 or pending_pdf_path):
                    try:
                        if pending_pdf_b64:
                            await chat.trigger_download(
                                pending_pdf_filename,
                                pending_pdf_b64,
                                "application/pdf",
                            )
                        if pending_pdf_path and pending_pdf_path.exists():
                            await chat.add_message(
                                "system",
                                f"Demo script saved to: {pending_pdf_path}"
                            )
                        else:
                            await chat.add_message("system", "Demo script downloaded!")
                    except Exception as e:
                        logger.error("PDF download trigger failed: %s", e)
                        if pending_pdf_path and pending_pdf_path.exists():
                            await chat.add_message(
                                "system",
                                f"Browser download failed, but file is saved at: {pending_pdf_path}"
                            )
                        else:
                            await chat.add_message("system", f"Download failed: {e}")

                elif action in ("pause", "resume", "skip", "quit"):
                    # These are handled inside the demo execution loop above
                    pass

                continue

            # ---- Handle user messages ----
            user_text = event_value
            logger.info("User message: %s", user_text[:80])

            # Check if this is a plan modification
            if current_plan and not demo_task:
                # User is modifying the plan
                await chat.set_status("Refining plan...", "busy")
                await chat.show_typing()
                await chat.disable_input()

                try:
                    current_plan = await planner.refine_plan(current_plan, user_text)
                    await chat.hide_typing()

                    # Show updated plan
                    plan_dict = _plan_to_dict(current_plan)
                    await chat.show_plan(plan_dict)
                    await chat.set_status("Plan ready", "")
                except Exception as e:
                    await chat.hide_typing()
                    await chat.add_message("system", f"Error refining plan: {e}")
                    await chat.set_status("Error", "error")
                    logger.exception("Plan refinement error")

                await chat.enable_input()
                continue

            # ---- Generate a new demo plan ----
            await chat.set_status("Planning...", "busy")
            await chat.show_typing()
            await chat.disable_input()
            await chat.reset_demo_ui_for_new_request()

            try:
                current_plan = await planner.create_plan(
                    customer_request=user_text,
                    max_steps=config.max_demo_steps,
                )
                await chat.hide_typing()

                # Show plan in chat
                plan_dict = _plan_to_dict(current_plan)
                await chat.show_plan(plan_dict)
                await chat.set_status("Plan ready", "")

                logger.info(
                    "Plan generated: %s (%d steps, %d sections)",
                    current_plan.title,
                    current_plan.total_steps,
                    len(current_plan.sections),
                )

            except Exception as e:
                await chat.hide_typing()
                await chat.add_message("system", f"Error generating plan: {e}")
                await chat.set_status("Error", "error")
                logger.exception("Plan generation error")

            await chat.enable_input()

    except KeyboardInterrupt:
        console.print("\n[dim]Shutting down...[/dim]")
    finally:
        # Cleanup
        try:
            await chat.destroy()
        except Exception:
            pass
        if schema_discovery:
            try:
                await schema_discovery.disconnect()
            except Exception:
                pass
        if learn_docs:
            try:
                await learn_docs.disconnect()
            except Exception:
                pass
        await browser.stop()
        console.print("[dim]Goodbye![/dim]")


def _plan_to_dict(plan) -> dict:
    """Convert a DemoPlan to a plain dict for the chat panel JS."""
    return {
        "title": plan.title,
        "estimated_duration_minutes": plan.estimated_duration_minutes,
        "total_steps": plan.total_steps,
        "sections": [
            {
                "title": section.title,
                "steps": [
                    {"title": step.title, "id": step.id}
                    for step in section.steps
                ],
            }
            for section in plan.sections
        ],
    }


def cli():
    """CLI entry point."""
    if "--setup" in sys.argv:
        from .setup import run_setup
        run_setup()
        return
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\n[dim]Interrupted. Goodbye![/dim]")


if __name__ == "__main__":
    cli()
