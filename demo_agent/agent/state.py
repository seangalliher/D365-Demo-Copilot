"""
D365 Demo Copilot — Demo State Machine

Manages the runtime state of a demo execution including:
- Current position (section, step, phase)
- Pause / resume / skip controls
- Execution history for replay
"""

from __future__ import annotations

import asyncio
import time
from dataclasses import dataclass, field
from enum import Enum
from typing import Optional


class DemoStatus(str, Enum):
    """Overall demo execution status."""
    IDLE = "idle"                 # No demo loaded
    PLANNING = "planning"         # Generating demo plan
    READY = "ready"               # Plan ready, not started
    TITLE_SLIDE = "title_slide"   # Showing title slide
    TELL_BEFORE = "tell_before"   # Showing pre-action narration
    SHOWING = "showing"           # Performing browser actions
    TELL_AFTER = "tell_after"     # Showing post-action narration
    VALUE = "value"               # Showing business value card
    TRANSITIONING = "transitioning"  # Between sections
    PAUSED = "paused"             # Paused by user
    COMPLETED = "completed"       # Demo finished
    ERROR = "error"               # Error occurred


@dataclass
class StepHistory:
    """Record of a completed step."""
    step_id: str
    section_id: str
    started_at: float
    completed_at: float
    skipped: bool = False


@dataclass
class DemoState:
    """
    Mutable runtime state for a demo execution.

    Thread-safe pause/resume via asyncio.Event.
    """
    status: DemoStatus = DemoStatus.IDLE
    current_section_index: int = 0
    current_step_index: int = 0       # Global step index (across all sections)
    current_local_step_index: int = 0  # Step index within current section
    current_phase: str = ""           # Current Tell-Show-Tell phase
    total_steps: int = 0
    total_sections: int = 0
    history: list[StepHistory] = field(default_factory=list)
    started_at: Optional[float] = None
    error_message: Optional[str] = None

    # Pause control
    _resume_event: asyncio.Event = field(default_factory=asyncio.Event)
    _previous_status: Optional[DemoStatus] = None

    # Step-advance control (Ctrl+Space / user-action driven)
    _advance_event: asyncio.Event = field(default_factory=asyncio.Event)
    _user_acted_event: asyncio.Event = field(default_factory=asyncio.Event)
    step_mode: bool = False  # When True, executor pauses after each step; False = auto-play

    def __post_init__(self):
        self._resume_event.set()  # Start in resumed state
        self._advance_event.clear()  # Not advanced yet
        self._user_acted_event.clear()

    def start(self, total_steps: int, total_sections: int):
        """Initialize state for a new demo run."""
        self.status = DemoStatus.TITLE_SLIDE
        self.current_section_index = 0
        self.current_step_index = 0
        self.current_local_step_index = 0
        self.current_phase = ""
        self.total_steps = total_steps
        self.total_sections = total_sections
        self.history = []
        self.started_at = time.time()
        self.error_message = None
        self._resume_event.set()

    def advance_step(self):
        """Move to the next step."""
        self.current_step_index += 1
        self.current_local_step_index += 1

    def advance_section(self):
        """Move to the next section."""
        self.current_section_index += 1
        self.current_local_step_index = 0

    def set_phase(self, phase: str):
        """Update the current Tell-Show-Tell phase."""
        self.current_phase = phase
        phase_to_status = {
            "tell_before": DemoStatus.TELL_BEFORE,
            "show": DemoStatus.SHOWING,
            "tell_after": DemoStatus.TELL_AFTER,
            "value": DemoStatus.VALUE,
            "transition": DemoStatus.TRANSITIONING,
        }
        self.status = phase_to_status.get(phase, self.status)

    def record_step(self, step_id: str, section_id: str, started_at: float, skipped: bool = False):
        """Record a completed step in history."""
        self.history.append(
            StepHistory(
                step_id=step_id,
                section_id=section_id,
                started_at=started_at,
                completed_at=time.time(),
                skipped=skipped,
            )
        )

    def complete(self):
        """Mark the demo as completed."""
        self.status = DemoStatus.COMPLETED

    def set_error(self, message: str):
        """Mark the demo as errored."""
        self.status = DemoStatus.ERROR
        self.error_message = message

    # ---- Pause / Resume ----

    def pause(self):
        """Pause the demo execution."""
        if self.status not in (DemoStatus.PAUSED, DemoStatus.COMPLETED, DemoStatus.IDLE):
            self._previous_status = self.status
            self.status = DemoStatus.PAUSED
            self._resume_event.clear()

    def resume(self):
        """Resume the demo execution."""
        if self.status == DemoStatus.PAUSED:
            self.status = self._previous_status or DemoStatus.SHOWING
            self._previous_status = None
            self._resume_event.set()

    @property
    def is_paused(self) -> bool:
        return self.status == DemoStatus.PAUSED

    async def wait_if_paused(self):
        """Await this to respect pause state. Returns immediately if not paused."""
        await self._resume_event.wait()

    # ---- Step Advance (Ctrl+Space / user-action) ----

    def signal_advance(self):
        """Signal the executor to advance to the next step (Ctrl+Space press)."""
        self._advance_event.set()

    def signal_user_acted(self):
        """Signal that the user manually performed the current step's actions."""
        self._user_acted_event.set()

    async def wait_for_advance(self):
        """Block until the user presses Ctrl+Space to advance. Resets the event after.

        No-op when ``step_mode`` is False (auto-play).
        """
        if not self.step_mode:
            return
        await self._advance_event.wait()
        self._advance_event.clear()

    async def wait_for_step_trigger(self) -> str:
        """Wait for either Ctrl+Space advance or user-action detection.

        Returns:
            ``'advance'`` -- user pressed Ctrl+Space; agent should execute SHOW phase.
            ``'user_acted'`` -- user performed actions manually; skip SHOW.

        No-op (returns ``'advance'``) when ``step_mode`` is False.
        """
        if not self.step_mode:
            return "advance"

        advance_task = asyncio.create_task(self._advance_event.wait())
        user_task = asyncio.create_task(self._user_acted_event.wait())

        done, pending = await asyncio.wait(
            {advance_task, user_task},
            return_when=asyncio.FIRST_COMPLETED,
        )

        for t in pending:
            t.cancel()
            try:
                await t
            except asyncio.CancelledError:
                pass

        result = "advance" if advance_task in done else "user_acted"
        self._advance_event.clear()
        self._user_acted_event.clear()
        return result

    # ---- Progress ----

    @property
    def progress_pct(self) -> float:
        """Current progress as a percentage (0-100)."""
        if self.total_steps == 0:
            return 0.0
        return (self.current_step_index / self.total_steps) * 100

    @property
    def elapsed_seconds(self) -> float:
        """Seconds since the demo started."""
        if not self.started_at:
            return 0.0
        return time.time() - self.started_at

    @property
    def elapsed_display(self) -> str:
        """Formatted elapsed time string."""
        secs = int(self.elapsed_seconds)
        mins, secs = divmod(secs, 60)
        return f"{mins}:{secs:02d}"

    def to_dict(self) -> dict:
        """Serialize state for display or logging."""
        return {
            "status": self.status.value,
            "section": f"{self.current_section_index + 1}/{self.total_sections}",
            "step": f"{self.current_step_index + 1}/{self.total_steps}",
            "phase": self.current_phase,
            "progress": f"{self.progress_pct:.0f}%",
            "elapsed": self.elapsed_display,
            "steps_completed": len(self.history),
        }
