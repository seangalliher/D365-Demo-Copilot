"""Tests for demo_agent.agent.state — DemoState machine."""

from __future__ import annotations

import asyncio
import time

import pytest

from demo_agent.agent.state import DemoState, DemoStatus, StepHistory


# ---- StepHistory ----

class TestStepHistory:

    def test_construction(self):
        h = StepHistory(
            step_id="s1",
            section_id="sec1",
            started_at=100.0,
            completed_at=105.0,
        )
        assert h.step_id == "s1"
        assert h.skipped is False

    def test_skipped_flag(self):
        h = StepHistory(
            step_id="s1",
            section_id="sec1",
            started_at=100.0,
            completed_at=105.0,
            skipped=True,
        )
        assert h.skipped is True


# ---- DemoState Initialization ----

class TestDemoStateInit:

    def test_default_status(self):
        state = DemoState()
        assert state.status == DemoStatus.IDLE

    def test_default_indices(self):
        state = DemoState()
        assert state.current_section_index == 0
        assert state.current_step_index == 0
        assert state.current_local_step_index == 0

    def test_resume_event_starts_set(self):
        state = DemoState()
        assert state._resume_event.is_set()

    def test_advance_event_starts_clear(self):
        state = DemoState()
        assert not state._advance_event.is_set()

    def test_user_acted_event_starts_clear(self):
        state = DemoState()
        assert not state._user_acted_event.is_set()

    def test_empty_history(self):
        state = DemoState()
        assert state.history == []

    def test_started_at_none(self):
        state = DemoState()
        assert state.started_at is None


# ---- start() ----

class TestDemoStateStart:

    def test_start_sets_title_slide(self):
        state = DemoState()
        state.start(10, 3)
        assert state.status == DemoStatus.TITLE_SLIDE

    def test_start_sets_totals(self):
        state = DemoState()
        state.start(10, 3)
        assert state.total_steps == 10
        assert state.total_sections == 3

    def test_start_resets_indices(self):
        state = DemoState()
        state.current_step_index = 5
        state.current_section_index = 2
        state.start(10, 3)
        assert state.current_step_index == 0
        assert state.current_section_index == 0
        assert state.current_local_step_index == 0

    def test_start_sets_started_at(self):
        state = DemoState()
        before = time.time()
        state.start(10, 3)
        after = time.time()
        assert before <= state.started_at <= after

    def test_start_clears_history(self):
        state = DemoState()
        state.history = [
            StepHistory("s1", "sec1", 0.0, 1.0),
        ]
        state.start(10, 3)
        assert state.history == []

    def test_start_clears_error(self):
        state = DemoState()
        state.error_message = "something broke"
        state.start(10, 3)
        assert state.error_message is None


# ---- set_phase() ----

class TestSetPhase:

    def test_tell_before(self):
        state = DemoState()
        state.set_phase("tell_before")
        assert state.status == DemoStatus.TELL_BEFORE
        assert state.current_phase == "tell_before"

    def test_show(self):
        state = DemoState()
        state.set_phase("show")
        assert state.status == DemoStatus.SHOWING

    def test_tell_after(self):
        state = DemoState()
        state.set_phase("tell_after")
        assert state.status == DemoStatus.TELL_AFTER

    def test_value(self):
        state = DemoState()
        state.set_phase("value")
        assert state.status == DemoStatus.VALUE

    def test_transition(self):
        state = DemoState()
        state.set_phase("transition")
        assert state.status == DemoStatus.TRANSITIONING

    def test_unknown_phase_keeps_status(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.set_phase("unknown_phase")
        assert state.status == DemoStatus.SHOWING
        assert state.current_phase == "unknown_phase"


# ---- advance_step / advance_section ----

class TestAdvance:

    def test_advance_step(self):
        state = DemoState()
        state.advance_step()
        assert state.current_step_index == 1
        assert state.current_local_step_index == 1

    def test_advance_section(self):
        state = DemoState()
        state.current_local_step_index = 3
        state.advance_section()
        assert state.current_section_index == 1
        assert state.current_local_step_index == 0


# ---- record_step ----

class TestRecordStep:

    def test_record_step_appends(self):
        state = DemoState()
        t = time.time()
        state.record_step("s1", "sec1", t)
        assert len(state.history) == 1
        h = state.history[0]
        assert h.step_id == "s1"
        assert h.section_id == "sec1"
        assert h.started_at == t
        assert h.completed_at >= t
        assert h.skipped is False

    def test_record_step_skipped(self):
        state = DemoState()
        state.record_step("s1", "sec1", time.time(), skipped=True)
        assert state.history[0].skipped is True

    def test_multiple_records(self):
        state = DemoState()
        state.record_step("s1", "sec1", time.time())
        state.record_step("s2", "sec1", time.time())
        assert len(state.history) == 2


# ---- complete / set_error ----

class TestCompletion:

    def test_complete(self):
        state = DemoState()
        state.complete()
        assert state.status == DemoStatus.COMPLETED

    def test_set_error(self):
        state = DemoState()
        state.set_error("boom")
        assert state.status == DemoStatus.ERROR
        assert state.error_message == "boom"


# ---- pause / resume ----

class TestPauseResume:

    def test_pause_changes_status(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.pause()
        assert state.status == DemoStatus.PAUSED
        assert state.is_paused

    def test_pause_saves_previous(self):
        state = DemoState()
        state.status = DemoStatus.TELL_BEFORE
        state.pause()
        assert state._previous_status == DemoStatus.TELL_BEFORE

    def test_pause_clears_resume_event(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.pause()
        assert not state._resume_event.is_set()

    def test_resume_restores_status(self):
        state = DemoState()
        state.status = DemoStatus.TELL_AFTER
        state.pause()
        state.resume()
        assert state.status == DemoStatus.TELL_AFTER
        assert not state.is_paused

    def test_resume_sets_event(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.pause()
        state.resume()
        assert state._resume_event.is_set()

    def test_pause_noop_when_already_paused(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.pause()
        state.pause()  # second pause is noop
        assert state._previous_status == DemoStatus.SHOWING

    def test_pause_noop_when_completed(self):
        state = DemoState()
        state.status = DemoStatus.COMPLETED
        state.pause()
        assert state.status == DemoStatus.COMPLETED

    def test_pause_noop_when_idle(self):
        state = DemoState()
        state.pause()
        assert state.status == DemoStatus.IDLE

    def test_resume_noop_when_not_paused(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.resume()  # Should be noop
        assert state.status == DemoStatus.SHOWING

    def test_resume_defaults_to_showing(self):
        state = DemoState()
        state.status = DemoStatus.SHOWING
        state.pause()
        state._previous_status = None  # Edge case
        state.resume()
        assert state.status == DemoStatus.SHOWING


# ---- Progress ----

class TestProgress:

    def test_progress_pct_zero(self):
        state = DemoState()
        assert state.progress_pct == 0.0

    def test_progress_pct_midway(self):
        state = DemoState()
        state.total_steps = 10
        state.current_step_index = 5
        assert state.progress_pct == 50.0

    def test_progress_pct_complete(self):
        state = DemoState()
        state.total_steps = 4
        state.current_step_index = 4
        assert state.progress_pct == 100.0

    def test_elapsed_display_no_start(self):
        state = DemoState()
        assert state.elapsed_display == "0:00"

    def test_elapsed_display_with_time(self):
        state = DemoState()
        state.started_at = time.time() - 65  # 1m5s ago
        display = state.elapsed_display
        assert display == "1:05"


# ---- to_dict ----

class TestToDict:

    def test_keys(self):
        state = DemoState()
        state.start(10, 3)
        d = state.to_dict()
        expected_keys = {
            "status", "section", "step", "phase",
            "progress", "elapsed", "steps_completed",
        }
        assert set(d.keys()) == expected_keys

    def test_values(self):
        state = DemoState()
        state.start(10, 3)
        state.current_step_index = 2
        state.current_section_index = 1
        d = state.to_dict()
        assert d["status"] == "title_slide"
        assert d["section"] == "2/3"
        assert d["step"] == "3/10"
        assert d["progress"] == "20%"
        assert d["steps_completed"] == 0


# ---- Async tests ----

class TestAsyncState:

    @pytest.mark.asyncio
    async def test_wait_for_step_trigger_auto_play(self):
        state = DemoState()
        state.step_mode = False
        result = await state.wait_for_step_trigger()
        assert result == "advance"

    @pytest.mark.asyncio
    async def test_signal_advance_sets_event(self):
        state = DemoState()
        state.signal_advance()
        assert state._advance_event.is_set()

    @pytest.mark.asyncio
    async def test_signal_user_acted_sets_event(self):
        state = DemoState()
        state.signal_user_acted()
        assert state._user_acted_event.is_set()

    @pytest.mark.asyncio
    async def test_wait_for_step_trigger_advance(self):
        state = DemoState()
        state.step_mode = True

        async def fire():
            await asyncio.sleep(0.05)
            state.signal_advance()

        asyncio.create_task(fire())
        result = await state.wait_for_step_trigger()
        assert result == "advance"

    @pytest.mark.asyncio
    async def test_wait_for_step_trigger_user_acted(self):
        state = DemoState()
        state.step_mode = True

        async def fire():
            await asyncio.sleep(0.05)
            state.signal_user_acted()

        asyncio.create_task(fire())
        result = await state.wait_for_step_trigger()
        assert result == "user_acted"

    @pytest.mark.asyncio
    async def test_wait_if_paused_returns_immediately(self):
        state = DemoState()
        # resume_event is set by default, should return immediately
        await asyncio.wait_for(state.wait_if_paused(), timeout=1.0)
