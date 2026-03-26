"""Tests for demo_agent.models.demo_plan — Pydantic data models."""

from __future__ import annotations

import pytest
from pydantic import ValidationError

from demo_agent.models.demo_plan import (
    ActionType,
    DemoPlan,
    DemoSection,
    DemoStep,
    StepAction,
    StepPhase,
    ValueHighlight,
)


# ---- ActionType Enum ----

class TestActionType:

    def test_all_action_types_exist(self):
        expected = {
            "navigate", "click", "fill", "select", "hover",
            "scroll", "wait", "screenshot", "spotlight", "custom_js",
        }
        assert {a.value for a in ActionType} == expected

    def test_action_type_is_str_enum(self):
        assert isinstance(ActionType.CLICK, str)
        assert ActionType.CLICK == "click"


# ---- StepPhase Enum ----

class TestStepPhase:

    def test_all_phases_exist(self):
        assert {p.value for p in StepPhase} == {"tell_before", "show", "tell_after"}


# ---- StepAction ----

class TestStepAction:

    def test_valid_construction(self, sample_action):
        assert sample_action.action_type == ActionType.CLICK
        assert sample_action.selector == 'button[data-id="save-btn"]'
        assert sample_action.description == "Click save"

    def test_coerce_int_value_to_str(self):
        action = StepAction(
            action_type=ActionType.FILL,
            selector="input#hours",
            value=8,
            description="Enter 8 hours",
        )
        assert action.value == "8"
        assert isinstance(action.value, str)

    def test_coerce_float_value_to_str(self):
        action = StepAction(
            action_type=ActionType.FILL,
            selector="input#rate",
            value=125.50,
            description="Enter rate",
        )
        assert action.value == "125.5"
        assert isinstance(action.value, str)

    def test_none_value_stays_none(self):
        action = StepAction(
            action_type=ActionType.CLICK,
            selector="button",
            value=None,
            description="Click",
        )
        assert action.value is None

    def test_string_value_unchanged(self):
        action = StepAction(
            action_type=ActionType.FILL,
            selector="input",
            value="hello",
            description="Type text",
        )
        assert action.value == "hello"

    def test_default_delays(self):
        action = StepAction(
            action_type=ActionType.CLICK,
            selector="button",
            description="Click",
        )
        assert action.delay_before_ms == 500
        assert action.delay_after_ms == 1000

    def test_custom_delays(self):
        action = StepAction(
            action_type=ActionType.CLICK,
            selector="button",
            description="Click",
            delay_before_ms=0,
            delay_after_ms=2000,
        )
        assert action.delay_before_ms == 0
        assert action.delay_after_ms == 2000

    def test_missing_description_raises(self):
        with pytest.raises(ValidationError):
            StepAction(
                action_type=ActionType.CLICK,
                selector="button",
            )


# ---- ValueHighlight ----

class TestValueHighlight:

    def test_valid_construction(self, sample_value_highlight):
        assert sample_value_highlight.title == "Time Savings"
        assert sample_value_highlight.metric_value == "40%"

    def test_default_position(self):
        vh = ValueHighlight(
            title="Test",
            description="Test value",
        )
        assert vh.position == "top-right"

    def test_optional_metric_defaults_none(self):
        vh = ValueHighlight(
            title="Test",
            description="Test value",
        )
        assert vh.metric_value is None
        assert vh.metric_label is None


# ---- DemoStep ----

class TestDemoStep:

    def test_valid_construction(self, sample_step):
        assert sample_step.id == "step_1"
        assert sample_step.title == "Save the record"
        assert len(sample_step.actions) == 1

    def test_default_pause_after(self, sample_step):
        assert sample_step.pause_after is False

    def test_default_caption_speed(self, sample_step):
        assert sample_step.caption_speed == 25

    def test_custom_caption_speed(self):
        step = DemoStep(
            id="s",
            title="T",
            tell_before="Before",
            tell_after="After",
            caption_speed=50,
        )
        assert step.caption_speed == 50

    def test_empty_actions_default(self):
        step = DemoStep(
            id="s",
            title="T",
            tell_before="Before",
            tell_after="After",
        )
        assert step.actions == []


# ---- DemoSection ----

class TestDemoSection:

    def test_valid_construction(self, sample_section):
        assert sample_section.id == "section_1"
        assert len(sample_section.steps) == 2

    def test_optional_bpc_reference(self, sample_section):
        assert sample_section.bpc_reference is None

    def test_optional_transition_text(self, sample_section):
        assert sample_section.transition_text is None


# ---- DemoPlan ----

class TestDemoPlan:

    def test_total_steps(self, sample_plan):
        assert sample_plan.total_steps == 4

    def test_total_steps_empty_sections(self):
        plan = DemoPlan(
            id="p",
            title="Empty",
            customer_request="test",
            sections=[],
        )
        assert plan.total_steps == 0

    def test_total_steps_with_empty_section(self):
        plan = DemoPlan(
            id="p",
            title="Test",
            customer_request="test",
            sections=[
                DemoSection(id="s", title="S", description="D", steps=[]),
            ],
        )
        assert plan.total_steps == 0

    def test_all_steps_returns_flat_list(self, sample_plan):
        steps = sample_plan.all_steps
        assert len(steps) == 4
        assert [s.id for s in steps] == ["s1", "s2", "s3", "s4"]

    def test_get_step_index_found(self, sample_plan):
        assert sample_plan.get_step_index("s1") == 0
        assert sample_plan.get_step_index("s3") == 2
        assert sample_plan.get_step_index("s4") == 3

    def test_get_step_index_not_found(self, sample_plan):
        assert sample_plan.get_step_index("nonexistent") == -1

    def test_default_duration(self):
        plan = DemoPlan(
            id="p",
            title="Test",
            customer_request="test",
        )
        assert plan.estimated_duration_minutes == 15

    def test_serialization_round_trip(self, sample_plan):
        data = sample_plan.model_dump()
        restored = DemoPlan(**data)
        assert restored.title == sample_plan.title
        assert restored.total_steps == sample_plan.total_steps
        assert restored.all_steps[0].id == "s1"
        assert restored.sections[1].steps[0].value_highlight is not None
        assert restored.sections[1].steps[0].value_highlight.metric_value == "2min"
