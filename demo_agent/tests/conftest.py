"""Shared fixtures for D365 Demo Copilot tests."""

from __future__ import annotations

import os

import pytest

from demo_agent.models.demo_plan import (
    ActionType,
    DemoPlan,
    DemoSection,
    DemoStep,
    StepAction,
    ValueHighlight,
)

# Environment variables that DemoConfig reads — cleared before each test
# so that no real credentials leak into the test environment.
_ENV_KEYS = [
    "D365_BASE_URL",
    "D365_FO_BASE_URL",
    "DATAVERSE_TENANT_ID",
    "DATAVERSE_CLIENT_ID",
    "DATAVERSE_CLIENT_SECRET",
    "MS_LEARN_MCP_URL",
    "AZURE_OPENAI_ENDPOINT",
    "AZURE_OPENAI_API_KEY",
    "AZURE_OPENAI_API_VERSION",
    "LLM_MODEL",
    "OPENAI_API_KEY",
    "GITHUB_TOKEN",
    "GITHUB_MODELS_BASE_URL",
    "GITHUB_COPILOT_MODEL",
    "ANTHROPIC_API_KEY",
    "BROWSER_HEADLESS",
    "BROWSER_SLOW_MO",
    "AUTH_STATE_PATH",
    "VOICE_ENABLED",
    "VOICE_PROVIDER",
    "VOICE_MODEL",
    "VOICE_NAME",
    "VOICE_SPEED",
]


@pytest.fixture(autouse=True)
def clean_env(monkeypatch):
    """Remove all D365/LLM env vars so DemoConfig tests are deterministic."""
    for key in _ENV_KEYS:
        monkeypatch.delenv(key, raising=False)


@pytest.fixture
def sample_action():
    """A minimal valid StepAction."""
    return StepAction(
        action_type=ActionType.CLICK,
        selector='button[data-id="save-btn"]',
        value=None,
        description="Click save",
    )


@pytest.fixture
def sample_step(sample_action):
    """A valid DemoStep with one action."""
    return DemoStep(
        id="step_1",
        title="Save the record",
        tell_before="Now we will save the record.",
        actions=[sample_action],
        tell_after="The record has been saved.",
    )


@pytest.fixture
def sample_value_highlight():
    """A valid ValueHighlight with a metric."""
    return ValueHighlight(
        title="Time Savings",
        description="Reduces manual data entry",
        metric_value="40%",
        metric_label="reduction in processing time",
    )


@pytest.fixture
def sample_section(sample_step):
    """A valid DemoSection with two steps."""
    step2 = DemoStep(
        id="step_2",
        title="Verify the record",
        tell_before="Let's verify the record was created.",
        actions=[],
        tell_after="The record is confirmed.",
    )
    return DemoSection(
        id="section_1",
        title="Record Management",
        description="Create and manage records",
        steps=[sample_step, step2],
    )


@pytest.fixture
def sample_plan():
    """A valid DemoPlan with 2 sections and 4 total steps."""
    section_a = DemoSection(
        id="sec_a",
        title="Project Setup",
        description="Create and configure a project",
        steps=[
            DemoStep(
                id="s1",
                title="Open Projects",
                tell_before="Navigate to the projects list.",
                actions=[
                    StepAction(
                        action_type=ActionType.NAVIGATE,
                        value="/main.aspx?pagetype=entitylist&etn=msdyn_project",
                        description="Open the projects list",
                    )
                ],
                tell_after="The projects list is now visible.",
            ),
            DemoStep(
                id="s2",
                title="Create a project",
                tell_before="Click New to create a project.",
                actions=[
                    StepAction(
                        action_type=ActionType.CLICK,
                        selector='button[data-id="edit-form-new-btn"]',
                        description="Click New",
                    )
                ],
                tell_after="A new project form is open.",
            ),
        ],
    )
    section_b = DemoSection(
        id="sec_b",
        title="Time Entry",
        description="Enter and submit time",
        steps=[
            DemoStep(
                id="s3",
                title="Enter hours",
                tell_before="Fill in the duration field.",
                actions=[
                    StepAction(
                        action_type=ActionType.FILL,
                        selector='input[data-id="msdyn_duration"]',
                        value="8",
                        description="Enter 8 hours",
                    )
                ],
                tell_after="Hours have been entered.",
                value_highlight=ValueHighlight(
                    title="Efficiency",
                    description="Quick time entry",
                    metric_value="2min",
                    metric_label="average entry time",
                ),
            ),
            DemoStep(
                id="s4",
                title="Submit entry",
                tell_before="Submit the time entry for approval.",
                actions=[],
                tell_after="Time entry submitted.",
                pause_after=True,
            ),
        ],
    )
    return DemoPlan(
        id="plan_1",
        title="Project Time Tracking",
        subtitle="End-to-end time entry demo",
        customer_request="Show me how to track project time",
        sections=[section_a, section_b],
    )
