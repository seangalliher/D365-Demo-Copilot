"""
D365 Demo Copilot — Demo Plan Data Models

Defines the structure for dynamically generated demo plans.
Each plan consists of sections (logical groupings) containing steps.
Each step follows the Tell-Show-Tell pattern with optional business
value highlights and specific browser actions.
"""

from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field, field_validator


class StepPhase(str, Enum):
    """The Tell-Show-Tell phase of a demo step."""
    TELL_BEFORE = "tell_before"   # Explain what we're about to show
    SHOW = "show"                 # Perform the action in the UI
    TELL_AFTER = "tell_after"     # Summarize what was demonstrated


class ActionType(str, Enum):
    """Types of browser actions the agent can perform."""
    NAVIGATE = "navigate"           # Navigate to a URL or page
    CLICK = "click"                 # Click an element
    FILL = "fill"                   # Type text into a field
    SELECT = "select"               # Select a dropdown option
    HOVER = "hover"                 # Hover over an element
    SCROLL = "scroll"               # Scroll the page or element
    WAIT = "wait"                   # Wait for an element to appear
    SCREENSHOT = "screenshot"       # Take a screenshot
    SPOTLIGHT = "spotlight"         # Highlight an element (no click)
    CUSTOM_JS = "custom_js"         # Execute custom JavaScript


class StepAction(BaseModel):
    """A single browser action within a demo step."""
    action_type: ActionType = Field(description="The type of browser action to perform")
    selector: Optional[str] = Field(
        default=None,
        description="CSS/XPath selector or aria label for the target element"
    )
    value: Optional[str] = Field(
        default=None,
        description="Value for fill/select actions, URL for navigate, JS for custom_js"
    )

    @field_validator("value", mode="before")
    @classmethod
    def _coerce_value_to_str(cls, v: object) -> Optional[str]:
        """LLMs sometimes return ints/floats for numeric fields — coerce to str."""
        if v is None:
            return None
        return str(v)
    description: str = Field(
        description="Human-readable description of what this action does"
    )
    tooltip: Optional[str] = Field(
        default=None,
        description="Optional tooltip text to show on the element during this action"
    )
    delay_before_ms: int = Field(
        default=500,
        description="Milliseconds to wait before performing this action"
    )
    delay_after_ms: int = Field(
        default=1000,
        description="Milliseconds to wait after performing this action"
    )


class ValueHighlight(BaseModel):
    """Business value callout to display during or after a step."""
    title: str = Field(description="Card title, e.g. 'Time Savings'")
    description: str = Field(description="Value description text")
    metric_value: Optional[str] = Field(
        default=None,
        description="Quantified metric, e.g. '40%' or '$2.3M'"
    )
    metric_label: Optional[str] = Field(
        default=None,
        description="Metric label, e.g. 'reduction in manual entry'"
    )
    position: str = Field(
        default="top-right",
        description="Card position: top-right, top-left, bottom-right, bottom-left, center-right"
    )


class DemoStep(BaseModel):
    """
    A single demo step following the Tell-Show-Tell pattern.

    Flow:
    1. tell_before — Caption text explaining what will be shown
    2. actions[]   — Browser actions to perform (with spotlight/tooltips)
    3. tell_after  — Caption text summarizing what was demonstrated
    4. value       — Optional business value callout card
    """
    id: str = Field(description="Unique step identifier, e.g. 'create_project_1'")
    title: str = Field(description="Short step title for progress indicator")
    tell_before: str = Field(
        description="Narration text shown BEFORE performing the action (Tell phase 1)"
    )
    actions: list[StepAction] = Field(
        default_factory=list,
        description="Ordered list of browser actions to perform (Show phase)"
    )
    tell_after: str = Field(
        description="Narration text shown AFTER performing the action (Tell phase 2)"
    )
    value_highlight: Optional[ValueHighlight] = Field(
        default=None,
        description="Optional business value callout displayed after the step"
    )
    pause_after: bool = Field(
        default=False,
        description="Whether to pause for user confirmation after this step"
    )
    caption_speed: int = Field(
        default=25,
        description="Typewriter speed in ms per character for caption animation"
    )


class DemoSection(BaseModel):
    """
    A logical grouping of demo steps (e.g. 'Project Creation', 'Time Entry').

    Sections provide narrative structure and can map to BPC process areas.
    """
    id: str = Field(description="Unique section identifier")
    title: str = Field(description="Section heading, e.g. 'Creating a New Project'")
    description: str = Field(description="Brief description of what this section covers")
    bpc_reference: Optional[str] = Field(
        default=None,
        description="BPC process sequence ID, e.g. '80.40.010'"
    )
    steps: list[DemoStep] = Field(
        default_factory=list,
        description="Ordered list of demo steps in this section"
    )
    transition_text: Optional[str] = Field(
        default=None,
        description="Text to show when transitioning to the next section"
    )


class DemoPlan(BaseModel):
    """
    A complete demo plan generated from a customer request.

    The plan contains:
    - Metadata (title, audience, duration estimate)
    - Sections with ordered steps
    - Each step follows Tell-Show-Tell pattern
    """
    id: str = Field(description="Unique plan identifier")
    title: str = Field(description="Demo title, e.g. 'Project Time Tracking in D365'")
    subtitle: Optional[str] = Field(
        default=None,
        description="Demo subtitle or audience context"
    )
    customer_request: str = Field(
        description="The original customer request that generated this plan"
    )
    estimated_duration_minutes: int = Field(
        default=15,
        description="Estimated demo duration in minutes"
    )
    d365_base_url: str = Field(
        default="https://projectopscoreagentimplemented.crm.dynamics.com",
        description="Base URL of the D365 environment"
    )
    sections: list[DemoSection] = Field(
        default_factory=list,
        description="Ordered list of demo sections"
    )
    closing_text: Optional[str] = Field(
        default=None,
        description="Closing narration text after the demo completes"
    )

    @property
    def total_steps(self) -> int:
        """Total number of steps across all sections."""
        return sum(len(s.steps) for s in self.sections)

    @property
    def all_steps(self) -> list[DemoStep]:
        """Flat list of all steps in order."""
        steps = []
        for section in self.sections:
            steps.extend(section.steps)
        return steps

    def get_step_index(self, step_id: str) -> int:
        """Get the global index of a step by its ID."""
        for i, step in enumerate(self.all_steps):
            if step.id == step_id:
                return i
        return -1
