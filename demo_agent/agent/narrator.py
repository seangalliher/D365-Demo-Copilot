"""
D365 Demo Copilot — Narrator

Generates contextual narration text for Tell-Show-Tell phases.
Can produce dynamic narration based on what's actually visible on
the D365 page, not just pre-scripted text.
"""

from __future__ import annotations

import logging
from typing import Optional

from openai import AsyncAzureOpenAI, AsyncOpenAI

logger = logging.getLogger("demo_agent.narrator")


class Narrator:
    """
    Generates dynamic narration text for demo steps.

    Used when the planner's static text needs to be supplemented with
    context-aware narration (e.g., referencing actual data values on screen).
    """

    def __init__(
        self,
        llm_client: AsyncAzureOpenAI | AsyncOpenAI,
        model: str = "gpt-4o",
    ):
        self.client = llm_client
        self.model = model

    async def generate_tell_before(
        self,
        step_title: str,
        step_actions_summary: str,
        section_context: str = "",
        page_context: str = "",
    ) -> str:
        """
        Generate the 'Tell Before' narration for a step.

        This text explains what the user is about to see.
        """
        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are narrating a live Dynamics 365 demonstration. "
                        "Generate a concise, engaging 1-2 sentence 'Tell' narration "
                        "that explains what the audience is about to see. "
                        "Use present tense. Be specific about the D365 capability. "
                        "Do NOT use markdown formatting — this is displayed as a caption overlay."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Step: {step_title}\n"
                        f"Actions: {step_actions_summary}\n"
                        f"Section context: {section_context}\n"
                        f"Current page: {page_context}\n\n"
                        f"Generate the 'Tell Before' narration."
                    ),
                },
            ],
            max_completion_tokens=150,
        )
        return response.choices[0].message.content.strip()

    async def generate_tell_after(
        self,
        step_title: str,
        what_was_shown: str,
        business_value: Optional[str] = None,
    ) -> str:
        """
        Generate the 'Tell After' narration for a step.

        This text summarizes what was just demonstrated and optionally
        connects it to business value.
        """
        value_context = ""
        if business_value:
            value_context = f"\nBusiness value to emphasize: {business_value}"

        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are narrating a live Dynamics 365 demonstration. "
                        "Generate a concise 1-2 sentence summary of what was just "
                        "demonstrated. Connect it to business outcomes when possible. "
                        "Use past tense for what happened, present tense for value. "
                        "Do NOT use markdown formatting — this is displayed as a caption overlay."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Step: {step_title}\n"
                        f"What was shown: {what_was_shown}\n"
                        f"{value_context}\n\n"
                        f"Generate the 'Tell After' narration."
                    ),
                },
            ],
            max_completion_tokens=150,
        )
        return response.choices[0].message.content.strip()

    async def generate_section_transition(
        self,
        completed_section: str,
        next_section: str,
        next_section_description: str,
    ) -> str:
        """
        Generate transition text between demo sections.
        """
        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are narrating a live Dynamics 365 demonstration. "
                        "Generate a smooth 1-2 sentence transition between demo sections. "
                        "Briefly acknowledge what was covered and preview what's next. "
                        "Keep it natural and engaging."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Just completed: {completed_section}\n"
                        f"Moving to: {next_section}\n"
                        f"Description: {next_section_description}\n\n"
                        f"Generate the transition narration."
                    ),
                },
            ],
            max_completion_tokens=120,
        )
        return response.choices[0].message.content.strip()

    async def generate_closing(
        self,
        demo_title: str,
        sections_covered: list[str],
        key_values: list[str],
    ) -> str:
        """
        Generate closing narration for the demo.
        """
        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are narrating a live Dynamics 365 demonstration. "
                        "Generate a compelling 2-3 sentence closing that summarizes "
                        "what was demonstrated and reinforces key business value. "
                        "End with an invitation for questions or deeper exploration."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"Demo: {demo_title}\n"
                        f"Sections covered: {', '.join(sections_covered)}\n"
                        f"Key values highlighted: {', '.join(key_values)}\n\n"
                        f"Generate the closing narration."
                    ),
                },
            ],
            max_completion_tokens=200,
        )
        return response.choices[0].message.content.strip()
