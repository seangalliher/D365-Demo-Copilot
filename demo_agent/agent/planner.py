"""
D365 Demo Copilot — LLM-Powered Demo Planner

Takes a customer's natural language request and generates a structured
DemoPlan with sections and steps, each following the Tell-Show-Tell pattern.
Uses BPC reference data to ground plans in real process sequences.
"""

from __future__ import annotations

import datetime
import json
import logging
import re
from pathlib import Path
from typing import Optional

from openai import AsyncAzureOpenAI, AsyncOpenAI

from ..models.demo_plan import DemoPlan

logger = logging.getLogger("demo_agent.planner")

PROMPTS_DIR = Path(__file__).parent.parent / "prompts"


class DemoPlanner:
    """
    Generates demo plans from customer requests using an LLM.

    The planner:
    1. Takes the customer's natural language request
    2. Maps it to relevant D365 Project Operations capabilities
    3. Generates a structured DemoPlan with Tell-Show-Tell steps
    4. Includes business value highlights and BPC references
    """

    def __init__(
        self,
        llm_client: AsyncAzureOpenAI | AsyncOpenAI,
        model: str = "gpt-4o",
        d365_base_url: str = "https://projectopscoreagentimplemented.crm.dynamics.com",
        schema_discovery=None,
        learn_docs=None,
    ):
        self.client = llm_client
        self.model = model
        self.d365_base_url = d365_base_url
        self.schema_discovery = schema_discovery
        self.learn_docs = learn_docs  # LearnDocsClient instance
        self._system_prompt: Optional[str] = None

    @property
    def system_prompt(self) -> str:
        """Load the planner system prompt from file."""
        if self._system_prompt is None:
            prompt_file = PROMPTS_DIR / "planner.md"
            if prompt_file.exists():
                self._system_prompt = prompt_file.read_text(encoding="utf-8")
            else:
                self._system_prompt = self._default_system_prompt()
        return self._system_prompt

    async def create_plan(
        self,
        customer_request: str,
        context: Optional[str] = None,
        max_steps: int = 15,
        schema_context: Optional[str] = None,
    ) -> DemoPlan:
        """
        Generate a demo plan from a customer request.

        Args:
            customer_request: What the customer wants to see demonstrated
            context: Optional additional context (customer industry, role, etc.)
            max_steps: Maximum number of steps to include
            schema_context: Optional Dataverse schema info for accurate selectors

        Returns:
            A structured DemoPlan ready for execution
        """
        logger.info("Generating demo plan for: %s", customer_request)

        # Auto-discover schemas via MCP if available and no context provided
        if not schema_context and self.schema_discovery:
            try:
                schemas = await self.schema_discovery.get_entity_schemas_for_request(
                    customer_request
                )
                if schemas:
                    schema_context = self.schema_discovery.format_schemas_for_prompt(schemas)
                    logger.info(
                        "Auto-discovered schema context for %d entities",
                        len(schemas),
                    )
            except Exception as e:
                logger.warning("Schema auto-discovery failed: %s", e)

        # Fetch MS Learn documentation context if available
        docs_context: Optional[str] = None
        if self.learn_docs:
            try:
                docs_context = await self.learn_docs.get_docs_for_request(
                    customer_request
                )
                if docs_context:
                    # Cap docs context to avoid exceeding model token limits
                    max_docs_chars = 8000
                    if len(docs_context) > max_docs_chars:
                        docs_context = docs_context[:max_docs_chars] + "\n\n[... truncated for brevity]"
                    logger.info(
                        "MS Learn docs context fetched (%d chars)",
                        len(docs_context),
                    )
            except Exception as e:
                logger.warning("MS Learn doc fetch failed: %s", e)

        user_message = self._build_user_message(
            customer_request, context, max_steps, schema_context, docs_context
        )

        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": user_message},
            ],
            max_completion_tokens=8000,
        )

        raw_json = response.choices[0].message.content
        logger.info(
            "LLM response — finish_reason=%s, content_length=%s",
            response.choices[0].finish_reason,
            len(raw_json) if raw_json else 0,
        )
        logger.debug("Raw plan response: %s", raw_json)

        plan_data = self._parse_json_response(raw_json)
        plan = DemoPlan(**plan_data)
        plan.d365_base_url = self.d365_base_url

        logger.info(
            "Plan generated: '%s' — %d sections, %d steps, ~%d min",
            plan.title,
            len(plan.sections),
            plan.total_steps,
            plan.estimated_duration_minutes,
        )

        return plan

    async def refine_plan(
        self,
        plan: DemoPlan,
        feedback: str,
    ) -> DemoPlan:
        """
        Refine an existing plan based on customer feedback.

        Args:
            plan: The current demo plan
            feedback: Customer's feedback or modification request

        Returns:
            A refined DemoPlan
        """
        logger.info("Refining plan based on feedback: %s", feedback)

        response = await self.client.chat.completions.create(
            model=self.model,
            messages=[
                {"role": "system", "content": self.system_prompt},
                {
                    "role": "user",
                    "content": (
                        f"Here is the current demo plan:\n\n"
                        f"```json\n{plan.model_dump_json(indent=2)}\n```\n\n"
                        f"The customer has requested this modification:\n"
                        f"{feedback}\n\n"
                        f"Generate an updated plan incorporating this feedback. "
                        f"Return the complete updated plan as JSON."
                    ),
                },
            ],
            max_completion_tokens=8000,
        )

        raw_json = response.choices[0].message.content
        logger.info(
            "Refine response — finish_reason=%s, content_length=%s",
            response.choices[0].finish_reason,
            len(raw_json) if raw_json else 0,
        )
        plan_data = self._parse_json_response(raw_json)
        refined = DemoPlan(**plan_data)
        refined.d365_base_url = self.d365_base_url

        logger.info("Plan refined: %d sections, %d steps", len(refined.sections), refined.total_steps)
        return refined

    @staticmethod
    def _parse_json_response(content: str | None) -> dict:
        """Safely extract and parse JSON from an LLM response.

        Handles:
        - None / empty content
        - Markdown ```json ... ``` fences
        - Leading/trailing whitespace
        """
        if not content or not content.strip():
            raise ValueError(
                "LLM returned empty content. The model may have refused "
                "or the response was truncated. Please try again."
            )

        text = content.strip()

        # Strip markdown code fences if present
        fence_match = re.search(
            r"```(?:json)?\s*\n?(.*?)\n?```", text, re.DOTALL
        )
        if fence_match:
            text = fence_match.group(1).strip()

        return json.loads(text)

    def _build_user_message(
        self,
        request: str,
        context: Optional[str],
        max_steps: int,
        schema_context: Optional[str] = None,
        docs_context: Optional[str] = None,
    ) -> str:
        """Build the user message for the LLM."""
        today = datetime.date.today()
        two_weeks_ago = today - datetime.timedelta(days=14)

        parts = [
            f"## Customer Demo Request\n{request}",
            f"\n## Constraints\n- Maximum {max_steps} total steps across all sections",
            f"- D365 base URL: {self.d365_base_url}",
            "- Deployment model: Resource/Non-stocked",
            "- Legal entities: Zava US (USD), Zava CA (CAD), Zava MX (MXN)",
            f"- **Today's date: {today.isoformat()} ({today.strftime('%A, %B %d, %Y')})**",
            f"- **ALL dates in demo actions MUST fall between {two_weeks_ago.isoformat()} and {today.isoformat()}** (the last two weeks). NEVER use dates from 2023 or any other year.",
            "- For option set / dropdown fields, use the **text label** (e.g., 'Work', 'Approved'), NOT the numeric option value code (e.g., '192350000').",
            "- For duration fields, use the format 'Xh Ym' (e.g., '8h 0m' for 8 hours) — NOT raw minutes.",
        ]
        if context:
            parts.insert(1, f"\n## Additional Context\n{context}")

        if schema_context:
            parts.append(f"\n{schema_context}")
            parts.append(
                "\n**IMPORTANT**: The schema above was fetched from the LIVE Dataverse environment. "
                "Use the exact logical field names from this schema when building selectors. "
                "For form fields, prefer `input[data-id=\"{logical_name}.fieldControl-text-box-text\"]` "
                "or `input[data-id=\"{logical_name}.fieldControl-duration-combobox-text\"]` patterns "
                "over `input[aria-label=\"...\"]` which can vary by language/customization."
            )

        if docs_context:
            parts.append(f"\n{docs_context}")

        parts.append(
            "\n\nGenerate a complete demo plan as JSON matching the DemoPlan schema. "
            "Ensure every step has meaningful tell_before and tell_after narration, "
            "concrete browser actions with realistic selectors, and at least 3 steps "
            "include business value highlights with quantified metrics."
        )

        return "\n".join(parts)

    @staticmethod
    def _default_system_prompt() -> str:
        """Fallback system prompt if the file isn't found."""
        return (
            "You are an expert Dynamics 365 Project Operations demo planner. "
            "Given a customer request, generate a structured demo plan in JSON format "
            "that follows the Tell-Show-Tell presentation pattern. "
            "Each step should have narration text, browser actions with CSS selectors, "
            "and optional business value highlights. "
            "Ground your plans in real D365 Project Operations functionality."
        )
