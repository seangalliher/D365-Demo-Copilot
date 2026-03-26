"""
D365 Demo Copilot — Script Generator

Generates a professional PDF demo script from captured step data.
Uses fpdf2 (pure Python PDF generation) to create a document with:
  - Title page
  - Section headers
  - Step pages with narration text and screenshots
  - Summary page
"""

from __future__ import annotations

import io
import logging
import re
from datetime import datetime
from typing import Optional

from fpdf import FPDF

from ..models.demo_plan import DemoPlan
from .script_recorder import StepCapture

logger = logging.getLogger("demo_agent.script_generator")

# ---- Color palette (Microsoft brand-aligned) ----
_BLUE = (0, 120, 212)       # #0078D4 — primary accent
_DARK = (30, 30, 30)        # #1E1E1E — title page background
_SECTION_BG = (37, 37, 38)  # #252526 — section header bar
_WHITE = (255, 255, 255)
_BODY = (40, 40, 40)        # Body text
_MUTED = (120, 120, 120)    # Labels and secondary text
_LIGHT_BG = (245, 245, 245) # Narration box background
_VALUE_BG = (230, 244, 255) # Value highlight box background
_BORDER = (220, 220, 220)   # Screenshot border


def _strip_html(text: str) -> str:
    """Remove HTML tags and decode entities for PDF-friendly plain text."""
    clean = re.sub(r"<[^>]+>", "", text)
    clean = clean.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    clean = clean.replace("&quot;", '"').replace("&#039;", "'")
    clean = clean.replace("\u2014", " -- ").replace("\u2013", "-").replace("\u2026", "...")
    # Replace smart quotes with ASCII equivalents
    clean = clean.replace("\u2018", "'").replace("\u2019", "'")
    clean = clean.replace("\u201c", '"').replace("\u201d", '"')
    # Strip emoji (fpdf2 Helvetica doesn't support them)
    clean = re.sub(
        r"[\U0001F300-\U0001F9FF\U0000200D\U00002600-\U000026FF\U00002700-\U000027BF]",
        "", clean,
    )
    return clean.strip()


class ScriptGenerator:
    """Generates a PDF demo script from captured step data."""

    def generate(
        self,
        plan: DemoPlan,
        captures: list[StepCapture],
        elapsed_display: str,
    ) -> bytes:
        """Generate the PDF and return raw PDF bytes."""
        pdf = FPDF(orientation="P", unit="mm", format="A4")
        pdf.set_auto_page_break(auto=True, margin=20)

        # Title page
        self._add_title_page(pdf, plan)

        # Step pages grouped by section
        current_section_id = None
        section_number = 0

        for capture in captures:
            # Section header when section changes
            if capture.section_id != current_section_id:
                current_section_id = capture.section_id
                section_number += 1
                self._add_section_header(
                    pdf,
                    capture.section_title,
                    capture.section_description,
                    section_number,
                )

            # Step content
            self._add_step(pdf, capture)

        # Summary page
        self._add_summary_page(pdf, plan, captures, elapsed_display)

        return pdf.output()

    # ---- Title Page ----

    def _add_title_page(self, pdf: FPDF, plan: DemoPlan) -> None:
        pdf.add_page()

        # Dark header bar
        pdf.set_fill_color(*_DARK)
        pdf.rect(0, 0, 210, 80, "F")

        # Brand line
        pdf.set_text_color(*_BLUE)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_y(20)
        pdf.cell(0, 6, "Dynamics 365 Demo Script", align="C", new_x="LMARGIN", new_y="NEXT")

        # Title
        pdf.set_text_color(*_WHITE)
        pdf.set_font("Helvetica", "B", 24)
        pdf.ln(6)
        title = _strip_html(plan.title)
        pdf.multi_cell(0, 12, title, align="C")

        # Subtitle
        if plan.subtitle:
            pdf.set_text_color(180, 180, 180)
            pdf.set_font("Helvetica", "", 12)
            pdf.ln(2)
            pdf.multi_cell(0, 7, _strip_html(plan.subtitle), align="C")

        # Metadata section (below dark bar)
        pdf.set_y(95)
        pdf.set_text_color(*_MUTED)
        pdf.set_font("Helvetica", "", 10)

        meta_items = [
            ("Date", datetime.now().strftime("%B %d, %Y")),
            ("Estimated Duration", f"{plan.estimated_duration_minutes} minutes"),
            ("Total Steps", str(plan.total_steps)),
        ]

        for label, value in meta_items:
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(*_MUTED)
            pdf.cell(40, 7, f"{label}:", new_x="END")
            pdf.set_font("Helvetica", "", 10)
            pdf.set_text_color(*_BODY)
            pdf.cell(0, 7, value, new_x="LMARGIN", new_y="NEXT")

        # Customer request
        if plan.customer_request:
            pdf.ln(6)
            pdf.set_font("Helvetica", "B", 10)
            pdf.set_text_color(*_MUTED)
            pdf.cell(0, 7, "Customer Request:", new_x="LMARGIN", new_y="NEXT")
            pdf.set_font("Helvetica", "I", 10)
            pdf.set_text_color(*_BODY)
            request_text = _strip_html(plan.customer_request)
            if len(request_text) > 300:
                request_text = request_text[:297] + "..."
            pdf.multi_cell(0, 6, f'"{request_text}"')

    # ---- Section Header ----

    def _add_section_header(
        self, pdf: FPDF, title: str, description: str, section_number: int
    ) -> None:
        pdf.add_page()

        # Blue accent bar
        pdf.set_fill_color(*_BLUE)
        pdf.rect(10, pdf.get_y(), 3, 18, "F")

        # Section title
        pdf.set_x(18)
        pdf.set_font("Helvetica", "B", 16)
        pdf.set_text_color(*_BODY)
        pdf.cell(
            0, 10,
            f"Section {section_number}: {_strip_html(title)}",
            new_x="LMARGIN", new_y="NEXT",
        )

        # Description
        pdf.set_x(18)
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*_MUTED)
        pdf.multi_cell(0, 6, _strip_html(description))

        # Separator
        pdf.ln(4)
        pdf.set_draw_color(*_BORDER)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(6)

    # ---- Step Content ----

    def _add_step(self, pdf: FPDF, capture: StepCapture) -> None:
        # Check if we need a new page (at least 60mm needed for step header + narration)
        if pdf.get_y() > 220:
            pdf.add_page()

        # Step header with blue accent
        y_start = pdf.get_y()
        pdf.set_fill_color(*_BLUE)
        pdf.rect(10, y_start, 2, 8, "F")

        pdf.set_x(16)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(*_BLUE)
        step_label = f"Step {capture.step_number}: {_strip_html(capture.step_title)}"
        if capture.skipped:
            step_label += " (Skipped)"
        pdf.cell(0, 8, step_label, new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

        # Narration: tell_before
        self._add_text_box(pdf, "NARRATION", capture.tell_before, _LIGHT_BG)

        # Screenshot
        if capture.screenshot_png and not capture.skipped:
            self._add_screenshot(pdf, capture.screenshot_png)

        # Summary: tell_after
        self._add_text_box(pdf, "SUMMARY", capture.tell_after, _LIGHT_BG)

        # Business value highlight
        if capture.value_highlight:
            self._add_value_box(pdf, capture.value_highlight)

        # Spacing between steps
        pdf.ln(6)

    def _add_text_box(
        self, pdf: FPDF, label: str, text: str, bg_color: tuple
    ) -> None:
        """Add a labeled text box with background fill."""
        if pdf.get_y() > 260:
            pdf.add_page()

        x = pdf.l_margin
        w = pdf.w - pdf.l_margin - pdf.r_margin

        # Label
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(*_MUTED)
        pdf.cell(0, 5, label, new_x="LMARGIN", new_y="NEXT")

        # Box background
        y_before = pdf.get_y()
        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(*_BODY)

        clean_text = _strip_html(text)
        # Calculate height by writing to a temporary position
        # Use multi_cell with dry_run to get height
        line_height = 6
        text_height = pdf.multi_cell(
            w - 8, line_height, clean_text, dry_run=True, output="HEIGHT"
        )
        box_height = text_height + 6  # padding

        # Check if box fits on page
        if y_before + box_height > pdf.h - pdf.b_margin:
            pdf.add_page()
            y_before = pdf.get_y()

        # Draw background
        pdf.set_fill_color(*bg_color)
        pdf.rect(x, y_before, w, box_height, "F")

        # Write text inside box
        pdf.set_xy(x + 4, y_before + 3)
        pdf.multi_cell(w - 8, line_height, clean_text)
        pdf.set_y(y_before + box_height + 2)

    def _add_screenshot(self, pdf: FPDF, png_bytes: bytes) -> None:
        """Add a screenshot image scaled to page width."""
        if not png_bytes:
            return

        # Check available space — if less than 50mm, start new page
        if pdf.get_y() > 200:
            pdf.add_page()

        available_width = pdf.w - pdf.l_margin - pdf.r_margin

        try:
            img_stream = io.BytesIO(png_bytes)
            y_before = pdf.get_y()

            # Draw border
            pdf.set_draw_color(*_BORDER)

            pdf.image(
                img_stream,
                x=pdf.l_margin,
                y=y_before,
                w=available_width,
            )

            # Get the actual image height as rendered
            img_height = pdf.get_y() - y_before
            if img_height <= 0:
                # Fallback: estimate based on typical 16:10 ratio
                img_height = available_width * 0.625

            # Draw border rectangle around the image
            pdf.rect(pdf.l_margin, y_before, available_width, img_height)

            pdf.set_y(y_before + img_height + 3)

        except Exception as e:
            logger.warning("[SCRIPT] Failed to embed screenshot: %s", e)
            pdf.set_font("Helvetica", "I", 9)
            pdf.set_text_color(*_MUTED)
            pdf.cell(0, 6, "[Screenshot unavailable]", new_x="LMARGIN", new_y="NEXT")

    def _add_value_box(self, pdf: FPDF, vh: dict) -> None:
        """Add a business value highlight box."""
        if pdf.get_y() > 250:
            pdf.add_page()

        x = pdf.l_margin
        w = pdf.w - pdf.l_margin - pdf.r_margin

        # Calculate content height
        pdf.set_font("Helvetica", "B", 10)
        title_text = _strip_html(vh.get("title", ""))
        desc_text = _strip_html(vh.get("description", ""))

        desc_height = pdf.multi_cell(
            w - 8, 5, desc_text, dry_run=True, output="HEIGHT"
        )
        box_height = 8 + desc_height + 6  # title + desc + padding
        if vh.get("metric_value"):
            box_height += 10

        y_before = pdf.get_y()
        if y_before + box_height > pdf.h - pdf.b_margin:
            pdf.add_page()
            y_before = pdf.get_y()

        # Blue-tinted background
        pdf.set_fill_color(*_VALUE_BG)
        pdf.set_draw_color(*_BLUE)
        pdf.rect(x, y_before, w, box_height, "DF")

        # Left accent line
        pdf.set_fill_color(*_BLUE)
        pdf.rect(x, y_before, 3, box_height, "F")

        # Title
        pdf.set_xy(x + 6, y_before + 3)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*_BLUE)
        pdf.cell(0, 6, f"Business Value: {title_text}", new_x="LMARGIN", new_y="NEXT")

        # Description
        pdf.set_x(x + 6)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*_BODY)
        pdf.multi_cell(w - 12, 5, desc_text)

        # Metric
        if vh.get("metric_value"):
            pdf.set_x(x + 6)
            pdf.set_font("Helvetica", "B", 14)
            pdf.set_text_color(*_BLUE)
            metric_str = vh["metric_value"]
            if vh.get("metric_label"):
                metric_str += f"  {vh['metric_label']}"
            pdf.cell(0, 8, metric_str, new_x="LMARGIN", new_y="NEXT")

        pdf.set_y(y_before + box_height + 3)

    # ---- Summary Page ----

    def _add_summary_page(
        self,
        pdf: FPDF,
        plan: DemoPlan,
        captures: list[StepCapture],
        elapsed_display: str,
    ) -> None:
        pdf.add_page()

        # Header
        pdf.set_font("Helvetica", "B", 20)
        pdf.set_text_color(*_BODY)
        pdf.cell(0, 12, "Demo Summary", new_x="LMARGIN", new_y="NEXT")

        # Separator
        pdf.set_draw_color(*_BLUE)
        pdf.set_line_width(0.5)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.set_line_width(0.2)
        pdf.ln(8)

        # Stats
        completed = sum(1 for c in captures if not c.skipped)
        skipped = sum(1 for c in captures if c.skipped)

        stats = [
            ("Steps Completed", str(completed)),
            ("Steps Skipped", str(skipped)),
            ("Elapsed Time", elapsed_display),
        ]

        for label, value in stats:
            pdf.set_font("Helvetica", "B", 11)
            pdf.set_text_color(*_MUTED)
            pdf.cell(50, 8, f"{label}:", new_x="END")
            pdf.set_font("Helvetica", "", 11)
            pdf.set_text_color(*_BODY)
            pdf.cell(0, 8, value, new_x="LMARGIN", new_y="NEXT")

        # Sections covered
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 12)
        pdf.set_text_color(*_BODY)
        pdf.cell(0, 8, "Sections Covered:", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

        seen_sections: set[str] = set()
        section_num = 0
        for capture in captures:
            if capture.section_id not in seen_sections:
                seen_sections.add(capture.section_id)
                section_num += 1
                pdf.set_font("Helvetica", "", 10)
                pdf.set_text_color(*_BODY)
                section_title = _strip_html(capture.section_title)
                pdf.cell(
                    0, 7,
                    f"  {section_num}. {section_title}",
                    new_x="LMARGIN", new_y="NEXT",
                )

        # Footer
        pdf.ln(12)
        pdf.set_draw_color(*_BORDER)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(6)
        pdf.set_font("Helvetica", "I", 9)
        pdf.set_text_color(*_MUTED)
        pdf.cell(0, 5, "Generated by D365 Demo Copilot", new_x="LMARGIN", new_y="NEXT")
        pdf.cell(
            0, 5,
            datetime.now().strftime("%B %d, %Y at %I:%M %p"),
            new_x="LMARGIN", new_y="NEXT",
        )
