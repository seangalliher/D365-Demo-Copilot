"""Tests for utility functions — _strip_html from script_generator and voice."""

from __future__ import annotations

from demo_agent.agent.script_generator import _strip_html as sg_strip
from demo_agent.agent.voice import _strip_html as voice_strip


class TestScriptGeneratorStripHtml:
    """Tests for script_generator._strip_html (PDF-friendly output)."""

    def test_removes_html_tags(self):
        assert sg_strip("<b>bold</b> text") == "bold text"

    def test_removes_nested_tags(self):
        assert sg_strip("<p>Hello <span class='x'>world</span></p>") == "Hello world"

    def test_decodes_amp(self):
        assert sg_strip("A &amp; B") == "A & B"

    def test_decodes_lt_gt(self):
        assert sg_strip("1 &lt; 2 &gt; 0") == "1 < 2 > 0"

    def test_decodes_quotes(self):
        assert sg_strip("&quot;hello&quot; &#039;world&#039;") == '"hello" \'world\''

    def test_converts_em_dash(self):
        # Replacement is " -- " so without surrounding spaces we get clean output
        assert sg_strip("here\u2014there") == "here -- there"

    def test_converts_ellipsis(self):
        assert sg_strip("loading\u2026") == "loading..."

    def test_strips_emoji(self):
        result = sg_strip("\U0001f4cb Plan ready")
        # Emoji should be removed, text preserved
        assert "Plan ready" in result
        assert "\U0001f4cb" not in result

    def test_empty_string(self):
        assert sg_strip("") == ""

    def test_whitespace_only(self):
        assert sg_strip("   ") == ""

    def test_plain_text_unchanged(self):
        assert sg_strip("Hello world") == "Hello world"


class TestVoiceStripHtml:
    """Tests for voice._strip_html (TTS-friendly output)."""

    def test_removes_html_tags(self):
        assert voice_strip("<b>bold</b> text") == "bold text"

    def test_decodes_entities(self):
        assert voice_strip("A &amp; B &lt; C") == "A & B < C"

    def test_converts_em_dash(self):
        # Voice version replaces \u2014 with " \u2014 " (adds surrounding spaces)
        result = voice_strip("here\u2014there")
        assert result == "here \u2014 there"

    def test_converts_ellipsis(self):
        assert voice_strip("loading\u2026") == "loading..."

    def test_empty_string(self):
        assert voice_strip("") == ""

    def test_plain_text_unchanged(self):
        assert voice_strip("Hello world") == "Hello world"
