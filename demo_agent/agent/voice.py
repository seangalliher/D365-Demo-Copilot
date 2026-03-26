"""
D365 Demo Copilot — Voice Narrator

Provides text-to-speech narration for demo presentations.
Audio is played through the browser via HTML5 Audio so the
audience hears the narration from the demo machine.

TTS Providers (resolution order):
  1. OpenAI TTS  — if VOICE_API_KEY or OPENAI_API_KEY is set
  2. Edge TTS    — free, no key needed (Microsoft Edge neural voices)

Architecture:
  Python (VoiceNarrator)
    → TTS provider → MP3 bytes
    → base64 encode → page.evaluate()
    → Browser Audio() playback
"""

from __future__ import annotations

import asyncio
import base64
import io
import logging
import os
import re
from typing import Optional

from playwright.async_api import Page

logger = logging.getLogger("demo_agent.voice")

# JavaScript injected into the browser to handle audio playback.
_VOICE_PLAYER_JS = """() => {
    if (window.__demoVoicePlayer) return;

    window.__demoVoicePlayer = {
        audio: null,

        play(b64, format) {
            return new Promise((resolve, reject) => {
                this.stop();
                const src = 'data:audio/' + (format || 'mp3') + ';base64,' + b64;
                this.audio = new Audio(src);
                this.audio.addEventListener('ended', () => {
                    this.audio = null;
                    resolve('ended');
                });
                this.audio.addEventListener('error', (e) => {
                    this.audio = null;
                    reject(e.message || 'Audio playback error');
                });
                this.audio.play().catch(err => {
                    this.audio = null;
                    reject(err.message || 'Audio play() rejected');
                });
            });
        },

        stop() {
            if (this.audio) {
                try {
                    this.audio.pause();
                    this.audio.currentTime = 0;
                    this.audio.src = '';
                } catch (e) {}
                this.audio = null;
            }
        },

        isPlaying() {
            return !!(this.audio && !this.audio.paused);
        }
    };

    console.log('[DemoVoice] Browser audio player initialized');
}"""


def _strip_html(text: str) -> str:
    """Remove HTML tags and decode entities for TTS-friendly plain text."""
    clean = re.sub(r"<[^>]+>", "", text)
    clean = clean.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
    clean = clean.replace("&quot;", '"').replace("&#039;", "'")
    clean = clean.replace("\u2014", " — ").replace("\u2026", "...")
    return clean.strip()


# ---- Edge TTS voice mapping ----
# Maps short friendly names to Microsoft Edge neural voice IDs.
EDGE_VOICE_MAP: dict[str, str] = {
    "nova": "en-US-AriaNeural",
    "alloy": "en-US-JennyNeural",
    "echo": "en-US-GuyNeural",
    "fable": "en-GB-SoniaNeural",
    "onyx": "en-US-DavisNeural",
    "shimmer": "en-US-SaraNeural",
}


async def _synthesize_edge_tts(text: str, voice: str, speed: float) -> bytes:
    """Synthesize speech using edge-tts (free, no API key)."""
    import edge_tts

    # Resolve friendly name to Edge voice ID
    edge_voice = EDGE_VOICE_MAP.get(voice, voice)
    # If the voice is already a full Edge voice ID, use it as-is
    if not edge_voice.endswith("Neural"):
        edge_voice = "en-US-AriaNeural"

    # Edge TTS rate: "+0%", "+20%", "-10%", etc.
    rate_pct = round((speed - 1.0) * 100)
    rate_str = f"+{rate_pct}%" if rate_pct >= 0 else f"{rate_pct}%"

    comm = edge_tts.Communicate(text, edge_voice, rate=rate_str)
    audio_buf = io.BytesIO()
    async for chunk in comm.stream():
        if chunk["type"] == "audio":
            audio_buf.write(chunk["data"])

    return audio_buf.getvalue()


async def _synthesize_openai_tts(
    client, text: str, model: str, voice: str, speed: float
) -> bytes:
    """Synthesize speech using OpenAI TTS API."""
    response = await client.audio.speech.create(
        model=model,
        voice=voice,
        input=text,
        speed=speed,
        response_format="mp3",
    )
    return response.content


def _create_openai_tts_client():
    """Try to create an OpenAI client for TTS. Returns None if no key available."""
    from openai import AsyncAzureOpenAI, AsyncOpenAI

    voice_key = os.getenv("VOICE_API_KEY", "")
    if voice_key:
        logger.info("TTS provider: OpenAI (VOICE_API_KEY)")
        return AsyncOpenAI(api_key=voice_key)

    voice_azure_ep = os.getenv("VOICE_AZURE_ENDPOINT", "")
    voice_azure_key = os.getenv("VOICE_AZURE_API_KEY", "")
    if voice_azure_ep and voice_azure_key:
        logger.info("TTS provider: Azure OpenAI (VOICE_AZURE_ENDPOINT)")
        return AsyncAzureOpenAI(
            azure_endpoint=voice_azure_ep,
            api_key=voice_azure_key,
            api_version=os.getenv("VOICE_AZURE_API_VERSION", "2024-10-21"),
        )

    openai_key = os.getenv("OPENAI_API_KEY", "")
    if openai_key:
        logger.info("TTS provider: OpenAI (OPENAI_API_KEY)")
        return AsyncOpenAI(api_key=openai_key)

    azure_ep = os.getenv("AZURE_OPENAI_ENDPOINT", "")
    azure_key = os.getenv("AZURE_OPENAI_API_KEY", "")
    if azure_ep and azure_key:
        logger.info("TTS provider: Azure OpenAI (AZURE_OPENAI_ENDPOINT)")
        return AsyncAzureOpenAI(
            azure_endpoint=azure_ep,
            api_key=azure_key,
            api_version=os.getenv("AZURE_OPENAI_API_VERSION", "2024-10-21"),
        )

    return None


class VoiceNarrator:
    """
    Text-to-speech narrator with browser playback.

    Uses OpenAI TTS when an API key is available, otherwise falls back
    to edge-tts (free Microsoft Edge neural voices, no key required).
    """

    def __init__(
        self,
        page: Page,
        model: str = "tts-1",
        voice: str = "nova",
        speed: float = 1.0,
        provider: str = "auto",
    ):
        self._page = page
        self.model = model
        self.voice = voice
        self.speed = speed
        self._enabled = False
        self._player_injected = False
        self._current_task: Optional[asyncio.Task] = None

        # Resolve provider
        self._openai_client = None
        if provider == "openai":
            self._openai_client = _create_openai_tts_client()
            self._provider = "openai" if self._openai_client else "edge"
        elif provider == "edge":
            self._provider = "edge"
        else:  # auto
            self._openai_client = _create_openai_tts_client()
            self._provider = "openai" if self._openai_client else "edge"

        if self._provider == "edge":
            edge_voice = EDGE_VOICE_MAP.get(voice, voice)
            logger.info("TTS provider: Edge TTS (free) — voice: %s", edge_voice)
        else:
            logger.info("TTS provider: OpenAI — model: %s, voice: %s", model, voice)

    @property
    def available(self) -> bool:
        """Voice is always available (edge-tts requires no key)."""
        return True

    @property
    def provider(self) -> str:
        """Current TTS provider: 'openai' or 'edge'."""
        return self._provider

    @property
    def enabled(self) -> bool:
        return self._enabled

    @enabled.setter
    def enabled(self, value: bool):
        self._enabled = value
        logger.info("Voice narration %s", "enabled" if value else "disabled")

    async def _ensure_player(self):
        """Inject the browser-side audio player if not already present."""
        if self._player_injected:
            try:
                exists = await self._page.evaluate(
                    "typeof window.__demoVoicePlayer !== 'undefined'"
                )
                if exists:
                    return
            except Exception:
                pass

        await self._page.evaluate(_VOICE_PLAYER_JS)
        self._player_injected = True

    async def _synthesize(self, text: str) -> bytes:
        """Synthesize speech using the configured provider."""
        if self._provider == "openai" and self._openai_client:
            return await _synthesize_openai_tts(
                self._openai_client, text, self.model, self.voice, self.speed
            )
        return await _synthesize_edge_tts(text, self.voice, self.speed)

    async def speak(self, text: str) -> bool:
        """
        Synthesize and play narration for the given text.

        Blocks until playback finishes or is stopped.
        Returns True if audio played to completion.
        """
        if not self._enabled:
            return False

        clean_text = _strip_html(text)
        if not clean_text:
            return False

        logger.info("[VOICE] Synthesizing (%s): %s", self._provider, clean_text[:80])

        try:
            audio_bytes = await self._synthesize(clean_text)
            if not audio_bytes:
                logger.warning("[VOICE] TTS returned empty audio")
                return False

            audio_b64 = base64.b64encode(audio_bytes).decode("ascii")

            await self._ensure_player()
            await self._page.evaluate(
                "(b64) => window.__demoVoicePlayer.play(b64, 'mp3')",
                audio_b64,
            )

            logger.info("[VOICE] Playback completed")
            return True

        except Exception as e:
            logger.warning("[VOICE] Speech synthesis/playback failed: %s", e)
            return False

    async def speak_async(self, text: str):
        """Start narration in the background (non-blocking)."""
        if not self._enabled:
            return

        await self.stop()
        self._current_task = asyncio.create_task(self.speak(text))

    async def stop(self):
        """Stop any currently playing narration."""
        if self._current_task and not self._current_task.done():
            self._current_task.cancel()
            try:
                await self._current_task
            except (asyncio.CancelledError, Exception):
                pass
            self._current_task = None

        try:
            await self._page.evaluate(
                "() => { if (window.__demoVoicePlayer) window.__demoVoicePlayer.stop(); }"
            )
        except Exception:
            pass

    async def wait_for_completion(self):
        """Wait for the current background speech to finish."""
        if self._current_task and not self._current_task.done():
            try:
                await self._current_task
            except (asyncio.CancelledError, Exception):
                pass
            self._current_task = None
