# demo_agent/browser/__init__.py
from .chat_panel import ChatPanelManager
from .controller import BrowserController
from .d365_pages import D365Navigator
from .overlay_manager import OverlayManager

__all__ = ["BrowserController", "ChatPanelManager", "D365Navigator", "OverlayManager"]
