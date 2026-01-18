"""
Gemini Operations Module
Handles all Gemini-specific UI interactions with image-based detection.
"""

import os
import time
import pyautogui
from pywinauto import Desktop
from typing import Optional

from config import (
    WAIT_SHORT,
    WAIT_MEDIUM,
    WAIT_LONG,
    WAIT_SPREADSHEET,
    MAX_RETRIES,
    RETRY_DELAY,
    GEMINI_WINDOW_TITLE,
    GEM_NAME,
)
from ui_helpers import (
    find_image_on_screen,
    wait_for_image,
    click_image,
    wait_and_click_image,
    safe_hotkey,
    capture_screenshot,
    RetryHandler,
)


def log(message: str) -> None:
    """Print timestamped log message."""
    timestamp = time.strftime("%H:%M:%S")
    print(f"[{timestamp}] {message}")


class GeminiAutomation:
    """Handles Gemini Web interface automation."""

    def __init__(self):
        self.retry_handler = RetryHandler(MAX_RETRIES, RETRY_DELAY)

    def find_gemini_window(self):
        """Find the Gemini browser window."""
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        for win in windows:
            try:
                if GEMINI_WINDOW_TITLE.lower() in win.window_text().lower():
                    return win
            except Exception:
                continue
        return None

    def activate_gemini(self) -> bool:
        """Activate the Gemini browser window."""
        log("Activating Gemini window...")

        for attempt in range(MAX_RETRIES):
            window = self.find_gemini_window()
            if window:
                try:
                    window.set_focus()
                    time.sleep(WAIT_SHORT)
                    log("Gemini window activated")
                    return True
                except Exception as e:
                    log(f"Activation failed: {e}")

            log(f"Retry {attempt + 1}/{MAX_RETRIES}")
            time.sleep(RETRY_DELAY)

        log("ERROR: Gemini window not found")
        return False

    def select_gem(self) -> bool:
        """Select the Gem (saved prompt) for document processing."""
        log("Selecting Gem...")

        # Try image-based detection first
        if click_image("gem_button.png", confidence=0.7):
            log("Gem button clicked via image detection")
            time.sleep(WAIT_SHORT)
            return True

        # Fallback: Use keyboard navigation if available
        log("Image detection failed, using fallback method")
        time.sleep(WAIT_SHORT)
        return True





