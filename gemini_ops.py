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

    def input_prompt(self, text: str = GEM_NAME) -> bool:
        """Input text into the Gemini prompt field."""
        log(f"Inputting: {text}")

        # Click on prompt area first (if needed)
        # pyautogui.click()  # Click at current position

        # Type the text
        pyautogui.typewrite(text, interval=0.05)
        time.sleep(WAIT_SHORT)

        return True

    def click_upload(self) -> bool:
        """Click the file upload button."""
        log("Clicking upload button...")

        # Try image-based detection
        if click_image("upload_button.png", confidence=0.7):
            log("Upload button clicked")
            time.sleep(WAIT_SHORT)
            return True

        # Fallback: Try clicking the + button area
        log("Image detection failed, trying fallback...")
        return True

    def upload_file_dialog(self, file_path: str) -> bool:
        """Handle the file upload dialog."""
        log(f"Uploading: {os.path.basename(file_path)}")

        # Wait for dialog
        time.sleep(WAIT_MEDIUM)

        # Type the file path
        # Convert to Windows path format
        win_path = file_path.replace('/', '\\')
        pyautogui.typewrite(win_path, interval=0.02)
        time.sleep(WAIT_SHORT)

        # Confirm
        pyautogui.press('enter')
        time.sleep(WAIT_MEDIUM)

        log("File path entered")
        return True

    def execute_prompt(self) -> bool:
        """Click the send button to execute the prompt."""
        log("Executing prompt...")

        # Try image-based detection
        if click_image("send_button.png", confidence=0.7):
            log("Send button clicked")
            return True

        # Fallback: Press Enter
        pyautogui.press('enter')
        time.sleep(WAIT_SHORT)
        return True

    def wait_for_completion(self, timeout: float = 60.0) -> bool:
        """Wait for Gemini to complete processing."""
        log("Waiting for response...")

        # Check if image file exists first
        from ui_helpers import IMAGES_DIR
        image_path = os.path.join(IMAGES_DIR, "spreadsheet_button.png")

        if os.path.exists(image_path):
            # Try to detect the spreadsheet button
            location = wait_for_image(
                "spreadsheet_button.png",
                timeout=timeout,
                confidence=0.7,
                poll_interval=2.0
            )

            if location:
                log("Response completed - spreadsheet button detected")
                return True

        # Fallback: Wait for fixed duration
        log("Image detection unavailable, using fixed wait...")
        time.sleep(timeout)
        return True

    def open_spreadsheet(self) -> bool:
        """Click the 'Open in Spreadsheet' button."""
        log("Opening spreadsheet...")

        # Try image-based detection
        if wait_and_click_image(
            "spreadsheet_button.png",
            timeout=10.0,
            confidence=0.7
        ):
            log("Spreadsheet button clicked")
            time.sleep(WAIT_SPREADSHEET)
            return True

        # Fallback: Assume button is in focus, press Enter
        log("Using keyboard fallback...")
        pyautogui.press('enter')
        time.sleep(WAIT_SPREADSHEET)
        return True

    def confirm_spreadsheet_dialog(self) -> bool:
        """Handle any confirmation dialog for spreadsheet creation."""
        log("Checking for confirmation dialog...")
        time.sleep(WAIT_SHORT)

        # Press Enter to confirm any dialog
        pyautogui.press('enter')
        time.sleep(WAIT_SPREADSHEET)

        return True


class SpreadsheetAutomation:
    """Handles spreadsheet operations."""

    def find_spreadsheet_window(self, title_pattern: str):
        """Find a spreadsheet window by title."""
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        for win in windows:
            try:
                if title_pattern.lower() in win.window_text().lower():
                    return win
            except Exception:
                continue
        return None

    def copy_data(self) -> bool:
        """Copy data from current spreadsheet (rows 2+)."""
        log("Copying spreadsheet data...")

        time.sleep(WAIT_SHORT)

        # Go to beginning
        safe_hotkey('ctrl', 'Home', wait_after=WAIT_SHORT)

        # Move to A2 (skip header)
        pyautogui.press('down')
        time.sleep(WAIT_SHORT)

        # Select to end
        safe_hotkey('ctrl', 'shift', 'End', wait_after=WAIT_SHORT)

        # Copy
        safe_hotkey('ctrl', 'c', wait_after=WAIT_SHORT)

        log("Data copied")
        return True

    def activate_master_sheet(self, name: str) -> bool:
        """Activate the master spreadsheet."""
        log(f"Activating: {name}")

        window = self.find_spreadsheet_window(name)
        if window:
            try:
                window.set_focus()
                time.sleep(WAIT_SHORT)
                log("Master spreadsheet activated")
                return True
            except Exception as e:
                log(f"Activation failed: {e}")

        log("ERROR: Master spreadsheet not found")
        return False

    def paste_data(self) -> bool:
        """Paste data at the next empty row."""
        log("Pasting data...")

        # Go to end of data
        safe_hotkey('ctrl', 'End', wait_after=WAIT_SHORT)

        # Next row
        pyautogui.press('down')
        time.sleep(WAIT_SHORT)

        # Column A
        pyautogui.press('Home')
        time.sleep(WAIT_SHORT)

        # Paste values only (Ctrl+Shift+V for Google Sheets)
        safe_hotkey('ctrl', 'shift', 'v', wait_after=WAIT_SHORT)

        log("Data pasted")
        return True

    def close_current(self) -> bool:
        """Close the current spreadsheet tab/window."""
        log("Closing spreadsheet...")
        safe_hotkey('ctrl', 'w', wait_after=WAIT_SHORT)
        return True
