"""
UI Helper functions for Gemini Automation
Provides image-based detection and robust UI interaction.
"""

import os
import time
import pyautogui
from typing import Optional, Tuple


# Directory for UI element screenshots
IMAGES_DIR = os.path.join(os.path.dirname(__file__), "images")


def ensure_images_dir():
    """Create images directory if it doesn't exist."""
    if not os.path.exists(IMAGES_DIR):
        os.makedirs(IMAGES_DIR)


def find_image_on_screen(
    image_name: str,
    confidence: float = 0.8,
    region: Optional[Tuple[int, int, int, int]] = None,
    grayscale: bool = True
) -> Optional[Tuple[int, int]]:
    """
    Find an image on screen and return its center coordinates.

    Args:
        image_name: Name of the image file in the images directory
        confidence: Match confidence threshold (0.0 to 1.0)
        region: Optional (x, y, width, height) to limit search area
        grayscale: Use grayscale matching for better performance

    Returns:
        (x, y) center coordinates if found, None otherwise
    """
    image_path = os.path.join(IMAGES_DIR, image_name)
    if not os.path.exists(image_path):
        print(f"Warning: Image not found: {image_path}")
        return None

    try:
        location = pyautogui.locateCenterOnScreen(
            image_path,
            confidence=confidence,
            region=region,
            grayscale=grayscale
        )
        return location
    except Exception as e:
        print(f"Image search error: {e}")
        return None


def wait_for_image(
    image_name: str,
    timeout: float = 30.0,
    confidence: float = 0.8,
    poll_interval: float = 1.0
) -> Optional[Tuple[int, int]]:
    """
    Wait for an image to appear on screen.

    Args:
        image_name: Name of the image file
        timeout: Maximum wait time in seconds
        confidence: Match confidence threshold
        poll_interval: Time between checks

    Returns:
        (x, y) coordinates if found within timeout, None otherwise
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        location = find_image_on_screen(image_name, confidence=confidence)
        if location:
            return location
        time.sleep(poll_interval)
    return None


def click_image(
    image_name: str,
    confidence: float = 0.8,
    clicks: int = 1,
    wait_after: float = 0.5
) -> bool:
    """
    Find and click on an image.

    Args:
        image_name: Name of the image file
        confidence: Match confidence threshold
        clicks: Number of clicks
        wait_after: Wait time after clicking

    Returns:
        True if clicked successfully, False otherwise
    """
    location = find_image_on_screen(image_name, confidence=confidence)
    if location:
        pyautogui.click(location[0], location[1], clicks=clicks)
        time.sleep(wait_after)
        return True
    return False


def wait_and_click_image(
    image_name: str,
    timeout: float = 30.0,
    confidence: float = 0.8,
    wait_after: float = 0.5
) -> bool:
    """
    Wait for an image to appear and then click it.

    Args:
        image_name: Name of the image file
        timeout: Maximum wait time
        confidence: Match confidence threshold
        wait_after: Wait time after clicking

    Returns:
        True if found and clicked, False if timeout
    """
    location = wait_for_image(image_name, timeout=timeout, confidence=confidence)
    if location:
        pyautogui.click(location[0], location[1])
        time.sleep(wait_after)
        return True
    return False


def safe_click(x: int, y: int, wait_after: float = 0.3) -> None:
    """Safely click at coordinates with wait."""
    pyautogui.click(x, y)
    time.sleep(wait_after)


def safe_type(text: str, interval: float = 0.05) -> None:
    """Safely type text with interval between keys."""
    pyautogui.typewrite(text, interval=interval)


def safe_hotkey(*keys, wait_after: float = 0.3) -> None:
    """Safely press hotkey combination with wait."""
    pyautogui.hotkey(*keys)
    time.sleep(wait_after)


def capture_screenshot(name: str) -> str:
    """
    Capture a screenshot and save it.

    Args:
        name: Name for the screenshot file

    Returns:
        Path to saved screenshot
    """
    ensure_images_dir()
    filename = f"{name}_{int(time.time())}.png"
    filepath = os.path.join(IMAGES_DIR, filename)
    pyautogui.screenshot(filepath)
    return filepath


def capture_region(name: str, region: Tuple[int, int, int, int]) -> str:
    """
    Capture a specific region of the screen.

    Args:
        name: Name for the screenshot file
        region: (x, y, width, height) region to capture

    Returns:
        Path to saved screenshot
    """
    ensure_images_dir()
    filename = f"{name}.png"
    filepath = os.path.join(IMAGES_DIR, filename)
    pyautogui.screenshot(filepath, region=region)
    return filepath


class RetryHandler:
    """Handle retries with exponential backoff."""

    def __init__(self, max_retries: int = 3, base_delay: float = 1.0):
        self.max_retries = max_retries
        self.base_delay = base_delay

    def execute(self, func, *args, **kwargs):
        """
        Execute a function with retries.

        Args:
            func: Function to execute
            *args, **kwargs: Arguments to pass to function

        Returns:
            Function result if successful

        Raises:
            Exception: Last exception if all retries fail
        """
        last_exception = None
        for attempt in range(self.max_retries):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                last_exception = e
                delay = self.base_delay * (2 ** attempt)
                print(f"Attempt {attempt + 1} failed: {e}. Retrying in {delay}s...")
                time.sleep(delay)
        raise last_exception


# Initialize images directory on import
ensure_images_dir()
