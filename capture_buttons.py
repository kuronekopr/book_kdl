"""
Button Capture Tool
Use this to capture UI elements for image-based automation.

Usage:
1. Run this script
2. Position your mouse over the target button
3. Press Enter to capture
4. The script will save a small region around the cursor
"""

import os
import time
import pyautogui
from ui_helpers import IMAGES_DIR, ensure_images_dir


def capture_button(name: str, size: int = 100) -> None:
    """
    Capture a button image at current mouse position.

    Args:
        name: Name for the button image
        size: Size of the capture region (square)
    """
    ensure_images_dir()

    # Get current mouse position
    x, y = pyautogui.position()

    # Calculate region (centered on mouse)
    half_size = size // 2
    region = (x - half_size, y - half_size, size, size)

    # Capture
    filepath = os.path.join(IMAGES_DIR, f"{name}.png")
    pyautogui.screenshot(filepath, region=region)

    print(f"Saved: {filepath}")
    print(f"Region: {region}")


def interactive_capture():
    """Interactive mode for capturing multiple buttons."""
    print("=" * 50)
    print("Button Capture Tool")
    print("=" * 50)
    print()
    print("Instructions:")
    print("1. Position your mouse over the target button")
    print("2. Enter a name for the button when prompted")
    print("3. The script will capture a 100x100 region")
    print("4. Type 'quit' to exit")
    print()

    buttons_to_capture = [
        ("gem_button", "Gem selection button (consult-style summary)"),
        ("upload_button", "Upload/+ button for adding files"),
        ("send_button", "Send/arrow button to execute"),
        ("spreadsheet_button", "Open in Spreadsheet button"),
    ]

    print("Suggested buttons to capture:")
    for name, desc in buttons_to_capture:
        print(f"  - {name}: {desc}")
    print()

    while True:
        name = input("Enter button name (or 'quit' to exit): ").strip()
        if name.lower() == 'quit':
            break

        if not name:
            print("Please enter a valid name")
            continue

        print(f"Position mouse over '{name}' and press Enter...")
        input()

        # Small delay to ensure mouse is in position
        time.sleep(0.5)

        try:
            capture_button(name)
            print(f"Successfully captured '{name}'")
        except Exception as e:
            print(f"Error capturing: {e}")

        print()

    print("Capture session ended.")


if __name__ == "__main__":
    interactive_capture()
