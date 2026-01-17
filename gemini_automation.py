"""
Gemini Automation Script
Automates document processing through Gemini Web interface
and aggregates results to a master spreadsheet.
"""

import os
import sys
import time
import glob
import pyautogui

from config import (
    GOOGLE_DRIVE_BOOK_KDL_PATH,
    MASTER_SPREADSHEET_NAME,
    WAIT_MEDIUM,
    WAIT_SPREADSHEET,
)
from gemini_ops import GeminiAutomation, SpreadsheetAutomation, log

# Fail-safe: move mouse to corner to abort
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.5


def get_docx_files() -> list:
    """Get list of .docx files from the Google Drive folder."""
    pattern = os.path.join(GOOGLE_DRIVE_BOOK_KDL_PATH, "*.docx")
    files = glob.glob(pattern)
    log(f"Found {len(files)} .docx files")
    for f in files:
        log(f"  - {os.path.basename(f)}")
    return files


def process_single_file(
    file_path: str,
    gemini: GeminiAutomation,
    spreadsheet: SpreadsheetAutomation
) -> bool:
    """Process a single DOCX file through Gemini."""
    log(f"\n{'='*50}")
    log(f"Processing: {os.path.basename(file_path)}")
    log(f"{'='*50}")

    try:
        # Step 1: Ensure Gemini is active
        if not gemini.activate_gemini():
            return False

        # Step 2: Select Gem
        gemini.select_gem()

        # Step 3: Input prompt
        gemini.input_prompt()

        # Step 4: Upload file
        gemini.click_upload()
        time.sleep(WAIT_MEDIUM)
        gemini.upload_file_dialog(file_path)

        # Step 5: Execute
        gemini.execute_prompt()

        # Step 6: Wait for response
        gemini.wait_for_completion(timeout=90.0)

        # Step 7: Open spreadsheet
        gemini.open_spreadsheet()
        gemini.confirm_spreadsheet_dialog()

        # Step 8: Copy data from generated spreadsheet
        time.sleep(WAIT_SPREADSHEET)
        spreadsheet.copy_data()

        # Step 9: Paste to master sheet
        if not spreadsheet.activate_master_sheet(MASTER_SPREADSHEET_NAME):
            log("WARNING: Could not activate master spreadsheet")
            return False

        spreadsheet.paste_data()

        # Step 10: Close temp spreadsheet and return to Gemini
        spreadsheet.close_current()

        log(f"Successfully processed: {os.path.basename(file_path)}")
        return True

    except Exception as e:
        log(f"ERROR processing file: {e}")
        return False


def main():
    """Main entry point for the automation."""
    log("=" * 60)
    log("GEMINI AUTOMATION - Document Processing")
    log("=" * 60)
    log(f"Source folder: {GOOGLE_DRIVE_BOOK_KDL_PATH}")
    log(f"Master sheet: {MASTER_SPREADSHEET_NAME}")

    # Verify source path exists
    if not os.path.exists(GOOGLE_DRIVE_BOOK_KDL_PATH):
        log(f"ERROR: Source path does not exist: {GOOGLE_DRIVE_BOOK_KDL_PATH}")
        log("Please update GOOGLE_DRIVE_BOOK_KDL_PATH in config.py")
        sys.exit(1)

    # Get list of files to process
    files = get_docx_files()
    if not files:
        log("No .docx files found to process")
        sys.exit(0)

    # Initialize automation handlers
    gemini = GeminiAutomation()
    spreadsheet = SpreadsheetAutomation()

    # Process each file
    successful = 0
    failed = 0
    failed_files = []

    for i, file_path in enumerate(files):
        log(f"\nProgress: {i+1}/{len(files)}")

        if process_single_file(file_path, gemini, spreadsheet):
            successful += 1
        else:
            failed += 1
            failed_files.append(os.path.basename(file_path))

        # Small delay between files
        time.sleep(WAIT_MEDIUM)

    # Summary
    log("\n" + "=" * 60)
    log("AUTOMATION COMPLETE")
    log("=" * 60)
    log(f"Successful: {successful}")
    log(f"Failed: {failed}")
    log(f"Total: {len(files)}")

    if failed_files:
        log("\nFailed files:")
        for f in failed_files:
            log(f"  - {f}")

    log("=" * 60)


if __name__ == "__main__":
    main()
