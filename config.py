# Gemini Automation Configuration
# Update these paths according to your environment

import os

# Google Drive book_kdl folder path (where .docx files are located)
# Examples:
#   Windows: "C:/Users/YourName/Google Drive/My Drive/book_kdl"
#   Or if using Google Drive for Desktop: "G:/My Drive/book_kdl"
GOOGLE_DRIVE_BOOK_KDL_PATH = "G:/マイドライブ/book_kdl"

# Master spreadsheet name for aggregating results
MASTER_SPREADSHEET_NAME = "KDL_20260115"

# Timing configurations (in seconds)
WAIT_SHORT = 1.0          # Short wait for UI transitions
WAIT_MEDIUM = 3.0         # Medium wait for file uploads
WAIT_LONG = 10.0          # Long wait for Gemini processing
WAIT_SPREADSHEET = 5.0    # Wait for spreadsheet to open

# Retry settings
MAX_RETRIES = 3
RETRY_DELAY = 2.0

# Gemini window title pattern
GEMINI_WINDOW_TITLE = "Gemini"

# Gem name to select
GEM_NAME = "GEM"
