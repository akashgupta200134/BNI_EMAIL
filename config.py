"""
config.py — Central configuration for the email automation.
Credentials are loaded from .env — never hardcode them here.
"""

import os
from dotenv import load_dotenv

# Load .env from the project root
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))


class Config:
    # ── File Paths ────────────────────────────────────────────────────────
    BASE_DIR        = os.path.dirname(os.path.abspath(__file__))
    EXCEL_FILE      = os.path.join(BASE_DIR, "data", "email_campaign.xlsx")
    STATE_FILE      = os.path.join(BASE_DIR, "data", "state.json")
    LOG_DIR         = os.path.join(BASE_DIR, "logs")
    SCREENSHOT_DIR  = os.path.join(BASE_DIR, "screenshots")

    # ── Outlook Credentials (loaded from .env) ────────────────────────────
    OUTLOOK_EMAIL    = os.getenv("OUTLOOK_EMAIL")
    OUTLOOK_PASSWORD = os.getenv("OUTLOOK_PASSWORD")

    # ── Outlook Web ───────────────────────────────────────────────────────
    OUTLOOK_URL     = "https://outlook.office.com"

    # ── Send Settings ─────────────────────────────────────────────────────
    DELAY_BETWEEN_EMAILS = 30       # seconds between each email send
    BATCH_SIZE           = 50       # emails per batch
    BATCH_PAUSE          = 60       # seconds pause between batches
    DAILY_LIMIT          = 200      # max emails per calendar day

    # ── Playwright Settings ───────────────────────────────────────────────
    HEADLESS         = False        # False = visible browser (needed for login)
    PAGE_TIMEOUT     = 20000        # ms — max wait for page elements
    SEND_WAIT        = 4000         # ms — wait after clicking Send
    COMPOSE_RETRIES  = 2            # retry count if compose window fails

    # ── Excel Column Indices (0-based) ────────────────────────────────────
    COL_NAME         = 0   # Column A
    COL_EMAIL        = 1   # Column B
    COL_STATUS       = 2   # Column C — Email Status
    COL_REMARK       = 3   # Column D — Remark

    # ── Status Values ─────────────────────────────────────────────────────
    REMARK_DONE      = "Done"
    STATUS_INVALID   = "Invalid Email"
    STATUS_NOT_FOUND = "Not Found"
    STATUS_SENT      = "Sent"

    # ── Selectors — Microsoft Login Page ─────────────────────────────────
    SEL_LOGIN_EMAIL  = 'input[type="email"]'
    SEL_LOGIN_NEXT   = 'input[type="submit"]'
    SEL_LOGIN_PASS   = 'input[type="password"]'
    SEL_LOGIN_SUBMIT = 'input[type="submit"]'
    SEL_STAY_SIGNED  = '#idBtn_Back'          # "No" on "Stay signed in?" prompt
    SEL_LOGIN_ERROR  = '#usernameError, #passwordError, .alert-error'

    # ── Selectors — Outlook Web ───────────────────────────────────────────
    # Run: python scripts/test_selectors.py  to verify these on your account
    # If selectors fail, the test script will dump all aria-labels to help fix them

    SEL_NEW_MAIL  = (
        '[aria-label="New mail"],'
        '[aria-label="New Mail"],'
        '[aria-label="New message"],'
        'button[title="New mail"]'
    )
    SEL_TO_FIELD  = (
        '[aria-label="To"],'
        'input[aria-label="To"],'
        'div[aria-label="To"]'
    )
    SEL_SUBJECT   = (
        '[aria-label="Add a subject"],'
        'input[aria-label="Add a subject"],'
        '[placeholder="Add a subject"],'
        '[aria-label="Subject"]'
    )
    SEL_BODY      = (
        '[aria-label="Message body, body text"],'
        '[aria-label="Message body"],'
        'div[contenteditable="true"][aria-label],'
        '[role="textbox"][aria-multiline="true"]'
    )
    SEL_SEND_BTN  = (
        '[aria-label="Send"],'
        'button[aria-label="Send"],'
        '[data-icon-name="Send"],'
        'button[title="Send"]'
    )
    SEL_INBOX     = '[aria-label="Mail"]'
