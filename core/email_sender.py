
"""
core/email_sender.py — Orchestrates the full send campaign using Playwright.
Selectors verified via Playwright codegen on this exact Outlook account.
"""

import time
import os
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout, Page
from config import Config
from utils.validator import is_valid_email
from utils.excel_handler import ExcelHandler
from utils.state_manager import StateManager


class EmailSender:

    def __init__(self, excel: ExcelHandler, state_manager: StateManager, logger, dry_run: bool = False):
        self.excel         = excel
        self.state_manager = state_manager
        self.logger        = logger
        self.dry_run       = dry_run

    # ─────────────────────────────────────────────────────────────────────────
    # Public entry point
    # ─────────────────────────────────────────────────────────────────────────

    def run(self, pending_rows: list, subject: str, template: str):
        state       = self.state_manager.load()
        sent_today  = state["count"]
        batch_count = 0

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=Config.HEADLESS)
            context = browser.new_context()
            page    = context.new_page()
            page.set_default_timeout(Config.PAGE_TIMEOUT)

            self._do_login(page)

            print(f"\n🚀 Starting campaign. {len(pending_rows)} emails to process.\n")

            for row in pending_rows:
                # Daily limit guard
                if sent_today >= Config.DAILY_LIMIT:
                    print(f"\n🚫 Daily limit of {Config.DAILY_LIMIT} reached. Stopping.")
                    self.logger.warning("Daily limit reached mid-campaign.")
                    break

                name_cell   = row[Config.COL_NAME]
                email_cell  = row[Config.COL_EMAIL]
                status_cell = row[Config.COL_STATUS]
                remark_cell = row[Config.COL_REMARK]

                name  = str(name_cell.value or "").strip()
                email = str(email_cell.value or "").strip()

                print(f"📧 [{sent_today + 1}/{Config.DAILY_LIMIT}] Processing: {name} <{email}>")

                # Pre-validate email format
                if not is_valid_email(email):
                    status_cell.value = Config.STATUS_INVALID
                    self.excel.save()
                    self.logger.warning(f"INVALID_EMAIL | {name} | {email}")
                    print(f"   ⚠️  Invalid email format — skipped, marked 'Invalid Email'")
                    continue

                # Personalise body
                body = template.replace("{{Name}}", name).replace("{{Email}}", email)

                success = self._send_with_retry(page, email, subject, body, name)

                if success:
                    remark_cell.value = Config.REMARK_DONE
                    sent_today += 1
                    batch_count += 1
                    state["count"] = sent_today
                    self.state_manager.save(state)
                    self.excel.save()
                    self.logger.info(f"SENT | {name} | {email}")
                    print(f"   ✅ Sent successfully ({sent_today} today)")
                else:
                    status_cell.value = Config.STATUS_NOT_FOUND
                    self.excel.save()
                    self.logger.error(f"FAILED | {name} | {email}")
                    print(f"   ❌ Failed — marked 'Not Found'")

                # Batch pause
                if batch_count > 0 and batch_count % Config.BATCH_SIZE == 0:
                    print(f"\n⏸️  Batch of {Config.BATCH_SIZE} complete. Pausing {Config.BATCH_PAUSE}s...\n")
                    self.logger.info(f"Batch pause after {batch_count} emails")
                    time.sleep(Config.BATCH_PAUSE)

                # Per-email delay
                elif sent_today < Config.DAILY_LIMIT:
                    remaining = len(pending_rows) - (pending_rows.index(row) + 1)
                    if remaining > 0:
                        print(f"   ⏳ Waiting {Config.DELAY_BETWEEN_EMAILS}s before next email...")
                        time.sleep(Config.DELAY_BETWEEN_EMAILS)

            browser.close()

    # ─────────────────────────────────────────────────────────────────────────
    # Login — exact flow from codegen recording
    # ─────────────────────────────────────────────────────────────────────────

    def _do_login(self, page: Page):
        print("\n🌐 Opening Outlook Web...")
        page.goto(Config.OUTLOOK_URL)

        # Already logged in?
        try:
            page.wait_for_selector('[aria-label="New mail"], [name="New mail"]', timeout=5000)
            print("✅ Already logged in.\n")
            return
        except PWTimeout:
            pass

        if not Config.OUTLOOK_EMAIL or not Config.OUTLOOK_PASSWORD:
            raise RuntimeError(
                "Outlook credentials not found. "
                "Make sure OUTLOOK_EMAIL and OUTLOOK_PASSWORD are set in your .env file."
            )

        print(f"🔐 Auto-logging in as: {Config.OUTLOOK_EMAIL}")

        try:
            # Step 1: Enter email
            # get_by_role("textbox", name="Enter your email, phone, or")
            page.get_by_role("textbox", name="Enter your email, phone, or").click()
            page.get_by_role("textbox", name="Enter your email, phone, or").fill(Config.OUTLOOK_EMAIL)
            page.get_by_role("button", name="Next").click()

            # Step 2: Enter password
            # get_by_role("textbox", name="Enter the password for")
            page.get_by_role("textbox", name="Enter the password for").wait_for(timeout=15000)
            page.get_by_role("textbox", name="Enter the password for").fill(Config.OUTLOOK_PASSWORD)
            page.get_by_role("textbox", name="Enter the password for").press("Enter")

            # Step 3: "Stay signed in?" — click No
            # get_by_role("checkbox", name="Don't show this again") + get_by_role("button", name="No")
            try:
                page.get_by_role("button", name="No").wait_for(timeout=8000)
                page.get_by_role("button", name="No").click()
                self.logger.info("Dismissed 'Stay signed in?' prompt")
            except PWTimeout:
                pass  # Prompt did not appear — fine

            # Step 4: Wait for inbox / New mail button
            page.get_by_role("button", name="New mail").wait_for(timeout=30000)
            print("✅ Login successful. Starting campaign...\n")
            self.logger.info(f"Logged in as {Config.OUTLOOK_EMAIL}")

        except PWTimeout:
            self._take_screenshot(page, "login_failed")
            raise RuntimeError(
                "Auto-login timed out.\n"
                "  • Check OUTLOOK_EMAIL and OUTLOOK_PASSWORD in .env\n"
                "  • MFA/2FA may be blocking the login\n"
                "  Screenshot saved to screenshots/login_failed.png"
            )
        except RuntimeError:
            raise
        except Exception as e:
            self._take_screenshot(page, "login_error")
            raise RuntimeError(f"Unexpected login error: {e}")

    # ─────────────────────────────────────────────────────────────────────────
    # Send with retry
    # ─────────────────────────────────────────────────────────────────────────

    def _send_with_retry(self, page: Page, to_email: str, subject: str, body: str, name: str) -> bool:
        for attempt in range(1, Config.COMPOSE_RETRIES + 2):
            try:
                self._compose_and_send(page, to_email, subject, body)
                return True
            except PWTimeout as e:
                self.logger.warning(f"TIMEOUT attempt {attempt} | {name} | {to_email} | {e}")
                print(f"   ⚠️  Timeout on attempt {attempt}/{Config.COMPOSE_RETRIES + 1}")
                if attempt <= Config.COMPOSE_RETRIES:
                    self._take_screenshot(page, f"timeout_{to_email}_{attempt}")
                    time.sleep(5)
                    self._recover_to_inbox(page)
            except Exception as e:
                self.logger.error(f"ERROR attempt {attempt} | {name} | {to_email} | {e}")
                print(f"   ⚠️  Error on attempt {attempt}: {e}")
                if attempt <= Config.COMPOSE_RETRIES:
                    self._take_screenshot(page, f"error_{to_email}_{attempt}")
                    time.sleep(5)
                    self._recover_to_inbox(page)

        return False

    # ─────────────────────────────────────────────────────────────────────────
    # Compose and send — exact selector flow from codegen recording
    # ─────────────────────────────────────────────────────────────────────────

    def _compose_and_send(self, page: Page, to_email: str, subject: str, body: str):
        self._recover_to_inbox(page)

        # ── Open compose ──────────────────────────────────────────────────
        # get_by_role("button", name="New mail")
        page.get_by_role("button", name="New mail").click()
        page.wait_for_timeout(1500)  # wait for compose panel to fully render

        # ── To field ──────────────────────────────────────────────────────
        # get_by_label("To", exact=True)
        # Use exact=True so it doesn't match "To" inside other labels
        page.get_by_label("To", exact=True).click()
        page.wait_for_timeout(400)
        page.get_by_label("To", exact=True).fill(to_email)
        page.wait_for_timeout(1000)  # let autocomplete appear

        # Press Enter to confirm the address as a recipient token
        # and close the autocomplete dropdown in one keystroke
        page.keyboard.press("Enter")
        page.wait_for_timeout(600)

        # ── Subject field ─────────────────────────────────────────────────
        # get_by_placeholder("Add a subject")
        # Always click directly — never rely on Tab to land here
        page.get_by_placeholder("Add a subject").click()
        page.wait_for_timeout(300)
        page.get_by_placeholder("Add a subject").fill(subject)
        page.wait_for_timeout(400)

        # ── Body field ────────────────────────────────────────────────────
        # get_by_label("Message body")
        # Always click directly to ensure focus lands in the body
        page.get_by_label("Message body").click()
        page.wait_for_timeout(500)

        # Type line by line so newlines render as paragraph breaks
        for line in body.split("\n"):
            page.keyboard.type(line, delay=10)
            page.keyboard.press("Enter")

        page.wait_for_timeout(600)

        # ── Send or discard (dry run) ─────────────────────────────────────
        if self.dry_run:
            print("   🧪 DRY RUN — compose filled, not sending")
            try:
                page.get_by_role("button", name="Discard").click()
                page.wait_for_timeout(800)
            except Exception:
                page.keyboard.press("Escape")
                page.wait_for_timeout(800)
        else:
            # get_by_label("Send", exact=True)
            page.get_by_label("Send", exact=True).click()
            page.wait_for_timeout(Config.SEND_WAIT)

    # ─────────────────────────────────────────────────────────────────────────
    # Helpers
    # ─────────────────────────────────────────────────────────────────────────

    def _recover_to_inbox(self, page: Page):
        """Navigate back to inbox if something went wrong."""
        try:
            page.get_by_role("button", name="New mail").wait_for(timeout=3000)
        except PWTimeout:
            self.logger.info("Recovering — navigating back to inbox")
            page.goto(f"{Config.OUTLOOK_URL}/mail/0/")
            page.get_by_role("button", name="New mail").wait_for(timeout=Config.PAGE_TIMEOUT)

    def _take_screenshot(self, page: Page, label: str):
        """Save a screenshot for debugging failures."""
        try:
            safe_label = label.replace("@", "_at_").replace(".", "_")
            path = os.path.join(Config.SCREENSHOT_DIR, f"{safe_label}.png")
            page.screenshot(path=path)
            self.logger.info(f"Screenshot saved: {path}")
        except Exception as e:
            self.logger.warning(f"Could not save screenshot: {e}")
