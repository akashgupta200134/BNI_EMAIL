"""
scripts/test_selectors.py — Interactive selector tester for Outlook Web.

Run this BEFORE the main campaign to verify all selectors work on your account.
Usage: python scripts/test_selectors.py

What it does:
  1. Opens Outlook Web and logs in automatically
  2. Opens a New Mail compose window
  3. Tries each selector one by one and reports Pass/Fail
  4. Dumps ALL aria-labels found on page so you can find correct ones
  5. Fills a test compose (To/Subject/Body) so you can visually confirm
  6. Does NOT send anything
"""

import os
import sys
import time
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
from config import Config


# ── Selectors to test ─────────────────────────────────────────────────────────
# Add / modify these based on what you find
SELECTORS_TO_TEST = {
    "New Mail Button":  Config.SEL_NEW_MAIL,
    "To Field":         Config.SEL_TO_FIELD,
    "Subject Field":    Config.SEL_SUBJECT,
    "Body Area":        Config.SEL_BODY,
    "Send Button":      Config.SEL_SEND_BTN,
}

# ── Colors for terminal output ────────────────────────────────────────────────
GREEN  = "\033[92m"
RED    = "\033[91m"
YELLOW = "\033[93m"
CYAN   = "\033[96m"
RESET  = "\033[0m"
BOLD   = "\033[1m"


def print_header(text):
    print(f"\n{BOLD}{CYAN}{'─' * 55}{RESET}")
    print(f"{BOLD}{CYAN}  {text}{RESET}")
    print(f"{BOLD}{CYAN}{'─' * 55}{RESET}")


def print_pass(label, selector):
    print(f"  {GREEN}✅ PASS{RESET}  {label}")
    print(f"         Selector: {CYAN}{selector}{RESET}")


def print_fail(label, selector, error=""):
    print(f"  {RED}❌ FAIL{RESET}  {label}")
    print(f"         Selector: {CYAN}{selector}{RESET}")
    if error:
        print(f"         Error:    {RED}{error[:80]}{RESET}")


def do_auto_login(page):
    """Auto login using credentials from .env"""
    print(f"\n🔐 Logging in as: {Config.OUTLOOK_EMAIL}")
    page.goto(Config.OUTLOOK_URL)

    # Already logged in?
    try:
        page.wait_for_selector(Config.SEL_NEW_MAIL, timeout=5000)
        print(f"  {GREEN}✅ Already logged in{RESET}")
        return
    except PWTimeout:
        pass

    # Fill email
    page.wait_for_selector(Config.SEL_LOGIN_EMAIL, timeout=15000)
    page.fill(Config.SEL_LOGIN_EMAIL, Config.OUTLOOK_EMAIL)
    page.click(Config.SEL_LOGIN_NEXT)

    # Fill password
    page.wait_for_selector(Config.SEL_LOGIN_PASS, timeout=15000)
    page.fill(Config.SEL_LOGIN_PASS, Config.OUTLOOK_PASSWORD)
    page.click(Config.SEL_LOGIN_SUBMIT)

    # "Stay signed in?" → No
    try:
        page.wait_for_selector(Config.SEL_STAY_SIGNED, timeout=6000)
        page.click(Config.SEL_STAY_SIGNED)
    except PWTimeout:
        pass

    page.wait_for_selector(Config.SEL_NEW_MAIL, timeout=30000)
    print(f"  {GREEN}✅ Login successful{RESET}")


def dump_aria_labels(page):
    """Print all aria-label values found on page — helps find correct selectors."""
    print_header("ALL aria-label VALUES ON PAGE")
    labels = page.evaluate("""
        () => {
            const els = document.querySelectorAll('[aria-label]');
            return Array.from(els).map(el => ({
                tag:   el.tagName,
                label: el.getAttribute('aria-label'),
                id:    el.id || '',
                name:  el.getAttribute('name') || '',
                type:  el.getAttribute('type') || '',
                cls:   el.className ? el.className.toString().substring(0, 50) : ''
            }));
        }
    """)
    for item in labels:
        print(f"  {YELLOW}{item['tag']:<12}{RESET} aria-label={CYAN}\"{item['label']}\"{RESET}"
              + (f"  id={item['id']}" if item['id'] else "")
              + (f"  name={item['name']}" if item['name'] else ""))


def dump_inputs(page):
    """Print all input/textarea elements — useful for finding To/Subject/Body."""
    print_header("ALL INPUT & TEXTAREA ELEMENTS ON PAGE")
    inputs = page.evaluate("""
        () => {
            const els = document.querySelectorAll('input, textarea, [contenteditable="true"]');
            return Array.from(els).map(el => ({
                tag:         el.tagName,
                type:        el.getAttribute('type') || '',
                aria:        el.getAttribute('aria-label') || '',
                placeholder: el.getAttribute('placeholder') || '',
                id:          el.id || '',
                name:        el.getAttribute('name') || '',
                role:        el.getAttribute('role') || '',
            }));
        }
    """)
    for item in inputs:
        print(f"  {YELLOW}{item['tag']:<12}{RESET}"
              + (f" type={item['type']}" if item['type'] else "")
              + (f" | aria-label={CYAN}\"{item['aria']}\"{RESET}" if item['aria'] else "")
              + (f" | placeholder=\"{item['placeholder']}\"" if item['placeholder'] else "")
              + (f" | id={item['id']}" if item['id'] else "")
              + (f" | role={item['role']}" if item['role'] else ""))


def test_selector(page, label, selector, timeout=8000):
    """Try to find an element. Returns True/False."""
    try:
        page.wait_for_selector(selector, timeout=timeout)
        print_pass(label, selector)
        return True
    except PWTimeout:
        print_fail(label, selector, "Element not found within timeout")
        return False
    except Exception as e:
        print_fail(label, selector, str(e))
        return False


def test_fill(page, selector, value, label):
    """Try to fill a field."""
    try:
        el = page.locator(selector)
        el.click()
        time.sleep(0.3)
        page.keyboard.type(value)
        time.sleep(0.3)
        print(f"  {GREEN}✅ Filled{RESET}  {label} → \"{value}\"")
        return True
    except Exception as e:
        print(f"  {RED}❌ Fill failed{RESET}  {label}: {e}")
        return False


def main():
    print_header("OUTLOOK WEB SELECTOR TESTER")
    print(f"  Account : {Config.OUTLOOK_EMAIL}")
    print(f"  URL     : {Config.OUTLOOK_URL}")
    print(f"\n  {YELLOW}This script does NOT send any emails.{RESET}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, slow_mo=200)
        page    = browser.new_page()
        page.set_default_timeout(10000)

        # ── Login ─────────────────────────────────────────────────────────
        try:
            do_auto_login(page)
        except Exception as e:
            print(f"\n{RED}❌ Login failed: {e}{RESET}")
            browser.close()
            return

        # ── Test inbox selectors ──────────────────────────────────────────
        print_header("STEP 1 — INBOX SELECTORS")
        inbox_ok = test_selector(page, "New Mail Button", Config.SEL_NEW_MAIL)

        if not inbox_ok:
            print(f"\n  {YELLOW}⚠️  Dumping aria-labels to help you find the correct selector...{RESET}")
            dump_aria_labels(page)
            input(f"\n  {YELLOW}>>> Fix SEL_NEW_MAIL in config.py, then press ENTER to retry...{RESET} ")
            test_selector(page, "New Mail Button (retry)", Config.SEL_NEW_MAIL)

        # ── Open compose window ───────────────────────────────────────────
        print_header("STEP 2 — OPENING COMPOSE WINDOW")
        try:
            page.click(Config.SEL_NEW_MAIL)
            page.wait_for_timeout(2000)
            print(f"  {GREEN}✅ Compose window opened{RESET}")
        except Exception as e:
            print(f"  {RED}❌ Could not open compose: {e}{RESET}")
            browser.close()
            return

        # ── Dump all inputs in compose ────────────────────────────────────
        print_header("STEP 3 — INPUTS FOUND IN COMPOSE WINDOW")
        dump_inputs(page)
        dump_aria_labels(page)

        # ── Test compose selectors ────────────────────────────────────────
        print_header("STEP 4 — COMPOSE FIELD SELECTORS")
        results = {}
        for label, selector in SELECTORS_TO_TEST.items():
            if label == "New Mail Button":
                continue
            results[label] = test_selector(page, label, selector)
            page.wait_for_timeout(400)

        # ── Try filling fields that passed ────────────────────────────────
        print_header("STEP 5 — FILLING FIELDS (visual test)")

        if results.get("To Field"):
            to_field = page.locator(Config.SEL_TO_FIELD).first
            to_field.click()
            page.wait_for_timeout(500)
            page.keyboard.type("test@example.com", delay=60)
            page.wait_for_timeout(1200)
            page.keyboard.press("Enter")   # confirm address token, dismiss dropdown
            page.wait_for_timeout(800)
            print(f"  {GREEN}✅ To field filled (address confirmed with Enter){RESET}")

        if results.get("Subject Field"):
            subject_locator = page.locator(Config.SEL_SUBJECT).first
            subject_locator.click()
            page.wait_for_timeout(400)
            subject_locator.fill("TEST — Selector Check (do not send)")
            page.wait_for_timeout(400)
            print(f"  {GREEN}✅ Subject field filled{RESET}")

        if results.get("Body Area"):
            body_locator = page.locator(Config.SEL_BODY).first
            body_locator.click()          # always explicitly click body
            page.wait_for_timeout(600)
            page.keyboard.type("This is a test message. No email will be sent.", delay=10)
            page.wait_for_timeout(400)
            print(f"  {GREEN}✅ Body field filled{RESET}")

        # ── Summary ───────────────────────────────────────────────────────
        print_header("SUMMARY")
        passed = sum(1 for v in results.values() if v)
        total  = len(results)
        color  = GREEN if passed == total else RED

        print(f"  {color}{BOLD}{passed}/{total} selectors working{RESET}")
        print()
        for label, ok in results.items():
            icon = f"{GREEN}✅" if ok else f"{RED}❌"
            print(f"  {icon}  {label}{RESET}")

        if passed < total:
            print(f"\n  {YELLOW}⚠️  Fix the failing selectors in config.py")
            print(f"     Use the aria-label dump above to find correct values.{RESET}")
        else:
            print(f"\n  {GREEN}🎉 All selectors working! You're ready to run main.py{RESET}")

        print(f"\n  {YELLOW}The compose window is still open — check the browser to visually confirm.{RESET}")
        input(f"\n  >>> Press ENTER to close browser and exit: ")
        browser.close()


if __name__ == "__main__":
    main()
