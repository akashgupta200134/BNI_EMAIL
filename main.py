"""
Outlook Web Email Automation - Main Entry Point
Usage: python main.py [--dry-run] [--reset-state]
"""

import argparse
import sys
from core.email_sender import EmailSender
from utils.logger import setup_logger
from utils.state_manager import StateManager
from utils.excel_handler import ExcelHandler
from config import Config

def parse_args():
    parser = argparse.ArgumentParser(description="Outlook Web Email Automation")
    parser.add_argument("--dry-run", action="store_true", help="Run without sending emails")
    parser.add_argument("--reset-state", action="store_true", help="Reset daily send counter")
    parser.add_argument("--excel", type=str, default=Config.EXCEL_FILE, help="Path to Excel file")
    return parser.parse_args()


def main():
    args = parse_args()

    # Setup
    logger = setup_logger()
    logger.info("=" * 60)
    logger.info("Email Campaign Started")
    logger.info(f"Dry Run: {args.dry_run}")
    logger.info(f"Excel File: {args.excel}")
    logger.info("=" * 60)

    if args.dry_run:
        print("\n⚠️  DRY RUN MODE — Emails will NOT be sent\n")

    # Validate credentials loaded from .env
    if not Config.OUTLOOK_EMAIL or not Config.OUTLOOK_PASSWORD:
        print("\n❌ Outlook credentials missing!")
        print("   Create a .env file in the project root with:")
        print("   OUTLOOK_EMAIL=your@email.com")
        print("   OUTLOOK_PASSWORD=yourpassword")
        sys.exit(1)
    logger.info(f"Credentials loaded for: {Config.OUTLOOK_EMAIL}")

    # Reset state if requested
    state_manager = StateManager(Config.STATE_FILE)
    if args.reset_state:
        state_manager.reset()
        print("✅ State reset. Daily counter cleared.")

    # Check daily limit before starting
    state = state_manager.load()
    if state["count"] >= Config.DAILY_LIMIT:
        print(f"\n🚫 Daily limit of {Config.DAILY_LIMIT} emails already reached for today.")
        print("   Use --reset-state to override, or wait until tomorrow.")
        logger.warning("Daily limit reached. Exiting.")
        sys.exit(0)

    remaining = Config.DAILY_LIMIT - state["count"]
    print(f"\n📊 Daily limit: {Config.DAILY_LIMIT} | Sent today: {state['count']} | Remaining: {remaining}")

    # Load Excel
    excel = ExcelHandler(args.excel)
    try:
        excel.load()
    except FileNotFoundError:
        print(f"\n❌ Excel file not found: {args.excel}")
        print("   Make sure 'email_campaign.xlsx' is in the data/ folder.")
        sys.exit(1)

    subject, template = excel.get_template(template_index=0)
    if not subject or not template:
        print("\n❌ Could not read template from Sheet2. Check your Excel file.")
        sys.exit(1)

    pending = excel.get_pending_rows()
    print(f"📋 Pending emails to send: {len(pending)}")

    if not pending:
        print("\n✅ No pending emails. All rows are marked Done or have errors.")
        sys.exit(0)

    # Start sender
    sender = EmailSender(
        excel=excel,
        state_manager=state_manager,
        logger=logger,
        dry_run=args.dry_run
    )

    try:
        sender.run(pending, subject, template)
    except KeyboardInterrupt:
        print("\n\n⚠️  Interrupted by user. Progress has been saved.")
        logger.warning("Campaign interrupted by user.")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        logger.exception("Unexpected error in main")
    finally:
        final_state = state_manager.load()
        print(f"\n📊 Session complete. Total sent today: {final_state['count']}/{Config.DAILY_LIMIT}")
        logger.info(f"Session complete. Total sent today: {final_state['count']}")


if __name__ == "__main__":
    main()
