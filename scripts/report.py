"""
scripts/report.py — Print a summary of campaign progress from the Excel file.

Usage: python scripts/report.py
"""

import os
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.excel_handler import ExcelHandler
from utils.state_manager import StateManager
from config import Config


def main():
    excel = ExcelHandler(Config.EXCEL_FILE)
    try:
        excel.load()
    except FileNotFoundError:
        print(f"❌ Excel file not found: {Config.EXCEL_FILE}")
        sys.exit(1)

    summary = excel.get_summary()
    state   = StateManager(Config.STATE_FILE).load()

    print("\n" + "=" * 45)
    print("     📊  EMAIL CAMPAIGN REPORT")
    print("=" * 45)
    print(f"  Total contacts   : {summary['total']}")
    print(f"  ✅ Done          : {summary['done']}")
    print(f"  📋 Pending       : {summary['pending']}")
    print(f"  ⚠️  Invalid email : {summary['invalid']}")
    print(f"  ❌ Not Found     : {summary['not_found']}")
    print("-" * 45)
    print(f"  📆 Sent today    : {state['count']} / {Config.DAILY_LIMIT}")
    print(f"  📅 Date          : {state['date']}")
    print("=" * 45 + "\n")


if __name__ == "__main__":
    main()
