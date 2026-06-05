"""
utils/logger.py — Sets up file + console logging for the campaign.
"""

import logging
import os
from datetime import date
from config import Config


def setup_logger(name: str = "email_campaign") -> logging.Logger:
    """
    Sets up a logger that writes to:
    - logs/campaign_YYYY-MM-DD.log  (file, appends)
    - console (INFO and above)
    """
    os.makedirs(Config.LOG_DIR, exist_ok=True)

    log_file = os.path.join(Config.LOG_DIR, f"campaign_{date.today()}.log")

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Avoid duplicate handlers if called multiple times
    if logger.handlers:
        return logger

    # ── File handler ──────────────────────────────────────────────────────
    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))

    # ── Console handler ───────────────────────────────────────────────────
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)   # Only warnings/errors on console
    ch.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    logger.addHandler(fh)
    logger.addHandler(ch)

    return logger
