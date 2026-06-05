"""
utils/validator.py — Email address validation helpers.
"""

import re


# RFC-5321 simplified pattern — covers 99% of real-world emails
_EMAIL_REGEX = re.compile(
    r"^(?!\.)"                          # no leading dot
    r"[a-zA-Z0-9._%+\-]+"             # local part
    r"@"
    r"(?!-)"                            # domain can't start with hyphen
    r"[a-zA-Z0-9\-]+"
    r"(\.[a-zA-Z0-9\-]+)*"
    r"\.[a-zA-Z]{2,}$"                 # TLD at least 2 chars
)


def is_valid_email(email: str) -> bool:
    """
    Returns True if email passes basic format validation.
    Does NOT do DNS/MX lookup — just syntax check.
    """
    if not email or not isinstance(email, str):
        return False

    email = email.strip()

    if len(email) > 254:
        return False

    # Must have exactly one @
    parts = email.split("@")
    if len(parts) != 2:
        return False

    local, domain = parts

    # Local part can't exceed 64 chars
    if len(local) > 64:
        return False

    return bool(_EMAIL_REGEX.match(email))


def sanitize_email(email: str) -> str:
    """Strip whitespace and lowercase the email."""
    return email.strip().lower() if email else ""
