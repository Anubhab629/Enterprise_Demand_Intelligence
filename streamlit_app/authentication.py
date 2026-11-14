
"""Streamlit-compatible authentication using EDI's existing credentials."""

from __future__ import annotations

import hmac
import os

import env


def verify_credentials(username: str, password: str) -> bool:
    """Validate dashboard credentials using the existing EDI credential rules."""
    expected_username = os.getenv("EDI_ADMIN_USERNAME", "").strip()
    expected_hash = os.getenv("EDI_ADMIN_PASSWORD_HASH", "").strip()

    if not expected_username or not expected_hash:
        return False

    supplied_hash = env.hash_password(password)

    return (
        hmac.compare_digest(username.strip(), expected_username)
        and hmac.compare_digest(supplied_hash, expected_hash)
    )
