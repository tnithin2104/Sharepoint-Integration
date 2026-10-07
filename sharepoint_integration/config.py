"""Configuration for the standalone SharePoint integration.

All values are read from environment variables. No credentials or
organisation-specific site values are hardcoded.
"""

from __future__ import annotations

import os
from typing import Optional


def _get_bool(name: str, default: bool = False) -> bool:
    return os.getenv(name, str(default)).strip().lower() in ("1", "true", "yes", "on")


def _get_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


# --- Authentication ---
SHAREPOINT_TENANT_ID: Optional[str] = os.getenv("SHAREPOINT_TENANT_ID")
SHAREPOINT_CLIENT_ID: Optional[str] = os.getenv("SHAREPOINT_CLIENT_ID")
SHAREPOINT_CLIENT_SECRET: Optional[str] = os.getenv("SHAREPOINT_CLIENT_SECRET")

# --- Source / READ site ---
SHAREPOINT_DOMAIN: Optional[str] = os.getenv("SHAREPOINT_DOMAIN")
SHAREPOINT_SITE_PATH: Optional[str] = os.getenv("SHAREPOINT_SITE_PATH")
SHAREPOINT_LIBRARY_NAME: str = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documents")
SHAREPOINT_BASE_FOLDER: str = os.getenv("SHAREPOINT_BASE_FOLDER", "")

# --- Output / WRITE site ---
SHAREPOINT_OUTPUT_DOMAIN: Optional[str] = os.getenv("SHAREPOINT_OUTPUT_DOMAIN")
SHAREPOINT_OUTPUT_SITE_PATH: Optional[str] = os.getenv("SHAREPOINT_OUTPUT_SITE_PATH")
SHAREPOINT_OUTPUT_LIBRARY: Optional[str] = os.getenv("SHAREPOINT_OUTPUT_LIBRARY")

# --- Optional ---
SHAREPOINT_SSL_VERIFY: bool = _get_bool("SHAREPOINT_SSL_VERIFY", True)
SHAREPOINT_CONNECT_TIMEOUT: int = _get_int("SHAREPOINT_CONNECT_TIMEOUT", 30)
SHAREPOINT_READ_TIMEOUT: int = _get_int("SHAREPOINT_READ_TIMEOUT", 120)
SHAREPOINT_MAX_RETRIES: int = _get_int("SHAREPOINT_MAX_RETRIES", 5)
SHAREPOINT_POOL_CONNECTIONS: int = _get_int("SHAREPOINT_POOL_CONNECTIONS", 20)
SHAREPOINT_POOL_MAXSIZE: int = _get_int("SHAREPOINT_POOL_MAXSIZE", 50)


def has_credentials() -> bool:
    """True when the three authentication variables are all set."""
    return bool(SHAREPOINT_TENANT_ID and SHAREPOINT_CLIENT_ID and SHAREPOINT_CLIENT_SECRET)


def require_credentials() -> None:
    """Raise RuntimeError if authentication credentials are missing."""
    if not has_credentials():
        raise RuntimeError(
            "Missing required environment variables: SHAREPOINT_TENANT_ID, "
            "SHAREPOINT_CLIENT_ID, SHAREPOINT_CLIENT_SECRET"
        )