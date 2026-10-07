#!/usr/bin/env python
"""Standalone smoke test for the SharePoint integration.

Verifies, in order:
    1. Configuration loads from environment variables
    2. Authentication / token acquisition works
    3. Source SharePoint site can be resolved
    4. Source document library / drive can be resolved
    5. Basic file / folder access works

Run from the repository root:

    uv run python verify_connection.py

Exit code is 0 on success, 1 on failure.
"""

from __future__ import annotations

import asyncio
import logging
import sys
from typing import Optional

from sharepoint_integration.config import (
    SHAREPOINT_BASE_FOLDER,
    SHAREPOINT_CLIENT_ID,
    SHAREPOINT_CLIENT_SECRET,
    SHAREPOINT_DOMAIN,
    SHAREPOINT_LIBRARY_NAME,
    SHAREPOINT_SITE_PATH,
    SHAREPOINT_TENANT_ID,
    has_credentials,
    require_credentials,
)
from sharepoint_integration.service import SharePointService

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s | %(message)s",
)
logger = logging.getLogger("verify_connection")

PASS = "PASS"
FAIL = "FAIL"
_results: list[tuple[str, str, str]] = []


def _record(step: str, status: str, detail: str = "") -> None:
    _results.append((step, status, detail))
    suffix = f" - {detail}" if detail else ""
    print(f"[{status}] {step}{suffix}")


def _redact(value: Optional[str]) -> str:
    if not value:
        return "(not set)"
    if len(value) <= 8:
        return "*" * len(value)
    return value[:4] + "*" * (len(value) - 4)


def step_1_configuration() -> bool:
    print("\n" + "=" * 60)
    print("STEP 1: Configuration loads")
    print("=" * 60)
    ok = True
    checks = [
        ("SHAREPOINT_TENANT_ID", SHAREPOINT_TENANT_ID),
        ("SHAREPOINT_CLIENT_ID", SHAREPOINT_CLIENT_ID),
        ("SHAREPOINT_CLIENT_SECRET", SHAREPOINT_CLIENT_SECRET),
        ("SHAREPOINT_DOMAIN", SHAREPOINT_DOMAIN),
        ("SHAREPOINT_SITE_PATH", SHAREPOINT_SITE_PATH),
        ("SHAREPOINT_LIBRARY_NAME", SHAREPOINT_LIBRARY_NAME),
        ("SHAREPOINT_BASE_FOLDER", SHAREPOINT_BASE_FOLDER or "(library root)"),
    ]
    for name, value in checks:
        print(f"  {name}: {_redact(value)}")
        if name in (
            "SHAREPOINT_TENANT_ID",
            "SHAREPOINT_CLIENT_ID",
            "SHAREPOINT_CLIENT_SECRET",
            "SHAREPOINT_DOMAIN",
            "SHAREPOINT_SITE_PATH",
        ) and not value:
            print(f"    ERROR: {name} is required")
            ok = False
    _record("Configuration loads", PASS if ok else FAIL)
    return ok


def step_2_authentication(service: SharePointService) -> bool:
    print("\n" + "=" * 60)
    print("STEP 2: Authentication / token acquisition")
    print("=" * 60)
    try:
        token = asyncio.run(service._get_access_token())
        if not token:
            _record("Authentication", FAIL, "no access_token returned")
            return False
        print(f"  Token obtained: {_redact(token)}")
        _record("Authentication", PASS)
        return True
    except Exception as e:
        logger.error("Authentication failed: %s", e)
        _record("Authentication", FAIL, str(e))
        return False


def step_3_resolve_site(service: SharePointService) -> bool:
    print("\n" + "=" * 60)
    print("STEP 3: Resolve source SharePoint site")
    print("=" * 60)
    try:
        site_id, drive_id = asyncio.run(service._get_site_and_drive_ids())
        if not site_id:
            _record("Resolve site", FAIL, "no site_id returned")
            return False
        print(f"  Site ID:  {site_id}")
        print(f"  Drive ID: {drive_id}")
        _record("Resolve site", PASS)
        return True
    except Exception as e:
        logger.error("Site resolution failed: %s", e)
        _record("Resolve site", FAIL, str(e))
        return False


def step_4_resolve_drive(service: SharePointService) -> bool:
    print("\n" + "=" * 60)
    print("STEP 4: Resolve source document library / drive")
    print("=" * 60)
    if service.drive_id:
        print(f"  Drive ID cached: {service.drive_id}")
        _record("Resolve drive", PASS)
        return True
    _record("Resolve drive", FAIL, "drive_id not set after site resolution")
    return False


def step_5_basic_access(service: SharePointService) -> bool:
    print("\n" + "=" * 60)
    print("STEP 5: Basic file / folder access")
    print("=" * 60)
    try:
        tag_counts = asyncio.run(service.get_file_count())
        total = tag_counts.get("TOTAL", 0)
        print(f"  File count by tag: {tag_counts}")
        print(f"  Total files: {total}")
        _record("Basic access", PASS)
        return True
    except Exception as e:
        logger.error("Basic access failed: %s", e)
        _record("Basic access", FAIL, str(e))
        return False


def main() -> int:
    print("SharePoint Integration - Connection Verification")
    print("=" * 60)

    if not step_1_configuration():
        print("\nConfiguration is incomplete. Fix the missing values and retry.")
        return 1

    try:
        require_credentials()
    except RuntimeError as e:
        print(f"\nERROR: {e}")
        return 1

    try:
        service = SharePointService()
    except Exception as e:
        logger.error("SharePointService initialization failed: %s", e)
        _record("Initialize service", FAIL, str(e))
        return 1
    _record("Initialize service", PASS)

    if not step_2_authentication(service):
        return 1
    if not step_3_resolve_site(service):
        return 1
    if not step_4_resolve_drive(service):
        return 1
    if not step_5_basic_access(service):
        return 1

    print("\n" + "=" * 60)
    print("RESULT: ALL CHECKS PASSED")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())