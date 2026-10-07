"""Standalone SharePoint integration package.

Public API:
    - SharePointService      (read/discovery from a source site)
    - SharePointFile         (metadata dataclass for discovered files)
    - SharePointUploadService (write/upload to an output site)
    - config                 (environment-driven configuration)
"""

from __future__ import annotations

from sharepoint_integration import config
from sharepoint_integration.service import SharePointFile, SharePointService
from sharepoint_integration.upload import SharePointUploadService

__all__ = [
    "config",
    "SharePointFile",
    "SharePointService",
    "SharePointUploadService",
]