"""Standalone SharePoint integration package.

Public API:
    - SharePointService      (read/discovery from a source site)
    - SharePointFile         (metadata dataclass for discovered files)
    - SharePointUploadService (write/upload to an output site)
    - SharePointDeltaSync    (Graph Delta API change detection)
    - DeltaSyncStore         (durable delta bookmark + known item IDs)
    - config                 (environment-driven configuration)
"""

from __future__ import annotations

# Import the services first: service.py and upload.py load the
# .env file (load_dotenv) as a side effect of their import, so
# config.py snapshots the environment *after* .env values are
# available regardless of which submodule is imported first.
from sharepoint_integration.service import SharePointFile, SharePointService
from sharepoint_integration.upload import SharePointUploadService
from sharepoint_integration import config
from sharepoint_integration.delta_sync import (
    DeltaSyncStore,
    DeltaSyncResult,
    SharePointDeltaSync,
    get_delta_sync,
)

__all__ = [
    "config",
    "SharePointFile",
    "SharePointService",
    "SharePointUploadService",
    "SharePointDeltaSync",
    "DeltaSyncStore",
    "DeltaSyncResult",
    "get_delta_sync",
]