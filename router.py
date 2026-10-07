"""FastAPI router exposing the standalone SharePoint integration.

This adapter only wraps the SharePoint read/upload services. It does not
implement any host application business logic (no RFP pipeline, no database,
no embeddings, no blob storage).

Mount or include this router in a FastAPI app when you want HTTP access:

    from fastapi import FastAPI
    from router import router

    app = FastAPI()
    app.include_router(router)

The core ``sharepoint_integration`` package does not depend on FastAPI.
"""

from __future__ import annotations

import logging
from typing import Optional

from fastapi import APIRouter, HTTPException, Query

from sharepoint_integration.service import SharePointFile, SharePointService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/sharepoint", tags=["sharepoint"])

# A single lazily-initialized service instance. Importing this module does
# not require credentials; the service is only built on first request.
_service: Optional[SharePointService] = None


def _get_service() -> SharePointService:
    global _service
    if _service is None:
        logger.info("Initializing SharePointService for first request")
        _service = SharePointService()
    return _service


@router.get("/health")
async def health() -> dict:
    """Basic liveness check - does not touch SharePoint."""
    logger.info("health check")
    return {"status": "ok"}


@router.get("/files", response_model=list[SharePointFile])
async def list_files(
    base_folder: Optional[str] = Query(default=None, description="Override the configured base folder"),
) -> list[SharePointFile]:
    """List all files under the configured base folder, grouped by tag."""
    try:
        service = _get_service()
        files = await service.discover_tagged_files(base_folder)
        logger.info("list_files succeeded: returned %d files", len(files))
        return files
    except Exception as e:
        logger.error("list_files failed: %s", e)
        raise HTTPException(status_code=500, detail=str(e)) from e


@router.get("/files/count")
async def file_count(
    base_folder: Optional[str] = Query(default=None, description="Override the configured base folder"),
) -> dict:
    """Return file counts grouped by tag (no file downloads)."""
    try:
        service = _get_service()
        counts = await service.get_file_count(base_folder)
        logger.info("file_count succeeded: %s", counts)
        return counts
    except Exception as e:
        logger.error("file_count failed: %s", e)
        raise HTTPException(status_code=500, detail=str(e)) from e


@router.get("/files/{relative_path:path}", response_model=SharePointFile)
async def get_file(relative_path: str) -> SharePointFile:
    """Fetch a file by its path relative to the base folder, or by a sharing URL."""
    try:
        service = _get_service()
        sp_file = await service.get_file_by_relative_path(relative_path)
        if sp_file is None:
            logger.info("get_file: not found: %s", relative_path)
            raise HTTPException(status_code=404, detail="File not found")
        logger.info("get_file succeeded: %s (%d bytes)", sp_file.name, sp_file.size_bytes)
        return sp_file
    except HTTPException:
        raise
    except Exception as e:
        logger.error("get_file failed: %s", e)
        raise HTTPException(status_code=500, detail=str(e)) from e