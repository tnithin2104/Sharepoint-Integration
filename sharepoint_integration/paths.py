"""Consistent document path fields for documents table."""

from __future__ import annotations

import os
from typing import Any, Dict, Optional
from urllib.parse import unquote, urlparse


def blob_url_to_relative(
    blob_url: Optional[str],
    blob_base_path: Optional[str] = None,
) -> Optional[str]:
    if not blob_url:
        return None
    base = blob_base_path or os.getenv(
        "BLOB_BASE_PATH",
        "https://smartrfpstg.blob.core.windows.net/srfp-processed-documents/",
    )
    if blob_url.startswith(base):
        return unquote(blob_url[len(base) :])
    parsed = urlparse(blob_url)
    parts = parsed.path.lstrip("/").split("/", 1)
    if len(parts) == 2:
        return unquote(parts[1])
    return unquote(parsed.path.lstrip("/"))


def normalize_sharepoint_source_path(
    path: str,
    base_folder: str = "",
) -> str:
    """Return path relative to the configured SharePoint base folder (e.g. RFP/)."""
    normalized = path.strip().replace("\\", "/")
    base = (base_folder or "").strip().strip("/")
    if base and normalized.lower().startswith(base.lower() + "/"):
        return normalized[len(base) + 1 :]
    return normalized


SHAREPOINT_BLOB_ROOT = "sharepoint"
SALESFORCE_BLOB_ROOT = "salesforce"


def source_patterns(enabled_sources: Optional[Dict[str, bool]]) -> Optional[list[str]]:
    """Resolve enabled source names for an enabled-source selection.

    Returns ``None`` when no selection is present (``enabled_sources is None``)
    so callers apply no restriction and preserve the legacy "search all"
    behaviour for jobs created before this feature.

    Returns a (possibly empty) list when a selection dict is present. An empty
    list means no sources are enabled, in which case callers must match zero
    rows rather than falling back to searching every source.
    """
    if enabled_sources is None:
        return None
    sources: list[str] = []
    for source, enabled in enabled_sources.items():
        if enabled:
            sources.append(source)
    return sources


def sharepoint_blob_tag(tag: str) -> str:
    """Blob folder for SharePoint files: sharepoint/{tag}/filename."""
    normalized = tag.strip().strip("/")
    if normalized.lower().startswith(f"{SHAREPOINT_BLOB_ROOT.lower()}/"):
        return normalized
    return f"{SHAREPOINT_BLOB_ROOT}/{normalized}"


def salesforce_blob_tag(tag: str) -> str:
    """Blob folder for Salesforce files: salesforce/{tag}/filename."""
    normalized = tag.strip().strip("/")
    if normalized.lower().startswith(f"{SALESFORCE_BLOB_ROOT.lower()}/"):
        return normalized
    return f"{SALESFORCE_BLOB_ROOT}/{normalized}"


def blob_tag_for_upload(tag: str, blob_folder_prefix: Optional[str] = None) -> str:
    """Resolve blob virtual-folder path from ingestion tag and optional prefix."""
    normalized = tag.strip().strip("/")
    if blob_folder_prefix:
        prefix = blob_folder_prefix.strip().strip("/")
        if normalized.lower().startswith(f"{prefix.lower()}/"):
            return normalized
        return f"{prefix}/{normalized}"
    return normalized


def format_source_path(source_path: str, source_prefix: Optional[str] = None) -> str:
    """Normalize and optionally prefix source paths (e.g. sharepoint/01/file.docx)."""
    normalized = source_path.strip().replace("\\", "/").lstrip("/")
    if not source_prefix:
        return normalized
    prefix = source_prefix.strip().strip("/")
    if normalized.lower().startswith(f"{prefix.lower()}/"):
        return normalized
    return f"{prefix}/{normalized}"


def build_document_path_fields(
    source_path: str,
    blob_url: Optional[str] = None,
    *,
    source_prefix: Optional[str] = None,
    blob_base_path: Optional[str] = None,
    data_source: Optional[str] = None,
) -> Dict[str, Optional[str]]:
    """
    path / doc_source_path = where the file came from (local or SharePoint).
    doc_target_path = full blob URL of the processed artifact (if uploaded).
    data_source = authoritative source identifier (sharepoint, salesforce, local_drive).
    """
    formatted_source = format_source_path(source_path, source_prefix)
    return {
        "path": formatted_source,
        "doc_source_path": formatted_source,
        "doc_target_path": blob_url,
        "data_source": data_source,
    }


def attach_blob_url_to_metadata(
    metadata: Optional[Dict[str, Any]],
    blob_url: Optional[str],
) -> Dict[str, Any]:
    meta = dict(metadata) if metadata else {}
    if blob_url:
        meta["blob_url"] = blob_url
    return meta
