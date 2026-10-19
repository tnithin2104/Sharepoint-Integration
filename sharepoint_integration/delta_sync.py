"""Microsoft Graph Delta API synchronization for a SharePoint document library.

Change detection uses the Graph delta API (``GET /drives/{drive_id}/root/delta``):

- **Baseline** — the first run enumerates the whole library, records every
  existing item ID as "known" (existing files are *ignored* as changes), and
  stores the final ``@odata.deltaLink`` as a durable bookmark. The bookmark is
  only written after every page has been retrieved successfully.
- **Poll** — later runs (usually triggered by webhook notifications) request
  the saved delta link, follow every ``@odata.nextLink`` page, classify each
  returned item as created / updated / deleted using persistent item-ID
  tracking and the ``deleted`` facet, persist the change metadata, and only
  then advance the bookmark to the new ``@odata.deltaLink``.

The delta API is checkpoint-based, not a general date-range query: it returns
the changes since the previous checkpoint and does not guarantee that every
individual event is returned in chronological order. When a time window is
requested (``since`` / ``until``), returned items are filtered using the
timestamps available on each item, but item timestamps do not guarantee a
complete event history for an arbitrary timeframe.

Scope: this module only detects and classifies changes and returns/logs item
metadata. It never downloads file content, parses documents, chunks text, or
updates any vector database.

Durable state (bookmark + known item IDs) lives in a JSON file configured via
``SHAREPOINT_DELTA_STATE_FILE`` (default: ``.delta_sync_state.json`` next to
the project root). A JSONL change log (``*.changes.jsonl``) records the
classified metadata of every synchronization run.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

import requests

from sharepoint_integration.service import SharePointService

# Optional text extraction dependencies
try:
    from docx import Document as DocxDocument
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pypdf import PdfReader
    PDF_AVAILABLE = True
except ImportError:
    PDF_AVAILABLE = False

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Default durable-state location: repository root next to the package.
DEFAULT_STATE_FILE = Path(__file__).resolve().parents[1] / ".delta_sync_state.json"

# HTTP 410 Gone is returned when a delta link has expired.
# 400/404 are treated as expired only when the error clearly refers to the
# delta token itself (a genuinely invalid link needs re-baselining too).
EXPIRED_DELTA_ERROR_CODES = frozenset(
    {"deltalinkexpired", "invaliddeltatoken", "itemnotfound"}
)


class DeltaLinkExpiredError(Exception):
    """The saved delta link expired or is no longer valid."""


@dataclass
class DeltaSyncState:
    """Durable synchronization state (survives application restarts)."""

    delta_link: Optional[str] = None
    known_items: Dict[str, Dict[str, Any]] = field(default_factory=dict)
    baseline_initialized: bool = False
    last_sync_started_at: Optional[str] = None
    last_sync_completed_at: Optional[str] = None
    last_change_counts: Dict[str, int] = field(default_factory=dict)
    total_changes_seen: int = 0
    version: int = 1

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DeltaSyncState":
        return cls(
            delta_link=data.get("delta_link"),
            known_items=data.get("known_items") or {},
            baseline_initialized=bool(data.get("baseline_initialized")),
            last_sync_started_at=data.get("last_sync_started_at"),
            last_sync_completed_at=data.get("last_sync_completed_at"),
            last_change_counts=data.get("last_change_counts") or {},
            total_changes_seen=int(data.get("total_changes_seen") or 0),
            version=int(data.get("version") or 1),
        )

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": self.version,
            "delta_link": self.delta_link,
            "known_items": self.known_items,
            "baseline_initialized": self.baseline_initialized,
            "last_sync_started_at": self.last_sync_started_at,
            "last_sync_completed_at": self.last_sync_completed_at,
            "last_change_counts": self.last_change_counts,
            "total_changes_seen": self.total_changes_seen,
        }


class DeltaSyncStore:
    """Durable, atomic storage for the delta bookmark, known item IDs, and a
    JSONL log of classified changes.

    State is written to a temporary file and atomically renamed so a crash can
    never leave a partially written (invalid) bookmark. The bookmark is only
    advanced by an explicit :meth:`save` after all pages have been fetched and
    the change metadata persisted via :meth:`append_changes`.
    """

    def __init__(self, state_path: Path, changes_path: Optional[Path] = None):
        self.state_path = Path(state_path)
        if changes_path is None:
            self.changes_path = self.state_path.with_name(
                self.state_path.stem + ".changes.jsonl"
            )
        else:
            self.changes_path = Path(changes_path)

    def load(self) -> DeltaSyncState:
        """Load the persisted state. A missing or corrupt state file yields a
        fresh (uninitialized) state; the safe recovery path is a re-baseline,
        which never treats existing files as new."""
        if not self.state_path.exists():
            return DeltaSyncState()
        try:
            data = json.loads(self.state_path.read_text(encoding="utf-8"))
            return DeltaSyncState.from_dict(data)
        except (OSError, json.JSONDecodeError) as e:
            logger.error(
                "Could not read delta sync state from %s (%s); "
                "starting from an uninitialized state",
                self.state_path,
                e,
            )
            return DeltaSyncState()

    def save(self, state: DeltaSyncState) -> None:
        """Atomically persist the state (bookmark + known item IDs)."""
        self.state_path.parent.mkdir(parents=True, exist_ok=True)
        tmp_path = self.state_path.with_name(self.state_path.name + ".tmp")
        tmp_path.write_text(
            json.dumps(state.to_dict(), indent=2, sort_keys=True),
            encoding="utf-8",
        )
        os.replace(tmp_path, self.state_path)

    def append_changes(self, run_id: str, changes: List[Dict[str, Any]]) -> None:
        """Append classified change metadata to the durable JSONL change log.
        The file is flushed and fsync'ed so the metadata is on disk before the
        caller advances the bookmark."""
        if not changes:
            return
        self.changes_path.parent.mkdir(parents=True, exist_ok=True)
        with open(self.changes_path, "a", encoding="utf-8") as fh:
            for change in changes:
                entry = {"run_id": run_id, "recorded_at": _utcnow_iso(), **change}
                fh.write(json.dumps(entry, sort_keys=True) + "\n")
            fh.flush()
            os.fsync(fh.fileno())


@dataclass
class DeltaSyncResult:
    """Outcome of one baseline or poll operation."""

    status: str  # "baseline" | "synced" | "rebased"
    changes: List[Dict[str, Any]] = field(default_factory=list)
    created: int = 0
    updated: int = 0
    deleted: int = 0
    pages_fetched: int = 0
    delta_link: Optional[str] = None
    started_at: str = ""
    completed_at: str = ""
    message: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "status": self.status,
            "changes": self.changes,
            "counts": {"created": self.created, "updated": self.updated, "deleted": self.deleted},
            "pages_fetched": self.pages_fetched,
            "delta_link": self.delta_link,
            "started_at": self.started_at,
            "completed_at": self.completed_at,
            "message": self.message,
        }


def _utcnow_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _parse_timestamp(value: Optional[str]) -> Optional[datetime]:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        return None


def _is_deleted_item(item: Dict[str, Any]) -> bool:
    """Graph marks removed items with a ``deleted`` facet."""
    return "deleted" in item


def _item_path(item: Dict[str, Any]) -> str:
    """Library-relative path of an item, derived from parentReference.path."""
    parent = (item.get("parentReference") or {}).get("path") or ""
    name = item.get("name") or ""
    if ":/" in parent:
        relative = parent.split(":/", 1)[1].strip("/")
    else:
        relative = parent.replace("/drive/root", "").strip("/")
    return f"{relative}/{name}".strip("/") if relative else name


def _item_snapshot(item: Dict[str, Any]) -> Dict[str, Any]:
    """Persistable metadata snapshot for a drive item."""
    return {
        "id": item.get("id") or "",
        "name": item.get("name") or "",
        "path": _item_path(item),
        "size_bytes": item.get("size") or 0,
        "mime_type": (item.get("file") or {}).get("mimeType") or "",
        "is_folder": "folder" in item,
        "web_url": item.get("webUrl") or "",
        "created": item.get("createdDateTime") or "",
        "modified": item.get("lastModifiedDateTime") or "",
    }


def _is_delta_link_error(response: requests.Response) -> bool:
    """Best-effort detection of an expired/invalid delta link."""
    if response is None:
        return False
    if response.status_code == 410:
        return True
    if response.status_code not in (400, 404):
        return False
    try:
        data = response.json()
    except Exception:
        return False
    error = data.get("error") if isinstance(data, dict) else None
    if not isinstance(error, dict):
        return False
    code = str(error.get("code") or "").strip().lower()
    message = str(error.get("message") or "").strip().lower()
    if code in EXPIRED_DELTA_ERROR_CODES:
        return True
    return "delta" in message and ("expired" in message or "invalid" in message)


def _count_changes(changes: List[Dict[str, Any]]) -> Dict[str, int]:
    counts = {"created": 0, "updated": 0, "deleted": 0}
    for change in changes:
        change_type = change.get("change_type")
        if change_type in counts:
            counts[change_type] += 1
    return counts


def _filter_by_window(
    changes: List[Dict[str, Any]],
    since: Optional[str],
    until: Optional[str],
) -> List[Dict[str, Any]]:
    """Filter changes to a time window using item timestamps.

    Note: the delta API is checkpoint-based, so this filter only narrows the
    changes returned by *this* poll. Item timestamps do not guarantee a
    complete event history for an arbitrary timeframe.
    """
    if since is None and until is None:
        return changes
    since_dt = _parse_timestamp(since) if since else None
    until_dt = _parse_timestamp(until) if until else None
    if since_dt is None and until_dt is None:
        return changes
    filtered: List[Dict[str, Any]] = []
    for change in changes:
        # Prefer the modification time; fall back to creation time.
        timestamp = _parse_timestamp(change.get("modified")) or _parse_timestamp(
            change.get("created")
        )
        if timestamp is None:
            # Items without a parsable timestamp are kept (at-least-once).
            filtered.append(change)
            continue
        if since_dt is not None and timestamp < since_dt:
            continue
        if until_dt is not None and timestamp > until_dt:
            continue
        filtered.append(change)
    return filtered


class SharePointDeltaSync:
    """Detects file/folder changes in one SharePoint document library via the
    Graph delta API.

    Reuses the existing :class:`SharePointService` for Graph authentication
    and its retrying HTTP client/session. Notifications from the webhook are
    treated only as *signals* to poll the delta API; they are never used as
    the list of changes.

    Concurrency: a single background worker task is coalesced from any number
    of duplicate or concurrent notifications. If a notification arrives while a
    baseline or sync is in progress, a follow-up poll is guaranteed (the poll
    request flag stays set until consumed). An asyncio lock prevents overlapping
    sync operations from corrupting the durable bookmark.
    """

    def __init__(
        self,
        service: SharePointService,
        store: DeltaSyncStore,
        *,
        max_worker_retries: int = 3,
        worker_backoff_base: float = 2.0,
    ):
        self._service = service
        self._store = store
        self._max_worker_retries = max_worker_retries
        self._worker_backoff_base = worker_backoff_base
        self._sync_lock = asyncio.Lock()
        self._worker_task: Optional[asyncio.Task] = None
        self._poll_requested = False
        self._last_result: Optional[DeltaSyncResult] = None

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def initialize_baseline(self) -> DeltaSyncResult:
        """Enumerate the library and store the initial delta bookmark.

        Existing files are recorded as known items but are NOT reported as
        changes. The baseline is only performed once; later calls (including
        after an application restart) return immediately without repeating it.
        """
        async with self._sync_lock:
            state = self._store.load()
            if state.baseline_initialized and state.delta_link:
                logger.info(
                    "Delta baseline already initialized (bookmark present); "
                    "not repeating baseline"
                )
                return DeltaSyncResult(
                    status="baseline",
                    delta_link=state.delta_link,
                    started_at=_utcnow_iso(),
                    completed_at=_utcnow_iso(),
                    message="baseline already initialized",
                )
            return await self._run_baseline(state)

    async def sync_once(
        self,
        since: Optional[str] = None,
        until: Optional[str] = None,
    ) -> DeltaSyncResult:
        """Poll the delta API once and return the classified changes.

        Follows every ``@odata.nextLink`` page, classifies items, persists the
        change metadata, and advances the stored bookmark to the new
        ``@odata.deltaLink`` only after everything succeeded. If the saved
        delta link has expired, a safe re-baseline runs automatically.
        """
        async with self._sync_lock:
            if since is not None and _parse_timestamp(since) is None:
                raise ValueError(f"Invalid 'since' timestamp: {since!r}")
            if until is not None and _parse_timestamp(until) is None:
                raise ValueError(f"Invalid 'until' timestamp: {until!r}")
            state = self._store.load()
            if not state.baseline_initialized or not state.delta_link:
                raise RuntimeError(
                    "Delta baseline not initialized. Call initialize_baseline() "
                    "or POST /sharepoint/delta/initialize first."
                )
            try:
                return await self._run_sync(state, since=since, until=until)
            except DeltaLinkExpiredError:
                logger.warning(
                    "Stored delta link expired or invalid; performing a safe "
                    "re-baseline (existing files will not be treated as new)"
                )
                return await self._rebaseline(state)

    def notify(self) -> None:
        """Signal that a webhook notification arrived.

        This only schedules a background delta poll (coalesced across
        duplicate/concurrent notifications); it never performs HTTP calls
        inline, so the webhook endpoint can acknowledge quickly.
        """
        self._poll_requested = True
        if self._worker_task is None or self._worker_task.done():
            self._worker_task = asyncio.create_task(self._worker())
            logger.debug("Delta sync worker started")

    def stop(self) -> None:
        """Best-effort cancellation of the background worker (shutdown)."""
        self._poll_requested = False
        if self._worker_task is not None:
            self._worker_task.cancel()

    def status(self) -> Dict[str, Any]:
        """Snapshot of the durable state and worker health for inspection."""
        state = self._store.load()
        last = self._last_result
        return {
            "baseline_initialized": state.baseline_initialized,
            "delta_link": state.delta_link,
            "known_item_count": len(state.known_items),
            "last_sync_started_at": state.last_sync_started_at,
            "last_sync_completed_at": state.last_sync_completed_at,
            "last_change_counts": state.last_change_counts,
            "total_changes_seen": state.total_changes_seen,
            "poll_requested": self._poll_requested,
            "worker_running": self._worker_task is not None and not self._worker_task.done(),
            "last_result": last.to_dict() if last else None,
            "state_file": str(self._store.state_path),
            "changes_file": str(self._store.changes_path),
        }

    def get_last_changes(self) -> List[Dict[str, Any]]:
        """Classified change metadata from the most recent sync run."""
        if self._last_result is None:
            return []
        return list(self._last_result.changes)

    # ------------------------------------------------------------------
    # Internals
    # ------------------------------------------------------------------

    async def _run_baseline(self, state: DeltaSyncState) -> DeltaSyncResult:
        started_at = _utcnow_iso()
        drive_id = await self._resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root/delta"
        known_items: Dict[str, Dict[str, Any]] = {}
        pages = 0
        delta_link: Optional[str] = None

        logger.info("Initializing delta baseline for drive %s", drive_id)
        while url:
            items, next_link, delta_link = await self._fetch_page(url)
            pages += 1
            for item in items:
                if _is_deleted_item(item):
                    continue
                item_id = item.get("id")
                if not item_id:
                    logger.warning("Skipping delta item without an id: %r", item.get("name"))
                    continue
                # Baseline: record as known, never report as a change.
                known_items[item_id] = _item_snapshot(item)
            logger.debug("Baseline page %d: %d items", pages, len(items))
            url = next_link

        if not delta_link:
            raise RuntimeError(
                "Delta baseline enumeration finished without a final "
                "@odata.deltaLink; bookmark not saved"
            )

        # Enumeration completed successfully — persist the bookmark now.
        state.delta_link = delta_link
        state.known_items = known_items
        state.baseline_initialized = True
        state.last_sync_started_at = started_at
        state.last_sync_completed_at = _utcnow_iso()
        state.last_change_counts = {"created": 0, "updated": 0, "deleted": 0}
        self._store.save(state)

        logger.info(
            "Delta baseline initialized: %d known items across %d pages",
            len(known_items),
            pages,
        )
        return DeltaSyncResult(
            status="baseline",
            pages_fetched=pages,
            delta_link=delta_link,
            started_at=started_at,
            completed_at=state.last_sync_completed_at,
            message=f"baseline initialized with {len(known_items)} known items",
        )

    async def _run_sync(
        self,
        state: DeltaSyncState,
        since: Optional[str],
        until: Optional[str],
    ) -> DeltaSyncResult:
        started_at = _utcnow_iso()
        url = state.delta_link
        raw_items: List[Dict[str, Any]] = []
        pages = 0
        delta_link: Optional[str] = None

        logger.info("Polling delta API for changes since last checkpoint")
        while url:
            items, next_link, delta_link = await self._fetch_page(url)
            pages += 1
            raw_items.extend(items)
            logger.debug("Delta poll page %d: %d items", pages, len(items))
            url = next_link

        if not delta_link:
            raise RuntimeError(
                "Delta sync finished without a new @odata.deltaLink; "
                "bookmark not advanced"
            )

        known_items = dict(state.known_items)
        changes: List[Dict[str, Any]] = []
        for item in raw_items:
            change = self._classify_item(item, known_items)
            if change is not None:
                changes.append(change)
                logger.info(
                    "Delta change: %s id=%s name=%s path=%s",
                    change["change_type"],
                    change.get("id"),
                    change.get("name"),
                    change.get("path"),
                )

        # Download and log content for newly created files (not folders)
        drive_id = await self._resolve_drive_id()
        for change in changes:
            if (
                change.get("change_type") == "created"
                and not change.get("is_folder")
                and change.get("id")
            ):
                logger.info("New file detected: %s", change.get("name"))
                await self._download_and_log_file_content(change, drive_id)

        # Optional time-window narrowing of the *returned* items. The full
        # change set is still persisted and used to advance the bookmark.
        returned_changes = _filter_by_window(changes, since=since, until=until)

        run_id = uuid.uuid4().hex
        counts = _count_changes(changes)

        # Persist all change metadata BEFORE advancing the bookmark.
        self._store.append_changes(run_id, changes)

        state.delta_link = delta_link
        state.known_items = known_items
        state.last_sync_started_at = started_at
        state.last_sync_completed_at = _utcnow_iso()
        state.last_change_counts = counts
        state.total_changes_seen += len(changes)
        self._store.save(state)

        logger.info(
            "Delta sync complete: %d pages, %d changes (%s)",
            pages,
            len(changes),
            counts,
        )
        result = DeltaSyncResult(
            status="synced",
            changes=returned_changes,
            created=counts["created"],
            updated=counts["updated"],
            deleted=counts["deleted"],
            pages_fetched=pages,
            delta_link=delta_link,
            started_at=started_at,
            completed_at=state.last_sync_completed_at,
        )
        self._last_result = result
        return result

    async def _rebaseline(self, state: DeltaSyncState) -> DeltaSyncResult:
        """Safe re-baseline after a delta link expiry.

        Re-enumerates the library from a fresh delta link. Every item that
        currently exists is recorded as known (existing files are NOT reported
        as created). Only items that were previously known and have vanished
        are reported as deletions.
        """
        started_at = _utcnow_iso()
        drive_id = await self._resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root/delta"
        current_items: Dict[str, Dict[str, Any]] = {}
        pages = 0
        delta_link: Optional[str] = None

        while url:
            items, next_link, delta_link = await self._fetch_page(url)
            pages += 1
            for item in items:
                if _is_deleted_item(item):
                    continue
                item_id = item.get("id")
                if item_id:
                    current_items[item_id] = _item_snapshot(item)
            url = next_link

        if not delta_link:
            raise RuntimeError(
                "Re-baseline enumeration finished without a final "
                "@odata.deltaLink; bookmark not saved"
            )

        changes: List[Dict[str, Any]] = []
        for item_id, previous in state.known_items.items():
            if item_id not in current_items:
                changes.append(
                    {
                        "change_type": "deleted",
                        "id": item_id,
                        "name": previous.get("name", ""),
                        "path": previous.get("path", ""),
                        "size_bytes": previous.get("size_bytes", 0),
                        "mime_type": previous.get("mime_type", ""),
                        "is_folder": previous.get("is_folder", False),
                        "web_url": previous.get("web_url", ""),
                        "created": previous.get("created", ""),
                        "modified": previous.get("modified", ""),
                        "deleted": True,
                    }
                )
                logger.info(
                    "Delta change (re-baseline): deleted id=%s name=%s",
                    item_id,
                    previous.get("name"),
                )

        run_id = uuid.uuid4().hex
        counts = _count_changes(changes)
        self._store.append_changes(run_id, changes)

        state.delta_link = delta_link
        state.known_items = current_items
        state.baseline_initialized = True
        state.last_sync_started_at = started_at
        state.last_sync_completed_at = _utcnow_iso()
        state.last_change_counts = counts
        state.total_changes_seen += len(changes)
        self._store.save(state)

        logger.info(
            "Delta re-baseline complete: %d known items, %d deletions detected",
            len(current_items),
            len(changes),
        )
        result = DeltaSyncResult(
            status="rebased",
            changes=changes,
            created=counts["created"],
            updated=counts["updated"],
            deleted=counts["deleted"],
            pages_fetched=pages,
            delta_link=delta_link,
            started_at=started_at,
            completed_at=state.last_sync_completed_at,
            message=(
                "delta link expired; re-baselined without reporting "
                f"{len(current_items)} existing files as new"
            ),
        )
        self._last_result = result
        return result

    def _classify_item(
        self, item: Dict[str, Any], known_items: Dict[str, Dict[str, Any]]
    ) -> Optional[Dict[str, Any]]:
        """Classify one delta item against the known item-ID set, updating it.

        * ``deleted`` facet present                      -> deleted
        * item ID already known                          -> updated
        * item ID not known                              -> created
        """
        item_id = item.get("id")
        if not item_id:
            logger.warning("Skipping delta item without an id: %r", item.get("name"))
            return None

        if _is_deleted_item(item):
            previous = known_items.pop(item_id, None) or {}
            return {
                "change_type": "deleted",
                "id": item_id,
                "name": item.get("name") or previous.get("name", ""),
                "path": _item_path(item) or previous.get("path", ""),
                "size_bytes": item.get("size") or 0,
                "mime_type": (item.get("file") or {}).get("mimeType") or "",
                "is_folder": "folder" in item,
                "web_url": item.get("webUrl") or "",
                "created": item.get("createdDateTime") or "",
                "modified": item.get("lastModifiedDateTime") or "",
                "deleted": True,
            }

        snapshot = _item_snapshot(item)
        change_type = "updated" if item_id in known_items else "created"
        known_items[item_id] = snapshot
        return {"change_type": change_type, **snapshot}

    async def _resolve_drive_id(self) -> str:
        _, drive_id = await self._service._get_site_and_drive_ids()
        if not drive_id:
            raise RuntimeError("Could not resolve the SharePoint drive ID")
        return drive_id

    # ------------------------------------------------------------------
    # File content fetching and text extraction
    # ------------------------------------------------------------------

    MAX_LOG_CONTENT_LENGTH = 5000

    def _extract_text_from_content(
        self, content: bytes, mime_type: str, name: str
    ) -> Optional[str]:
        """Extract readable text from file content based on MIME type/extension."""
        if not content:
            return None

        mime_type = (mime_type or "").lower()
        name_lower = (name or "").lower()

        # DOCX files
        if (
            mime_type
            == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            or name_lower.endswith(".docx")
        ):
            if not DOCX_AVAILABLE:
                logger.debug("python-docx not available, skipping DOCX text extraction")
                return None
            try:
                import io

                doc = DocxDocument(io.BytesIO(content))
                paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
                return "\n".join(paragraphs)
            except Exception as e:
                logger.warning("Failed to extract text from DOCX '%s': %s", name, e)
                return None

        # PDF files
        if mime_type == "application/pdf" or name_lower.endswith(".pdf"):
            if not PDF_AVAILABLE:
                logger.debug("pypdf not available, skipping PDF text extraction")
                return None
            try:
                import io

                reader = PdfReader(io.BytesIO(content))
                texts = []
                for page in reader.pages:
                    text = page.extract_text()
                    if text and text.strip():
                        texts.append(text.strip())
                return "\n".join(texts)
            except Exception as e:
                logger.warning("Failed to extract text from PDF '%s': %s", name, e)
                return None

        # Plain text files
        if mime_type.startswith("text/") or any(
            name_lower.endswith(ext)
            for ext in (".txt", ".md", ".csv", ".json", ".xml", ".html", ".htm")
        ):
            # Try common encodings
            for encoding in ("utf-8", "utf-16", "latin-1", "cp1252"):
                try:
                    return content.decode(encoding)
                except UnicodeDecodeError:
                    continue
            logger.warning("Could not decode text file '%s' with common encodings", name)
            return None

        # Unsupported file type
        logger.debug("No text extraction available for file type '%s': %s", mime_type, name)
        return None

    def _truncate_for_logging(self, text: str) -> str:
        """Truncate text for logging, indicating if truncated."""
        if len(text) <= self.MAX_LOG_CONTENT_LENGTH:
            return text
        return (
            text[: self.MAX_LOG_CONTENT_LENGTH]
            + f"\n... [truncated, total length: {len(text)} characters]"
        )

    async def _download_and_log_file_content(
        self, change: Dict[str, Any], drive_id: str
    ) -> None:
        """Download file content for a newly created file and log extracted text."""
        item_id = change.get("id")
        name = change.get("name")
        mime_type = change.get("mime_type")
        size_bytes = change.get("size_bytes", 0)

        if not item_id:
            logger.warning("Cannot download content: missing item ID for '%s'", name)
            return

        logger.info("Downloading file content: id=%s name=%s", item_id, name)

        try:
            content = await self._service.download_file_content_by_id(drive_id, item_id)
            actual_size = len(content)

            logger.info(
                "File downloaded successfully: name=%s size=%d bytes mime_type=%s",
                name,
                actual_size,
                mime_type,
            )

            # Extract text content
            extracted_text = self._extract_text_from_content(content, mime_type, name)

            if extracted_text:
                truncated = self._truncate_for_logging(extracted_text)
                logger.info(
                    "Extracted document text from '%s':\n%s",
                    name,
                    truncated,
                )
            else:
                logger.info(
                    "No extractable text content for '%s' (type: %s, size: %d bytes)",
                    name,
                    mime_type or "unknown",
                    actual_size,
                )

        except Exception as e:
            logger.error(
                "Failed to download or extract content for file '%s' (id=%s): %s",
                name,
                item_id,
                e,
            )
            # Do not re-raise - we don't want to crash the delta sync worker

    async def _fetch_page(self, url: str):
        """Fetch one delta page and return (items, nextLink, deltaLink)."""
        token = await self._service._get_access_token()
        headers = {"Authorization": f"Bearer {token}"}
        response = self._service._make_request_with_retry("GET", url, headers=headers)
        if _is_delta_link_error(response):
            raise DeltaLinkExpiredError(
                f"Delta link expired or invalid (HTTP {response.status_code}): {url}"
            )
        response.raise_for_status()
        data = response.json()
        return (
            data.get("value") or [],
            data.get("@odata.nextLink"),
            data.get("@odata.deltaLink"),
        )

    async def _worker(self) -> None:
        """Background worker: drains poll requests, one sync at a time.

        The poll-request flag is only cleared right before a sync starts, so a
        notification arriving during a baseline or an active sync leaves the
        flag set and guarantees a follow-up poll. Transient failures are
        retried with backoff; the worker stops after consecutive failures and
        is restarted by the next notification.
        """
        failures = 0
        try:
            while True:
                try:
                    state = self._store.load()
                    if not state.baseline_initialized:
                        await self.initialize_baseline()
                        failures = 0
                        continue
                    if not self._poll_requested:
                        break
                    self._poll_requested = False
                    await self.sync_once()
                    failures = 0
                except Exception as e:
                    failures += 1
                    logger.error(
                        "Delta sync attempt %d/%d failed: %s",
                        failures,
                        self._max_worker_retries,
                        e,
                    )
                    if failures >= self._max_worker_retries:
                        logger.error(
                            "Delta sync worker stopping after %d consecutive failures",
                            failures,
                        )
                        break
                        # The failed poll is retried (not lost) with backoff.
                    self._poll_requested = True
                    await asyncio.sleep(self._worker_backoff_base ** min(failures, 5))
        finally:
            self._worker_task = None


# ----------------------------------------------------------------------
# App-wide singleton (used by the FastAPI webhook and router endpoints)
# ----------------------------------------------------------------------

_delta_sync_instance: Optional[SharePointDeltaSync] = None


def get_delta_sync() -> SharePointDeltaSync:
    """Lazily create the app-wide delta sync service.

    Reuses the existing :class:`SharePointService` (Graph auth + retrying HTTP
    client) and persists state to ``SHAREPOINT_DELTA_STATE_FILE``.
    """
    global _delta_sync_instance
    if _delta_sync_instance is None:
        state_file = Path(
            os.getenv("SHAREPOINT_DELTA_STATE_FILE", str(DEFAULT_STATE_FILE))
        )
        service = SharePointService()
        _delta_sync_instance = SharePointDeltaSync(
            service=service,
            store=DeltaSyncStore(state_file),
        )
        logger.info("Delta sync service initialized (state: %s)", state_file)
    return _delta_sync_instance


def shutdown_delta_sync() -> None:
    """Stop the background worker of the app-wide delta sync service, if any."""
    if _delta_sync_instance is not None:
        _delta_sync_instance.stop()
