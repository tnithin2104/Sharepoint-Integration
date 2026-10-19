# sharepoint-integration

Standalone Microsoft Graph SharePoint integration for reading files from one
SharePoint site and uploading artifacts to another.

This package was migrated from a larger host application. It contains only the
SharePoint functionality — no RFP pipeline, no database, no embeddings, no
blob storage, no Salesforce integration.

## Features

- **Read** — discover files by tag, count files, download file content,
  resolve files by relative path or sharing URL.
- **Upload** — upload artifacts (e.g. ZIP files) to an output SharePoint site,
  organized by user and job folder, with resumable chunked upload for large
  files.
- **Delta sync** — detect new/modified/deleted files in a document library
  with the Microsoft Graph Delta API, driven by webhook notifications.
- **Optional FastAPI router** (`router.py`) exposing the read endpoints over
  HTTP. The core package does **not** depend on FastAPI.
- **Webhook listener** — receives change notifications from Microsoft Graph
  when files or folders in a SharePoint document library change.

## Installation

```bash
uv sync
# Optional, for the FastAPI router:
uv sync --extra fastapi
```

## Configuration

Copy `.env.example` to `.env` and fill in real values:

```bash
cp .env.example .env
```

### Required variables:

| Variable                   | Description                                  |
| -------------------------- | -------------------------------------------- |
| `SHAREPOINT_TENANT_ID`     | Azure AD tenant ID                           |
| `SHAREPOINT_CLIENT_ID`     | Azure AD application (client) ID             |
| `SHAREPOINT_CLIENT_SECRET` | Azure AD application secret                  |
| `SHAREPOINT_DOMAIN`        | SharePoint domain, e.g. `contoso.sharepoint.com` |
| `SHAREPOINT_SITE_PATH`     | Site path, e.g. `/sites/DemoProposalHub`    |

### Source / READ site:

| Variable                   | Description                                  |
| -------------------------- | -------------------------------------------- |
| `SHAREPOINT_LIBRARY_NAME`  | Document library name (default `Documents`) |
| `SHAREPOINT_BASE_FOLDER`   | Base folder relative to the library root     |

### Output / WRITE site:

| Variable                        | Description                                  |
| ------------------------------- | -------------------------------------------- |
| `SHAREPOINT_OUTPUT_DOMAIN`      | Output SharePoint domain                     |
| `SHAREPOINT_OUTPUT_SITE_PATH`   | Output site path                             |
| `SHAREPOINT_OUTPUT_LIBRARY`     | Output document library name                 |

### Webhook configuration:

| Variable                           | Description                                  |
| ---------------------------------- | -------------------------------------------- |
| `SHAREPOINT_WEBHOOK_CLIENT_STATE`  | Secret value used to validate incoming notifications (optional but recommended for security) |

### Delta sync configuration:

| Variable                   | Default                   | Description                              |
| -------------------------- | ------------------------- | ---------------------------------------- |
| `SHAREPOINT_DELTA_STATE_FILE` | `.delta_sync_state.json` (project root) | Durable file holding the `@odata.deltaLink` bookmark and known item IDs |

Set `SHAREPOINT_LIBRARY_NAME` to the document library to watch
(e.g. `graph-api-tests`). The delta sync tracks the same library
that `SharePointService` reads from.

### Optional:

| Variable                       | Default | Description                       |
| ------------------------------ | ------- | --------------------------------- |
| `SHAREPOINT_SSL_VERIFY`        | `true`  | Verify TLS certificates           |
| `SHAREPOINT_CONNECT_TIMEOUT`   | `30`    | Connect timeout (seconds)         |
| `SHAREPOINT_READ_TIMEOUT`      | `120`   | Read timeout (seconds)            |
| `SHAREPOINT_MAX_RETRIES`       | `5`     | HTTP retry count                  |
| `SHAREPOINT_POOL_CONNECTIONS`  | `20`    | HTTP connection pool size         |
| `SHAREPOINT_POOL_MAXSIZE`      | `50`    | HTTP max connections              |

## Usage

### Verify connection

```bash
uv run python verify_connection.py
```

Runs a smoke test: configuration load, token acquisition, site/drive
resolution, and file counting.

### Programmatic

```python
import asyncio
from sharepoint_integration import SharePointService

async def main():
    service = SharePointService()
    files = await service.discover_tagged_files()
    print(f"Found {len(files)} files")

asyncio.run(main())
```

### FastAPI router

```python
from fastapi import FastAPI
from router import router

app = FastAPI()
app.include_router(router)
```

Endpoints:

- `GET /sharepoint/health`
- `GET /sharepoint/files`
- `GET /sharepoint/files/count`
- `GET /sharepoint/files/{relative_path}`
- `GET /sharepoint/delta/status`
- `POST /sharepoint/delta/initialize`
- `POST /sharepoint/delta/sync`

### Webhook endpoint

Microsoft Graph change notifications can be received at:

- `POST /webhooks/sharepoint`

This endpoint supports:
1. **Validation token handshake** - Responds with the validation token as plain text when Microsoft Graph validates the notification URL
2. **Change notifications** - Processes file/folder change notifications from Microsoft Graph
3. **Client state validation** - Optionally validates the `clientState` value against `SHAREPOINT_WEBHOOK_CLIENT_STATE` environment variable
4. **Quick acknowledgment** - Returns HTTP 202 Accepted immediately after validation (does not perform lengthy processing)

Notifications are only *signals*: each notification schedules a
background Delta API poll (see [Delta synchronization](#delta-synchronization)).

### Delta synchronization

The integration detects file/folder changes with the Microsoft Graph
Delta API (`GET /drives/{drive_id}/root/delta`). The tracked library is
the source library configured via `SHAREPOINT_LIBRARY_NAME` (e.g.
`graph-api-tests`).

**1. Initialize the baseline (one time)**

```bash
curl -X POST http://127.0.0.1:8000/sharepoint/delta/initialize
```

This enumerates the whole library, follows every `@odata.nextLink`
page, records all existing item IDs as "known" (existing files are
**not** reported as changes), and — only after the enumeration
completes successfully — saves the final `@odata.deltaLink` as the
initial bookmark. The baseline is never repeated on application
restart; if a notification arrives before the baseline exists, the
background worker performs it automatically first.

**2. Trigger a synchronization**

- Automatically: every webhook notification to
  `POST /webhooks/sharepoint` schedules a background poll.
- Manually: `POST /sharepoint/delta/sync` polls now and returns the
  classified changes:

```bash
curl -X POST "http://127.0.0.1:8000/sharepoint/delta/sync"
# Optional time window (see caveat below):
curl -X POST "http://127.0.0.1:8000/sharepoint/delta/sync?since=2026-10-01T00:00:00Z&until=2026-10-02T00:00:00Z"
```

**3. Inspect state**

```bash
curl http://127.0.0.1:8000/sharepoint/delta/status
```

**Where the bookmark is stored**

The `@odata.deltaLink` bookmark, the known item IDs, sync timestamps,
and change counts are persisted in `SHAREPOINT_DELTA_STATE_FILE`
(default `.delta_sync_state.json` next to the project root, atomically
written). A JSONL change log (`.delta_sync_state.changes.jsonl`)
records the classified metadata of every synchronization run. Both
survive application restarts.

**Classification**

- Item ID not seen before → `created`
- Item ID already known → `updated`
- `deleted` facet present → `deleted`

Only item metadata (id, name, path, size, timestamps, webUrl) is
returned/logged. File content is never downloaded, and no parsing,
chunking, embeddings, or vector-database updates happen here.

**Time windows — important caveat**

The Delta API is checkpoint-based, not a general date-range query:
each poll returns the changes since the previous checkpoint and does
not guarantee that every individual event is returned in
chronological order. The optional `since`/`until` parameters only
filter the items returned by the *current* poll using the timestamps
available on each item (`lastModifiedDateTime`, falling back to
`createdDateTime`); they do **not** guarantee a complete event
history for an arbitrary timeframe. Synchronization timestamps are
recorded on every run (`last_sync_started_at`,
`last_sync_completed_at`) for correlation.

**Reliability guarantees**

- The bookmark only advances after every page has been retrieved and
  all change metadata has been persisted; a failed or incomplete run
  never overwrites a valid bookmark (retry the same window next time).
- Duplicate or concurrent webhook notifications are coalesced into a
  single poll; a notification arriving during a baseline or an active
  sync triggers a follow-up poll so nothing is lost.
- An expired/invalid delta link (HTTP 410, or a delta-token error)
  triggers a safe re-baseline: the library is re-enumerated, existing
  files are recorded as known (not reported as new), and only items
  that vanished are reported as deletions.

### Running the API

Start the development server with auto-reload:

```bash
uv run uvicorn main:app --reload
```

Open `http://127.0.0.1:8000/docs` for the interactive Swagger UI.

### Verifying SharePoint data

1. **Health check** — `GET /sharepoint/health` returns `{"status":"ok"}` without touching SharePoint.
2. **File count** — `GET /sharepoint/files/count` returns a JSON object mapping each tag folder to a file count plus `TOTAL`. This confirms the source site and library resolve correctly.
3. **List files** — `GET /sharepoint/files` returns the full list of discovered files with metadata (`name`, `path`, `tag`, `size_bytes`, `mime_type`, `download_url`, `web_url`, `file_id`, `created`, `modified`).
4. **Single file** — `GET /sharepoint/files/{relative_path}` fetches one file by path or sharing URL.
5. **Download via code** — call `download_file_to_temp(file)` or `get_file_content(file)` from `sharepoint_integration.service` to fetch bytes.

All successful actions and errors are logged. Check the console output to confirm token acquisition, site/drive resolution, folder traversal, and download sizes.

## Layout

```
sharepoint-integration/
├── sharepoint_integration/
│   ├── __init__.py    # public package API
│   ├── config.py      # environment-driven configuration
│   ├── service.py     # SharePoint read service (Graph API)
│   ├── delta_sync.py  # Graph Delta API change detection
│   ├── upload.py      # SharePoint upload service (Graph API)
│   └── paths.py       # path helpers
├── router.py          # optional FastAPI adapter
├── main.py            # FastAPI application entry point (includes webhook endpoint)
├── verify_connection.py
├── pyproject.toml
├── .env.example
└── README.md
```

## Azure AD application permissions

The integration uses **application-level** (app-only) permissions via the
Azure AD client-credentials flow. Grant the following on the target
SharePoint sites in the Azure portal:

| Permission             | Type         | Where needed |
| ---------------------- | ------------ | ------------ |
| `Sites.Read.All`       | Application  | Source / READ site |
| `Sites.ReadWrite.All`  | Application  | Output / WRITE site |

After granting permissions, an admin must consent to them.

## Verification

### Connection smoke test

```bash
uv run python verify_connection.py
```

This checks configuration load, token acquisition, site/drive resolution,
and basic file access against the configured source site.

### Source-site READ verification

Run `uv run python verify_connection.py` and confirm the source site resolves
and file counts are returned.

### Output-site WRITE verification

Use the upload service against a safe test destination (see below) before
uploading to a production folder.

### Safe upload test

```python
import asyncio
from sharepoint_integration.upload import SharePointUploadService

async def main():
    svc = SharePointUploadService()
    result = await svc.upload_artifacts(
        user_email="test.user@example.com",
        job_id="test-job-123",
        file_name="test_upload.txt",
        zip_bytes=b"hello world",
        zip_filename="test_upload.txt",
    )
    print(result)

asyncio.run(main())
```

Upload only to a dedicated test folder. Do not overwrite existing files
or write into a production location.

## Testing

```bash
uv run pytest
```

Test suite includes tests for the webhook endpoint and the
Delta API synchronization under `tests/` (all Graph calls are
mocked; no real credentials or network access are needed):

- `tests/test_webhook.py` — webhook validation handshake,
  clientState checks, malformed payloads.
- `tests/test_delta_sync.py` — baseline initialization, multi-page
  enumeration, create/update/delete classification, duplicate and
  concurrent notifications, failed requests, restart persistence,
  expired bookmark re-baselining, time-window filtering, and the
  change log.

## Webhook Implementation Details

The webhook endpoint implements Microsoft Graph's change notification requirements:

### Validation Handshake
When Microsoft Graph validates the notification URL, it sends a `validationToken` query parameter.
The endpoint returns this token as plain text with HTTP 200 and content type `text/plain`.

### Notification Processing
For actual notifications:
- Parses the Microsoft Graph notification payload
- Validates the `clientState` value against `SHAREPOINT_WEBHOOK_CLIENT_STATE` (when configured)
- Logs useful, non-sensitive metadata
- Returns HTTP 202 Accepted promptly

### Security
- Uses constant-time comparison for clientState validation to prevent timing attacks
- Never logs sensitive values like tokens or clientState
- Handles malformed JSON and invalid payloads gracefully

### Limitations
Currently, the endpoint:
- Acknowledges notifications quickly but does not perform document processing
- Does not implement downloading, generating embeddings, or indexing
- Leaves a clear integration point for adding a durable queue in future implementations