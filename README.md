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
- **Optional FastAPI router** (`router.py`) exposing the read endpoints over
  HTTP. The core package does **not** depend on FastAPI.

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

Required variables:

| Variable                   | Description                                  |
| -------------------------- | -------------------------------------------- |
| `SHAREPOINT_TENANT_ID`     | Azure AD tenant ID                           |
| `SHAREPOINT_CLIENT_ID`     | Azure AD application (client) ID             |
| `SHAREPOINT_CLIENT_SECRET` | Azure AD application secret                  |
| `SHAREPOINT_DOMAIN`        | SharePoint domain, e.g. `contoso.sharepoint.com` |
| `SHAREPOINT_SITE_PATH`     | Site path, e.g. `/sites/DemoProposalHub`    |

Source / READ site:

| Variable                   | Description                                  |
| -------------------------- | -------------------------------------------- |
| `SHAREPOINT_LIBRARY_NAME`  | Document library name (default `Documents`) |
| `SHAREPOINT_BASE_FOLDER`   | Base folder relative to the library root     |

Output / WRITE site:

| Variable                        | Description                                  |
| ------------------------------- | -------------------------------------------- |
| `SHAREPOINT_OUTPUT_DOMAIN`      | Output SharePoint domain                     |
| `SHAREPOINT_OUTPUT_SITE_PATH`   | Output site path                             |
| `SHAREPOINT_OUTPUT_LIBRARY`     | Output document library name                 |

Optional:

| Variable                       | Default | Description                       |
| ------------------------------ | ------- | --------------------------------- |
| `SHAREPOINT_SSL_VERIFY`        | `true`  | Verify TLS certificates           |
| `SHAREPOINT_CONNECT_TIMEOUT`   | `30`    | Connect timeout (seconds)         |
| `SHAREPOINT_READ_TIMEOUT`      | `120`   | Read timeout (seconds)            |
| `SHAREPOINT_MAX_RETRIES`       | `5`     | HTTP retry count                  |

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
│   ├── upload.py      # SharePoint upload service (Graph API)
│   └── paths.py       # path helpers
├── router.py          # optional FastAPI adapter
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

No test suite is bundled yet; add tests under `tests/` as needed.