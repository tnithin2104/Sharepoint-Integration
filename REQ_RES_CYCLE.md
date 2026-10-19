# SharePoint Integration — Request/Response Cycle & Graph API Reference

A complete, line-by-line documentation of every operation that happens in this
repository when a request flows through the system, from an incoming HTTP
call (FastAPI) or programmatic call, down through the service layer, into the
Microsoft Graph API, and back as a response.

---

## 1. Repository Overview

| Layer | File | Responsibility |
|---|---|---|
| Entry point | `main.py` | Creates the FastAPI app, attaches the lifespan handler, includes `router.py`. Hosts the `/webhooks/sharepoint` endpoint, which schedules a background Delta API poll on each notification. |
| HTTP adapter | `router.py` | FastAPI `APIRouter` (`/sharepoint` prefix). Lazily instantiates `SharePointService`. Exposes health, list, count, get-by-path, and delta sync endpoints (`/delta/initialize`, `/delta/sync`, `/delta/status`). |
| Config | `sharepoint_integration/config.py` | Reads all settings from environment variables. Provides `has_credentials()` and `require_credentials()` helpers. |
| Read service | `sharepoint_integration/service.py` | `SharePointService` — authenticates, resolves site/drive IDs, discovers files by tag, counts files, downloads file content, resolves sharing URLs. |
| Delta sync | `sharepoint_integration/delta_sync.py` | `SharePointDeltaSync` — Graph Delta API change detection: one-time baseline, checkpoint polling, create/update/delete classification, durable `@odata.deltaLink` bookmark + known item IDs, expired-link re-baselining, coalesced webhook-driven worker. |
| Write service | `sharepoint_integration/upload.py` | `SharePointUploadService` — resolves drive, ensures folders exist, uploads bytes (simple PUT for ≤4 MB, resumable session for >4 MB). |
| Data model | `sharepoint_integration/service.py` | `SharePointFile` dataclass — holds all metadata for a discovered/downloaded file. |
| Package API | `sharepoint_integration/__init__.py` | Re-exports `SharePointService`, `SharePointFile`, `SharePointUploadService`, `SharePointDeltaSync`, `DeltaSyncStore`, `config`. |
| Smoke test | `verify_connection.py` | Standalone script: 5-step verification of config → token → site → drive → basic access. |
| Dependencies | `pyproject.toml` / `requirements.txt` | `requests`, `urllib3`, `python-dotenv`; optional `fastapi` extra. |
| Config template | `.env.example` | All environment variable keys with empty values + comments. |

> **Note:** `sharepoint_integration/paths.py` was listed as a visible/open tab
> in the IDE but does **not exist** on disk. No code imports it. It appears to
> be a stale editor artifact.

---

## 2. Configuration (Environment Variables)

Loaded via `python-dotenv` from `.env` (located at the repository root, i.e.
`Path(__file__).resolve().parents[1] / ".env"` loaded in both `config.py`'s
package and `service.py`/`upload.py` directly).

### Authentication (required)

| Variable | Description |
|---|---|
| `SHAREPOINT_TENANT_ID` | Azure AD tenant GUID. |
| `SHAREPOINT_CLIENT_ID` | Azure AD app (service principal) registration client ID. |
| `SHAREPOINT_CLIENT_SECRET` | Client secret / app password. |

### Source / READ site

| Variable | Default | Description |
|---|---|---|
| `SHAREPOINT_DOMAIN` | *(none)* | e.g. `contoso.sharepoint.com`. |
| `SHAREPOINT_SITE_PATH` | *(none)* | e.g. `/sites/DemoProposalHub`. |
| `SHAREPOINT_LIBRARY_NAME` | `Documents` | Target document library (mapped to a Graph drive). |
| `SHAREPOINT_BASE_FOLDER` | *(empty = library root)* | Folder path relative to the library root. |

### Output / WRITE site

| Variable | Default | Description |
|---|---|---|
| `SHAREPOINT_OUTPUT_DOMAIN` | *(none)* | Output site domain. |
| `SHAREPOINT_OUTPUT_SITE_PATH` | *(none)* | e.g. `/sites/RFP-Outputs`. |
| `SHAREPOINT_OUTPUT_LIBRARY` | *(none)* | Output document library name. |

### Optional tuning (shared by both services)

| Variable | Default | Description |
|---|---|---|
| `SHAREPOINT_SSL_VERIFY` | `false` (service.py) / `true` (config.py) | When `false`, SSL certificate verification is disabled and a custom context with TLS 1.2 minimum is used. |
| `SHAREPOINT_CONNECT_TIMEOUT` | `30` | TCP connect timeout (seconds). |
| `SHAREPOINT_READ_TIMEOUT` | `120` | TCP read timeout (seconds). |
| `SHAREPOINT_MAX_RETRIES` | `5` | urllib3 `Retry.total` (transport-layer retries). |
| `SHAREPOINT_POOL_CONNECTIONS` | `20` | urllib3 `pool_connections`. |
| `SHAREPOINT_POOL_MAXSIZE` | `50` | urllib3 `pool_maxsize`. |

---

## 3. Authentication: OAuth 2.0 Client Credentials Flow

**Both services** (`SharePointService` and `SharePointUploadService`) authenticate
to Microsoft Graph using the **application (app-only) / client-credentials** grant.

### Token request

```
POST https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token
Content-Type: application/x-www-form-urlencoded

client_id={client_id}
&client_secret={client_secret}
&scope=https://graph.microsoft.com/.default
&grant_type=client_credentials
```

| Field | Value |
|---|---|
| **URL** | `https://login.microsoftonline.com/{SHAREPOINT_TENANT_ID}/oauth2/v2.0/token` |
| **Method** | `POST` |
| **Request body** | `client_id`, `client_secret`, `scope=https://graph.microsoft.com/.default`, `grant_type=client_credentials` |
| **Response** | JSON: `{ "token_type": "Bearer", "access_token": "...", "expires_in": 3599, ... }` |
| **Cached in** | `SharePointService.access_token` / `SharePointUploadService._access_token` |

- In `SharePointService` the token call goes through `_make_request_with_retry()`
  which adds `User-Agent`, `Accept`, `Accept-Encoding` headers and applies the
  application's retry strategy.
- In `SharePointUploadService` the token call is a direct `self._session.post()`
  with no application-level retry (transport layer only).
- On `401 Unauthorized` responses from any Graph call, the token is invalidated
  (`access_token = None`) and re-fetched before retrying.

---

## 4. Graph API Endpoint Catalog

Every Microsoft Graph endpoint called by the codebase, grouped by service.

### 4.1 `SharePointService` (READ — `service.py`)

| # | Endpoint | Method | Purpose |
|---|---|---|---|
| 1 | `https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token` | `POST` | Acquire OAuth2 app token. *(auth)* |
| 2 | `https://graph.microsoft.com/v1.0/sites/{domain}:{site_path}?$select=id` | `GET` | Resolve the **site ID** for the source site. |
| 3 | `https://graph.microsoft.com/v1.0/sites/{site_id}/drives` | `GET` | List all drives (document libraries) in the site; finds the one matching `SHAREPOINT_LIBRARY_NAME`. |
| 4 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}` | `GET` | Fetch a single `driveItem`'s metadata — used to refresh the `@microsoft.graph.downloadUrl` (which expires). |
| 5 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children` | `GET` | List direct children (folders + files) of the **drive root**. Used when `base_folder` is empty. |
| 6 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{path}:/children` | `GET` | List direct children of a **subfolder path**. Used by `_get_folders_in_path` (tags) and `_get_all_files_recursive`. |
| 7 | `https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem` | `GET` | Resolve a SharePoint **sharing URL** to a `driveItem`. `share_id` is `u!{base64url-encoded sharing-url}` (the `u!` prefix is required by the Graph sharing API). |
| 8 | *(direct download URL from `@microsoft.graph.downloadUrl`)* | `GET` | Download file **bytes** via the temporary anonymous download URL embedded in the drive item. First attempt. |
| 9 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_id}/content` | `GET` | **Fallback / authoritative** download endpoint. Returns the raw file bytes (redirects to a temporary download URL). |
| 10 | *(paginated `@odata.nextLink)`* | `GET` | Follows OData continuation links when a folder listing or recursive scan returns more items than fit on one page. |

**Query params used on listing calls:**

- `$select=id,name,size,file,folder,webUrl,lastModifiedDateTime,createdDateTime,@microsoft.graph.downloadUrl`
  — used in `_get_modified_files_ordered_recursive_internal` to limit fields.
- `$orderby=lastModifiedDateTime desc` — orders children by most recently modified.
- `$select=id` — used on the site resolution call to minimize payload.

### 4.2 `SharePointUploadService` (WRITE — `upload.py`)

| # | Endpoint | Method | Purpose |
|---|---|---|---|
| 1 | `https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token` | `POST` | Acquire OAuth2 app token. *(auth)* |
| 2 | `https://graph.microsoft.com/v1.0/sites/{domain}:{site_path}?$select=id` | `GET` | Resolve the **site ID** for the output site. |
| 3 | `https://graph.microsoft.com/v1.0/sites/{site_id}/drives` | `GET` | List drives; finds the one matching `SHAREPOINT_OUTPUT_LIBRARY`. |
| 4 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}` | `GET` | Check whether a folder **already exists** (HTTP 200 = exists, 404 = missing). |
| 5 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children` | `POST` | Create a folder **directly under the drive root** (parent_path is empty). Body: `{ "name": folderName, "folder": {}, "@microsoft.graph.conflictBehavior": "fail" }`. |
| 6 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{parent_path}:/children` | `POST` | Create a folder **inside a parent subfolder**. Same body shape. |
| 7 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{folder_path}?$select=webUrl,name` | `GET` | Retrieve the `webUrl` of the job folder (for returning a browseable SharePoint link). |
| 8 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_path}:/content` | `PUT` | **Simple upload** — for files ≤ 4 MB (`SIMPLE_UPLOAD_LIMIT`). The entire file body is sent as `application/octet-stream`. |
| 9 | `https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{sp_path}:/createUploadSession` | `POST` | **Resumable upload** — for files > 4 MB. Returns an `uploadUrl` (a pre-authenticated PUT target). Body: `{ "item": { "@microsoft.graph.conflictBehavior": "replace", "name": fileName } }`. |
| 10 | *(the `uploadUrl` returned from createUploadSession)* | `PUT` | **Chunked upload** — sends 10 MB chunks (`CHUNK_SIZE`) with `Content-Range: bytes {start}-{end}/{file_size}` headers. Repeated until all bytes are sent. |

---

## 5. Request-Response Cycle — Read Path (FastAPI HTTP)

This is the full lifecycle of an HTTP request hitting the FastAPI adapter.

```
[External HTTP client]
        │
        ▼
┌─────────────────────────────────────────────────────────────┐
│  main.py                                                     │
│  - FastAPI app created with lifespan (log start/shutdown)    │
│  - app.include_router(router) — mounts /sharepoint routes   │
│  - @app.get("/") → { service, docs, router }                 │
└─────────────────────────────────────────────────────────────┘
        │  GET /sharepoint/...
        ▼
┌─────────────────────────────────────────────────────────────┐
│  router.py                                                    │
│                                                             │
│  _get_service()  ──► lazy singleton:                        │
│    first request → SharePointService()                      │
│      │                                                      │
│      ▼  (inside SharePointService.__init__)                  │
│      load .env → read env vars                              │
│      build requests.Session with Retry + HTTPAdapter        │
│      (no network yet — token obtained on first API call)    │
│                                                             │
│  Route handlers:                                            │
│    GET /sharepoint/health       → {"status": "ok"}          │
│    GET /sharepoint/files        → discover_tagged_files()   │
│    GET /sharepoint/files/count  → get_file_count()          │
│    GET /sharepoint/files/{path} → get_file_by_relative_path()│
│                                                             │
│  Each handler:                                              │
│    try: service = _get_service()                            │
│         result = await service.<method>(...)                │
│    except: HTTPException(500, detail=str(e))                │
└─────────────────────────────────────────────────────────────┘
        │  await service.<method>
        ▼
┌─────────────────────────────────────────────────────────────┐
│  service.py — SharePointService                               │
│  (the heavy lifting for READ operations)                     │
└─────────────────────────────────────────────────────────────┘
```

### 5.1 Detailed step-by-step: `GET /sharepoint/files` (list all files by tag)

1. **Request received** by FastAPI on `router.py:51` (`list_files`).
2. **`_get_service()`** (router.py:36-41) — if the module-level `_service`
   global is `None`, a new `SharePointService()` is created. This constructor:
   - Reads `.env`.
   - Reads `SHAREPOINT_TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`, `DOMAIN`,
     `SITE_PATH`, `LIBRARY_NAME`, `BASE_FOLDER`.
   - Normalizes `base_folder_path`.
   - Creates a `requests.Session` with a `Retry`/`HTTPAdapter` (pools + retries).
   - Initializes `access_token`, `site_id`, `drive_id` to `None`.
   - No network I/O yet.
3. **`await service.discover_tagged_files(base_folder)`** (service.py:387).
   - **Step A — `_get_site_and_drive_ids()`** (service.py:256):
     - Checks if `site_id` and `drive_id` are cached; if so, returns immediately.
     - Calls **`_get_access_token()`** → `POST /token` (endpoint #1) → stores
       `self.access_token`.
     - Calls **GET `sites/{domain}:{site_path}?$select=id`** (endpoint #2) →
       extracts `site_id` from JSON.
     - Calls **GET `sites/{site_id}/drives`** (endpoint #3) → iterates the
       `value` array to find the drive whose `name` equals `LIBRARY_NAME` →
       stores `self.drive_id`.
   - **Step B — `_get_folders_in_path(base_folder)`** (service.py:434):
     - Builds the **children URL** via `_children_url()` (service.py:419):
       - If folder is empty/root → `…/drives/{drive_id}/root/children`.
       - Otherwise → `…/drives/{drive_id}/root:/{quoted(path)}:/children`.
     - Loops: **GET children** (endpoint #5 or #6) → filters items where
       `'folder' in item` → follows `@odata.nextLink` (endpoint #10) for paging.
   - **Step C — For each tag folder**, calls **`_get_all_files_recursive()`**
     (service.py:457):
     - Same children URL logic.
     - Loops over pages (endpoint #5/6 + #10).
     - For each item:
       - If it's a folder → recurse with `folder_path = parent + "/" + name`.
       - If it's a file → builds a `SharePointFile` dataclass with `name`,
         `path`, `tag`, `size_bytes`, `mime_type`, `download_url`,
         `web_url`, `file_id`, `created`, `modified`.
   - **`return`** `List[SharePointFile]`.
4. **Handler** serializes the list to JSON (FastAPI Pydantic response model)
   and returns HTTP 200.

> **Key detail:** Every Graph call goes through `_make_request_with_retry()`
> which:
> - Injects `User-Agent`, `Accept: application/json`, `Accept-Encoding`.
> - Retries on 429 (respecting `Retry-After`), 500/502/503/504, SSL errors,
>   connection errors, and timeouts (3 attempts, exponential backoff).
> - On 401, clears the token so the next call re-authenticates.

### 5.2 Detailed step-by-step: `GET /sharepoint/files/{relative_path}` (get one file)

1. **Request** hits `router.py:81` (`get_file`).
2. `_get_service()` returns/creates the service.
3. **`await service.get_file_by_relative_path(relative_path)`** (service.py:610):
   - If `relative_path` starts with `http://` or `https://`, it's treated as a
     **sharing URL** → goes to Section 5.3 below.
   - Otherwise:
     - `_get_site_and_drive_ids()` — cached or fresh (Section 5.1, Steps A).
     - `_get_access_token()` — cached or fresh.
     - Builds `full_path = base_folder_path + "/" + relative_path` (if base
       folder is set) or just `relative_path`.
     - Builds the **item URL** via `_item_url()` (service.py:427):
       `"…/drives/{drive_id}/root:/{quoted(path)}"`.
     - **GET item** (endpoint #4) → if 404, returns `None`. Otherwise builds a
       `SharePointFile` via `_file_from_drive_item()`.
4. If result is `None` → `HTTPException(404)`. Otherwise → 200 with JSON.

### 5.3 Detailed step-by-step: sharing URL resolution

1. `get_file_by_relative_path()` detects `https://` prefix → calls
   `get_file_by_sharing_url()` (service.py:582).
2. **`_encode_sharing_url()`** (service.py:577) → base64url-encodes the sharing
   URL and prefixes with `u!`.
3. **`GET https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem`**
   (endpoint #7) with `Authorization: Bearer {token}`.
4. If 404 → returns `None` (sharing URL doesn't resolve). Otherwise maps the
   `driveItem` JSON to a `SharePointFile` via `_file_from_drive_item()`.

### 5.4 Detailed step-by-step: file download (`download_file_to_temp`)

1. Called **programmatically** (not exposed as a router endpoint).
2. Creates a `tempfile.NamedTemporaryFile` with the file's suffix.
3. Calls **`_download_with_retry()`** (service.py:328):
   - **Attempt loop** (default `max_retries=1` → 2 total):
     - **Preferred path** (download_url present):
       - **GET download_url** (endpoint #8 — the `@microsoft.graph.downloadUrl`
         blob URL) → if 401, refresh token + call
         `_refresh_download_url()` (endpoint #4) to get a new URL, then retry.
     - **Fallback path** (download_url empty/unavailable):
       - **GET `…/drives/{drive_id}/items/{file_id}/content`**
         (endpoint #9, with `allow_redirects=True`).
   - Returns `response.content` (raw bytes).
4. Writes bytes to the temp file, closes it, returns the temp file path.

### 5.5 Detailed step-by-step: modified-files query (`get_files_modified_after`)

1. Calls `_get_site_and_drive_ids()` + `_get_access_token()`.
2. `_get_folders_in_path()` lists tag folders (endpoint #5/6).
3. For each tag folder → **`_get_modified_files_ordered_recursive()`**
   (service.py:800):
   - Wraps the internal call in `asyncio.wait_for(..., timeout=300)`.
4. **`_get_modified_files_ordered_recursive_internal()`** (service.py:818):
   - Sets `params = {"$orderby": "lastModifiedDateTime desc", "$select": "..."}`.
   - Loops over pages (endpoint #5/6 with `@odata.nextLink`).
   - For each item:
     - If folder → always recurses (does **not** skip based on folder mtime).
     - If file → parses `lastModifiedDateTime` (ISO 8601, `Z`→`+00:00`) → if
       newer than `after_time` → adds a `SharePointFile` to the result list.
   - **Note:** `params` is set to `None` after the first page because
     `@odata.nextLink` already contains its own query parameters.

---

## 6. Request-Response Cycle — Write Path (Upload)

Triggered programmatically via `SharePointUploadService.upload_artifacts()`.

### 6.1 Step-by-step: `upload_artifacts(user_email, job_id, file_name, zip_bytes)`

0. **Async wrapper**: `upload_artifacts` is an `async` method that runs the
   synchronous HTTP work in a thread via `asyncio.to_thread(_do_upload)`
   (upload.py:358). This keeps the caller's event loop unblocked.

1. **Build folder structure** (upload.py:319-328):
   - `user_folder` = `_email_to_folder_name(user_email)` → e.g.
     `"nithin@accellor.com"` → `"Nithin"`.
   - `job_folder` = `_job_folder_name(job_id, file_name)` →
     `"{short_id}_{cleaned_stem}"` (short_id = first 8 chars of job_id with
     dashes removed; cleaned stem = alphanumeric+underscore, max 50 chars).
   - `folder` = `"RFP Responses/{user_folder}/{job_folder}"`.
   - `sp_path` = `"{folder}/{zip_filename}"`.

2. **`_do_upload()` runs in thread:**
   - **A. `_resolve_drive_sync()`** (upload.py:134):
     - Cached `_drive_id` → return immediately.
     - **`GET sites/{domain}:{site_path}?$select=id`** (endpoint #2) →
       `self._site_id`.
     - **`GET sites/{site_id}/drives`** (endpoint #3) → find drive matching
       `SHAREPOINT_OUTPUT_LIBRARY` → `self._drive_id`.
   - **B. `_ensure_folder_sync(drive_id, folder)`** (upload.py:169) — recursive:
     - **`GET drives/{drive_id}/root:/{folder_path}?$select=webUrl,name`**
       (endpoint #4) → if 200, folder exists, return.
     - If 404, split path into parent + name → recurse to create parent first.
     - **`POST …/root:/{parent}:/children`** or **`POST …/root/children`**
       (endpoint #5/6) with body
       `{ "name": folderName, "folder": {}, "@microsoft.graph.conflictBehavior": "fail" }`.
     - If POST returns 409 (conflict = folder exists), return successfully.
   - **C. `_upload_bytes_sync(drive_id, zip_bytes, sp_path, zip_filename)`**
     (upload.py:233):
     - **If file_size ≤ 4 MB** (`SIMPLE_UPLOAD_LIMIT`):
       - **`PUT drives/{drive_id}/root:/{sp_path}:/content`** (endpoint #8)
         with `Content-Type: application/octet-stream`, body = full file bytes.
       - Returns the Graph `driveItem` JSON.
     - **If file_size > 4 MB** (resumable):
       - **`POST drives/{drive_id}/root:/{sp_path}:/createUploadSession`**
         (endpoint #9) with body
         `{ "item": { "@microsoft.graph.conflictBehavior": "replace", "name": fileName } }`.
       - Extract `uploadUrl` from response JSON.
       - **Loop** sending 10 MB chunks (`CHUNK_SIZE = 10*1024*1024`) via
         `PUT {uploadUrl}` with headers:
         `Content-Length: {len(chunk)}` and
         `Content-Range: bytes {offset}-{end}/{file_size}`.
       - Each chunk upload is confirmed; on `.raise_for_status()` failure the
         whole upload aborts.
     - Returns the final JSON response (driveItem metadata).
   - **D. `_get_folder_web_url_sync(drive_id, folder)`** (upload.py:201):
     - **`GET drives/{drive_id}/root:/{folder_path}?$select=webUrl,name`**
       (endpoint #7) → returns the `webUrl` for the job folder.
   - **E. Fallback**: if no `webUrl`, builds a SharePoint browser URL via
     `_folder_library_url()` — constructs an `AllItems.aspx?id=...` link
     matching what users see in the browser.
   - **Returns** a dict merging the upload result with `webUrl`, `folderWebUrl`,
     `fileWebUrl`, `folderPath`.

3. **`upload_artifacts`** receives the result dict, logs the job/email/web_url,
   and returns it. On any exception, logs the error and returns `None`.

---

## 7. Data Model: `SharePointFile` (service.py:36-48)

A `@dataclass` returned by all read/discovery methods and used as the FastAPI
response model in `router.py`.

| Field | Type | Source (Graph driveItem property) |
|---|---|---|
| `name` | `str` | `item.name` |
| `path` | `str` | Constructed: `folder_path + "/" + name` |
| `tag` | `str` | Derived from the tag folder name (first path segment) |
| `size_bytes` | `int` | `item.size` |
| `mime_type` | `str` | `item.file.mimeType` |
| `download_url` | `str` | `item["@microsoft.graph.downloadUrl"]` (temporary, expires) |
| `web_url` | `str` | `item.webUrl` |
| `file_id` | `str` | `item.id` |
| `created` | `str` | `item.createdDateTime` (ISO 8601) |
| `modified` | `str` | `item.lastModifiedDateTime` (ISO 8601) |

---

## 8. Cross-Cutting Concerns

### 8.1 Retry & Timeout Strategy

| Layer | Mechanism | Config |
|---|---|---|
| **Transport / urllib3** | `HTTPAdapter(max_retries=Retry(total=N, backoff_factor=1, status_forcelist=[429,500,502,503,504], respect_retry_after_header=True))` | `SHAREPOINT_MAX_RETRIES` (5), `SHAREPOINT_POOL_CONNECTIONS` (20), `SHAREPOINT_POOL_MAXSIZE` (50). |
| **Application-level** (`_make_request_with_retry` in service.py) | 3 attempts, exponential backoff (`2**attempt` seconds), explicit handling of 401 (token refresh), 429 (Retry-After), 5xx, SSL errors, connection errors, timeouts. | Hardcoded `max_attempts=3`, `base_delay=2`. |
| **Download-specific** (`_download_with_retry`) | 2 attempts by default, re-fetches download URL on 401. | `max_retries=1` default. |
| **Per-folder timeout** (modified-files scan) | `asyncio.wait_for(..., timeout=300)` per folder. | 5 min. |
| **Session timeout** | `(connect, read)` = `(30s, 120s)` by default. | `SHAREPOINT_CONNECT_TIMEOUT`, `SHAREPOINT_READ_TIMEOUT`. |

### 8.2 HTTP Session

- **`SharePointService`**: Uses `requests.Session()` mounted with
  `HTTPAdapter` for both `http://` and `https://`. The async methods call
  `self.session.request()` **synchronously** (the `async` keyword only
  enables `await asyncio.sleep()` for backoff — the actual HTTP call is
  blocking). This is a known design trait; for true async I/O an
  `httpx.AsyncClient` would be needed.
- **`SharePointUploadService`**: Uses `requests.Session()` with a smaller
  `HTTPAdapter(pool_connections=5, pool_maxsize=10)`.

### 8.3 SSL Configuration

- When `SHAREPOINT_SSL_VERIFY=false`:
  - A custom `ssl.SSLContext` is created with `verify_mode = CERT_NONE`,
    `check_hostname = False`, `OP_NO_SSLv2`, `OP_NO_SSLv3`, and
    `minimum_version = TLSv1_2`.
  - urllib3 `InsecureRequestWarning` is suppressed.
- **Note:** The SSL context is created but **never actually attached to the
  requests session** in `service.py` (the `_create_session` method builds the
  context but does not pass it to `HTTPAdapter` via `ssl_context=`). The
  session's `verify` flag is therefore not set explicitly; `requests` defaults
  to `verify=True`. This means SSL verification is **still active by default**
  even when `SHAREPOINT_SSL_VERIFY=false`. This is a latent bug / inconsistency
  between `service.py` (which has its own env-based settings) and
  `config.py` (which defaults `SHAREPOINT_SSL_VERIFY` to `true`).

### 8.4 Paging

Both `_get_folders_in_path` and `_get_all_files_recursive` (and all variants)
follow `@odata.nextLink` in a `while list_url:` loop, ensuring all items are
fetched across multiple OData pages.

### 8.5 Token Caching & Invalidation

| Event | Action |
|---|---|
| First API call | `_get_access_token()` → POST to token endpoint → cache `self.access_token`. |
| `site_id`/`drive_id` cached | `_get_site_and_drive_ids()` returns cached values on subsequent calls. |
| HTTP 401 from Graph | `self.access_token = None`; the next `_get_access_token()` call re-fetches. |
| Download URL expires | `_refresh_download_url()` re-fetches the single driveItem to get a fresh `@microsoft.graph.downloadUrl`. |

---

## 9. Startup Sequence (server boot to first request)

```
1.  uvicorn main:app  OR  python main.py
2.  main.py: FastAPI(__name__) created → lifespan context manager wired
3.  main.py: app.include_router(router) → /sharepoint/* routes registered
4.  (process idle — no SharePoint connection yet)
5.  ───── first HTTP request arrives ─────
6.  router.py: _get_service() → SharePointService.__init__()
       ├─ load_dotenv(".env")
       ├─ read env vars (tenant/client/site/library/base_folder)
       ├─ create requests.Session + Retry + HTTPAdapter
       └─ log config (domain, base folder, SSL, timeouts)
7.  router.py: handler calls await service.<method>(...)
8.  service.py: _get_site_and_drive_ids()
       ├─ _get_access_token() → POST /token (Graph auth endpoint)
       ├─ GET sites/{domain}:{path} → site_id
       └─ GET sites/{site_id}/drives → drive_id (matched by name)
9.  service.py: business logic method (discover / get / count / download)
10. router.py: HTTP 200/404/500 response returned to client
```

---

## 10. Programmatic Usage Paths

### 10.1 Read (programmatic)

```
SharePointService()
  → discover_tagged_files()      → GET token, GET site, GET drives, GET children×N, recurse
  → find_file_by_name(name)      → calls discover_tagged_files(), filters by name
  → find_file_by_relative_path() → calls discover_tagged_files(), filters by path
  → get_file_by_relative_path()  → GET item (single), OR GET shares/{id}/driveItem
  → get_file_by_sharing_url()    → GET shares/{id}/driveItem
  → get_file_content(file)       → _download_with_retry() → GET downloadUrl or GET …/content
  → download_file_to_temp(file)  → get_file_content() → write to tempfile → return path
  → get_file_count()             → GET children×N, count (no downloads)
  → get_files_modified_after()   → GET children×N ordered, filter by mtime
```

### 10.2 Write (programmatic)

```
SharePointUploadService()
  → upload_artifacts(email, job_id, file_name, zip_bytes)
     → asyncio.to_thread(_do_upload)
        → _resolve_drive_sync()    → POST /token, GET site, GET drives
        → _ensure_folder_sync()   → GET folder (check), POST folder (create, recurse)
        → _upload_bytes_sync()    → PUT /content  (≤4MB)  OR  POST createUploadSession + PUT chunks (>4MB)
        → _get_folder_web_url_sync() → GET folder webUrl
        → return { **upload_result, webUrl, folderWebUrl, fileWebUrl, folderPath }
```

---

## 11. Error Handling Summary

| Situation | Where handled | Behavior |
|---|---|---|
| Missing credentials | `SharePointUploadService.__init__` | `raise RuntimeError(...)` immediately at construction. |
| Graph 401 Unauthorized | `_make_request_with_retry` / `_download_with_retry` / `_request` | Token cleared (`= None`); retried with fresh token. |
| Graph 404 (file/folder/sharing URL) | `get_file_by_relative_path`, `get_file_by_sharing_url` | Returns `None` → router raises `HTTPException(404)`. |
| Graph 409 Conflict (folder create) | `_ensure_folder_sync` | Silently treated as "folder already exists" → returns. |
| Graph 429 (rate limit) | `_make_request_with_retry` | Sleeps `Retry-After` header value, then retries. |
| Graph 5xx transient | `_make_request_with_retry` | Exponential backoff (`2**attempt`), 3 attempts. |
| Transport SSL / Connection / Timeout | `_make_request_with_retry` | Exponential backoff, 3 attempts, then re-raise. |
| Upload to folder with no `webUrl` | `upload_artifacts._do_upload` | Falls back to `_folder_library_url()` (synthesized browser URL). |
| `upload_artifacts` with empty `user_email` | `upload_artifacts` | Returns `None` immediately, logs a warning. |
| Any uncaught exception in a router handler | `router.py` (each endpoint) | `logger.error(...)` + `HTTPException(500, detail=str(e))`. |
| `get_files_modified_after` top-level failure | `service.py:794` | **Catches and returns `[]`** — does not raise. Logs the error with elapsed time. |

---

## 12. Quick Reference: All Graph API Calls (Flat List)

| Service | Call | Endpoint (method) | Method file:line |
|---|---|---|---|
| Both | Acquire token | `POST …/oauth2/v2.0/token` | service.py:233 / upload.py:105 |
| Read | Resolve site | `GET v1.0/sites/{domain}:{path}?$select=id` | service.py:269 |
| Read | List drives | `GET v1.0/sites/{site_id}/drives` | service.py:275 |
| Read | Get item metadata | `GET v1.0/drives/{drive_id}/items/{item_id}` | service.py:309 |
| Read | List children (root) | `GET v1.0/drives/{drive_id}/root/children` | service.py:424 |
| Read | List children (path) | `GET v1.0/drives/{drive_id}/root:/{path}:/children` | service.py:425 |
| Read | Resolve sharing URL | `GET v1.0/shares/{u!base64url}/driveItem` | service.py:591 |
| Read | Download (temporary URL) | `GET {@microsoft.graph.downloadUrl}` | service.py:339 |
| Read | Download (content endpoint) | `GET v1.0/drives/{drive_id}/items/{item_id}/content` | service.py:359 |
| Read | Follow paging | `GET {@odata.nextLink}` | service.py:450, 497, 735, 878 |
| Read | Ordered+filtered listing | `GET …/children?$orderby=…&$select=…` | service.py:834 |
| Write | Resolve site | `GET v1.0/sites/{domain}:{path}?$select=id` | upload.py:141 |
| Write | List drives | `GET v1.0/sites/{site_id}/drives` | upload.py:148 |
| Write | Check folder exists | `GET v1.0/drives/{drive_id}/root:/{path}` | upload.py:172 |
| Write | Create folder (root) | `POST v1.0/drives/{drive_id}/root/children` | upload.py:187 |
| Write | Create folder (parent) | `POST v1.0/drives/{drive_id}/root:/{parent}:/children` | upload.py:184 |
| Write | Get folder webUrl | `GET v1.0/drives/{drive_id}/root:/{path}?$select=webUrl,name` | upload.py:205 |
| Write | Simple upload | `PUT v1.0/drives/{drive_id}/root:/{path}:/content` | upload.py:241 |
| Write | Create upload session | `POST v1.0/drives/{drive_id}/root:/{path}:/createUploadSession` | upload.py:257 |
| Write | Chunked upload | `PUT {uploadUrl}` w/ `Content-Range` | upload.py:285 |
| Verify | All of the above (subset) | `verify_connection.py` smoke test | verify_connection.py:92-155 |
