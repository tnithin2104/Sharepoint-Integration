"""
Upload artifacts to the shared SmartRFP-Outputs SharePoint team site
via Microsoft Graph API.

Uses Application-level permissions (Sites.ReadWrite.All) with client_credentials.
Files are organized by user name under a shared document library so the
whole team can access all generated artifacts.

Target structure:
  SmartRFP-Outputs / Documents / RFP Responses /
    {User Name} / {short_job_id}_{doc_name_truncated} / artifacts.zip
"""
from __future__ import annotations

import asyncio
import logging
import os
import re
from pathlib import Path, PurePosixPath

import requests
from dotenv import load_dotenv
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

load_dotenv(Path(__file__).resolve().parents[1] / ".env")

logger = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
ROOT_FOLDER = "RFP Responses"
CHUNK_SIZE = 10 * 1024 * 1024  # 10 MB
SIMPLE_UPLOAD_LIMIT = 4 * 1024 * 1024  # 4 MB
MAX_DOC_NAME_LEN = 50


def _email_to_folder_name(email: str) -> str:
    """Convert 'sravan.gangi@accellor.com' -> 'Sravan Gangi'."""
    local_part = email.split("@")[0]
    parts = local_part.replace("_", ".").split(".")
    return " ".join(p.capitalize() for p in parts)


def _job_folder_name(job_id: str, doc_name: str) -> str:
    """
    Build a readable job folder name.
    'a1b2c3d4_Enterprise_Cloud_Migration_RFP_Response_2026_Acme'
    """
    short_id = job_id.replace("-", "")[:8]
    stem = PurePosixPath(doc_name).stem
    clean = re.sub(r"[^\w]+", "_", stem).strip("_")
    while "__" in clean:
        clean = clean.replace("__", "_")
    if len(clean) > MAX_DOC_NAME_LEN:
        clean = clean[:MAX_DOC_NAME_LEN].rstrip("_")
    return f"{short_id}_{clean}"


class SharePointUploadService:

    def __init__(self):
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET")
        self.sp_domain = os.getenv(
            "SHAREPOINT_OUTPUT_DOMAIN",
        )
        self.sp_site_path = os.getenv(
            "SHAREPOINT_OUTPUT_SITE_PATH",
        )
        self.sp_library_name = os.getenv(
            "SHAREPOINT_OUTPUT_LIBRARY",
        )

        self._access_token: str | None = None
        self._drive_id: str | None = None
        self._site_id: str | None = None
        self._session = self._create_session()

        if not all([self.tenant_id, self.client_id, self.client_secret]):
            raise RuntimeError(
                "SharePointUploadService: missing SHAREPOINT_TENANT_ID / "
                "SHAREPOINT_CLIENT_ID / SHAREPOINT_CLIENT_SECRET"
            )
        logger.info(
            "SharePointUploadService initialized (site=%s%s)",
            self.sp_domain, self.sp_site_path,
        )

    @staticmethod
    def _create_session() -> requests.Session:
        session = requests.Session()
        retry = Retry(
            total=3,
            backoff_factor=1,
            status_forcelist=[429, 500, 502, 503, 504],
            allowed_methods=["GET", "POST", "PUT"],
            respect_retry_after_header=True,
            raise_on_status=False,
        )
        adapter = HTTPAdapter(max_retries=retry, pool_connections=5, pool_maxsize=10)
        session.mount("https://", adapter)
        return session

    def _get_token_sync(self) -> str:
        if self._access_token:
            return self._access_token

        url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        resp = self._session.post(url, data={
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        })
        resp.raise_for_status()
        self._access_token = resp.json()["access_token"]
        logger.info("SharePointUpload: acquired app token")
        return self._access_token

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._get_token_sync()}"}

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        resp = self._session.request(method, url, **kwargs)
        if resp.status_code == 401:
            self._access_token = None
            kwargs.setdefault("headers", {})["Authorization"] = (
                f"Bearer {self._get_token_sync()}"
            )
            resp = self._session.request(method, url, **kwargs)
        return resp

    def _resolve_drive_sync(self) -> str:
        """Resolve and cache the drive ID for the output site's document library."""
        if self._drive_id:
            return self._drive_id

        headers = {**self._headers(), "Content-Type": "application/json"}

        site_url = (
            f"{GRAPH_BASE}/sites/{self.sp_domain}:{self.sp_site_path}?$select=id"
        )
        resp = self._request("GET", site_url, headers=headers)
        resp.raise_for_status()
        self._site_id = resp.json()["id"]

        drives_url = f"{GRAPH_BASE}/sites/{self._site_id}/drives"
        resp = self._request("GET", drives_url, headers=headers)
        resp.raise_for_status()

        for drive in resp.json().get("value", []):
            if drive["name"] == self.sp_library_name:
                self._drive_id = drive["id"]
                break

        if not self._drive_id:
            raise RuntimeError(
                f"Drive '{self.sp_library_name}' not found in site "
                f"{self.sp_domain}{self.sp_site_path}"
            )

        logger.info(
            "SharePointUpload: resolved drive %s (site=%s)",
            self._drive_id, self._site_id,
        )
        return self._drive_id

    def _ensure_folder_sync(self, drive_id: str, folder_path: str) -> None:
        headers = {**self._headers(), "Content-Type": "application/json"}

        check_url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
        resp = self._request("GET", check_url, headers=headers)
        if resp.status_code == 200:
            return

        parts = folder_path.replace("\\", "/").split("/")
        parent_path = "/".join(parts[:-1])
        folder_name = parts[-1]

        if parent_path:
            self._ensure_folder_sync(drive_id, parent_path)
            create_url = (
                f"{GRAPH_BASE}/drives/{drive_id}/root:/{parent_path}:/children"
            )
        else:
            create_url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"

        body = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }
        resp = self._request("POST", create_url, headers=headers, json=body)

        if resp.status_code == 409:
            return
        resp.raise_for_status()
        logger.info("SharePointUpload: created folder %s", folder_path)

    def _get_folder_web_url_sync(self, drive_id: str, folder_path: str) -> str:
        """Return a browseable SharePoint folder URL (not the uploaded file)."""
        headers = self._headers()
        check_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
            f"?$select=webUrl,name"
        )
        resp = self._request("GET", check_url, headers=headers)
        resp.raise_for_status()
        return (resp.json().get("webUrl") or "").strip()

    def _folder_library_url(self, folder_path: str) -> str:
        """
        Build the SharePoint library folder view URL (AllItems.aspx).
        Matches the browser path users expect under Shared Documents.
        """
        from urllib.parse import quote

        # Graph drive "Documents" is exposed as "Shared Documents" in site URLs.
        library_segment = (
            "Shared Documents"
            if self.sp_library_name.lower() in ("documents", "shared documents")
            else self.sp_library_name
        )
        site_path = self.sp_site_path if self.sp_site_path.startswith("/") else f"/{self.sp_site_path}"
        folder_id = f"{site_path}/{library_segment}/{folder_path}".replace("//", "/")
        encoded_id = quote(folder_id, safe="")
        return (
            f"https://{self.sp_domain}{site_path}/{quote(library_segment)}"
            f"/Forms/AllItems.aspx?id={encoded_id}"
        )

    def _upload_bytes_sync(
        self, drive_id: str, file_bytes: bytes, sp_path: str, file_name: str,
    ) -> dict:
        headers = self._headers()
        file_size = len(file_bytes)

        if file_size <= SIMPLE_UPLOAD_LIMIT:
            upload_url = (
                f"{GRAPH_BASE}/drives/{drive_id}/root:/{sp_path}:/content"
            )
            resp = self._request(
                "PUT", upload_url,
                headers={**headers, "Content-Type": "application/octet-stream"},
                data=file_bytes,
            )
            resp.raise_for_status()
            logger.info(
                "SharePointUpload: uploaded %s (simple PUT, %d KB)",
                file_name, file_size // 1024,
            )
            return resp.json()

        # Resumable upload session for large files
        session_url = (
            f"{GRAPH_BASE}/drives/{drive_id}/root:/{sp_path}:/createUploadSession"
        )
        resp = self._request(
            "POST", session_url,
            headers={**headers, "Content-Type": "application/json"},
            json={
                "item": {
                    "@microsoft.graph.conflictBehavior": "replace",
                    "name": file_name,
                }
            },
        )
        resp.raise_for_status()
        upload_url = resp.json()["uploadUrl"]
        logger.info(
            "SharePointUpload: upload session created for %s (%d MB)",
            file_name, file_size // (1024 * 1024),
        )

        offset = 0
        last_resp = None
        while offset < file_size:
            end = min(offset + CHUNK_SIZE, file_size) - 1
            chunk = file_bytes[offset : end + 1]
            chunk_headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{file_size}",
            }
            last_resp = self._session.put(
                upload_url, headers=chunk_headers, data=chunk,
            )
            last_resp.raise_for_status()
            pct = int((end + 1) / file_size * 100)
            logger.info(
                "SharePointUpload: %s chunk %d%% (%d/%d bytes)",
                file_name, pct, end + 1, file_size,
            )
            offset = end + 1

        logger.info(
            "SharePointUpload: upload complete - %s (%d MB)",
            file_name, file_size // (1024 * 1024),
        )
        return last_resp.json() if last_resp else {}

    async def upload_artifacts(
        self,
        user_email: str,
        job_id: str,
        file_name: str,
        zip_bytes: bytes,
        zip_filename: str = "artifacts.zip",
    ) -> dict | None:
        """
        Upload the artifacts ZIP to the shared SharePoint team site.
        Non-blocking: runs the sync HTTP calls in a thread pool.
        Returns the Graph API response dict, or None on failure.

        Target: SmartRFP-Outputs / Documents /
                  RFP Responses / {User Name} /
                  {short_id}_{doc_name} / artifacts.zip
        """
        if not user_email:
            logger.warning(
                "SharePoint upload skipped: no user_email for job %s", job_id,
            )
            return None

        user_folder = _email_to_folder_name(user_email)
        job_folder = _job_folder_name(job_id, file_name)
        folder = f"{ROOT_FOLDER}/{user_folder}/{job_folder}"
        sp_path = f"{folder}/{zip_filename}"

        def _do_upload():
            drive_id = self._resolve_drive_sync()
            self._ensure_folder_sync(drive_id, folder)
            file_result = self._upload_bytes_sync(
                drive_id, zip_bytes, sp_path, zip_filename,
            )
            # Prefer the job-folder browse URL so "View in SharePoint" opens
            # the folder view instead of downloading artifacts.zip.
            folder_web_url = ""
            try:
                folder_web_url = self._get_folder_web_url_sync(drive_id, folder)
            except Exception as e:
                logger.warning(
                    "SharePointUpload: could not resolve folder webUrl for %s: %s",
                    folder, e,
                )
            if not folder_web_url:
                folder_web_url = self._folder_library_url(folder)

            return {
                **(file_result or {}),
                "webUrl": folder_web_url,
                "folderWebUrl": folder_web_url,
                "fileWebUrl": (file_result or {}).get("webUrl", ""),
                "folderPath": folder,
            }

        try:
            result = await asyncio.to_thread(_do_upload)
            web_url = result.get("webUrl", "")
            logger.info(
                "SharePointUpload: artifacts uploaded for job %s -> %s (user: %s)",
                job_id, web_url, user_email,
            )
            return result
        except Exception as e:
            logger.error(
                "SharePoint upload failed for job %s (user: %s): %s",
                job_id, user_email, e,
            )
            return None


def _create_global_instance():
    try:
        return SharePointUploadService()
    except Exception as e:
        logger.warning("SharePointUploadService could not be initialized: %s", e)
        return None


sharepoint_upload_service: SharePointUploadService | None = _create_global_instance()
