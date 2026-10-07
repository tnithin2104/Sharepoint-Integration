import os
import requests
import logging
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import tempfile
import uuid
from dataclasses import dataclass
from urllib.parse import quote, unquote, urlparse
import base64
import re
import time
from datetime import datetime, timezone
import urllib3
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import ssl
import asyncio
from dotenv import load_dotenv

load_dotenv(Path(__file__).resolve().parents[1] / ".env")

logger = logging.getLogger(__name__)
 
# Disable SSL warnings for development
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
 
# Environment variable configuration
SHAREPOINT_SSL_VERIFY = os.getenv("SHAREPOINT_SSL_VERIFY", "false").lower() == "true"
SHAREPOINT_CONNECT_TIMEOUT = int(os.getenv("SHAREPOINT_CONNECT_TIMEOUT", "30"))
SHAREPOINT_READ_TIMEOUT = int(os.getenv("SHAREPOINT_READ_TIMEOUT", "120"))
SHAREPOINT_MAX_RETRIES = int(os.getenv("SHAREPOINT_MAX_RETRIES", "5"))
SHAREPOINT_POOL_CONNECTIONS = int(os.getenv("SHAREPOINT_POOL_CONNECTIONS", "20"))
SHAREPOINT_POOL_MAXSIZE = int(os.getenv("SHAREPOINT_POOL_MAXSIZE", "50"))
 
@dataclass
class SharePointFile:
    """Represents a file from SharePoint with metadata"""
    name: str
    path: str
    tag: str
    size_bytes: int
    mime_type: str
    download_url: str
    web_url: str
    file_id: str
    created: str
    modified: str
 
class SharePointService:
    """Service for interacting with SharePoint via Microsoft Graph API"""
   
    def __init__(self):
        self.tenant_id = os.getenv("SHAREPOINT_TENANT_ID")
        self.client_id = os.getenv("SHAREPOINT_CLIENT_ID")
        self.client_secret = os.getenv("SHAREPOINT_CLIENT_SECRET")
        self.domain = os.getenv("SHAREPOINT_DOMAIN")
        self.site_path = os.getenv("SHAREPOINT_SITE_PATH")
        self.library_name = os.getenv("SHAREPOINT_LIBRARY_NAME", "Documents")
        self.base_folder_path = self._normalize_folder_path(
            os.getenv("SHAREPOINT_BASE_FOLDER", "")
        )
       
        self.access_token = None
        self.site_id = None
        self.drive_id = None
        
        self.session = self._create_session()
        
        logger.info(f"SharePoint service initialized for {self.domain}{self.site_path}")
        logger.info(
            f"Base folder (relative to document library): '{self.base_folder_path or '(library root)'}'"
        )
        logger.info(f"SSL verify: {SHAREPOINT_SSL_VERIFY}, Timeout: {SHAREPOINT_CONNECT_TIMEOUT}s/{SHAREPOINT_READ_TIMEOUT}s")

    @staticmethod
    def _normalize_folder_path(folder_path: Optional[str]) -> str:
        """
        Normalize a SharePoint folder path for Graph API drive paths.
        Accepts empty string (library root), relative paths like '01/subfolder',
        or full SharePoint URLs (extracts path after 'Shared Documents').
        """
        if not folder_path:
            return ""

        path = folder_path.strip().strip('"').strip("'")
        if not path:
            return ""

        if path.lower().startswith(("http://", "https://")):
            parsed = urlparse(path)
            path = unquote(parsed.path or "")

        # Strip site prefix when a full server-relative path is pasted
        site_path = os.getenv("SHAREPOINT_SITE_PATH", "").rstrip("/")
        if site_path and path.lower().startswith(site_path.lower()):
            path = path[len(site_path) :]

        # Drive already targets the document library — strip library segment if present
        for marker in ("/Shared Documents", "/Shared%20Documents", "Shared Documents"):
            if marker.lower() in path.lower():
                parts = re.split(r"/Shared Documents/?", path, flags=re.IGNORECASE)
                path = parts[-1] if parts else ""
                break

        return path.strip().strip("/")
       
    def _create_session(self) -> requests.Session:
        """Create a requests session with retry strategy and SSL configuration"""
        session = requests.Session()
        
        # Configure retry strategy
        retry_strategy = Retry(
            total=SHAREPOINT_MAX_RETRIES,  # Total number of retries
            backoff_factor=1,  # Wait 1, 2, 4, 8, 16 seconds between retries
            status_forcelist=[429, 500, 502, 503, 504],  # HTTP status codes to retry on
            allowed_methods=["HEAD", "GET", "POST", "PUT", "DELETE", "OPTIONS", "TRACE"],
            respect_retry_after_header=True,
            raise_on_status=False
        )
        
        # Create adapter with retry strategy and connection pooling
        adapter = HTTPAdapter(
            max_retries=retry_strategy, 
            pool_connections=SHAREPOINT_POOL_CONNECTIONS,  # Increased connection pool
            pool_maxsize=SHAREPOINT_POOL_MAXSIZE,     # Increased max connections
            pool_block=False     # Don't block when pool is full
        )
        
        # Mount adapter for both HTTP and HTTPS
        session.mount("http://", adapter)
        session.mount("https://", adapter)
        
        # Configure SSL context with more lenient settings
        if not SHAREPOINT_SSL_VERIFY:
            ssl_context = ssl.create_default_context()
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE
            ssl_context.options |= ssl.OP_NO_SSLv2
            ssl_context.options |= ssl.OP_NO_SSLv3
            ssl_context.minimum_version = ssl.TLSVersion.TLSv1_2
        
        # Set session timeout and SSL configuration
        session.timeout = (SHAREPOINT_CONNECT_TIMEOUT, SHAREPOINT_READ_TIMEOUT)  # Increased read timeout to 2 minutes
        
        return session
    
    def _make_request_with_retry(self, method: str, url: str, **kwargs) -> requests.Response:
        """Make HTTP request with comprehensive error handling and retry logic"""
        max_attempts = 3
        base_delay = 2
        
        for attempt in range(max_attempts):
            try:
                # Ensure proper headers
                headers = kwargs.get('headers', {})
                headers.update({
                    'User-Agent': 'Smart-RFP-SharePoint-Service/1.0',
                    'Accept': 'application/json',
                    'Accept-Encoding': 'gzip, deflate'
                })
                kwargs['headers'] = headers
                
                # Make request
                response = self.session.request(method, url, **kwargs)
                
                # Handle specific error cases
                if response.status_code == 401:
                    logger.warning(f"Authentication failed on attempt {attempt + 1}, refreshing token...")
                    self.access_token = None
                    # Don't raise here, let the caller handle token refresh
                    return response
                
                elif response.status_code in [500, 502, 503, 504]:
                    logger.warning(f"Server error {response.status_code} on attempt {attempt + 1}")
                    if attempt < max_attempts - 1:
                        delay = base_delay ** attempt
                        logger.info(f"Waiting {delay} seconds before retry...")
                        time.sleep(delay)
                        continue
                
                elif response.status_code == 429:
                    retry_after = response.headers.get('Retry-After', base_delay)
                    logger.warning(f"Rate limited, waiting {retry_after} seconds...")
                    time.sleep(int(retry_after))
                    continue
                
                # Success or other status codes
                return response
                
            except requests.exceptions.SSLError as e:
                logger.error(f"SSL error on attempt {attempt + 1}: {e}")
                if attempt < max_attempts - 1:
                    delay = base_delay ** attempt
                    logger.info(f"Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    continue
                else:
                    raise
                    
            except requests.exceptions.ConnectionError as e:
                logger.error(f"Connection error on attempt {attempt + 1}: {e}")
                if attempt < max_attempts - 1:
                    delay = base_delay ** attempt
                    logger.info(f"Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    continue
                else:
                    raise
                    
            except requests.exceptions.Timeout as e:
                logger.error(f"Timeout error on attempt {attempt + 1}: {e}")
                if attempt < max_attempts - 1:
                    delay = base_delay ** attempt
                    logger.info(f"Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    continue
                else:
                    raise
                    
            except Exception as e:
                logger.error(f"Unexpected error on attempt {attempt + 1}: {e}")
                if attempt < max_attempts - 1:
                    delay = base_delay ** attempt
                    logger.info(f"Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    continue
                else:
                    raise
        
        # If we get here, all attempts failed
        raise Exception(f"All {max_attempts} attempts failed for {method} {url}")
    
    async def _refresh_token_if_needed(self):
        """Refresh access token if it's None or expired"""
        if not self.access_token:
            await self._get_access_token()
       
    async def _get_access_token(self) -> str:
        """Get OAuth2 access token for Microsoft Graph API"""
        if self.access_token:
            return self.access_token
           
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        token_data = {
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default',
            'grant_type': 'client_credentials'
        }
       
        try:
            response = self._make_request_with_retry('POST', token_url, data=token_data)
            response.raise_for_status()
            self.access_token = response.json().get("access_token")
            logger.info("SharePoint access token retrieved successfully")
            return self.access_token
        except Exception as e:
            logger.error(f"Failed to get SharePoint access token: {e}")
            raise
 
    async def _get_site_and_drive_ids(self) -> Tuple[str, str]:
        """Get SharePoint site ID and drive ID"""
        if self.site_id and self.drive_id:
            return self.site_id, self.drive_id
           
        token = await self._get_access_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
       
        try:
            # Get site ID
            site_info_url = f"https://graph.microsoft.com/v1.0/sites/{self.domain}:{self.site_path}?$select=id"
            site_response = self._make_request_with_retry('GET', site_info_url, headers=headers)
            site_response.raise_for_status()
            self.site_id = site_response.json().get("id")
           
            # Get drive ID for the target library
            drives_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
            drives_response = self._make_request_with_retry('GET', drives_url, headers=headers)
            drives_response.raise_for_status()
            drives = drives_response.json().get("value", [])
           
            for drive in drives:
                if drive.get("name") == self.library_name:
                    self.drive_id = drive.get("id")
                    break
                   
            if not self.drive_id:
                raise ValueError(f"Drive named '{self.library_name}' not found in site")
               
            logger.info(f"SharePoint site ID: {self.site_id}, drive ID: {self.drive_id}")
            return self.site_id, self.drive_id
            
        except Exception as e:
            logger.error(f"Error getting site and drive IDs: {e}")
            raise
 
    async def _refresh_download_url(self, sharepoint_file: SharePointFile) -> str:
        """
        Refresh the download URL for a SharePoint file by making a new API call.
        Returns the new download URL.
        """
        try:
            await self._get_site_and_drive_ids()
            token = await self._get_access_token()
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
           
            # Get fresh file metadata including new download URL
            file_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{sharepoint_file.file_id}"
            response = self._make_request_with_retry('GET', file_url, headers=headers)
            response.raise_for_status()
           
            item = response.json()
            new_download_url = item.get('@microsoft.graph.downloadUrl', '')
           
            if not new_download_url:
                raise ValueError(f"No download URL available for file {sharepoint_file.name}")
               
            # Update the file object with new download URL
            sharepoint_file.download_url = new_download_url
            logger.info(f"Refreshed download URL for file '{sharepoint_file.name}'")
            return new_download_url
           
        except Exception as e:
            logger.error(f"Error refreshing download URL for file '{sharepoint_file.name}': {e}")
            raise
 
    async def _download_with_retry(self, sharepoint_file: SharePointFile, max_retries: int = 1) -> bytes:
        """
        Download file content with retry logic for expired URLs.
        Returns file content as bytes.
        """
        last_exception = None
       
        for attempt in range(max_retries + 1):  # +1 to include initial attempt
            try:
                # Preferred path: if we have a direct download URL, use it
                if sharepoint_file.download_url:
                    response = self._make_request_with_retry('GET', sharepoint_file.download_url)
                    if response.status_code == 401:
                        logger.warning(f"Download URL expired for '{sharepoint_file.name}', attempt {attempt + 1}/{max_retries}, refreshing token...")
                        self.access_token = None
                        await self._get_access_token()
                        # Refresh the download URL
                        await self._refresh_download_url(sharepoint_file)
                        # Retry with new URL
                        response = self._make_request_with_retry('GET', sharepoint_file.download_url)
                    response.raise_for_status()
                    logger.info(
                        "Successfully downloaded file '%s' (%d bytes) via direct URL on attempt %d",
                        sharepoint_file.name, len(response.content), attempt + 1,
                    )
                    return response.content

                # Fallback path: use Graph /content endpoint (authoritative)
                await self._get_site_and_drive_ids()
                token = await self._get_access_token()
                headers = { 'Authorization': f'Bearer {token}' }
                content_url = f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/items/{sharepoint_file.file_id}/content"
                response = self._make_request_with_retry('GET', content_url, headers=headers, allow_redirects=True)
                if response.status_code == 401:
                    logger.warning(f"/content request unauthorized for '{sharepoint_file.name}', attempt {attempt + 1}/{max_retries}, refreshing token...")
                    self.access_token = None
                    await self._get_access_token()
                    response = self._make_request_with_retry('GET', content_url, headers={ 'Authorization': f'Bearer {self.access_token}' }, allow_redirects=True)
                response.raise_for_status()
                logger.info(
                    "Successfully downloaded file '%s' (%d bytes) via /content on attempt %d",
                    sharepoint_file.name, len(response.content), attempt + 1,
                )
                return response.content

            except requests.exceptions.RequestException as e:
                last_exception = e
                logger.warning(f"Download attempt {attempt + 1} failed for '{sharepoint_file.name}': {e}")
                if attempt < max_retries:
                    wait_time = 2 ** attempt
                    logger.info(f"Waiting {wait_time} seconds before retry...")
                    time.sleep(wait_time)
                else:
                    logger.error(f"Final download attempt failed for '{sharepoint_file.name}': {e}")
       
        # All retries failed
        logger.error(f"All download attempts failed for '{sharepoint_file.name}'")
        raise last_exception or Exception(f"Failed to download file '{sharepoint_file.name}' after {max_retries} attempts")
 
    async def discover_tagged_files(self, base_folder: str = None) -> List[SharePointFile]:
        """
        Discover all files in SharePoint organized by tags (folder names under base folder).
        Returns list of SharePointFile objects with tag assignments.
        """
        await self._get_site_and_drive_ids()
        token = await self._get_access_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        discovered_files = []
        base_folder = self._normalize_folder_path(base_folder or self.base_folder_path)
        try:
            # Get tag folders (direct subfolders under base folder)
            tag_folders = await self._get_folders_in_path(base_folder, headers)
            logger.info(f"Found {len(tag_folders)} tag folders under {base_folder}")
            
            for tag_folder in tag_folders:
                tag_name = tag_folder['name']
                tag_folder_path = f"{base_folder}/{tag_name}"
                # Recursively get all files in this tag folder and its subfolders
                files_in_tag = await self._get_all_files_recursive(tag_folder_path, headers, tag_name)
                logger.info(f"Tag '{tag_name}': Found {len(files_in_tag)} files")
                discovered_files.extend(files_in_tag)
            
            logger.info(f"Discovered {len(discovered_files)} total files across all tags in SharePoint")
            return discovered_files
        except Exception as e:
            logger.error(f"Error discovering files from SharePoint: {e}")
            raise
 
    def _children_url(self, folder_path: str) -> str:
        """Build the correct Graph API children URL for a folder path.
        Uses /root/children for root, /root:/{path}:/children for subfolders."""
        path = folder_path.strip().strip("/")
        if not path:
            return f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root/children"
        return f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{quote(path)}:/children"

    def _item_url(self, item_path: str) -> str:
        """Build the correct Graph API item URL for a file/folder path."""
        path = item_path.strip().strip("/")
        if not path:
            return f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root"
        return f"https://graph.microsoft.com/v1.0/drives/{self.drive_id}/root:/{quote(path)}"

    async def _get_folders_in_path(self, folder_path: str, headers: Dict) -> List[Dict]:
        """Get direct subfolders in a given path"""
        list_url = self._children_url(folder_path)
        
        all_folders = []
        while list_url:
            try:
                response = self._make_request_with_retry('GET', list_url, headers=headers)
                response.raise_for_status()
                response_data = response.json()
                items = response_data.get("value", [])
                # Filter only folders
                folders = [item for item in items if 'folder' in item]
                all_folders.extend(folders)
                
                # Check for next page
                list_url = response_data.get("@odata.nextLink")
            except Exception as e:
                logger.error(f"Error getting folders from {folder_path}: {e}")
                break
        
        return all_folders

    async def _get_all_files_recursive(self, folder_path: str, headers: Dict, tag: str) -> List[SharePointFile]:
        """Recursively get all files in a folder and its subfolders"""
        files = []
        page_count = 0
        try:
            list_url = self._children_url(folder_path)
            
            while list_url:
                page_count += 1
                try:
                    response = self._make_request_with_retry('GET', list_url, headers=headers)
                    response.raise_for_status()
                    response_data = response.json()
                    items = response_data.get("value", [])
                    
                    logger.debug(f"Page {page_count}: Found {len(items)} items in {folder_path}")
                    
                    for item in items:
                        if 'folder' in item:
                            # It's a subfolder, recurse into it
                            subfolder_path = f"{folder_path}/{item['name']}"
                            subfiles = await self._get_all_files_recursive(subfolder_path, headers, tag)
                            files.extend(subfiles)
                        else:
                            # It's a file, add to list
                            sharepoint_file = SharePointFile(
                                name=item.get('name', ''),
                                path=f"{folder_path}/{item.get('name', '')}",
                                tag=tag,
                                size_bytes=item.get('size', 0),
                                mime_type=item.get('file', {}).get('mimeType', ''),
                                download_url=item.get('@microsoft.graph.downloadUrl', ''),
                                web_url=item.get('webUrl', ''),
                                file_id=item.get('id', ''),
                                created=item.get('createdDateTime', ''),
                                modified=item.get('lastModifiedDateTime', '')
                            )
                            files.append(sharepoint_file)
                    
                    # Check for next page
                    list_url = response_data.get("@odata.nextLink")
                    if list_url:
                        logger.debug(f"Following next page link for {folder_path}")
                        
                except Exception as e:
                    logger.error(f"Error on page {page_count} for {folder_path}: {e}")
                    break
                
        except Exception as e:
            logger.error(f"Error getting files from {folder_path}: {e}")
        
        if page_count > 1:
            logger.info(f"Retrieved {len(files)} files from {folder_path} across {page_count} pages")
        
        return files
 
    async def find_file_by_name(self, filename: str) -> Optional[SharePointFile]:
        """
        Find a specific file by name across all tag folders.
        Returns the first match found.
        """
        all_files = await self.discover_tagged_files()
       
        for file in all_files:
            if file.name.lower() == filename.lower():
                return file
               
        return None
 
    async def find_file_by_relative_path(self, relative_path: str) -> Optional[SharePointFile]:
        """
        Find a specific file by its path relative to the base folder (e.g., 'CX/BE Team Presos/file.pptx').
        Returns the first match found (case-insensitive).
        """
        all_files = await self.discover_tagged_files()
        # Normalize input for comparison
        normalized_input = relative_path.strip().replace("\\", "/").lower()
        for file in all_files:
            # Remove base_folder_path from file.path, if present
            if file.path.lower().startswith(self.base_folder_path.lower() + "/"):
                rel_path = file.path[len(self.base_folder_path)+1:]
            else:
                rel_path = file.path
            if rel_path.replace("\\", "/").lower() == normalized_input:
                return file
        return None
 
    def _file_from_drive_item(self, item: dict, relative_hint: str = "") -> Optional[SharePointFile]:
        """Map a Graph driveItem JSON payload to SharePointFile."""
        if "file" not in item:
            return None
        name = item.get("name", "")
        parent_path = (item.get("parentReference") or {}).get("path") or ""
        library_rel = ""
        marker = "/root:"
        if marker in parent_path:
            library_rel = parent_path.split(marker, 1)[1].lstrip("/")
        full_path = f"{library_rel}/{name}".strip("/") if library_rel else name

        rel_to_base = full_path
        base = (self.base_folder_path or "").strip("/")
        if base and full_path.lower().startswith(base.lower() + "/"):
            rel_to_base = full_path[len(base) + 1 :]
        elif relative_hint and not relative_hint.lower().startswith("http"):
            rel_to_base = relative_hint.strip().replace("\\", "/")

        tag = rel_to_base.split("/", 1)[0] if "/" in rel_to_base else (rel_to_base or name)
        return SharePointFile(
            name=name,
            path=full_path,
            tag=tag,
            size_bytes=item.get("size", 0),
            mime_type=item.get("file", {}).get("mimeType", ""),
            download_url=item.get("@microsoft.graph.downloadUrl", ""),
            web_url=item.get("webUrl", ""),
            file_id=item.get("id", ""),
            created=item.get("createdDateTime", ""),
            modified=item.get("lastModifiedDateTime", ""),
        )

    @staticmethod
    def _encode_sharing_url(sharing_url: str) -> str:
        b64 = base64.urlsafe_b64encode(sharing_url.encode("utf-8")).decode("ascii").rstrip("=")
        return f"u!{b64}"

    async def get_file_by_sharing_url(self, sharing_url: str) -> Optional[SharePointFile]:
        """Resolve a SharePoint sharing / :x: / :b: / :w: link to a drive item."""
        await self._get_site_and_drive_ids()
        token = await self._get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        share_id = self._encode_sharing_url(sharing_url.strip())
        url = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem"
        try:
            response = self._make_request_with_retry("GET", url, headers=headers)
            if response.status_code == 404:
                logger.warning("Sharing URL did not resolve to a drive item")
                return None
            response.raise_for_status()
            item = response.json()
            sp_file = self._file_from_drive_item(item)
            if sp_file:
                logger.info(
                    "Resolved sharing URL to '%s' (path=%s, tag=%s)",
                    sp_file.name, sp_file.path, sp_file.tag,
                )
            return sp_file
        except Exception as e:
            logger.error("Error resolving SharePoint sharing URL: %s", e)
            return None

    async def get_file_by_relative_path(self, relative_path: str) -> Optional[SharePointFile]:
        """
        Fetch a file by path relative to the SharePoint base folder, or by a sharing URL.
        """
        rel_path = relative_path.strip().replace("\\", "/")
        if rel_path.lower().startswith(("http://", "https://")):
            return await self.get_file_by_sharing_url(rel_path)

        await self._get_site_and_drive_ids()
        token = await self._get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
        full_path = f"{self.base_folder_path}/{rel_path}" if self.base_folder_path else rel_path
        file_url = self._item_url(full_path)
        try:
            response = self._make_request_with_retry("GET", file_url, headers=headers)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            item = response.json()
            return self._file_from_drive_item(item, relative_hint=rel_path)
        except Exception as e:
            logger.error(f"Error fetching file by relative path '{relative_path}': {e}")
            return None
 
    async def download_file_to_temp(self, sharepoint_file: SharePointFile) -> str:
        """
        Download a SharePoint file to a temporary location.
        Returns the path to the temporary file.
        """
        try:
            file_extension = Path(sharepoint_file.name).suffix
            temp_file = tempfile.NamedTemporaryFile(
                delete=False,
                suffix=file_extension,
                prefix=f"sharepoint_{uuid.uuid4().hex[:8]}_"
            )
           
            # Use the improved download method with retry logic
            file_content = await self._download_with_retry(sharepoint_file)
            logger.info(
                "download_file_to_temp: '%s' (%d bytes) -> %s",
                sharepoint_file.name, len(file_content), temp_file.name,
            )
            temp_file.write(file_content)
            temp_file.close()
            logger.info(f"Downloaded SharePoint file '{sharepoint_file.name}' to {temp_file.name}")
            return temp_file.name
           
        except Exception as e:
            logger.error(f"Error downloading SharePoint file '{sharepoint_file.name}': {e}")
            raise
 
    async def get_file_content(self, sharepoint_file: SharePointFile) -> bytes:
        """Get file content as bytes"""
        try:
            return await self._download_with_retry(sharepoint_file)
        except Exception as e:
            logger.error(f"Error getting content for SharePoint file '{sharepoint_file.name}': {e}")
            raise
 
    async def list_files_in_folder_by_relative_path(self, relative_folder_path: str) -> list:
        """
        List all files in a folder (relative to base folder) recursively, returning SharePointFile objects.
        """
        await self._get_site_and_drive_ids()
        token = await self._get_access_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        # Normalize the path
        rel_path = relative_folder_path.strip().replace("\\", "/")
        # Compose the full path under the base folder
        full_path = f"{self.base_folder_path}/{rel_path}" if self.base_folder_path else rel_path
        tag = rel_path.split("/", 1)[0] if "/" in rel_path else rel_path
        files = await self._get_all_files_recursive(full_path, headers, tag)
        return files
 
    async def get_file_count(self, base_folder: str = None) -> Dict[str, int]:
        """
        Get a quick count of files in SharePoint organized by tags without downloading them.
        Returns a dictionary with tag names as keys and file counts as values.
        """
        await self._get_site_and_drive_ids()
        token = await self._get_access_token()
        headers = {
            'Authorization': f'Bearer {token}',
            'Content-Type': 'application/json'
        }
        
        tag_counts = {}
        base_folder = self._normalize_folder_path(base_folder or self.base_folder_path)
        
        try:
            # Get tag folders (direct subfolders under base folder)
            tag_folders = await self._get_folders_in_path(base_folder, headers)
            logger.info(f"Found {len(tag_folders)} tag folders under {base_folder}")
            
            total_files = 0
            for tag_folder in tag_folders:
                tag_name = tag_folder['name']
                tag_folder_path = f"{base_folder}/{tag_name}"
                # Count files recursively
                file_count = await self._count_files_recursive(tag_folder_path, headers)
                tag_counts[tag_name] = file_count
                total_files += file_count
                logger.info(f"Tag '{tag_name}': {file_count} files")
            
            tag_counts['TOTAL'] = total_files
            logger.info(f"Total file count: {total_files}")
            return tag_counts
            
        except Exception as e:
            logger.error(f"Error counting files from SharePoint: {e}")
            raise

    async def _count_files_recursive(self, folder_path: str, headers: Dict) -> int:
        """Recursively count files in a folder and its subfolders without downloading them"""
        file_count = 0
        try:
            list_url = self._children_url(folder_path)
            
            while list_url:
                try:
                    response = self._make_request_with_retry('GET', list_url, headers=headers)
                    response.raise_for_status()
                    response_data = response.json()
                    items = response_data.get("value", [])
                    
                    for item in items:
                        if 'folder' in item:
                            # It's a subfolder, recurse into it
                            subfolder_path = f"{folder_path}/{item['name']}"
                            subfile_count = await self._count_files_recursive(subfolder_path, headers)
                            file_count += subfile_count
                        else:
                            # It's a file, count it
                            file_count += 1
                    
                    # Check for next page
                    list_url = response_data.get("@odata.nextLink")
                    
                except Exception as e:
                    logger.error(f"Error counting files from {folder_path}: {e}")
                    break
                
        except Exception as e:
            logger.error(f"Error counting files from {folder_path}: {e}")
        
        return file_count
 
    async def get_files_modified_after(self, after_time: datetime, base_folder: str = None) -> List[SharePointFile]:
        """
        Get files modified after a specific time using Graph API ordering and early-stop.
        Avoids unsupported $filter on children and does not use local fallbacks.
        """
        start_time = time.time()
        try:
            await self._get_site_and_drive_ids()
            await self._refresh_token_if_needed()
            token = await self._get_access_token()
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }

            modified_files = []
            base_folder = self._normalize_folder_path(base_folder or self.base_folder_path)

            logger.info(f"🔍 Starting modified files search: base_folder='{base_folder}', after_time='{after_time.isoformat()}'")

            # Get tag folders (top-level under base folder)
            tag_folders = await self._get_folders_in_path(base_folder, headers)
            logger.info(f"📁 Found {len(tag_folders)} tag folders to check")

            for i, tag_folder in enumerate(tag_folders, 1):
                try:
                    tag_name = tag_folder['name']
                    tag_folder_path = f"{base_folder}/{tag_name}"

                    logger.info(f"📂 Checking {i}/{len(tag_folders)}: {tag_name}")
                    files_in_tag = await self._get_modified_files_ordered_recursive(
                        tag_folder_path, headers, tag_name, after_time
                    )
                    
                    if files_in_tag:
                        logger.info(f"{tag_name}: {len(files_in_tag)} files")
                    
                    modified_files.extend(files_in_tag)
                    
                except Exception as e:
                    logger.error(f"❌ Error in {tag_folder.get('name', 'unknown')}: {e}")
                    continue

            elapsed_time = time.time() - start_time
            logger.info(f"🎯 Search completed: {len(modified_files)} files found in {elapsed_time:.1f}s")
            
            return modified_files

        except Exception as e:
            elapsed_time = time.time() - start_time
            logger.error(f"❌ Error getting modified files from SharePoint after {elapsed_time:.1f}s: {e}")
            # Return empty list instead of raising to prevent complete failure
            return []

    async def _get_modified_files_ordered_recursive(self, folder_path: str, headers: Dict, tag: str, after_time: datetime) -> List[SharePointFile]:
        """
        Recursively list children ordered by lastModifiedDateTime desc and collect files
        newer than after_time. Always checks subfolders regardless of folder modification time.
        """
        try:
            # Add timeout to prevent getting stuck on any single folder
            return await asyncio.wait_for(
                self._get_modified_files_ordered_recursive_internal(folder_path, headers, tag, after_time),
                timeout=300  # 5 minutes timeout per folder
            )
        except asyncio.TimeoutError:
            logger.error(f"⏰ Timeout while processing folder: {folder_path}")
            return []
        except Exception as e:
            logger.error(f"Error processing folder {folder_path}: {e}")
            return []

    async def _get_modified_files_ordered_recursive_internal(self, folder_path: str, headers: Dict, tag: str, after_time: datetime) -> List[SharePointFile]:
        """
        Internal implementation of recursive file discovery
        """
        files: List[SharePointFile] = []

        list_url = self._children_url(folder_path)
        params = {
            "$orderby": "lastModifiedDateTime desc",
            "$select": "id,name,size,file,folder,webUrl,lastModifiedDateTime,createdDateTime,@microsoft.graph.downloadUrl",
        }

        try:
            subfolder_count = 0
            while list_url:
                try:
                    response = self._make_request_with_retry('GET', list_url, headers=headers, params=params)
                    params = None  # nextLink already contains its own query params
                    response.raise_for_status()
                    response_data = response.json()
                    items = response_data.get("value", [])

                    for item in items:
                        item_modified_iso = item.get('lastModifiedDateTime')
                        item_modified_dt = None
                        if item_modified_iso:
                            try:
                                item_modified_dt = datetime.fromisoformat(item_modified_iso.replace('Z', '+00:00'))
                            except Exception:
                                item_modified_dt = None

                        # If it's a folder, ALWAYS recurse into it to check for modified files
                        if 'folder' in item:
                            subfolder_count += 1
                            subfolder_path = f"{folder_path}/{item.get('name', '')}"
                            # Only log every 5th subfolder to avoid flooding
                            if subfolder_count % 5 == 1:
                                logger.info(f"  📁 Scanning subfolder {subfolder_count}: {item.get('name', '')}")
                            subfiles = await self._get_modified_files_ordered_recursive(subfolder_path, headers, tag, after_time)
                            files.extend(subfiles)
                            continue

                        # It's a file - check if it's modified after our threshold
                        if item_modified_dt and item_modified_dt > after_time:
                            sharepoint_file = SharePointFile(
                                name=item.get('name', ''),
                                path=f"{folder_path}/{item.get('name', '')}",
                                tag=tag,
                                size_bytes=item.get('size', 0),
                                mime_type=item.get('file', {}).get('mimeType', ''),
                                download_url=item.get('@microsoft.graph.downloadUrl', ''),
                                web_url=item.get('webUrl', ''),
                                file_id=item.get('id', ''),
                                created=item.get('createdDateTime', ''),
                                modified=item.get('lastModifiedDateTime', '')
                            )
                            files.append(sharepoint_file)
                            logger.info(f"📄 Found: {sharepoint_file.path}")

                    # Always follow next page to check all items
                    list_url = response_data.get("@odata.nextLink")
                        
                except Exception as e:
                    logger.error(f"Error listing ordered items from {folder_path}: {e}")
                    break

        except Exception as e:
            logger.error(f"Error listing ordered items from {folder_path}: {e}")

        # Log summary if we scanned many subfolders
        if subfolder_count > 0:
            logger.info(f"  Scanned {subfolder_count} subfolders in {folder_path.split('/')[-1]}")

        return files

    async def get_files_modified_after_fallback(self, after_time: datetime, base_folder: str = None) -> List[SharePointFile]:
        """
        Fallback method: Get all files and filter locally if Graph API filter fails
        """
        logger.info("Using fallback method - getting all files and filtering locally")
        all_files = await self.discover_tagged_files(base_folder)
        
        # Filter files modified after the specified time
        modified_files = []
        for file in all_files:
            try:
                modified_time = datetime.fromisoformat(file.modified.replace('Z', '+00:00'))
                if modified_time > after_time:
                    modified_files.append(file)
            except Exception as e:
                logger.warning(f"Error parsing modified time for {file.name}: {e}")
        
        logger.info(f"Fallback: Found {len(modified_files)} files modified after {after_time}")
        return modified_files
 
def _create_global_instance():
    """Lazy-create so import doesn't crash when env vars are missing."""
    try:
        return SharePointService()
    except Exception as e:
        logger.warning(f"SharePointService could not be initialized: {e}")
        return None

sharepoint_service = _create_global_instance()


if __name__ == "__main__":
    import sys

    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")

    svc = SharePointService()

    async def _test():
        print("=" * 60)
        print("STEP 1: Authenticate")
        print("=" * 60)
        token = await svc._get_access_token()
        print(f"  Token obtained: {token[:20]}...{token[-10:]}")

        print()
        print("=" * 60)
        print("STEP 2: Get Site & Drive IDs")
        print("=" * 60)
        site_id, drive_id = await svc._get_site_and_drive_ids()
        print(f"  Site ID:  {site_id}")
        print(f"  Drive ID: {drive_id}")

        print()
        print("=" * 60)
        print("STEP 3: Discover files")
        print("=" * 60)
        files = await svc.discover_tagged_files()
        print(f"  Found {len(files)} files total")
        for f in files[:20]:
            size_kb = f.size_bytes / 1024
            print(f"    [{f.tag}] {f.name} ({size_kb:.0f} KB)")
        if len(files) > 20:
            print(f"    ... and {len(files) - 20} more")

    asyncio.run(_test())