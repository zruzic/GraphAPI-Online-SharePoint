"""
GraphAPI Client for Microsoft Graph API - SharePoint Operations
Supports authentication, file management, document sets, security, and pages operations.
"""

import requests
import json
import os
from typing import Optional, Dict, Any, List, Tuple
from urllib.parse import urlencode


class GraphAPIClient:
    """Client for Microsoft Graph API operations on SharePoint"""

    def __init__(self, client_id: str, client_secret: str, tenant_id: str, tenant_name: str):
        """
        Initialize GraphAPI Client with Azure AD credentials.

        Args:
            client_id: Azure AD Application ID
            client_secret: Azure AD Application Secret
            tenant_id: Azure AD Tenant ID
            tenant_name: Tenant name (e.g., 'contoso')
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self.tenant_name = tenant_name
        self.base_url = "https://graph.microsoft.com/v1.0"
        self.access_token = None
        self.site_id = None
        self.drive_id = None
        self.session = requests.Session()

    def get_access_token(self) -> str:
        """
        Get access token from Azure AD using client credentials flow.

        Returns:
            str: Access token for API requests
        """
        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        data = {
            'grant_type': 'client_credentials',
            'client_id': self.client_id,
            'client_secret': self.client_secret,
            'scope': 'https://graph.microsoft.com/.default'
        }

        response = requests.post(token_url, data=data)
        response.raise_for_status()

        self.access_token = response.json()['access_token']
        return self.access_token

    def _get_headers(self, content_type: str = 'application/json') -> Dict[str, str]:
        """Get HTTP headers with authorization."""
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': content_type
        }

    def _request(self, method: str, url: str, **kwargs) -> requests.Response:
        """Make HTTP request to Graph API."""
        if not self.access_token:
            self.get_access_token()

        if 'headers' not in kwargs:
            kwargs['headers'] = self._get_headers()

        response = self.session.request(method, url, **kwargs)
        response.raise_for_status()
        return response

    # ==================== Authentication & Setup ====================

    def get_site_id(self, site_name: str) -> str:
        """
        Get Site ID from site name.

        Args:
            site_name: SharePoint site name

        Returns:
            str: Site ID
        """
        url = f"{self.base_url}/sites/{self.tenant_name}.sharepoint.com:/sites/{site_name}"
        response = self._request('GET', url)
        self.site_id = response.json()['id']
        return self.site_id

    def get_drive_id(self) -> str:
        """
        Get Drive ID (default document library) for the site.

        Returns:
            str: Drive ID
        """
        if not self.site_id:
            raise ValueError("Site ID not set. Call get_site_id first.")

        url = f"{self.base_url}/sites/{self.site_id}/drives"
        response = self._request('GET', url)
        drives = response.json()['value']

        # Get the first drive (default document library)
        self.drive_id = drives[0]['id']
        return self.drive_id

    # ==================== File Management ====================

    def create_folder(self, folder_name: str, parent_id: str = 'root') -> Dict[str, Any]:
        """Create a folder."""
        if not self.drive_id:
            raise ValueError("Drive ID not set. Call get_drive_id first.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{parent_id}/children"

        data = {
            "name": folder_name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "rename"
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def delete_file(self, file_id: str) -> None:
        """Delete a file."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}"
        self._request('DELETE', url)

    def rename_file(self, file_id: str, new_name: str) -> Dict[str, Any]:
        """Rename a file."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}"

        data = {"name": new_name}

        response = self._request('PATCH', url, json=data)
        return response.json()

    def copy_file(self, file_id: str, new_name: str, destination_id: str = 'root') -> Dict[str, Any]:
        """Copy a file."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/copy"

        data = {
            "parentReference": {
                "driveId": self.drive_id,
                "id": destination_id
            },
            "name": new_name
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def move_file(self, file_id: str, destination_folder_id: str) -> Dict[str, Any]:
        """Move a file to a different folder."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}"

        data = {
            "parentReference": {
                "id": destination_folder_id
            }
        }

        response = self._request('PATCH', url, json=data)
        return response.json()

    def get_file_metadata(self, file_id: str) -> Dict[str, Any]:
        """Get file metadata."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}"
        response = self._request('GET', url)
        return response.json()

    def get_folder_contents(self, folder_id: str) -> List[Dict[str, Any]]:
        """Get contents of a folder."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{folder_id}/children"
        response = self._request('GET', url)
        return response.json()['value']

    def search_files(self, search_term: str) -> List[Dict[str, Any]]:
        """Search for files by name."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/root/search(q='{search_term}')"
        response = self._request('GET', url)
        return response.json()['value']

    def download_file(self, file_id: str, output_path: str) -> None:
        """Download a file."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/content"
        response = self._request('GET', url)

        with open(output_path, 'wb') as f:
            f.write(response.content)

    def upload_file(self, file_path: str, folder_id: str = 'root') -> Dict[str, Any]:
        """Upload a file (simple upload for files < 4MB)."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        filename = os.path.basename(file_path)

        url = f"{self.base_url}/drives/{self.drive_id}/items/{folder_id}:/{filename}:/content"

        with open(file_path, 'rb') as f:
            response = self._request('PUT', url, data=f.read(),
                                   headers={'Authorization': f'Bearer {self.access_token}'})

        return response.json()

    def create_upload_session(self, filename: str, folder_id: str = 'root') -> Dict[str, Any]:
        """Create upload session for large files."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{folder_id}:/{filename}:/createUploadSession"

        data = {
            "item": {
                "@microsoft.graph.conflictBehavior": "rename",
                "name": filename
            }
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def upload_chunk(self, session_url: str, chunk_data: bytes,
                    start: int, end: int, total_size: int) -> Dict[str, Any]:
        """Upload a chunk for large file upload."""
        headers = {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Length': str(len(chunk_data)),
            'Content-Range': f'bytes {start}-{end}/{total_size}'
        }

        response = self._request('PUT', session_url, data=chunk_data, headers=headers)
        return response.json()

    def get_upload_session_status(self, session_url: str) -> Dict[str, Any]:
        """Get upload session status."""
        response = self._request('GET', session_url)
        return response.json()

    def cancel_upload_session(self, session_url: str) -> None:
        """Cancel an upload session."""
        self._request('DELETE', session_url)

    # ==================== Document Sets ====================

    def get_all_document_sets(self, list_id: str) -> List[Dict[str, Any]]:
        """Get all document sets in a list."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items"
        params = {'$filter': "contentType/name eq 'Document Set'"}

        response = self._request('GET', url, params=params)
        return response.json()['value']

    def get_document_set_details(self, list_id: str, docset_id: str) -> Dict[str, Any]:
        """Get document set details."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{docset_id}"
        response = self._request('GET', url)
        return response.json()

    def create_document_set(self, list_id: str, title: str) -> Dict[str, Any]:
        """Create a document set."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items"

        data = {
            "fields": {
                "Title": title,
                "ContentTypeId": "0x0120D520"
            }
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def update_document_set(self, list_id: str, docset_id: str, title: str) -> Dict[str, Any]:
        """Update document set properties."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{docset_id}"

        data = {
            "fields": {
                "Title": title,
                "ContentTypeId": "0x0120D520"
            }
        }

        response = self._request('PATCH', url, json=data)
        return response.json()

    def get_documents_in_set(self, docset_id: str) -> List[Dict[str, Any]]:
        """Get documents in a document set."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{docset_id}/children"
        response = self._request('GET', url)
        return response.json()['value']

    def delete_document_set(self, list_id: str, docset_id: str) -> None:
        """Delete a document set."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{docset_id}"
        self._request('DELETE', url)

    # ==================== Security & Sharing ====================

    def get_item_permissions(self, file_id: str) -> List[Dict[str, Any]]:
        """Get item permissions."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/permissions"
        response = self._request('GET', url)
        return response.json()['value']

    def share_with_user(self, file_id: str, email: str, role: str = 'edit') -> Dict[str, Any]:
        """Share item with a user."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/invite"

        data = {
            "recipients": [{"email": email}],
            "roles": [role],
            "requireSignIn": True
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def create_sharing_link(self, file_id: str, link_type: str = 'view',
                           scope: str = 'anonymous') -> Dict[str, Any]:
        """Create a sharing link (view or edit)."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/createLink"

        data = {
            "type": link_type,
            "scope": scope
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def delete_permission(self, file_id: str, permission_id: str) -> None:
        """Delete a permission."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/permissions/{permission_id}"
        self._request('DELETE', url)

    def update_permission_role(self, file_id: str, permission_id: str, role: str = 'read') -> Dict[str, Any]:
        """Update permission role."""
        if not self.drive_id:
            raise ValueError("Drive ID not set.")

        url = f"{self.base_url}/drives/{self.drive_id}/items/{file_id}/permissions/{permission_id}"

        data = {"roles": [role]}

        response = self._request('PATCH', url, json=data)
        return response.json()

    # ==================== Pages Management ====================

    def get_all_pages(self) -> List[Dict[str, Any]]:
        """Get all pages in the site."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages"
        response = self._request('GET', url)
        return response.json()['value']

    def get_page_details(self, page_id: str) -> Dict[str, Any]:
        """Get page details."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages/{page_id}"
        response = self._request('GET', url)
        return response.json()

    def create_page(self, name: str, title: str) -> Dict[str, Any]:
        """Create a new page."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages"

        data = {
            "name": name,
            "title": title,
            "layoutWebpartId": "3eb3e627-5144-4667-83d5-7662c6abb714"
        }

        response = self._request('POST', url, json=data)
        return response.json()

    def update_page(self, page_id: str, title: str, description: str = '') -> Dict[str, Any]:
        """Update page title and description."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages/{page_id}"

        data = {
            "title": title,
            "description": description
        }

        response = self._request('PATCH', url, json=data)
        return response.json()

    def publish_page(self, page_id: str) -> None:
        """Publish a page."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages/{page_id}/publish"
        self._request('POST', url)

    def delete_page(self, page_id: str) -> None:
        """Delete a page."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages/{page_id}"
        self._request('DELETE', url)

    def add_web_part_to_page(self, page_id: str, web_part_data: Dict[str, Any]) -> Dict[str, Any]:
        """Add a web part to a page."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/pages/{page_id}/webparts"

        data = {"webPartData": web_part_data}

        response = self._request('POST', url, json=data)
        return response.json()

    # ==================== List & List Items ====================

    def get_lists(self) -> List[Dict[str, Any]]:
        """Get all lists in the site."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists"
        response = self._request('GET', url)
        return response.json()['value']

    def create_list_item(self, list_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Create a list item."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items"

        data = {"fields": fields}

        response = self._request('POST', url, json=data)
        return response.json()

    def get_list_item(self, list_id: str, item_id: str) -> Dict[str, Any]:
        """Get a list item."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{item_id}"
        response = self._request('GET', url)
        return response.json()

    def update_list_item(self, list_id: str, item_id: str, fields: Dict[str, Any]) -> Dict[str, Any]:
        """Update a list item."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{item_id}"

        data = {"fields": fields}

        response = self._request('PATCH', url, json=data)
        return response.json()

    def delete_list_item(self, list_id: str, item_id: str) -> None:
        """Delete a list item."""
        if not self.site_id:
            raise ValueError("Site ID not set.")

        url = f"{self.base_url}/sites/{self.site_id}/lists/{list_id}/items/{item_id}"
        self._request('DELETE', url)
