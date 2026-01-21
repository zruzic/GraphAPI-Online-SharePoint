#!/usr/bin/env python3
"""
Example: File Management Operations using GraphAPI Client

This script demonstrates common file operations:
- Create folders
- Upload files
- Download files
- Rename/Copy/Move files
- Search for files
"""

import sys
import os
from graphapi_client import GraphAPIClient

def main():
    # Configuration
    CLIENT_ID = "your_client_id"
    CLIENT_SECRET = "your_client_secret"
    TENANT_ID = "your_tenant_id"
    TENANT_NAME = "your_tenant_name"  # e.g., 'contoso'
    SITE_NAME = "your_site_name"

    try:
        # Initialize client
        print("[*] Initializing GraphAPI Client...")
        client = GraphAPIClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID, TENANT_NAME)

        # Get access token
        print("[*] Authenticating...")
        client.get_access_token()
        print("[+] Authentication successful")

        # Get Site ID
        print(f"[*] Getting Site ID for site: {SITE_NAME}...")
        site_id = client.get_site_id(SITE_NAME)
        print(f"[+] Site ID: {site_id}")

        # Get Drive ID
        print("[*] Getting Drive ID...")
        drive_id = client.get_drive_id()
        print(f"[+] Drive ID: {drive_id}")

        # Create a folder
        print("[*] Creating folder 'TestFolder'...")
        folder_response = client.create_folder("TestFolder")
        folder_id = folder_response['id']
        print(f"[+] Folder created: {folder_response['name']} (ID: {folder_id})")

        # Get folder contents
        print(f"[*] Getting folder contents...")
        contents = client.get_folder_contents("root")
        print(f"[+] Found {len(contents)} items in root folder")
        for item in contents[:5]:
            print(f"   - {item.get('name')} (ID: {item.get('id')})")

        # Search for files
        print("[*] Searching for files with 'test' in name...")
        search_results = client.search_files("test")
        print(f"[+] Found {len(search_results)} results")
        for result in search_results[:5]:
            print(f"   - {result.get('name')}")

        # Example: Upload file (if file exists)
        if os.path.exists("sample.txt"):
            print("[*] Uploading 'sample.txt'...")
            upload_response = client.upload_file("sample.txt", folder_id)
            file_id = upload_response['id']
            print(f"[+] File uploaded: {upload_response['name']} (ID: {file_id})")

            # Get file metadata
            print("[*] Getting file metadata...")
            metadata = client.get_file_metadata(file_id)
            print(f"[+] File size: {metadata.get('size')} bytes")
            print(f"[+] Created: {metadata.get('createdDateTime')}")

            # Rename file
            print("[*] Renaming file to 'sample_renamed.txt'...")
            rename_response = client.rename_file(file_id, "sample_renamed.txt")
            print(f"[+] File renamed: {rename_response['name']}")

            # Copy file
            print("[*] Copying file...")
            copy_response = client.copy_file(file_id, "sample_copy.txt", folder_id)
            print(f"[+] File copy initiated (Operation ID: {copy_response.get('id')})")

            # Download file
            print("[*] Downloading file...")
            client.download_file(file_id, "downloaded_sample.txt")
            print("[+] File downloaded to 'downloaded_sample.txt'")

        print("\n[✓] File operations example completed successfully!")

    except Exception as e:
        print(f"[!] Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
