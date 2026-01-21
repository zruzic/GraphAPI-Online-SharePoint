#!/usr/bin/env python3
"""
Example: SharePoint Pages Management using GraphAPI Client

This script demonstrates page operations:
- Create pages
- Get page details
- Update page properties
- Publish pages
- Delete pages
- Add web parts
"""

import sys
from graphapi_client import GraphAPIClient

def main():
    # Configuration
    CLIENT_ID = "your_client_id"
    CLIENT_SECRET = "your_client_secret"
    TENANT_ID = "your_tenant_id"
    TENANT_NAME = "your_tenant_name"
    SITE_NAME = "your_site_name"

    try:
        # Initialize and authenticate
        print("[*] Initializing GraphAPI Client...")
        client = GraphAPIClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID, TENANT_NAME)

        print("[*] Authenticating...")
        client.get_access_token()

        print(f"[*] Getting Site ID for site: {SITE_NAME}...")
        site_id = client.get_site_id(SITE_NAME)
        print(f"[+] Site ID: {site_id}")

        # Get all pages
        print("[*] Retrieving all pages...")
        pages = client.get_all_pages()
        print(f"[+] Found {len(pages)} pages")
        for page in pages[:5]:
            print(f"   - {page.get('name')} (Title: {page.get('title')})")

        # Create a new page
        print("[*] Creating new page...")
        page_response = client.create_page("DemoPage.aspx", "Demo Page Title")
        page_id = page_response['id']
        print(f"[+] Page created: {page_response.get('name')}")
        print(f"    ID: {page_id}")
        print(f"    URL: {page_response.get('webUrl')}")

        # Get page details
        print(f"[*] Getting page details...")
        page_details = client.get_page_details(page_id)
        print(f"[+] Page Title: {page_details.get('title')}")
        print(f"[+] Description: {page_details.get('description')}")
        print(f"[+] Status: {page_details.get('publishingState')}")

        # Update page
        print("[*] Updating page title and description...")
        update_response = client.update_page(
            page_id,
            title="Updated Demo Page",
            description="This is an updated demo page with more information"
        )
        print(f"[+] Page updated: {update_response.get('title')}")

        # Add a web part to the page
        print("[*] Adding web part to page...")
        web_part_data = {
            "id": "webpartid",
            "instanceId": "00000000-0000-0000-0000-000000000000",
            "title": "Welcome Text",
            "serverProcessedContent": {
                "htmlStrings": {},
                "searchablePlainTexts": ["Welcome to this page"],
                "imageSources": [],
                "links": []
            },
            "dataVersion": "1.0",
            "properties": {
                "text": "Welcome to this page"
            }
        }

        try:
            webpart_response = client.add_web_part_to_page(page_id, web_part_data)
            print(f"[+] Web part added successfully")
        except Exception as e:
            print(f"[!] Web part addition error (may be expected): {e}")

        # Publish the page
        print("[*] Publishing page...")
        try:
            client.publish_page(page_id)
            print("[+] Page published successfully")
        except Exception as e:
            print(f"[!] Publishing error: {e}")

        # Get updated page details
        print("[*] Getting final page details...")
        final_details = client.get_page_details(page_id)
        print(f"[+] Final Title: {final_details.get('title')}")
        print(f"[+] Status: {final_details.get('publishingState')}")

        # Delete the page (optional - comment out to keep the page)
        # print("[*] Deleting page...")
        # client.delete_page(page_id)
        # print("[+] Page deleted")

        print("\n[✓] Pages management example completed successfully!")

    except Exception as e:
        print(f"[!] Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
