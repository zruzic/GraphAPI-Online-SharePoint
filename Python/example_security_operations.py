#!/usr/bin/env python3
"""
Example: Security & Sharing Operations using GraphAPI Client

This script demonstrates security operations:
- Get item permissions
- Share files with users
- Create sharing links
- Update permission roles
- Delete permissions
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
    FILE_ID = "file_id_to_share"  # Replace with actual file ID
    USER_EMAIL = "user@example.com"  # Replace with actual user email

    try:
        # Initialize and authenticate
        print("[*] Initializing GraphAPI Client...")
        client = GraphAPIClient(CLIENT_ID, CLIENT_SECRET, TENANT_ID, TENANT_NAME)

        print("[*] Authenticating...")
        client.get_access_token()

        print(f"[*] Getting Site ID for site: {SITE_NAME}...")
        site_id = client.get_site_id(SITE_NAME)
        print(f"[+] Site ID: {site_id}")

        print("[*] Getting Drive ID...")
        drive_id = client.get_drive_id()
        print(f"[+] Drive ID: {drive_id}")

        # Get file metadata first
        print(f"[*] Getting file metadata for {FILE_ID}...")
        try:
            file_metadata = client.get_file_metadata(FILE_ID)
            print(f"[+] File: {file_metadata.get('name')}")
        except Exception as e:
            print(f"[!] Note: File ID might not exist. Error: {e}")
            print("[*] Using FILE_ID as example for permissions operations...")

        # Get current permissions
        print(f"[*] Getting current permissions for file...")
        try:
            permissions = client.get_item_permissions(FILE_ID)
            print(f"[+] Found {len(permissions)} permissions:")
            for perm in permissions:
                print(f"   - Permission ID: {perm.get('id')}")
                if 'grantedToIdentities' in perm:
                    identities = perm['grantedToIdentities']
                    if identities:
                        print(f"     Granted to: {identities[0].get('user', {}).get('displayName')}")
                print(f"     Roles: {perm.get('roles')}")
        except Exception as e:
            print(f"[*] Could not retrieve permissions (file may not exist): {e}")

        # Share with user (Edit access)
        print(f"\n[*] Sharing file with user (Edit access)...")
        try:
            share_response = client.share_with_user(FILE_ID, USER_EMAIL, role='edit')
            print(f"[+] Share invitation sent to {USER_EMAIL}")
            if 'value' in share_response:
                for grant in share_response['value']:
                    print(f"   Permission ID: {grant.get('id')}")
                    print(f"   Roles: {grant.get('roles')}")
        except Exception as e:
            print(f"[!] Share operation error (file may not exist): {e}")

        # Share with user (Read-only access)
        print(f"\n[*] Sharing file with user (Read-only access)...")
        try:
            share_readonly = client.share_with_user(FILE_ID, USER_EMAIL, role='read')
            print(f"[+] Read-only share invitation sent")
        except Exception as e:
            print(f"[!] Read-only share error: {e}")

        # Create anonymous sharing link (View)
        print(f"\n[*] Creating anonymous view-only sharing link...")
        try:
            link_view = client.create_sharing_link(FILE_ID, link_type='view', scope='anonymous')
            print(f"[+] View link created:")
            if 'link' in link_view:
                print(f"   URL: {link_view['link'].get('webUrl')}")
        except Exception as e:
            print(f"[!] View link creation error: {e}")

        # Create anonymous sharing link (Edit)
        print(f"\n[*] Creating anonymous edit sharing link...")
        try:
            link_edit = client.create_sharing_link(FILE_ID, link_type='edit', scope='anonymous')
            print(f"[+] Edit link created:")
            if 'link' in link_edit:
                print(f"   URL: {link_edit['link'].get('webUrl')}")
        except Exception as e:
            print(f"[!] Edit link creation error: {e}")

        # Update permission role
        print(f"\n[*] Updating permission role...")
        if permissions and len(permissions) > 0:
            perm_id = permissions[0].get('id')
            try:
                updated_perm = client.update_permission_role(FILE_ID, perm_id, role='read')
                print(f"[+] Permission updated to read-only")
                print(f"   New roles: {updated_perm.get('roles')}")
            except Exception as e:
                print(f"[!] Update permission error: {e}")

        # Delete permission
        print(f"\n[*] Deleting permission...")
        if permissions and len(permissions) > 1:
            perm_id = permissions[1].get('id')
            try:
                client.delete_permission(FILE_ID, perm_id)
                print(f"[+] Permission deleted")
            except Exception as e:
                print(f"[!] Delete permission error: {e}")

        print("\n[✓] Security & Sharing example completed!")

    except Exception as e:
        print(f"[!] Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
