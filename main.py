#!/usr/bin/env python3
"""
Azure AD App Registration to SharePoint Reference Implementation

This is a minimal reference showing how to connect to SharePoint using:
- Azure AD App Registration with Sites.Selected permissions
- Client credentials flow (app-only authentication)
- Microsoft Graph API

Configuration: env_SiteSelected.txt
"""

from sharepoint_client import SharePointClient
from azure_auth import AzureAuthenticator
from config import Config

def main():
    """Reference implementation for SharePoint connection"""
    print("Azure AD App Registration → SharePoint Connection")
    print("=" * 50)
    
    try:
        # 1. Validate configuration
        Config.validate_config()
        print("✓ Configuration loaded")
        
        # 2. Test Azure authentication
        authenticator = AzureAuthenticator()
        if not authenticator.test_connection():
            print("✗ Azure authentication failed")
            return False
        print("✓ Azure authentication successful")
        
        # 3. Connect to SharePoint
        sp_client = SharePointClient()
        if not sp_client.connect():
            print("✗ SharePoint connection failed")
            return False
        print("✓ SharePoint connection successful")
        
        # 4. Get site information
        site_info = sp_client.get_site_info()
        print(f"✓ Connected to: {site_info.get('displayName', 'Unknown Site')}")
        
        # 5. List files (basic operation)
        files = sp_client.list_files("Documents")
        if files:
            print(f"✓ Found {len(files)} files in Documents library")
            for file in files[:3]:  # Show first 3 files
                size_mb = file['size'] / (1024 * 1024) if file['size'] > 0 else 0
                print(f"  - {file['name']} ({size_mb:.2f} MB)")
        
        print("\n🎉 SharePoint connection successful!")
        return True
        
    except Exception as e:
        print(f"✗ Error: {e}")
        return False

if __name__ == "__main__":
    success = main()
    exit(0 if success else 1) 