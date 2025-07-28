# Azure AD App Registration to SharePoint Reference

A minimal reference implementation showing how to connect to SharePoint using Azure AD App Registration with Sites.Selected permissions and client credentials flow.

## Features

- ✅ Azure AD App Registration authentication using MSAL
- ✅ Microsoft Graph API integration (modern approach)
- ✅ OAuth 2.0 client credentials flow (app-only authentication)
- ✅ Sites.Selected permissions support
- ✅ Basic SharePoint operations (connect, list files, upload)

## Prerequisites

- Python 3.7+
- Azure Active Directory tenant
- SharePoint Online site
- Azure App Registration with Sites.Selected permissions

## Quick Start

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Configure environment:**
   Update `env_SiteSelected.txt` with your Azure app details:
   ```
   AZURE_CLIENT_ID=your_client_id
   AZURE_CLIENT_SECRET=your_client_secret
   AZURE_TENANT_ID=your_tenant_id
   SHAREPOINT_SITE_URL=https://tenant.sharepoint.com/sites/sitename
   SHAREPOINT_TENANT_NAME=tenant
   ```

3. **Run:**
   ```bash
   python3 main.py
   ```

## Azure Setup

### 1. Create App Registration
1. Go to [Azure Portal](https://portal.azure.com) → Azure Active Directory → App registrations
2. Click "New registration"
3. Name: e.g., "SharePoint Connector"
4. Account types: "Accounts in this organizational directory only"
5. Click "Register"

### 2. Configure Permissions
1. Go to "API permissions"
2. Add Microsoft Graph → Application permissions:
   - `Sites.Selected` (for specific site access)
   - `Files.ReadWrite.All` (for file operations)
3. **Grant admin consent**

### 3. Create Client Secret
1. Go to "Certificates & secrets"
2. Create new client secret
3. Copy the **VALUE** (not ID)

### 4. Grant Site Access
Since using Sites.Selected, you need to grant specific site access:
```bash
# Use Microsoft Graph PowerShell or REST API
# Grant permission to your specific SharePoint site
```

## Project Structure

```
├── env_SiteSelected.txt     # Configuration
├── config.py               # Config loader
├── azure_auth.py           # Azure authentication
├── sharepoint_client.py    # SharePoint operations
├── main.py                 # Reference implementation
└── requirements.txt        # Dependencies
```

## Key Classes

### AzureAuthenticator
Handles Azure AD authentication using MSAL with client credentials flow.

### SharePointClient  
Provides SharePoint operations via Microsoft Graph API.

## Usage Example

```python
from sharepoint_client import SharePointClient

# Initialize and connect
sp_client = SharePointClient()
sp_client.connect()

# List files
files = sp_client.list_files("Documents")

# Upload file
sp_client.upload_file("local_file.pdf", "remote_file.pdf")
```

## Security Notes

- Uses modern OAuth 2.0 client credentials flow
- No user interaction required (app-only authentication)
- Client secret should be secured (use Azure Key Vault in production)
- Sites.Selected provides principle of least privilege

## Dependencies

- `msal` - Microsoft Authentication Library
- `requests` - HTTP library
- `python-dotenv` - Environment variables
- `python-docx` - Word document creation

## License

Reference implementation for educational purposes. 