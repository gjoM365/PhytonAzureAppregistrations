import os
from dotenv import load_dotenv

# Load environment variables from env_SiteSelected_example.txt file (with actual credentials)
# For production, use env_SiteSelected.txt as template
load_dotenv('env_SiteSelected_example.txt')
#load_dotenv('env_example.txt')

class Config:
    """Configuration class for Azure App Registration and SharePoint settings"""
    
    # Azure App Registration settings - Sites.Selected configuration
    AZURE_CLIENT_ID = os.getenv('AZURE_CLIENT_ID')
    AZURE_CLIENT_SECRET = os.getenv('AZURE_CLIENT_SECRET')
    AZURE_TENANT_ID = os.getenv('AZURE_TENANT_ID')
    
    # SharePoint settings
    SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL')
    SHAREPOINT_TENANT_NAME = os.getenv('SHAREPOINT_TENANT_NAME')
    
    # Microsoft Graph API settings
    GRAPH_API_ENDPOINT = 'https://graph.microsoft.com'
    SCOPE = ['https://graph.microsoft.com/.default']
    AUTHORITY = f'https://login.microsoftonline.com/{AZURE_TENANT_ID}'
    
    @classmethod
    def validate_config(cls):
        """Validate that all required configuration values are present"""
        required_vars = [
            'AZURE_CLIENT_ID',
            'AZURE_CLIENT_SECRET', 
            'AZURE_TENANT_ID',
            'SHAREPOINT_SITE_URL',
            'SHAREPOINT_TENANT_NAME'
        ]
        
        missing_vars = []
        for var in required_vars:
            if not getattr(cls, var):
                missing_vars.append(var)
        
        if missing_vars:
            raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")
        
        return True 