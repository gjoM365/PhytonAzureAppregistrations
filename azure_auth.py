import msal
import requests
import logging
from config import Config

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class AzureAuthenticator:
    """
    Azure App Registration authenticator using MSAL
    Uses OAuth 2.0 client credentials flow for app-only authentication
    """
    
    def __init__(self):
        """Initialize the Azure authenticator with app registration details"""
        Config.validate_config()
        
        self.client_id = Config.AZURE_CLIENT_ID
        self.client_secret = Config.AZURE_CLIENT_SECRET
        self.tenant_id = Config.AZURE_TENANT_ID
        self.authority = Config.AUTHORITY
        self.scope = Config.SCOPE
        
        # Create MSAL confidential client app
        self.app = msal.ConfidentialClientApplication(
            client_id=self.client_id,
            client_credential=self.client_secret,
            authority=self.authority
        )
        
        self.access_token = None
    
    def get_access_token(self):
        """Get access token using client credentials flow"""
        try:
            # Try to get token from cache first
            result = self.app.acquire_token_silent(self.scope, account=None)
            
            if not result:
                # Get new token using client credentials
                result = self.app.acquire_token_for_client(scopes=self.scope)
            
            if "access_token" in result:
                self.access_token = result["access_token"]
                return self.access_token
            else:
                error_msg = f"Failed to acquire token: {result.get('error')}"
                logger.error(error_msg)
                raise Exception(error_msg)
                
        except Exception as e:
            logger.error(f"Error acquiring access token: {str(e)}")
            raise
    
    def get_auth_headers(self):
        """Get authorization headers for API requests"""
        if not self.access_token:
            self.get_access_token()
        
        return {
            'Authorization': f'Bearer {self.access_token}',
            'Content-Type': 'application/json'
        }
    
    def test_connection(self):
        """Test the Azure connection by calling Microsoft Graph API"""
        try:
            headers = self.get_auth_headers()
            
            # Test with sites endpoint (appropriate for app-only authentication)
            response = requests.get(
                f"{Config.GRAPH_API_ENDPOINT}/v1.0/sites",
                headers=headers
            )
            
            return response.status_code == 200
                    
        except Exception as e:
            logger.error(f"Error testing Azure connection: {str(e)}")
            return False 