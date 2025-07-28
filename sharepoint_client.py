import requests
import logging
import os
import tempfile
from datetime import datetime
from urllib.parse import quote
from azure_auth import AzureAuthenticator
from config import Config
from docx import Document

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class SharePointClient:
    """
    SharePoint client using Microsoft Graph API with Azure App Registration authentication
    Reference implementation for Sites.Selected permissions
    """
    
    def __init__(self):
        """Initialize SharePoint client with Azure authentication"""
        self.authenticator = AzureAuthenticator()
        self.site_url = Config.SHAREPOINT_SITE_URL
        self.tenant_name = Config.SHAREPOINT_TENANT_NAME
        self.graph_endpoint = Config.GRAPH_API_ENDPOINT
        
        # Extract site path from URL for Graph API calls
        self.site_path = self._extract_site_path()
        
    def _extract_site_path(self):
        """Extract site path from SharePoint URL for Graph API"""
        if '/sites/' in self.site_url:
            return self.site_url.split('.sharepoint.com')[1]
        else:
            raise ValueError("Invalid SharePoint site URL format")
    
    def connect(self):
        """Establish connection to SharePoint site"""
        try:
            # Test Azure authentication first
            if not self.authenticator.test_connection():
                return False
            
            # Test SharePoint site access
            site_info = self.get_site_info()
            return site_info is not None
                
        except Exception as e:
            logger.error(f"Error connecting to SharePoint: {str(e)}")
            return False
    
    def get_site_info(self):
        """Get basic information about the SharePoint site"""
        try:
            headers = self.authenticator.get_auth_headers()
            
            # Use Graph API to get site information
            site_id = f"{self.tenant_name}.sharepoint.com:{self.site_path}"
            url = f"{self.graph_endpoint}/v1.0/sites/{site_id}"
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.json()
            else:
                logger.error(f"Failed to get site info: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"Error getting site info: {str(e)}")
            return None
    
    def list_files(self, library_name="Documents", folder_path=""):
        """List files in a specific document library and folder"""
        try:
            headers = self.authenticator.get_auth_headers()
            
            # Get site ID first
            site_info = self.get_site_info()
            if not site_info:
                return None
            
            site_id = site_info['id']
            
            # Construct the path for the files
            path = f"/lists/{library_name}/items?$expand=fields,driveItem&$filter=fields/FileRef ne null"
            url = f"{self.graph_endpoint}/v1.0/sites/{site_id}{path}"
            
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                items = response.json().get('value', [])
                files = []
                
                for item in items:
                    if 'driveItem' in item and item['driveItem']:
                        file_info = {
                            'name': item['fields'].get('FileLeafRef', 'Unknown'),
                            'url': item['fields'].get('FileRef', ''),
                            'size': item['driveItem'].get('size', 0),
                            'modified': item['fields'].get('Modified', ''),
                            'id': item['driveItem'].get('id', '')
                        }
                        files.append(file_info)
                
                return files
            else:
                logger.error(f"Failed to list files: {response.status_code}")
                return None
                
        except Exception as e:
            logger.error(f"Error listing files: {str(e)}")
            return None
    
    def upload_file(self, local_file_path, remote_filename=None, library_name="Documents"):
        """Upload a file to SharePoint document library"""
        try:
            headers = self.authenticator.get_auth_headers()
            
            # Get site ID first
            site_info = self.get_site_info()
            if not site_info:
                return False
            
            site_id = site_info['id']
            
            # Use the filename from path if not provided
            if not remote_filename:
                remote_filename = os.path.basename(local_file_path)
            
            # Read the file content
            with open(local_file_path, 'rb') as file:
                file_content = file.read()
            
            # Get the drive ID for the document library
            drive_url = f"{self.graph_endpoint}/v1.0/sites/{site_id}/drives"
            drive_response = requests.get(drive_url, headers=headers)
            
            if drive_response.status_code != 200:
                return False
            
            drives = drive_response.json().get('value', [])
            documents_drive = None
            
            # Find the Documents drive
            for drive in drives:
                if drive.get('name') == library_name:
                    documents_drive = drive
                    break
            
            if not documents_drive:
                return False
            
            drive_id = documents_drive['id']
            
            # Upload the file
            upload_url = f"{self.graph_endpoint}/v1.0/sites/{site_id}/drives/{drive_id}/root:/{remote_filename}:/content"
            
            upload_headers = {
                'Authorization': headers['Authorization'],
                'Content-Type': 'application/octet-stream'
            }
            
            upload_response = requests.put(upload_url, headers=upload_headers, data=file_content)
            
            if upload_response.status_code in [200, 201]:
                return upload_response.json()
            else:
                logger.error(f"Failed to upload file: {upload_response.status_code}")
                return False
                
        except Exception as e:
            logger.error(f"Error uploading file: {str(e)}")
            return False 