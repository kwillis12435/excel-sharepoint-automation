# Configuration settings for the application
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# SharePoint Settings
SHAREPOINT_SITE_URL = os.getenv('SHAREPOINT_SITE_URL', 'https://arrowres.sharepoint.com')
SHAREPOINT_CLIENT_ID = os.getenv('SHAREPOINT_CLIENT_ID')
SHAREPOINT_CLIENT_SECRET = os.getenv('SHAREPOINT_CLIENT_SECRET')

# SharePoint Paths
SHAREPOINT_BASE_PATH = "/sites/Discovery-Biology/Studies"

# Local Paths
EXCEL_FILE_PATH = os.getenv('EXCEL_FILE_PATH', 'downloaded_studies')
RESULTS_OUTPUT_PATH = os.getenv('RESULTS_OUTPUT_PATH', 'compiled_results')

# Validate required environment variables
required_vars = ['SHAREPOINT_CLIENT_ID', 'SHAREPOINT_CLIENT_SECRET']
missing_vars = [var for var in required_vars if not os.getenv(var)]
if missing_vars:
    raise ValueError(f"Missing required environment variables: {', '.join(missing_vars)}")