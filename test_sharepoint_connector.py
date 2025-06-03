from src.sharepoint.connector import SharePointConnector
from dotenv import load_dotenv
import os

load_dotenv(dotenv_path="src/.env")


SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
CLIENT_ID = os.getenv("SHAREPOINT_CLIENT_ID")
CLIENT_SECRET = os.getenv("SHAREPOINT_CLIENT_SECRET")
TENANT_ID = os.getenv("SHAREPOINT_TENANT_ID")


FOLDER_PATH = "/sites/Discovery-Biology/Studies/2024/01 - 2024/2024011201 (hALK7_AAV8_KD_7)"  # Example

def main():
    connector = SharePointConnector(SITE_URL, CLIENT_ID, CLIENT_SECRET, TENANT_ID)
    try:
        connector.connect()
        print("Connection successful!")
        files = connector.list_files(FOLDER_PATH)
        print("Files in folder:", files)
    except Exception as e:
        print("Error connecting to SharePoint:", e)

if __name__ == "__main__":
    main() 