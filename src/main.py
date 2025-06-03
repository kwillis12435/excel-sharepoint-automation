# Contents of /excel-sharepoint-automation/excel-sharepoint-automation/src/main.py

import os
from config import settings
from sharepoint.connector import SharePointConnector
from excel.parser import ExcelParser
from excel.compiler import ResultsCompiler

def main():
    # Initialize SharePoint connector
    sp_connector = SharePointConnector(settings.SHAREPOINT_URL, settings.SHAREPOINT_CREDENTIALS)
    
    # Connect to SharePoint
    sp_connector.connect()
    
    # List and download Excel files
    excel_files = sp_connector.list_files()
    for file in excel_files:
        sp_connector.download_file(file)

    # Parse and compile results from downloaded Excel files
    parser = ExcelParser()
    compiler = ResultsCompiler()
    
    for file in os.listdir(settings.DOWNLOAD_DIRECTORY):
        if file.endswith('.xlsx'):
            data = parser.parse_file(os.path.join(settings.DOWNLOAD_DIRECTORY, file))
            compiler.compile_results(data)

    # Save the compiled results
    compiler.save_results(settings.RESULTS_FILE_PATH)

if __name__ == "__main__":
    main()