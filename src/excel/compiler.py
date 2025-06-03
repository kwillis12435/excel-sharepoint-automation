from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File
import pandas as pd
from datetime import datetime
import os
import config  # Import the config module

class ResultsCompiler:
    def __init__(self):
        self.results = []

    def compile_results(self, data_frames):
        """Compiles results from multiple data frames into a single results table."""
        for df in data_frames:
            self.results.append(df)
        # Assuming all data frames have the same structure, concatenate them
        compiled_results = pd.concat(self.results, ignore_index=True)
        return compiled_results

    def save_results(self, compiled_results, output_path):
        """Saves the compiled results to an Excel file."""
        compiled_results.to_excel(output_path, index=False)

class SharePointExcelManager:
    def __init__(self):
        """
        Initialize SharePoint connection using config values
        """
        self.site_url = config.SHAREPOINT_SITE_URL
        self.ctx = ClientContext(self.site_url).with_credentials(
            ClientCredential(
                config.SHAREPOINT_CLIENT_ID,
                config.SHAREPOINT_CLIENT_SECRET
            )
        )

    def list_years(self, base_path):
        # Implementation of list_years method
        pass

    def list_months(self, year):
        # Implementation of list_months method
        pass

    def list_studies(self, year, month):
        # Implementation of list_studies method
        pass

    def download_excel(self, file_url, local_path):
        # Implementation of download_excel method
        pass

    def process_all_studies(self):
        """
        Process all studies across all years and months
        """
        os.makedirs(config.EXCEL_FILE_PATH, exist_ok=True)
        
        for year in self.list_years(config.SHAREPOINT_BASE_PATH):
            for month in self.list_months(year):
                studies = self.list_studies(year, month)
                for study in studies:
                    file_url = f"{config.SHAREPOINT_BASE_PATH}/{year}/{month}/{study}"
                    local_path = os.path.join(config.EXCEL_FILE_PATH, f"{year}_{month}_{study}")
                    self.download_excel(file_url, local_path)
                    print(f"Downloaded: {study} from {month} {year}")

def main():
    # Initialize the manager with config values
    sp_manager = SharePointExcelManager()

    # Process all studies
    sp_manager.process_all_studies()

if __name__ == "__main__":
    main()

# Example usage for specific operations
sp_manager = SharePointExcelManager()

# List all years
years = sp_manager.list_years()
print(years)

# List months for a specific year
months = sp_manager.list_months("2024")
print(months)

# List studies for a specific year/month
studies = sp_manager.list_studies("2024", "March")
print(studies)