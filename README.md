# Excel SharePoint Automation

This project automates the process of compiling results tables from multiple Excel files stored on SharePoint. It connects to SharePoint, retrieves the necessary Excel files, and compiles the data into a single results table.

## Project Structure

```
excel-sharepoint-automation
├── src
│   ├── main.py            # Entry point of the application
│   ├── config.py          # Configuration settings for the application
│   ├── sharepoint         # Package for SharePoint-related functionalities
│   │   ├── __init__.py
│   │   └── connector.py    # Handles SharePoint connections and file operations
│   ├── excel              # Package for Excel-related functionalities
│   │   ├── __init__.py
│   │   ├── parser.py       # Parses Excel files and extracts data
│   │   └── compiler.py      # Compiles data from multiple Excel files
│   └── utils              # Package for utility functions
│       ├── __init__.py
│       └── helpers.py      # Contains helper functions for logging and data formatting
├── tests                  # Directory for unit tests
│   ├── __init__.py
│   ├── test_sharepoint.py  # Unit tests for SharePoint connector
│   └── test_excel.py       # Unit tests for Excel parser and compiler
├── config                 # Configuration files
│   └── settings.json       # JSON file for API endpoints and configuration parameters
├── requirements.txt       # Project dependencies
├── .gitignore             # Files and directories to ignore in version control
└── README.md              # Project documentation
```

## Setup Instructions

1. Clone the repository:
   ```
   git clone <repository-url>
   cd excel-sharepoint-automation
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Configure your SharePoint credentials and file paths in `src/config.py`.

4. Run the application:
   ```
   python src/main.py
   ```

## Usage Guidelines

- Ensure you have access to the SharePoint site and the necessary permissions to download files.
- Modify the configuration settings as needed to match your SharePoint environment.
- Use the provided unit tests to verify the functionality of the application.

## Contributing

Contributions are welcome! Please submit a pull request or open an issue for any enhancements or bug fixes.