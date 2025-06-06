import pandas as pd
import os
from openpyxl import load_workbook
import re
import json
import csv
from datetime import datetime
import traceback
from typing import Dict, List, Optional, Tuple, Any
#reset

# ========================== CONFIGURATION ==========================
class Config:
    """
    Central configuration class containing all constants and settings.
    This makes it easy to modify paths, cell locations, and parameters without 
    hunting through the entire codebase.
    """
    # Main folder containing all study subfolders to process
    MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    DEBUG = False  # Set to True to see detailed debug output during processing
    
    # Excel sheet names we're looking for in the workbooks
    PROCEDURE_SHEET = "Procedure Request Form"  # Contains study metadata
    RELATIVE_EXPRESSION_SHEETS = [              # Contains the actual expression data
        "Compiled Indiv. & Grp.",
        "Compiled Indiv. & Grp",
        "Compiled Indiv & Grp.",
        "Compiled Indiv & Grp",
        "Calcs Norm to D1 & Ctrl",
        "Calcs Norm to Pre & Control",
        "Calcs Norm to D1 & Ctrl."
    ]
    LAR_SHEET = "LAR Sheet"  # Contains additional metadata
    
    # Specific cell locations in the Procedure Request Form sheet
    STUDY_NAME_CELL = "C14"          # Where study name is stored
    SCREENING_MODEL_CELL = "M6"      # Where screening model is stored  
    STUDY_CODE_CELL = "M12"          # Where study code is stored
    TISSUES_START_ROW = 17           # Row where tissue list starts
    TISSUES_COLUMN = "S"             # Column containing tissue names
    TRIGGERS_START_ROW = 80          # Row where trigger list starts
    TRIGGERS_COLUMN = "B"            # Column containing trigger names
    DOSES_COLUMN = "D"               # Column containing doses (parallel to triggers)
    
    # Settings for relative expression data extraction
    REL_EXP_SEARCH_ROWS = range(120, 135)  # Rows to search for "Relative Expression" header
    TARGET_START_COLUMN = 6                # Column F - where targets start
    TARGET_COLUMN_SPACING = 4              # Targets are spaced 4 columns apart (F, J, N, etc.)
    MAX_TARGETS = 30                       # Safety limit to prevent infinite loops
    MAX_TRIGGERS = 20                      # Safety limit for trigger extraction

# ========================== UTILITY FUNCTIONS ==========================
def debug_print(*args, **kwargs):
    """
    Print debug messages only when DEBUG mode is enabled.
    This allows us to add debugging throughout the code without cluttering 
    normal output.
    """
    if Config.DEBUG:
        print(*args, **kwargs)

def is_empty_or_zero(value: Any) -> bool:
    """
    Check if a cell value should be considered empty.
    Excel cells can contain None, empty strings, spaces, or zeros.
    This function standardizes the check across all extraction functions.
    """
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    if value == 0 or str(value).strip() == "0":
        return True
    return False

def normalize_string(text: str) -> str:
    """
    Remove spaces, special characters, and convert to lowercase.
    Used for fuzzy string matching when trigger names might have 
    slight variations between metadata and results sheets.
    """
    if not text:
        return ""
    return ''.join(c for c in str(text).lower().strip() if c.isalnum())

def format_timepoint(timepoint: str) -> str:
    """
    Ensure timepoint starts with 'D' (e.g., "14" becomes "D14").
    Standardizes timepoint format across different data sources.
    """
    if not timepoint:
        return timepoint
    
    timepoint = str(timepoint).strip()
    if not timepoint.startswith('D') and timepoint and timepoint[0].isdigit():
        return f"D{timepoint}"
    return timepoint

def convert_to_numeric(value: Any) -> str:
    """
    Convert values to consistent numeric format with 4 decimal places.
    Handles various input types (None, strings, numbers) and ensures
    consistent formatting in the final CSV output.
    """
    if value is None:
        return ""
    try:
        float_val = float(str(value).strip())
        return f"{float_val:.4f}" if float_val != 0 else "0.0000"
    except (ValueError, TypeError):
        return str(value).strip()

def safe_workbook_operation(file_path: str, operation_func, *args, **kwargs):
    """
    Safely open Excel workbooks and ensure they're properly closed.
    This wrapper function handles errors and cleanup for all Excel operations,
    preventing memory leaks and file locking issues.
    """
    try:
        wb = load_workbook(file_path, data_only=True, read_only=kwargs.get('read_only', True))
        result = operation_func(wb, *args)
        wb.close()
        return result
    except Exception as e:
        print(f"Error processing {file_path}: {e}")
        if Config.DEBUG:
            traceback.print_exc()
        return None

# ========================== EXCEL DATA EXTRACTION ==========================
class ExcelExtractor:
    """
    Utility class containing reusable methods for extracting data from Excel worksheets.
    This centralizes common patterns like extracting columns, finding text, etc.
    """
    
    @staticmethod
    def extract_column_values(ws, start_row: int, column: str, stop_on_empty: bool = True) -> List[str]:
        """
        Extract all values from a single column starting at a specific row.
        
        Args:
            ws: Excel worksheet object
            start_row: Row number to start extraction (1-indexed)
            column: Column letter (e.g., 'B', 'S')
            stop_on_empty: Whether to stop when hitting an empty cell
            
        Returns:
            List of extracted string values
        """
        values = []
        row = start_row
        while row <= ws.max_row:
            cell_value = ws[f"{column}{row}"].value
            if stop_on_empty and is_empty_or_zero(cell_value):
                break
            if not is_empty_or_zero(cell_value):
                values.append(str(cell_value).strip())
            row += 1
        return values

    @staticmethod
    def extract_paired_columns(ws, start_row: int, col1: str, col2: str) -> Tuple[List[str], List[str]]:
        """
        Extract values from two parallel columns (e.g., triggers and doses).
        Stops when the first column becomes empty, ensuring both lists have same length.
        
        Args:
            ws: Excel worksheet object
            start_row: Row number to start extraction
            col1: First column letter (e.g., 'B' for triggers)
            col2: Second column letter (e.g., 'D' for doses)
            
        Returns:
            Tuple of (values_from_col1, values_from_col2)
        """
        values1, values2 = [], []
        row = start_row
        while row <= ws.max_row:
            val1 = ws[f"{col1}{row}"].value
            if is_empty_or_zero(val1):
                break
            val2 = ws[f"{col2}{row}"].value
            values1.append(str(val1).strip())
            values2.append(val2 if val2 is not None else None)
            row += 1
        return values1, values2

    @staticmethod
    def find_cell_with_text(ws, search_text: str, search_rows: range = None) -> Optional[Tuple[int, int]]:
        """
        Find the first cell containing specific text (case-insensitive).
        
        Args:
            ws: Excel worksheet object
            search_text: Text to search for
            search_rows: Optional range of rows to limit search
            
        Returns:
            Tuple of (row, column) if found, None otherwise
        """
        search_range = search_rows or range(1, min(ws.max_row, 200))
        for row in search_range:
            for col in range(1, min(ws.max_column, 50)):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str) and search_text.lower() in cell_value.lower():
                    return row, col
        return None

    @staticmethod
    def extract_targets_from_row(ws, row: int) -> Tuple[List[str], List[int]]:
        """
        Extract target gene names and their column positions from a specific row.
        Targets are typically spaced every 4 columns (F, J, N, R, etc.).
        
        Args:
            ws: Excel worksheet object
            row: Row number containing target names
            
        Returns:
            Tuple of (target_names, column_numbers)
        """
        targets, target_columns = [], []
        col_start = Config.TARGET_START_COLUMN  # Start at column F (6)
        zero_count = 0
        
        while len(targets) < Config.MAX_TARGETS:
            target_value = ws.cell(row=row, column=col_start).value
            
            if is_empty_or_zero(target_value):
                zero_count += 1
                if zero_count >= 5:  # Stop after 5 consecutive empty cells
                    break
            else:
                zero_count = 0
                targets.append(str(target_value).strip())
                target_columns.append(col_start)
            
            col_start += Config.TARGET_COLUMN_SPACING  # Move to next target column
        
        return targets, target_columns

# ========================== STUDY METADATA EXTRACTION ==========================
def extract_study_metadata(wb, folder_name: str) -> Dict[str, Any]:
    """
    Extract all metadata from the 'Procedure Request Form' sheet.
    This includes study name, code, screening model, tissues, triggers, doses, and timepoint.
    
    Args:
        wb: Excel workbook object
        folder_name: Name of the study folder (used as fallback for study code)
        
    Returns:
        Dictionary containing all extracted metadata
    """
    if Config.PROCEDURE_SHEET not in wb.sheetnames:
        return _extract_fallback_metadata(folder_name)
    
    ws = wb[Config.PROCEDURE_SHEET]
    
    # Extract basic metadata from specific cells
    study_name = ws[Config.STUDY_NAME_CELL].value
    screening_model = _determine_screening_model(study_name, ws[Config.SCREENING_MODEL_CELL].value)
    study_code = _extract_study_code(ws[Config.STUDY_CODE_CELL].value, folder_name)
    
    # Extract lists of data (tissues, triggers with doses)
    tissues = _extract_unique_tissues(ws)
    triggers, doses = ExcelExtractor.extract_paired_columns(
        ws, Config.TRIGGERS_START_ROW, Config.TRIGGERS_COLUMN, Config.DOSES_COLUMN
    )
    
    # Create mapping between triggers and their corresponding doses
    trigger_dose_map = _create_trigger_dose_map(triggers, doses)
    
    # Extract timepoint information
    timepoint = _extract_timepoint(ws)
    
    return {
        "study_name": study_name,
        "study_code": study_code,
        "screening_model": screening_model,
        "tissues": tissues,
        "trigger_dose_map": trigger_dose_map,
        "timepoint": timepoint,
    }

def _extract_fallback_metadata(folder_name: str) -> Dict[str, Any]:
    """
    Extract minimal metadata when the procedure sheet is not available.
    Attempts to get study code from folder name using regex pattern.
    """
    study_code = None
    match = re.match(r"(\d{10})", folder_name)  # Look for 10-digit study code
    if match:
        study_code = match.group(1)
    
    return {
        "study_name": None,
        "study_code": study_code,
        "screening_model": None,
        "tissues": [],
        "trigger_dose_map": {},
        "timepoint": None,
    }

def _determine_screening_model(study_name: str, model_cell_value: str) -> str:
    """
    Determine screening model based on study name or specific cell value.
    If study name contains 'aav', automatically set to 'AAV'.
    """
    if study_name and "aav" in str(study_name).lower():
        return "AAV"
    return model_cell_value

def _extract_study_code(cell_value: Any, folder_name: str) -> Optional[str]:
    """
    Extract 10-digit study code from cell value or folder name.
    Validates that the code is exactly 10 digits.
    """
    if cell_value and re.fullmatch(r"\d{10}", str(cell_value)):
        return str(cell_value)
    
    # Fallback to folder name
    match = re.match(r"(\d{10})", folder_name)
    return match.group(1) if match else None

def _extract_unique_tissues(ws) -> List[str]:
    """
    Extract unique tissue types from the tissues column.
    Removes duplicates while preserving order.
    """
    tissues = ExcelExtractor.extract_column_values(
        ws, Config.TISSUES_START_ROW, Config.TISSUES_COLUMN
    )
    return list(dict.fromkeys(tissues))  # Remove duplicates, preserve order

def _create_trigger_dose_map(triggers: List[str], doses: List[Any]) -> Dict[str, Any]:
    """
    Create a dictionary mapping each trigger to its corresponding dose.
    Ensures both lists are the same length by padding doses with None.
    """
    # Ensure lists are same length
    while len(doses) < len(triggers):
        doses.append(None)
    return {str(trigger): dose for trigger, dose in zip(triggers, doses[:len(triggers)])}

def _extract_timepoint(ws) -> Optional[str]:
    """
    Extract timepoint from the worksheet by searching for 'day' or 'timepoint' headers.
    Finds the last non-empty value in that column.
    """
    # Search for timepoint header cell
    timepoint_location = ExcelExtractor.find_cell_with_text(ws, "day")
    if not timepoint_location:
        timepoint_location = ExcelExtractor.find_cell_with_text(ws, "timepoint")
    
    if not timepoint_location:
        return None
    
    header_row, header_col = timepoint_location
    
    # Find the last non-empty value in that column
    last_val = None
    for row in range(header_row + 1, ws.max_row + 1):
        val = ws.cell(row=row, column=header_col).value
        if val is not None and str(val).strip():
            last_val = val
    
    return format_timepoint(str(last_val).strip()) if last_val else None

# ========================== RELATIVE EXPRESSION DATA EXTRACTION ==========================
def extract_relative_expression_data(wb) -> Optional[Dict[str, Any]]:
    """
    Extract relative expression data from the results sheet.
    This is the main data we're interested in - target genes vs triggers with expression values.
    
    The data structure looks like:
    - Row with target names (F, J, N, etc.)
    - Rows with trigger names in column B, and corresponding data in columns G,H,I then K,L,M etc.
    
    Returns:
        Dictionary with 'targets' list and 'relative_expression_data' nested dict
    """
    sheet_name = _find_relative_expression_sheet(wb)
    if not sheet_name:
        return None
    
    ws = wb[sheet_name]
    print(f"Using sheet: '{sheet_name}'")
    
    # Find the "Relative Expression by Groups" section header
    rel_exp_location = ExcelExtractor.find_cell_with_text(
        ws, "relative expression", Config.REL_EXP_SEARCH_ROWS
    )
    if not rel_exp_location:
        print(f"Relative Expression section not found in sheet {sheet_name}")
        return None
    
    rel_exp_row, _ = rel_exp_location
    
    # Extract target gene names (usually 2 rows below the header)
    target_row = rel_exp_row + 2
    targets, target_columns = ExcelExtractor.extract_targets_from_row(ws, target_row)
    
    if not targets:
        print(f"No targets found in row {target_row}")
        return None
    
    # Extract trigger names (usually 3 rows below the header, in column B)
    trigger_start_row = target_row + 3
    triggers = ExcelExtractor.extract_column_values(
        ws, trigger_start_row, "B", stop_on_empty=False
    )[:Config.MAX_TRIGGERS]
    
    print(f"Found targets: {targets}")
    print(f"Found triggers: {triggers}")
    
    # Extract the actual expression data for each trigger-target combination
    triggers_data = _extract_trigger_target_data(ws, triggers, targets, target_columns, trigger_start_row)
    
    # Remove triggers with no data
    clean_triggers_data = {k: v for k, v in triggers_data.items() if v}
    
    return {
        "targets": targets,
        "relative_expression_data": clean_triggers_data
    }

def _find_relative_expression_sheet(wb) -> Optional[str]:
    """
    Find the sheet containing relative expression data from the list of possible names.
    Tries exact matches first, then case-insensitive keyword matching.
    """
    # Try exact matches first
    for sheet_name in Config.RELATIVE_EXPRESSION_SHEETS:
        if sheet_name in wb.sheetnames:
            return sheet_name
    
    # Try case-insensitive and keyword matching
    for sheet in wb.sheetnames:
        sheet_lower = sheet.lower()
        if (("compiled" in sheet_lower and ("indiv" in sheet_lower or "grp" in sheet_lower)) or
            ("calcs" in sheet_lower and "norm" in sheet_lower and "ctrl" in sheet_lower)):
            return sheet
    
    return None

def _extract_trigger_target_data(ws, triggers: List[str], targets: List[str], 
                                target_columns: List[int], trigger_start_row: int) -> Dict[str, Dict[str, Dict[str, Any]]]:
    """
    Extract the actual expression values for each trigger-target combination.
    
    For each target at column X, the data is in columns X+1, X+2, X+3:
    - X+1: relative expression value
    - X+2: low confidence interval  
    - X+3: high confidence interval
    
    Args:
        ws: Excel worksheet
        triggers: List of trigger names
        targets: List of target gene names
        target_columns: List of column numbers where targets are located
        trigger_start_row: Row where trigger data begins
        
    Returns:
        Nested dictionary: {trigger: {target: {rel_exp, low, high}}}
    """
    triggers_data = {}
    
    for trigger_idx, trigger in enumerate(triggers):
        if is_empty_or_zero(trigger):
            continue
            
        trigger_row = trigger_start_row + trigger_idx
        triggers_data[trigger] = {}
        
        for target_idx, target in enumerate(targets):
            base_col = target_columns[target_idx]
            
            # Extract the three data values for this trigger-target combination
            values = {
                "rel_exp": ws.cell(row=trigger_row, column=base_col + 1).value,  # Relative expression
                "low": ws.cell(row=trigger_row, column=base_col + 2).value,      # Low CI
                "high": ws.cell(row=trigger_row, column=base_col + 3).value      # High CI
            }
            
            # Skip combinations with no data
            if all(v is None for v in values.values()):
                continue
            
            triggers_data[trigger][target] = values
            
            debug_print(f"  {trigger} + {target}: {values}")
    
    return triggers_data

# ========================== STRING MATCHING UTILITIES ==========================
class StringMatcher:
    """
    Utility class for finding the best match between strings.
    Used to match trigger names between metadata and results sheets
    when they might have slight variations.
    """
    
    @staticmethod
    def find_best_match(target: str, candidates: List[str]) -> Optional[str]:
        """
        Find the best matching string using multiple strategies:
        1. Exact match (case-insensitive)
        2. Prefix match (target is at start of candidate)
        3. Normalized match (alphanumeric characters only)
        
        Args:
            target: String to find a match for
            candidates: List of possible matches
            
        Returns:
            Best matching candidate string, or None if no match found
        """
        if not target or not candidates:
            return None
        
        target_clean = target.lower().strip()
        
        # Strategy 1: Exact match
        for candidate in candidates:
            if candidate.lower().strip() == target_clean:
                return candidate
        
        # Strategy 2: Prefix match
        for candidate in candidates:
            if candidate.lower().strip().startswith(target_clean):
                return candidate
        
        # Strategy 3: Normalized match (alphanumeric only)
        target_norm = normalize_string(target)
        for candidate in candidates:
            if normalize_string(candidate).startswith(target_norm):
                return candidate
        
        return None

# ========================== STUDY PROCESSING ==========================
def process_study_folder(study_folder: str) -> Optional[Dict[str, Any]]:
    """
    Process a single study folder and extract all available data.
    
    Each study folder should contain:
    - {folder_name}.xlsm (metadata file)
    - Results/{results_file}.xlsm (results file)
    
    Args:
        study_folder: Path to the study folder
        
    Returns:
        Dictionary containing all extracted data, or None if no data found
    """
    folder_name = os.path.basename(study_folder)
    info_file = os.path.join(study_folder, f"{folder_name}.xlsm")
    results_folder = os.path.join(study_folder, "Results")
    
    print(f"\nProcessing study: {folder_name}")
    
    study_data = {}
    
    # Extract metadata from the main study file
    if os.path.exists(info_file):
        metadata = safe_workbook_operation(info_file, extract_study_metadata, folder_name)
        if metadata:
            study_data.update(metadata)
            print("Extracted metadata fields:")
            for k, v in metadata.items():
                print(f"  {k}: {v}")
    else:
        print(f"Info file not found: {info_file}")
    
    # Extract data from the results file
    results_file = _find_results_file(results_folder)
    if results_file:
        # Extract LAR sheet data (additional metadata)
        lar_data = _extract_lar_data(results_file)
        if lar_data:
            study_data["lar_data"] = lar_data
        
        # Extract the main relative expression data
        rel_exp_data = safe_workbook_operation(
            results_file, extract_relative_expression_data, read_only=False
        )
        if rel_exp_data:
            study_data["relative_expression"] = rel_exp_data
            print(f"Extracted relative expression data: {len(rel_exp_data['targets'])} targets, "
                  f"{len(rel_exp_data['relative_expression_data'])} triggers")
    
    return study_data if study_data else None

def _find_results_file(results_folder: str) -> Optional[str]:
    """
    Find the first .xlsm file in the Results folder.
    There should typically be only one results file per study.
    """
    if not os.path.exists(results_folder):
        return None
    
    for file in os.listdir(results_folder):
        if file.endswith(".xlsm"):
            return os.path.join(results_folder, file)
    
    return None

def _extract_lar_data(results_file: str) -> Optional[Dict[str, str]]:
    """
    Extract additional metadata from the LAR Sheet.
    Searches for keyword-value pairs in the first 20 rows.
    """
    try:
        if Config.LAR_SHEET not in pd.ExcelFile(results_file).sheet_names:
            return None
        
        df = pd.read_excel(results_file, sheet_name=Config.LAR_SHEET, header=None)
        
        fields = {}
        for i, row in df.iterrows():
            if i >= 20:  # Limit search to first 20 rows
                break
            for col in range(len(row)):
                cell = str(row[col]).strip().lower()
                # Look for cells containing these keywords
                if any(keyword in cell for keyword in ["trigger", "dose", "tissue", "timepoint"]):
                    field_name = next(k for k in ["trigger", "dose", "tissue", "timepoint"] if k in cell)
                    field_value = str(row[col + 1]).strip() if col + 1 < len(row) else ""
                    if field_value != "nan":
                        fields[field_name] = field_value
        
        return fields if fields else None
    except Exception as e:
        print(f"Error extracting LAR data: {e}")
        return None

# ========================== CSV EXPORT ==========================
def export_to_csv(all_study_data: List[Dict[str, Any]], output_path: str):
    """
    Export all study data to a CSV file in a standardized format.
    
    The CSV will have one row per trigger-target combination with columns:
    - study_name, study_code, screening_model, gene_target, trigger, dose,
      timepoint, tissue, avg_rel_exp, avg_rel_exp_lsd, avg_rel_exp_hsd
    """
    header = [
        "study_name", "study_code", "screening_model", "gene_target", "trigger", 
        "dose", "timepoint", "tissue", "avg_rel_exp", "avg_rel_exp_lsd", "avg_rel_exp_hsd"
    ]
    
    csv_rows = [header]
    stats = {"studies_processed": 0, "studies_with_data": 0, "total_rows": 0}
    
    for study in all_study_data:
        stats["studies_processed"] += 1
        rows_added = _process_study_for_csv(study, csv_rows)
        if rows_added > 0:
            stats["studies_with_data"] += 1
            stats["total_rows"] += rows_added
    
    # Write CSV file
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv.writer(csvfile).writerows(csv_rows)
    
    _print_export_summary(stats, output_path)

def _process_study_for_csv(study: Dict[str, Any], csv_rows: List[List[str]]) -> int:
    """
    Process a single study for CSV export.
    Matches triggers from metadata with triggers from results data,
    then creates one CSV row per trigger-target combination.
    """
    if "relative_expression" not in study:
        return 0
    
    study_info = _extract_study_info_for_csv(study)
    rel_exp_data = study["relative_expression"]["relative_expression_data"]
    
    rows_added = 0
    # Process each trigger from the metadata
    for trigger, dose in study_info["trigger_dose_map"].items():
        # Find the matching trigger in the results data (may have slight name variations)
        matching_trigger = StringMatcher.find_best_match(trigger, list(rel_exp_data.keys()))
        
        if not matching_trigger:
            continue
        
        # For each target that has data for this trigger
        trigger_data = rel_exp_data[matching_trigger]
        for target, values in trigger_data.items():
            # Skip if no actual values
            if all(not values.get(key) for key in ['rel_exp', 'low', 'high']):
                continue
            
            # Create a CSV row for this trigger-target combination
            row = _create_csv_row(study_info, trigger, dose, target, values)
            csv_rows.append(row)
            rows_added += 1
    
    return rows_added

def _extract_study_info_for_csv(study: Dict[str, Any]) -> Dict[str, Any]:
    """
    Extract and format study information for CSV export.
    Handles formatting of timepoint and tissue data.
    """
    timepoint = study.get("timepoint", "")
    if timepoint and not timepoint.startswith('D') and timepoint.strip().isdigit():
        timepoint = f"D{timepoint.strip()}"
    
    # Use tissue from metadata, fall back to LAR data
    tissue = ""
    if study.get("tissues"):
        tissue = study["tissues"][0]
    elif "lar_data" in study and "tissue" in study["lar_data"]:
        tissue = study["lar_data"]["tissue"]
    
    return {
        "study_name": study.get("study_name", ""),
        "study_code": f"'{study.get('study_code', '')}'" if study.get("study_code") else "",  # Prevent scientific notation
        "screening_model": study.get("screening_model", ""),
        "trigger_dose_map": study.get("trigger_dose_map", {}),
        "timepoint": timepoint,
        "tissue": tissue
    }

def _create_csv_row(study_info: Dict[str, Any], trigger: str, dose: Any, 
                   target: str, values: Dict[str, Any]) -> List[str]:
    """
    Create a single CSV row for one trigger-target combination.
    """
    return [
        study_info["study_name"],
        study_info["study_code"],
        study_info["screening_model"],
        target,                                    # gene_target
        trigger,                                   # trigger (from metadata)
        dose,                                      # dose (from metadata)
        study_info["timepoint"],
        study_info["tissue"],
        convert_to_numeric(values.get("rel_exp")),  # avg_rel_exp
        convert_to_numeric(values.get("low")),      # avg_rel_exp_lsd  
        convert_to_numeric(values.get("high"))      # avg_rel_exp_hsd
    ]

def _print_export_summary(stats: Dict[str, int], output_path: str):
    """Print summary statistics after CSV export."""
    print(f"\nExport Summary:")
    print(f"- Total studies processed: {stats['studies_processed']}")
    print(f"- Studies with data: {stats['studies_with_data']}")
    print(f"- Total data rows: {stats['total_rows']}")
    print(f"- Output file: {output_path}")

# ========================== MAIN EXECUTION ==========================
def main():
    """
    Main execution function.
    
    Process flow:
    1. Find all study folders in the month directory
    2. Process each folder to extract metadata and results data
    3. Export all data to JSON (raw) and CSV (formatted) files
    """
    # Find all subdirectories in the month folder (each is a study)
    study_folders = [
        os.path.join(Config.MONTH_FOLDER, name)
        for name in os.listdir(Config.MONTH_FOLDER)
        if os.path.isdir(os.path.join(Config.MONTH_FOLDER, name))
    ]
    
    if not study_folders:
        print("No studies found in the month folder.")
        return
    
    print(f"Processing {len(study_folders)} study folders")
    
    # Process all studies
    all_study_data = []
    for study_folder in study_folders:
        study_data = process_study_folder(study_folder)
        if study_data:
            all_study_data.append(study_data)
    
    # Generate output files with timestamp
    timestamp = datetime.now().strftime("%Y%m%d")
    base_output_dir = os.path.dirname(Config.MONTH_FOLDER)
    month_name = os.path.basename(Config.MONTH_FOLDER).split(' ')[0]
    
    # Export raw data to JSON
    json_output_path = os.path.join(base_output_dir, f"study_metadata_{month_name}_{timestamp}.json")
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {json_output_path}")
    
    # Export formatted data to CSV
    csv_output_path = os.path.join(base_output_dir, f"study_data_{month_name}_{timestamp}.csv")
    export_to_csv(all_study_data, csv_output_path)

if __name__ == "__main__":
    main()