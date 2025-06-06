import pandas as pd
import os
import re
import json
import csv
from datetime import datetime
import traceback
from typing import Dict, List, Optional, Tuple, Any
from openpyxl import load_workbook

class Config:
    """
    Central configuration class containing all constants and settings.
    This makes it easy to modify paths, cell locations, and parameters without 
    hunting through the entire codebase.
    """
    # Main folder containing all study subfolders to process
    MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    DEBUG = False  # Set to True to see detailed debug output during processing
    MAX_STUDIES = 10  # LIMIT TO 10 STUDIES FOR TESTING
    
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
    
    # Known dose types to search for
    DOSE_TYPES = {'SQ', 'IV', 'IM', 'intratracheal', 'subcutaneous', 'intravenous', 'intramuscular'}
    
    # Known tissue types to differentiate from gene targets
    TISSUE_TYPES = {
        'lung', 'cerebellum', 'cortex', 'eye', 'liver', 'heart', 'kidney', 'spleen',
        'brain', 'muscle', 'skin', 'blood', 'plasma', 'serum', 'bone', 'fat',
        'adipose', 'pancreas', 'stomach', 'intestine', 'colon', 'bladder',
        'prostate', 'ovary', 'uterus', 'hippocampus', 'cerebral cortex',
        'frontal cortex', 'motor cortex', 'brainstem', 'midbrain', 'spinal cord',
        'thoracic spinal cord', 'lumbar spinal cord', 'cervical spinal cord',
        'skeletal muscle', 'cardiac muscle', 'diaphragm', 'quadriceps',
        'gastrocnemius', 'tibialis anterior', 'retina', 'optic nerve'
    }

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

def detect_dose_type(text: str) -> Optional[str]:
    """
    Detect dose type (SQ, IV, IM, intratracheal) from text.
    
    Args:
        text: Text to search for dose type indicators
        
    Returns:
        Detected dose type or None if not found
    """
    if not text:
        return None
    
    text_lower = str(text).lower().strip()
    
    # Check for exact matches first
    for dose_type in Config.DOSE_TYPES:
        if dose_type.lower() in text_lower:
            # Return standardized format
            if dose_type.lower() in ['sq', 'subcutaneous']:
                return 'SQ'
            elif dose_type.lower() in ['iv', 'intravenous']:
                return 'IV'
            elif dose_type.lower() in ['im', 'intramuscular']:
                return 'IM'
            elif dose_type.lower() == 'intratracheal':
                return 'Intratracheal'
    
    return None

def is_tissue_name(text: str) -> bool:
    """
    Check if a string represents a tissue name rather than a gene target.
    Only returns True for exact (case-insensitive, stripped) matches to known tissue types.
    """
    if not text:
        return False
    text_lower = str(text).lower().strip()
    # Only allow exact matches
    return text_lower in Config.TISSUE_TYPES


def classify_target_or_tissue(text: str, procedure_tissues: List[str]) -> Tuple[str, Optional[str], Optional[str]]:
    """
    Classify a text string as either a tissue or gene target.
    Only classifies as tissue on exact match to procedure tissues or known tissue types.
    """
    if not text:
        return 'target', None, None
    text_clean = str(text).strip()
    # Only exact match to procedure tissues
    for proc_tissue in procedure_tissues:
        if normalize_string(text_clean) == normalize_string(proc_tissue):
            return 'tissue', text_clean, None
    # Only exact match to known tissue types
    if is_tissue_name(text_clean):
        return 'tissue', text_clean, None
    # Otherwise, treat as gene target
    return 'target', None, text_clean


def safe_workbook_operation(file_path: str, operation_func, *args, **kwargs):
    """
    Safely open Excel workbooks and ensure they're properly closed.
    Always uses read_only mode unless explicitly disabled.
    """
    try:
        # Default to read_only=True unless explicitly set to False
        read_only = kwargs.pop('read_only', True)
        wb = load_workbook(file_path, data_only=True, read_only=read_only)
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
    def extract_column_values(ws, start_row: int, column: str, stop_on_empty: bool = True, max_empty_cells: int = 1) -> List[str]:
        """
        Extract all values from a single column starting at a specific row.
        
        Args:
            ws: Excel worksheet object
            start_row: Row number to start extraction (1-indexed)
            column: Column letter (e.g., 'B', 'S')
            stop_on_empty: Whether to stop when hitting empty cells
            max_empty_cells: Number of consecutive empty cells to tolerate before stopping
            
        Returns:
            List of extracted string values
        """
        values = []
        row = start_row
        empty_count = 0
        
        while row <= ws.max_row:
            cell_value = ws[f"{column}{row}"].value
            
            if is_empty_or_zero(cell_value):
                empty_count += 1
                if stop_on_empty and empty_count >= max_empty_cells:
                    break
            else:
                # Reset empty counter when we find a value
                empty_count = 0
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
    def extract_targets_from_row(ws, row: int, procedure_tissues: List[str] = None) -> Tuple[List[str], List[int], List[str]]:
        """
        Extract target gene names and their column positions from a specific row.
        Also identifies any tissues found in the target row.
        Targets are typically spaced every 4 columns (F, J, N, R, etc.).
        
        Args:
            ws: Excel worksheet object
            row: Row number containing target names
            procedure_tissues: List of tissues from Procedure Request Form for comparison
            
        Returns:
            Tuple of (target_names, column_numbers, found_tissues)
        """
        if procedure_tissues is None:
            procedure_tissues = []
            
        targets, target_columns, found_tissues = [], [], []
        col_start = Config.TARGET_START_COLUMN  # Start at column F (6)
        zero_count = 0
        
        while len(targets) + len(found_tissues) < Config.MAX_TARGETS:
            cell_value = ws.cell(row=row, column=col_start).value
            
            if is_empty_or_zero(cell_value):
                zero_count += 1
                if zero_count >= 5:  # Stop after 5 consecutive empty cells
                    break
            else:
                zero_count = 0
                text_clean = str(cell_value).strip()
                
                # Classify as tissue or target
                classification, tissue_name, target_name = classify_target_or_tissue(
                    text_clean, procedure_tissues
                )
                print(f"Target row cell '{text_clean}': classified as {classification}")
                if classification == 'tissue':
                    found_tissues.append(tissue_name)
                else:
                    targets.append(target_name)
                    target_columns.append(col_start)
            
            col_start += Config.TARGET_COLUMN_SPACING  # Move to next target column
        
        return targets, target_columns, found_tissues

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

def _create_trigger_dose_map(triggers: List[str], doses: List[Any]) -> Dict[str, Dict[str, Any]]:
    """
    Create a dictionary mapping each trigger to its corresponding dose and detected dose type.
    Ensures both lists are the same length by padding doses with None.
    Now also detects dose types (SQ, IV, IM, intratracheal).
    
    Returns:
        Dictionary mapping trigger -> {"dose": dose_value, "dose_type": detected_type}
    """
    # Ensure lists are same length
    while len(doses) < len(triggers):
        doses.append(None)
    
    trigger_dose_map = {}
    for trigger, dose in zip(triggers, doses[:len(triggers)]):
        dose_str = str(dose) if dose is not None else ""
        detected_dose_type = detect_dose_type(dose_str)
        
        trigger_dose_map[str(trigger)] = {
            "dose": dose,
            "dose_type": detected_dose_type
        }
        
        if detected_dose_type:
            debug_print(f"Detected dose type '{detected_dose_type}' for trigger '{trigger}': {dose_str}")
    
    return trigger_dose_map

def _extract_timepoint(ws) -> Optional[str]:
    """
    Extract the timepoint value using a systematic approach:
    1. First check known locations for specific study types
    2. Scan column A for timepoint headers or patterns
    3. When a timepoint header is found, search downward to find the LAST value
    4. The last value in a timepoint sequence is the actual timepoint we want
    """
    print("\nExtracting timepoint from worksheet...")
    
    # STRATEGY 1: Check specific known locations for certain study types
    known_locations = [
        (24, 1),   # A24 (common location)
        (16, 4),   # D16 for rIL33_8_Alternaria
        (23, 4),   # D23 for rHDM_Pilot_2
        (55, 1),   # A55 for D154
        (145, 4),  # D145 for mTSHR_7
        (154, 1),   # A154 possibly containing timepoint
        (24, 4)    # D24 possibly containing timepoint
    ]
    
    for row, col in known_locations:
        try:
            cell_value = ws.cell(row=row, column=col).value
            if cell_value:
                cell_text = str(cell_value).lower()
                match = re.search(r'd(\d+)', cell_text)
                if match and match.group(1):
                    print(f"Found direct timepoint in cell ({row},{col}): D{match.group(1)}")
                    # Continue checking other cells to find the LAST value
                    continue
        except:
            continue
    
    # STRATEGY 2: Find any timepoint header in column A
    timepoint_headers = []
    for row in range(1, min(ws.max_row + 1, 100)):
        try:
            cell_value = ws.cell(row=row, column=1).value
            if not cell_value:
                continue
                
            cell_text = str(cell_value).lower()
            # Check for common timepoint header patterns
            if any(keyword in cell_text for keyword in ['day', 'timepoint', 'sacrifice', 'necropsy']):
                print(f"Found timepoint header in A{row}: {cell_value}")
                timepoint_headers.append(row)
        except:
            continue
    
    # STRATEGY 3: For each header, find the LAST value in that section
    timepoint_values = []
    
    for header_row in timepoint_headers:
        last_value = None
        last_row = None
        empty_count = 0
        
        # Search up to 30 rows below the header
        for row in range(header_row + 1, min(header_row + 30, ws.max_row + 1)):
            try:
                cell_value = ws.cell(row=row, column=1).value
                
                if is_empty_or_zero(cell_value):
                    empty_count += 1
                    if empty_count >= 3:  # Stop after 3 consecutive empty cells
                        break
                else:
                    empty_count = 0
                    cell_text = str(cell_value).strip().lower()
                    
                    # Skip rows that look like headers
                    if any(keyword in cell_text for keyword in ['day', 'timepoint', 'date']):
                        continue
                        
                    # Look for day patterns (D16, D23, etc.)
                    d_match = re.search(r'd(\d+)', cell_text)
                    num_match = re.search(r'^(\d+)$', cell_text)
                    
                    if d_match:
                        last_value = f"D{d_match.group(1)}"
                        last_row = row
                        print(f"Found timepoint in A{row}: {last_value}")
                    elif num_match and not any(skip in cell_text for skip in ['/', '-']):  # Avoid dates
                        last_value = f"D{num_match.group(1)}"
                        last_row = row
                        print(f"Found numeric timepoint in A{row}: {last_value}")
            except:
                continue
        
        if last_value:
            print(f"Final timepoint in section A{last_row}: {last_value}")
            timepoint_values.append((last_row, last_value))
    
    # If we found multiple timepoint values, use the one with highest row number
    # This ensures we get the LAST value in the sheet
    if timepoint_values:
        timepoint_values.sort(key=lambda x: x[0], reverse=True)
        print(f"Selected final timepoint: {timepoint_values[0][1]} (from row A{timepoint_values[0][0]})")
        return timepoint_values[0][1]
    
    # STRATEGY 4: Brute force scan for any D## pattern
    # This is our last resort if the other strategies fail
    d_pattern_rows = []
    
    for row in range(1, min(ws.max_row + 1, 200)):
        try:
            cell_value = ws.cell(row=row, column=1).value
            if cell_value:
                cell_text = str(cell_value).lower()
                match = re.search(r'd(\d+)', cell_text)
                if match:
                    day_num = int(match.group(1))
                    if 1 <= day_num <= 365:  # Valid day range
                        d_pattern_rows.append((row, f"D{day_num}"))
                        print(f"Found D-pattern in A{row}: D{day_num}")
        except:
            continue
    
    # Get the last D-pattern in the sheet
    if d_pattern_rows:
        d_pattern_rows.sort(key=lambda x: x[0], reverse=True)
        print(f"Selected D-pattern: {d_pattern_rows[0][1]} (from row A{d_pattern_rows[0][0]})")
        return d_pattern_rows[0][1]
    
    # STRATEGY 5: Special case for rIL33_8_Alternaria and similar
    # Known to have timepoint info in specific cells
    special_case_cells = [
        (16, 4), (17, 4), (18, 4),  # D column near row 16
        (23, 4), (24, 4), (25, 4),  # D column near row 23
        (145, 1), (146, 1), (147, 1)  # A column near row 145
    ]
    
    for row, col in special_case_cells:
        try:
            cell_value = ws.cell(row=row, column=col).value
            if cell_value:
                cell_text = str(cell_value).lower()
                # Very specific pattern matching for these cells
                match = re.search(r'd\s*(\d+)', cell_text)
                if match:
                    timepoint = f"D{match.group(1)}"
                    print(f"Found special case timepoint in ({row},{col}): {timepoint}")
                    return timepoint
        except:
            continue
    
    print("No valid timepoint found after exhaustive search")
    return None

# ========================== RELATIVE EXPRESSION DATA EXTRACTION ==========================
def extract_relative_expression_data(wb, procedure_tissues: List[str] = None) -> Optional[Dict[str, Any]]:
    """
    Extract relative expression data from the results sheet.
    This is the main data we're interested in - target genes vs triggers with expression values.
    
    The data structure looks like:
    - Row with target names (F, J, N, etc.)
    - Rows with trigger names in column B, and corresponding data in columns G,H,I then K,L,M etc.
    
    Args:
        wb: Excel workbook object
        procedure_tissues: List of tissues from Procedure Request Form for comparison
    
    Returns:
        Dictionary with 'targets' list, 'relative_expression_data' nested dict, and 'found_tissues'
    """
    if procedure_tissues is None:
        procedure_tissues = []
        
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
    targets, target_columns, found_tissues = ExcelExtractor.extract_targets_from_row(
        ws, target_row, procedure_tissues
    )
    
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
    if found_tissues:
        print(f"Found tissues in target row: {found_tissues}")
    
    # Extract the actual expression data for each trigger-target combination
    triggers_data = _extract_trigger_target_data(ws, triggers, targets, target_columns, trigger_start_row)
    
    # Remove triggers with no data
    clean_triggers_data = {k: v for k, v in triggers_data.items() if v}
    
    return {
        "targets": targets,
        "relative_expression_data": clean_triggers_data,
        "found_tissues": found_tissues
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
            print(f"Extracted LAR data: {lar_data}")
        
        # Extract the main relative expression data
        procedure_tissues = study_data.get("tissues", [])
        rel_exp_data = safe_workbook_operation(
            results_file, 
            extract_relative_expression_data,
            procedure_tissues
        )
        
        if rel_exp_data:
            study_data["relative_expression"] = rel_exp_data
            print(f"Extracted relative expression data:")
            print(f"  - Targets: {len(rel_exp_data.get('targets', []))} targets")
            print(f"  - Triggers: {len(rel_exp_data.get('relative_expression_data', {}))} triggers")
            # Print a sample of the data structure
            if rel_exp_data.get('relative_expression_data'):
                first_trigger = next(iter(rel_exp_data['relative_expression_data']))
                print(f"  - Sample data structure for trigger '{first_trigger}':")
                if rel_exp_data['relative_expression_data'][first_trigger]:
                    first_target = next(iter(rel_exp_data['relative_expression_data'][first_trigger]))
                    print(f"    {first_target}: {rel_exp_data['relative_expression_data'][first_trigger][first_target]}")
                else:
                    print("    No target data found")
        else:
            print("  No relative expression data found")
    
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
    - study_name, study_code, screening_model, gene_target, trigger, dose, dose_type,
      timepoint, tissue, avg_rel_exp, avg_rel_exp_lsd, avg_rel_exp_hsd
    """
    header = [
        "study_name", "study_code", "screening_model", "gene_target", "trigger", 
        "dose", "dose_type", "timepoint", "tissue", "avg_rel_exp", "avg_rel_exp_lsd", "avg_rel_exp_hsd"
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
    Handles both older data format and new optimized format.
    
    Args:
        study: Study data dictionary
        csv_rows: List of CSV rows to append to
        
    Returns:
        Number of rows added to the CSV
    """
    # Skip studies with no relative expression data
    if "relative_expression" not in study:
        print(f"Skipping study with no relative expression data: {study.get('study_name')}")
        return 0
    
    print(f"\nProcessing CSV data for study: {study.get('study_name')}")
    study_info = _extract_study_info_for_csv(study)
    print(f"  Study info: {study_info}")
    
    rel_exp_data = study["relative_expression"].get("relative_expression_data", {})
    
    if not rel_exp_data:
        print(f"  No relative expression data found in study")
        return 0
    
    print(f"  Found {len(rel_exp_data)} triggers in relative expression data")
    print(f"  Trigger names in relative expression: {list(rel_exp_data.keys())}")
    print(f"  Trigger names in metadata: {list(study_info['trigger_dose_map'].keys())}")
    
    rows_added = 0
    
    # Process each trigger from results data, not just from metadata
    for trigger in rel_exp_data.keys():
        # Get trigger info from metadata if available, otherwise create empty info
        trigger_info = study_info["trigger_dose_map"].get(trigger, {"dose": "", "dose_type": ""})
        trigger_data = rel_exp_data[trigger]
        
        print(f"  Processing trigger: {trigger} with {len(trigger_data)} targets")
        
        # Process each target for this trigger
        for target, value in trigger_data.items():
            # Check if value is a dictionary (old format) or float (new format)
            if isinstance(value, dict):
                # Old format: {"rel_exp": val, "low": val, "high": val}
                row = _create_csv_row(
                    study_info, trigger, trigger_info, target, 
                    {"rel_exp": value.get("rel_exp"), "low": value.get("low"), "high": value.get("high")}
                )
            else:
                # New format: direct float value
                row = _create_csv_row(
                    study_info, trigger, trigger_info, target, 
                    {"rel_exp": value, "low": None, "high": None}
                )
            
            csv_rows.append(row)
            rows_added += 1
    
    print(f"  Added {rows_added} rows to CSV for study {study.get('study_name')}")
    return rows_added

def _extract_study_info_for_csv(study: Dict[str, Any]) -> Dict[str, Any]:
    """
    Extract and format study information for CSV export.
    """
    timepoint = study.get("timepoint", "")
    if timepoint and not timepoint.startswith('D') and timepoint.strip().isdigit():
        timepoint = f"D{timepoint.strip()}"
    
    tissue = ""
    if study.get("tissues"):
        tissue = study["tissues"][0]
    elif "lar_data" in study and "tissue" in study["lar_data"]:
        tissue = study["lar_data"]["tissue"]
    
    if not tissue and "relative_expression" in study:
        found_tissues = study["relative_expression"].get("found_tissues", [])
        if found_tissues:
            tissue = found_tissues[0]
    
    return {
        "study_name": study.get("study_name", ""),
        "study_code": f"'{study.get('study_code', '')}'" if study.get("study_code") else "",
        "screening_model": study.get("screening_model", ""),
        "trigger_dose_map": study.get("trigger_dose_map", {}),
        "timepoint": timepoint,
        "tissue": tissue
    }

def _create_csv_row(study_info: Dict[str, Any], trigger: str, trigger_info: Dict[str, Any], 
                   target: str, values: Dict[str, Any]) -> List[str]:
    """
    Create a single CSV row for one trigger-target combination.
    
    CSV columns (in order):
    1. study_name: Name of the study
    2. study_code: 10-digit code (quoted to prevent Excel issues)
    3. screening_model: Type of screening used
    4. gene_target: Name of the target gene
    5. trigger: Name of the trigger (e.g., siRNA sequence)
    6. dose: Dose amount for this trigger
    7. dose_type: Detected dose type (SQ, IV, IM, Intratracheal)
    8. timepoint: Time point of measurement (e.g., D14)
    9. tissue: Tissue type tested
    10. avg_rel_exp: Average relative expression value
    11. avg_rel_exp_lsd: Lower standard deviation
    12. avg_rel_exp_hsd: Higher standard deviation
    
    All numeric values are formatted to 4 decimal places.
    """
    # Handle the case where values is directly a float (new format)
    if not isinstance(values, dict):
        rel_exp = values
        low = None
        high = None
    else:
        rel_exp = values.get("rel_exp")
        low = values.get("low")
        high = values.get("high")
        
    return [
        study_info["study_name"],
        study_info["study_code"],
        study_info["screening_model"],
        target,                                # gene_target
        trigger,                               # trigger (from metadata)
        str(trigger_info.get("dose", "")),     # dose (from metadata)
        str(trigger_info.get("dose_type", "")), # dose_type (detected)
        study_info["timepoint"],
        study_info["tissue"],
        convert_to_numeric(rel_exp),           # avg_rel_exp 
        convert_to_numeric(low),               # avg_rel_exp_lsd
        convert_to_numeric(high)               # avg_rel_exp_hsd
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
    """
    study_folders = [
        os.path.join(Config.MONTH_FOLDER, name)
        for name in os.listdir(Config.MONTH_FOLDER)
        if os.path.isdir(os.path.join(Config.MONTH_FOLDER, name))
    ]
    
    if not study_folders:
        print("No studies found in the month folder.")
        return

    print(f"Processing {min(len(study_folders), Config.MAX_STUDIES)} study folders")
    
    all_study_data = []
    for study_folder in study_folders[:Config.MAX_STUDIES]:
        study_data = process_study_folder(study_folder)
        if study_data:
            all_study_data.append(study_data)

    timestamp = datetime.now().strftime("%Y%m%d")
    base_output_dir = os.path.dirname(Config.MONTH_FOLDER)
    month_name = os.path.basename(Config.MONTH_FOLDER).split(' ')[0]
    
    json_output_path = os.path.join(base_output_dir, f"study_metadata_{month_name}_{timestamp}.json")
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {json_output_path}")
    
    csv_output_path = os.path.join(base_output_dir, f"study_data_{month_name}_{timestamp}.csv")
    export_to_csv(all_study_data, csv_output_path)

if __name__ == "__main__":
    main()