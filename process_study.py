import pandas as pd
import os
import re
import json
import csv
from datetime import datetime
import traceback
import logging
from typing import Dict, List, Optional, Tuple, Any
from openpyxl import load_workbook

# ========================== LOGGING SETUP ==========================
def setup_logging(log_file_path: str) -> logging.Logger:
    """
    Set up comprehensive logging for the study processing script.
    Creates both file and console handlers with detailed formatting.
    """
    # Create logger
    logger = logging.getLogger('StudyProcessor')
    logger.setLevel(logging.DEBUG)
    
    # Clear any existing handlers
    logger.handlers.clear()
    
    # Create formatters
    detailed_formatter = logging.Formatter(
        '%(asctime)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    console_formatter = logging.Formatter(
        '%(levelname)s - %(message)s'
    )
    
    # File handler - logs everything
    file_handler = logging.FileHandler(log_file_path, mode='w', encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(detailed_formatter)
    logger.addHandler(file_handler)
    
    # Console handler - logs INFO and above
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    return logger

# Global logger instance
logger = None

def init_logger(log_file_path: str):
    """Initialize the global logger instance."""
    global logger
    logger = setup_logging(log_file_path)

class Config:
    """
    Central configuration class containing all constants and settings.
    This makes it easy to modify paths, cell locations, and parameters without 
    hunting through the entire codebase.
        """
    # Main folder containing all study subfolders to process
    MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    DEBUG = True  # Set to True to see detailed debug output during processing
    
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
        'gastrocnemius', 'tibialis anterior', 'retina', 'optic nerve','adrenal gland',
        'sciatic nerve','striatum','lymph node','diaphragm', 'kidney cortex','gastroc','triceps',
        'apex','left ventricle','right ventricle','left atrium','right atrium','medial lobe', 'aorta',
        'rlung','llung','macrophage','tri','gast','gst','hrt','iWAT','thalamus','TSC','Cer','Ctx','Left Lateral Lobe',
        'Right Lateral Lobe', 'Medial Lobe'
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
    # Also log to file if logger is available
    if logger:
        message = ' '.join(str(arg) for arg in args)
        logger.debug(message)

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

def extract_dose_from_trigger_name(trigger_name: str) -> Optional[str]:
    """
    Extract dose value from trigger name, specifically looking for mpk values.
    
    Args:
        trigger_name: The trigger name to search for dose information
        
    Returns:
        Extracted dose value or None if not found
    """
    if not trigger_name:
        return None
    
    text = str(trigger_name).lower().strip()
    
    # Look for mpk patterns (e.g., "5mpk", "10 mpk", "2.5mpk")
    mpk_patterns = [
        r'(\d+(?:\.\d+)?)\s*mpk',  # Matches "5mpk", "10 mpk", "2.5mpk"
        r'(\d+(?:\.\d+)?)\s*mg/kg',  # Matches "5mg/kg", "10 mg/kg"
    ]
    
    for pattern in mpk_patterns:
        match = re.search(pattern, text)
        if match:
            dose_value = match.group(1)
            if logger:
                logger.debug(f"Extracted dose '{dose_value} mpk' from trigger name: '{trigger_name}'")
            return f"{dose_value} mpk"
    
    # Look for other dose patterns (ug, mg, etc.)
    dose_patterns = [
        r'(\d+(?:\.\d+)?)\s*(ug|μg|mg|g)\b',  # Matches "250ug", "5mg", etc.
        r'(\d+(?:\.\d+)?)\s*(ul|μl|ml|l)\b',  # Matches "250ul", "5ml", etc.
    ]
    
    for pattern in dose_patterns:
        match = re.search(pattern, text)
        if match:
            dose_value = match.group(1)
            dose_unit = match.group(2)
            extracted_dose = f"{dose_value} {dose_unit}"
            if logger:
                logger.debug(f"Extracted dose '{extracted_dose}' from trigger name: '{trigger_name}'")
            return extracted_dose
    
    return None

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
    Simple rule: If it's a tissue, keep it as tissue. If it's a target, keep it as target.
    No more complex splitting - one item = one classification.
    """
    if not text:
        return 'target', None, None
    text_clean = str(text).strip()
    
    # First check if the entire string is a tissue (exact match to procedure tissues)
    for proc_tissue in procedure_tissues:
        if normalize_string(text_clean) == normalize_string(proc_tissue):
            return 'tissue', text_clean, None
    
    # Check if the entire string is a known tissue type
    if is_tissue_name(text_clean):
        return 'tissue', text_clean, None
    
    # Check if any words in the text are tissue names (but keep the whole thing as tissue)
    words = text_clean.split()
    if len(words) > 1:  # Only check multi-word strings
        # Check for multi-word tissue names first
        all_tissue_types = list(Config.TISSUE_TYPES) + procedure_tissues
        sorted_tissues = sorted(all_tissue_types, key=len, reverse=True)
        
        for tissue_type in sorted_tissues:
            if tissue_type.lower() in text_clean.lower():
                if Config.DEBUG:
                    print(f"    Found tissue phrase '{tissue_type}' in '{text_clean}' - treating whole thing as tissue")
                return 'tissue', text_clean, None
        
        # Check individual words for tissue matches
        for word in words:
            word_clean = word.strip()
            if is_tissue_name(word_clean):
                if Config.DEBUG:
                    print(f"    Found tissue word '{word_clean}' in '{text_clean}' - treating whole thing as tissue")
                return 'tissue', text_clean, None
            
            # Also check against procedure tissues
            for proc_tissue in procedure_tissues:
                if normalize_string(word_clean) == normalize_string(proc_tissue):
                    if Config.DEBUG:
                        print(f"    Found procedure tissue word '{word_clean}' in '{text_clean}' - treating whole thing as tissue")
                    return 'tissue', text_clean, None
    
    # Otherwise, treat as gene target
    if Config.DEBUG:
        print(f"    Result: TARGET - '{text_clean}'")
    return 'target', None, text_clean


def safe_workbook_operation(file_path: str, operation_func, *args, **kwargs):
    """
    Safely open Excel workbooks and ensure they're properly closed.
    Always uses read_only mode unless explicitly disabled.
    ALWAYS uses data_only=True to evaluate formulas properly.
    """
    try:
        # Default to read_only=True unless explicitly set to False
        read_only = kwargs.pop('read_only', True)
        
        # Always use data_only=True to evaluate formulas
        # This is critical for trigger extraction since column B contains formulas
        if logger:
            logger.debug(f"Opening workbook: {file_path} (read_only={read_only}, data_only=True)")
        
        wb = load_workbook(file_path, data_only=True, read_only=read_only)
        result = operation_func(wb, *args)
        wb.close()
        
        if logger:
            logger.debug(f"Successfully processed workbook: {file_path}")
        
        return result
                
    except Exception as e:
        error_msg = f"Error processing {file_path}: {e}"
        print(error_msg)
        if logger:
            logger.error(error_msg)
            logger.error(f"Full traceback: {traceback.format_exc()}")
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
            # Get cell and handle potential Excel formula issues
            cell = ws[f"{column}{row}"]
            cell_value = cell.value
            
            # Cell value should be properly evaluated with data_only=True
            
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
    def extract_targets_from_row(ws, row: int, procedure_tissues: List[str] = None) -> Tuple[List[str], List[int], List[str], List[str], List[int]]:
        """
        Extract target gene names and tissue names with their column positions from a specific row.
        Now tracks both targets and tissues for data extraction.
        Targets are typically spaced every 4 columns (F, J, N, R, etc.).
        
        Args:
            ws: Excel worksheet object
            row: Row number containing target names
            procedure_tissues: List of tissues from Procedure Request Form for comparison
            
        Returns:
            Tuple of (target_names, target_columns, found_tissues, tissue_names_for_data, tissue_columns_for_data)
        """
        if procedure_tissues is None:
            procedure_tissues = []
            
        targets, target_columns, found_tissues = [], [], []
        tissue_names_for_data, tissue_columns_for_data = [], []  # NEW: Track tissues for data extraction
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
                
                # Classify as tissue or target (simplified - no more "both")
                classification, tissue_name, target_name = classify_target_or_tissue(
                    text_clean, procedure_tissues
                )
                print(f"Target row cell '{text_clean}': classified as {classification}")
                if logger:
                    logger.debug(f"Target row cell '{text_clean}': classified as {classification}")
                
                if classification == 'tissue':
                    found_tissues.append(tissue_name)
                    # Also add to data extraction lists
                    tissue_names_for_data.append(tissue_name)
                    tissue_columns_for_data.append(col_start)
                    if logger:
                        logger.debug(f"  Added to tissues: '{tissue_name}' (will extract data)")
                else:  # target
                    targets.append(target_name)
                    target_columns.append(col_start)
                    if logger:
                        logger.debug(f"  Added to targets: '{target_name}'")
            
            col_start += Config.TARGET_COLUMN_SPACING  # Move to next target column
        
        return targets, target_columns, found_tissues, tissue_names_for_data, tissue_columns_for_data

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
    if logger:
        logger.info(f"Extracting metadata for study: {folder_name}")
        logger.debug(f"Available sheets: {wb.sheetnames}")
    
    if Config.PROCEDURE_SHEET not in wb.sheetnames:
        if logger:
            logger.warning(f"'{Config.PROCEDURE_SHEET}' sheet not found in {folder_name}, using fallback metadata")
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
    
    # Debug: Print extracted triggers
    print(f"Extracted {len(triggers)} triggers from metadata:")
    if logger:
        logger.info(f"Extracted {len(triggers)} triggers from metadata for {folder_name}")
    for i, trigger in enumerate(triggers):
        dose = doses[i] if i < len(doses) else None
        print(f"  {i+1}: '{trigger}' -> dose: {dose}")
        if logger:
            logger.debug(f"  Trigger {i+1}: '{trigger}' -> dose: {dose}")
    
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
    Now also detects dose types (SQ, IV, IM, intratracheal) and extracts doses from trigger names.
    
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
        
        # If no dose from metadata, try to extract from trigger name
        final_dose = dose
        if not dose or dose_str.strip() == "" or dose_str.lower() == "none":
            extracted_dose = extract_dose_from_trigger_name(trigger)
            if extracted_dose:
                final_dose = extracted_dose
                if logger:
                    logger.info(f"No metadata dose for trigger '{trigger}', using extracted dose: '{extracted_dose}'")
        
        trigger_dose_map[str(trigger)] = {
            "dose": final_dose,
            "dose_type": detected_dose_type
        }
        
        if detected_dose_type:
            debug_print(f"Detected dose type '{detected_dose_type}' for trigger '{trigger}': {dose_str}")
        
        if logger:
            logger.debug(f"Trigger dose mapping: '{trigger}' -> dose: '{final_dose}', type: '{detected_dose_type}'")
    
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
    
    if logger:
        logger.info("Extracting relative expression data")
        logger.debug(f"Available sheets: {wb.sheetnames}")
        
    sheet_name = _find_relative_expression_sheet(wb)
    if not sheet_name:
        if logger:
            logger.warning("No relative expression sheet found")
        return None
    
    ws = wb[sheet_name]
    print(f"Using sheet: '{sheet_name}'")
    if logger:
        logger.info(f"Using sheet: '{sheet_name}' for relative expression data")
    
    # Find the data section - handle both standard and calculation sheets
    rel_exp_location = None
    target_row = None
    
    # For "Calcs" sheets, try pattern-based detection instead of header search
    if "calc" in sheet_name.lower():
        print(f"Detected calculation sheet - using pattern-based target detection")
        if logger:
            logger.info(f"Detected calculation sheet '{sheet_name}' - using pattern-based detection")
        
        # For calc sheets, look for the pattern of target names in specific rows
        # Try common target row locations for calc sheets
        potential_target_rows = [15, 16, 17, 18, 19, 20, 25, 30]
        
        for test_row in potential_target_rows:
            # Look for non-empty cells in the target columns that could be gene/target names
            test_targets, test_columns, _, _, _ = ExcelExtractor.extract_targets_from_row(
                ws, test_row, procedure_tissues
            )
            if test_targets:
                target_row = test_row
                print(f"Found targets in row {target_row}: {test_targets}")
                if logger:
                    logger.info(f"Pattern detection found targets in row {target_row}: {test_targets}")
                break
                
        if target_row is None:
            print(f"No target pattern found in calculation sheet {sheet_name}")
            if logger:
                logger.warning(f"No target pattern found in calculation sheet {sheet_name}")
            return None
            
    else:
        # For standard sheets, use header search
        rel_exp_location = ExcelExtractor.find_cell_with_text(
            ws, "relative expression", Config.REL_EXP_SEARCH_ROWS
        )
        if not rel_exp_location:
            print(f"Relative Expression section not found in sheet {sheet_name}")
            if logger:
                logger.warning(f"Relative Expression section not found in sheet {sheet_name}")
            return None
        
        rel_exp_row, _ = rel_exp_location
        target_row = rel_exp_row + 2
    targets, target_columns, found_tissues, tissue_names_for_data, tissue_columns_for_data = ExcelExtractor.extract_targets_from_row(
        ws, target_row, procedure_tissues
    )
    
    # Combine targets and tissues for data extraction
    all_data_names = targets + tissue_names_for_data
    all_data_columns = target_columns + tissue_columns_for_data
    
    if not all_data_names:
        print(f"No targets or tissues found for data extraction in row {target_row}")
        return None
    
    # Extract trigger names - adjust based on sheet type
    if "calc" in sheet_name.lower():
        # For calc sheets, triggers are usually 1-2 rows below targets
        trigger_start_row = target_row + 1
        # First try 1 row below
        triggers = ExcelExtractor.extract_column_values(
            ws, trigger_start_row, "B", stop_on_empty=False
        )[:Config.MAX_TRIGGERS]
        
        # If no triggers found, try 2 rows below
        if not triggers or all(is_empty_or_zero(t) for t in triggers):
            trigger_start_row = target_row + 2
            triggers = ExcelExtractor.extract_column_values(
                ws, trigger_start_row, "B", stop_on_empty=False
            )[:Config.MAX_TRIGGERS]
            
        # If still no triggers, try 3 rows below
        if not triggers or all(is_empty_or_zero(t) for t in triggers):
            trigger_start_row = target_row + 3
            triggers = ExcelExtractor.extract_column_values(
                ws, trigger_start_row, "B", stop_on_empty=False
            )[:Config.MAX_TRIGGERS]
    else:
        # For standard sheets, triggers are usually 3 rows below the header
        trigger_start_row = target_row + 3
        triggers = ExcelExtractor.extract_column_values(
            ws, trigger_start_row, "B", stop_on_empty=False
        )[:Config.MAX_TRIGGERS]
    
    print(f"Found targets: {targets}")
    print(f"Found tissues for data extraction: {tissue_names_for_data}")
    print(f"Found {len(triggers)} triggers in relative expression data:")
    if logger:
        logger.info(f"Found {len(targets)} gene targets, {len(tissue_names_for_data)} tissue targets, and {len(triggers)} triggers in relative expression data")
        logger.debug(f"Gene targets: {targets}")
        logger.debug(f"Tissue targets: {tissue_names_for_data}")
    for i, trigger in enumerate(triggers):
        print(f"  {i+1}: '{trigger}'")
        if logger:
            logger.debug(f"  Trigger {i+1}: '{trigger}'")
    
    if found_tissues:
        print(f"Found tissues in target row: {found_tissues}")
        if logger:
            logger.info(f"Found tissues in target row: {found_tissues}")
    
    # Extract the actual expression data for each trigger-target combination (including tissues)
    triggers_data = _extract_trigger_target_data(ws, triggers, all_data_names, all_data_columns, trigger_start_row)
    
    # Remove triggers with no data
    clean_triggers_data = {k: v for k, v in triggers_data.items() if v}
    
    return {
        "targets": targets,
        "tissue_targets": tissue_names_for_data,  # NEW: Tissues that have expression data
        "relative_expression_data": clean_triggers_data,
        "found_tissues": found_tissues,
        "all_data_items": all_data_names  # NEW: Combined list of all items with data
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
    if logger:
        logger.info(f"Starting processing of study: {folder_name}")
        logger.debug(f"Study folder path: {study_folder}")
        logger.debug(f"Expected info file: {info_file}")
        logger.debug(f"Expected results folder: {results_folder}")
    
    study_data = {}

    # Extract metadata from the main study file
    if os.path.exists(info_file):
        if logger:
            logger.debug(f"Found info file: {info_file}")
        metadata = safe_workbook_operation(info_file, extract_study_metadata, folder_name)
        if metadata:
            study_data.update(metadata)
            print("Extracted metadata fields:")
            if logger:
                logger.info(f"Successfully extracted metadata for {folder_name}")
            for k, v in metadata.items():
                print(f"  {k}: {v}")
                if logger:
                    logger.debug(f"  Metadata {k}: {v}")
        else:
            if logger:
                logger.error(f"Failed to extract metadata from {info_file}")
    else:
        print(f"Info file not found: {info_file}")
        if logger:
            logger.warning(f"Info file not found: {info_file}")
    
    # Extract data from the results file
    results_file = _find_results_file(results_folder)
    if results_file:
        if logger:
            logger.debug(f"Found results file: {results_file}")
    else:
        if logger:
            logger.warning(f"No results file found in {results_folder}")
    
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
    if logger:
        logger.info(f"Starting CSV export to: {output_path}")
        logger.info(f"Processing {len(all_study_data)} studies for CSV export")
    
    header = [
        "study_name", "study_code", "screening_model", "gene_target", "item_type", "trigger", 
        "dose", "dose_type", "timepoint", "tissue", "avg_rel_exp", "avg_rel_exp_lsd", "avg_rel_exp_hsd"
    ]
    
    csv_rows = [header]
    stats = {
        "studies_processed": 0, 
        "studies_with_data": 0, 
        "studies_excluded": 0,
        "total_rows": 0,
        "excluded_studies": []
    }
    
    for study in all_study_data:
        stats["studies_processed"] += 1
        study_name = study.get('study_name', 'Unknown')
        study_code = study.get('study_code', 'Unknown')
        
        if logger:
            logger.debug(f"Processing study for CSV: {study_name} ({study_code})")
        
        rows_added = _process_study_for_csv(study, csv_rows)
        if rows_added > 0:
            stats["studies_with_data"] += 1
            stats["total_rows"] += rows_added
        else:
            stats["studies_excluded"] += 1
            stats["excluded_studies"].append({
                "name": study_name,
                "code": study_code,
                "reason": "No valid data rows generated"
            })
    
    # Write CSV file
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv.writer(csvfile).writerows(csv_rows)
    
    if logger:
        logger.info(f"CSV file written successfully: {output_path}")
        logger.info("="*60)
        logger.info("CSV EXPORT SUMMARY")
        logger.info("="*60)
        logger.info(f"Studies processed for CSV: {stats['studies_processed']}")
        logger.info(f"Studies included in CSV: {stats['studies_with_data']}")
        logger.info(f"Studies excluded from CSV: {stats['studies_excluded']}")
        logger.info(f"Total data rows in CSV: {stats['total_rows']}")
        
        if stats["excluded_studies"]:
            logger.warning("STUDIES EXCLUDED FROM CSV:")
            for excluded in stats["excluded_studies"]:
                logger.warning(f"  - {excluded['name']} ({excluded['code']}): {excluded['reason']}")
    
    _print_export_summary(stats, output_path)

def _process_study_for_csv(study: Dict[str, Any], csv_rows: List[List[str]]) -> int:
    """
    Process a single study for CSV export.
    Groups all rows for a study by tissue (primary) and then by target (secondary).
    If tissue is missing, groups by target.
    """
    study_name = study.get('study_name', 'Unknown')
    study_code = study.get('study_code', 'Unknown')
    
    if "relative_expression" not in study:
        exclusion_reason = "No relative expression data found"
        print(f"Skipping study with no relative expression data: {study_name}")
        if logger:
            logger.warning(f"EXCLUDED FROM CSV: {study_name} ({study_code}) - {exclusion_reason}")
        return 0

    print(f"\nProcessing CSV data for study: {study_name}")
    if logger:
        logger.info(f"Processing CSV data for study: {study_name} ({study_code})")
    
    study_info = _extract_study_info_for_csv(study)
    print(f"  Study info: {study_info}")
    if logger:
        logger.debug(f"Study info for CSV: {study_info}")

    rel_exp_data = study["relative_expression"].get("relative_expression_data", {})
    if not rel_exp_data:
        exclusion_reason = "Relative expression data structure is empty"
        print(f"  No relative expression data found in study")
        if logger:
            logger.warning(f"EXCLUDED FROM CSV: {study_name} ({study_code}) - {exclusion_reason}")
        return 0

    print(f"  Found {len(rel_exp_data)} triggers in relative expression data")
    print(f"  Trigger names in relative expression: {list(rel_exp_data.keys())}")
    print(f"  Trigger names in metadata: {list(study_info['trigger_dose_map'].keys())}")
    
    if logger:
        logger.debug(f"Found {len(rel_exp_data)} triggers in relative expression data")
        logger.debug(f"Trigger names in relative expression: {list(rel_exp_data.keys())}")
        logger.debug(f"Trigger names in metadata: {list(study_info['trigger_dose_map'].keys())}")

        # Collect all rows for this study
    study_rows = []
    
    # Get lists of gene targets and tissue targets for proper classification
    gene_targets = study.get("relative_expression", {}).get("targets", [])
    tissue_targets = study.get("relative_expression", {}).get("tissue_targets", [])
    
    for trigger in rel_exp_data.keys():
            # Get dose info from metadata mapping
            trigger_info = study_info["trigger_dose_map"].get(trigger, {"dose": "", "dose_type": ""})
            
            # Try to extract dose from trigger name if not available in mapping
            extracted_dose = extract_dose_from_trigger_name(trigger)
            if extracted_dose and not trigger_info.get("dose"):
                trigger_info = {
                    "dose": extracted_dose,
                    "dose_type": trigger_info.get("dose_type", "")
                }
                if logger:
                    logger.info(f"Using extracted dose '{extracted_dose}' for trigger '{trigger}'")
            elif trigger_info.get("dose") and extracted_dose:
                # Both sources available - log for comparison
                if logger:
                    logger.debug(f"Trigger '{trigger}': metadata dose='{trigger_info.get('dose')}', extracted dose='{extracted_dose}'")
            
            trigger_data = rel_exp_data[trigger]
            
            if logger:
                logger.debug(f"Processing trigger '{trigger}' with {len(trigger_data)} items, final dose: '{trigger_info.get('dose', 'None')}'")
            
            for item_name, value in trigger_data.items():
                # Determine if this item is a gene target or tissue target
                if item_name in gene_targets:
                    item_type = "gene_target"
                elif item_name in tissue_targets:
                    item_type = "tissue_target"
                else:
                    item_type = "unknown"  # Fallback
                
                if isinstance(value, dict):
                    row = _create_csv_row(
                        study_info, trigger, trigger_info, item_name, item_type,
                        {"rel_exp": value.get("rel_exp"), "low": value.get("low"), "high": value.get("high")}
                    )
                else:
                    row = _create_csv_row(
                        study_info, trigger, trigger_info, item_name, item_type,
                        {"rel_exp": value, "low": None, "high": None}
                    )
                study_rows.append(row)
                
                if logger:
                    logger.debug(f"Added CSV row: {trigger} -> {item_name} ({item_type}) = {value}, dose = {trigger_info.get('dose', 'None')}")

    # Determine index for tissue and target columns in the row
    tissue_idx = 9  # 0-based index for 'tissue' in the row (shifted due to new item_type column)
    target_idx = 3  # 0-based index for 'gene_target' in the row
    item_type_idx = 4  # 0-based index for 'item_type' in the row

    # Sort rows: by tissue (if present), then by item type, then by target
    def sort_key(row):
        tissue = row[tissue_idx] or "ZZZ"  # Put missing tissues at the end
        item_type = row[item_type_idx] or "ZZZ"
        target = row[target_idx] or "ZZZ"
        return (tissue, item_type, target)

    study_rows.sort(key=sort_key)

    # Append sorted rows to the main CSV
    for row in study_rows:
        csv_rows.append(row)

    print(f"  Added {len(study_rows)} rows to CSV for study {study_name}")
    if logger:
        logger.info(f"INCLUDED IN CSV: {study_name} ({study_code}) - Added {len(study_rows)} rows")
    return len(study_rows)

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
                   target: str, item_type: str, values: Dict[str, Any]) -> List[str]:
    """
    Create a single CSV row for one trigger-target combination.
    
    CSV columns (in order):
    1. study_name: Name of the study
    2. study_code: 10-digit code (quoted to prevent Excel issues)
    3. screening_model: Type of screening used
    4. gene_target: Name of the target gene or tissue
    5. item_type: Type of item (gene_target or tissue_target)
    6. trigger: Name of the trigger (e.g., siRNA sequence)
    7. dose: Dose amount for this trigger
    8. dose_type: Detected dose type (SQ, IV, IM, Intratracheal)
    9. timepoint: Time point of measurement (e.g., D14)
    10. tissue: Tissue type tested
    11. avg_rel_exp: Average relative expression value
    12. avg_rel_exp_lsd: Lower standard deviation
    13. avg_rel_exp_hsd: Higher standard deviation
    
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
        target,                                # gene_target (now used for both genes and tissues)
        item_type,                             # item_type (gene_target or tissue_target)
        trigger,                               # trigger (from metadata)
        str(trigger_info.get("dose", "")),     # dose (from metadata)
        str(trigger_info.get("dose_type", "")), # dose_type (detected)
        study_info["timepoint"],
        study_info["tissue"],
        convert_to_numeric(rel_exp),           # avg_rel_exp 
        convert_to_numeric(low),               # avg_rel_exp_lsd
        convert_to_numeric(high)               # avg_rel_exp_hsd
    ]

def _print_export_summary(stats: Dict[str, Any], output_path: str):
    """Print summary statistics after CSV export."""
    print(f"\nExport Summary:")
    print(f"- Total studies processed: {stats['studies_processed']}")
    print(f"- Studies included in CSV: {stats['studies_with_data']}")
    print(f"- Studies excluded from CSV: {stats.get('studies_excluded', 0)}")
    print(f"- Total data rows: {stats['total_rows']}")
    print(f"- Output file: {output_path}")
    
    if stats.get('excluded_studies'):
        print(f"\nExcluded studies:")
        for excluded in stats['excluded_studies']:
            print(f"  - {excluded['name']} ({excluded['code']}): {excluded['reason']}")

# ========================== MAIN EXECUTION ==========================
def main():
    """
    Main execution function.
    """
    # Set up logging first
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_output_dir = os.path.dirname(Config.MONTH_FOLDER)
    month_name = os.path.basename(Config.MONTH_FOLDER).split(' ')[0]
    log_file_path = os.path.join(base_output_dir, f"study_processing_log_{month_name}_{timestamp}.log")
    
    init_logger(log_file_path)
    
    logger.info("="*80)
    logger.info("STUDY PROCESSING SESSION STARTED")
    logger.info("="*80)
    logger.info(f"Processing folder: {Config.MONTH_FOLDER}")
    logger.info(f"Log file: {log_file_path}")
    logger.info(f"Debug mode: {Config.DEBUG}")
    
    study_folders = [
        os.path.join(Config.MONTH_FOLDER, name)
        for name in os.listdir(Config.MONTH_FOLDER)
        if os.path.isdir(os.path.join(Config.MONTH_FOLDER, name))
    ]
    
    if not study_folders:
        error_msg = "No studies found in the month folder."
        print(error_msg)
        logger.error(error_msg)
        return

    print(f"Processing {len(study_folders)} study folders")
    logger.info(f"Found {len(study_folders)} study folders to process")
    logger.debug(f"Study folders: {[os.path.basename(f) for f in study_folders]}")
    
    # Track processing statistics
    processing_stats = {
        "total_folders": len(study_folders),
        "successful_processing": 0,
        "failed_processing": 0,
        "included_in_csv": 0,
        "excluded_from_csv": 0,
        "failed_folders": [],
        "excluded_studies": []
    }
    
    all_study_data = []
    for i, study_folder in enumerate(study_folders, 1):
        folder_name = os.path.basename(study_folder)
        logger.info(f"[{i}/{len(study_folders)}] Processing: {folder_name}")
        
        try:
            study_data = process_study_folder(study_folder)
            if study_data:
                all_study_data.append(study_data)
                processing_stats["successful_processing"] += 1
                logger.info(f"Successfully processed: {folder_name}")
            else:
                processing_stats["failed_processing"] += 1
                processing_stats["failed_folders"].append(folder_name)
                logger.error(f"Failed to extract any data from: {folder_name}")
        except Exception as e:
            processing_stats["failed_processing"] += 1
            processing_stats["failed_folders"].append(folder_name)
            error_msg = f"Exception processing {folder_name}: {e}"
            print(f"ERROR: {error_msg}")
            logger.error(error_msg)
            logger.error(f"Full traceback: {traceback.format_exc()}")

    # Log processing summary
    logger.info("="*80)
    logger.info("PROCESSING SUMMARY")
    logger.info("="*80)
    logger.info(f"Total folders found: {processing_stats['total_folders']}")
    logger.info(f"Successfully processed: {processing_stats['successful_processing']}")
    logger.info(f"Failed processing: {processing_stats['failed_processing']}")
    
    if processing_stats["failed_folders"]:
        logger.warning(f"Failed folders: {processing_stats['failed_folders']}")

    # Save JSON output
    json_timestamp = datetime.now().strftime("%Y%m%d")
    json_output_path = os.path.join(base_output_dir, f"study_metadata_{month_name}_{json_timestamp}.json")
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {json_output_path}")
    logger.info(f"Wrote study metadata to {json_output_path}")
    
    # Export to CSV with detailed logging
    csv_output_path = os.path.join(base_output_dir, f"study_data_{month_name}_{json_timestamp}.csv")
    logger.info("="*80)
    logger.info("CSV EXPORT STARTING")
    logger.info("="*80)
    
    export_to_csv(all_study_data, csv_output_path)
    
    # Create final summary report
    _create_final_summary_report(processing_stats, all_study_data, log_file_path)
    
    # Final summary
    logger.info("="*80)
    logger.info("SESSION COMPLETED")
    logger.info("="*80)
    logger.info(f"Log file saved to: {log_file_path}")
    print(f"\nDetailed log saved to: {log_file_path}")

def _create_final_summary_report(processing_stats: Dict[str, Any], all_study_data: List[Dict[str, Any]], log_file_path: str):
    """Create a comprehensive summary report at the end of the log file."""
    if not logger:
        return
    
    logger.info("="*80)
    logger.info("FINAL PROCESSING REPORT")
    logger.info("="*80)
    
    # Overall statistics
    logger.info("OVERALL STATISTICS:")
    logger.info(f"  Total study folders found: {processing_stats['total_folders']}")
    logger.info(f"  Successfully processed: {processing_stats['successful_processing']}")
    logger.info(f"  Failed to process: {processing_stats['failed_processing']}")
    logger.info(f"  Success rate: {processing_stats['successful_processing']/processing_stats['total_folders']*100:.1f}%")
    
    # Failed folders
    if processing_stats["failed_folders"]:
        logger.warning("FOLDERS THAT FAILED PROCESSING:")
        for folder in processing_stats["failed_folders"]:
            logger.warning(f"  - {folder}")
    
    # Studies with data breakdown
    studies_with_metadata = sum(1 for study in all_study_data if any(k in study for k in ['study_name', 'study_code', 'tissues']))
    studies_with_rel_exp = sum(1 for study in all_study_data if 'relative_expression' in study)
    studies_with_both = sum(1 for study in all_study_data if 'relative_expression' in study and any(k in study for k in ['study_name', 'study_code']))
    
    logger.info("DATA EXTRACTION BREAKDOWN:")
    logger.info(f"  Studies with metadata: {studies_with_metadata}")
    logger.info(f"  Studies with relative expression data: {studies_with_rel_exp}")
    logger.info(f"  Studies with both metadata and expression data: {studies_with_both}")
    
    # Detailed study breakdown
    logger.info("DETAILED STUDY BREAKDOWN:")
    for study in all_study_data:
        study_name = study.get('study_name', 'Unknown')
        study_code = study.get('study_code', 'Unknown')
        
        has_metadata = any(k in study for k in ['study_name', 'study_code', 'tissues', 'trigger_dose_map'])
        has_rel_exp = 'relative_expression' in study
        
        if has_rel_exp:
            rel_exp_data = study['relative_expression'].get('relative_expression_data', {})
            num_triggers = len(rel_exp_data)
            num_targets = len(study['relative_expression'].get('targets', []))
            total_data_points = sum(len(trigger_data) for trigger_data in rel_exp_data.values())
        else:
            num_triggers = num_targets = total_data_points = 0
        
        status = []
        if has_metadata:
            status.append("metadata")
        if has_rel_exp:
            status.append(f"rel_exp({num_triggers}t×{num_targets}g={total_data_points}pts)")
        
        status_str = ", ".join(status) if status else "no_data"
        logger.info(f"  {study_name} ({study_code}): {status_str}")
    
    logger.info("="*80)

if __name__ == "__main__":
    main()