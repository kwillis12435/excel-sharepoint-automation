import pandas as pd
import os
from openpyxl import load_workbook
import re
import json
import csv
from datetime import datetime
import traceback
from typing import Dict, List, Optional, Tuple, Any

# ========================== CONFIGURATION ==========================
class Config:
    MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    DEBUG = False
    
    # Excel sheet names and locations
    PROCEDURE_SHEET = "Procedure Request Form"
    RELATIVE_EXPRESSION_SHEETS = [
        "Compiled Indiv. & Grp.",
        "Compiled Indiv. & Grp",
        "Compiled Indiv & Grp.",
        "Compiled Indiv & Grp",
        "Calcs Norm to D1 & Ctrl",
        "Calcs Norm to Pre & Control",
        "Calcs Norm to D1 & Ctrl."
    ]
    LAR_SHEET = "LAR Sheet"
    
    # Cell locations
    STUDY_NAME_CELL = "C14"
    SCREENING_MODEL_CELL = "M6"
    STUDY_CODE_CELL = "M12"
    TISSUES_START_ROW = 17
    TISSUES_COLUMN = "S"
    TRIGGERS_START_ROW = 80
    TRIGGERS_COLUMN = "B"
    DOSES_COLUMN = "D"
    
    # Relative expression data
    REL_EXP_SEARCH_ROWS = range(120, 135)
    TARGET_START_COLUMN = 6  # Column F
    TARGET_COLUMN_SPACING = 4
    MAX_TARGETS = 30
    MAX_TRIGGERS = 20

# ========================== UTILITY FUNCTIONS ==========================
def debug_print(*args, **kwargs):
    """Print debug messages only when DEBUG is enabled"""
    if Config.DEBUG:
        print(*args, **kwargs)

def is_empty_or_zero(value: Any) -> bool:
    """Check if a value is None, empty string, or zero"""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    if value == 0 or str(value).strip() == "0":
        return True
    return False

def normalize_string(text: str) -> str:
    """Remove spaces, special characters, and convert to lowercase"""
    if not text:
        return ""
    return ''.join(c for c in str(text).lower().strip() if c.isalnum())

def format_timepoint(timepoint: str) -> str:
    """Format timepoint to start with 'D' if it doesn't already"""
    if not timepoint:
        return timepoint
    
    timepoint = str(timepoint).strip()
    if not timepoint.startswith('D') and timepoint and timepoint[0].isdigit():
        return f"D{timepoint}"
    return timepoint

def convert_to_numeric(value: Any) -> str:
    """Convert a value to a numeric format with consistent decimal places"""
    if value is None:
        return ""
    try:
        float_val = float(str(value).strip())
        return f"{float_val:.4f}" if float_val != 0 else "0.0000"
    except (ValueError, TypeError):
        return str(value).strip()

def safe_workbook_operation(file_path: str, operation_func, *args, **kwargs):
    """Safely perform workbook operations with proper cleanup"""
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
    @staticmethod
    def extract_column_values(ws, start_row: int, column: str, stop_on_empty: bool = True) -> List[str]:
        """Extract values from a column starting at a specific row"""
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
        """Extract paired values from two columns"""
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
        """Find a cell containing specific text"""
        search_range = search_rows or range(1, min(ws.max_row, 200))
        for row in search_range:
            for col in range(1, min(ws.max_column, 50)):
                cell_value = ws.cell(row=row, column=col).value
                if cell_value and isinstance(cell_value, str) and search_text.lower() in cell_value.lower():
                    return row, col
        return None

    @staticmethod
    def extract_targets_from_row(ws, row: int) -> Tuple[List[str], List[int]]:
        """Extract targets and their column positions from a row"""
        targets, target_columns = [], []
        col_start = Config.TARGET_START_COLUMN
        zero_count = 0
        
        while len(targets) < Config.MAX_TARGETS:
            target_value = ws.cell(row=row, column=col_start).value
            
            if is_empty_or_zero(target_value):
                zero_count += 1
                if zero_count >= 5:
                    break
            else:
                zero_count = 0
                targets.append(str(target_value).strip())
                target_columns.append(col_start)
            
            col_start += Config.TARGET_COLUMN_SPACING
        
        return targets, target_columns

# ========================== STUDY METADATA EXTRACTION ==========================
def extract_study_metadata(wb, folder_name: str) -> Dict[str, Any]:
    """Extract all study metadata from the procedure request form"""
    if Config.PROCEDURE_SHEET not in wb.sheetnames:
        return _extract_fallback_metadata(folder_name)
    
    ws = wb[Config.PROCEDURE_SHEET]
    
    # Basic metadata
    study_name = ws[Config.STUDY_NAME_CELL].value
    screening_model = _determine_screening_model(study_name, ws[Config.SCREENING_MODEL_CELL].value)
    study_code = _extract_study_code(ws[Config.STUDY_CODE_CELL].value, folder_name)
    
    # Extract lists
    tissues = _extract_unique_tissues(ws)
    triggers, doses = ExcelExtractor.extract_paired_columns(
        ws, Config.TRIGGERS_START_ROW, Config.TRIGGERS_COLUMN, Config.DOSES_COLUMN
    )
    
    # Create trigger-dose mapping
    trigger_dose_map = _create_trigger_dose_map(triggers, doses)
    
    # Extract timepoint
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
    """Extract basic metadata when procedure sheet is not available"""
    study_code = None
    match = re.match(r"(\d{10})", folder_name)
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
    """Determine screening model based on study name or cell value"""
    if study_name and "aav" in str(study_name).lower():
        return "AAV"
    return model_cell_value

def _extract_study_code(cell_value: Any, folder_name: str) -> Optional[str]:
    """Extract study code from cell or folder name"""
    if cell_value and re.fullmatch(r"\d{10}", str(cell_value)):
        return str(cell_value)
    
    match = re.match(r"(\d{10})", folder_name)
    return match.group(1) if match else None

def _extract_unique_tissues(ws) -> List[str]:
    """Extract unique tissue types"""
    tissues = ExcelExtractor.extract_column_values(
        ws, Config.TISSUES_START_ROW, Config.TISSUES_COLUMN
    )
    return list(dict.fromkeys(tissues))  # Remove duplicates while preserving order

def _create_trigger_dose_map(triggers: List[str], doses: List[Any]) -> Dict[str, Any]:
    """Create mapping between triggers and doses"""
    # Ensure lists are same length
    while len(doses) < len(triggers):
        doses.append(None)
    return {str(trigger): dose for trigger, dose in zip(triggers, doses[:len(triggers)])}

def _extract_timepoint(ws) -> Optional[str]:
    """Extract timepoint from worksheet"""
    # Search for timepoint header
    timepoint_location = ExcelExtractor.find_cell_with_text(ws, "day")
    if not timepoint_location:
        timepoint_location = ExcelExtractor.find_cell_with_text(ws, "timepoint")
    
    if not timepoint_location:
        return None
    
    header_row, header_col = timepoint_location
    
    # Find last non-empty value in the column
    last_val = None
    for row in range(header_row + 1, ws.max_row + 1):
        val = ws.cell(row=row, column=header_col).value
        if val is not None and str(val).strip():
            last_val = val
    
    return format_timepoint(str(last_val).strip()) if last_val else None

# ========================== RELATIVE EXPRESSION DATA EXTRACTION ==========================
def extract_relative_expression_data(wb) -> Optional[Dict[str, Any]]:
    """Extract relative expression data from workbook"""
    sheet_name = _find_relative_expression_sheet(wb)
    if not sheet_name:
        return None
    
    ws = wb[sheet_name]
    print(f"Using sheet: '{sheet_name}'")
    
    # Find relative expression section
    rel_exp_location = ExcelExtractor.find_cell_with_text(
        ws, "relative expression", Config.REL_EXP_SEARCH_ROWS
    )
    if not rel_exp_location:
        print(f"Relative Expression section not found in sheet {sheet_name}")
        return None
    
    rel_exp_row, _ = rel_exp_location
    
    # Extract targets and triggers
    target_row = rel_exp_row + 2
    targets, target_columns = ExcelExtractor.extract_targets_from_row(ws, target_row)
    
    if not targets:
        print(f"No targets found in row {target_row}")
        return None
    
    trigger_start_row = target_row + 3
    triggers = ExcelExtractor.extract_column_values(
        ws, trigger_start_row, "B", stop_on_empty=False
    )[:Config.MAX_TRIGGERS]
    
    print(f"Found targets: {targets}")
    print(f"Found triggers: {triggers}")
    
    # Extract data for each trigger-target combination
    triggers_data = _extract_trigger_target_data(ws, triggers, targets, target_columns, trigger_start_row)
    
    # Clean up empty triggers
    clean_triggers_data = {k: v for k, v in triggers_data.items() if v}
    
    return {
        "targets": targets,
        "relative_expression_data": clean_triggers_data
    }

def _find_relative_expression_sheet(wb) -> Optional[str]:
    """Find the sheet containing relative expression data"""
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
    """Extract data for each trigger-target combination"""
    triggers_data = {}
    
    for trigger_idx, trigger in enumerate(triggers):
        if is_empty_or_zero(trigger):
            continue
            
        trigger_row = trigger_start_row + trigger_idx
        triggers_data[trigger] = {}
        
        for target_idx, target in enumerate(targets):
            base_col = target_columns[target_idx]
            
            # Data columns: rel_exp, low, high
            values = {
                "rel_exp": ws.cell(row=trigger_row, column=base_col + 1).value,
                "low": ws.cell(row=trigger_row, column=base_col + 2).value,
                "high": ws.cell(row=trigger_row, column=base_col + 3).value
            }
            
            # Skip if all values are empty
            if all(v is None for v in values.values()):
                continue
            
            triggers_data[trigger][target] = values
            
            debug_print(f"  {trigger} + {target}: {values}")
    
    return triggers_data

# ========================== STRING MATCHING UTILITIES ==========================
class StringMatcher:
    @staticmethod
    def find_best_match(target: str, candidates: List[str]) -> Optional[str]:
        """Find the best matching string using multiple strategies"""
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
    """Process a single study folder and extract all data"""
    folder_name = os.path.basename(study_folder)
    info_file = os.path.join(study_folder, f"{folder_name}.xlsm")
    results_folder = os.path.join(study_folder, "Results")
    
    print(f"\nProcessing study: {folder_name}")
    
    study_data = {}
    
    # Extract metadata
    if os.path.exists(info_file):
        metadata = safe_workbook_operation(info_file, extract_study_metadata, folder_name)
        if metadata:
            study_data.update(metadata)
            print("Extracted metadata fields:")
            for k, v in metadata.items():
                print(f"  {k}: {v}")
    else:
        print(f"Info file not found: {info_file}")
    
    # Extract results data
    results_file = _find_results_file(results_folder)
    if results_file:
        # Extract LAR data
        lar_data = _extract_lar_data(results_file)
        if lar_data:
            study_data["lar_data"] = lar_data
        
        # Extract relative expression data
        rel_exp_data = safe_workbook_operation(
            results_file, extract_relative_expression_data, read_only=False
        )
        if rel_exp_data:
            study_data["relative_expression"] = rel_exp_data
            print(f"Extracted relative expression data: {len(rel_exp_data['targets'])} targets, "
                  f"{len(rel_exp_data['relative_expression_data'])} triggers")
    
    return study_data if study_data else None

def _find_results_file(results_folder: str) -> Optional[str]:
    """Find the first .xlsm file in the Results folder"""
    if not os.path.exists(results_folder):
        return None
    
    for file in os.listdir(results_folder):
        if file.endswith(".xlsm"):
            return os.path.join(results_folder, file)
    
    return None

def _extract_lar_data(results_file: str) -> Optional[Dict[str, str]]:
    """Extract data from LAR Sheet"""
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
    """Export study data to CSV format"""
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
    """Process a single study for CSV export"""
    if "relative_expression" not in study:
        return 0
    
    study_info = _extract_study_info_for_csv(study)
    rel_exp_data = study["relative_expression"]["relative_expression_data"]
    
    rows_added = 0
    for trigger, dose in study_info["trigger_dose_map"].items():
        matching_trigger = StringMatcher.find_best_match(trigger, list(rel_exp_data.keys()))
        
        if not matching_trigger:
            continue
        
        trigger_data = rel_exp_data[matching_trigger]
        for target, values in trigger_data.items():
            if all(not values.get(key) for key in ['rel_exp', 'low', 'high']):
                continue
            
            row = _create_csv_row(study_info, trigger, dose, target, values)
            csv_rows.append(row)
            rows_added += 1
    
    return rows_added

def _extract_study_info_for_csv(study: Dict[str, Any]) -> Dict[str, Any]:
    """Extract and format study information for CSV export"""
    timepoint = study.get("timepoint", "")
    if timepoint and not timepoint.startswith('D') and timepoint.strip().isdigit():
        timepoint = f"D{timepoint.strip()}"
    
    tissue = ""
    if study.get("tissues"):
        tissue = study["tissues"][0]
    elif "lar_data" in study and "tissue" in study["lar_data"]:
        tissue = study["lar_data"]["tissue"]
    
    return {
        "study_name": study.get("study_name", ""),
        "study_code": f"'{study.get('study_code', '')}'" if study.get("study_code") else "",
        "screening_model": study.get("screening_model", ""),
        "trigger_dose_map": study.get("trigger_dose_map", {}),
        "timepoint": timepoint,
        "tissue": tissue
    }

def _create_csv_row(study_info: Dict[str, Any], trigger: str, dose: Any, 
                   target: str, values: Dict[str, Any]) -> List[str]:
    """Create a single CSV row"""
    return [
        study_info["study_name"],
        study_info["study_code"],
        study_info["screening_model"],
        target,
        trigger,
        dose,
        study_info["timepoint"],
        study_info["tissue"],
        convert_to_numeric(values.get("rel_exp")),
        convert_to_numeric(values.get("low")),
        convert_to_numeric(values.get("high"))
    ]

def _print_export_summary(stats: Dict[str, int], output_path: str):
    """Print export summary statistics"""
    print(f"\nExport Summary:")
    print(f"- Total studies processed: {stats['studies_processed']}")
    print(f"- Studies with data: {stats['studies_with_data']}")
    print(f"- Total data rows: {stats['total_rows']}")
    print(f"- Output file: {output_path}")

# ========================== MAIN EXECUTION ==========================
def main():
    """Main execution function"""
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
    
    # Generate output files
    timestamp = datetime.now().strftime("%Y%m%d")
    base_output_dir = os.path.dirname(Config.MONTH_FOLDER)
    month_name = os.path.basename(Config.MONTH_FOLDER).split(' ')[0]
    
    # Export to JSON
    json_output_path = os.path.join(base_output_dir, f"study_metadata_{month_name}_{timestamp}.json")
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {json_output_path}")
    
    # Export to CSV
    csv_output_path = os.path.join(base_output_dir, f"study_data_{month_name}_{timestamp}.csv")
    export_to_csv(all_study_data, csv_output_path)

if __name__ == "__main__":
    main()