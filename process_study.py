import pandas as pd
import os
from openpyxl import load_workbook
import re
import json
import csv
from datetime import datetime
import traceback

# Configuration
MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
DEBUG = False  # Set to True for detailed debug output

def debug_print(*args, **kwargs):
    """Helper function to print debug messages only when DEBUG is enabled"""
    if DEBUG:
        print(*args, **kwargs)

def extract_study_metadata_by_cell(info_file, folder_name):
    """Extract study metadata from the Procedure Request Form sheet"""
    wb = load_workbook(info_file, data_only=True, read_only=True)
    sheet_name = "Procedure Request Form"
    study_name = None
    study_code = None
    triggers = []
    doses = []
    screening_model = None
    tissues = []
    timepoint = None

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Study name from C14
        study_name = ws["C14"].value

        # Screening model logic
        if study_name and "aav" in str(study_name).lower():
            screening_model = "AAV"
        else:
            screening_model = ws["M6"].value

        # Tissues from S17 down
        row = 17
        seen_tissues = set()
        while True:
            cell_value = ws[f"S{row}"].value
            if is_empty_or_zero(cell_value):
                break
            if cell_value not in seen_tissues:
                tissues.append(cell_value)
                seen_tissues.add(cell_value)
            row += 1

        # Try to get study code from M12
        study_code_candidate = ws["M12"].value
        if study_code_candidate and re.fullmatch(r"\d{10}", str(study_code_candidate)):
            study_code = str(study_code_candidate)
        else:
            match = re.match(r"(\d{10})", folder_name)
            if match:
                study_code = match.group(1)

        # Extract triggers from B80 down, and doses from D80 down
        row = 80
        while True:
            trigger_cell = ws[f"B{row}"].value
            if is_empty_or_zero(trigger_cell):
                break
            triggers.append(trigger_cell)
            dose_cell = ws[f"D{row}"].value
            doses.append(dose_cell if dose_cell is not None else None)
            row += 1

        # --- Timepoint extraction logic ---
        timepoint = None
        header_row = None
        header_col = None
        # Search for header cell containing 'Day #' or 'timepoint' (case-insensitive)
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                val = str(cell.value).lower() if cell.value is not None else ""
                if ("day" in val and "#" in val) or ("timepoint" in val):
                    header_row = cell.row
                    header_col = cell.column
                    break
            if header_row:
                break
                
        # If found, collect all non-empty cells below in the same column
        if header_row and header_col:
            last_val = None
            for r in range(header_row + 1, ws.max_row + 1):
                v = ws.cell(row=r, column=header_col).value
                if v is not None and str(v).strip() != "":
                    last_val = v
            if last_val is not None:
                timepoint = format_timepoint(str(last_val).strip())
        # --- End timepoint extraction ---
    else:
        # Fallback: Try to extract study code from folder name
        match = re.match(r"(\d{10})", folder_name)
        if match:
            study_code = match.group(1)

    # Ensure doses list matches triggers list in length
    while len(doses) < len(triggers):
        doses.append(None)
    if len(doses) > len(triggers):
        doses = doses[:len(triggers)]

    # Create trigger-dose mapping
    trigger_dose_map = {str(trigger): dose for trigger, dose in zip(triggers, doses)}

    fields = {
        "study_name": study_name,
        "study_code": study_code,
        "screening_model": screening_model,
        "tissues": tissues,
        "trigger_dose_map": trigger_dose_map,
        "timepoint": timepoint,
    }
    wb.close()
    return fields

def extract_relative_expression_data(results_file):
    """
    Extract relative expression data from the 'Compiled Indiv. & Grp.' or 'Calcs Norm to D1 & Ctrl' sheet.
    Returns a dictionary with targets and their corresponding trigger data.
    """
    try:
        # Load the workbook without read_only mode to use cell references
        wb = load_workbook(results_file, data_only=True, read_only=False)
        
        print(f"Available sheets in {os.path.basename(results_file)}: {wb.sheetnames}")
        
        # Try to find the correct sheet name
        sheet_name = find_relative_expression_sheet(wb)
        if not sheet_name:
            print(f"Could not find any suitable sheet in {results_file}")
            wb.close()
            return None
            
        print(f"Using sheet: '{sheet_name}'")
        ws = wb[sheet_name]
        
        # Find the relative expression section row
        rel_exp_row = find_relative_expression_row(ws)
        if not rel_exp_row:
            print(f"Relative Expression section not found in sheet {sheet_name}")
            wb.close()
            return None
        
        # Extract targets and their columns
        target_row = rel_exp_row + 2
        targets, target_columns = extract_targets(ws, target_row)
        if not targets:
            print(f"No targets found in row {target_row}")
            wb.close()
            return None
        
        print(f"Found targets: {targets}")
        
        # Extract trigger data
        trigger_start_row = target_row + 3
        triggers = extract_triggers(ws, trigger_start_row)
        print(f"Found triggers in Column B: {triggers}")
        
        # Extract data for each trigger-target combination
        triggers_data = extract_trigger_target_data(ws, triggers, targets, target_columns, trigger_start_row)
        
        # Clean up empty targets
        clean_triggers_data = {trigger: target_data for trigger, target_data in triggers_data.items() if target_data}
        
        wb.close()
        
        return {
            "targets": targets,
            "relative_expression_data": clean_triggers_data
        }
        
    except Exception as e:
        print(f"Error extracting relative expression data: {e}")
        traceback.print_exc()
        return None

def find_relative_expression_sheet(wb):
    """Find the sheet containing relative expression data"""
    # Try exact match with known sheet names
    target_sheet_names = [
        "Compiled Indiv. & Grp.",
        "Compiled Indiv. & Grp",
        "Compiled Indiv & Grp.",
        "Compiled Indiv & Grp",
        "Calcs Norm to D1 & Ctrl",
        "Calcs Norm to Pre & Control",
        "Calcs Norm to D1 & Ctrl."
    ]
    
    for target_name in target_sheet_names:
        if target_name in wb.sheetnames:
            return target_name
    
    # If exact match not found, try case-insensitive search
    for sheet in wb.sheetnames:
        if ("compiled" in sheet.lower() and ("indiv" in sheet.lower() or "grp" in sheet.lower())) or \
           ("calcs" in sheet.lower() and "norm" in sheet.lower() and "ctrl" in sheet.lower()):
            return sheet
    
    # Last resort - try any sheet with relevant keywords
    for sheet in wb.sheetnames:
        if "result" in sheet.lower() or "data" in sheet.lower() or "expression" in sheet.lower():
            print(f"Trying alternative sheet: {sheet}")
            return sheet
    
    return None

def find_relative_expression_row(ws):
    """Find the row containing 'Relative Expression' header"""
    # First try specific rows where it's commonly found
    for check_row in range(120, 135):
        cell_value = ws.cell(row=check_row, column=1).value  # Column A
        debug_print(f"Cell A{check_row} value: '{cell_value}'")
        if cell_value and isinstance(cell_value, str) and "relative expression" in cell_value.lower():
            print(f"Found 'Relative Expression' at A{check_row}: '{cell_value}'")
            return check_row
    
    # If not found in common rows, search the entire sheet
    print("Relative Expression section not found in common location. Searching entire sheet...")
    for row_idx in range(1, min(ws.max_row, 200)):  # Limit search to first 200 rows
        cell_value = ws.cell(row=row_idx, column=1).value  # Column A
        if cell_value and isinstance(cell_value, str) and "relative expression" in cell_value.lower():
            print(f"Found 'Relative Expression' at A{row_idx}: '{cell_value}'")
            return row_idx
    
    return None

def extract_targets(ws, target_row):
    """Extract targets and their column positions"""
    debug_print(f"Looking for targets in row {target_row}")
    
    # Debug: Check what's in the target row
    for col_num in range(1, 20):  # Check columns A through S
        col_letter = chr(ord('A') + col_num - 1)
        cell_value = ws.cell(row=target_row, column=col_num).value
        if cell_value:
            debug_print(f"  {col_letter}{target_row}: '{cell_value}'")
    
    # Extract targets starting from F (col 6), then J, N, etc. (every 4 columns)
    targets = []
    target_columns = []
    col_start = 6  # Column F (1-indexed: A=1, B=2, ..., F=6)
    zero_count = 0  # Counter for consecutive zeros
    
    while True:
        col_letter = chr(ord('A') + col_start - 1)
        target_value = ws.cell(row=target_row, column=col_start).value
        
        debug_print(f"Checking {col_letter}{target_row} for target: '{target_value}'")
        
        # Check if value is zero or empty
        if is_empty_or_zero(target_value):
            zero_count += 1
            if zero_count >= 5:
                debug_print(f"Found {zero_count} consecutive zeros/empty cells - stopping target search")
                break
        else:
            # Reset zero counter if we find a non-zero value
            zero_count = 0
            targets.append(str(target_value).strip())
            target_columns.append(col_start)
        
        col_start += 4  # Move to next target (4 columns apart)
        
        # Safety break to avoid infinite loop
        if len(targets) > 30:  # Increased limit for safety
            break
    
    # If no targets found, try with wider search
    if not targets:
        print(f"No targets found using standard spacing. Trying alternative approach...")
        col_spacing = 3  # Try different column spacing
        col_start = 6    # Start at column F again
        
        while col_start < ws.max_column:
            target_value = ws.cell(row=target_row, column=col_start).value
            if target_value and not is_empty_or_zero(target_value):
                targets.append(str(target_value).strip())
                target_columns.append(col_start)
            
            col_start += col_spacing
            if len(targets) > 30:  # Safety limit
                break
    
    return targets, target_columns

def extract_triggers(ws, trigger_start_row):
    """Extract triggers from column B"""
    trigger_col = 2  # Column B
    row = trigger_start_row
    triggers = []
    
    # Extract triggers from Column B
    while row < min(ws.max_row, trigger_start_row + 20):  # Limit to 20 triggers
        trigger_value = ws.cell(row=row, column=trigger_col).value
        if trigger_value and not is_empty_or_zero(trigger_value):
            triggers.append(str(trigger_value).strip())
        row += 1
    
    return triggers

def extract_trigger_target_data(ws, triggers, targets, target_columns, trigger_start_row):
    """Extract data for each trigger-target combination"""
    triggers_data = {}
    
    # For each trigger, extract the data for each target
    for trigger_idx, trigger in enumerate(triggers):
        trigger_row = trigger_start_row + trigger_idx
        triggers_data[trigger] = {}
        
        for target_idx, target in enumerate(targets):
            # Calculate column positions for this target
            base_col = target_columns[target_idx]
            # For target in column F (6), the data is in G, H, I (7, 8, 9)
            rel_exp_col = base_col + 1    # G, K, O, etc.
            low_col = base_col + 2        # H, L, P, etc.
            high_col = base_col + 3       # I, M, Q, etc.
            
            # Extract values
            rel_exp_val = ws.cell(row=trigger_row, column=rel_exp_col).value
            low_val = ws.cell(row=trigger_row, column=low_col).value
            high_val = ws.cell(row=trigger_row, column=high_col).value
            
            # Debug column letters for troubleshooting
            if DEBUG:
                rel_exp_letter = chr(ord('A') + rel_exp_col - 1)
                low_letter = chr(ord('A') + low_col - 1)
                high_letter = chr(ord('A') + high_col - 1)
                debug_print(f"  {trigger} + {target}: {rel_exp_letter}{trigger_row}={rel_exp_val}, {low_letter}{trigger_row}={low_val}, {high_letter}{trigger_row}={high_val}")
            
            # Skip adding data if all values are null/empty
            if rel_exp_val is None and low_val is None and high_val is None:
                debug_print(f"  Skipping empty data for {trigger} + {target}")
                continue
            
            triggers_data[trigger][target] = {
                "rel_exp": rel_exp_val,
                "low": low_val,
                "high": high_val
            }
    
    return triggers_data

def export_to_csv(all_study_data, output_path):
    """
    Export the study data to a CSV file using metadata as source of truth.
    The function will:
    1. Use trigger_dose_map from metadata as primary source
    2. Match relative expression data with metadata triggers
    3. Format the data in the required CSV structure
    4. Convert numeric values properly
    """
    csv_rows = []
    header = [
        "study_name", "study_code", "screening_model", "gene_target", "trigger", "dose", "timepoint", "tissue", "avg_rel_exp", "avg_rel_exp_lsd", "avg_rel_exp_hsd"
    ]
    csv_rows.append(header)
    
    studies_processed = 0
    studies_with_data = 0
    
    for study in all_study_data:
        studies_processed += 1
        study_metadata = {
            "study_name": study.get("study_name", ""),
            "study_code": study.get("study_code", ""),
            "screening_model": study.get("screening_model", ""),
            "trigger_dose_map": study.get("trigger_dose_map", {}),
            "timepoint": study.get("timepoint", ""),
            "tissues": study.get("tissues", [])
        }
        
        print(f"\n{'='*80}")
        print(f"Processing study: {study_metadata['study_name']} ({study_metadata['study_code']})")
        print(f"Triggers from metadata: {list(study_metadata['trigger_dose_map'].keys())}")
        
        # Format timepoint
        if study_metadata["timepoint"] and not study_metadata["timepoint"].startswith('D'):
            if study_metadata["timepoint"].strip().isdigit():
                study_metadata["timepoint"] = f"D{study_metadata['timepoint'].strip()}"
            elif study_metadata["timepoint"].strip() and study_metadata["timepoint"].strip()[0].isdigit():
                study_metadata["timepoint"] = f"D{study_metadata['timepoint'].strip()}"
        
        # Get tissue (prefer metadata, fallback to LAR data)
        tissue = study_metadata["tissues"][0] if study_metadata["tissues"] else ""
        if not tissue and "lar_data" in study and "tissue" in study["lar_data"]:
            tissue = study["lar_data"]["tissue"]
        
        rel_exp_data = study.get("relative_expression", {})
        if not rel_exp_data:
            print(f"No relative expression data for study: {study_metadata['study_name']}")
            continue
            
        rel_exp_results = rel_exp_data.get("relative_expression_data", {})
        if not rel_exp_results:
            print(f"No relative expression results for study: {study_metadata['study_name']}")
            continue
            
        print(f"Triggers from results file: {list(rel_exp_results.keys())}")
        
        studies_with_data += 1
        rows_added = 0
        
        # Process each trigger from metadata
        for trigger, dose in study_metadata["trigger_dose_map"].items():
            print(f"\nProcessing metadata trigger: {trigger}")
            # Find matching trigger in relative expression data (case-insensitive and removing whitespace)
            matching_trigger = None
            trigger_clean = str(trigger).lower().strip()
            
            # First try exact match
            for rel_trigger in rel_exp_results:
                rel_trigger_clean = str(rel_trigger).lower().strip()
                if trigger_clean == rel_trigger_clean:
                    matching_trigger = rel_trigger
                    print(f"Found exact match: {rel_trigger}")
                    break
                    
            # Then try prefix match (trigger at start of string)
            if matching_trigger is None:
                for rel_trigger in rel_exp_results:
                    rel_trigger_clean = str(rel_trigger).lower().strip()
                    if rel_trigger_clean.startswith(trigger_clean):
                        matching_trigger = rel_trigger
                        print(f"Found prefix match: {rel_trigger}")
                        break
            
            # If still no match, try normalized match (removing spaces and special chars)
            if matching_trigger is None:
                trigger_norm = ''.join(c for c in trigger_clean if c.isalnum())
                for rel_trigger in rel_exp_results:
                    rel_trigger_norm = ''.join(c for c in str(rel_trigger).lower().strip() if c.isalnum())
                    if rel_trigger_norm.startswith(trigger_norm):
                        matching_trigger = rel_trigger
                        print(f"Found normalized match: {rel_trigger}")
                        break
            
            if matching_trigger is None:
                print(f"  No expression data found for trigger: {trigger}")
                print(f"  Available triggers: {list(rel_exp_results.keys())}")
                continue
                
            trigger_data = rel_exp_results[matching_trigger]
            print(f"  Processing trigger: {trigger} (Found {len(trigger_data)} targets)")
            
            # Process each target for this trigger
            for target, values in trigger_data.items():
                # Skip if no valid values
                if all(not values.get(key) for key in ['rel_exp', 'low', 'high']):
                    print(f"  Skipping {target} - no valid values")
                    continue
                
                # Debug values
                print(f"  Target {target} values: {values}")
                
                # Convert values to numeric format
                rel_exp = convert_to_numeric(values.get("rel_exp"))
                low = convert_to_numeric(values.get("low"))
                high = convert_to_numeric(values.get("high"))
                
                print(f"  Converted values - rel_exp: {rel_exp}, low: {low}, high: {high}")
                      # Ensure study_code is formatted as a raw string to prevent scientific notation
                study_code = f"'{study_metadata['study_code']}'" if study_metadata["study_code"] else ""
                
                row = [
                    study_metadata["study_name"],    # study_name
                    study_code,                      # study_code (as string to prevent scientific notation)
                    study_metadata["screening_model"], # screening_model
                    target,                          # gene_target
                    trigger,                         # trigger (use metadata version)
                    dose,                           # dose (from metadata)
                    study_metadata["timepoint"],     # timepoint
                    tissue,                         # tissue
                    rel_exp,                       # avg_rel_exp
                    low,                           # avg_rel_exp_lsd
                    high                           # avg_rel_exp_hsd
                ]
                csv_rows.append(row)
                rows_added += 1
                print(f"  Added row for {trigger} - {target}")
        
        print(f"Added {rows_added} rows for study {study_metadata['study_name']}")
        print('='*80)
    
    # Write to CSV file
    with open(output_path, 'w', newline='', encoding='utf-8') as csvfile:
        csv_writer = csv.writer(csvfile)
        csv_writer.writerows(csv_rows)
    
    print(f"\nExport Summary:")
    print(f"- Total studies processed: {studies_processed}")
    print(f"- Studies with data: {studies_with_data}")
    print(f"- Total data rows: {len(csv_rows) - 1}")  # Subtract header row
    print(f"- Output file: {output_path}")

def main():
    # List all study folders in the month folder
    study_folders = [
        os.path.join(MONTH_FOLDER, name)
        for name in os.listdir(MONTH_FOLDER)
        if os.path.isdir(os.path.join(MONTH_FOLDER, name))
    ]
    
    if not study_folders:
        print("No studies found in the month folder.")
        return
    
    print(f"Processing {len(study_folders)} study folders")
    
    all_study_data = []
    for study_folder in study_folders:
        study_data = process_study_folder(study_folder)
        if study_data:
            all_study_data.append(study_data)
            print("\nStudy metadata:")
            print(f"Triggers in metadata: {list(study_data['trigger_dose_map'].keys())}")
            if 'relative_expression' in study_data:
                print("Relative expression triggers:", list(study_data['relative_expression']['relative_expression_data'].keys()))

    # Generate timestamp for output files
    timestamp = datetime.now().strftime("%Y%m%d")
    base_output_dir = os.path.dirname(MONTH_FOLDER)
    month_name = os.path.basename(MONTH_FOLDER).split(' ')[0]  # Extract month number
    
    # Write to JSON file
    json_output_path = os.path.join(
        base_output_dir, f"study_metadata_{month_name}_{timestamp}_test.json"
    )
    with open(json_output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {json_output_path}")
    
    # Write to CSV file
    csv_output_path = os.path.join(
        base_output_dir, f"study_data_{month_name}_{timestamp}_test.csv"
    )
    export_to_csv(all_study_data, csv_output_path)

def is_empty_or_zero(value):
    """Check if a value is None, empty string, or zero"""
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    if value == 0 or str(value).strip() == "0":
        return True
    return False

def format_timepoint(timepoint):
    """Format timepoint to start with 'D' if it doesn't already"""
    if not timepoint:
        return timepoint
    
    timepoint = str(timepoint).strip()
    if not timepoint.startswith('D'):
        # Check if it's just a number (possibly with whitespace)
        if timepoint.isdigit():
            return f"D{timepoint}"
        # Or if it starts with a number after whitespace
        elif timepoint and timepoint[0].isdigit():
            return f"D{timepoint}"
    return timepoint

def convert_to_numeric(value):
    """Convert a value to a numeric format with consistent decimal places"""
    if value is None:
        return ""
    try:
        # Try to convert to float first
        float_val = float(str(value).strip())
        # Format with consistent decimal places
        return f"{float_val:.4f}" if float_val != 0 else "0.0000"
    except (ValueError, TypeError):
        return str(value).strip()

def normalize_string(text):
    """Remove spaces, special characters, and convert to lowercase"""
    if text is None:
        return ""
    return ''.join(c for c in str(text).lower().strip() if c.isalnum())

if __name__ == "__main__":
    main()