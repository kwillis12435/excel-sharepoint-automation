import pandas as pd
import os
from openpyxl import load_workbook
import re
import json

# Path to the month folder
MONTH_FOLDER = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"

def extract_study_metadata_by_cell(info_file, folder_name):
    wb = load_workbook(info_file, data_only=True, read_only=True)
    sheet_name = "Procedure Request Form"
    study_name = None
    study_code = None
    triggers = []
    doses = []
    screening_model = None
    tissues = []

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
            if cell_value is None or (isinstance(cell_value, str) and not cell_value.strip()):
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
            if trigger_cell is None or (isinstance(trigger_cell, str) and not trigger_cell.strip()):
                break
            triggers.append(trigger_cell)
            dose_cell = ws[f"D{row}"].value
            doses.append(dose_cell if dose_cell is not None else None)
            row += 1
    else:
        study_name = None
        screening_model = None
        tissues = []
        match = re.match(r"(\d{10})", folder_name)
        if match:
            study_code = match.group(1)
        else:
            study_code = None

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
    }
    wb.close()
    return fields

def extract_relative_expression_data(results_file):
    """
    Extract relative expression data from the 'Compiled Indiv. & Grp.' sheet.
    Returns a dictionary with targets and their corresponding trigger data.
    """
    try:
        # Load the workbook without read_only mode to use cell references
        wb = load_workbook(results_file, data_only=True, read_only=False)
        
        # Debug: Print all sheet names
        print(f"Available sheets in {os.path.basename(results_file)}: {wb.sheetnames}")
        
        # Try to find the correct sheet name (case insensitive)
        sheet_name = None
        target_sheet_names = [
            "Compiled Indiv. & Grp.",
            "Compiled Indiv. & Grp",
            "Compiled Indiv & Grp.",
            "Compiled Indiv & Grp"
        ]
        
        for target_name in target_sheet_names:
            if target_name in wb.sheetnames:
                sheet_name = target_name
                break
        
        # If exact match not found, try case-insensitive search
        if not sheet_name:
            for sheet in wb.sheetnames:
                if "compiled" in sheet.lower() and ("indiv" in sheet.lower() or "grp" in sheet.lower()):
                    sheet_name = sheet
                    break
        
        if not sheet_name:
            print(f"No 'Compiled Indiv. & Grp.' sheet found in {results_file}")
            wb.close()
            return None
            
        print(f"Using sheet: '{sheet_name}'")
        ws = wb[sheet_name]
        
        # Look for "Relative Expression by Groups" in nearby cells
        rel_exp_row = None
        for check_row in range(120, 135):
            cell_value = ws.cell(row=check_row, column=1).value  # Column A
            print(f"Cell A{check_row} value: '{cell_value}'")
            if cell_value and isinstance(cell_value, str) and "relative expression" in cell_value.lower():
                rel_exp_row = check_row
                print(f"Found 'Relative Expression by Groups' at A{check_row}: '{cell_value}'")
                break
        
        if not rel_exp_row:
            print("Relative Expression by Groups section not found between rows 120-135")
            wb.close()
            return None
        
        # Extract targets starting from row rel_exp_row + 2 (so if found at 125, look at 127)
        target_row = rel_exp_row + 2
        print(f"Looking for targets in row {target_row}")
        
        # Debug: Check what's in the target row
        for col_num in range(1, 20):  # Check columns A through S
            col_letter = chr(ord('A') + col_num - 1)
            cell_value = ws.cell(row=target_row, column=col_num).value
            if cell_value:
                print(f"  {col_letter}{target_row}: '{cell_value}'")
        
        # Extract targets starting from F127, then J127, N127, etc. (every 4 columns)
        targets = []
        target_columns = []
        col_start = 6  # Column F (1-indexed: A=1, B=2, ..., F=6)
        zero_count = 0  # Counter for consecutive zeros
        
        while True:
            # Convert column number to letter for debugging only
            col_letter = chr(ord('A') + col_start - 1)
            target_value = ws.cell(row=target_row, column=col_start).value
            
            print(f"Checking {col_letter}{target_row} for target: '{target_value}'")
            
            # Check if value is zero or empty
            is_zero_or_empty = (target_value is None or 
                               (isinstance(target_value, str) and not target_value.strip()) or
                               target_value == 0 or 
                               str(target_value).strip() == "0")
            
            if is_zero_or_empty:
                zero_count += 1
                # Stop if we encounter 5 consecutive zeros
                if zero_count >= 5:
                    print(f"Found {zero_count} consecutive zeros/empty cells - stopping target search")
                    break
            else:
                # Reset zero counter if we find a non-zero value
                zero_count = 0
                targets.append(str(target_value).strip())
                target_columns.append(col_start)
            
            col_start += 4  # Move to next target (4 columns apart)
            
            # Safety break to avoid infinite loop
            if len(targets) > 20:  # Increased limit for safety
                break
        
        if not targets:
            print(f"No targets found in row {target_row}")
            wb.close()
            return None
        
        print(f"Found targets: {targets}")
        print(f"Target columns: {target_columns}")
        
        # Extract trigger data starting from row target_row + 3 (130 if target_row is 127)
        trigger_start_row = target_row + 3  # Start at row 130
        print(f"Looking for triggers starting at row {trigger_start_row}")
        
        # Debug: Check what's in column B around the trigger start row
        for check_row in range(trigger_start_row - 2, trigger_start_row + 10):
            trigger_value = ws.cell(row=check_row, column=2).value  # Column B = 2
            if trigger_value:
                print(f"  B{check_row}: '{trigger_value}'")
        
        # Extract trigger data
        triggers_data = {}
        
        # Get all trigger names from column B starting at row 130
        row = trigger_start_row
        triggers = []
        while True:
            trigger_cell = ws.cell(row=row, column=2).value  # Column B = 2
            if trigger_cell is None or (isinstance(trigger_cell, str) and not trigger_cell.strip()):
                break
            triggers.append(str(trigger_cell).strip())
            row += 1
            
            # Safety break to avoid infinite loop
            if len(triggers) > 20:
                break
        
        print(f"Found triggers: {triggers}")
        
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
                
                # For debugging, convert to letter notation
                rel_exp_letter = chr(ord('A') + rel_exp_col - 1)
                low_letter = chr(ord('A') + low_col - 1)
                high_letter = chr(ord('A') + high_col - 1)
                
                # Extract values using cell() method
                rel_exp_val = ws.cell(row=trigger_row, column=rel_exp_col).value
                low_val = ws.cell(row=trigger_row, column=low_col).value
                high_val = ws.cell(row=trigger_row, column=high_col).value
                
                print(f"  {trigger} + {target}: {rel_exp_letter}{trigger_row}={rel_exp_val}, {low_letter}{trigger_row}={low_val}, {high_letter}{trigger_row}={high_val}")
                
                triggers_data[trigger][target] = {
                    "rel_exp": rel_exp_val,
                    "low": low_val,
                    "high": high_val
                }
        
        wb.close()
        
        return {
            "targets": targets,
            "relative_expression_data": triggers_data
        }
        
    except Exception as e:
        print(f"Error extracting relative expression data: {e}")
        import traceback
        traceback.print_exc()
        return None

def extract_lar_sheet_fields(df):
    # Try to find relevant fields in the first 20 rows
    fields = {}
    for i, row in df.iterrows():
        for col in range(len(row)):
            cell = str(row[col]).strip().lower()
            if "trigger" in cell:
                fields["trigger"] = str(row[col+1]).strip() if col+1 < len(row) else ""
            if "dose" in cell:
                fields["dose"] = str(row[col+1]).strip() if col+1 < len(row) else ""
            if "tissue" in cell:
                fields["tissue"] = str(row[col+1]).strip() if col+1 < len(row) else ""
            if "timepoint" in cell:
                fields["timepoint"] = str(row[col+1]).strip() if col+1 < len(row) else ""
    return fields

def process_study_folder(study_folder):
    folder_name = os.path.basename(study_folder)
    info_file = os.path.join(study_folder, f"{folder_name}.xlsm")
    results_folder = os.path.join(study_folder, "Results")
    # Find the first .xlsm file in Results
    results_file = None
    if os.path.exists(results_folder):
        for f in os.listdir(results_folder):
            if f.endswith(".xlsm"):
                results_file = os.path.join(results_folder, f)
                break

    print(f"\nProcessing study: {folder_name}")

    study_data = {}

    if os.path.exists(info_file):
        try:
            fields = extract_study_metadata_by_cell(info_file, folder_name)
            study_data.update(fields)
            print("Extracted metadata fields:")
            for k, v in fields.items():
                print(f"  {k}: {v}")
        except Exception as e:
            print(f"Error reading info file: {e}")
    else:
        print(f"Info file not found for {folder_name}")

    # Extract from LAR Sheet in results file
    if results_file and os.path.exists(results_file):
        try:
            xls = pd.ExcelFile(results_file)
            if "LAR Sheet" in xls.sheet_names:
                df_lar = pd.read_excel(results_file, sheet_name="LAR Sheet", header=None)
                lar_fields = extract_lar_sheet_fields(df_lar)
                study_data["lar_data"] = lar_fields
                print("Extracted LAR fields:")
                for k, v in lar_fields.items():
                    print(f"  {k}: {v}")
        except Exception as e:
            print(f"Error reading LAR sheet: {e}")
    
    # Extract relative expression data
    if results_file and os.path.exists(results_file):
        rel_exp_data = extract_relative_expression_data(results_file)
        if rel_exp_data:
            study_data["relative_expression"] = rel_exp_data
            print("Extracted relative expression data:")
            print(f"  Targets: {rel_exp_data['targets']}")
            print(f"  Number of triggers: {len(rel_exp_data['relative_expression_data'])}")
            # Print sample data for first trigger and target
            if rel_exp_data['relative_expression_data'] and rel_exp_data['targets']:
                first_trigger = list(rel_exp_data['relative_expression_data'].keys())[0]
                first_target = rel_exp_data['targets'][0]
                sample_data = rel_exp_data['relative_expression_data'][first_trigger][first_target]
                print(f"  Sample data ({first_trigger}, {first_target}): {sample_data}")
    
    return study_data

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

    all_study_data = []
    for study_folder in study_folders:
        study_data = process_study_folder(study_folder)
        if study_data:
            all_study_data.append(study_data)

    # Write to JSON file in the Discovery Biology - 2024 folder
    output_path = os.path.join(
        os.path.dirname(MONTH_FOLDER), "study_metadata_01-2024.json"
    )
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(all_study_data, f, indent=2, ensure_ascii=False)
    print(f"\nWrote study metadata to {output_path}")

if __name__ == "__main__":
    main()