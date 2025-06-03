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
        wb = load_workbook(results_file, data_only=True, read_only=True)
        sheet_name = "Compiled Indiv. & Grp."
        
        if sheet_name not in wb.sheetnames:
            print(f"Sheet '{sheet_name}' not found in {results_file}")
            wb.close()
            return None
            
        ws = wb[sheet_name]
        
        # Check if A125 contains "Relative Expression by Groups"
        if ws["A125"].value != "Relative Expression by Groups":
            print("Relative Expression by Groups section not found at A125")
            wb.close()
            return None
        
        # Extract targets starting from F127, then J127, N127, etc. (every 4 columns)
        targets = []
        target_columns = []
        col_start = 6  # Column F (1-indexed: A=1, B=2, ..., F=6)
        
        while True:
            # Convert column number to letter
            col_letter = chr(ord('A') + col_start - 1)
            target_cell = f"{col_letter}127"
            target_value = ws[target_cell].value
            
            if target_value is None or (isinstance(target_value, str) and not target_value.strip()):
                break
                
            targets.append(str(target_value).strip())
            target_columns.append(col_start)
            col_start += 4  # Move to next target (4 columns apart)
        
        if not targets:
            print("No targets found in row 127")
            wb.close()
            return None
        
        print(f"Found targets: {targets}")
        
        # Extract trigger data starting from row 129
        triggers_data = {}
        
        # First, get all trigger names from column E (assuming triggers are in column E starting from row 129)
        row = 129
        triggers = []
        while True:
            trigger_cell = ws[f"E{row}"].value
            if trigger_cell is None or (isinstance(trigger_cell, str) and not trigger_cell.strip()):
                break
            triggers.append(str(trigger_cell).strip())
            row += 1
        
        print(f"Found triggers: {triggers}")
        
        # For each trigger, extract the data for each target
        for trigger_idx, trigger in enumerate(triggers):
            trigger_row = 129 + trigger_idx
            triggers_data[trigger] = {}
            
            for target_idx, target in enumerate(targets):
                # Calculate column positions for this target
                base_col = target_columns[target_idx]
                rel_exp_col = chr(ord('A') + base_col)      # G, K, O, etc.
                low_col = chr(ord('A') + base_col + 1)      # H, L, P, etc.
                high_col = chr(ord('A') + base_col + 2)     # I, M, Q, etc.
                
                # Extract values
                rel_exp_val = ws[f"{rel_exp_col}{trigger_row}"].value
                low_val = ws[f"{low_col}{trigger_row}"].value
                high_val = ws[f"{high_col}{trigger_row}"].value
                
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