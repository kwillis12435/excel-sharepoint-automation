#!/usr/bin/env python3
"""
Debug script to examine exactly what's happening with target detection and data extraction.
"""

import os
import sys
sys.path.append('.')

from process_study import (
    process_study_folder, init_logger, init_targets, 
    safe_workbook_operation, extract_relative_expression_data,
    ExcelExtractor, Config
)

def debug_single_study():
    """Debug a single study to see exactly what's happening."""
    
    # Initialize systems
    init_logger("debug_single_study.log")
    init_targets()
    
    # Focus on rIL33_8_Alternaria
    study_name = "rIL33_8_Alternaria"
    base_path = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    
    # Find the study folder
    study_folder = None
    for folder_name in os.listdir(base_path):
        if study_name.lower() in folder_name.lower():
            study_folder = os.path.join(base_path, folder_name)
            print(f"Found study folder: {folder_name}")
            break
    
    if not study_folder:
        print(f"❌ Study folder not found for: {study_name}")
        return
    
    # Get the results file
    results_folder = os.path.join(study_folder, "Results")
    results_file = None
    if os.path.exists(results_folder):
        for file in os.listdir(results_folder):
            if file.endswith(".xlsm"):
                results_file = os.path.join(results_folder, file)
                break
    
    if not results_file:
        print(f"❌ No results file found")
        return
    
    print(f"📁 Results file: {results_file}")
    
    # Debug the target detection process
    def debug_target_detection(wb):
        print(f"\n🎯 DEBUGGING TARGET DETECTION")
        print(f"Available sheets: {wb.sheetnames}")
        
        # Find the relative expression sheet
        sheet_name = None
        for possible_name in Config.RELATIVE_EXPRESSION_SHEETS:
            if possible_name in wb.sheetnames:
                sheet_name = possible_name
                break
        
        if not sheet_name:
            for sheet in wb.sheetnames:
                sheet_lower = sheet.lower()
                if ("compiled" in sheet_lower and ("indiv" in sheet_lower or "grp" in sheet_lower)):
                    sheet_name = sheet
                    break
        
        if not sheet_name:
            print("❌ No relative expression sheet found")
            return None
        
        print(f"📊 Using sheet: '{sheet_name}'")
        ws = wb[sheet_name]
        
        # Look for the "Relative Expression" header
        rel_exp_location = ExcelExtractor.find_cell_with_text(
            ws, "relative expression", Config.REL_EXP_SEARCH_ROWS
        )
        
        if rel_exp_location:
            rel_exp_row, _ = rel_exp_location
            target_row = rel_exp_row + 2
            print(f"🔍 Found 'Relative Expression' at row {rel_exp_row}, target row: {target_row}")
        else:
            print("❌ 'Relative Expression' header not found")
            return None
        
        # Debug what's in the target row
        print(f"\n📋 EXAMINING TARGET ROW {target_row}:")
        for col in range(Config.TARGET_START_COLUMN, Config.TARGET_START_COLUMN + 20, Config.TARGET_COLUMN_SPACING):
            cell_value = ws.cell(row=target_row, column=col).value
            if cell_value:
                print(f"  Column {col}: '{cell_value}'")
        
        # Extract targets using our function
        targets, target_columns, found_tissues, tissue_names_for_data, tissue_columns_for_data = ExcelExtractor.extract_targets_from_row(
            ws, target_row, []
        )
        
        print(f"\n📊 EXTRACTION RESULTS:")
        print(f"  Gene targets: {targets}")
        print(f"  Target columns: {target_columns}")
        print(f"  Tissue targets: {tissue_names_for_data}")
        print(f"  Tissue columns: {tissue_columns_for_data}")
        print(f"  Found tissues: {found_tissues}")
        
        # Debug trigger extraction
        trigger_start_row = target_row + 3
        print(f"\n🎯 EXAMINING TRIGGER AREA (starting row {trigger_start_row}):")
        for row in range(trigger_start_row, trigger_start_row + 10):
            trigger_cell = ws.cell(row=row, column=2).value  # Column B
            if trigger_cell:
                print(f"  Row {row}: '{trigger_cell}'")
        
        return {
            'sheet_name': sheet_name,
            'target_row': target_row,
            'trigger_start_row': trigger_start_row,
            'targets': targets,
            'target_columns': target_columns,
            'tissue_names_for_data': tissue_names_for_data,
            'tissue_columns_for_data': tissue_columns_for_data
        }
    
    # Debug data extraction
    def debug_data_extraction(wb, debug_info):
        if not debug_info:
            return
        
        print(f"\n💾 DEBUGGING DATA EXTRACTION")
        ws = wb[debug_info['sheet_name']]
        
        # Extract triggers
        raw_triggers = ExcelExtractor.extract_column_values(
            ws, debug_info['trigger_start_row'], "B", stop_on_empty=False
        )[:Config.MAX_TRIGGERS]
        
        print(f"📋 Raw triggers found: {raw_triggers}")
        
        # Filter triggers
        triggers = []
        for trigger in raw_triggers:
            trigger_str = str(trigger).strip()
            if (trigger_str.lower() not in ['x', 'n/a', 'blank', 'none', '--', '-', ''] and
                not trigger_str.startswith('#')):
                triggers.append(trigger_str)
        
        print(f"📋 Filtered triggers: {triggers}")
        
        # Combine all data items
        all_data_names = debug_info['targets'] + debug_info['tissue_names_for_data']
        all_data_columns = debug_info['target_columns'] + debug_info['tissue_columns_for_data']
        
        print(f"📊 All data items: {all_data_names}")
        print(f"📊 All data columns: {all_data_columns}")
        
        # Extract data for first few combinations to see the pattern
        print(f"\n🔍 SAMPLE DATA EXTRACTION:")
        for trigger_idx, trigger in enumerate(triggers[:3]):  # Just first 3 triggers
            trigger_row = debug_info['trigger_start_row'] + trigger_idx
            print(f"\n  Trigger: '{trigger}' (row {trigger_row})")
            
            for target_idx, target in enumerate(all_data_names):
                if target_idx >= len(all_data_columns):
                    break
                base_col = all_data_columns[target_idx]
                
                # Extract the three data values
                rel_exp = ws.cell(row=trigger_row, column=base_col + 1).value
                low = ws.cell(row=trigger_row, column=base_col + 2).value
                high = ws.cell(row=trigger_row, column=base_col + 3).value
                
                print(f"    {target} (col {base_col}): rel_exp={rel_exp}, low={low}, high={high}")
                
                if not all(v is None for v in [rel_exp, low, high]):
                    print(f"      ✅ Has data!")
                else:
                    print(f"      ❌ No data")
    
    # Run the debugging
    debug_info = safe_workbook_operation(results_file, debug_target_detection)
    safe_workbook_operation(results_file, debug_data_extraction, debug_info)
    
    # Now run the full extraction to compare
    print(f"\n🔄 RUNNING FULL EXTRACTION:")
    study_data = process_study_folder(study_folder)
    
    if study_data and 'relative_expression' in study_data:
        rel_exp_data = study_data['relative_expression']
        triggers_data = rel_exp_data.get('relative_expression_data', {})
        
        print(f"\n📈 FULL EXTRACTION RESULTS:")
        print(f"  Targets found: {rel_exp_data.get('targets', [])}")
        print(f"  Tissue targets found: {rel_exp_data.get('tissue_targets', [])}")
        print(f"  Triggers with data: {list(triggers_data.keys())}")
        
        total_combinations = 0
        for trigger, target_data in triggers_data.items():
            combo_count = len(target_data)
            total_combinations += combo_count
            print(f"    {trigger}: {combo_count} targets")
        
        print(f"  Total combinations: {total_combinations}")
    else:
        print(f"❌ Full extraction failed or no data found")

if __name__ == "__main__":
    debug_single_study() 