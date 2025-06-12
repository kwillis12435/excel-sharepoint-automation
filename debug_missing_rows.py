#!/usr/bin/env python3
"""
Debug script to identify why experimental processing is missing rows.
"""

import os
import sys
sys.path.append('.')
from process_study import process_study_folder, Config, extract_relative_expression_data, safe_workbook_operation

def debug_missing_rows():
    """Debug why we're missing rows compared to manual processing."""
    
    base_folder = Config.MONTH_FOLDER
    
    # Test a few specific studies that show differences
    test_studies = [
        "2024010101 (G2rCNS_21)",      # 30 vs 54 rows (-24)
        "2024010506 (mF12_phos_38)",   # 0 vs 11 rows (missing completely)
        "2024013001 (hACVR2b_SEAP_KD_1)", # 0 vs 60 rows (missing completely)
        "2024010105 (mChiral_6)",      # 7 vs 24 rows (-17)
        "2024010806 (mSOD1_50_ICV)"    # 52 vs 104 rows (-52)
    ]
    
    for study_name in test_studies:
        study_folder = os.path.join(base_folder, study_name)
        
        if not os.path.exists(study_folder):
            print(f"❌ Study folder not found: {study_name}")
            continue
            
        print(f"\n{'='*60}")
        print(f"🔍 DEBUGGING: {study_name}")
        print(f"{'='*60}")
        
        # Check if study processes at all
        try:
            result = process_study_folder(study_folder)
            if not result:
                print("❌ Study failed to process - no data returned")
                continue
                
            print(f"✅ Study processed successfully")
            print(f"   Study name: {result.get('study_name', 'Unknown')}")
            
            # Check relative expression data specifically
            if 'relative_expression' not in result:
                print("❌ No relative expression data found")
                continue
                
            rel_exp = result['relative_expression']
            rel_exp_data = rel_exp.get('relative_expression_data', {})
            
            print(f"📊 Relative Expression Data:")
            print(f"   Gene targets: {len(rel_exp.get('targets', []))}")
            print(f"   Tissue targets: {len(rel_exp.get('tissue_targets', []))}")
            print(f"   Triggers with data: {len(rel_exp_data)}")
            
            # Count total data points
            total_data_points = 0
            for trigger, trigger_data in rel_exp_data.items():
                total_data_points += len(trigger_data)
                print(f"     {trigger}: {len(trigger_data)} items")
            
            print(f"   Total data points: {total_data_points}")
            
            # Check what items are being found
            all_targets = rel_exp.get('targets', [])
            all_tissue_targets = rel_exp.get('tissue_targets', [])
            print(f"   All gene targets: {all_targets}")
            print(f"   All tissue targets: {all_tissue_targets}")
            
            # Manual check of the results file
            results_folder = os.path.join(study_folder, "Results")
            results_file = None
            if os.path.exists(results_folder):
                for file in os.listdir(results_folder):
                    if file.endswith(".xlsm"):
                        results_file = os.path.join(results_folder, file)
                        break
            
            if results_file:
                print(f"\n🔍 Manual inspection of results file:")
                try:
                    procedure_tissues = result.get("tissues", [])
                    manual_rel_exp = safe_workbook_operation(
                        results_file, 
                        debug_extract_relative_expression,
                        procedure_tissues
                    )
                    if manual_rel_exp:
                        print(f"   Manual extraction found: {manual_rel_exp}")
                except Exception as e:
                    print(f"   Manual extraction failed: {e}")
            
        except Exception as e:
            print(f"❌ Exception processing study: {e}")
            import traceback
            traceback.print_exc()

def debug_extract_relative_expression(wb, procedure_tissues):
    """Manual debug version of relative expression extraction."""
    
    print(f"   Available sheets: {wb.sheetnames}")
    
    # Try to find calculation sheets
    calc_sheets = [s for s in wb.sheetnames if 'calc' in s.lower() or 'compiled' in s.lower()]
    print(f"   Calculation sheets found: {calc_sheets}")
    
    if not calc_sheets:
        return None
    
    # Use the first calculation sheet
    sheet_name = calc_sheets[0]
    ws = wb[sheet_name]
    print(f"   Using sheet: {sheet_name}")
    print(f"   Sheet dimensions: {ws.max_row} rows × {ws.max_column} columns")
    
    # Look for data in the target columns (F, J, N, etc.)
    target_columns = [6, 10, 14, 18, 22, 26, 30]  # F, J, N, R, V, Z, AD
    
    # Scan for potential target rows
    potential_target_rows = []
    for row in range(100, min(200, ws.max_row + 1)):
        non_empty_count = 0
        row_content = []
        for col in target_columns[:5]:  # Check first 5 target columns
            try:
                cell_value = ws.cell(row=row, column=col).value
                if cell_value and str(cell_value).strip():
                    non_empty_count += 1
                    row_content.append(str(cell_value).strip())
            except:
                continue
        
        if non_empty_count >= 2:
            potential_target_rows.append((row, row_content))
            print(f"   Potential target row {row}: {row_content}")
    
    # Look for trigger data in column B
    trigger_rows = []
    for row in range(120, min(200, ws.max_row + 1)):
        try:
            cell_value = ws.cell(row=row, column=2).value  # Column B
            if cell_value and str(cell_value).strip() and str(cell_value).strip() not in ['', 'x', 'X']:
                trigger_rows.append((row, str(cell_value).strip()))
        except:
            continue
    
    print(f"   Found {len(trigger_rows)} potential triggers in column B")
    if trigger_rows:
        print(f"   First 5 triggers: {[t[1] for t in trigger_rows[:5]]}")
    
    return {
        "potential_targets": len(potential_target_rows),
        "potential_triggers": len(trigger_rows),
        "target_rows": potential_target_rows[:3],  # First 3 for debugging
        "trigger_samples": trigger_rows[:5]        # First 5 for debugging
    }

if __name__ == "__main__":
    debug_missing_rows() 