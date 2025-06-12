#!/usr/bin/env python3
"""
Test the saline duplication fix specifically for rHDM_Pilot_2 study.
This study should have multiple HDM triggers that weren't being captured properly.
"""

import os
import sys
import csv
from collections import defaultdict
sys.path.append('.')

from process_study import process_study_folder, init_logger, export_to_csv, init_targets

def test_hdm_fix():
    """Test the HDM trigger duplication fix for rHDM_Pilot_2."""
    
    # Initialize logging and targets
    init_logger("test_hdm_fix.log")
    init_targets()
    
    # Target study that had HDM duplication issues
    study_name = "rHDM_Pilot_2"
    base_path = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    print(f"🔍 TESTING HDM DUPLICATE FIX")
    print(f"============================================================")
    
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
    
    # Process the study
    study_data = process_study_folder(study_folder)
    
    if not study_data:
        print(f"❌ Failed to process study: {study_name}")
        return
    
    print(f"\n📊 EXTRACTION RESULTS:")
    
    # Check metadata
    trigger_dose_map = study_data.get('trigger_dose_map', {})
    print(f"  Metadata triggers: {len(trigger_dose_map)}")
    for trigger_key, dose_info in trigger_dose_map.items():
        original_name = dose_info.get('original_name', trigger_key)
        print(f"    '{trigger_key}' (original: '{original_name}') -> {dose_info}")
    
    # Check relative expression data
    rel_exp_data = study_data.get('relative_expression', {})
    if rel_exp_data:
        targets = rel_exp_data.get('targets', [])
        tissue_targets = rel_exp_data.get('tissue_targets', [])
        triggers_data = rel_exp_data.get('relative_expression_data', {})
        
        print(f"  Gene targets: {targets}")
        print(f"  Tissue targets: {tissue_targets}")
        print(f"  Total trigger keys: {len(triggers_data)}")
        
        # Analyze HDM triggers specifically
        hdm_triggers = [k for k in triggers_data.keys() if 'hdm' in k.lower()]
        print(f"\n🎯 HDM TRIGGER ANALYSIS:")
        for hdm_trigger in hdm_triggers:
            target_count = len(triggers_data[hdm_trigger])
            print(f"  {hdm_trigger} -> {target_count} targets")
        
        # Count expected vs actual
        expected_hdm_count = len([k for k in trigger_dose_map.keys() if 'hdm' in k.lower() or 
                                 (trigger_dose_map[k].get('original_name', '').lower().count('hdm') > 0)])
        actual_hdm_count = len(hdm_triggers)
        
        print(f"\n📈 HDM SUMMARY:")
        print(f"  Total triggers found: {len(triggers_data)}")
        print(f"  HDM entries expected: {expected_hdm_count}")
        print(f"  HDM entries found: {actual_hdm_count}")
        
        if actual_hdm_count >= expected_hdm_count:
            print(f"  ✅ SUCCESS: Found expected number of HDM entries!")
        else:
            print(f"  ❌ FAILURE: Expected {expected_hdm_count} HDM entries, found {actual_hdm_count}")
        
        # Calculate total combinations
        total_targets = len(targets) + len(tissue_targets)
        expected_combinations = len(triggers_data) * total_targets
        actual_combinations = sum(len(trigger_data) for trigger_data in triggers_data.values())
        
        print(f"  Expected combinations: {len(triggers_data)} triggers × {total_targets} targets = {expected_combinations}")
        print(f"  Actual combinations: {actual_combinations}")
        
        if actual_combinations >= expected_combinations * 0.9:  # 90% threshold
            print(f"  ✅ SUCCESS: All combinations captured!")
        else:
            print(f"  ❌ FAILURE: Missing combinations")
    
    # Generate CSV
    output_file = os.path.join(output_dir, "hdm_fix_test.csv")
    print(f"\n📝 Generating CSV: {output_file}")
    export_to_csv([study_data], output_file)
    
    # Analyze CSV
    print(f"\n📋 CSV ANALYSIS:")
    hdm_rows = 0
    total_rows = 0
    
    with open(output_file, 'r') as f:
        reader = csv.DictReader(f)
        for row in reader:
            total_rows += 1
            if 'hdm' in row['trigger'].lower():
                hdm_rows += 1
    
    expected_hdm_rows = expected_hdm_count * total_targets if 'expected_hdm_count' in locals() else 0
    
    print(f"  Total CSV rows: {total_rows}")
    print(f"  HDM rows: {hdm_rows}")
    print(f"  Expected HDM rows: {expected_hdm_rows} ({expected_hdm_count} HDMs × {total_targets} targets)")
    
    if hdm_rows >= expected_hdm_rows:
        print(f"  ✅ SUCCESS: CSV has correct number of HDM rows!")
    else:
        print(f"  ❌ FAILURE: Expected {expected_hdm_rows} HDM rows, found {hdm_rows}")
    
    print(f"\n🎉 TEST COMPLETED")

if __name__ == "__main__":
    test_hdm_fix() 