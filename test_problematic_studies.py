#!/usr/bin/env python3
"""
Test script to verify fixes for problematic studies, focusing on repeated trigger data in CSV output.
"""

import os
import sys
import csv
from collections import defaultdict
sys.path.append('.')

from process_study import process_study_folder, init_logger, export_to_csv, init_targets

def test_problematic_studies():
    """Test the specific studies that were problematic, focusing on repeated trigger data."""
    
    # Initialize logging
    init_logger("test_problematic_studies.log")
    
    # Initialize target detection system
    init_targets()
    
    # List of problematic studies to test
    problematic_studies = [
        "rIL33_8_Alternaria",  # Didn't find saline twice
        "rHDM_Pilot_2",        # Didn't find HDM 3 times
        "hIL33_AAV8_KD_4",     # Only found one instance of triggers
        "rPCSK6_7",            # Only found half the triggers
        "mChiral_6",           # Found days instead of triggers
        "hIL33_AAV8_KD_5"      # Only found one instance of data/triggers
    ]
    
    base_path = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    output_dir = "test_output"
    os.makedirs(output_dir, exist_ok=True)
    
    all_study_data = []
    
    for study_name in problematic_studies:
        print(f"\n{'='*80}")
        print(f"🔍 Testing study: {study_name}")
        print(f"{'='*80}")
        
        # Find the study folder
        study_folder = None
        for folder_name in os.listdir(base_path):
            if study_name.lower() in folder_name.lower():
                study_folder = os.path.join(base_path, folder_name)
                print(f"Found study folder: {folder_name}")
                break
        
        if not study_folder:
            print(f"❌ Study folder not found for: {study_name}")
            continue
        
        try:
            study_data = process_study_folder(study_folder)
            
            if study_data:
                print(f"\n📊 STUDY DATA SUMMARY:")
                print(f"  Study name: {study_data.get('study_name')}")
                print(f"  Study code: {study_data.get('study_code')}")
                
                # Check metadata triggers
                trigger_dose_map = study_data.get('trigger_dose_map', {})
                print(f"\n📋 METADATA TRIGGERS: {len(trigger_dose_map)} triggers")
                for i, (trigger, dose_info) in enumerate(trigger_dose_map.items(), 1):
                    print(f"  {i}: '{trigger}' -> {dose_info}")
                
                # Check relative expression data
                rel_exp_data = study_data.get('relative_expression', {})
                if rel_exp_data:
                    rel_exp_triggers = rel_exp_data.get('relative_expression_data', {})
                    print(f"\n🧬 RELATIVE EXPRESSION DATA: {len(rel_exp_triggers)} triggers")
                    print(f"  Gene targets: {rel_exp_data.get('targets', [])}")
                    print(f"  Tissue targets: {rel_exp_data.get('tissue_targets', [])}")
                    
                    # Show which triggers have data
                    print(f"\n🎯 TRIGGERS WITH DATA:")
                    for trigger in rel_exp_triggers.keys():
                        data_count = len(rel_exp_triggers[trigger])
                        print(f"  - {trigger}: {data_count} target(s)")
                    
                    # Calculate total data points for CSV
                    total_data_points = 0
                    for trigger_data in rel_exp_triggers.values():
                        total_data_points += len(trigger_data)
                    
                    print(f"\n📈 TOTAL DATA POINTS: {total_data_points}")
                    print(f"  Expected CSV rows: {total_data_points}")
                    
                    # Check if enhanced matching was used
                    if rel_exp_data.get('enhanced_matching'):
                        print(f"  ✅ Enhanced matching was used")
                        if rel_exp_data.get('positional_mapping'):
                            print(f"  ✅ Positional mapping was used")
                    
                    # Compare with expected
                    expected_triggers = len(trigger_dose_map)
                    actual_triggers = len(rel_exp_triggers)
                    
                    print(f"\n🔄 COMPARISON:")
                    print(f"  Metadata triggers: {expected_triggers}")
                    print(f"  Expression triggers: {actual_triggers}")
                    print(f"  Success rate: {actual_triggers}/{expected_triggers} ({actual_triggers/expected_triggers*100:.1f}%)")
                    
                    if actual_triggers >= expected_triggers * 0.8:  # 80% success rate
                        print(f"  🎉 SUCCESS: Good trigger capture rate!")
                    elif actual_triggers > 1:
                        print(f"  ⚠️ PARTIAL: Some triggers captured")
                    else:
                        print(f"  ❌ FAILURE: Very low trigger capture")
                    
                    # Show raw triggers found for debugging
                    if 'raw_triggers_found' in rel_exp_data:
                        print(f"\n🔍 RAW TRIGGERS FOUND:")
                        for i, trigger in enumerate(rel_exp_data['raw_triggers_found'], 1):
                            print(f"  {i}: '{trigger}'")
                    
                    # Add to list for CSV generation
                    all_study_data.append(study_data)
                else:
                    print(f"\n❌ NO RELATIVE EXPRESSION DATA FOUND")
            else:
                print(f"\n❌ STUDY PROCESSING FAILED")
                
        except Exception as e:
            print(f"\n💥 EXCEPTION: {e}")
            import traceback
            traceback.print_exc()
    
    # Generate CSV for all studies
    if all_study_data:
        output_file = os.path.join(output_dir, "problematic_studies_test.csv")
        print(f"\n📝 Generating CSV file: {output_file}")
        export_to_csv(all_study_data, output_file)
        
        # Analyze CSV output
        print(f"\n📊 CSV ANALYSIS:")
        trigger_counts = defaultdict(int)
        target_counts = defaultdict(int)
        study_counts = defaultdict(int)
        
        with open(output_file, 'r') as f:
            reader = csv.DictReader(f)
            for row in reader:
                study = row['study_name']
                trigger = row['trigger']
                target = row['gene_target']
                item_type = row['item_type']
                
                # Count triggers per study
                study_counts[study] += 1
                
                # Count unique triggers per study
                key = f"{study}_{trigger}"
                trigger_counts[key] += 1
                
                # Count unique targets per study
                target_key = f"{study}_{target}_{item_type}"
                target_counts[target_key] += 1
        
        # Print analysis
        print("\n📈 STUDY STATISTICS:")
        for study in study_counts:
            print(f"\n  Study: {study}")
            print(f"    Total rows: {study_counts[study]}")
            
            # Count unique triggers
            study_triggers = {k.split('_', 1)[1] for k in trigger_counts.keys() if k.startswith(f"{study}_")}
            print(f"    Unique triggers: {len(study_triggers)}")
            
            # Count unique targets
            study_targets = {k.split('_', 1)[1] for k in target_counts.keys() if k.startswith(f"{study}_")}
            print(f"    Unique targets: {len(study_targets)}")
            
            # Show trigger repetition
            print(f"    Trigger repetition:")
            for trigger in study_triggers:
                count = trigger_counts[f"{study}_{trigger}"]
                print(f"      - {trigger}: {count} times")
            
            # Show target repetition
            print(f"    Target repetition:")
            for target in study_targets:
                count = target_counts[f"{study}_{target}"]
                print(f"      - {target}: {count} times")

if __name__ == "__main__":
    test_problematic_studies() 