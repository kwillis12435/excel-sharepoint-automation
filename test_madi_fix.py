#!/usr/bin/env python3
"""
Test script to verify the mAdi_110_HDAC11_1 study fix.
"""

import os
import sys
sys.path.append('.')

from process_study import process_study_folder, init_logger

def test_madi_fix():
    """Test the specific study that was problematic."""
    
    # Initialize logging
    init_logger("test_madi_fix.log")
    
    # Path to the specific study
    study_code = "2024011202"
    base_path = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    study_folder = None
    
    # Find the study folder
    for folder_name in os.listdir(base_path):
        if folder_name.startswith(study_code):
            study_folder = os.path.join(base_path, folder_name)
            print(f"Found study folder: {folder_name}")
            break
    
    if not study_folder:
        print(f"Study folder not found for code: {study_code}")
        return
    
    print(f"\n🔍 Testing fix for study: {os.path.basename(study_folder)}")
    print("="*60)
    
    try:
        study_data = process_study_folder(study_folder)
        
        if study_data:
            print(f"\n📊 STUDY DATA SUMMARY:")
            print(f"  Study name: {study_data.get('study_name')}")
            print(f"  Study code: {study_data.get('study_code')}")
            
            # Check metadata triggers
            trigger_dose_map = study_data.get('trigger_dose_map', {})
            print(f"\n📋 METADATA TRIGGERS: {len(trigger_dose_map)} triggers")
            
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
                
            else:
                print(f"\n❌ NO RELATIVE EXPRESSION DATA FOUND")
        else:
            print(f"\n❌ STUDY PROCESSING FAILED")
            
    except Exception as e:
        print(f"\n💥 EXCEPTION: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_madi_fix() 