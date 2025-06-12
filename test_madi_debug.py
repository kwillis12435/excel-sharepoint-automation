#!/usr/bin/env python3
"""
Debug script for mAdi_110_HDAC11_1 study to understand why it's not capturing data properly.
"""

import os
import sys
from process_study import process_study_folder, Config

def debug_madi_study():
    """Debug the specific mAdi_110_HDAC11_1 study."""
    
    # Find the study folder
    month_folder = Config.MONTH_FOLDER
    study_name = "2024011202 (mAdi_110_HDAC11_1)"
    
    study_folder = os.path.join(month_folder, study_name)
    
    if not os.path.exists(study_folder):
        print(f"❌ Study folder not found: {study_folder}")
        # Try to find any folder with mAdi_110_HDAC11_1 in the name
        for folder in os.listdir(month_folder):
            if "mAdi_110_HDAC11_1" in folder:
                study_folder = os.path.join(month_folder, folder)
                print(f"✅ Found study folder: {study_folder}")
                break
        else:
            print("❌ Could not find any folder containing 'mAdi_110_HDAC11_1'")
            return
    
    print(f"🔍 Debugging study: {study_folder}")
    print("="*60)
    
    # Process the study with debug enabled
    Config.DEBUG = True
    
    try:
        study_data = process_study_folder(study_folder)
        
        if study_data:
            print(f"\n📊 STUDY DATA SUMMARY:")
            print(f"  Study name: {study_data.get('study_name')}")
            print(f"  Study code: {study_data.get('study_code')}")
            print(f"  Tissues: {study_data.get('tissues')}")
            print(f"  Trigger dose map: {len(study_data.get('trigger_dose_map', {}))} triggers")
            
            # Show trigger dose map details
            trigger_dose_map = study_data.get('trigger_dose_map', {})
            print(f"\n📋 METADATA TRIGGERS:")
            for i, (trigger, dose_info) in enumerate(trigger_dose_map.items(), 1):
                print(f"  {i}: '{trigger}' -> {dose_info}")
            
            # Show relative expression data
            rel_exp_data = study_data.get('relative_expression', {})
            if rel_exp_data:
                print(f"\n🧬 RELATIVE EXPRESSION DATA:")
                print(f"  Targets: {rel_exp_data.get('targets', [])}")
                print(f"  Tissue targets: {rel_exp_data.get('tissue_targets', [])}")
                print(f"  Triggers with data: {list(rel_exp_data.get('relative_expression_data', {}).keys())}")
                
                # Show raw triggers found for debugging
                if 'raw_triggers_found' in rel_exp_data:
                    print(f"  Raw triggers found in results sheet: {rel_exp_data['raw_triggers_found']}")
                
                # Count total data points
                triggers_data = rel_exp_data.get('relative_expression_data', {})
                total_data_points = 0
                for trigger_data in triggers_data.values():
                    total_data_points += len(trigger_data)
                print(f"  Total data points: {total_data_points}")
            else:
                print(f"\n❌ NO RELATIVE EXPRESSION DATA FOUND")
        else:
            print(f"\n❌ NO STUDY DATA EXTRACTED")
            
    except Exception as e:
        print(f"\n💥 ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_madi_study() 