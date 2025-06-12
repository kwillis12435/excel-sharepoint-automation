#!/usr/bin/env python3
"""
Test script to diagnose specific studies that should be processing but aren't appearing in CSV.
"""

import os
import sys
sys.path.append('.')
from process_study import process_study_folder, Config

def test_specific_studies():
    """Test specific studies that are known to exist but missing from CSV."""
    
    base_folder = Config.MONTH_FOLDER
    
    # Studies that exist but are missing from CSV based on the comparison
    missing_studies = [
        "2024013001 (hACVR2b_SEAP_KD_1)",
        "2024013101 (hACVR2b_SEAP_KD_2)", 
        "2024012501 (hIL33_AAV8_KD_4)",
        "2024012601 (hIL33_AAV8_KD_5)",
        "2024013102 (hIL33_AAV8_KD_6)"
    ]
    
    print("Testing specific studies that should be processed but are missing from CSV...")
    print("=" * 80)
    
    for study_name in missing_studies:
        study_folder = os.path.join(base_folder, study_name)
        
        print(f"\n🔍 Testing: {study_name}")
        print("-" * 50)
        
        if not os.path.exists(study_folder):
            print(f"❌ Folder not found: {study_folder}")
            continue
            
        print(f"✅ Folder exists: {study_folder}")
        
        # Check file structure
        info_file = os.path.join(study_folder, f"{study_name}.xlsm")
        results_folder = os.path.join(study_folder, "Results")
        
        print(f"📁 Info file exists: {os.path.exists(info_file)}")
        print(f"📁 Results folder exists: {os.path.exists(results_folder)}")
        
        if os.path.exists(results_folder):
            results_files = [f for f in os.listdir(results_folder) if f.endswith(('.xlsm', '.xlsx'))]
            print(f"📁 Results files found: {len(results_files)}")
            if results_files:
                print(f"   - {results_files[0]}")
        
        # Try to process the study
        print("🔄 Attempting to process study...")
        try:
            study_data = process_study_folder(study_folder)
            
            if study_data is None:
                print("❌ process_study_folder returned None")
                continue
                
            print("✅ Study processing succeeded")
            
            # Check what data was extracted
            has_metadata = any(k in study_data for k in ['study_name', 'study_code', 'tissues'])
            has_rel_exp = 'relative_expression' in study_data
            
            print(f"📊 Has metadata: {has_metadata}")
            print(f"📊 Has relative expression: {has_rel_exp}")
            
            if has_metadata:
                print(f"   - Study name: {study_data.get('study_name', 'None')}")
                print(f"   - Study code: {study_data.get('study_code', 'None')}")
                print(f"   - Tissues: {len(study_data.get('tissues', []))} found")
            
            if has_rel_exp:
                rel_exp = study_data['relative_expression']
                targets = rel_exp.get('targets', [])
                tissue_targets = rel_exp.get('tissue_targets', [])
                rel_exp_data = rel_exp.get('relative_expression_data', {})
                
                print(f"   - Gene targets: {len(targets)}")
                print(f"   - Tissue targets: {len(tissue_targets)}")
                print(f"   - Triggers with data: {len(rel_exp_data)}")
                
                if targets:
                    print(f"   - Example gene targets: {targets[:3]}")
                if tissue_targets:
                    print(f"   - Example tissue targets: {tissue_targets[:3]}")
                if rel_exp_data:
                    first_trigger = next(iter(rel_exp_data))
                    first_trigger_data = rel_exp_data[first_trigger]
                    print(f"   - Example trigger: '{first_trigger}' -> {len(first_trigger_data)} items")
            
            # Check if this would be included in CSV
            would_be_included = has_rel_exp and rel_exp_data
            print(f"🎯 Would be included in CSV: {would_be_included}")
            
            if not would_be_included:
                print("❓ Reason for exclusion:")
                if not has_rel_exp:
                    print("   - No relative expression data found")
                elif not rel_exp_data:
                    print("   - Relative expression data structure is empty")
                    
        except Exception as e:
            print(f"❌ Exception during processing: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    test_specific_studies() 