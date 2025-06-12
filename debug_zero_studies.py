#!/usr/bin/env python3
"""
Debug why specific studies are returning 0 rows when they should have data.
"""

import os
import sys
sys.path.append('.')
from process_study import process_study_folder, Config

def debug_zero_row_studies():
    """Debug the studies that show 0 rows in experimental but many in manual."""
    
    base_folder = Config.MONTH_FOLDER
    
    # Focus on the biggest mismatches that show 0 in experimental
    zero_row_studies = [
        "hmSNCA_002_ICV",           # 0 vs 72
        "hACVR2b_SEAP_KD_1",        # 0 vs 60  
        "hPDE3B_AAV8_LO_1",         # 0 vs 33
        "hPDE3B_AAV8_LO_2",         # 0 vs 33
        "mGPR75_001_ICV",           # 0 vs 39
    ]
    
    print("🔍 DEBUGGING ZERO-ROW STUDIES")
    print("="*60)
    
    for study_name in zero_row_studies:
        print(f"\n{'='*50}")
        print(f"🔍 DEBUGGING: {study_name}")
        print(f"{'='*50}")
        
        # Look for the study folder (try different patterns)
        possible_folders = []
        
        # Check all folders that might match
        for folder in os.listdir(base_folder):
            if study_name.lower() in folder.lower():
                possible_folders.append(folder)
        
        if not possible_folders:
            print(f"❌ No folder found containing '{study_name}'")
            
            # List all folders to see what's available
            print(f"📁 Available folders containing similar names:")
            all_folders = os.listdir(base_folder)
            
            # Look for partial matches
            words = study_name.lower().split('_')
            for folder in all_folders:
                folder_lower = folder.lower()
                if any(word in folder_lower for word in words if len(word) > 2):
                    print(f"   - {folder}")
            continue
        
        print(f"📁 Found {len(possible_folders)} possible folder(s):")
        for folder in possible_folders:
            print(f"   - {folder}")
        
        # Process the first matching folder
        study_folder = os.path.join(base_folder, possible_folders[0])
        
        try:
            print(f"\n🔍 Processing: {possible_folders[0]}")
            result = process_study_folder(study_folder)
            
            if not result:
                print("❌ Study failed to process - no data returned")
                continue
            
            print(f"✅ Study processed successfully")
            print(f"   Study name: {result.get('study_name', 'Unknown')}")
            
            # Check what data was found
            has_metadata = bool(result.get('study_name') or result.get('study_code'))
            has_rel_exp = 'relative_expression' in result
            
            print(f"   Has metadata: {has_metadata}")
            print(f"   Has relative expression: {has_rel_exp}")
            
            if has_rel_exp:
                rel_exp = result['relative_expression']
                rel_exp_data = rel_exp.get('relative_expression_data', {})
                targets = rel_exp.get('targets', [])
                tissue_targets = rel_exp.get('tissue_targets', [])
                
                print(f"   Gene targets: {len(targets)} - {targets}")
                print(f"   Tissue targets: {len(tissue_targets)} - {tissue_targets}")
                print(f"   Triggers with data: {len(rel_exp_data)}")
                
                if rel_exp_data:
                    print(f"   Triggers found: {list(rel_exp_data.keys())}")
                    
                    # Count total data points
                    total_points = sum(len(trigger_data) for trigger_data in rel_exp_data.values())
                    print(f"   Total data points: {total_points}")
                else:
                    print("   ❌ No relative expression data found")
            else:
                print("   ❌ No relative expression data structure")
                
        except Exception as e:
            print(f"❌ Exception processing study: {e}")
            import traceback
            traceback.print_exc()

if __name__ == "__main__":
    debug_zero_row_studies() 