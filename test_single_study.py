#!/usr/bin/env python3
"""
Test script to debug the specific mF12_phos_38 study that's causing hanging.
"""

import os
import sys
sys.path.append('.')
from process_study import process_study_folder, Config

def test_mf12_phos_38():
    """Test just the problematic mF12_phos_38 study."""
    
    base_folder = Config.MONTH_FOLDER
    study_folder = os.path.join(base_folder, "2024010506 (mF12_phos_38)")
    
    if not os.path.exists(study_folder):
        print(f"Study folder not found: {study_folder}")
        return
    
    print(f"Testing single study: {study_folder}")
    print("=" * 60)
    
    try:
        result = process_study_folder(study_folder)
        if result:
            print("✅ SUCCESS - Study processed successfully")
            print(f"Study name: {result.get('study_name', 'Unknown')}")
            print(f"Study code: {result.get('study_code', 'Unknown')}")
            if 'relative_expression' in result:
                rel_exp = result['relative_expression']
                print(f"Relative expression data found:")
                print(f"  - Targets: {len(rel_exp.get('targets', []))}")
                print(f"  - Triggers: {len(rel_exp.get('relative_expression_data', {}))}")
            else:
                print("No relative expression data found")
        else:
            print("❌ FAILED - No data returned")
    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_mf12_phos_38() 