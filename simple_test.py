#!/usr/bin/env python3
"""
Simple test to verify the new extract_relative_expression_data function is working
"""

import sys
import os
import glob
sys.path.append(os.getcwd())

from process_study import extract_relative_expression_data, safe_workbook_operation

def test_function_call():
    """Test if the new function is being called correctly"""
    
    # Find any study folder to test with
    base_folder = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    
    if not os.path.exists(base_folder):
        print(f"❌ Base folder not found: {base_folder}")
        return
    
    # Find the first study folder with a Results directory
    study_folders = [f for f in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, f))]
    
    for folder_name in study_folders[:3]:  # Test first 3 folders
        study_folder = os.path.join(base_folder, folder_name)
        results_folder = os.path.join(study_folder, "Results")
        
        if os.path.exists(results_folder):
            # Find any .xlsm file in Results
            xlsm_files = glob.glob(os.path.join(results_folder, "*.xlsm"))
            if xlsm_files:
                results_file = xlsm_files[0]
                print(f"🔍 Testing with: {folder_name}")
                print(f"   Results file: {os.path.basename(results_file)}")
                
                # Test the function directly
                try:
                    result = safe_workbook_operation(
                        results_file,
                        extract_relative_expression_data,
                        []  # empty procedure_tissues
                    )
                    
                    if result:
                        print(f"✅ Function returned data: {len(result.get('targets', []))} targets")
                        print(f"   Triggers found: {len(result.get('relative_expression_data', {}))}")
                        return  # Success, stop testing
                    else:
                        print("❌ Function returned None")
                        
                except Exception as e:
                    print(f"❌ Error calling function: {e}")
                    import traceback
                    traceback.print_exc()
                
                print()  # Add spacing between tests
    
    print("❌ No suitable test files found")

if __name__ == "__main__":
    test_function_call() 