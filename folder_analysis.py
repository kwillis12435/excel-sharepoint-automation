#!/usr/bin/env python3
"""
Diagnostic script to analyze study folder structure and identify processing issues.
"""

import os
from pathlib import Path

def analyze_study_folders():
    """Analyze all study folders to identify potential processing issues."""
    
    base_folder = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\01 - 2024"
    
    if not os.path.exists(base_folder):
        print(f"Base folder not found: {base_folder}")
        return
    
    print(f"Analyzing study folders in: {base_folder}")
    print("=" * 80)
    
    # Get all subdirectories
    study_folders = [
        os.path.join(base_folder, name)
        for name in os.listdir(base_folder)
        if os.path.isdir(os.path.join(base_folder, name))
    ]
    
    print(f"Found {len(study_folders)} total directories")
    print()
    
    # Analyze each folder
    issues = {
        "missing_main_file": [],
        "missing_results_folder": [],
        "missing_results_file": [],
        "wrong_file_extension": [],
        "complete_studies": []
    }
    
    for study_folder in study_folders:
        folder_name = os.path.basename(study_folder)
        
        # Check for main study file
        expected_main_file = os.path.join(study_folder, f"{folder_name}.xlsm")
        xlsx_alternative = os.path.join(study_folder, f"{folder_name}.xlsx")
        
        has_main_file = os.path.exists(expected_main_file)
        has_xlsx_alternative = os.path.exists(xlsx_alternative)
        
        # Check for Results folder and file
        results_folder = os.path.join(study_folder, "Results")
        has_results_folder = os.path.exists(results_folder)
        
        results_file = None
        if has_results_folder:
            # Look for any .xlsm or .xlsx file in Results folder
            try:
                for file in os.listdir(results_folder):
                    if file.endswith(('.xlsm', '.xlsx')):
                        results_file = os.path.join(results_folder, file)
                        break
            except PermissionError:
                pass
        
        # Categorize the study
        if not has_main_file and not has_xlsx_alternative:
            issues["missing_main_file"].append({
                "folder": folder_name,
                "expected": f"{folder_name}.xlsm",
                "has_xlsx": has_xlsx_alternative
            })
        elif not has_main_file and has_xlsx_alternative:
            issues["wrong_file_extension"].append({
                "folder": folder_name,
                "expected": f"{folder_name}.xlsm",
                "found": f"{folder_name}.xlsx"
            })
        elif not has_results_folder:
            issues["missing_results_folder"].append({
                "folder": folder_name,
                "has_main": has_main_file
            })
        elif not results_file:
            issues["missing_results_file"].append({
                "folder": folder_name,
                "has_main": has_main_file,
                "has_results_folder": has_results_folder
            })
        else:
            issues["complete_studies"].append({
                "folder": folder_name,
                "main_file": f"{folder_name}.xlsm" if has_main_file else f"{folder_name}.xlsx",
                "results_file": os.path.basename(results_file)
            })
    
    # Print summary
    print("ANALYSIS SUMMARY:")
    print("=" * 50)
    print(f"Complete studies (should process): {len(issues['complete_studies'])}")
    print(f"Missing main file: {len(issues['missing_main_file'])}")
    print(f"Wrong file extension (.xlsx instead of .xlsm): {len(issues['wrong_file_extension'])}")
    print(f"Missing Results folder: {len(issues['missing_results_folder'])}")
    print(f"Missing results file: {len(issues['missing_results_file'])}")
    print()
    
    # Print details for each category
    if issues["complete_studies"]:
        print("COMPLETE STUDIES (should be processed):")
        print("-" * 40)
        for study in sorted(issues["complete_studies"], key=lambda x: x["folder"]):
            print(f"  ✓ {study['folder']}")
            print(f"    Main: {study['main_file']}")
            print(f"    Results: {study['results_file']}")
        print()
    
    if issues["wrong_file_extension"]:
        print("STUDIES WITH .xlsx INSTEAD OF .xlsm:")
        print("-" * 40)
        for study in sorted(issues["wrong_file_extension"], key=lambda x: x["folder"]):
            print(f"  ⚠ {study['folder']}")
            print(f"    Expected: {study['expected']}")
            print(f"    Found: {study['found']}")
        print()
    
    if issues["missing_main_file"]:
        print("STUDIES MISSING MAIN FILE:")
        print("-" * 40)
        for study in sorted(issues["missing_main_file"], key=lambda x: x["folder"]):
            print(f"  ✗ {study['folder']}")
            print(f"    Expected: {study['expected']}")
        print()
    
    if issues["missing_results_folder"]:
        print("STUDIES MISSING RESULTS FOLDER:")
        print("-" * 40)
        for study in sorted(issues["missing_results_folder"], key=lambda x: x["folder"]):
            print(f"  ✗ {study['folder']}")
        print()
    
    if issues["missing_results_file"]:
        print("STUDIES MISSING RESULTS FILE:")
        print("-" * 40)
        for study in sorted(issues["missing_results_file"], key=lambda x: x["folder"]):
            print(f"  ✗ {study['folder']}")
        print()
    
    # Check specific studies mentioned in comparison
    print("CHECKING SPECIFIC STUDIES FROM COMPARISON:")
    print("-" * 50)
    missing_studies = [
        "hACVR2b_SEAP_KD_1", "hACVR2b_SEAP_KD_2", "hIL33_AAV8_KD_4", 
        "hIL33_AAV8_KD_5", "hIL33_AAV8_KD_6", "hmSNCA_002_ICV",
        "hPDE3B_AAV8_LO_1", "hPDE3B_AAV8_LO_2", "mGPR75_001_ICV"
    ]
    
    for study_name in missing_studies:
        found = False
        for study_folder in study_folders:
            folder_name = os.path.basename(study_folder)
            if study_name in folder_name or folder_name.startswith(study_name):
                print(f"  ✓ Found folder for {study_name}: {folder_name}")
                found = True
                break
        if not found:
            print(f"  ✗ No folder found for: {study_name}")

if __name__ == "__main__":
    analyze_study_folders() 