#!/usr/bin/env python3
"""
Find potential folder matches for studies that appear in manual data but not in experimental results.
"""

import os
import re
from difflib import SequenceMatcher
from process_study import Config

def similarity(a, b):
    """Calculate similarity between two strings."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def find_missing_studies():
    """Find potential matches for missing studies."""
    
    base_folder = Config.MONTH_FOLDER
    
    # Studies that show up in manual data but are missing from experimental (0 rows)
    missing_studies = [
        'hmSNCA_002_ICV',      # 72 rows in manual
        'hACVR2b_SEAP_KD_1',   # 60 rows in manual  
        'hPDE3B_AAV8_LO_1',    # 33 rows in manual
        'hPDE3B_AAV8_LO_2',    # 33 rows in manual
        'hIL33_AAV8_KD_6',     # 45 rows in manual
        'mGPR75_001_ICV',      # 39 rows in manual
        'mAlk7_Pharma_4',      # 36 rows in manual
        'mG3TD_67',            # 26 rows in manual
        'rMac_Aif1_4',         # 22 rows in manual
        'rINHBA_liver_02',     # 18 rows in manual
        'mSNCA_02_ICV',        # 12 rows in manual
        'mF12_Phos_38',        # 11 rows in manual (case issue?)
        'mG3TD_065',           # 10 rows in manual
        'mG3TD_066',           # 8 rows in manual
        'rSQL_60',             # 6 rows in manual
    ]
    
    # Get all available folders
    all_folders = [f for f in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, f))]
    
    print("🔍 FINDING MISSING STUDIES")
    print("="*80)
    print(f"Looking for {len(missing_studies)} missing studies in {len(all_folders)} folders")
    print()
    
    potential_matches = {}
    
    for missing_study in missing_studies:
        print(f"📋 Missing: {missing_study}")
        
        # Try different matching strategies
        matches = []
        
        # Strategy 1: Exact case-insensitive substring match
        for folder in all_folders:
            if missing_study.lower() in folder.lower():
                matches.append((folder, 1.0, "exact_substring"))
        
        # Strategy 2: Remove common differences and match
        normalized_missing = re.sub(r'[_\-\s]', '', missing_study.lower())
        for folder in all_folders:
            normalized_folder = re.sub(r'[_\-\s]', '', folder.lower())
            if normalized_missing in normalized_folder or normalized_folder in normalized_missing:
                similarity_score = similarity(normalized_missing, normalized_folder)
                if similarity_score > 0.6:
                    matches.append((folder, similarity_score, "normalized"))
        
        # Strategy 3: Split into words and check for significant overlap
        missing_words = set(missing_study.lower().split('_'))
        for folder in all_folders:
            folder_words = set(re.split(r'[_\-\s\(\)]', folder.lower()))
            
            # Remove common words and numbers
            missing_words_clean = {w for w in missing_words if len(w) > 2 and not w.isdigit()}
            folder_words_clean = {w for w in folder_words if len(w) > 2 and not w.isdigit()}
            
            if missing_words_clean and folder_words_clean:
                overlap = len(missing_words_clean & folder_words_clean)
                total_words = len(missing_words_clean)
                if overlap >= max(1, total_words - 1):  # Allow for 1 word difference
                    score = overlap / total_words
                    matches.append((folder, score, "word_overlap"))
        
        # Remove duplicates and sort by score
        unique_matches = {}
        for folder, score, method in matches:
            if folder not in unique_matches or unique_matches[folder][0] < score:
                unique_matches[folder] = (score, method)
        
        sorted_matches = sorted(unique_matches.items(), key=lambda x: x[1][0], reverse=True)
        
        if sorted_matches:
            print(f"   ✅ Found {len(sorted_matches)} potential matches:")
            for folder, (score, method) in sorted_matches[:5]:  # Show top 5
                print(f"      {score:.2f} - {folder} ({method})")
            potential_matches[missing_study] = sorted_matches[0][0]  # Best match
        else:
            print(f"   ❌ No potential matches found")
        
        print()
    
    # Summary of recommendations
    print("📋 RECOMMENDED FOLDER MAPPINGS:")
    print("="*50)
    for missing_study, best_folder in potential_matches.items():
        print(f"'{missing_study}': '{best_folder}',")
    
    print(f"\n✅ Found potential matches for {len(potential_matches)}/{len(missing_studies)} missing studies")
    
    return potential_matches

if __name__ == "__main__":
    find_missing_studies() 