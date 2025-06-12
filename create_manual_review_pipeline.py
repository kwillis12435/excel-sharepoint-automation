#!/usr/bin/env python3
"""
Manual Review Pipeline for Study Data Processing

This script creates a comprehensive review system to identify studies that need manual attention
and categorizes them by the type of issue that needs to be resolved.
"""

import os
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple, Any
import json
from datetime import datetime
import csv
from process_study import Config, process_study_folder

class ManualReviewPipeline:
    """
    Pipeline to identify and categorize studies that need manual review.
    """
    
    def __init__(self, experimental_csv: str, manual_excel: str):
        """Initialize with paths to experimental and manual datasets."""
        self.experimental_csv = experimental_csv
        self.manual_excel = manual_excel
        self.experimental_df = None
        self.manual_df = None
        self.issues = {
            'missing_folders': [],           # Studies in manual but no folder found
            'name_mismatches': [],           # Studies with potential folder name issues  
            'missing_rel_exp': [],           # Folder exists, metadata OK, but no rel exp data
            'partial_data': [],              # Some data but significantly less than manual
            'sheet_format_issues': [],       # Calculation sheets or format problems
            'trigger_filtering_issues': [],  # Valid triggers being filtered out
            'exact_matches': []              # Studies working perfectly (for validation)
        }
        
    def load_datasets(self):
        """Load both datasets for comparison."""
        print("📊 Loading datasets...")
        
        # Load experimental data (CSV)
        self.experimental_df = pd.read_csv(self.experimental_csv)
        print(f"   ✅ Experimental CSV: {len(self.experimental_df)} rows")
        
        # Load manual data (Excel)
        self.manual_df = pd.read_excel(self.manual_excel)
        print(f"   ✅ Manual Excel: {len(self.manual_df)} rows")
        
    def analyze_study_discrepancies(self):
        """Analyze discrepancies between datasets and categorize issues."""
        print("\n🔍 Analyzing study discrepancies...")
        
        # Group by study name
        exp_studies = self.experimental_df.groupby('study_name').size().to_dict()
        manual_studies = self.manual_df.groupby('study_name').size().to_dict()
        
        all_studies = set(exp_studies.keys()) | set(manual_studies.keys())
        
        print(f"   Found {len(exp_studies)} studies in experimental data")
        print(f"   Found {len(manual_studies)} studies in manual data")
        print(f"   Total unique studies: {len(all_studies)}")
        
        for study_name in sorted(all_studies):
            exp_rows = exp_studies.get(study_name, 0)
            manual_rows = manual_studies.get(study_name, 0)
            
            self._categorize_study_issue(study_name, exp_rows, manual_rows)
    
    def _categorize_study_issue(self, study_name: str, exp_rows: int, manual_rows: int):
        """Categorize a single study's issue type."""
        
        if exp_rows == manual_rows and exp_rows > 0:
            # Perfect match
            self.issues['exact_matches'].append({
                'study_name': study_name,
                'rows': exp_rows,
                'status': 'Perfect match'
            })
            
        elif exp_rows == 0 and manual_rows > 0:
            # Missing from experimental - need to determine why
            folder_status = self._check_folder_exists(study_name)
            
            if not folder_status['exists']:
                if folder_status['similar_folders']:
                    self.issues['name_mismatches'].append({
                        'study_name': study_name,
                        'manual_rows': manual_rows,
                        'similar_folders': folder_status['similar_folders'],
                        'status': 'Potential folder name mismatch'
                    })
                else:
                    self.issues['missing_folders'].append({
                        'study_name': study_name,
                        'manual_rows': manual_rows,
                        'status': 'No folder found'
                    })
            else:
                # Folder exists but no data extracted - investigate why
                investigation = self._investigate_folder_processing(folder_status['folder_path'])
                
                if investigation['has_metadata'] and not investigation['has_rel_exp']:
                    self.issues['missing_rel_exp'].append({
                        'study_name': study_name,
                        'manual_rows': manual_rows,
                        'folder_path': folder_status['folder_path'],
                        'investigation': investigation,
                        'status': 'Folder exists, metadata OK, missing relative expression data'
                    })
                elif investigation['error']:
                    self.issues['sheet_format_issues'].append({
                        'study_name': study_name,
                        'manual_rows': manual_rows,
                        'folder_path': folder_status['folder_path'],
                        'error': investigation['error'],
                        'status': 'Processing error or sheet format issue'
                    })
                    
        elif exp_rows > 0 and manual_rows > exp_rows:
            # Partial data - getting some but not all
            diff_pct = ((manual_rows - exp_rows) / manual_rows) * 100
            
            if diff_pct > 50:  # Missing more than 50% of expected data
                self.issues['partial_data'].append({
                    'study_name': study_name,
                    'exp_rows': exp_rows,
                    'manual_rows': manual_rows,
                    'missing_pct': diff_pct,
                    'status': f'Missing {diff_pct:.1f}% of expected data'
                })
            else:
                # Minor differences - might be trigger filtering issues
                self.issues['trigger_filtering_issues'].append({
                    'study_name': study_name,
                    'exp_rows': exp_rows,
                    'manual_rows': manual_rows,
                    'missing_pct': diff_pct,
                    'status': f'Minor data loss ({diff_pct:.1f}%), possibly trigger filtering'
                })
    
    def _check_folder_exists(self, study_name: str) -> Dict[str, Any]:
        """Check if a folder exists for the study and find similar folders."""
        base_folder = Config.MONTH_FOLDER
        all_folders = [f for f in os.listdir(base_folder) if os.path.isdir(os.path.join(base_folder, f))]
        
        # Check for exact match
        for folder in all_folders:
            if study_name.lower() in folder.lower():
                return {
                    'exists': True,
                    'folder_path': os.path.join(base_folder, folder),
                    'folder_name': folder,
                    'similar_folders': []
                }
        
        # Look for similar folders
        similar_folders = []
        study_words = set(study_name.lower().replace('_', ' ').split())
        
        for folder in all_folders:
            folder_words = set(folder.lower().replace('_', ' ').replace('(', ' ').replace(')', ' ').split())
            
            # Remove numbers and common words
            study_words_clean = {w for w in study_words if len(w) > 2 and not w.isdigit()}
            folder_words_clean = {w for w in folder_words if len(w) > 2 and not w.isdigit()}
            
            if study_words_clean and folder_words_clean:
                overlap = len(study_words_clean & folder_words_clean)
                total_words = len(study_words_clean)
                
                if overlap >= max(1, total_words - 1):  # Allow for 1 word difference
                    score = overlap / total_words
                    similar_folders.append((folder, score))
        
        # Sort by similarity score
        similar_folders.sort(key=lambda x: x[1], reverse=True)
        
        return {
            'exists': False,
            'folder_path': None,
            'folder_name': None,
            'similar_folders': [f[0] for f in similar_folders[:3]]  # Top 3 matches
        }
    
    def _investigate_folder_processing(self, folder_path: str) -> Dict[str, Any]:
        """Investigate why a folder isn't producing relative expression data."""
        try:
            # Try to process the study
            study_data = process_study_folder(folder_path)
            
            if not study_data:
                return {'error': 'process_study_folder returned None', 'has_metadata': False, 'has_rel_exp': False}
            
            has_metadata = any(key in study_data for key in ['study_name', 'study_code', 'tissues', 'trigger_dose_map'])
            has_rel_exp = 'relative_expression' in study_data
            
            result = {
                'error': None,
                'has_metadata': has_metadata,
                'has_rel_exp': has_rel_exp,
                'study_keys': list(study_data.keys())
            }
            
            if has_rel_exp:
                rel_exp_data = study_data['relative_expression']
                result['rel_exp_details'] = {
                    'has_targets': bool(rel_exp_data.get('targets')),
                    'has_data': bool(rel_exp_data.get('relative_expression_data')),
                    'num_targets': len(rel_exp_data.get('targets', [])),
                    'num_triggers': len(rel_exp_data.get('relative_expression_data', {}))
                }
            
            return result
            
        except Exception as e:
            return {'error': str(e), 'has_metadata': False, 'has_rel_exp': False}
    
    def print_summary(self):
        """Print a summary of findings."""
        print("\n📊 PIPELINE SUMMARY")
        print("="*40)
        
        total_issues = sum(len(issues) for issues in self.issues.values()) - len(self.issues['exact_matches'])
        
        print(f"✅ Exact matches: {len(self.issues['exact_matches'])}")
        print(f"⚠️  Issues needing review: {total_issues}")
        print()
        
        for issue_type, issues in self.issues.items():
            if issues and issue_type != 'exact_matches':
                print(f"   {issue_type.replace('_', ' ').title()}: {len(issues)}")

def main():
    """Main function to run the manual review pipeline."""
    # File paths - update these to match your actual file locations
    experimental_csv = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\study_data_01_20250612.csv"
    manual_excel = r"C:\Users\kwillis\OneDrive - Arrowhead Pharmaceuticals Inc\Discovery Biology - 2024\LO_study_data (1).xlsx"
    
    # Create and run pipeline
    pipeline = ManualReviewPipeline(experimental_csv, manual_excel)
    pipeline.load_datasets()
    pipeline.analyze_study_discrepancies()
    pipeline.print_summary()
    
    print(f"\n🎯 NEXT STEPS:")
    print(f"1. Focus on 'name_mismatches' first - these are likely quick wins")
    print(f"2. Investigate 'missing_rel_exp' - folders exist but no data extracted")
    print(f"3. Review 'partial_data' for trigger filtering issues")

if __name__ == "__main__":
    main() 