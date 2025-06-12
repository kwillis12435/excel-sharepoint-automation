#!/usr/bin/env python3
"""
Study Data Analysis Script

This script analyzes CSV or Excel files containing study data and provides comprehensive
statistics including:
- Number of rows per study
- Unique targets per study
- Unique triggers per study
- Data completeness metrics
- Comparison capabilities between datasets

Usage:
    python analyze_study_data.py <file_path> [--output <output_file>]
    
Example:
    python analyze_study_data.py experimental_data.csv
    python analyze_study_data.py manual_data.xlsx --output analysis_report.txt
    python analyze_study_data.py experimental_data.csv manual_data.xlsx --compare
"""

import pandas as pd
import argparse
import sys
from pathlib import Path
from typing import Dict, List, Tuple, Any, Optional
import json
from datetime import datetime

class StudyDataAnalyzer:
    """
    Comprehensive analyzer for study data files (CSV or Excel).
    Provides detailed statistics and comparison capabilities.
    """
    
    def __init__(self, file_path: str):
        """Initialize analyzer with a data file."""
        self.file_path = Path(file_path)
        self.df = None
        self.stats = {}
        self.load_data()
    
    def load_data(self):
        """Load data from CSV or Excel file."""
        try:
            if self.file_path.suffix.lower() == '.csv':
                self.df = pd.read_csv(self.file_path)
                print(f"✓ Loaded CSV file: {self.file_path}")
            elif self.file_path.suffix.lower() in ['.xlsx', '.xls']:
                self.df = pd.read_excel(self.file_path)
                print(f"✓ Loaded Excel file: {self.file_path}")
            else:
                raise ValueError(f"Unsupported file format: {self.file_path.suffix}")
            
            print(f"  Total rows: {len(self.df)}")
            print(f"  Columns: {list(self.df.columns)}")
            
        except Exception as e:
            print(f"✗ Error loading {self.file_path}: {e}")
            sys.exit(1)
    
    def detect_column_mapping(self) -> Dict[str, str]:
        """
        Automatically detect column mappings to handle different column names.
        Returns mapping of standard names to actual column names.
        """
        columns = [col.lower().strip() for col in self.df.columns]
        mapping = {}
        
        # Study identification columns
        for study_col in ['study_name', 'study', 'name']:
            for col in self.df.columns:
                if study_col in col.lower():
                    mapping['study_name'] = col
                    break
        
        for code_col in ['study_code', 'code', 'study_id']:
            for col in self.df.columns:
                if code_col in col.lower():
                    mapping['study_code'] = col
                    break
        
        # Target columns
        for target_col in ['gene_target', 'target', 'gene']:
            for col in self.df.columns:
                if target_col in col.lower():
                    mapping['target'] = col
                    break
        
        # Trigger columns
        for trigger_col in ['trigger', 'treatment', 'compound']:
            for col in self.df.columns:
                if trigger_col in col.lower():
                    mapping['trigger'] = col
                    break
        
        # Tissue columns
        for tissue_col in ['tissue', 'organ']:
            for col in self.df.columns:
                if tissue_col in col.lower():
                    mapping['tissue'] = col
                    break
        
        # Item type columns
        for type_col in ['item_type', 'type', 'category']:
            for col in self.df.columns:
                if type_col in col.lower():
                    mapping['item_type'] = col
                    break
        
        # Expression data columns
        for expr_col in ['avg_rel_exp', 'rel_exp', 'expression', 'value']:
            for col in self.df.columns:
                if expr_col in col.lower() and 'avg' in col.lower():
                    mapping['expression'] = col
                    break
        
        print(f"  Detected column mapping: {mapping}")
        return mapping
    
    def analyze(self) -> Dict[str, Any]:
        """
        Perform comprehensive analysis of the study data.
        Returns detailed statistics dictionary.
        """
        print(f"\n📊 Analyzing data from: {self.file_path.name}")
        print("="*60)
        
        column_map = self.detect_column_mapping()
        
        # Basic statistics
        total_rows = len(self.df)
        
        # Study-level analysis
        study_stats = self._analyze_studies(column_map)
        
        # Target analysis
        target_stats = self._analyze_targets(column_map)
        
        # Trigger analysis
        trigger_stats = self._analyze_triggers(column_map)
        
        # Tissue analysis
        tissue_stats = self._analyze_tissues(column_map)
        
        # Data completeness analysis
        completeness_stats = self._analyze_completeness(column_map)
        
        self.stats = {
            'file_info': {
                'file_path': str(self.file_path),
                'file_name': self.file_path.name,
                'total_rows': total_rows,
                'columns': list(self.df.columns),
                'column_mapping': column_map
            },
            'study_stats': study_stats,
            'target_stats': target_stats,
            'trigger_stats': trigger_stats,
            'tissue_stats': tissue_stats,
            'completeness_stats': completeness_stats,
            'analysis_timestamp': datetime.now().isoformat()
        }
        
        return self.stats
    
    def _analyze_studies(self, column_map: Dict[str, str]) -> Dict[str, Any]:
        """Analyze study-level statistics."""
        if 'study_name' not in column_map and 'study_code' not in column_map:
            return {'error': 'No study identifier columns found'}
        
        # Use study_name if available, otherwise study_code
        study_col = column_map.get('study_name') or column_map.get('study_code')
        
        study_groups = self.df.groupby(study_col)
        
        study_details = []
        for study_name, group in study_groups:
            study_info = {
                'name': study_name,
                'total_rows': len(group),
                'unique_targets': 0,
                'unique_triggers': 0,
                'unique_tissues': 0
            }
            
            # Count unique targets
            if 'target' in column_map:
                target_col = column_map['target']
                unique_targets = group[target_col].dropna().nunique()
                study_info['unique_targets'] = unique_targets
                study_info['targets'] = sorted(group[target_col].dropna().unique().tolist())
            
            # Count unique triggers
            if 'trigger' in column_map:
                trigger_col = column_map['trigger']
                unique_triggers = group[trigger_col].dropna().nunique()
                study_info['unique_triggers'] = unique_triggers
                study_info['triggers'] = sorted(group[trigger_col].dropna().unique().tolist())
            
            # Count unique tissues
            if 'tissue' in column_map:
                tissue_col = column_map['tissue']
                unique_tissues = group[tissue_col].dropna().nunique()
                study_info['unique_tissues'] = unique_tissues
                study_info['tissues'] = sorted(group[tissue_col].dropna().unique().tolist())
            
            # Check for item types if available
            if 'item_type' in column_map:
                item_types = group[column_map['item_type']].value_counts().to_dict()
                study_info['item_types'] = item_types
            
            study_details.append(study_info)
        
        # Sort by total rows (descending)
        study_details.sort(key=lambda x: x['total_rows'], reverse=True)
        
        return {
            'total_studies': len(study_details),
            'studies': study_details,
            'total_rows_all_studies': sum(s['total_rows'] for s in study_details)
        }
    
    def _analyze_targets(self, column_map: Dict[str, str]) -> Dict[str, Any]:
        """Analyze target-related statistics."""
        if 'target' not in column_map:
            return {'error': 'No target column found'}
        
        target_col = column_map['target']
        targets = self.df[target_col].dropna()
        
        target_counts = targets.value_counts()
        
        return {
            'total_unique_targets': len(target_counts),
            'total_target_measurements': len(targets),
            'target_frequency': target_counts.to_dict(),
            'top_10_targets': target_counts.head(10).to_dict(),
            'targets_list': sorted(targets.unique().tolist())
        }
    
    def _analyze_triggers(self, column_map: Dict[str, str]) -> Dict[str, Any]:
        """Analyze trigger-related statistics."""
        if 'trigger' not in column_map:
            return {'error': 'No trigger column found'}
        
        trigger_col = column_map['trigger']
        triggers = self.df[trigger_col].dropna()
        
        trigger_counts = triggers.value_counts()
        
        return {
            'total_unique_triggers': len(trigger_counts),
            'total_trigger_measurements': len(triggers),
            'trigger_frequency': trigger_counts.to_dict(),
            'top_10_triggers': trigger_counts.head(10).to_dict(),
            'triggers_list': sorted(triggers.unique().tolist())
        }
    
    def _analyze_tissues(self, column_map: Dict[str, str]) -> Dict[str, Any]:
        """Analyze tissue-related statistics."""
        if 'tissue' not in column_map:
            return {'error': 'No tissue column found'}
        
        tissue_col = column_map['tissue']
        tissues = self.df[tissue_col].dropna()
        
        tissue_counts = tissues.value_counts()
        
        return {
            'total_unique_tissues': len(tissue_counts),
            'total_tissue_measurements': len(tissues),
            'tissue_frequency': tissue_counts.to_dict(),
            'tissues_list': sorted(tissues.unique().tolist())
        }
    
    def _analyze_completeness(self, column_map: Dict[str, str]) -> Dict[str, Any]:
        """Analyze data completeness."""
        completeness = {}
        
        for col_name, actual_col in column_map.items():
            if actual_col in self.df.columns:
                total_rows = len(self.df)
                non_null_rows = self.df[actual_col].notna().sum()
                completeness[col_name] = {
                    'total_rows': total_rows,
                    'non_null_rows': int(non_null_rows),
                    'completeness_pct': round((non_null_rows / total_rows) * 100, 2)
                }
        
        # Check for expression data completeness
        if 'expression' in column_map:
            expr_col = column_map['expression']
            numeric_values = pd.to_numeric(self.df[expr_col], errors='coerce').notna().sum()
            completeness['expression_numeric'] = {
                'total_rows': len(self.df),
                'numeric_values': int(numeric_values),
                'numeric_pct': round((numeric_values / len(self.df)) * 100, 2)
            }
        
        return completeness
    
    def print_summary(self):
        """Print a comprehensive summary of the analysis."""
        if not self.stats:
            print("❌ No analysis performed yet. Run analyze() first.")
            return
        
        print(f"\n📋 ANALYSIS SUMMARY: {self.stats['file_info']['file_name']}")
        print("="*80)
        
        # File info
        print(f"📁 File: {self.stats['file_info']['file_path']}")
        print(f"📊 Total rows: {self.stats['file_info']['total_rows']:,}")
        print(f"📋 Columns: {len(self.stats['file_info']['columns'])}")
        
        # Study statistics
        study_stats = self.stats['study_stats']
        if 'error' not in study_stats:
            print(f"\n🔬 STUDY STATISTICS:")
            print(f"   Total studies: {study_stats['total_studies']}")
            print(f"   Average rows per study: {study_stats['total_rows_all_studies'] / study_stats['total_studies']:.1f}")
            
            print(f"\n📊 STUDIES BY ROW COUNT:")
            for study in study_stats['studies'][:10]:  # Top 10
                name = study['name'][:40] + '...' if len(study['name']) > 40 else study['name']
                print(f"   {name:<45} {study['total_rows']:>6} rows")
            
            if len(study_stats['studies']) > 10:
                print(f"   ... and {len(study_stats['studies']) - 10} more studies")
        
        # Target statistics
        target_stats = self.stats['target_stats']
        if 'error' not in target_stats:
            print(f"\n🎯 TARGET STATISTICS:")
            print(f"   Unique targets: {target_stats['total_unique_targets']}")
            print(f"   Total measurements: {target_stats['total_target_measurements']:,}")
            
            print(f"\n🔝 TOP TARGETS:")
            for target, count in list(target_stats['top_10_targets'].items())[:5]:
                print(f"   {target:<30} {count:>6} measurements")
        
        # Trigger statistics
        trigger_stats = self.stats['trigger_stats']
        if 'error' not in trigger_stats:
            print(f"\n💉 TRIGGER STATISTICS:")
            print(f"   Unique triggers: {trigger_stats['total_unique_triggers']}")
            print(f"   Total measurements: {trigger_stats['total_trigger_measurements']:,}")
            
            print(f"\n🔝 TOP TRIGGERS:")
            for trigger, count in list(trigger_stats['top_10_triggers'].items())[:5]:
                print(f"   {trigger:<30} {count:>6} measurements")
        
        # Tissue statistics
        tissue_stats = self.stats['tissue_stats']
        if 'error' not in tissue_stats:
            print(f"\n🧬 TISSUE STATISTICS:")
            print(f"   Unique tissues: {tissue_stats['total_unique_tissues']}")
            print(f"   Total measurements: {tissue_stats['total_tissue_measurements']:,}")
        
        # Completeness statistics
        completeness = self.stats['completeness_stats']
        print(f"\n✅ DATA COMPLETENESS:")
        for field, stats in completeness.items():
            if isinstance(stats, dict) and 'completeness_pct' in stats:
                print(f"   {field:<20} {stats['completeness_pct']:>6.1f}% complete")
        
        print("="*80)
    
    def export_detailed_report(self, output_file: str):
        """Export detailed analysis to a text file."""
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(f"DETAILED STUDY DATA ANALYSIS REPORT\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"File: {self.stats['file_info']['file_path']}\n")
            f.write("="*80 + "\n\n")
            
            # Study details
            study_stats = self.stats['study_stats']
            if 'error' not in study_stats:
                f.write("DETAILED STUDY BREAKDOWN:\n")
                f.write("-"*40 + "\n")
                
                for study in study_stats['studies']:
                    f.write(f"\nStudy: {study['name']}\n")
                    f.write(f"  Total rows: {study['total_rows']}\n")
                    f.write(f"  Unique targets: {study['unique_targets']}\n")
                    f.write(f"  Unique triggers: {study['unique_triggers']}\n")
                    f.write(f"  Unique tissues: {study['unique_tissues']}\n")
                    
                    if 'targets' in study:
                        f.write(f"  Targets: {', '.join(study['targets'])}\n")
                    if 'triggers' in study:
                        f.write(f"  Triggers: {', '.join(study['triggers'])}\n")
                    if 'tissues' in study:
                        f.write(f"  Tissues: {', '.join(study['tissues'])}\n")
                    if 'item_types' in study:
                        f.write(f"  Item types: {study['item_types']}\n")
            
            # Export full statistics as JSON for machine reading
            f.write(f"\n\nFULL STATISTICS (JSON):\n")
            f.write("-"*40 + "\n")
            json.dump(self.stats, f, indent=2, default=str)
        
        print(f"✓ Detailed report exported to: {output_file}")

def compare_datasets(analyzer1: StudyDataAnalyzer, analyzer2: StudyDataAnalyzer):
    """Compare two datasets and highlight differences."""
    print(f"\n🔄 DATASET COMPARISON")
    print("="*80)
    
    stats1 = analyzer1.stats
    stats2 = analyzer2.stats
    
    print(f"Dataset 1: {stats1['file_info']['file_name']}")
    print(f"Dataset 2: {stats2['file_info']['file_name']}")
    
    # Compare basic metrics
    print(f"\n📊 BASIC COMPARISON:")
    print(f"{'Metric':<25} {'Dataset 1':<15} {'Dataset 2':<15} {'Difference'}")
    print("-"*70)
    
    metrics = [
        ('Total rows', 'file_info', 'total_rows'),
        ('Total studies', 'study_stats', 'total_studies'),
        ('Unique targets', 'target_stats', 'total_unique_targets'),
        ('Unique triggers', 'trigger_stats', 'total_unique_triggers'),
        ('Unique tissues', 'tissue_stats', 'total_unique_tissues')
    ]
    
    for metric_name, section, key in metrics:
        try:
            val1 = stats1[section][key]
            val2 = stats2[section][key]
            diff = val1 - val2
            diff_str = f"{diff:+d}" if diff != 0 else "0"
            print(f"{metric_name:<25} {val1:<15} {val2:<15} {diff_str}")
        except KeyError:
            print(f"{metric_name:<25} {'N/A':<15} {'N/A':<15} {'N/A'}")
    
    # Compare studies
    if ('study_stats' in stats1 and 'studies' in stats1['study_stats'] and
        'study_stats' in stats2 and 'studies' in stats2['study_stats']):
        
        studies1 = {s['name']: s['total_rows'] for s in stats1['study_stats']['studies']}
        studies2 = {s['name']: s['total_rows'] for s in stats2['study_stats']['studies']}
        
        all_studies = set(studies1.keys()) | set(studies2.keys())
        
        print(f"\n📋 STUDY-BY-STUDY COMPARISON:")
        print(f"{'Study Name':<40} {'Dataset 1':<12} {'Dataset 2':<12} {'Difference'}")
        print("-"*80)
        
        for study in sorted(all_studies):
            rows1 = studies1.get(study, 0)
            rows2 = studies2.get(study, 0)
            diff = rows1 - rows2
            
            status = ""
            if study not in studies1:
                status = "(only in dataset 2)"
            elif study not in studies2:
                status = "(only in dataset 1)"
            elif diff != 0:
                status = f"({diff:+d})"
            
            print(f"{study[:39]:<40} {rows1:<12} {rows2:<12} {status}")
    
    print("="*80)

def main():
    """Main function to handle command line arguments and run analysis."""
    parser = argparse.ArgumentParser(
        description='Analyze study data from CSV or Excel files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python analyze_study_data.py data.csv
  python analyze_study_data.py data.xlsx --output report.txt
  python analyze_study_data.py experimental.csv manual.xlsx --compare
        """
    )
    
    parser.add_argument('files', nargs='+', help='Input file(s) to analyze')
    parser.add_argument('--output', '-o', help='Output file for detailed report')
    parser.add_argument('--compare', '-c', action='store_true', 
                       help='Compare two datasets (requires exactly 2 input files)')
    parser.add_argument('--json', '-j', help='Export statistics as JSON file')
    
    args = parser.parse_args()
    
    if args.compare and len(args.files) != 2:
        print("❌ Comparison mode requires exactly 2 input files")
        sys.exit(1)
    
    # Analyze files
    analyzers = []
    for file_path in args.files:
        print(f"\n🔍 Processing: {file_path}")
        analyzer = StudyDataAnalyzer(file_path)
        analyzer.analyze()
        analyzer.print_summary()
        analyzers.append(analyzer)
        
        # Export detailed report if requested
        if args.output and len(args.files) == 1:
            analyzer.export_detailed_report(args.output)
        
        # Export JSON if requested
        if args.json and len(args.files) == 1:
            with open(args.json, 'w') as f:
                json.dump(analyzer.stats, f, indent=2, default=str)
            print(f"✓ Statistics exported to JSON: {args.json}")
    
    # Compare datasets if requested
    if args.compare and len(analyzers) == 2:
        compare_datasets(analyzers[0], analyzers[1])
    
    print(f"\n✅ Analysis complete!")

if __name__ == "__main__":
    main() 