#!/usr/bin/env python3
"""
Test script for the new tissue extraction logic
"""

# Import the necessary functions from process_study.py
import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from process_study import classify_target_or_tissue, Config

def test_tissue_extraction():
    """Test the tissue extraction logic with various cases"""
    
    test_cases = [
        # Format: (input_text, expected_classification, expected_tissue, expected_target)
        ("adrenal gland", "tissue", "adrenal gland", None),
        ("APOE4 adrenal gland", "both", "adrenal gland", "APOE4"),
        ("liver APOE4", "both", "liver", "APOE4"),
        ("APOE4 liver BACE1", "both", "liver", "APOE4 BACE1"),
        ("adrenal gland liver", "tissue", "adrenal gland liver", None),
        ("APOE4", "target", None, "APOE4"),
        ("liver", "tissue", "liver", None),
        ("sciatic nerve", "tissue", "sciatic nerve", None),
        ("APOE4 sciatic nerve BACE1", "both", "sciatic nerve", "APOE4 BACE1"),
        ("left ventricle", "tissue", "left ventricle", None),
        ("APOE4 left ventricle", "both", "left ventricle", "APOE4"),
        ("kidney cortex APOE4", "both", "kidney cortex", "APOE4"),
        ("lymph node", "tissue", "lymph node", None),
        ("APOE4 lymph node BACE1 liver", "both", "lymph node liver", "APOE4 BACE1"),
    ]
    
    print("Testing tissue extraction logic:")
    print("="*60)
    
    for i, (input_text, expected_class, expected_tissue, expected_target) in enumerate(test_cases, 1):
        print(f"\nTest {i}: '{input_text}'")
        
        # Test with empty procedure_tissues list
        result_class, result_tissue, result_target = classify_target_or_tissue(input_text, [])
        
        print(f"  Expected: {expected_class} | tissue: '{expected_tissue}' | target: '{expected_target}'")
        print(f"  Actual:   {result_class} | tissue: '{result_tissue}' | target: '{result_target}'")
        
        # Check if results match expectations
        success = (result_class == expected_class and 
                  result_tissue == expected_tissue and 
                  result_target == expected_target)
        
        print(f"  Result: {'✓ PASS' if success else '✗ FAIL'}")
        
        if not success:
            print(f"  *** MISMATCH ***")

if __name__ == "__main__":
    test_tissue_extraction() 