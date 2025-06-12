#!/usr/bin/env python3
"""
Test script to verify trigger name cleaning functionality.
"""

import sys
sys.path.append('.')

from process_study import clean_trigger_name, init_logger

def test_trigger_cleaning():
    """Test various trigger name cleaning scenarios."""
    
    init_logger("test_trigger_cleaning.log")
    
    test_cases = [
        # Good cases that should be cleaned
        ("250uL  HDM 5ug (D1, 3)", False, "HDM", "5ug"),
        ("200uL  Saline NA (D1, 3)", False, "Saline", None),
        ("AC005120 2x5mpk (D1, 3)", False, "AC005120", "2x5mpk"),
        ("AC00008 + saline", False, "AC00008 + saline", None),
        ("PBS", False, "PBS", None),
        ("ACSF", False, "ACSF", None),
        ("Vehicle", False, "Vehicle", None),
        
        # Cases with existing dose (should not extract dose)
        ("AC005120 2x5mpk (D1, 3)", True, "AC005120", None),
        ("250uL  HDM 5ug (D1, 3)", True, "HDM", None),
        
        # Corrupted cases that should be filtered out
        ("AC006365 4 x", False, "", None),
        ("AC007163/kg", False, "", None),
        ("B-hIL11/hIL11RA/mL", False, "", None),
        ("4 x", False, "", None),
        ("/kg", False, "", None),
        ("/mL", False, "", None),
        ("123", False, "", None),
        
        # Edge cases
        ("", False, "", None),
        ("   ", False, "", None),
        ("AC123456", False, "AC123456", None),
    ]
    
    print("🧪 TESTING TRIGGER NAME CLEANING")
    print("=" * 80)
    
    passed = 0
    failed = 0
    
    for i, (input_name, has_dose, expected_clean, expected_dose) in enumerate(test_cases, 1):
        try:
            actual_clean, actual_dose = clean_trigger_name(input_name, has_dose)
            
            # Check results
            clean_match = actual_clean == expected_clean
            dose_match = actual_dose == expected_dose
            
            if clean_match and dose_match:
                status = "✅ PASS"
                passed += 1
            else:
                status = "❌ FAIL"
                failed += 1
            
            print(f"{i:2d}. {status}")
            print(f"    Input: '{input_name}' (has_dose={has_dose})")
            print(f"    Expected: clean='{expected_clean}', dose='{expected_dose}'")
            print(f"    Actual:   clean='{actual_clean}', dose='{actual_dose}'")
            
            if not clean_match:
                print(f"    ❌ Clean name mismatch!")
            if not dose_match:
                print(f"    ❌ Dose extraction mismatch!")
            print()
            
        except Exception as e:
            print(f"{i:2d}. ❌ ERROR: {e}")
            print(f"    Input: '{input_name}' (has_dose={has_dose})")
            print()
            failed += 1
    
    print("=" * 80)
    print(f"📊 RESULTS: {passed} passed, {failed} failed")
    
    if failed == 0:
        print("🎉 All tests passed!")
    else:
        print(f"⚠️ {failed} tests failed")

if __name__ == "__main__":
    test_trigger_cleaning() 