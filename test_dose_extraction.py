import re

def extract_dose_from_trigger_name(trigger_name):
    if not trigger_name:
        return None
    
    text = str(trigger_name).lower().strip()
    
    # Look for mpk patterns
    mpk_patterns = [
        r'(\d+(?:\.\d+)?)\s*mpk',
        r'(\d+(?:\.\d+)?)\s*mg/kg',
    ]
    
    for pattern in mpk_patterns:
        match = re.search(pattern, text)
        if match:
            dose_value = match.group(1)
            return f'{dose_value} mpk'
    
    # Look for other dose patterns
    dose_patterns = [
        r'(\d+(?:\.\d+)?)\s*(ug|μg|mg|g)\b',
        r'(\d+(?:\.\d+)?)\s*(ul|μl|ml|l)\b',
    ]
    
    for pattern in dose_patterns:
        match = re.search(pattern, text)
        if match:
            dose_value = match.group(1)
            dose_unit = match.group(2)
            return f'{dose_value} {dose_unit}'
    
    return None

# Test cases
test_triggers = [
    'siRNA_5mpk_treatment',
    'control_10 mpk dose',
    'test_2.5mpk_sample',
    'compound_15mg/kg_dose',
    'sample_250ug_treatment',
    'control_5ml_volume',
    'no_dose_info',
    'ASO_3mpk_D14',
    'Vehicle_Control',
    'Treatment_1.5mpk',
    'Dose_20mg/kg_group'
]

print('Testing dose extraction:')
for trigger in test_triggers:
    result = extract_dose_from_trigger_name(trigger)
    print(f'  {trigger:<25} -> {result}') 