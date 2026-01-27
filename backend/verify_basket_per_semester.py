#!/usr/bin/env python3
"""Verify that timetables show only the correct baskets for each semester"""
import pandas as pd
import os

# Expected baskets per semester
EXPECTED_BASKETS = {
    1: ['ELECTIVE_B1'],
    3: ['ELECTIVE_B3'],
    5: ['ELECTIVE_B4', 'ELECTIVE_B5'],
    7: ['ELECTIVE_B6', 'ELECTIVE_B7', 'ELECTIVE_B8', 'ELECTIVE_B9']
}

timetables_dir = 'output_timetables'
all_timetables = [
    ('sem1_CSE_timetable.xlsx', 1),
    ('sem1_DSAI_timetable.xlsx', 1),
    ('sem1_ECE_timetable.xlsx', 1),
    ('sem3_CSE_timetable.xlsx', 3),
    ('sem3_DSAI_timetable.xlsx', 3),
    ('sem3_ECE_timetable.xlsx', 3),
    ('sem5_CSE_timetable.xlsx', 5),
    ('sem5_DSAI_timetable.xlsx', 5),
    ('sem5_ECE_timetable.xlsx', 5),
    ('sem7_CSE_timetable.xlsx', 7),
    ('sem7_DSAI_timetable.xlsx', 7),
    ('sem7_ECE_timetable.xlsx', 7),
]

print("=" * 80)
print("VERIFYING BASKET ALLOCATION PER SEMESTER")
print("=" * 80)

all_passed = True

for filename, semester in all_timetables:
    filepath = os.path.join(timetables_dir, filename)
    if not os.path.exists(filepath):
        print(f"\n[SKIP] {filename} not found")
        continue
    
    expected = set(EXPECTED_BASKETS[semester])
    
    try:
        df = pd.read_excel(filepath, sheet_name='Course_Information')
        
        # Find basket column
        basket_col = None
        for col in df.columns:
            if any('BASKET' in str(val).upper() for val in df[col].values if pd.notna(val)):
                basket_col = col
                break
        
        if basket_col is None:
            print(f"\n[SKIP] {filename}: No basket column found")
            continue
        
        # Get all basket values
        baskets_found = set()
        for val in df[basket_col].values:
            if pd.notna(val):
                val_str = str(val).strip()
                # Check if it's an elective basket (starts with ELECTIVE_B, HSS_B, etc.)
                if 'ELECTIVE_B' in val_str:
                    baskets_found.add(val_str)
        
        print(f"\n{filename} (Semester {semester}):")
        print(f"  Expected: {sorted(expected)}")
        print(f"  Found:    {sorted(baskets_found)}")
        
        # Check for incorrect baskets
        unexpected = baskets_found - expected
        missing = expected - baskets_found
        
        if unexpected:
            print(f"  [✗ FAIL] Unexpected baskets: {sorted(unexpected)}")
            all_passed = False
        if missing:
            print(f"  [⚠ WARNING] Missing expected baskets: {sorted(missing)}")
        if not unexpected and not missing:
            print(f"  [✓ PASS] All baskets correct!")
            
    except Exception as e:
        print(f"\n[ERROR] {filename}: {e}")

print("\n" + "=" * 80)
if all_passed:
    print("✓ ALL TIMETABLES HAVE CORRECT BASKET ALLOCATION")
else:
    print("✗ SOME TIMETABLES HAVE INCORRECT BASKETS")
print("=" * 80)
