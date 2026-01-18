#!/usr/bin/env python3
"""
Regenerate ALL semester timetables to ensure consistent Excel column structure
"""
import os
import sys
import glob
import pandas as pd

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(__file__))

from app import export_semester_timetable_with_baskets, load_all_data

# Load the input data
print("Loading input files...")
dfs = load_all_data(force_reload=True)

# Define all semesters and branches
SEMESTERS = [1, 3, 5, 7]
BRANCHES = ['CSE', 'DSAI', 'ECE']

output_dir = os.path.join(os.path.dirname(__file__), 'output_timetables')

print("\n" + "="*60)
print("REGENERATING ALL SEMESTER TIMETABLES")
print("="*60)

# Delete old files first
print("\nCleaning old timetable files...")
for sem in SEMESTERS:
    for branch in BRANCHES:
        pattern = os.path.join(output_dir, f'sem{sem}_{branch}_timetable_baskets.xlsx')
        for f in glob.glob(pattern):
            try:
                os.remove(f)
                print(f"  Deleted: {os.path.basename(f)}")
            except Exception as e:
                print(f"  [WARN] Could not delete {os.path.basename(f)}: {e}")

# Also remove lock files
for lockfile in glob.glob(os.path.join(output_dir, '~$*.xlsx')) + glob.glob(os.path.join(output_dir, '.~*.xlsx')):
    try:
        os.remove(lockfile)
    except:
        pass

print("\n" + "="*60)
print("GENERATION IN PROGRESS")
print("="*60)

results = {}

# Generate timetables for all semesters and branches
for semester in SEMESTERS:
    print(f"\n[SEMESTER {semester}]")
    results[semester] = {}
    
    for branch in BRANCHES:
        try:
            print(f"  Generating {branch}...", end=" ", flush=True)
            export_semester_timetable_with_baskets(dfs, semester, branch)
            print("[OK]")
            results[semester][branch] = 'SUCCESS'
        except Exception as e:
            print(f"[ERROR]")
            print(f"    {str(e)}")
            results[semester][branch] = f'ERROR: {str(e)}'

print("\n" + "="*60)
print("VERIFICATION")
print("="*60)

# Verify generated files and column structure
all_verified = True

for semester in SEMESTERS:
    print(f"\n[SEMESTER {semester}]")
    
    for branch in BRANCHES:
        filepath = os.path.join(output_dir, f'sem{semester}_{branch}_timetable_baskets.xlsx')
        
        if not os.path.exists(filepath):
            print(f"  {branch}: [MISSING FILE]")
            all_verified = False
            continue
        
        try:
            # For CSE, check both sections
            if branch == 'CSE':
                df_a = pd.read_excel(filepath, sheet_name='Section_A')
                df_b = pd.read_excel(filepath, sheet_name='Section_B')
                cols_a = list(df_a.columns)
                cols_b = list(df_b.columns)
                
                expected_cols = ['Time Slot', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']
                
                if cols_a == expected_cols and cols_b == expected_cols:
                    print(f"  {branch}: [OK] Columns match (Section A & B)")
                else:
                    print(f"  {branch}: [MISMATCH]")
                    if cols_a != expected_cols:
                        print(f"    Section A: {cols_a}")
                    if cols_b != expected_cols:
                        print(f"    Section B: {cols_b}")
                    all_verified = False
            else:
                df = pd.read_excel(filepath, sheet_name='Timetable')
                cols = list(df.columns)
                expected_cols = ['Time Slot', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri']
                
                if cols == expected_cols:
                    print(f"  {branch}: [OK] Columns match")
                else:
                    print(f"  {branch}: [MISMATCH]")
                    print(f"    Got: {cols}")
                    print(f"    Expected: {expected_cols}")
                    all_verified = False
                    
        except Exception as e:
            print(f"  {branch}: [ERROR] {str(e)}")
            all_verified = False

print("\n" + "="*60)
if all_verified:
    print("RESULT: ALL TIMETABLES REGENERATED SUCCESSFULLY!")
else:
    print("RESULT: SOME ISSUES DETECTED - SEE ABOVE")
print("="*60)
