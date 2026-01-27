#!/usr/bin/env python3
"""
Regenerate semester 7 timetables using the corrected Excel writing logic
"""
import os
import sys
import glob

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(__file__))

from app import export_consolidated_semester_timetable, load_all_data

# Load the input data
print("Loading input files...")
dfs = load_all_data(force_reload=True)

# Generate timetables for semester 7
print("Generating semester 7 timetables...")
for branch in ['CSE', 'DSAI', 'ECE']:
    try:
        print(f"  - Generating {branch} semester 7...")
        export_consolidated_semester_timetable(dfs, 7, branch)
        print(f"    [OK] {branch} semester 7 completed")
    except Exception as e:
        print(f"    [ERROR] {branch}: {str(e)}")

print("\nRegeneration complete!")
print("\nVerifying generated files...")
output_dir = os.path.join(os.path.dirname(__file__), 'output_timetables')
sem7_files = glob.glob(os.path.join(output_dir, 'sem7_*.xlsx'))
print(f"Found {len(sem7_files)} semester 7 files:")
for f in sorted(sem7_files):
    print(f"  - {os.path.basename(f)}")

# Verify column structure
print("\nVerifying column structure...")
import pandas as pd
for filepath in sorted(sem7_files):
    try:
        df = pd.read_excel(filepath, sheet_name=0)
        print(f"\n{os.path.basename(filepath)}:")
        print(f"  Columns: {list(df.columns)}")
        print(f"  Shape: {df.shape}")
    except Exception as e:
        print(f"  ERROR: {e}")
