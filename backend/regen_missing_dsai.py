#!/usr/bin/env python3
"""Regenerate specific missing timetables"""
import sys
sys.path.insert(0, '.')
from app import export_consolidated_semester_timetable, load_all_data
import os

# Load data
print("[LOAD] Loading input data...")
dfs = load_all_data()
if not dfs:
    print("[ERROR] Failed to load data")
    exit(1)

print("[OK] Data loaded successfully")

# Generate specific timetables
missing_timetables = [
    (3, 'DSAI'),
    (5, 'DSAI')
]

output_dir = 'output_timetables'
os.makedirs(output_dir, exist_ok=True)

for semester, branch in missing_timetables:
    print(f"\n[TARGET] Generating Semester {semester}, Branch {branch}...")
    try:
        filename = os.path.join(output_dir, f'sem{semester}_{branch}_timetable.xlsx')
        export_consolidated_semester_timetable(dfs, semester, branch)
        print(f"[OK] Generated {filename}")
    except Exception as e:
        print(f"[ERROR] Failed: {e}")
        import traceback
        traceback.print_exc()
