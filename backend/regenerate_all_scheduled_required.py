#!/usr/bin/env python3
"""
Regenerate ALL semester timetables with updated L/T/P columns showing scheduled/required format.
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

# Define all semesters and branches
SEMESTERS = [1, 3, 5, 7]
BRANCHES = ['CSE', 'DSAI', 'ECE']

output_dir = os.path.join(os.path.dirname(__file__), 'output_timetables')

print("\n" + "="*80)
print("REGENERATING ALL TIMETABLES WITH SCHEDULED/REQUIRED FORMAT")
print("="*80)

# Delete old files first
print("\nCleaning old timetable files...")
for sem in SEMESTERS:
    for branch in BRANCHES:
        old_file = os.path.join(output_dir, f'sem{sem}_{branch}_timetable.xlsx')
        if os.path.exists(old_file):
            try:
                os.remove(old_file)
                print(f"  Deleted: sem{sem}_{branch}_timetable.xlsx")
            except Exception as e:
                print(f"  [WARN] Could not delete sem{sem}_{branch}_timetable.xlsx: {e}")

print("\nRegenerating all timetables...")
total = len(SEMESTERS) * len(BRANCHES)
count = 0

for sem in SEMESTERS:
    for branch in BRANCHES:
        count += 1
        print(f"\n[{count}/{total}] Generating Semester {sem} - {branch}...")
        
        try:
            # Call the main export function
            result = export_consolidated_semester_timetable(
                dfs=dfs,
                semester=sem,
                branch=branch
            )
            
            if result:
                print(f"  ✓ Successfully generated sem{sem}_{branch}_timetable.xlsx")
            else:
                print(f"  ✗ Failed to generate sem{sem}_{branch}_timetable.xlsx")
                
        except Exception as e:
            print(f"  ✗ Error: {str(e)}")
            import traceback
            traceback.print_exc()

print("\n" + "="*80)
print("REGENERATION COMPLETE")
print("="*80)
print("\nAll timetables now show L/T/P columns in 'Scheduled/Required' format")
print(f"Output directory: {output_dir}")
