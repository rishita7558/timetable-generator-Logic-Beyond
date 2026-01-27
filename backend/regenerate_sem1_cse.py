#!/usr/bin/env python3
"""
Regenerate just sem1_CSE_timetable.xlsx with the new L/T/P columns.
"""
import os
import sys

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(__file__))

from app import export_consolidated_semester_timetable, load_all_data

# Load the input data
print("Loading input files...")
dfs = load_all_data(force_reload=True)

output_dir = os.path.join(os.path.dirname(__file__), 'output_timetables')

print("\n" + "="*80)
print("REGENERATING sem1_CSE_timetable.xlsx")
print("="*80)

try:
    result = export_consolidated_semester_timetable(
        dfs=dfs,
        semester=1,
        branch='CSE'
    )
    
    if result:
        print("\n✓ Successfully generated sem1_CSE_timetable.xlsx")
        print("\nThe file now includes Lectures/Week, Tutorials/Week, Labs/Week columns")
        print("in the CORE COURSES section.")
    else:
        print("\n✗ Failed to generate sem1_CSE_timetable.xlsx")
        
except Exception as e:
    print(f"\n✗ Error: {str(e)}")
    import traceback
    traceback.print_exc()

print("=" * 80)
