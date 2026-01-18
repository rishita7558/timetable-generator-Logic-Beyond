#!/usr/bin/env python3
"""
Regenerate all timetables with the fixed column structure
"""
import os
import sys
import glob
import shutil
from pathlib import Path

# Add backend to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import (
    OUTPUT_DIR, load_all_data, export_semester_timetable_with_baskets
)
import pandas as pd

def regenerate_all():
    """Regenerate all timetables"""
    print("=" * 80)
    print("REGENERATING ALL TIMETABLES WITH FIXED COLUMN STRUCTURE")
    print("=" * 80)
    
    # Delete old timetables
    print("\n[CLEANUP] Removing old timetable files...")
    old_files = glob.glob(os.path.join(OUTPUT_DIR, "sem*_*_timetable_baskets.xlsx"))
    for f in old_files:
        try:
            os.remove(f)
            print(f"  [OK] Deleted {os.path.basename(f)}")
        except Exception as e:
            print(f"  [FAIL] Failed to delete {os.path.basename(f)}: {e}")
    
    # Load data
    print("\n[DATA] Loading course data...")
    dfs = load_all_data(force_reload=True)
    if not dfs:
        print("[FAIL] Could not load data frames")
        return False
    
    print("[OK] Data loaded successfully")
    
    # Generate for all branches and semesters
    semesters = [1, 3, 5, 7]
    branches = ['CSE', 'DSAI', 'ECE']
    
    generated_count = 0
    failed_count = 0
    
    for sem in semesters:
        print(f"\n{'='*60}")
        print(f"SEMESTER {sem}")
        print(f"{'='*60}")
        
        for branch in branches:
            try:
                print(f"\n[GEN] Generating {branch} semester {sem} timetable...")
                
                result = export_semester_timetable_with_baskets(dfs, sem, branch)
                
                if result:
                    filename = f"sem{sem}_{branch}_timetable_baskets.xlsx"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    
                    # Verify the file exists and check column structure
                    if os.path.exists(filepath):
                        try:
                            sheet_name = 'Section_A' if branch == 'CSE' else 'Timetable'
                            df = pd.read_excel(filepath, sheet_name=sheet_name)
                            columns = list(df.columns)
                            print(f"[OK] Generated {filename}")
                            print(f"     Columns: {columns}")
                            
                            # Check for unwanted columns
                            unwanted = [col for col in columns if col in ['level_0', 'Unnamed: 0', 'Unnamed: 1']]
                            if unwanted:
                                print(f"     [WARN] Found unwanted columns: {unwanted}")
                                failed_count += 1
                            else:
                                print(f"     [OK] Column structure clean")
                                generated_count += 1
                        except Exception as e:
                            print(f"[FAIL] Could not read {filename}: {e}")
                            failed_count += 1
                    else:
                        print(f"[FAIL] File not created: {filename}")
                        failed_count += 1
                else:
                    print(f"[FAIL] Generation returned None for {branch}")
                    failed_count += 1
                    
            except Exception as e:
                print(f"[FAIL] Error generating {branch} semester {sem}: {e}")
                import traceback
                traceback.print_exc()
                failed_count += 1
    
    print(f"\n{'='*80}")
    print("REGENERATION COMPLETE")
    print(f"{'='*80}")
    print(f"Generated: {generated_count}")
    print(f"Failed: {failed_count}")
    print(f"Total: {generated_count + failed_count}")
    
    return failed_count == 0

if __name__ == '__main__':
    success = regenerate_all()
    sys.exit(0 if success else 1)
