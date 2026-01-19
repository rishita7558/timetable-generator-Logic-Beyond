"""
Regenerate timetables for Semesters 1 and 3 (to include tutorials for ELECTIVE_B1 and ELECTIVE_B3)
"""
import sys
import os
from app import export_semester_timetable_with_baskets, load_all_data, export_mid_semester_timetables

def main():
    print("=" * 80)
    print("Regenerating Semester 1 and 3 Timetables")
    print("=" * 80)
    
    # Load data once
    dfs = load_all_data(force_reload=True)
    if not dfs:
        print("[ERROR] Failed to load CSV data")
        return False
    
    # Semester 1 - CSE, DSAI, ECE
    print("\n" + "=" * 80)
    print("SEMESTER 1")
    print("=" * 80)
    
    for branch in ['CSE', 'DSAI', 'ECE']:
        print(f"\n[START] Generating {branch} semester 1...")
        success = export_semester_timetable_with_baskets(dfs, 1, branch)
        if success:
            print(f"[OK] {branch} semester 1 completed")
        else:
            print(f"[FAIL] {branch} semester 1 failed")
        
        # Half semester timetables
        print(f"\n[START] Generating {branch} semester 1 half-semester timetables...")
        export_mid_semester_timetables(dfs, 1, branch)
        print(f"[OK] {branch} semester 1 half-semester completed")
    
    # Semester 3 - CSE, DSAI, ECE
    print("\n" + "=" * 80)
    print("SEMESTER 3")
    print("=" * 80)
    
    for branch in ['CSE', 'DSAI', 'ECE']:
        print(f"\n[START] Generating {branch} semester 3...")
        success = export_semester_timetable_with_baskets(dfs, 3, branch)
        if success:
            print(f"[OK] {branch} semester 3 completed")
        else:
            print(f"[FAIL] {branch} semester 3 failed")
        
        # Half semester timetables
        print(f"\n[START] Generating {branch} semester 3 half-semester timetables...")
        export_mid_semester_timetables(dfs, 3, branch)
        print(f"[OK] {branch} semester 3 half-semester completed")
    
    print("\n" + "=" * 80)
    print("Regeneration complete!")
    print("=" * 80)
    
    # Verify files
    print("\nVerifying generated files...")
    from pathlib import Path
    output_dir = Path(__file__).parent / "output_timetables"
    
    for sem in [1, 3]:
        files = list(output_dir.glob(f"sem{sem}_*.xlsx"))
        print(f"Found {len(files)} semester {sem} files:")
        for f in sorted(files):
            print(f"  - {f.name}")
    
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
