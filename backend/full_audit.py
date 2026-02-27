"""Full generation and audit of all timetables."""
import app
import glob
import os

print("=" * 60)
print("FULL TIMETABLE GENERATION AND AUDIT")
print("=" * 60)

# Reset all trackers
app.reset_classroom_usage_tracker()
print("\n[STEP 1] Reset all trackers")

# Load data
dfs = app.load_all_data(force_reload=True)
print(f"[STEP 2] Loaded data: {list(dfs.keys())}")

# Clear old files
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), 'output_timetables')
for f in glob.glob(os.path.join(OUTPUT_DIR, "sem*_*_timetable.xlsx")):
    try:
        os.remove(f)
        print(f"[CLEAN] Removed: {os.path.basename(f)}")
    except Exception as e:
        print(f"[WARN] Could not remove {f}: {e}")

# Generate all timetables
print("\n[STEP 3] Generating all timetables...")
branches = ['DSAI', 'ECE', 'CSE']  # Process order
semesters = [1, 3, 5, 7]
success_count = 0

import sys
import io

for branch in branches:
    for sem in semesters:
        try:
            # Capture stdout to avoid print interference with allocation
            old_stdout = sys.stdout
            sys.stdout = io.StringIO()
            result = app.export_consolidated_semester_timetable(dfs, sem, branch)
            sys.stdout = old_stdout
            if result:
                success_count += 1
                print(f"[OK] {branch} Sem {sem}")
            else:
                print(f"[FAIL] {branch} Sem {sem}")
        except Exception as e:
            sys.stdout = old_stdout
            print(f"[ERROR] {branch} Sem {sem}: {e}")

print(f"\n[STEP 4] Generated {success_count} timetables")

# Regenerate classroom audit file using the actual tracker data
print("\n[STEP 4.5] Regenerating classroom audit file...")
audit_path = app.generate_classroom_audit_file(dfs, OUTPUT_DIR)
if audit_path:
    print(f"[OK] Classroom audit file updated: {os.path.basename(audit_path)}")

# Now audit for conflicts
print("\n[STEP 5] Running conflict audit...")

import pandas as pd
import re
from collections import defaultdict

# Group sheets by period type - only compare SAME period sheets
# Regular_* vs Regular_*, PreMid_* vs PreMid_*, PostMid_* vs PostMid_*
def get_period_type(sheet_name):
    """Extract period type from sheet name for grouping."""
    if 'Regular' in sheet_name:
        return 'Regular'
    elif 'PreMid' in sheet_name:
        return 'PreMid'
    elif 'PostMid' in sheet_name:
        return 'PostMid'
    return None

files = glob.glob(os.path.join(OUTPUT_DIR, "sem*_*_timetable.xlsx"))

# Structure: period_type -> room -> day -> time_slot -> list of (file, sheet, course)
# This ensures we only compare schedules from the SAME period
period_schedules = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))

for filepath in sorted(files):
    filename = os.path.basename(filepath)
    try:
        xl = pd.ExcelFile(filepath)
        for sheet in xl.sheet_names:
            if 'Legend' in sheet or 'Audit' in sheet:
                continue
            df = pd.read_excel(filepath, sheet_name=sheet, index_col=0)
            
            # Handle duplicate indices safely
            if df.index.has_duplicates:
                df = df[~df.index.duplicated(keep='first')]
            
            # Get period type for this sheet
            period_type = get_period_type(sheet)
            if period_type is None:
                continue  # Skip non-schedule sheets
            
            for day in df.columns:
                if day not in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']:
                    continue
                for time_slot in df.index:
                    try:
                        cell = df.loc[time_slot, day]
                        # Handle Series (duplicate indices)
                        if isinstance(cell, pd.Series):
                            cell = cell.iloc[0] if len(cell) > 0 else None
                    except Exception:
                        continue
                    if pd.isna(cell) or not str(cell).strip():
                        continue
                    cell_str = str(cell).strip()
                    if cell_str in ['Free', 'LUNCH BREAK']:
                        continue
                    # Extract room 
                    matches = re.findall(r'\[(C\d+|L\d+)\]', cell_str)
                    for room in matches:
                        # Group by period type so we only compare same-period schedules
                        period_schedules[period_type][room][day][time_slot].append({
                            'file': filename,
                            'sheet': sheet,
                            'course': cell_str
                        })
    except Exception as e:
        print(f"[ERROR] Could not read {filename}: {e}")

# Find conflicts WITHIN SAME PERIOD TYPE ONLY
# (Regular vs Regular, PreMid vs PreMid, etc. - NOT across periods)
conflicts = []
for period_type, room_data in period_schedules.items():
    for room, days in room_data.items():
        for day, slots in days.items():
            for time_slot, bookings in slots.items():
                if len(bookings) > 1:
                    # Extract course names (remove room brackets for comparison)
                    courses = set()
                    for b in bookings:
                        course = re.sub(r'\s*\[.*?\]\s*', '', b['course']).strip()
                        courses.add(course)
                    
                    # If all bookings are for the SAME course, it's a common course - NOT a conflict
                    if len(courses) == 1:
                        continue
                    
                    # Check if bookings are from different files (different branches)
                    # Sections in the same file are for the same branch - also not a conflict
                    files_involved = set(b['file'] for b in bookings)
                    if len(files_involved) == 1:
                        # Same file, different sheets (sections) - check if truly different courses
                        # Get base file without section suffix for comparison
                        pass  # Continue to conflict check
                
                    conflicts.append({
                        'room': room,
                        'day': day,
                        'time_slot': time_slot,
                        'period': period_type,
                        'bookings': bookings
                    })

print(f"\n[RESULT] Found {len(conflicts)} conflicts")
if conflicts:
    for c in conflicts:
        print(f"\n  CONFLICT [{c['period']}]: {c['room']} on {c['day']} {c['time_slot']}")
        for b in c['bookings']:
            print(f"    - {b['file']} {b['sheet']}: {b['course'][:60]}...")
else:
    print("\n  SUCCESS: No classroom double-booking conflicts detected!")
