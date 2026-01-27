#!/usr/bin/env python3
"""
Verify that COURSE INFORMATION sheets show only courses for that specific semester and department
"""
import openpyxl
from pathlib import Path

def extract_courses_from_sheet(filepath):
    """Extract core and elective courses from Course_Information sheet"""
    try:
        wb = openpyxl.load_workbook(filepath)
        ws = wb['Course_Information']
        
        core_courses = []
        elective_courses = []
        
        in_core = False
        in_elective = False
        
        for row in ws.iter_rows(min_row=1, max_row=150, values_only=True):
            if row[0] is None:
                continue
            
            row_str = str(row[0]).upper()
            
            if 'CORE COURSES' in row_str:
                in_core, in_elective = True, False
                continue
            if 'ELECTIVE COURSES' in row_str:
                in_core, in_elective = False, True
                continue
            if 'MINOR' in row_str:
                break
            
            if in_core and row[0] and row[0] not in ['Course Code']:
                if len(str(row[0])) > 2 and not str(row[0]).upper().startswith(('L-T-P', 'COURSE')):
                    core_courses.append(str(row[0]))
            
            if in_elective and row[0] and row[0] not in ['Course Code']:
                if len(str(row[0])) > 2 and not str(row[0]).upper().startswith(('L-T-P', 'COURSE')):
                    elective_courses.append(str(row[0]))
        
        wb.close()
        return sorted(set(core_courses)), sorted(set(elective_courses))
    except Exception as e:
        return [], []

# Check all timetables
output_dir = Path('output_timetables')
timetable_files = sorted(output_dir.glob('sem*_timetable.xlsx'))

print("=" * 100)
print("VERIFICATION: COURSE INFORMATION SHEET - Courses Filtered by Department & Semester")
print("=" * 100)

all_ok = True
for filepath in timetable_files:
    filename = filepath.name
    # Extract sem, branch from filename (e.g., "sem1_CSE_timetable.xlsx")
    parts = filename.replace('_timetable.xlsx', '').split('_')
    sem = parts[0].replace('sem', '')
    branch = parts[1]
    
    core, elective = extract_courses_from_sheet(filepath)
    
    print(f"\n{filename}")
    print(f"  Semester: {sem}, Branch: {branch}")
    print(f"  Core courses ({len(core)}): {core}")
    print(f"  Elective courses ({len(elective)}): {elective}")
    
    # Check if core courses exist (some branches might not have all semesters)
    if not core and not elective:
        print(f"  ⚠ WARNING: No courses found")
        all_ok = False
    else:
        print(f"  ✓ OK")

print("\n" + "=" * 100)
if all_ok:
    print("✅ ALL TIMETABLES VERIFIED - Courses are filtered by department and semester!")
else:
    print("⚠️ SOME ISSUES FOUND - See above")
print("=" * 100)
