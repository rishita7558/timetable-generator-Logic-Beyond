#!/usr/bin/env python3
import openpyxl

wb = openpyxl.load_workbook('output_timetables/sem1_CSE_timetable.xlsx')
ws = wb['Course_Information']

print("=" * 80)
print("COURSE INFORMATION SHEET - sem1_CSE_timetable.xlsx")
print("=" * 80)

in_core = False
in_elective = False
core_courses = []
elective_courses = []

for i, row in enumerate(ws.iter_rows(min_row=1, max_row=100, values_only=True), 1):
    if row[0] is None:
        continue
    
    row_str = str(row[0]).upper()
    
    if 'CORE COURSES' in row_str:
        in_core = True
        in_elective = False
        print(f"\n✓ CORE COURSES section found at row {i}")
        continue
    
    if 'ELECTIVE COURSES' in row_str:
        in_core = False
        in_elective = True
        print(f"✓ ELECTIVE COURSES section found at row {i}")
        print(f"  Total core courses found: {len(core_courses)}")
        print(f"  Core courses: {core_courses}")
        continue
    
    if 'MINOR COURSES' in row_str:
        in_elective = False
        print(f"✓ MINOR COURSES section found at row {i}")
        print(f"  Total elective courses found: {len(elective_courses)}")
        print(f"  Elective courses: {elective_courses}")
        break
    
    if in_core and row[0] and row[0] not in ['Course Code', 'Course Name']:
        if len(str(row[0])) > 0 and not str(row[0]).startswith('L-T-P'):
            core_courses.append(str(row[0]))
    
    if in_elective and row[0] and row[0] not in ['Course Code', 'Course Name']:
        if len(str(row[0])) > 0 and not str(row[0]).startswith('L-T-P'):
            elective_courses.append(str(row[0]))

print("\n" + "=" * 80)
print("SUMMARY:")
print(f"Core courses in sem1_CSE: {core_courses}")
print(f"Elective courses in sem1_CSE: {elective_courses}")
print("=" * 80)
