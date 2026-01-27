#!/usr/bin/env python3
"""Final verification focused on the core issue: CS304 and CS307 in CSE sem3"""
import pandas as pd
import os

file_path = 'output_timetables/sem3_CSE_timetable.xlsx'

print("\n" + "=" * 80)
print("FINAL VERIFICATION: CS304 and CS307 in sem3_CSE_timetable.xlsx")
print("=" * 80)

if not os.path.exists(file_path):
    print(f"[ERROR] File not found: {file_path}")
    exit(1)

try:
    excel_file = pd.ExcelFile(file_path)
    
    # Find course information sheet
    course_info_sheet = None
    for sheet in excel_file.sheet_names:
        if 'course' in sheet.lower() and 'information' in sheet.lower():
            course_info_sheet = sheet
            break
    
    if not course_info_sheet:
        print(f"[ERROR] No Course Information sheet found")
        exit(1)
    
    df = pd.read_excel(file_path, sheet_name=course_info_sheet)
    
    print(f"\nOpened: {file_path}")
    print(f"Sheet: {course_info_sheet}")
    print(f"\nCourse Information content (first 15 rows):")
    print("=" * 80)
    print(df.head(15).to_string())
    
    print("\n" + "=" * 80)
    print("SEARCHING FOR PROBLEMATIC COURSES:")
    print("=" * 80)
    
    # Search for CS304 and CS307
    for course_code in ['CS304', 'CS307']:
        # Convert all cells to string and search
        found = False
        for col in df.columns:
            for idx, val in enumerate(df[col]):
                if isinstance(val, str) and course_code in val:
                    print(f"\n[FOUND] {course_code} in row {idx}, column {col}:")
                    print(f"  Value: {val}")
                    found = True
                elif not pd.isna(val):
                    val_str = str(val)
                    if course_code in val_str:
                        print(f"\n[FOUND] {course_code} in row {idx}, column {col}:")
                        print(f"  Value: {val_str}")
                        found = True
        
        if not found:
            print(f"\n[✓ PASS] {course_code} is NOT present in the timetable (CORRECT!)")
    
    # Show what core courses ARE in the sheet
    print("\n" + "=" * 80)
    print("CORE COURSES THAT ARE IN THE SHEET:")
    print("=" * 80)
    
    # Find the core courses section
    core_found = False
    core_courses = []
    for idx, row in df.iterrows():
        row_str = str(row.values).upper()
        if 'CORE COURSES' in row_str:
            core_found = True
            continue
        if core_found and 'ELECTIVE' in row_str:
            break
        if core_found and 'COURSE CODE' not in row_str:
            # Try to extract course codes
            for val in row.values:
                if pd.notna(val):
                    val_str = str(val).strip()
                    if len(val_str) >= 4 and val_str[0].isalpha() and val_str[1:4].isalpha():
                        if val_str[:2] in ['CS', 'MA', 'DS', 'EC', 'HS', 'DA']:
                            core_courses.append(val_str)
    
    unique_core = list(set(core_courses))
    unique_core.sort()
    print(f"Core courses found: {unique_core}")
    
    if 'CS304' not in unique_core and 'CS307' not in unique_core:
        print("\n[✓ PASS] CS304 and CS307 are NOT in the core courses!")
    else:
        print("\n[✗ FAIL] CS304 or CS307 found in core courses")
    
except Exception as e:
    print(f"[ERROR] {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("VERIFICATION COMPLETE")
print("=" * 80 + "\n")
