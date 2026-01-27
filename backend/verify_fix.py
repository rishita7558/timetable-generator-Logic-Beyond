#!/usr/bin/env python3
"""Verify that CS304 and CS307 are NOT appearing in CSE courses"""
import pandas as pd
import openpyxl
import os

# Check the timetable files
timetables_dir = 'output_timetables'

if not os.path.exists(timetables_dir):
    print(f"[ERROR] Directory {timetables_dir} not found")
    exit(1)

print("=" * 80)
print("VERIFICATION: Checking if CS304 and CS307 are in CSE sem3 timetable")
print("=" * 80)

file_to_check = os.path.join(timetables_dir, 'sem3_CSE_timetable.xlsx')

if not os.path.exists(file_to_check):
    print(f"[ERROR] File {file_to_check} not found")
    exit(1)

# Read the COURSE INFORMATION sheet
try:
    excel_file = pd.ExcelFile(file_to_check)
    print(f"\n[OK] Opened {file_to_check}")
    print(f"Sheets: {excel_file.sheet_names}")
    
    # Try to read COURSE INFORMATION (check for different naming)
    course_info_sheet = None
    for sheet in excel_file.sheet_names:
        if 'course' in sheet.lower() and 'information' in sheet.lower():
            course_info_sheet = sheet
            break
    
    if course_info_sheet:
        df = pd.read_excel(file_to_check, sheet_name=course_info_sheet)
        print(f"\n[INFO] COURSE INFORMATION sheet has {len(df)} rows")
        print("\nFirst 20 rows:")
        print(df.head(20).to_string())
        
        # Check if CS304 or CS307 are in the data
        print("\n" + "=" * 80)
        print("CHECKING FOR CS304 AND CS307:")
        print("=" * 80)
        
        for course in ['CS304', 'CS307']:
            matches = df[df.isin([course]).any(axis=1)]
            if len(matches) > 0:
                print(f"\n[ERROR] Found {course} in COURSE INFORMATION:")
                print(matches.to_string())
            else:
                print(f"[OK] {course} NOT found in COURSE INFORMATION")
    else:
        print(f"[ERROR] Sheet 'COURSE INFORMATION' not found")
        print(f"Available sheets: {excel_file.sheet_names}")
        
except Exception as e:
    print(f"[ERROR] Failed to read file: {e}")
    import traceback
    traceback.print_exc()

print("\n" + "=" * 80)
print("VERIFICATION COMPLETE")
print("=" * 80)
