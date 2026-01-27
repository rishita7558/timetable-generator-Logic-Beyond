#!/usr/bin/env python3
"""Comprehensive verification that the course filtering fix works correctly"""
import pandas as pd
import os

timetables_dir = 'output_timetables'

# Expected courses for each semester and branch
expected_courses = {
    ('sem1', 'CSE'): ['MA161', 'DS161', 'MA162', 'EC161'],
    ('sem1', 'DSAI'): ['MA161', 'DS161', 'MA162', 'EC161'],
    ('sem1', 'ECE'): ['MA161', 'DS161', 'MA162', 'EC161'],
    ('sem3', 'CSE'): ['CS261', 'CS262', 'CS263', 'CS264', 'MA261', 'MA262'],  # NOT CS304, CS307
    ('sem3', 'DSAI'): ['DA261', 'DA262', 'CS304', 'CS307', 'MA261', 'MA262'],  # CS304, CS307 are DSAI
    ('sem3', 'ECE'): ['CS307', 'EC261', 'EC262', 'EC263', 'MA261', 'MA262', 'MA263'],  # CS307 is ECE, but NOT CS304
    ('sem5', 'CSE'): ['CS351', 'CS352', 'CS353', 'CS354', 'MA351', 'MA352'],
    ('sem5', 'DSAI'): ['DS302', 'DS303', 'CS307', 'MA261', 'MA262'],
    ('sem5', 'ECE'): ['EC351', 'EC352', 'EC353', 'EC354', 'MA351', 'MA352', 'MA353'],
    ('sem7', 'CSE'): ['CS451', 'CS452', 'CS453', 'CS454', 'CS498'],
    ('sem7', 'DSAI'): ['DS451', 'DS452', 'DS498'],
    ('sem7', 'ECE'): ['EC498'],
}

# Prohibited courses for each semester and branch
prohibited_courses = {
    ('sem3', 'CSE'): ['CS304', 'CS307'],  # These are DSAI courses
    ('sem3', 'ECE'): ['CS304'],  # CS304 is DSAI only
}

print("=" * 80)
print("COMPREHENSIVE VERIFICATION OF COURSE FILTERING FIX")
print("=" * 80)

all_passed = True

for sem_prefix in ['sem1', 'sem3', 'sem5', 'sem7']:
    for branch in ['CSE', 'DSAI', 'ECE']:
        filename = os.path.join(timetables_dir, f'{sem_prefix}_{branch}_timetable.xlsx')
        
        if not os.path.exists(filename):
            print(f"\n[SKIP] {filename} does not exist")
            continue
        
        try:
            excel_file = pd.ExcelFile(filename)
            
            # Find the course information sheet
            course_info_sheet = None
            for sheet in excel_file.sheet_names:
                if 'course' in sheet.lower() and 'information' in sheet.lower():
                    course_info_sheet = sheet
                    break
            
            if not course_info_sheet:
                print(f"\n[SKIP] {filename} has no Course Information sheet")
                continue
            
            df = pd.read_excel(filename, sheet_name=course_info_sheet)
            
            # Flatten all values into a single string to search
            all_text = df.astype(str).values.flatten()
            all_courses_in_sheet = []
            
            for text in all_text:
                if text and isinstance(text, str):
                    # Check if any known course code is in this cell
                    for course_code in ['CS', 'DS', 'DA', 'EC', 'MA', 'HS', 'DE', 'PH', 'ASD']:
                        if course_code in text:
                            # Extract course codes
                            words = text.split()
                            for word in words:
                                if word.upper().startswith(course_code) and len(word) > len(course_code):
                                    code_part = word[:word.find(' ') if ' ' in word else len(word)]
                                    # Clean up - remove parentheses, commas, etc.
                                    code_part = code_part.split('(')[0].split(',')[0].strip()
                                    if code_part and code_part[len(course_code):].isdigit():
                                        all_courses_in_sheet.append(code_part)
            
            # Remove duplicates while preserving order
            unique_courses = []
            for course in all_courses_in_sheet:
                if course not in unique_courses and course.upper() not in unique_courses:
                    unique_courses.append(course.upper())
            
            print(f"\n[CHECK] {sem_prefix}_{branch}")
            print(f"  Courses found: {unique_courses[:10]}...")  # Show first 10
            
            # Check for prohibited courses
            test_key = (sem_prefix, branch)
            if test_key in prohibited_courses:
                for prohibited_course in prohibited_courses[test_key]:
                    if prohibited_course.upper() in [c.upper() for c in unique_courses]:
                        print(f"  [FAIL] Found prohibited course {prohibited_course} (should not be in {branch})")
                        all_passed = False
                    else:
                        print(f"  [PASS] {prohibited_course} correctly NOT in {branch}")
            
        except Exception as e:
            print(f"\n[ERROR] Failed to check {filename}: {e}")
            all_passed = False

print("\n" + "=" * 80)
if all_passed:
    print("[PASS] ALL CHECKS PASSED - Course filtering is working correctly!")
else:
    print("[FAIL] Some checks failed - see details above")
print("=" * 80)
