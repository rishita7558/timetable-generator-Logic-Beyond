#!/usr/bin/env python3
"""
Debug script to understand the timetable structure
"""
import os
import sys
sys.path.insert(0, os.path.dirname(__file__))

from app import load_all_data, separate_courses_by_type, generate_semester_schedule

dfs = load_all_data()

# Generate semester 1, CSE schedule to inspect the structure
semester = 1
branch = 'CSE'

semester_courses = dfs['course'][(dfs['course']['Semester'] == semester)].copy()
core_courses, elective_courses, minor_courses = separate_courses_by_type(dfs, semester, branch)

print("=" * 80)
print(f"SEMESTER {semester}, BRANCH {branch}")
print("=" * 80)

print(f"\nCore courses: {core_courses[:3]}")  # Print first 3

# Generate schedule for section A
schedule_a = generate_semester_schedule(
    dfs=dfs,
    semester_id=semester,
    section='A',
    courses=semester_courses,
    branch=branch,
    time_config=None
)

if schedule_a is not None and not schedule_a.empty:
    print(f"\nSchedule shape: {schedule_a.shape}")
    print(f"Schedule columns (first 5): {schedule_a.columns.tolist()[:5]}")
    print(f"\nSchedule index (time slots): {schedule_a.index.tolist()}")
    
    # Pick a core course and see what it looks like
    if core_courses:
        test_course = core_courses[0]
        if test_course in schedule_a.columns:
            print(f"\n\nCourse {test_course} entries:")
            for idx, val in schedule_a[test_course].items():
                if val and str(val).strip() != '':
                    print(f"  {idx}: {val}")
        else:
            print(f"\nCourse {test_course} not in schedule")
            print(f"Available courses: {[c for c in schedule_a.columns if c in core_courses][:3]}")
else:
    print("\nSchedule is empty!")
