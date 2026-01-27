"""Simplified debug script"""
import pandas as pd
import sys
sys.path.insert(0, r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend')

from app import load_all_data, separate_courses_by_type, allocate_electives_by_baskets
from app import generate_section_schedule_with_elective_baskets, allocate_classrooms_for_timetable
from app import get_course_info

# Suppress output
import os, contextlib

with contextlib.redirect_stdout(open(os.devnull, 'w')):
    dfs = load_all_data()

semester, branch = 1, 'CSE'

course_info = get_course_info(dfs)
course_baskets = separate_courses_by_type(dfs, semester, branch)
core_courses = course_baskets.get('core_courses', [])

elective_courses_all = course_baskets.get('elective_courses', [])
elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)

basket_courses_map = {}
for basket_name, alloc in (basket_allocations or {}).items():
    courses_in_basket = alloc.get('all_courses_in_basket', [])
    if courses_in_basket:
        basket_courses_map[basket_name] = courses_in_basket

regular_section_a = generate_section_schedule_with_elective_baskets(
    dfs, semester, 'A', elective_allocations, branch, 
    time_config=None, basket_allocations=basket_allocations
)

classroom_data = dfs.get('classroom')
if classroom_data is not None and not classroom_data.empty:
    regular_section_a = allocate_classrooms_for_timetable(
        regular_section_a, classroom_data, course_info, semester, branch, 'A', basket_courses_map
    )

print(f"\n=== COUNTING LOGIC TEST FOR {branch} SEM{semester} ===")
print(f"DataFrame shape: {regular_section_a.shape}")
print(f"Time slots: {list(regular_section_a.index[:5])}")
print(f"Columns: {list(regular_section_a.columns[:5])}")
print()

# Test on first course
test_course = core_courses[0] if core_courses else None
if not test_course:
    print("No core courses found!")
    sys.exit(1)

print(f"Testing course: {test_course}")
info = course_info.get(test_course, {})
ltpsc = info.get('ltpsc', 'N/A')
print(f"LTPSC: {ltpsc}")

# Parse LTPSC
req_l, req_t, req_p = 0, 0, 0
if ltpsc != 'N/A':
    try:
        parts = ltpsc.split('-')
        if len(parts) >= 3:
            req_l, req_t, req_p = int(parts[0]), int(parts[1]), int(parts[2])
    except:
        pass
print(f"Required: L={req_l}, T={req_t}, P={req_p}")

# Check if in DataFrame
if test_course not in regular_section_a.columns:
    print(f"ERROR: Course {test_course} not in DataFrame!")
    print(f"Available columns: {list(regular_section_a.columns[:10])}")
    sys.exit(1)

print(f"\n✓ Course found in DataFrame")
course_col = regular_section_a[test_course]
print(f"Column length: {len(course_col)}")

# Check non-empty
non_empty = course_col.dropna()
non_empty = non_empty[non_empty != '']
non_free = non_empty[~non_empty.astype(str).str.lower().str.contains('free')]
print(f"Non-empty entries: {len(non_empty)}")
print(f"Non-'free' entries: {len(non_free)}")

# Show entries
print(f"\nEntries with non-empty values:")
for idx in range(len(course_col)):
    val = course_col.iloc[idx]
    if pd.notna(val) and str(val).strip():
        time_slot = str(regular_section_a.index[idx])
        print(f"  [{idx:2d}] {time_slot:20s} -> {str(val)[:40]}")

# Now test the counting logic
print(f"\nCounting logic:")
sched_l, sched_t, sched_p = 0, 0, 0
processed = set()

for idx in range(len(course_col)):
    if idx in processed:
        continue
    
    entry = course_col.iloc[idx]
    if pd.isna(entry) or entry == '' or 'free' in str(entry).lower():
        continue
    
    time_slot_label = str(regular_section_a.index[idx]).lower()
    
    # Tutorial detection
    if any(t in time_slot_label for t in ['14:30-15:30', '17:00-18:00', '18:00-18:30', '18:30-20:00']):
        sched_t += 1
        processed.add(idx)
        print(f"  [{idx:2d}] -> TUTORIAL")
    else:
        # Check consecutive
        if idx + 1 < len(course_col):
            next_entry = course_col.iloc[idx + 1]
            if pd.notna(next_entry) and str(next_entry) != '' and 'free' not in str(next_entry).lower():
                sched_p += 2
                processed.add(idx)
                processed.add(idx + 1)
                print(f"  [{idx:2d}] -> LAB (2 hours)")
            else:
                sched_l += 1
                processed.add(idx)
                print(f"  [{idx:2d}] -> LECTURE")
        else:
            sched_l += 1
            processed.add(idx)
            print(f"  [{idx:2d}] -> LECTURE")

print(f"\nFinal counts: L={sched_l}/{req_l}, T={sched_t}/{req_t}, P={sched_p}/{req_p}")
