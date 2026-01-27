"""Test the actual counting logic on EC261"""
import pandas as pd
import sys
import os, contextlib
sys.path.insert(0, '.')

from app import load_all_data, separate_courses_by_type, allocate_electives_by_baskets
from app import generate_section_schedule_with_elective_baskets, allocate_classrooms_for_timetable, get_course_info

with contextlib.redirect_stdout(open(os.devnull, 'w')):
    dfs = load_all_data()

semester, branch = 3, 'ECE'
course_code = 'EC261'

course_info = get_course_info(dfs)
course_baskets = separate_courses_by_type(dfs, semester, branch)

elective_courses_all = course_baskets.get('elective_courses', [])
elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)

basket_courses_map = {}
for basket_name, alloc in (basket_allocations or {}).items():
    courses_in_basket = alloc.get('all_courses_in_basket', [])
    if courses_in_basket:
        basket_courses_map[basket_name] = courses_in_basket

regular_section = generate_section_schedule_with_elective_baskets(
    dfs, semester, 'Whole', elective_allocations, branch, 
    time_config=None, basket_allocations=basket_allocations
)

classroom_data = dfs.get('classroom')
if classroom_data is not None and not classroom_data.empty:
    regular_section = allocate_classrooms_for_timetable(
        regular_section, classroom_data, course_info, semester, branch, 'Whole', basket_courses_map
    )

print(f"Timetable structure: {regular_section.shape}")
print(f"Index: {list(regular_section.index)}")
print(f"Columns: {list(regular_section.columns)}")

print(f"\n\nRaw timetable for {course_code}:")
print("="*80)
for slot_idx, time_slot in enumerate(regular_section.index):
    for col_idx, day in enumerate(regular_section.columns):
        cell = regular_section.loc[time_slot, day]
        if pd.notna(cell) and isinstance(cell, str) and course_code in cell:
            print(f"[{slot_idx}] {str(time_slot):20s} {day:5s}: {cell[:50]}")

print(f"\n\nNow running the actual counting logic:")
print("="*80)

# Parse LTPSC
info = course_info.get(course_code, {})
ltpsc = info.get('ltpsc', 'N/A')
try:
    parts = ltpsc.split('-')
    req_l = int(parts[0])
    req_t = int(parts[1])
    req_p = int(parts[2])
except:
    req_l, req_t, req_p = 0, 0, 0

print(f"LTPSC: {ltpsc} -> L={req_l}, T={req_t}, P={req_p}")

# Run the counting logic
sched_l, sched_t, sched_p = 0, 0, 0
processed_cells = set()

for slot_idx, time_slot in enumerate(regular_section.index):
    time_slot_str = str(time_slot).lower()
    
    for day in regular_section.columns:
        if (slot_idx, day) in processed_cells:
            continue
        
        cell_value = regular_section.loc[time_slot, day]
        
        if pd.isna(cell_value) or cell_value == '' or 'free' in str(cell_value).lower():
            continue
        
        cell_str = str(cell_value).lower()
        
        if course_code.lower() not in cell_str:
            continue
        
        print(f"\n[{slot_idx}] {str(time_slot):20s} {day:5s}: {str(cell_value)[:50]}")
        
        # Check for explicit markers
        if '(lab)' in cell_str:
            print(f"    -> Has (Lab) marker")
            if slot_idx + 1 < len(regular_section.index):
                next_slot = regular_section.index[slot_idx + 1]
                next_cell = regular_section.loc[next_slot, day]
                
                if pd.notna(next_cell) and course_code.lower() in str(next_cell).lower() and '(lab)' in str(next_cell).lower():
                    sched_p += 2
                    processed_cells.add((slot_idx, day))
                    processed_cells.add((slot_idx + 1, day))
                    print(f"    -> LAB (2 hours, consecutive)")
                else:
                    sched_p += 1
                    processed_cells.add((slot_idx, day))
                    print(f"    -> LAB (1 hour)")
            else:
                sched_p += 1
                processed_cells.add((slot_idx, day))
                print(f"    -> LAB (1 hour, last)")
        elif '(tutorial)' in cell_str:
            sched_t += 1
            processed_cells.add((slot_idx, day))
            print(f"    -> Has (Tutorial) marker -> TUTORIAL")
        else:
            is_tutorial = any(t in time_slot_str for t in ['14:30-15:30', '17:00-18:00', '18:00-18:30', '18:30-20:00'])
            
            if is_tutorial:
                sched_t += 1
                processed_cells.add((slot_idx, day))
                print(f"    -> Tutorial time slot -> TUTORIAL")
            else:
                if slot_idx + 1 < len(regular_section.index):
                    next_slot = regular_section.index[slot_idx + 1]
                    next_cell = regular_section.loc[next_slot, day]
                    
                    if pd.notna(next_cell) and str(next_cell) != '' and 'free' not in str(next_cell).lower():
                        if course_code.lower() in str(next_cell).lower():
                            sched_p += 2
                            processed_cells.add((slot_idx, day))
                            processed_cells.add((slot_idx + 1, day))
                            print(f"    -> Consecutive slots -> LAB (2 hours)")
                        else:
                            sched_l += 1
                            processed_cells.add((slot_idx, day))
                            print(f"    -> Next slot has different course -> LECTURE")
                    else:
                        sched_l += 1
                        processed_cells.add((slot_idx, day))
                        print(f"    -> Next slot free -> LECTURE")
                else:
                    sched_l += 1
                    processed_cells.add((slot_idx, day))
                    print(f"    -> Last slot -> LECTURE")

# Adjustment
if req_l in [2, 3] and sched_l >= 2:
    sched_l = req_l
    print(f"\nAdjustment: L={sched_l} (matched to required)")

print(f"\n\nFinal counts: L={sched_l}, T={sched_t}, P={sched_p}")
print(f"Required:     L={req_l}, T={req_t}, P={req_p}")
print(f"Display:      L={sched_l}/{req_l}, T={sched_t}/{req_t}, P={sched_p}/{req_p}")
