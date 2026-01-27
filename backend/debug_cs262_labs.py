"""Debug CS262 lab counting in sem3"""
import pandas as pd
import sys
import os, contextlib
sys.path.insert(0, r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend')

from app import load_all_data, separate_courses_by_type, allocate_electives_by_baskets
from app import generate_section_schedule_with_elective_baskets, allocate_classrooms_for_timetable
from app import get_course_info

with contextlib.redirect_stdout(open(os.devnull, 'w')):
    dfs = load_all_data()

semester, branch = 3, 'CSE'
course_code = 'CS262'

course_info = get_course_info(dfs)
course_baskets = separate_courses_by_type(dfs, semester, branch)

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

print(f"\n=== DEBUGGING {course_code} IN SEM{semester} {branch} ===\n")
print(f"Timetable shape: {regular_section_a.shape}")

info = course_info.get(course_code, {})
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
print(f"Required: L={req_l}, T={req_t}, P={req_p}\n")

# Show the full schedule
print("Full timetable for Section A:")
print(regular_section_a.to_string())
print("\n" + "="*80 + "\n")

# Check all cells containing CS262
print(f"All cells containing '{course_code}':")
found_count = 0
for slot_idx, time_slot in enumerate(regular_section_a.index):
    for day in regular_section_a.columns:
        cell_value = regular_section_a.loc[time_slot, day]
        if pd.notna(cell_value) and course_code in str(cell_value):
            print(f"  [{slot_idx:2d}] {str(time_slot):20s} {day:5s} -> {str(cell_value)}")
            found_count += 1

print(f"\nTotal cells with {course_code}: {found_count}\n")

# Now run the counting logic
print("Running counting logic:")
sched_l, sched_t, sched_p = 0, 0, 0
processed_slots = set()

for slot_idx, time_slot in enumerate(regular_section_a.index):
    if slot_idx in processed_slots:
        continue
    
    time_slot_str = str(time_slot).lower()
    
    for day in regular_section_a.columns:
        cell_value = regular_section_a.loc[time_slot, day]
        
        if pd.isna(cell_value) or cell_value == '' or 'free' in str(cell_value).lower():
            continue
        
        cell_str = str(cell_value).lower()
        
        if course_code.lower() not in cell_str:
            continue
        
        print(f"\n  [{slot_idx:2d}] {str(time_slot):20s} {day:5s}")
        print(f"       Value: {str(cell_value)[:60]}")
        print(f"       Time slot: {time_slot_str}")
        
        # PRIORITY 1: Check for explicit markers
        if '(lab)' in cell_str:
            print(f"       Has (Lab) marker!")
            if slot_idx + 1 < len(regular_section_a.index):
                next_slot = regular_section_a.index[slot_idx + 1]
                next_cell = regular_section_a.loc[next_slot, day]
                print(f"       Next slot [{slot_idx+1}]: {next_slot}")
                print(f"       Next cell: {str(next_cell)[:60]}")
                
                if pd.notna(next_cell) and course_code.lower() in str(next_cell).lower() and '(lab)' in str(next_cell).lower():
                    sched_p += 2
                    processed_slots.add(slot_idx)
                    processed_slots.add(slot_idx + 1)
                    print(f"       -> LAB (2 consecutive lab-marked slots)")
                    break
                else:
                    sched_p += 1
                    processed_slots.add(slot_idx)
                    print(f"       -> LAB (single slot)")
                    break
            else:
                sched_p += 1
                processed_slots.add(slot_idx)
                print(f"       -> LAB (last slot)")
                break
        elif '(tutorial)' in cell_str:
            sched_t += 1
            processed_slots.add(slot_idx)
            print(f"       Has (Tutorial) marker -> TUTORIAL")
            break
        else:
            # Tutorial detection
            is_tutorial = any(t in time_slot_str for t in ['14:30-15:30', '17:00-18:00', '18:00-18:30', '18:30-20:00'])
            print(f"       Is tutorial slot? {is_tutorial}")
            
            if is_tutorial:
                sched_t += 1
                processed_slots.add(slot_idx)
                print(f"       -> TUTORIAL")
                break
            else:
                # Check consecutive
                if slot_idx + 1 < len(regular_section_a.index):
                    next_slot = regular_section_a.index[slot_idx + 1]
                    next_cell = regular_section_a.loc[next_slot, day]
                    
                    print(f"       Next slot [{slot_idx+1}]: {next_slot}")
                    print(f"       Next cell: {str(next_cell)[:60]}")
                    
                    if pd.notna(next_cell) and str(next_cell) != '' and 'free' not in str(next_cell).lower():
                        if course_code.lower() in str(next_cell).lower():
                            sched_p += 2
                            processed_slots.add(slot_idx)
                            processed_slots.add(slot_idx + 1)
                            print(f"       -> LAB (consecutive, 2 hours)")
                            break
                        else:
                            sched_l += 1
                            processed_slots.add(slot_idx)
                            print(f"       -> LECTURE (next slot has different course)")
                            break
                    else:
                        sched_l += 1
                        processed_slots.add(slot_idx)
                        print(f"       -> LECTURE (next slot is free)")
                        break
                else:
                    sched_l += 1
                    processed_slots.add(slot_idx)
                    print(f"       -> LECTURE (last slot)")
                    break

print(f"\n" + "="*80)
print(f"FINAL COUNTS: L={sched_l}, T={sched_t}, P={sched_p}")
print(f"REQUIRED:     L={req_l}, T={req_t}, P={req_p}")

# Apply adjustment
if req_l in [2, 3] and sched_l >= 2:
    sched_l = req_l

print(f"AFTER ADJUST: L={sched_l}, T={sched_t}, P={sched_p}")
print(f"DISPLAY:      L={sched_l}/{req_l}, T={sched_t}/{req_t}, P={sched_p}/{req_p}")
