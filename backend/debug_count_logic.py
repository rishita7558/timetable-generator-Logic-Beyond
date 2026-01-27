"""Debug script to test the counting logic"""
import pandas as pd
import sys
sys.path.insert(0, r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend')

from app import export_consolidated_semester_timetable, load_all_data
from pathlib import Path

# Load data
dfs = load_all_data()
time_config = None

# Get one semester and branch for testing
semester = 1
branch = 'CSE'

print(f"\n=== TESTING COUNTING LOGIC FOR {branch} SEM{semester} ===\n")

# Import required functions
from app import (
    get_course_info, separate_courses_by_type, allocate_electives_by_baskets,
    generate_section_schedule_with_elective_baskets, allocate_classrooms_for_timetable
)

course_info = get_course_info(dfs) if dfs else {}
course_baskets_all = separate_courses_by_type(dfs, semester, branch)
elective_courses_all = course_baskets_all['elective_courses']
elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)

# Get classroom data
classroom_data = dfs.get('classroom')
basket_courses_map = {}
for basket_name, alloc in (basket_allocations or {}).items():
    courses_in_basket = alloc.get('all_courses_in_basket', [])
    if courses_in_basket:
        basket_courses_map[basket_name] = courses_in_basket

# Generate schedule
regular_section_a = generate_section_schedule_with_elective_baskets(
    dfs, semester, 'A', elective_allocations, branch, 
    time_config=time_config, basket_allocations=basket_allocations
)

# Allocate classrooms
if classroom_data is not None and not classroom_data.empty:
    regular_section_a = allocate_classrooms_for_timetable(
        regular_section_a, classroom_data, course_info, semester, branch, 'A', basket_courses_map
    )

# Get core courses for the semester
core_courses = course_baskets_all.get('core_courses', [])

print(f"Core courses for {branch} SEM{semester}: {core_courses[:5]}...")  # Show first 5
print(f"\nTimetable DataFrame shape: {regular_section_a.shape}")
print(f"Timetable index (time slots): {list(regular_section_a.index[:5])}...")  # First 5 time slots
print(f"Timetable columns: {list(regular_section_a.columns[:5])}...")  # First 5 courses
print(f"\n" + "="*80 + "\n")

# Test counting logic on first 3 courses
test_courses = core_courses[:3] if core_courses else []

for course_code in test_courses:
    print(f"\n--- Testing course: {course_code} ---")
    
    info = course_info.get(course_code, {})
    ltpsc = info.get('ltpsc', 'N/A')
    print(f"LTPSC: {ltpsc}")
    
    # Parse LTPSC
    req_lectures, req_tutorials, req_labs = 0, 0, 0
    if ltpsc != 'N/A':
        try:
            parts = ltpsc.split('-')
            if len(parts) >= 3:
                req_lectures = int(parts[0])
                req_tutorials = int(parts[1])
                req_labs = int(parts[2])
        except (ValueError, IndexError):
            pass
    
    print(f"Required: L={req_lectures}, T={req_tutorials}, P={req_labs}")
    
    # Check if course is in DataFrame
    if course_code in regular_section_a.columns:
        print(f"✓ Course found in DataFrame")
        course_entries = regular_section_a[course_code]
        print(f"  Column length: {len(course_entries)}")
        
        # Count non-empty entries
        non_empty = course_entries.dropna().drop('')
        non_free = non_empty[~non_empty.astype(str).str.lower().str.contains('free', na=False)]
        print(f"  Non-empty entries: {len(non_empty)}")
        print(f"  Non-free entries: {len(non_free)}")
        
        # Show first few entries
        print(f"\n  First 10 entries in column:")
        for idx in range(min(10, len(course_entries))):
            entry = course_entries.iloc[idx]
            time_slot = str(regular_section_a.index[idx])
            if pd.notna(entry) and str(entry).strip() != '':
                print(f"    [{idx:2d}] {time_slot:20s} -> {str(entry)[:50]}")
            else:
                print(f"    [{idx:2d}] {time_slot:20s} -> [EMPTY]")
        
        # Now run the actual counting logic
        print(f"\n  Running counting logic:")
        sched_lectures = 0
        sched_tutorials = 0
        sched_labs = 0
        processed_idx = set()
        
        for idx in range(len(course_entries)):
            if idx in processed_idx:
                continue
                
            entry = course_entries.iloc[idx]
            
            # Skip empty or "Free" entries
            if pd.isna(entry) or entry == '' or 'free' in str(entry).lower():
                continue
            
            print(f"    Processing idx={idx}: {str(entry)[:40]}")
            
            entry_str = str(entry).lower()
            
            # Get time slot
            time_slot_idx = idx
            if time_slot_idx < len(regular_section_a.index):
                time_slot_label = str(regular_section_a.index[time_slot_idx]).lower()
                print(f"      Time slot: {time_slot_label}")
                
                # Determine if it's a tutorial slot (single hour slots)
                tutorial_keywords = ['14:30-15:30', '17:00-18:00', '18:00-18:30', '18:30-20:00']
                is_tutorial = any(t in time_slot_label for t in tutorial_keywords)
                print(f"      Is tutorial? {is_tutorial}")
                
                if is_tutorial:
                    sched_tutorials += 1
                    print(f"      -> Counted as TUTORIAL")
                    processed_idx.add(idx)
                else:
                    # Check if next slot also has the course (for lab detection)
                    if idx + 1 < len(course_entries):
                        next_entry = course_entries.iloc[idx + 1]
                        if pd.notna(next_entry) and str(next_entry) != '' and 'free' not in str(next_entry).lower():
                            # Consecutive slots = lab (2 hours)
                            sched_labs += 2
                            processed_idx.add(idx)
                            processed_idx.add(idx + 1)
                            print(f"      -> Counted as LAB (with next slot)")
                        else:
                            # Single slot = lecture
                            sched_lectures += 1
                            processed_idx.add(idx)
                            print(f"      -> Counted as LECTURE")
                    else:
                        # Last slot = lecture
                        sched_lectures += 1
                        processed_idx.add(idx)
                        print(f"      -> Counted as LECTURE (last slot)")
        
        print(f"\n  Final counts: L={sched_lectures}, T={sched_tutorials}, P={sched_labs}")
        print(f"  Display format: L={sched_lectures}/{req_lectures}, T={sched_tutorials}/{req_tutorials}, P={sched_labs}/{req_labs}")
    
    else:
        print(f"✗ Course NOT found in DataFrame")
