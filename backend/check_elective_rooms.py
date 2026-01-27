#!/usr/bin/env python3
"""Check elective room allocations in the consolidated timetable"""
import pandas as pd
import re

# Read Regular Section A and B
df_a = pd.read_excel("output_timetables/sem1_CSE_timetable.xlsx", sheet_name="Regular_Section_A")
df_b = pd.read_excel("output_timetables/sem1_CSE_timetable.xlsx", sheet_name="Regular_Section_B")

print("=" * 80)
print("ELECTIVE ROOM ALLOCATION CHECK - SEM1 CSE")
print("=" * 80)

# Function to extract course and room from cell
def extract_course_room(cell_value):
    if pd.isna(cell_value):
        return None, None
    text = str(cell_value)
    # Pattern: COURSE_CODE [ROOM]
    match = re.search(r'([A-Z]{2}\d{3})\s*\[([^\]]+)\]', text)
    if match:
        return match.group(1), match.group(2)
    return None, None

# Check each time slot
print("\nSection A - Elective Baskets:")
for idx, row in df_a.iterrows():
    time_slot = row.get('Time Slot', idx)
    for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']:
        if day in df_a.columns:
            cell_value = row[day]
            if pd.notna(cell_value) and 'ELECTIVE_B1' in str(cell_value).upper():
                # Extract individual courses
                text = str(cell_value)
                # Find all course-room pairs
                matches = re.findall(r'([A-Z]{2}\d{3})\s*\[([^\]]+)\]', text)
                if matches:
                    print(f"\n{day} {time_slot}:")
                    for course, room in matches:
                        print(f"  {course} -> {room}")
                    # Check for duplicates
                    rooms = [room for _, room in matches]
                    if len(rooms) != len(set(rooms)):
                        print(f"  ❌ CONFLICT: Multiple courses assigned to same room!")

print("\n" + "=" * 80)
print("Section B - Elective Baskets:")
for idx, row in df_b.iterrows():
    time_slot = row.get('Time Slot', idx)
    for day in ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']:
        if day in df_b.columns:
            cell_value = row[day]
            if pd.notna(cell_value) and 'ELECTIVE_B1' in str(cell_value).upper():
                # Extract individual courses
                text = str(cell_value)
                # Find all course-room pairs
                matches = re.findall(r'([A-Z]{2}\d{3})\s*\[([^\]]+)\]', text)
                if matches:
                    print(f"\n{day} {time_slot}:")
                    for course, room in matches:
                        print(f"  {course} -> {room}")
                    # Check for duplicates
                    rooms = [room for _, room in matches]
                    if len(rooms) != len(set(rooms)):
                        print(f"  ❌ CONFLICT: Multiple courses assigned to same room!")
