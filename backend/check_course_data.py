#!/usr/bin/env python3
import pandas as pd

# Load course data
course_df = pd.read_csv('temp_inputs/course_data.csv')

# Look at CS304 and CS307
print("=" * 80)
print("CS304 and CS307 in course data:")
print("=" * 80)

for code in ['CS304', 'CS307']:
    matches = course_df[course_df['Course Code'] == code]
    if not matches.empty:
        for _, row in matches.iterrows():
            print(f"\nCourse Code: {row['Course Code']}")
            print(f"  Semester: {row['Semester']}")
            print(f"  Department: {row['Department']}")
            print(f"  Elective: {row['Elective (Yes/No)']}")
            print(f"  Half Semester: {row.get('Half Semester (Yes/No)', 'N/A')}")
            print(f"  Post mid-sem: {row.get('Post mid-sem', 'N/A')}")
    else:
        print(f"\n{code}: NOT FOUND")

print("\n" + "=" * 80)
print("All semester 3 courses with their departments and elective status:")
print("=" * 80)

sem3 = course_df[course_df['Semester'] == 3]
for _, row in sem3.iterrows():
    print(f"{row['Course Code']:6} - Dept: {row['Department']:30} - Elective: {row['Elective (Yes/No)']}")
