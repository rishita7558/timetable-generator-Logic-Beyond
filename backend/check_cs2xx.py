import openpyxl
import sys

wb = openpyxl.load_workbook(r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend\output_timetables\sem3_CSE_timetable.xlsx')
ws = wb.worksheets[0]

print("All courses in CORE COURSES section:")
found_cs262 = False
for row in ws.iter_rows(min_row=1, max_row=100, max_col=7, values_only=True):
    if row and row[0]:
        course = str(row[0])
        if 'CS2' in course:  # Any CS2xx course
            print(f"  {course}: LTPSC={row[2]}, L={row[4]}, T={row[5]}, P={row[6]}")
            if 'CS262' in course:
                found_cs262 = True

if not found_cs262:
    print("\nCS262 NOT FOUND in the file!")
    sys.exit(1)
