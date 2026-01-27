import openpyxl

wb = openpyxl.load_workbook(r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend\output_timetables\sem3_CSE_timetable.xlsx')
ws = wb.worksheets[0]

print("Searching for CS262 in CORE COURSES...")
for row_idx, row in enumerate(ws.iter_rows(min_row=40, max_row=100, values_only=True), start=40):
    if row and row[0] and 'CS262' in str(row[0]):
        print(f"Row {row_idx}: {row}")
        print(f"  Course: {row[0]}")
        print(f"  Name: {row[1]}")
        print(f"  LTPSC: {row[2]}")
        print(f"  Term: {row[3]}")
        print(f"  Lectures: {row[4]}")
        print(f"  Tutorials: {row[5]}")
        print(f"  Labs: {row[6]}")
