import pandas as pd

print('=== SEMESTER 1 ELECTIVE_B1 ===')
path1 = r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend\output_timetables\sem1_CSE_timetable_baskets.xlsx'
df1 = pd.read_excel(path1, sheet_name='Classroom_Allocation')
print('Columns:', list(df1.columns))
b1 = df1[df1['Basket'] == 'ELECTIVE_B1']
print('Total B1 rows:', len(b1))
if 'Session Type' in df1.columns:
    print('B1 Session types:', b1['Session Type'].value_counts().to_dict())
    tutorials = b1[b1['Session Type'] == 'Tutorial']
    print('B1 Tutorial rows:', len(tutorials))
    print('\nSample B1 tutorials:')
    print(tutorials[['Course', 'Basket', 'Day', 'Time Slot', 'Room Number']].head())

print('\n=== SEMESTER 3 ELECTIVE_B3 ===')
path3 = r'C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend\output_timetables\sem3_CSE_timetable_baskets.xlsx'
df3 = pd.read_excel(path3, sheet_name='Classroom_Allocation')
b3 = df3[df3['Basket'] == 'ELECTIVE_B3']
print('Total B3 rows:', len(b3))
if 'Session Type' in df3.columns:
    print('B3 Session types:', b3['Session Type'].value_counts().to_dict())
    tutorials = b3[b3['Session Type'] == 'Tutorial']
    print('B3 Tutorial rows:', len(tutorials))
    print('\nSample B3 tutorials:')
    print(tutorials[['Course', 'Basket', 'Day', 'Time Slot', 'Room Number']].head())
