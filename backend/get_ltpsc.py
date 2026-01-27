import pandas as pd

courses = ['CS307', 'EC261', 'MA261', 'MA262']
df = pd.read_csv('temp_inputs/course_data.csv')

for c in courses:
    row = df[df['Course Code']==c]
    if len(row) > 0:
        ltpsc = row['LTPSC'].values[0]
        print(f"{c}: {ltpsc}")
