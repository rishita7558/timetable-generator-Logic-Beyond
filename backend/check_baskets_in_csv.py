#!/usr/bin/env python3
"""Check which baskets are in which semester in the input data"""
import pandas as pd

df = pd.read_csv('temp_inputs/course_data.csv')
elec = df[df['Elective (Yes/No)'].str.upper() == 'YES']

print("=" * 80)
print("ELECTIVE BASKETS BY SEMESTER IN INPUT DATA")
print("=" * 80)

for sem in [1, 3, 5, 7]:
    sem_elec = elec[elec['Semester'] == sem]
    baskets = sem_elec['Basket'].unique()
    basket_list = sorted([b for b in baskets if pd.notna(b)])
    print(f"\nSemester {sem}: {basket_list}")
    print(f"  Course count: {len(sem_elec)}")
    
    # Show sample courses per basket
    for basket in basket_list:
        courses = sem_elec[sem_elec['Basket'] == basket]['Course Code'].tolist()
        print(f"    {basket}: {courses[:5]}{'...' if len(courses) > 5 else ''}")
