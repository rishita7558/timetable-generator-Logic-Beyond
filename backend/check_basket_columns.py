#!/usr/bin/env python3
"""Check if basket columns are populated correctly in timetables"""
import pandas as pd
import os

timetables_dir = 'output_timetables'
files_to_check = ['sem1_CSE_timetable.xlsx', 'sem3_CSE_timetable.xlsx', 'sem5_DSAI_timetable.xlsx', 'sem7_ECE_timetable.xlsx']

print("=" * 80)
print("CHECKING BASKET COLUMN IN ELECTIVE COURSES")
print("=" * 80)

for filename in files_to_check:
    filepath = os.path.join(timetables_dir, filename)
    if not os.path.exists(filepath):
        print(f"\n[SKIP] {filename} not found")
        continue
    
    print(f"\n{'=' * 80}")
    print(f"FILE: {filename}")
    print('=' * 80)
    
    try:
        df = pd.read_excel(filepath, sheet_name='Course_Information')
        
        # Find the elective courses section
        elec_idx = None
        for idx, row in df.iterrows():
            if any('ELECTIVE COURSES' in str(val).upper() for val in row.values if pd.notna(val)):
                elec_idx = idx
                break
        
        if elec_idx is not None:
            # Show the next 10 rows after the "ELECTIVE COURSES" header
            print("\nElective Courses Section:")
            print(df.iloc[elec_idx:elec_idx+12].to_string())
            
            # Check if basket column has values
            basket_col_idx = None
            header_row = df.iloc[elec_idx + 1]
            for col_idx, val in enumerate(header_row.values):
                if pd.notna(val) and 'BASKET' in str(val).upper():
                    basket_col_idx = col_idx
                    break
            
            if basket_col_idx is not None:
                print(f"\n[✓] Basket column found at column index {basket_col_idx}")
                # Check if there are non-empty basket values
                basket_values = df.iloc[elec_idx+2:elec_idx+12, basket_col_idx]
                non_empty = basket_values[basket_values.notna()].tolist()
                print(f"[INFO] Sample basket values: {non_empty[:5]}")
                if non_empty:
                    print(f"[✓] Basket column is populated with values!")
                else:
                    print(f"[✗] Basket column is EMPTY!")
            else:
                print(f"[✗] Basket column header NOT FOUND")
        else:
            print("[SKIP] No elective courses section found")
            
    except Exception as e:
        print(f"[ERROR] Failed to read {filename}: {e}")

print("\n" + "=" * 80)
print("CHECK COMPLETE")
print("=" * 80)
