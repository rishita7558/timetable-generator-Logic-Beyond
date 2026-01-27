"""Check LTPSC for problem courses"""
import sys
sys.path.insert(0, '.')
from app import load_all_data, get_course_info

dfs = load_all_data() if False else None

# Load just course info
course_data = pd.read_csv('temp_inputs/course_data.csv')

problem_courses = ['CS307', 'EC261', 'MA261', 'MA262']

print("LTPSC values for problem courses:")
print("="*60)
for course in problem_courses:
    row = course_data[course_data['Course Code'] == course]
    if len(row) > 0:
        ltpsc = row['LTPSC'].values[0]
        parts = str(ltpsc).split('-')
        l, t, p = int(parts[0]), int(parts[1]), int(parts[2])
        print(f"{course}: LTPSC={ltpsc} (L={l}, T={t}, P={p})")
    else:
        print(f"{course}: NOT FOUND")
