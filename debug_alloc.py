from backend.app import load_all_data, allocate_classrooms_for_timetable, _TIMETABLE_CLASSROOM_ALLOCATIONS
import pandas as pd, os, pprint
INPUT_DIR = os.path.join(os.getcwd(), 'backend', 'temp_inputs')
load_all_data(force_reload=True)
schedule = pd.DataFrame(index=['09:00-10:30'], columns=['Mon']).fillna('Free')
schedule.loc['09:00-10:30', 'Mon'] = 'CS300'
course_info = {'CS300': {'semester': '3', 'branch': 'Computer Science and Engineering', 'is_elective': False}}
_TIMETABLE_CLASSROOM_ALLOCATIONS.clear()
classrooms_df = pd.read_csv(os.path.join(INPUT_DIR, 'classroom_data.csv'))
res = allocate_classrooms_for_timetable(schedule, classrooms_df, course_info, semester='3', branch='CSE', section='A')
print('SCHEDULE_CELL:', res.loc['09:00-10:30', 'Mon'])
print('\n--- ALLOC MAP ---')
for k, v in _TIMETABLE_CLASSROOM_ALLOCATIONS.items():
    print('TIMETABLE KEY:', k)
    for rec_key, rec in v.items():
        print(rec_key, rec)
