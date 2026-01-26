import os
import shutil
import pandas as pd
from backend.app import (
    load_all_data,
    estimate_course_enrollment,
    allocate_classrooms_for_timetable,
    _TIMETABLE_CLASSROOM_ALLOCATIONS,
    separate_courses_by_mid_semester,
)

INPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'temp_inputs')


def _backup_and_write_csv(filename, content):
    orig_path = os.path.join(INPUT_DIR, filename)
    backup_path = orig_path + '.bak'
    if os.path.exists(orig_path):
        shutil.copy(orig_path, backup_path)
    with open(orig_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return backup_path if os.path.exists(backup_path) else None


def _restore_csv(filename, backup_path):
    path = os.path.join(INPUT_DIR, filename)
    if backup_path and os.path.exists(backup_path):
        shutil.move(backup_path, path)
    else:
        if os.path.exists(path):
            os.remove(path)


def test_estimate_enrollment_from_student_data(tmp_path):
    # Create student data with 30 students in Semester 3, Department ECE
    rows = ['Roll No,Name,Semester,Department'] + [f'{i},Student{i},3,ECE' for i in range(30)]
    backup = _backup_and_write_csv('student_data.csv', '\n'.join(rows) + '\n')

    # Provide minimal course_info
    course_info = {'EC100': {'semester': '3', 'branch': 'Electronics and Communication Engineering', 'is_elective': False}}

    # Force reload data
    dfs = load_all_data(force_reload=True)
    estimates = estimate_course_enrollment(course_info)

    assert estimates['EC100'] == 30

    _restore_csv('student_data.csv', backup)


def test_allocate_multiple_rooms_for_large_enrollment(tmp_path):
    # Create student data with 120 students in Semester 3, Department CSE
    rows = ['Roll No,Name,Semester,Department'] + [f'{i},Student{i},3,CSE' for i in range(120)]
    backup_students = _backup_and_write_csv('student_data.csv', '\n'.join(rows) + '\n')

    # Create classroom data with multiple rooms smaller than total capacity
    csv_rooms = 'Room Number,Type,Capacity,Facilities\nC201,classroom,50,Proj\nC202,classroom,50,Proj\nC203,classroom,50,Proj\n'
    backup_rooms = _backup_and_write_csv('classroom_data.csv', csv_rooms)

    # Build a simple schedule dataframe
    import pandas as pd
    schedule = pd.DataFrame(index=['09:00-10:30'], columns=['Mon'], dtype=object).fillna('Free')
    schedule.loc['09:00-10:30', 'Mon'] = 'CS300'  # core course

    # Course info for CS300 in sem 3, CSE
    course_info = {'CS300': {'semester': '3', 'branch': 'Computer Science and Engineering', 'is_elective': False}}

    # Clear allocations tracker
    _TIMETABLE_CLASSROOM_ALLOCATIONS.clear()

    # Read classrooms df from input
    classrooms_df = pd.read_csv(os.path.join(INPUT_DIR, 'classroom_data.csv'))

    schedule_with_rooms = allocate_classrooms_for_timetable(schedule, classrooms_df, course_info, semester='3', branch='CSE', section='A')

    cell = schedule_with_rooms.loc['09:00-10:30', 'Mon']
    # Expect multiple rooms listed for the course
    assert 'CS300' in cell
    # There should be multiple room numbers in brackets
    assert '[' in cell and ']' in cell

    # Verify TIMETABLE allocations contain split entries for CS300
    found_split = False
    for k, v in _TIMETABLE_CLASSROOM_ALLOCATIONS.items():
        for rec_key, rec in v.items():
            if rec.get('course') and 'CS300' in rec.get('course'):
                if rec.get('split', False):
                    found_split = True

    assert found_split

    _restore_csv('student_data.csv', backup_students)
    _restore_csv('classroom_data.csv', backup_rooms)


def test_separate_courses_by_mid_semester_covers_all_cases():
    courses = [
        # Full-semester course -> both pre and post
        {
            'Course Code': 'CS101',
            'Semester': '3',
            'Department': 'CSE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'NO',
            'Post mid-sem': 'NO',
            'Common': 'No',
        },
        # Half-semester, post-mid only
        {
            'Course Code': 'CS102',
            'Semester': '3',
            'Department': 'CSE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'YES',
            'Post mid-sem': 'YES',
            'Common': 'No',
        },
        # Half-semester, pre-mid only because Post mid-sem is blank
        {
            'Course Code': 'CS103',
            'Semester': '3',
            'Department': 'CSE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'YES',
            'Post mid-sem': '',
            'Common': 'No',
        },
        # Half-semester, pre-mid only because Post mid-sem is NO
        {
            'Course Code': 'CS104',
            'Semester': '3',
            'Department': 'CSE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'YES',
            'Post mid-sem': 'NO',
            'Common': 'No',
        },
        # Common course from another department should still appear for CSE
        {
            'Course Code': 'CS105',
            'Semester': '3',
            'Department': 'ECE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'NO',
            'Post mid-sem': 'NO',
            'Common': 'YES',
        },
        # Different semester to ensure filtering by semester works
        {
            'Course Code': 'CS106',
            'Semester': '5',
            'Department': 'CSE',
            'Elective (Yes/No)': 'No',
            'Half Semester': 'NO',
            'Post mid-sem': 'NO',
            'Common': 'No',
        },
    ]

    dfs = {'course': pd.DataFrame(courses)}

    result = separate_courses_by_mid_semester(dfs, semester_id=3, branch='CSE')
    pre_mid_codes = set(result['pre_mid_courses']['Course Code'])
    post_mid_codes = set(result['post_mid_courses']['Course Code'])

    # Full-semester and common courses appear in both
    assert 'CS101' in pre_mid_codes and 'CS101' in post_mid_codes
    assert 'CS105' in pre_mid_codes and 'CS105' in post_mid_codes

    # Half-semester pre-only cases
    assert 'CS103' in pre_mid_codes and 'CS103' not in post_mid_codes
    assert 'CS104' in pre_mid_codes and 'CS104' not in post_mid_codes

    # Half-semester post-only case
    assert 'CS102' in post_mid_codes and 'CS102' not in pre_mid_codes

    # Other semesters should be excluded entirely
    assert 'CS106' not in pre_mid_codes and 'CS106' not in post_mid_codes
