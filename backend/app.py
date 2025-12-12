from flask import Flask, render_template, request, jsonify, send_file
import os
import re
import pandas as pd
import random
import zipfile
import glob
from werkzeug.utils import secure_filename
from datetime import datetime, timedelta
import math
import traceback
import shutil
import time
import hashlib

app = Flask(__name__)
_SEMESTER_ELECTIVE_ALLOCATIONS = {}

# Configuration
INPUT_DIR = os.path.join(os.getcwd(), "temp_inputs")
OUTPUT_DIR = os.path.join(os.getcwd(), "output_timetables")
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Cache variables with file hashes to detect changes
_cached_data_frames = None
_cached_timestamp = 0
_file_hashes = {}
_EXAM_SCHEDULE_FILES = set()

# Allowed file extensions
ALLOWED_EXTENSIONS = {'csv'}

_CLASSROOM_USAGE_TRACKER = {}
_TIMETABLE_CLASSROOM_ALLOCATIONS = {}

def initialize_classroom_usage_tracker():
    """Initialize the global classroom usage tracker"""
    global _CLASSROOM_USAGE_TRACKER
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    time_slots = ['09:00-10:30', '10:30-12:00', '12:00-13:00', '13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00']
    
    _CLASSROOM_USAGE_TRACKER = {}
    for day in days:
        _CLASSROOM_USAGE_TRACKER[day] = {}
        for time_slot in time_slots:
            _CLASSROOM_USAGE_TRACKER[day][time_slot] = set()

    print(f"   [SCHOOL] Initialized classroom tracker: {len(days)} days × {len(time_slots)} time slots")

def reset_classroom_usage_tracker():
    """Reset the classroom usage tracker (call before generating new timetables)"""
    global _CLASSROOM_USAGE_TRACKER, _TIMETABLE_CLASSROOM_ALLOCATIONS
    _CLASSROOM_USAGE_TRACKER = {}
    _TIMETABLE_CLASSROOM_ALLOCATIONS = {}
    initialize_classroom_usage_tracker()
    print("[RESET] Classroom usage tracker reset for new timetable generation")

def get_exam_schedule_files():
    """Get list of exam schedule files that should be displayed"""
    global _EXAM_SCHEDULE_FILES
    return list(_EXAM_SCHEDULE_FILES)

def add_exam_schedule_file(filename):
    """Add a filename to the list of exam schedules to display"""
    global _EXAM_SCHEDULE_FILES
    _EXAM_SCHEDULE_FILES.add(filename)

def clear_exam_schedule_files():
    """Clear the list of exam schedules to display"""
    global _EXAM_SCHEDULE_FILES
    _EXAM_SCHEDULE_FILES.clear()

def remove_exam_schedule_file(filename):
    """Remove a filename from the list of exam schedules to display"""
    global _EXAM_SCHEDULE_FILES
    if filename in _EXAM_SCHEDULE_FILES:
        _EXAM_SCHEDULE_FILES.remove(filename)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_file_hash(filepath):
    """Calculate MD5 hash of a file to detect changes"""
    try:
        with open(filepath, 'rb') as f:
            return hashlib.md5(f.read()).hexdigest()
    except:
        return None

def has_files_changed():
    """Check if any input files have changed since last load"""
    global _file_hashes
    
    if not os.path.exists(INPUT_DIR):
        return True
        
    current_files = os.listdir(INPUT_DIR)
    current_hashes = {}
    
    for file in current_files:
        filepath = os.path.join(INPUT_DIR, file)
        current_hashes[file] = get_file_hash(filepath)
    
    # If number of files changed
    if set(current_files) != set(_file_hashes.keys()):
        print("[DIR] File count changed")
        _file_hashes = current_hashes
        return True
    
    # If file contents changed
    for file, current_hash in current_hashes.items():
        if file not in _file_hashes or _file_hashes[file] != current_hash:
            print(f"[FILE] File content changed: {file}")
            _file_hashes = current_hashes
            return True
    
    return False

def find_csv_file(filename):
    """Find CSV file with flexible search (case-insensitive, space-insensitive)."""
    if not os.path.exists(INPUT_DIR):
        return None
        
    files = os.listdir(INPUT_DIR)
    filename_clean = filename.lower().replace(' ', '')
    
    for file in files:
        file_clean = file.lower().replace(' ', '')
        if file_clean == filename_clean:
            return os.path.join(INPUT_DIR, file)
    
    # Try more flexible matching for common patterns
    for file in files:
        file_clean = file.lower().replace(' ', '').replace('_', '').replace('-', '')
        filename_clean_alt = filename_clean.replace('_', '').replace('-', '')
        if file_clean == filename_clean_alt:
            return os.path.join(INPUT_DIR, file)
    
    return None

def load_all_data(force_reload=False):
    """Load CSV files and return dataframes with force reload option"""
    global _cached_data_frames
    global _cached_timestamp
    
    # Always check if files have changed
    files_changed = has_files_changed()
    
    # Use caching but allow force reload or if files changed
    if (not force_reload and 
        not files_changed and
        '_cached_data_frames' in globals() and 
        _cached_data_frames is not None):

        current_time = time.time()
        # Cache for 30 seconds max
        if current_time - _cached_timestamp < 30:
            print("[FOLDER] Using cached data frames (files unchanged)")
            return _cached_data_frames
        else:
            print("[FOLDER] Cache expired, reloading data")
    
    if files_changed:
        print("[FOLDER] Files changed, reloading data")
    
    required_files = [
        "course_data.csv",
        "faculty_availability.csv",
        "classroom_data.csv",
        "student_data.csv",
        "exams_data.csv"
    ]
    dfs = {}
    
    print("[FOLDER] Loading CSV files...")
    print(f"[DIR] Input directory contents: {os.listdir(INPUT_DIR) if os.path.exists(INPUT_DIR) else 'Directory not found'}")
    
    # Update file hashes
    if os.path.exists(INPUT_DIR):
        for file in os.listdir(INPUT_DIR):
            filepath = os.path.join(INPUT_DIR, file)
            _file_hashes[file] = get_file_hash(filepath)
    
    for f in required_files:
        file_path = find_csv_file(f)
        if not file_path:
            print(f"[FAIL] CSV not found: {f}")
            # Try to find any similar file
            files = os.listdir(INPUT_DIR) if os.path.exists(INPUT_DIR) else []
            similar_files = [file for file in files if f.split('_')[0] in file.lower()]
            if similar_files:
                print(f"   💡 Similar files found: {similar_files}")
            return None
        
        try:
            key = f.replace("_data.csv", "").replace(".csv", "")
            dfs[key] = pd.read_csv(file_path)
            print(f"[OK] Loaded {f} from {file_path} ({len(dfs[key])} rows)")
            
            # Show sample data for verification
            if not dfs[key].empty:
                print(f"   Columns: {list(dfs[key].columns)}")
                if 'course' in key:
                    print(f"   First 3 courses:")
                    for i, row in dfs[key].head(3).iterrows():
                        print(f"     {i+1}. {row['Course Code'] if 'Course Code' in row else 'N/A'} - Semester: {row.get('Semester', 'N/A')} - Branch: {row.get('Branch', 'N/A')} - Elective: {row.get('Elective (Yes/No)', 'N/A')}")
                
        except Exception as e:
            print(f"[FAIL] Error loading {f}: {e}")
            return None
    
    # Validate required columns in course data
    if 'course' in dfs:
        required_course_columns = ['Course Code', 'Semester', 'LTPSC']
        missing_columns = [col for col in required_course_columns if col not in dfs['course'].columns]
        if missing_columns:
            print(f"[FAIL] Missing columns in course_data: {missing_columns}")
            print(f"   Available columns: {list(dfs['course'].columns)}")
            return None
    
    # Cache the results
    _cached_data_frames = dfs
    _cached_timestamp = time.time()
    
    print("[OK] All CSV files loaded successfully!")
    return dfs

def get_course_info(dfs):
    """Extract course information from course data for frontend display with proper department mapping"""
    course_info = {}
    if 'course' in dfs:
        for _, course in dfs['course'].iterrows():
            course_code = course['Course Code']
            
            # FIXED: Map department based on course code prefix
            department = map_department_from_course_code(course_code)
            
            is_elective = course.get('Elective (Yes/No)', 'No').upper() == 'YES'
            course_type = 'Elective' if is_elective else 'Core'
            
            # FIX: Use 'Faculty' column instead of 'Instructor'
            instructor = course.get('Faculty', 'Unknown')
            
            course_info[course_code] = {
                'name': course.get('Course Name', 'Unknown Course'),
                'credits': course.get('Credits', 0),
                'type': course_type,
                'instructor': instructor,  # This will now use the correct Faculty column
                'department': department,  # Use mapped department
                'is_elective': is_elective,
                'branch': department,  # Use department as branch for compatibility
                'is_common_elective': is_elective,
                'ltpsc': course.get('LTPSC', ''),
                'ltpsc_components': parse_ltpsc(course.get('LTPSC', '')) if 'LTPSC' in course else None
            }
            
            # Debug logging
            print(f"   [NOTE] Course {course_code}: Department = {department}")
            
    return course_info

def map_department_from_course_code(course_code):
    """Map department based on course code prefix"""
    if not isinstance(course_code, str):
        return 'Department of Arts, Science and Design'
    
    course_code_upper = course_code.upper().strip()
    
    if course_code_upper.startswith('CS'):
        return 'Computer Science and Engineering'
    elif course_code_upper.startswith('EC'):
        return 'Electronics and Communication Engineering'
    elif course_code_upper.startswith('DS'):
        return 'Data Science and Artificial Intelligence'
    elif course_code_upper.startswith('DA'):  # Also handle DA prefix for DSAI
        return 'Data Science and Artificial Intelligence'
    else:
        return 'Department of Arts, Science and Design'

def get_departments_from_data(dfs):
    """Extract unique departments from course data"""
    if 'course' in dfs and 'Department' in dfs['course'].columns:
        departments = dfs['course']['Department'].unique()
        # Filter out None/NaN and return as list
        return [dept for dept in departments if pd.notna(dept) and dept != '']
    else:
        return ['CSE', 'DSAI', 'ECE']  # fallback

def separate_courses_by_type(dfs, semester_id, branch=None):
    """Separate courses into core and elective baskets for a given semester and branch"""
    if 'course' not in dfs:
        return {'core_courses': [], 'elective_courses': []}
    
    try:
        # Normalize branch name to match Department values in data
        def normalize_branch_name(branch_raw):
            if not branch_raw:
                return None
            br = str(branch_raw).strip().upper()
            if br in ['CSE', 'CS', 'COMPUTER SCIENCE', 'COMPUTER SCIENCE AND ENGINEERING']:
                return 'Computer Science and Engineering'
            if br in ['ECE', 'EC', 'ELECTRONICS', 'ELECTRONICS AND COMMUNICATION ENGINEERING']:
                return 'Electronics and Communication Engineering'
            if br in ['DSAI', 'DS', 'DA', 'DATA SCIENCE', 'DATA SCIENCE AND ARTIFICIAL INTELLIGENCE']:
                return 'Data Science and Artificial Intelligence'
            # Fallback: return original for exact matches in data
            return branch_raw

        # Filter courses for the semester
        sem_courses = dfs['course'][
            dfs['course']['Semester'].astype(str).str.strip() == str(semester_id)
        ].copy()
        
        if sem_courses.empty:
            return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}
        
        # ENHANCED: Filter by department if specified - only include courses for the specific department
        if branch and 'Department' in sem_courses.columns:
            normalized_branch = normalize_branch_name(branch)

            # Build robust department mask that tolerates abbreviations and missing/incorrect Department field
            dept_series = sem_courses['Department'].astype(str).fillna('').str.strip()
            branch_upper = str(branch).strip().upper()

            # Map by course code prefix as a reliable fallback
            inferred_departments = sem_courses['Course Code'].apply(map_department_from_course_code)

            # Accept rows where:
            # - Department equals normalized full name, OR
            # - Department equals branch abbreviation, OR
            # - Inferred department from code equals normalized full name
            dept_match = (
                (dept_series == normalized_branch) |
                (dept_series.str.upper() == branch_upper) |
                (inferred_departments == normalized_branch)
            )

            # Include department-specific cores and all electives
            is_elective = sem_courses['Elective (Yes/No)'].astype(str).str.upper() == 'YES'
            sem_courses = sem_courses[(dept_match & ~is_elective) | is_elective].copy()
        
        if sem_courses.empty:
            return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}
        
        # Separate core and elective courses
        core_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() != 'YES'
        ].copy()
        
        elective_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() == 'YES'
        ].copy()
        
        print(f"   [STATS] Course separation for Semester {semester_id}, Department {branch or 'All'}:")
        print(f"      Core courses: {len(core_courses)}")
        print(f"      Elective courses: {len(elective_courses)}")
        
        # ENHANCED: Debug info about department-specific courses
        if branch:
            dept_core_courses = core_courses[core_courses['Department'] == branch]
            print(f"      Department-specific core courses for {branch}: {len(dept_core_courses)}")
            if not dept_core_courses.empty:
                print(f"      Department core courses: {dept_core_courses['Course Code'].tolist()}")
        
        return {
            'core_courses': core_courses,
            'elective_courses': elective_courses
        }
        
    except Exception as e:
        print(f"[WARN] Error separating courses by type: {e}")
        traceback.print_exc()
        return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}

def separate_courses_by_mid_semester(dfs, semester_id, branch=None):
    """Separate courses into pre-mid and post-mid based on Half Semester and Post mid-sem columns
    
    Logic:
    - If Half Semester = Yes: Schedule in BOTH pre-mid and post-mid
    - If Half Semester = No:
        - If Post mid-sem = No: Schedule in PRE-MID ONLY
        - If Post mid-sem = Yes: Schedule in POST-MID ONLY
    """
    if 'course' not in dfs:
        return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}
    
    try:
        # Filter courses for the semester
        sem_courses = dfs['course'][
            dfs['course']['Semester'].astype(str).str.strip() == str(semester_id)
        ].copy()
        
        if sem_courses.empty:
            return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}
        
        # Filter by department if specified
        if branch and 'Department' in sem_courses.columns:
            normalized_branch = branch.strip()
            # Include department-specific courses and common courses
            dept_match = sem_courses['Department'].astype(str).str.strip() == normalized_branch

            # Robust detection of a 'common' column (handles misspellings like 'Comman' or variations)
            common_col = None
            for col in sem_courses.columns:
                col_low = str(col).lower()
                if 'common' in col_low or 'comman' in col_low:
                    common_col = col
                    break

            if common_col is not None:
                common_courses = sem_courses[common_col].astype(str).str.upper().str.strip() == 'YES'
            else:
                # If there's no common column, treat none as common
                common_courses = pd.Series([False] * len(sem_courses), index=sem_courses.index)

            sem_courses = sem_courses[dept_match | common_courses].copy()
        
        if sem_courses.empty:
            return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}
        
        # Check if required columns exist (try flexible matching)
        post_mid_col = None
        for col in sem_courses.columns:
            if 'post' in str(col).lower() and 'mid' in str(col).lower():
                post_mid_col = col
                break
        
        half_sem_col = None
        for col in sem_courses.columns:
            if 'half' in str(col).lower() and 'semester' in str(col).lower():
                half_sem_col = col
                break
        
        if post_mid_col is None:
            print(f"[WARN] 'Post mid-sem' column not found for semester {semester_id}")
            print(f"   Available columns: {list(sem_courses.columns)}")
            return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}
        
        if half_sem_col is None:
            print(f"[WARN] 'Half Semester' column not found for semester {semester_id}")
            print(f"   Available columns: {list(sem_courses.columns)}")
            return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}
        
        # Debug: Print sample of column values
        print(f"   [DEBUG] Using columns:")
        print(f"      - Half Semester: '{half_sem_col}'")
        print(f"      - Post mid-sem: '{post_mid_col}'")
        sample_half = sem_courses[half_sem_col].astype(str).str.strip().str.upper().unique()[:10]
        sample_post = sem_courses[post_mid_col].astype(str).str.strip().str.upper().unique()[:10]
        print(f"      Half Semester values: {list(sample_half)}")
        print(f"      Post mid-sem values: {list(sample_post)}")
        print(f"      Total courses: {len(sem_courses)}")
        
        # Normalize column values
        sem_courses[half_sem_col] = sem_courses[half_sem_col].astype(str).str.strip().str.upper()
        sem_courses[post_mid_col] = sem_courses[post_mid_col].astype(str).str.strip().str.upper()
        
        # ======================================================
        # CORRECTED LOGIC (Based on Half Semester):
        # ======================================================
        
        # RULE 1: Courses with Half Semester = NO
        # These go in BOTH pre-mid and post-mid
        half_sem_no_courses = sem_courses[
            sem_courses[half_sem_col] == 'NO'
        ].copy()
        
        # RULE 2: Courses with Half Semester = YES
        # Rule 2a: Post mid-sem = NO → PRE-MID ONLY
        half_sem_yes_pre_mid = sem_courses[
            (sem_courses[half_sem_col] == 'YES') &
            (sem_courses[post_mid_col] == 'NO')
        ].copy()
        
        # Rule 2b: Post mid-sem = YES → POST-MID ONLY
        half_sem_yes_post_mid = sem_courses[
            (sem_courses[half_sem_col] == 'YES') &
            (sem_courses[post_mid_col] == 'YES')
        ].copy()
        
        # FINAL SEPARATION:
        # PRE-MID: Half Sem = NO courses + Half Sem = YES with Post mid = NO courses
        pre_mid_courses = pd.concat([half_sem_no_courses, half_sem_yes_pre_mid], ignore_index=True)
        pre_mid_courses = pre_mid_courses.drop_duplicates(subset=['Course Code'])
        
        # POST-MID: Half Sem = NO courses + Half Sem = YES with Post mid = YES courses
        post_mid_courses = pd.concat([half_sem_no_courses, half_sem_yes_post_mid], ignore_index=True)
        post_mid_courses = post_mid_courses.drop_duplicates(subset=['Course Code'])
        
        print(f"   [STATS] Course separation for Semester {semester_id}, Department {branch or 'All'}:")
        print(f"      Total courses: {len(sem_courses)}")
        print(f"      Half Semester = NO (both pre & post): {len(half_sem_no_courses)}")
        print(f"      Half Semester = YES, Post mid = NO (pre-mid only): {len(half_sem_yes_pre_mid)}")
        print(f"      Half Semester = YES, Post mid = YES (post-mid only): {len(half_sem_yes_post_mid)}")
        print(f"      → Pre-mid courses total: {len(pre_mid_courses)}")
        print(f"      → Post-mid courses total: {len(post_mid_courses)}")
        
        return {
            'pre_mid_courses': pre_mid_courses,
            'post_mid_courses': post_mid_courses
        }
        
    except Exception as e:
        print(f"[WARN] Error separating courses by mid-semester: {e}")
        traceback.print_exc()
        return {'pre_mid_courses': pd.DataFrame(), 'post_mid_courses': pd.DataFrame()}

def parse_ltpsc(ltpsc_string):
    """Parse L-T-P-S-C string and return components"""
    # Default values: 2 lectures, 1 tutorial, 0 practicals
    default_ltpsc = {'L': 2, 'T': 1, 'P': 0, 'S': 0, 'C': 2}
    
    # Handle empty or None values
    if not ltpsc_string or pd.isna(ltpsc_string) or str(ltpsc_string).strip() == '':
        return default_ltpsc
    
    try:
        parts = str(ltpsc_string).strip().split('-')
        if len(parts) == 5:
            return {
                'L': int(parts[0]) if parts[0].strip() else 2,  # Lectures per week
                'T': int(parts[1]) if parts[1].strip() else 1,  # Tutorials per week
                'P': int(parts[2]) if parts[2].strip() else 0,  # Practicals per week
                'S': int(parts[3]) if parts[3].strip() else 0,  # S credits
                'C': int(parts[4]) if parts[4].strip() else 2   # Total credits
            }
        else:
            return default_ltpsc
    except:
        return default_ltpsc

def enforce_elective_day_separation(basket_allocations):
    """Enforce that elective lectures and tutorials are on different days"""
    print("[DEBUG] ENFORCING ELECTIVE DAY SEPARATION...")
    
    for basket_name, allocation in basket_allocations.items():
        lecture_days = set(day for day, time in allocation['lectures'])
        tutorial_day = allocation['tutorial'][0]
        
        if tutorial_day in lecture_days:
            print(f"   [FAIL] VIOLATION: Basket '{basket_name}' has tutorial on same day as lectures")
            print(f"      Lecture days: {lecture_days}, Tutorial day: {tutorial_day}")
            
            # Find an alternative tutorial day
            all_days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
            available_days = [day for day in all_days if day not in lecture_days]
            
            if available_days:
                new_tutorial_day = available_days[0]
                # Find a tutorial time slot for the new day
                tutorial_times = ['14:30-15:30']  # Assuming 1-hour tutorial slots
                new_tutorial_slot = (new_tutorial_day, tutorial_times[0])
                
                allocation['tutorial'] = new_tutorial_slot
                allocation['tutorial_day'] = new_tutorial_day
                allocation['days_separated'] = True
                
                print(f"   [OK] FIXED: Moved tutorial to {new_tutorial_day}")
            else:
                print(f"   [WARN]  CANNOT FIX: No available days for tutorial")
        else:
            print(f"   [OK] VALID: Basket '{basket_name}' has proper day separation")
    
    return basket_allocations

def schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, lecture_times, tutorial_times, lab_times=None, branch=None):
    """Schedule core courses strictly adhering to LTPSC structure"""
    if core_courses.empty:
        return used_slots
    
    course_day_usage = {}
    
    # Lab times are handled as consecutive slot pairs (2-hour labs use 2 consecutive 1.5-hour slots)
    # No need for separate lab_times parameter - handled internally
    
    # ENHANCED: Filter core courses to only include department-specific ones
    if branch and 'Department' in core_courses.columns:
        dept_core_courses = core_courses[core_courses['Department'] == branch].copy()
        print(f"   [COURSES] Scheduling {len(dept_core_courses)} department-specific core courses for {branch}...")
        if not dept_core_courses.empty:
            print(f"      Courses to schedule: {dept_core_courses['Course Code'].tolist()}")
    else:
        dept_core_courses = core_courses.copy()
        print(f"   [COURSES] Scheduling {len(dept_core_courses)} core courses...")
    
    if dept_core_courses.empty:
        print(f"   [INFO] No department-specific core courses found for {branch}")
        return used_slots
    
    # Parse LTPSC for core courses - STRICTLY ADHERE TO LTPSC STRUCTURE
    for _, course in dept_core_courses.iterrows():
        course_code = course['Course Code']
        ltpsc_str = course.get('LTPSC', '')
        
        # Parse LTPSC - defaults to 2 lectures, 1 tutorial if empty
        ltpsc = parse_ltpsc(ltpsc_str)
        L = ltpsc['L']
        T = ltpsc['T']
        P = ltpsc['P']
        
        # FIX: NEVER schedule labs for MA (Mathematics) courses
        is_math_course = course_code.startswith('MA') or 'Mathematics' in str(course.get('Department', ''))
        if is_math_course:
            P = 0  # Force no labs for math courses
            print(f"      [DISABLE] MA Course detected: {course_code} - Labs disabled regardless of LTPSC")
        
        # Determine lectures needed: if L is 2 or 3, schedule 2 lectures
        if L == 2 or L == 3:
            lectures_needed = 2
        elif L == 1:
            lectures_needed = 1
        else:
            # For other values, schedule min(L, 2) lectures
            lectures_needed = min(L, 2)
        
        # Determine tutorials needed: if T is 1, schedule 1 tutorial
        tutorials_needed = 1 if T >= 1 else 0
        
        # Determine labs needed: if P is 1 or more, schedule 1 lab (2 hours)
        # BUT NEVER for MA courses (already handled above)
        # If P is 0, do NOT schedule any labs regardless of other factors
        labs_needed = 1 if P >= 1 else 0
        
        # Track which days we've used for this course
        course_day_usage[course_code] = {'lectures': set(), 'tutorials': set(), 'labs': set()}
        
        print(f"      Scheduling {course_code} (LTPSC: {ltpsc_str} -> L={L}, T={T}, P={P}):")
        print(f"         → {lectures_needed} lectures, {tutorials_needed} tutorial, {labs_needed} lab")
        
        # Schedule lectures (1.5 hours each)
        lectures_scheduled = 0
        max_lecture_attempts = 200
        
        while lectures_scheduled < lectures_needed and max_lecture_attempts > 0:
            max_lecture_attempts -= 1
            
            # Find days where this course doesn't have a lecture yet
            available_days = [d for d in days if d not in course_day_usage[course_code]['lectures']]
            if not available_days:
                print(f"      [WARN] Cannot schedule more lectures for {course_code} - no available days left")
                break
            
            day = random.choice(available_days)
            time_slot = random.choice(lecture_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = course_code
                used_slots.add(key)
                course_day_usage[course_code]['lectures'].add(day)
                lectures_scheduled += 1
                print(f"      [OK] Scheduled lecture {lectures_scheduled} for {course_code} on {day} at {time_slot}")
        
        # Schedule tutorial (1 hour) if needed
        tutorials_scheduled = 0
        if tutorials_needed > 0:
            max_tutorial_attempts = 100
        
        while tutorials_scheduled < tutorials_needed and max_tutorial_attempts > 0:
            max_tutorial_attempts -= 1
            
            # Find days where this course doesn't have a tutorial AND doesn't have a lecture
            available_days = [d for d in days if d not in course_day_usage[course_code]['tutorials'] and d not in course_day_usage[course_code]['lectures']]
            if not available_days:
                # If no completely free days, try days without tutorial but with lecture
                available_days = [d for d in days if d not in course_day_usage[course_code]['tutorials']]
                if not available_days:
                    print(f"      [WARN] Cannot schedule tutorial for {course_code} - no available days left")
                    break
            
            day = random.choice(available_days)
            time_slot = random.choice(tutorial_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                used_slots.add(key)
                course_day_usage[course_code]['tutorials'].add(day)
                tutorials_scheduled += 1
                print(f"      [OK] Scheduled tutorial for {course_code} on {day} at {time_slot}")
        
        # Schedule lab (STRICTLY 2 hours) if needed - use consecutive 1.5-hour slots
        # NOTE: If P=0 in LTPSC, NO labs will be scheduled for this course
        # Labs are ALWAYS scheduled for exactly 2 hours using two consecutive slots
        # BUT NEVER for MA courses
        labs_scheduled = 0
        if labs_needed > 0 and P > 0 and not is_math_course:
            max_lab_attempts = 100
            
            # Define 2-hour lab slot pairs (consecutive 1.5-hour slots that form a 2-hour lab)
            # Labs MUST be scheduled for exactly 2 hours - no exceptions
            lab_slot_pairs = [
                (['13:00-14:30', '14:30-15:30'], '13:00-15:30'),  # 2 hours: 13:00-15:30
                (['15:30-17:00', '17:00-18:00'], '15:30-18:00'),  # 2 hours: 15:30-18:00
            ]
            
            while labs_scheduled < labs_needed and max_lab_attempts > 0:
                max_lab_attempts -= 1
                
                # Find days where this course doesn't have a lab
                available_days = [d for d in days if d not in course_day_usage[course_code]['labs']]
                if not available_days:
                    print(f"      [WARN] Cannot schedule lab for {course_code} - no available days left")
                    break
                
                day = random.choice(available_days)
                # Try each lab slot pair
                lab_scheduled = False
                for slot_pair, lab_display_time in lab_slot_pairs:
                    # Check if both consecutive slots are free
                    slot1, slot2 = slot_pair
                    key1 = (day, slot1)
                    key2 = (day, slot2)
                    
                    if (key1 not in used_slots and key2 not in used_slots and
                        schedule.loc[slot1, day] == 'Free' and schedule.loc[slot2, day] == 'Free'):
                        # Mark both slots as lab
                        schedule.loc[slot1, day] = f"{course_code} (Lab)"
                        schedule.loc[slot2, day] = f"{course_code} (Lab)"
                        used_slots.add(key1)
                        used_slots.add(key2)
                        course_day_usage[course_code]['labs'].add(day)
                        labs_scheduled += 1
                        lab_scheduled = True
                        print(f"      [OK] Scheduled lab for {course_code} on {day} at {lab_display_time} (using slots {slot1} and {slot2})")
                        break
                
                if not lab_scheduled:
                    # Try next attempt
                    continue
        elif P == 0:
            print(f"      [SKIP] No labs scheduled for {course_code} (P=0 in LTPSC)")
        elif is_math_course:
            print(f"      [DISABLE] Skipping lab for MA course {course_code} (P={P} in LTPSC but labs disabled for Mathematics)")
        
        # Summary
        if lectures_scheduled < lectures_needed:
            print(f"      [FAIL] Could only schedule {lectures_scheduled}/{lectures_needed} lectures for {course_code}")
        if tutorials_needed > 0 and tutorials_scheduled < tutorials_needed:
            print(f"      [FAIL] Could only schedule {tutorials_scheduled}/{tutorials_needed} tutorials for {course_code}")
        if labs_needed > 0 and labs_scheduled < labs_needed and not is_math_course:
            print(f"      [FAIL] Could only schedule {labs_scheduled}/{labs_needed} labs for {course_code}")
        if lectures_scheduled == lectures_needed and tutorials_scheduled == tutorials_needed and labs_scheduled == labs_needed:
            print(f"      [OK] Successfully scheduled {course_code} according to LTPSC structure")

    return used_slots

def get_common_elective_slots():
    """Define common elective slots that will be used for both sections"""
    # Return slots for 2 lectures + 1 tutorial per elective
    return [
        # Monday slots
        ('Mon', '09:00-10:30'), ('Mon', '10:30-12:00'), ('Mon', '13:00-14:30'), 
        ('Mon', '14:30-15:30'), ('Mon', '15:30-17:00'), ('Mon', '17:00-18:00'),
        # Tuesday slots
        ('Tue', '09:00-10:30'), ('Tue', '10:30-12:00'), ('Tue', '13:00-14:30'),
        ('Tue', '14:30-15:30'), ('Tue', '15:30-17:00'), ('Tue', '17:00-18:00'),
        # Wednesday slots
        ('Wed', '09:00-10:30'), ('Wed', '10:30-12:00'), ('Wed', '13:00-14:30'),
        ('Wed', '14:30-15:30'), ('Wed', '15:30-17:00'), ('Wed', '17:00-18:00'),
        # Thursday slots
        ('Thu', '09:00-10:30'), ('Thu', '10:30-12:00'), ('Thu', '13:00-14:30'),
        ('Thu', '14:30-15:30'), ('Thu', '15:30-17:00'), ('Thu', '17:00-18:00'),
        # Friday slots
        ('Fri', '09:00-10:30'), ('Fri', '10:30-12:00'), ('Fri', '13:00-14:30'),
        ('Fri', '14:30-15:30'), ('Fri', '15:30-17:00'), ('Fri', '17:00-18:00')
    ]

def get_basket_time_slots():
    """Define time slots for elective baskets"""
    return [
        ('Mon', '09:00-10:30'), ('Mon', '13:00-14:30'),
        ('Tue', '09:00-10:30'), ('Tue', '13:00-14:30'),
        ('Wed', '09:00-10:30'), ('Wed', '13:00-14:30'),
        ('Thu', '09:00-10:30'), ('Thu', '13:00-14:30'),
        ('Fri', '09:00-10:30'), ('Fri', '13:00-14:30'),
        ('Mon', '14:30-15:30'), ('Tue', '14:30-15:30'),
        ('Wed', '14:30-15:30'), ('Thu', '14:30-15:30'),
        ('Fri', '14:30-15:30')
    ]

def schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section, branch=None):
    """Schedule elective courses in IDENTICAL COMMON slots for ALL branches and sections"""
    elective_scheduled = 0
    branch_info = f" for {branch}" if branch else ""
    
    print(f"   📅 Applying IDENTICAL COMMON slots for Section {section}{branch_info}:")
    
    # Track which basket slots we've already scheduled
    scheduled_basket_slots = set()
    
    for course_code, allocation in elective_allocations.items():
        if allocation is None:
            continue
            
        basket_name = allocation['basket_name']
        lectures = allocation['lectures']
        tutorial = allocation['tutorial']
        all_courses = allocation['all_courses_in_basket']
        days_separated = allocation['days_separated']
        
        print(f"      [BASKET] Basket '{basket_name}' - IDENTICAL across all branches:")
        print(f"         Courses: {', '.join(all_courses)}")
        print(f"         Days Separation: {'[OK]' if days_separated else '[FAIL]'}")
        
        # Use basket name directly
        basket_display = basket_name
        tutorial_display = f"{basket_name} (Tutorial)"
        
        # Schedule lectures - IDENTICAL for all branches and sections
        for day, time_slot in lectures:
            slot_key = (basket_name, day, time_slot, 'lecture')
            
            if slot_key in scheduled_basket_slots:
                continue
                
            key = (day, time_slot)
            
            if schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = basket_display
                used_slots.add(key)
                scheduled_basket_slots.add(slot_key)
                elective_scheduled += 1
                print(f"         [OK] COMMON LECTURE: {day} {time_slot}")
                print(f"                SAME for ALL branches & sections")
            else:
                print(f"         [FAIL] LECTURE CONFLICT: {day} {time_slot} - {schedule.loc[time_slot, day]}")
        
        # Schedule tutorial - IDENTICAL for all branches and sections (only if T >= 1)
        if tutorial:
            day, time_slot = tutorial
            slot_key = (basket_name, day, time_slot, 'tutorial')
            
            if slot_key not in scheduled_basket_slots:
                key = (day, time_slot)
                
                if schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = tutorial_display
                    used_slots.add(key)
                    scheduled_basket_slots.add(slot_key)
                    elective_scheduled += 1
                    print(f"         [OK] COMMON TUTORIAL: {day} {time_slot}")
                    print(f"                SAME for ALL branches & sections")
                else:
                    print(f"         [FAIL] TUTORIAL CONFLICT: {day} {time_slot} - {schedule.loc[time_slot, day]}")
        else:
            print(f"         [INFO] No tutorial scheduled (T=0 in LTPSC)")
    
    print(f"   [OK] Scheduled {elective_scheduled} IDENTICAL COMMON elective sessions")
    return used_slots

def generate_section_schedule_with_elective_baskets(dfs, semester_id, section, elective_allocations, branch=None, time_config=None, basket_allocations=None):
    """Generate schedule with basket-based elective allocation - COMMON slots across branches.
    Allows overriding time slots via time_config: {
        'morning_slots': [..], 'lunch_slots': [..], 'afternoon_slots': [..],
        'lecture_times': [..], 'tutorial_times': [..]
    }
    basket_allocations: Optional dict of basket allocations to ensure all required baskets are scheduled."""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"   [TARGET] Generating BASKET-BASED schedule for Semester {semester_id}, Section {section}{branch_info}")
    print(f"   [PIN] Using COMMON elective basket slots (same for all branches)")
    
    if 'course' not in dfs:
        print("[FAIL] Course data not available")
        return None
    
    try:
        # Get only the core courses for this specific branch
        course_baskets = separate_courses_by_type(dfs, semester_id, branch)
        core_courses = course_baskets['core_courses']
        
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        
        # Time slot structure
        if time_config:
            morning_slots = time_config.get('morning_slots', ['09:00-10:30', '10:30-12:00'])
            lunch_slots = time_config.get('lunch_slots', ['12:00-13:00'])
            afternoon_slots = time_config.get('afternoon_slots', ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00'])
        else:
            morning_slots = ['09:00-10:30', '10:30-12:00']
            lunch_slots = ['12:00-13:00']
            afternoon_slots = ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00']
        all_slots = morning_slots + lunch_slots + afternoon_slots
        
        # Lecture slots (1.5 hours)
        if time_config and time_config.get('lecture_times'):
            lecture_times = time_config['lecture_times']
        else:
            lecture_times = ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00']
        
        # Tutorial slots (1 hour)
        if time_config and time_config.get('tutorial_times'):
            tutorial_times = time_config['tutorial_times']
        else:
            tutorial_times = ['14:30-15:30', '17:00-18:00']
        
        # Lab slots (2 hours) - represented as pairs of consecutive 1.5-hour slots
        # Labs will use: ['13:00-14:30', '14:30-15:30'] or ['15:30-17:00', '17:00-18:00']
        lab_times = None  # Will be handled as slot pairs in the scheduling function
        
        # Create schedule template
        schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
        # Mark lunch break label across provided lunch slots
        for lunch_slot in lunch_slots:
            if lunch_slot in schedule.index:
                schedule.loc[lunch_slot] = 'LUNCH BREAK'

        used_slots = set()

        # Schedule elective courses FIRST using COMMON basket allocation
        if elective_allocations:
            print(f"   📅 Applying COMMON basket elective slots for Section {section}:")
            # Show what we're scheduling
            unique_baskets = set()
            for allocation in elective_allocations.values():
                if allocation:
                    unique_baskets.add(allocation['basket_name'])
            
            print(f"   [BASKET] Baskets to schedule: {list(unique_baskets)}")
            
            used_slots = schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section, branch)
        
        # ENSURE all required baskets are scheduled from basket_allocations
        # This ensures baskets are scheduled even if courses aren't in elective_allocations
        if basket_allocations:
            scheduled_basket_names = set()
            if elective_allocations:
                for allocation in elective_allocations.values():
                    if allocation and allocation.get('basket_name'):
                        scheduled_basket_names.add(allocation['basket_name'])
            
            # Schedule any baskets from basket_allocations that weren't scheduled via elective_allocations
            for basket_name, basket_allocation in basket_allocations.items():
                if basket_name not in scheduled_basket_names:
                    print(f"   🔧 Scheduling basket '{basket_name}' directly from basket_allocations...")
                    # Create a temporary allocation structure to schedule this basket
                    temp_allocation = {
                        'basket_name': basket_name,
                        'lectures': basket_allocation['lectures'],
                        'tutorial': basket_allocation.get('tutorial'),
                        'all_courses_in_basket': basket_allocation.get('courses', []),
                        'days_separated': basket_allocation.get('days_separated', True)
                    }
                    # Schedule using a dummy course code key
                    dummy_key = f"__BASKET_{basket_name}__"
                    temp_elective_allocations = {dummy_key: temp_allocation}
                    used_slots = schedule_electives_by_baskets(temp_elective_allocations, schedule, used_slots, section, branch)
        
        # Schedule core courses AFTER electives - these are branch-specific
        # IMPORTANT: Filter out any elective courses that are already scheduled in baskets
        if not core_courses.empty:
            # Get list of elective course codes from baskets to exclude them
            elective_course_codes = set()
            if elective_allocations:
                for allocation in elective_allocations.values():
                    if allocation and 'all_courses_in_basket' in allocation:
                        elective_course_codes.update(allocation['all_courses_in_basket'])
            
            # Filter out elective courses from core_courses
            if elective_course_codes:
                core_courses_filtered = core_courses[~core_courses['Course Code'].isin(elective_course_codes)].copy()
                excluded_count = len(core_courses) - len(core_courses_filtered)
                if excluded_count > 0:
                    print(f"   [DISABLE] Excluded {excluded_count} elective course(s) already scheduled in baskets: {', '.join(elective_course_codes)}")
                core_courses = core_courses_filtered
            
            if not core_courses.empty:
                print(f"   [COURSES] Scheduling {len(core_courses)} BRANCH-SPECIFIC core courses for {branch}...")
                used_slots = schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, 
                                                                    lecture_times, tutorial_times, None, branch)
            else:
                print(f"   [INFO] No core courses to schedule after filtering electives")
        
        return schedule
        
    except Exception as e:
        print(f"[FAIL] Error generating basket-based schedule: {e}")
        traceback.print_exc()
        return None

def generate_mid_semester_schedule(dfs, semester_id, section, courses_df, branch=None, time_config=None, schedule_type='pre_mid', elective_allocations=None):
    """Generate pre-mid or post-mid schedule with elective basket support"""
    schedule_type_name = "PRE-MID" if schedule_type == 'pre_mid' else "POST-MID"
    branch_info = f", Branch {branch}" if branch else ""
    print(f"   [TARGET] Generating {schedule_type_name} schedule for Semester {semester_id}, Section {section}{branch_info}")
    
    if courses_df.empty:
        print(f"   [INFO] No courses to schedule for {schedule_type_name}")
        return None
    
    try:
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        
        # Time slot structure
        if time_config:
            morning_slots = time_config.get('morning_slots', ['09:00-10:30', '10:30-12:00'])
            lunch_slots = time_config.get('lunch_slots', ['12:00-13:00'])
            afternoon_slots = time_config.get('afternoon_slots', ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00'])
        else:
            morning_slots = ['09:00-10:30', '10:30-12:00']
            lunch_slots = ['12:00-13:00']
            afternoon_slots = ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00']
        all_slots = morning_slots + lunch_slots + afternoon_slots
        
        # Lecture slots (1.5 hours)
        if time_config and time_config.get('lecture_times'):
            lecture_times = time_config['lecture_times']
        else:
            lecture_times = ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00']
        
        # Tutorial slots (1 hour)
        if time_config and time_config.get('tutorial_times'):
            tutorial_times = time_config['tutorial_times']
        else:
            tutorial_times = ['14:30-15:30', '17:00-18:00']
        
        # Create schedule template
        schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
        # Mark lunch break label across provided lunch slots
        for lunch_slot in lunch_slots:
            if lunch_slot in schedule.index:
                schedule.loc[lunch_slot] = 'LUNCH BREAK'

        used_slots = set()
        
        # SEPARATE courses into core and electives
        if 'Elective (Yes/No)' in courses_df.columns:
            core_courses = courses_df[
                courses_df['Elective (Yes/No)'].astype(str).str.upper() != 'YES'
            ].copy()
            elective_courses = courses_df[
                courses_df['Elective (Yes/No)'].astype(str).str.upper() == 'YES'
            ].copy()
        else:
            core_courses = courses_df.copy()
            elective_courses = pd.DataFrame()
        
        print(f"   [COURSES] Scheduling {len(courses_df)} courses for {schedule_type_name}...")
        print(f"      Core courses: {len(core_courses)}, Elective courses: {len(elective_courses)}")
        
        # SCHEDULE ELECTIVES FIRST using basket logic
        if not elective_courses.empty and elective_allocations:
            print(f"   [BASKET] Scheduling elective baskets for {schedule_type_name}...")
            used_slots = schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section, branch)
        elif not elective_courses.empty:
            print(f"   [WARN] No basket allocations provided for electives, scheduling as core courses")
        
        # SCHEDULE CORE COURSES - iterate over core_courses only
        print(f"   [CORE] Scheduling {len(core_courses)} core courses...")
        for _, course in core_courses.iterrows():
            course_code = course['Course Code']
            ltpsc_str = course.get('LTPSC', '')
            
            # Parse LTPSC - defaults to 2 lectures, 1 tutorial if empty
            ltpsc = parse_ltpsc(ltpsc_str)
            L = ltpsc['L']
            T = ltpsc['T']
            P = ltpsc['P']
            
            # FIX: NEVER schedule labs for MA (Mathematics) courses
            is_math_course = course_code.startswith('MA') or 'Mathematics' in str(course.get('Department', ''))
            if is_math_course:
                P = 0  # Force no labs for math courses
                print(f"      [DISABLE] MA Course detected: {course_code} - Labs disabled regardless of LTPSC")
            
            # Determine lectures needed: if L is 2 or 3, schedule 2 lectures
            if L == 2 or L == 3:
                lectures_needed = 2
            elif L == 1:
                lectures_needed = 1
            else:
                # For other values, schedule min(L, 2) lectures
                lectures_needed = min(L, 2)
            
            # Determine tutorials needed: if T is 1, schedule 1 tutorial
            tutorials_needed = 1 if T >= 1 else 0
            
            # Determine labs needed: if P is 1 or more, schedule 1 lab (2 hours)
            # BUT NEVER for MA courses (already handled above)
            # If P is 0, do NOT schedule any labs regardless of other factors
            labs_needed = 1 if P >= 1 else 0
            
            print(f"      Scheduling {course_code} ({schedule_type_name}) - LTPSC: {ltpsc_str} -> L={L}, T={T}, P={P}:")
            print(f"         >> {lectures_needed} lectures, {tutorials_needed} tutorial, {labs_needed} lab")
            
            # Track which days we've used for this course
            course_day_usage = {'lectures': set(), 'tutorials': set(), 'labs': set()}
            
            # Schedule lectures (1.5 hours each)
            lectures_scheduled = 0
            max_lecture_attempts = 100
            
            while lectures_scheduled < lectures_needed and max_lecture_attempts > 0:
                max_lecture_attempts -= 1
                
                # Find days where this course doesn't have a lecture yet
                available_days = [d for d in days if d not in course_day_usage['lectures']]
                if not available_days:
                    print(f"      [WARN] Cannot schedule more lectures for {course_code} - no available days left")
                    break
                
                day = random.choice(available_days)
                time_slot = random.choice(lecture_times)
                key = (day, time_slot)
                
                if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = course_code
                    used_slots.add(key)
                    course_day_usage['lectures'].add(day)
                    lectures_scheduled += 1
                    print(f"      [OK] Scheduled lecture {lectures_scheduled} for {course_code} on {day} at {time_slot}")
            
            # Schedule tutorial (1 hour) if needed
            tutorials_scheduled = 0
            if tutorials_needed > 0:
                max_tutorial_attempts = 50
            
            while tutorials_scheduled < tutorials_needed and max_tutorial_attempts > 0:
                max_tutorial_attempts -= 1
                
                # Find days where this course doesn't have a tutorial AND doesn't have a lecture
                available_days = [d for d in days if d not in course_day_usage['tutorials'] and d not in course_day_usage['lectures']]
                if not available_days:
                    # If no completely free days, try days without tutorial but with lecture
                    available_days = [d for d in days if d not in course_day_usage['tutorials']]
                    if not available_days:
                        print(f"      [WARN] Cannot schedule tutorial for {course_code} - no available days left")
                        break
                
                day = random.choice(available_days)
                time_slot = random.choice(tutorial_times)
                key = (day, time_slot)
                
                if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                    used_slots.add(key)
                    course_day_usage['tutorials'].add(day)
                    tutorials_scheduled += 1
                    print(f"      [OK] Scheduled tutorial for {course_code} on {day} at {time_slot}")
            
            # Schedule lab (STRICTLY 2 hours) if needed - use consecutive 1.5-hour slots
            # NOTE: If P=0 in LTPSC, NO labs will be scheduled for this course
            labs_scheduled = 0
            if labs_needed > 0 and P > 0 and not is_math_course:
                max_lab_attempts = 50
                
                # Define 2-hour lab slot pairs (consecutive 1.5-hour slots that form a 2-hour lab)
                lab_slot_pairs = [
                    (['13:00-14:30', '14:30-15:30'], '13:00-15:30'),  # 2 hours: 13:00-15:30
                    (['15:30-17:00', '17:00-18:00'], '15:30-18:00'),  # 2 hours: 15:30-18:00
                ]
                
                while labs_scheduled < labs_needed and max_lab_attempts > 0:
                    max_lab_attempts -= 1
                    
                    # Find days where this course doesn't have a lab
                    available_days = [d for d in days if d not in course_day_usage['labs']]
                    if not available_days:
                        print(f"      [WARN] Cannot schedule lab for {course_code} - no available days left")
                        break
                    
                    day = random.choice(available_days)
                    # Try each lab slot pair
                    lab_scheduled = False
                    for slot_pair, lab_display_time in lab_slot_pairs:
                        # Check if both consecutive slots are free
                        slot1, slot2 = slot_pair
                        key1 = (day, slot1)
                        key2 = (day, slot2)
                        
                        if (key1 not in used_slots and key2 not in used_slots and
                            schedule.loc[slot1, day] == 'Free' and schedule.loc[slot2, day] == 'Free'):
                            # Mark both slots as lab
                            schedule.loc[slot1, day] = f"{course_code} (Lab)"
                            schedule.loc[slot2, day] = f"{course_code} (Lab)"
                            used_slots.add(key1)
                            used_slots.add(key2)
                            course_day_usage['labs'].add(day)
                            labs_scheduled += 1
                            lab_scheduled = True
                            print(f"      [OK] Scheduled lab for {course_code} on {day} at {lab_display_time} (using slots {slot1} and {slot2})")
                            break
                    
                    if not lab_scheduled:
                        # Try next attempt
                        continue
            elif P == 0:
                print(f"      [SKIP] No labs scheduled for {course_code} (P=0 in LTPSC)")
            elif is_math_course:
                print(f"      [DISABLE] Skipping lab for MA course {course_code} (P={P} in LTPSC but labs disabled for Mathematics)")
            
            # Summary
            if lectures_scheduled < lectures_needed:
                print(f"      [FAIL] Could only schedule {lectures_scheduled}/{lectures_needed} lectures for {course_code}")
            if tutorials_needed > 0 and tutorials_scheduled < tutorials_needed:
                print(f"      [FAIL] Could only schedule {tutorials_scheduled}/{tutorials_needed} tutorials for {course_code}")
            if labs_needed > 0 and labs_scheduled < labs_needed and not is_math_course:
                print(f"      [FAIL] Could only schedule {labs_scheduled}/{labs_needed} labs for {course_code}")
            if lectures_scheduled == lectures_needed and tutorials_scheduled == tutorials_needed and labs_scheduled == labs_needed:
                print(f"      [OK] Successfully scheduled {course_code} according to LTPSC structure")
        
        return schedule
        
    except Exception as e:
        print(f"[FAIL] Error generating {schedule_type} schedule: {e}")
        traceback.print_exc()
        return None

def create_course_summary(dfs, semester, branch=None):
    """Create a summary sheet showing core vs elective courses"""
    if 'course' not in dfs:
        return pd.DataFrame()
    
    sem_courses = dfs['course'][
        dfs['course']['Semester'].astype(str).str.strip() == str(semester)
    ].copy()
    
    # Filter by branch if specified
    if branch and 'Department' in sem_courses.columns:
        sem_courses = sem_courses[
            (sem_courses['Department'] == branch) | 
            (sem_courses['Elective (Yes/No)'].str.upper() == 'YES')
        ]
    
    if sem_courses.empty:
        return pd.DataFrame()
    
    # Add course type classification and parse LTPSC
    sem_courses['Course Type'] = sem_courses['Elective (Yes/No)'].apply(
        lambda x: 'Elective' if str(x).upper() == 'YES' else 'Core'
    )
    
    # Add branch specificity info
    sem_courses['Branch Specificity'] = sem_courses.apply(
        lambda row: 'Common for All Branches' if row['Course Type'] == 'Elective' else f"Department: {row.get('Department', 'General')}",
        axis=1
    )
    
    # Parse LTPSC for detailed information
    ltpsc_data = sem_courses['LTPSC'].apply(parse_ltpsc)
    sem_courses['Lectures/Week'] = ltpsc_data.apply(lambda x: x['L'])
    sem_courses['Tutorials/Week'] = ltpsc_data.apply(lambda x: x['T'])
    sem_courses['Practicals/Week'] = ltpsc_data.apply(lambda x: x['P'])
    sem_courses['Total Credits'] = ltpsc_data.apply(lambda x: x['C'])
    
    summary_columns = ['Course Code', 'Course Name', 'Course Type', 'Branch Specificity', 'LTPSC', 'Lectures/Week', 'Tutorials/Week', 'Total Credits', 'Instructor', 'Department']
    available_columns = [col for col in summary_columns if col in sem_courses.columns]
    
    return sem_courses[available_columns]

def create_common_slots_info(basket_allocations, semester):
    """Create information sheet about common slots for all departments"""
    info_data = []
    
    for basket_name, allocation in basket_allocations.items():
        lecture_days = allocation['lecture_days']
        tutorial_day = allocation['tutorial_day']
        
        # Check if lectures and tutorial are on different days
        days_separated = tutorial_day not in lecture_days
        separation_status = "[OK] DIFFERENT DAYS" if days_separated else "[FAIL] SAME DAY"
        
        info_data.append({
            'Semester': f'Semester {semester}',
            'Basket Name': basket_name,
            'Lecture Slot 1': f"{allocation['lectures'][0][0]} {allocation['lectures'][0][1]}",
            'Lecture Slot 2': f"{allocation['lectures'][1][0]} {allocation['lectures'][1][1]}",
            'Tutorial Slot': f"{allocation['tutorial'][0]} {allocation['tutorial'][1]}",
            'Lecture Days': ', '.join(lecture_days),
            'Tutorial Day': tutorial_day,
            'Days Separation': separation_status,
            'Courses in Basket': ', '.join(allocation['courses']),
            'Common for All Departments': 'Yes',
            'Common for Both Sections': 'Yes',
            'Session Type': '2 Lectures + 1 Tutorial per course'
        })
    
    return pd.DataFrame(info_data)

def allocate_electives_by_baskets(elective_courses, semester_id):
    """Allocate elective courses to COMMON time slots for ALL branches and sections of a semester"""
    print(f"[TARGET] Allocating COMMON elective slots for Semester {semester_id} (ALL branches & sections)...")
    
    # Group electives by basket
    basket_groups = {}
    for _, course in elective_courses.iterrows():
        # Normalize basket name to ensure matching (e.g., 'elective_b4', trailing spaces, etc.)
        raw_basket = course.get('Basket', 'ELECTIVE_B1')
        basket = str(raw_basket).strip().upper() if pd.notna(raw_basket) else 'ELECTIVE_B1'
        if basket not in basket_groups:
            basket_groups[basket] = []
        basket_groups[basket].append(course)
    
    print(f"   Found {len(basket_groups)} elective baskets: {list(basket_groups.keys())}")
    
    # FILTER BASKETS BY SEMESTER - Only schedule required baskets
    if semester_id == 3:
        # Semester 3: Only ELECTIVE_B3, exclude ELECTIVE_B5
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket == 'ELECTIVE_B3']
        print(f"   [TARGET] Semester 3: Scheduling only ELECTIVE_B3, excluding ELECTIVE_B5")
    elif semester_id == 5:
        # Semester 5: Schedule ELECTIVE_B5 and ELECTIVE_B4 - ALWAYS schedule both even if no courses found
        required_baskets = ['ELECTIVE_B5', 'ELECTIVE_B4']
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket in required_baskets]
        # Ensure both baskets are scheduled even if not found in course data
        for req_basket in required_baskets:
            if req_basket not in baskets_to_schedule:
                print(f"   [WARN] {req_basket} not found in course data, but will be scheduled anyway for Semester 5")
                baskets_to_schedule.append(req_basket)
                # Create empty basket group so it can be scheduled
                if req_basket not in basket_groups:
                    basket_groups[req_basket] = []
        print(f"   [TARGET] Semester 5: Scheduling BOTH ELECTIVE_B5 and ELECTIVE_B4")
    elif semester_id == 7:
        # FIXED: Semester 7: Schedule BOTH ELECTIVE_B6 and ELECTIVE_B7
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket in ['ELECTIVE_B6', 'ELECTIVE_B7']]
        print(f"   [TARGET] Semester 7: Scheduling BOTH ELECTIVE_B6 and ELECTIVE_B7")
    else:
        # Other semesters: Schedule all baskets
        baskets_to_schedule = list(basket_groups.keys())
        print(f"   [TARGET] Semester {semester_id}: Scheduling all baskets")
    
    print(f"   [LIST] Baskets to schedule: {baskets_to_schedule}")
    
    # FIXED COMMON SLOTS for ALL branches and sections - ADDED B6 and B7
    common_slots_mapping = {
        'ELECTIVE_B1': {
            'lectures': [('Mon', '09:00-10:30'), ('Wed', '09:00-10:30')],
            'tutorial': ('Fri', '14:30-15:30')
        },
        'ELECTIVE_B2': {
            'lectures': [('Tue', '09:00-10:30'), ('Thu', '09:00-10:30')],
            'tutorial': ('Mon', '14:30-15:30')
        },
        'ELECTIVE_B3': {
            'lectures': [('Mon', '13:00-14:30'), ('Wed', '13:00-14:30')],
            'tutorial': ('Tue', '14:30-15:30')
        },
        'ELECTIVE_B4': {
            'lectures': [('Tue', '13:00-14:30'), ('Thu', '13:00-14:30')],
            'tutorial': ('Wed', '14:30-15:30')
        },
        'ELECTIVE_B5': {
            'lectures': [('Mon', '15:30-17:00'), ('Wed', '15:30-17:00')],
            'tutorial': ('Thu', '14:30-15:30')
        },
        'ELECTIVE_B6': {
            'lectures': [('Mon', '09:00-10:30'), ('Wed', '13:00-14:30')],
            'tutorial': ('Tue', '14:30-15:30')
        },
        'ELECTIVE_B7': {
            'lectures': [('Tue', '09:00-10:30'), ('Thu', '13:00-14:30')],
            'tutorial': ('Wed', '14:30-15:30')
        },
        'HSS_B1': {
            'lectures': [('Mon', '15:30-17:00'), ('Wed', '15:30-17:00')],
            'tutorial': ('Thu', '14:30-15:30')
        },
        'HSS_B2': {
            'lectures': [('Tue', '15:30-17:00'), ('Thu', '15:30-17:00')],
            'tutorial': ('Fri', '14:30-15:30')
        }
    }
    
    elective_allocations = {}
    basket_allocations = {}
    
    for basket_name in sorted(baskets_to_schedule):  # Only iterate through filtered baskets
        if basket_name not in basket_groups:
            print(f"   [WARN] Basket {basket_name} not found in course data")
            continue
            
        basket_courses = basket_groups[basket_name]
        course_codes = [course['Course Code'] for course in basket_courses] if basket_courses else []
        
        # Parse LTPSC for the first course in the basket (assume all courses in basket have same LTPSC)
        # If LTPSC is empty or no courses, default to 2 lectures and 1 tutorial
        if basket_courses and len(basket_courses) > 0:
            first_course = basket_courses[0]
            ltpsc_str = first_course.get('LTPSC', '')
            ltpsc = parse_ltpsc(ltpsc_str)
            L = ltpsc['L']
            T = ltpsc['T']
            P = ltpsc['P']
        else:
            # Empty basket - use default LTPSC: 2 lectures, 1 tutorial
            ltpsc_str = '2-1-0-0-3'
            ltpsc = parse_ltpsc(ltpsc_str)
            L = 2
            T = 1
            P = 0
            print(f"   [WARN] Basket '{basket_name}' has no courses, using default LTPSC: 2-1-0-0-3")
        
        # Determine lectures needed: if L is 2 or 3, schedule 2 lectures
        if L == 2 or L == 3:
            lectures_needed = 2
        elif L == 1:
            lectures_needed = 1
        else:
            lectures_needed = min(L, 2)
        
        # Determine tutorials needed: if T is 1, schedule 1 tutorial
        tutorials_needed = 1 if T >= 1 else 0
        
        print(f"   [BASKET] Basket '{basket_name}' LTPSC: {ltpsc_str} -> L={L}, T={T}, P={P}")
        print(f"      → Scheduling {lectures_needed} lectures, {tutorials_needed} tutorial")

        if basket_name in common_slots_mapping:
            fixed_slots = common_slots_mapping[basket_name]
            # Use only the number of lectures needed based on LTPSC
            lectures_allocated = fixed_slots['lectures'][:lectures_needed]
            
            # FIX: Ensure tutorial is allocated when needed
            if tutorials_needed > 0:
                tutorial_allocated = fixed_slots['tutorial']
                print(f"      [OK] Tutorial allocation: {tutorial_allocated}")
            else:
                tutorial_allocated = None
                print(f"      [WARN] No tutorial needed (T={T})")
        else:
            print(f"      [FAIL] No fixed slots found for basket '{basket_name}'")
        
        # Verify day separation
        lecture_days = set(day for day, time in lectures_allocated)
        tutorial_day = tutorial_allocated[0] if tutorial_allocated else None
        days_separated = tutorial_day not in lecture_days if tutorial_day else True
        
        # Store allocation for ALL courses in this basket
        # If no courses, create a dummy allocation entry to ensure basket is scheduled
        if course_codes:
            for course_code in course_codes:
                elective_allocations[course_code] = {
                    'basket_name': basket_name,
                    'lectures': lectures_allocated,
                    'tutorial': tutorial_allocated,
                    'all_courses_in_basket': course_codes,
                    'for_all_branches': True,
                    'for_both_sections': True,
                    'common_for_semester': True,
                    'common_for_all_departments': True,
                    'lecture_days': list(lecture_days),
                    'tutorial_day': tutorial_day,
                    'days_separated': days_separated,
                    'fixed_common_slots': True,
                    'ltpsc': ltpsc_str,
                    'lectures_needed': lectures_needed,
                    'tutorials_needed': tutorials_needed
                }
        else:
            # Empty basket - create a dummy allocation to ensure it's scheduled
            dummy_key = f"__BASKET_{basket_name}__"
            elective_allocations[dummy_key] = {
                'basket_name': basket_name,
                'lectures': lectures_allocated,
                'tutorial': tutorial_allocated,
                'all_courses_in_basket': [],
                'for_all_branches': True,
                'for_both_sections': True,
                'common_for_semester': True,
                'common_for_all_departments': True,
                'lecture_days': list(lecture_days),
                'tutorial_day': tutorial_day,
                'days_separated': days_separated,
                'fixed_common_slots': True,
                'ltpsc': ltpsc_str,
                'lectures_needed': lectures_needed,
                'tutorials_needed': tutorials_needed
            }
            print(f"   [NOTE] Created dummy allocation for empty basket '{basket_name}' to ensure scheduling")
        
        basket_allocations[basket_name] = {
            'lectures': lectures_allocated,
            'tutorial': tutorial_allocated,
            'courses': course_codes,
            'common_for_all_departments': True,
            'lecture_days': list(lecture_days),
            'tutorial_day': tutorial_day,
            'days_separated': days_separated,
            'fixed_common_slots': True,
            'ltpsc': ltpsc_str,
            'lectures_needed': lectures_needed,
            'tutorials_needed': tutorials_needed
        }
        
        print(f"   [BASKET] FIXED COMMON SLOTS for Basket '{basket_name}':")
        print(f"      [PIN] SAME FOR ALL BRANCHES & SECTIONS (LTPSC-based)")
        for i, (day, time_slot) in enumerate(lectures_allocated, 1):
            print(f"      Lecture {i}: {day} {time_slot}")
        if tutorial_allocated:
            print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
        print(f"      Days Separation: {'[OK] DIFFERENT DAYS' if days_separated else '[FAIL] SAME DAY'}")
        print(f"      Courses: {', '.join(course_codes)}")
    
    # Log excluded baskets
    excluded_baskets = set(basket_groups.keys()) - set(baskets_to_schedule)
    if excluded_baskets:
        print(f"   [DISABLE] Excluded baskets for Semester {semester_id}: {list(excluded_baskets)}")
    
    return elective_allocations, basket_allocations

def create_common_basket_summary(basket_allocations, semester, branch=None):
    """Create basket summary highlighting common slots"""
    summary_data = []
    
    for basket_name, allocation in basket_allocations.items():
        summary_data.append({
            'Basket Name': basket_name,
            'Lecture Slot 1': f"{allocation['lectures'][0][0]} {allocation['lectures'][0][1]}",
            'Lecture Slot 2': f"{allocation['lectures'][1][0]} {allocation['lectures'][1][1]}",
            'Tutorial Slot': f"{allocation['tutorial'][0]} {allocation['tutorial'][1]}",
            'Courses in Basket': ', '.join(allocation['courses']),
            'Common for All Branches': '[OK] YES',
            'Common for Both Sections': '[OK] YES', 
            'Days Separation': '[OK] YES' if allocation['days_separated'] else '[FAIL] NO',
            'Semester': f'Semester {semester}',
            'Applicable Branches': 'CSE, DSAI, ECE (ALL)',
            'Slot Type': 'FIXED COMMON SLOTS'
        })
    
    return pd.DataFrame(summary_data)


def create_detailed_common_slots_info(basket_allocations, semester):
    """Create detailed information about common slots"""
    info_data = []
    
    for basket_name, allocation in basket_allocations.items():
        info_data.append({
            'Semester': f'Semester {semester}',
            'Basket Name': basket_name,
            'Lecture 1 Day': allocation['lectures'][0][0],
            'Lecture 1 Time': allocation['lectures'][0][1],
            'Lecture 2 Day': allocation['lectures'][1][0],
            'Lecture 2 Time': allocation['lectures'][1][1],
            'Tutorial Day': allocation['tutorial'][0],
            'Tutorial Time': allocation['tutorial'][1],
            'Courses': ', '.join(allocation['courses']),
            'Common for Branches': 'CSE, DSAI, ECE (ALL)',
            'Common for Sections': 'A & B (BOTH)',
            'Days Separation': '[OK] Achieved' if allocation['days_separated'] else '[FAIL] Not Achieved',
            'Slot Consistency': '[OK] IDENTICAL across all',
            'Notes': 'FIXED COMMON TIMETABLE SLOTS'
        })
    
    return pd.DataFrame(info_data)

def create_semester_rules_sheet(semester, basket_allocations):
    """Create a sheet explaining semester-specific elective rules"""
    rules_data = []
    
    if semester == 3:
        rules_data.append({
            'Semester': 'Semester 3',
            'Rule': 'Schedule only ELECTIVE_B3',
            'Exclusion': 'Exclude ELECTIVE_B5',
            'Reason': 'Curriculum requirement - Semester 3 focuses on B3 electives',
            'Scheduled Baskets': ', '.join(basket_allocations.keys()) if basket_allocations else 'None',
            'Status': '[OK] Applied'
        })
    elif semester == 5:
        rules_data.append({
            'Semester': 'Semester 5', 
            'Rule': 'Schedule only ELECTIVE_B5',
            'Exclusion': 'Exclude ELECTIVE_B3',
            'Reason': 'Curriculum requirement - Semester 5 focuses on B5 electives',
            'Scheduled Baskets': ', '.join(basket_allocations.keys()) if basket_allocations else 'None',
            'Status': '[OK] Applied'
        })
    else:
        rules_data.append({
            'Semester': f'Semester {semester}',
            'Rule': 'Schedule all elective baskets',
            'Exclusion': 'None',
            'Reason': 'No specific restrictions for this semester',
            'Scheduled Baskets': ', '.join(basket_allocations.keys()) if basket_allocations else 'None',
            'Status': '[OK] Applied'
        })
    
    return pd.DataFrame(rules_data)

def print_classroom_allocation_summary(semester, branch):
    """Print a summary of classroom allocations"""
    global _TIMETABLE_CLASSROOM_ALLOCATIONS
    timetable_key = f"{branch}_sem{semester}"
    
    room_usage = {}
    for key, allocations in _TIMETABLE_CLASSROOM_ALLOCATIONS.items():
        if key.startswith(timetable_key):
            for allocation_key, allocation in allocations.items():
                room = allocation['classroom']
                if room not in room_usage:
                    room_usage[room] = 0
                room_usage[room] += 1
    
    print(f"   [STATS] Classroom usage summary for {branch} Semester {semester}:")
    for room, count in sorted(room_usage.items()):
        print(f"      {room}: {count} sessions")

def create_classroom_allocation_detail_with_tracking(timetable_schedules, classrooms_df, semester, branch):
    """Create detailed classroom allocation information with global tracking"""
    allocation_data = []
    global _TIMETABLE_CLASSROOM_ALLOCATIONS
    
    for i, schedule in enumerate(timetable_schedules, 1):
        section = 'A' if i == 1 else 'B'
        timetable_key = f"{branch}_sem{semester}_sec{section}"
        
        for day in schedule.columns:
            for time_slot in schedule.index:
                cell_value = str(schedule.loc[time_slot, day])
                
                if cell_value not in ['Free', 'LUNCH BREAK'] and '[' in cell_value and ']' in cell_value:
                    # Extract course and room
                    course_match = re.search(r'^(.*?)\s*\[', cell_value)
                    room_match = re.search(r'\[(.*?)\]', cell_value)
                    
                    if course_match and room_match:
                        course = course_match.group(1).strip()
                        room_number = room_match.group(1)
                        
                        # Get room details
                        room_details = classrooms_df[classrooms_df['Room Number'] == room_number]
                        capacity = room_details['Capacity'].iloc[0] if not room_details.empty else 'Unknown'
                        room_type = room_details['Type'].iloc[0] if not room_details.empty else 'Unknown'
                        
                        allocation_data.append({
                            'Semester': semester,
                            'Branch': branch,
                            'Section': section,
                            'Day': day,
                            'Time Slot': time_slot,
                            'Course': course,
                            'Room Number': room_number,
                            'Room Type': room_type,
                            'Capacity': capacity,
                            'Facilities': room_details['Facilities'].iloc[0] if not room_details.empty else 'Unknown',
                            'Allocation Type': 'Global Tracking'
                        })
    
    return pd.DataFrame(allocation_data)

def export_semester_timetable_with_baskets(dfs, semester, branch=None, time_config=None):
    """Export timetable using IDENTICAL COMMON elective slots for ALL branches and sections with classroom allocation.
    Accepts optional time_config to override slot timings."""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"\n[STATS] Generating timetable for Semester {semester}{branch_info}...")
    
    # Reset classroom tracker at the start of generation for this branch/semester
    reset_classroom_usage_tracker()
    
    # Show semester-specific basket rules
    if semester == 3:
        print(f"   [TARGET] SEMESTER 3 RULE: Scheduling only ELECTIVE_B3, excluding ELECTIVE_B5")
    elif semester == 5:
        print(f"   [TARGET] SEMESTER 5 RULE: Scheduling BOTH ELECTIVE_B5 and ELECTIVE_B4")
    elif semester == 7:
        print(f"   [TARGET] SEMESTER 7 RULE: Scheduling BOTH ELECTIVE_B6 and ELECTIVE_B7")
    else:
        print(f"   [TARGET] SEMESTER {semester}: Scheduling all elective baskets")
    
    try:
        # Get ALL elective courses for this semester (without branch filter)
        course_baskets_all = separate_courses_by_type(dfs, semester)
        elective_courses_all = course_baskets_all['elective_courses']
        
        print(f"[TARGET] Elective courses for Semester {semester} (COMMON for ALL): {len(elective_courses_all)}")
        
        # Allocate electives using FIXED COMMON slots (with semester filtering)
        elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)
        
        print(f"   📅 FINAL SCHEDULED BASKETS for Semester {semester}:")
        if basket_allocations:
            for basket_name, allocation in basket_allocations.items():
                status = "[OK] VALID" if allocation['days_separated'] else "[FAIL] INVALID"
                print(f"      {basket_name}: {status}")
        
        # Generate schedules - these will have IDENTICAL elective slots
        # Pass basket_allocations to ensure all required baskets are scheduled
        section_a = generate_section_schedule_with_elective_baskets(dfs, semester, 'A', elective_allocations, branch, time_config=time_config, basket_allocations=basket_allocations)
        section_b = generate_section_schedule_with_elective_baskets(dfs, semester, 'B', elective_allocations, branch, time_config=time_config, basket_allocations=basket_allocations)
        
        if section_a is None or section_b is None:
            return False

        # ALLOCATE CLASSROOMS for both sections with proper tracking
        course_info = get_course_info(dfs) if dfs else {}
        classroom_data = dfs.get('classroom')
        
        if classroom_data is not None and not classroom_data.empty:
            print("[SCHOOL] Allocating classrooms with global tracking...")
            section_a_with_rooms = allocate_classrooms_for_timetable(
                section_a, classroom_data, course_info, semester, branch, 'A'
            )
            section_b_with_rooms = allocate_classrooms_for_timetable(
                section_b, classroom_data, course_info, semester, branch, 'B'
            )
            
            # Check if classroom allocation was successful
            has_classroom_allocation_a = check_for_classroom_allocation(section_a_with_rooms)
            has_classroom_allocation_b = check_for_classroom_allocation(section_b_with_rooms)
            has_classroom_allocation = has_classroom_allocation_a or has_classroom_allocation_b
            
            if has_classroom_allocation:
                print("   [OK] Classroom allocation completed with global tracking")
                # Print allocation summary
                print_classroom_allocation_summary(semester, branch)
            else:
                print("   [WARN] Classroom allocation attempted but no rooms were assigned")
        else:
            section_a_with_rooms = section_a
            section_b_with_rooms = section_b
            print("[WARN]  No classroom data available for allocation")

        # Create filename
        filename = f"sem{semester}_{branch}_timetable_baskets.xlsx" if branch else f"sem{semester}_timetable_baskets.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Save schedules with classroom allocation
            section_a_with_rooms.to_excel(writer, sheet_name='Section_A')
            section_b_with_rooms.to_excel(writer, sheet_name='Section_B')
            
            # ========== ADDED: COMPREHENSIVE VERIFICATION SHEETS ==========
            print("[STATS] Generating comprehensive verification sheets...")
            
            # Get course info for verification
            course_info = get_course_info(dfs) if dfs else {}
            
            # 1. Create detailed verification sheet for Section A
            if not section_a_with_rooms.empty:
                print(f"   Creating verification sheet for Section A...")
                verification_a = create_timetable_verification_sheet(
                    section_a_with_rooms, course_info, classroom_data, semester, branch, 'A'
                )
                if not verification_a.empty:
                    verification_a.to_excel(writer, sheet_name='Verification_A', index=False)
                    print(f"   [OK] Created Verification_A sheet with {len(verification_a)} entries")
            
            # 2. Create detailed verification sheet for Section B
            if not section_b_with_rooms.empty:
                print(f"   Creating verification sheet for Section B...")
                verification_b = create_timetable_verification_sheet(
                    section_b_with_rooms, course_info, classroom_data, semester, branch, 'B'
                )
                if not verification_b.empty:
                    verification_b.to_excel(writer, sheet_name='Verification_B', index=False)
                    print(f"   [OK] Created Verification_B sheet with {len(verification_b)} entries")
            
            # 3. Create room allocation summary
            if classroom_data is not None and not classroom_data.empty:
                print(f"   Creating room allocation summary...")
                room_summary = create_room_allocation_summary_verification(
                    section_a_with_rooms, section_b_with_rooms, classroom_data
                )
                if not room_summary.empty:
                    room_summary.to_excel(writer, sheet_name='Room_Allocation', index=False)
                    print(f"   [OK] Created Room_Allocation sheet with {len(room_summary)} rooms")
            
            # 4. Create LTPSC compliance summary
            print(f"   Creating LTPSC compliance summary...")
            ltpsc_summary = create_ltpsc_compliance_summary(
                dfs, semester, branch, section_a_with_rooms, section_b_with_rooms
            )
            if not ltpsc_summary.empty:
                ltpsc_summary.to_excel(writer, sheet_name='LTPSC_Compliance', index=False)
                print(f"   [OK] Created LTPSC_Compliance sheet with {len(ltpsc_summary)} courses")
            
            # 5. Create executive summary
            print(f"   Creating executive summary...")
            exec_summary = create_executive_summary(
                dfs, semester, branch, section_a_with_rooms, section_b_with_rooms, basket_allocations
            )
            if not exec_summary.empty:
                exec_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)
                print(f"   [OK] Created Executive_Summary sheet")
            
            # ========== EXISTING SHEETS ==========
            # Add basket allocation summary
            if basket_allocations:
                basket_summary = create_common_basket_summary(basket_allocations, semester, branch)
                basket_summary.to_excel(writer, sheet_name='Basket_Allocation', index=False)
            
            # Add course summary
            course_summary = create_course_summary(dfs, semester, branch)
            if not course_summary.empty:
                course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
            
            # Add basket courses details
            basket_courses_sheet = create_basket_courses_sheet(basket_allocations)
            if not basket_courses_sheet.empty:
                basket_courses_sheet.to_excel(writer, sheet_name='Basket_Courses', index=False)
            
            # Add classroom utilization report
            if classroom_data is not None and not classroom_data.empty:
                classroom_report = create_classroom_utilization_report(
                    classroom_data, [section_a_with_rooms, section_b_with_rooms], []
                )
                classroom_report.to_excel(writer, sheet_name='Classroom_Utilization', index=False)
                
                # Add detailed classroom allocation with global tracking info
                classroom_allocation_detail = create_classroom_allocation_detail_with_tracking(
                    [section_a_with_rooms, section_b_with_rooms], classroom_data, semester, branch
                )
                classroom_allocation_detail.to_excel(writer, sheet_name='Classroom_Allocation', index=False)

            # Persist configuration if provided
            if time_config:
                try:
                    config_items = [{'Parameter': k, 'Value': str(v)} for k, v in time_config.items()]
                    pd.DataFrame(config_items).to_excel(writer, sheet_name='Configuration', index=False)
                except Exception as _:
                    pass
            # ========== END OF ADDED SECTION ==========
        
        success_message = f"[OK] Timetable saved: {filename}"
        if classroom_data is not None and not classroom_data.empty:
            success_message += " (with classroom allocation)"
        
        print(success_message)
        print(f"[STATS] Added comprehensive verification sheets for easy inspection")
        return True
        
    except Exception as e:
        print(f"[FAIL] Error generating timetable: {e}")
        traceback.print_exc()
        return False

def allocate_mid_semester_electives_by_baskets(elective_courses, semester_id):
    """Allocate elective courses in pre-mid/post-mid to COMMON time slots for ALL branches and sections
    
    Similar to allocate_electives_by_baskets but for mid-semester courses
    """
    if elective_courses.empty:
        print(f"   [INFO] No elective courses to allocate for Semester {semester_id}")
        return {}
    
    print(f"   [BASKET] Allocating elective slots for mid-semester, Semester {semester_id}...")
    
    # Group electives by basket
    basket_groups = {}
    for _, course in elective_courses.iterrows():
        raw_basket = course.get('Basket', 'ELECTIVE_B1')
        basket = str(raw_basket).strip().upper() if pd.notna(raw_basket) else 'ELECTIVE_B1'
        if basket not in basket_groups:
            basket_groups[basket] = []
        basket_groups[basket].append(course)
    
    print(f"      Found {len(basket_groups)} elective baskets: {list(basket_groups.keys())}")
    
    # FILTER BASKETS BY SEMESTER (same logic as regular timetables)
    if semester_id == 3:
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket == 'ELECTIVE_B3']
        print(f"      [TARGET] Semester 3: Scheduling only ELECTIVE_B3")
    elif semester_id == 5:
        required_baskets = ['ELECTIVE_B5', 'ELECTIVE_B4']
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket in required_baskets]
        for req_basket in required_baskets:
            if req_basket not in baskets_to_schedule:
                baskets_to_schedule.append(req_basket)
                if req_basket not in basket_groups:
                    basket_groups[req_basket] = []
        print(f"      [TARGET] Semester 5: Scheduling BOTH ELECTIVE_B5 and ELECTIVE_B4")
    elif semester_id == 7:
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket in ['ELECTIVE_B6', 'ELECTIVE_B7']]
        print(f"      [TARGET] Semester 7: Scheduling BOTH ELECTIVE_B6 and ELECTIVE_B7")
    else:
        baskets_to_schedule = list(basket_groups.keys())
        print(f"      [TARGET] Semester {semester_id}: Scheduling all baskets")
    
    # COMMON SLOTS mapping (same as regular timetables)
    common_slots_mapping = {
        'ELECTIVE_B1': {
            'lectures': [('Mon', '09:00-10:30'), ('Wed', '09:00-10:30')],
            'tutorial': ('Fri', '14:30-15:30')
        },
        'ELECTIVE_B2': {
            'lectures': [('Tue', '09:00-10:30'), ('Thu', '09:00-10:30')],
            'tutorial': ('Mon', '14:30-15:30')
        },
        'ELECTIVE_B3': {
            'lectures': [('Mon', '13:00-14:30'), ('Wed', '13:00-14:30')],
            'tutorial': ('Tue', '14:30-15:30')
        },
        'ELECTIVE_B4': {
            'lectures': [('Tue', '13:00-14:30'), ('Thu', '13:00-14:30')],
            'tutorial': ('Wed', '14:30-15:30')
        },
        'ELECTIVE_B5': {
            'lectures': [('Mon', '15:30-17:00'), ('Wed', '15:30-17:00')],
            'tutorial': ('Thu', '14:30-15:30')
        },
        'ELECTIVE_B6': {
            'lectures': [('Mon', '09:00-10:30'), ('Wed', '13:00-14:30')],
            'tutorial': ('Tue', '14:30-15:30')
        },
        'ELECTIVE_B7': {
            'lectures': [('Tue', '09:00-10:30'), ('Thu', '13:00-14:30')],
            'tutorial': ('Wed', '14:30-15:30')
        },
        'HSS_B1': {
            'lectures': [('Mon', '15:30-17:00'), ('Wed', '15:30-17:00')],
            'tutorial': ('Thu', '14:30-15:30')
        },
        'HSS_B2': {
            'lectures': [('Tue', '15:30-17:00'), ('Thu', '15:30-17:00')],
            'tutorial': ('Fri', '14:30-15:30')
        }
    }
    
    elective_allocations = {}
    
    for basket_name in sorted(baskets_to_schedule):
        if basket_name not in basket_groups:
            continue
        
        basket_courses = basket_groups[basket_name]
        course_codes = [course['Course Code'] for course in basket_courses] if basket_courses else []
        
        # Parse LTPSC from first course in basket
        if basket_courses and len(basket_courses) > 0:
            first_course = basket_courses[0]
            ltpsc_str = first_course.get('LTPSC', '')
            ltpsc = parse_ltpsc(ltpsc_str)
            L = ltpsc['L']
            T = ltpsc['T']
        else:
            ltpsc_str = '2-1-0-0-3'
            ltpsc = parse_ltpsc(ltpsc_str)
            L = 2
            T = 1
        
        # Determine lectures and tutorials needed
        if L == 2 or L == 3:
            lectures_needed = 2
        elif L == 1:
            lectures_needed = 1
        else:
            lectures_needed = min(L, 2)
        
        tutorials_needed = 1 if T >= 1 else 0
        
        print(f"      [BASKET] Basket '{basket_name}' LTPSC: {ltpsc_str} -> L={L}, T={T}")
        print(f"         → Scheduling {lectures_needed} lectures, {tutorials_needed} tutorial")
        
        if basket_name in common_slots_mapping:
            fixed_slots = common_slots_mapping[basket_name]
            lectures_allocated = fixed_slots['lectures'][:lectures_needed]
            tutorial_allocated = fixed_slots['tutorial'] if tutorials_needed > 0 else None
            
            # Calculate if lectures and tutorials are on separate days
            if tutorial_allocated:
                lecture_days = set([day for day, _ in lectures_allocated])
                tutorial_day = tutorial_allocated[0]
                days_separated = tutorial_day not in lecture_days
            else:
                days_separated = True  # No tutorial, so separation doesn't apply
            
            # Build allocation for this basket - ENSURE all_courses_in_basket is a list of strings
            allocation = {
                'basket_name': basket_name,
                'lectures': lectures_allocated,
                'tutorial': tutorial_allocated,
                'courses': course_codes,
                'all_courses_in_basket': course_codes,  # Use course codes list, not series objects
                'lecture_days': list(set([day for day, _ in lectures_allocated])),
                'tutorial_day': tutorial_allocated[0] if tutorial_allocated else None,
                'days_separated': days_separated
            }
            
            elective_allocations[basket_name] = allocation
            print(f"         → Lectures: {lectures_allocated}, Tutorial: {tutorial_allocated}, Days Separated: {days_separated}")
        else:
            print(f"      [WARN] Basket '{basket_name}' not in common slots mapping")
    
    print(f"   [OK] Elective allocation complete: {len(elective_allocations)} baskets allocated")
    return elective_allocations

def export_mid_semester_timetables(dfs, semester, branch=None, time_config=None):
    """Export separate pre-mid and post-mid timetables
    
    Note: CSE has Section A and Section B
           DSAI and ECE are treated as a whole (no sections)
    """
    branch_info = f", Branch {branch}" if branch else ""
    print(f"\n[STATS] Generating MID-SEMESTER timetables for Semester {semester}{branch_info}...")
    
    try:
        # Separate courses into pre-mid and post-mid
        mid_semester_courses = separate_courses_by_mid_semester(dfs, semester, branch)
        pre_mid_courses = mid_semester_courses['pre_mid_courses']
        post_mid_courses = mid_semester_courses['post_mid_courses']
        
        print(f"[LIST] Course distribution:")
        print(f"   Pre-mid courses: {len(pre_mid_courses)}")
        print(f"   Post-mid courses: {len(post_mid_courses)}")
        
        # ALLOCATE ELECTIVES TO BASKETS for pre-mid and post-mid
        print(f"\n[BASKET] Allocating electives to baskets...")
        if not pre_mid_courses.empty:
            pre_mid_electives = pre_mid_courses[
                pre_mid_courses['Elective (Yes/No)'].astype(str).str.upper() == 'YES'
            ] if 'Elective (Yes/No)' in pre_mid_courses.columns else pd.DataFrame()
            pre_mid_elective_allocations = allocate_mid_semester_electives_by_baskets(pre_mid_electives, semester)
        else:
            pre_mid_elective_allocations = {}
        
        if not post_mid_courses.empty:
            post_mid_electives = post_mid_courses[
                post_mid_courses['Elective (Yes/No)'].astype(str).str.upper() == 'YES'
            ] if 'Elective (Yes/No)' in post_mid_courses.columns else pd.DataFrame()
            post_mid_elective_allocations = allocate_mid_semester_electives_by_baskets(post_mid_electives, semester)
        else:
            post_mid_elective_allocations = {}
        
        # Determine if this branch has sections (only CSE has sections A and B)
        has_sections = (branch == 'CSE')
        if has_sections:
            sections = ['A', 'B']
            print(f"[INFO] CSE branch detected - will generate timetables for Section A and Section B")
        else:
            sections = ['Whole']
            print(f"[INFO] {branch} branch detected - will generate single timetable (no sections)")
        
        # Generate pre-mid timetables
        pre_mid_sections = {}
        if not pre_mid_courses.empty:
            print(f"[RESET] Generating PRE-MID timetables...")
            print(f"   Courses to schedule: {pre_mid_courses['Course Code'].tolist()}")
            
            for section in sections:
                section_label = f"Section {section}" if has_sections else "Whole Branch"
                try:
                    pre_mid_sections[section] = generate_mid_semester_schedule(
                        dfs, semester, section, pre_mid_courses, branch, time_config, 'pre_mid', pre_mid_elective_allocations
                    )
                    print(f"   [OK] Pre-mid {section_label}: {'Generated' if pre_mid_sections[section] is not None else 'Failed'}")
                except Exception as e:
                    print(f"   [FAIL] Error generating pre-mid {section_label}: {e}")
                    traceback.print_exc()
                    pre_mid_sections[section] = None
        else:
            print(f"[WARN] No pre-mid courses to schedule for Semester {semester}, Branch {branch}")
        
        # Generate post-mid timetables
        post_mid_sections = {}
        if not post_mid_courses.empty:
            print(f"[RESET] Generating POST-MID timetables...")
            print(f"   Courses to schedule: {post_mid_courses['Course Code'].tolist()}")
            
            for section in sections:
                section_label = f"Section {section}" if has_sections else "Whole Branch"
                try:
                    post_mid_sections[section] = generate_mid_semester_schedule(
                        dfs, semester, section, post_mid_courses, branch, time_config, 'post_mid', post_mid_elective_allocations
                    )
                    print(f"   [OK] Post-mid {section_label}: {'Generated' if post_mid_sections[section] is not None else 'Failed'}")
                except Exception as e:
                    print(f"   [FAIL] Error generating post-mid {section_label}: {e}")
                    traceback.print_exc()
                    post_mid_sections[section] = None
        else:
            print(f"[WARN] No post-mid courses to schedule for Semester {semester}, Branch {branch}")
        
        # Create filenames
        base_filename = f"sem{semester}_{branch}_" if branch else f"sem{semester}_"
        pre_mid_filename = base_filename + "pre_mid_timetable.xlsx"
        post_mid_filename = base_filename + "post_mid_timetable.xlsx"
        
        pre_mid_filepath = os.path.join(OUTPUT_DIR, pre_mid_filename)
        post_mid_filepath = os.path.join(OUTPUT_DIR, post_mid_filename)
        
        # Determine sheet names based on whether this branch has sections
        if has_sections:
            sheet_names = {'A': 'Section_A', 'B': 'Section_B', 'Whole': 'Timetable'}
        else:
            sheet_names = {'A': 'Timetable', 'B': None, 'Whole': 'Timetable'}
        
        # Save pre-mid timetable
        pre_mid_all_valid = all(pre_mid_sections.get(s) is not None for s in sections)
        if pre_mid_all_valid:
            try:
                with pd.ExcelWriter(pre_mid_filepath, engine='openpyxl') as writer:
                    for section in sections:
                        pre_mid_data = pre_mid_sections[section]
                        # Reset index to ensure Time Slot is a column, not index
                        pre_mid_reset = pre_mid_data.reset_index()
                        pre_mid_reset = pre_mid_reset.rename(columns={'index': 'Time Slot'})
                        sheet_name = sheet_names.get(section, f'Section_{section}')
                        pre_mid_reset.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # ========== ADDED: VERIFICATION STATISTICS ==========
                    print(f"[STATS] Generating PRE-MID verification sheets...")
                    
                    # Get course info for verification
                    course_info = get_course_info(dfs) if dfs else {}
                    classroom_data = dfs.get('classroom', None) if dfs else None
                    
                    # Create verification sheets for each section
                    for section in sections:
                        pre_mid_reset = pre_mid_sections[section].reset_index().rename(columns={'index': 'Time Slot'})
                        if not pre_mid_reset.empty:
                            section_label = f"Section {section}" if has_sections else "Timetable"
                            print(f"   Creating Pre-Mid Verification sheet for {section_label}...")
                            verification = create_timetable_verification_sheet(
                                pre_mid_reset, course_info, classroom_data, semester, branch, section
                            )
                            if not verification.empty:
                                sheet_name = f'Verification_{section}' if has_sections else 'Verification'
                                verification.to_excel(writer, sheet_name=sheet_name, index=False)
                                print(f"   [OK] Created Pre-Mid Verification sheet with {len(verification)} entries")
                    
                    # Create room allocation summary
                    if classroom_data is not None and not classroom_data.empty:
                        print(f"   Creating room allocation summary...")
                        all_sections_data = [pre_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                        room_summary = create_room_allocation_summary_verification(
                            all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame(), classroom_data
                        )
                        if not room_summary.empty:
                            room_summary.to_excel(writer, sheet_name='Room_Allocation', index=False)
                            print(f"   [OK] Created Room_Allocation sheet with {len(room_summary)} rooms")
                    
                    # Create LTPSC compliance summary
                    print(f"   Creating LTPSC compliance summary...")
                    all_sections_data = [pre_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                    ltpsc_summary = create_ltpsc_compliance_summary(
                        dfs, semester, branch, all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame()
                    )
                    if not ltpsc_summary.empty:
                        ltpsc_summary.to_excel(writer, sheet_name='LTPSC_Compliance', index=False)
                        print(f"   [OK] Created LTPSC_Compliance sheet with {len(ltpsc_summary)} courses")
                    
                    # Create executive summary
                    print(f"   Creating executive summary...")
                    all_sections_data = [pre_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                    exec_summary = create_executive_summary(
                        dfs, semester, branch, all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame(), {}
                    )
                    if not exec_summary.empty:
                        exec_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)
                        print(f"   [OK] Created Executive_Summary sheet")
                    
                    # Add course summary
                    pre_mid_summary = create_mid_semester_summary(pre_mid_courses, 'Pre-Mid', semester, branch)
                    pre_mid_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
                    
                    print(f"[STATS] Added PRE-MID verification sheets for easy inspection")
                    print(f"[OK] PRE-MID timetable saved: {pre_mid_filename}")
            except Exception as e:
                print(f"[FAIL] Error saving pre-mid timetable: {e}")
                traceback.print_exc()
        else:
            print(f"[WARN] Cannot save pre-mid timetable - Not all sections generated successfully")
        
        # Save post-mid timetable
        post_mid_all_valid = all(post_mid_sections.get(s) is not None for s in sections)
        if post_mid_all_valid:
            try:
                with pd.ExcelWriter(post_mid_filepath, engine='openpyxl') as writer:
                    for section in sections:
                        post_mid_data = post_mid_sections[section]
                        # Reset index to ensure Time Slot is a column, not index
                        post_mid_reset = post_mid_data.reset_index()
                        post_mid_reset = post_mid_reset.rename(columns={'index': 'Time Slot'})
                        sheet_name = sheet_names.get(section, f'Section_{section}')
                        post_mid_reset.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    # ========== ADDED: VERIFICATION STATISTICS ==========
                    print(f"[STATS] Generating POST-MID verification sheets...")
                    
                    # Get course info for verification
                    course_info = get_course_info(dfs) if dfs else {}
                    classroom_data = dfs.get('classroom', None) if dfs else None
                    
                    # Create verification sheets for each section
                    for section in sections:
                        post_mid_reset = post_mid_sections[section].reset_index().rename(columns={'index': 'Time Slot'})
                        if not post_mid_reset.empty:
                            section_label = f"Section {section}" if has_sections else "Timetable"
                            print(f"   Creating Post-Mid Verification sheet for {section_label}...")
                            verification = create_timetable_verification_sheet(
                                post_mid_reset, course_info, classroom_data, semester, branch, section
                            )
                            if not verification.empty:
                                sheet_name = f'Verification_{section}' if has_sections else 'Verification'
                                verification.to_excel(writer, sheet_name=sheet_name, index=False)
                                print(f"   [OK] Created Post-Mid Verification sheet with {len(verification)} entries")
                    
                    # Create room allocation summary
                    if classroom_data is not None and not classroom_data.empty:
                        print(f"   Creating room allocation summary...")
                        all_sections_data = [post_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                        room_summary = create_room_allocation_summary_verification(
                            all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame(), classroom_data
                        )
                        if not room_summary.empty:
                            room_summary.to_excel(writer, sheet_name='Room_Allocation', index=False)
                            print(f"   [OK] Created Room_Allocation sheet with {len(room_summary)} rooms")
                    
                    # Create LTPSC compliance summary
                    print(f"   Creating LTPSC compliance summary...")
                    all_sections_data = [post_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                    ltpsc_summary = create_ltpsc_compliance_summary(
                        dfs, semester, branch, all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame()
                    )
                    if not ltpsc_summary.empty:
                        ltpsc_summary.to_excel(writer, sheet_name='LTPSC_Compliance', index=False)
                        print(f"   [OK] Created LTPSC_Compliance sheet with {len(ltpsc_summary)} courses")
                    
                    # Create executive summary
                    print(f"   Creating executive summary...")
                    all_sections_data = [post_mid_sections[s].reset_index().rename(columns={'index': 'Time Slot'}) for s in sections]
                    exec_summary = create_executive_summary(
                        dfs, semester, branch, all_sections_data[0], all_sections_data[1] if len(all_sections_data) > 1 else pd.DataFrame(), {}
                    )
                    if not exec_summary.empty:
                        exec_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)
                        print(f"   [OK] Created Executive_Summary sheet")
                    
                    # Add course summary
                    post_mid_summary = create_mid_semester_summary(post_mid_courses, 'Post-Mid', semester, branch)
                    post_mid_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
                    
                    # Add distribution logic sheet
                    distribution_sheet = create_mid_semester_distribution_sheet(
                        pre_mid_courses, post_mid_courses, semester, branch
                    )
                    distribution_sheet.to_excel(writer, sheet_name='Distribution_Logic', index=False)
                    
                    # Add combined overview
                    combined_sheet = create_combined_mid_semester_sheet(
                        pre_mid_courses, post_mid_courses, semester, branch
                    )
                    combined_sheet.to_excel(writer, sheet_name='All_Courses_Overview', index=False)
                    
                    print(f"[STATS] Added POST-MID verification sheets for easy inspection")
                    print(f"[OK] POST-MID timetable saved: {post_mid_filename}")
            except Exception as e:
                print(f"[FAIL] Error saving post-mid timetable: {e}")
                traceback.print_exc()
        else:
            print(f"[WARN] Cannot save post-mid timetable - Not all sections generated successfully")
        
        return {
            'pre_mid_success': pre_mid_all_valid,
            'post_mid_success': post_mid_all_valid,
            'pre_mid_filename': pre_mid_filename if pre_mid_all_valid else None,
            'post_mid_filename': post_mid_filename if post_mid_all_valid else None
        }
        
    except Exception as e:
        print(f"[FAIL] Error generating mid-semester timetables: {e}")
        traceback.print_exc()
        return {
            'pre_mid_success': False,
            'post_mid_success': False,
            'pre_mid_filename': None,
            'post_mid_filename': None
        }

def create_mid_semester_summary(courses_df, schedule_type, semester, branch=None):
    """Create summary sheet for mid-semester timetable
    
    Shows inclusion reasoning based on Half Semester and Post mid-sem values
    """
    if courses_df.empty:
        return pd.DataFrame()
    
    summary_data = []
    
    for _, course in courses_df.iterrows():
        ltpsc_str = course.get('LTPSC', '')
        ltpsc = parse_ltpsc(ltpsc_str)
        credits = course.get('Credits', 'N/A')
        half_sem = course.get('Half Semester (Yes/No)', 'N/A')
        post_mid = course.get('Post mid-sem', 'N/A')
        
        # Add reasoning based on new logic
        if str(half_sem).upper().strip() == 'YES':
            reasoning = "Half Semester = Yes (scheduled in both pre-mid and post-mid)"
        elif str(half_sem).upper().strip() == 'NO':
            if str(post_mid).upper().strip() == 'NO':
                reasoning = f"Half Semester = No, Post mid-sem = No (pre-mid only)"
            elif str(post_mid).upper().strip() == 'YES':
                reasoning = f"Half Semester = No, Post mid-sem = Yes (post-mid only)"
            else:
                reasoning = "Half Semester = No (check Post mid-sem value)"
        else:
            reasoning = f"Check Half Semester ({half_sem}) and Post mid-sem ({post_mid}) values"
        
        summary_data.append({
            'Schedule Type': schedule_type,
            'Semester': semester,
            'Branch': branch or 'All',
            'Course Code': course['Course Code'],
            'Course Name': course.get('Course Name', 'Unknown'),
            'Credits': credits,
            'LTPSC': ltpsc_str,
            'Lectures/Week': ltpsc['L'],
            'Tutorials/Week': ltpsc['T'],
            'Practicals/Week': ltpsc['P'],
            'Faculty': course.get('Faculty', 'Unknown'),
            'Department': course.get('Department', 'Unknown'),
            'Half Semester': half_sem,
            'Post mid-sem': post_mid,
            'Elective': course.get('Elective (Yes/No)', 'N/A'),
            'Inclusion Reasoning': reasoning
        })
    
    return pd.DataFrame(summary_data)

def create_mid_semester_distribution_sheet(pre_mid_courses, post_mid_courses, semester, branch=None):
    """Create a sheet showing the complete distribution logic"""
    if pre_mid_courses.empty and post_mid_courses.empty:
        return pd.DataFrame()
    
    # Get all courses from the original data to see complete picture
    dfs = load_all_data(force_reload=False)
    if dfs is None or 'course' not in dfs:
        return pd.DataFrame()
    
    # Get all courses for this semester and branch
    all_courses = dfs['course'][
        dfs['course']['Semester'].astype(str).str.strip() == str(semester)
    ].copy()
    
    if branch and 'Department' in all_courses.columns:
        normalized_branch = branch.strip()
        dept_match = all_courses['Department'].astype(str).str.strip() == normalized_branch
        common_courses = all_courses.get('Comman', '').astype(str).str.upper() == 'YES'
        all_courses = all_courses[dept_match | common_courses].copy()
    
    all_courses['Credits'] = pd.to_numeric(all_courses['Credits'], errors='coerce')
    
    distribution_data = []
    
    for _, course in all_courses.iterrows():
        course_code = course['Course Code']
        credits = course.get('Credits', 'N/A')
        post_mid = course.get('Post mid-sem', 'N/A')
        
        # Check if in pre-mid
        in_pre_mid = course_code in pre_mid_courses['Course Code'].values if not pre_mid_courses.empty else False
        
        # Check if in post-mid
        in_post_mid = course_code in post_mid_courses['Course Code'].values if not post_mid_courses.empty else False
        
        # Determine scheduling logic
        if in_pre_mid and in_post_mid:
            schedule_type = "Both (non-2-credit with Post mid-sem='No')"
        elif in_pre_mid:
            schedule_type = "Pre-Mid Only"
        elif in_post_mid:
            schedule_type = "Post-Mid Only"
        else:
            schedule_type = "Not Scheduled (Error)"
        
        # Detailed reasoning
        if credits == 2:
            if str(post_mid).upper() == 'NO':
                logic = "2-credit + Post mid-sem='No' → Pre-Mid Only"
            else:  # 'YES'
                logic = "2-credit + Post mid-sem='Yes' → Post-Mid Only"
        else:  # non-2-credit
            if str(post_mid).upper() == 'NO':
                logic = f"{credits}-credit + Post mid-sem='No' → BOTH Pre-Mid & Post-Mid"
            else:  # 'YES'
                logic = f"{credits}-credit + Post mid-sem='Yes' → Post-Mid Only"
        
        distribution_data.append({
            'Course Code': course_code,
            'Course Name': course.get('Course Name', 'Unknown'),
            'Credits': credits,
            'Post mid-sem': post_mid,
            'Scheduled In': schedule_type,
            'Logic Applied': logic,
            'In Pre-Mid': 'Yes' if in_pre_mid else 'No',
            'In Post-Mid': 'Yes' if in_post_mid else 'No',
            'Semester': semester,
            'Branch': branch or 'All'
        })
    
    return pd.DataFrame(distribution_data)

def create_combined_mid_semester_sheet(pre_mid_courses, post_mid_courses, semester, branch=None):
    """Create combined sheet showing all courses with their mid-semester status"""
    if pre_mid_courses.empty and post_mid_courses.empty:
        return pd.DataFrame()
    
    combined_data = []
    
    # Add pre-mid courses
    for _, course in pre_mid_courses.iterrows():
        combined_data.append({
            'Course Code': course['Course Code'],
            'Course Name': course.get('Course Name', 'Unknown'),
            'Credits': course.get('Credits', 'N/A'),
            'Semester': semester,
            'Branch': branch or 'All',
            'Schedule Type': 'Pre-Mid',
            'Post mid-sem': 'No',
            'Half Semester': course.get('Half Semester (Yes/No)', 'N/A'),
            'Department': course.get('Department', 'Unknown'),
            'Faculty': course.get('Faculty', 'Unknown')
        })
    
    # Add post-mid courses
    for _, course in post_mid_courses.iterrows():
        combined_data.append({
            'Course Code': course['Course Code'],
            'Course Name': course.get('Course Name', 'Unknown'),
            'Credits': course.get('Credits', 'N/A'),
            'Semester': semester,
            'Branch': branch or 'All',
            'Schedule Type': 'Post-Mid',
            'Post mid-sem': 'Yes' if course.get('Credits') == 2 else 'Any',
            'Half Semester': course.get('Half Semester (Yes/No)', 'N/A'),
            'Department': course.get('Department', 'Unknown'),
            'Faculty': course.get('Faculty', 'Unknown')
        })
    
    return pd.DataFrame(combined_data)

def check_for_classroom_allocation(schedule_df):
    """Check if classroom allocation was successful by looking for [Room] pattern"""
    for col in schedule_df.columns:
        for val in schedule_df[col]:
            if isinstance(val, str) and '[' in val and ']' in val:
                return True
    return False

def calculate_classroom_usage_for_exams(exam_schedule_df):
    """Calculate classroom usage statistics for exams"""
    used_rooms = set()
    total_sessions = 0
    
    for _, exam in exam_schedule_df.iterrows():
        if exam['status'] == 'Scheduled' and 'classroom' in exam:
            classrooms = str(exam['classroom']).split(', ')
            used_rooms.update(classrooms)
            total_sessions += len(classrooms)
    
    return {
        'used_rooms': len(used_rooms),
        'total_sessions': total_sessions,
        'rooms_list': list(used_rooms)
    }

def calculate_timetable_classroom_usage(timetable_schedules):
    """Calculate classroom usage statistics for timetables"""
    used_rooms = set()
    total_sessions = 0
    
    for schedule in timetable_schedules:
        for day in schedule.columns:
            for time_slot in schedule.index:
                cell_value = str(schedule.loc[time_slot, day])
                # Extract room numbers from format "Course [Room]"
                if '[' in cell_value and ']' in cell_value:
                    room_match = re.search(r'\[(.*?)\]', cell_value)
                    if room_match:
                        room_number = room_match.group(1)
                        used_rooms.add(room_number)
                        total_sessions += 1
    
    
    return {
        'used_rooms': len(used_rooms),
        'total_sessions': total_sessions,
        'rooms_list': list(used_rooms)
    }

def create_classroom_allocation_detail(timetable_schedules, classrooms_df):
    """Create detailed classroom allocation information"""
    allocation_data = []
    
    for i, schedule in enumerate(timetable_schedules, 1):
        section = 'A' if i == 1 else 'B'
        
        for day in schedule.columns:
            for time_slot in schedule.index:
                cell_value = str(schedule.loc[time_slot, day])
                
                if cell_value not in ['Free', 'LUNCH BREAK'] and '[' in cell_value and ']' in cell_value:
                    # Extract course and room
                    course_match = re.search(r'^(.*?)\s*\[', cell_value)
                    room_match = re.search(r'\[(.*?)\]', cell_value)
                    
                    if course_match and room_match:
                        course = course_match.group(1).strip()
                        room_number = room_match.group(1)
                        
                        # Get room details
                        room_details = classrooms_df[classrooms_df['Room Number'] == room_number]
                        capacity = room_details['Capacity'].iloc[0] if not room_details.empty else 'Unknown'
                        room_type = room_details['Type'].iloc[0] if not room_details.empty else 'Unknown'
                        
                        allocation_data.append({
                            'Section': section,
                            'Day': day,
                            'Time Slot': time_slot,
                            'Course': course,
                            'Room Number': room_number,
                            'Room Type': room_type,
                            'Capacity': capacity,
                            'Facilities': room_details['Facilities'].iloc[0] if not room_details.empty else 'Unknown'
                        })
    
    return pd.DataFrame(allocation_data)

def create_basket_courses_sheet(basket_allocations):
    """Create detailed sheet showing all courses in each basket"""
    summary_data = []
    
    for basket_name, allocation in basket_allocations.items():
        for course_code in allocation['courses']:
            summary_data.append({
                'Basket Name': basket_name,
                'Course Code': course_code,
                'Lecture Slots': ', '.join([f"{day} {time}" for day, time in allocation['lectures']]),
                'Tutorial Slot': f"{allocation['tutorial'][0]} {allocation['tutorial'][1]}",
                'Total Courses in Basket': len(allocation['courses']),
                'Common for All Branches': 'Yes',
                'Common for Both Sections': 'Yes'
            })
    
    return pd.DataFrame(summary_data)

def export_semester_timetable(dfs, semester, branch=None):
    """Export timetable using basket-based elective allocation with COMMON slots"""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"\n[STATS] Generating BASKET-BASED timetable for Semester {semester}{branch_info}...")
    print(f"[TARGET] Using COMMON elective basket slots across all branches")
    
    try:
        # CRITICAL: Get ALL elective courses for this semester ONCE (without branch filter)
        # This ensures COMMON allocation for all branches
        course_baskets_all = separate_courses_by_type(dfs, semester)  # No branch filter
        elective_courses_all = course_baskets_all['elective_courses']
        
        print(f"[TARGET] Elective courses found for semester {semester} (COMMON for ALL branches): {len(elective_courses_all)}")
        if not elective_courses_all.empty:
            print("   Common elective courses:", elective_courses_all['Course Code'].tolist())
            # Show basket distribution
            basket_counts = elective_courses_all['Basket'].value_counts()
            print("   Basket distribution:")
            for basket, count in basket_counts.items():
                courses = elective_courses_all[elective_courses_all['Basket'] == basket]['Course Code'].tolist()
                print(f"      {basket}: {count} courses - {courses}")
        
        # Allocate electives by baskets - COMMON for all branches
        elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)
        
        print(f"   📅 COMMON BASKET ALLOCATIONS for Semester {semester}:")
        for basket_name, allocation in basket_allocations.items():
            print(f"      {basket_name}:")
            print(f"         Lectures: {allocation['lectures']}")
            print(f"         Tutorial: {allocation['tutorial']}")
            print(f"         Courses: {allocation['courses']}")
        
        # Generate schedules for both sections with IDENTICAL basket slots
        section_a = generate_section_schedule_with_elective_baskets(dfs, semester, 'A', elective_allocations, branch)
        section_b = generate_section_schedule_with_elective_baskets(dfs, semester, 'B', elective_allocations, branch)
        
        if section_a is None or section_b is None:
            return False

        # ALLOCATE CLASSROOMS for both sections
        course_info = get_course_info(dfs) if dfs else {}
        classroom_data = dfs.get('classroom')
        
        if classroom_data is not None and not classroom_data.empty:
            print("[SCHOOL] Allocating classrooms...")
            section_a_with_rooms = allocate_classrooms_for_timetable(
                section_a, classroom_data, course_info, semester, branch, 'A'
            )
            section_b_with_rooms = allocate_classrooms_for_timetable(
                section_b, classroom_data, course_info, semester, branch, 'B'
            )
        else:
            section_a_with_rooms = section_a
            section_b_with_rooms = section_b
            print("[WARN]  No classroom data available for allocation")

        # Create filename
        if branch:
            filename = f"sem{semester}_{branch}_timetable.xlsx"
        else:
            filename = f"sem{semester}_timetable.xlsx"
            
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a_with_rooms.to_excel(writer, sheet_name='Section_A')
            section_b_with_rooms.to_excel(writer, sheet_name='Section_B')
            
            # NEW: Add comprehensive statistics sheets
            if dfs:
                course_info = get_course_info(dfs)
                classroom_df = dfs.get('classroom', pd.DataFrame())
                
                # Statistics for Section A
                print(f"[STATS] Creating statistics sheet for Section A...")
                stats_a = create_timetable_statistics_sheet(
                    section_a_with_rooms, course_info, classroom_df, semester, branch, 'A'
                )
                stats_a.to_excel(writer, sheet_name='Statistics_A', index=False)
                
                # Statistics for Section B
                print(f"[STATS] Creating statistics sheet for Section B...")
                stats_b = create_timetable_statistics_sheet(
                    section_b_with_rooms, course_info, classroom_df, semester, branch, 'B'
                )
                stats_b.to_excel(writer, sheet_name='Statistics_B', index=False)
                
                # Room allocation summary
                if not classroom_df.empty:
                    print(f"[STATS] Creating room allocation summary...")
                    room_summary = create_room_allocation_summary(section_a_with_rooms, classroom_df)
                    room_summary.to_excel(writer, sheet_name='Room_Summary', index=False)
                
                # Comprehensive executive summary
                print(f"[STATS] Creating executive summary...")
                exec_summary = create_comprehensive_summary(
                    dfs, semester, branch, section_a_with_rooms, section_b_with_rooms, basket_allocations
                )
                exec_summary.to_excel(writer, sheet_name='Executive_Summary', index=False)
            
            # Add basket allocation summary
            basket_summary = create_basket_summary(basket_allocations, semester, branch)
            basket_summary.to_excel(writer, sheet_name='Basket_Allocation', index=False)
            
            # Add course summary
            course_summary = create_course_summary(dfs, semester, branch)
            if not course_summary.empty:
                course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
            
            # Add basket course details
            basket_courses_sheet = create_basket_courses_sheet(basket_allocations)
            basket_courses_sheet.to_excel(writer, sheet_name='Basket_Courses', index=False)
            
            # Add branch-specific info sheet
            branch_info_sheet = create_branch_info_sheet(dfs, semester, branch)
            if not branch_info_sheet.empty:
                branch_info_sheet.to_excel(writer, sheet_name='Branch_Info', index=False)
        
        print(f"[OK] Basket-based timetable saved: {filename}")
        print(f"[STATS] Added comprehensive statistics sheets for easy verification")
        return True
        
    except Exception as e:
        print(f"[FAIL] Error generating basket-based timetable: {e}")
        traceback.print_exc()
        return False

def create_basket_summary(basket_allocations, semester, branch=None):
    """Create a summary of basket allocations"""
    summary_data = []
    
    for basket_name, allocation in basket_allocations.items():
        day, time_slot = allocation['slot']
        courses = allocation['courses']
        
        summary_data.append({
            'Basket Name': basket_name,
            'Day': day,
            'Time Slot': time_slot,
            'Courses in Basket': ', '.join(courses),
            'Number of Courses': len(courses),
            'Sections': 'A & B (Common)',
            'Branches': branch if branch else 'ALL',
            'Semester': f'Semester {semester}'
        })
    
    return pd.DataFrame(summary_data)

def create_branch_info_sheet(dfs, semester, branch):
    """Create a sheet showing department-specific course information"""
    if 'course' not in dfs:
        return pd.DataFrame()
    
    # Get all courses for the semester
    sem_courses = dfs['course'][
        dfs['course']['Semester'].astype(str).str.strip() == str(semester)
    ].copy()
    
    if sem_courses.empty:
        return pd.DataFrame()
    
    # Separate department-specific and common courses
    dept_specific_courses = sem_courses[
        (sem_courses['Department'] == branch) & 
        (sem_courses['Elective (Yes/No)'].str.upper() != 'YES')
    ].copy()
    
    common_elective_courses = sem_courses[
        sem_courses['Elective (Yes/No)'].str.upper() == 'YES'
    ].copy()
    
    # Create summary
    summary_data = []
    
    # Add department-specific courses
    for _, course in dept_specific_courses.iterrows():
        summary_data.append({
            'Course Type': 'Department-Specific Core',
            'Course Code': course['Course Code'],
            'Course Name': course.get('Course Name', 'Unknown'),
            'Department': course.get('Department', 'Unknown'),
            'LTPSC': course.get('LTPSC', 'N/A'),
            'Credits': course.get('Credits', 'N/A'),
            'Instructor': course.get('Instructor', 'Unknown')
        })
    
    # Add common elective courses
    for _, course in common_elective_courses.iterrows():
        summary_data.append({
            'Course Type': 'Common Elective',
            'Course Code': course['Course Code'],
            'Course Name': course.get('Course Name', 'Unknown'),
            'Department': 'ALL (Common)',
            'LTPSC': course.get('LTPSC', 'N/A'),
            'Credits': course.get('Credits', 'N/A'),
            'Instructor': course.get('Instructor', 'Unknown')
        })
    
    return pd.DataFrame(summary_data)

def clean_table_html(html):
    """Clean and format the HTML table for better display"""
    html = html.replace('border="1"', '')
    html = html.replace('class="dataframe"', 'class="timetable-table"')
    html = html.replace('<thead>', '<thead class="timetable-head">')
    html = html.replace('<tbody>', '<tbody class="timetable-body">')
    return html

def extract_unique_courses(df):
    """Extract unique course codes from a timetable dataframe"""
    courses = set()
    for col in df.columns:
        for value in df[col]:
            if isinstance(value, str) and value not in ['Free', 'LUNCH BREAK']:
                # Handle both regular courses and elective/tutorial marked courses
                clean_value = value.replace(' (Elective)', '').replace(' (Tutorial)', '')
                course_code = extract_course_code(clean_value)
                if course_code:
                    courses.add(course_code)
    return list(courses)

def extract_unique_courses_with_baskets(df, elective_allocations=None):
    """Extract unique course codes and basket names from a timetable dataframe"""
    courses = set()
    baskets = set()
    
    for col in df.columns:
        for value in df[col]:
            if isinstance(value, str) and value not in ['Free', 'LUNCH BREAK']:
                # Extract clean course code (remove classroom info, tutorial markers, etc.)
                clean_value = value
                
                # Remove classroom allocation info
                if '[' in clean_value:
                    clean_value = clean_value.split('[')[0].strip()
                
                # Remove tutorial marker
                clean_value = clean_value.replace(' (Tutorial)', '')
                
                # Check if this is a basket entry
                basket_names = ['ELECTIVE_B1', 'ELECTIVE_B2', 'ELECTIVE_B3', 'ELECTIVE_B4', 'ELECTIVE_B5', 'ELECTIVE_B6', 'ELECTIVE_B7', 'HSS_B1', 'HSS_B2']
                is_basket = any(basket in clean_value for basket in basket_names)
                
                if is_basket:
                    # Extract basket name - get the actual basket name from the string
                    for basket_name in basket_names:
                        if basket_name in clean_value:
                            baskets.add(basket_name)
                            break
                    
                    # ADDED: Also add all courses from this basket to the courses set
                    if elective_allocations:
                        for course_code, allocation in elective_allocations.items():
                            if allocation and allocation.get('basket_name') in baskets:
                                courses.add(course_code)
                else:
                    # Handle regular courses - extract course code
                    course_code = extract_course_code(clean_value)
                    if course_code:
                        courses.add(course_code)
    
    return list(courses), list(baskets)

def extract_course_code(text):
    """Extract course code from text"""
    import re
    course_pattern = r'[A-Z]{2,3}\d{3}'
    match = re.search(course_pattern, text)
    return match.group(0) if match else None

def generate_course_colors(courses, course_info):
    """Generate unique colors for each course, with different shades for core vs elective"""
    colors = [
        '#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FECA57', '#FF9FF3',
        '#54A0FF', '#5F27CD', '#00D2D3', '#FF9F43', '#10AC84', '#EE5A24',
        '#0984E3', '#00B894', '#6C5CE7', '#E84393', '#FDCB6E', '#00CEC9',
        '#A29BFE', '#FD79A8', '#FDCB6E', '#E17055', '#00B894', '#0984E3'
    ]
    
    elective_colors = [
        '#FF9F93', '#88E0D0', '#85D1E5', '#C8E6C9', '#FFE082', '#FFC8DD',
        '#90CAF9', '#B39DDB', '#80DEEA', '#FFCC80', '#81C784', '#FF8A65',
        '#64B5F6', '#4DB6AC', '#9575CD', '#F48FB1', '#FFF176', '#4DD0E1'
    ]
    
    course_colors = {}
    core_courses = []
    elective_courses = []
    
    # Separate core and elective courses
    for course in sorted(courses):
        if course in course_info:
            if course_info[course].get('is_elective', False):
                elective_courses.append(course)
            else:
                core_courses.append(course)
        else:
            core_courses.append(course)  # Default to core if info not available
    
    # Assign colors - core courses get primary colors, electives get lighter shades
    for i, course in enumerate(core_courses):
        course_colors[course] = colors[i % len(colors)]
    
    for i, course in enumerate(elective_courses):
        course_colors[course] = elective_colors[i % len(elective_colors)]
    
    return course_colors

def allocate_classrooms_for_timetable(schedule_df, classrooms_df, course_info, semester, branch, section):
    """Allocate classrooms to timetable sessions with proper tracking across all timetables"""
    print(f"[SCHOOL] Allocating classrooms for {branch} Semester {semester} Section {section}...")
    
    if classrooms_df is None or classrooms_df.empty:
        print("   [WARN] No classroom data available")
        return schedule_df
    
    # Initialize global tracker if not exists
    global _CLASSROOM_USAGE_TRACKER
    if not _CLASSROOM_USAGE_TRACKER:
        initialize_classroom_usage_tracker()
    
    # Filter available classrooms (exclude labs, recreation, library, etc.)
    available_classrooms = classrooms_df[
        (classrooms_df['Type'].str.contains('classroom', case=False, na=False)) |
        (classrooms_df['Type'].str.contains('auditorium', case=False, na=False))
    ].copy()
    
    # Filter lab rooms separately (rooms starting with "L")
    available_lab_rooms = classrooms_df[
        classrooms_df['Room Number'].astype(str).str.startswith('L', na=False)
    ].copy()
    
    # Convert capacity to numeric, handle 'nil' values
    available_classrooms['Capacity'] = pd.to_numeric(available_classrooms['Capacity'], errors='coerce')
    available_classrooms = available_classrooms.dropna(subset=['Capacity'])
    
    available_lab_rooms['Capacity'] = pd.to_numeric(available_lab_rooms['Capacity'], errors='coerce')
    available_lab_rooms = available_lab_rooms.dropna(subset=['Capacity'])
    
    if available_classrooms.empty:
        print("   [WARN] No suitable classrooms found after filtering")
        return schedule_df
    
    print(f"   Available classrooms: {len(available_classrooms)}")
    print(f"   Available lab rooms: {len(available_lab_rooms)}")
    
    non_lab_mask = ~available_classrooms['Room Number'].astype(str).str.startswith('L', na=False)
    primary_classrooms = available_classrooms[non_lab_mask].copy()
    fallback_lab_classrooms = available_classrooms[~non_lab_mask].copy()
    
    print(f"   Primary classrooms (non-L prefix): {len(primary_classrooms)}")
    print(f"   Fallback classrooms (L prefix): {len(fallback_lab_classrooms)}")
    
    # Create a copy of schedule with classroom allocation
    schedule_with_rooms = schedule_df.copy()
    course_preferred_classrooms = {}

    def room_available(room_number, day_key, slot_key):
        return room_number not in _CLASSROOM_USAGE_TRACKER.get(day_key, {}).get(slot_key, set())

    def reserve_room(room_number, day_key, slot_key):
        if day_key not in _CLASSROOM_USAGE_TRACKER:
            _CLASSROOM_USAGE_TRACKER[day_key] = {}
        if slot_key not in _CLASSROOM_USAGE_TRACKER[day_key]:
            _CLASSROOM_USAGE_TRACKER[day_key][slot_key] = set()
        _CLASSROOM_USAGE_TRACKER[day_key][slot_key].add(room_number)

    def room_available_for_lab_pair(room_number, day_key, slot_one, slot_two):
        return room_available(room_number, day_key, slot_one) and room_available(room_number, day_key, slot_two)
    
    def allocate_regular_classroom(enrollment_value, day_key, slot_key):
        classroom_choice = None
        if not primary_classrooms.empty:
            classroom_choice = find_suitable_classroom_with_tracking(
                primary_classrooms, enrollment_value, day_key, slot_key, _CLASSROOM_USAGE_TRACKER
            )
        if classroom_choice is None and not fallback_lab_classrooms.empty:
            classroom_choice = find_suitable_classroom_with_tracking(
                fallback_lab_classrooms, enrollment_value, day_key, slot_key, _CLASSROOM_USAGE_TRACKER
            )
            if classroom_choice:
                print(f"         [WARN] Falling back to lab-prefixed room {classroom_choice} for {day_key} {slot_key}")
        return classroom_choice
    
    # Estimate student numbers for courses
    course_enrollment = estimate_course_enrollment(course_info)
    
    # Track allocations for this specific timetable
    timetable_key = f"{branch}_sem{semester}_sec{section}"
    if timetable_key not in _TIMETABLE_CLASSROOM_ALLOCATIONS:
        _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key] = {}
    
    # Track which lab slots have been processed to avoid double allocation
    processed_lab_slots = set()
    
    # Define lab slot pairs (consecutive slots that form a 2-hour lab)
    lab_slot_pairs = {
        '13:00-14:30': '14:30-15:30',  # First slot -> Second slot
        '15:30-17:00': '17:00-18:00',  # First slot -> Second slot
    }
    
    # Allocate classrooms for each time slot
    allocation_count = 0
    for day in schedule_df.columns:
        for time_slot in schedule_df.index:
            course_value = schedule_df.loc[time_slot, day]
            
            # Skip free slots, lunch breaks
            if course_value in ['Free', 'LUNCH BREAK']:
                continue
            
            # Skip if this slot was already processed as part of a lab pair
            if (day, time_slot) in processed_lab_slots:
                continue
            
            # Handle both regular courses and basket entries
            if isinstance(course_value, str):
                # Check if this is a lab slot
                is_lab = ' (Lab)' in course_value
                
                # For basket entries, use a standard enrollment
                if any(basket in course_value for basket in ['ELECTIVE_B', 'HSS_B', 'PROF_B', 'OE_B']):
                    enrollment = 40  # Standard enrollment for elective baskets
                    course_display = course_value
                else:
                    # Regular course - extract clean course code (handle Tutorial, Lab suffixes)
                    clean_course_code = course_value.replace(' (Tutorial)', '').replace(' (Lab)', '')
                    enrollment = course_enrollment.get(clean_course_code, 50)
                    course_display = course_value
            else:
                continue
            
            course_key = course_display
            preferred_classroom = course_preferred_classrooms.get(course_key)
            
            # If this is a lab and it's the first slot of a lab pair, find a classroom available for BOTH slots
            if is_lab and time_slot in lab_slot_pairs:
                second_slot = lab_slot_pairs[time_slot]
                second_course_value = schedule_df.loc[second_slot, day]
                
                # Check if second slot also has the same lab course
                if isinstance(second_course_value, str) and second_course_value == course_value:
                    if preferred_classroom and room_available_for_lab_pair(preferred_classroom, day, time_slot, second_slot):
                        reserve_room(preferred_classroom, day, time_slot)
                        reserve_room(preferred_classroom, day, second_slot)
                        schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{preferred_classroom}]"
                        schedule_with_rooms.loc[second_slot, day] = f"{course_display} [{preferred_classroom}]"
                        
                        allocation_key_1 = f"{day}_{time_slot}"
                        allocation_key_2 = f"{day}_{second_slot}"
                        _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key_1] = {
                            'course': course_display,
                            'classroom': preferred_classroom,
                            'enrollment': enrollment
                        }
                        _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key_2] = {
                            'course': course_display,
                            'classroom': preferred_classroom,
                            'enrollment': enrollment
                        }
                        
                        processed_lab_slots.add((day, second_slot))
                        
                        allocation_count += 2
                        print(f"      🔁 Reusing {preferred_classroom} for lab pair {day} {time_slot} & {second_slot} ({enrollment} students)")
                    else:
                        # Find a LAB ROOM (starting with "L") that's available for BOTH slots
                        # Labs must use lab rooms, not regular classrooms
                        suitable_classroom = find_suitable_classroom_for_lab_pair(
                            available_lab_rooms, enrollment, day, time_slot, second_slot, _CLASSROOM_USAGE_TRACKER
                        )
                        
                        if suitable_classroom:
                            # Allocate the same classroom to BOTH slots
                            # Mark both slots as used in the tracker
                            reserve_room(suitable_classroom, day, time_slot)
                            reserve_room(suitable_classroom, day, second_slot)
                            
                            # Update schedule with same classroom for both slots
                            schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{suitable_classroom}]"
                            schedule_with_rooms.loc[second_slot, day] = f"{course_display} [{suitable_classroom}]"
                            if course_key not in course_preferred_classrooms:
                                course_preferred_classrooms[course_key] = suitable_classroom
                            
                            # Track both allocations
                            allocation_key_1 = f"{day}_{time_slot}"
                            allocation_key_2 = f"{day}_{second_slot}"
                            _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key_1] = {
                                'course': course_display,
                                'classroom': suitable_classroom,
                                'enrollment': enrollment
                            }
                            _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key_2] = {
                                'course': course_display,
                                'classroom': suitable_classroom,
                                'enrollment': enrollment
                            }
                            
                            # Mark second slot as processed
                            processed_lab_slots.add((day, second_slot))
                            
                            allocation_count += 2
                            print(f"      [OK] {day} {time_slot} & {second_slot}: {course_display} → {suitable_classroom} ({enrollment} students) [Lab pair - same room]")
                        else:
                            print(f"      [WARN]  {day} {time_slot}: No classroom available for lab pair ({time_slot} & {second_slot})")
                else:
                    # Not a proper lab pair, treat as regular course
                    if preferred_classroom and room_available(preferred_classroom, day, time_slot):
                        reserve_room(preferred_classroom, day, time_slot)
                        schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{preferred_classroom}]"
                        allocation_key = f"{day}_{time_slot}"
                        _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key] = {
                            'course': course_display,
                            'classroom': preferred_classroom,
                            'enrollment': enrollment
                        }
                        allocation_count += 1
                        print(f"      🔁 Reusing {preferred_classroom} for {course_display} on {day} {time_slot} ({enrollment} students)")
                    else:
                        suitable_classroom = allocate_regular_classroom(enrollment, day, time_slot)
                        if suitable_classroom:
                            schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{suitable_classroom}]"
                            allocation_key = f"{day}_{time_slot}"
                            _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key] = {
                                'course': course_display,
                                'classroom': suitable_classroom,
                                'enrollment': enrollment
                            }
                            if course_key not in course_preferred_classrooms:
                                course_preferred_classrooms[course_key] = suitable_classroom
                            allocation_count += 1
                            print(f"      [OK] {day} {time_slot}: {course_display} → {suitable_classroom} ({enrollment} students)")
                        else:
                            print(f"      [WARN]  {day} {time_slot}: No classroom available for {course_display} ({enrollment} students)")
            else:
                # Regular course (not a lab pair) - find classroom normally
                if preferred_classroom and room_available(preferred_classroom, day, time_slot):
                    reserve_room(preferred_classroom, day, time_slot)
                    schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{preferred_classroom}]"
                    
                    allocation_key = f"{day}_{time_slot}"
                    _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key] = {
                        'course': course_display,
                        'classroom': preferred_classroom,
                        'enrollment': enrollment
                    }
                    
                    allocation_count += 1
                    print(f"      🔁 Reusing {preferred_classroom} for {course_display} on {day} {time_slot} ({enrollment} students)")
                else:
                    suitable_classroom = allocate_regular_classroom(enrollment, day, time_slot)
                    
                    if suitable_classroom:
                        # Update schedule with classroom in format "Course [Room]"
                        schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{suitable_classroom}]"
                        
                        # Track this allocation
                        allocation_key = f"{day}_{time_slot}"
                        _TIMETABLE_CLASSROOM_ALLOCATIONS[timetable_key][allocation_key] = {
                            'course': course_display,
                            'classroom': suitable_classroom,
                            'enrollment': enrollment
                        }
                        if course_key not in course_preferred_classrooms:
                            course_preferred_classrooms[course_key] = suitable_classroom
                        
                        allocation_count += 1
                        print(f"      [OK] {day} {time_slot}: {course_display} → {suitable_classroom} ({enrollment} students)")
                    else:
                        print(f"      [WARN]  {day} {time_slot}: No classroom available for {course_display} ({enrollment} students)")
    
    print(f"   [SCHOOL] Total classroom allocations: {allocation_count}")
    return schedule_with_rooms

def find_suitable_classroom_for_lab_pair(lab_rooms_df, enrollment, day, time_slot1, time_slot2, classroom_usage_tracker):
    """Find a suitable LAB ROOM (starting with "L") that's available for BOTH slots of a lab pair"""
    if lab_rooms_df.empty:
        print(f"         [WARN] No lab rooms available for lab pair")
        return None
    
    # Ensure we only use lab rooms (rooms starting with "L")
    lab_rooms_df = lab_rooms_df[
        lab_rooms_df['Room Number'].astype(str).str.startswith('L', na=False)
    ].copy()
    
    if lab_rooms_df.empty:
        print(f"         [WARN] No lab rooms (starting with 'L') found")
        return None
    
    # Filter lab rooms that can accommodate the enrollment
    suitable_rooms = lab_rooms_df[lab_rooms_df['Capacity'] >= enrollment].copy()
    
    if suitable_rooms.empty:
        # If no room can accommodate, find the largest available lab room
        suitable_rooms = lab_rooms_df.nlargest(3, 'Capacity')
        if suitable_rooms.empty:
            print(f"         [WARN] No lab rooms available for {enrollment} students")
            return None
        print(f"         [WARN] Using largest available lab room for {enrollment} students (lab pair)")
    
    # Sort by capacity (prefer smallest adequate room first)
    suitable_rooms = suitable_rooms.sort_values('Capacity')
    
    # Shuffle to ensure variety
    suitable_rooms = suitable_rooms.sample(frac=1).reset_index(drop=True)
    
    # Check availability in global tracker for BOTH slots
    for _, room in suitable_rooms.iterrows():
        room_number = room['Room Number']
        
        # Check if room is available for BOTH time slots
        available_slot1 = (day in classroom_usage_tracker and 
                          time_slot1 in classroom_usage_tracker[day] and
                          room_number not in classroom_usage_tracker[day][time_slot1])
        
        available_slot2 = (day in classroom_usage_tracker and 
                          time_slot2 in classroom_usage_tracker[day] and
                          room_number not in classroom_usage_tracker[day][time_slot2])
        
        if available_slot1 and available_slot2:
            # Mark room as used in global tracker for BOTH slots
            if day not in classroom_usage_tracker:
                classroom_usage_tracker[day] = {}
            if time_slot1 not in classroom_usage_tracker[day]:
                classroom_usage_tracker[day][time_slot1] = set()
            if time_slot2 not in classroom_usage_tracker[day]:
                classroom_usage_tracker[day][time_slot2] = set()
            
            classroom_usage_tracker[day][time_slot1].add(room_number)
            classroom_usage_tracker[day][time_slot2].add(room_number)
            print(f"         [PIN] Allocated {room_number} for lab pair {day} {time_slot1} & {time_slot2} (Capacity: {room['Capacity']})")
            return room_number
    
    # If all suitable rooms are booked, try larger lab rooms
    larger_rooms = lab_rooms_df[lab_rooms_df['Capacity'] > enrollment].copy()
    if not larger_rooms.empty:
        larger_rooms = larger_rooms.sort_values('Capacity')
        # Shuffle to ensure variety
        larger_rooms = larger_rooms.sample(frac=1).reset_index(drop=True)
        for _, room in larger_rooms.iterrows():
            room_number = room['Room Number']
            
            available_slot1 = (day in classroom_usage_tracker and 
                              time_slot1 in classroom_usage_tracker[day] and
                              room_number not in classroom_usage_tracker[day][time_slot1])
            
            available_slot2 = (day in classroom_usage_tracker and 
                              time_slot2 in classroom_usage_tracker[day] and
                              room_number not in classroom_usage_tracker[day][time_slot2])
            
            if available_slot1 and available_slot2:
                if day not in classroom_usage_tracker:
                    classroom_usage_tracker[day] = {}
                if time_slot1 not in classroom_usage_tracker[day]:
                    classroom_usage_tracker[day][time_slot1] = set()
                if time_slot2 not in classroom_usage_tracker[day]:
                    classroom_usage_tracker[day][time_slot2] = set()
                
                classroom_usage_tracker[day][time_slot1].add(room_number)
                classroom_usage_tracker[day][time_slot2].add(room_number)
                print(f"         [RESET] Using larger room {room_number} for lab pair {day} {time_slot1} & {time_slot2} (Capacity: {room['Capacity']})")
                return room_number
    
    return None

def find_suitable_classroom_with_tracking(classrooms_df, enrollment, day, time_slot, classroom_usage_tracker):
    """Find a suitable classroom based on capacity and availability with global tracking"""
    if classrooms_df.empty:
        return None
    
    # Filter classrooms that can accommodate the enrollment
    suitable_rooms = classrooms_df[classrooms_df['Capacity'] >= enrollment].copy()
    
    if suitable_rooms.empty:
        # If no room can accommodate, find the largest available
        suitable_rooms = classrooms_df.nlargest(3, 'Capacity')  # Get top 3 largest
        if suitable_rooms.empty:
            return None
        print(f"         [WARN] Using largest available room for {enrollment} students")
    
    # Sort by capacity (prefer smallest adequate room first to preserve larger rooms)
    suitable_rooms = suitable_rooms.sort_values('Capacity')
    
    # Shuffle to ensure different classrooms are selected for different slots
    # This prevents always selecting the same first available room
    suitable_rooms = suitable_rooms.sample(frac=1).reset_index(drop=True)
    
    # Check availability in global tracker
    for _, room in suitable_rooms.iterrows():
        room_number = room['Room Number']
        
        # Check if room is already booked at this SPECIFIC day and time_slot in global tracker
        if day in classroom_usage_tracker and time_slot in classroom_usage_tracker[day]:
            if room_number not in classroom_usage_tracker[day][time_slot]:
                # Mark room as used in global tracker for this specific day and time slot
                classroom_usage_tracker[day][time_slot].add(room_number)
                print(f"         [PIN] Allocated {room_number} for {day} {time_slot} (Capacity: {room['Capacity']})")
                return room_number
        else:
            # Initialize if not exists (shouldn't happen, but safety check)
            if day not in classroom_usage_tracker:
                classroom_usage_tracker[day] = {}
            if time_slot not in classroom_usage_tracker[day]:
                classroom_usage_tracker[day][time_slot] = set()
            classroom_usage_tracker[day][time_slot].add(room_number)
            print(f"         [PIN] Allocated {room_number} for {day} {time_slot} (Capacity: {room['Capacity']}) - initialized")
            return room_number
    
    # If all suitable rooms are booked, try larger rooms
    larger_rooms = classrooms_df[classrooms_df['Capacity'] > enrollment].copy()
    if not larger_rooms.empty:
        larger_rooms = larger_rooms.sort_values('Capacity')
        # Shuffle to ensure variety
        larger_rooms = larger_rooms.sample(frac=1).reset_index(drop=True)
        for _, room in larger_rooms.iterrows():
            room_number = room['Room Number']
            if day in classroom_usage_tracker and time_slot in classroom_usage_tracker[day]:
                if room_number not in classroom_usage_tracker[day][time_slot]:
                    classroom_usage_tracker[day][time_slot].add(room_number)
                    print(f"         [RESET] Using larger room {room_number} for {enrollment} students (Capacity: {room['Capacity']})")
                    return room_number
    
    return None

def find_suitable_classroom(classrooms_df, enrollment, day, time_slot, classroom_usage):
    """Find a suitable classroom based on capacity and availability"""
    if classrooms_df.empty:
        return None
    
    # Filter classrooms that can accommodate the enrollment
    suitable_rooms = classrooms_df[classrooms_df['Capacity'] >= enrollment].copy()
    
    if suitable_rooms.empty:
        # If no room can accommodate, find the largest available
        suitable_rooms = classrooms_df.nlargest(1, 'Capacity')
        if suitable_rooms.empty:
            return None
        print(f"         [WARN] Using largest available room for {enrollment} students")
    
    # Sort by capacity (prefer smallest adequate room first)
    suitable_rooms = suitable_rooms.sort_values('Capacity')
    
    # Check availability
    for _, room in suitable_rooms.iterrows():
        room_number = room['Room Number']
        
        # Check if room is already booked at this time
        if room_number not in classroom_usage[day][time_slot]:
            return room_number
    
    return None

def estimate_course_enrollment(course_info):
    """Estimate student enrollment for courses (can be enhanced with actual student data)"""
    enrollment_estimates = {}
    for course_code, info in course_info.items():
        if info.get('is_elective', False):
            # Electives typically have smaller enrollment
            enrollment_estimates[course_code] = 40
        else:
            # Core courses have larger enrollment
            if 'Semester' in course_code:
                # Estimate based on typical semester sizes
                enrollment_estimates[course_code] = 60
            else:
                enrollment_estimates[course_code] = 50
    return enrollment_estimates

def calculate_rooms_needed(enrollment, classrooms_df):
    """Calculate how many rooms are needed for an exam based on enrollment"""
    # Sort classrooms by capacity (descending)
    sorted_rooms = classrooms_df.sort_values('Capacity', ascending=False)
    
    remaining_students = enrollment
    allocated_rooms = []
    
    for _, room in sorted_rooms.iterrows():
        if remaining_students <= 0:
            break
            
        room_capacity = room['Capacity']
        room_number = room['Room Number']
        
        # Use this room
        allocated_rooms.append(room_number)
        remaining_students -= room_capacity
    
    return allocated_rooms

def create_classroom_utilization_report(classrooms_df, timetable_schedules, exam_schedules):
    """Create a comprehensive classroom utilization report"""
    print("[STATS] Generating classroom utilization report...")
    
    utilization_data = []
    
    for _, classroom in classrooms_df.iterrows():
        room_number = classroom['Room Number']
        capacity = classroom['Capacity']
        room_type = classroom['Type']
        
        # Calculate timetable usage
        timetable_usage = calculate_timetable_usage(room_number, timetable_schedules)
        
        # Calculate exam usage
        exam_usage = calculate_exam_usage(room_number, exam_schedules)
        
        utilization_data.append({
            'Room Number': room_number,
            'Type': room_type,
            'Capacity': capacity,
            'Weekly Hours (Timetable)': timetable_usage['weekly_hours'],
            'Daily Avg Hours (Timetable)': timetable_usage['daily_avg_hours'],
            'Exam Sessions': exam_usage['exam_sessions'],
            'Utilization Rate (%)': calculate_utilization_rate(
                timetable_usage['weekly_hours'], exam_usage['exam_sessions']
            ),
            'Facilities': classroom.get('Facilities', 'None')
        })
    
    return pd.DataFrame(utilization_data)

def calculate_timetable_usage(room_number, timetable_schedules):
    """Calculate how much a room is used in timetables"""
    weekly_hours = 0
    
    for schedule in timetable_schedules:
        for day in schedule.columns:
            for time_slot in schedule.index:
                cell_value = str(schedule.loc[time_slot, day])
                if f"[{room_number}]" in cell_value:
                    # Calculate hours from time slot
                    hours = calculate_time_slot_hours(time_slot)
                    weekly_hours += hours
    
    daily_avg_hours = weekly_hours / 5  # 5 days per week
    return {'weekly_hours': weekly_hours, 'daily_avg_hours': daily_avg_hours}

def calculate_exam_usage(room_number, exam_schedules):
    """Calculate how much a room is used for exams"""
    exam_sessions = 0
    
    for schedule in exam_schedules:
        for _, exam in schedule.iterrows():
            if exam['status'] == 'Scheduled':
                classrooms = str(exam.get('classroom', ''))
                if room_number in classrooms:
                    exam_sessions += 1
    
    return {'exam_sessions': exam_sessions}

def calculate_time_slot_hours(time_slot):
    """Calculate duration in hours from time slot string"""
    try:
        if '-' in time_slot:
            start, end = time_slot.split('-')
            start_hour = int(start.split(':')[0]) + int(start.split(':')[1]) / 60
            end_hour = int(end.split(':')[0]) + int(end.split(':')[1]) / 60
            return end_hour - start_hour
    except:
        pass
    return 1.5  # Default fallback

def calculate_utilization_rate(weekly_hours, exam_sessions):
    """Calculate classroom utilization rate"""
    # Assuming 40 hours per week available (8 hours × 5 days)
    max_weekly_hours = 40
    timetable_utilization = (weekly_hours / max_weekly_hours) * 100
    
    # Exam utilization (each exam session counts as 1)
    exam_utilization = min(exam_sessions * 5, 100)  # Cap at 100%
    
    return min(timetable_utilization + exam_utilization, 100)

def create_timetable_statistics_sheet(schedule_df, course_info, classrooms_df, semester, branch, section):
    """Create detailed statistics sheet for timetable verification"""
    
    print(f"[STATS] Creating statistics sheet for Semester {semester}, {branch}, Section {section}...")
    
    # Extract all scheduled courses from the timetable
    scheduled_courses = extract_scheduled_courses_from_timetable(schedule_df)
    
    statistics_data = []
    
    for course_code, course_data in scheduled_courses.items():
        # Get course info
        info = course_info.get(course_code, {})
        
        # Parse LTPSC
        ltpsc_str = info.get('ltpsc', '3-0-0-0-3')
        ltpsc = parse_ltpsc(ltpsc_str)
        
        # Calculate scheduled sessions from timetable
        lecture_count = course_data.get('lecture_count', 0)
        tutorial_count = course_data.get('tutorial_count', 0)
        lab_count = course_data.get('lab_count', 0)
        
        # Check LTPSC compliance
        lectures_compliant = "[OK]" if lecture_count == ltpsc['L'] else f"[FAIL] ({lecture_count}/{ltpsc['L']})"
        tutorials_compliant = "[OK]" if tutorial_count == ltpsc['T'] else f"[FAIL] ({tutorial_count}/{ltpsc['T']})"
        labs_compliant = "[OK]" if lab_count == ltpsc['P'] else f"[FAIL] ({lab_count}/{ltpsc['P']})"
        
        # Get allocated rooms
        allocated_rooms = ', '.join(course_data.get('rooms', [])) if course_data.get('rooms') else 'Not Allocated'
        
        # Check for combined classes (courses with same code in different sections at same time)
        combined_class = check_combined_class(course_code, schedule_df, section, semester, branch)
        
        statistics_data.append({
            'Course Code': f"**{course_code}**",
            'Course Name': info.get('name', 'Unknown'),
            'Faculty': info.get('instructor', 'Unknown'),
            'LTPSC': ltpsc_str,
            'Lectures/Week': f"{lecture_count}/{ltpsc['L']}",
            'Tutorials/Week': f"{tutorial_count}/{ltpsc['T']}",
            'Labs/Week': f"{lab_count}/{ltpsc['P']}",
            'LTPSC Compliance': f"{lectures_compliant} | {tutorials_compliant} | {labs_compliant}",
            'Combined Class': 'Yes' if combined_class else 'No',
            'Allocation Status': '[OK] Complete' if all([
                lecture_count == ltpsc['L'],
                tutorial_count == ltpsc['T'],
                lab_count == ltpsc['P']
            ]) else '[WARN] Partial',
            'Allocated Rooms': allocated_rooms,
            'Room Utilization': calculate_room_utilization(course_data.get('rooms', []), classrooms_df)
        })
    
    # Add summary row
    total_courses = len(scheduled_courses)
    compliant_courses = sum(1 for course in statistics_data if course['Allocation Status'] == '[OK] Complete')
    
    statistics_data.append({
        'Course Code': '**SUMMARY**',
        'Course Name': f'Total Courses: {total_courses}',
        'Faculty': f'Fully Compliant: {compliant_courses}',
        'LTPSC': f'Compliance Rate: {(compliant_courses/total_courses)*100:.1f}%' if total_courses > 0 else 'N/A',
        'Lectures/Week': f'Lectures: {sum(c.get("lecture_count", 0) for c in scheduled_courses.values())}',
        'Tutorials/Week': f'Tutorials: {sum(c.get("tutorial_count", 0) for c in scheduled_courses.values())}',
        'Labs/Week': f'Labs: {sum(c.get("lab_count", 0) for c in scheduled_courses.values())}',
        'LTPSC Compliance': 'OVERALL',
        'Combined Class': '--',
        'Allocation Status': '[OK]' if compliant_courses == total_courses else f'[WARN] {total_courses - compliant_courses} issues',
        'Allocated Rooms': '--',
        'Room Utilization': '--'
    })
    
    return pd.DataFrame(statistics_data)

def extract_scheduled_courses_from_timetable(schedule_df):
    """Extract all scheduled courses from timetable with session counts"""
    
    scheduled_courses = {}
    
    for day in schedule_df.columns:
        for time_slot in schedule_df.index:
            cell_value = str(schedule_df.loc[time_slot, day])
            
            if cell_value in ['Free', 'LUNCH BREAK']:
                continue
            
            # Extract course code and session type
            course_code, session_type, rooms = parse_timetable_cell(cell_value)
            
            if not course_code:
                continue
            
            if course_code not in scheduled_courses:
                scheduled_courses[course_code] = {
                    'lecture_count': 0,
                    'tutorial_count': 0,
                    'lab_count': 0,
                    'rooms': set()
                }
            
            # Count session types
            if '(Tutorial)' in cell_value:
                scheduled_courses[course_code]['tutorial_count'] += 1
            elif '(Lab)' in cell_value:
                scheduled_courses[course_code]['lab_count'] += 1
            else:
                scheduled_courses[course_code]['lecture_count'] += 1
            
            # Add rooms
            if rooms:
                scheduled_courses[course_code]['rooms'].add(rooms)
    
    return scheduled_courses

def parse_timetable_cell(cell_value):
    """Parse timetable cell to extract course code and room"""
    
    if not isinstance(cell_value, str):
        return None, None, None
    
    # Remove classroom info
    if '[' in cell_value and ']' in cell_value:
        # Extract course part and room part
        course_part = cell_value.split('[')[0].strip()
        room_part = '[' + cell_value.split('[')[1]
        
        # Clean course code
        course_clean = course_part.replace(' (Tutorial)', '').replace(' (Lab)', '')
        
        # Try to extract course code pattern
        import re
        course_pattern = r'[A-Z]{2,3}\d{3}'
        match = re.search(course_pattern, course_clean)
        
        if match:
            return match.group(0), course_part, room_part.strip('[]')
        else:
            # Check if it's a basket
            if any(basket in course_clean.upper() for basket in ['ELECTIVE_', 'HSS_', 'PROF_', 'OE_']):
                return course_clean, course_part, room_part.strip('[]')
    
    # For cells without room allocation
    course_clean = cell_value.replace(' (Tutorial)', '').replace(' (Lab)', '')
    
    import re
    course_pattern = r'[A-Z]{2,3}\d{3}'
    match = re.search(course_pattern, course_clean)
    
    if match:
        return match.group(0), cell_value, None
    elif any(basket in course_clean.upper() for basket in ['ELECTIVE_', 'HSS_', 'PROF_', 'OE_']):
        return course_clean, cell_value, None
    
    return None, None, None

def check_combined_class(course_code, schedule_df, section, semester, branch):
    """Check if this course has combined classes with other sections"""
    # This would need cross-section comparison
    # For now, return False - can be enhanced with multi-section analysis
    return False

def calculate_room_utilization(rooms, classrooms_df):
    """Calculate room utilization percentage"""
    if not rooms or classrooms_df.empty:
        return 'N/A'
    
    total_capacity = 0
    used_capacity = 0
    
    for room in rooms:
        room_info = classrooms_df[classrooms_df['Room Number'] == room]
        if not room_info.empty:
            capacity = room_info['Capacity'].iloc[0]
            if isinstance(capacity, (int, float)):
                total_capacity += capacity
                used_capacity += min(capacity, 60)  # Assuming 60 students per course
    
    if total_capacity > 0:
        utilization = (used_capacity / total_capacity) * 100
        return f'{utilization:.1f}%'
    
    return 'N/A'

def create_room_allocation_summary(schedule_df, classrooms_df):
    """Create detailed room allocation summary"""
    
    room_allocations = {}
    
    for day in schedule_df.columns:
        for time_slot in schedule_df.index:
            cell_value = str(schedule_df.loc[time_slot, day])
            
            if '[' in cell_value and ']' in cell_value:
                try:
                    # Extract room
                    room_match = re.search(r'\[(.*?)\]', cell_value)
                    if room_match:
                        room = room_match.group(1)
                        
                        # Extract course
                        course_part = cell_value.split('[')[0].strip()
                        
                        if room not in room_allocations:
                            room_allocations[room] = {
                                'total_sessions': 0,
                                'days_used': set(),
                                'courses': set(),
                                'time_slots': []
                            }
                        
                        room_allocations[room]['total_sessions'] += 1
                        room_allocations[room]['days_used'].add(day)
                        room_allocations[room]['courses'].add(course_part)
                        room_allocations[room]['time_slots'].append(f"{day} {time_slot}")
                except:
                    continue
    
    # Create summary DataFrame
    summary_data = []
    
    for room, allocation in room_allocations.items():
        # Get room details
        room_info = classrooms_df[classrooms_df['Room Number'] == room]
        capacity = room_info['Capacity'].iloc[0] if not room_info.empty else 'Unknown'
        room_type = room_info['Type'].iloc[0] if not room_info.empty else 'Unknown'
        
        summary_data.append({
            'Room Number': room,
            'Type': room_type,
            'Capacity': capacity,
            'Total Sessions': allocation['total_sessions'],
            'Days Used': len(allocation['days_used']),
            'Courses Assigned': len(allocation['courses']),
            'Utilization (Sessions/Day)': f"{allocation['total_sessions']/5:.1f}",
            'Sample Schedule': ', '.join(allocation['time_slots'][:3]) + ('...' if len(allocation['time_slots']) > 3 else '')
        })
    
    return pd.DataFrame(summary_data)

def create_comprehensive_summary(dfs, semester, branch, section_a_df, section_b_df, basket_allocations):
    """Create comprehensive summary sheet with all key metrics"""
    
    summary_data = []
    
    # 1. Course Allocation Summary
    course_info = get_course_info(dfs) if dfs else {}
    scheduled_a = extract_scheduled_courses_from_timetable(section_a_df)
    scheduled_b = extract_scheduled_courses_from_timetable(section_b_df)
    
    total_courses = len(set(list(scheduled_a.keys()) + list(scheduled_b.keys())))
    core_courses = sum(1 for code in scheduled_a.keys() if code in course_info and not course_info[code].get('is_elective', False))
    elective_courses = total_courses - core_courses
    
    summary_data.append({
        'Category': 'Course Allocation',
        'Metric': 'Total Courses Scheduled',
        'Value': total_courses,
        'Details': f'Core: {core_courses}, Elective: {elective_courses}'
    })
    
    # 2. LTPSC Compliance
    compliant_a = sum(1 for course_code, data in scheduled_a.items() 
                     if check_ltpsc_compliance(course_code, data, course_info))
    compliant_b = sum(1 for course_code, data in scheduled_b.items() 
                     if check_ltpsc_compliance(course_code, data, course_info))
    
    compliance_rate = ((compliant_a + compliant_b) / (len(scheduled_a) + len(scheduled_b))) * 100 if (len(scheduled_a) + len(scheduled_b)) > 0 else 0
    
    summary_data.append({
        'Category': 'LTPSC Compliance',
        'Metric': 'Overall Compliance Rate',
        'Value': f'{compliance_rate:.1f}%',
        'Details': f'Section A: {compliant_a}/{len(scheduled_a)}, Section B: {compliant_b}/{len(scheduled_b)}'
    })
    
    # 3. Basket Allocation
    if basket_allocations:
        basket_count = len(basket_allocations)
        basket_courses = sum(len(basket['courses']) for basket in basket_allocations.values())
        
        summary_data.append({
            'Category': 'Elective Baskets',
            'Metric': 'Basket Coverage',
            'Value': f'{basket_count} baskets',
            'Details': f'{basket_courses} courses across all baskets'
        })
    
    # 4. Room Utilization
    classroom_data = dfs.get('classroom', pd.DataFrame())
    if not classroom_data.empty:
        used_rooms_a = set()
        used_rooms_b = set()
        
        for df in [section_a_df, section_b_df]:
            for day in df.columns:
                for time_slot in df.index:
                    cell_value = str(df.loc[time_slot, day])
                    if '[' in cell_value and ']' in cell_value:
                        room_match = re.search(r'\[(.*?)\]', cell_value)
                        if room_match:
                            if df is section_a_df:
                                used_rooms_a.add(room_match.group(1))
                            else:
                                used_rooms_b.add(room_match.group(1))
        
        total_rooms = len(classroom_data)
        used_rooms = len(used_rooms_a.union(used_rooms_b))
        
        summary_data.append({
            'Category': 'Room Utilization',
            'Metric': 'Classroom Usage',
            'Value': f'{used_rooms}/{total_rooms} rooms',
            'Details': f'Utilization: {(used_rooms/total_rooms)*100:.1f}%'
        })
    
    # 5. Time Slot Coverage
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    time_slots = section_a_df.index.tolist()
    
    occupied_slots_a = sum(1 for day in days for slot in time_slots 
                          if section_a_df.loc[slot, day] not in ['Free', 'LUNCH BREAK'])
    occupied_slots_b = sum(1 for day in days for slot in time_slots 
                          if section_b_df.loc[slot, day] not in ['Free', 'LUNCH BREAK'])
    total_slots = len(days) * len(time_slots)
    
    summary_data.append({
        'Category': 'Schedule Density',
        'Metric': 'Time Slot Utilization',
        'Value': f'{(occupied_slots_a + occupied_slots_b)/(total_slots*2)*100:.1f}%',
        'Details': f'A: {occupied_slots_a}/{total_slots}, B: {occupied_slots_b}/{total_slots}'
    })
    
    return pd.DataFrame(summary_data)

def check_ltpsc_compliance(course_code, scheduled_data, course_info):
    """Check if scheduled sessions match LTPSC requirements"""
    
    info = course_info.get(course_code, {})
    ltpsc_str = info.get('ltpsc', '3-0-0-0-3')
    ltpsc = parse_ltpsc(ltpsc_str)
    
    return (scheduled_data.get('lecture_count', 0) == ltpsc['L'] and
            scheduled_data.get('tutorial_count', 0) == ltpsc['T'] and
            scheduled_data.get('lab_count', 0) == ltpsc['P'])

def create_timetable_verification_sheet(schedule_df, course_info, classrooms_df, semester, branch, section):
    """Create detailed verification sheet matching the requested format"""
    
    print(f"[STATS] Creating verification sheet for Semester {semester}, {branch}, Section {section}...")
    
    # Extract all scheduled courses from the timetable
    scheduled_courses = extract_scheduled_courses_from_timetable(schedule_df)
    
    verification_data = []
    
    for course_code, course_data in scheduled_courses.items():
        # Get course info
        info = course_info.get(course_code, {})
        
        # Parse LTPSC
        ltpsc_str = info.get('ltpsc', '3-0-0-0-3')
        ltpsc = parse_ltpsc(ltpsc_str)
        
        # Calculate scheduled sessions from timetable
        lecture_count = course_data.get('lecture_count', 0)
        tutorial_count = course_data.get('tutorial_count', 0)
        lab_count = course_data.get('lab_count', 0)
        
        # Get allocated rooms
        allocated_rooms = ', '.join(course_data.get('rooms', [])) if course_data.get('rooms') else 'Not Allocated'
        
        # Format course code with ** for emphasis
        formatted_course_code = f"**{course_code}**"
        
        # Check if it's an elective basket
        is_basket = any(basket in course_code.upper() for basket in ['ELECTIVE_', 'HSS_', 'PROF_', 'OE_'])
        
        # Format course name
        course_name = info.get('name', 'Unknown')
        if is_basket:
            course_name = "Elective Basket"
        
        # Format faculty
        faculty = info.get('instructor', 'Unknown')
        if is_basket:
            faculty = '–'  # Dash for elective baskets
        
        # Format LTPSC
        formatted_ltpsc = ltpsc_str
        
        # Format lectures/tutorials per week
        if is_basket:
            lect_tuts_week = f"0/0"  # Baskets show 0/0
        else:
            lect_tuts_week = f"{lecture_count}/{tutorial_count}"
        
        # Format labs per week
        labs_week = f"{lab_count}/1" if lab_count > 0 else "0/0"
        
        # Combined class status
        combined_class = "No"  # Default - can be enhanced
        
        # Allocation status
        allocation = "Complete" if (lecture_count >= ltpsc['L'] and 
                                   tutorial_count >= ltpsc['T'] and 
                                   lab_count >= ltpsc['P']) else "Partial"
        
        # Room allocation
        room = allocated_rooms
        
        verification_data.append({
            'Course Code': formatted_course_code,
            'Course Name': course_name,
            'Faculty': faculty,
            'LTPSC': formatted_ltpsc,
            'Lect/Tuts/Week': lect_tuts_week,
            'Labs/Week': labs_week,
            'Combined Class': combined_class,
            'Allocation': allocation,
            'Room': room
        })
    
    # Add summary row
    if verification_data:
        total_courses = len(verification_data)
        complete_allocations = sum(1 for course in verification_data if course['Allocation'] == 'Complete')
        
        verification_data.append({
            'Course Code': '**SUMMARY**',
            'Course Name': f'Total Courses: {total_courses}',
            'Faculty': f'Complete: {complete_allocations}',
            'LTPSC': f'Rate: {(complete_allocations/total_courses)*100:.1f}%',
            'Lect/Tuts/Week': f'Lectures: {sum(c.get("lecture_count", 0) for c in scheduled_courses.values())}',
            'Labs/Week': f'Labs: {sum(c.get("lab_count", 0) for c in scheduled_courses.values())}',
            'Combined Class': '--',
            'Allocation': '[OK]' if complete_allocations == total_courses else f'[WARN] {total_courses - complete_allocations} issues',
            'Room': '--'
        })
    
    return pd.DataFrame(verification_data)

def create_room_allocation_summary_verification(section_a_df, section_b_df, classrooms_df):
    """Create room allocation summary for verification"""
    
    print("[STATS] Creating room allocation summary...")
    
    # Track room usage across both sections
    room_usage = {}
    
    for df, section in [(section_a_df, 'A'), (section_b_df, 'B')]:
        for day in df.columns:
            for time_slot in df.index:
                cell_value = str(df.loc[time_slot, day])
                
                if '[' in cell_value and ']' in cell_value:
                    try:
                        # Extract course and room
                        course_part = cell_value.split('[')[0].strip()
                        room_match = re.search(r'\[(.*?)\]', cell_value)
                        if room_match:
                            room = room_match.group(1)
                            
                            if room not in room_usage:
                                room_usage[room] = {
                                    'total_sessions': 0,
                                    'sections': set(),
                                    'courses': set(),
                                    'time_slots': []
                                }
                            
                            room_usage[room]['total_sessions'] += 1
                            room_usage[room]['sections'].add(section)
                            room_usage[room]['courses'].add(course_part)
                            room_usage[room]['time_slots'].append(f"{day} {time_slot}")
                    except:
                        continue
    
    # Create summary DataFrame
    summary_data = []
    
    for room, usage in room_usage.items():
        # Get room details
        room_info = classrooms_df[classrooms_df['Room Number'] == room]
        capacity = room_info['Capacity'].iloc[0] if not room_info.empty else 'Unknown'
        room_type = room_info['Type'].iloc[0] if not room_info.empty else 'Unknown'
        facilities = room_info['Facilities'].iloc[0] if not room_info.empty and 'Facilities' in room_info.columns else 'Standard'
        
        summary_data.append({
            'Room Number': room,
            'Type': room_type,
            'Capacity': capacity,
            'Facilities': facilities,
            'Total Sessions': usage['total_sessions'],
            'Sections': ', '.join(sorted(usage['sections'])),
            'Courses Assigned': len(usage['courses']),
            'Sample Courses': ', '.join(list(usage['courses'])[:3]) + ('...' if len(usage['courses']) > 3 else ''),
            'Utilization (Sessions/Day)': f"{usage['total_sessions']/5:.1f}"
        })
    
    return pd.DataFrame(summary_data).sort_values('Room Number')

def create_ltpsc_compliance_summary(dfs, semester, branch, section_a_df, section_b_df):
    """Create LTPSC compliance summary sheet"""
    
    print("[STATS] Creating LTPSC compliance summary...")
    
    course_info = get_course_info(dfs) if dfs else {}
    scheduled_a = extract_scheduled_courses_from_timetable(section_a_df)
    scheduled_b = extract_scheduled_courses_from_timetable(section_b_df)
    
    compliance_data = []
    
    # Combine courses from both sections
    all_courses = set(list(scheduled_a.keys()) + list(scheduled_b.keys()))
    
    for course_code in sorted(all_courses):
        info = course_info.get(course_code, {})
        ltpsc_str = info.get('ltpsc', '3-0-0-0-3')
        ltpsc = parse_ltpsc(ltpsc_str)
        
        # Get scheduled counts from both sections
        scheduled_lectures_a = scheduled_a.get(course_code, {}).get('lecture_count', 0)
        scheduled_tutorials_a = scheduled_a.get(course_code, {}).get('tutorial_count', 0)
        scheduled_labs_a = scheduled_a.get(course_code, {}).get('lab_count', 0)
        
        scheduled_lectures_b = scheduled_b.get(course_code, {}).get('lecture_count', 0)
        scheduled_tutorials_b = scheduled_b.get(course_code, {}).get('tutorial_count', 0)
        scheduled_labs_b = scheduled_b.get(course_code, {}).get('lab_count', 0)
        
        # Check compliance
        lectures_ok = (scheduled_lectures_a >= ltpsc['L'] and scheduled_lectures_b >= ltpsc['L'])
        tutorials_ok = (scheduled_tutorials_a >= ltpsc['T'] and scheduled_tutorials_b >= ltpsc['T'])
        labs_ok = (scheduled_labs_a >= ltpsc['P'] and scheduled_labs_b >= ltpsc['P'])
        
        compliance_status = "[OK] FULLY COMPLIANT" if lectures_ok and tutorials_ok and labs_ok else "[WARN] PARTIAL"
        
        compliance_data.append({
            'Course Code': course_code,
            'Course Name': info.get('name', 'Unknown'),
            'Required LTPSC': ltpsc_str,
            'Required (L/T/P)': f"{ltpsc['L']}/{ltpsc['T']}/{ltpsc['P']}",
            'Section A (L/T/P)': f"{scheduled_lectures_a}/{scheduled_tutorials_a}/{scheduled_labs_a}",
            'Section B (L/T/P)': f"{scheduled_lectures_b}/{scheduled_tutorials_b}/{scheduled_labs_b}",
            'Lectures Status': '[OK]' if lectures_ok else '[FAIL]',
            'Tutorials Status': '[OK]' if tutorials_ok else '[FAIL]',
            'Labs Status': '[OK]' if labs_ok else '[FAIL]',
            'Overall Compliance': compliance_status,
            'Notes': 'Meets requirements' if lectures_ok and tutorials_ok and labs_ok else 'Check scheduling'
        })
    
    return pd.DataFrame(compliance_data)

def create_executive_summary(dfs, semester, branch, section_a_df, section_b_df, basket_allocations):
    """Create executive summary sheet"""
    
    print("[STATS] Creating executive summary...")
    
    summary_data = []
    
    # 1. Basic Information
    summary_data.append({'Category': 'Basic Information', 'Metric': 'Semester', 'Value': semester, 'Details': f'Branch: {branch}'})
    summary_data.append({'Category': 'Basic Information', 'Metric': 'Generation Date', 'Value': datetime.now().strftime('%Y-%m-%d %H:%M'), 'Details': 'Timetable generation timestamp'})
    
    # 2. Course Statistics
    course_info = get_course_info(dfs) if dfs else {}
    scheduled_a = extract_scheduled_courses_from_timetable(section_a_df)
    scheduled_b = extract_scheduled_courses_from_timetable(section_b_df)
    
    total_courses = len(set(list(scheduled_a.keys()) + list(scheduled_b.keys())))
    core_courses = sum(1 for code in scheduled_a.keys() if code in course_info and not course_info[code].get('is_elective', False))
    elective_courses = total_courses - core_courses
    
    summary_data.append({'Category': 'Course Statistics', 'Metric': 'Total Courses', 'Value': total_courses, 'Details': f'Core: {core_courses}, Elective: {elective_courses}'})
    
    # 3. LTPSC Compliance
    compliant_courses = 0
    for course_code in set(list(scheduled_a.keys()) + list(scheduled_b.keys())):
        info = course_info.get(course_code, {})
        ltpsc_str = info.get('ltpsc', '3-0-0-0-3')
        ltpsc = parse_ltpsc(ltpsc_str)
        
        scheduled_a_data = scheduled_a.get(course_code, {})
        scheduled_b_data = scheduled_b.get(course_code, {})
        
        if (scheduled_a_data.get('lecture_count', 0) >= ltpsc['L'] and
            scheduled_a_data.get('tutorial_count', 0) >= ltpsc['T'] and
            scheduled_a_data.get('lab_count', 0) >= ltpsc['P'] and
            scheduled_b_data.get('lecture_count', 0) >= ltpsc['L'] and
            scheduled_b_data.get('tutorial_count', 0) >= ltpsc['T'] and
            scheduled_b_data.get('lab_count', 0) >= ltpsc['P']):
            compliant_courses += 1
    
    compliance_rate = (compliant_courses / total_courses * 100) if total_courses > 0 else 0
    summary_data.append({'Category': 'LTPSC Compliance', 'Metric': 'Compliance Rate', 'Value': f'{compliance_rate:.1f}%', 'Details': f'{compliant_courses}/{total_courses} courses fully compliant'})
    
    # 4. Basket Allocation
    if basket_allocations:
        basket_count = len(basket_allocations)
        basket_courses = sum(len(basket['courses']) for basket in basket_allocations.values())
        summary_data.append({'Category': 'Elective Baskets', 'Metric': 'Baskets Scheduled', 'Value': basket_count, 'Details': f'{basket_courses} courses in baskets'})
    
    # 5. Room Utilization
    classroom_data = dfs.get('classroom', pd.DataFrame())
    if not classroom_data.empty:
        used_rooms = set()
        for df in [section_a_df, section_b_df]:
            for day in df.columns:
                for time_slot in df.index:
                    cell_value = str(df.loc[time_slot, day])
                    if '[' in cell_value and ']' in cell_value:
                        room_match = re.search(r'\[(.*?)\]', cell_value)
                        if room_match:
                            used_rooms.add(room_match.group(1))
        
        total_rooms = len(classroom_data)
        summary_data.append({'Category': 'Room Utilization', 'Metric': 'Rooms Used', 'Value': f'{len(used_rooms)}/{total_rooms}', 'Details': f'Utilization: {(len(used_rooms)/total_rooms)*100:.1f}%'})
    
    # 6. Schedule Density
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    time_slots = section_a_df.index.tolist()
    
    occupied_slots_a = sum(1 for day in days for slot in time_slots if section_a_df.loc[slot, day] not in ['Free', 'LUNCH BREAK'])
    occupied_slots_b = sum(1 for day in days for slot in time_slots if section_b_df.loc[slot, day] not in ['Free', 'LUNCH BREAK'])
    total_slots = len(days) * len(time_slots)
    
    density_a = (occupied_slots_a / total_slots) * 100
    density_b = (occupied_slots_b / total_slots) * 100
    avg_density = (density_a + density_b) / 2
    
    summary_data.append({'Category': 'Schedule Density', 'Metric': 'Time Slot Utilization', 'Value': f'{avg_density:.1f}%', 'Details': f'A: {density_a:.1f}%, B: {density_b:.1f}%'})
    
    # 7. Quality Assessment
    quality_score = (compliance_rate * 0.4) + (avg_density * 0.3) + (100 if basket_allocations else 80) * 0.3
    quality_status = "[OK] EXCELLENT" if quality_score >= 90 else "🟡 GOOD" if quality_score >= 75 else "[WARN] NEEDS REVIEW"
    
    summary_data.append({'Category': 'Quality Assessment', 'Metric': 'Overall Quality', 'Value': quality_status, 'Details': f'Score: {quality_score:.1f}/100'})
    
    return pd.DataFrame(summary_data)

def allocate_classrooms_for_exams(exam_schedule_df, classrooms_df, course_data_df):
    """Allocate classrooms for exams based on enrollment estimates"""
    print("[SCHOOL] Allocating classrooms for exams...")
    
    if exam_schedule_df.empty or classrooms_df.empty:
        print("[WARN]  No exam schedule or classroom data available")
        return exam_schedule_df
    
    exam_schedule_with_rooms = exam_schedule_df.copy()
    
    # Estimate enrollment for each exam
    enrollment_estimates = estimate_exam_enrollment(exam_schedule_df, course_data_df)
    
    # Track classroom usage by date and session
    classroom_usage = {}
    
    for idx, exam in exam_schedule_df.iterrows():
        if exam['status'] != 'Scheduled':
            continue
        
        date = exam['date']
        session = exam['session']
        course_code = exam['course_code']
        
        # Create usage key
        usage_key = f"{date}_{session}"
        if usage_key not in classroom_usage:
            classroom_usage[usage_key] = set()
        
        # Get enrollment estimate
        enrollment = enrollment_estimates.get(course_code, 50)
        
        # Find suitable classroom
        suitable_classroom = find_suitable_classroom_for_exam(
            classrooms_df, enrollment, classroom_usage[usage_key]
        )
        
        if suitable_classroom:
            # Mark classroom as used
            classroom_usage[usage_key].add(suitable_classroom)
            
            # Update schedule with classroom
            exam_schedule_with_rooms.at[idx, 'classroom'] = suitable_classroom
            
            # Add capacity info
            room_capacity = classrooms_df[classrooms_df['Room Number'] == suitable_classroom]['Capacity'].iloc[0]
            exam_schedule_with_rooms.at[idx, 'capacity_info'] = f"{enrollment}/{room_capacity}"
            
            print(f"[OK] {course_code} on {date} ({session}): {suitable_classroom} ({enrollment} students)")
        else:
            print(f"[WARN]  No classroom available for {course_code} on {date} ({session})")
    
    return exam_schedule_with_rooms

def estimate_exam_enrollment(exam_schedule_df, course_data_df):
    """Estimate enrollment for exams based on course data"""
    enrollment_estimates = {}
    
    if course_data_df.empty:
        # Default estimates if no course data
        for _, exam in exam_schedule_df.iterrows():
            if exam['status'] == 'Scheduled':
                enrollment_estimates[exam['course_code']] = 50
        return enrollment_estimates
    
    # Match courses and estimate enrollment
    for _, exam in exam_schedule_df.iterrows():
        if exam['status'] != 'Scheduled':
            continue
        
        course_code = exam['course_code']
        
        # Try to find course in course data
        course_match = course_data_df[course_data_df['Course Code'] == course_code]
        
        if not course_match.empty:
            # Estimate based on course type
            is_elective = course_match.iloc[0].get('Elective (Yes/No)', 'No').upper() == 'YES'
            if is_elective:
                enrollment_estimates[course_code] = 40  # Smaller for electives
            else:
                enrollment_estimates[course_code] = 60  # Larger for core courses
        else:
            # Default estimate
            enrollment_estimates[course_code] = 50
    
    return enrollment_estimates

def find_suitable_classroom_for_exam(classrooms_df, enrollment, used_classrooms):
    """Find a suitable classroom for an exam"""
    # Filter available classrooms (not already used in this slot)
    available_rooms = classrooms_df[
        ~classrooms_df['Room Number'].isin(used_classrooms)
    ].copy()
    
    if available_rooms.empty:
        return None
    
    # Filter by capacity
    suitable_rooms = available_rooms[available_rooms['Capacity'] >= enrollment]
    
    if suitable_rooms.empty:
        # Use largest available room if none meet capacity
        suitable_rooms = available_rooms.nlargest(1, 'Capacity')
    
    # Prefer non-lab rooms
    non_lab_rooms = suitable_rooms[
        ~suitable_rooms['Room Number'].astype(str).str.startswith('L', na=False)
    ]
    
    if not non_lab_rooms.empty:
        return non_lab_rooms.iloc[0]['Room Number']
    elif not suitable_rooms.empty:
        return suitable_rooms.iloc[0]['Room Number']
    
    return None

def create_configuration_sheet(config):
    """Create configuration information sheet"""
    config_items = []
    
    # Add main configuration parameters
    config_items.append({'Parameter': 'Maximum Exams Per Day', 'Value': config.get('max_exams_per_day', 2)})
    config_items.append({'Parameter': 'Session Duration (minutes)', 'Value': config.get('session_duration', 180)})
    config_items.append({'Parameter': 'Include Weekends', 'Value': 'Yes' if config.get('include_weekends', False) else 'No'})
    config_items.append({'Parameter': 'Department Conflict Strictness', 'Value': config.get('department_conflict', 'moderate')})
    config_items.append({'Parameter': 'Preference Weight', 'Value': config.get('preference_weight', 'medium')})
    config_items.append({'Parameter': 'Session Balance', 'Value': config.get('session_balance', 'strict')})
    config_items.append({'Parameter': 'Morning Start Time', 'Value': config.get('morning_start', '09:00')})
    config_items.append({'Parameter': 'Afternoon Start Time', 'Value': config.get('afternoon_start', '14:00')})
    
    # Add constraints if present
    if config.get('constraints'):
        constraints = config['constraints']
        config_items.append({'Parameter': 'Allowed Departments', 'Value': ', '.join(constraints.get('departments', []))})
        config_items.append({'Parameter': 'Allowed Exam Types', 'Value': ', '.join(constraints.get('examTypes', []))})
        config_items.append({'Parameter': 'Rules Applied', 'Value': ', '.join(constraints.get('rules', []))})
    
    return pd.DataFrame(config_items)

def save_exam_schedule(schedule_df, start_date, end_date, config=None):
    """Save exam schedule with classroom allocation"""
    try:
        filename = f"exam_schedule_{start_date.strftime('%d-%m-%Y')}_to_{end_date.strftime('%d-%m-%Y')}.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Organized exam schedule with classrooms
            schedule_df.to_excel(writer, sheet_name='Exam_Schedule', index=False)
            
            # Add classroom allocation summary if classrooms were allocated
            if 'classroom' in schedule_df.columns:
                classroom_summary = create_exam_classroom_summary(schedule_df)
                classroom_summary.to_excel(writer, sheet_name='Exam_Classrooms', index=False)
            
            # Add configuration sheet
            if config:
                config_sheet = create_configuration_sheet(config)
                config_sheet.to_excel(writer, sheet_name='Configuration', index=False)
            
            # Add exam summary
            exam_summary = create_exam_summary(schedule_df)
            exam_summary.to_excel(writer, sheet_name='Exam_Summary', index=False)
            
            # Add department summary
            dept_summary = create_department_summary(schedule_df)
            if not dept_summary.empty:
                dept_summary.to_excel(writer, sheet_name='Department_Summary', index=False)
        
        return filename
        
    except Exception as e:
        print(f"[FAIL] Error saving exam schedule: {e}")
        return None

# Debug endpoints
@app.route('/debug/current-data')
def debug_current_data():
    """Debug endpoint to show currently loaded data"""
    try:
        data_frames = load_all_data()
        if data_frames is None:
            return jsonify({
                'success': False,
                'message': 'No data loaded'
            })
        
        debug_info = {
            'success': True,
            'loaded_files': list(data_frames.keys()),
            'file_sizes': {},
            'sample_data': {}
        }
        
        for key, df in data_frames.items():
            debug_info['file_sizes'][key] = {
                'rows': len(df),
                'columns': list(df.columns),
                'memory_usage': df.memory_usage(deep=True).sum()
            }
            
            # Add sample data (first 3 rows)
            if not df.empty:
                debug_info['sample_data'][key] = df.head(3).to_dict('records')
        
        # Add timetable info
                # Look for ALL timetable files (basket, pre-mid, post-mid)
        basket_files = glob.glob(os.path.join(OUTPUT_DIR, "*_baskets.xlsx"))
        pre_mid_files = glob.glob(os.path.join(OUTPUT_DIR, "*_pre_mid_timetable.xlsx"))
        post_mid_files = glob.glob(os.path.join(OUTPUT_DIR, "*_post_mid_timetable.xlsx"))
        
        # Combine all files
        excel_files = basket_files + pre_mid_files + post_mid_files
        
        print(f"[DIR] Looking for timetable files in {OUTPUT_DIR}")
        print(f"[FILE] Found {len(basket_files)} basket, {len(pre_mid_files)} pre-mid, {len(post_mid_files)} post-mid files")
        debug_info['generated_timetables'] = {
            'count': len(excel_files),
            'files': [os.path.basename(f) for f in excel_files]
        }
        
        return jsonify(debug_info)
        
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'traceback': traceback.format_exc()
        })

@app.route('/debug/clear-cache')
def debug_clear_cache():
    """Debug endpoint to clear cached data"""
    global _cached_data_frames, _cached_timestamp, _file_hashes
    
    _cached_data_frames = None
    _cached_timestamp = 0
    _file_hashes = {}
    
    # Clear input directory
    if os.path.exists(INPUT_DIR):
        try:
            shutil.rmtree(INPUT_DIR)
            os.makedirs(INPUT_DIR, exist_ok=True)
        except Exception as e:
            return jsonify({
                'success': False,
                'error': f'Error clearing input directory: {str(e)}'
            })
    
    return jsonify({
        'success': True,
        'message': 'Cache cleared successfully',
        'cache_cleared': True,
        'input_dir_cleared': True
    })

@app.route('/debug/file-matching')
def debug_file_matching():
    """Debug endpoint to check file matching"""
    available_files = os.listdir(INPUT_DIR) if os.path.exists(INPUT_DIR) else []
    required_files = [
        "course_data.csv",
        "faculty_availability.csv",
        "classroom_data.csv",
        "student_data.csv",
        "exams_data.csv"
    ]
    
    matching_results = {}
    
    for required_file in required_files:
        required_clean = required_file.lower().replace(' ', '').replace('_', '').replace('-', '')
        matches = []
        
        for uploaded_file in available_files:
            uploaded_clean = uploaded_file.lower().replace(' ', '').replace('_', '').replace('-', '')
            is_match = (
                required_clean in uploaded_clean or 
                uploaded_clean in required_clean or
                any(part in uploaded_clean for part in required_file.split('_'))
            )
            
            matches.append({
                'uploaded_file': uploaded_file,
                'uploaded_clean': uploaded_clean,
                'is_match': is_match,
                'match_type': 'exact' if uploaded_clean == required_clean else 'partial' if is_match else 'none'
            })
        
        matching_results[required_file] = {
            'required_clean': required_clean,
            'matches': matches,
            'has_match': any(match['is_match'] for match in matches)
        }
    
    return jsonify({
        'available_files': available_files,
        'required_files': required_files,
        'matching_results': matching_results,
        'all_files_matched': all(result['has_match'] for result in matching_results.values())
    })

@app.route('/')
def index():
    return render_template('index.html')

def convert_dataframe_to_html_with_baskets(df, table_id, course_colors, basket_colors, course_info):
    """Convert dataframe to HTML with special handling for basket entries, classroom allocation, and color coding"""
    # Create a copy to avoid modifying the original
    df_display = df.copy()
    
    basket_keywords = ['ELECTIVE_', 'HSS_', 'PROF_', 'OE_']

    def is_basket_entry(text):
        if not isinstance(text, str):
            return False
        normalized = text.upper().replace(' (TUTORIAL)', '')
        return any(keyword in normalized for keyword in basket_keywords)

    def should_hide_classroom_info(course_label):
        if not isinstance(course_label, str):
            return False
        base_label = course_label.replace(' (Tutorial)', '').strip()
        # Hide for basket labels (elective groupings)
        if is_basket_entry(base_label):
            return True
        
        course_code = extract_course_code(base_label)
        if course_code and course_code in course_info:
            return course_info[course_code].get('is_elective', False)
        return False
    
    def build_course_title(course_label):
        """
        Build a tooltip/title for regular courses including Course Name and LTPSC if available.
        For basket entries, provide a concise basket tooltip.
        """
        if not isinstance(course_label, str):
            return ""
        if is_basket_entry(course_label):
            base = course_label.replace(' (Tutorial)', '')
            return f"{base} • Elective basket"
        base = course_label.replace(' (Tutorial)', '').replace(' (Lab)', '').strip()
        code = extract_course_code(base) or base
        info = course_info.get(code, {})
        name = info.get('name', '').strip()
        ltpsc = info.get('ltpsc', '').strip()
        parts = []
        if name:
            parts.append(name)
        if ltpsc:
            parts.append(f"LTPSC: {ltpsc}")
        return " • ".join(parts)

    def style_cell_with_classroom_and_colors(val):
        if isinstance(val, str) and val not in ['Free', 'LUNCH BREAK']:
            # Check for classroom allocation format "Course [Room]"
            if '[' in val and ']' in val:
                try:
                    # Extract course and classroom
                    course_part = val.split('[')[0].strip()
                    room_part = '[' + val.split('[')[1]  # Get everything after first [
                    
                    # Determine if classroom info should be hidden
                    hide_classroom = should_hide_classroom_info(course_part)
                    
                    # Check if it's a basket entry
                    is_basket = is_basket_entry(course_part)
                    title_attr = build_course_title(course_part)
                    
                    if is_basket:
                        basket_key = course_part.replace(' (Tutorial)', '')
                        basket_color = basket_colors.get(basket_key, '#cccccc')
                        if '(Tutorial)' in course_part:
                            if hide_classroom:
                                return f'<span class="basket-entry basket-tutorial" style="background-color: {basket_color}" title="{title_attr}">{course_part}</span>'
                            return f'<span class="basket-entry basket-tutorial" style="background-color: {basket_color}" title="{title_attr}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                        if hide_classroom:
                            return f'<span class="basket-entry elective-basket" style="background-color: {basket_color}" title="{title_attr}">{course_part}</span>'
                        return f'<span class="basket-entry elective-basket" style="background-color: {basket_color}" title="{title_attr}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                    else:
                        # Regular course with classroom - get course color
                        clean_course = course_part.replace(' (Tutorial)', '')
                        course_color = course_colors.get(clean_course, '#cccccc')
                        if hide_classroom:
                            return f'<span class="regular-course" style="background-color: {course_color}" title="{title_attr}">{course_part}</span>'
                        return f'<span class="course-with-room" style="background-color: {course_color}" title="{title_attr}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                except:
                    return val
            
            # Handle courses without classroom allocation but with color coding
            for basket_keyword in basket_keywords:
                if basket_keyword in val.upper():
                    basket_key = val.replace(' (Tutorial)', '')
                    basket_color = basket_colors.get(basket_key, '#cccccc')
                    if '(Tutorial)' in val:
                        return f'<span class="basket-entry basket-tutorial" style="background-color: {basket_color}" title="{build_course_title(val)}">{val}</span>'
                    return f'<span class="basket-entry elective-basket" style="background-color: {basket_color}" title="{build_course_title(val)}">{val}</span>'
            
            # Regular courses without classrooms - apply course color
            if val not in ['Free', 'LUNCH BREAK']:
                clean_course = val.replace(' (Tutorial)', '')
                course_color = course_colors.get(clean_course, '#cccccc')
                return f'<span class="regular-course" style="background-color: {course_color}" title="{build_course_title(val)}">{val}</span>'
                
        return val
    
    # Apply styling to all cells
    for col in df_display.columns:
        df_display[col] = df_display[col].apply(style_cell_with_classroom_and_colors)
    
    # Convert to HTML
    html = df_display.to_html(
        classes='timetable-table', 
        index=True,  # Keep index to show time slots
        escape=False,
        border=0,
        table_id=table_id
    )
    
    return clean_table_html(html)

def generate_basket_colors(baskets):
    """
    Generate consistent colors for elective baskets
    """
    basket_colors = {
        'ELECTIVE_B1': '#FF6B6B',
        'ELECTIVE_B2': '#4ECDC4', 
        'ELECTIVE_B3': '#45B7D1',
        'ELECTIVE_B4': '#96CEB4',
        'HSS_B1': '#FECA57',
        'HSS_B2': '#FF9FF3',
        'HSS_B3': '#54A0FF',
        'HSS_B4': '#5F27CD',
        'PROF_B1': '#00D2D3',
        'PROF_B2': '#FF9F43',
        'OE_B1': '#10AC84',
        'OE_B2': '#EE5A24'
    }
    
    # For any baskets not in predefined colors, generate consistent colors
    for basket in baskets:
        if basket not in basket_colors:
            # Generate color based on basket name hash
            hash_val = sum(ord(char) for char in basket)
            hue = hash_val % 360
            basket_colors[basket] = f'hsl({hue}, 70%, 65%)'
    
    return basket_colors


@app.route('/timetables')
def get_timetables():
    try:
        timetables = []
        # Look for ALL timetable files (basket, pre-mid, post-mid)
        basket_files = glob.glob(os.path.join(OUTPUT_DIR, "*_baskets.xlsx"))
        pre_mid_files = glob.glob(os.path.join(OUTPUT_DIR, "*_pre_mid_timetable.xlsx"))
        post_mid_files = glob.glob(os.path.join(OUTPUT_DIR, "*_post_mid_timetable.xlsx"))
        
        # Combine all files
        excel_files = basket_files + pre_mid_files + post_mid_files
        
        print(f"[DIR] Looking for timetable files in {OUTPUT_DIR}")
        print(f"[FILE] Found {len(basket_files)} basket, {len(pre_mid_files)} pre-mid, {len(post_mid_files)} post-mid files")
        print(f"[FILE] Total files: {len(excel_files)}")
        
        # Load course data for course information - force reload to get latest data
        data_frames = load_all_data(force_reload=True)
        course_info = get_course_info(data_frames) if data_frames else {}
        
        # Generate colors for all courses
        all_courses = set()
        all_baskets = set()
        if data_frames and 'course' in data_frames:
            all_courses = set(data_frames['course']['Course Code'].unique())
            
            # Extract basket information from course data
            if 'Basket' in data_frames['course'].columns:
                all_baskets = set(data_frames['course']['Basket'].dropna().unique())
        
        course_colors = generate_course_colors(all_courses, course_info)
        basket_colors = generate_basket_colors(all_baskets)
        
        for file_path in excel_files:
            filename = os.path.basename(file_path)
            
            # Determine timetable type
            timetable_type = 'regular'
            if '_baskets' in filename:
                timetable_type = 'basket'
            elif '_pre_mid_timetable' in filename or 'pre_mid' in filename.lower():
                timetable_type = 'pre_mid'
            elif '_post_mid_timetable' in filename or 'post_mid' in filename.lower():
                timetable_type = 'post_mid'
            
            print(f"📖 Processing {timetable_type} timetable file: {filename}")
            
            try:
                # Extract semester and branch from filename
                if '_' in filename and filename.count('_') >= 2:
                    # Format: semX_BRANCH_timetable_TYPE.xlsx
                    parts = filename.split('_')
                    sem_part = parts[0].replace('sem', '')
                    branch = parts[1]
                    sem = int(sem_part)
                else:
                    # Legacy format: semX_timetable_TYPE.xlsx
                    sem_part = filename.split('sem')[1].split('_')[0]
                    sem = int(sem_part)
                    branch = None
                
                print(f"📖 Reading {timetable_type} timetable file: {filename} (Branch: {branch}, Semester: {sem})")
                
                # Read both sections from the Excel file
                # For mid-semester timetables, check if both sections exist
                # Try to read Section_A; if missing, fall back to a generic 'Timetable' sheet
                try:
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                except Exception:
                    try:
                        # Some exported timetables use a single sheet named 'Timetable'
                        df_a = pd.read_excel(file_path, sheet_name='Timetable')
                        print(f"   [INFO] Using 'Timetable' sheet as Section_A for {filename}")
                    except Exception:
                        print(f"   [WARN] No Section_A or Timetable sheet in {filename}")
                        continue

                # Section_B may be absent for single-sheet timetables; treat as empty
                try:
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                except Exception:
                    # Not an error — many timetables have only one sheet
                    df_b = pd.DataFrame()
                
                # Clean any unintended index columns like "Unnamed: 0" and set proper index
                def _clean_section_df(df):
                    if df.empty:
                        return df
                    unnamed_cols = [col for col in df.columns if str(col).startswith('Unnamed')]
                    if unnamed_cols:
                        df = df.drop(columns=unnamed_cols)
                    if 'Time Slot' in df.columns:
                        df = df.set_index('Time Slot')
                    return df
                df_a = _clean_section_df(df_a)
                if not df_b.empty:
                    df_b = _clean_section_df(df_b)
                
                # Check if classrooms are allocated (look for [Room] pattern in any cell)
                has_classroom_allocation = False
                for df in [df_a, df_b]:
                    if df.empty:
                        continue
                    for col in df.columns:
                        for val in df[col]:
                            if isinstance(val, str) and '[' in val and ']' in val:
                                has_classroom_allocation = True
                                break
                        if has_classroom_allocation:
                            break
                    if has_classroom_allocation:
                        break
                
                # Try to read basket allocations if available (only for basket timetables)
                basket_allocations = {}
                basket_courses_map = {}
                if timetable_type == 'basket':
                    try:
                        basket_df = pd.read_excel(file_path, sheet_name='Basket_Allocation')
                        for _, row in basket_df.iterrows():
                            basket_name = row['Basket Name']
                            courses_in_basket = row['Courses in Basket'].split(', ')
                            basket_allocations[basket_name] = {
                                'courses': courses_in_basket,
                                'slot': (row['Day'], row['Time Slot'])
                            }
                            # Build basket courses map for legends
                            basket_courses_map[basket_name] = courses_in_basket
                    except:
                        print(f"   [WARN] No basket allocation sheet found in {filename}")
                
                # Try to read classroom allocation details
                classroom_allocation_details = []
                try:
                    classroom_df = pd.read_excel(file_path, sheet_name='Classroom_Allocation')
                    classroom_allocation_details = classroom_df.to_dict('records')
                    print(f"   [SCHOOL] Found classroom allocation details: {len(classroom_allocation_details)} entries")
                except:
                    print(f"   [WARN] No classroom allocation sheet found in {filename}")
                
                # Try to read configuration details to reflect current settings
                configuration_summary = {}
                try:
                    config_df = pd.read_excel(file_path, sheet_name='Configuration')
                    if not config_df.empty and {'Parameter', 'Value'}.issubset(config_df.columns):
                        configuration_summary = dict(zip(config_df['Parameter'], config_df['Value']))
                    else:
                        configuration_summary = config_df.to_dict('records')
                    print(f"   ⚙️ Loaded configuration details for {filename}")
                except:
                    print(f"   [WARN] No configuration sheet found in {filename}")
                    configuration_summary = {}
                
                # Convert to HTML tables with basket-aware, classroom-aware, and color-aware processing
                html_a = convert_dataframe_to_html_with_baskets(df_a, f"sem{sem}_{branch}_A" if branch else f"sem{sem}_A", course_colors, basket_colors, course_info)
                html_b = convert_dataframe_to_html_with_baskets(df_b, f"sem{sem}_{branch}_B" if branch else f"sem{sem}_B", course_colors, basket_colors, course_info) if not df_b.empty else ""
                
                # Extract unique courses AND baskets from the actual schedule
                unique_courses_a, unique_baskets_a = extract_unique_courses_with_baskets(df_a, basket_allocations)
                unique_courses_b, unique_baskets_b = extract_unique_courses_with_baskets(df_b, basket_allocations) if not df_b.empty else ([], [])
                
                # Get course basket information for this semester and branch
                course_baskets = separate_courses_by_type(data_frames, sem, branch) if data_frames else {'core_courses': [], 'elective_courses': []}
                
                # ENHANCED: Build comprehensive basket courses map including ALL elective courses for this semester
                if not course_baskets['elective_courses'].empty and 'Basket' in course_baskets['elective_courses'].columns:
                    for _, course in course_baskets['elective_courses'].iterrows():
                        raw_basket = course.get('Basket', 'Unknown')
                        basket = str(raw_basket).strip().upper() if pd.notna(raw_basket) else 'Unknown'
                        course_code = course['Course Code']
                        if basket not in basket_courses_map:
                            basket_courses_map[basket] = []
                        if course_code not in basket_courses_map[basket]:
                            basket_courses_map[basket].append(course_code)
                
                # ENHANCED: Create comprehensive course lists for legends including ALL basket courses
                all_core_courses = course_baskets['core_courses']['Course Code'].tolist() if not course_baskets['core_courses'].empty else []
                all_elective_courses = course_baskets['elective_courses']['Course Code'].tolist() if not course_baskets['elective_courses'].empty else []
                
                # FIXED: Remove duplicate basket entries and empty baskets
                # Combine scheduled courses with all basket courses for complete legends
                legend_courses_a = set(unique_courses_a)
                legend_courses_b = set(unique_courses_b)
                
                # For mid-semester timetables, add all courses from the schedule
                if timetable_type in ['pre_mid', 'post_mid']:
                    # Add all courses from the schedule for mid-semester timetables
                    all_courses_in_schedule = set(unique_courses_a).union(set(unique_courses_b))
                    legend_courses_a.update(all_courses_in_schedule)
                    legend_courses_b.update(all_courses_in_schedule)
                
                supplemental_baskets = []
                if sem == 5:
                    supplemental_baskets = ['ELECTIVE_B4']

                # Add all elective courses from baskets that appear in the schedule
                for basket_name in unique_baskets_a:
                    if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                        legend_courses_a.update(basket_courses_map[basket_name])
                
                for basket_name in unique_baskets_b:
                    if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                        legend_courses_b.update(basket_courses_map[basket_name])

                # Ensure semester 5 includes ELECTIVE_B4 courses even if slots share across sections
                if supplemental_baskets:
                    for basket_name in supplemental_baskets:
                        if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                            legend_courses_a.update(basket_courses_map[basket_name])
                            legend_courses_b.update(basket_courses_map[basket_name])
                
                # FIXED: Create clean basket lists without duplicates
                clean_baskets_a = [basket for basket in unique_baskets_a if basket in basket_courses_map and basket_courses_map[basket]]
                clean_baskets_b = [basket for basket in unique_baskets_b if basket in basket_courses_map and basket_courses_map[basket]]

                if supplemental_baskets:
                    for basket_name in supplemental_baskets:
                        if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                            if basket_name not in clean_baskets_a:
                                clean_baskets_a.append(basket_name)
                            if basket_name not in clean_baskets_b:
                                clean_baskets_b.append(basket_name)
                
                print(f"   🎨 Color coding: {len(course_colors)} course colors, {len(basket_colors)} basket colors")
                print(f"   [STATS] Legend courses A: {len(legend_courses_a)}, Baskets A: {clean_baskets_a}")
                print(f"   [STATS] Legend courses B: {len(legend_courses_b)}, Baskets B: {clean_baskets_b}")
                
                # Compute scheduled core courses (non-basket) from the actual timetable
                elective_code_set = set(all_elective_courses)
                scheduled_core_courses_a = [code for code in unique_courses_a if code not in elective_code_set]
                scheduled_core_courses_b = [code for code in unique_courses_b if code not in elective_code_set]
                
                def build_course_legend_entries(course_codes):
                    legend_entries = []
                    for code in sorted(course_codes):
                        info = course_info.get(code, {})
                        ltpsc_value = info.get('ltpsc', '') if info else ''
                        legend_entries.append({
                            'code': code,
                            'name': info.get('name', ''),
                            'ltpsc': ltpsc_value,
                            'display': f"{code} ({ltpsc_value})" if ltpsc_value else code
                        })
                    return legend_entries
                
                course_legends_a = build_course_legend_entries(legend_courses_a)
                course_legends_b = build_course_legend_entries(legend_courses_b)
                
                # Add timetable for Section A
                timetable_data = {
                    'semester': sem,
                    'section': 'A',
                    'branch': branch,
                    'filename': filename,
                    'html': html_a,
                    'courses': list(legend_courses_a),  # Use enhanced course list
                    'baskets': clean_baskets_a,  # Use cleaned basket list
                    'basket_courses_map': basket_courses_map,
                    'course_info': course_info,
                    'course_colors': course_colors,
                    'basket_colors': basket_colors,
                    'core_courses': all_core_courses,
                    'elective_courses': all_elective_courses,
                    'scheduled_core_courses': scheduled_core_courses_a,
                    'is_basket_timetable': (timetable_type == 'basket'),
                    'is_pre_mid_timetable': (timetable_type == 'pre_mid'),
                    'is_post_mid_timetable': (timetable_type == 'post_mid'),
                    'all_basket_courses': basket_courses_map,  # Include all basket courses for legends
                    'has_classroom_allocation': has_classroom_allocation,
                    'classroom_details': classroom_allocation_details,
                    'configuration': configuration_summary,
                    'course_legends': course_legends_a,
                    'timetable_type': timetable_type  # Add type for frontend filtering
                }
                timetables.append(timetable_data)
                
                # Add timetable for Section B (if it exists)
                if not df_b.empty and html_b:
                    timetables.append({
                        'semester': sem,
                        'section': 'B',
                        'branch': branch,
                        'filename': filename,
                        'html': html_b,
                        'courses': list(legend_courses_b),  # Use enhanced course list
                        'baskets': clean_baskets_b,  # Use cleaned basket list
                        'basket_courses_map': basket_courses_map,
                        'course_info': course_info,
                        'course_colors': course_colors,
                        'basket_colors': basket_colors,
                        'core_courses': all_core_courses,
                        'elective_courses': all_elective_courses,
                        'scheduled_core_courses': scheduled_core_courses_b,
                        'is_basket_timetable': (timetable_type == 'basket'),
                        'is_pre_mid_timetable': (timetable_type == 'pre_mid'),
                        'is_post_mid_timetable': (timetable_type == 'post_mid'),
                        'all_basket_courses': basket_courses_map,  # Include all basket courses for legends
                        'has_classroom_allocation': has_classroom_allocation,
                        'classroom_details': classroom_allocation_details,
                        'configuration': configuration_summary,
                        'course_legends': course_legends_b,
                        'timetable_type': timetable_type  # Add type for frontend filtering
                    })
                
                print(f"[OK] Loaded {timetable_type.upper()} timetable: {filename}")
                print(f"   Type: {timetable_type}")
                print(f"   Classroom allocation: {'[OK] YES' if has_classroom_allocation else '[FAIL] NO'}")
                
            except Exception as e:
                print(f"[FAIL] Error reading {timetable_type} timetable {filename}: {e}")
                traceback.print_exc()
                continue
        
        print(f"[STATS] Total timetables loaded: {len(timetables)}")
        return jsonify(timetables)
        
    except Exception as e:
        print(f"[FAIL] Error in /timetables: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error loading timetables: {str(e)}'})
    
@app.route('/download/<filename>')
def download_timetable(filename):
    file_path = os.path.join(OUTPUT_DIR, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        return jsonify({'success': False, 'message': 'File not found'})

@app.route('/download-all')
def download_all_timetables():
    try:
        zip_path = os.path.join(OUTPUT_DIR, 'all_timetables.zip')
        
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for file in glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx")):
                zipf.write(file, os.path.basename(file))
        
        return send_file(zip_path, as_attachment=True)
        
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error creating zip: {str(e)}'})

@app.route('/stats')
def get_stats():
    try:
        # Count generated timetables
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        total_timetables = len(excel_files) * 2  # 2 sections per file
        
        # Count courses, faculty, and classrooms from uploaded files
        course_count = 0
        faculty_count = 0
        classroom_count = 0
        
        # Load data to get accurate counts - force reload
        data_frames = load_all_data(force_reload=True)
        if data_frames:
            if 'course' in data_frames:
                course_count = len(data_frames['course'])
            if 'faculty_availability' in data_frames:
                faculty_count = len(data_frames['faculty_availability'])
            if 'classroom' in data_frames:
                classroom_count = len(data_frames['classroom'])
        
        return jsonify({
            'total_timetables': total_timetables,
            'total_courses': course_count,
            'total_faculty': faculty_count,
            'total_classrooms': classroom_count
        })
        
    except Exception as e:
        print(f"[FAIL] Error loading stats: {e}")
        return jsonify({
            'total_timetables': 0,
            'total_courses': 0,
            'total_faculty': 0,
            'total_classrooms': 0
        })

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        print("=" * 50)
        print("📤 RECEIVED FILE UPLOAD REQUEST")
        print("=" * 50)
        
        if 'files' not in request.files:
            return jsonify({
                'success': False,
                'message': 'No files provided'
            }), 400
        
        files = request.files.getlist('files')
        uploaded_files = []
        
        print(f"📦 Received {len(files)} files")
        
        # Clear input directory first using shutil for better file handling
        if os.path.exists(INPUT_DIR):
            try:
                shutil.rmtree(INPUT_DIR)
                os.makedirs(INPUT_DIR, exist_ok=True)
                print("🗑️ Cleared input directory")
            except Exception as e:
                print(f"[WARN] Could not clear input directory: {e}")
        
        for file in files:
            if file and file.filename != '' and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(INPUT_DIR, filename)
                file.save(filepath)
                uploaded_files.append(filename)
                print(f"[OK] Uploaded: {filename} -> {filepath}")
        
        if not uploaded_files:
            return jsonify({
                'success': False,
                'message': 'No valid CSV files uploaded'
            }), 400
        
        # Verify we have all required files with flexible matching
        required_files = [
            "course_data.csv",
            "faculty_availability.csv",
            "classroom_data.csv", 
            "student_data.csv",
            "exams_data.csv"
        ]
        
        missing_files = []
        available_files = os.listdir(INPUT_DIR)
        
        print(f"[DIR] Available files after upload: {available_files}")
        
        for required_file in required_files:
            found = False
            required_clean = required_file.lower().replace(' ', '').replace('_', '').replace('-', '')
            
            for uploaded_file in available_files:
                uploaded_clean = uploaded_file.lower().replace(' ', '').replace('_', '').replace('-', '')
                if (required_clean in uploaded_clean or 
                    uploaded_clean in required_clean or
                    any(part in uploaded_clean for part in required_file.split('_'))):
                    found = True
                    print(f"[OK] Matched {required_file} with {uploaded_file}")
                    break
            
            if not found:
                missing_files.append(required_file)
                print(f"[FAIL] No match found for {required_file}")
        
        if missing_files:
            return jsonify({
                'success': False,
                'message': f'Missing required files: {", ".join(missing_files)}',
                'uploaded_files': uploaded_files,
                'missing_files': missing_files
            }), 400
        
        # Clear the common elective allocations cache when new files are uploaded
        global _SEMESTER_ELECTIVE_ALLOCATIONS
        _SEMESTER_ELECTIVE_ALLOCATIONS = {}
        print("🗑️ Cleared common elective allocations cache")
        
        # Reset classroom usage tracker for new generation
        reset_classroom_usage_tracker()
        print("[RESET] Reset classroom usage tracker for new timetable generation")
        
        # Generate ONLY basket timetables for all branches and semesters
        branches = ['CSE', 'DSAI', 'ECE']
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        # Load the newly uploaded data immediately (force reload) before computing allocations
        data_frames = load_all_data(force_reload=True)
        if data_frames is None:
            return jsonify({
                'success': False,
                'message': 'Failed to load CSV data after upload',
                'uploaded_files': uploaded_files
            }), 400
        
        # Use the basket-based generation endpoint instead of individual course generation
        for sem in target_semesters:
            for branch in branches:
                try:
                    print(f"[RESET] Generating BASKET timetable for {branch} Semester {sem}...")
                    
                    # Use basket-based scheduling
                    success = export_semester_timetable_with_baskets(data_frames, sem, branch)
                    
                    filename = f"sem{sem}_{branch}_timetable_baskets.xlsx"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    
                    if success and os.path.exists(filepath):
                        success_count += 1
                        generated_files.append(filename)
                        print(f"[OK] Successfully generated BASKET timetable: {filename}")
                    else:
                        print(f"[FAIL] Basket timetable not created: {filename}")
                        
                except Exception as e:
                    print(f"[FAIL] Error generating basket timetable for {branch} semester {sem}: {e}")
                    traceback.print_exc()

        return jsonify({
            'success': True,
            'message': f'Successfully uploaded {len(uploaded_files)} files and generated {success_count} BASKET timetables!',
            'uploaded_files': uploaded_files,
            'generated_count': success_count,
            'files': generated_files
        })
        
    except Exception as e:
        print(f"[FAIL] Error uploading files: {e}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'message': f'Error uploading files: {str(e)}'
        }), 500


def export_semester_timetable_with_baskets_common(dfs, semester, branch, common_elective_allocations):
    """Export timetable using pre-allocated common basket slots"""
    try:
        print(f"[STATS] Generating timetable for Semester {semester}, Branch {branch} with COMMON basket slots...")
        
        # Generate schedules using the COMMON basket allocations
        section_a = generate_section_schedule_with_elective_baskets(dfs, semester, 'A', common_elective_allocations, branch)
        section_b = generate_section_schedule_with_elective_baskets(dfs, semester, 'B', common_elective_allocations, branch)
        
        if section_a is None or section_b is None:
            return False

        filename = f"sem{semester}_{branch}_timetable_baskets.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
            # Create basket allocations summary from the common allocations
            basket_allocations = {}
            for allocation in common_elective_allocations.values():
                if allocation:
                    basket_name = allocation['basket_name']
                    if basket_name not in basket_allocations:
                        basket_allocations[basket_name] = {
                            'lectures': allocation['lectures'],
                            'tutorial': allocation['tutorial'],
                            'courses': allocation['all_courses_in_basket']
                        }
            
            basket_summary = create_basket_summary(basket_allocations, semester, branch)
            basket_summary.to_excel(writer, sheet_name='Basket_Allocation', index=False)
            
            course_summary = create_course_summary(dfs, semester, branch)
            if not course_summary.empty:
                course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
        
        print(f"[OK] Generated: {filename} with COMMON basket slots")
        return True
        
    except Exception as e:
        print(f"[FAIL] Error: {e}")
        return False
    
@app.route('/exam-schedule', methods=['POST'])
def generate_exam_schedule():
    """Generate conflict-free exam timetable with configuration and classroom allocation"""
    try:
        print("[RESET] Starting exam schedule generation with classroom allocation...")
        
        data = request.json
        exam_period_start = datetime.strptime(data['start_date'], '%d/%m/%Y')
        exam_period_end = datetime.strptime(data['end_date'], '%d/%m/%Y')
        
        # Get configuration from request
        config = data.get('config', {})
        max_exams_per_day = config.get('max_exams_per_day', 2)
        include_weekends = config.get('include_weekends', False)
        
        print(f"⚙️ Configuration: {max_exams_per_day} exams/day, weekends: {include_weekends}")
        
        # Load exam data
        data_frames = load_all_data(force_reload=True)
        if not data_frames or 'exams' not in data_frames:
            return jsonify({'success': False, 'message': 'No exam data found in CSV files'})
        
        exams_df = data_frames['exams']
        
        # Check if we have any exams to schedule
        if exams_df.empty:
            return jsonify({'success': False, 'message': 'No exam data available in CSV'})
        
        print(f"[LIST] Found {len(exams_df)} exams in CSV data")
        
        # Generate exam schedule with configuration
        session_duration = int(config.get('session_duration', 180) or 180)
        department_conflict = config.get('department_conflict', 'moderate')
        preference_weight = config.get('preference_weight', 'medium')
        session_balance = config.get('session_balance', 'strict')
        constraints = config.get('constraints')
        morning_start = config.get('morning_start', '09:00')
        afternoon_start = config.get('afternoon_start', '14:00')
        
        exam_schedule = schedule_exams_conflict_free(
            exams_df, exam_period_start, exam_period_end, 
            max_exams_per_day, include_weekends,
            session_duration=session_duration,
            department_conflict=department_conflict,
            preference_weight=preference_weight,
            session_balance=session_balance,
            constraints=constraints,
            morning_start=morning_start,
            afternoon_start=afternoon_start
        )
        
        # Check if DataFrame is empty properly
        if exam_schedule is None:
            return jsonify({'success': False, 'message': 'Scheduling algorithm failed to generate any schedule'})
        
        if exam_schedule.empty:
            scheduled_count = len(exam_schedule[exam_schedule['status'] == 'Scheduled'])
            if scheduled_count == 0:
                return jsonify({
                    'success': False, 
                    'message': 'No exams could be scheduled. Try increasing exam period or reducing constraints.'
                })
        
        # ALLOCATE CLASSROOMS for exams
        if data_frames and 'classroom' in data_frames:
            print("[SCHOOL] Allocating classrooms for exams...")
            exam_schedule_with_rooms = allocate_classrooms_for_exams(
                exam_schedule, data_frames['classroom'], data_frames.get('course', pd.DataFrame())
            )
            
            # Add classroom utilization info
            classroom_usage = calculate_classroom_usage_for_exams(exam_schedule_with_rooms)
            print(f"   [STATS] Classroom utilization: {classroom_usage['used_rooms']} rooms used, {classroom_usage['total_sessions']} exam sessions")
        else:
            exam_schedule_with_rooms = exam_schedule
            print("[WARN]  No classroom data available for exam allocation")
        
        # Save to Excel
        filename = save_exam_schedule(exam_schedule_with_rooms, exam_period_start, exam_period_end, config)
        
        # Add this file to the list of schedules to display
        add_exam_schedule_file(filename)
        
        scheduled_count = len(exam_schedule[exam_schedule['status'] == 'Scheduled'])
        total_days = len(exam_schedule['date'].unique())
        
        response_message = f'Exam schedule generated successfully! Scheduled {scheduled_count} exams over {total_days} days.'
        if 'classroom' in data_frames:
            response_message += ' Classroom allocation completed.'
        
        return jsonify({
            'success': True,
            'message': response_message,
            'filename': filename,
            'schedule': exam_schedule_with_rooms.to_dict('records'),
            'config_used': config,
            'is_new_generation': True,
            'classroom_allocated': 'classroom' in data_frames
        })
        
    except Exception as e:
        print(f"[FAIL] Error generating exam schedule: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})
    
@app.route('/generate', methods=['POST'])
def generate_all_timetables():
    """Main generation endpoint - generates basket, pre-mid, and post-mid timetables"""
    try:
        # Call the basket generation which now includes pre-mid and post-mid
        return generate_timetables_with_baskets()
    except Exception as e:
        print(f"[FAIL] Error in /generate endpoint: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/generate-timetable', methods=['POST'])
def generate_timetable_with_config():
    """Generate regular classroom timetables with optional dynamic time configuration."""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}

        if not semester or not branch:
            return jsonify({'success': False, 'message': 'semester and branch are required'}), 400

        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500

        # Generate with provided configuration
        success = export_semester_timetable_with_baskets(dfs, semester, branch, time_config=time_config)
        filename = f"sem{semester}_{branch}_timetable_baskets.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)

        if success and os.path.exists(filepath):
            return jsonify({'success': True, 'message': 'Timetable generated', 'file': filename, 'configuration': time_config})
        return jsonify({'success': False, 'message': 'Failed to generate timetable'}), 500
    except Exception as e:
        print(f"[FAIL] Error generating timetable: {e}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

def schedule_exams_conflict_free(exams_df, start_date, end_date, max_exams_per_day=2, 
                               include_weekends=False, session_duration=180,
                               department_conflict='moderate', preference_weight='medium',
                               session_balance='strict', constraints=None,
                               morning_start='09:00', afternoon_start='14:00'):
    """Generate conflict-free exam schedule with multiple exams per slot"""
    try:
        try:
            session_duration = int(session_duration)
        except (TypeError, ValueError):
            session_duration = 180
        session_duration = max(60, session_duration)
        
        def build_time_slot(start_time_str, fallback):
            try:
                start_dt = datetime.strptime(start_time_str, '%H:%M')
            except Exception:
                start_dt = datetime.strptime(fallback, '%H:%M')
            end_dt = start_dt + timedelta(minutes=session_duration)
            return f"{start_dt.strftime('%H:%M')} - {end_dt.strftime('%H:%M')}"
        
        morning_slot = build_time_slot(morning_start or '09:00', '09:00')
        afternoon_slot = build_time_slot(afternoon_start or '14:00', '14:00')
        
        time_slots = {
            'Morning': morning_slot,
            'Afternoon': afternoon_slot
        }
        
        # VALIDATE max_exams_per_day parameter
        max_exams_per_day = max(1, min(4, max_exams_per_day))  # Ensure between 1-4
        
        print("📅 Generating conflict-free exam schedule with multiple exams per slot...")
        print(f"⚙️ Configuration: {max_exams_per_day} exams/slot, weekends: {include_weekends}")
        print(f"⌛ Session duration: {session_duration} minutes | Slots → Morning: {morning_slot}, Afternoon: {afternoon_slot}")
        print(f"⚙️ Conflict: {department_conflict}, Preference: {preference_weight}, Balance: {session_balance}")
        
        # Set default constraints if none provided
        if constraints is None:
            constraints = {
                'departments': ['CSE', 'DSAI', 'ECE', 'Mathematics', 'Physics', 'Humanities'],
                'examTypes': ['Theory', 'Lab'],
                'rules': ['gapDays', 'sessionLimit', 'preferMorning']
            }
        
        # Convert date strings to datetime objects
        exams_df = exams_df.copy()
        exams_df['Exam Duration (minutes)'] = session_duration
        
        # Parse dates and handle missing values
        exams_df['Preferred Exam Date'] = pd.to_datetime(
            exams_df['Preferred Exam Date'], errors='coerce'
        )
        exams_df['Alternate Exam Date'] = pd.to_datetime(
            exams_df['Alternate Exam Date'], errors='coerce'
        )
        
        # Remove exams with invalid dates
        exams_df = exams_df.dropna(subset=['Preferred Exam Date', 'Alternate Exam Date'])
        
        if exams_df.empty:
            print("[FAIL] No valid exam data found")
            return None
        
        # Initialize schedule
        day_slots = {}
        
        # Generate all possible dates in exam period as date objects (not datetime)
        all_dates = []
        current_date = start_date.date()  # Convert to date object
        end_date_obj = end_date.date()    # Convert to date object
        
        while current_date <= end_date_obj:
            # Check if we should include weekends based on configuration
            if include_weekends or current_date.weekday() < 5:  # 0-4 = Monday-Friday
                all_dates.append(current_date)
                day_slots[current_date] = {
                    'Morning': [],
                    'Afternoon': []
                }
            current_date += timedelta(days=1)
        
        if not all_dates:
            print("[FAIL] No valid dates in exam period")
            return None
        
        print(f"📅 Exam period: {start_date.date()} to {end_date.date()} ({len(all_dates)} days)")
        
        # Extract department from course code (fallback if no department column)
        def extract_department(course_code):
            if course_code.startswith('CS'):
                return 'CSE'
            elif course_code.startswith('DS') or course_code.startswith('DA'):
                return 'DSAI'
            elif course_code.startswith('EC'):
                return 'ECE'
            elif course_code.startswith('MA'):
                return 'Mathematics'
            elif course_code.startswith('PH'):
                return 'Physics'
            elif course_code.startswith('HS'):
                return 'Humanities'
            else:
                return 'General'
        
        # Add department if not present
        if 'Department' not in exams_df.columns:
            exams_df['Department'] = exams_df['Course Code'].apply(extract_department)
        
        # Filter exams based on constraints
        original_count = len(exams_df)
        if constraints and 'departments' in constraints:
            allowed_departments = constraints['departments']
            exams_df = exams_df[exams_df['Department'].isin(allowed_departments)]
            print(f"[LIST] Filtered to {len(exams_df)} exams from allowed departments: {allowed_departments}")
        
        if constraints and 'examTypes' in constraints:
            allowed_exam_types = constraints['examTypes']
            exams_df = exams_df[exams_df['Exam Type'].isin(allowed_exam_types)]
            print(f"[LIST] Filtered to {len(exams_df)} exams of allowed types: {allowed_exam_types}")
        
        if exams_df.empty:
            print(f"[FAIL] No exams remain after applying constraints (had {original_count} exams)")
            return None
        
        # Add semester if not present (extract from course data if available)
        if 'Semester' not in exams_df.columns:
            exams_df['Semester'] = 'N/A'  # Default value
        
        exams_df['Duration_Hours'] = exams_df['Exam Duration (minutes)'] / 60
        
        # Apply preference weight in sorting
        if preference_weight == 'high':
            # Strong preference for requested dates
            exams_df = exams_df.sort_values([
                'Preferred Exam Date', 'Alternate Exam Date', 'Duration_Hours'
            ], ascending=[True, True, False])
        elif preference_weight == 'low':
            # Focus on optimization rather than preferences
            exams_df = exams_df.sort_values([
                'Duration_Hours', 'Preferred Exam Date'
            ], ascending=[False, True])
        else:  # medium (default)
            # Balance between preferences and optimization
            exams_df = exams_df.sort_values([
                'Preferred Exam Date', 'Duration_Hours'
            ], ascending=[True, False])
        
        # Remove duplicate course codes (keep first occurrence)
        exams_df = exams_df.drop_duplicates(subset=['Course Code'], keep='first')
        
        print(f"[NOTE] Processing {len(exams_df)} unique exams...")
        
        # Schedule exams with multiple attempts and fallbacks
        max_attempts = 3
        attempt = 0
        best_schedule = None
        best_success_rate = 0
        
        while attempt < max_attempts:
            scheduled_exams = []
            failed_exams = []
            
            # Reset day_slots for this attempt
            current_day_slots = {}
            for date in all_dates:
                current_day_slots[date] = {
                    'Morning': [],
                    'Afternoon': []
                }
            
            print(f"[RESET] Scheduling attempt {attempt + 1}/{max_attempts}")
            
            for _, exam in exams_df.iterrows():
                exam_code = exam['Course Code']
                preferred_date = exam['Preferred Exam Date'].date()
                alternate_date = exam['Alternate Exam Date'].date()
                duration_hours = exam['Duration_Hours']
                department = exam['Department']
                exam_type = exam['Exam Type']
                
                scheduled_date = None
                scheduled_session = None
                
                # Try preferred date first with relaxed constraints on later attempts
                date_priority = [preferred_date, alternate_date] + all_dates
                
                for date in date_priority:
                    if date not in current_day_slots:
                        continue
                        
                    # Try both sessions
                    for session in ['Morning', 'Afternoon']:
                        current_slot_exams = current_day_slots[date][session]
                        
                        # Check capacity - NOW CONFIGURABLE (1-4 exams per slot)
                        if len(current_slot_exams) >= max_exams_per_day:
                            continue
                        
                        # Check for conflicts with different strictness levels
                        has_conflict = False
                        
                        if attempt == 0:
                            # First attempt: strict conflict detection
                            has_conflict = has_student_conflict_strict(date, session, exam_code, department, current_day_slots, exams_df, max_exams_per_day)
                        elif attempt == 1:
                            # Second attempt: moderate conflict detection
                            has_conflict = has_student_conflict_moderate(date, session, exam_code, department, current_day_slots, exams_df, max_exams_per_day)
                        else:
                            # Final attempt: lenient conflict detection
                            has_conflict = has_student_conflict_lenient(date, session, exam_code, department, current_day_slots, exams_df, max_exams_per_day)
                        
                        if has_conflict:
                            continue
                            
                        # Apply session balancing
                        if not is_session_balanced(date, session, exam_code, current_day_slots, session_balance, max_exams_per_day):
                            continue
                            
                        scheduled_date = date
                        scheduled_session = session
                        break
                    
                    if scheduled_date:
                        break
                
                if scheduled_date and scheduled_session:
                    duration_hours_value = float(duration_hours)
                    duration_hours_label = f"{duration_hours_value:.2f}".rstrip('0').rstrip('.')
                    exam_slot = {
                        'course_code': exam_code,
                        'course_name': exam.get('Course Name', 'Unknown Course'),
                        'exam_type': exam_type,
                        'duration': f"{duration_hours_label} hours",
                        'duration_minutes': exam['Exam Duration (minutes)'],
                        'department': department,
                        'semester': exam.get('Semester', 'N/A'),
                        'date': scheduled_date.strftime('%d-%m-%Y'),
                        'day': scheduled_date.strftime('%A'),
                        'session': scheduled_session,
                        'time_slot': time_slots[scheduled_session],
                        'original_preferred': preferred_date.strftime('%d-%m-%Y'),
                        'status': 'Scheduled'
                    }
                    
                    current_day_slots[scheduled_date][scheduled_session].append(exam_slot)
                    scheduled_exams.append(exam_slot)
                    print(f"[OK] Scheduled {exam_code} on {scheduled_date} ({scheduled_session}) - {department}")
                else:
                    failed_exams.append(exam_code)
                    print(f"[FAIL] Failed to schedule {exam_code} - {department}")
            
            # Calculate success rate for this attempt
            success_rate = len(scheduled_exams) / len(exams_df)
            print(f"[STATS] Attempt {attempt + 1}: {len(scheduled_exams)}/{len(exams_df)} exams scheduled ({success_rate:.1%})")
            
            # Store the best schedule so far
            if success_rate > best_success_rate:
                best_success_rate = success_rate
                best_schedule = (current_day_slots.copy(), scheduled_exams.copy(), failed_exams.copy())
            
            # If we scheduled all exams successfully, break early
            if success_rate >= 0.95:  # 95% success rate is excellent
                print(f"🎉 Excellent schedule found with {success_rate:.1%} success rate")
                break
            
            # If we have acceptable success rate, we can break early on later attempts
            if attempt >= 1 and success_rate >= 0.8:
                print(f"[OK] Acceptable schedule found with {success_rate:.1%} success rate")
                break
            
            attempt += 1
        
        # Use the best schedule found
        if best_schedule:
            current_day_slots, scheduled_exams, failed_exams = best_schedule
            print(f"🏆 Using best schedule with {best_success_rate:.1%} success rate")
        else:
            print("[FAIL] No acceptable schedule found after all attempts")
            return None
        
        # Create final schedule dataframe
        schedule_data = []
        for date in sorted(current_day_slots.keys()):
            # Add morning session exams
            for exam in current_day_slots[date]['Morning']:
                schedule_data.append(exam)
            
            # Add afternoon session exams  
            for exam in current_day_slots[date]['Afternoon']:
                schedule_data.append(exam)
            
            # Add empty day marker if no exams scheduled
            if (not current_day_slots[date]['Morning'] and 
                not current_day_slots[date]['Afternoon']):
                schedule_data.append({
                    'course_code': 'No Exam',
                    'course_name': 'Free',
                    'exam_type': '',
                    'duration': '',
                    'duration_minutes': 0,
                    'department': '',
                    'semester': '',
                    'date': date.strftime('%d-%m-%Y'),
                    'day': date.strftime('%A'),
                    'session': '',
                    'time_slot': '',
                    'original_preferred': '',
                    'status': 'Free'
                })
        
        schedule_df = pd.DataFrame(schedule_data)
        
        print(f"[STATS] Exam scheduling completed: {len(scheduled_exams)} scheduled, {len(failed_exams)} failed")
        if failed_exams:
            print(f"[FAIL] Failed exams: {failed_exams}")
            print(f"💡 Try increasing exam period or maximum exams per day")
        
        return schedule_df
        
    except Exception as e:
        print(f"[FAIL] Error in exam scheduling: {e}")
        traceback.print_exc()
        return None

def has_student_conflict_strict(date, session, exam_code, department, day_slots, exams_df, max_exams_per_day):
    """Strict conflict detection - avoids any potential student overlaps"""
    slot_exams = day_slots[date][session]
    
    if not slot_exams:
        return False
    
    course_prefix = exam_code[:2]
    
    for scheduled_exam in slot_exams:
        scheduled_code = scheduled_exam['course_code']
        scheduled_dept = scheduled_exam['department']
        
        # Conflict: Same department
        if scheduled_dept == department:
            return True
            
        # Conflict: Same course prefix (likely same student group)
        if scheduled_code[:2] == course_prefix:
            return True
    
    return False

def has_student_conflict_moderate(date, session, exam_code, department, day_slots, exams_df, max_exams_per_day):
    """Moderate conflict detection - allows some overlaps"""
    slot_exams = day_slots[date][session]
    
    if not slot_exams:
        return False
    
    course_prefix = exam_code[:2]
    
    for scheduled_exam in slot_exams:
        scheduled_code = scheduled_exam['course_code']
        scheduled_dept = scheduled_exam['department']
        
        # Only conflict if same department AND same course level
        if scheduled_dept == department:
            # Try to detect course level
            if len(exam_code) >= 5 and len(scheduled_code) >= 5:
                try:
                    exam_level = int(exam_code[2:5]) // 100
                    scheduled_level = int(scheduled_code[2:5]) // 100
                    if exam_level == scheduled_level:
                        return True
                except:
                    # If level detection fails, be conservative
                    if scheduled_code[:3] == exam_code[:3]:
                        return True
            else:
                # Fallback: conflict if first 3 characters match
                if scheduled_code[:3] == exam_code[:3]:
                    return True
    
    return False

def has_student_conflict_lenient(date, session, exam_code, department, day_slots, exams_df, max_exams_per_day):
    """Lenient conflict detection - only prevents obvious conflicts"""
    slot_exams = day_slots[date][session]
    
    if not slot_exams:
        return False
    
    # Only conflict if exact same course (shouldn't happen due to deduplication)
    for scheduled_exam in slot_exams:
        if scheduled_exam['course_code'] == exam_code:
            return True
    
    return False

def has_student_conflict_lenient(date, session, exam_code, department, day_slots, exams_df):
    """Lenient conflict detection - only prevents obvious conflicts"""
    slot_exams = day_slots[date][session]
    
    if not slot_exams:
        return False
    
    # Only conflict if exact same course (shouldn't happen due to deduplication)
    for scheduled_exam in slot_exams:
        if scheduled_exam['course_code'] == exam_code:
            return True
    
    return False

def is_session_balanced(date, session, exam_code, day_slots, session_balance, max_exams_per_day):
    """Check if session assignment maintains balance"""
    morning_count = len(day_slots[date]['Morning'])
    afternoon_count = len(day_slots[date]['Afternoon'])
    
    if session_balance == 'strict':
        # Strict: sessions must be within 1 exam of each other
        if session == 'Morning' and morning_count > afternoon_count + 1:
            return False
        if session == 'Afternoon' and afternoon_count > morning_count + 1:
            return False
    
    elif session_balance == 'flexible':
        # Flexible: Allow some imbalance
        if session == 'Morning' and morning_count > afternoon_count + 2:
            return False
        if session == 'Afternoon' and afternoon_count > morning_count + 2:
            return False
    
    # 'none' balance mode always returns True
    return True

def has_student_conflict(date, session, exam_code, department, day_slots, exams_df):
    """Check if scheduling this exam would cause student conflicts - PERMISSIVE VERSION"""
    # Get all exams already scheduled in this slot
    slot_exams = day_slots[date][session]
    
    # If no exams in this slot, no conflict
    if not slot_exams:
        return False
    
    # Extract course prefix and number for better conflict detection
    course_prefix = exam_code[:2]  # e.g., 'CS' from 'CS101'
    
    try:
        course_number = int(exam_code[2:5])  # e.g., 101 from 'CS101'
        course_level = course_number // 100  # e.g., 1 from 101 (100-level course)
    except:
        course_level = 0
    
    for scheduled_exam in slot_exams:
        scheduled_code = scheduled_exam['course_code']
        scheduled_dept = scheduled_exam['department']
        
        # Conflict 1: Same exact course (shouldn't happen due to deduplication)
        if scheduled_code == exam_code:
            return True
            
        # Conflict 2: Same department AND same course level (students likely in same year)
        if (scheduled_dept == department and 
            scheduled_code[:2] == course_prefix and
            len(scheduled_code) >= 5):
            try:
                scheduled_number = int(scheduled_code[2:5])
                scheduled_level = scheduled_number // 100
                if scheduled_level == course_level:
                    return True
            except:
                # If we can't parse course numbers, be conservative
                if scheduled_code[:3] == exam_code[:3]:
                    return True
        
        # Conflict 3: Core courses from same department in same slot
        # (Allow electives to run concurrently since students choose different electives)
        if (scheduled_dept == department and 
            is_core_course(exam_code) and 
            is_core_course(scheduled_code)):
            return True
    
    return False

def is_core_course(course_code):
    """Determine if a course is likely a core course based on naming patterns"""
    # Core courses typically don't have special suffixes
    core_indicators = ['101', '102', '201', '202', '301', '302', '151', '152', '251', '252']
    
    for indicator in core_indicators:
        if indicator in course_code:
            return True
    
    # Courses with basic numbering are usually core
    if len(course_code) == 5 and course_code[2:5].isdigit():
        course_num = int(course_code[2:5])
        return course_num < 400  # Lower numbers are usually core courses
    
    return False

def has_conflict(date, exam_code, department, day_slots, department_conflict='moderate'):
    """Check if scheduling this exam would cause conflicts based on configuration"""
    current_exams = day_slots[date]
    
    # If no exams scheduled yet, no conflict
    if not current_exams:
        return False
    
    # Apply department conflict rules based on configuration
    if department_conflict == 'strict':
        # Strict: No same department on same day
        for scheduled_exam in current_exams:
            scheduled_dept = scheduled_exam.get('department', 'General')
            if scheduled_dept == department:
                return True
                
    elif department_conflict == 'moderate':
        # Moderate: Allow 1 same department per day
        same_department_count = sum(1 for exam in current_exams if exam.get('department') == department)
        if same_department_count >= 1:
            return True
            
    elif department_conflict == 'lenient':
        # Lenient: Allow multiple same departments, but limit based on total exams
        same_department_count = sum(1 for exam in current_exams if exam.get('department') == department)
        if same_department_count >= 2:  # Allow up to 2 same department exams
            return True
    
    # Check for session balance (avoid too many exams of same type)
    theory_count = sum(1 for exam in current_exams if exam.get('exam_type') == 'Theory')
    lab_count = sum(1 for exam in current_exams if exam.get('exam_type') == 'Lab')
    
    # If we're adding a theory exam and already have 2 theory exams, conflict
    # if theory_count >= 2 and exam_type == 'Theory':
    #     return True
    
    # If we're adding a lab exam and already have 2 lab exams, conflict
    # if lab_count >= 2 and exam_type == 'Lab':
    #     return True
    
    return False

def assign_session(day_exams, session_balance='strict', exam_type=None, duration_hours=0):
    """Assign morning or afternoon session based on configuration"""
    morning_count = sum(1 for exam in day_exams if exam.get('session') == 'Morning')
    afternoon_count = sum(1 for exam in day_exams if exam.get('session') == 'Afternoon')
    
    if session_balance == 'strict':
        # Strict: Always balance sessions
        if morning_count <= afternoon_count:
            return 'Morning'
        else:
            return 'Afternoon'
            
    elif session_balance == 'flexible':
        # Flexible: Allow some imbalance
        if morning_count <= afternoon_count + 1:
            return 'Morning'
        else:
            return 'Afternoon'
            
    elif session_balance == 'none':
        # None: No balancing, prefer morning for long exams if configured
        if exam_type == 'Theory' and duration_hours >= 3:
            return 'Morning'  # Prefer morning for long theory exams
        else:
            # Simple round-robin
            if len(day_exams) % 2 == 0:
                return 'Morning'
            else:
                return 'Afternoon'
    
    else:
        # Default: strict balancing
        if morning_count <= afternoon_count:
            return 'Morning'
        else:
            return 'Afternoon'
    
def get_time_slot(session):
    """Get time slot based on session"""
    if session == 'Morning':
        return '09:00 - 12:00'
    else:
        return '14:00 - 17:00'

def create_exam_classroom_summary(exam_schedule_df):
    """Create summary of classroom allocation for exams"""
    summary_data = []
    
    scheduled_exams = exam_schedule_df[exam_schedule_df['status'] == 'Scheduled']
    
    for _, exam in scheduled_exams.iterrows():
        summary_data.append({
            'Course Code': exam['course_code'],
            'Course Name': exam.get('course_name', 'Unknown'),
            'Date': exam['date'],
            'Session': exam['session'],
            'Time Slot': exam.get('time_slot', ''),
            'Allocated Classrooms': exam.get('classroom', 'Not Allocated'),
            'Estimated Enrollment': exam.get('capacity_info', 'Unknown'),
            'Department': exam.get('department', 'Unknown'),
            'Duration': exam.get('duration', '3 hours')
        })
    
    return pd.DataFrame(summary_data)

def save_exam_schedule(schedule_df, start_date, end_date, config=None):
    """Save exam schedule with classroom allocation"""
    try:
        filename = f"exam_schedule_{start_date.strftime('%d-%m-%Y')}_to_{end_date.strftime('%d-%m-%Y')}.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Organized exam schedule with classrooms
            organized_data = []
            for date in sorted(schedule_df['date'].unique()):
                date_exams = schedule_df[schedule_df['date'] == date]
                
                # Morning session
                morning_exams = date_exams[date_exams['session'] == 'Morning']
                for _, exam in morning_exams.iterrows():
                    organized_data.append(exam.to_dict())
                
                # Afternoon session  
                afternoon_exams = date_exams[date_exams['session'] == 'Afternoon']
                for _, exam in afternoon_exams.iterrows():
                    organized_data.append(exam.to_dict())
            
            organized_df = pd.DataFrame(organized_data)
            organized_df.to_excel(writer, sheet_name='Exam_Schedule', index=False)
            
            # Add classroom allocation summary
            if 'classroom' in organized_df.columns:
                classroom_summary = create_exam_classroom_summary(organized_df)
                classroom_summary.to_excel(writer, sheet_name='Exam_Classrooms', index=False)
            
            # Add configuration sheet
            if config:
                config_sheet = create_configuration_sheet(config)
                config_sheet.to_excel(writer, sheet_name='Configuration', index=False)
            
            # Add exam summary
            exam_summary = create_exam_summary(schedule_df)
            exam_summary.to_excel(writer, sheet_name='Exam_Summary', index=False)
            
            # Add department summary
            dept_summary = create_department_summary(schedule_df)
            if not dept_summary.empty:
                dept_summary.to_excel(writer, sheet_name='Department_Summary', index=False)
        
        return filename
        
    except Exception as e:
        print(f"[FAIL] Error saving exam schedule: {e}")
        return None

def create_configuration_sheet(config):
    """Create configuration information sheet"""
    config_info = {
        'Parameter': [
            'Maximum Exams Per Day',
            'Session Duration (minutes)',
            'Include Weekends',
            'Department Conflict Strictness',
            'Preference Weight',
            'Session Balance',
            'Allowed Departments',
            'Allowed Exam Types',
            'Additional Rules'
        ],
        'Value': [
            config.get('max_exams_per_day', 2),
            config.get('session_duration', 180),
            'Yes' if config.get('include_weekends', False) else 'No',
            config.get('department_conflict', 'moderate'),
            config.get('preference_weight', 'medium'),
            config.get('session_balance', 'strict'),
            ', '.join(config.get('constraints', {}).get('departments', [])),
            ', '.join(config.get('constraints', {}).get('examTypes', [])),
            ', '.join(config.get('constraints', {}).get('rules', []))
        ]
    }
    
    return pd.DataFrame(config_info)

def create_exam_summary(schedule_df):
    """Create exam schedule summary"""
    scheduled_exams = schedule_df[schedule_df['status'] == 'Scheduled']
    
    summary = {
        'Total Exams Scheduled': [len(scheduled_exams)],
        'Exam Period': [f"{schedule_df['date'].min()} to {schedule_df['date'].max()}"],
        'Morning Sessions': [len(scheduled_exams[scheduled_exams['session'] == 'Morning'])],
        'Afternoon Sessions': [len(scheduled_exams[scheduled_exams['session'] == 'Afternoon'])],
        'Theory Exams': [len(scheduled_exams[scheduled_exams['exam_type'] == 'Theory'])],
        'Lab Exams': [len(scheduled_exams[scheduled_exams['exam_type'] == 'Lab'])],
        'Departments Involved': [scheduled_exams['department'].nunique()]
    }
    
    return pd.DataFrame(summary)

def create_department_summary(schedule_df):
    """Create department-wise exam summary"""
    scheduled_exams = schedule_df[schedule_df['status'] == 'Scheduled']
    
    if scheduled_exams.empty:
        return pd.DataFrame()
    
    dept_summary = scheduled_exams.groupby('department').agg({
        'course_code': 'count',
        'duration_minutes': 'sum',
        'semester': lambda x: x.nunique()
    }).reset_index()
    
    dept_summary.columns = ['Department', 'Number of Exams', 'Total Duration (min)', 'Semesters Involved']
    dept_summary['Total Duration (hours)'] = dept_summary['Total Duration (min)'] / 60
    
    return dept_summary

@app.route('/exam-timetables')
def get_exam_timetables():
    """Get generated exam timetables - only shows schedules that are marked for display"""
    try:
        # Only get files that are in our display list
        exam_files_to_display = get_exam_schedule_files()
        exam_timetables = []
        
        print(f"[DIR] Looking for {len(exam_files_to_display)} exam schedules to display")
        
        for filename in exam_files_to_display:
            file_path = os.path.join(OUTPUT_DIR, filename)
            if not os.path.exists(file_path):
                print(f"[WARN] File not found, removing from display list: {filename}")
                remove_exam_schedule_file(filename)
                continue
                
            try:
                # Read exam schedule
                schedule_df = pd.read_excel(file_path, sheet_name='Exam_Schedule')
                
                # Try to read configuration
                configuration_summary = {}
                try:
                    config_df = pd.read_excel(file_path, sheet_name='Configuration')
                    if not config_df.empty and {'Parameter', 'Value'}.issubset(config_df.columns):
                        configuration_summary = dict(zip(config_df['Parameter'], config_df['Value']))
                    else:
                        configuration_summary = config_df.to_dict('records')
                    print(f"   ⚙️ Loaded configuration for {filename}")
                except:
                    configuration_summary = {}
                
                # Check if classrooms are allocated
                has_classroom_allocation = 'classroom' in schedule_df.columns
                
                # Convert to HTML with classroom highlighting
                html_table = schedule_df.to_html(
                    classes='exam-timetable-table',
                    index=False,
                    escape=False
                )
                
                # Clean HTML and add classroom styling
                html_table = clean_exam_table_html(html_table)
                
                # Add classroom highlighting if available
                if has_classroom_allocation:
                    html_table = html_table.replace('<td>', '<td class="exam-cell">')
                    # Highlight cells with classroom allocation
                    html_table = re.sub(
                        r'<td class="exam-cell">(.*?)(C\d{3}[^<]*)(.*?)</td>',
                        r'<td class="exam-cell with-classroom"><strong>\1</strong><br><small class="exam-classroom">\2</small></td>',
                        html_table
                    )
                
                exam_timetables.append({
                    'filename': filename,
                    'html': html_table,
                    'schedule_data': schedule_df.to_dict('records'),
                    'period': filename.replace('exam_schedule_', '').replace('.xlsx', '').replace('_', ' to '),
                    'file_exists': True,
                    'has_classroom_allocation': has_classroom_allocation,
                    'configuration': configuration_summary
                })
                
            except Exception as e:
                print(f"[FAIL] Error reading {filename}: {e}")
                remove_exam_schedule_file(filename)
                continue
        
        print(f"[STATS] Loaded {len(exam_timetables)} exam timetables for display")
        return jsonify(exam_timetables)
        
    except Exception as e:
        print(f"[FAIL] Error loading exam timetables: {e}")
        return jsonify([])
    
@app.route('/exam-timetables/all')
def get_all_exam_timetables():
    """Get ALL exam schedule files (for revisiting previous schedules)"""
    try:
        exam_files = glob.glob(os.path.join(OUTPUT_DIR, "exam_schedule_*.xlsx"))
        exam_timetables = []
        
        for file_path in exam_files:
            filename = os.path.basename(file_path)
            try:
                # Read exam schedule
                schedule_df = pd.read_excel(file_path, sheet_name='Exam_Schedule')
                
                # Try to read configuration
                configuration_summary = {}
                try:
                    config_df = pd.read_excel(file_path, sheet_name='Configuration')
                    if not config_df.empty and {'Parameter', 'Value'}.issubset(config_df.columns):
                        configuration_summary = dict(zip(config_df['Parameter'], config_df['Value']))
                    else:
                        configuration_summary = config_df.to_dict('records')
                    print(f"   ⚙️ Loaded configuration for {filename}")
                except:
                    configuration_summary = {}
                
                # Convert to HTML
                html_table = schedule_df.to_html(
                    classes='exam-timetable-table',
                    index=False,
                    escape=False
                )
                
                # Clean HTML
                html_table = clean_exam_table_html(html_table)
                
                # Check if this file is currently in display list
                is_currently_displayed = filename in get_exam_schedule_files()
                
                exam_timetables.append({
                    'filename': filename,
                    'html': html_table,
                    'schedule_data': schedule_df.to_dict('records'),
                    'period': filename.replace('exam_schedule_', '').replace('.xlsx', '').replace('_', ' to '),
                    'is_currently_displayed': is_currently_displayed,
                    'file_exists': True,
                    'configuration': configuration_summary
                })
                
            except Exception as e:
                print(f"[FAIL] Error reading {filename}: {e}")
                continue
        
        return jsonify(exam_timetables)
        
    except Exception as e:
        print(f"[FAIL] Error loading all exam timetables: {e}")
        return jsonify([])

@app.route('/exam-timetables/add-to-display', methods=['POST'])
def add_exam_to_display():
    """Add an exam schedule to the display list"""
    try:
        data = request.json
        filename = data.get('filename')
        
        if filename:
            add_exam_schedule_file(filename)
            return jsonify({'success': True, 'message': f'Added {filename} to display'})
        else:
            return jsonify({'success': False, 'message': 'No filename provided'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/exam-timetables/remove-from-display', methods=['POST'])
def remove_exam_from_display():
    """Remove an exam schedule from the display list"""
    try:
        data = request.json
        filename = data.get('filename')
        
        if filename:
            remove_exam_schedule_file(filename)
            return jsonify({'success': True, 'message': f'Removed {filename} from display'})
        else:
            return jsonify({'success': False, 'message': 'No filename provided'})
            
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/exam-timetables/clear-display', methods=['POST'])
def clear_exam_display():
    """Clear all exam schedules from display list"""
    try:
        clear_exam_schedule_files()
        return jsonify({'success': True, 'message': 'Cleared all exam schedules from display'})
    except Exception as e:
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

def clean_exam_table_html(html):
    """Clean and format exam timetable HTML"""
    html = html.replace('border="1"', 'border="0"')
    html = html.replace('class="dataframe"', 'class="exam-timetable-table"')
    html = html.replace('<thead>', '<thead class="exam-timetable-head">')
    html = html.replace('<tbody>', '<tbody class="exam-timetable-body">')
    return html

@app.route('/generate-with-baskets', methods=['POST'])
def generate_timetables_with_baskets():
    try:
        print("[RESET] Starting basket-based timetable generation with COMMON slots...")
        
        # Clear existing basket timetables first
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*_baskets.xlsx"))
        for file in excel_files:
            try:
                os.remove(file)
                print(f"🗑️ Removed old basket file: {file}")
            except Exception as e:
                print(f"[WARN] Could not remove {file}: {e}")

        # Load data
        data_frames = load_all_data(force_reload=True)
        if data_frames is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'})

        # Generate timetables using basket approach with COMMON slots
        departments = get_departments_from_data(data_frames)
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        # First, create COMMON basket allocations for each semester
        common_basket_allocations = {}
        for sem in target_semesters:
            course_baskets_all = separate_courses_by_type(data_frames, sem)
            elective_courses_all = course_baskets_all['elective_courses']
            if not elective_courses_all.empty:
                elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, sem)
                common_basket_allocations[sem] = elective_allocations
                print(f"📦 Created COMMON basket allocations for semester {sem}")
        
        # Generate timetables for each branch using COMMON allocations
        for branch in departments:
            for sem in target_semesters:
                try:
                    if sem in common_basket_allocations:
                        success = export_semester_timetable_with_baskets_common(
                            data_frames, sem, branch, common_basket_allocations[sem]
                        )
                    else:
                        success = export_semester_timetable_with_baskets(data_frames, sem, branch)
                    
                    filename = f"sem{sem}_{branch}_timetable_baskets.xlsx"
                    
                    if success:
                        success_count += 1
                        generated_files.append(filename)
                        print(f"[OK] Successfully generated: {filename} with COMMON basket slots")
                        
                except Exception as e:
                    print(f"[FAIL] Error generating basket timetable for {branch} semester {sem}: {e}")
        
        # Also generate pre-mid and post-mid timetables
        print("\n" + "="*80)
        print("[RESET] Generating PRE-MID and POST-MID timetables...")
        print("="*80)
        pre_mid_count = 0
        post_mid_count = 0
        
        for branch in departments:
            for sem in target_semesters:
                print(f"\n📅 Processing Semester {sem}, Branch {branch} for mid-semester timetables...")
                try:
                    # Generate mid-semester timetables (pre-mid and post-mid)
                    mid_result = export_mid_semester_timetables(data_frames, sem, branch)
                    
                    print(f"   Result for Semester {sem}, Branch {branch}:")
                    print(f"      Pre-mid success: {mid_result.get('pre_mid_success', False)}")
                    print(f"      Post-mid success: {mid_result.get('post_mid_success', False)}")
                    print(f"      Pre-mid filename: {mid_result.get('pre_mid_filename', 'None')}")
                    print(f"      Post-mid filename: {mid_result.get('post_mid_filename', 'None')}")
                    
                    if mid_result.get('pre_mid_success', False):
                        pre_mid_count += 1
                        generated_files.append(mid_result['pre_mid_filename'])
                        print(f"[OK] Successfully generated PRE-MID: {mid_result['pre_mid_filename']}")
                    else:
                        print(f"[WARN] Pre-mid generation failed or no courses for Semester {sem}, Branch {branch}")
                    
                    if mid_result.get('post_mid_success', False):
                        post_mid_count += 1
                        generated_files.append(mid_result['post_mid_filename'])
                        print(f"[OK] Successfully generated POST-MID: {mid_result['post_mid_filename']}")
                    else:
                        print(f"[WARN] Post-mid generation failed or no courses for Semester {sem}, Branch {branch}")
                        
                except Exception as e:
                    print(f"[FAIL] Error generating mid-semester timetables for {branch} semester {sem}: {e}")
                    traceback.print_exc()
        
        print(f"\n[STATS] Mid-semester generation summary:")
        print(f"   Pre-mid timetables generated: {pre_mid_count}")
        print(f"   Post-mid timetables generated: {post_mid_count}")
        print("="*80 + "\n")

        total_count = success_count + pre_mid_count + post_mid_count
        return jsonify({
            'success': True, 
            'message': f'Successfully generated {success_count} basket timetables, {pre_mid_count} pre-mid, and {post_mid_count} post-mid timetables!',
            'generated_count': total_count,
            'basket_count': success_count,
            'pre_mid_count': pre_mid_count,
            'post_mid_count': post_mid_count,
            'files': generated_files
        })
    
        
    except Exception as e:
        print(f"[FAIL] Error in basket generation endpoint: {e}")
        traceback.print_exc()

@app.route('/generate-mid-semester', methods=['POST'])
def generate_mid_semester_timetables():
    """Generate separate pre-mid and post-mid timetables"""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}

        if not semester:
            return jsonify({'success': False, 'message': 'semester is required'}), 400

        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500

        # Generate mid-semester timetables
        result = export_mid_semester_timetables(dfs, semester, branch, time_config=time_config)
        
        response = {
            'success': result['pre_mid_success'] or result['post_mid_success'],
            'message': 'Mid-semester timetables generated',
            'pre_mid_generated': result['pre_mid_success'],
            'post_mid_generated': result['post_mid_success']
        }
        
        if result['pre_mid_success']:
            response['pre_mid_file'] = result['pre_mid_filename']
        if result['post_mid_success']:
            response['post_mid_file'] = result['post_mid_filename']
        
        return jsonify(response)
        
    except Exception as e:
        print(f"[FAIL] Error generating mid-semester timetables: {e}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

@app.route('/generate-mid-semester-timetables', methods=['POST'])
def generate_mid_semester_timetables_endpoint():  # Changed function name
    """Generate separate pre-mid and post-mid timetables"""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}
        
        if not semester:
            return jsonify({'success': False, 'message': 'semester is required'}), 400
        
        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500
        
        print(f"[RESET] Generating mid-semester timetables for Semester {semester}, Branch {branch}...")
        
        # Generate pre-mid and post-mid timetables
        results = export_mid_semester_timetables(dfs, semester, branch, time_config)
        
        generated_files = []
        if results['pre_mid_success'] and results['pre_mid_filename']:
            generated_files.append(results['pre_mid_filename'])
        
        if results['post_mid_success'] and results['post_mid_filename']:
            generated_files.append(results['post_mid_filename'])
        
        if generated_files:
            message = f'Generated {len(generated_files)} mid-semester timetables: {", ".join(generated_files)}'
            return jsonify({
                'success': True,
                'message': message,
                'files': generated_files,
                'pre_mid_success': results['pre_mid_success'],
                'post_mid_success': results['post_mid_success']
            })
        else:
            return jsonify({
                'success': False,
                'message': 'Failed to generate mid-semester timetables',
                'pre_mid_success': results['pre_mid_success'],
                'post_mid_success': results['post_mid_success']
            })
        
    except Exception as e:
        print(f"[FAIL] Error generating mid-semester timetables: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500
    
@app.route('/generate-pre-mid-timetable', methods=['POST'])
def generate_pre_mid_timetable():
    """Generate pre-mid semester timetable"""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}

        if not semester or not branch:
            return jsonify({'success': False, 'message': 'semester and branch are required'}), 400

        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500

        # Generate pre-mid timetable
        result = export_mid_semester_timetables(dfs, semester, branch, time_config=time_config)
        
        if result['pre_mid_success'] and result['pre_mid_filename']:
            return jsonify({
                'success': True, 
                'message': 'Pre-mid timetable generated successfully!',
                'filename': result['pre_mid_filename'],
                'timetable_type': 'pre_mid'
            })
        else:
            return jsonify({'success': False, 'message': 'Failed to generate pre-mid timetable'}), 500
    except Exception as e:
        print(f"[FAIL] Error generating pre-mid timetable: {e}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

@app.route('/generate-post-mid-timetable', methods=['POST'])
def generate_post_mid_timetable():
    """Generate post-mid semester timetable"""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}

        if not semester or not branch:
            return jsonify({'success': False, 'message': 'semester and branch are required'}), 400

        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500

        # Generate post-mid timetable
        result = export_mid_semester_timetables(dfs, semester, branch, time_config=time_config)
        
        if result['post_mid_success'] and result['post_mid_filename']:
            return jsonify({
                'success': True, 
                'message': 'Post-mid timetable generated successfully!',
                'filename': result['post_mid_filename'],
                'timetable_type': 'post_mid'
            })
        else:
            return jsonify({'success': False, 'message': 'Failed to generate post-mid timetable'}), 500
    except Exception as e:
        print(f"[FAIL] Error generating post-mid timetable: {e}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500

@app.route('/generate-both-mid-timetables', methods=['POST'])
def generate_both_mid_timetables():
    """Generate both pre-mid and post-mid timetables"""
    try:
        data = request.json or {}
        semester = int(data.get('semester'))
        branch = data.get('branch')
        time_config = data.get('time_config') or {}

        if not semester or not branch:
            return jsonify({'success': False, 'message': 'semester and branch are required'}), 400

        # Load latest data
        dfs = load_all_data(force_reload=True)
        if dfs is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'}), 500

        # Generate both timetables
        result = export_mid_semester_timetables(dfs, semester, branch, time_config=time_config)
        
        generated_files = []
        if result['pre_mid_success'] and result['pre_mid_filename']:
            generated_files.append(result['pre_mid_filename'])
        if result['post_mid_success'] and result['post_mid_filename']:
            generated_files.append(result['post_mid_filename'])
        
        if generated_files:
            return jsonify({
                'success': True, 
                'message': f'Generated {len(generated_files)} mid-semester timetable(s)',
                'files': generated_files,
                'pre_mid_success': result['pre_mid_success'],
                'post_mid_success': result['post_mid_success']
            })
        else:
            return jsonify({'success': False, 'message': 'Failed to generate any mid-semester timetables'}), 500
    except Exception as e:
        print(f"[FAIL] Error generating mid-semester timetables: {e}")
        return jsonify({'success': False, 'message': f'Error: {str(e)}'}), 500
    
if __name__ == '__main__':
    print("🚀 Starting Timetable Generator with Comprehensive Statistics...")
    print(f"[FOLDER] Input directory: {INPUT_DIR}")
    print(f"[FOLDER] Output directory: {OUTPUT_DIR}")
    app.run(debug=True, port=5000)