from flask import Flask, render_template, request, jsonify, send_file
import os
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
        print("📁 File count changed")
        _file_hashes = current_hashes
        return True
    
    # If file contents changed
    for file, current_hash in current_hashes.items():
        if file not in _file_hashes or _file_hashes[file] != current_hash:
            print(f"📄 File content changed: {file}")
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
            print("📂 Using cached data frames (files unchanged)")
            return _cached_data_frames
        else:
            print("📂 Cache expired, reloading data")
    
    if files_changed:
        print("📂 Files changed, reloading data")
    
    required_files = [
        "course_data.csv",
        "faculty_availability.csv",
        "classroom_data.csv",
        "student_data.csv",
        "exams_data.csv"
    ]
    dfs = {}
    
    print("📂 Loading CSV files...")
    print(f"📁 Input directory contents: {os.listdir(INPUT_DIR) if os.path.exists(INPUT_DIR) else 'Directory not found'}")
    
    # Update file hashes
    if os.path.exists(INPUT_DIR):
        for file in os.listdir(INPUT_DIR):
            filepath = os.path.join(INPUT_DIR, file)
            _file_hashes[file] = get_file_hash(filepath)
    
    for f in required_files:
        file_path = find_csv_file(f)
        if not file_path:
            print(f"❌ CSV not found: {f}")
            # Try to find any similar file
            files = os.listdir(INPUT_DIR) if os.path.exists(INPUT_DIR) else []
            similar_files = [file for file in files if f.split('_')[0] in file.lower()]
            if similar_files:
                print(f"   💡 Similar files found: {similar_files}")
            return None
        
        try:
            key = f.replace("_data.csv", "").replace(".csv", "")
            dfs[key] = pd.read_csv(file_path)
            print(f"✅ Loaded {f} from {file_path} ({len(dfs[key])} rows)")
            
            # Show sample data for verification
            if not dfs[key].empty:
                print(f"   Columns: {list(dfs[key].columns)}")
                if 'course' in key:
                    print(f"   First 3 courses:")
                    for i, row in dfs[key].head(3).iterrows():
                        print(f"     {i+1}. {row['Course Code'] if 'Course Code' in row else 'N/A'} - Semester: {row.get('Semester', 'N/A')} - Branch: {row.get('Branch', 'N/A')} - Elective: {row.get('Elective (Yes/No)', 'N/A')}")
                
        except Exception as e:
            print(f"❌ Error loading {f}: {e}")
            return None
    
    # Validate required columns in course data
    if 'course' in dfs:
        required_course_columns = ['Course Code', 'Semester', 'LTPSC']
        missing_columns = [col for col in required_course_columns if col not in dfs['course'].columns]
        if missing_columns:
            print(f"❌ Missing columns in course_data: {missing_columns}")
            print(f"   Available columns: {list(dfs['course'].columns)}")
            return None
    
    # Cache the results
    _cached_data_frames = dfs
    _cached_timestamp = time.time()
    
    print("✅ All CSV files loaded successfully!")
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
                'is_common_elective': is_elective
            }
            
            # Debug logging
            print(f"   📝 Course {course_code}: Department = {department}")
            
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
        # Filter courses for the semester
        sem_courses = dfs['course'][
            dfs['course']['Semester'].astype(str).str.strip() == str(semester_id)
        ].copy()
        
        if sem_courses.empty:
            return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}
        
        # ENHANCED: Filter by department if specified - only include courses for the specific department
        if branch and 'Department' in sem_courses.columns:
            # Include courses that are either:
            # 1. Department-specific core courses for this branch
            # 2. Elective courses (common for all departments)
            sem_courses = sem_courses[
                ((sem_courses['Department'] == branch) & 
                 (sem_courses['Elective (Yes/No)'].str.upper() != 'YES')) |
                (sem_courses['Elective (Yes/No)'].str.upper() == 'YES')
            ].copy()
        
        if sem_courses.empty:
            return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}
        
        # Separate core and elective courses
        core_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() != 'YES'
        ].copy()
        
        elective_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() == 'YES'
        ].copy()
        
        print(f"   📊 Course separation for Semester {semester_id}, Department {branch or 'All'}:")
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
        print(f"⚠️ Error separating courses by type: {e}")
        traceback.print_exc()
        return {'core_courses': pd.DataFrame(), 'elective_courses': pd.DataFrame()}

def parse_ltpsc(ltpsc_string):
    """Parse L-T-P-S-C string and return components"""
    try:
        parts = ltpsc_string.split('-')
        if len(parts) == 5:
            return {
                'L': int(parts[0]),  # Lectures per week
                'T': int(parts[1]),  # Tutorials per week
                'P': int(parts[2]),  # Practicals per week
                'S': int(parts[3]),  # S credits
                'C': int(parts[4])   # Total credits
            }
        else:
            return {'L': 3, 'T': 0, 'P': 0, 'S': 0, 'C': 3}
    except:
        return {'L': 3, 'T': 0, 'P': 0, 'S': 0, 'C': 3}

def enforce_elective_day_separation(basket_allocations):
    """Enforce that elective lectures and tutorials are on different days"""
    print("🔍 ENFORCING ELECTIVE DAY SEPARATION...")
    
    for basket_name, allocation in basket_allocations.items():
        lecture_days = set(day for day, time in allocation['lectures'])
        tutorial_day = allocation['tutorial'][0]
        
        if tutorial_day in lecture_days:
            print(f"   ❌ VIOLATION: Basket '{basket_name}' has tutorial on same day as lectures")
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
                
                print(f"   ✅ FIXED: Moved tutorial to {new_tutorial_day}")
            else:
                print(f"   ⚠️  CANNOT FIX: No available days for tutorial")
        else:
            print(f"   ✅ VALID: Basket '{basket_name}' has proper day separation")
    
    return basket_allocations

def schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, lecture_times, tutorial_times, branch=None):
    """Schedule core courses with exactly 2 lectures and 1 tutorial per course"""
    if core_courses.empty:
        return used_slots
    
    course_day_usage = {}
    
    # ENHANCED: Filter core courses to only include department-specific ones
    if branch and 'Department' in core_courses.columns:
        dept_core_courses = core_courses[core_courses['Department'] == branch].copy()
        print(f"   📚 Scheduling {len(dept_core_courses)} department-specific core courses for {branch}...")
        if not dept_core_courses.empty:
            print(f"      Courses to schedule: {dept_core_courses['Course Code'].tolist()}")
    else:
        dept_core_courses = core_courses.copy()
        print(f"   📚 Scheduling {len(dept_core_courses)} core courses...")
    
    if dept_core_courses.empty:
        print(f"   ℹ️ No department-specific core courses found for {branch}")
        return used_slots
    
    # Parse LTPSC for core courses - ENFORCE 2 lectures + 1 tutorial
    for _, course in dept_core_courses.iterrows():
        course_code = course['Course Code']
        
        # ENFORCE: Each course gets exactly 2 lectures and 1 tutorial
        lectures_needed = 2
        tutorials_needed = 1
        
        # Track which days we've used for this course
        course_day_usage[course_code] = {'lectures': set(), 'tutorials': set()}
        
        print(f"      Scheduling {lectures_needed} lectures and {tutorials_needed} tutorial for {course_code} (Department: {course.get('Department', 'General')})...")
        
        # Schedule lectures (1.5 hours each) - ENFORCE 2 lectures on different days
        lectures_scheduled = 0
        max_lecture_attempts = 200
        
        while lectures_scheduled < lectures_needed and max_lecture_attempts > 0:
            max_lecture_attempts -= 1
            
            # Find days where this course doesn't have a lecture yet
            available_days = [d for d in days if d not in course_day_usage[course_code]['lectures']]
            if not available_days:
                # If no available days left, we can't schedule without violating the rule
                print(f"      ⚠️ Cannot schedule more lectures for {course_code} - no available days left")
                break
            
            day = random.choice(available_days)
            time_slot = random.choice(lecture_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = course_code
                used_slots.add(key)
                course_day_usage[course_code]['lectures'].add(day)
                lectures_scheduled += 1
                print(f"      ✅ Scheduled lecture {lectures_scheduled} for {course_code} on {day} at {time_slot}")
        
        # Schedule tutorial (1 hour) - ENFORCE 1 tutorial on a different day
        tutorials_scheduled = 0
        max_tutorial_attempts = 100
        
        while tutorials_scheduled < tutorials_needed and max_tutorial_attempts > 0:
            max_tutorial_attempts -= 1
            
            # Find days where this course doesn't have a tutorial AND doesn't have a lecture
            available_days = [d for d in days if d not in course_day_usage[course_code]['tutorials'] and d not in course_day_usage[course_code]['lectures']]
            if not available_days:
                # If no completely free days, try days without tutorial but with lecture
                available_days = [d for d in days if d not in course_day_usage[course_code]['tutorials']]
                if not available_days:
                    print(f"      ⚠️ Cannot schedule tutorial for {course_code} - no available days left")
                    break
            
            day = random.choice(available_days)
            time_slot = random.choice(tutorial_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                used_slots.add(key)
                course_day_usage[course_code]['tutorials'].add(day)
                tutorials_scheduled += 1
                print(f"      ✅ Scheduled tutorial for {course_code} on {day} at {time_slot}")
        
        if lectures_scheduled < lectures_needed:
            print(f"      ❌ Could only schedule {lectures_scheduled}/{lectures_needed} lectures for {course_code}")
        if tutorials_scheduled < tutorials_needed:
            print(f"      ❌ Could only schedule {tutorials_scheduled}/{tutorials_needed} tutorials for {course_code}")
        if lectures_scheduled == lectures_needed and tutorials_scheduled == tutorials_needed:
            print(f"      ✅ Successfully scheduled {course_code} with 2 lectures and 1 tutorial")

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

    """Allocate elective courses to COMMON time slots for ALL departments of the semester"""
    print(f"🎯 Allocating elective courses to COMMON SLOTS for Semester {semester_id}...")
    
    # Group electives by basket
    basket_groups = {}
    for _, course in elective_courses.iterrows():
        basket = course.get('Basket', 'ELECTIVE_B1')  # Default basket if not specified
        if basket not in basket_groups:
            basket_groups[basket] = []
        basket_groups[basket].append(course)
    
    print(f"   Found {len(basket_groups)} elective baskets: {list(basket_groups.keys())}")
    
    # Define COMMON time slots for ALL departments - 2 lectures + 1 tutorial per basket
    common_lecture_slots = [
        ('Mon', '09:00-10:30'), ('Mon', '13:00-14:30'), 
        ('Tue', '09:00-10:30'), ('Tue', '13:00-14:30'),
        ('Wed', '09:00-10:30'), ('Wed', '13:00-14:30'),
        ('Thu', '09:00-10:30'), ('Thu', '13:00-14:30'),
        ('Fri', '09:00-10:30'), ('Fri', '13:00-14:30')
    ]
    
    common_tutorial_slots = [
        ('Mon', '14:30-15:30'), ('Tue', '14:30-15:30'),
        ('Wed', '14:30-15:30'), ('Thu', '14:30-15:30'),
        ('Fri', '14:30-15:30')
    ]
    
    elective_allocations = {}
    basket_allocations = {}
    
    # Use FIXED allocation to ensure consistency across ALL departments
    lecture_idx = 0
    tutorial_idx = 0
    
    for basket_name in sorted(basket_groups.keys()):  # Sort for consistent ordering
        basket_courses = basket_groups[basket_name]
        course_codes = [course['Course Code'] for course in basket_courses]
        
        # Allocate 2 lectures and 1 tutorial for this basket - COMMON for ALL departments
        lectures_allocated = []
        tutorial_allocated = None
        
        # Allocate 2 lectures from common slots
        for _ in range(2):
            if lecture_idx < len(common_lecture_slots):
                lectures_allocated.append(common_lecture_slots[lecture_idx])
                lecture_idx += 1
            else:
                print(f"   ⚠️ Not enough common lecture slots available for basket '{basket_name}'")
        
        # Allocate 1 tutorial from common slots
        if tutorial_idx < len(common_tutorial_slots):
            tutorial_allocated = common_tutorial_slots[tutorial_idx]
            tutorial_idx += 1
        else:
            print(f"   ⚠️ Not enough common tutorial slots available for basket '{basket_name}'")
        
        if len(lectures_allocated) == 2 and tutorial_allocated:
            # Store allocation for all courses in this basket - COMMON for ALL departments
            for course_code in course_codes:
                elective_allocations[course_code] = {
                    'basket_name': basket_name,
                    'lectures': lectures_allocated,
                    'tutorial': tutorial_allocated,
                    'all_courses_in_basket': course_codes,
                    'for_all_branches': True,
                    'for_both_sections': True,
                    'common_for_semester': True,
                    'common_for_all_departments': True  # NEW: Explicitly mark as common for all departments
                }
            
            basket_allocations[basket_name] = {
                'lectures': lectures_allocated,
                'tutorial': tutorial_allocated,
                'courses': course_codes,
                'common_for_all_departments': True
            }
            
            print(f"   🗂️ COMMON SLOTS for Basket '{basket_name}' (ALL Departments):")
            for i, (day, time_slot) in enumerate(lectures_allocated, 1):
                print(f"      Lecture {i}: {day} {time_slot}")
            print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
            print(f"      Courses: {', '.join(course_codes)}")
        else:
            print(f"   ❌ Could not allocate all required COMMON slots for basket '{basket_name}'")
    
    return elective_allocations, basket_allocations

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
        
        print(f"      🗂️ Basket '{basket_name}' - IDENTICAL across all branches:")
        print(f"         Courses: {', '.join(all_courses)}")
        print(f"         Days Separation: {'✅' if days_separated else '❌'}")
        
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
                print(f"         ✅ COMMON LECTURE: {day} {time_slot}")
                print(f"                SAME for ALL branches & sections")
            else:
                print(f"         ❌ LECTURE CONFLICT: {day} {time_slot} - {schedule.loc[time_slot, day]}")
        
        # Schedule tutorial - IDENTICAL for all branches and sections
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
                    print(f"         ✅ COMMON TUTORIAL: {day} {time_slot}")
                    print(f"                SAME for ALL branches & sections")
                else:
                    print(f"         ❌ TUTORIAL CONFLICT: {day} {time_slot} - {schedule.loc[time_slot, day]}")
    
    print(f"   ✅ Scheduled {elective_scheduled} IDENTICAL COMMON elective sessions")
    return used_slots

def generate_section_schedule_with_elective_baskets(dfs, semester_id, section, elective_allocations, branch=None):
    """Generate schedule with basket-based elective allocation - COMMON slots across branches"""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"   🎯 Generating BASKET-BASED schedule for Semester {semester_id}, Section {section}{branch_info}")
    print(f"   📍 Using COMMON elective basket slots (same for all branches)")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
    
    try:
        # Get only the core courses for this specific branch
        course_baskets = separate_courses_by_type(dfs, semester_id, branch)
        core_courses = course_baskets['core_courses']
        
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        
        # Time slot structure
        morning_slots = ['09:00-10:30', '10:30-12:00']
        lunch_slots = ['12:00-13:00']
        afternoon_slots = ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00']
        all_slots = morning_slots + lunch_slots + afternoon_slots
        
        # Lecture slots (1.5 hours)
        lecture_times = ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00']
        
        # Tutorial slots (1 hour)
        tutorial_times = ['14:30-15:30', '17:00-18:00']
        
        # Create schedule template
        schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
        schedule.loc['12:00-13:00'] = 'LUNCH BREAK'

        used_slots = set()

        # Schedule elective courses FIRST using COMMON basket allocation
        if elective_allocations:
            print(f"   📅 Applying COMMON basket elective slots for Section {section}:")
            # Show what we're scheduling
            unique_baskets = set()
            for allocation in elective_allocations.values():
                if allocation:
                    unique_baskets.add(allocation['basket_name'])
            
            print(f"   🗂️ Baskets to schedule: {list(unique_baskets)}")
            
            used_slots = schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section, branch)
        
        # Schedule core courses AFTER electives - these are branch-specific
        if not core_courses.empty:
            print(f"   📚 Scheduling {len(core_courses)} BRANCH-SPECIFIC core courses for {branch}...")
            used_slots = schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, 
                                                            lecture_times, tutorial_times, branch)
        
        return schedule
        
    except Exception as e:
        print(f"❌ Error generating basket-based schedule: {e}")
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
        separation_status = "✅ DIFFERENT DAYS" if days_separated else "❌ SAME DAY"
        
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
    print(f"🎯 Allocating COMMON elective slots for Semester {semester_id} (ALL branches & sections)...")
    
    # Group electives by basket
    basket_groups = {}
    for _, course in elective_courses.iterrows():
        basket = course.get('Basket', 'ELECTIVE_B1')
        if basket not in basket_groups:
            basket_groups[basket] = []
        basket_groups[basket].append(course)
    
    print(f"   Found {len(basket_groups)} elective baskets: {list(basket_groups.keys())}")
    
    # FILTER BASKETS BY SEMESTER - Only schedule required baskets
    if semester_id == 3:
        # Semester 3: Only ELECTIVE_B3, exclude ELECTIVE_B5
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket == 'ELECTIVE_B3']
        print(f"   🎯 Semester 3: Scheduling only ELECTIVE_B3, excluding ELECTIVE_B5")
    elif semester_id == 5:
        # Semester 5: Only ELECTIVE_B5, exclude ELECTIVE_B3
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket == 'ELECTIVE_B5']
        print(f"   🎯 Semester 5: Scheduling only ELECTIVE_B5, excluding ELECTIVE_B3")
    elif semester_id == 7:
        # FIXED: Semester 7: Schedule BOTH ELECTIVE_B6 and ELECTIVE_B7
        baskets_to_schedule = [basket for basket in basket_groups.keys() if basket in ['ELECTIVE_B6', 'ELECTIVE_B7']]
        print(f"   🎯 Semester 7: Scheduling BOTH ELECTIVE_B6 and ELECTIVE_B7")
    else:
        # Other semesters: Schedule all baskets
        baskets_to_schedule = list(basket_groups.keys())
        print(f"   🎯 Semester {semester_id}: Scheduling all baskets")
    
    print(f"   📋 Baskets to schedule: {baskets_to_schedule}")
    
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
            print(f"   ⚠️ Basket {basket_name} not found in course data")
            continue
            
        basket_courses = basket_groups[basket_name]
        course_codes = [course['Course Code'] for course in basket_courses]
        
        # Get the FIXED common slots for this basket
        if basket_name in common_slots_mapping:
            fixed_slots = common_slots_mapping[basket_name]
            lectures_allocated = fixed_slots['lectures']
            tutorial_allocated = fixed_slots['tutorial']
        else:
            # Fallback for baskets not in mapping
            lectures_allocated = [('Mon', '09:00-10:30'), ('Wed', '09:00-10:30')]
            tutorial_allocated = ('Fri', '14:30-15:30')
            print(f"   ⚠️ Using fallback slots for unknown basket: {basket_name}")
        
        # Verify day separation
        lecture_days = set(day for day, time in lectures_allocated)
        tutorial_day = tutorial_allocated[0]
        days_separated = tutorial_day not in lecture_days
        
        # Store allocation for ALL courses in this basket
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
                'fixed_common_slots': True
            }
        
        basket_allocations[basket_name] = {
            'lectures': lectures_allocated,
            'tutorial': tutorial_allocated,
            'courses': course_codes,
            'common_for_all_departments': True,
            'lecture_days': list(lecture_days),
            'tutorial_day': tutorial_day,
            'days_separated': days_separated,
            'fixed_common_slots': True
        }
        
        print(f"   🗂️ FIXED COMMON SLOTS for Basket '{basket_name}':")
        print(f"      📍 SAME FOR ALL BRANCHES & SECTIONS")
        for i, (day, time_slot) in enumerate(lectures_allocated, 1):
            print(f"      Lecture {i}: {day} {time_slot}")
        print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
        print(f"      Days Separation: {'✅ DIFFERENT DAYS' if days_separated else '❌ SAME DAY'}")
        print(f"      Courses: {', '.join(course_codes)}")
    
    # Log excluded baskets
    excluded_baskets = set(basket_groups.keys()) - set(baskets_to_schedule)
    if excluded_baskets:
        print(f"   🚫 Excluded baskets for Semester {semester_id}: {list(excluded_baskets)}")
    
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
            'Common for All Branches': '✅ YES',
            'Common for Both Sections': '✅ YES', 
            'Days Separation': '✅ YES' if allocation['days_separated'] else '❌ NO',
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
            'Days Separation': '✅ Achieved' if allocation['days_separated'] else '❌ Not Achieved',
            'Slot Consistency': '✅ IDENTICAL across all',
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
            'Status': '✅ Applied'
        })
    elif semester == 5:
        rules_data.append({
            'Semester': 'Semester 5', 
            'Rule': 'Schedule only ELECTIVE_B5',
            'Exclusion': 'Exclude ELECTIVE_B3',
            'Reason': 'Curriculum requirement - Semester 5 focuses on B5 electives',
            'Scheduled Baskets': ', '.join(basket_allocations.keys()) if basket_allocations else 'None',
            'Status': '✅ Applied'
        })
    else:
        rules_data.append({
            'Semester': f'Semester {semester}',
            'Rule': 'Schedule all elective baskets',
            'Exclusion': 'None',
            'Reason': 'No specific restrictions for this semester',
            'Scheduled Baskets': ', '.join(basket_allocations.keys()) if basket_allocations else 'None',
            'Status': '✅ Applied'
        })
    
    return pd.DataFrame(rules_data)

def export_semester_timetable_with_baskets(dfs, semester, branch=None):
    """Export timetable using IDENTICAL COMMON elective slots for ALL branches and sections with classroom allocation"""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"\n📊 Generating timetable for Semester {semester}{branch_info}...")
    print(f"🎯 Using IDENTICAL COMMON elective slots for ALL branches & sections of Semester {semester}")
    
    # Show semester-specific basket rules
    if semester == 3:
        print(f"   🎯 SEMESTER 3 RULE: Scheduling only ELECTIVE_B3, excluding ELECTIVE_B5")
    elif semester == 5:
        print(f"   🎯 SEMESTER 5 RULE: Scheduling only ELECTIVE_B5, excluding ELECTIVE_B3")
    elif semester == 7:
        print(f"   🎯 SEMESTER 7 RULE: Scheduling BOTH ELECTIVE_B6 and ELECTIVE_B7")
    else:
        print(f"   🎯 SEMESTER {semester}: Scheduling all elective baskets")
    
    try:
        # Get ALL elective courses for this semester (without branch filter)
        course_baskets_all = separate_courses_by_type(dfs, semester)
        elective_courses_all = course_baskets_all['elective_courses']
        
        print(f"🎯 Elective courses for Semester {semester} (COMMON for ALL): {len(elective_courses_all)}")
        if not elective_courses_all.empty:
            print("   All courses found:", elective_courses_all['Course Code'].tolist())
            # Show basket distribution
            basket_counts = elective_courses_all['Basket'].value_counts()
            print("   Basket distribution in data:")
            for basket, count in basket_counts.items():
                courses = elective_courses_all[elective_courses_all['Basket'] == basket]['Course Code'].tolist()
                print(f"      {basket}: {count} courses - {courses}")
        
        # Allocate electives using FIXED COMMON slots (with semester filtering)
        elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses_all, semester)
        
        print(f"   📅 FINAL SCHEDULED BASKETS for Semester {semester}:")
        if basket_allocations:
            for basket_name, allocation in basket_allocations.items():
                status = "✅ VALID" if allocation['days_separated'] else "❌ INVALID"
                print(f"      {basket_name}: {status}")
                print(f"         Lectures: {allocation['lectures']}")
                print(f"         Tutorial: {allocation['tutorial']}")
                print(f"         📍 SAME FOR: CSE, DSAI, ECE - Sections A & B")
        else:
            print(f"      No elective baskets scheduled for Semester {semester}")
        
        # Generate schedules - these will have IDENTICAL elective slots
        section_a = generate_section_schedule_with_elective_baskets(dfs, semester, 'A', elective_allocations, branch)
        section_b = generate_section_schedule_with_elective_baskets(dfs, semester, 'B', elective_allocations, branch)
        
        if section_a is None or section_b is None:
            return False

        # ALLOCATE CLASSROOMS for both sections
        course_info = get_course_info(dfs) if dfs else {}
        classroom_data = dfs.get('classroom')
        
        if classroom_data is not None and not classroom_data.empty:
            print("🏫 Allocating classrooms for timetable sessions...")
            section_a_with_rooms = allocate_classrooms_for_timetable(section_a, classroom_data, course_info)
            section_b_with_rooms = allocate_classrooms_for_timetable(section_b, classroom_data, course_info)
            
            # Check if classroom allocation was successful
            has_classroom_allocation_a = check_for_classroom_allocation(section_a_with_rooms)
            has_classroom_allocation_b = check_for_classroom_allocation(section_b_with_rooms)
            has_classroom_allocation = has_classroom_allocation_a or has_classroom_allocation_b
            
            if has_classroom_allocation:
                print("   ✅ Classroom allocation completed successfully")
            else:
                print("   ⚠️ Classroom allocation attempted but no rooms were assigned")
        else:
            section_a_with_rooms = section_a
            section_b_with_rooms = section_b
            print("⚠️  No classroom data available for allocation")

        # Create filename
        filename = f"sem{semester}_{branch}_timetable_baskets.xlsx" if branch else f"sem{semester}_timetable_baskets.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            # Save schedules with classroom allocation
            section_a_with_rooms.to_excel(writer, sheet_name='Section_A')
            section_b_with_rooms.to_excel(writer, sheet_name='Section_B')
            
            # Enhanced basket summary showing common slots
            basket_summary = create_common_basket_summary(basket_allocations, semester, branch)
            basket_summary.to_excel(writer, sheet_name='Basket_Allocation', index=False)
            
            course_summary = create_course_summary(dfs, semester, branch)
            if not course_summary.empty:
                course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
            
            basket_courses_sheet = create_basket_courses_sheet(basket_allocations)
            basket_courses_sheet.to_excel(writer, sheet_name='Basket_Courses', index=False)
            
            # Detailed common slots info
            common_slots_info = create_detailed_common_slots_info(basket_allocations, semester)
            common_slots_info.to_excel(writer, sheet_name='Common_Slots_Info', index=False)
            
            # Add semester-specific rules sheet
            semester_rules = create_semester_rules_sheet(semester, basket_allocations)
            semester_rules.to_excel(writer, sheet_name='Semester_Rules', index=False)
            
            # Add classroom allocation summary if classrooms were allocated
            if classroom_data is not None and not classroom_data.empty:
                classroom_report = create_classroom_utilization_report(
                    classroom_data, [section_a_with_rooms, section_b_with_rooms], []
                )
                classroom_report.to_excel(writer, sheet_name='Classroom_Utilization', index=False)
                
                # Add detailed classroom allocation
                classroom_allocation_detail = create_classroom_allocation_detail(
                    [section_a_with_rooms, section_b_with_rooms], classroom_data
                )
                classroom_allocation_detail.to_excel(writer, sheet_name='Classroom_Allocation', index=False)
        
        success_message = f"✅ Timetable with semester-specific elective rules saved: {filename}"
        if classroom_data is not None and not classroom_data.empty:
            success_message += " (with classroom allocation)"
        
        print(success_message)
        return True
        
    except Exception as e:
        print(f"❌ Error generating timetable: {e}")
        traceback.print_exc()
        return False

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
    print(f"\n📊 Generating BASKET-BASED timetable for Semester {semester}{branch_info}...")
    print(f"🎯 Using COMMON elective basket slots across all branches")
    
    try:
        # CRITICAL: Get ALL elective courses for this semester ONCE (without branch filter)
        # This ensures COMMON allocation for all branches
        course_baskets_all = separate_courses_by_type(dfs, semester)  # No branch filter
        elective_courses_all = course_baskets_all['elective_courses']
        
        print(f"🎯 Elective courses found for semester {semester} (COMMON for ALL branches): {len(elective_courses_all)}")
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

        # Create filename
        if branch:
            filename = f"sem{semester}_{branch}_timetable.xlsx"
        else:
            filename = f"sem{semester}_timetable.xlsx"
            
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
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
        
        print(f"✅ Basket-based timetable saved: {filename}")
        return True
        
    except Exception as e:
        print(f"❌ Error generating basket-based timetable: {e}")
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

def create_course_summary(dfs, semester, branch=None):
    """Create a summary sheet showing core vs elective courses"""
    if 'course' not in dfs:
        return pd.DataFrame()
    
    sem_courses = dfs['course'][
        dfs['course']['Semester'].astype(str).str.strip() == str(semester)
    ].copy()
    
    # ENHANCED: Filter by branch if specified - only include department-specific courses and electives
    if branch and 'Department' in sem_courses.columns:
        sem_courses = sem_courses[
            ((sem_courses['Department'] == branch) & 
             (sem_courses['Elective (Yes/No)'].str.upper() != 'YES')) |
            (sem_courses['Elective (Yes/No)'].str.upper() == 'YES')
        ]
    
    if sem_courses.empty:
        return pd.DataFrame()
    
    # Add course type classification and parse LTPSC
    sem_courses['Course Type'] = sem_courses['Elective (Yes/No)'].apply(
        lambda x: 'Elective' if str(x).upper() == 'YES' else 'Core'
    )
    
    # ENHANCED: Add branch specificity info
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

def allocate_classrooms_for_timetable(schedule_df, classrooms_df, course_info):
    """Allocate classrooms to timetable sessions based on capacity and facilities"""
    print("🏫 Allocating classrooms for timetable...")
    
    if classrooms_df is None or classrooms_df.empty:
        print("   ⚠️ No classroom data available")
        return schedule_df
    
    # Filter available classrooms (exclude labs, recreation, library, etc.)
    available_classrooms = classrooms_df[
        (classrooms_df['Type'].str.contains('classroom', case=False, na=False)) |
        (classrooms_df['Type'].str.contains('auditorium', case=False, na=False))
    ].copy()
    
    # Convert capacity to numeric, handle 'nil' values
    available_classrooms['Capacity'] = pd.to_numeric(available_classrooms['Capacity'], errors='coerce')
    available_classrooms = available_classrooms.dropna(subset=['Capacity'])
    
    if available_classrooms.empty:
        print("   ⚠️ No suitable classrooms found after filtering")
        return schedule_df
    
    print(f"   Available classrooms: {len(available_classrooms)}")
    print(f"   Classroom list: {available_classrooms['Room Number'].tolist()}")
    
    # Create a copy of schedule with classroom allocation
    schedule_with_rooms = schedule_df.copy()
    
    # Track classroom usage to avoid double-booking
    classroom_usage = {}
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    
    # Initialize classroom usage tracking
    for day in days:
        classroom_usage[day] = {}
        for time_slot in schedule_df.index:
            classroom_usage[day][time_slot] = set()
    
    # Estimate student numbers for courses
    course_enrollment = estimate_course_enrollment(course_info)
    
    # Allocate classrooms for each time slot
    allocation_count = 0
    for day in days:
        for time_slot in schedule_df.index:
            course_value = schedule_df.loc[time_slot, day]
            
            # Skip free slots, lunch breaks
            if course_value in ['Free', 'LUNCH BREAK']:
                continue
            
            # Handle both regular courses and basket entries
            if isinstance(course_value, str):
                # For basket entries, use a standard enrollment
                if any(basket in course_value for basket in ['ELECTIVE_B', 'HSS_B', 'PROF_B', 'OE_B']):
                    enrollment = 40  # Standard enrollment for elective baskets
                    course_display = course_value
                else:
                    # Regular course - extract clean course code
                    clean_course_code = course_value.replace(' (Tutorial)', '')
                    enrollment = course_enrollment.get(clean_course_code, 50)
                    course_display = course_value
            else:
                continue
            
            # Find suitable classroom
            suitable_classroom = find_suitable_classroom(
                available_classrooms, enrollment, day, time_slot, classroom_usage
            )
            
            if suitable_classroom:
                # Update schedule with classroom in format "Course [Room]"
                schedule_with_rooms.loc[time_slot, day] = f"{course_display} [{suitable_classroom}]"
                classroom_usage[day][time_slot].add(suitable_classroom)
                allocation_count += 1
                print(f"      ✅ {day} {time_slot}: {course_display} → {suitable_classroom} ({enrollment} students)")
            else:
                print(f"      ⚠️  {day} {time_slot}: No classroom available for {course_display} ({enrollment} students)")
    
    print(f"   🏫 Total classroom allocations: {allocation_count}")
    return schedule_with_rooms

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
        print(f"         ⚠️ Using largest available room for {enrollment} students")
    
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

def find_suitable_classroom(classrooms_df, enrollment, day, time_slot, classroom_usage):
    """Find a suitable classroom based on capacity and availability"""
    # Filter classrooms that can accommodate the enrollment
    suitable_rooms = classrooms_df[classrooms_df['Capacity'] >= enrollment].copy()
    
    if suitable_rooms.empty:
        # If no room can accommodate, find the largest available
        suitable_rooms = classrooms_df.nlargest(1, 'Capacity')
        if suitable_rooms.empty:
            return None
    
    # Sort by capacity (prefer smallest adequate room first)
    suitable_rooms = suitable_rooms.sort_values('Capacity')
    
    # Check availability
    for _, room in suitable_rooms.iterrows():
        room_number = room['Room Number']
        
        # Check if room is already booked at this time
        if (room_number not in classroom_usage[day][time_slot]):
            return room_number
    
    return None

def allocate_classrooms_for_exams(exam_schedule_df, classrooms_df, courses_df):
    """Allocate classrooms for exam schedule based on enrollment and duration"""
    print("🏫 Allocating classrooms for exams...")
    
    # Get all available classrooms
    available_classrooms = classrooms_df.copy()
    available_classrooms['Capacity'] = pd.to_numeric(available_classrooms['Capacity'], errors='coerce')
    available_classrooms = available_classrooms.dropna(subset=['Capacity'])
    
    # Create exam schedule with classroom allocation
    exam_schedule_with_rooms = exam_schedule_df.copy()
    
    # Track classroom usage for exams
    exam_classroom_usage = {}
    
    # Group exams by date and session
    for date in exam_schedule_df['date'].unique():
        date_str = str(date)
        exam_classroom_usage[date_str] = {'Morning': set(), 'Afternoon': set()}
    
    # Estimate exam enrollment (can be enhanced with actual student data)
    exam_enrollment = estimate_exam_enrollment(exam_schedule_df, courses_df)
    
    # Allocate classrooms for each exam
    for idx, exam in exam_schedule_df.iterrows():
        if exam['status'] != 'Scheduled':
            continue
            
        course_code = exam['course_code']
        date = exam['date']
        session = exam['session']
        duration = exam.get('duration_minutes', 180)
        
        # Get enrollment estimate
        enrollment = exam_enrollment.get(course_code, 60)
        
        # Calculate number of rooms needed based on capacity
        rooms_needed = calculate_rooms_needed(enrollment, available_classrooms)
        
        allocated_rooms = []
        for room_num in rooms_needed:
            # Check if room is available
            if room_num not in exam_classroom_usage[str(date)][session]:
                allocated_rooms.append(room_num)
                exam_classroom_usage[str(date)][session].add(room_num)
        
        if allocated_rooms:
            room_str = ', '.join(allocated_rooms)
            exam_schedule_with_rooms.at[idx, 'classroom'] = room_str
            exam_schedule_with_rooms.at[idx, 'capacity_info'] = f"{enrollment} students"
            print(f"      ✅ {date} {session}: {course_code} → {room_str} ({enrollment} students)")
        else:
            exam_schedule_with_rooms.at[idx, 'classroom'] = 'NOT ALLOCATED'
            exam_schedule_with_rooms.at[idx, 'capacity_info'] = f"{enrollment} students"
            print(f"      ⚠️  {date} {session}: No classroom available for {course_code}")
    
    return exam_schedule_with_rooms

def estimate_exam_enrollment(exam_schedule_df, courses_df):
    """Estimate student enrollment for exams"""
    enrollment_estimates = {}
    
    for _, exam in exam_schedule_df.iterrows():
        if exam['status'] == 'Scheduled':
            course_code = exam['course_code']
            department = exam.get('department', 'General')
            
            # Simple estimation based on department and course type
            if department in ['CSE', 'DSAI', 'ECE']:
                enrollment_estimates[course_code] = 60  # Core department courses
            elif 'Elective' in course_code:
                enrollment_estimates[course_code] = 40  # Elective courses
            else:
                enrollment_estimates[course_code] = 50  # Default
    
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
    print("📊 Generating classroom utilization report...")
    
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
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
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
    
    def style_cell_with_classroom_and_colors(val):
        if isinstance(val, str) and val not in ['Free', 'LUNCH BREAK']:
            # Check for classroom allocation format "Course [Room]"
            if '[' in val and ']' in val:
                try:
                    # Extract course and classroom
                    course_part = val.split('[')[0].strip()
                    room_part = '[' + val.split('[')[1]  # Get everything after first [
                    
                    # Check if it's a basket
                    basket_names = ['ELECTIVE_B1', 'ELECTIVE_B2', 'ELECTIVE_B3', 'ELECTIVE_B4', 'ELECTIVE_B5', 'ELECTIVE_B6', 'ELECTIVE_B7', 'HSS_B1', 'HSS_B2']
                    is_basket = any(basket in course_part for basket in basket_names)
                    
                    if is_basket:
                        # Get basket color
                        basket_color = basket_colors.get(course_part.replace(' (Tutorial)', ''), '#cccccc')
                        if '(Tutorial)' in course_part:
                            return f'<span class="basket-entry basket-tutorial" style="background-color: {basket_color}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                        else:
                            return f'<span class="basket-entry elective-basket" style="background-color: {basket_color}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                    else:
                        # Regular course with classroom - get course color
                        clean_course = course_part.replace(' (Tutorial)', '')
                        course_color = course_colors.get(clean_course, '#cccccc')
                        return f'<span class="course-with-room" style="background-color: {course_color}">{course_part}<br><small class="classroom-info">{room_part}</small></span>'
                except:
                    return val
            
            # Handle courses without classroom allocation but with color coding
            basket_names = ['ELECTIVE_B1', 'ELECTIVE_B2', 'ELECTIVE_B3', 'ELECTIVE_B4', 'ELECTIVE_B5', 'ELECTIVE_B6', 'ELECTIVE_B7', 'HSS_B1', 'HSS_B2']
            for basket_name in basket_names:
                if basket_name in val and '(Tutorial)' not in val:
                    basket_color = basket_colors.get(basket_name, '#cccccc')
                    return f'<span class="basket-entry elective-basket" style="background-color: {basket_color}">{val}</span>'
                elif basket_name in val and '(Tutorial)' in val:
                    basket_color = basket_colors.get(basket_name.replace(' (Tutorial)', ''), '#cccccc')
                    return f'<span class="basket-entry basket-tutorial" style="background-color: {basket_color}">{val}</span>'
            
            # Regular courses without classrooms - apply course color
            if val not in ['Free', 'LUNCH BREAK']:
                clean_course = val.replace(' (Tutorial)', '')
                course_color = course_colors.get(clean_course, '#cccccc')
                return f'<span class="regular-course" style="background-color: {course_color}">{val}</span>'
                
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
        # Only look for basket timetables
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*_baskets.xlsx"))
        
        print(f"📁 Looking for BASKET timetable files in {OUTPUT_DIR}")
        print(f"📄 Found {len(excel_files)} basket Excel files: {excel_files}")
        
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
            if 'sem' in filename and 'timetable_baskets' in filename:
                try:
                    # Extract semester and branch from filename
                    if '_' in filename and filename.count('_') >= 2:
                        # Format: semX_BRANCH_timetable_baskets.xlsx
                        parts = filename.split('_')
                        sem_part = parts[0].replace('sem', '')
                        branch = parts[1]
                        sem = int(sem_part)
                    else:
                        # Legacy format: semX_timetable_baskets.xlsx
                        sem_part = filename.split('sem')[1].split('_')[0]
                        sem = int(sem_part)
                        branch = None
                    
                    print(f"📖 Reading BASKET timetable file: {filename} (Branch: {branch})")
                    
                    # Read both sections from the Excel file
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                    
                    # Check if classrooms are allocated (look for [Room] pattern in any cell)
                    has_classroom_allocation = False
                    for df in [df_a, df_b]:
                        for col in df.columns:
                            for val in df[col]:
                                if isinstance(val, str) and '[' in val and ']' in val:
                                    has_classroom_allocation = True
                                    break
                            if has_classroom_allocation:
                                break
                        if has_classroom_allocation:
                            break
                    
                    # Try to read basket allocations if available
                    basket_allocations = {}
                    basket_courses_map = {}
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
                        print(f"   ⚠️ No basket allocation sheet found in {filename}")
                    
                    # Try to read classroom allocation details
                    classroom_allocation_details = []
                    try:
                        classroom_df = pd.read_excel(file_path, sheet_name='Classroom_Allocation')
                        classroom_allocation_details = classroom_df.to_dict('records')
                        print(f"   🏫 Found classroom allocation details: {len(classroom_allocation_details)} entries")
                    except:
                        print(f"   ⚠️ No classroom allocation sheet found in {filename}")
                    
                    # Convert to HTML tables with basket-aware, classroom-aware, and color-aware processing
                    html_a = convert_dataframe_to_html_with_baskets(df_a, f"sem{sem}_{branch}_A" if branch else f"sem{sem}_A", course_colors, basket_colors, course_info)
                    html_b = convert_dataframe_to_html_with_baskets(df_b, f"sem{sem}_{branch}_B" if branch else f"sem{sem}_B", course_colors, basket_colors, course_info)
                    
                    # Extract unique courses AND baskets from the actual schedule
                    unique_courses_a, unique_baskets_a = extract_unique_courses_with_baskets(df_a, basket_allocations)
                    unique_courses_b, unique_baskets_b = extract_unique_courses_with_baskets(df_b, basket_allocations)
                    
                    # Get course basket information for this semester and branch
                    course_baskets = separate_courses_by_type(data_frames, sem, branch) if data_frames else {'core_courses': [], 'elective_courses': []}
                    
                    # ENHANCED: Build comprehensive basket courses map including ALL elective courses for this semester
                    if not course_baskets['elective_courses'].empty and 'Basket' in course_baskets['elective_courses'].columns:
                        for _, course in course_baskets['elective_courses'].iterrows():
                            basket = course.get('Basket', 'Unknown')
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
                    
                    # Add all elective courses from baskets that appear in the schedule
                    for basket_name in unique_baskets_a:
                        if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                            legend_courses_a.update(basket_courses_map[basket_name])
                    
                    for basket_name in unique_baskets_b:
                        if basket_name in basket_courses_map and basket_courses_map[basket_name]:
                            legend_courses_b.update(basket_courses_map[basket_name])
                    
                    # FIXED: Create clean basket lists without duplicates
                    clean_baskets_a = [basket for basket in unique_baskets_a if basket in basket_courses_map and basket_courses_map[basket]]
                    clean_baskets_b = [basket for basket in unique_baskets_b if basket in basket_courses_map and basket_courses_map[basket]]
                    
                    print(f"   🎨 Color coding: {len(course_colors)} course colors, {len(basket_colors)} basket colors")
                    print(f"   📊 Legend courses A: {len(legend_courses_a)}, Baskets A: {clean_baskets_a}")
                    print(f"   📊 Legend courses B: {len(legend_courses_b)}, Baskets B: {clean_baskets_b}")
                    
                    # Add timetable for Section A
                    timetables.append({
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
                        'is_basket_timetable': True,
                        'all_basket_courses': basket_courses_map,  # Include all basket courses for legends
                        'has_classroom_allocation': has_classroom_allocation,
                        'classroom_details': classroom_allocation_details
                    })
                    
                    # Add timetable for Section B
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
                        'is_basket_timetable': True,
                        'all_basket_courses': basket_courses_map,  # Include all basket courses for legends
                        'has_classroom_allocation': has_classroom_allocation,
                        'classroom_details': classroom_allocation_details
                    })
                    
                    print(f"✅ Loaded BASKET timetable: {filename}")
                    print(f"   Baskets found: {clean_baskets_a}")
                    print(f"   Classroom allocation: {'✅ YES' if has_classroom_allocation else '❌ NO'}")
                    print(f"   All basket courses: {basket_courses_map}")
                    
                except Exception as e:
                    print(f"❌ Error reading basket timetable {filename}: {e}")
                    continue
        
        print(f"📊 Total BASKET timetables loaded: {len(timetables)}")
        return jsonify(timetables)
        
    except Exception as e:
        print(f"❌ Error in /timetables: {e}")
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
        print(f"❌ Error loading stats: {e}")
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
                print(f"⚠️ Could not clear input directory: {e}")
        
        for file in files:
            if file and file.filename != '' and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(INPUT_DIR, filename)
                file.save(filepath)
                uploaded_files.append(filename)
                print(f"✅ Uploaded: {filename} -> {filepath}")
        
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
        
        print(f"📁 Available files after upload: {available_files}")
        
        for required_file in required_files:
            found = False
            required_clean = required_file.lower().replace(' ', '').replace('_', '').replace('-', '')
            
            for uploaded_file in available_files:
                uploaded_clean = uploaded_file.lower().replace(' ', '').replace('_', '').replace('-', '')
                if (required_clean in uploaded_clean or 
                    uploaded_clean in required_clean or
                    any(part in uploaded_clean for part in required_file.split('_'))):
                    found = True
                    print(f"✅ Matched {required_file} with {uploaded_file}")
                    break
            
            if not found:
                missing_files.append(required_file)
                print(f"❌ No match found for {required_file}")
        
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
                    print(f"🔄 Generating BASKET timetable for {branch} Semester {sem}...")
                    
                    # Use basket-based scheduling
                    success = export_semester_timetable_with_baskets(data_frames, sem, branch)
                    
                    filename = f"sem{sem}_{branch}_timetable_baskets.xlsx"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    
                    if success and os.path.exists(filepath):
                        success_count += 1
                        generated_files.append(filename)
                        print(f"✅ Successfully generated BASKET timetable: {filename}")
                    else:
                        print(f"❌ Basket timetable not created: {filename}")
                        
                except Exception as e:
                    print(f"❌ Error generating basket timetable for {branch} semester {sem}: {e}")
                    traceback.print_exc()

        return jsonify({
            'success': True,
            'message': f'Successfully uploaded {len(uploaded_files)} files and generated {success_count} BASKET timetables!',
            'uploaded_files': uploaded_files,
            'generated_count': success_count,
            'files': generated_files
        })
        
    except Exception as e:
        print(f"❌ Error uploading files: {e}")
        traceback.print_exc()
        return jsonify({
            'success': False,
            'message': f'Error uploading files: {str(e)}'
        }), 500


def export_semester_timetable_with_baskets_common(dfs, semester, branch, common_elective_allocations):
    """Export timetable using pre-allocated common basket slots"""
    try:
        print(f"📊 Generating timetable for Semester {semester}, Branch {branch} with COMMON basket slots...")
        
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
        
        print(f"✅ Generated: {filename} with COMMON basket slots")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False
    
@app.route('/exam-schedule', methods=['POST'])
def generate_exam_schedule():
    """Generate conflict-free exam timetable with configuration and classroom allocation"""
    try:
        print("🔄 Starting exam schedule generation with classroom allocation...")
        
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
        
        print(f"📋 Found {len(exams_df)} exams in CSV data")
        
        # Generate exam schedule with configuration
        exam_schedule = schedule_exams_conflict_free(
            exams_df, exam_period_start, exam_period_end, 
            max_exams_per_day, include_weekends
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
            print("🏫 Allocating classrooms for exams...")
            exam_schedule_with_rooms = allocate_classrooms_for_exams(
                exam_schedule, data_frames['classroom'], data_frames.get('course', pd.DataFrame())
            )
            
            # Add classroom utilization info
            classroom_usage = calculate_classroom_usage_for_exams(exam_schedule_with_rooms)
            print(f"   📊 Classroom utilization: {classroom_usage['used_rooms']} rooms used, {classroom_usage['total_sessions']} exam sessions")
        else:
            exam_schedule_with_rooms = exam_schedule
            print("⚠️  No classroom data available for exam allocation")
        
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
        print(f"❌ Error generating exam schedule: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})
    

def schedule_exams_conflict_free(exams_df, start_date, end_date, max_exams_per_day=2, 
                               include_weekends=False, department_conflict='moderate',
                               preference_weight='medium', session_balance='strict', 
                               constraints=None):
    """Generate conflict-free exam schedule with multiple exams per slot"""
    try:
        # VALIDATE max_exams_per_day parameter
        max_exams_per_day = max(1, min(4, max_exams_per_day))  # Ensure between 1-4
        
        print("📅 Generating conflict-free exam schedule with multiple exams per slot...")
        print(f"⚙️ Configuration: {max_exams_per_day} exams/slot, weekends: {include_weekends}")
        print(f"⚙️ Conflict: {department_conflict}, Preference: {preference_weight}, Balance: {session_balance}")
        
        # Set default constraints if none provided
        if constraints is None:
            constraints = {
                'departments': ['CSE', 'DSAI', 'ECE', 'Mathematics', 'Physics', 'Humanities'],
                'examTypes': ['Theory', 'Lab'],
                'rules': ['gapDays', 'sessionLimit', 'preferMorning']
            }
        
        # Define time slots
        time_slots = {
            'Morning': '09:00 - 12:00',
            'Afternoon': '14:00 - 17:00'
        }
        
        # Convert date strings to datetime objects
        exams_df = exams_df.copy()
        
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
            print("❌ No valid exam data found")
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
            print("❌ No valid dates in exam period")
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
            print(f"📋 Filtered to {len(exams_df)} exams from allowed departments: {allowed_departments}")
        
        if constraints and 'examTypes' in constraints:
            allowed_exam_types = constraints['examTypes']
            exams_df = exams_df[exams_df['Exam Type'].isin(allowed_exam_types)]
            print(f"📋 Filtered to {len(exams_df)} exams of allowed types: {allowed_exam_types}")
        
        if exams_df.empty:
            print(f"❌ No exams remain after applying constraints (had {original_count} exams)")
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
        
        print(f"📝 Processing {len(exams_df)} unique exams...")
        
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
            
            print(f"🔄 Scheduling attempt {attempt + 1}/{max_attempts}")
            
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
                    exam_slot = {
                        'course_code': exam_code,
                        'course_name': exam.get('Course Name', 'Unknown Course'),
                        'exam_type': exam_type,
                        'duration': f"{duration_hours} hours",
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
                    print(f"✅ Scheduled {exam_code} on {scheduled_date} ({scheduled_session}) - {department}")
                else:
                    failed_exams.append(exam_code)
                    print(f"❌ Failed to schedule {exam_code} - {department}")
            
            # Calculate success rate for this attempt
            success_rate = len(scheduled_exams) / len(exams_df)
            print(f"📊 Attempt {attempt + 1}: {len(scheduled_exams)}/{len(exams_df)} exams scheduled ({success_rate:.1%})")
            
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
                print(f"✅ Acceptable schedule found with {success_rate:.1%} success rate")
                break
            
            attempt += 1
        
        # Use the best schedule found
        if best_schedule:
            current_day_slots, scheduled_exams, failed_exams = best_schedule
            print(f"🏆 Using best schedule with {best_success_rate:.1%} success rate")
        else:
            print("❌ No acceptable schedule found after all attempts")
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
        
        print(f"📊 Exam scheduling completed: {len(scheduled_exams)} scheduled, {len(failed_exams)} failed")
        if failed_exams:
            print(f"❌ Failed exams: {failed_exams}")
            print(f"💡 Try increasing exam period or maximum exams per day")
        
        return schedule_df
        
    except Exception as e:
        print(f"❌ Error in exam scheduling: {e}")
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
        print(f"❌ Error saving exam schedule: {e}")
        return None
        
    except Exception as e:
        print(f"❌ Error saving exam schedule: {e}")
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

def generate_section_schedule_with_elective_baskets(dfs, semester_id, section, elective_allocations, branch=None):
    """Generate schedule with basket-based elective allocation"""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"   Generating basket-based schedule for Semester {semester_id}, Section {section}{branch_info}...")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
    
    try:
        # Get only the core courses for this specific branch
        course_baskets = separate_courses_by_type(dfs, semester_id, branch)
        core_courses = course_baskets['core_courses']
        
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        
        # Time slot structure
        morning_slots = ['09:00-10:30', '10:30-12:00']
        lunch_slots = ['12:00-13:00']
        afternoon_slots = ['13:00-14:30', '14:30-15:30', '15:30-17:00', '17:00-18:00']
        all_slots = morning_slots + lunch_slots + afternoon_slots
        
        # Create schedule template
        schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
        schedule.loc['12:00-13:00'] = 'LUNCH BREAK'

        used_slots = set()

        # Schedule elective courses using basket allocation
        if elective_allocations:
            print(f"   🎯 Applying BASKET elective allocation for Section {section}:")
            used_slots = schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section, branch)
        
        # Schedule core courses after electives
        if not core_courses.empty:
            print(f"   📚 Scheduling {len(core_courses)} core courses for Section {section}, Branch {branch}...")
            used_slots = schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, 
                                                            ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00'],
                                                            ['14:30-15:30', '17:00-18:00'], branch)
        
        return schedule
        
    except Exception as e:
        print(f"❌ Error generating basket-based schedule: {e}")
        traceback.print_exc()
        return None
    
@app.route('/exam-timetables')
def get_exam_timetables():
    """Get generated exam timetables - only shows schedules that are marked for display"""
    try:
        # Only get files that are in our display list
        exam_files_to_display = get_exam_schedule_files()
        exam_timetables = []
        
        print(f"📁 Looking for {len(exam_files_to_display)} exam schedules to display")
        
        for filename in exam_files_to_display:
            file_path = os.path.join(OUTPUT_DIR, filename)
            if not os.path.exists(file_path):
                print(f"⚠️ File not found, removing from display list: {filename}")
                remove_exam_schedule_file(filename)
                continue
                
            try:
                # Read exam schedule
                schedule_df = pd.read_excel(file_path, sheet_name='Exam_Schedule')
                
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
                    'has_classroom_allocation': has_classroom_allocation
                })
                
            except Exception as e:
                print(f"❌ Error reading {filename}: {e}")
                remove_exam_schedule_file(filename)
                continue
        
        print(f"📊 Loaded {len(exam_timetables)} exam timetables for display")
        return jsonify(exam_timetables)
        
    except Exception as e:
        print(f"❌ Error loading exam timetables: {e}")
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
                    'file_exists': True
                })
                
            except Exception as e:
                print(f"❌ Error reading {filename}: {e}")
                continue
        
        return jsonify(exam_timetables)
        
    except Exception as e:
        print(f"❌ Error loading all exam timetables: {e}")
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
        print("🔄 Starting basket-based timetable generation with COMMON slots...")
        
        # Clear existing basket timetables first
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*_baskets.xlsx"))
        for file in excel_files:
            try:
                os.remove(file)
                print(f"🗑️ Removed old basket file: {file}")
            except Exception as e:
                print(f"⚠️ Could not remove {file}: {e}")

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
                        print(f"✅ Successfully generated: {filename} with COMMON basket slots")
                        
                except Exception as e:
                    print(f"❌ Error generating basket timetable for {branch} semester {sem}: {e}")

        return jsonify({
            'success': True, 
            'message': f'Successfully generated {success_count} basket-based timetables with COMMON slots!',
            'generated_count': success_count,
            'files': generated_files
        })
        
    except Exception as e:
        print(f"❌ Error in basket generation endpoint: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})
    
if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)