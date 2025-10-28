from flask import Flask, render_template, request, jsonify, send_file
import os
import pandas as pd
import random
import zipfile
import glob
from werkzeug.utils import secure_filename
import traceback
import shutil
import time
import hashlib

app = Flask(__name__)

# Configuration
INPUT_DIR = os.path.join(os.getcwd(), "temp_inputs")
OUTPUT_DIR = os.path.join(os.getcwd(), "output_timetables")
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Cache variables with file hashes to detect changes
_cached_data_frames = None
_cached_timestamp = 0
_file_hashes = {}

# Allowed file extensions
ALLOWED_EXTENSIONS = {'csv'}

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
    """Extract course information from course data for frontend display"""
    course_info = {}
    if 'course' in dfs:
        for _, course in dfs['course'].iterrows():
            course_code = course['Course Code']
            is_elective = course.get('Elective (Yes/No)', 'No').upper() == 'YES'
            course_type = 'Elective' if is_elective else 'Core'
            department = course.get('Department', 'General')
            
            course_info[course_code] = {
                'name': course.get('Course Name', 'Unknown Course'),
                'credits': course.get('Credits', 0),
                'type': course_type,
                'instructor': course.get('Instructor', 'Unknown'),
                'department': department,
                'is_elective': is_elective,
                'branch': department,  # Use department as branch for compatibility
                'is_common_elective': is_elective
            }
    return course_info

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

def schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, lecture_times, tutorial_times, branch=None):
    """Schedule core courses with proper LTPSC structure - only department-specific courses"""
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
    
    # Parse LTPSC for core courses
    for _, course in dept_core_courses.iterrows():
        course_code = course['Course Code']
        ltpsc = parse_ltpsc(course['LTPSC'])
        lectures_needed = ltpsc['L']
        tutorials_needed = ltpsc['T']
        
        course_day_usage[course_code] = {'lectures': set(), 'tutorials': set()}
        
        print(f"      Scheduling {lectures_needed} lectures and {tutorials_needed} tutorials for {course_code} (Department: {course.get('Department', 'General')})...")
        
        # Schedule lectures (1.5 hours each)
        lectures_scheduled = 0
        max_lecture_attempts = 100
        
        while lectures_scheduled < lectures_needed and max_lecture_attempts > 0:
            max_lecture_attempts -= 1
            available_days = [d for d in days if d not in course_day_usage[course_code]['lectures']]
            if not available_days:
                course_day_usage[course_code]['lectures'] = set()
                available_days = days.copy()
            
            day = random.choice(available_days)
            time_slot = random.choice(lecture_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = course_code
                used_slots.add(key)
                course_day_usage[course_code]['lectures'].add(day)
                lectures_scheduled += 1
        
        # Schedule tutorials (1 hour each)
        tutorials_scheduled = 0
        max_tutorial_attempts = 50
        
        while tutorials_scheduled < tutorials_needed and max_tutorial_attempts > 0:
            max_tutorial_attempts -= 1
            available_days = [d for d in days if d not in course_day_usage[course_code]['tutorials']]
            if not available_days:
                course_day_usage[course_code]['tutorials'] = set()
                available_days = days.copy()
            
            day = random.choice(available_days)
            time_slot = random.choice(tutorial_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                used_slots.add(key)
                course_day_usage[course_code]['tutorials'].add(day)
                tutorials_scheduled += 1
        
        if lectures_scheduled < lectures_needed:
            print(f"      ⚠️ Could only schedule {lectures_scheduled}/{lectures_needed} lectures for {course_code}")
        if tutorials_scheduled < tutorials_needed:
            print(f"      ⚠️ Could only schedule {tutorials_scheduled}/{tutorials_needed} tutorials for {course_code}")

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

def allocate_electives_to_common_slots(elective_courses, semester_id, branch=None):
    """Allocate elective courses to common time slots for all branches and sections - 2 lectures + 1 tutorial per elective"""
    common_slots = get_common_elective_slots()
    elective_allocations = {}
    
    branch_info = f" for Branch {branch}" if branch else ""
    print(f"🎯 Allocating elective courses to COMMON SLOTS for ALL BRANCHES - Semester {semester_id}{branch_info}...")
    print(f"   Each elective will get 2 lectures and 1 tutorial (common for ALL branches and sections)")
    
    # Separate lecture and tutorial slots
    lecture_slots = [slot for slot in common_slots if any(time in slot[1] for time in ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00'])]
    tutorial_slots = [slot for slot in common_slots if any(time in slot[1] for time in ['14:30-15:30', '17:00-18:00'])]
    
    # Shuffle to get random allocation
    random.shuffle(lecture_slots)
    random.shuffle(tutorial_slots)
    
    lecture_idx = 0
    tutorial_idx = 0
    
    for _, course in elective_courses.iterrows():
        course_code = course['Course Code']
        
        # Allocate 2 lectures and 1 tutorial for each elective
        lectures_allocated = []
        tutorial_allocated = None
        
        # Allocate 2 lectures
        for _ in range(2):
            if lecture_idx < len(lecture_slots):
                lectures_allocated.append(lecture_slots[lecture_idx])
                lecture_idx += 1
            else:
                print(f"   ⚠️ Not enough lecture slots available for {course_code}")
        
        # Allocate 1 tutorial
        if tutorial_idx < len(tutorial_slots):
            tutorial_allocated = tutorial_slots[tutorial_idx]
            tutorial_idx += 1
        else:
            print(f"   ⚠️ Not enough tutorial slots available for {course_code}")
        
        if len(lectures_allocated) == 2 and tutorial_allocated:
            elective_allocations[course_code] = {
                'lectures': lectures_allocated,
                'tutorial': tutorial_allocated,
                'for_all_branches': True,
                'for_both_sections': True
            }
            print(f"   ✅ Allocated COMMON elective {course_code} for ALL BRANCHES:")
            for i, (day, time_slot) in enumerate(lectures_allocated, 1):
                print(f"      Lecture {i}: {day} {time_slot}")
            print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
        else:
            print(f"   ❌ Could not allocate all required slots for elective {course_code}")
            elective_allocations[course_code] = None
    
    return elective_allocations

def schedule_electives_in_common_slots(elective_allocations, schedule, used_slots, section):
    """Schedule elective courses in their pre-allocated common slots"""
    elective_scheduled = 0
    
    for course_code, allocation in elective_allocations.items():
        if allocation is not None:
            # Schedule lectures
            for day, time_slot in allocation['lectures']:
                key = (day, time_slot)
                
                # Check if the slot is available
                if schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = f"{course_code} (Elective)"
                    used_slots.add(key)
                    elective_scheduled += 1
                    print(f"      ✅ Scheduled COMMON elective lecture {course_code} at {day} {time_slot} for Section {section}")
                else:
                    print(f"      ⚠️ Common slot occupied for {course_code} at {day} {time_slot}: {schedule.loc[time_slot, day]}")
            
            # Schedule tutorial
            if allocation['tutorial']:
                day, time_slot = allocation['tutorial']
                key = (day, time_slot)
                
                if schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                    used_slots.add(key)
                    elective_scheduled += 1
                    print(f"      ✅ Scheduled COMMON elective tutorial {course_code} at {day} {time_slot} for Section {section}")
                else:
                    print(f"      ⚠️ Common slot occupied for {course_code} tutorial at {day} {time_slot}: {schedule.loc[time_slot, day]}")
        else:
            print(f"      ❌ No allocation found for elective {course_code}")
    
    print(f"   ✅ Scheduled {elective_scheduled} COMMON elective sessions for Section {section}")
    return used_slots

def generate_section_schedule_with_electives(dfs, semester_id, section, elective_allocations, branch=None):
    """Generate schedule with pre-allocated elective slots and branch-specific core courses"""
    branch_info = f", Branch {branch}" if branch else ""
    print(f"   Generating coordinated schedule for Semester {semester_id}, Section {section}{branch_info}...")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
    
    try:
        # Separate courses into core and elective baskets
        course_baskets = separate_courses_by_type(dfs, semester_id, branch)
        core_courses = course_baskets['core_courses']
        elective_courses = course_baskets['elective_courses']
        
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

        # Schedule elective courses FIRST to ensure they get common slots
        if not elective_courses.empty and elective_allocations:
            print(f"   🎯 Scheduling {len(elective_courses)} elective courses for Section {section} (COMMON for ALL BRANCHES)...")
            used_slots = schedule_electives_in_common_slots(elective_allocations, schedule, used_slots, section)
        
        # Schedule core courses after electives - these will be department-specific
        if not core_courses.empty:
            print(f"   📚 Scheduling {len(core_courses)} core courses for Section {section}, Branch {branch}...")
            used_slots = schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, lecture_times, tutorial_times, branch)
        
        return schedule
        
    except Exception as e:
        print(f"❌ Error generating coordinated schedule: {e}")
        traceback.print_exc()
        return None

def create_elective_summary(elective_allocations):
    """Create a summary of elective course allocations"""
    summary_data = []
    
    for course_code, allocation in elective_allocations.items():
        if allocation:
            # Add lecture allocations
            for i, (day, time_slot) in enumerate(allocation['lectures'], 1):
                summary_data.append({
                    'Elective Course': course_code,
                    'Session Type': f'Lecture {i}',
                    'Day': day,
                    'Time Slot': time_slot,
                    'Duration': '1.5 hours',
                    'Sections': 'A & B (Common Slot)',
                    'Branches': 'ALL (Common for All Branches)'  # NEW: Show common for all branches
                })
            
            # Add tutorial allocation
            if allocation['tutorial']:
                day, time_slot = allocation['tutorial']
                summary_data.append({
                    'Elective Course': course_code,
                    'Session Type': 'Tutorial',
                    'Day': day,
                    'Time Slot': time_slot,
                    'Duration': '1 hour',
                    'Sections': 'A & B (Common Slot)',
                    'Branches': 'ALL (Common for All Branches)'  # NEW: Show common for all branches
                })
        else:
            summary_data.append({
                'Elective Course': course_code,
                'Session Type': 'Not Scheduled',
                'Day': 'Not Scheduled',
                'Time Slot': 'Not Scheduled', 
                'Duration': 'N/A',
                'Sections': 'None',
                'Branches': 'None'
            })
    
    return pd.DataFrame(summary_data)

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

def create_elective_summary(elective_allocations):
    """Create a summary of elective course allocations"""
    summary_data = []
    
    for course_code, allocation in elective_allocations.items():
        if allocation:
            # Add lecture allocations
            for i, (day, time_slot) in enumerate(allocation['lectures'], 1):
                summary_data.append({
                    'Elective Course': course_code,
                    'Session Type': f'Lecture {i}',
                    'Day': day,
                    'Time Slot': time_slot,
                    'Duration': '1.5 hours',
                    'Sections': 'A & B (Common Slot)',
                    'Branches': 'ALL (Common for All Branches)'
                })
            
            # Add tutorial allocation
            if allocation['tutorial']:
                day, time_slot = allocation['tutorial']
                summary_data.append({
                    'Elective Course': course_code,
                    'Session Type': 'Tutorial',
                    'Day': day,
                    'Time Slot': time_slot,
                    'Duration': '1 hour',
                    'Sections': 'A & B (Common Slot)',
                    'Branches': 'ALL (Common for All Branches)'
                })
        else:
            summary_data.append({
                'Elective Course': course_code,
                'Session Type': 'Not Scheduled',
                'Day': 'Not Scheduled',
                'Time Slot': 'Not Scheduled', 
                'Duration': 'N/A',
                'Sections': 'None',
                'Branches': 'None'
            })
    
    return pd.DataFrame(summary_data)

def export_semester_timetable(dfs, semester, branch=None):
    branch_info = f", Branch {branch}" if branch else ""
    print(f"\n📊 Generating COORDINATED timetable for Semester {semester}{branch_info}...")
    
    try:
        # First, identify all elective courses for this semester (common for all branches)
        # ENHANCED: Get electives without branch filter to get ALL electives for the semester
        course_baskets_all = separate_courses_by_type(dfs, semester)
        elective_courses_all = course_baskets_all['elective_courses']
        
        print(f"🎯 Elective courses found for semester {semester} (COMMON for ALL branches): {len(elective_courses_all)}")
        if not elective_courses_all.empty:
            print("   Common elective courses:", elective_courses_all['Course Code'].tolist())
        
        # Allocate elective courses to common slots (2 lectures + 1 tutorial each) - COMMON FOR ALL BRANCHES
        common_elective_allocations = {}
        if not elective_courses_all.empty:
            common_elective_allocations = allocate_electives_to_common_slots(elective_courses_all, semester, branch)
        
        # Generate schedules for both sections with coordinated electives and branch-specific cores
        section_a = generate_section_schedule_with_electives(dfs, semester, 'A', common_elective_allocations, branch)
        section_b = generate_section_schedule_with_electives(dfs, semester, 'B', common_elective_allocations, branch)
        
        if section_a is None or section_b is None:
            print(f"❌ Failed to generate timetable for semester {semester}{branch_info}")
            return False

        # Create filename with branch information
        if branch:
            filename = f"sem{semester}_{branch}_timetable.xlsx"
        else:
            filename = f"sem{semester}_timetable.xlsx"
            
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
            # Add course summary sheet
            course_summary = create_course_summary(dfs, semester, branch)
            if not course_summary.empty:
                course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
            
            # Add elective coordination sheet
            if common_elective_allocations:
                elective_summary = create_elective_summary(common_elective_allocations)
                elective_summary.to_excel(writer, sheet_name='Elective_Coordination', index=False)
            
            # Add branch-specific info sheet
            branch_info_sheet = create_branch_info_sheet(dfs, semester, branch)
            if not branch_info_sheet.empty:
                branch_info_sheet.to_excel(writer, sheet_name='Branch_Info', index=False)
            
        print(f"✅ Coordinated timetable saved: {filename}")
        return True
        
    except Exception as e:
        print(f"❌ Error generating timetable for semester {semester}{branch_info}: {e}")
        traceback.print_exc()
        return False
    

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

@app.route('/generate', methods=['POST'])
def generate_timetables():
    try:
        print("🔄 Starting timetable generation...")
        
        # Clear existing timetables first
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        for file in excel_files:
            try:
                os.remove(file)
                print(f"🗑️ Removed old file: {file}")
            except Exception as e:
                print(f"⚠️ Could not remove {file}: {e}")

        # Load data and generate timetables - ALWAYS force reload
        print("📂 Forcing data reload for generation...")
        data_frames = load_all_data(force_reload=True)
        if data_frames is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'})

        # Generate timetables for all branches and semesters
        departments = get_departments_from_data(data_frames)
        print(f"📋 Generating timetables for departments: {departments}")
        
        branches = departments
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        for branch in branches:
            for sem in target_semesters:
                try:
                    print(f"🔄 Generating timetable for {branch} Semester {sem}...")
                    success = export_semester_timetable(data_frames, sem, branch)
                    filename = f"sem{sem}_{branch}_timetable.xlsx"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    
                    if success and os.path.exists(filepath):
                        success_count += 1
                        generated_files.append(filename)
                        print(f"✅ Successfully generated: {filename}")
                    else:
                        print(f"❌ File not created: {filename}")
                        
                except Exception as e:
                    print(f"❌ Error generating timetable for {branch} semester {sem}: {e}")

        return jsonify({
            'success': True, 
            'message': f'Successfully generated {success_count} timetables!',
            'generated_count': success_count,
            'files': generated_files
        })
        
    except Exception as e:
        print(f"❌ Error in generate endpoint: {e}")
        traceback.print_exc()
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/timetables')
def get_timetables():
    try:
        timetables = []
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        
        print(f"📁 Looking for timetable files in {OUTPUT_DIR}")
        print(f"📄 Found {len(excel_files)} Excel files: {excel_files}")
        
        # Load course data for course information - force reload to get latest data
        data_frames = load_all_data(force_reload=True)
        course_info = get_course_info(data_frames) if data_frames else {}
        
        # Generate colors for all courses
        all_courses = set()
        if data_frames and 'course' in data_frames:
            all_courses = set(data_frames['course']['Course Code'].unique())
        course_colors = generate_course_colors(all_courses, course_info)
        
        for file_path in excel_files:
            filename = os.path.basename(file_path)
            if 'sem' in filename and 'timetable' in filename:
                try:
                    # Extract semester and branch from filename
                    if '_' in filename and filename.count('_') >= 2:
                        # Format: semX_BRANCH_timetable.xlsx
                        parts = filename.split('_')
                        sem_part = parts[0].replace('sem', '')
                        branch = parts[1]
                        sem = int(sem_part)
                    else:
                        # Legacy format: semX_timetable.xlsx
                        sem_part = filename.split('sem')[1].split('_')[0]
                        sem = int(sem_part)
                        branch = None
                    
                    print(f"📖 Reading timetable file: {filename} (Branch: {branch})")
                    
                    # Read both sections from the Excel file
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                    
                    # Convert to HTML tables with proper formatting and color coding
                    html_a = df_a.to_html(
                        classes='timetable-table', 
                        index=False, 
                        escape=False,
                        border=0,
                        table_id=f"sem{sem}_{branch}_A" if branch else f"sem{sem}_A"
                    )
                    html_b = df_b.to_html(
                        classes='timetable-table', 
                        index=False, 
                        escape=False,
                        border=0,
                        table_id=f"sem{sem}_{branch}_B" if branch else f"sem{sem}_B"
                    )
                    
                    # Clean up the HTML tables
                    html_a = clean_table_html(html_a)
                    html_b = clean_table_html(html_b)
                    
                    # Extract unique courses for this timetable
                    unique_courses_a = extract_unique_courses(df_a)
                    unique_courses_b = extract_unique_courses(df_b)
                    
                    # Get course basket information for this semester and branch
                    course_baskets = separate_courses_by_type(data_frames, sem, branch) if data_frames else {'core_courses': [], 'elective_courses': []}
                    
                    # Add timetable for Section A
                    timetables.append({
                        'semester': sem,
                        'section': 'A',
                        'branch': branch,
                        'filename': filename,
                        'html': html_a,
                        'courses': unique_courses_a,
                        'course_info': course_info,
                        'course_colors': course_colors,
                        'core_courses': course_baskets['core_courses']['Course Code'].tolist() if not course_baskets['core_courses'].empty else [],
                        'elective_courses': course_baskets['elective_courses']['Course Code'].tolist() if not course_baskets['elective_courses'].empty else []
                    })
                    
                    # Add timetable for Section B
                    timetables.append({
                        'semester': sem,
                        'section': 'B',
                        'branch': branch,
                        'filename': filename,
                        'html': html_b,
                        'courses': unique_courses_b,
                        'course_info': course_info,
                        'course_colors': course_colors,
                        'core_courses': course_baskets['core_courses']['Course Code'].tolist() if not course_baskets['core_courses'].empty else [],
                        'elective_courses': course_baskets['elective_courses']['Course Code'].tolist() if not course_baskets['elective_courses'].empty else []
                    })
                    
                    print(f"✅ Loaded timetable: {filename}")
                    
                except Exception as e:
                    print(f"❌ Error reading {filename}: {e}")
                    continue
        
        print(f"📊 Total timetables loaded: {len(timetables)}")
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
        
        # Clear any cached data to force reload
        global _cached_data_frames
        if '_cached_data_frames' in globals():
            _cached_data_frames = None
            print("🗑️ Cleared cached data frames")
        
        # Load data with force reload
        print("🔄 Loading data with force reload...")
        data_frames = load_all_data(force_reload=True)
        if data_frames is None:
            return jsonify({
                'success': False,
                'message': 'Files uploaded but failed to load data for timetable generation'
            }), 400
        
        # Clear existing timetables
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        for file in excel_files:
            try:
                os.remove(file)
                print(f"🗑️ Removed old timetable: {file}")
            except Exception as e:
                print(f"⚠️ Could not remove {file}: {e}")
        
        # Generate timetables for all branches and semesters
        branches = ['CSE', 'DSAI', 'ECE']
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        for branch in branches:
            for sem in target_semesters:
                try:
                    print(f"🔄 Generating timetable for {branch} Semester {sem}...")
                    success = export_semester_timetable(data_frames, sem, branch)
                    filename = f"sem{sem}_{branch}_timetable.xlsx"
                    filepath = os.path.join(OUTPUT_DIR, filename)
                    
                    if success and os.path.exists(filepath):
                        success_count += 1
                        generated_files.append(filename)
                        print(f"✅ Successfully generated: {filename}")
                    else:
                        print(f"❌ File not created: {filename}")
                        
                except Exception as e:
                    print(f"❌ Error generating timetable for {branch} semester {sem}: {e}")
                    traceback.print_exc()

        print("=" * 50)
        print("✅ UPLOAD AND GENERATION COMPLETED")
        print(f"📁 Uploaded: {len(uploaded_files)} files")
        print(f"📊 Generated: {success_count} timetables")
        print("=" * 50)

        return jsonify({
            'success': True,
            'message': f'Successfully uploaded {len(uploaded_files)} files and generated {success_count} timetables!',
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

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)