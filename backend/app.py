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
                        print(f"     {i+1}. {row['Course Code'] if 'Course Code' in row else 'N/A'} - Semester: {row.get('Semester', 'N/A')}")
                
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
            
            course_info[course_code] = {
                'name': course.get('Course Name', 'Unknown Course'),
                'credits': course.get('Credits', 0),
                'type': course_type,
                'instructor': course.get('Instructor', 'Unknown'),
                'department': course.get('Department', 'Unknown'),
                'is_elective': is_elective
            }
    return course_info

def separate_courses_by_type(dfs, semester_id):
    """Separate courses into core and elective baskets for a given semester"""
    if 'course' not in dfs:
        return {'core_courses': [], 'elective_courses': []}
    
    try:
        # Filter courses for the semester
        sem_courses = dfs['course'][
            dfs['course']['Semester'].astype(str).str.strip() == str(semester_id)
        ].copy()
        
        if sem_courses.empty:
            return {'core_courses': [], 'elective_courses': []}
        
        # Separate core and elective courses
        core_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() != 'YES'
        ].copy()
        
        elective_courses = sem_courses[
            sem_courses['Elective (Yes/No)'].str.upper() == 'YES'
        ].copy()
        
        print(f"   📊 Course separation for Semester {semester_id}:")
        print(f"      Core courses: {len(core_courses)}")
        print(f"      Elective courses: {len(elective_courses)}")
        
        return {
            'core_courses': core_courses,
            'elective_courses': elective_courses
        }
        
    except Exception as e:
        print(f"⚠️ Error separating courses by type: {e}")
        return {'core_courses': [], 'elective_courses': []}

def schedule_core_courses(core_courses, schedule, used_slots, days, times):
    """Schedule core courses with proper constraints"""
    if core_courses.empty:
        return used_slots
    
    available_times = [t for t in times if t != '12:30-14:00']
    course_day_usage = {}
    
    # Parse LTPSC for core courses
    if 'LTPSC' in core_courses.columns:
        try:
            core_courses[['L', 'T', 'P', 'S', 'C']] = core_courses['LTPSC'].str.split('-', expand=True)
            core_courses[['L', 'T', 'P']] = core_courses[['L', 'T', 'P']].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
        except Exception as e:
            print(f"⚠️ Error parsing LTPSC for core courses: {e}")
            core_courses['L'] = 3
            core_courses['T'] = 0
            core_courses['P'] = 0
    else:
        core_courses['L'] = 3
        core_courses['T'] = 0
        core_courses['P'] = 0

    total_lectures_needed = core_courses['L'].sum()
    lectures_scheduled_total = 0
    
    for _, course in core_courses.iterrows():
        course_code = course['Course Code']
        lectures_needed = course['L']
        course_day_usage[course_code] = set()
        lectures_scheduled = 0
        max_attempts = 100
        
        print(f"      Scheduling {lectures_needed} core lectures for {course_code}...")
        
        while lectures_scheduled < lectures_needed and max_attempts > 0:
            max_attempts -= 1
            available_days = [d for d in days if d not in course_day_usage[course_code]]
            if not available_days:
                course_day_usage[course_code] = set()
                available_days = days.copy()
            
            day = random.choice(available_days)
            time_slot = random.choice(available_times)
            key = (day, time_slot)
            
            if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = course_code
                used_slots.add(key)
                course_day_usage[course_code].add(day)
                lectures_scheduled += 1
                lectures_scheduled_total += 1

        if lectures_scheduled < lectures_needed:
            print(f"      ⚠️ Could only schedule {lectures_scheduled}/{lectures_needed} core lectures for {course_code}")

    print(f"   ✅ Scheduled {lectures_scheduled_total}/{total_lectures_needed} core lectures")
    return used_slots

def pre_allocate_elective_slots(elective_courses):
    """Pre-allocate elective courses to common time slots for both sections"""
    common_elective_slots = [
        ('Thu', '15:30-17:00'),  # Common elective slot 1
        ('Fri', '14:00-15:30'),  # Common elective slot 2
        ('Thu', '14:00-15:30'),  # Common elective slot 3
        ('Fri', '15:30-17:00')   # Common elective slot 4
    ]
    
    allocations = {}
    
    print("🎯 Pre-allocating elective courses to common slots...")
    print(f"   Available elective slots: {common_elective_slots}")
    print(f"   Elective courses to allocate: {len(elective_courses)}")
    
    for slot_index, (day, time_slot) in enumerate(common_elective_slots):
        if slot_index < len(elective_courses):
            course = elective_courses.iloc[slot_index]
            course_code = course['Course Code']
            allocations[course_code] = {
                'day': day,
                'time_slot': time_slot,
                'slot_name': f"Elective_Slot_{slot_index + 1}"
            }
            print(f"   ✅ Allocated {course_code} to {day} {time_slot}")
        else:
            break
    
    # Handle case where we have more electives than slots
    if len(elective_courses) > len(common_elective_slots):
        print(f"   ⚠️ Warning: {len(elective_courses) - len(common_elective_slots)} elective courses cannot be scheduled due to limited slots")
        for i in range(len(common_elective_slots), len(elective_courses)):
            course = elective_courses.iloc[i]
            course_code = course['Course Code']
            allocations[course_code] = None
            print(f"   ❌ No slot available for {course_code}")
    
    return allocations

def schedule_pre_allocated_electives(elective_courses, schedule, used_slots, elective_allocations, section):
    """Schedule elective courses using pre-allocated common slots"""
    elective_scheduled = 0
    
    for _, course in elective_courses.iterrows():
        course_code = course['Course Code']
        
        if course_code in elective_allocations and elective_allocations[course_code] is not None:
            allocation = elective_allocations[course_code]
            day = allocation['day']
            time_slot = allocation['time_slot']
            key = (day, time_slot)
            
            # Check if the slot is available
            if schedule.loc[time_slot, day] == 'Free':
                schedule.loc[time_slot, day] = f"{course_code} (Elective)"
                used_slots.add(key)
                elective_scheduled += 1
                print(f"      ✅ Scheduled elective {course_code} at {day} {time_slot} for Section {section}")
            else:
                print(f"      ⚠️ Pre-allocated slot occupied for {course_code} at {day} {time_slot}: {schedule.loc[time_slot, day]}")
                # Try to find an alternative slot
                alternative_slot_found = False
                for alt_day in ['Thu', 'Fri']:
                    for alt_time in ['14:00-15:30', '15:30-17:00']:
                        alt_key = (alt_day, alt_time)
                        if alt_key not in used_slots and schedule.loc[alt_time, alt_day] == 'Free':
                            schedule.loc[alt_time, alt_day] = f"{course_code} (Elective)"
                            used_slots.add(alt_key)
                            elective_scheduled += 1
                            alternative_slot_found = True
                            print(f"      ✅ Rescheduled elective {course_code} to {alt_day} {alt_time} for Section {section}")
                            break
                    if alternative_slot_found:
                        break
        else:
            print(f"      ❌ No allocation found for elective {course_code}")
    
    print(f"   ✅ Scheduled {elective_scheduled} elective courses for Section {section}")
    return used_slots

def generate_section_schedule_with_electives(dfs, semester_id, section, elective_allocations):
    """Generate schedule with pre-allocated elective slots"""
    print(f"   Generating coordinated schedule for Semester {semester_id}, Section {section}...")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
    
    try:
        # Separate courses into core and elective baskets
        course_baskets = separate_courses_by_type(dfs, semester_id)
        core_courses = course_baskets['core_courses']
        elective_courses = course_baskets['elective_courses']
        
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
        times = ['9:00-10:30', '10:30-12:00', '12:30-14:00', '14:00-15:30', '15:30-17:00', '17:00-18:30']
        
        # Create schedule template
        schedule = pd.DataFrame(index=times, columns=days, dtype=object).fillna('Free')
        schedule.loc['12:30-14:00'] = 'LUNCH BREAK'
        
        used_slots = set()

        # Schedule core courses first
        if not core_courses.empty:
            print(f"   📚 Scheduling {len(core_courses)} core courses for Section {section}...")
            used_slots = schedule_core_courses(core_courses, schedule, used_slots, days, times)
        
        # Schedule elective courses using pre-allocated slots
        if not elective_courses.empty:
            print(f"   🎯 Scheduling {len(elective_courses)} elective courses for Section {section}...")
            used_slots = schedule_pre_allocated_electives(elective_courses, schedule, used_slots, elective_allocations, section)
        
        # Fill remaining slots with tutorials/labs if needed
        schedule = fill_remaining_slots(schedule, used_slots, core_courses, elective_courses, days, times)
        
        return schedule
        
    except Exception as e:
        print(f"❌ Error generating coordinated schedule: {e}")
        traceback.print_exc()
        return None

def fill_remaining_slots(schedule, used_slots, core_courses, elective_courses, days, times):
    """Fill remaining slots with tutorials, labs, or free periods"""
    available_times = [t for t in times if t != '12:30-14:00']
    
    # Schedule tutorials for courses that have them
    all_courses = pd.concat([core_courses, elective_courses], ignore_index=True)
    
    for _, course in all_courses.iterrows():
        course_code = course['Course Code']
        tutorials_needed = course.get('T', 0)
        
        if tutorials_needed > 0:
            tutorials_scheduled = 0
            max_attempts = 50
            
            while tutorials_scheduled < tutorials_needed and max_attempts > 0:
                max_attempts -= 1
                day = random.choice(days)
                time_slot = random.choice(available_times)
                key = (day, time_slot)
                
                if key not in used_slots and schedule.loc[time_slot, day] == 'Free':
                    schedule.loc[time_slot, day] = f"{course_code} (Tutorial)"
                    used_slots.add(key)
                    tutorials_scheduled += 1
    
    return schedule

def create_elective_summary(elective_allocations):
    """Create a summary of elective course allocations"""
    summary_data = []
    
    for course_code, allocation in elective_allocations.items():
        if allocation:
            summary_data.append({
                'Elective Course': course_code,
                'Day': allocation['day'],
                'Time Slot': allocation['time_slot'],
                'Slot Name': allocation['slot_name'],
                'Sections': 'A & B (Common Slot)'
            })
        else:
            summary_data.append({
                'Elective Course': course_code,
                'Day': 'Not Scheduled',
                'Time Slot': 'Not Scheduled', 
                'Slot Name': 'No Slot Available',
                'Sections': 'None'
            })
    
    return pd.DataFrame(summary_data)

def export_semester_timetable(dfs, semester):
    print(f"\n📊 Generating COORDINATED timetable for Semester {semester}...")
    
    try:
        # First, identify all elective courses for this semester
        course_baskets = separate_courses_by_type(dfs, semester)
        elective_courses = course_baskets['elective_courses']
        
        print(f"🎯 Elective courses found for semester {semester}: {len(elective_courses)}")
        if not elective_courses.empty:
            print("   Elective courses:", elective_courses['Course Code'].tolist())
        
        # Pre-allocate common elective slots
        common_elective_allocations = pre_allocate_elective_slots(elective_courses)
        
        # Generate schedules for both sections with coordinated electives
        section_a = generate_section_schedule_with_electives(dfs, semester, 'A', common_elective_allocations)
        section_b = generate_section_schedule_with_electives(dfs, semester, 'B', common_elective_allocations)
        
        if section_a is None or section_b is None:
            print(f"❌ Failed to generate timetable for semester {semester}")
            return False

        filename = f"sem{semester}_timetable.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
            # Add course summary sheet
            course_summary = create_course_summary(dfs, semester)
            course_summary.to_excel(writer, sheet_name='Course_Summary', index=False)
            
            # Add elective coordination sheet
            elective_summary = create_elective_summary(common_elective_allocations)
            elective_summary.to_excel(writer, sheet_name='Elective_Coordination', index=False)
            
        print(f"✅ Coordinated timetable saved: {filename}")
        
        # Verify elective coordination
        verify_elective_coordination(section_a, section_b, elective_courses)
        
        return True
        
    except Exception as e:
        print(f"❌ Error generating timetable for semester {semester}: {e}")
        return False

def verify_elective_coordination(section_a, section_b, elective_courses):
    """Verify that elective courses are scheduled in the same slots for both sections"""
    print("🔍 Verifying elective course coordination between sections...")
    
    elective_course_codes = elective_courses['Course Code'].tolist()
    coordination_issues = []
    
    for course_code in elective_course_codes:
        # Find elective in Section A
        a_slot = find_course_slot(section_a, course_code)
        # Find elective in Section B  
        b_slot = find_course_slot(section_b, course_code)
        
        if a_slot and b_slot:
            if a_slot == b_slot:
                print(f"   ✅ {course_code}: Both sections at {a_slot}")
            else:
                coordination_issues.append(f"{course_code}: Section A at {a_slot}, Section B at {b_slot}")
                print(f"   ❌ {course_code}: MISMATCH - Section A at {a_slot}, Section B at {b_slot}")
        elif a_slot and not b_slot:
            coordination_issues.append(f"{course_code}: Only in Section A at {a_slot}")
            print(f"   ❌ {course_code}: Only scheduled in Section A at {a_slot}")
        elif not a_slot and b_slot:
            coordination_issues.append(f"{course_code}: Only in Section B at {b_slot}")
            print(f"   ❌ {course_code}: Only scheduled in Section B at {b_slot}")
        else:
            coordination_issues.append(f"{course_code}: Not scheduled in either section")
            print(f"   ⚠️ {course_code}: Not scheduled in either section")
    
    if coordination_issues:
        print(f"🚨 Found {len(coordination_issues)} coordination issues")
    else:
        print("🎉 All elective courses are properly coordinated between sections!")

def find_course_slot(schedule, course_code):
    """Find the time slot for a given course in a schedule"""
    for time_slot in schedule.index:
        for day in schedule.columns:
            cell_value = schedule.loc[time_slot, day]
            if isinstance(cell_value, str) and course_code in cell_value:
                return f"{day} {time_slot}"
    return None

def create_course_summary(dfs, semester):
    """Create a summary sheet showing core vs elective courses"""
    if 'course' not in dfs:
        return pd.DataFrame()
    
    sem_courses = dfs['course'][
        dfs['course']['Semester'].astype(str).str.strip() == str(semester)
    ].copy()
    
    if sem_courses.empty:
        return pd.DataFrame()
    
    # Add course type classification
    sem_courses['Course Type'] = sem_courses['Elective (Yes/No)'].apply(
        lambda x: 'Elective' if str(x).upper() == 'YES' else 'Core'
    )
    
    summary_columns = ['Course Code', 'Course Name', 'Course Type', 'LTPSC', 'Credits', 'Instructor']
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
                # Handle both regular courses and elective marked courses
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

        # Generate timetables for semesters 1,3,5,7
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        for sem in target_semesters:
            try:
                print(f"🔄 Generating timetable for semester {sem}...")
                success = export_semester_timetable(data_frames, sem)
                filename = f"sem{sem}_timetable.xlsx"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                if success and os.path.exists(filepath):
                    success_count += 1
                    generated_files.append(filename)
                    print(f"✅ Successfully generated: {filename}")
                else:
                    print(f"❌ File not created: {filename}")
                    
            except Exception as e:
                print(f"❌ Error generating timetable for semester {sem}: {e}")

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
                    sem_part = filename.split('sem')[1].split('_')[0]
                    sem = int(sem_part)
                    
                    print(f"📖 Reading timetable file: {filename}")
                    
                    # Read both sections from the Excel file
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                    
                    # Convert to HTML tables with proper formatting and color coding
                    html_a = df_a.to_html(
                        classes='timetable-table', 
                        index=False, 
                        escape=False,
                        border=0,
                        table_id=f"sem{sem}_A"
                    )
                    html_b = df_b.to_html(
                        classes='timetable-table', 
                        index=False, 
                        escape=False,
                        border=0,
                        table_id=f"sem{sem}_B"
                    )
                    
                    # Clean up the HTML tables
                    html_a = clean_table_html(html_a)
                    html_b = clean_table_html(html_b)
                    
                    # Extract unique courses for this timetable
                    unique_courses_a = extract_unique_courses(df_a)
                    unique_courses_b = extract_unique_courses(df_b)
                    
                    # Get course basket information for this semester
                    course_baskets = separate_courses_by_type(data_frames, sem) if data_frames else {'core_courses': [], 'elective_courses': []}
                    
                    timetables.append({
                        'semester': sem,
                        'section': 'A',
                        'filename': filename,
                        'html': html_a,
                        'courses': unique_courses_a,
                        'course_info': course_info,
                        'course_colors': course_colors,
                        'core_courses': course_baskets['core_courses']['Course Code'].tolist() if not course_baskets['core_courses'].empty else [],
                        'elective_courses': course_baskets['elective_courses']['Course Code'].tolist() if not course_baskets['elective_courses'].empty else []
                    })
                    
                    timetables.append({
                        'semester': sem,
                        'section': 'B',
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
                'missing_files': missing_files,
                'available_files': available_files
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
        
        # Generate timetables for semesters 1,3,5,7
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        for sem in target_semesters:
            try:
                print(f"🔄 Generating timetable for semester {sem}...")
                success = export_semester_timetable(data_frames, sem)
                filename = f"sem{sem}_timetable.xlsx"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                if success and os.path.exists(filepath):
                    success_count += 1
                    generated_files.append(filename)
                    print(f"✅ Successfully generated: {filename}")
                else:
                    print(f"❌ File not created: {filename}")
                    
            except Exception as e:
                print(f"❌ Error generating timetable for semester {sem}: {e}")
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

@app.route('/debug/files')
def debug_files():
    """Debug endpoint to check what files are available"""
    try:
        input_files = []
        if os.path.exists(INPUT_DIR):
            input_files = os.listdir(INPUT_DIR)
        
        output_files = []
        if os.path.exists(OUTPUT_DIR):
            output_files = os.listdir(OUTPUT_DIR)
        
        return jsonify({
            'input_dir': INPUT_DIR,
            'input_files': input_files,
            'output_dir': OUTPUT_DIR,
            'output_files': output_files,
            'input_dir_exists': os.path.exists(INPUT_DIR),
            'output_dir_exists': os.path.exists(OUTPUT_DIR),
            'cached_data': _cached_data_frames is not None,
            'file_hashes': _file_hashes
        })
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/debug/courses')
def debug_courses():
    """Debug endpoint to check course data"""
    try:
        data_frames = load_all_data(force_reload=True)
        if data_frames is None or 'course' not in data_frames:
            return jsonify({'error': 'No course data available'})
        
        course_df = data_frames['course']
        return jsonify({
            'columns': list(course_df.columns),
            'semesters': course_df['Semester'].unique().tolist() if 'Semester' in course_df.columns else [],
            'total_courses': len(course_df),
            'sample_data': course_df.head(10).to_dict('records')
        })
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/debug/current-data')
def debug_current_data():
    """Debug endpoint to check currently loaded data"""
    try:
        data_frames = load_all_data(force_reload=True)
        if data_frames is None:
            return jsonify({'error': 'No data available'})
        
        debug_info = {}
        for key, df in data_frames.items():
            debug_info[key] = {
                'shape': df.shape,
                'columns': list(df.columns),
                'sample_data': df.head(3).to_dict('records')
            }
        
        return jsonify(debug_info)
    except Exception as e:
        return jsonify({'error': str(e)})

@app.route('/debug/clear-cache')
def debug_clear_cache():
    """Debug endpoint to clear cache"""
    global _cached_data_frames
    global _file_hashes
    _cached_data_frames = None
    _file_hashes = {}
    return jsonify({'success': True, 'message': 'Cache cleared'})

@app.route('/debug/reset-all')
def debug_reset_all():
    """Debug endpoint to reset everything"""
    global _cached_data_frames
    global _file_hashes
    
    # Clear input directory
    if os.path.exists(INPUT_DIR):
        shutil.rmtree(INPUT_DIR)
        os.makedirs(INPUT_DIR, exist_ok=True)
    
    # Clear output directory  
    if os.path.exists(OUTPUT_DIR):
        shutil.rmtree(OUTPUT_DIR)
        os.makedirs(OUTPUT_DIR, exist_ok=True)
    
    # Clear cache
    _cached_data_frames = None
    _file_hashes = {}
    
    return jsonify({'success': True, 'message': 'Complete reset done'})

@app.route('/debug')
def debug_page():
    return render_template('debug.html')

if __name__ == '__main__':
    print("🚀 Starting Timetable Generator Web Application...")
    print(f"🌐 Access the application at: http://127.0.0.1:5000")
    print(f"📁 Input directory: {INPUT_DIR}")
    print(f"📁 Output directory: {OUTPUT_DIR}")
    print("🔧 Debug endpoints available:")
    print("   /debug/files - Check uploaded files")
    print("   /debug/courses - Check course data") 
    print("   /debug/current-data - Check loaded data")
    print("   /debug/clear-cache - Clear cache")
    print("   /debug/reset-all - Reset everything")
    print("   /debug - Debug interface")
    
    app.run(debug=True, host='127.0.0.1', port=5000)