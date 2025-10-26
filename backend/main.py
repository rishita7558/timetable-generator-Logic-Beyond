import os
import pandas as pd
import random

# Input / Output
INPUT_DIR = os.path.join(os.getcwd(), "temp_inputs")
OUTPUT_DIR = os.path.join(os.getcwd(), "output_timetables")
os.makedirs(OUTPUT_DIR, exist_ok=True)

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

# --- Load CSVs ---
def load_all_data():
    required_files = [
        "course_data.csv",
        "faculty_availability.csv",
        "classroom_data.csv",
        "student_data.csv",
        "exams_data.csv"
    ]
    dfs = {}
    
    print("📂 Loading CSV files...")
    
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
            print(f"✅ Loaded {f} ({len(dfs[key])} rows)")
            
            # Show basic info about the loaded data
            if not dfs[key].empty:
                print(f"   Columns: {list(dfs[key].columns)}")
                if 'course' in key:
                    print(f"   Courses: {len(dfs[key])}")
                
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
    
    print("✅ All CSV files loaded successfully!")
    return dfs

def separate_courses_by_type(dfs, semester_id):
    """Separate courses into core and elective baskets for a given semester"""
    if 'course' not in dfs:
        return {'core_courses': [], 'elective_courses': []}
    
    try:
        # Filter courses for the semester
        sem_courses = dfs['course'][dfs['course']['Semester'] == semester_id].copy()
        
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

def allocate_electives_to_common_slots(elective_courses, semester_id):
    """Allocate elective courses to common time slots for both sections - 2 lectures + 1 tutorial per elective"""
    common_slots = get_common_elective_slots()
    elective_allocations = {}
    
    print(f"🎯 Allocating elective courses to common slots for Semester {semester_id}...")
    print(f"   Each elective will get 2 lectures and 1 tutorial (common for both sections)")
    
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
                'for_both_sections': True
            }
            print(f"   ✅ Allocated {course_code}:")
            for i, (day, time_slot) in enumerate(lectures_allocated, 1):
                print(f"      Lecture {i}: {day} {time_slot}")
            print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
        else:
            print(f"   ❌ Could not allocate all required slots for {course_code}")
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
                    print(f"      ✅ Scheduled elective lecture {course_code} at {day} {time_slot} for Section {section}")
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
                    print(f"      ✅ Scheduled elective tutorial {course_code} at {day} {time_slot} for Section {section}")
                else:
                    print(f"      ⚠️ Common slot occupied for {course_code} tutorial at {day} {time_slot}: {schedule.loc[time_slot, day]}")
        else:
            print(f"      ❌ No allocation found for elective {course_code}")
    
    print(f"   ✅ Scheduled {elective_scheduled} elective sessions for Section {section}")
    return used_slots

# --- Core timetable generation ---
def generate_section_schedule_with_electives(dfs, semester_id, section, elective_allocations):
    """Generate schedule with pre-allocated elective slots"""
    print(f"   Generating coordinated schedule for Semester {semester_id}, Section {section}...")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
        
    # Separate courses into core and elective baskets
    course_baskets = separate_courses_by_type(dfs, semester_id)
    core_courses = course_baskets['core_courses']
    elective_courses = course_baskets['elective_courses']
    
    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    
    # LOGICAL TIME SLOT STRUCTURE
    # Morning Session
    morning_slots = [
        '09:00-10:30',  # 1.5-hour lecture
        '10:30-12:00'   # 1.5-hour lecture
    ]
    
    # Lunch Break
    lunch_slots = [
        '12:00-13:00'   # 1-hour lunch
    ]
    
    # Afternoon Session
    afternoon_slots = [
        '13:00-14:30',  # 1.5-hour lecture
        '14:30-15:30',  # 1-hour tutorial
        '15:30-17:00',  # 1.5-hour lecture
        '17:00-18:00'   # 1-hour tutorial
    ]
    
    # All time slots in chronological order
    all_slots = morning_slots + lunch_slots + afternoon_slots
    
    # Lecture slots (1.5 hours)
    lecture_times = ['09:00-10:30', '10:30-12:00', '13:00-14:30', '15:30-17:00']
    
    # Tutorial slots (1 hour)
    tutorial_times = ['14:30-15:30', '17:00-18:00']
    
    # Create schedule template
    schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
    # Mark lunch break
    schedule.loc['12:00-13:00'] = 'LUNCH BREAK'

    used_slots = set()
    course_day_usage = {}

    # Schedule elective courses FIRST to ensure they get common slots
    if not elective_courses.empty:
        print(f"   🎯 Scheduling {len(elective_courses)} elective courses for Section {section}...")
        used_slots = schedule_electives_in_common_slots(elective_allocations, schedule, used_slots, section)
    
    # Schedule core courses after electives
    if not core_courses.empty:
        print(f"   📚 Scheduling {len(core_courses)} core courses for Section {section}...")
        
        # Parse LTPSC for core courses
        for _, course in core_courses.iterrows():
            course_code = course['Course Code']
            ltpsc = parse_ltpsc(course['LTPSC'])
            lectures_needed = ltpsc['L']
            tutorials_needed = ltpsc['T']
            
            course_day_usage[course_code] = {'lectures': set(), 'tutorials': set()}
            
            print(f"      Scheduling {lectures_needed} lectures and {tutorials_needed} tutorials for {course_code}...")
            
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

    return schedule

# --- Export ---
def export_semester_timetable(dfs, semester):
    print(f"\n📊 Generating COORDINATED timetable for Semester {semester}...")
    
    try:
        # First, identify all elective courses for this semester
        course_baskets = separate_courses_by_type(dfs, semester)
        elective_courses = course_baskets['elective_courses']
        
        print(f"🎯 Elective courses found for semester {semester}: {len(elective_courses)}")
        if not elective_courses.empty:
            print("   Elective courses:", elective_courses['Course Code'].tolist())
        
        # Allocate elective courses to common slots (2 lectures + 1 tutorial each)
        common_elective_allocations = allocate_electives_to_common_slots(elective_courses, semester)
        
        # Generate schedules for both sections with coordinated electives
        section_a = generate_section_schedule_with_electives(dfs, semester, 'A', common_elective_allocations)
        section_b = generate_section_schedule_with_electives(dfs, semester, 'B', common_elective_allocations)
        
        if section_a is None or section_b is None:
            print(f"❌ Failed to generate timetable for semester {semester}")
            return

        filename = f"sem{semester}_timetable.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
        print(f"✅ Timetable saved: {filename}")
        
    except Exception as e:
        print(f"❌ Error generating timetable for semester {semester}: {e}")