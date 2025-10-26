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

# --- Core timetable generation ---
def generate_section_schedule(dfs, semester_id, section):
    print(f"   Generating schedule for Semester {semester_id}, Section {section}...")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
        
    sem_courses = dfs['course'][dfs['course']['Semester'] == semester_id].copy()
    if sem_courses.empty:
        print(f"⚠️ No courses found for semester {semester_id}")
        return None

    print(f"   Found {len(sem_courses)} courses for semester {semester_id}")

    days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri']
    # CORRECTED time slots with 1-hour lunch break:
    # Lectures: 1.5 hours (9:00-10:30, 10:30-12:00, 13:00-14:30, 14:30-16:00, 16:00-17:30)
    # Tutorials: 1 hour (12:00-13:00, 14:30-15:30, 16:00-17:00, 17:00-18:00)
    lecture_times = ['9:00-10:30', '10:30-12:00', '13:00-14:30', '14:30-16:00', '16:00-17:30']
    tutorial_times = ['12:00-13:00', '14:30-15:30', '16:00-17:00', '17:00-18:00']
    
    # Create schedule template with all time slots
    all_slots = lecture_times + tutorial_times
    schedule = pd.DataFrame(index=all_slots, columns=days, dtype=object).fillna('Free')
    # 1-hour lunch break at 12:00-13:00
    schedule.loc['12:00-13:00'] = 'LUNCH BREAK'

    used_slots = set()
    course_day_usage = {}

    # Schedule lectures and tutorials based on LTPSC
    for _, course in sem_courses.iterrows():
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
    print(f"\n📊 Generating timetable for Semester {semester}...")
    
    try:
        section_a = generate_section_schedule(dfs, semester, 'A')
        section_b = generate_section_schedule(dfs, semester, 'B')
        
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