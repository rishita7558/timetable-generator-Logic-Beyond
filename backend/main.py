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
    times = ['9:00-10:30', '10:30-12:00', '12:30-14:00', '14:00-15:30', '15:30-17:00', '17:00-18:30']
    
    # Create schedule template
    schedule = pd.DataFrame(index=times, columns=days, dtype=object).fillna('Free')
    schedule.loc['12:30-14:00'] = 'LUNCH BREAK'

    # Parse LTPSC if the column exists
    if 'LTPSC' in sem_courses.columns:
        try:
            sem_courses[['L', 'T', 'P', 'S', 'C']] = sem_courses['LTPSC'].str.split('-', expand=True)
            sem_courses[['L', 'T', 'P']] = sem_courses[['L', 'T', 'P']].apply(pd.to_numeric, errors='coerce').fillna(0).astype(int)
        except Exception as e:
            print(f"⚠️ Error parsing LTPSC: {e}")
            # Set default values if parsing fails
            sem_courses['L'] = 3  # Default 3 lectures per course
            sem_courses['T'] = 0
            sem_courses['P'] = 0
    else:
        print("⚠️ LTPSC column not found, using default values")
        sem_courses['L'] = 3  # Default 3 lectures per course
        sem_courses['T'] = 0
        sem_courses['P'] = 0

    used_slots = set()
    course_day_usage = {}

    # Schedule lectures
    available_times = [t for t in times if t != '12:30-14:00']
    
    total_lectures_needed = sem_courses['L'].sum()
    lectures_scheduled_total = 0
    
    for _, course in sem_courses.iterrows():
        course_code = course['Course Code']
        lectures_needed = course['L']
        course_day_usage[course_code] = set()
        lectures_scheduled = 0
        max_attempts = 100  # Increased attempts for better scheduling
        
        print(f"      Scheduling {lectures_needed} lectures for {course_code}...")
        
        while lectures_scheduled < lectures_needed and max_attempts > 0:
            max_attempts -= 1
            available_days = [d for d in days if d not in course_day_usage[course_code]]
            if not available_days:
                # Reset day usage if all days are used
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
            print(f"      ⚠️ Could only schedule {lectures_scheduled}/{lectures_needed} lectures for {course_code}")

    print(f"   ✅ Scheduled {lectures_scheduled_total}/{total_lectures_needed} total lectures")
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