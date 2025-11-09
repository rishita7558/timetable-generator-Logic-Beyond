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

def allocate_electives_by_baskets(elective_courses, semester_id):
    """Allocate elective courses based on basket groups with common time slots"""
    print(f"🎯 Allocating elective courses by baskets for Semester {semester_id}...")
    
    # Group electives by basket
    basket_groups = {}
    for _, course in elective_courses.iterrows():
        basket = course.get('Basket', 'ELECTIVE_B1')  # Default basket if not specified
        if basket not in basket_groups:
            basket_groups[basket] = []
        basket_groups[basket].append(course)
    
    print(f"   Found {len(basket_groups)} elective baskets: {list(basket_groups.keys())}")
    
    # Define available time slots for baskets (2 lectures + 1 tutorial per basket)
    basket_lecture_slots = [
        ('Mon', '09:00-10:30'), ('Mon', '13:00-14:30'), ('Mon', '15:30-17:00'),
        ('Tue', '09:00-10:30'), ('Tue', '13:00-14:30'), ('Tue', '15:30-17:00'),
        ('Wed', '09:00-10:30'), ('Wed', '13:00-14:30'), ('Wed', '15:30-17:00'),
        ('Thu', '09:00-10:30'), ('Thu', '13:00-14:30'), ('Thu', '15:30-17:00'),
        ('Fri', '09:00-10:30'), ('Fri', '13:00-14:30'), ('Fri', '15:30-17:00')
    ]
    
    basket_tutorial_slots = [
        ('Mon', '14:30-15:30'), ('Mon', '17:00-18:00'),
        ('Tue', '14:30-15:30'), ('Tue', '17:00-18:00'),
        ('Wed', '14:30-15:30'), ('Wed', '17:00-18:00'),
        ('Thu', '14:30-15:30'), ('Thu', '17:00-18:00'),
        ('Fri', '14:30-15:30'), ('Fri', '17:00-18:00')
    ]
    
    elective_allocations = {}
    basket_allocations = {}
    
    # Use FIXED allocation (don't shuffle) to ensure consistency
    lecture_idx = 0
    tutorial_idx = 0
    
    for basket_name in sorted(basket_groups.keys()):  # Sort for consistent ordering
        basket_courses = basket_groups[basket_name]
        course_codes = [course['Course Code'] for course in basket_courses]
        
        # Allocate 2 lectures and 1 tutorial for this basket
        lectures_allocated = []
        tutorial_allocated = None
        
        # Allocate 2 lectures (1.5 hours each)
        for _ in range(2):
            if lecture_idx < len(basket_lecture_slots):
                lectures_allocated.append(basket_lecture_slots[lecture_idx])
                lecture_idx += 1
            else:
                print(f"   ⚠️ Not enough lecture slots available for basket '{basket_name}'")
        
        # Allocate 1 tutorial (1 hour)
        if tutorial_idx < len(basket_tutorial_slots):
            tutorial_allocated = basket_tutorial_slots[tutorial_idx]
            tutorial_idx += 1
        else:
            print(f"   ⚠️ Not enough tutorial slots available for basket '{basket_name}'")
        
        if len(lectures_allocated) == 2 and tutorial_allocated:
            # Store allocation for all courses in this basket
            for course_code in course_codes:
                elective_allocations[course_code] = {
                    'basket_name': basket_name,
                    'lectures': lectures_allocated,
                    'tutorial': tutorial_allocated,
                    'all_courses_in_basket': course_codes,
                    'for_all_branches': True,
                    'for_both_sections': True,
                    'common_for_semester': True
                }
            
            basket_allocations[basket_name] = {
                'lectures': lectures_allocated,
                'tutorial': tutorial_allocated,
                'courses': course_codes
            }
            
            print(f"   🗂️ Basket '{basket_name}' allocated:")
            for i, (day, time_slot) in enumerate(lectures_allocated, 1):
                print(f"      Lecture {i}: {day} {time_slot}")
            print(f"      Tutorial: {tutorial_allocated[0]} {tutorial_allocated[1]}")
            print(f"      Courses: {', '.join(course_codes)}")
        else:
            print(f"   ❌ Could not allocate all required slots for basket '{basket_name}'")
    
    return elective_allocations, basket_allocations

def schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section):
    """Schedule elective courses based on basket allocations"""
    elective_scheduled = 0
    
    # Track which basket slots we've already scheduled (to avoid duplicates)
    scheduled_basket_slots = set()
    
    for course_code, allocation in elective_allocations.items():
        if allocation is None:
            continue
            
        basket_name = allocation['basket_name']
        lectures = allocation['lectures']
        tutorial = allocation['tutorial']
        all_courses = allocation['all_courses_in_basket']
        
        # Format basket display with all courses
        basket_display = f"Basket: {basket_name}"
        tutorial_display = f"Basket: {basket_name} (Tutorial)"
        
        # Schedule lectures
        for day, time_slot in lectures:
            slot_key = (basket_name, day, time_slot, 'lecture')
            
            # Skip if we've already scheduled this basket lecture slot
            if slot_key in scheduled_basket_slots:
                continue
                
            key = (day, time_slot)
            
            if schedule.loc[time_slot, day] == 'Free':
                # Schedule as basket with clear basket name
                schedule.loc[time_slot, day] = basket_display
                used_slots.add(key)
                scheduled_basket_slots.add(slot_key)
                elective_scheduled += 1
                
                print(f"      ✅ BASKET LECTURE '{basket_name}' at {day} {time_slot} - Section {section}")
                print(f"         Courses: {', '.join(all_courses)}")
            else:
                print(f"      ❌ LECTURE SLOT CONFLICT: Basket '{basket_name}' at {day} {time_slot} already has: {schedule.loc[time_slot, day]}")
        
        # Schedule tutorial
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
                    
                    print(f"      ✅ BASKET TUTORIAL '{basket_name}' at {day} {time_slot} - Section {section}")
                else:
                    print(f"      ❌ TUTORIAL SLOT CONFLICT: Basket '{basket_name}' at {day} {time_slot} already has: {schedule.loc[time_slot, day]}")
    
    print(f"   ✅ Scheduled {elective_scheduled} elective basket sessions for Section {section}")
    return used_slots

def schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, lecture_times, tutorial_times):
    """Schedule core courses with proper LTPSC structure"""
    if core_courses.empty:
        return used_slots
    
    course_day_usage = {}
    
    print(f"   📚 Scheduling {len(core_courses)} core courses...")
    
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

    return used_slots

def generate_section_schedule_with_elective_baskets(dfs, semester_id, section, elective_allocations):
    """Generate schedule with basket-based elective allocation"""
    print(f"   🎯 Generating BASKET-BASED schedule for Semester {semester_id}, Section {section}")
    
    if 'course' not in dfs:
        print("❌ Course data not available")
        return None
    
    try:
        # Get courses for this semester
        course_baskets = separate_courses_by_type(dfs, semester_id)
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

        # Schedule elective courses FIRST using basket allocation
        if elective_allocations:
            print(f"   📅 Applying basket elective slots for Section {section}:")
            # Show what we're scheduling
            unique_baskets = set()
            for allocation in elective_allocations.values():
                if allocation:
                    unique_baskets.add(allocation['basket_name'])
            
            print(f"   🗂️ Baskets to schedule: {list(unique_baskets)}")
            
            used_slots = schedule_electives_by_baskets(elective_allocations, schedule, used_slots, section)
        
        # Schedule core courses AFTER electives
        if not core_courses.empty:
            print(f"   📚 Scheduling {len(core_courses)} core courses for Section {section}...")
            used_slots = schedule_core_courses_with_tutorials(core_courses, schedule, used_slots, days, 
                                                            lecture_times, tutorial_times)
        
        return schedule
        
    except Exception as e:
        print(f"❌ Error generating basket-based schedule: {e}")
        return None

def create_basket_summary(basket_allocations, semester):
    """Create a summary of basket allocations"""
    summary_data = []
    
    for basket_name, allocation in basket_allocations.items():
        # Add lecture allocations
        for i, (day, time_slot) in enumerate(allocation['lectures'], 1):
            summary_data.append({
                'Basket Name': basket_name,
                'Session Type': f'Lecture {i}',
                'Day': day,
                'Time Slot': time_slot,
                'Courses in Basket': ', '.join(allocation['courses']),
                'Number of Courses': len(allocation['courses']),
                'Sections': 'A & B (Common)',
                'Semester': f'Semester {semester}'
            })
        
        # Add tutorial allocation
        if allocation['tutorial']:
            day, time_slot = allocation['tutorial']
            summary_data.append({
                'Basket Name': basket_name,
                'Session Type': 'Tutorial',
                'Day': day,
                'Time Slot': time_slot,
                'Courses in Basket': ', '.join(allocation['courses']),
                'Number of Courses': len(allocation['courses']),
                'Sections': 'A & B (Common)',
                'Semester': f'Semester {semester}'
            })
    
    return pd.DataFrame(summary_data)

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
                'Common for Both Sections': 'Yes'
            })
    
    return pd.DataFrame(summary_data)

# --- Export ---
def export_semester_timetable(dfs, semester):
    print(f"\n📊 Generating BASKET-BASED timetable for Semester {semester}...")
    print(f"🎯 Using elective basket slots")
    
    try:
        # First, identify all elective courses for this semester
        course_baskets = separate_courses_by_type(dfs, semester)
        elective_courses = course_baskets['elective_courses']
        
        print(f"🎯 Elective courses found for semester {semester}: {len(elective_courses)}")
        if not elective_courses.empty:
            print("   Elective courses:", elective_courses['Course Code'].tolist())
            # Show basket distribution
            basket_counts = elective_courses['Basket'].value_counts()
            print("   Basket distribution:")
            for basket, count in basket_counts.items():
                courses = elective_courses[elective_courses['Basket'] == basket]['Course Code'].tolist()
                print(f"      {basket}: {count} courses - {courses}")
        
        # Allocate elective courses by baskets
        elective_allocations, basket_allocations = allocate_electives_by_baskets(elective_courses, semester)
        
        print(f"   📅 BASKET ALLOCATIONS for Semester {semester}:")
        for basket_name, allocation in basket_allocations.items():
            print(f"      {basket_name}:")
            print(f"         Lectures: {allocation['lectures']}")
            print(f"         Tutorial: {allocation['tutorial']}")
            print(f"         Courses: {allocation['courses']}")
        
        # Generate schedules for both sections with basket slots
        section_a = generate_section_schedule_with_elective_baskets(dfs, semester, 'A', elective_allocations)
        section_b = generate_section_schedule_with_elective_baskets(dfs, semester, 'B', elective_allocations)
        
        if section_a is None or section_b is None:
            print(f"❌ Failed to generate timetable for semester {semester}")
            return

        filename = f"sem{semester}_timetable_baskets.xlsx"
        filepath = os.path.join(OUTPUT_DIR, filename)
        
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            section_a.to_excel(writer, sheet_name='Section_A')
            section_b.to_excel(writer, sheet_name='Section_B')
            
            # Add basket allocation summary
            basket_summary = create_basket_summary(basket_allocations, semester)
            basket_summary.to_excel(writer, sheet_name='Basket_Allocation', index=False)
            
            # Add basket course details
            basket_courses_sheet = create_basket_courses_sheet(basket_allocations)
            basket_courses_sheet.to_excel(writer, sheet_name='Basket_Courses', index=False)
            
        print(f"✅ Basket-based timetable saved: {filename}")
        
    except Exception as e:
        print(f"❌ Error generating basket-based timetable for semester {semester}: {e}")

# --- Main execution ---
if __name__ == '__main__':
    print("🚀 Starting Basket-Based Timetable Generation...")
    
    # Load data
    data_frames = load_all_data()
    if data_frames is None:
        print("❌ Failed to load data. Exiting.")
        exit(1)
    
    # Generate timetables for target semesters
    target_semesters = [1, 3, 5, 7]
    
    for sem in target_semesters:
        export_semester_timetable(data_frames, sem)
    
    print("🎉 Basket-based timetable generation completed!")