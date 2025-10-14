from flask import Flask, render_template, request, jsonify, send_file
import os
import pandas as pd
import random
import zipfile
import glob

app = Flask(__name__)

# Configuration
INPUT_DIR = os.path.join(os.getcwd(), "temp_inputs")
OUTPUT_DIR = os.path.join(os.getcwd(), "output_timetables")
os.makedirs(INPUT_DIR, exist_ok=True)
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

def clean_table_html(html):
    """Clean and format the HTML table for better display"""
    # Remove default pandas styling and add our classes
    html = html.replace('border="1"', '')
    html = html.replace('class="dataframe"', 'class="timetable-table"')
    html = html.replace('<thead>', '<thead class="timetable-head">')
    html = html.replace('<tbody>', '<tbody class="timetable-body">')
    return html

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_timetables():
    try:
        # Clear existing timetables first
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        for file in excel_files:
            try:
                os.remove(file)
                print(f"🗑️ Removed old file: {file}")
            except Exception as e:
                print(f"⚠️ Could not remove {file}: {e}")

        # Load data and generate timetables
        data_frames = load_all_data()
        if data_frames is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'})

        # Generate timetables for semesters 1,3,5,7
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        generated_files = []
        
        for sem in target_semesters:
            try:
                export_semester_timetable(data_frames, sem)
                filename = f"sem{sem}_timetable.xlsx"
                filepath = os.path.join(OUTPUT_DIR, filename)
                
                if os.path.exists(filepath):
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
        return jsonify({'success': False, 'message': f'Error: {str(e)}'})

@app.route('/timetables')
def get_timetables():
    try:
        timetables = []
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        
        for file_path in excel_files:
            filename = os.path.basename(file_path)
            # Extract semester info from filename
            if 'sem' in filename and 'timetable' in filename:
                try:
                    # Extract semester number from filename
                    sem_part = filename.split('sem')[1].split('_')[0]
                    sem = int(sem_part)
                    
                    # Read both sections from the Excel file
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                    
                    # Convert to HTML tables with proper formatting
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
                    
                    timetables.append({
                        'semester': sem,
                        'section': 'A',
                        'filename': filename,
                        'html': html_a
                    })
                    
                    timetables.append({
                        'semester': sem,
                        'section': 'B',
                        'filename': filename,
                        'html': html_b
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
        
        # Count courses from course_data.csv
        course_count = 0
        faculty_count = 0
        classroom_count = 0
        
        course_file = os.path.join(INPUT_DIR, "course_data.csv")
        faculty_file = os.path.join(INPUT_DIR, "faculty_availability.csv")
        classroom_file = os.path.join(INPUT_DIR, "classroom_data.csv")
        
        if os.path.exists(course_file):
            df = pd.read_csv(course_file)
            course_count = len(df)
        
        if os.path.exists(faculty_file):
            df = pd.read_csv(faculty_file)
            faculty_count = len(df)
            
        if os.path.exists(classroom_file):
            df = pd.read_csv(classroom_file)
            classroom_count = len(df)
        
        return jsonify({
            'total_timetables': total_timetables,
            'total_courses': course_count,
            'total_faculty': faculty_count,
            'total_classrooms': classroom_count
        })
        
    except Exception as e:
        return jsonify({
            'total_timetables': 0,
            'total_courses': 0,
            'total_faculty': 0,
            'total_classrooms': 0
        })

@app.route('/debug-timetables')
def debug_timetables():
    """Debug endpoint to check timetable files"""
    try:
        excel_files = glob.glob(os.path.join(OUTPUT_DIR, "*.xlsx"))
        debug_info = {
            'output_dir': OUTPUT_DIR,
            'excel_files': excel_files,
            'files_count': len(excel_files)
        }
        
        for file_path in excel_files:
            filename = os.path.basename(file_path)
            debug_info[filename] = {
                'exists': os.path.exists(file_path),
                'size': os.path.getsize(file_path),
                'sheets': []
            }
            
            try:
                # Check what sheets are available
                xl_file = pd.ExcelFile(file_path)
                debug_info[filename]['sheets'] = xl_file.sheet_names
                
                # Check first few rows of each sheet
                for sheet in xl_file.sheet_names:
                    df = pd.read_excel(file_path, sheet_name=sheet)
                    debug_info[filename][f'{sheet}_sample'] = df.head(3).to_dict('records')
                    
            except Exception as e:
                debug_info[filename]['error'] = str(e)
        
        return jsonify(debug_info)
        
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    print("Starting Timetable Generator Web Application...")
    print(f"Access the application at: http://127.0.0.1:5000")
    print(f"Input directory: {INPUT_DIR}")
    print(f"Output directory: {OUTPUT_DIR}")
    app.run(debug=True, host='127.0.0.1', port=5000)