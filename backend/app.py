from flask import Flask, render_template, request, jsonify, send_file
import os
import pandas as pd
import random
from main import load_all_data, export_semester_timetable, OUTPUT_DIR
import zipfile
import glob

app = Flask(__name__)

# Configuration
INPUT_DIR = os.path.join(os.getcwd(), "temp_inputs")
OUTPUT_DIR = os.path.join(os.getcwd(), "output_timetables")
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_timetables():
    try:
        # Load data and generate timetables
        data_frames = load_all_data()
        if data_frames is None:
            return jsonify({'success': False, 'message': 'Failed to load CSV data'})

        # Generate timetables for semesters 1,3,5,7
        target_semesters = [1, 3, 5, 7]
        success_count = 0
        
        for sem in target_semesters:
            try:
                export_semester_timetable(data_frames, sem)
                success_count += 1
            except Exception as e:
                print(f"Error generating timetable for semester {sem}: {e}")

        return jsonify({
            'success': True, 
            'message': f'Successfully generated {success_count} timetables!',
            'generated_count': success_count
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
            # Extract semester and section info from filename
            if 'sem' in filename and 'timetable' in filename:
                sem = filename.split('sem')[1].split('_')[0]
                
                # Read both sections from the Excel file
                try:
                    df_a = pd.read_excel(file_path, sheet_name='Section_A')
                    df_b = pd.read_excel(file_path, sheet_name='Section_B')
                    
                    # Convert to HTML tables
                    html_a = df_a.to_html(classes='timetable-table', index=False, escape=False)
                    html_b = df_b.to_html(classes='timetable-table', index=False, escape=False)
                    
                    timetables.append({
                        'semester': int(sem),
                        'section': 'A',
                        'filename': filename,
                        'html': html_a
                    })
                    
                    timetables.append({
                        'semester': int(sem),
                        'section': 'B',
                        'filename': filename,
                        'html': html_b
                    })
                    
                except Exception as e:
                    print(f"Error reading {filename}: {e}")
        
        return jsonify(timetables)
        
    except Exception as e:
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

if __name__ == '__main__':
    print("Starting Timetable Generator Web Application...")
    print(f"Access the application at: http://127.0.0.1:5000")
    print(f"Input directory: {INPUT_DIR}")
    print(f"Output directory: {OUTPUT_DIR}")
    app.run(debug=True, host='127.0.0.1', port=5000)