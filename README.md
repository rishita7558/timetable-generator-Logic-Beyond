# Automated Timetable Generator - IIIT Dharwad

Automates **class timetables**, **basket (elective) schedules**for IIIT Dharwad. The system uses faculty availability, classroom capacity, and course constraints to produce conflict-free timetables and downloadable Excel outputs through a web dashboard.

---

## ✨ Key Features

- Branch-specific timetables for CSE, DSAI, and ECE across semesters 1, 3, 5, 7
- Pre-mid, post-mid, and basket timetable generation with automatic course splits
- Multi-section output (Section A, Section B, and common courses where applicable)
- Drag-and-drop CSV upload with required-file validation
- On-the-fly classroom allocation with room badges in the UI
- Interactive dashboard filters (branch, semester, section, type, grid/list view)
- Download-all, print, and per-file downloads directly from the UI
- Validation sheets and utilization summaries included in Excel outputs

---

## 🧰 Tech Stack

- **Backend:** Python, Flask, Pandas, OpenPyXL
- **Frontend:** HTML, CSS, JavaScript
- **Icons:** Font Awesome

---

## 📋 Prerequisites

- Python 3.8+ (3.12 works fine)
- `pip` (Python package manager)

---

## 📂 Input CSV Files (required)

Place the following files in `backend/temp_inputs/` (or upload from the UI):

- course_data.csv
- faculty_availability.csv
- classroom_data.csv
- student_data.csv

Optional:
- exams_data.csv (exam functionality is currently disabled)
- minor_data.csv
- room_availability.csv

---

## 📦 Output Files

Generated Excel files are saved in `backend/output_timetables/`.

When you click **Generate All Timetables**:

- 24 mid-semester timetables (4 semesters × 3 branches × 2 types)
- 12 basket timetables (Excel files with separate sheets for each section)
- Classroom allocation and utilization sheets

*Note: Exam schedule generation is currently disabled. See [EXAM_TIMETABLE_DISABLED.md](EXAM_TIMETABLE_DISABLED.md) for re-enabling instructions.*

---

## 🧭 Using the Web Interface

1. Start the server.
2. Open http://localhost:5000.
3. Use **Generate All Timetables** for full generation.
4. Use filters (branch/semester/section/type) to explore results.
5. Upload new CSVs from **Upload Files** and click **Process Files**.
   - Note: `exams_data.csv` is no longer required (exam functionality is disabled)
   - All other required files must be present for generation to proceed

---

## 🗂️ Project Structure (key paths)

```
timetable-generator-Logic-Beyond/
├── START_SERVER.bat
├── START_SERVER.ps1
├── START_WITH_VENV.bat
├── START_WITH_VENV.ps1
└── timetable-generator-Logic-Beyond/
    └── backend/
        ├── app.py
        ├── main.py
        ├── requirements.txt
        ├── temp_inputs/
        ├── output_timetables/
        ├── templates/
        └── static/
```

---

## ⚙️ Configuration

- **Default port:** 5000
- **Input folder:** backend/temp_inputs/
- **Output folder:** backend/output_timetables/

To change the port, edit `app.py` and update the `app.run(...)` call.# 🎓 Automated Timetable Scheduler – IIIT Dharwad

A **smart web-based application** that automates the creation of academic and examination timetables for **IIIT Dharwad**.  
It intelligently considers faculty availability, classroom capacity, course constraints, and student group schedules to generate **conflict-free and optimized** timetables in seconds.

---

## ✨ Key Features

- ✅ Automatic generation of **class and exam timetables**  
- ⚡ Built-in **conflict detection** for overlapping classes or unavailable faculty  
- 🏫 Dynamic **classroom allocation** based on student strength *(partially implemented)*  
- 📤 Export timetables in **PDF / Excel / CSV** formats  
- 🧭 **User-friendly interface** for administrators and academic coordinators  
- 🎨 Customizable appearance, layout, and accessibility options  
- 📅 Supports **multiple semesters (1, 3, 5, 7)** and sections (A & B) for all departments  
- 🔎 Smart filtering, color coding, and real-time validation

---

## 🛠️ Technology Stack

**Backend:** Python | Flask | Pandas | OpenPyXL  
**Frontend:** HTML5 | CSS3 | JavaScript  
**Styling:** Custom CSS (Glass Morphism effects)  
**Icons:** Font Awesome

---

## 📋 Prerequisites

- Python 3.8 or higher  
- `pip` (Python package manager)

---

## 📦 Installation & Setup

Follow these steps carefully to set up and run the application locally.

```bash
# Step 1: Clone the Repository
git clone https://github.com/your-username/Automated-Time-Table-IIIT-DHARWAD.git

# Step 2: Navigate to the Project Directory (backend)
cd Automated-Time-Table-IIIT-DHARWAD/timetable_generator/backend

# Step 3: Create a Virtual Environment
# Windows (recommended)
py -m venv venv
# or (cross-platform)
python -m venv venv

# Step 4: Activate the Virtual Environment
# On Windows (PowerShell)
venv\Scripts\Activate.ps1
# On Windows (cmd)
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate

# Step 5: Install Dependencies
pip install flask pandas openpyxl werkzeug

# Step 6: Run the Application
python app.py
# or (Windows)
py app.py
