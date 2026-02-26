# Automated Timetable Generator - IIIT Dharwad

A **smart web-based application** that automates the creation of academic timetables for **IIIT Dharwad**. The system intelligently considers faculty availability, classroom capacity, course constraints, and student group schedules to generate **conflict-free and optimized** timetables in seconds.

> **For Testing Teams:** Please refer to [USER_MANUAL.md](USER_MANUAL.md) for detailed usage instructions covering all features.

---

## ✨ Key Features

| Feature | Description |
|---------|-------------|
| **Branch-Specific Timetables** | Generates timetables for CSE, DSAI, and ECE across semesters 1, 3, 5, 7 |
| **Pre-Mid/Post-Mid Split** | Automatic course division for half-semester courses |
| **Multi-Section Support** | Section A, Section B, and whole-branch courses |
| **Basket/Elective Scheduling** | Groups elective courses into common time slots |
| **Classroom Allocation** | Dynamic room assignment based on capacity |
| **Conflict Detection** | Prevents faculty and room double-booking |
| **Drag-and-Drop Upload** | Easy CSV file upload with validation |
| **Interactive Filters** | Filter by branch, semester, section, or type |
| **Excel Export** | Download formatted Excel files with multiple sheets |
| **Theme Customization** | Multiple themes and accessibility options |

---

## 🛠️ Technology Stack

| Layer | Technologies |
|-------|--------------|
| **Backend** | Python 3.8+, Flask, Pandas, OpenPyXL |
| **Frontend** | HTML5, CSS3, JavaScript |
| **Styling** | Custom CSS with Glass Morphism effects |
| **Icons** | Font Awesome |

---

## 📋 Prerequisites

- Python 3.8 or higher (3.12 recommended)
- `pip` (Python package manager)
- Modern web browser (Chrome, Firefox, Edge, Safari)

---

## 🚀 Quick Start

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
```

---

## 📂 Required CSV Files

Upload these files via the web interface or place them in `backend/temp_inputs/`:

| File | Purpose | Required |
|------|---------|----------|
| `course_data.csv` | Course info, faculty, credits, elective flags | ✅ Yes |
| `faculty_availability.csv` | Faculty names | ✅ Yes |
| `classroom_data.csv` | Room numbers, types, capacities | ✅ Yes |
| `student_data.csv` | Student enrollment data | ✅ Yes |
| `minor_data.csv` | Minor course registrations | Optional |

> See [USER_MANUAL.md](USER_MANUAL.md) for detailed CSV format specifications.

---

## 📦 Output

Generated Excel files are saved in `backend/output_timetables/`:

- **24 Mid-Semester Timetables** (4 semesters × 3 branches × 2 types)
- **12 Basket Timetables** (elective schedules per semester/branch)
- **Allocation & Validation Sheets**

---

## 🧭 Using the Web Interface

1. **Start** the server and open http://localhost:5000
2. **Upload** CSV files via the Upload section
3. **Click** "Generate All Timetables"
4. **Filter** results by branch, semester, section, or type
5. **Download** individual files or all as ZIP

---

## 🗂️ Project Structure

```
timetable-generator-Logic-Beyond/
├── START_SERVER.bat          # Quick start scripts
├── START_SERVER.ps1
├── START_WITH_VENV.bat
├── START_WITH_VENV.ps1
└── timetable-generator-Logic-Beyond/
    ├── USER_MANUAL.md        # Detailed user documentation
    ├── README.md             # This file
    └── backend/
        ├── app.py            # Main Flask application
        ├── requirements.txt  # Python dependencies
        ├── temp_inputs/      # Input CSV files
        ├── output_timetables/# Generated Excel files
        ├── templates/        # HTML templates
        ├── static/           # CSS, JS, assets
        └── test_data/        # Sample test data
```

---

## ⚙️ Configuration

| Setting | Default | Description |
|---------|---------|-------------|
| Port | 5000 | Edit `app.run()` in app.py to change |
| Input folder | `backend/temp_inputs/` | CSV files location |
| Output folder | `backend/output_timetables/` | Generated files |

---

## 📖 Documentation

- **[USER_MANUAL.md](USER_MANUAL.md)** - Complete user documentation with screenshots
- **[SETUP_AND_RUN_GUIDE.md](../SETUP_AND_RUN_GUIDE.md)** - Detailed setup instructions

---

## ⚠️ Known Limitations

- Exam schedule generation is currently **disabled**
- Weekend scheduling not supported
- Single-user local deployment (multi-user requires server deployment)

---

## 🐛 Troubleshooting

| Issue | Solution |
|-------|----------|
| No timetables generated | Ensure all 4 required CSV files are uploaded |
| Page won't load | Check if server is running; try port 5000 |
| Electives not showing | Verify `Elective (Yes/No)` column in course_data.csv |
| Download fails | Check browser download settings; verify files exist |

For more troubleshooting, see the [USER_MANUAL.md](USER_MANUAL.md#14-troubleshooting).

---

## 📞 Support

- Use **Help & Support** in the application sidebar
- Access debug endpoints: `/debug/current-data`, `/debug/file-matching`
- Clear cache if data seems stale

---

*IIIT Dharwad Automated Timetable Generator v1.0*
