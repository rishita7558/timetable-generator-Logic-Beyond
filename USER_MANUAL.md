# User Manual - IIIT Dharwad Timetable Generator

**Version:** 1.0  
**Last Updated:** February 2026

---

## Table of Contents

1. [Introduction](#1-introduction)
2. [System Requirements](#2-system-requirements)
3. [Installation & Setup](#3-installation--setup)
4. [Starting the Server](#4-starting-the-server)
5. [Dashboard Overview](#5-dashboard-overview)
6. [Uploading CSV Files](#6-uploading-csv-files)
7. [CSV File Format Specifications](#7-csv-file-format-specifications)
8. [Generating Timetables](#8-generating-timetables)
9. [Using Filters](#9-using-filters)
10. [Viewing Timetables](#10-viewing-timetables)
11. [Downloading & Exporting](#11-downloading--exporting)
12. [Settings & Customization](#12-settings--customization)
13. [Understanding the Output](#13-understanding-the-output)
14. [Troubleshooting](#14-troubleshooting)
15. [Frequently Asked Questions (FAQ)](#15-frequently-asked-questions-faq)

---

## 1. Introduction

The **IIIT Dharwad Automated Timetable Generator** is a web-based application that automates the creation of academic timetables. The system intelligently handles:

- Faculty availability constraints
- Classroom capacity and allocation
- Course scheduling (lectures, labs, tutorials)
- Multi-section support (Section A, Section B)
- Elective/Basket course scheduling
- Pre-mid and Post-mid semester course splits

### Key Features

| Feature | Description |
|---------|-------------|
| **Branch-specific Timetables** | Generates timetables for CSE, DSAI, and ECE branches |
| **Semester Support** | Supports semesters 1, 3, 5, and 7 |
| **Section Support** | Handles Section A, Section B, and whole-branch courses |
| **Pre-Mid/Post-Mid Split** | Automatically splits courses by semester period |
| **Basket/Elective Scheduling** | Groups elective courses into common time slots |
| **Classroom Allocation** | Automatically assigns rooms based on capacity |
| **Conflict Detection** | Prevents faculty and room double-booking |
| **Excel Export** | Downloads timetables in formatted Excel files |

---

## 2. System Requirements

### Minimum Requirements

- **Operating System:** Windows 10/11, macOS 10.14+, or Linux
- **Python:** Version 3.8 or higher (3.12 recommended)
- **RAM:** 4 GB minimum
- **Disk Space:** 500 MB free space
- **Web Browser:** Chrome, Firefox, Edge, or Safari (latest versions)

### Required Python Packages

```
flask
pandas
openpyxl
werkzeug
```

---

## 3. Installation & Setup

### Step 1: Clone or Download the Project

```bash
# If using Git
git clone <repository-url>

# Or extract the downloaded ZIP file to your preferred location
```

### Step 2: Navigate to the Project Directory

```bash
cd timetable-generator-Logic-Beyond
```

### Step 3: Create a Virtual Environment (Recommended)

**Windows (PowerShell):**
```powershell
py -m venv venv
.\venv\Scripts\Activate.ps1
```

**Windows (Command Prompt):**
```cmd
py -m venv venv
venv\Scripts\activate
```

**macOS/Linux:**
```bash
python3 -m venv venv
source venv/bin/activate
```

### Step 4: Install Dependencies

```bash
pip install flask pandas openpyxl werkzeug
```

Or use the requirements file:

```bash
pip install -r timetable-generator-Logic-Beyond/backend/requirements.txt
```

---

## 4. Starting the Server

### Option 1: Using Batch/PowerShell Scripts (Recommended for Windows)

From the root `timetable-generator-Logic-Beyond` folder:

| Script | Purpose |
|--------|---------|
| `START_SERVER.bat` | Starts server (Command Prompt) |
| `START_SERVER.ps1` | Starts server (PowerShell) |
| `START_WITH_VENV.bat` | Starts with virtual environment activation |
| `START_WITH_VENV.ps1` | Starts with virtual environment (PowerShell) |

Simply double-click the appropriate script.

### Option 2: Manual Start

```bash
cd timetable-generator-Logic-Beyond/backend
python app.py
```

### Accessing the Application

Once the server is running, open your web browser and navigate to:

```
http://localhost:5000
```

You should see the Timetable Dashboard homepage.

---

## 5. Dashboard Overview

The dashboard is divided into several key areas:

### 5.1 Navigation Sidebar (Left)

The sidebar provides quick navigation to:

| Section | Description |
|---------|-------------|
| **Dashboard** | Main overview showing all timetables |
| **Computer Science (CSE)** | CSE branch timetables for semesters 1, 3, 5, 7 |
| **Data Science & AI (DSAI)** | DSAI branch timetables for semesters 1, 3, 5, 7 |
| **Electronics & Communication (ECE)** | ECE branch timetables for semesters 1, 3, 5, 7 |
| **Upload Files** | CSV file upload section |
| **Settings** | UI and theme customization |
| **Help & Support** | Quick start guide and troubleshooting |

### 5.2 Header Section (Top)

- **Generate All Timetables** - Main button to generate all timetables
- **Refresh** - Reload the current view

### 5.3 Statistics Overview

Four stat cards showing:
- 📅 **Total Timetables** - Number of generated timetables
- 📚 **Total Courses** - Number of courses in the system
- 👨‍🏫 **Faculty Members** - Number of faculty loaded
- 🏛️ **Total Classrooms** - Available classroom count

### 5.4 Control Section

Filters and action buttons:
- Branch filter
- Semester filter
- Section filter
- Timetable type filter (Pre-Mid/Post-Mid/Regular)
- View mode toggle (Grid/List)
- Upload, Print, and Download buttons

### 5.5 Timetables Display Area

Shows generated timetables in either grid or list view.

### 5.6 Quick Actions Panel

Direct access to:
- 📤 Upload Files
- 📊 Export Data
- ⚙️ Settings
- ❓ Help & Support

---

## 6. Uploading CSV Files

### Required Files

The following CSV files are **required** for timetable generation:

| File | Purpose |
|------|---------|
| `course_data.csv` | Course information, faculty assignments, credits |
| `faculty_availability.csv` | Faculty names (availability constraints) |
| `classroom_data.csv` | Room numbers, types, and capacities |
| `student_data.csv` | Student enrollment data |

### Optional Files

| File | Purpose |
|------|---------|
| `minor_data.csv` | Minor course registrations |
| `exams_data.csv` | Exam scheduling (currently disabled) |
| `room_availability.csv` | Additional room constraints |

### Upload Process

1. Click **"Upload Files"** in the sidebar or Quick Actions panel
2. The upload section will expand showing required files
3. **Drag and drop** your CSV files onto the upload area, OR
4. Click **"Browse Files"** to select files manually
5. The system will validate each file as it's uploaded
6. File status indicators show:
   - ✅ **Green checkmark** - File uploaded successfully
   - ⚠️ **Yellow warning** - File may have issues
   - ❌ **Red X** - File missing or invalid
7. Once all required files show green checkmarks, click **"Process Files"**

### File Upload Tips

- Ensure CSV files are properly formatted with headers
- File names must match exactly (case-insensitive)
- Files are automatically validated upon upload
- Previously uploaded files persist until replaced

---

## 7. CSV File Format Specifications

### 7.1 course_data.csv

**Purpose:** Contains all course information for scheduling.

| Column | Required | Description | Example Values |
|--------|----------|-------------|----------------|
| `Course Code` | Yes | Unique course identifier | CS161, MA262 |
| `Course Name` | Yes | Full course name | Problem Solving |
| `Semester` | Yes | Semester number | 1, 3, 5, 7 |
| `Department` | Yes | Branch/department | CSE, DSAI, ECE |
| `LTPSC` | Yes | Lecture-Tutorial-Practical-Self-Credits | 3-0-2-0-4 |
| `Credits` | Yes | Credit hours | 2, 3, 4 |
| `Faculty` | Yes | Instructor name(s) | "Sunil P V, Vivekraj" |
| `Registered Students` | No | Student count | 160 |
| `Elective (Yes/No)` | Yes | Is this an elective? | Yes, No |
| `Half Semester (Yes/No)` | Yes | Half-semester course? | Yes, No |
| `Basket` | No | Elective basket group | ELECTIVE_B1, HSS_B3, None |
| `Post mid-sem` | No | Taught in post-mid? | Yes, No |
| `Common` | No | Common to all sections? | Yes, No |

**Example:**
```csv
Course Code,Course Name,Semester,Department,LTPSC,Credits,Faculty,Registered Students,Elective (Yes/No),Half Semester (Yes/No),Basket,Post mid-sem,Common
MA161,Statistics,1,CSE,2-0-0-0-2,2,Ramesh Athe,160,No,Yes,None,No,Yes
CS161,Problem Solving,1,CSE,3-0-2-0-4,4,"Sunil P V, Sunil C K",160,No,No,None,No,No
```

### 7.2 faculty_availability.csv

**Purpose:** Lists faculty members for course assignment validation.

| Column | Required | Description |
|--------|----------|-------------|
| `FACULTY NAME` | Yes | Full name of faculty member |

**Example:**
```csv
FACULTY NAME
Anand Barangi
Animesh Roy
Sunil P V
Abdul Wahid
```

### 7.3 classroom_data.csv

**Purpose:** Defines available rooms and their capacities.

| Column | Required | Description | Example Values |
|--------|----------|-------------|----------------|
| `Room Number` | Yes | Room identifier | C101, L105 |
| `Type` | Yes | Room type | classroom, Software Lab, Hardware Lab |
| `Capacity` | Yes | Seating capacity | 40, 96, 120 |
| `Facilities` | No | Available equipment | Projector, Computers |
| `exam capacity` | No | Capacity for exams | 48 |

**Example:**
```csv
Room Number,Type,Capacity,Facilities,exam capacity
C101,classroom,96,Projector,48
L105,Hardware Lab,40,Equipment,0
L106,Software Lab,40,Computers,0
```

### 7.4 student_data.csv

**Purpose:** Contains student enrollment information.

| Column | Required | Description |
|--------|----------|-------------|
| `Roll No` | Yes | Student roll number |
| `Name` | Yes | Student name |
| `Semester` | Yes | Current semester |
| `Department` | Yes | Branch/department |

**Example:**
```csv
Roll No,Name,Semester,Department
25BCS001,Aarav Sharma,1,CSE
25BCS002,Aditya Patel,1,CSE
25BDS001,Anjali Singh,1,DSAI
```

### 7.5 minor_data.csv (Optional)

**Purpose:** Defines minor course registrations.

| Column | Required | Description |
|--------|----------|-------------|
| `MINOR COURSE` | Yes | Minor course name |
| `SEMESTER` | Yes | Semester offered |
| `REGISTERED STUDENTS` | Yes | Number enrolled |

**Example:**
```csv
MINOR COURSE ,SEMESTER,REGISTERED STUDENTS
Generative Ai,3,140
Cybersecurity,3,36
Design,3,15
```

---

## 8. Generating Timetables

### Full Generation (Recommended)

1. Ensure all required CSV files are uploaded
2. Click the **"Generate All Timetables"** button in the header
3. A loading overlay will appear showing progress
4. Wait for generation to complete (typically 10-30 seconds)
5. Success notification will appear when complete
6. Timetables will automatically display in the main area

### What Gets Generated

When you click "Generate All Timetables":

| Output | Count | Description |
|--------|-------|-------------|
| **Pre-Mid Timetables** | 12 | 4 semesters × 3 branches (Section A & B combined) |
| **Post-Mid Timetables** | 12 | 4 semesters × 3 branches |
| **Basket Timetables** | 12 | Elective schedules per semester/branch |
| **Allocation Sheets** | Multiple | Classroom and faculty allocation reports |

### Generation Process Details

The system performs these steps:

1. **Data Loading** - Reads all CSV files
2. **Validation** - Checks for conflicts and missing data
3. **Course Assignment** - Places courses based on LTPSC values
4. **Lab Scheduling** - Assigns consecutive slots for lab sessions
5. **Elective Grouping** - Groups basket courses in common slots
6. **Room Allocation** - Assigns classrooms based on capacity
7. **Conflict Resolution** - Ensures no faculty/room double-booking
8. **Excel Generation** - Creates formatted Excel output files

---

## 9. Using Filters

### Branch Filter

Filter timetables by department:
- **All Branches** - Show all departments
- **Computer Science** - CSE only
- **Data Science & AI** - DSAI only
- **Electronics & Communication** - ECE only

### Semester Filter

Filter by academic semester:
- **All Semesters** - Show all
- **Semester 1, 3, 5, 7** - Specific semester

### Section Filter

Filter by student section:
- **All Sections** - Show all
- **Section A** - First section
- **Section B** - Second section
- **Whole Branch** - Courses common to all students

### Timetable Type Filter

Filter by schedule period:
- **All Types** - Show everything
- **Regular Timetables** - Full semester schedules
- **Pre-Mid Timetables** - First half of semester
- **Post-Mid Timetables** - Second half of semester

### View Mode

Toggle between:
- **Grid View** - Card-based visual layout
- **List View** - Compact table format

---

## 10. Viewing Timetables

### Grid View

In grid view, each timetable appears as a card showing:
- **Header** - Branch, semester, and section info
- **Preview Grid** - Mini timetable preview
- **Actions** - Download and view buttons
- **Room Badge** - Assigned classroom indicator

### List View

List view shows a compact table with:
- Timetable name/description
- Branch and semester
- Section
- Download link

### Timetable Card Details

Click on any timetable card to see:
- Full schedule grid (Monday - Friday)
- All time slots (07:30 - 20:00)
- Course codes with room assignments
- Color-coded course types
- Faculty assignments

### Time Slot Schedule

The standard day schedule:

| Slot | Time |
|------|------|
| 1 | 07:30 - 09:00 |
| 2 | 09:00 - 10:30 |
| 3 | 10:30 - 12:00 |
| 4 | 12:00 - 13:00 (SHORT) |
| LUNCH | 13:00 - 14:30 |
| 5 | 14:30 - 15:30 (SHORT) |
| 6 | 15:30 - 17:00 |
| 7 | 17:00 - 18:00 (SHORT) |
| 8 | 18:30 - 20:00 |

---

## 11. Downloading & Exporting

### Download Individual Timetable

1. Locate the timetable card
2. Click the **Download** icon (⬇️) on the card
3. Excel file will download to your browser's download folder

### Download All Timetables

1. Click **"Download All"** button in the control section
2. A ZIP file containing all generated Excel files will download
3. Extract the ZIP to access individual files

### Print Timetables

1. Click **"Print All"** button
2. Browser print dialog will open
3. Select printer and settings
4. Click Print

### Output Files Location

Generated files are saved to:
```
timetable-generator-Logic-Beyond/backend/output_timetables/
```

### Excel File Contents

Each Excel file includes:
- **Main Timetable Sheet** - Schedule grid with course codes
- **Section A Sheet** - Section A specific schedule
- **Section B Sheet** - Section B specific schedule
- **Room Allocation Sheet** - Classroom assignments
- **Validation Sheet** - Data validation summary

---

## 12. Settings & Customization

### Accessing Settings

Click **Settings** in the sidebar or Quick Actions panel.

### Theme Selection

Choose from multiple color themes:
- Light themes for daytime use
- Dark themes for reduced eye strain
- High contrast options for accessibility

### Appearance Options

| Setting | Options | Description |
|---------|---------|-------------|
| **Font Size** | Small, Medium, Large, Extra Large | Text size throughout the app |
| **Density** | Compact, Comfortable, Spacious | UI element spacing |

### Accessibility Options

| Setting | Description |
|---------|-------------|
| **High Contrast Mode** | Enhanced visibility for better readability |
| **Reduce Animations** | Disables transitions for performance |

### Layout Options

| Setting | Description |
|---------|-------------|
| **Compact Mode** | Reduces padding and margins |
| **Collapse Sidebar** | Minimizes sidebar for more workspace |

### Saving Settings

- Settings are saved automatically in your browser
- Click **"Reset to Defaults"** to restore original settings
- Click **"Apply Changes"** to close the settings modal

---

## 13. Understanding the Output

### Timetable Naming Convention

Files follow this naming pattern:
```
Sem{X}_{Branch}_{Section}_{Type}_timetable.xlsx
```

Examples:
- `Sem1_CSE_A_pre_mid_timetable.xlsx`
- `Sem3_DSAI_B_post_mid_timetable.xlsx`
- `Sem5_ECE_basket_timetable.xlsx`

### Course Entry Format

Timetable cells show:
```
COURSE_CODE [ROOM]
Faculty Name
```

Example:
```
CS161 [C101]
Sunil P V
```

### Color Coding

Courses are color-coded by type:
- **Blue** - Core courses
- **Green** - Lab sessions
- **Purple** - Electives
- **Orange** - Tutorials

### Room Allocation Badges

Room numbers appear in brackets:
- `[C101]` - Classroom 101
- `[L106]` - Lab 106
- `[C004]` - Auditorium

---

## 14. Troubleshooting

### Common Issues

#### Issue: "No timetables generated"

**Possible Causes:**
- Required CSV files not uploaded
- CSV format errors
- Empty data files

**Solutions:**
1. Check that all 4 required files are uploaded
2. Verify CSV headers match expected format
3. Ensure files have data rows (not just headers)

#### Issue: "Page won't load"

**Possible Causes:**
- Server not running
- Port already in use

**Solutions:**
1. Check terminal for error messages
2. Try restarting the server
3. Use a different port (edit app.py)

#### Issue: "Elective courses not showing"

**Possible Causes:**
- Elective column not set correctly

**Solutions:**
1. Ensure `Elective (Yes/No)` column has "Yes" for electives
2. Check `Basket` column has valid basket names (e.g., ELECTIVE_B1)

#### Issue: "Faculty conflicts detected"

**Possible Causes:**
- Same faculty assigned to multiple concurrent courses

**Solutions:**
1. Review faculty assignments in course_data.csv
2. Check faculty_availability.csv for constraints
3. Regenerate timetables

#### Issue: "Download not working"

**Possible Causes:**
- Browser blocking downloads
- Files not generated

**Solutions:**
1. Check browser download settings
2. Ensure timetables were generated successfully
3. Check output_timetables folder for files

### Debug Tools

Access debug information at:
- `/debug/current-data` - View loaded data
- `/debug/file-matching` - Check file parsing
- `/debug/clear-cache` - Clear cached data

### Clearing Cache

If data seems stale:
1. Click **Help & Support** in sidebar
2. Click **"Clear Cache"** button
3. Refresh the page
4. Re-upload files if needed

---

## 15. Frequently Asked Questions (FAQ)

### Q1: How long does timetable generation take?

**A:** Typical generation takes 10-30 seconds depending on the number of courses and constraints. Complex schedules with many electives may take longer.

### Q2: Can I edit generated timetables?

**A:** The application generates read-only Excel files. To make changes, modify the input CSV files and regenerate.

### Q3: How do I add a new course?

**A:** Add a new row to `course_data.csv` with all required columns, then upload the updated file and regenerate timetables.

### Q4: What happens to old timetables when I regenerate?

**A:** Old files in `output_timetables/` are overwritten. Download any important files before regenerating.

### Q5: Why are some rooms not being used?

**A:** The system prioritizes rooms based on:
- Capacity matching student count
- Room type (classroom vs. lab)
- Previous allocations for consistency

### Q6: Can I schedule weekend classes?

**A:** Currently, the system schedules Monday through Friday only. Weekend scheduling requires code modifications.

### Q7: How do I report a bug?

**A:** Document the issue with:
- Steps to reproduce
- CSV file samples (if relevant)
- Screenshots of the error
- Browser console errors

Submit through your institution's feedback channel.

### Q8: Can multiple users use the system simultaneously?

**A:** The application runs locally. For multi-user access, deploy to a server and ensure proper session handling.

### Q9: How do I backup my data?

**A:** Regularly backup:
- All CSV files in `temp_inputs/`
- Generated files in `output_timetables/`
- Any custom configuration

### Q10: Is exam scheduling available?

**A:** Exam timetable functionality is currently **disabled**. Contact the development team for information about re-enabling this feature.

---

## Quick Reference Card

### Keyboard Shortcuts

| Action | Shortcut |
|--------|----------|
| Refresh Page | `F5` or `Ctrl+R` |
| Print | `Ctrl+P` |

### Important URLs

| URL | Purpose |
|-----|---------|
| `http://localhost:5000` | Main dashboard |
| `http://localhost:5000/debug/current-data` | Debug data view |
| `http://localhost:5000/stats` | Statistics API |

### File Locations

| Item | Path |
|------|------|
| Application | `timetable-generator-Logic-Beyond/backend/` |
| Input Files | `backend/temp_inputs/` |
| Output Files | `backend/output_timetables/` |
| Test Data | `backend/test_data/` |

---

## Support & Contact

For additional assistance:
1. Use the **Help & Support** section in the application
2. Check the **Debug Info** for technical details
3. Contact your institution's IT support team

---

*This manual is for IIIT Dharwad Automated Timetable Generator v1.0*
