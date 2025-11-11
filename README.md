# 🎓 Automated Timetable Scheduler – IIIT Dharwad

A **smart web-based application** that automates the creation of academic and examination timetables for **IIIT Dharwad**.  
It intelligently considers faculty availability, classroom capacity, course constraints, and student group schedules to generate **conflict-free and optimized** timetables in seconds.

---

## 🏷️ Badges

![Version](https://img.shields.io/badge/Version-1.0.0-blue.svg)  
![Python](https://img.shields.io/badge/Python-3.8%2B-green.svg)  
![Flask](https://img.shields.io/badge/Flask-2.3.3-lightgrey.svg)  
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

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

# If you're using a local folder path instead:
# cd C:\timetable-generator-Logic-Beyond\timetable-generator-Logic-Beyond\backend

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
# If requirements.txt exists:
pip install -r requirements.txt
# If not, install main packages manually:
pip install flask pandas openpyxl werkzeug

# Step 6: Run the Application
python app.py
# or (Windows)
py app.py
