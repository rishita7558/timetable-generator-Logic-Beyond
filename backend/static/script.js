// Global variables
let currentTimetables = [];
let currentView = 'grid';

// DOM Elements
const generateBtn = document.getElementById('generate-btn');
const refreshBtn = document.getElementById('refresh-btn');
const downloadAllBtn = document.getElementById('download-all-btn');
const emptyGenerateBtn = document.getElementById('empty-generate-btn');
const timetablesContainer = document.getElementById('timetables-container');
const emptyState = document.getElementById('empty-state');
const loadingOverlay = document.getElementById('loading-overlay');
const notificationContainer = document.getElementById('notification-container');

// Stats elements
const totalTimetablesEl = document.getElementById('total-timetables');
const totalCoursesEl = document.getElementById('total-courses');
const totalFacultyEl = document.getElementById('total-faculty');
const totalClassroomsEl = document.getElementById('total-classrooms');

// Filter elements
const semesterFilter = document.getElementById('semester-filter');
const sectionFilter = document.getElementById('section-filter');
const viewMode = document.getElementById('view-mode');

// Initialize application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

function initializeApp() {
    // Event listeners
    generateBtn.addEventListener('click', generateTimetables);
    refreshBtn.addEventListener('click', refreshAll);
    downloadAllBtn.addEventListener('click', downloadAllTimetables);
    emptyGenerateBtn.addEventListener('click', generateTimetables);
    
    // Filter listeners
    semesterFilter.addEventListener('change', filterTimetables);
    sectionFilter.addEventListener('change', filterTimetables);
    viewMode.addEventListener('change', changeViewMode);
    
    // View toggle buttons
    document.querySelectorAll('.view-btn').forEach(btn => {
        btn.addEventListener('click', function() {
            document.querySelectorAll('.view-btn').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentView = this.dataset.view;
            renderTimetables();
        });
    });
    
    // Quick actions
    document.querySelectorAll('.action-card').forEach(card => {
        card.addEventListener('click', function() {
            const action = this.dataset.action;
            handleQuickAction(action);
        });
    });
    
    // Load initial data
    loadStats();
    loadTimetables();
}

// API Functions
async function generateTimetables() {
    showLoading(true);
    
    try {
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const result = await response.json();
        
        if (result.success) {
            showNotification('Timetables generated successfully!', 'success');
            await loadTimetables();
            await loadStats();
        } else {
            showNotification(result.message || 'Error generating timetables', 'error');
        }
    } catch (error) {
        showNotification('Error generating timetables: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function loadTimetables() {
    try {
        const response = await fetch('/timetables');
        const timetables = await response.json();
        
        currentTimetables = timetables;
        renderTimetables();
    } catch (error) {
        console.error('Error loading timetables:', error);
        showNotification('Error loading timetables', 'error');
    }
}

async function loadStats() {
    try {
        const response = await fetch('/stats');
        const stats = await response.json();
        
        totalTimetablesEl.textContent = stats.total_timetables;
        totalCoursesEl.textContent = stats.total_courses;
        totalFacultyEl.textContent = stats.total_faculty;
        totalClassroomsEl.textContent = stats.total_classrooms;
    } catch (error) {
        console.error('Error loading stats:', error);
    }
}

// Rendering Functions
function renderTimetables() {
    if (currentTimetables.length === 0) {
        emptyState.style.display = 'block';
        timetablesContainer.innerHTML = '';
        return;
    }
    
    emptyState.style.display = 'none';
    
    const filteredTimetables = filterTimetablesData();
    
    if (currentView === 'grid') {
        renderGridView(filteredTimetables);
    } else if (currentView === 'list') {
        renderListView(filteredTimetables);
    }
}

function renderGridView(timetables) {
    let html = '<div class="timetables-grid">';
    
    timetables.forEach(timetable => {
        html += `
            <div class="timetable-card">
                <div class="timetable-header">
                    <h3>Semester ${timetable.semester} - Section ${timetable.section}</h3>
                    <div class="timetable-actions">
                        <button class="action-btn" onclick="downloadTimetable('${timetable.filename}')" title="Download">
                            <i class="fas fa-download"></i>
                        </button>
                        <button class="action-btn" onclick="printTimetable(${timetable.semester}, '${timetable.section}')" title="Print">
                            <i class="fas fa-print"></i>
                        </button>
                    </div>
                </div>
                <div class="timetable-content">
                    ${timetable.html}
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    timetablesContainer.innerHTML = html;
}

function renderListView(timetables) {
    let html = '<div class="timetables-list">';
    
    timetables.forEach(timetable => {
        html += `
            <div class="timetable-item">
                <div class="timetable-header">
                    <h3>Semester ${timetable.semester} - Section ${timetable.section}</h3>
                    <div class="timetable-actions">
                        <button class="btn btn-outline" onclick="downloadTimetable('${timetable.filename}')">
                            <i class="fas fa-download"></i> Download
                        </button>
                    </div>
                </div>
                <div class="timetable-content">
                    ${timetable.html}
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    timetablesContainer.innerHTML = html;
}

// Filtering Functions
function filterTimetablesData() {
    const semester = semesterFilter.value;
    const section = sectionFilter.value;
    
    return currentTimetables.filter(timetable => {
        const semesterMatch = semester === 'all' || timetable.semester === parseInt(semester);
        const sectionMatch = section === 'all' || timetable.section === section;
        return semesterMatch && sectionMatch;
    });
}

function filterTimetables() {
    renderTimetables();
}

function changeViewMode() {
    currentView = viewMode.value;
    renderTimetables();
}

// Action Functions
async function downloadAllTimetables() {
    try {
        const response = await fetch('/download-all');
        const blob = await response.blob();
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'all_timetables.zip';
        
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showNotification('All timetables downloaded successfully!', 'success');
    } catch (error) {
        showNotification('Error downloading timetables: ' + error.message, 'error');
    }
}

function downloadTimetable(filename) {
    window.open(`/download/${filename}`, '_blank');
}

function printTimetable(semester, section) {
    showNotification(`Printing Semester ${semester} - Section ${section}`, 'info');
    // In a real implementation, this would open a print dialog for the specific timetable
    window.print();
}

function handleQuickAction(action) {
    switch(action) {
        case 'export':
            downloadAllTimetables();
            break;
        case 'settings':
            showNotification('Settings panel would open here', 'info');
            break;
        case 'help':
            showNotification('Help and support information', 'info');
            break;
        case 'feedback':
            showNotification('Feedback form would open here', 'info');
            break;
    }
}

function refreshAll() {
    loadTimetables();
    loadStats();
    showNotification('Data refreshed successfully!', 'success');
}

// UI Helper Functions
function showLoading(show) {
    if (show) {
        loadingOverlay.classList.add('active');
        // Simulate progress
        simulateProgress();
    } else {
        loadingOverlay.classList.remove('active');
    }
}

function simulateProgress() {
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    let progress = 0;
    
    const interval = setInterval(() => {
        if (progress < 90) {
            progress += Math.random() * 10;
            progressFill.style.width = progress + '%';
            progressText.textContent = Math.round(progress) + '%';
        }
    }, 200);
    
    // Clear interval when loading is done
    setTimeout(() => {
        clearInterval(interval);
        progressFill.style.width = '100%';
        progressText.textContent = '100%';
    }, 3000);
}

function showNotification(message, type = 'info') {
    const notification = document.createElement('div');
    notification.className = `notification ${type}`;
    notification.innerHTML = `
        <i class="fas fa-${getNotificationIcon(type)}"></i>
        <span>${message}</span>
    `;
    
    notificationContainer.appendChild(notification);
    
    setTimeout(() => {
        notification.remove();
    }, 5000);
}

function getNotificationIcon(type) {
    switch(type) {
        case 'success': return 'check-circle';
        case 'error': return 'exclamation-circle';
        case 'warning': return 'exclamation-triangle';
        default: return 'info-circle';
    }
}