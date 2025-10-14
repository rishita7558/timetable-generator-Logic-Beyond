// Global variables
let currentTimetables = [];
let currentView = 'grid';
let currentSemesterFilter = 'all';
let currentSectionFilter = 'all';

// Course information database
const courseDatabase = {
    'MA101': { name: 'Mathematics I', credits: 4, type: 'Core' },
    'DS101': { name: 'Data Structures', credits: 4, type: 'Core' },
    'MA102': { name: 'Mathematics II', credits: 4, type: 'Core' },
    'EC101': { name: 'Electronics', credits: 3, type: 'Core' },
    'CS101': { name: 'Computer Programming', credits: 4, type: 'Core' },
    'HS101': { name: 'Communication Skills', credits: 2, type: 'Core' },
    'CS151': { name: 'Programming Lab', credits: 2, type: 'Lab' },
    'MA261': { name: 'Probability & Statistics', credits: 4, type: 'Core' },
    'CS261': { name: 'Algorithms', credits: 4, type: 'Core' },
    'CS263': { name: 'Database Systems', credits: 4, type: 'Core' },
    'CS264': { name: 'Computer Networks', credits: 4, type: 'Core' },
    'CS309': { name: 'Software Engineering', credits: 4, type: 'Core' },
    'CS303': { name: 'Machine Learning', credits: 4, type: 'Core' },
    'CS304': { name: 'Operating Systems', credits: 4, type: 'Core' },
    'CS461': { name: 'Artificial Intelligence', credits: 4, type: 'Elective' },
    'DS456': { name: 'Data Science', credits: 4, type: 'Elective' },
    'EC456': { name: 'Embedded Systems', credits: 4, type: 'Elective' },
    'DS401': { name: 'Big Data Analytics', credits: 4, type: 'Elective' }
};

// DOM Elements
const generateBtn = document.getElementById('generate-btn');
const refreshBtn = document.getElementById('refresh-btn');
const downloadAllBtn = document.getElementById('download-all-btn');
const emptyGenerateBtn = document.getElementById('empty-generate-btn');
const timetablesContainer = document.getElementById('timetables-container');
const emptyState = document.getElementById('empty-state');
const loadingOverlay = document.getElementById('loading-overlay');
const notificationContainer = document.getElementById('notification-container');
const sectionTitle = document.getElementById('section-title');

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
    console.log("🚀 Initializing Timetable Application...");
    
    // Event listeners
    generateBtn.addEventListener('click', generateTimetables);
    refreshBtn.addEventListener('click', refreshAll);
    downloadAllBtn.addEventListener('click', downloadAllTimetables);
    emptyGenerateBtn.addEventListener('click', generateTimetables);
    
    // Filter listeners
    semesterFilter.addEventListener('change', function() {
        currentSemesterFilter = this.value;
        updateSectionTitle();
        renderTimetables();
    });
    
    sectionFilter.addEventListener('change', function() {
        currentSectionFilter = this.value;
        updateSectionTitle();
        renderTimetables();
    });
    
    viewMode.addEventListener('change', changeViewMode);
    
    // Sidebar navigation
    setupSidebarNavigation();
    
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
    
    // Print button
    const printAllBtn = document.getElementById('print-all-btn');
    if (printAllBtn) {
        printAllBtn.addEventListener('click', printAllTimetables);
    }
    
    // Debug button (add this to your HTML or use browser console)
    window.debugApp = debugApp;
    
    // Load initial data
    loadStats();
    loadTimetables();
    
    console.log("✅ Application initialized successfully");
}

// Sidebar Navigation
function setupSidebarNavigation() {
    const navItems = document.querySelectorAll('.sidebar-nav a[data-semester]');
    
    navItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            
            // Remove active class from all items
            document.querySelectorAll('.nav-item').forEach(nav => {
                nav.classList.remove('active');
            });
            
            // Add active class to clicked item
            this.parentElement.classList.add('active');
            
            // Get semester filter
            const semester = this.getAttribute('data-semester');
            currentSemesterFilter = semester;
            
            // Update semester filter dropdown
            semesterFilter.value = semester;
            
            // Update section title and render
            updateSectionTitle();
            renderTimetables();
        });
    });
}

function updateSectionTitle() {
    if (currentSemesterFilter === 'all') {
        sectionTitle.textContent = 'All Timetables';
    } else {
        sectionTitle.textContent = `Semester ${currentSemesterFilter} Timetables`;
    }
    
    // Add section info if filtered
    if (currentSectionFilter !== 'all') {
        sectionTitle.textContent += ` - Section ${currentSectionFilter}`;
    }
}

// API Functions
async function generateTimetables() {
    showLoading(true);
    
    try {
        console.log("🔄 Starting timetable generation...");
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const result = await response.json();
        console.log("📦 Generation result:", result);
        
        if (result.success) {
            showNotification(`✅ ${result.message}`, 'success');
            console.log('📁 Generated files:', result.files);
            
            // Reload data
            await loadTimetables();
            await loadStats();
            
            // If no files were generated, show debug info
            if (result.generated_count === 0) {
                showNotification('⚠️ No timetables were generated. Check the console for details.', 'warning');
            }
        } else {
            showNotification(`❌ ${result.message}`, 'error');
            console.error('Generation failed:', result.message);
        }
    } catch (error) {
        console.error('❌ Error generating timetables:', error);
        showNotification('❌ Error generating timetables: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function loadTimetables() {
    try {
        console.log("🔄 Loading timetables...");
        const response = await fetch('/timetables');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const timetables = await response.json();
        
        console.log(`📊 Received ${timetables.length} timetables`);
        
        currentTimetables = timetables;
        renderTimetables();
        
        // Show notification if no timetables
        if (timetables.length === 0) {
            console.log("ℹ️ No timetables available");
            showNotification('No timetables found. Click "Generate All Timetables" to create them.', 'info');
        } else {
            console.log("✅ Timetables loaded successfully");
        }
        
    } catch (error) {
        console.error('❌ Error loading timetables:', error);
        showNotification('❌ Error loading timetables: ' + error.message, 'error');
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
        
        console.log("📈 Stats loaded:", stats);
    } catch (error) {
        console.error('Error loading stats:', error);
    }
}

// Color Coding and Legend Functions
function applyColorCoding(tableElement) {
    const cells = tableElement.querySelectorAll('td');
    
    cells.forEach(cell => {
        const text = cell.textContent.trim();
        
        // Skip header cells, empty cells, and special slots
        if (!text || text === 'Free' || text === 'LUNCH BREAK' || cell.cellIndex === 0) {
            cell.classList.add('empty-cell');
            return;
        }
        
        // Extract course code (assuming format like "MA101" or similar)
        const courseCode = extractCourseCode(text);
        
        if (courseCode && courseDatabase[courseCode]) {
            // Add color coding class
            cell.classList.add(`course-${courseCode}`);
            
            // Add tooltip with course info
            const courseInfo = courseDatabase[courseCode];
            cell.title = `${courseCode}: ${courseInfo.name} (${courseInfo.credits} credits)`;
            
            // Make cell clickable for more info
            cell.style.cursor = 'help';
        }
    });
}

function extractCourseCode(text) {
    // Match common course code patterns like MA101, CS101, etc.
    const coursePattern = /[A-Z]{2,3}\d{3}/;
    const match = text.match(coursePattern);
    return match ? match[0] : text;
}

function createLegend(semester, section) {
    const timetable = currentTimetables.find(t => 
        t.semester === semester && t.section === section
    );
    
    if (!timetable) return '';
    
    // Extract unique courses from the timetable
    const tempDiv = document.createElement('div');
    tempDiv.innerHTML = timetable.html;
    const cells = tempDiv.querySelectorAll('td');
    const uniqueCourses = new Set();
    
    cells.forEach(cell => {
        const text = cell.textContent.trim();
        if (text && text !== 'Free' && text !== 'LUNCH BREAK') {
            const courseCode = extractCourseCode(text);
            if (courseCode && courseDatabase[courseCode]) {
                uniqueCourses.add(courseCode);
            }
        }
    });
    
    if (uniqueCourses.size === 0) return '';
    
    // Create legend HTML
    let legendHtml = `
        <div class="timetable-legend">
            <div class="legend-title">
                <i class="fas fa-palette"></i>
                Course Legend - Semester ${semester}, Section ${section}
            </div>
            <div class="legend-grid">
    `;
    
    // Sort courses alphabetically
    const sortedCourses = Array.from(uniqueCourses).sort();
    
    sortedCourses.forEach(courseCode => {
        const courseInfo = courseDatabase[courseCode];
        legendHtml += `
            <div class="legend-item">
                <div class="legend-color course-${courseCode}"></div>
                <span class="legend-course-code">${courseCode}</span>
                <span class="legend-course-name">${courseInfo.name} (${courseInfo.credits} cr)</span>
            </div>
        `;
    });
    
    legendHtml += `
            </div>
        </div>
    `;
    
    return legendHtml;
}

function enhanceTables() {
    document.querySelectorAll('.timetable-table').forEach(table => {
        // Apply color coding
        applyColorCoding(table);
        
        // Find the parent timetable card and add legend
        const timetableCard = table.closest('.timetable-card') || table.closest('.timetable-item');
        if (timetableCard) {
            const header = timetableCard.querySelector('.timetable-header h3');
            if (header) {
                const match = header.textContent.match(/Semester (\d+) - Section ([AB])/);
                if (match) {
                    const semester = parseInt(match[1]);
                    const section = match[2];
                    const legend = createLegend(semester, section);
                    if (legend) {
                        // Remove existing legend if any
                        const existingLegend = timetableCard.querySelector('.timetable-legend');
                        if (existingLegend) {
                            existingLegend.remove();
                        }
                        // Add new legend
                        timetableCard.insertAdjacentHTML('beforeend', legend);
                    }
                }
            }
        }
    });
    
    // Add hover effects to colored cells
    const coloredCells = document.querySelectorAll('td[class*="course-"]');
    coloredCells.forEach(cell => {
        cell.addEventListener('mouseenter', function() {
            this.style.transform = 'scale(1.05)';
            this.style.zIndex = '2';
            this.style.boxShadow = '0 4px 12px rgba(0,0,0,0.3)';
        });
        
        cell.addEventListener('mouseleave', function() {
            this.style.transform = 'scale(1)';
            this.style.boxShadow = 'none';
        });
        
        // Add click handler for course info
        cell.addEventListener('click', function() {
            const courseCode = extractCourseCode(this.textContent);
            if (courseCode && courseDatabase[courseCode]) {
                const courseInfo = courseDatabase[courseCode];
                showNotification(`${courseCode}: ${courseInfo.name} - ${courseInfo.credits} credits (${courseInfo.type})`, 'info');
            }
        });
    });
}

// Rendering Functions
function renderTimetables() {
    console.log(`🎨 Rendering ${currentTimetables.length} timetables with filters:`, {
        semester: currentSemesterFilter,
        section: currentSectionFilter,
        view: currentView
    });
    
    if (currentTimetables.length === 0) {
        emptyState.style.display = 'block';
        timetablesContainer.innerHTML = '';
        return;
    }
    
    emptyState.style.display = 'none';
    
    const filteredTimetables = filterTimetablesData();
    console.log(`🔍 Filtered to ${filteredTimetables.length} timetables`);
    
    if (filteredTimetables.length === 0) {
        timetablesContainer.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">
                    <i class="fas fa-search"></i>
                </div>
                <h3>No Timetables Found</h3>
                <p>No timetables match the current filters. Try adjusting your selection.</p>
                <button class="btn btn-outline" onclick="currentSemesterFilter='all'; currentSectionFilter='all'; semesterFilter.value='all'; sectionFilter.value='all'; updateSectionTitle(); renderTimetables();">
                    Show All Timetables
                </button>
            </div>
        `;
        return;
    }
    
    if (currentView === 'grid') {
        renderGridView(filteredTimetables);
    } else if (currentView === 'list') {
        renderListView(filteredTimetables);
    } else if (currentView === 'compact') {
        renderCompactView(filteredTimetables);
    }
    
    console.log("✅ Timetables rendered successfully");
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
    
    // Enhance tables with color coding and legends
    enhanceTables();
}

function renderListView(timetables) {
    let html = '<div class="timetables-list">';
    
    timetables.forEach(timetable => {
        html += `
            <div class="timetable-item" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); margin-bottom: 1.5rem; overflow: hidden;">
                <div class="timetable-header" style="background: linear-gradient(135deg, var(--primary), var(--primary-dark)); color: white; padding: 1.25rem; display: flex; justify-content: space-between; align-items: center;">
                    <h3 style="margin: 0; font-size: 1.1rem; font-weight: 600;">Semester ${timetable.semester} - Section ${timetable.section}</h3>
                    <div class="timetable-actions">
                        <button class="btn btn-outline" onclick="downloadTimetable('${timetable.filename}')" style="background: rgba(255, 255, 255, 0.2); color: white; border: 1px solid rgba(255, 255, 255, 0.5);">
                            <i class="fas fa-download"></i> Download
                        </button>
                    </div>
                </div>
                <div class="timetable-content" style="padding: 1rem;">
                    ${timetable.html}
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    timetablesContainer.innerHTML = html;
    
    // Enhance tables with color coding and legends
    enhanceTables();
}

function renderCompactView(timetables) {
    let html = '<div class="timetables-compact">';
    
    timetables.forEach(timetable => {
        html += `
            <div class="compact-card" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); padding: 1.5rem; margin-bottom: 1rem;">
                <div style="display: flex; justify-content: between; align-items: center; margin-bottom: 1rem;">
                    <h4 style="margin: 0; color: var(--dark);">Semester ${timetable.semester} - Section ${timetable.section}</h4>
                    <div>
                        <button class="btn btn-outline btn-sm" onclick="downloadTimetable('${timetable.filename}')" style="padding: 0.25rem 0.5rem; font-size: 0.8rem;">
                            <i class="fas fa-download"></i>
                        </button>
                    </div>
                </div>
                <div style="overflow-x: auto;">
                    ${timetable.html}
                </div>
            </div>
        `;
    });
    
    html += '</div>';
    timetablesContainer.innerHTML = html;
    
    // Enhance tables with color coding and legends
    enhanceTables();
}

// Filtering Functions
function filterTimetablesData() {
    return currentTimetables.filter(timetable => {
        const semesterMatch = currentSemesterFilter === 'all' || timetable.semester === parseInt(currentSemesterFilter);
        const sectionMatch = currentSectionFilter === 'all' || timetable.section === currentSectionFilter;
        return semesterMatch && sectionMatch;
    });
}

function changeViewMode() {
    currentView = viewMode.value;
    renderTimetables();
}

// Action Functions
async function downloadAllTimetables() {
    try {
        showNotification('📦 Preparing download...', 'info');
        
        const response = await fetch('/download-all');
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.message || 'Failed to download files');
        }
        
        const blob = await response.blob();
        
        if (blob.size === 0) {
            throw new Error('Download file is empty - no timetables available');
        }
        
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.style.display = 'none';
        a.href = url;
        a.download = 'all_timetables.zip';
        
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
        
        showNotification('✅ All timetables downloaded successfully!', 'success');
    } catch (error) {
        showNotification('❌ Error downloading timetables: ' + error.message, 'error');
    }
}

function downloadTimetable(filename) {
    window.open(`/download/${filename}`, '_blank');
    showNotification(`📥 Downloading ${filename}...`, 'info');
}

function printTimetable(semester, section) {
    showNotification(`🖨️ Printing Semester ${semester} - Section ${section}`, 'info');
    
    // Create a print-friendly version
    const printWindow = window.open('', '_blank');
    const timetable = currentTimetables.find(t => t.semester === semester && t.section === section);
    
    if (timetable) {
        printWindow.document.write(`
            <!DOCTYPE html>
            <html>
            <head>
                <title>Semester ${semester} - Section ${section} Timetable</title>
                <style>
                    body { font-family: Arial, sans-serif; margin: 20px; }
                    h1 { color: #333; text-align: center; }
                    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
                    th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
                    th { background-color: #f5f5f5; font-weight: bold; }
                    .timetable-header { background: #4361ee; color: white; padding: 15px; text-align: center; }
                </style>
            </head>
            <body>
                <div class="timetable-header">
                    <h1>IIIT Dharwad - Semester ${semester} - Section ${section}</h1>
                    <p>Generated on ${new Date().toLocaleDateString()}</p>
                </div>
                ${timetable.html}
                <script>
                    window.onload = function() { window.print(); }
                </script>
            </body>
            </html>
        `);
        printWindow.document.close();
    }
}

function printAllTimetables() {
    showNotification('🖨️ Preparing all timetables for printing...', 'info');
    
    const printWindow = window.open('', '_blank');
    const filteredTimetables = filterTimetablesData();
    
    let printContent = `
        <!DOCTYPE html>
        <html>
        <head>
            <title>All Timetables - IIIT Dharwad</title>
            <style>
                body { font-family: Arial, sans-serif; margin: 20px; }
                h1 { color: #333; text-align: center; }
                h2 { color: #4361ee; margin-top: 30px; }
                table { width: 100%; border-collapse: collapse; margin: 20px 0; page-break-inside: avoid; }
                th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
                th { background-color: #f5f5f5; font-weight: bold; }
                .timetable-section { margin-bottom: 40px; }
                @media print {
                    .timetable-section { page-break-after: always; }
                }
            </style>
        </head>
        <body>
            <h1>IIIT Dharwad - All Timetables</h1>
            <p>Generated on ${new Date().toLocaleDateString()}</p>
    `;
    
    filteredTimetables.forEach(timetable => {
        printContent += `
            <div class="timetable-section">
                <h2>Semester ${timetable.semester} - Section ${timetable.section}</h2>
                ${timetable.html}
            </div>
        `;
    });
    
    printContent += `
            <script>
                window.onload = function() { window.print(); }
            </script>
        </body>
        </html>
    `;
    
    printWindow.document.write(printContent);
    printWindow.document.close();
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
    showNotification('🔄 Data refreshed successfully!', 'success');
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
    
    // Auto remove after 5 seconds
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
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

// Debug function to check current state
function debugApp() {
    console.log('=== DEBUG STATE ===');
    console.log('Current timetables:', currentTimetables);
    console.log('Current filters - Semester:', currentSemesterFilter, 'Section:', currentSectionFilter);
    console.log('Current view:', currentView);
    console.log('Filtered timetables:', filterTimetablesData());
    console.log('Course database:', courseDatabase);
    console.log('===================');
}

// Export functions for global access
window.downloadTimetable = downloadTimetable;
window.printTimetable = printTimetable;
window.debugApp = debugApp;