// Global variables
let currentTimetables = [];
let currentView = 'grid';
let currentSemesterFilter = 'all';
let currentSectionFilter = 'all';

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
    
    // Load initial data
    loadStats();
    loadTimetables();
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
        const response = await fetch('/generate', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const result = await response.json();
        
        if (result.success) {
            showNotification('Timetables generated successfully!', 'success');
            console.log('Generated files:', result.files);
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
        console.log("🔄 Loading timetables...");
        const response = await fetch('/timetables');
        
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        
        const timetables = await response.json();
        
        console.log(`📊 Received ${timetables.length} timetables:`, timetables);
        
        currentTimetables = timetables;
        renderTimetables();
        
        // Show notification if no timetables
        if (timetables.length === 0) {
            showNotification('No timetables found. Please generate them first.', 'info');
        }
        
    } catch (error) {
        console.error('❌ Error loading timetables:', error);
        showNotification('Error loading timetables: ' + error.message, 'error');
        
        // Try debug endpoint
        try {
            const debugResponse = await fetch('/debug-timetables');
            const debugInfo = await debugResponse.json();
            console.log('🐛 Debug info:', debugInfo);
        } catch (debugError) {
            console.error('Debug endpoint failed:', debugError);
        }
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
    
    if (filteredTimetables.length === 0) {
        timetablesContainer.innerHTML = `
            <div class="empty-state">
                <div class="empty-icon">
                    <i class="fas fa-search"></i>
                </div>
                <h3>No Timetables Found</h3>
                <p>No timetables match the current filters. Try adjusting your selection.</p>
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
    
    // Enhance tables after rendering
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
    
    // Enhance tables after rendering
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
    
    // Enhance tables after rendering
    enhanceTables();
}

function enhanceTables() {
    // Add hover effects and better styling to tables
    document.querySelectorAll('.timetable-table').forEach(table => {
        // Ensure proper styling
        table.style.width = '100%';
        table.style.borderCollapse = 'collapse';
        
        // Add hover effects to cells
        const cells = table.querySelectorAll('td:not(:first-child)');
        cells.forEach(cell => {
            if (cell.textContent.trim() && cell.textContent.trim() !== 'Free' && cell.textContent.trim() !== 'LUNCH BREAK') {
                cell.style.cursor = 'pointer';
                cell.addEventListener('mouseenter', function() {
                    this.style.transform = 'scale(1.02)';
                    this.style.zIndex = '1';
                    this.style.boxShadow = '0 2px 8px rgba(0,0,0,0.15)';
                });
                cell.addEventListener('mouseleave', function() {
                    this.style.transform = 'scale(1)';
                    this.style.boxShadow = 'none';
                });
            }
        });
    });
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
        showNotification('Preparing download...', 'info');
        
        const response = await fetch('/download-all');
        if (!response.ok) {
            throw new Error('Failed to download files');
        }
        
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
    showNotification(`Downloading ${filename}...`, 'info');
}

function printTimetable(semester, section) {
    showNotification(`Printing Semester ${semester} - Section ${section}`, 'info');
    
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
    showNotification('Preparing all timetables for printing...', 'info');
    
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
            // You could redirect to a help page or open a modal
            break;
        case 'feedback':
            showNotification('Feedback form would open here', 'info');
            // You could open a feedback modal or redirect
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
function debugState() {
    console.log('=== DEBUG STATE ===');
    console.log('Current timetables:', currentTimetables);
    console.log('Current filters - Semester:', currentSemesterFilter, 'Section:', currentSectionFilter);
    console.log('Current view:', currentView);
    console.log('Filtered timetables:', filterTimetablesData());
    console.log('===================');
}

// Export functions for global access
window.downloadTimetable = downloadTimetable;
window.printTimetable = printTimetable;
window.debugState = debugState;