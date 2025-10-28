// Global variables
let currentTimetables = [];
let currentView = 'grid';
let currentSemesterFilter = 'all';
let currentSectionFilter = 'all';
let currentBranchFilter = 'all';
let uploadedFiles = [];
let isUploadSectionVisible = false;

// Course information database - will be populated from server data
let courseDatabase = {};

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

// Filter elements
const branchFilter = document.getElementById('branch-filter');
const semesterFilter = document.getElementById('semester-filter');
const sectionFilter = document.getElementById('section-filter');
const viewMode = document.getElementById('view-mode');

// Initialize application
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
});

// Settings Initialization Function
function setupSettings() {
    const settingsBtn = document.getElementById('settings-nav-btn');
    if (settingsBtn) {
        settingsBtn.addEventListener('click', function(e) {
            e.preventDefault();
            showSettingsModal();
        });
    }

    // Also make the quick action settings button work
    const settingsActionCard = document.querySelector('.action-card[data-action="settings"]');
    if (settingsActionCard) {
        settingsActionCard.addEventListener('click', function() {
            showSettingsModal();
        });
    }

    console.log("✅ Settings system initialized");
}

function initializeApp() {
    console.log("🚀 Initializing Timetable Application...");
    
    // Event listeners
    generateBtn.addEventListener('click', generateTimetables);
    refreshBtn.addEventListener('click', refreshAll);
    downloadAllBtn.addEventListener('click', downloadAllTimetables);
    emptyGenerateBtn.addEventListener('click', generateTimetables);
    
    // Initialize all systems in correct order
    setupFileUpload();
    setupSettings(); // This should work now
    setupHelpSupport(); // Make sure this exists too
    
    // Filter listeners
    branchFilter.addEventListener('change', function() {
        currentBranchFilter = this.value;
        updateSectionTitle();
        renderTimetables();
    });
    
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
    
    // Upload files button
    const uploadFilesBtn = document.getElementById('upload-files-btn');
    if (uploadFilesBtn) {
        uploadFilesBtn.addEventListener('click', showUploadSection);
    }
    
    // Apply saved settings
    initializeSettings();
    
    // Load initial data
    loadStats();
    loadTimetables();
    
    console.log("✅ Application initialized successfully");
}

// File Upload Functions
function setupFileUpload() {
    const uploadArea = document.getElementById('upload-area');
    const fileInput = document.getElementById('file-input');
    const browseBtn = document.getElementById('browse-btn');
    const cancelUploadBtn = document.getElementById('cancel-upload-btn');
    const processFilesBtn = document.getElementById('process-files-btn');
    const uploadNavBtn = document.getElementById('upload-nav-btn');

    // Browse button click
    if (browseBtn) {
        browseBtn.addEventListener('click', () => fileInput.click());
    }

    // File input change
    if (fileInput) {
        fileInput.addEventListener('change', handleFileSelect);
    }

    // Drag and drop events
    if (uploadArea) {
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.add('drag-over');
            }, false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            uploadArea.addEventListener(eventName, () => {
                uploadArea.classList.remove('drag-over');
            }, false);
        });

        // Drop event
        uploadArea.addEventListener('drop', handleDrop, false);
    }

    // Cancel upload
    if (cancelUploadBtn) {
        cancelUploadBtn.addEventListener('click', hideUploadSection);
    }

    // Process files
    if (processFilesBtn) {
        processFilesBtn.addEventListener('click', processUploadedFiles);
    }

    // Navigation upload button
    if (uploadNavBtn) {
        uploadNavBtn.addEventListener('click', function(e) {
            e.preventDefault();
            showUploadSection();
        });
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    handleFiles(files);
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
}

function handleFiles(files) {
    const validFiles = Array.from(files).filter(file => {
        const isValidCSV = file.name.toLowerCase().endsWith('.csv');
        if (!isValidCSV) {
            showNotification(`❌ ${file.name} is not a CSV file`, 'error');
            return false;
        }
        return true;
    });

    if (validFiles.length > 0) {
        uploadedFiles = [...uploadedFiles, ...validFiles];
        updateFileList();
        updateUploadArea();
        updateProcessButton();
        showNotification(`✅ Added ${validFiles.length} file(s)`, 'success');
    }
}

function updateFileList() {
    const uploadedFilesContainer = document.getElementById('uploaded-files');
    const requiredFiles = document.querySelectorAll('.file-item.required');
    
    if (!uploadedFilesContainer) return;
    
    // Clear existing uploaded files display
    const existingUploaded = uploadedFilesContainer.querySelectorAll('.uploaded-file-item');
    existingUploaded.forEach(item => item.remove());
    
    // Update required files status
    requiredFiles.forEach(item => {
        const fileName = item.dataset.file;
        const fileStatus = item.querySelector('.file-status');
        const isUploaded = uploadedFiles.some(file => 
            file.name.toLowerCase().replace(/[ _-]/g, '') === fileName.toLowerCase().replace(/[ _-]/g, '') ||
            fileName.toLowerCase().replace(/[ _-]/g, '').includes(file.name.toLowerCase().replace(/[ _-]/g, '')) ||
            file.name.toLowerCase().replace(/[ _-]/g, '').includes(fileName.toLowerCase().replace(/[ _-]/g, ''))
        );
        
        if (isUploaded) {
            item.classList.add('uploaded');
            fileStatus.textContent = 'Uploaded';
            fileStatus.classList.add('uploaded');
            fileStatus.classList.remove('missing');
        } else {
            item.classList.remove('uploaded');
            fileStatus.textContent = 'Missing';
            fileStatus.classList.add('missing');
            fileStatus.classList.remove('uploaded');
        }
    });
    
    // Display uploaded files
    uploadedFiles.forEach((file, index) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'uploaded-file-item';
        fileItem.innerHTML = `
            <i class="fas fa-file-csv"></i>
            <div class="file-info">
                <div class="file-name">${file.name}</div>
                <div class="file-size">${formatFileSize(file.size)}</div>
            </div>
            <button class="remove-file" data-index="${index}">
                <i class="fas fa-times"></i>
            </button>
        `;
        uploadedFilesContainer.appendChild(fileItem);
    });
    
    // Add event listeners to remove buttons
    document.querySelectorAll('.remove-file').forEach(btn => {
        btn.addEventListener('click', function() {
            const index = parseInt(this.dataset.index);
            removeFile(index);
        });
    });
}

function removeFile(index) {
    const removedFile = uploadedFiles[index];
    uploadedFiles.splice(index, 1);
    updateFileList();
    updateUploadArea();
    updateProcessButton();
    showNotification(`🗑️ Removed ${removedFile.name}`, 'info');
}

function updateUploadArea() {
    const uploadArea = document.getElementById('upload-area');
    if (!uploadArea) return;
    
    if (uploadedFiles.length > 0) {
        uploadArea.classList.add('has-files');
        const uploadContent = uploadArea.querySelector('.upload-content');
        uploadContent.innerHTML = `
            <i class="fas fa-check-circle"></i>
            <h3>${uploadedFiles.length} File(s) Ready</h3>
            <p>Drag and drop more files or click to browse</p>
            <button class="btn btn-outline" id="browse-btn">
                <i class="fas fa-folder-open"></i>
                Add More Files
            </button>
        `;
        
        // Reattach browse button event listener
        const newBrowseBtn = uploadArea.querySelector('#browse-btn');
        newBrowseBtn.addEventListener('click', () => {
            document.getElementById('file-input').click();
        });
    } else {
        uploadArea.classList.remove('has-files');
        const uploadContent = uploadArea.querySelector('.upload-content');
        uploadContent.innerHTML = `
            <i class="fas fa-cloud-upload-alt"></i>
            <h3>Drag & Drop Files Here</h3>
            <p>Supported files: CSV</p>
            <button class="btn btn-outline" id="browse-btn">
                <i class="fas fa-folder-open"></i>
                Browse Files
            </button>
        `;
        
        // Reattach browse button event listener
        const newBrowseBtn = uploadArea.querySelector('#browse-btn');
        newBrowseBtn.addEventListener('click', () => {
            document.getElementById('file-input').click();
        });
    }
}

function updateProcessButton() {
    const processFilesBtn = document.getElementById('process-files-btn');
    if (!processFilesBtn) return;
    
    const requiredFiles = [
        'course_data.csv',
        'faculty_availability.csv', 
        'classroom_data.csv',
        'student_data.csv',
        'exams_data.csv'
    ];
    
    const hasAllRequired = requiredFiles.every(requiredFile => {
        // More flexible matching for file names
        const requiredFileClean = requiredFile.toLowerCase().replace(/[ _-]/g, '');
        return uploadedFiles.some(file => {
            const fileNameClean = file.name.toLowerCase().replace(/[ _-]/g, '');
            return fileNameClean.includes(requiredFileClean) || requiredFileClean.includes(fileNameClean);
        });
    });
    
    processFilesBtn.disabled = !hasAllRequired;
    
    if (hasAllRequired) {
        processFilesBtn.innerHTML = '<i class="fas fa-cog"></i> Process Files & Generate Timetables';
    } else {
        processFilesBtn.innerHTML = '<i class="fas fa-cog"></i> Process Files';
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

async function processUploadedFiles() {
    if (uploadedFiles.length === 0) {
        showNotification('❌ No files to process', 'error');
        return;
    }

    showUploadProgress(true);
    
    try {
        const formData = new FormData();
        uploadedFiles.forEach(file => {
            formData.append('files', file);
        });

        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();
        
        if (result.success) {
            showNotification(`✅ ${result.message}`, 'success');
            console.log('📁 Uploaded files:', result.uploaded_files);
            console.log('📊 Generated timetables:', result.generated_count);
            
            // Clear uploaded files
            uploadedFiles = [];
            updateFileList();
            updateUploadArea();
            updateProcessButton();
            
            // Hide upload section
            hideUploadSection();
            
            // Reload timetables and stats
            await loadTimetables();
            await loadStats();
            
            // Verify data was loaded correctly
            await verifyDataLoad();
            
        } else {
            showNotification(`❌ ${result.message}`, 'error');
            // Show available files for debugging
            if (result.available_files) {
                console.log('Available files:', result.available_files);
            }
        }
    } catch (error) {
        console.error('❌ Error uploading files:', error);
        showNotification('❌ Error uploading files: ' + error.message, 'error');
    } finally {
        showUploadProgress(false);
    }
}

function showUploadSection() {
    const uploadSection = document.getElementById('upload-section');
    const timetablesSection = document.querySelector('.timetables-section');
    const controlsSection = document.querySelector('.controls-section');
    const quickActions = document.querySelector('.quick-actions');
    
    if (uploadSection) {
        uploadSection.style.display = 'block';
    }
    if (timetablesSection) timetablesSection.style.display = 'none';
    if (controlsSection) controlsSection.style.display = 'none';
    if (quickActions) quickActions.style.display = 'none';
    
    isUploadSectionVisible = true;
    
    // Update navigation
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
    });
    const uploadNavBtn = document.querySelector('#upload-nav-btn');
    if (uploadNavBtn) {
        uploadNavBtn.parentElement.classList.add('active');
    }
}

function hideUploadSection() {
    const uploadSection = document.getElementById('upload-section');
    const timetablesSection = document.querySelector('.timetables-section');
    const controlsSection = document.querySelector('.controls-section');
    const quickActions = document.querySelector('.quick-actions');
    
    if (uploadSection) {
        uploadSection.style.display = 'none';
    }
    if (timetablesSection) timetablesSection.style.display = 'block';
    if (controlsSection) controlsSection.style.display = 'flex';
    if (quickActions) quickActions.style.display = 'block';
    
    isUploadSectionVisible = false;
    
    // Reset to dashboard view
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
    });
    const dashboardNav = document.querySelector('[data-semester="all"]');
    if (dashboardNav) {
        dashboardNav.parentElement.classList.add('active');
    }
}

function showUploadProgress(show) {
    const uploadOverlay = document.getElementById('upload-overlay');
    const progressFill = document.getElementById('upload-progress-fill');
    const progressText = document.getElementById('upload-progress-text');
    
    if (uploadOverlay && progressFill && progressText) {
        if (show) {
            uploadOverlay.classList.add('active');
            // Simulate upload progress
            simulateUploadProgress(progressFill, progressText);
        } else {
            uploadOverlay.classList.remove('active');
            progressFill.style.width = '0%';
            progressText.textContent = '0%';
        }
    }
}

function simulateUploadProgress(progressFill, progressText) {
    let progress = 0;
    
    const interval = setInterval(() => {
        if (progress < 90) {
            progress += Math.random() * 15;
            progress = Math.min(progress, 90);
            progressFill.style.width = progress + '%';
            progressText.textContent = Math.round(progress) + '%';
        }
    }, 200);
    
    // Clear interval when upload is done
    setTimeout(() => {
        clearInterval(interval);
        progressFill.style.width = '100%';
        progressText.textContent = '100%';
    }, 2000);
}

// Debug Functions
// Debug Functions
async function verifyDataLoad() {
    try {
        const response = await fetch('/debug/current-data');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        console.log('📊 Current loaded data:', data);
        return data;
    } catch (error) {
        console.error('Error verifying data:', error);
        // Show user-friendly message
        if (error.message.includes('404')) {
            console.log('🔧 Debug endpoints not available - this is normal in production');
            showNotification('🔧 Debug features not available', 'info');
        } else {
            showNotification('❌ Error verifying data: ' + error.message, 'error');
        }
        return null;
    }
}

async function clearCache() {
    try {
        const response = await fetch('/debug/clear-cache');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const result = await response.json();
        console.log('🗑️ Cache cleared:', result);
        showNotification('✅ Cache cleared successfully', 'success');
        return result;
    } catch (error) {
        console.error('Error clearing cache:', error);
        if (error.message.includes('404')) {
            console.log('🔧 Debug endpoints not available - this is normal in production');
            showNotification('🔧 Debug features not available', 'info');
        } else {
            showNotification('❌ Error clearing cache: ' + error.message, 'error');
        }
        return null;
    }
}

async function debugFileMatching() {
    try {
        const response = await fetch('/debug/file-matching');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const result = await response.json();
        console.log('🔍 File matching debug:', result);
        return result;
    } catch (error) {
        console.error('Error debugging file matching:', error);
        if (error.message.includes('404')) {
            console.log('🔧 Debug endpoints not available');
            // Fallback to client-side file matching debug
            debugFileMatchingClient();
        } else {
            showNotification('❌ Error debugging file matching: ' + error.message, 'error');
        }
        return null;
    }
}

function debugFileMatchingClient() {
    console.log('=== CLIENT-SIDE FILE MATCHING DEBUG ===');
    const requiredFiles = [
        'course_data.csv',
        'faculty_availability.csv', 
        'classroom_data.csv',
        'student_data.csv',
        'exams_data.csv'
    ];
    
    requiredFiles.forEach(requiredFile => {
        const requiredFileClean = requiredFile.toLowerCase().replace(/[ _-]/g, '');
        console.log(`Required: ${requiredFile} -> ${requiredFileClean}`);
        
        uploadedFiles.forEach(file => {
            const fileNameClean = file.name.toLowerCase().replace(/[ _-]/g, '');
            const matches = fileNameClean.includes(requiredFileClean) || requiredFileClean.includes(fileNameClean);
            console.log(`  ${file.name} -> ${fileNameClean} : ${matches ? '✅ MATCH' : '❌ NO MATCH'}`);
        });
    });
    console.log('==========================');
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
            
            // Get semester and branch filters
            const semester = this.getAttribute('data-semester');
            const branch = this.getAttribute('data-branch');
            
            currentSemesterFilter = semester;
            currentBranchFilter = branch;
            
            // Update filter dropdowns
            if (semesterFilter) semesterFilter.value = semester;
            if (branchFilter) branchFilter.value = branch;
            
            // Hide upload section if visible
            if (isUploadSectionVisible) {
                hideUploadSection();
            }
            
            // Update section title and render
            updateSectionTitle();
            renderTimetables();
        });
    });
}

function updateSectionTitle() {
    if (sectionTitle) {
        let title = '';
        
        if (currentBranchFilter === 'all') {
            title = 'All Timetables';
        } else {
            const branchNames = {
                'CSE': 'Computer Science',
                'DSAI': 'Data Science & AI', 
                'ECE': 'Electronics & Communication'
            };
            title = `${branchNames[currentBranchFilter]} Timetables`;
        }
        
        if (currentSemesterFilter !== 'all') {
            title += ` - Semester ${currentSemesterFilter}`;
        }
        
        if (currentSectionFilter !== 'all') {
            title += ` - Section ${currentSectionFilter}`;
        }
        
        sectionTitle.textContent = title;
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
            
        } else {
            showNotification(`❌ ${result.message}`, 'error');
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
        
        // Update course database with server data
        if (timetables.length > 0 && timetables[0].course_info) {
            courseDatabase = timetables[0].course_info;
        }
        
        renderTimetables();
        
    } catch (error) {
        console.error('❌ Error loading timetables:', error);
        showNotification('❌ Error loading timetables: ' + error.message, 'error');
    }
}

async function loadStats() {
    try {
        const response = await fetch('/stats');
        const stats = await response.json();
        
        // Update stats elements
        const totalTimetablesEl = document.getElementById('total-timetables');
        const totalCoursesEl = document.getElementById('total-courses');
        const totalFacultyEl = document.getElementById('total-faculty');
        const totalClassroomsEl = document.getElementById('total-classrooms');
        
        if (totalTimetablesEl) totalTimetablesEl.textContent = stats.total_timetables;
        if (totalCoursesEl) totalCoursesEl.textContent = stats.total_courses;
        if (totalFacultyEl) totalFacultyEl.textContent = stats.total_faculty;
        if (totalClassroomsEl) totalClassroomsEl.textContent = stats.total_classrooms;
        
    } catch (error) {
        console.error('Error loading stats:', error);
    }
}

// Enhanced Time Slot Functions
function enhanceTimeSlotHeaders() {
    document.querySelectorAll('.timetable-table').forEach(table => {
        const rows = table.querySelectorAll('tr');
        
        rows.forEach((row) => {
            const timeCell = row.cells[0];
            if (timeCell && timeCell.textContent) {
                const timeText = timeCell.textContent.trim();
                
                // Skip if already processed or if it's a header row
                if (timeCell.classList.contains('time-slot-processed') || 
                    timeCell.textContent.includes('Time') || 
                    !timeText.includes('-')) {
                    return;
                }
                
                // Clear existing classes and add base classes
                timeCell.className = '';
                timeCell.classList.add('time-slot', 'time-slot-processed');
                
                let duration = '?';
                let timeClass = '';
                
                // Determine time slot type and duration
                if (timeText.includes('09:00-10:30') || timeText.includes('10:30-12:00') || 
                    timeText.includes('13:00-14:30') || timeText.includes('15:30-17:00')) {
                    // All 1.5-hour lecture slots
                    timeClass = 'lecture-slot';
                    duration = '1.5h';
                } else if (timeText.includes('14:30-15:30') || timeText.includes('17:00-18:00')) {
                    // All 1-hour tutorial slots
                    timeClass = 'tutorial-slot';
                    duration = '1h';
                } else if (timeText.includes('12:00-13:00')) {
                    // Lunch break
                    timeClass = 'lunch-break';
                    duration = '1h';
                }
                
                // Apply the determined classes
                timeClass.split(' ').forEach(cls => {
                    if (cls) timeCell.classList.add(cls);
                });
                
                // Update the cell content with duration
                timeCell.innerHTML = `<div class="time-slot-content">${timeText}<span class="slot-duration">${duration}</span></div>`;
            }
        });
    });
}

// Color Coding and Legend Functions
function applyDynamicColorCoding(tableElement, courseColors) {
    const cells = tableElement.querySelectorAll('td');
    
    cells.forEach(cell => {
        const text = cell.textContent.trim();
        
        // Skip header cells, empty cells, and special slots
        if (!text || text === 'Free' || text === 'LUNCH BREAK' || cell.cellIndex === 0) {
            cell.classList.add('empty-cell');
            return;
        }
        
        // Extract course code - handle both regular and elective/tutorial versions
        let courseCode = extractCourseCode(text);
        
        // For elective tutorials, use the base course code without "(Tutorial)"
        if (text.includes('(Tutorial)') && courseCode) {
            // Remove "(Tutorial)" suffix but keep the same course code
            const baseCourseCode = text.replace(' (Tutorial)', '');
            courseCode = extractCourseCode(baseCourseCode) || courseCode;
            cell.classList.add('tutorial'); // Add tutorial class for styling
        }
        
        // For elective lectures, use the base course code without "(Elective)"
        if (text.includes('(Elective)') && courseCode) {
            const baseCourseCode = text.replace(' (Elective)', '');
            courseCode = extractCourseCode(baseCourseCode) || courseCode;
        }
        
        if (courseCode && courseColors[courseCode]) {
            // Apply dynamic color
            const color = courseColors[courseCode];
            cell.style.background = color;
            cell.style.color = getContrastColor(color);
            cell.style.fontWeight = '600';
            cell.style.border = '2px solid white';
            
            // Add tooltip with course info
            const courseInfo = courseDatabase[courseCode];
            if (courseInfo) {
                let sessionType = 'Lecture';
                if (text.includes('(Tutorial)')) sessionType = 'Tutorial';
                if (text.includes('(Elective)')) sessionType = 'Elective Lecture';
                
                cell.title = `${courseCode}: ${courseInfo.name} (${sessionType})\n${courseInfo.credits} credits - ${courseInfo.instructor}`;
            }
            
            // Make cell clickable for more info
            cell.style.cursor = 'pointer';
            cell.classList.add('colored-cell');
            
            // Add elective indicator for both lectures and tutorials
            if (text.includes('(Elective)') || text.includes('(Tutorial)')) {
                cell.classList.add('elective-slot');
                if (text.includes('(Elective)')) {
                    cell.title += "\n⚡ Common Elective Slot (Same for both sections)";
                } else if (text.includes('(Tutorial)')) {
                    cell.title += "\n⚡ Common Tutorial Slot (Same for both sections)";
                }
            }
        }
    });
}

function getContrastColor(hexcolor) {
    // Remove the # if present
    hexcolor = hexcolor.replace("#", "");
    
    // Convert to RGB
    const r = parseInt(hexcolor.substr(0, 2), 16);
    const g = parseInt(hexcolor.substr(2, 2), 16);
    const b = parseInt(hexcolor.substr(4, 2), 16);
    
    // Calculate luminance
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    
    // Return black or white based on luminance
    return luminance > 0.5 ? '#000000' : '#FFFFFF';
}

function extractCourseCode(text) {
    // Match common course code patterns like MA101, CS101, etc.
    // Handle both regular courses and elective/tutorial marked courses
    const cleanText = text.replace(' (Elective)', '').replace(' (Tutorial)', '');
    
    // More robust course code pattern matching
    const coursePattern = /[A-Z]{2,3}\d{3}[A-Z]?/; // Handles codes like CS101, MA101A
    const match = cleanText.match(coursePattern);
    
    if (match) {
        return match[0];
    }
    
    // If no pattern match, return the original text but log for debugging
    console.log('⚠️ No course code pattern found for:', text);
    return cleanText;
}

function debugCourseDuplicates(timetable) {
    console.log('🔍 Debugging course duplicates for:', {
        semester: timetable.semester,
        section: timetable.section,
        branch: timetable.branch
    });
    
    console.log('All courses:', timetable.courses);
    console.log('Unique courses:', [...new Set(timetable.courses)]);
    console.log('Core courses:', timetable.core_courses);
    console.log('Elective courses:', timetable.elective_courses);
    
    // Check for duplicates
    const duplicates = timetable.courses.filter((course, index) => 
        timetable.courses.indexOf(course) !== index
    );
    
    if (duplicates.length > 0) {
        console.log('❌ Found duplicates:', duplicates);
    } else {
        console.log('✅ No duplicates found');
    }
}

function getUniqueCourses(courses) {
    if (!courses || !Array.isArray(courses)) return [];
    
    // Use Set for basic deduplication
    const uniqueCourses = [...new Set(courses)];
    
    // Additional filtering to ensure no duplicates
    const seen = new Set();
    const result = [];
    
    uniqueCourses.forEach(course => {
        if (course && !seen.has(course)) {
            seen.add(course);
            result.push(course);
        }
    });
    
    console.log(`🔄 Deduplicated courses: ${courses.length} -> ${result.length}`);
    return result;
}

function createEnhancedLegend(semester, section, courses, courseColors, courseInfo, coreCourses = [], electiveCourses = []) {
    if (!courses || courses.length === 0) return '';
    
    // Get unique courses with aggressive deduplication
    const uniqueCourses = getUniqueCourses(courses);
    
    // Separate courses into core and elective for better organization
    const uniqueCoreCourses = getUniqueCourses(coreCourses);
    const uniqueElectiveCourses = getUniqueCourses(electiveCourses);
    
    // Filter to only include courses that actually exist in the timetable
    const coreCourseList = uniqueCoreCourses.filter(course => uniqueCourses.includes(course));
    const electiveCourseList = uniqueElectiveCourses.filter(course => uniqueCourses.includes(course));
    const otherCourses = uniqueCourses.filter(course => 
        !coreCourseList.includes(course) && !electiveCourseList.includes(course)
    );
    
    console.log(`📊 Legend for Semester ${semester}, Section ${section}:`, {
        allCourses: uniqueCourses,
        coreCourses: coreCourseList,
        electiveCourses: electiveCourseList,
        otherCourses: otherCourses
    });
    
    // If no courses to display, return empty
    if (coreCourseList.length === 0 && electiveCourseList.length === 0 && otherCourses.length === 0) {
        console.log('⚠️ No courses to display in legend');
        return '';
    }
    
    let legendHtml = `
        <div class="timetable-legend">
            <div class="legend-title">
                <i class="fas fa-palette"></i>
                Course Legend - Semester ${semester}, Section ${section}
            </div>
    `;
    
    // Core Courses Section
    if (coreCourseList.length > 0) {
        legendHtml += `
            <div class="legend-section">
                <div class="legend-section-title">
                    <i class="fas fa-book"></i>
                    Core Courses (${coreCourseList.length})
                </div>
                <div class="legend-grid">
        `;
        
        coreCourseList.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color, 'core');
        });
        
        legendHtml += `
                </div>
            </div>
        `;
    }
    
    // Elective Courses Section
    if (electiveCourseList.length > 0) {
        legendHtml += `
            <div class="legend-section elective">
                <div class="legend-section-title elective">
                    <i class="fas fa-clipboard-list"></i>
                    Elective Courses (${electiveCourseList.length}) - Common for Both Sections
                </div>
                <div class="legend-grid">
        `;
        
        electiveCourseList.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color, 'elective');
        });
        
        legendHtml += `
                </div>
                <div style="margin-top: 0.5rem; font-size: 0.8rem; color: var(--gray);">
                    <i class="fas fa-info-circle"></i>
                    Electives have 2 lectures + 1 tutorial (same slots for both sections)
                </div>
            </div>
        `;
    }
    
    // Other Courses Section (if any)
    if (otherCourses.length > 0) {
        legendHtml += `
            <div class="legend-section">
                <div class="legend-section-title">
                    <i class="fas fa-graduation-cap"></i>
                    Other Courses (${otherCourses.length})
                </div>
                <div class="legend-grid">
        `;
        
        otherCourses.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color, 'other');
        });
        
        legendHtml += `
                </div>
            </div>
        `;
    }
    
    legendHtml += `</div>`;
    return legendHtml;
}

function createLegendItem(courseCode, courseInfo, color, type = 'core') {
    const courseName = courseInfo ? courseInfo.name : 'Unknown Course';
    const credits = courseInfo ? courseInfo.credits : '?';
    const instructor = courseInfo ? courseInfo.instructor : 'Unknown';
    const courseType = courseInfo ? courseInfo.type : 'Core';
    
    let additionalInfo = '';
    if (type === 'elective') {
        additionalInfo = '<br><small style="color: var(--secondary);">⚡ 2 lectures + 1 tutorial (Common slots)</small>';
    }
    
    return `
        <div class="legend-item ${type}">
            <div class="legend-color" style="background: ${color};"></div>
            <span class="legend-course-code">${courseCode}</span>
            <span class="legend-course-name">
                ${courseName} (${credits} cr)
                <br><small>${instructor} • ${courseType}</small>
                ${additionalInfo}
            </span>
        </div>
    `;
}

function enhanceTables() {
    console.log("🎨 Enhancing tables with color coding and time slots...");
    
    document.querySelectorAll('.timetable-table').forEach(table => {
        // Find the parent timetable card and get course information
        const timetableCard = table.closest('.timetable-card') || table.closest('.timetable-item');
        if (timetableCard) {
            const header = timetableCard.querySelector('.timetable-header h3');
            if (header) {
                const match = header.textContent.match(/Semester (\d+) - Section ([AB])/);
                if (match) {
                    const semester = parseInt(match[1]);
                    const section = match[2];
                    
                    // Find the timetable data
                    const timetable = currentTimetables.find(t => 
                        t.semester === semester && t.section === section
                    );
                    
                    if (timetable) {
                        // Debug duplicates
                        debugCourseDuplicates(timetable);
                        
                        if (timetable.course_colors) {
                            // Apply dynamic color coding
                            applyDynamicColorCoding(table, timetable.course_colors);
                            
                            // Create enhanced legend with course type separation
                            if (timetable.courses) {
                                const legend = createEnhancedLegend(
                                    semester, 
                                    section, 
                                    timetable.courses, 
                                    timetable.course_colors, 
                                    timetable.course_info,
                                    timetable.core_courses || [],
                                    timetable.elective_courses || []
                                );
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
                }
            }
        }
    });
    
    // Enhance time slot headers
    enhanceTimeSlotHeaders();
    
    // Add enhanced hover effects
    const electiveCells = document.querySelectorAll('.elective-slot');
    electiveCells.forEach(cell => {
        cell.addEventListener('mouseenter', function() {
            this.style.transform = 'scale(1.08)';
            this.style.zIndex = '3';
            this.style.boxShadow = '0 6px 20px rgba(114, 9, 183, 0.4)';
        });
        
        cell.addEventListener('mouseleave', function() {
            this.style.transform = 'scale(1)';
            this.style.boxShadow = 'none';
        });
    });
    
    // Add hover effects and click handlers to all colored cells
    const coloredCells = document.querySelectorAll('.colored-cell');
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
                const isElective = this.classList.contains('elective-slot');
                const electiveText = isElective ? '<br><strong>⚡ Common Elective Slot (Same for both sections)</strong>' : '';
                
                showNotification(
                    `<strong>${courseCode}: ${courseInfo.name}</strong><br>
                    Credits: ${courseInfo.credits}<br>
                    Type: ${courseInfo.type}<br>
                    Instructor: ${courseInfo.instructor}<br>
                    Department: ${courseInfo.department}${electiveText}`, 
                    'info'
                );
            }
        });
    });
}

// Rendering Functions
function renderTimetables() {
    console.log(`🎨 Rendering ${currentTimetables.length} timetables with filters:`, {
        branch: currentBranchFilter,
        semester: currentSemesterFilter,
        section: currentSectionFilter,
        view: currentView
    });
    
    if (!timetablesContainer) return;
    
    if (currentTimetables.length === 0) {
        if (emptyState) emptyState.style.display = 'block';
        timetablesContainer.innerHTML = '';
        return;
    }
    
    if (emptyState) emptyState.style.display = 'none';
    
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
                <button class="btn btn-outline" onclick="resetFilters()">
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
    
    // Enhance tables with color coding and legends
    enhanceTables();
}

function resetFilters() {
    currentBranchFilter = 'all';
    currentSemesterFilter = 'all';
    currentSectionFilter = 'all';
    
    if (branchFilter) branchFilter.value = 'all';
    if (semesterFilter) semesterFilter.value = 'all';
    if (sectionFilter) sectionFilter.value = 'all';
    
    updateSectionTitle();
    renderTimetables();
}

function renderGridView(timetables) {
    let html = '<div class="timetables-grid">';
    
    timetables.forEach(timetable => {
        const branchClass = `branch-${timetable.branch?.toLowerCase() || 'general'}`;
        const branchBadge = timetable.branch ? `<span class="branch-badge ${timetable.branch.toLowerCase()}">${timetable.branch}</span>` : '';
        
        html += `
            <div class="timetable-card ${branchClass}">
                <div class="timetable-header">
                    <h3>Semester ${timetable.semester} - Section ${timetable.section} ${branchBadge}</h3>
                    <div class="timetable-actions">
                        <button class="action-btn" onclick="downloadTimetable('${timetable.filename}')" title="Download">
                            <i class="fas fa-download"></i>
                        </button>
                        <button class="action-btn" onclick="printTimetable(${timetable.semester}, '${timetable.section}', '${timetable.branch || ''}')" title="Print">
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
        const branchBadge = timetable.branch ? `<span class="branch-badge ${timetable.branch.toLowerCase()}">${timetable.branch}</span>` : '';
        
        html += `
            <div class="timetable-item" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); margin-bottom: 1.5rem; overflow: hidden;">
                <div class="timetable-header" style="background: linear-gradient(135deg, var(--primary), var(--primary-dark)); color: white; padding: 1.25rem; display: flex; justify-content: space-between; align-items: center;">
                    <h3 style="margin: 0; font-size: 1.1rem; font-weight: 600;">Semester ${timetable.semester} - Section ${timetable.section} ${branchBadge}</h3>
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
}

function renderCompactView(timetables) {
    let html = '<div class="timetables-compact">';
    
    timetables.forEach(timetable => {
        const branchBadge = timetable.branch ? `<span class="branch-badge ${timetable.branch.toLowerCase()}">${timetable.branch}</span>` : '';
        
        html += `
            <div class="compact-card" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); padding: 1.5rem; margin-bottom: 1rem;">
                <div style="display: flex; justify-content: between; align-items: center; margin-bottom: 1rem;">
                    <h4 style="margin: 0; color: var(--dark);">Semester ${timetable.semester} - Section ${timetable.section} ${branchBadge}</h4>
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
}

// Filtering Functions
function filterTimetablesData() {
    return currentTimetables.filter(timetable => {
        const branchMatch = currentBranchFilter === 'all' || timetable.branch === currentBranchFilter;
        const semesterMatch = currentSemesterFilter === 'all' || timetable.semester === parseInt(currentSemesterFilter);
        const sectionMatch = currentSectionFilter === 'all' || timetable.section === currentSectionFilter;
        return branchMatch && semesterMatch && sectionMatch;
    });
}

function changeViewMode() {
    if (viewMode) {
        currentView = viewMode.value;
        renderTimetables();
    }
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
                    .elective-slot { border: 2px dashed #7209b7 !important; background: #f8f8f8 !important; }
                    .time-slot { background: #f5f5f5 !important; font-weight: bold; text-align: right !important; padding-right: 12px !important; }
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
                .elective-slot { border: 2px dashed #7209b7 !important; background: #f8f8f8 !important; }
                .time-slot { background: #f5f5f5 !important; font-weight: bold; text-align: right !important; padding-right: 12px !important; }
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
        case 'upload':
            showUploadSection();
            break;
        case 'export':
            downloadAllTimetables();
            break;
        case 'settings':
            showSettingsModal(); // Now opens settings modal directly
            break;
        case 'help':
            showHelpModal(); // Now opens help modal directly
            break;
        case 'feedback':
            showFeedbackModal();
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
    if (loadingOverlay) {
        if (show) {
            loadingOverlay.classList.add('active');
            // Simulate progress
            simulateProgress();
        } else {
            loadingOverlay.classList.remove('active');
        }
    }
}

function simulateProgress() {
    const progressFill = document.getElementById('progress-fill');
    const progressText = document.getElementById('progress-text');
    
    if (progressFill && progressText) {
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
}

function showNotification(message, type = 'info') {
    if (!notificationContainer) return;
    
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
    console.log('Uploaded files:', uploadedFiles);
    console.log('Upload section visible:', isUploadSectionVisible);
    console.log('Course database:', courseDatabase);
    console.log('Filtered timetables:', filterTimetablesData());
    console.log('===================');
}

// Debug file matching
function debugFileMatching() {
    console.log('=== FILE MATCHING DEBUG ===');
    const requiredFiles = [
        'course_data.csv',
        'faculty_availability.csv', 
        'classroom_data.csv',
        'student_data.csv',
        'exams_data.csv'
    ];
    
    requiredFiles.forEach(requiredFile => {
        const requiredFileClean = requiredFile.toLowerCase().replace(/[ _-]/g, '');
        console.log(`Required: ${requiredFile} -> ${requiredFileClean}`);
        
        uploadedFiles.forEach(file => {
            const fileNameClean = file.name.toLowerCase().replace(/[ _-]/g, '');
            const matches = fileNameClean.includes(requiredFileClean) || requiredFileClean.includes(fileNameClean);
            console.log(`  ${file.name} -> ${fileNameClean} : ${matches ? '✅ MATCH' : '❌ NO MATCH'}`);
        });
    });
    console.log('==========================');
}

// Enhanced Settings Functions
function showSettingsModal() {
    const settingsModal = document.getElementById('settings-modal');
    if (settingsModal) {
        settingsModal.style.display = 'flex';
        loadCurrentSettings();
        updateSettingsStats();
    }
}

function closeSettingsModal() {
    const settingsModal = document.getElementById('settings-modal');
    if (settingsModal) {
        settingsModal.style.display = 'none';
    }
}

function loadCurrentSettings() {
    // Load saved settings from localStorage or use defaults
    const defaultView = localStorage.getItem('defaultView') || 'grid';
    const colorTheme = localStorage.getItem('colorTheme') || 'default';
    const notifications = localStorage.getItem('notifications') !== 'false';
    
    document.getElementById('default-view').value = defaultView;
    document.getElementById('color-theme').value = colorTheme;
    document.getElementById('notifications-toggle').checked = notifications;
    
    // Update last updated timestamp
    document.getElementById('last-updated').textContent = new Date().toLocaleString();
}

function updateSettingsStats() {
    // Update statistics in settings modal
    const timetableCount = currentTimetables.length;
    const courseCount = Object.keys(courseDatabase).length;
    
    document.getElementById('settings-timetable-count').textContent = timetableCount;
    document.getElementById('settings-course-count').textContent = courseCount;
}

function saveSettings() {
    const defaultView = document.getElementById('default-view').value;
    const colorTheme = document.getElementById('color-theme').value;
    const notifications = document.getElementById('notifications-toggle').checked;
    const electiveStrategy = document.getElementById('elective-strategy').value;
    
    // Save to localStorage
    localStorage.setItem('defaultView', defaultView);
    localStorage.setItem('colorTheme', colorTheme);
    localStorage.setItem('notifications', notifications);
    localStorage.setItem('electiveStrategy', electiveStrategy);
    
    // Apply settings
    applySettings();
    
    showNotification('✅ Settings saved successfully!', 'success');
    closeSettingsModal();
}

function applySettings() {
    const defaultView = localStorage.getItem('defaultView') || 'grid';
    const colorTheme = localStorage.getItem('colorTheme') || 'default';
    
    // Apply view mode
    if (viewMode) {
        viewMode.value = defaultView;
        currentView = defaultView;
        renderTimetables();
    }
    
    // Apply theme (simplified - you can expand this)
    document.body.setAttribute('data-theme', colorTheme);
}

function exportConfiguration() {
    const settings = {
        defaultView: localStorage.getItem('defaultView') || 'grid',
        colorTheme: localStorage.getItem('colorTheme') || 'default',
        notifications: localStorage.getItem('notifications') !== 'false',
        electiveStrategy: localStorage.getItem('electiveStrategy') || 'common',
        exportDate: new Date().toISOString()
    };
    
    const dataStr = JSON.stringify(settings, null, 2);
    const dataBlob = new Blob([dataStr], {type: 'application/json'});
    
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'timetable-settings.json';
    link.click();
    
    URL.revokeObjectURL(url);
    showNotification('📁 Settings exported successfully!', 'success');
}

function backupData() {
    const backupData = {
        timetables: currentTimetables,
        courseDatabase: courseDatabase,
        settings: {
            defaultView: localStorage.getItem('defaultView'),
            colorTheme: localStorage.getItem('colorTheme'),
            notifications: localStorage.getItem('notifications')
        },
        backupDate: new Date().toISOString()
    };
    
    const dataStr = JSON.stringify(backupData, null, 2);
    const dataBlob = new Blob([dataStr], {type: 'application/json'});
    
    const url = URL.createObjectURL(dataBlob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `timetable-backup-${new Date().toISOString().split('T')[0]}.json`;
    link.click();
    
    URL.revokeObjectURL(url);
    showNotification('💾 Data backup created successfully!', 'success');
}

function resetToDefaults() {
    if (confirm('Are you sure you want to reset all settings to defaults? This action cannot be undone.')) {
        localStorage.clear();
        applySettings();
        showNotification('🔄 Settings reset to defaults!', 'info');
        closeSettingsModal();
    }
}

// Initialize settings when app loads
function initializeSettings() {
    applySettings();
}
// Initialize settings when the app loads
// Add this to your initializeApp() function:
// setupSettings();
// Help & Support Functions
function setupHelpSupport() {
    const helpNavBtn = document.getElementById('help-nav-btn');
    if (helpNavBtn) {
        helpNavBtn.addEventListener('click', function(e) {
            e.preventDefault();
            showHelpModal();
        });
    }

    // Also make the quick action help button work
    const helpActionCard = document.querySelector('.action-card[data-action="help"]');
    if (helpActionCard) {
        helpActionCard.addEventListener('click', function() {
            showHelpModal();
        });
    }
}

function showHelpModal() {
    const helpModal = document.getElementById('help-modal');
    if (helpModal) {
        helpModal.style.display = 'flex';
    }
}

function closeHelpModal() {
    const helpModal = document.getElementById('help-modal');
    if (helpModal) {
        helpModal.style.display = 'none';
    }
}

function showDebugInfo() {
    debugApp();
    verifyDataLoad();
    showNotification('🔧 Debug information logged to console', 'info');
}

function clearCacheAndReload() {
    showNotification('🔄 Clearing cache and reloading...', 'info');
    
    clearCache().then(() => {
        setTimeout(() => {
            loadTimetables();
            loadStats();
            showNotification('✅ Cache cleared and data reloaded', 'success');
        }, 1000);
    });
}

// Add keyboard shortcut for help (F1)
document.addEventListener('keydown', function(e) {
    if (e.key === 'F1') {
        e.preventDefault();
        showHelpModal();
    }
});

// Initialize help system when app loads
// Add this to your initializeApp() function:
// setupHelpSupport();

// Export functions for global access
window.downloadTimetable = downloadTimetable;
window.printTimetable = printTimetable;
window.debugApp = debugApp;
window.showUploadSection = showUploadSection;
window.hideUploadSection = hideUploadSection;
window.debugFileMatching = debugFileMatching;
window.verifyDataLoad = verifyDataLoad;
window.clearCache = clearCache;
window.resetFilters = resetFilters;

