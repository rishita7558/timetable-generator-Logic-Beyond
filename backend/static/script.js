// Global variables
let currentTimetables = [];
let currentView = 'grid';
let currentSemesterFilter = 'all';
let currentSectionFilter = 'all';
let currentBranchFilter = 'all';
let currentTimetableTypeFilter = 'all';
let uploadedFiles = [];
let isUploadSectionVisible = false;
let currentExamTimetables = [];
let isExamSectionVisible = false;
let allExamTimetables = [];
let showAllSchedules = false;

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
const timetableTypeFilter = document.getElementById('timetable-type-filter');
const viewMode = document.getElementById('view-mode');

// Helper Functions
// Add this helper function at the top with other helper functions
function parseDDMMYYYY(dateStr) {
    if (!dateStr) return new Date(0);
    
    const parts = dateStr.split('-');
    if (parts.length !== 3) return new Date(0);
    
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // Months are 0-indexed in JavaScript
    const year = parseInt(parts[2], 10);
    
    return new Date(year, month, day);
}

function sortDates(dates) {
    return dates.sort((a, b) => parseDDMMYYYY(a) - parseDDMMYYYY(b));
}

function formatDateForDisplay(dateStr) {
    // If it's already in dd-mm-yyyy format, return as is
    if (typeof dateStr === 'string' && dateStr.includes('-')) {
        return dateStr;
    }
    // If it's in yyyy-mm-dd format, convert to dd-mm-yyyy
    if (typeof dateStr === 'string' && dateStr.includes('-')) {
        const parts = dateStr.split('-');
        if (parts.length === 3 && parts[0].length === 4) {
            return `${parts[2]}-${parts[1]}-${parts[0]}`;
        }
    }
    return dateStr;
}

function safeGet(object, path, defaultValue = null) {
    return path.split('.').reduce((obj, key) => (obj && obj[key] !== undefined) ? obj[key] : defaultValue, object);
}

function safeArray(array) {
    return Array.isArray(array) ? array : [];
}

function safeString(str) {
    return typeof str === 'string' ? str : 'N/A';
}

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
    themeConfig.init();
    uiConfig.init();
    
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
    
    if (timetableTypeFilter) {
        timetableTypeFilter.addEventListener('change', function() {
            currentTimetableTypeFilter = this.value;
            updateSectionTitle();
            renderTimetables();
        });
    }
    
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
    
    initializeExamSystem();

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
    showLoading(true, {
        title: 'Generating timetables',
        subtitle: 'Building schedules and assigning rooms...',
        iconClass: 'fas fa-bolt fa-spin'
    });
    
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
        
        const data = await response.json();
        
        // Validate and normalize the response data
        if (Array.isArray(data)) {
            // Direct array response
            currentTimetables = data;
        } else if (data && Array.isArray(data.timetables)) {
            // Response with timetables property
            currentTimetables = data.timetables;
        } else if (data && data.data && Array.isArray(data.data)) {
            // Response with data property containing array
            currentTimetables = data.data;
        } else {
            // Unexpected response structure
            console.warn('⚠️ Unexpected timetables response structure:', data);
            currentTimetables = [];
        }
        
        console.log(`📊 Loaded ${currentTimetables.length} timetables`);
        
        // Update course database with server data
        if (currentTimetables.length > 0 && currentTimetables[0].course_info) {
            courseDatabase = currentTimetables[0].course_info;
        }
        
        renderTimetables();
        
    } catch (error) {
        console.error('❌ Error loading timetables:', error);
        showNotification('❌ Error loading timetables: ' + error.message, 'error');
        // Ensure currentTimetables is always an array
        currentTimetables = [];
        renderTimetables();
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
        const usableHintEl = document.getElementById('usable-classrooms-hint');
        
        if (totalTimetablesEl) totalTimetablesEl.textContent = stats.total_timetables;
        if (totalCoursesEl) totalCoursesEl.textContent = stats.total_courses;
        if (totalFacultyEl) totalFacultyEl.textContent = stats.total_faculty;
        if (totalClassroomsEl) totalClassroomsEl.textContent = stats.total_classrooms;
        
        // Display usable classrooms hint
        if (usableHintEl && stats.usable_classrooms) {
            usableHintEl.textContent = `(${stats.usable_classrooms} available for scheduling)`;
        }
        
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
                if (timeText.includes('07:30-09:00') || timeText.includes('18:30-20:00') ||
                    timeText.includes('09:00-10:30') || timeText.includes('10:30-12:00') ||
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

function setupEnhancedTooltips() {
    // Create tooltip element
    let tooltip = document.getElementById('enhanced-tooltip');
    if (!tooltip) {
        tooltip = document.createElement('div');
        tooltip.id = 'enhanced-tooltip';
        tooltip.className = 'enhanced-tooltip';
        document.body.appendChild(tooltip);
    }

    // Add event listeners to all colored cells
    document.addEventListener('mouseover', function(e) {
        const cell = e.target.closest('.colored-cell');
        if (cell && cell.hasAttribute('data-course-code')) {
            showEnhancedTooltip(cell, e.clientX, e.clientY);
        }
    });

    document.addEventListener('mousemove', function(e) {
        const tooltip = document.getElementById('enhanced-tooltip');
        if (tooltip.style.display === 'block') {
            tooltip.style.left = (e.clientX + 15) + 'px';
            tooltip.style.top = (e.clientY + 15) + 'px';
        }
    });

    document.addEventListener('mouseout', function(e) {
        const cell = e.target.closest('.colored-cell');
        if (cell) {
            hideEnhancedTooltip();
        }
    });
}
function setupExamTooltips() {
    const examCards = document.querySelectorAll('.exam-card-time');
    
    examCards.forEach((card) => {
        const tooltip = card.querySelector('.exam-tooltip-enhanced');
        
        if (tooltip) {
            card.addEventListener('mouseenter', function(e) {
                const rect = this.getBoundingClientRect();
                const tooltipRect = tooltip.getBoundingClientRect();
                const viewportWidth = window.innerWidth;
                const viewportHeight = window.innerHeight;
                
                // Reset all positioning
                tooltip.style.left = '50%';
                tooltip.style.right = 'auto';
                tooltip.style.bottom = 'calc(100% + 10px)';
                tooltip.style.top = 'auto';
                tooltip.classList.remove('position-below');
                
                // Calculate ideal position
                let idealLeft = rect.left + (rect.width / 2);
                let leftPosition = idealLeft - (tooltipRect.width / 2);
                
                // Adjust for left boundary
                if (leftPosition < 10) {
                    leftPosition = 10;
                }
                // Adjust for right boundary
                else if (leftPosition + tooltipRect.width > viewportWidth - 10) {
                    leftPosition = viewportWidth - tooltipRect.width - 10;
                }
                
                // Check if tooltip fits above the card
                const spaceAbove = rect.top - 20;
                const spaceBelow = viewportHeight - rect.bottom - 20;
                
                if (spaceAbove >= tooltipRect.height || spaceAbove >= spaceBelow) {
                    // Position above
                    tooltip.style.bottom = 'calc(100% + 10px)';
                    tooltip.style.top = 'auto';
                } else {
                    // Position below
                    tooltip.style.bottom = 'auto';
                    tooltip.style.top = 'calc(100% + 10px)';
                    tooltip.classList.add('position-below');
                }
                
                // Apply horizontal positioning
                tooltip.style.left = `${leftPosition}px`;
                tooltip.style.transform = 'translateX(0)';
                
                // Show tooltip
                tooltip.style.opacity = '1';
                tooltip.style.visibility = 'visible';
            });
            
            card.addEventListener('mouseleave', function() {
                tooltip.style.opacity = '0';
                tooltip.style.visibility = 'hidden';
            });
            
            // Click handler is already in the HTML onclick attribute
        }
    });
}

function showEnhancedTooltip(cell, x, y) {
    const tooltip = document.getElementById('enhanced-tooltip');
    if (!tooltip) return;

    const courseCode = cell.getAttribute('data-course-code');
    const courseName = cell.getAttribute('data-course-name');
    const sessionType = cell.getAttribute('data-session-type');
    const credits = cell.getAttribute('data-credits');
    const instructor = cell.getAttribute('data-instructor');
    const department = cell.getAttribute('data-department');
    const courseType = cell.getAttribute('data-course-type');
    const isCommon = cell.getAttribute('data-is-common') === 'true';

    // Determine session icon and color
    let sessionIcon = '📚'; // Default lecture icon
    let sessionColor = '#4361ee'; // Default blue
    
    if (sessionType.includes('Tutorial')) {
        sessionIcon = '👨‍🏫';
        sessionColor = '#4cc9f0'; // Cyan for tutorials
    } else if (sessionType.includes('Elective')) {
        sessionIcon = '⭐';
        sessionColor = '#7209b7'; // Purple for electives
    }

    const isElective = courseType === 'Elective' || sessionType.includes('Elective');

    tooltip.innerHTML = `
        <div class="tooltip-header" style="border-left: 4px solid ${sessionColor}">
            <div class="tooltip-title">
                <span class="session-icon">${sessionIcon}</span>
                <div>
                    <div class="course-code">${courseCode}</div>
                    <div class="session-type">${sessionType}</div>
                </div>
            </div>
            ${isCommon ? '<div class="common-badge">COMMON SLOT</div>' : ''}
        </div>
        <div class="tooltip-body">
            <div class="course-name">${courseName}</div>
            <div class="tooltip-details">
                <div class="detail-item">
                    <i class="fas fa-user-graduate"></i>
                    <span>${instructor}</span>
                </div>
                <div class="detail-item">
                    <i class="fas fa-university"></i>
                    <span>${department}</span>
                </div>
                <div class="detail-item">
                    <i class="fas fa-star"></i>
                    <span>${credits} Credits • ${courseType}</span>
                </div>
                ${isElective ? `
                <div class="detail-item highlight">
                    <i class="fas fa-users"></i>
                    <span>Common for all branches & sections</span>
                </div>
                ` : ''}
            </div>
        </div>
        <div class="tooltip-footer">
            <small>Click for more details</small>
        </div>
    `;

    tooltip.style.left = (x + 15) + 'px';
    tooltip.style.top = (y + 15) + 'px';
    tooltip.style.display = 'block';
    tooltip.style.opacity = '1';
}

function hideEnhancedTooltip() {
    const tooltip = document.getElementById('enhanced-tooltip');
    if (tooltip) {
        tooltip.style.opacity = '0';
        setTimeout(() => {
            tooltip.style.display = 'none';
        }, 200);
    }
}

// Color Coding and Legend Functions
// Color Coding and Legend Functions
function applyDynamicColorCoding(tableElement, courseColors, basketColors = {}) {
    const cells = tableElement.querySelectorAll('td');
    
    cells.forEach(cell => {
        const text = cell.textContent.trim();
        
        // Skip header cells, empty cells, and special slots
        if (!text || text === 'Free' || text === 'LUNCH BREAK' || cell.cellIndex === 0) {
            // Clear the content for free slots but keep the cell structure
            if (text === 'Free') {
                cell.textContent = ''; // Clear the "Free" text
                cell.classList.add('empty-slot'); // Add class for empty slots
            } else if (text === 'LUNCH BREAK') {
                cell.classList.add('lunch-break-slot'); // Keep lunch break styling
            }
            return;
        }
        
        // Check if this is a basket entry (ELECTIVE_B1, HSS_B3, etc.)
        const isBasket = isBasketEntry(text);
        const isBasketTutorial = text.includes('(Tutorial)') && isBasketEntry(text.replace(' (Tutorial)', ''));
        
        if (isBasket || isBasketTutorial) {
            // Handle basket entries
            const basketName = isBasketTutorial ? text.replace(' (Tutorial)', '') : text;
            const color = basketColors[basketName] || courseColors[basketName] || getDefaultBasketColor(basketName);
            
            if (color) {
                cell.style.background = color;
                cell.style.color = getContrastColor(color);
                cell.style.fontWeight = '600';
                cell.style.border = '2px solid white';
                
                // Add basket tooltip
                cell.setAttribute('data-basket-name', basketName);
                cell.setAttribute('data-session-type', isBasketTutorial ? 'Basket Tutorial' : 'Basket Lecture');
                cell.setAttribute('data-is-basket', 'true');
                
                cell.style.cursor = 'pointer';
                cell.classList.add('basket-slot');
                
                if (isBasketTutorial) {
                    cell.classList.add('basket-tutorial-session');
                } else {
                    cell.classList.add('basket-lecture-session');
                }
            }
        } else {
            // Handle regular course entries
            // Extract course code - handle both regular and tutorial versions
            let courseCode = extractCourseCode(text);
            
            // For tutorials, use the base course code without "(Tutorial)"
            if (text.includes('(Tutorial)') && courseCode) {
                const baseCourseCode = text.replace(' (Tutorial)', '');
                courseCode = extractCourseCode(baseCourseCode) || courseCode;
                cell.classList.add('tutorial-session');
            }
            
            if (courseCode && courseColors[courseCode]) {
                // Apply dynamic color - SAME COLOR for all sessions of the course
                const color = courseColors[courseCode];
                cell.style.background = color;
                cell.style.color = getContrastColor(color);
                cell.style.fontWeight = '600';
                cell.style.border = '2px solid white';
                
                // Add enhanced tooltip with course info
                const courseInfo = courseDatabase[courseCode];
                if (courseInfo) {
                    let sessionType = 'Lecture';
                    if (text.includes('(Tutorial)')) sessionType = 'Tutorial';
                    
                    // Enhanced tooltip data attributes
                    cell.setAttribute('data-course-code', courseCode);
                    cell.setAttribute('data-course-name', courseInfo.name);
                    cell.setAttribute('data-session-type', sessionType);
                    cell.setAttribute('data-credits', courseInfo.credits);
                    cell.setAttribute('data-instructor', courseInfo.instructor);
                    cell.setAttribute('data-department', courseInfo.department);
                    cell.setAttribute('data-course-type', courseInfo.type);
                }
                
                // Make cell clickable for more info
                cell.style.cursor = 'pointer';
                cell.classList.add('colored-cell');
            }
        }
    });
}

function isBasketEntry(text) {
    // Check if the text matches basket naming patterns
    const basketPatterns = [
        /^ELECTIVE_B\d+$/,
        /^HSS_B\d+$/,
        /^PROF_B\d+$/,
        /^OE_B\d+$/,
        /^BASKET_\w+$/,
        /^B_\d+$/
    ];
    
    return basketPatterns.some(pattern => pattern.test(text));
}

function getDefaultBasketColor(basketName) {
    // Generate consistent colors for baskets
    const basketColors = {
        'ELECTIVE_B1': '#FF6B6B',
        'ELECTIVE_B2': '#4ECDC4', 
        'ELECTIVE_B3': '#45B7D1',
        'ELECTIVE_B4': '#96CEB4',
        'HSS_B1': '#FECA57',
        'HSS_B2': '#FF9FF3',
        'HSS_B3': '#54A0FF',
        'HSS_B4': '#5F27CD',
        'PROF_B1': '#00D2D3',
        'PROF_B2': '#FF9F43',
        'OE_B1': '#10AC84',
        'OE_B2': '#EE5A24'
    };
    
    // If basket not in predefined colors, generate a consistent color
    if (basketColors[basketName]) {
        return basketColors[basketName];
    }
    
    // Generate color based on basket name hash
    let hash = 0;
    for (let i = 0; i < basketName.length; i++) {
        hash = basketName.charCodeAt(i) + ((hash << 5) - hash);
    }
    
    const hue = hash % 360;
    return `hsl(${hue}, 70%, 65%)`;
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
    // Handle both regular courses and tutorial marked courses
    const cleanText = text.replace(' (Tutorial)', '');
    
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
    console.log('Baskets:', timetable.baskets);
    console.log('Basket courses map:', timetable.basket_courses_map);
    
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

function createEnhancedLegend(semester, section, courses, courseColors, courseInfo, coreCourses = [], electiveCourses = [], baskets = [], basketCoursesMap = {}, basketColors = {}, timetable = null) {
    if (!courses || courses.length === 0) return '';
    
    // Get unique courses with aggressive deduplication
    const uniqueCourses = getUniqueCourses(courses);
    
    // Separate courses into core and elective for better organization
    const uniqueCoreCourses = getUniqueCourses(coreCourses);
    const uniqueElectiveCourses = getUniqueCourses(electiveCourses);
    
    // Determine department filter from timetable branch
    const branchAbbrev = timetable?.branch || null;
    const mapBranchToDepartment = (abbr) => {
        if (!abbr) return null;
        const key = String(abbr).trim().toUpperCase();
        if (key === 'CSE' || key === 'CS') return 'Computer Science and Engineering';
        if (key === 'ECE' || key === 'EC') return 'Electronics and Communication Engineering';
        if (key === 'DSAI' || key === 'DS' || key === 'DA') return 'Data Science and Artificial Intelligence';
        return null;
    };
    const allowedDepartment = mapBranchToDepartment(branchAbbrev);

    // Infer department from course code as a robust fallback
    const inferDepartmentFromCode = (code) => {
        if (!code || typeof code !== 'string') return null;
        const upper = code.trim().toUpperCase();
        if (upper.startsWith('CS')) return 'Computer Science and Engineering';
        if (upper.startsWith('EC')) return 'Electronics and Communication Engineering';
        if (upper.startsWith('DS') || upper.startsWith('DA')) return 'Data Science and Artificial Intelligence';
        return null;
    };

    // Filter to only include courses that actually exist in the timetable
    // Prefer backend-provided scheduled core courses if available
    let finalCoreCourses = Array.isArray(timetable?.scheduled_core_courses) && timetable.scheduled_core_courses.length > 0
        ? getUniqueCourses(timetable.scheduled_core_courses)
        : [];

    // Fallback: derive from timetable courses excluding basket courses
    if (finalCoreCourses.length === 0) {
        const gatherBasketCourses = (mapObj) => {
            if (!mapObj) return [];
            const out = [];
            Object.values(mapObj).forEach(list => {
                if (Array.isArray(list)) out.push(...list);
            });
            return out;
        };
        const basketCoursesFromMap = gatherBasketCourses(basketCoursesMap);
        const basketCoursesFromAll = gatherBasketCourses(timetable?.all_basket_courses);
        const allElectiveCodes = new Set([
            ...basketCoursesFromMap,
            ...basketCoursesFromAll
        ]);
        finalCoreCourses = uniqueCourses.filter(code => !allElectiveCodes.has(code));
    }

    // Include all scheduled non-basket courses irrespective of department
    const electiveCourseList = uniqueElectiveCourses.filter(course => uniqueCourses.includes(course));
    // We no longer show "Other Courses" in the legend; keep empty
    const otherCourses = [];
    
    console.log(`📊 Legend for Semester ${semester}, Section ${section}:`, {
        allCourses: uniqueCourses,
        coreCourses: finalCoreCourses,
        electiveCourses: electiveCourseList,
        otherCourses: otherCourses,
        baskets: baskets,
        basketCoursesMap: basketCoursesMap,
        allBasketCourses: timetable?.all_basket_courses
    });
    
    // If no courses to display, return empty
    if (finalCoreCourses.length === 0 && electiveCourseList.length === 0 && baskets.length === 0) {
        console.log('⚠️ No courses to display in legend');
        return '';
    }
    
    let legendHtml = `
        <div class="timetable-legend">
            <div class="legend-title">
                <i class="fas fa-palette"></i>
                Scheduled Courses - Semester ${semester}, Section ${section}
            </div>
    `;
    
    // Core Courses Section — always render the section; show a friendly message if empty
    legendHtml += `
        <div class="legend-section">
            <div class="legend-section-title">
                <i class="fas fa-book"></i>
                Scheduled Core Courses (${finalCoreCourses.length})
            </div>
            <div class="legend-grid">
    `;

    if (finalCoreCourses.length > 0) {
        finalCoreCourses.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color, 'core');
        });
    } else {
        legendHtml += `
            <div class="legend-empty">
                <i class="fas fa-info-circle"></i>
                <span>No core courses detected for this timetable.</span>
            </div>
        `;
    }

    legendHtml += `
            </div>
        </div>
    `;
    
    // Elective Baskets Section - Filter based on semester
    if (baskets.length > 0) {
        // Filter baskets based on semester
        let filteredBaskets = baskets;
        
        // Enforce allowed elective baskets per semester
        const allowedBasketsBySemester = {
            1: ['ELECTIVE_B1'],
            3: ['ELECTIVE_B3'],
            5: ['ELECTIVE_B4', 'ELECTIVE_B5'],
            7: ['ELECTIVE_B6', 'ELECTIVE_B7', 'ELECTIVE_B8', 'ELECTIVE_B9']
        };

        if (allowedBasketsBySemester[semester]) {
            // Show ALL allowed baskets for this semester (even if not present in the schedule)
            filteredBaskets = allowedBasketsBySemester[semester].slice();
            console.log(`🎯 Semester ${semester} - Showing allowed baskets:`, filteredBaskets);
        } // For other semesters, show all baskets (no filtering)

        
        if (filteredBaskets.length > 0) {
            legendHtml += `
                <div class="legend-section elective-baskets">
                    <div class="legend-section-header elective">
                        <div class="legend-section-title elective">
                            <i class="fas fa-clipboard-list"></i>
                            Elective Baskets (${filteredBaskets.length})
                        </div>
                        <div class="elective-semester-info">
                            <i class="fas fa-info-circle"></i>
                            <span>Common for All Branches & Sections</span>
                        </div>
                    </div>
                    <div class="elective-baskets-container">
            `;
            
            filteredBaskets.sort().forEach(basketName => {
                // Combine courses from both basketCoursesMap and all_basket_courses
                const basketCoursesFromMap = basketCoursesMap[basketName] || [];
                const basketCoursesFromAll = (timetable && timetable.all_basket_courses && timetable.all_basket_courses[basketName]) || [];
                
                // Merge and deduplicate courses
                const allBasketCourses = [...new Set([...basketCoursesFromMap, ...basketCoursesFromAll])];
                
                const color = basketColors[basketName] || courseColors[basketName] || getDefaultBasketColor(basketName);
                
                // Determine basket type for styling
                let basketType = 'general';
                if (basketName.includes('ELECTIVE')) basketType = 'professional';
                else if (basketName.includes('HSS')) basketType = 'hss';
                else if (basketName.includes('PROF')) basketType = 'professional';
                else if (basketName.includes('OE')) basketType = 'open';
                
                legendHtml += `
                    <div class="legend-basket ${basketType}-basket">
                        <div class="basket-header" style="border-left-color: ${color}">
                            <div class="basket-header-content">
                                <div class="basket-color-info">
                                    <div class="legend-color" style="background: ${color};"></div>
                                    <span class="basket-name"><strong>${basketName}</strong></span>
                                </div>
                                <div class="basket-stats">
                                    <span class="course-count">${allBasketCourses.length} courses</span>
                                </div>
                            </div>
                        </div>
                        <div class="basket-courses">
                `;
                
                if (allBasketCourses.length > 0) {
                    allBasketCourses.forEach(courseCode => {
                        const info = courseInfo[courseCode] || { name: courseCode };
                        // Pass in basket-specific allocations map and basketName so course items can show allocated rooms
                        legendHtml += createBasketCourseItem(courseCode, info, color, timetable && timetable.basket_course_allocations, basketName);
                    });
                } else {
                    legendHtml += `
                        <div class="no-courses">
                            <i class="fas fa-exclamation-circle"></i>
                            <span>No courses available in this basket</span>
                        </div>
                    `;
                }
                
                legendHtml += `
                        </div>
                        <div class="basket-footer">
                            <div class="basket-schedule-info">
                                <i class="fas fa-clock"></i>
                                <span>2 Lectures + 1 Tutorial per week (Different rooms)</span>
                            </div>
                            <div class="basket-note">
                                <i class="fas fa-info-circle"></i>
                                <span>📌 Each elective course meets in different rooms for each session type to efficiently manage classroom capacity</span>
                            </div>
                        </div>
                    </div>
                `;
            });
            
            // Add semester-specific guidance
            let semesterGuidance = '';
            if (semester === 3) {
                semesterGuidance = `
                    <div class="semester-guidance">
                        <div class="guidance-header">
                            <i class="fas fa-graduation-cap"></i>
                            <strong>Semester 3 Elective Guidance</strong>
                        </div>
                        <div class="guidance-content">
                            <p>Choose one elective basket from the available B3 options above.</p>
                            <ul>
                                <li>Each basket contains related elective courses</li>
                                <li>Courses are common across all branches and sections</li>
                                <li>Schedule includes 2 lectures and 1 tutorial per week</li>
                            </ul>
                        </div>
                    </div>
                `;
            } else if (semester === 5) {
                semesterGuidance = `
                    <div class="semester-guidance">
                        <div class="guidance-header">
                            <i class="fas fa-graduation-cap"></i>
                            <strong>Semester 5 Elective Guidance</strong>
                        </div>
                        <div class="guidance-content">
                            <p>Choose one elective basket from the available B5 options above.</p>
                            <ul>
                                <li>Each basket contains specialized elective courses</li>
                                <li>Courses are common across all branches and sections</li>
                                <li>Schedule includes 2 lectures and 1 tutorial per week</li>
                                <li>Focus on advanced topics in your chosen specialization</li>
                            </ul>
                        </div>
                    </div>
                `;
            }
            
            legendHtml += `
                    </div>
                    ${semesterGuidance}
                </div>
            `;
        }
    }
    
    // Individual Elective Courses Section (if any exist outside baskets)
    if (electiveCourseList.length > 0) {
        legendHtml += `
            <div class="legend-section elective">
                <div class="legend-section-title elective">
                    <i class="fas fa-star"></i>
                    Individual Elective Courses (${electiveCourseList.length})
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
            </div>
        `;
    }
    
    // Removed "Other Courses" section by design
    
    legendHtml += `</div>`;
    return legendHtml;
}

function createLegendItem(courseCode, courseInfo, color, type = 'core') {
    const courseName = courseInfo ? courseInfo.name : 'Unknown Course';
    const credits = courseInfo ? courseInfo.credits : '?';
    const instructor = courseInfo ? courseInfo.instructor : 'Unknown';
    const courseType = courseInfo ? courseInfo.type : 'Core';
    const ltpsc = courseInfo && courseInfo.ltpsc ? courseInfo.ltpsc : '';
    
    // Get term type (Pre-Mid, Post-Mid, or Full Sem)
    const termType = courseInfo && courseInfo.term_type ? courseInfo.term_type : 'Full Sem';
    
    // Add type-specific icons
    let typeIcon = 'fas fa-book'; // Default for core
    if (type === 'elective') typeIcon = 'fas fa-star';
    else if (type === 'other') typeIcon = 'fas fa-graduation-cap';
    
    return `
        <div class="legend-item ${type}">
            <div class="legend-color" style="background: ${color};"></div>
            <div class="legend-item-content">
                <div class="legend-course-header">
                    <span class="legend-course-code">${courseCode}</span>
                    <span class="course-credits">${credits} cr</span>
                </div>
                <span class="legend-course-name">${courseName}</span>
                <div class="legend-course-details">
                    <i class="${typeIcon}"></i>
                    <span>${instructor} • ${courseType}${ltpsc ? ` • ${ltpsc}` : ''} • ${termType}</span>
                </div>
            </div>
        </div>
    `;
}

function createBasketCourseItem(courseCode, courseInfo, basketColor, allocationsMap = {}, basketName = null) {
    const courseName = courseInfo ? courseInfo.name : 'Unknown Course';
    const credits = courseInfo ? courseInfo.credits : '?';
    const instructor = courseInfo ? courseInfo.instructor : 'Unknown';
    const ltpsc = courseInfo && courseInfo.ltpsc ? courseInfo.ltpsc : '';
    
    // Get term type (Pre-Mid, Post-Mid, or Full Sem)
    const termType = courseInfo && courseInfo.term_type ? courseInfo.term_type : 'Full Sem';

    // Build a compact single-line meta
    const metaLine = [
        courseCode,
        `${credits} cr`,
        courseName,
        instructor,
        ltpsc ? `${ltpsc}` : null,
        termType
    ].filter(Boolean).join(' • ');
    
    // Determine allocated schedule for this course (if provided)
    let roomDisplay = '';
    try {
        if (allocationsMap && basketName && allocationsMap[basketName]) {
            const alloc = allocationsMap[basketName][courseCode];
            if (alloc) {
                // Check if it's the new format with day/time info
                if (Array.isArray(alloc) && alloc.length > 0 && typeof alloc[0] === 'object' && alloc[0].room) {
                    // New format: [{room: 'C101', day: 'Monday', time: '09:00-10:30'}, ...]
                    
                    // Group allocations by room+time+type (avoid merging tutorials with lectures)
                    const grouped = {};
                    alloc.forEach(item => {
                        const sessionType = item.type || '';
                        const key = `${item.room}|${item.time}|${sessionType}`;
                        if (!grouped[key]) {
                            grouped[key] = {
                                room: item.room,
                                time: item.time,
                                type: sessionType,
                                days: []
                            };
                        }
                        grouped[key].days.push(item.day);
                    });
                    
                    const groupedList = Object.values(grouped);

                    // Render a simple schedule list (one line per session)
                    const scheduleLines = groupedList.map(item => {
                        const daysText = item.days.join(' & ');
                        return `<div class="basket-course-schedule-line">
                                    <span class="room-chip">${item.room}</span>
                                    <span class="schedule-text">${daysText} ${item.time}${item.type ? ` • ${item.type}` : ''}</span>
                                </div>`;
                    }).join('');

                    roomDisplay = `
                        <div class="basket-course-room simple-schedule">
                            <div class="room-label">
                                <i class="fas fa-door-open"></i>
                                Class Schedule
                            </div>
                            <div class="basket-course-schedule">${scheduleLines}</div>
                        </div>
                    `;
                } else if (Array.isArray(alloc) && alloc.length > 1) {
                    // Old format: Multiple rooms (array of strings)
                    const roomList = alloc.map(room => `<div class="basket-course-schedule-line"><span class="room-chip">${room}</span></div>`).join('');
                    roomDisplay = `
                        <div class="basket-course-room simple-schedule">
                            <div class="room-label">
                                <i class="fas fa-door-open"></i>
                                Class Rooms
                            </div>
                            <div class="basket-course-schedule">${roomList}</div>
                        </div>
                    `;
                } else if (Array.isArray(alloc)) {
                    // Old format: Single room in array
                    roomDisplay = `
                        <div class="basket-course-room simple-schedule">
                            <div class="room-label">
                                <i class="fas fa-door-open"></i>
                                Room
                            </div>
                            <div class="basket-course-schedule">
                                <div class="basket-course-schedule-line">
                                    <span class="room-chip">${alloc[0]}</span>
                                </div>
                            </div>
                        </div>
                    `;
                } else {
                    // Old format: Single room string
                    roomDisplay = `
                        <div class="basket-course-room simple-schedule">
                            <div class="room-label">
                                <i class="fas fa-door-open"></i>
                                Room
                            </div>
                            <div class="basket-course-schedule">
                                <div class="basket-course-schedule-line">
                                    <span class="room-chip">${alloc}</span>
                                </div>
                            </div>
                        </div>
                    `;
                }
            } else {
                roomDisplay = `<div class="basket-course-room unallocated"><i class="fas fa-exclamation-triangle"></i> Room not yet allocated</div>`;
            }
        }
    } catch (e) {
        roomDisplay = '';
    }
    
    return `
        <div class="basket-course-item" style="border-left-color: ${basketColor}">
            <div class="basket-course-meta">${metaLine}</div>
            ${roomDisplay}
        </div>
    `;
}

function enhanceTables() {
    console.log("🎨 Enhancing tables with color coding and time slots...");
    
    document.querySelectorAll('.timetable-table').forEach(table => {
        // Enhanced "Unnamed: 0" removal - more comprehensive approach
        const allCells = table.querySelectorAll('th, td');
        
        allCells.forEach(cell => {
            const text = cell.textContent.trim();
            // Remove any cell containing "Unnamed:" or "Unnamed: 0"
            if (text.includes('Unnamed:') || text === 'Unnamed: 0' || text === 'Unnamed:0') {
                cell.textContent = ''; // Make it completely empty
                cell.classList.add('empty-header');
                
                // Add minimal styling to ensure it looks clean
                cell.style.padding = '2px';
                cell.style.minWidth = '10px';
            }
        });

        // Additional cleanup for the first column (time slots)
        const rows = table.querySelectorAll('tr');
        rows.forEach(row => {
            const firstCell = row.querySelector('th:first-child, td:first-child');
            if (firstCell && firstCell.textContent.trim() === '') {
                firstCell.classList.add('time-slot-header');
                // Ensure proper styling for time slot headers
                firstCell.style.fontWeight = 'bold';
                firstCell.style.textAlign = 'right';
                firstCell.style.paddingRight = '12px';
            }
        });

        // Find the parent timetable card and get course information
        const timetableCard = table.closest('.timetable-card') || table.closest('.timetable-item');
        if (timetableCard) {
            const header = timetableCard.querySelector('.timetable-header h3');
            if (header) {
                    // Extract semester and section from the header. Section can appear as 'Section A' or 'Whole Branch'
                    const semMatch = header.textContent.match(/Semester\s+(\d+)/i);
                    const sectionABMatch = header.textContent.match(/Section\s+(A|B)/i);
                    const wholeMatch = header.textContent.match(/Whole\s+Branch/i);
                    if (semMatch) {
                    const semester = parseInt(semMatch[1]);
                    let section = null;
                    if (sectionABMatch) section = sectionABMatch[1];
                    else if (wholeMatch) section = 'Whole';
                    
                    // Find the timetable data - include branch via table id to avoid cross-branch mixups
                    const tableId = table.getAttribute('id');
                    const timetable = currentTimetables.find(t => {
                        const normalizedSectionId = (section.toLowerCase && section.toLowerCase() === 'whole') ? 'whole' : section;
                        const expectedId = t.branch ? `sem${t.semester}_${t.branch}_${normalizedSectionId}` : `sem${t.semester}_${normalizedSectionId}`;
                        return expectedId === tableId;
                    }) || currentTimetables.find(t => t.semester === semester && String(t.section || '').toLowerCase() === String(section || '').toLowerCase()) || currentTimetables.find(t => {
                        // Extra fallback: normalize timetable table id and compare to tableId directly
                        const normalized = (t.branch ? `sem${t.semester}_${t.branch}_${String(t.section || '').toLowerCase()}` : `sem${t.semester}_${String(t.section || '').toLowerCase()}`);
                        return normalized === (tableId || '').toLowerCase();
                    });
                    
                    if (timetable) {
                        // Debug duplicates
                        debugCourseDuplicates(timetable);
                        
                        if (timetable.course_colors) {
                            // Apply dynamic color coding
                            applyDynamicColorCoding(table, timetable.course_colors);
                            
                            // Create enhanced legend with course type separation
                            if (timetable.courses) {
                                console.log(`📋 Found ${timetable.courses.length} courses for Semester ${semester} Section ${section} (Branch: ${timetable.branch})`);
                                const legend = createEnhancedLegend(
                                    semester, 
                                    section, 
                                    timetable.courses, 
                                    timetable.course_colors, 
                                    timetable.course_info,
                                    timetable.core_courses || [],
                                    timetable.elective_courses || [],
                                    timetable.baskets || [],
                                    timetable.basket_courses_map || {},
                                    timetable.basket_colors || {},
                                    timetable // Pass the entire timetable object
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
    setupEnhancedTooltips();
    
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

    const allowedViews = ['grid', 'list'];
    if (!allowedViews.includes(currentView)) {
        console.warn(`⚠️ Invalid view "${currentView}" detected, falling back to grid view`);
        currentView = 'grid';
        if (viewMode) viewMode.value = 'grid';
    }
    
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
    }
    
    // Enhance tables with color coding and legends
    enhanceTables();
}

function resetFilters() {
    currentBranchFilter = 'all';
    currentSemesterFilter = 'all';
    currentSectionFilter = 'all';
    currentTimetableTypeFilter = 'all';
    
    if (branchFilter) branchFilter.value = 'all';
    if (semesterFilter) semesterFilter.value = 'all';
    if (sectionFilter) sectionFilter.value = 'all';
    if (timetableTypeFilter) timetableTypeFilter.value = 'all';
    
    updateSectionTitle();
    renderTimetables();
}

function renderGridView(timetables) {
    let html = '<div class="timetables-grid">';
    
    timetables.forEach(timetable => {
        const branchClass = `branch-${timetable.branch?.toLowerCase() || 'general'}`;
        const branchBadge = timetable.branch ? `<span class="branch-badge ${timetable.branch.toLowerCase()}">${timetable.branch}</span>` : '';
        
        // Add timetable type badge
        let typeBadge = '';
        if (timetable.is_pre_mid_timetable || timetable.timetable_type === 'pre_mid') {
            typeBadge = '<span class="timetable-type-badge pre-mid">PRE-MID</span>';
        } else if (timetable.is_post_mid_timetable || timetable.timetable_type === 'post_mid') {
            typeBadge = '<span class="timetable-type-badge post-mid">POST-MID</span>';
        } else if (timetable.is_basket_timetable || timetable.timetable_type === 'basket') {
            typeBadge = '<span class="timetable-type-badge basket">REGULAR</span>';
        }
        
        // Determine display label for section
        const sectionLabelDisplay = (timetable.section === 'A' || timetable.section === 'B') ? ` - Section ${timetable.section}` : (timetable.section === 'Whole' ? ' - Whole Branch' : '');
        html += `
            <div class="timetable-card ${branchClass}">
                <div class="timetable-header">
                    <h3>Semester ${timetable.semester}${sectionLabelDisplay} ${branchBadge} ${typeBadge}</h3>
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
        // Determine display label for section (List view)
        const sectionLabelDisplay = (timetable.section === 'A' || timetable.section === 'B') ? ` - Section ${timetable.section}` : (timetable.section === 'Whole' ? ' - Whole Branch' : '');
        
        // Add timetable type badge
        let typeBadge = '';
        if (timetable.is_pre_mid_timetable || timetable.timetable_type === 'pre_mid') {
            typeBadge = '<span class="timetable-type-badge pre-mid">PRE-MID</span>';
        } else if (timetable.is_post_mid_timetable || timetable.timetable_type === 'post_mid') {
            typeBadge = '<span class="timetable-type-badge post-mid">POST-MID</span>';
        } else if (timetable.is_basket_timetable || timetable.timetable_type === 'basket') {
            typeBadge = '<span class="timetable-type-badge basket">REGULAR</span>';
        }
        
        html += `
            <div class="timetable-item" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); margin-bottom: 1.5rem; overflow: hidden;">
                <div class="timetable-header" style="background: linear-gradient(135deg, var(--primary), var(--primary-dark)); color: white; padding: 1.25rem; display: flex; justify-content: space-between; align-items: center;">
                    <h3 style="margin: 0; font-size: 1.1rem; font-weight: 600;">Semester ${timetable.semester}${sectionLabelDisplay} ${branchBadge} ${typeBadge}</h3>
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
        // Determine display label for section (Compact view)
        const sectionLabelDisplay = (timetable.section === 'A' || timetable.section === 'B') ? ` - Section ${timetable.section}` : (timetable.section === 'Whole' ? ' - Whole Branch' : '');
        
        // Add timetable type badge
        let typeBadge = '';
        if (timetable.is_pre_mid_timetable || timetable.timetable_type === 'pre_mid') {
            typeBadge = '<span class="timetable-type-badge pre-mid">PRE-MID</span>';
        } else if (timetable.is_post_mid_timetable || timetable.timetable_type === 'post_mid') {
            typeBadge = '<span class="timetable-type-badge post-mid">POST-MID</span>';
        } else if (timetable.is_basket_timetable || timetable.timetable_type === 'basket') {
            typeBadge = '<span class="timetable-type-badge basket">REGULAR</span>';
        }
        
        html += `
            <div class="compact-card" style="background: white; border-radius: var(--radius); box-shadow: var(--shadow); padding: 1.5rem; margin-bottom: 1rem;">
                <div style="display: flex; justify-content: between; align-items: center; margin-bottom: 1rem;">
                    <h4 style="margin: 0; color: var(--dark);">Semester ${timetable.semester}${sectionLabelDisplay} ${branchBadge} ${typeBadge}</h4>
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
    if (!Array.isArray(currentTimetables)) {
        console.error('❌ currentTimetables is not an array:', currentTimetables);
        currentTimetables = [];
        return [];
    }
    
    return currentTimetables.filter(timetable => {
        const branchMatch = currentBranchFilter === 'all' || timetable.branch === currentBranchFilter;
        const semesterMatch = currentSemesterFilter === 'all' || timetable.semester === parseInt(currentSemesterFilter);
        const sectionMatch = currentSectionFilter === 'all' || timetable.section === currentSectionFilter;
        
        // Determine timetable type
        let timetableType = 'regular';
        if (timetable.is_pre_mid_timetable || timetable.timetable_type === 'pre_mid') {
            timetableType = 'pre_mid';
        } else if (timetable.is_post_mid_timetable || timetable.timetable_type === 'post_mid') {
            timetableType = 'post_mid';
        }
        
        const typeMatch = currentTimetableTypeFilter === 'all' || 
                         timetableType === currentTimetableTypeFilter ||
                         (currentTimetableTypeFilter === 'regular' && timetable.timetable_type === 'basket');
        
        return branchMatch && semesterMatch && sectionMatch && typeMatch;
    });
}

function changeViewMode() {
    if (viewMode) {
        const selectedView = viewMode.value;
        currentView = ['grid', 'list'].includes(selectedView) ? selectedView : 'grid';
        if (currentView !== selectedView) {
            viewMode.value = currentView;
        }
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
    const sectionLabel = (section === 'A' || section === 'B') ? `Section ${section}` : (section === 'Whole' ? 'Whole Branch' : '');
    showNotification(`🖨️ Printing Semester ${semester} ${sectionLabel}`, 'info');
    
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
                    <h1>IIIT Dharwad - Semester ${semester}${sectionLabel ? ' - ' + sectionLabel : ''}</h1>
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
function showLoading(show, options = {}) {
    if (!loadingOverlay) return;
    const titleEl = document.getElementById('loading-title');
    const subtitleEl = document.getElementById('loading-subtitle');
    const iconEl = document.getElementById('loading-icon');

    const defaultTitle = 'Generating Timetables';
    const defaultSubtitle = 'This may take a few moments...';
    const defaultIcon = 'fas fa-spinner fa-spin';

    if (show) {
        if (titleEl && options.title) titleEl.textContent = options.title;
        if (subtitleEl && options.subtitle) subtitleEl.textContent = options.subtitle;
        if (iconEl) iconEl.className = options.iconClass || defaultIcon;

        const progressFill = document.getElementById('progress-fill');
        const progressText = document.getElementById('progress-text');
        if (progressFill && progressText) {
            progressFill.style.width = '0%';
            progressText.textContent = '0%';
        }

        loadingOverlay.classList.add('active');
        // Simulate progress
        simulateProgress();
    } else {
        loadingOverlay.classList.remove('active');
        if (titleEl) titleEl.textContent = defaultTitle;
        if (subtitleEl) subtitleEl.textContent = defaultSubtitle;
        if (iconEl) iconEl.className = defaultIcon;
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

function showNotification(message, type = 'info', duration = 5000, isHTML = false) {
    if (!notificationContainer) return;
    
    const notification = document.createElement('div');
    notification.className = `notification ${type} ${isHTML ? 'html-content' : ''}`;
    
    if (isHTML) {
        notification.innerHTML = message;
        notification.style.cursor = 'pointer';
        notification.addEventListener('click', function() {
            if (this.parentNode) {
                this.remove();
            }
        });
    } else {
        const cleanMessage = stripLeadingEmojis(message);
        notification.innerHTML = `
            <i class="fas fa-${getNotificationIcon(type)}"></i>
            <span>${cleanMessage}</span>
        `;
    }
    
    notificationContainer.appendChild(notification);
    
    // Auto remove after duration
    setTimeout(() => {
        if (notification.parentNode) {
            notification.remove();
        }
    }, duration);
}

function getNotificationIcon(type) {
    switch(type) {
        case 'success': return 'check-circle';
        case 'error': return 'exclamation-circle';
        case 'warning': return 'exclamation-triangle';
        default: return 'info-circle';
    }
}

// Strip leading emojis so we don't show duplicates alongside the icon
function stripLeadingEmojis(text) {
    if (!text) return '';
    // Remove leading emoji/pictographic characters and surrounding whitespace
    return text.replace(/^[\p{Extended_Pictographic}\p{Emoji_Presentation}\uFE0F\u200D\s]+/u, '').trimStart();
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
        themeConfig.setupCompactThemeSelector(); // Refresh theme selector
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

// Pastel Theme Management System
const themeConfig = {
    currentTheme: 'default',
    themes: {
        'default': {
            name: 'Pastel Blue',
            colors: ['#4361ee', '#4cc9f0', '#7209b7']
        },
        'sage': {
            name: 'Sage Green', 
            colors: ['#2a9d8f', '#8ac926', '#f4a261']
        },
        'lavender': {
            name: 'Lavender',
            colors: ['#7209b7', '#b5179e', '#f72585']
        },
        'sand': {
            name: 'Warm Sand',
            colors: ['#f78c19', '#ff9e00', '#ffb703']
        },
        'slate': {
            name: 'Cool Slate', 
            colors: ['#495057', '#6c757d', '#adb5bd']
        },
        'rose': {
            name: 'Blush Rose',
            colors: ['#f72585', '#b5179e', '#7209b7']
        }
    },
    
    init() {
        console.log("🎨 Initializing theme system...");
        this.loadTheme();
        this.setupCompactThemeSelector();
        this.applyTheme(this.currentTheme);
    },
    
    loadTheme() {
        const savedTheme = localStorage.getItem('selectedTheme');
        if (savedTheme && this.themes[savedTheme]) {
            this.currentTheme = savedTheme;
            console.log("📁 Loaded saved theme:", savedTheme);
        } else {
            this.currentTheme = 'default';
            console.log("📁 Using default theme");
        }
    },
    
    setupCompactThemeSelector() {
        const themeContainer = document.getElementById('theme-selector');
        if (!themeContainer) {
            console.error("❌ Theme selector container not found!");
            return;
        }
        
        console.log("🔄 Setting up theme selector with themes:", Object.keys(this.themes));
        
        themeContainer.innerHTML = Object.entries(this.themes).map(([themeKey, theme]) => {
            console.log("🎨 Processing theme:", themeKey, theme);
            return `
                <div class="theme-option-compact ${themeKey === this.currentTheme ? 'active' : ''}" 
                     data-theme="${themeKey}" onclick="themeConfig.selectTheme('${themeKey}')">
                    <div class="theme-preview-colors-compact">
                        <div class="theme-color-primary-compact" style="background: ${theme.colors[0]}"></div>
                        <div class="theme-color-secondary-compact" style="background: ${theme.colors[1]}"></div>
                        <div class="theme-color-accent-compact" style="background: ${theme.colors[2]}"></div>
                    </div>
                    <div class="theme-name-compact">${theme.name}</div>
                </div>
            `;
        }).join('');
        
        console.log("✅ Theme selector setup complete");
    },
    
    selectTheme(theme) {
        console.log("🎨 Selecting theme:", theme);
        if (!this.themes[theme]) {
            console.error("❌ Theme not found:", theme);
            return;
        }
        
        this.currentTheme = theme;
        this.applyTheme(theme);
        this.setupCompactThemeSelector(); // Refresh the selector
        showNotification(`🎨 Switched to ${this.themes[theme].name} theme`, 'success');
    },
    
    applyTheme(theme) {
        console.log("🎨 Applying theme:", theme);
        document.body.setAttribute('data-theme', theme);
        localStorage.setItem('selectedTheme', theme);
        this.currentTheme = theme;
    },
    
    saveTheme(theme) {
        localStorage.setItem('selectedTheme', theme);
        this.currentTheme = theme;
    }
};

// UI Configuration System
const uiConfig = {
    settings: {
        compactMode: false,
        animations: true,
        sidebarCollapsed: false,
        highContrast: false,
        fontSize: 'medium',
        density: 'comfortable'
    },
    
    init() {
        this.loadSettings();
        this.applyAllSettings();
    },
    
    loadSettings() {
        const saved = localStorage.getItem('uiSettings');
        if (saved) {
            this.settings = { ...this.settings, ...JSON.parse(saved) };
        }
    },
    
    saveSettings() {
        localStorage.setItem('uiSettings', JSON.stringify(this.settings));
    },
    
    applyAllSettings() {
        this.applyCompactMode();
        this.applyAnimations();
        this.applySidebarState();
        this.applyHighContrast();
        this.applyFontSize();
        this.applyDensity();
    },
    
    // Compact Mode
    applyCompactMode() {
        if (this.settings.compactMode) {
            document.body.classList.add('compact-mode');
        } else {
            document.body.classList.remove('compact-mode');
        }
    },
    
    toggleCompactMode() {
        this.settings.compactMode = !this.settings.compactMode;
        this.applyCompactMode();
        this.saveSettings();
    },
    
    // Animations
    applyAnimations() {
        if (!this.settings.animations) {
            document.body.classList.add('no-animations');
        } else {
            document.body.classList.remove('no-animations');
        }
    },
    
    toggleAnimations() {
        this.settings.animations = !this.settings.animations;
        this.applyAnimations();
        this.saveSettings();
    },
    
    // Sidebar
    applySidebarState() {
        if (this.settings.sidebarCollapsed) {
            document.body.classList.add('sidebar-collapsed');
        } else {
            document.body.classList.remove('sidebar-collapsed');
        }
    },
    
    toggleSidebar() {
        this.settings.sidebarCollapsed = !this.settings.sidebarCollapsed;
        this.applySidebarState();
        this.saveSettings();
    },
    
    // High Contrast
    applyHighContrast() {
        if (this.settings.highContrast) {
            document.body.classList.add('high-contrast');
        } else {
            document.body.classList.remove('high-contrast');
        }
    },
    
    toggleHighContrast() {
        this.settings.highContrast = !this.settings.highContrast;
        this.applyHighContrast();
        this.saveSettings();
    },
    
    // Font Size
    applyFontSize() {
        document.body.className = document.body.className.replace(/\bfont-\w+\b/g, '');
        document.body.classList.add(`font-${this.settings.fontSize}`);
    },
    
    setFontSize(size) {
        this.settings.fontSize = size;
        this.applyFontSize();
        this.saveSettings();
    },
    
    // Density
    applyDensity() {
        document.body.className = document.body.className.replace(/\bdensity-\w+\b/g, '');
        document.body.classList.add(`density-${this.settings.density}`);
    },
    
    setDensity(density) {
        this.settings.density = density;
        this.applyDensity();
        this.saveSettings();
    }
};

function resetUISettings() {
    if (confirm('Reset all UI settings to defaults?')) {
        localStorage.removeItem('selectedTheme');
        localStorage.removeItem('uiSettings');
        themeConfig.init();
        uiConfig.init();
        showNotification('🔄 UI settings reset to defaults', 'success');
    }
}

// Update loadCurrentSettings function
function loadCurrentSettings() {
    // Update theme and UI controls
    const fontSizeEl = document.getElementById('font-size');
    if (fontSizeEl) fontSizeEl.value = uiConfig.settings.fontSize;
    
    const densityEl = document.getElementById('density');
    if (densityEl) densityEl.value = uiConfig.settings.density;
    
    const highContrastEl = document.getElementById('high-contrast-toggle');
    if (highContrastEl) highContrastEl.checked = uiConfig.settings.highContrast;
    
    const animationsEl = document.getElementById('animations-toggle');
    if (animationsEl) animationsEl.checked = uiConfig.settings.animations;
    
    const compactEl = document.getElementById('compact-toggle');
    if (compactEl) compactEl.checked = uiConfig.settings.compactMode;
    
    const sidebarEl = document.getElementById('sidebar-toggle');
    if (sidebarEl) sidebarEl.checked = uiConfig.settings.sidebarCollapsed;
}

function initializeExamSystem() {
    console.log("📝 Initializing exam system...");
    
    // Exam navigation
    setupExamNavigation();
    
    // Event listeners for exam section
    document.getElementById('generate-exam-schedule-btn')?.addEventListener('click', generateExamSchedule);
    document.getElementById('exam-cancel-btn')?.addEventListener('click', hideExamSection);
    document.getElementById('refresh-exam-btn')?.addEventListener('click', loadExamTimetables);
    document.getElementById('download-all-exam-btn')?.addEventListener('click', downloadAllExamTimetables);
    document.getElementById('exam-empty-generate-btn')?.addEventListener('click', showExamGenerateSection);
    
    // Date input formatting
    setupDateInputs();
    
    console.log("✅ Exam system initialized");
}

function setupExamNavigation() {
    const examNavItems = document.querySelectorAll('a[data-section^="exam-"]');
    
    examNavItems.forEach(item => {
        item.addEventListener('click', function(e) {
            e.preventDefault();
            
            const section = this.getAttribute('data-section');
            
            // Hide all exam sections first
            document.querySelectorAll('.exam-section').forEach(sec => {
                sec.style.display = 'none';
            });
            
            // Hide other main sections
            document.querySelector('.timetables-section').style.display = 'none';
            document.querySelector('.controls-section').style.display = 'none';
            document.querySelector('.quick-actions').style.display = 'none';
            if (document.getElementById('upload-section')) {
                document.getElementById('upload-section').style.display = 'none';
            }
            
            // Show selected exam section
            if (section === 'exam-generate') {
                document.getElementById('exam-generate-section').style.display = 'block';
                showExamGenerateSection();
            } else if (section === 'exam-view') {
                document.getElementById('exam-view-section').style.display = 'block';
                loadExamTimetables();
            }
            
            isExamSectionVisible = true;
            
            // Update navigation
            document.querySelectorAll('.nav-item').forEach(navItem => {
                navItem.classList.remove('active');
            });
            this.parentElement.classList.add('active');
        });
    });
}

function showExamGenerateSection() {
    if (document.getElementById('exam-generate-section')) {
        document.getElementById('exam-generate-section').style.display = 'block';
    }
    if (document.getElementById('exam-view-section')) {
        document.getElementById('exam-view-section').style.display = 'none';
    }
    
    // Hide other sections
    document.querySelector('.timetables-section').style.display = 'none';
    document.querySelector('.controls-section').style.display = 'none';
    document.querySelector('.quick-actions').style.display = 'none';
}

function hideExamSection() {
    document.querySelectorAll('.exam-section').forEach(section => {
        section.style.display = 'none';
    });
    
    // Show main sections
    document.querySelector('.timetables-section').style.display = 'block';
    document.querySelector('.controls-section').style.display = 'flex';
    document.querySelector('.quick-actions').style.display = 'block';
    
    isExamSectionVisible = false;
    
    // Reset to dashboard
    document.querySelectorAll('.nav-item').forEach(item => {
        item.classList.remove('active');
    });
    const dashboardNav = document.querySelector('[data-semester="all"]');
    if (dashboardNav) {
        dashboardNav.parentElement.classList.add('active');
    }
}

function setupDateInputs() {
    // Add date validation and formatting
    const dateInputs = document.querySelectorAll('.date-input');
    dateInputs.forEach(input => {
        input.addEventListener('blur', function() {
            const value = this.value.trim();
            if (value && isValidDate(value)) {
                this.classList.remove('error');
                this.classList.add('success');
            } else if (value) {
                this.classList.add('error');
                this.classList.remove('success');
            }
        });
    });
}

function isValidDate(dateString) {
    // DD/MM/YYYY format validation
    const regex = /^(\d{2})\/(\d{2})\/(\d{4})$/;
    if (!regex.test(dateString)) return false;
    
    const parts = dateString.split('/');
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);
    const year = parseInt(parts[2], 10);
    
    if (month < 1 || month > 12) return false;
    if (day < 1 || day > 31) return false;
    
    const date = new Date(year, month - 1, day);
    return date.getDate() === day && date.getMonth() === month - 1 && date.getFullYear() === year;
}

async function generateExamSchedule() {
    const startDateInput = document.getElementById('exam-period-start');
    const endDateInput = document.getElementById('exam-period-end');
    
    if (!startDateInput || !endDateInput) {
        showNotification('❌ Exam scheduling form not loaded properly', 'error');
        return;
    }
    
    const startDate = startDateInput.value;
    const endDate = endDateInput.value;
    
    // Validate dates
    if (!startDate || !endDate) {
        showNotification('❌ Please enter both start and end dates', 'error');
        return;
    }
    
    if (!isValidDate(startDate) || !isValidDate(endDate)) {
        showNotification('❌ Please enter valid dates in DD/MM/YYYY format', 'error');
        return;
    }
    
    // Ensure configuration is initialized
    if (!examConfig.current) {
        examConfig.loadConfig();
    }
    
    // Update configuration from UI
    examConfig.updateConfigFromUI();
    
    // Validate configuration
    const validationErrors = examConfig.validateConfig();
    if (validationErrors.length > 0) {
        showNotification(`❌ Configuration errors: ${validationErrors.join(', ')}`, 'error');
        return;
    }
    
    showLoading(true, {
        title: 'Generating exam schedule',
        subtitle: 'Optimizing slots and allocating rooms...',
        iconClass: 'fas fa-calendar-alt fa-spin'
    });
    
    try {
        const config = examConfig.getConfigForAPI();
        
        const response = await fetch('/exam-schedule', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                start_date: startDate,
                end_date: endDate,
                config: config
            })
        });
        
        const result = await response.json();
        
        if (result.success) {
            showNotification(`✅ ${result.message}`, 'success');
            console.log('📊 Exam schedule generated with config:', config);
            
            // Safely update preview
            try {
                updateExamPreview(result.schedule, config);
            } catch (previewError) {
                console.error('❌ Error updating preview:', previewError);
                showNotification('⚠️ Schedule generated but preview update failed', 'warning');
            }
            
            // Reload exam timetables
            await loadExamTimetables();
            
        } else {
            // Show detailed error message
            let errorMessage = result.message || 'Unknown error occurred';
            if (errorMessage.includes('No exam data found')) {
                errorMessage += ' - Please check if exams_data.csv is uploaded and contains valid data';
            } else if (errorMessage.includes('No exams could be scheduled')) {
                errorMessage += ' - Try increasing the exam period or maximum exams per day';
            }
            
            showNotification(`❌ ${errorMessage}`, 'error');
            
            // Show empty preview with error message
            const previewContent = document.getElementById('exam-preview-content');
            if (previewContent) {
                previewContent.innerHTML = `
                    <div class="empty-preview error">
                        <i class="fas fa-exclamation-triangle"></i>
                        <h4>Scheduling Failed</h4>
                        <p>${errorMessage}</p>
                        <div class="error-suggestions">
                            <h5>Suggestions:</h5>
                            <ul>
                                <li>Increase the exam period duration</li>
                                <li>Increase maximum exams per day</li>
                                <li>Include weekends in scheduling</li>
                                <li>Check if exams_data.csv has valid data</li>
                                <li>Relax department conflict settings</li>
                            </ul>
                        </div>
                    </div>
                `;
            }
        }
    } catch (error) {
        console.error('❌ Error generating exam schedule:', error);
        showNotification('❌ Error generating exam schedule: ' + error.message, 'error');
        
        // Show error in preview
        const previewContent = document.getElementById('exam-preview-content');
        if (previewContent) {
            previewContent.innerHTML = `
                <div class="empty-preview error">
                    <i class="fas fa-exclamation-circle"></i>
                    <h4>Network Error</h4>
                    <p>Failed to connect to the server. Please try again.</p>
                </div>
            `;
        }
    } finally {
        showLoading(false);
    }
}

// Enhanced updateExamPreview function with improved list view support
function updateExamPreview(schedule, config) {
    const previewContent = document.getElementById('exam-preview-content');
    const downloadBtn = document.getElementById('download-preview-btn');
    
    if (!previewContent) {
        console.error('❌ Preview content element not found');
        return;
    }
    
    // Use safe helper for schedule data
    const safeSchedule = safeArray(schedule);
    
    // Check if schedule data is valid
    if (!schedule || !Array.isArray(schedule) || schedule.length === 0) {
        previewContent.innerHTML = `
            <div class="empty-preview">
                <i class="fas fa-calendar-alt"></i>
                <h4>No Schedule Generated Yet</h4>
                <p>Configure the settings above and click "Generate Exam Schedule" to create your exam timetable</p>
            </div>
        `;
        if (downloadBtn) downloadBtn.style.display = 'none';
        return;
    }
    
    // Filter only scheduled exams using safe helpers
    const scheduledExams = safeSchedule.filter(e => safeGet(e, 'status') === 'Scheduled');
    
    if (scheduledExams.length === 0) {
        previewContent.innerHTML = `
            <div class="empty-preview">
                <i class="fas fa-exclamation-triangle"></i>
                <h4>No Exams Scheduled</h4>
                <p>The schedule was generated but no exams could be scheduled. Try adjusting your configuration.</p>
            </div>
        `;
        if (downloadBtn) downloadBtn.style.display = 'none';
        return;
    }
    
    // Show download button
    if (downloadBtn) downloadBtn.style.display = 'inline-flex';
    
    // Calculate statistics using safe helpers
    const uniqueDates = [...new Set(safeSchedule.map(e => safeString(safeGet(e, 'date'))))];
    const uniqueScheduledDates = [...new Set(scheduledExams.map(e => safeString(safeGet(e, 'date'))))];
    
    // Sort dates properly
    const sortedDates = sortDates(uniqueDates);
    const sortedScheduledDates = sortDates(uniqueScheduledDates);
    
    const totalDays = sortedDates.length;
    const daysWithExams = sortedScheduledDates.length;
    const freeDays = totalDays - daysWithExams;

    // Initialize examsByDate safely
    const examsByDate = {};
    scheduledExams.forEach(exam => {
        const date = safeString(safeGet(exam, 'date'));
        if (date) {
            if (!examsByDate[date]) {
                examsByDate[date] = [];
            }
            examsByDate[date].push(exam);
        }
    });

    // Initialize html variable here
    let html = '';

    // Enhanced statistics with beautiful summary
    html += `
        <div class="preview-summary">
            <div class="summary-grid">
                <div class="summary-stat">
                    <div class="summary-number">${scheduledExams.length}</div>
                    <div class="summary-label">Total Exams</div>
                </div>
                <div class="summary-stat">
                    <div class="summary-number">${daysWithExams}</div>
                    <div class="summary-label">Exam Days</div>
                </div>
                <div class="summary-stat">
                    <div class="summary-number">${freeDays}</div>
                    <div class="summary-label">Free Days</div>
                </div>
                <div class="summary-stat">
                    <div class="summary-number">${scheduledExams.filter(e => safeGet(e, 'session') === 'Morning').length}</div>
                    <div class="summary-label">Morning Sessions</div>
                </div>
                <div class="summary-stat">
                    <div class="summary-number">${scheduledExams.filter(e => safeGet(e, 'session') === 'Afternoon').length}</div>
                    <div class="summary-label">Afternoon Sessions</div>
                </div>
            </div>
        </div>
        
        <!-- View Toggle -->
        <div class="view-toggle-container">
            <div class="view-toggle-group">
                <button class="view-toggle-btn active" data-view="daily">
                    <i class="fas fa-calendar-day"></i>
                    <span>Daily View</span>
                </button>
                <button class="view-toggle-btn" data-view="list">
                    <i class="fas fa-list"></i>
                    <span>List View</span>
                </button>
            </div>
            <div class="view-actions">
                <button class="action-btn" id="export-list-btn" title="Export to Excel">
                    <i class="fas fa-file-excel"></i>
                </button>
                <button class="action-btn" id="print-list-btn" title="Print List">
                    <i class="fas fa-print"></i>
                </button>
            </div>
        </div>
        
        <div class="daily-schedule-full" id="daily-schedule-view">
    `;
    
    // Create daily schedule view with sorted dates
    sortedDates.forEach(date => {
        const dayExams = examsByDate[date] || [];
        if (dayExams.length === 0) return;
        
        const dayName = safeString(safeGet(dayExams[0], 'day'));
        const formattedDate = formatDateForDisplay(date);
        
        // Group exams by session
        const morningExams = dayExams.filter(e => safeGet(e, 'session') === 'Morning');
        const afternoonExams = dayExams.filter(e => safeGet(e, 'session') === 'Afternoon');
        
        html += `
            <div class="day-slot-full">
                <div class="day-header-full">
                    <div>
                        <h4>${dayName}</h4>
                        <div class="day-date-full">${formattedDate}</div>
                    </div>
                    <div class="day-stats-full">
                        <span class="session-badge morning">${morningExams.length} Morning</span>
                        <span class="session-badge afternoon">${afternoonExams.length} Afternoon</span>
                    </div>
                </div>
                <div class="day-content-full">
                    <div class="time-slots-container">
                        <!-- Morning Slot -->
                        <div class="time-slot-group morning-slot">
                            <div class="time-slot-header">
                                <i class="fas fa-sun"></i>
                                <div class="slot-info">
                                    <span class="slot-title">Morning Session</span>
                                    <span class="slot-time">09:00 - 12:00</span>
                                </div>
                                <span class="slot-count">${morningExams.length} exams</span>
                            </div>
                            <div class="slot-exams-grid">
                                ${morningExams.length > 0 ? 
                                    morningExams.map(exam => createExamCardWithTime(exam)).join('') : 
                                    '<div class="no-exams">No exams scheduled</div>'
                                }
                            </div>
                        </div>
                        
                        <!-- Afternoon Slot -->
                        <div class="time-slot-group afternoon-slot">
                            <div class="time-slot-header">
                                <i class="fas fa-cloud-sun"></i>
                                <div class="slot-info">
                                    <span class="slot-title">Afternoon Session</span>
                                    <span class="slot-time">14:00 - 17:00</span>
                                </div>
                                <span class="slot-count">${afternoonExams.length} exams</span>
                            </div>
                            <div class="slot-exams-grid">
                                ${afternoonExams.length > 0 ? 
                                    afternoonExams.map(exam => createExamCardWithTime(exam)).join('') : 
                                    '<div class="no-exams">No exams scheduled</div>'
                                }
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    html += `</div>`;
    
    // Add improved list view (hidden by default)
    html += `
        <div class="list-schedule-view" id="list-schedule-view" style="display: none;">
            ${createListViewTable(scheduledExams)}
        </div>
    `;
    
    previewContent.innerHTML = html;
    
    // Setup exam tooltips after rendering
    setTimeout(() => {
        setupExamTooltips();
        setupListViewInteractions();
    }, 100);
    
    // Add download functionality
    if (downloadBtn) {
        downloadBtn.onclick = function() {
            showNotification('📥 Preparing exam schedule download...', 'info');
        };
    }
    
    // Add view toggle functionality
    setupViewToggle();
}

// IMPROVED: Create list view table with better UI
function createListViewTable(scheduledExams) {
    // Sort exams by date, then by session
    const sortedExams = scheduledExams.sort((a, b) => {
        const dateA = parseDDMMYYYY(a.date);
        const dateB = parseDDMMYYYY(b.date);
        
        if (dateA.getTime() !== dateB.getTime()) {
            return dateA - dateB;
        }
        
        // If same date, sort by session (Morning before Afternoon)
        if (a.session === 'Morning' && b.session === 'Afternoon') return -1;
        if (a.session === 'Afternoon' && b.session === 'Morning') return 1;
        return 0;
    });
    
    return `
        <div class="list-view-container">
            <div class="list-view-header">
                <div class="list-view-info">
                    <i class="fas fa-info-circle"></i>
                    <span>Showing ${sortedExams.length} scheduled exams</span>
                </div>
                <div class="list-view-controls">
                    <button class="control-btn" id="toggle-departments">
                        <i class="fas fa-layer-group"></i>
                        <span>Group by Department</span>
                    </button>
                    <button class="control-btn" id="filter-session">
                        <i class="fas fa-filter"></i>
                        <span>Filter Session</span>
                    </button>
                </div>
            </div>
            
            <div class="table-container">
                <table class="exam-list-table">
                    <thead>
                        <tr>
                            <th class="date-col">Date</th>
                            <th class="day-col">Day</th>
                            <th class="session-col">Session</th>
                            <th class="time-col">Time</th>
                            <th class="course-col">Course</th>
                            <th class="type-col">Type</th>
                            <th class="duration-col">Duration</th>
                            <th class="dept-col">Department</th>
                            <th class="semester-col">Semester</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sortedExams.map((exam, index) => createListViewRow(exam, index)).join('')}
                    </tbody>
                </table>
            </div>
            
            <div class="list-view-footer">
                <div class="footer-stats">
                    <div class="footer-stat">
                        <span class="stat-value">${new Set(sortedExams.map(e => e.date)).size}</span>
                        <span class="stat-label">Exam Days</span>
                    </div>
                    <div class="footer-stat">
                        <span class="stat-value">${sortedExams.filter(e => e.session === 'Morning').length}</span>
                        <span class="stat-label">Morning Exams</span>
                    </div>
                    <div class="footer-stat">
                        <span class="stat-value">${sortedExams.filter(e => e.session === 'Afternoon').length}</span>
                        <span class="stat-label">Afternoon Exams</span>
                    </div>
                </div>
            </div>
        </div>
    `;
}

// IMPROVED: Create list view row with better spacing and visual design
function createListViewRow(exam, index) {
    const formattedDate = formatDateForDisplay(exam.date);
    const sessionClass = exam.session.toLowerCase();
    const deptClass = `dept-${exam.department.toLowerCase().replace(/\s+/g, '-')}`;
    const rowClass = index % 2 === 0 ? 'even-row' : 'odd-row';
    
    return `
        <tr class="exam-list-row ${rowClass} ${sessionClass}-row" data-exam-id="${exam.course_code}-${exam.date}">
            <td class="date-cell">
                <div class="date-content">
                    <div class="date-main">${formattedDate}</div>
                </div>
            </td>
            <td class="day-cell">
                <span class="day-name">${exam.day}</span>
            </td>
            <td class="session-cell">
                <div class="session-indicator ${sessionClass}">
                    <i class="fas ${sessionClass === 'morning' ? 'fa-sun' : 'fa-cloud-sun'}"></i>
                    <span>${exam.session}</span>
                </div>
            </td>
            <td class="time-cell">
                <div class="time-slot">${exam.time_slot || '09:00-12:00'}</div>
            </td>
            <td class="course-cell">
                <div class="course-info">
                    <div class="course-code">${exam.course_code}</div>
                    <div class="course-name">${exam.course_name}</div>
                </div>
            </td>
            <td class="type-cell">
                <span class="exam-type-badge ${exam.exam_type.toLowerCase()}">${exam.exam_type}</span>
            </td>
            <td class="duration-cell">
                <div class="duration-display">
                    <i class="fas fa-clock"></i>
                    <span>${exam.duration}</span>
                </div>
            </td>
            <td class="dept-cell">
                <span class="dept-badge ${deptClass}">${exam.department}</span>
            </td>
            <td class="semester-cell">
                <span class="semester-pill">Sem ${exam.semester || 'N/A'}</span>
            </td>
        </tr>
    `;
}

// NEW: Setup list view interactions
function setupListViewInteractions() {
    // Add hover effects and click interactions
    const examRows = document.querySelectorAll('.exam-list-row');
    
    examRows.forEach(row => {
        row.addEventListener('mouseenter', function() {
            this.style.transform = 'translateY(-2px)';
            this.style.boxShadow = '0 4px 12px rgba(0,0,0,0.1)';
        });
        
        row.addEventListener('mouseleave', function() {
            this.style.transform = 'translateY(0)';
            this.style.boxShadow = 'none';
        });
        
        row.addEventListener('click', function() {
            // Toggle detailed view
            this.classList.toggle('expanded');
        });
    });
    
    // Setup control buttons
    const toggleDeptBtn = document.getElementById('toggle-departments');
    const filterSessionBtn = document.getElementById('filter-session');
    
    if (toggleDeptBtn) {
        toggleDeptBtn.addEventListener('click', function() {
            document.querySelector('.table-container').classList.toggle('group-by-dept');
            this.classList.toggle('active');
        });
    }
    
    if (filterSessionBtn) {
        filterSessionBtn.addEventListener('click', function() {
            // Implement session filtering
            showNotification('Session filter functionality coming soon!', 'info');
        });
    }
    
    // Setup export and print buttons
    const exportBtn = document.getElementById('export-list-btn');
    const printBtn = document.getElementById('print-list-btn');
    
    if (exportBtn) {
        exportBtn.addEventListener('click', function() {
            showNotification('Export to Excel functionality coming soon!', 'info');
        });
    }
    
    if (printBtn) {
        printBtn.addEventListener('click', function() {
            window.print();
        });
    }
}

function getCourseDetails(courseCode, semester, department) {
    if (!courseDatabase || Object.keys(courseDatabase).length === 0) {
        console.warn('⚠️ Course database not loaded');
        return {};
    }
    
    // Try exact match first in course database
    let courseDetails = courseDatabase[courseCode];
    
    // If not found, try to find by course code pattern in course database
    if (!courseDetails) {
        const matchingCourse = Object.entries(courseDatabase).find(([code, details]) => {
            return code.includes(courseCode) || courseCode.includes(code) || 
                   details.name && details.name.includes(courseCode);
        });
        
        if (matchingCourse) {
            courseDetails = matchingCourse[1];
        }
    }
    
    // If still not found, try to find by semester and department in course database
    if (!courseDetails) {
        const semesterCourses = Object.values(courseDatabase).filter(course => {
            return course.semester == semester && course.department === department;
        });
        
        if (semesterCourses.length > 0) {
            // Find the most relevant course by name similarity
            courseDetails = semesterCourses[0];
        }
    }
    
    return courseDetails || {};
}

function createExamCardWithTime(exam) {
    if (!exam) return '';
    
    const department = safeString(safeGet(exam, 'department', 'general'));
    const deptClass = `dept-${department.toLowerCase().replace(' ', '-')}`;
    const courseName = safeString(safeGet(exam, 'course_name', 'Unknown Course'));
    const duration = safeString(safeGet(exam, 'duration', 'N/A'));
    const courseCode = safeString(safeGet(exam, 'course_code', 'N/A'));
    const examType = safeString(safeGet(exam, 'exam_type', 'N/A'));
    const timeSlot = safeString(safeGet(exam, 'time_slot', 'Time N/A'));
    const semesterFromExam = safeString(safeGet(exam, 'semester', 'N/A'));
    const creditsFromExam = safeString(safeGet(exam, 'credits', 'N/A'));
    const instructorFromExam = safeString(safeGet(exam, 'instructor', 'Not assigned'));
    const preferredDate = safeString(safeGet(exam, 'original_preferred', 'Not specified'));
    const ltpscFromExam = safeString(safeGet(exam, 'ltpsc', 'N/A'));
    const courseTypeFromExam = safeString(safeGet(exam, 'course_type', 'Core'));
    
    // Get enhanced data from course database (course_data.csv)
    const courseDetails = getCourseDetails(courseCode, semesterFromExam, department);
    
    // Priority: course_data.csv > exams_data.csv
    const finalInstructor = courseDetails.instructor || instructorFromExam;
    const finalCredits = courseDetails.credits || creditsFromExam;
    const finalLTPSC = courseDetails.ltpsc || ltpscFromExam;
    const finalCourseType = courseDetails.type || courseTypeFromExam;
    const finalSemester = courseDetails.semester || semesterFromExam;
    
    // Get session for additional styling
    const session = safeString(safeGet(exam, 'session', ''));
    const sessionClass = session.toLowerCase();
    
    return `
        <div class="exam-card-time ${deptClass} ${sessionClass}" style="position: relative;" 
             onclick="showExamCardNotification('${courseCode}', '${courseName.replace(/'/g, "\\'")}', '${finalInstructor.replace(/'/g, "\\'")}', '${finalSemester}', '${finalCredits}', '${finalLTPSC}', '${department}', '${examType}', '${timeSlot}', '${session}')">
            <div class="exam-time-badge">${timeSlot}</div>
            <div class="exam-card-content">
                <div class="exam-header">
                    <span class="exam-code">${courseCode}</span>
                    <span class="exam-type">${examType}</span>
                </div>
                <div class="exam-details">
                    <div class="exam-name">${courseName}</div>
                    <div class="exam-meta">
                        <span><i class="fas fa-clock"></i> ${duration}</span>
                        <span><i class="fas fa-building"></i> ${department}</span>
                        <span><i class="fas fa-graduation-cap"></i> Sem ${finalSemester}</span>
                    </div>
                </div>
            </div>
            
            <!-- Enhanced Exam Tooltip -->
            <div class="exam-tooltip-enhanced ${deptClass}">
                <div class="exam-tooltip-header">
                    <div class="exam-tooltip-code">${courseCode}</div>
                    <div class="exam-tooltip-type">${examType} Exam</div>
                </div>
                <div class="exam-tooltip-name">${courseName}</div>
                
                <div class="exam-tooltip-details">
                    <!-- Semester and LTPSC prominently displayed -->
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item full-width highlight">
                            <i class="fas fa-graduation-cap"></i>
                            <span><strong>Semester ${finalSemester}</strong> • ${finalCredits} Credits</span>
                        </div>
                    </div>
                    
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item full-width highlight">
                            <i class="fas fa-book"></i>
                            <span><strong>LTPSC Structure:</strong> ${finalLTPSC}</span>
                        </div>
                    </div>
                    
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item">
                            <i class="fas fa-user-tie"></i>
                            <span><strong>Instructor:</strong> ${finalInstructor}</span>
                        </div>
                        <div class="tooltip-item">
                            <i class="fas fa-tag"></i>
                            <span><strong>Type:</strong> ${finalCourseType}</span>
                        </div>
                    </div>
                    
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item">
                            <i class="fas fa-calendar-alt"></i>
                            <span><strong>Semester:</strong> ${finalSemester}</span>
                        </div>
                        <div class="tooltip-item">
                            <i class="fas fa-building"></i>
                            <span><strong>Department:</strong> ${department}</span>
                        </div>
                    </div>
                    
                    ${preferredDate && preferredDate !== 'Not specified' ? `
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item full-width">
                            <i class="fas fa-calendar-check"></i>
                            <span><strong>Preferred Date:</strong> ${preferredDate}</span>
                        </div>
                    </div>
                    ` : ''}
                    
                    <!-- Data Source Info -->
                    <div class="exam-tooltip-row">
                        <div class="tooltip-item full-width" style="font-size: 0.7rem; color: var(--gray);">
                            <i class="fas fa-database"></i>
                            <span>Data: ${courseDetails.semester ? 'course_data.csv' : 'exams_data.csv'}</span>
                        </div>
                    </div>
                </div>
                
                <div class="exam-tooltip-footer">
                    <div class="session-indicator ${sessionClass}">
                        <i class="fas ${session === 'Morning' ? 'fa-sun' : 'fa-cloud-sun'}"></i>
                        ${session} Session • ${duration}
                    </div>
                    <div class="click-hint">
                        <small><i class="fas fa-mouse-pointer"></i> Click for details</small>
                    </div>
                </div>
            </div>
        </div>
    `;
}
 
function showExamCardNotification(courseCode, courseName, instructor, semester, credits, ltpsc, department, examType, timeSlot, session) {
    // Parse LTPSC for better display
    const ltpscParts = ltpsc.split('-');
    const ltpscDisplay = ltpscParts.length === 5 ? 
        `L:${ltpscParts[0]} T:${ltpscParts[1]} P:${ltpscParts[2]} S:${ltpscParts[3]} C:${ltpscParts[4]}` : 
        ltpsc;
    
    const sessionIcon = session === 'Morning' ? '☀️' : '🌤️';
    
    const notificationHTML = `
        <div class="exam-notification-content">
            <div class="exam-notification-header">
                <div class="exam-notification-title">
                    <span class="exam-notification-code">${courseCode}</span>
                    <span class="exam-notification-type">${examType} Exam</span>
                </div>
                <div class="exam-notification-session ${session.toLowerCase()}">
                    ${sessionIcon} ${session} Session
                </div>
            </div>
            
            <div class="exam-notification-body">
                <div class="exam-notification-course">${courseName}</div>
                
                <div class="exam-notification-details">
                    <div class="detail-row">
                        <div class="detail-item">
                            <i class="fas fa-user-tie"></i>
                            <span><strong>Instructor:</strong> ${instructor}</span>
                        </div>
                        <div class="detail-item">
                            <i class="fas fa-graduation-cap"></i>
                            <span><strong>Semester:</strong> ${semester}</span>
                        </div>
                    </div>
                    
                    <div class="detail-row">
                        <div class="detail-item">
                            <i class="fas fa-star"></i>
                            <span><strong>Credits:</strong> ${credits}</span>
                        </div>
                        <div class="detail-item">
                            <i class="fas fa-building"></i>
                            <span><strong>Department:</strong> ${department}</span>
                        </div>
                    </div>
                    
                    <div class="detail-row">
                        <div class="detail-item full-width">
                            <i class="fas fa-book"></i>
                            <span><strong>LTPSC Structure:</strong> ${ltpscDisplay}</span>
                        </div>
                    </div>
                    
                    <div class="detail-row">
                        <div class="detail-item full-width">
                            <i class="fas fa-clock"></i>
                            <span><strong>Time Slot:</strong> ${timeSlot}</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <div class="exam-notification-footer">
                <small>Click anywhere to dismiss</small>
            </div>
        </div>
    `;
    
    showNotification(notificationHTML, 'info', 8000, true);
}

function setupExamTooltips() {
    // This function will be called after the exam preview is rendered
    const examCards = document.querySelectorAll('.exam-card-time');
    
    examCards.forEach(card => {
        // Prevent tooltip from going off-screen
        const tooltip = card.querySelector('.exam-tooltip');
        if (tooltip) {
            card.addEventListener('mouseenter', function() {
                const rect = this.getBoundingClientRect();
                const tooltipRect = tooltip.getBoundingClientRect();
                
                // Check if tooltip would go off the left edge
                if (rect.left - tooltipRect.width / 2 < 10) {
                    tooltip.style.left = '0';
                    tooltip.style.transform = 'translateX(0)';
                } 
                // Check if tooltip would go off the right edge
                else if (rect.left + tooltipRect.width / 2 > window.innerWidth - 10) {
                    tooltip.style.left = 'auto';
                    tooltip.style.right = '0';
                    tooltip.style.transform = 'translateX(0)';
                } else {
                    tooltip.style.left = '50%';
                    tooltip.style.right = 'auto';
                    tooltip.style.transform = 'translateX(-50%)';
                }
            });
        }
    });
}

function createExamCardFull(exam) {
    const deptClass = `dept-${exam.department.toLowerCase().replace(' ', '-')}`;
    
    return `
        <div class="exam-card-full ${deptClass}">
            <div class="exam-card-header-full">
                <div class="exam-code-full">${exam.course_code}</div>
                <div class="exam-type-full">${exam.exam_type}</div>
            </div>
            <div class="exam-details-full">
                <div class="exam-name-full">${exam.course_name}</div>
                <div class="exam-meta-full">
                    <span class="meta-item-full">
                        <i class="fas fa-clock"></i>
                        ${exam.duration}
                    </span>
                    <span class="meta-item-full">
                        <i class="fas fa-building"></i>
                        ${exam.department}
                    </span>
                    <span class="meta-item-full">
                        <i class="fas fa-graduation-cap"></i>
                        Sem ${exam.semester}
                    </span>
                    <span class="meta-item-full">
                        <i class="fas fa-calendar-check"></i>
                        Preferred: ${exam.original_preferred}
                    </span>
                </div>
            </div>
        </div>
    `;
}

function createExamCard(exam) {
    const deptClass = `dept-${exam.department.toLowerCase().replace(' ', '-')}`;
    
    return `
        <div class="exam-card ${deptClass}">
            <div class="exam-card-header">
                <div class="exam-code">${exam.course_code}</div>
                <div class="exam-type">${exam.exam_type}</div>
            </div>
            <div class="exam-details">
                <div class="exam-name">${exam.course_name}</div>
                <div class="exam-meta">
                    <span class="meta-item">
                        <i class="fas fa-clock"></i>
                        ${exam.duration}
                    </span>
                    <span class="meta-item">
                        <i class="fas fa-building"></i>
                        ${exam.department}
                    </span>
                    <span class="meta-item">
                        <i class="fas fa-graduation-cap"></i>
                        Sem ${exam.semester}
                    </span>
                </div>
            </div>
        </div>
    `;
}

function setupViewToggle() {
    const dailyViewBtn = document.getElementById('toggle-daily-view');
    const listViewBtn = document.getElementById('toggle-list-view');
    const dailyView = document.getElementById('daily-schedule-view');
    const listView = document.getElementById('list-schedule-view');
    
    if (!dailyViewBtn || !listViewBtn || !dailyView || !listView) return;
    
    dailyViewBtn.addEventListener('click', function() {
        this.classList.add('active');
        listViewBtn.classList.remove('active');
        dailyView.style.display = 'block';
        listView.style.display = 'none';
    });
    
    listViewBtn.addEventListener('click', function() {
        this.classList.add('active');
        dailyViewBtn.classList.remove('active');
        dailyView.style.display = 'none';
        listView.style.display = 'block';
    });
    
    // Set default active state
    dailyViewBtn.classList.add('active');
}

async function loadExamTimetables(showAll = false) {
    try {
        showAllSchedules = showAll;
        
        const endpoint = showAll ? '/exam-timetables/all' : '/exam-timetables';
        const response = await fetch(endpoint);
        const examTimetables = await response.json();
        
        if (showAll) {
            allExamTimetables = examTimetables;
        } else {
            currentExamTimetables = examTimetables;
        }
        
        renderExamTimetables();
        
    } catch (error) {
        console.error('❌ Error loading exam timetables:', error);
        showNotification('❌ Error loading exam timetables: ' + error.message, 'error');
    }
}

// Add function to toggle between current and all schedules
function toggleScheduleView() {
    showAllSchedules = !showAllSchedules;
    loadExamTimetables(showAllSchedules);
    
    const toggleBtn = document.getElementById('toggle-schedule-view');
    if (toggleBtn) {
        if (showAllSchedules) {
            toggleBtn.innerHTML = '<i class="fas fa-eye"></i> Show Current Only';
            toggleBtn.classList.remove('btn-outline');
            toggleBtn.classList.add('btn-primary');
            showNotification('📂 Showing all previously generated schedules', 'info');
        } else {
            toggleBtn.innerHTML = '<i class="fas fa-history"></i> Show All Schedules';
            toggleBtn.classList.remove('btn-primary');
            toggleBtn.classList.add('btn-outline');
            showNotification('📋 Showing current schedules only', 'info');
        }
    }
}

function renderExamTimetables() {
    const container = document.getElementById('exam-timetables-container');
    const emptyState = document.getElementById('exam-empty-state');
    
    if (!container) return;
    
    const timetablesToRender = showAllSchedules ? allExamTimetables : currentExamTimetables;
    
    if (timetablesToRender.length === 0) {
        if (emptyState) {
            emptyState.style.display = 'block';
            if (showAllSchedules) {
                emptyState.querySelector('h3').textContent = 'No Exam Schedules Found';
                emptyState.querySelector('p').textContent = 'No previously generated exam schedules were found in the system.';
            }
        }
        container.innerHTML = '';
        return;
    }
    
    if (emptyState) emptyState.style.display = 'none';
    
    // Get current view mode
    const activeViewRaw = document.querySelector('.toggle-btn.active')?.dataset.view || 'session';
    const activeView = activeViewRaw === 'daily' ? 'daily' : 'session';
    
    let html = '';
    
    // Add header for all schedules view
    if (showAllSchedules) {
        html += `
            <div class="schedule-view-header">
                <div class="view-info">
                    <i class="fas fa-history"></i>
                    <div>
                        <h3>All Generated Exam Schedules</h3>
                        <p>Showing ${timetablesToRender.length} previously generated schedule(s)</p>
                    </div>
                </div>
                <div class="view-actions">
                    <button class="btn btn-outline" onclick="clearAllSchedulesFromDisplay()">
                        <i class="fas fa-ban"></i>
                        Clear All from Display
                    </button>
                </div>
            </div>
        `;
    }
    
    timetablesToRender.forEach(timetable => {
        const scheduledExams = timetable.schedule_data ? 
            timetable.schedule_data.filter(e => e.status === 'Scheduled') : [];
        
        if (scheduledExams.length === 0) {
            html += `
                <div class="no-exams-message">
                    <i class="fas fa-calendar-times"></i>
                    <h4>No Exams Scheduled</h4>
                    <p>This timetable period has no scheduled exams</p>
                </div>
            `;
            return;
        }
        
        if (activeView === 'session') {
            html += renderSessionView(timetable, scheduledExams);
        } else {
            html += renderDailyView(timetable, scheduledExams);
        }
    });
    
    container.innerHTML = html;
    
    // Add view mode toggle functionality
    setupViewModeToggle();
}

// Update the session view to include management buttons for all schedules
function renderSessionView(timetable, scheduledExams) {
    const isCurrentlyDisplayed = timetable.is_currently_displayed !== false;
    
    let managementButtons = '';
    if (showAllSchedules) {
        if (isCurrentlyDisplayed) {
            managementButtons = `
                <button class="btn btn-sm btn-warning" onclick="removeScheduleFromDisplay('${timetable.filename}')">
                    <i class="fas fa-eye-slash"></i>
                    Remove from Display
                </button>
            `;
        } else {
            managementButtons = `
                <button class="btn btn-sm btn-success" onclick="addScheduleToDisplay('${timetable.filename}')">
                    <i class="fas fa-eye"></i>
                    Add to Display
                </button>
            `;
        }
    }
    
    // ... rest of the session view rendering code ...
    
    return `
        <div class="exam-timetable-card ${!isCurrentlyDisplayed ? 'not-displayed' : ''}">
            <div class="exam-timetable-header">
                <div>
                    <h3>Exam Schedule - ${timetable.period}</h3>
                    ${!isCurrentlyDisplayed ? '<span class="status-badge not-displayed-badge">Not in Display</span>' : ''}
                </div>
                <div class="exam-actions">
                    ${managementButtons}
                    <button class="action-btn" onclick="downloadExamTimetable('${timetable.filename}')" title="Download">
                        <i class="fas fa-download"></i>
                    </button>
                    <button class="action-btn" onclick="printExamTimetable('${timetable.filename}')" title="Print">
                        <i class="fas fa-print"></i>
                    </button>
                </div>
            </div>
            <!-- ... rest of the session view content ... -->
        </div>
    `;
}

// Add management functions
async function addScheduleToDisplay(filename) {
    try {
        const response = await fetch('/exam-timetables/add-to-display', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ filename })
        });
        
        const result = await response.json();
        if (result.success) {
            showNotification('✅ Schedule added to display', 'success');
            // Reload both views
            await loadExamTimetables(showAllSchedules);
            await loadExamTimetables(false);
        } else {
            showNotification('❌ ' + result.message, 'error');
        }
    } catch (error) {
        console.error('Error adding schedule to display:', error);
        showNotification('❌ Error adding schedule to display', 'error');
    }
}

async function removeScheduleFromDisplay(filename) {
    try {
        const response = await fetch('/exam-timetables/remove-from-display', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ filename })
        });
        
        const result = await response.json();
        if (result.success) {
            showNotification('✅ Schedule removed from display', 'success');
            // Reload both views
            await loadExamTimetables(showAllSchedules);
            await loadExamTimetables(false);
        } else {
            showNotification('❌ ' + result.message, 'error');
        }
    } catch (error) {
        console.error('Error removing schedule from display:', error);
        showNotification('❌ Error removing schedule from display', 'error');
    }
}

async function clearAllSchedulesFromDisplay() {
    if (!confirm('Are you sure you want to remove all schedules from display? This will only hide them, not delete the files.')) {
        return;
    }
    
    try {
        const response = await fetch('/exam-timetables/clear-display', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            }
        });
        
        const result = await response.json();
        if (result.success) {
            showNotification('✅ All schedules removed from display', 'success');
            // Reload both views
            await loadExamTimetables(showAllSchedules);
            await loadExamTimetables(false);
        } else {
            showNotification('❌ ' + result.message, 'error');
        }
    } catch (error) {
        console.error('Error clearing schedules from display:', error);
        showNotification('❌ Error clearing schedules from display', 'error');
    }
}

// Update the exam view section initialization
function initializeExamSystem() {
    console.log("📝 Initializing exam system...");
    
    // Initialize configuration system
    examConfig.init();
    
    // Exam navigation
    setupExamNavigation();
    
    // Event listeners for exam section
    document.getElementById('generate-exam-schedule-btn')?.addEventListener('click', generateExamSchedule);
    document.getElementById('exam-cancel-btn')?.addEventListener('click', hideExamSection);
    document.getElementById('refresh-exam-btn')?.addEventListener('click', () => loadExamTimetables(showAllSchedules));
    document.getElementById('download-all-exam-btn')?.addEventListener('click', downloadAllExamTimetables);
    document.getElementById('exam-empty-generate-btn')?.addEventListener('click', showExamGenerateSection);
    
    // Add toggle button for schedule view
    const toggleViewBtn = document.getElementById('toggle-schedule-view');
    if (!toggleViewBtn) {
        // Add the button to the exam controls if it doesn't exist
        const examControls = document.querySelector('.exam-controls .control-actions');
        if (examControls) {
            examControls.innerHTML += `
                <button class="btn btn-outline" id="toggle-schedule-view">
                    <i class="fas fa-history"></i>
                    Show All Schedules
                </button>
            `;
            document.getElementById('toggle-schedule-view').addEventListener('click', toggleScheduleView);
        }
    } else {
        toggleViewBtn.addEventListener('click', toggleScheduleView);
    }
    
    // Date input formatting
    setupDateInputs();
    
    console.log("✅ Exam system initialized");
}

function renderExamsTable(exams) {
    // Sort exams by date
    const sortedExams = exams.sort((a, b) => parseDDMMYYYY(a.date) - parseDDMMYYYY(b.date));
    
    return `
        <table class="exams-table">
            <thead>
                <tr>
                    <th>Date & Day</th>
                    <th>Course Code</th>
                    <th>Course Name</th>
                    <th>Duration</th>
                    <th>Department</th>
                    <th>Exam Type</th>
                    <th>Time Slot</th>
                </tr>
            </thead>
            <tbody>
                ${sortedExams.map(exam => {
                    const formattedDate = formatDateForDisplay(exam.date);
                    return `
                    <tr>
                        <td class="date-cell">
                            <div>${formattedDate}</div>
                            <div style="font-size: 0.8rem; color: var(--gray);">${exam.day}</div>
                        </td>
                        <td><span class="course-code">${exam.course_code}</span></td>
                        <td>${exam.course_name}</td>
                        <td>${exam.duration}</td>
                        <td><span class="department">${exam.department}</span></td>
                        <td><span class="exam-type">${exam.exam_type}</span></td>
                        <td class="time-cell">${exam.time_slot}</td>
                    </tr>
                    `;
                }).join('')}
            </tbody>
        </table>
    `;
}

function renderDailyView(timetable, scheduledExams) {
    // Group exams by date
    const examsByDate = {};
    scheduledExams.forEach(exam => {
        const date = exam.date;
        if (!examsByDate[date]) {
            examsByDate[date] = [];
        }
        examsByDate[date].push(exam);
    });
    
    // Sort dates
    const sortedDates = sortDates(Object.keys(examsByDate));
    
    let dailyHtml = '';
    sortedDates.forEach(date => {
        const dayExams = examsByDate[date];
        const dayName = dayExams[0].day;
        const morningExams = dayExams.filter(e => e.session === 'Morning');
        const afternoonExams = dayExams.filter(e => e.session === 'Afternoon');
        const formattedDate = formatDateForDisplay(date);
        
        dailyHtml += `
            <div class="daily-card">
                <div class="daily-header">
                    <div>
                        <div class="daily-date">${formattedDate}</div>
                        <div class="daily-day">${dayName}</div>
                    </div>
                    <div class="daily-stats">
                        <span class="session-badge morning">${morningExams.length} Morning</span>
                        <span class="session-badge afternoon">${afternoonExams.length} Afternoon</span>
                    </div>
                </div>
                <div class="daily-sessions">
                    <div class="daily-session morning">
                        <div class="session-title-small">
                            <i class="fas fa-sun" style="color: #2a87e4ff;"></i>
                            Morning Session (09:00-12:00)
                        </div>
                        <div class="session-exams">
                            ${morningExams.length > 0 ? 
                                morningExams.map(exam => renderDailyExamCard(exam)).join('') :
                                '<div class="no-exams-message"><small>No morning exams</small></div>'
                            }
                        </div>
                    </div>
                    <div class="daily-session afternoon">
                        <div class="session-title-small">
                            <i class="fas fa-cloud-sun" style="color: #f78819ff;"></i>
                            Afternoon Session (14:00-17:00)
                        </div>
                        <div class="session-exams">
                            ${afternoonExams.length > 0 ? 
                                afternoonExams.map(exam => renderDailyExamCard(exam)).join('') :
                                '<div class="no-exams-message"><small>No afternoon exams</small></div>'
                            }
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    return `
        <div class="exam-timetable-card">
            <div class="exam-timetable-header">
                <h3>Exam Schedule - ${timetable.period} (Daily View)</h3>
                <div class="exam-actions">
                    <button class="action-btn" onclick="downloadExamTimetable('${timetable.filename}')" title="Download">
                        <i class="fas fa-download"></i>
                    </button>
                </div>
            </div>
            <div class="daily-view">
                ${dailyHtml}
            </div>
        </div>
    `;
}

function renderDailyExamCard(exam) {
    const formattedDate = formatDateForDisplay(exam.date);
    return `
        <div class="daily-exam-card">
            <div class="daily-exam-header">
                <div class="daily-exam-code">${exam.course_code}</div>
                <div class="daily-exam-meta">
                    <span>${exam.duration}</span>
                    <span>${exam.exam_type}</span>
                </div>
            </div>
            <div class="daily-exam-name">${exam.course_name}</div>
            <div class="daily-exam-details">
                <span><i class="fas fa-building"></i> ${exam.department}</span>
                <span><i class="fas fa-calendar"></i> ${formattedDate}</span>
                <span><i class="fas fa-clock"></i> ${exam.time_slot}</span>
            </div>
        </div>
    `;
}

function renderCompactView(timetable, scheduledExams) {
    // Sort exams by date
    const sortedExams = scheduledExams.sort((a, b) => parseDDMMYYYY(a.date) - parseDDMMYYYY(b.date));
    
    return `
        <div class="exam-timetable-card">
            <div class="exam-timetable-header">
                <h3>Exam Schedule - ${timetable.period} (Compact View)</h3>
                <div class="exam-actions">
                    <button class="action-btn" onclick="downloadExamTimetable('${timetable.filename}')" title="Download">
                        <i class="fas fa-download"></i>
                    </button>
                </div>
            </div>
            <div class="compact-view">
                <table class="compact-table">
                    <thead>
                        <tr>
                            <th>Date</th>
                            <th>Day</th>
                            <th>Session</th>
                            <th>Course Code</th>
                            <th>Course Name</th>
                            <th>Duration</th>
                            <th>Department</th>
                            <th>Type</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${sortedExams.map(exam => {
                            const formattedDate = formatDateForDisplay(exam.date);
                            return `
                            <tr>
                                <td>${formattedDate}</td>
                                <td>${exam.day}</td>
                                <td><span class="session-badge ${exam.session.toLowerCase()}">${exam.session}</span></td>
                                <td><strong>${exam.course_code}</strong></td>
                                <td>${exam.course_name}</td>
                                <td>${exam.duration}</td>
                                <td>${exam.department}</td>
                                <td>${exam.exam_type}</td>
                            </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            </div>
        </div>
    `;
}

function setupViewModeToggle() {
    const toggleButtons = document.querySelectorAll('.toggle-btn');
    
    toggleButtons.forEach(btn => {
        btn.addEventListener('click', function() {
            // Update active button
            toggleButtons.forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            
            // Re-render with new view mode
            renderExamTimetables();
        });
    });
}

// Add event listener for the toggle view button
document.getElementById('toggle-view-mode')?.addEventListener('click', function() {
    const toggleButtons = document.querySelectorAll('.toggle-btn');
    const currentView = document.querySelector('.toggle-btn.active')?.dataset.view;
    const nextView = currentView === 'session' ? 'daily' : 'session';

    const nextButton = document.querySelector(`[data-view="${nextView}"]`);
    if (nextButton) {
        toggleButtons.forEach(b => b.classList.remove('active'));
        nextButton.classList.add('active');
        renderExamTimetables();
    }
});

function enhanceExamTables() {
    document.querySelectorAll('.exam-timetable-table').forEach(table => {
        const rows = table.querySelectorAll('tr');
        
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            cells.forEach(cell => {
                const text = cell.textContent.trim();
                
                if (text === 'Morning') {
                    cell.classList.add('session-morning');
                } else if (text === 'Afternoon') {
                    cell.classList.add('session-afternoon');
                } else if (text === 'No Exam') {
                    cell.classList.add('no-exam');
                }
                
                // Add department colors
                if (text === 'CSE') {
                    cell.classList.add('dept-cse');
                } else if (text === 'DSAI') {
                    cell.classList.add('dept-dsai');
                } else if (text === 'ECE') {
                    cell.classList.add('dept-ece');
                }
            });
        });
    });
}

function downloadExamTimetable(filename) {
    window.open(`/download/${filename}`, '_blank');
    showNotification(`📥 Downloading ${filename}...`, 'info');
}

function downloadAllExamTimetables() {
    if (currentExamTimetables.length === 0) {
        showNotification('❌ No exam timetables available to download', 'error');
        return;
    }
    
    showNotification('📦 Preparing exam timetable download...', 'info');
    
    // Download each exam timetable individually
    currentExamTimetables.forEach(timetable => {
        downloadExamTimetable(timetable.filename);
    });
}

function printExamTimetable(filename) {
    showNotification(`🖨️ Preparing ${filename} for printing...`, 'info');
    // For now, just download - printing can be implemented later
    downloadExamTimetable(filename);
}

// Configuration Management System
const examConfig = {
    defaults: {
        maxExamsPerDay: 2,
        sessionDuration: 180,
        includeWeekends: false,
        departmentConflict: 'moderate',
        preferenceWeight: 'medium',
        sessionBalance: 'strict',
        constraints: {
            departments: ['CSE', 'DSAI', 'ECE', 'Mathematics', 'Physics', 'Humanities'],
            examTypes: ['Theory', 'Lab'],
            rules: ['gapDays', 'sessionLimit', 'preferMorning']
        }
    },
    
    current: null,
    
    init() {
        this.loadConfig();
        this.setupConfigUI();
        this.setupEventListeners();
        this.updateConfigStatus();
    },
    
    loadConfig() {
        const savedConfig = localStorage.getItem('examConfig');
        if (savedConfig) {
            try {
                const parsedConfig = JSON.parse(savedConfig);
                // Ensure all default properties exist
                this.current = this.mergeWithDefaults(parsedConfig);
                console.log('📁 Loaded saved exam configuration:', this.current);
            } catch (e) {
                console.error('❌ Error loading saved config, using defaults:', e);
                this.current = { ...this.defaults };
            }
        } else {
            this.current = { ...this.defaults };
            console.log('📁 Using default exam configuration');
        }
    },
    
    mergeWithDefaults(savedConfig) {
        // Deep merge to ensure all default properties exist
        const merged = { ...this.defaults };
        
        // Merge top-level properties
        Object.keys(savedConfig).forEach(key => {
            if (key === 'constraints') {
                // Deep merge constraints
                merged.constraints = { ...this.defaults.constraints, ...savedConfig.constraints };
            } else {
                merged[key] = savedConfig[key];
            }
        });
        
        return merged;
    },
    
    saveConfig() {
        if (!this.current) {
            console.error('❌ Cannot save: configuration not initialized');
            return;
        }
        localStorage.setItem('examConfig', JSON.stringify(this.current));
        console.log('💾 Saved exam configuration:', this.current);
        this.updateConfigStatus('saved');
        showNotification('✅ Configuration saved successfully!', 'success');
    },
    
    resetConfig() {
        if (confirm('Are you sure you want to reset all configuration settings to defaults?')) {
            this.current = { ...this.defaults };
            localStorage.removeItem('examConfig');
            this.applyConfigToUI();
            this.updateConfigStatus('reset');
            showNotification('🔄 Configuration reset to defaults', 'info');
        }
    },
    
    setupConfigUI() {
        this.setupTabSystem();
        this.applyConfigToUI();
    },
    
    setupTabSystem() {
        const tabButtons = document.querySelectorAll('.tab-btn');
        const tabPanes = document.querySelectorAll('.tab-pane');
        
        tabButtons.forEach(btn => {
            btn.addEventListener('click', () => {
                const tabId = btn.dataset.tab;
                
                // Update active tab button
                tabButtons.forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                
                // Show corresponding tab pane
                tabPanes.forEach(pane => pane.classList.remove('active'));
                const targetPane = document.getElementById(`${tabId}-tab`);
                if (targetPane) {
                    targetPane.classList.add('active');
                }
            });
        });
    },
    
    applyConfigToUI() {
        if (!this.current) {
            console.error('❌ Cannot apply config: configuration not initialized');
            return;
        }
        
        // Basic Settings
        const maxExamsSelect = document.getElementById('max-exams-per-day');
        const sessionDurationSelect = document.getElementById('session-duration');
        if (maxExamsSelect) maxExamsSelect.value = this.current.maxExamsPerDay;
        if (sessionDurationSelect) sessionDurationSelect.value = this.current.sessionDuration;
        
        // Advanced Settings
        const includeWeekendsSelect = document.getElementById('include-weekends');
        const deptConflictSelect = document.getElementById('department-conflict');
        const preferenceSelect = document.getElementById('preference-weight');
        const sessionBalanceSelect = document.getElementById('session-balance');
        
        if (includeWeekendsSelect) includeWeekendsSelect.value = this.current.includeWeekends;
        if (deptConflictSelect) deptConflictSelect.value = this.current.departmentConflict;
        if (preferenceSelect) preferenceSelect.value = this.current.preferenceWeight;
        if (sessionBalanceSelect) sessionBalanceSelect.value = this.current.sessionBalance;
        
        // Constraints
        this.applyConstraintsToUI();
    },
    
    applyConstraintsToUI() {
        if (!this.current || !this.current.constraints) {
            console.error('❌ Cannot apply constraints: constraints not initialized');
            return;
        }
        
        const constraints = this.current.constraints;
        
        // Department constraints
        const setCheckbox = (id, value) => {
            const checkbox = document.getElementById(id);
            if (checkbox) checkbox.checked = value;
        };
        
        setCheckbox('constraint-cse', constraints.departments?.includes('CSE') ?? true);
        setCheckbox('constraint-dsai', constraints.departments?.includes('DSAI') ?? true);
        setCheckbox('constraint-ece', constraints.departments?.includes('ECE') ?? true);
        setCheckbox('constraint-math', constraints.departments?.includes('Mathematics') ?? true);
        setCheckbox('constraint-physics', constraints.departments?.includes('Physics') ?? true);
        setCheckbox('constraint-humanities', constraints.departments?.includes('Humanities') ?? true);
        
        // Exam type constraints
        setCheckbox('constraint-theory', constraints.examTypes?.includes('Theory') ?? true);
        setCheckbox('constraint-lab', constraints.examTypes?.includes('Lab') ?? true);
        
        // Additional rules
        setCheckbox('rule-gap-days', constraints.rules?.includes('gapDays') ?? true);
        setCheckbox('rule-session-limit', constraints.rules?.includes('sessionLimit') ?? true);
        setCheckbox('rule-prefer-morning', constraints.rules?.includes('preferMorning') ?? true);
    },
    
    updateConfigFromUI() {
        if (!this.current) {
            this.current = { ...this.defaults };
        }
        
        // Basic Settings
        const maxExamsSelect = document.getElementById('max-exams-per-day');
        const sessionDurationSelect = document.getElementById('session-duration');
        if (maxExamsSelect) this.current.maxExamsPerDay = parseInt(maxExamsSelect.value);
        if (sessionDurationSelect) this.current.sessionDuration = parseInt(sessionDurationSelect.value);
        
        // Advanced Settings
        const includeWeekendsSelect = document.getElementById('include-weekends');
        const deptConflictSelect = document.getElementById('department-conflict');
        const preferenceSelect = document.getElementById('preference-weight');
        const sessionBalanceSelect = document.getElementById('session-balance');
        
        if (includeWeekendsSelect) this.current.includeWeekends = includeWeekendsSelect.value === 'true';
        if (deptConflictSelect) this.current.departmentConflict = deptConflictSelect.value;
        if (preferenceSelect) this.current.preferenceWeight = preferenceSelect.value;
        if (sessionBalanceSelect) this.current.sessionBalance = sessionBalanceSelect.value;
        
        // Constraints
        this.updateConstraintsFromUI();
        
        console.log('🔄 Updated configuration from UI:', this.current);
    },
    
    updateConstraintsFromUI() {
        if (!this.current.constraints) {
            this.current.constraints = { ...this.defaults.constraints };
        }
        
        const constraints = {
            departments: [],
            examTypes: [],
            rules: []
        };
        
        // Department constraints
        const getCheckboxValue = (id) => {
            const checkbox = document.getElementById(id);
            return checkbox ? checkbox.checked : false;
        };
        
        if (getCheckboxValue('constraint-cse')) constraints.departments.push('CSE');
        if (getCheckboxValue('constraint-dsai')) constraints.departments.push('DSAI');
        if (getCheckboxValue('constraint-ece')) constraints.departments.push('ECE');
        if (getCheckboxValue('constraint-math')) constraints.departments.push('Mathematics');
        if (getCheckboxValue('constraint-physics')) constraints.departments.push('Physics');
        if (getCheckboxValue('constraint-humanities')) constraints.departments.push('Humanities');
        
        // Exam type constraints
        if (getCheckboxValue('constraint-theory')) constraints.examTypes.push('Theory');
        if (getCheckboxValue('constraint-lab')) constraints.examTypes.push('Lab');
        
        // Additional rules
        if (getCheckboxValue('rule-gap-days')) constraints.rules.push('gapDays');
        if (getCheckboxValue('rule-session-limit')) constraints.rules.push('sessionLimit');
        if (getCheckboxValue('rule-prefer-morning')) constraints.rules.push('preferMorning');
        
        this.current.constraints = constraints;
    },
    
    updateConfigStatus(status = '') {
        const statusElement = document.getElementById('config-status');
        if (!statusElement) return;
        
        switch(status) {
            case 'saved':
                statusElement.className = 'config-status saved';
                statusElement.innerHTML = '<i class="fas fa-check-circle"></i><span>Configuration saved and applied</span>';
                break;
            case 'reset':
                statusElement.className = 'config-status';
                statusElement.innerHTML = '<i class="fas fa-info-circle"></i><span>Using default configuration settings</span>';
                break;
            case 'unsaved':
                statusElement.className = 'config-status error';
                statusElement.innerHTML = '<i class="fas fa-exclamation-triangle"></i><span>Unsaved changes - click Save to apply</span>';
                break;
            default:
                const hasSavedConfig = localStorage.getItem('examConfig');
                if (hasSavedConfig) {
                    statusElement.className = 'config-status saved';
                    statusElement.innerHTML = '<i class="fas fa-check-circle"></i><span>Using saved configuration settings</span>';
                } else {
                    statusElement.className = 'config-status';
                    statusElement.innerHTML = '<i class="fas fa-info-circle"></i><span>Using default configuration settings</span>';
                }
        }
    },
    
    setupEventListeners() {
        // Save configuration button
        const saveBtn = document.getElementById('save-config-btn');
        if (saveBtn) {
            saveBtn.addEventListener('click', () => {
                this.updateConfigFromUI();
                this.saveConfig();
            });
        }
        
        // Reset configuration button
        const resetBtn = document.getElementById('reset-config-btn');
        if (resetBtn) {
            resetBtn.addEventListener('click', () => {
                this.resetConfig();
            });
        }
        
        // Auto-save on input changes (optional)
        const configInputs = document.querySelectorAll('#exam-generate-section select, #exam-generate-section input');
        configInputs.forEach(input => {
            input.addEventListener('change', () => {
                this.updateConfigFromUI();
                this.updateConfigStatus('unsaved');
            });
        });
    },
    
    validateConfig() {
        const errors = [];
        
        // Check if configuration is initialized
        if (!this.current) {
            errors.push('Configuration not initialized');
            return errors;
        }
        
        const startDate = document.getElementById('exam-period-start')?.value;
        const endDate = document.getElementById('exam-period-end')?.value;
        
        // Date validation
        if (!startDate || !endDate) {
            errors.push('Please enter both start and end dates');
        } else if (!isValidDate(startDate) || !isValidDate(endDate)) {
            errors.push('Please enter valid dates in DD/MM/YYYY format');
        } else {
            const start = new Date(startDate.split('/').reverse().join('-'));
            const end = new Date(endDate.split('/').reverse().join('-'));
            if (start >= end) {
                errors.push('End date must be after start date');
            }
        }
        
        // Department constraints validation (with safe access)
        const departments = this.current.constraints?.departments;
        if (!departments || departments.length === 0) {
            errors.push('At least one department must be selected');
        }
        
        // Exam type constraints validation (with safe access)
        const examTypes = this.current.constraints?.examTypes;
        if (!examTypes || examTypes.length === 0) {
            errors.push('At least one exam type must be selected');
        }
        
        return errors;
    },
    
    getConfigForAPI() {
        if (!this.current) {
            console.warn('⚠️ No configuration found, using defaults for API');
            return { ...this.defaults };
        }
        return {
            max_exams_per_day: this.current.maxExamsPerDay || this.defaults.maxExamsPerDay,
            session_duration: this.current.sessionDuration || this.defaults.sessionDuration,
            include_weekends: this.current.includeWeekends !== undefined ? this.current.includeWeekends : this.defaults.includeWeekends,
            department_conflict: this.current.departmentConflict || this.defaults.departmentConflict,
            preference_weight: this.current.preferenceWeight || this.defaults.preferenceWeight,
            session_balance: this.current.sessionBalance || this.defaults.sessionBalance,
            constraints: this.current.constraints || this.defaults.constraints
        };
    }
};

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
window.downloadExamTimetable = downloadExamTimetable;
window.printExamTimetable = printExamTimetable;
window.showExamGenerateSection = showExamGenerateSection;