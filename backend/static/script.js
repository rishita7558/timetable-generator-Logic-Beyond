// Global variables
let currentTimetables = [];
let currentView = 'grid';
let currentSemesterFilter = 'all';
let currentSectionFilter = 'all';
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
    
    // New upload-related event listeners
    setupFileUpload();
    
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
    
    // Upload files button
    const uploadFilesBtn = document.getElementById('upload-files-btn');
    if (uploadFilesBtn) {
        uploadFilesBtn.addEventListener('click', showUploadSection);
    }
    
    // Debug button
    window.debugApp = debugApp;
    
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
            if (semesterFilter) {
                semesterFilter.value = semester;
            }
            
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
        
        // Update course database with server data
        if (timetables.length > 0 && timetables[0].course_info) {
            courseDatabase = timetables[0].course_info;
            console.log('📚 Course database updated from server:', courseDatabase);
        }
        
        renderTimetables();
        
        // Show notification if no timetables
        if (timetables.length === 0) {
            console.log("ℹ️ No timetables available");
            showNotification('No timetables found. Upload CSV files and click "Generate All Timetables" to create them.', 'info');
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
        
        if (totalTimetablesEl) totalTimetablesEl.textContent = stats.total_timetables;
        if (totalCoursesEl) totalCoursesEl.textContent = stats.total_courses;
        if (totalFacultyEl) totalFacultyEl.textContent = stats.total_faculty;
        if (totalClassroomsEl) totalClassroomsEl.textContent = stats.total_classrooms;
        
        console.log("📈 Stats loaded:", stats);
    } catch (error) {
        console.error('Error loading stats:', error);
    }
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
        
        // Extract course code
        const courseCode = extractCourseCode(text);
        
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
                cell.title = `${courseCode}: ${courseInfo.name} (${courseInfo.credits} credits) - ${courseInfo.instructor}`;
            }
            
            // Make cell clickable for more info
            cell.style.cursor = 'pointer';
            cell.classList.add('colored-cell');
            
            // Add elective indicator
            if (text.includes('(Elective)')) {
                cell.classList.add('elective-slot');
                cell.title += " - Common Elective Slot (Same for both sections)";
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
    const coursePattern = /[A-Z]{2,3}\d{3}/;
    const match = text.match(coursePattern);
    return match ? match[0] : text;
}

function createEnhancedLegend(semester, section, courses, courseColors, courseInfo, coreCourses = [], electiveCourses = []) {
    if (!courses || courses.length === 0) return '';
    
    // Separate courses into core and elective for better organization
    const coreCourseList = coreCourses.filter(course => courses.includes(course));
    const electiveCourseList = electiveCourses.filter(course => courses.includes(course));
    const otherCourses = courses.filter(course => 
        !coreCourseList.includes(course) && !electiveCourseList.includes(course)
    );
    
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
                    Core Courses
                </div>
                <div class="legend-grid">
        `;
        
        coreCourseList.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color);
        });
        
        legendHtml += `
                </div>
            </div>
        `;
    }
    
    // Elective Courses Section
    if (electiveCourseList.length > 0) {
        legendHtml += `
            <div class="legend-section">
                <div class="legend-section-title elective">
                    <i class="fas fa-clipboard-list"></i>
                    Elective Basket
                </div>
                <div class="legend-grid">
        `;
        
        electiveCourseList.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color);
        });
        
        legendHtml += `
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
                    Other Courses
                </div>
                <div class="legend-grid">
        `;
        
        otherCourses.sort().forEach(courseCode => {
            const info = courseInfo[courseCode];
            const color = courseColors[courseCode] || '#CCCCCC';
            legendHtml += createLegendItem(courseCode, info, color);
        });
        
        legendHtml += `
                </div>
            </div>
        `;
    }
    
    legendHtml += `</div>`;
    return legendHtml;
}

function createLegendItem(courseCode, courseInfo, color) {
    const courseName = courseInfo ? courseInfo.name : 'Unknown Course';
    const credits = courseInfo ? courseInfo.credits : '?';
    const instructor = courseInfo ? courseInfo.instructor : 'Unknown';
    const courseType = courseInfo ? courseInfo.type : 'Core';
    
    return `
        <div class="legend-item ${courseType.toLowerCase()}">
            <div class="legend-color" style="background: ${color};"></div>
            <span class="legend-course-code">${courseCode}</span>
            <span class="legend-course-name">
                ${courseName} (${credits} cr)
                <br><small>${instructor} • ${courseType}</small>
            </span>
        </div>
    `;
}

function enhanceTables() {
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
                    
                    if (timetable && timetable.course_colors) {
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
    });
    
    // Add enhanced hover effects for elective courses
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
    
    // Enhance tables with color coding and legends
    enhanceTables();
    
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

// Export functions for global access
window.downloadTimetable = downloadTimetable;
window.printTimetable = printTimetable;
window.debugApp = debugApp;
window.showUploadSection = showUploadSection;
window.hideUploadSection = hideUploadSection;
window.debugFileMatching = debugFileMatching;