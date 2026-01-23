        let horseData = [];
        let filteredData = [];
        let currentSort = { column: 'lastTrainingDate', order: 'desc' };
        let allHorseDetailData = {};
        let currentHorseDetailData = [];
        let currentHorseDetailSort = { column: 'date', order: 'desc' };
        let currentTypeFilter = 'all';
        let currentHorseRawName = ''; // Store raw name for data lookup

        // View: 'active' or 'historic'
        let currentView = 'active';

        // Multi-sheet variables
        let allSheets = {};
        let sheetNames = [];
        let currentSheetName = 'Default';
        
        // Column visibility settings
        let columnVisibility = {
            date: true, horse: true, type: true, track: true, surface: true, distance: true,
            avgSpeed: true, maxSpeed: true, best1f: true, best2f: true, best3f: true, best4f: true,
            best5f: true, best6f: true, best7f: true, maxHR: true, fastRecovery: true, fastQuality: true,
            fastPercent: true, recovery15: true, quality15: true, hr15Percent: true, maxSL: true,
            slGallop: true, sfGallop: true, slWork: true, sfWork: true, hr2min: true, hr5min: true,
            symmetry: true, regularity: true, bpm120: true, zone5: true, age: true, sex: true,
            temp: true, distanceCol: true, trotHR: true, walkHR: true, notes: true
        };

        // Default column order
        let columnOrder = [
            'date', 'horse', 'type', 'track', 'surface', 'distance',
            'avgSpeed', 'maxSpeed', 'best1f', 'best2f', 'best3f', 'best4f',
            'best5f', 'best6f', 'best7f', 'maxHR', 'fastRecovery', 'fastQuality',
            'fastPercent', 'recovery15', 'quality15', 'hr15Percent', 'maxSL',
            'slGallop', 'sfGallop', 'slWork', 'sfWork', 'hr2min', 'hr5min',
            'symmetry', 'regularity', 'bpm120', 'zone5', 'age', 'sex',
            'temp', 'distanceCol', 'trotHR', 'walkHR', 'notes'
        ];
        
        // LocalStorage functions for multi-sheet functionality
        function loadAllSheets() {
            try {
                const stored = localStorage.getItem('trainingSheets');
                if (stored) {
                    allSheets = JSON.parse(stored);
                    sheetNames = Object.keys(allSheets);
                    if (sheetNames.length > 0) {
                        currentSheetName = sheetNames[0];
                    }
                }
                updateSheetDropdown();
            } catch (error) {
                console.error('Error loading sheets:', error);
            }
        }
        
        function saveAllSheets() {
            try {
                localStorage.setItem('trainingSheets', JSON.stringify(allSheets));
            } catch (error) {
                console.error('Error saving sheets:', error);
            }
        }

        // Sync sheet data to Redis backend
        function syncSheetDataToRedis(sessionId) {
            // Validate inputs
            if (!sessionId) {
                console.warn('Cannot sync sheet data: no session ID available');
                return;
            }

            if (!allSheets || Object.keys(allSheets).length === 0) {
                console.warn('Cannot sync sheet data: no sheet data available');
                return;
            }

            const sheetData = {
                allSheets: allSheets,
                sheetNames: sheetNames,
                currentSheetName: currentSheetName
            };

            fetch(`/api/session/${sessionId}/sheets`, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(sheetData)
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(data => {
                if (data.success) {
                    console.log('Sheet data synced to Redis successfully');
                } else {
                    console.error('Failed to sync sheet data:', data.error);
                    // Don't show user error for sync failures - it's background operation
                }
            })
            .catch(error => {
                console.error('Error syncing sheet data to Redis:', error);
                // Don't show user error for sync failures - it's background operation
            });
        }
        
        function getCurrentSheetData() {
            return allSheets[currentSheetName] || { horseData: [], allHorseDetailData: {} };
        }
        
        function setCurrentSheetData(horseData, allHorseDetailData) {
            allSheets[currentSheetName] = { horseData, allHorseDetailData };
            saveAllSheets();
        }
        
        function loadCurrentSheet() {
            const sheetData = getCurrentSheetData();
            horseData = sheetData.horseData || [];
            allHorseDetailData = sheetData.allHorseDetailData || {};
            filteredData = [...horseData];
        }
        
        function getFastRecoveryColor(value) {
            if (!value || value === '-') return null;
            
            // Check if value contains letters (finish position in race)
            if (/[a-zA-Z]/.test(value)) return null;
            
            const numValue = parseFloat(value);
            if (isNaN(numValue)) return null;
            
            if (numValue >= 140) return '#fdeaea';
            if (numValue >= 125) return '#fff3cd';
            if (numValue >= 119) return '#f9f7e3';
            if (numValue >= 101) return '#d4edda';
            return '#d1ecf1';
        }
        
        function getRecovery15Color(value) {
            if (!value || value === '-') return null;
            
            const numValue = parseFloat(value);
            if (isNaN(numValue)) return null;
            
            if (numValue >= 116) return '#fdeaea';
            if (numValue >= 102) return '#fff3cd';
            if (numValue >= 81) return '#d4edda';
            return '#d1ecf1';
        }
        
        function getBest5FColor(timeStr) {
            if (!timeStr || timeStr === '-' || !isValidTime(timeStr)) return null;
            
            const seconds = timeToSeconds(timeStr);
            
            if (seconds <= 60) return '#d1ecf1';
            if (seconds <= 65) return '#d4edda';
            if (seconds <= 70) return '#f9f7e3';
            if (seconds <= 75) return '#fff3cd';
            return '#fdeaea';
        }
        
        // Column visibility functions
        function loadColumnVisibility() {
            try {
                const saved = localStorage.getItem('arioneoColumnVisibility');
                if (saved) {
                    const parsedVisibility = JSON.parse(saved);
                    columnVisibility = {...columnVisibility, ...parsedVisibility};
                    // Ensure notes column is always visible (new feature)
                    columnVisibility.notes = true;
                    console.log('Loaded column visibility preferences:', Object.keys(parsedVisibility).length, 'columns');
                }
                
                const savedOrder = localStorage.getItem('arioneoColumnOrder');
                if (savedOrder) {
                    const parsedOrder = JSON.parse(savedOrder);
                    if (Array.isArray(parsedOrder) && parsedOrder.length > 0) {
                        columnOrder = parsedOrder;
                        // Ensure notes column is in the order (may be missing from old preferences)
                        if (!columnOrder.includes('notes')) {
                            columnOrder.push('notes');
                        }
                        console.log('Loaded column order preferences:', parsedOrder.length, 'columns');
                    }
                }
                
                // Save current preferences version for future compatibility
                localStorage.setItem('arioneoPreferencesVersion', '1.0');
                
                updateColumnVisibilityUI();
            } catch (error) {
                console.error('Error loading column preferences:', error);
                // Reset to defaults if preferences are corrupted
                localStorage.removeItem('arioneoColumnVisibility');
                localStorage.removeItem('arioneoColumnOrder');
            }
        }
        
        function saveColumnVisibility() {
            localStorage.setItem('arioneoColumnVisibility', JSON.stringify(columnVisibility));
        }
        
        function saveColumnOrder() {
            localStorage.setItem('arioneoColumnOrder', JSON.stringify(columnOrder));
            console.log('Saved column order preferences');
        }
        
        // Save and load user preferences
        function saveUserPreferences() {
            const preferences = {
                columnVisibility: columnVisibility,
                columnOrder: columnOrder,
                currentSort: currentSort,
                lastUpdated: new Date().toISOString()
            };
            localStorage.setItem('arioneoUserPreferences', JSON.stringify(preferences));
            console.log('Saved user preferences');
        }
        
        function loadUserPreferences() {
            try {
                const saved = localStorage.getItem('arioneoUserPreferences');
                if (saved) {
                    const preferences = JSON.parse(saved);
                    if (preferences.columnVisibility) {
                        columnVisibility = {...columnVisibility, ...preferences.columnVisibility};
                        // Ensure notes column is always visible (new feature)
                        columnVisibility.notes = true;
                        // Update the UI to reflect loaded preferences
                        updateColumnVisibilityUI();
                    }
                    if (preferences.columnOrder && Array.isArray(preferences.columnOrder)) {
                        columnOrder = preferences.columnOrder;
                        // Ensure notes column is in the order
                        if (!columnOrder.includes('notes')) {
                            columnOrder.push('notes');
                        }
                        // Update column order UI if it exists
                        if (typeof populateColumnOrderList === 'function') {
                            populateColumnOrderList();
                        }
                    }
                    // Sort is always defaulted to most recent training - don't restore from preferences
                    console.log('Loaded user preferences from:', preferences.lastUpdated);
                }
            } catch (error) {
                console.error('Error loading user preferences:', error);
            }
        }
        
        function updateColumnVisibilityUI() {
            Object.keys(columnVisibility).forEach(col => {
                const checkbox = document.getElementById(`col-${col}`);
                if (checkbox) {
                    checkbox.checked = columnVisibility[col];
                }
            });
        }
        
        function toggleColumnVisibility(column, visible) {
            columnVisibility[column] = visible;
            saveColumnVisibility();
            saveUserPreferences(); // Save comprehensive preferences
            updateTableColumnVisibility();
        }
        
        function updateTableColumnVisibility() {
            const table = document.getElementById('horseDetailTable');
            if (!table) return;
            
            // Rebuild the entire table with proper column order
            buildTableHeader();
            displayHorseDetailData();

            // Coordinate visual updates to prevent race conditions
            setTimeout(() => {
                applyMobileStyles();
                updateScrollbar();
            }, 100); // Single coordinated update
        }
        
        function buildTableHeader() {
            const table = document.getElementById('horseDetailTable');
            const thead = table.querySelector('thead');
            if (!thead) return;
            
            const columnNames = {
                date: 'Date', horse: 'Horse', type: 'Type', track: 'Track', surface: 'Surface', distance: 'Distance',
                avgSpeed: 'Avg Speed', maxSpeed: 'Max Speed', best1f: 'Best 1F', best2f: 'Best 2F', best3f: 'Best 3F', best4f: 'Best 4F',
                best5f: 'Best 5F', best6f: 'Best 6F', best7f: 'Best 7F', maxHR: 'Max HR', fastRecovery: 'Fast Recovery', fastQuality: 'Fast Quality',
                fastPercent: 'Fast %', recovery15: '15 Recovery', quality15: '15 Quality', hr15Percent: 'HR 15%', maxSL: 'Max SL',
                slGallop: 'SL Gallop', sfGallop: 'SF Gallop', slWork: 'SL Work', sfWork: 'SF Work', hr2min: 'HR 2 min', hr5min: 'HR 5 min',
                symmetry: 'Symmetry', regularity: 'Regularity', bpm120: '120bpm', zone5: 'Zone 5', age: 'Age', sex: 'Sex',
                temp: 'Temp', distanceCol: 'Distance (Col)', trotHR: 'Trot HR', walkHR: 'Walk HR', notes: 'Notes'
            };

            const headerHTML = columnOrder.map(col => {
                const display = columnVisibility[col] ? '' : 'style="display: none;"';
                return `<th onclick="sortHorseTable('${col}')" ${display}>${columnNames[col] || col} <span class="sort-indicator">↕</span></th>`;
            }).join('');

            // Add Actions column header
            thead.innerHTML = `<tr>${headerHTML}<th style="width: 60px;">Actions</th></tr>`;
        }
        
        function populateColumnOrderList() {
            const container = document.getElementById('columnOrderList');
            if (!container) return;

            const columnNames = {
                date: 'Date', horse: 'Horse', type: 'Type', track: 'Track', surface: 'Surface', distance: 'Distance',
                avgSpeed: 'Avg Speed', maxSpeed: 'Max Speed', best1f: 'Best 1F', best2f: 'Best 2F', best3f: 'Best 3F', best4f: 'Best 4F',
                best5f: 'Best 5F', best6f: 'Best 6F', best7f: 'Best 7F', maxHR: 'Max HR', fastRecovery: 'Fast Recovery', fastQuality: 'Fast Quality',
                fastPercent: 'Fast %', recovery15: '15 Recovery', quality15: '15 Quality', hr15Percent: 'HR 15%', maxSL: 'Max SL',
                slGallop: 'SL Gallop', sfGallop: 'SF Gallop', slWork: 'SL Work', sfWork: 'SF Work', hr2min: 'HR 2 min', hr5min: 'HR 5 min',
                symmetry: 'Symmetry', regularity: 'Regularity', bpm120: '120bpm', zone5: 'Zone 5', age: 'Age', sex: 'Sex',
                temp: 'Temp', distanceCol: 'Distance (Col)', trotHR: 'Trot HR', walkHR: 'Walk HR', notes: 'Notes'
            };
            
            container.innerHTML = columnOrder.map((col, index) => `
                <div class="column-item" data-column="${col}" data-index="${index}">
                    <input type="checkbox" ${columnVisibility[col] ? 'checked' : ''}>
                    <span class="column-name">${columnNames[col] || col}</span>
                    <button class="move-up" ${index === 0 ? 'disabled' : ''}>▲</button>
                    <button class="move-down" ${index === columnOrder.length - 1 ? 'disabled' : ''}>▼</button>
                </div>
            `).join('');
            
            // Add event listeners for checkboxes and move buttons
            const columnItems = container.querySelectorAll('.column-item');
            columnItems.forEach((item, index) => {
                const checkbox = item.querySelector('input[type="checkbox"]');
                const moveUpBtn = item.querySelector('.move-up');
                const moveDownBtn = item.querySelector('.move-down');
                const column = item.dataset.column;
                
                checkbox.addEventListener('change', function(e) {
                    columnVisibility[column] = e.target.checked;
                    saveColumnVisibility();
                    updateTableColumnVisibility();
                });
                
                moveUpBtn.addEventListener('click', function() {
                    if (index > 0) {
                        moveColumn(index, index - 1);
                    }
                });
                
                moveDownBtn.addEventListener('click', function() {
                    if (index < columnOrder.length - 1) {
                        moveColumn(index, index + 1);
                    }
                });
            });
        }
        
        function moveColumn(fromIndex, toIndex) {
            const movedColumn = columnOrder[fromIndex];
            columnOrder.splice(fromIndex, 1);
            columnOrder.splice(toIndex, 0, movedColumn);
            
            saveColumnOrder();
            saveUserPreferences(); // Save comprehensive preferences
            populateColumnOrderList();
            updateTableColumnVisibility();
        }
        
        console.log("🐎 Horse Racing Analyzer (Arioneo US) loaded successfully!");
        
        // Enhanced custom scrollbar for horizontal scrolling
        function initializeTableNavigation() {
            const tableContainer = document.getElementById('horseDetailTableContainer');
            const customScrollTrack = document.getElementById('customScrollTrack');
            const customScrollThumb = document.getElementById('customScrollThumb');
            
            if (!tableContainer || !customScrollTrack || !customScrollThumb) return;
            
            let isDragging = false;
            let startX = 0;
            let startScrollLeft = 0;
            
            function updateScrollbar() {
                const { scrollLeft, scrollWidth, clientWidth } = tableContainer;
                
                if (scrollWidth <= clientWidth) {
                    customScrollTrack.style.display = 'none';
                    return;
                }
                
                customScrollTrack.style.display = 'block';
                
                const thumbWidth = Math.max(20, (clientWidth / scrollWidth) * customScrollTrack.offsetWidth);
                const thumbPosition = (scrollLeft / (scrollWidth - clientWidth)) * (customScrollTrack.offsetWidth - thumbWidth);
                
                customScrollThumb.style.width = thumbWidth + 'px';
                customScrollThumb.style.left = thumbPosition + 'px';
            }
            
            // Track click to jump to position
            customScrollTrack.addEventListener('click', (e) => {
                if (e.target === customScrollThumb) return;
                
                const rect = customScrollTrack.getBoundingClientRect();
                const clickX = e.clientX - rect.left;
                const trackWidth = customScrollTrack.offsetWidth;
                const thumbWidth = customScrollThumb.offsetWidth;
                
                const newThumbPosition = Math.max(0, Math.min(clickX - thumbWidth / 2, trackWidth - thumbWidth));
                const scrollPercentage = newThumbPosition / (trackWidth - thumbWidth);
                const newScrollLeft = scrollPercentage * (tableContainer.scrollWidth - tableContainer.clientWidth);
                
                tableContainer.scrollTo({
                    left: newScrollLeft,
                    behavior: 'smooth'
                });
            });
            
            // Thumb dragging
            customScrollThumb.addEventListener('mousedown', (e) => {
                isDragging = true;
                startX = e.clientX;
                startScrollLeft = tableContainer.scrollLeft;
                customScrollThumb.style.cursor = 'grabbing';
                e.preventDefault();
            });
            
            document.addEventListener('mousemove', (e) => {
                if (!isDragging) return;
                
                const deltaX = e.clientX - startX;
                const trackWidth = customScrollTrack.offsetWidth;
                const thumbWidth = customScrollThumb.offsetWidth;
                const scrollableTrackWidth = trackWidth - thumbWidth;
                
                const scrollRatio = deltaX / scrollableTrackWidth;
                const maxScroll = tableContainer.scrollWidth - tableContainer.clientWidth;
                const newScrollLeft = Math.max(0, Math.min(startScrollLeft + (scrollRatio * maxScroll), maxScroll));
                
                tableContainer.scrollLeft = newScrollLeft;
            });
            
            document.addEventListener('mouseup', () => {
                if (isDragging) {
                    isDragging = false;
                    customScrollThumb.style.cursor = 'grab';
                }
            });
            
            // Update scrollbar when table scrolls
            tableContainer.addEventListener('scroll', updateScrollbar);
            
            // Update when content changes
            const observer = new MutationObserver(() => {
                setTimeout(updateScrollbar, 50);
            });
            observer.observe(tableContainer, { 
                childList: true, 
                subtree: true 
            });
            
            // Update when window resizes
            window.addEventListener('resize', updateScrollbar);
            
            // Initial update with sufficient delay to ensure table is rendered
            setTimeout(updateScrollbar, 150);
        }
        
        
        // Add file upload event listener with debugging
        const fileInput = document.getElementById('fileInput');
        if (fileInput) {
            console.log('File input found, adding event listener');
            fileInput.addEventListener('change', handleFileUpload);
        } else {
            console.error('File input not found!');
        }

        // Add Arioneo CSV upload event listener
        const arioneoFileInput = document.getElementById('arioneoFileInput');
        if (arioneoFileInput) {
            console.log('Arioneo file input found, adding event listener');
            arioneoFileInput.addEventListener('change', handleArioneoUpload);
        }

        // Initialize multi-sheet functionality
        loadAllSheets();

        // Load horse mapping for filters
        loadHorseFilters();
        
        // Check for latest session on page load
        loadLatestSession();
        document.getElementById('horseFilter').addEventListener('input', filterData);
        document.getElementById('ageFilter').addEventListener('change', filterData);
        document.getElementById('exportCsv').addEventListener('click', exportToCsv);
        document.getElementById('exportAllTraining').addEventListener('click', exportAllTrainingData);

        // Column visibility event listeners
        document.getElementById('toggleColumnVisibility').addEventListener('click', function() {
            document.getElementById('columnVisibilityPanel').style.display = 'block';
            populateColumnOrderList();
        });
        
        document.getElementById('closeColumnPanel').addEventListener('click', function() {
            document.getElementById('columnVisibilityPanel').style.display = 'none';
        });
        
        document.getElementById('selectAllColumns').addEventListener('click', function() {
            Object.keys(columnVisibility).forEach(col => {
                columnVisibility[col] = true;
            });
            saveColumnVisibility();
            populateColumnOrderList();
            updateTableColumnVisibility();
        });
        
        document.getElementById('deselectAllColumns').addEventListener('click', function() {
            Object.keys(columnVisibility).forEach(col => {
                columnVisibility[col] = false;
            });
            saveColumnVisibility();
            populateColumnOrderList();
            updateTableColumnVisibility();
        });
        
        document.getElementById('resetColumns').addEventListener('click', function() {
            Object.keys(columnVisibility).forEach(col => {
                columnVisibility[col] = true;
            });
            // Reset column order to default
            columnOrder = [
                'date', 'horse', 'type', 'track', 'surface', 'distance',
                'avgSpeed', 'maxSpeed', 'best1f', 'best2f', 'best3f', 'best4f',
                'best5f', 'best6f', 'best7f', 'maxHR', 'fastRecovery', 'fastQuality',
                'fastPercent', 'recovery15', 'quality15', 'hr15Percent', 'maxSL',
                'slGallop', 'sfGallop', 'slWork', 'sfWork', 'hr2min', 'hr5min',
                'symmetry', 'regularity', 'bpm120', 'zone5', 'age', 'sex',
                'temp', 'distanceCol', 'trotHR', 'walkHR'
            ];
            saveColumnVisibility();
            saveColumnOrder();
            populateColumnOrderList();
            updateTableColumnVisibility();
        });
        
        // Load user preferences on page load
        loadUserPreferences();
        loadColumnVisibility();
        // Ensure UI is updated after loading preferences
        updateColumnVisibilityUI();

        // Mobile-specific styling and column freezing
        function applyMobileStyles() {
            // Debug logging
            console.log('Window width:', window.innerWidth);
            console.log('Applying mobile styles...');
            
            // Check if we're on mobile (increase threshold to catch more devices)
            if (window.innerWidth <= 1024) {
                console.log('Mobile detected, applying styles...');
                
                // Header styling
                const header = document.querySelector('.header');
                if (header) {
                    header.style.padding = '15px 10px';
                    console.log('Header styled');
                }
                
                const headerH1 = document.querySelector('.header h1');
                if (headerH1) {
                    headerH1.style.fontSize = '1.4em';
                }
                
                const headerP = document.querySelector('.header p');
                if (headerP) {
                    headerP.style.fontSize = '0.9em';
                }
                
                // Controls styling disabled - using pure CSS instead
                // const controls = document.querySelector('.controls');
                // CSS handles all mobile controls styling now
                
                // Control groups styling disabled - using pure CSS instead
                // const controlGroups = document.querySelectorAll('.control-group');
                // CSS handles all control group styling now
                
                // Labels styling disabled - using pure CSS instead
                // const labels = document.querySelectorAll('.control-group label');
                // CSS handles all label styling now
                
                // Select and input styling disabled - using pure CSS instead
                // const selects = document.querySelectorAll('select, input[type="text"]');
                // CSS handles all input styling now
                /*
                selects.forEach(element => {
                    element.style.padding = '6px 8px';
                    element.style.fontSize = '14px';
                    element.style.height = '32px';
                    element.style.lineHeight = '1.2';
                    element.style.marginBottom = '4px';
                    
                    // Make text inputs smaller, but horse filter gets full width
                    if (element.type === 'text' || element.tagName === 'INPUT') {
                        if (element.id === 'horseFilter') {
                            element.style.width = '100%'; // Full width for horse filter
                            element.style.maxWidth = '100%';
                        } else {
                            element.style.width = '100px';
                            element.style.maxWidth = '100px';
                        }
                    }
                });
                
                // Button styling - much more compact (but skip file upload button)
                const buttons = document.querySelectorAll('.export-btn, .back-button');
                buttons.forEach(button => {
                    button.style.padding = '6px 10px';
                    button.style.fontSize = '13px';
                    button.style.minHeight = '30px';
                    button.style.height = '30px';
                    button.style.lineHeight = '1.2';
                    button.style.marginBottom = '4px';
                });
                
                // Don't modify file upload button at all - leave it exactly as original

                // Main page controls layout - re-enable button grouping only
                const mainControls = document.querySelector('#mainView .controls');
                if (mainControls) {
                    console.log('Found main controls, children:', mainControls.children.length);

                    // Find upload and export button containers with better debugging
                    const uploadGroup = Array.from(mainControls.children).find(child =>
                        child.querySelector('.file-label') || child.querySelector('input[type="file"]')
                    );
                    const exportGroup = Array.from(mainControls.children).find(child =>
                        child.querySelector('#exportCsv') || child.querySelector('.export-btn')
                    );

                    console.log('Upload group found:', !!uploadGroup);
                    console.log('Export group found:', !!exportGroup);
                    
                    if (uploadGroup && exportGroup) {
                        // Create a mobile button row using our new CSS class
                        const buttonRow = document.createElement('div');
                        buttonRow.classList.add('mobile-button-row');

                        // Move buttons into the row
                        buttonRow.appendChild(uploadGroup);
                        buttonRow.appendChild(exportGroup);

                        // Insert the button row at the beginning of controls
                        mainControls.insertBefore(buttonRow, mainControls.firstChild);
                    }

                    // Group Training Sheet dropdown with Add New button
                    const sheetGroup = Array.from(mainControls.children).find(child => child.querySelector('#sheetSelector'));
                    const addNewGroup = Array.from(mainControls.children).find(child => child.querySelector('.add-new-btn'));

                    console.log('Sheet group found:', !!sheetGroup);
                    console.log('Add New group found:', !!addNewGroup);

                    if (sheetGroup && addNewGroup) {
                        // Create a mobile button row for sheet controls
                        const sheetRow = document.createElement('div');
                        sheetRow.classList.add('mobile-button-row');

                        // Move elements into the row
                        sheetRow.appendChild(sheetGroup);
                        sheetRow.appendChild(addNewGroup);

                        // Find where to insert (after the upload/export row)
                        const uploadRow = mainControls.querySelector('.mobile-button-row');
                        if (uploadRow) {
                            mainControls.insertBefore(sheetRow, uploadRow.nextSibling);
                        }
                    }

                    // Group Sort by and Filter by Age on the same row
                    const sortGroup = Array.from(mainControls.children).find(child => child.querySelector('#sortBy'));
                    const ageFilterGroup = Array.from(mainControls.children).find(child => child.querySelector('#ageFilter'));

                    if (sortGroup && ageFilterGroup) {
                        // Create a mobile button row for filters
                        const filterRow = document.createElement('div');
                        filterRow.classList.add('mobile-button-row');

                        // Move elements into the row
                        filterRow.appendChild(sortGroup);
                        filterRow.appendChild(ageFilterGroup);

                        // Insert after other controls
                        const lastRow = mainControls.querySelectorAll('.mobile-button-row');
                        if (lastRow.length > 0) {
                            const lastElement = lastRow[lastRow.length - 1];
                            mainControls.insertBefore(filterRow, lastElement.nextSibling);
                        }
                    }
                }
                
                // Horse detail page controls layout - use flex with fixed sizes
                const horseDetailControls = document.querySelector('.horse-detail-view .controls');
                if (horseDetailControls) {
                    horseDetailControls.style.display = 'flex';
                    horseDetailControls.style.flexDirection = 'column';
                    horseDetailControls.style.gap = '8px';
                    
                    // Main button and Age filter on same row
                    const backButtonGroup = Array.from(horseDetailControls.children).find(child => child.querySelector('.back-button'));
                    const ageFilterGroup = Array.from(horseDetailControls.children).find(child => child.querySelector('#horseAgeFilter'));
                    
                    if (backButtonGroup && ageFilterGroup) {
                        // Create a flex row for main button and age filter
                        const firstRow = document.createElement('div');
                        firstRow.style.display = 'flex';
                        firstRow.style.gap = '8px';
                        firstRow.style.alignItems = 'center';
                        
                        // Set fixed widths to prevent stretching
                        backButtonGroup.style.width = '90px';
                        backButtonGroup.style.flex = 'none';
                        ageFilterGroup.style.flex = '1';
                        
                        // Make main button same length as export button and align with filter
                        const mainButton = backButtonGroup.querySelector('.back-button');
                        if (mainButton) {
                            mainButton.style.width = '85px';
                            mainButton.style.padding = '8px 10px';
                            mainButton.style.fontSize = '13px';
                            mainButton.style.minHeight = '32px';
                            mainButton.style.height = '32px';
                            mainButton.style.boxSizing = 'border-box';
                            mainButton.style.marginTop = '2px';
                        }
                        
                        firstRow.appendChild(backButtonGroup);
                        firstRow.appendChild(ageFilterGroup);
                        horseDetailControls.insertBefore(firstRow, horseDetailControls.firstChild);
                    }
                    
                    // Export and show/hide buttons on same row
                    const exportGroup = Array.from(horseDetailControls.children).find(child => child.querySelector('#exportHorseCsv'));
                    const toggleGroup = Array.from(horseDetailControls.children).find(child => child.querySelector('#toggleColumnVisibility'));
                    
                    if (exportGroup && toggleGroup) {
                        const secondRow = document.createElement('div');
                        secondRow.style.display = 'flex';
                        secondRow.style.gap = '8px';
                        secondRow.style.alignItems = 'center';
                        
                        // Set fixed widths to prevent stretching
                        exportGroup.style.width = '90px';
                        exportGroup.style.flex = 'none';
                        toggleGroup.style.width = '120px';
                        toggleGroup.style.flex = 'none';
                        
                        secondRow.appendChild(exportGroup);
                        secondRow.appendChild(toggleGroup);
                        horseDetailControls.appendChild(secondRow);
                    }
                }
                
                // Horse detail header
                const horseHeader = document.querySelector('.horse-detail-header');
                if (horseHeader) {
                    horseHeader.style.padding = '15px 10px';
                }
                
                const horseHeaderH1 = document.querySelector('.horse-detail-header h1');
                if (horseHeaderH1) {
                    horseHeaderH1.style.fontSize = '1.4em';
                    horseHeaderH1.style.marginBottom = '5px';
                }
                
                // Column freezing removed - not working properly
                */
                // END OF DISABLED MOBILE STYLING - CSS handles spacing, JS handles button grouping only

                // Main table text size - increase by 2 points  
                const mainTableCells = document.querySelectorAll('#horseTable th, #horseTable td');
                mainTableCells.forEach(cell => {
                    if (!cell.querySelector('span')) { // Don't change cells with sort indicators
                        cell.style.fontSize = '14px'; // Increased by 2 points from 12px
                    }
                });
                
                // Horse detail table text size
                const horseDetailCells = document.querySelectorAll('.horse-detail-view th, .horse-detail-view td');
                horseDetailCells.forEach(cell => {
                    cell.style.fontSize = '13px';
                });
                
            }
        }
        
        // Position lost numbers based on screen size
        setTimeout(() => {
            const lostNumbers = document.querySelector('.lost-numbers');
            if (lostNumbers) {
                // Remove from current location
                lostNumbers.remove();
                
                // Create new div with proper positioning
                const newLostNumbers = document.createElement('div');
                newLostNumbers.className = 'lost-numbers';
                newLostNumbers.textContent = '4 8 15 16 23 42';
                
                if (window.innerWidth <= 1024) {
                    // Mobile: check if we're on horse detail page
                    const isHorseDetailPage = document.querySelector('.horse-detail-view') && 
                                            document.querySelector('.horse-detail-view').style.display !== 'none';
                    
                    if (!isHorseDetailPage) {
                        // Mobile main page: below table
                        newLostNumbers.style.cssText = `
                            position: relative !important;
                            text-align: right !important;
                            padding: 8px 15px 6px 0 !important;
                            margin: 3px 0 0 0 !important;
                            font-family: 'Courier New', 'IBM Plex Mono', monospace !important;
                            font-size: 11px !important;
                            color: #6a9a7b !important;
                            opacity: 0.75 !important;
                            font-weight: normal !important;
                            letter-spacing: 2px !important;
                            user-select: none !important;
                            pointer-events: none !important;
                            background: transparent !important;
                        `;
                        document.body.appendChild(newLostNumbers);
                    }
                    // If on mobile horse detail page, don't add lost numbers (hide them)
                } else {
                    // Desktop: fixed at bottom right
                    newLostNumbers.style.cssText = `
                        position: fixed !important;
                        bottom: 10px !important;
                        right: 15px !important;
                        font-family: 'Courier New', 'IBM Plex Mono', monospace !important;
                        font-size: 11px !important;
                        color: #6a9a7b !important;
                        opacity: 0.75 !important;
                        font-weight: normal !important;
                        letter-spacing: 2px !important;
                        user-select: none !important;
                        pointer-events: none !important;
                        background: transparent !important;
                        z-index: 1000 !important;
                    `;
                    document.body.appendChild(newLostNumbers);
                }
                
                console.log('Lost numbers positioned for', window.innerWidth <= 1024 ? 'mobile' : 'desktop');
            }
        }, 200);

        // NO FREEZING - removed all sticky functionality
        
        // Remove shadows from first columns (always applied)
        const removeShadowsStyle = document.createElement('style');
        removeShadowsStyle.textContent = `
            table th:first-child,
            table td:first-child {
                box-shadow: none !important;
            }
        `;
        document.head.appendChild(removeShadowsStyle);
        
        // Desktop button alignment improvements (only apply on desktop)
        if (window.innerWidth > 1024) {
            const desktopButtonStyle = document.createElement('style');
            desktopButtonStyle.textContent = `
                .controls {
                    display: flex !important;
                    flex-wrap: wrap !important;
                    justify-content: center !important;
                    align-items: end !important;
                    gap: 20px !important;
                }
                .control-group {
                    display: flex !important;
                    flex-direction: column !important;
                    align-items: center !important;
                    text-align: center !important;
                }
                .control-group label {
                    margin-bottom: 5px !important;
                }
                .control-group select,
                .control-group input {
                    text-align: center !important;
                }
                .export-btn, .back-button {
                    margin-top: auto !important;
                }
            `;
            document.head.appendChild(desktopButtonStyle);
        }
        
        // Mobile-only horse column width adjustment
        if (window.innerWidth <= 1024) {
            const mobileHorseWidthStyle = document.createElement('style');
            mobileHorseWidthStyle.textContent = `
                .horse-name-cell {
                    width: 120px !important;
                    max-width: 120px !important;
                }
                .horse-name-col {
                    width: 120px !important;
                    max-width: 120px !important;
                }
            `;
            document.head.appendChild(mobileHorseWidthStyle);
        }


        // Apply mobile styles on load and resize with delay to ensure DOM is ready
        setTimeout(() => applyMobileStyles(), 500);
        window.addEventListener('resize', applyMobileStyles);
        window.addEventListener('orientationchange', applyMobileStyles);

        function handleFileUpload(event) {
            console.log('File upload triggered!', event);
            const file = event.target.files[0];
            if (!file) {
                console.log('No file selected');
                return;
            }
            
            console.log('File selected:', file.name, 'Size:', file.size, 'Type:', file.type);
            
            // Show upload progress
            const uploadStatus = document.createElement('div');
            uploadStatus.innerHTML = '<p>Uploading and processing file...</p>';
            uploadStatus.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border: 2px solid #3498db; border-radius: 8px; z-index: 9999; box-shadow: 0 4px 8px rgba(0,0,0,0.2);';
            document.body.appendChild(uploadStatus);
            
            const formData = new FormData();
            formData.append('excel', file);
            
            fetch('/api/upload', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                console.log('Upload response status:', response.status);
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(data => {
                document.body.removeChild(uploadStatus);
                console.log('Upload response data:', data);
                
                if (data.success) {
                    // Update page with new data
                    horseData = data.data.horseData;
                    allHorseDetailData = data.data.allHorseDetailData;

                    // Store in current sheet
                    setCurrentSheetData(horseData, allHorseDetailData);

                    // Store and sync session data
                    if (data.sessionId) {
                        localStorage.setItem('currentSessionId', data.sessionId);
                        syncSheetDataToRedis(data.sessionId);
                    }

                    // Calculate Last Work dates for each horse
                    calculateLastWorkDates();

                    updateHorseFilter();
                    updateAgeFilter();

                    // Default sort to most recent training
                    currentSort = { column: 'lastTrainingDate', order: 'desc' };
                    filterData();
                    document.getElementById('exportCsv').disabled = false;
                    document.getElementById('exportAllTraining').disabled = false;

                    // Show share URL
                    const shareUrl = `${window.location.origin}/share/${data.sessionId}`;
                    showShareDialog(shareUrl);

                    console.log('File uploaded successfully. Share URL:', shareUrl);
                } else {
                    console.error('Upload failed:', data.error);
                    alert('Error uploading file: ' + data.error);
                }
            })
            .catch(error => {
                document.body.removeChild(uploadStatus);
                console.error('Error uploading file:', error);
                alert('Error uploading file: ' + error.message + '\n\nPlease check the browser console for more details.');
            });
        }

        function calculateLastWorkDates() {
            // No longer needed - server now provides lastTrainingDate directly
            // Kept as empty function to avoid breaking existing calls
        }

        // ============================================
        // ARIONEO CSV UPLOAD HANDLER
        // ============================================
        function handleArioneoUpload(event) {
            console.log('Arioneo CSV upload triggered!', event);
            const file = event.target.files[0];
            if (!file) {
                console.log('No file selected');
                return;
            }

            console.log('Arioneo file selected:', file.name);

            // Show upload progress
            const uploadStatus = document.createElement('div');
            uploadStatus.innerHTML = '<p>Processing Arioneo data...</p><p style="font-size: 12px; color: #666;">Applying transformations and merging with existing data</p>';
            uploadStatus.style.cssText = 'position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); background: white; padding: 20px; border: 2px solid #27ae60; border-radius: 8px; z-index: 9999; box-shadow: 0 4px 8px rgba(0,0,0,0.2); text-align: center;';
            document.body.appendChild(uploadStatus);

            const formData = new FormData();
            formData.append('csv', file);

            fetch('/api/upload/arioneo', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                return response.json();
            })
            .then(data => {
                document.body.removeChild(uploadStatus);
                console.log('Arioneo upload response:', data);

                if (data.success) {
                    // Update page with new data
                    horseData = data.data.horseData;
                    allHorseDetailData = data.data.allHorseDetailData;

                    // Store in current sheet
                    setCurrentSheetData(horseData, allHorseDetailData);

                    // Store session ID
                    if (data.sessionId) {
                        localStorage.setItem('currentSessionId', data.sessionId);
                    }

                    // Calculate Last Work dates
                    calculateLastWorkDates();

                    // Refresh filters
                    updateHorseFilter();
                    updateAgeFilter();
                    loadHorseFilters();

                    // Default sort to most recent training
                    currentSort = { column: 'lastTrainingDate', order: 'desc' };
                    filterData();
                    document.getElementById('exportCsv').disabled = false;
                    document.getElementById('exportAllTraining').disabled = false;

                    // Show success message
                    alert(`Success!\n\n${data.message}\n\nTotal horses: ${data.totalHorses}\nTotal entries: ${data.totalEntries}`);

                    console.log('Arioneo data processed successfully');
                } else {
                    alert('Error processing file: ' + data.error);
                }

                // Reset the file input
                event.target.value = '';
            })
            .catch(error => {
                document.body.removeChild(uploadStatus);
                console.error('Error uploading Arioneo file:', error);
                alert('Error processing file: ' + error.message);
                event.target.value = '';
            });
        }

        // ============================================
        // CSV UPLOAD MODAL FUNCTIONS
        // ============================================
        function openCsvUploadModal() {
            const modal = document.getElementById('csvUploadModal');
            modal.style.display = 'flex';
            initCsvDropzone();
        }

        function closeCsvUploadModal() {
            const modal = document.getElementById('csvUploadModal');
            modal.style.display = 'none';
            // Clear selected file display
            const selectedFileDiv = document.getElementById('csvSelectedFile');
            selectedFileDiv.innerHTML = '';
            selectedFileDiv.classList.remove('has-file');
        }

        function initCsvDropzone() {
            const dropzone = document.getElementById('csvDropzone');
            const fileInput = document.getElementById('arioneoFileInput');

            // Remove old listeners by cloning
            const newDropzone = dropzone.cloneNode(true);
            dropzone.parentNode.replaceChild(newDropzone, dropzone);

            // Click to select file
            newDropzone.addEventListener('click', () => {
                fileInput.click();
            });

            // Drag events
            newDropzone.addEventListener('dragover', (e) => {
                e.preventDefault();
                newDropzone.classList.add('dragover');
            });

            newDropzone.addEventListener('dragleave', () => {
                newDropzone.classList.remove('dragover');
            });

            newDropzone.addEventListener('drop', (e) => {
                e.preventDefault();
                newDropzone.classList.remove('dragover');

                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    const file = files[0];
                    if (file.name.endsWith('.csv') || file.name.endsWith('.xlsx')) {
                        handleDroppedCsvFile(file);
                    } else {
                        alert('Please drop a CSV or Excel file.');
                    }
                }
            });

            // Handle file input change (from click)
            fileInput.onchange = function(e) {
                if (e.target.files.length > 0) {
                    handleDroppedCsvFile(e.target.files[0]);
                }
            };
        }

        function handleDroppedCsvFile(file) {
            // Show selected file
            const selectedFileDiv = document.getElementById('csvSelectedFile');
            selectedFileDiv.innerHTML = `
                <span class="csv-file-name">📄 ${file.name}</span>
                <button class="csv-file-remove" onclick="clearCsvSelection()">&times;</button>
            `;
            selectedFileDiv.classList.add('has-file');

            // Process the file
            closeCsvUploadModal();

            // Create a fake event for handleArioneoUpload
            const fakeEvent = {
                target: {
                    files: [file],
                    value: file.name
                }
            };
            // Reset value function
            fakeEvent.target.value = '';

            handleArioneoUpload(fakeEvent);
        }

        function clearCsvSelection() {
            const selectedFileDiv = document.getElementById('csvSelectedFile');
            selectedFileDiv.innerHTML = '';
            selectedFileDiv.classList.remove('has-file');
            document.getElementById('arioneoFileInput').value = '';
        }

        // Close modal on outside click
        document.addEventListener('click', (e) => {
            const modal = document.getElementById('csvUploadModal');
            if (e.target === modal) {
                closeCsvUploadModal();
            }
            const noteModal = document.getElementById('addNoteModal');
            if (e.target === noteModal) {
                closeAddNoteModal();
            }
        });

        // ============================================
        // ADD NOTE MODAL FUNCTIONS
        // ============================================
        function openAddNoteModal() {
            const modal = document.getElementById('addNoteModal');
            modal.style.display = 'flex';
            // Set default date to today
            const today = new Date().toISOString().split('T')[0];
            document.getElementById('noteDate').value = today;
            document.getElementById('noteText').value = '';
        }

        function closeAddNoteModal() {
            const modal = document.getElementById('addNoteModal');
            modal.style.display = 'none';
            document.getElementById('noteDate').value = '';
            document.getElementById('noteText').value = '';
        }

        async function submitNote() {
            const dateInput = document.getElementById('noteDate');
            const noteInput = document.getElementById('noteText');

            const dateValue = dateInput.value;
            const noteText = noteInput.value.trim();

            if (!dateValue) {
                alert('Please select a date.');
                return;
            }

            if (!noteText) {
                alert('Please enter a note.');
                return;
            }

            // Format date as MM/DD/YYYY for consistency with other entries
            const dateParts = dateValue.split('-');
            const formattedDate = `${dateParts[1]}/${dateParts[2]}/${dateParts[0]}`;

            const horseName = currentHorseRawName;
            if (!horseName) {
                alert('No horse selected.');
                return;
            }

            try {
                const response = await fetch('/api/notes', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        horseName: horseName,
                        date: formattedDate,
                        note: noteText
                    })
                });

                const data = await response.json();

                if (data.success) {
                    closeAddNoteModal();
                    // Add the note to the current detail data and re-render
                    const noteEntry = {
                        date: formattedDate,
                        horse: horseName,
                        type: 'Note',
                        notes: noteText,
                        isNote: true,
                        // Empty values for other columns
                        track: '-', surface: '-', distance: '-', avgSpeed: '-', maxSpeed: '-',
                        best1f: '-', best2f: '-', best3f: '-', best4f: '-', best5f: '-',
                        best6f: '-', best7f: '-', maxHR: '-', fastRecovery: '-', fastQuality: '-',
                        fastPercent: '-', recovery15: '-', quality15: '-', hr15Percent: '-',
                        maxSL: '-', slGallop: '-', sfGallop: '-', slWork: '-', sfWork: '-',
                        hr2min: '-', hr5min: '-', symmetry: '-', regularity: '-', bpm120: '-',
                        zone5: '-', age: '-', sex: '-', temp: '-', distanceCol: '-',
                        trotHR: '-', walkHR: '-'
                    };

                    // Add to allHorseDetailData for persistence
                    if (!allHorseDetailData[horseName]) {
                        allHorseDetailData[horseName] = [];
                    }
                    allHorseDetailData[horseName].push(noteEntry);

                    // Re-fetch and display the horse data
                    sortHorseTable(currentHorseDetailSort.column);
                } else {
                    alert('Error saving note: ' + (data.error || 'Unknown error'));
                }
            } catch (error) {
                console.error('Error saving note:', error);
                alert('Error saving note: ' + error.message);
            }
        }

        async function deleteNote(horseName, date) {
            try {
                const response = await fetch('/api/notes', {
                    method: 'DELETE',
                    headers: {
                        'Content-Type': 'application/json'
                    },
                    body: JSON.stringify({
                        horseName: horseName,
                        date: date
                    })
                });

                const data = await response.json();

                if (data.success) {
                    // Remove the note from local data
                    if (allHorseDetailData[horseName]) {
                        allHorseDetailData[horseName] = allHorseDetailData[horseName].filter(
                            entry => !(entry.isNote && entry.date === date)
                        );
                    }

                    // Re-render the table
                    sortHorseTable(currentHorseDetailSort.column);
                    return true;
                } else {
                    alert('Error deleting note: ' + (data.error || 'Unknown error'));
                    return false;
                }
            } catch (error) {
                console.error('Error deleting note:', error);
                alert('Error deleting note: ' + error.message);
                return false;
            }
        }

        function showEditNoteModal(encodedHorse, encodedDate, encodedNote) {
            const horseName = decodeURIComponent(atob(encodedHorse));
            const date = decodeURIComponent(atob(encodedDate));
            const noteText = decodeURIComponent(atob(encodedNote));

            let modal = document.getElementById('editNoteModal');
            if (!modal) {
                modal = document.createElement('div');
                modal.id = 'editNoteModal';
                modal.className = 'add-note-modal';
                modal.innerHTML = `
                    <div class="add-note-content">
                        <div class="add-note-header">
                            <h2>Edit Note</h2>
                            <button class="add-note-close" onclick="closeEditNoteModal()">&times;</button>
                        </div>
                        <div class="add-note-body">
                            <div class="add-note-field">
                                <label for="editNoteDate">Date <span class="required">*</span></label>
                                <input type="date" id="editNoteDate" required>
                            </div>
                            <div class="add-note-field">
                                <label for="editNoteText">Note <span class="required">*</span></label>
                                <textarea id="editNoteText" rows="4" placeholder="Enter your note..." required></textarea>
                            </div>
                            <input type="hidden" id="editNoteHorse">
                            <input type="hidden" id="editNoteOriginalDate">
                            <div class="add-note-actions" style="justify-content: space-between;">
                                <button class="add-note-cancel" style="background: #dc3545; color: white; border-color: #dc3545;" onclick="deleteNoteFromModal()">Delete</button>
                                <div style="display: flex; gap: 12px;">
                                    <button class="add-note-cancel" onclick="closeEditNoteModal()">Cancel</button>
                                    <button class="add-note-submit" onclick="saveNoteEdit()">Save Changes</button>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                document.body.appendChild(modal);
            }

            // Convert MM/DD/YYYY to YYYY-MM-DD for date input
            const dateParts = date.split('/');
            const dateValue = dateParts.length === 3 ?
                `${dateParts[2]}-${dateParts[0].padStart(2, '0')}-${dateParts[1].padStart(2, '0')}` : '';

            document.getElementById('editNoteHorse').value = horseName;
            document.getElementById('editNoteOriginalDate').value = date;
            document.getElementById('editNoteDate').value = dateValue;
            document.getElementById('editNoteText').value = noteText;

            modal.style.display = 'flex';

            modal.onclick = function(e) {
                if (e.target === modal) {
                    closeEditNoteModal();
                }
            };
        }

        function closeEditNoteModal() {
            const modal = document.getElementById('editNoteModal');
            if (modal) modal.style.display = 'none';
        }

        async function deleteNoteFromModal() {
            const horseName = document.getElementById('editNoteHorse').value;
            const originalDate = document.getElementById('editNoteOriginalDate').value;

            if (!confirm(`Delete this note?`)) {
                return;
            }

            const success = await deleteNote(horseName, originalDate);
            if (success) {
                closeEditNoteModal();
            }
        }

        async function saveNoteEdit() {
            const horseName = document.getElementById('editNoteHorse').value;
            const originalDate = document.getElementById('editNoteOriginalDate').value;
            const newDateValue = document.getElementById('editNoteDate').value;
            const newNoteText = document.getElementById('editNoteText').value.trim();

            if (!newDateValue) {
                alert('Please select a date.');
                return;
            }

            if (!newNoteText) {
                alert('Please enter a note.');
                return;
            }

            // Convert YYYY-MM-DD to MM/DD/YYYY
            const dateParts = newDateValue.split('-');
            const newDate = `${dateParts[1]}/${dateParts[2]}/${dateParts[0]}`;

            try {
                // Delete old note
                const deleteResponse = await fetch('/api/notes', {
                    method: 'DELETE',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ horseName, date: originalDate })
                });

                // Add updated note
                const addResponse = await fetch('/api/notes', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ horseName, date: newDate, note: newNoteText })
                });

                const addData = await addResponse.json();

                if (addData.success) {
                    // Update local data
                    if (allHorseDetailData[horseName]) {
                        // Remove old note
                        allHorseDetailData[horseName] = allHorseDetailData[horseName].filter(
                            entry => !(entry.isNote && entry.date === originalDate)
                        );
                        // Add updated note
                        allHorseDetailData[horseName].push({
                            date: newDate,
                            horse: horseName,
                            type: 'Note',
                            notes: newNoteText,
                            isNote: true,
                            track: '-', surface: '-', distance: '-', avgSpeed: '-', maxSpeed: '-',
                            best1f: '-', best2f: '-', best3f: '-', best4f: '-', best5f: '-',
                            best6f: '-', best7f: '-', maxHR: '-', fastRecovery: '-', fastQuality: '-',
                            fastPercent: '-', recovery15: '-', quality15: '-', hr15Percent: '-',
                            maxSL: '-', slGallop: '-', sfGallop: '-', slWork: '-', sfWork: '-',
                            hr2min: '-', hr5min: '-', symmetry: '-', regularity: '-', bpm120: '-',
                            zone5: '-', age: '-', sex: '-', temp: '-', distanceCol: '-',
                            trotHR: '-', walkHR: '-'
                        });
                    }

                    closeEditNoteModal();
                    sortHorseTable(currentHorseDetailSort.column);
                } else {
                    alert('Error saving note: ' + (addData.error || 'Unknown error'));
                }
            } catch (error) {
                console.error('Error saving note:', error);
                alert('Error saving note: ' + error.message);
            }
        }

        // ============================================
        // CLEAR DATA FUNCTION
        // ============================================
        async function clearAllData() {
            if (!confirm('Are you sure you want to clear ALL training data?\n\nThis will remove all horse training entries.\nYour horse-owner mappings will be preserved.\n\nYou can upload fresh data after clearing.')) {
                return;
            }

            try {
                const response = await fetch('/api/session/clear', {
                    method: 'DELETE'
                });

                const data = await response.json();

                if (data.success) {
                    // Clear local data
                    horseData = [];
                    allHorseDetailData = {};
                    allSheets = {};
                    sheetNames = [];
                    currentSheetName = null;

                    // Clear display
                    renderTable([]);

                    alert('All training data has been cleared.\n\nYou can now upload fresh Arioneo CSV data.');
                } else {
                    alert('Error clearing data: ' + data.error);
                }
            } catch (error) {
                console.error('Error clearing data:', error);
                alert('Error clearing data: ' + error.message);
            }
        }

        // ============================================
        // HORSE FILTER FUNCTIONS (Owner/Country)
        // ============================================
        async function loadHorseFilters() {
            try {
                const response = await fetch('/api/horses');
                if (!response.ok) return;

                const data = await response.json();

                // Populate owner filter
                const ownerSelect = document.getElementById('ownerFilter');
                if (ownerSelect && data.owners) {
                    ownerSelect.innerHTML = '<option value="">All Owners</option>' +
                        data.owners.map(o => `<option value="${o}">${o}</option>`).join('');
                }

                // Populate country filter
                const countrySelect = document.getElementById('countryFilter');
                if (countrySelect && data.countries) {
                    countrySelect.innerHTML = '<option value="">All Countries</option>' +
                        '<option value="-">- (No Country)</option>' +
                        data.countries.map(c => `<option value="${c}">${c}</option>`).join('');
                }

                console.log('Loaded filters - Owners:', data.owners?.length || 0, 'Countries:', data.countries?.length || 0);
            } catch (error) {
                console.error('Error loading horse filters:', error);
            }
        }

        function applyFilters() {
            filterData();
        }

        // ============================================
        // HORSE MANAGEMENT MODAL
        // ============================================
        function showHorseManagement() {
            // Create modal if it doesn't exist
            let modal = document.getElementById('horseManagementModal');
            if (!modal) {
                modal = document.createElement('div');
                modal.id = 'horseManagementModal';
                modal.className = 'modal';
                modal.innerHTML = `
                    <div class="modal-content" style="max-width: 800px; max-height: 90vh; overflow-y: auto;">
                        <h2 style="margin-top: 0;">Manage Horses</h2>

                        <div style="margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px;">
                            <h4 style="margin-top: 0;">Add/Edit Horse</h4>
                            <div style="display: grid; grid-template-columns: 1fr 1fr 1fr auto auto; gap: 10px; align-items: end;">
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">Horse Name</label>
                                    <input type="text" id="horseNameInput" placeholder="Horse name..." style="width: 100%; padding: 8px;">
                                </div>
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">Owner</label>
                                    <input type="text" id="horseOwnerInput" placeholder="Owner..." style="width: 100%; padding: 8px;">
                                </div>
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">Country</label>
                                    <input type="text" id="horseCountryInput" placeholder="Country..." style="width: 100%; padding: 8px;">
                                </div>
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">Status</label>
                                    <select id="horseStatusInput" style="width: 100%; padding: 8px;">
                                        <option value="active">Active</option>
                                        <option value="historic">Historic</option>
                                    </select>
                                </div>
                                <button onclick="saveHorseMapping()" class="export-btn" style="height: 38px;">Save</button>
                            </div>
                        </div>

                        <div style="margin-bottom: 20px; padding: 15px; background: #e8f4f8; border-radius: 8px; border: 1px solid #3498db;">
                            <h4 style="margin-top: 0;">Rename Horse</h4>
                            <p style="font-size: 12px; color: #666; margin-bottom: 10px;">
                                Change a horse's name (e.g., when an unnamed horse gets named). Training data will transfer to the new name.
                            </p>
                            <div style="display: grid; grid-template-columns: 1fr 1fr auto; gap: 10px; align-items: end;">
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">Current Name</label>
                                    <select id="renameOldName" style="width: 100%; padding: 8px;">
                                        <option value="">-- Select horse --</option>
                                    </select>
                                </div>
                                <div>
                                    <label style="display: block; margin-bottom: 5px; font-weight: bold;">New Name</label>
                                    <input type="text" id="renameNewName" placeholder="Enter new name..." style="width: 100%; padding: 8px;">
                                </div>
                                <button onclick="renameHorse()" class="upload-btn" style="height: 38px;">Rename</button>
                            </div>
                        </div>

                        <div style="margin-bottom: 20px; padding: 15px; background: #f8f9fa; border-radius: 8px;">
                            <h4 style="margin-top: 0;">Bulk Import</h4>
                            <p style="font-size: 12px; color: #666; margin-bottom: 10px;">
                                Upload CSV/Excel with columns: Horse Name, Owner, Country
                            </p>
                            <input type="file" id="horseMappingFile" accept=".csv,.xlsx" style="margin-right: 10px;">
                            <button onclick="importHorseMappings()" class="export-btn">Import</button>
                        </div>

                        <div>
                            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
                                <h4 style="margin: 0;">Current Mappings</h4>
                                <button onclick="showMergeHorsesUI()" class="upload-btn" style="padding: 6px 12px;">Merge Horses</button>
                            </div>
                            <div style="display: flex; gap: 10px; margin-bottom: 10px; align-items: center;">
                                <select id="mappingStatusFilter" onchange="filterHorseMappings()" style="padding: 8px; border-radius: 4px; border: 1px solid #ddd;">
                                    <option value="active">Active Horses</option>
                                    <option value="historic">Historic Horses</option>
                                    <option value="all">All Horses</option>
                                </select>
                                <input type="text" id="mappingSearchInput" placeholder="Search horses..." oninput="filterHorseMappings()" style="flex: 1; padding: 8px; border-radius: 4px; border: 1px solid #ddd;">
                            </div>
                            <div id="mergeHorsesPanel" style="display: none; background: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 15px; border: 2px solid #3498db;">
                                <h5 style="margin: 0 0 10px 0;">Merge Horses (combine names for the same horse)</h5>
                                <p style="font-size: 12px; color: #666; margin-bottom: 10px;">Select the horses to merge, then choose which name to use as the primary display name.</p>
                                <div id="mergeHorsesList" style="max-height: 200px; overflow-y: auto; background: white; border: 1px solid #ddd; border-radius: 4px; padding: 10px; margin-bottom: 10px;"></div>
                                <div style="margin-bottom: 10px;">
                                    <label style="font-weight: bold;">Primary Name (display name):</label>
                                    <select id="primaryHorseSelect" style="width: 100%; padding: 8px; margin-top: 5px;"></select>
                                </div>
                                <div style="display: flex; gap: 10px;">
                                    <button onclick="executeMerge()" class="upload-btn">Merge Selected</button>
                                    <button onclick="hideMergeHorsesUI()" class="cancel-btn">Cancel</button>
                                </div>
                            </div>
                            <div id="horseMappingList" style="max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px;">
                                <table style="width: 100%; border-collapse: collapse;">
                                    <thead>
                                        <tr style="background: #f1f1f1; position: sticky; top: 0;">
                                            <th style="padding: 10px; text-align: left;">Horse</th>
                                            <th style="padding: 10px; text-align: left;">Owner</th>
                                            <th style="padding: 10px; text-align: left;">Country</th>
                                            <th style="padding: 10px; text-align: center;">Status</th>
                                            <th style="padding: 10px; width: 80px;">Actions</th>
                                        </tr>
                                    </thead>
                                    <tbody id="horseMappingTableBody">
                                        <tr><td colspan="5" style="padding: 20px; text-align: center; color: #666;">Loading...</td></tr>
                                    </tbody>
                                </table>
                            </div>
                        </div>

                        <div class="modal-footer" style="margin-top: 20px; text-align: right;">
                            <button onclick="closeHorseManagement()" class="cancel-btn">Close</button>
                        </div>
                    </div>
                `;
                document.body.appendChild(modal);
            }

            modal.style.display = 'flex';
            loadHorseMappingList();

            // Close modal when clicking on background
            modal.onclick = function(e) {
                if (e.target === modal) {
                    closeHorseManagement();
                }
            };
        }

        function closeHorseManagement() {
            const modal = document.getElementById('horseManagementModal');
            if (modal) modal.style.display = 'none';
        }

        async function loadHorseMappingList() {
            try {
                const response = await fetch('/api/horses');
                const data = await response.json();

                const tbody = document.getElementById('horseMappingTableBody');
                if (!tbody) return;

                if (!data.horses || data.horses.length === 0) {
                    tbody.innerHTML = '<tr><td colspan="5" style="padding: 20px; text-align: center; color: #666;">No horses mapped yet</td></tr>';
                    return;
                }

                // Store horses data for merge UI and filtering
                window.horseMappingData = data.horses;

                // Apply current filter
                filterHorseMappings();

                // Populate the rename dropdown with ALL horses
                const renameSelect = document.getElementById('renameOldName');
                if (renameSelect) {
                    renameSelect.innerHTML = '<option value="">-- Select horse --</option>' +
                        data.horses.map(h => {
                            const encodedName = btoa(encodeURIComponent(h.name));
                            const displayName = h.displayName || h.name;
                            return `<option value="${encodedName}">${displayName}</option>`;
                        }).join('');
                }

            } catch (error) {
                console.error('Error loading horse mappings:', error);
            }
        }

        function filterHorseMappings() {
            const tbody = document.getElementById('horseMappingTableBody');
            if (!tbody || !window.horseMappingData) return;

            const statusFilter = document.getElementById('mappingStatusFilter')?.value || 'active';
            const searchTerm = (document.getElementById('mappingSearchInput')?.value || '').toLowerCase().trim();

            // Filter horses based on status and search
            let filteredHorses = window.horseMappingData.filter(h => {
                const isHistoric = h.isHistoric || false;

                // Status filter
                if (statusFilter === 'active' && isHistoric) return false;
                if (statusFilter === 'historic' && !isHistoric) return false;

                // Search filter
                if (searchTerm) {
                    const displayName = (h.displayName || h.name || '').toLowerCase();
                    const owner = (h.owner || '').toLowerCase();
                    const aliases = (h.aliases || []).join(' ').toLowerCase();
                    if (!displayName.includes(searchTerm) && !owner.includes(searchTerm) && !aliases.includes(searchTerm)) {
                        return false;
                    }
                }

                return true;
            });

            if (filteredHorses.length === 0) {
                tbody.innerHTML = '<tr><td colspan="5" style="padding: 20px; text-align: center; color: #666;">No horses found</td></tr>';
                return;
            }

            tbody.innerHTML = filteredHorses.map(h => {
                    const isHistoric = h.isHistoric || false;
                    const statusColor = isHistoric ? '#e74c3c' : '#27ae60';
                    const statusText = isHistoric ? 'Historic' : 'Active';
                    const toggleText = isHistoric ? 'Make Active' : 'Make Historic';
                    const displayName = h.displayName || h.name;
                    // Use base64 encoding to safely pass horse names with special characters
                    const encodedName = btoa(encodeURIComponent(h.name));

                    // Show aliases if any
                    const aliases = h.aliases || [];
                    const aliasDisplay = aliases.length > 0
                        ? `<div style="font-size: 11px; color: #666; margin-top: 2px;">Also: ${aliases.join(', ')}</div>`
                        : '';

                    return `
                    <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">
                            ${displayName}
                            ${aliasDisplay}
                        </td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">${h.owner || '-'}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">${h.country || '-'}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee; text-align: center;">
                            <span style="color: ${statusColor}; font-weight: bold; margin-right: 8px;">${statusText}</span>
                            <button onclick="toggleHorseStatusEncoded('${encodedName}', ${!isHistoric})" style="padding: 2px 6px; cursor: pointer; font-size: 11px;">${toggleText}</button>
                        </td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee; text-align: center;">
                            <button onclick="editHorseMappingEncoded('${encodedName}')" style="padding: 4px 8px; margin-right: 4px; cursor: pointer;">Edit</button>
                            ${aliases.length > 0 ? `<button onclick="showUnmergeUI('${encodedName}')" style="padding: 4px 8px; margin-right: 4px; cursor: pointer; color: #e67e22;" title="Unmerge aliases">⇔</button>` : ''}
                            <button onclick="deleteHorseMappingEncoded('${encodedName}')" style="padding: 4px 8px; cursor: pointer; color: red;">X</button>
                        </td>
                    </tr>`;
                }).join('');
        }

        async function renameHorse() {
            const oldNameSelect = document.getElementById('renameOldName');
            const newNameInput = document.getElementById('renameNewName');

            if (!oldNameSelect.value) {
                alert('Please select a horse to rename');
                return;
            }

            const oldName = decodeURIComponent(atob(oldNameSelect.value));
            const newName = newNameInput.value.trim();

            if (!newName) {
                alert('Please enter a new name');
                return;
            }

            if (oldName.toLowerCase() === newName.toLowerCase()) {
                alert('New name must be different from the current name');
                return;
            }

            try {
                const response = await fetch('/api/horses/rename', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ oldName, newName })
                });

                const data = await response.json();
                if (data.success) {
                    alert(`Renamed "${oldName}" to "${newName}"`);
                    newNameInput.value = '';
                    oldNameSelect.value = '';
                    loadHorseMappingList();
                    loadLatestSession();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error renaming horse:', error);
                alert('Error renaming horse');
            }
        }

        async function saveHorseMapping() {
            const name = document.getElementById('horseNameInput').value.trim();
            const owner = document.getElementById('horseOwnerInput').value.trim();
            const country = document.getElementById('horseCountryInput').value.trim();
            const status = document.getElementById('horseStatusInput')?.value || 'active';
            const isHistoric = status === 'historic';

            if (!name) {
                alert('Horse name is required');
                return;
            }

            try {
                const response = await fetch('/api/horses', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name, owner, country, isHistoric })
                });

                const data = await response.json();
                if (data.success) {
                    // Clear inputs
                    document.getElementById('horseNameInput').value = '';
                    document.getElementById('horseOwnerInput').value = '';
                    document.getElementById('horseCountryInput').value = '';
                    document.getElementById('horseStatusInput').value = 'active';

                    // Refresh list and filters
                    loadHorseMappingList();
                    loadHorseFilters();

                    // Refresh main data display
                    loadLatestSession();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error saving horse mapping:', error);
                alert('Error saving horse mapping');
            }
        }

        // Toggle horse between active and historic
        async function toggleHorseStatus(name, makeHistoric) {
            try {
                // First get the current horse data
                const response = await fetch('/api/horses');
                const data = await response.json();
                const horse = data.horses.find(h => h.name === name);

                if (!horse) {
                    alert('Horse not found');
                    return;
                }

                // Update with new status
                const updateResponse = await fetch('/api/horses', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({
                        name: horse.name,
                        owner: horse.owner,
                        country: horse.country,
                        isHistoric: makeHistoric
                    })
                });

                const updateData = await updateResponse.json();
                if (updateData.success) {
                    loadHorseMappingList();
                    loadLatestSession();
                } else {
                    alert('Error: ' + updateData.error);
                }
            } catch (error) {
                console.error('Error toggling horse status:', error);
                alert('Error updating horse status');
            }
        }

        function editHorseMapping(name) {
            // Pre-fill the form with existing data
            fetch('/api/horses')
                .then(r => r.json())
                .then(data => {
                    const horse = data.horses.find(h => h.name === name);
                    if (horse) {
                        document.getElementById('horseNameInput').value = horse.name;
                        document.getElementById('horseOwnerInput').value = horse.owner || '';
                        document.getElementById('horseCountryInput').value = horse.country || '';
                        document.getElementById('horseStatusInput').value = horse.isHistoric ? 'historic' : 'active';
                    }
                });
        }

        async function deleteHorseMapping(name) {
            if (!confirm(`Delete mapping for "${name}"?`)) return;

            try {
                const response = await fetch(`/api/horses/${encodeURIComponent(name)}`, {
                    method: 'DELETE'
                });

                const data = await response.json();
                if (data.success) {
                    loadHorseMappingList();
                    loadHorseFilters();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error deleting horse mapping:', error);
            }
        }

        // Wrapper functions that decode base64-encoded horse names
        function toggleHorseStatusEncoded(encodedName, makeHistoric) {
            const name = decodeURIComponent(atob(encodedName));
            toggleHorseStatus(name, makeHistoric);
        }

        function editHorseMappingEncoded(encodedName) {
            const name = decodeURIComponent(atob(encodedName));
            editHorseMapping(name);
        }

        function deleteHorseMappingEncoded(encodedName) {
            const name = decodeURIComponent(atob(encodedName));
            deleteHorseMapping(name);
        }

        // ============================================
        // HORSE MERGE FUNCTIONS
        // ============================================

        function showMergeHorsesUI() {
            const panel = document.getElementById('mergeHorsesPanel');
            const list = document.getElementById('mergeHorsesList');
            const select = document.getElementById('primaryHorseSelect');

            if (!window.horseMappingData || window.horseMappingData.length === 0) {
                alert('No horses available to merge');
                return;
            }

            // Build checkboxes for each horse
            list.innerHTML = window.horseMappingData.map(h => {
                const encodedName = btoa(encodeURIComponent(h.name));
                const displayName = h.displayName || h.name;
                const aliasInfo = h.aliases && h.aliases.length > 0
                    ? ` <span style="color: #666; font-size: 11px;">(has ${h.aliases.length} alias${h.aliases.length > 1 ? 'es' : ''})</span>`
                    : '';
                return `
                    <label style="display: block; padding: 5px; cursor: pointer; border-bottom: 1px solid #eee;">
                        <input type="checkbox" class="merge-horse-checkbox" value="${encodedName}" onchange="updatePrimarySelect()">
                        ${displayName}${aliasInfo}
                    </label>
                `;
            }).join('');

            select.innerHTML = '<option value="">-- Select horses first --</option>';
            panel.style.display = 'block';
        }

        function hideMergeHorsesUI() {
            document.getElementById('mergeHorsesPanel').style.display = 'none';
        }

        function updatePrimarySelect() {
            const checkboxes = document.querySelectorAll('.merge-horse-checkbox:checked');
            const select = document.getElementById('primaryHorseSelect');

            if (checkboxes.length < 2) {
                select.innerHTML = '<option value="">-- Select at least 2 horses --</option>';
                return;
            }

            select.innerHTML = Array.from(checkboxes).map(cb => {
                const name = decodeURIComponent(atob(cb.value));
                const horse = window.horseMappingData.find(h => h.name === name);
                const displayName = horse ? (horse.displayName || horse.name) : name;
                return `<option value="${cb.value}">${displayName}</option>`;
            }).join('');
        }

        async function executeMerge() {
            const checkboxes = document.querySelectorAll('.merge-horse-checkbox:checked');
            const primarySelect = document.getElementById('primaryHorseSelect');

            if (checkboxes.length < 2) {
                alert('Please select at least 2 horses to merge');
                return;
            }

            if (!primarySelect.value) {
                alert('Please select a primary name');
                return;
            }

            const primaryName = decodeURIComponent(atob(primarySelect.value));
            const aliasNames = Array.from(checkboxes)
                .map(cb => decodeURIComponent(atob(cb.value)))
                .filter(name => name !== primaryName);

            if (aliasNames.length === 0) {
                alert('No aliases to merge');
                return;
            }

            try {
                const response = await fetch('/api/horses/merge', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ primaryName, aliasNames })
                });

                const data = await response.json();
                if (data.success) {
                    alert(`Merged successfully! "${primaryName}" now includes: ${aliasNames.join(', ')}`);
                    hideMergeHorsesUI();
                    loadHorseMappingList();
                    loadLatestSession(); // Refresh data to apply merge
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error merging horses:', error);
                alert('Error merging horses');
            }
        }

        function showUnmergeUI(encodedName) {
            const name = decodeURIComponent(atob(encodedName));
            const horse = window.horseMappingData.find(h => h.name === name);

            if (!horse || !horse.aliases || horse.aliases.length === 0) {
                alert('This horse has no aliases to unmerge');
                return;
            }

            const aliasToRemove = prompt(
                `"${horse.displayName || horse.name}" has these aliases:\n\n${horse.aliases.join('\n')}\n\nEnter the alias name to remove:`
            );

            if (!aliasToRemove) return;

            if (!horse.aliases.includes(aliasToRemove)) {
                alert('Alias not found. Please enter the exact alias name.');
                return;
            }

            executeUnmerge(name, aliasToRemove);
        }

        async function executeUnmerge(primaryName, aliasName) {
            try {
                const response = await fetch('/api/horses/unmerge', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ primaryName, aliasName })
                });

                const data = await response.json();
                if (data.success) {
                    alert(`Removed "${aliasName}" from "${primaryName}"`);
                    loadHorseMappingList();
                    loadLatestSession();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error unmerging horse:', error);
                alert('Error unmerging horse');
            }
        }

        async function importHorseMappings() {
            const fileInput = document.getElementById('horseMappingFile');
            if (!fileInput.files || !fileInput.files[0]) {
                alert('Please select a file');
                return;
            }

            const formData = new FormData();
            formData.append('csv', fileInput.files[0]);

            try {
                const response = await fetch('/api/horses/import', {
                    method: 'POST',
                    body: formData
                });

                const data = await response.json();
                if (data.success) {
                    alert(data.message);
                    fileInput.value = '';
                    loadHorseMappingList();
                    loadHorseFilters();
                    loadLatestSession();
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error importing horse mappings:', error);
                alert('Error importing file');
            }
        }

        // ============================================
        // EDIT TRAINING ENTRY
        // ============================================
        function showEditTrainingModal(horse, date, currentData) {
            let modal = document.getElementById('editTrainingModal');
            if (!modal) {
                modal = document.createElement('div');
                modal.id = 'editTrainingModal';
                modal.className = 'modal';
                modal.innerHTML = `
                    <div class="modal-content" style="max-width: 500px;">
                        <h2 style="margin-top: 0;">Edit Training Entry</h2>
                        <p id="editTrainingInfo" style="color: #666; margin-bottom: 20px;"></p>

                        <div class="form-group">
                            <label>Type</label>
                            <select id="editType" style="width: 100%; padding: 8px;">
                                <option value="">-- Select --</option>
                                <option value="Work">Work</option>
                                <option value="Work - G">Work - G (Gate)</option>
                                <option value="Race">Race</option>
                                <option value="">Gallop/Other</option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label>Track</label>
                            <input type="text" id="editTrack" style="width: 100%; padding: 8px;">
                        </div>

                        <div class="form-group">
                            <label>Surface</label>
                            <select id="editSurface" style="width: 100%; padding: 8px;">
                                <option value="">-- Select --</option>
                                <option value="D">Dirt</option>
                                <option value="T">Turf</option>
                                <option value="AWT">All Weather</option>
                                <option value="Sand">Sand</option>
                                <option value="N/A">N/A</option>
                            </select>
                        </div>

                        <div class="form-group">
                            <label>Notes</label>
                            <textarea id="editNotes" rows="3" style="width: 100%; padding: 8px;" placeholder="Add notes about this training..."></textarea>
                        </div>

                        <input type="hidden" id="editHorse">
                        <input type="hidden" id="editDate">

                        <div class="modal-footer" style="display: flex; justify-content: space-between;">
                            <button onclick="deleteTrainingEntry()" class="cancel-btn" style="background: #dc3545; color: white;">Delete</button>
                            <div>
                                <button onclick="closeEditTrainingModal()" class="cancel-btn">Cancel</button>
                                <button onclick="saveTrainingEdit()" class="upload-btn">Save Changes</button>
                            </div>
                        </div>
                    </div>
                `;
                document.body.appendChild(modal);
            }

            // Populate fields
            document.getElementById('editTrainingInfo').textContent = `${horse} - ${date}`;
            document.getElementById('editHorse').value = horse;
            document.getElementById('editDate').value = date;
            document.getElementById('editType').value = currentData.type || '';
            document.getElementById('editTrack').value = currentData.track || '';
            document.getElementById('editSurface').value = currentData.surface || '';
            document.getElementById('editNotes').value = currentData.notes || '';

            modal.style.display = 'flex';

            // Close modal when clicking on background
            modal.onclick = function(e) {
                if (e.target === modal) {
                    closeEditTrainingModal();
                }
            };
        }

        function closeEditTrainingModal() {
            const modal = document.getElementById('editTrainingModal');
            if (modal) modal.style.display = 'none';
        }

        // Wrapper to decode base64-encoded horse name and date
        function showEditTrainingModalEncoded(encodedHorse, encodedDate, currentData) {
            const horse = decodeURIComponent(atob(encodedHorse));
            const date = decodeURIComponent(atob(encodedDate));
            showEditTrainingModal(horse, date, currentData);
        }

        async function saveTrainingEdit() {
            const horse = document.getElementById('editHorse').value;
            const date = document.getElementById('editDate').value;
            const type = document.getElementById('editType').value;
            const track = document.getElementById('editTrack').value;
            const surface = document.getElementById('editSurface').value;
            const notes = document.getElementById('editNotes').value;

            try {
                const response = await fetch('/api/training/edit', {
                    method: 'PUT',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ horse, date, type, track, surface, notes })
                });

                const data = await response.json();
                if (data.success) {
                    closeEditTrainingModal();

                    // Refresh the data
                    loadLatestSession();

                    alert('Training entry updated');
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error saving training edit:', error);
                alert('Error saving changes');
            }
        }

        async function deleteTrainingEntry() {
            const horse = document.getElementById('editHorse').value;
            const date = document.getElementById('editDate').value;

            if (!confirm(`Are you sure you want to delete this training entry for ${horse} on ${date}?`)) {
                return;
            }

            try {
                const response = await fetch('/api/training/delete', {
                    method: 'DELETE',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ horse, date })
                });

                const data = await response.json();
                if (data.success) {
                    closeEditTrainingModal();
                    loadLatestSession();
                    alert('Training entry deleted');
                } else {
                    alert('Error: ' + data.error);
                }
            } catch (error) {
                console.error('Error deleting training entry:', error);
                alert('Error deleting entry');
            }
        }

        function parseCSV(csvText) {
            const lines = csvText.split('\n');
            const result = [];
            
            for (let i = 0; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line) {
                    const row = [];
                    let current = '';
                    let inQuotes = false;
                    
                    for (let j = 0; j < line.length; j++) {
                        const char = line[j];
                        if (char === '"') {
                            inQuotes = !inQuotes;
                        } else if (char === ',' && !inQuotes) {
                            row.push(current.trim());
                            current = '';
                        } else {
                            current += char;
                        }
                    }
                    row.push(current.trim());
                    result.push(row);
                }
            }
            
            return result;
        }

        function isValidTime(timeStr) {
            if (!timeStr) return false;
            const str = timeStr.toString().trim();
            if (str === '-' || str === '' || str === 'NaN') return false;
            return /^\d{2}:\d{2}\.\d{2}$/.test(str);
        }

        function timeToSeconds(timeStr) {
            const parts = timeStr.toString().trim().split(':');
            const minutes = parseInt(parts[0]);
            const secondsParts = parts[1].split('.');
            const seconds = parseInt(secondsParts[0]);
            const hundredths = parseInt(secondsParts[1]);
            
            return minutes * 60 + seconds + hundredths / 100;
        }

        function updateHorseFilter() {
            const horseList = document.getElementById('horseList');
            const horses = horseData.map(horse => horse.name).sort();
            
            horseList.innerHTML = '';
            horses.forEach(horseName => {
                const option = document.createElement('option');
                option.value = horseName;
                horseList.appendChild(option);
            });
        }

        function updateAgeFilter() {
            const ageFilter = document.getElementById('ageFilter');
            const ages = [...new Set(horseData.map(horse => horse.age).filter(age => age !== null))].sort((a, b) => a - b);
            
            ageFilter.innerHTML = '<option value="">All Ages</option>';
            ages.forEach(age => {
                const option = document.createElement('option');
                option.value = age;
                option.textContent = age;
                ageFilter.appendChild(option);
            });
        }

        function filterData() {
            const horseFilter = document.getElementById('horseFilter').value.toLowerCase().trim();
            const ageFilter = document.getElementById('ageFilter').value;
            const ownerFilter = document.getElementById('ownerFilter')?.value || '';
            const countryFilter = document.getElementById('countryFilter')?.value || '';

            filteredData = horseData.filter(horse => {
                // Filter by active/historic view
                const isHistoric = horse.isHistoric || false;
                if (currentView === 'active' && isHistoric) {
                    return false;
                }
                if (currentView === 'historic' && !isHistoric) {
                    return false;
                }

                // Search by both name and displayName
                const searchName = (horse.displayName || horse.name).toLowerCase();
                if (horseFilter && !searchName.includes(horseFilter) && !horse.name.toLowerCase().includes(horseFilter)) {
                    return false;
                }
                if (ageFilter && horse.age !== parseInt(ageFilter)) {
                    return false;
                }
                if (ownerFilter && horse.owner !== ownerFilter) {
                    return false;
                }
                if (countryFilter) {
                    if (countryFilter === '-') {
                        // Filter for horses with no country
                        if (horse.country && horse.country.trim() !== '') {
                            return false;
                        }
                    } else if (horse.country !== countryFilter) {
                        return false;
                    }
                }
                return true;
            });

            sortData();
        }

        // Mobile menu toggles
        document.addEventListener('DOMContentLoaded', function() {
            // Main page toggle
            const menuBtn = document.getElementById('mobileMenuBtn');
            const controls = document.getElementById('mainControls');

            if (menuBtn && controls) {
                menuBtn.addEventListener('click', function() {
                    controls.classList.toggle('expanded');
                    menuBtn.textContent = controls.classList.contains('expanded') ? '✕' : '⋯';
                });
            }

            // Detail page toggle
            const detailMenuBtn = document.getElementById('detailMenuBtn');
            const detailControls = document.getElementById('detailControls');

            if (detailMenuBtn && detailControls) {
                detailMenuBtn.addEventListener('click', function() {
                    detailControls.classList.toggle('expanded');
                    detailMenuBtn.textContent = detailControls.classList.contains('expanded') ? '✕' : '⋯';
                });
            }
        });

        function switchView() {
            const viewSelector = document.getElementById('viewSelector');
            currentView = viewSelector.value;
            filterData();
        }

        function sortTable(column) {
            if (currentSort.column === column) {
                currentSort.order = currentSort.order === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort.column = column;
                // Default to 'desc' for lastTrainingDate column (most recent first), 'asc' for others
                currentSort.order = column === 'lastTrainingDate' ? 'desc' : 'asc';
            }

            sortData();
        }

        function sortData() {
            filteredData.sort((a, b) => {
                let aVal = a[currentSort.column];
                let bVal = b[currentSort.column];
                
                if (currentSort.column === 'best1f' || currentSort.column === 'best5f') {
                    aVal = aVal ? timeToSeconds(aVal) : Infinity;
                    bVal = bVal ? timeToSeconds(bVal) : Infinity;
                } else if (currentSort.column === 'lastTrainingDate') {
                    aVal = (a.lastTrainingDate || a.lastWorkDate) ? new Date(a.lastTrainingDate || a.lastWorkDate) : new Date(0);
                    bVal = (b.lastTrainingDate || b.lastWorkDate) ? new Date(b.lastTrainingDate || b.lastWorkDate) : new Date(0);
                } else if (currentSort.column === 'age' || currentSort.column === 'fastRecovery' || currentSort.column === 'recovery15min') {
                    aVal = parseFloat(aVal) || 0;
                    bVal = parseFloat(bVal) || 0;
                } else {
                    aVal = aVal || '';
                    bVal = bVal || '';
                }
                
                if (currentSort.order === 'asc') {
                    return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
                } else {
                    return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
                }
            });
            
            displayData();
        }

        function displayData() {
            const tbody = document.getElementById('horseTableBody');

            if (filteredData.length === 0) {
                const viewType = currentView === 'historic' ? 'historic' : 'active';
                tbody.innerHTML = `<tr><td colspan="9" class="no-data">No ${viewType} horses match the current filters.</td></tr>`;
                return;
            }

            tbody.innerHTML = filteredData.map((horse, index) => {
                const best5fStyle = horse.best5fColor ? `style="background-color: ${horse.best5fColor}; color: #000;"` : '';
                const fastRecoveryStyle = horse.fastRecoveryColor ? `style="background-color: ${horse.fastRecoveryColor}; color: #000;"` : '';
                const recovery15Style = horse.recovery15Color ? `style="background-color: ${horse.recovery15Color}; color: #000;"` : '';
                const displayName = horse.displayName || horse.name;
                return `
                    <tr class="clickable-row" data-index="${index}">
                        <td class="horse-name-cell"><span style="color: inherit; font-weight: bold;">${displayName}</span></td>
                        <td>${horse.owner || '-'}</td>
                        <td>${horse.country || '-'}</td>
                        <td class="last-work-cell">${horse.lastTrainingDate || horse.lastWorkDate || '-'}</td>
                        <td class="age-cell age-col">${horse.age || '-'}</td>
                        <td class="time-cell">${horse.best1f || '-'}</td>
                        <td class="time-cell best5f-col" ${best5fStyle}>${horse.best5f || '-'}</td>
                        <td class="recovery-cell fast-recovery-col" ${fastRecoveryStyle}>${horse.fastRecovery || '-'}</td>
                        <td class="recovery15-cell" ${recovery15Style}>${horse.recovery15min || '-'}</td>
                    </tr>
                `;
            }).join('');

            // Add click handlers using event delegation - use index to get horse directly
            tbody.querySelectorAll('.clickable-row').forEach(row => {
                row.onclick = function() {
                    const index = parseInt(this.getAttribute('data-index'));
                    const horse = filteredData[index];
                    if (horse) {
                        console.log('Clicked row - index:', index, 'horse.name:', horse.name, 'displayName:', horse.displayName);
                        showHorseDetail(horse.name, horse.displayName);
                    } else {
                        console.error('No horse found at index:', index, 'filteredData length:', filteredData.length);
                    }
                };
            });
        }

        function exportToCsv() {
            if (filteredData.length === 0) {
                alert('No data to export');
                return;
            }

            // Create worksheet data with headers
            const headers = ['Horse Name', 'Owner', 'Country', 'Last Training', 'Age', '1F', '5F', 'Fast', '15 min'];
            const wsData = [headers];

            filteredData.forEach(horse => {
                const displayName = horse.displayName || horse.name;
                wsData.push([
                    displayName,
                    horse.owner || '',
                    horse.country || '',
                    horse.lastTrainingDate || horse.lastWorkDate || '',
                    horse.age || '',
                    horse.best1f || '',
                    horse.best5f || '',
                    horse.fastRecovery || '',
                    horse.recovery15min || ''
                ]);
            });

            // Create workbook and worksheet
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(wsData);

            // Set column widths
            ws['!cols'] = [
                { wch: 25 }, // Horse Name
                { wch: 20 }, // Owner
                { wch: 12 }, // Country
                { wch: 14 }, // Last Training
                { wch: 6 },  // Age
                { wch: 10 }, // 1F
                { wch: 10 }, // 5F
                { wch: 8 },  // Fast
                { wch: 8 }   // 15 min
            ];

            // Force time columns to be text format to preserve display
            const timeColumns = [5, 6, 7, 8]; // 1F, 5F, Fast, 15 min (0-indexed)
            for (let row = 1; row <= filteredData.length; row++) {
                timeColumns.forEach(col => {
                    const cellRef = XLSX.utils.encode_cell({ r: row, c: col });
                    if (ws[cellRef] && ws[cellRef].v) {
                        ws[cellRef].t = 's'; // Force text type
                        ws[cellRef].z = '@'; // Text format
                    }
                });
            }

            XLSX.utils.book_append_sheet(wb, ws, 'Training Summary');

            // Generate filename with date
            const date = new Date().toISOString().split('T')[0];
            const filename = `horse_training_summary_${date}.xlsx`;

            // Download
            XLSX.writeFile(wb, filename);
        }

        // Export all training data for all horses (active and historic) with multi-sheet Excel
        async function exportAllTrainingData() {
            // Filter out invalid horse names (like default sheet names)
            const invalidNames = ['worksheet', 'sheet1', 'sheet2', 'sheet3', 'data', 'default'];
            const allHorses = Object.keys(allHorseDetailData).filter(name => {
                const lowerName = name.toLowerCase().trim();
                return lowerName && !invalidNames.includes(lowerName) && !lowerName.startsWith('sheet');
            });

            if (allHorses.length === 0) {
                alert('No training data to export');
                return;
            }

            // Show loading indicator
            const btn = document.getElementById('exportAllTraining');
            const originalText = btn.textContent;
            btn.textContent = 'Exporting...';
            btn.disabled = true;

            try {
                const workbook = new ExcelJS.Workbook();

                // Define headers for training data (Owner and Country added after Horse)
                const headers = [
                    'Date', 'Horse', 'Owner', 'Country', 'Type', 'Track', 'Surface', 'Distance', 'Avg Speed', 'Max Speed',
                    'Best 1F', 'Best 2F', 'Best 3F', 'Best 4F', 'Best 5F', 'Best 6F', 'Best 7F',
                    'Max HR', 'Fast Recovery', 'Fast Quality', 'Fast %', '15 Recovery', '15 Quality',
                    'HR 15%', 'Max SL', 'SL Gallop', 'SF Gallop', 'SL Work', 'SF Work', 'HR 2 min',
                    'HR 5 min', 'Symmetry', 'Regularity', '120bpm', 'Zone 5', 'Age', 'Sex', 'Temp',
                    'Distance (Col)', 'Trot HR', 'Walk HR', 'Notes'
                ];

                // Column indices (1-based for ExcelJS) - adjusted for Owner and Country columns
                const BEST5F_COL = 15;
                const FAST_RECOVERY_COL = 19;
                const RECOVERY15_COL = 22;

                // Build owner/country lookup from horseData
                const horseLookup = {};
                horseData.forEach(h => {
                    horseLookup[h.name] = { owner: h.owner || '', country: h.country || '' };
                    if (h.displayName && h.displayName !== h.name) {
                        horseLookup[h.displayName] = { owner: h.owner || '', country: h.country || '' };
                    }
                });

                // Helper function to add styled header row
                function addHeaderRow(worksheet) {
                    const headerRow = worksheet.addRow(headers);
                    headerRow.font = { bold: true };
                    headerRow.eachCell(cell => {
                        cell.fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: 'FFE0E0E0' },
                            bgColor: { argb: 'FFE0E0E0' }
                        };
                    });
                }

                // Helper function to parse date string to Date object
                function parseDate(dateStr) {
                    if (!dateStr || dateStr === '-' || dateStr === '') return null;

                    // Try parsing various date formats
                    let date = new Date(dateStr);

                    // If invalid, try MM/DD/YYYY format
                    if (isNaN(date.getTime())) {
                        const parts = dateStr.split('/');
                        if (parts.length === 3) {
                            // Assume MM/DD/YYYY
                            date = new Date(parts[2], parts[0] - 1, parts[1]);
                        }
                    }

                    return isNaN(date.getTime()) ? null : date;
                }

                // Helper function to add a data row with color coding
                function addDataRow(worksheet, rowData) {
                    // Parse date to proper Date object for Excel
                    const dateValue = parseDate(rowData.date);

                    const row = worksheet.addRow([
                        dateValue || rowData.date || '', rowData.horse || '', rowData.owner || '', rowData.country || '',
                        rowData.type || '', rowData.track || '', rowData.surface || '', rowData.distance || '',
                        rowData.avgSpeed || '', rowData.maxSpeed || '', rowData.best1f || '', rowData.best2f || '',
                        rowData.best3f || '', rowData.best4f || '', rowData.best5f || '', rowData.best6f || '',
                        rowData.best7f || '', rowData.maxHR || '', rowData.fastRecovery || '', rowData.fastQuality || '',
                        rowData.fastPercent || '', rowData.recovery15 || '', rowData.quality15 || '',
                        rowData.hr15Percent || '', rowData.maxSL || '', rowData.slGallop || '', rowData.sfGallop || '',
                        rowData.slWork || '', rowData.sfWork || '', rowData.hr2min || '', rowData.hr5min || '',
                        rowData.symmetry || '', rowData.regularity || '', rowData.bpm120 || '', rowData.zone5 || '',
                        rowData.age || '', rowData.sex || '', rowData.temp || '', rowData.distanceCol || '',
                        rowData.trotHR || '', rowData.walkHR || '', rowData.notes || ''
                    ]);

                    // Format date cell if we have a valid date
                    if (dateValue) {
                        row.getCell(1).numFmt = 'MM/DD/YYYY';
                    }

                    // Apply color coding to Best 5F
                    const best5fColor = getBest5FColor(rowData.best5f);
                    if (best5fColor) {
                        const argb = ('FF' + best5fColor.replace('#', '')).toUpperCase();
                        row.getCell(BEST5F_COL).fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: argb },
                            bgColor: { argb: argb }
                        };
                    }

                    // Apply color coding to Fast Recovery
                    const fastRecoveryColor = getFastRecoveryColor(rowData.fastRecovery);
                    if (fastRecoveryColor) {
                        const argb = ('FF' + fastRecoveryColor.replace('#', '')).toUpperCase();
                        row.getCell(FAST_RECOVERY_COL).fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: argb },
                            bgColor: { argb: argb }
                        };
                    }

                    // Apply color coding to 15 Recovery
                    const recovery15Color = getRecovery15Color(rowData.recovery15);
                    if (recovery15Color) {
                        const argb = ('FF' + recovery15Color.replace('#', '')).toUpperCase();
                        row.getCell(RECOVERY15_COL).fill = {
                            type: 'pattern',
                            pattern: 'solid',
                            fgColor: { argb: argb },
                            bgColor: { argb: argb }
                        };
                    }

                    return row;
                }

                // Helper to set column widths
                function setColumnWidths(worksheet) {
                    worksheet.columns.forEach((column, i) => {
                        column.width = i === 39 ? 30 : 12; // Notes column wider
                    });
                }

                // Sheet 1: All Training combined
                const allTrainingSheet = workbook.addWorksheet('All Training');
                addHeaderRow(allTrainingSheet);

                // Collect all training data and sort by date (most recent first)
                let allTrainingData = [];
                allHorses.forEach(horseName => {
                    const horseTraining = allHorseDetailData[horseName] || [];
                    const horseInfo = horseLookup[horseName] || { owner: '', country: '' };
                    horseTraining.forEach(entry => {
                        allTrainingData.push({
                            ...entry,
                            horse: entry.horse || horseName,
                            owner: horseInfo.owner,
                            country: horseInfo.country
                        });
                    });
                });

                // Sort by date descending
                allTrainingData.sort((a, b) => {
                    const dateA = a.date ? new Date(a.date.split('/').reverse().join('-')) : new Date(0);
                    const dateB = b.date ? new Date(b.date.split('/').reverse().join('-')) : new Date(0);
                    return dateB - dateA;
                });

                // Add all training data to the first sheet
                allTrainingData.forEach(rowData => {
                    addDataRow(allTrainingSheet, rowData);
                });
                setColumnWidths(allTrainingSheet);

                // Create individual sheets for each horse
                // Sort horses alphabetically for consistent ordering
                const sortedHorses = [...allHorses].sort((a, b) => a.localeCompare(b));

                sortedHorses.forEach(horseName => {
                    // Sanitize sheet name (Excel has restrictions)
                    let sheetName = horseName
                        .replace(/[\\/*?:\[\]]/g, '') // Remove invalid chars
                        .substring(0, 31); // Max 31 chars

                    // Ensure unique sheet name
                    let uniqueName = sheetName;
                    let counter = 1;
                    while (workbook.getWorksheet(uniqueName)) {
                        uniqueName = sheetName.substring(0, 28) + `_${counter}`;
                        counter++;
                    }

                    const horseSheet = workbook.addWorksheet(uniqueName);
                    addHeaderRow(horseSheet);

                    const horseTraining = allHorseDetailData[horseName] || [];
                    const horseInfo = horseLookup[horseName] || { owner: '', country: '' };

                    // Sort horse training by date descending
                    const sortedTraining = [...horseTraining].sort((a, b) => {
                        const dateA = a.date ? new Date(a.date.split('/').reverse().join('-')) : new Date(0);
                        const dateB = b.date ? new Date(b.date.split('/').reverse().join('-')) : new Date(0);
                        return dateB - dateA;
                    });

                    sortedTraining.forEach(rowData => {
                        addDataRow(horseSheet, {
                            ...rowData,
                            horse: rowData.horse || horseName,
                            owner: horseInfo.owner,
                            country: horseInfo.country
                        });
                    });
                    setColumnWidths(horseSheet);
                });

                // Generate filename with date
                const date = new Date().toISOString().split('T')[0];
                const filename = `all_training_data_${date}.xlsx`;

                // Download using ExcelJS buffer
                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = filename;
                link.click();
                URL.revokeObjectURL(url);

                // Show success message
                const totalEntries = allTrainingData.length;
                const totalHorses = sortedHorses.length;
                alert(`Export complete!\n\n${totalHorses} horses\n${totalEntries} total training entries\n${totalHorses + 1} sheets created`);

            } catch (error) {
                console.error('Export error:', error);
                alert('Error exporting data: ' + error.message);
            } finally {
                // Restore button
                btn.textContent = originalText;
                btn.disabled = false;
            }
        }


        // Helper: find horse data with fuzzy matching
        function findHorseData(horseName) {
            // Try exact match first
            if (allHorseDetailData[horseName]) {
                return allHorseDetailData[horseName];
            }

            // Try case-insensitive match
            const keys = Object.keys(allHorseDetailData);
            const lowerName = horseName.toLowerCase();
            let match = keys.find(k => k.toLowerCase() === lowerName);
            if (match) return allHorseDetailData[match];

            // Try matching after stripping non-alphanumeric chars
            const stripName = horseName.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
            match = keys.find(k => k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase() === stripName);
            if (match) return allHorseDetailData[match];

            // Try partial match (name contains or is contained)
            match = keys.find(k => {
                const kStrip = k.replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
                return kStrip.includes(stripName) || stripName.includes(kStrip);
            });
            if (match) return allHorseDetailData[match];

            console.log('No match found for:', horseName, 'Available keys:', keys);
            return [];
        }

        // Called from table row click - uses index to avoid special char issues
        function showHorseDetailByIndex(index) {
            const horse = filteredData[index];
            if (horse) {
                showHorseDetail(horse.name, horse.displayName);
            }
        }

        function showHorseDetail(horseName, displayNameOverride) {
            console.log('Clicked horse:', horseName);
            console.log('Available horse detail data:', Object.keys(allHorseDetailData));

            // Store the raw name for data lookups
            currentHorseRawName = horseName;

            // Get the display name from horseData or use override
            const horse = horseData.find(h => h.name === horseName);
            const displayName = displayNameOverride || horse?.displayName || horseName;
            document.getElementById('horseDetailTitle').textContent = `${displayName} - Training Details`;

            // Use fuzzy matching to find the data
            currentHorseDetailData = findHorseData(horseName);
            console.log('Found data entries:', currentHorseDetailData.length);
            
            document.getElementById('mainView').style.display = 'none';
            document.getElementById('horseDetailView').style.display = 'block';
            
            // Reset scroll position to top and left
            window.scrollTo(0, 0);
            
            // Reset table scroll position to left
            const tableContainer = document.querySelector('.horse-detail-view .table-container');
            if (tableContainer) {
                tableContainer.scrollLeft = 0;
            }
            
            // Initialize navigation after showing the view
            setTimeout(() => {
                initializeTableNavigation();
                initializeFrozenHeader();
                applyMobileStyles(); // Apply mobile styles to horse detail view
            }, 100);
            
            // Reset to default sort (date descending)
            currentHorseDetailSort = { column: 'date', order: 'desc' };
            document.getElementById('horseSortBy').value = 'date';
            const horseSortOrderEl = document.getElementById('horseSortOrder');
            if (horseSortOrderEl) horseSortOrderEl.value = 'desc';
            
            // Reset filters
            document.getElementById('horseAgeFilter').value = '';
            document.getElementById('typeFilter').value = 'all';
            currentTypeFilter = 'all';
            
            // Populate age filter for this horse
            updateHorseAgeFilter(horseName);
            
            document.getElementById('horseSortBy').addEventListener('change', updateHorseDetailSort);
            const horseSortOrderElement = document.getElementById('horseSortOrder');
            if (horseSortOrderElement) horseSortOrderElement.addEventListener('change', updateHorseDetailSort);
            document.getElementById('horseAgeFilter').addEventListener('change', updateHorseDetailSort);
            document.getElementById('typeFilter').addEventListener('change', updateHorseDetailFilter);
            document.getElementById('exportHorseCsv').addEventListener('click', exportHorseDataToCsv);
            
            sortHorseDetailData();
        }

        function showMainView() {
            document.getElementById('horseDetailView').style.display = 'none';
            document.getElementById('mainView').style.display = 'block';

            document.getElementById('horseFilter').value = '';

            // Reset sort to most recent training (default view)
            currentSort = { column: 'lastTrainingDate', order: 'desc' };

            filterData();
        }

        function updateHorseAgeFilter(horseName) {
            const ageFilter = document.getElementById('horseAgeFilter');
            const horseDetailData = findHorseData(horseName);
            const ages = [...new Set(horseDetailData.map(row => row.age).filter(age => age && age !== '-'))].sort((a, b) => a - b);
            
            ageFilter.innerHTML = '<option value="">All Ages</option>';
            ages.forEach(age => {
                const option = document.createElement('option');
                option.value = age;
                option.textContent = age;
                ageFilter.appendChild(option);
            });
        }

        function updateHorseDetailFilter() {
            currentTypeFilter = document.getElementById('typeFilter').value;
            sortHorseDetailData();
        }

        function updateHorseDetailSort() {
            const sortBy = document.getElementById('horseSortBy').value;
            const sortOrderEl = document.getElementById('horseSortOrder');
            const sortOrder = sortOrderEl ? sortOrderEl.value : 'desc'; // Default to desc for horse details
            
            currentHorseDetailSort = { column: sortBy, order: sortOrder };
            sortHorseDetailData();
        }

        function sortHorseTable(column) {
            if (currentHorseDetailSort.column === column) {
                currentHorseDetailSort.order = currentHorseDetailSort.order === 'asc' ? 'desc' : 'asc';
            } else {
                currentHorseDetailSort.column = column;
                currentHorseDetailSort.order = 'asc';
            }
            
            document.getElementById('horseSortBy').value = column;
            const horseSortOrderElem = document.getElementById('horseSortOrder');
            if (horseSortOrderElem) horseSortOrderElem.value = currentHorseDetailSort.order;
            
            sortHorseDetailData();
        }

        function sortHorseDetailData() {
            // Use the stored raw name for data lookup, not the display name from title
            let dataToSort = findHorseData(currentHorseRawName);
            
            // Apply age filter
            const ageFilter = document.getElementById('horseAgeFilter').value;
            if (ageFilter) {
                dataToSort = dataToSort.filter(row => row.age === ageFilter);
            }
            
            // Apply type filter
            if (currentTypeFilter === 'work') {
                dataToSort = dataToSort.filter(row => row.isWork === true);
            } else if (currentTypeFilter === 'race') {
                dataToSort = dataToSort.filter(row => row.isRace === true);
            }
            
            currentHorseDetailData = [...dataToSort];
            
            currentHorseDetailData.sort((a, b) => {
                let aVal = a[currentHorseDetailSort.column];
                let bVal = b[currentHorseDetailSort.column];
                
                if (currentHorseDetailSort.column === 'best1f' || currentHorseDetailSort.column === 'best2f' || currentHorseDetailSort.column === 'best3f' || currentHorseDetailSort.column === 'best4f' || currentHorseDetailSort.column === 'best5f' || currentHorseDetailSort.column === 'best6f' || currentHorseDetailSort.column === 'best7f') {
                    aVal = aVal && isValidTime(aVal) ? timeToSeconds(aVal) : Infinity;
                    bVal = bVal && isValidTime(bVal) ? timeToSeconds(bVal) : Infinity;
                } else if (currentHorseDetailSort.column === 'age' || currentHorseDetailSort.column === 'distance' || currentHorseDetailSort.column === 'avgSpeed' || currentHorseDetailSort.column === 'maxSpeed' || currentHorseDetailSort.column === 'maxHR' || currentHorseDetailSort.column === 'fastRecovery' || currentHorseDetailSort.column === 'recovery15') {
                    aVal = parseFloat(aVal) || 0;
                    bVal = parseFloat(bVal) || 0;
                } else if (currentHorseDetailSort.column === 'date') {
                    aVal = aVal ? new Date(aVal) : new Date(0);
                    bVal = bVal ? new Date(bVal) : new Date(0);
                } else {
                    aVal = aVal || '';
                    bVal = bVal || '';
                }
                
                if (currentHorseDetailSort.order === 'asc') {
                    return aVal < bVal ? -1 : aVal > bVal ? 1 : 0;
                } else {
                    return aVal > bVal ? -1 : aVal < bVal ? 1 : 0;
                }
            });
            
            buildTableHeader();
            displayHorseDetailData();
            
            // Initialize frozen header after data is displayed
            setTimeout(() => {
                initializeFrozenHeader();
            }, 100);
        }

        function displayHorseDetailData() {
            const tbody = document.getElementById('horseDetailTableBody');
            
            if (currentHorseDetailData.length === 0) {
                tbody.innerHTML = '<tr><td colspan="38" class="no-data">No training data available for this horse.</td></tr>';
                return;
            }
            
            tbody.innerHTML = currentHorseDetailData.map((row, index) => {
                let rowClass = '';
                if (row.isNote) {
                    rowClass = 'note-row';
                } else if (row.isWork) {
                    rowClass = 'work-row';
                } else if (row.isRace) {
                    rowClass = 'race-row';
                }
                
                const best5fColor = getBest5FColor(row.best5f);
                const fastRecoveryColor = getFastRecoveryColor(row.fastRecovery);
                const recovery15Color = getRecovery15Color(row.recovery15);
                
                const best5fStyle = best5fColor ? `background-color: ${best5fColor}; color: #000;` : '';
                const fastRecoveryStyle = fastRecoveryColor ? `background-color: ${fastRecoveryColor}; color: #000;` : '';
                const recovery15Style = recovery15Color ? `background-color: ${recovery15Color}; color: #000;` : '';
                
                // Escape quotes for JSON in onclick
                const rowDataStr = JSON.stringify({
                    type: row.type || '',
                    track: row.track || '',
                    surface: row.surface || '',
                    notes: row.notes || ''
                }).replace(/'/g, "\\'").replace(/"/g, '&quot;');

                const cellData = {
                    date: row.date || '-',
                    horse: `<strong>${row.horse || '-'}</strong>`,
                    type: `<div class="type-cell-content">${row.type || '-'}</div>`,
                    track: row.track || '-',
                    surface: row.surface || '-',
                    distance: row.distance || '-',
                    avgSpeed: row.avgSpeed && !isNaN(parseFloat(row.avgSpeed)) ? parseFloat(row.avgSpeed).toFixed(1) : (row.avgSpeed || '-'),
                    maxSpeed: row.maxSpeed && !isNaN(parseFloat(row.maxSpeed)) ? parseFloat(row.maxSpeed).toFixed(1) : (row.maxSpeed || '-'),
                    best1f: row.best1f || '-',
                    best2f: row.best2f || '-',
                    best3f: row.best3f || '-',
                    best4f: row.best4f || '-',
                    best5f: row.best5f || '-',
                    best6f: row.best6f || '-',
                    best7f: row.best7f || '-',
                    maxHR: row.maxHR || '-',
                    fastRecovery: row.fastRecovery || '-',
                    fastQuality: row.fastQuality || '-',
                    fastPercent: row.fastPercent || '-',
                    recovery15: row.recovery15 || '-',
                    quality15: row.quality15 || '-',
                    hr15Percent: row.hr15Percent || '-',
                    maxSL: row.maxSL || '-',
                    slGallop: row.slGallop || '-',
                    sfGallop: row.sfGallop || '-',
                    slWork: row.slWork && !isNaN(parseFloat(row.slWork)) ? parseFloat(row.slWork).toFixed(2) : (row.slWork || '-'),
                    sfWork: row.sfWork || '-',
                    hr2min: row.hr2min || '-',
                    hr5min: row.hr5min || '-',
                    symmetry: row.symmetry || '-',
                    regularity: row.regularity || '-',
                    bpm120: row.bpm120 || '-',
                    zone5: row.zone5 || '-',
                    age: row.age || '-',
                    sex: row.sex || '-',
                    temp: row.temp || '-',
                    distanceCol: row.distanceCol || '-',
                    trotHR: row.trotHR || '-',
                    walkHR: row.walkHR || '-',
                    notes: row.notes ? `<span title="${row.notes}" style="cursor: help; max-width: 150px; display: inline-block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${row.notes.length > 20 ? row.notes.substring(0, 20) + '...' : row.notes}</span>` : '-'
                };

                const rowHTML = columnOrder.map(col => {
                    const display = columnVisibility[col] ? '' : 'style="display: none;"';
                    let cellStyle = '';
                    let cellClass = '';

                    // Apply special styling for specific columns
                    if (col === 'best5f' && best5fStyle) {
                        cellStyle = `style="${best5fStyle}${display ? ' ' + display.substring(7) : ''}"`;
                        cellClass = 'time-cell';
                    } else if (col === 'fastRecovery' && fastRecoveryStyle) {
                        cellStyle = `style="${fastRecoveryStyle}${display ? ' ' + display.substring(7) : ''}"`;
                        cellClass = 'recovery-cell';
                    } else if (col === 'recovery15' && recovery15Style) {
                        cellStyle = `style="${recovery15Style}${display ? ' ' + display.substring(7) : ''}"`;
                        cellClass = 'recovery15-cell';
                    } else if (col === 'horse') {
                        cellClass = 'horse-name-cell';
                        cellStyle = display;
                    } else if (col === 'maxSpeed') {
                        cellClass = 'speed-cell';
                        cellStyle = display;
                    } else if (col === 'age') {
                        cellClass = 'age-cell';
                        cellStyle = display;
                    } else if (col === 'type') {
                        cellClass = 'type-cell';
                        cellStyle = display;
                    } else if (col.includes('best') && col !== 'best5f') {
                        cellClass = 'time-cell';
                        cellStyle = display;
                    } else {
                        cellStyle = display;
                    }

                    return `<td class="${cellClass}" ${cellStyle}>${cellData[col] || '-'}</td>`;
                }).join('');

                // Add Edit button at the end - use base64 encoding for horse name to handle special characters
                const encodedHorse = btoa(encodeURIComponent(row.horse || ''));
                const encodedDate = btoa(encodeURIComponent(row.date || ''));

                let actionBtn;
                if (row.isNote) {
                    // Show Edit button for notes
                    const encodedNote = btoa(encodeURIComponent(row.notes || ''));
                    actionBtn = `<td style="text-align: center;"><button onclick="showEditNoteModal('${encodedHorse}', '${encodedDate}', '${encodedNote}')" style="padding: 2px 8px; font-size: 11px; cursor: pointer;">Edit</button></td>`;
                } else {
                    // Show Edit button for regular entries
                    actionBtn = `<td style="text-align: center;"><button onclick="showEditTrainingModalEncoded('${encodedHorse}', '${encodedDate}', JSON.parse(this.dataset.row))" data-row="${rowDataStr}" style="padding: 2px 8px; font-size: 11px; cursor: pointer;">Edit</button></td>`;
                }

                return `<tr class="${rowClass}">${rowHTML}${actionBtn}</tr>`;
            }).join('');
        }
        
        async function exportHorseDataToCsv() {
            if (currentHorseDetailData.length === 0) {
                alert('No data to export');
                return;
            }

            const horseName = document.getElementById('horseDetailTitle').textContent.replace(' - Training Details', '');

            // Use ExcelJS for proper styling support
            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Training Data');

            // Define headers
            const headers = [
                'Date', 'Horse', 'Type', 'Track', 'Surface', 'Distance', 'Avg Speed', 'Max Speed',
                'Best 1F', 'Best 2F', 'Best 3F', 'Best 4F', 'Best 5F', 'Best 6F', 'Best 7F',
                'Max HR', 'Fast Recovery', 'Fast Quality', 'Fast %', '15 Recovery', '15 Quality',
                'HR 15%', 'Max SL', 'SL Gallop', 'SF Gallop', 'SL Work', 'SF Work', 'HR 2 min',
                'HR 5 min', 'Symmetry', 'Regularity', '120bpm', 'Zone 5', 'Age', 'Sex', 'Temp',
                'Distance (Col)', 'Trot HR', 'Walk HR', 'Notes'
            ];

            // Add header row with bold styling
            const headerRow = worksheet.addRow(headers);
            headerRow.font = { bold: true };
            headerRow.eachCell(cell => {
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFE0E0E0' },
                    bgColor: { argb: 'FFE0E0E0' }
                };
            });

            // Column indices (1-based for ExcelJS)
            const BEST5F_COL = 13;       // Best 5F
            const FAST_RECOVERY_COL = 17; // Fast Recovery
            const RECOVERY15_COL = 20;    // 15 Recovery

            // Helper function to parse date string to Date object
            function parseDate(dateStr) {
                if (!dateStr || dateStr === '-' || dateStr === '') return null;
                let date = new Date(dateStr);
                if (isNaN(date.getTime())) {
                    const parts = dateStr.split('/');
                    if (parts.length === 3) {
                        date = new Date(parts[2], parts[0] - 1, parts[1]);
                    }
                }
                return isNaN(date.getTime()) ? null : date;
            }

            // Add data rows with coloring
            currentHorseDetailData.forEach(rowData => {
                const dateValue = parseDate(rowData.date);

                const row = worksheet.addRow([
                    dateValue || rowData.date || '', rowData.horse || '', rowData.type || '',
                    rowData.track || '', rowData.surface || '', rowData.distance || '',
                    rowData.avgSpeed || '', rowData.maxSpeed || '', rowData.best1f || '', rowData.best2f || '',
                    rowData.best3f || '', rowData.best4f || '', rowData.best5f || '', rowData.best6f || '',
                    rowData.best7f || '', rowData.maxHR || '', rowData.fastRecovery || '', rowData.fastQuality || '',
                    rowData.fastPercent || '', rowData.recovery15 || '', rowData.quality15 || '',
                    rowData.hr15Percent || '', rowData.maxSL || '', rowData.slGallop || '', rowData.sfGallop || '',
                    rowData.slWork || '', rowData.sfWork || '', rowData.hr2min || '', rowData.hr5min || '',
                    rowData.symmetry || '', rowData.regularity || '', rowData.bpm120 || '', rowData.zone5 || '',
                    rowData.age || '', rowData.sex || '', rowData.temp || '', rowData.distanceCol || '',
                    rowData.trotHR || '', rowData.walkHR || '', rowData.notes || ''
                ]);

                // Format date cell
                if (dateValue) {
                    row.getCell(1).numFmt = 'MM/DD/YYYY';
                }

                // Apply color coding to Best 5F
                const best5fColor = getBest5FColor(rowData.best5f);
                if (best5fColor) {
                    const argb = ('FF' + best5fColor.replace('#', '')).toUpperCase();
                    row.getCell(BEST5F_COL).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: argb },
                        bgColor: { argb: argb }
                    };
                }

                // Apply color coding to Fast Recovery
                const fastRecoveryColor = getFastRecoveryColor(rowData.fastRecovery);
                if (fastRecoveryColor) {
                    const argb = ('FF' + fastRecoveryColor.replace('#', '')).toUpperCase();
                    row.getCell(FAST_RECOVERY_COL).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: argb },
                        bgColor: { argb: argb }
                    };
                }

                // Apply color coding to 15 Recovery
                const recovery15Color = getRecovery15Color(rowData.recovery15);
                if (recovery15Color) {
                    const argb = ('FF' + recovery15Color.replace('#', '')).toUpperCase();
                    row.getCell(RECOVERY15_COL).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: argb },
                        bgColor: { argb: argb }
                    };
                }
            });

            // Auto-fit columns (approximate widths)
            worksheet.columns.forEach((column, i) => {
                column.width = i === 39 ? 30 : 12; // Notes column wider
            });

            // Generate filename
            const date = new Date().toISOString().split('T')[0];
            const filename = `${horseName}_training_data_${date}.xlsx`;

            // Download using ExcelJS buffer
            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = filename;
            link.click();
            URL.revokeObjectURL(url);
        }
        
        // Load latest session on page load
        function loadLatestSession() {
            console.log('Loading latest session...');
            fetch('/api/latest')
                .then(response => response.json())
                .then(data => {
                    console.log('Latest session response:', data);
                    if (data.sessionId) {
                        console.log('Loading session:', data.sessionId);
                        localStorage.setItem('currentSessionId', data.sessionId);
                        loadSessionData(data.sessionId);
                    } else {
                        console.log('No latest session found, loading from current sheet');
                        loadCurrentSheet();
                        updateDisplays();
                    }
                })
                .catch(error => {
                    console.error('Error loading latest session:', error);
                    console.log('Falling back to current sheet data');
                    loadCurrentSheet();
                    updateDisplays();
                });
        }
        
        function initializeFrozenHeader() {
            const tableContainer = document.getElementById('horseDetailTableContainer');
            const originalTable = document.getElementById('horseDetailTable');
            const fixedHeader = document.getElementById('fixedHeader');
            
            if (!tableContainer || !originalTable || !fixedHeader) return;
            
            const thead = originalTable.querySelector('thead');
            if (!thead) return;
            
            // Create cloned header
            function createFixedHeader() {
                fixedHeader.innerHTML = '';
                
                const clonedTable = document.createElement('table');
                clonedTable.className = originalTable.className;
                clonedTable.style.cssText = originalTable.style.cssText;
                
                const clonedThead = thead.cloneNode(true);
                clonedTable.appendChild(clonedThead);
                fixedHeader.appendChild(clonedTable);
                
                // Copy event listeners for sorting
                const originalHeaders = thead.querySelectorAll('th[onclick]');
                const clonedHeaders = clonedThead.querySelectorAll('th[onclick]');
                
                originalHeaders.forEach((header, index) => {
                    if (clonedHeaders[index]) {
                        const onclickValue = header.getAttribute('onclick');
                        if (onclickValue) {
                            clonedHeaders[index].setAttribute('onclick', onclickValue);
                        }
                    }
                });
            }
            
            // Update header positioning and visibility
            function updateFixedHeader() {
                const containerRect = tableContainer.getBoundingClientRect();
                const theadRect = thead.getBoundingClientRect();
                const scrollTop = tableContainer.scrollTop;
                
                // Show fixed header when original header scrolls out of view
                if (scrollTop > 10 && theadRect.bottom < containerRect.top) {
                    fixedHeader.style.display = 'block';
                    
                    // Sync column widths precisely
                    const originalCells = thead.querySelectorAll('th');
                    const fixedCells = fixedHeader.querySelectorAll('th');
                    const fixedTable = fixedHeader.querySelector('table');
                    
                    // Copy computed styles from original table
                    const originalComputedStyle = window.getComputedStyle(originalTable);
                    fixedTable.style.width = originalComputedStyle.width;
                    
                    originalCells.forEach((cell, index) => {
                        if (fixedCells[index]) {
                            const cellComputedStyle = window.getComputedStyle(cell);
                            const width = cell.offsetWidth;
                            const padding = cellComputedStyle.padding;
                            const fontSize = cellComputedStyle.fontSize;
                            const fontFamily = cellComputedStyle.fontFamily;
                            
                            fixedCells[index].style.cssText = `
                                width: ${width}px !important;
                                min-width: ${width}px !important;
                                max-width: ${width}px !important;
                                padding: ${padding};
                                font-size: ${fontSize};
                                font-family: ${fontFamily};
                                background: #34495e;
                                color: white;
                                border: 1px solid #2c3e50;
                                font-weight: bold;
                                text-align: center;
                                cursor: pointer;
                                line-height: 1.2;
                                white-space: nowrap;
                                vertical-align: middle;
                            `;
                        }
                    });
                } else {
                    fixedHeader.style.display = 'none';
                }
            }
            
            // Initialize
            createFixedHeader();
            updateFixedHeader();
            
            // Add scroll listener
            tableContainer.addEventListener('scroll', updateFixedHeader);
            
            // Update header on window resize
            window.addEventListener('resize', () => {
                setTimeout(() => {
                    createFixedHeader();
                    updateFixedHeader();
                }, 100);
            });
        }
        
        // Load session data
        // Filter out invalid horse names (default Excel sheet names)
        function isValidHorseName(name) {
            if (!name) return false;
            const lowerName = name.toLowerCase().trim();
            const invalidNames = ['worksheet', 'sheet1', 'sheet2', 'sheet3', 'data', 'default'];
            return lowerName && !invalidNames.includes(lowerName) && !lowerName.startsWith('sheet');
        }

        function loadSessionData(sessionId) {
            fetch(`/api/session/${sessionId}`)
                .then(response => response.json())
                .then(data => {
                    if (data.horseData && data.allHorseDetailData) {
                        // Filter out invalid horse names from existing data
                        horseData = data.horseData.filter(h => isValidHorseName(h.name));
                        allHorseDetailData = {};
                        Object.keys(data.allHorseDetailData).forEach(name => {
                            if (isValidHorseName(name)) {
                                allHorseDetailData[name] = data.allHorseDetailData[name];
                            }
                        });

                        // Load sheet metadata from Redis if available, else from localStorage
                        if (data.allSheets && data.sheetNames) {
                            allSheets = data.allSheets;
                            sheetNames = data.sheetNames;
                            if (data.currentSheetName) {
                                currentSheetName = data.currentSheetName;
                            }
                            updateSheetDropdown();
                        } else {
                            // Fallback to localStorage for backwards compatibility
                            loadAllSheets();
                        }

                        // Update allSheets with server data (which has edits/notes applied)
                        // This ensures loadCurrentSheet uses the correct data
                        if (currentSheetName && allSheets) {
                            allSheets[currentSheetName] = {
                                horseData: data.horseData,
                                allHorseDetailData: data.allHorseDetailData
                            };
                        }

                        // Load the sheet data (now uses server data with edits)
                        loadCurrentSheet();
                        console.log('Loaded data for sheet:', currentSheetName, 'horses:', horseData.length);

                        // Calculate Last Work dates for loaded session
                        calculateLastWorkDates();

                        updateHorseFilter();
                        updateAgeFilter();

                        // Default sort to most recent training
                        currentSort = { column: 'lastTrainingDate', order: 'desc' };
                        filterData();

                        // Apply user preferences (column visibility only, not sort)
                        loadUserPreferences();
                        updateColumnVisibilityUI();
                        updateTableColumnVisibility();

                        document.getElementById('exportCsv').disabled = false;
                    document.getElementById('exportAllTraining').disabled = false;

                        // Refresh horse detail view if it's currently visible
                        const horseDetailView = document.getElementById('horseDetailView');
                        if (horseDetailView && horseDetailView.style.display !== 'none' && currentHorseRawName) {
                            sortHorseDetailData();
                        }
                    }
                })
                .catch(error => {
                    console.error('Error loading session data:', error);
                });
        }
        
        // Show share dialog
        function showShareDialog(shareUrl) {
            const dialog = document.createElement('div');
            dialog.innerHTML = `
                <div style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 10000; display: flex; align-items: center; justify-content: center;">
                    <div style="background: white; padding: 30px; border-radius: 12px; max-width: 500px; text-align: center; box-shadow: 0 10px 30px rgba(0,0,0,0.3);">
                        <h3 style="margin: 0 0 20px 0; color: #2c3e50;">File Uploaded Successfully!</h3>
                        <p style="margin: 0 0 20px 0; color: #666;">Share this link with others to let them view the analysis results:</p>
                        <div style="background: #f8f9fa; padding: 15px; border-radius: 6px; margin: 0 0 20px 0; border: 1px solid #e9ecef;">
                            <input type="text" value="${shareUrl}" readonly style="width: 100%; border: none; background: none; font-family: monospace; font-size: 14px; color: #2c3e50; text-align: center;">
                        </div>
                        <div style="display: flex; gap: 10px; justify-content: center;">
                            <button onclick="copyShareUrl('${shareUrl}')" style="background: #3498db; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer;">Copy Link</button>
                            <button onclick="closeShareDialog()" style="background: #95a5a6; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer;">Close</button>
                        </div>
                    </div>
                </div>
            `;
            dialog.id = 'shareDialog';
            document.body.appendChild(dialog);
        }
        
        // Copy share URL to clipboard
        function copyShareUrl(url) {
            navigator.clipboard.writeText(url).then(() => {
                alert('Share URL copied to clipboard!');
            }).catch(err => {
                console.error('Could not copy text: ', err);
                // Fallback: select the text
                const input = document.querySelector('#shareDialog input');
                input.select();
                input.setSelectionRange(0, 99999);
                document.execCommand('copy');
                alert('Share URL copied to clipboard!');
            });
        }
        
        // Close share dialog
        function closeShareDialog() {
            const dialog = document.getElementById('shareDialog');
            if (dialog) {
                document.body.removeChild(dialog);
            }
        }
        
        // Multi-sheet management functions
        function updateSheetDropdown() {
            const dropdown = document.getElementById('sheetSelector');
            dropdown.innerHTML = '';
            
            if (sheetNames.length === 0) {
                const option = document.createElement('option');
                option.value = 'Default';
                option.textContent = 'Default';
                dropdown.appendChild(option);
                updateDeleteButtonState();
                return;
            }
            
            sheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                if (name === currentSheetName) {
                    option.selected = true;
                }
                dropdown.appendChild(option);
            });
            
            updateDeleteButtonState();
        }
        
        function switchSheet() {
            const dropdown = document.getElementById('sheetSelector');
            const selectedSheet = dropdown.value;

            if (selectedSheet !== currentSheetName) {
                console.log('Switching from', currentSheetName, 'to', selectedSheet);
                currentSheetName = selectedSheet;

                loadCurrentSheet();
                console.log('Loaded sheet data:', {
                    horsesCount: horseData.length,
                    detailDataKeys: Object.keys(allHorseDetailData),
                    firstHorse: horseData[0]?.name
                });

                // Calculate Last Work dates for the newly loaded sheet
                calculateLastWorkDates();
                console.log('After calculating last work dates, first horse:', horseData[0]);

                updateDisplays();
                updateDeleteButtonState();

                console.log('Sheet switch completed to', currentSheetName);
            }
        }
        
        function showAddNewModal() {
            document.getElementById('addNewModal').style.display = 'flex';
            document.getElementById('sheetTitle').value = '';
            document.getElementById('newSheetFile').value = '';
        }
        
        function closeAddNewModal() {
            document.getElementById('addNewModal').style.display = 'none';
        }
        
        function closeModalOnOutsideClick(event, modalId = 'addNewModal') {
            if (event.target === event.currentTarget) {
                if (modalId === 'deleteSheetModal') {
                    closeDeleteSheetModal();
                } else {
                    closeAddNewModal();
                }
            }
        }
        
        // Sheet deletion functions
        function showDeleteSheetConfirm() {
            // Don't allow deleting if there's only one sheet or it's Default
            if (sheetNames.length <= 1 || currentSheetName === 'Default') {
                if (currentSheetName === 'Default') {
                    alert('Cannot delete the Default sheet.');
                } else {
                    alert('Cannot delete the last remaining sheet.');
                }
                return;
            }
            
            document.getElementById('deleteSheetName').textContent = currentSheetName;
            document.getElementById('deleteSheetModal').style.display = 'flex';
        }
        
        function closeDeleteSheetModal() {
            document.getElementById('deleteSheetModal').style.display = 'none';
        }
        
        function confirmDeleteSheet() {
            if (currentSheetName === 'Default') {
                alert('Cannot delete the Default sheet.');
                closeDeleteSheetModal();
                return;
            }
            
            if (sheetNames.length <= 1) {
                alert('Cannot delete the last remaining sheet.');
                closeDeleteSheetModal();
                return;
            }
            
            const sheetToDelete = currentSheetName;
            
            // Remove from allSheets and sheetNames
            delete allSheets[sheetToDelete];
            sheetNames = sheetNames.filter(name => name !== sheetToDelete);
            
            // Switch to the first available sheet
            if (sheetNames.length > 0) {
                currentSheetName = sheetNames[0];
            } else {
                currentSheetName = 'Default';
                sheetNames = ['Default'];
                allSheets['Default'] = { horseData: [], allHorseDetailData: {} };
            }
            
            // Save changes and update UI
            saveAllSheets();
            loadCurrentSheet();

            // Calculate Last Work dates for the newly loaded sheet
            calculateLastWorkDates();

            updateDisplays();
            updateSheetDropdown();
            updateDeleteButtonState();
            closeDeleteSheetModal();

            // Sync sheet data to Redis if we have a current session
            const urlParams = new URLSearchParams(window.location.search);
            const sessionId = urlParams.get('session') || localStorage.getItem('currentSessionId');
            if (sessionId) {
                syncSheetDataToRedis(sessionId);
            }
            
            alert(`Sheet "${sheetToDelete}" has been deleted.`);
        }
        
        function updateDeleteButtonState() {
            const deleteBtn = document.querySelector('.delete-sheet-btn');
            if (deleteBtn) {
                // Disable delete button if only one sheet or current is Default
                const shouldDisable = sheetNames.length <= 1 || currentSheetName === 'Default';
                deleteBtn.disabled = shouldDisable;
                deleteBtn.title = shouldDisable ? 
                    (currentSheetName === 'Default' ? 'Cannot delete Default sheet' : 'Cannot delete last sheet') : 
                    'Delete current sheet';
            }
        }
        
        function uploadNewSheet() {
            const title = document.getElementById('sheetTitle').value.trim();
            const fileInput = document.getElementById('newSheetFile');
            
            if (!title) {
                alert('Please enter a title for the sheet.');
                return;
            }
            
            if (!fileInput.files || fileInput.files.length === 0) {
                alert('Please select a file to upload.');
                return;
            }
            
            if (sheetNames.includes(title)) {
                if (!confirm(`A sheet named "${title}" already exists. Do you want to replace it?`)) {
                    return;
                }
            }
            
            currentSheetName = title;
            const file = fileInput.files[0];
            processFileForNewSheet(file);
        }
        
        function processFileForNewSheet(file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    
                    const sheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[sheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    parseAndStoreNewSheetData(jsonData);
                    
                } catch (error) {
                    console.error('Error reading file:', error);
                    alert('Error reading the file. Please make sure it\'s a valid Excel file.');
                }
            };
            reader.readAsArrayBuffer(file);
        }
        
        function parseAndStoreNewSheetData(jsonData) {
            try {
                const newHorseData = [];
                const newAllHorseDetailData = {};
                
                // Debug: Log the first row to see actual column names
                if (jsonData.length > 0) {
                    console.log('First row columns:', Object.keys(jsonData[0]));
                    console.log('First row sample:', jsonData[0]);
                }
                
                // Use the same parsing logic as the original file upload
                jsonData.forEach(row => {
                    const processedRow = parseRow(row);
                    if (processedRow) {
                        const horseName = processedRow.horse;
                        if (!newAllHorseDetailData[horseName]) {
                            newAllHorseDetailData[horseName] = [];
                        }
                        newAllHorseDetailData[horseName].push(processedRow);
                    }
                });
                
                // Generate summary data for the main view
                for (let horseName in newAllHorseDetailData) {
                    const horseRows = newAllHorseDetailData[horseName];
                    const summaryData = generateHorseSummary(horseName, horseRows);
                    if (summaryData) {
                        newHorseData.push(summaryData);
                    }
                }
                
                // Store the new sheet data
                allSheets[currentSheetName] = { 
                    horseData: newHorseData, 
                    allHorseDetailData: newAllHorseDetailData 
                };
                
                if (!sheetNames.includes(currentSheetName)) {
                    sheetNames.push(currentSheetName);
                }
                
                saveAllSheets();
                loadCurrentSheet();
                updateDisplays();
                updateSheetDropdown();
                updateDeleteButtonState();
                closeAddNewModal();

                // Sync sheet data to Redis if we have a current session
                const urlParams = new URLSearchParams(window.location.search);
                const sessionId = urlParams.get('session') || localStorage.getItem('currentSessionId');
                if (sessionId) {
                    syncSheetDataToRedis(sessionId);
                }

                alert(`Sheet "${currentSheetName}" has been successfully added!`);
                
            } catch (error) {
                console.error('Error processing sheet data:', error);
                alert('Error processing the sheet data. Please check the file format.');
            }
        }
        
        function updateDisplays() {
            filterData();
            updateHorseFilter();
            updateAgeFilter();
        }
        
        // Helper function to generate horse summary (reused from existing logic)
        function generateHorseSummary(horseName, horseRows) {
            if (!horseRows || horseRows.length === 0) return null;

            // Get the most recent training entry (sorted by date, most recent first)
            const sortedRows = [...horseRows].sort((a, b) => new Date(b.date) - new Date(a.date));
            const lastTraining = sortedRows[0];

            // Get data from most recent training
            const lastTrainingDate = lastTraining ? lastTraining.date : 'N/A';
            const best1f = lastTraining ? lastTraining.best1f : '-';
            const best5f = lastTraining ? lastTraining.best5f : '-';
            const fastRecovery = lastTraining ? lastTraining.fastRecovery : '-';
            const recovery15min = lastTraining ? lastTraining.recovery15 : '-';
            const age = lastTraining ? lastTraining.age : 'N/A';

            return {
                name: horseName,
                age: age,
                lastTrainingDate: lastTrainingDate,
                best1f: best1f,
                best5f: best5f,
                fastRecovery: fastRecovery,
                recovery15min: recovery15min,
                best5fColor: getBest5FColor(best5f),
                fastRecoveryColor: getFastRecoveryColor(fastRecovery),
                recovery15Color: getRecovery15Color(recovery15min)
            };
        }
        
        function findBestTime(rows, field) {
            let best = null;
            rows.forEach(row => {
                if (row[field] && row[field] !== '-') {
                    if (!best || parseTime(row[field]) < parseTime(best)) {
                        best = row[field];
                    }
                }
            });
            return best || '-';
        }
        
        function findDateOfBestTime(rows, field) {
            let bestTime = null;
            let bestDate = null;
            rows.forEach(row => {
                if (row[field] && row[field] !== '-') {
                    if (!bestTime || parseTime(row[field]) < parseTime(bestTime)) {
                        bestTime = row[field];
                        bestDate = row.date;
                    }
                }
            });
            return bestDate || '-';
        }
        
        function findLatestValue(rows, field) {
            for (let row of rows) {
                if (row[field] && row[field] !== '-') {
                    return row[field];
                }
            }
            return '-';
        }
        
        function findHighestValue(rows, field) {
            let highest = null;
            rows.forEach(row => {
                if (row[field] && row[field] !== '-') {
                    const val = parseFloat(row[field]);
                    if (!isNaN(val) && (highest === null || val > highest)) {
                        highest = val;
                    }
                }
            });
            return highest !== null ? highest.toString() : '-';
        }
        
        // Parse individual row from Excel data
        function parseRow(row) {
            try {
                // Map the Excel columns to our internal structure
                const parsedRow = {};
                
                // Create a flexible column mapping that handles variations
                const columnMappings = {
                    'date': ['Date', 'date', 'DATE', 'Date/Time'],
                    'horse': ['Horse', 'horse', 'HORSE', 'Horse Name', 'HorseName', 'Name'],
                    'type': ['Type', 'type', 'TYPE', 'Entry Type', 'EntryType'],
                    'track': ['Track', 'track', 'TRACK'],
                    'surface': ['Surface', 'surface', 'SURFACE'],
                    'distance': ['Distance', 'distance', 'DISTANCE', 'Dist'],
                    'avgSpeed': ['Avg Speed', 'Average Speed', 'AvgSpeed', 'avgSpeed', 'AVG_SPEED'],
                    'maxSpeed': ['Max Speed', 'Maximum Speed', 'MaxSpeed', 'maxSpeed', 'MAX_SPEED'],
                    'best1f': ['Best 1F', 'Best1F', 'best1f', '1F', '1f'],
                    'best2f': ['Best 2F', 'Best2F', 'best2f', '2F', '2f'],
                    'best3f': ['Best 3F', 'Best3F', 'best3f', '3F', '3f'],
                    'best4f': ['Best 4F', 'Best4F', 'best4f', '4F', '4f'],
                    'best5f': ['Best 5F', 'Best5F', 'best5f', '5F', '5f'],
                    'best6f': ['Best 6F', 'Best6F', 'best6f', '6F', '6f'],
                    'best7f': ['Best 7F', 'Best7F', 'best7f', '7F', '7f'],
                    'maxHR': ['Max HR', 'Maximum HR', 'MaxHR', 'maxHR', 'MAX_HR', 'Heart Rate Max'],
                    'fastRecovery': ['Fast Recovery', 'FastRecovery', 'fastRecovery', 'FAST_RECOVERY'],
                    'fastQuality': ['Fast Quality', 'FastQuality', 'fastQuality', 'FAST_QUALITY'],
                    'fastPercent': ['Fast %', 'Fast Percent', 'FastPercent', 'fastPercent', 'FAST_PERCENT'],
                    'recovery15': ['15 Recovery', '15Recovery', 'recovery15', 'RECOVERY_15', 'Recovery 15'],
                    'quality15': ['15 Quality', '15Quality', 'quality15', 'QUALITY_15', 'Quality 15'],
                    'hr15Percent': ['HR 15%', 'HR15%', 'hr15Percent', 'HR_15_PERCENT'],
                    'maxSL': ['Max SL', 'MaxSL', 'maxSL', 'MAX_SL'],
                    'slGallop': ['SL Gallop', 'SLGallop', 'slGallop', 'SL_GALLOP'],
                    'sfGallop': ['SF Gallop', 'SFGallop', 'sfGallop', 'SF_GALLOP'],
                    'slWork': ['SL Work', 'SLWork', 'slWork', 'SL_WORK'],
                    'sfWork': ['SF Work', 'SFWork', 'sfWork', 'SF_WORK'],
                    'hr2min': ['HR 2 min', 'HR2min', 'hr2min', 'HR_2_MIN'],
                    'hr5min': ['HR 5 min', 'HR5min', 'hr5min', 'HR_5_MIN'],
                    'symmetry': ['Symmetry', 'symmetry', 'SYMMETRY'],
                    'regularity': ['Regularity', 'regularity', 'REGULARITY'],
                    'bpm120': ['120bpm', '120 BPM', 'bpm120', 'BPM_120'],
                    'zone5': ['Zone 5', 'Zone5', 'zone5', 'ZONE_5'],
                    'age': ['Age', 'age', 'AGE'],
                    'sex': ['Sex', 'sex', 'SEX', 'Gender'],
                    'temp': ['Temp', 'temp', 'TEMP', 'Temperature'],
                    'trotHR': ['Trot HR', 'TrotHR', 'trotHR', 'TROT_HR'],
                    'walkHR': ['Walk HR', 'WalkHR', 'walkHR', 'WALK_HR']
                };
                
                // Find matching columns using flexible mapping
                for (const [internalCol, possibleNames] of Object.entries(columnMappings)) {
                    let foundValue = null;
                    
                    for (const possibleName of possibleNames) {
                        if (row.hasOwnProperty(possibleName)) {
                            foundValue = row[possibleName];
                            break;
                        }
                    }
                    
                    if (foundValue !== null) {
                        parsedRow[internalCol] = foundValue || '-';
                    } else {
                        parsedRow[internalCol] = '-';
                    }
                }
                
                // Debug: Log column mapping results for first few rows
                if (Math.random() < 0.1) { // Log ~10% of rows for debugging
                    console.log('Parsed row sample:', {
                        originalKeys: Object.keys(row),
                        parsedData: parsedRow,
                        horse: parsedRow.horse
                    });
                }
                
                // Ensure horse name exists
                if (!parsedRow.horse || parsedRow.horse === '-') {
                    return null;
                }
                
                // Clean up the data
                parsedRow.horse = parsedRow.horse.toString().trim();
                
                // Determine if it's a work or race
                if (parsedRow.type) {
                    const typeStr = parsedRow.type.toString().toLowerCase();
                    parsedRow.isWork = typeStr.includes('work');
                    parsedRow.isRace = typeStr.includes('race');
                }
                
                // Format date
                if (parsedRow.date) {
                    parsedRow.date = formatDate(parsedRow.date);
                }
                
                return parsedRow;
            } catch (error) {
                console.error('Error parsing row:', error, row);
                return null;
            }
        }
        
        // Helper function to format dates
        function formatDate(dateValue) {
            try {
                if (!dateValue) return '-';
                
                // If it's already a formatted string, return as is
                if (typeof dateValue === 'string' && dateValue.includes('/')) {
                    return dateValue;
                }
                
                // If it's an Excel date number
                if (typeof dateValue === 'number') {
                    const date = new Date((dateValue - 25569) * 86400 * 1000);
                    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
                }
                
                // Try to parse as date
                const date = new Date(dateValue);
                if (!isNaN(date.getTime())) {
                    return (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear();
                }
                
                return dateValue.toString();
            } catch (error) {
                return dateValue ? dateValue.toString() : '-';
            }
        }
        
