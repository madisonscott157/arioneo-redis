// Race Chart Upload Module
// Handles bulk PDF upload, parsing, review, and saving of race results

class RaceChartUploader {
    constructor() {
        this.modal = null;
        this.results = [];
        this.existingHorses = [];
        this.init();
    }

    init() {
        this.createModal();
        this.bindEvents();
    }

    createModal() {
        // Create modal HTML
        const modalHTML = `
            <div id="raceUploadModal" class="race-upload-modal" style="display: none;">
                <div class="race-upload-content">
                    <div class="race-upload-header">
                        <h2>Upload Race Charts</h2>
                        <button class="race-upload-close" onclick="raceUploader.closeModal()">&times;</button>
                    </div>

                    <div class="race-upload-body">
                        <!-- Step 1: File Upload -->
                        <div id="raceUploadStep1" class="race-upload-step">
                            <div class="race-upload-dropzone" id="raceDropzone">
                                <input type="file" id="racePdfInput" accept=".pdf" multiple style="display: none;">
                                <div class="dropzone-content">
                                    <div class="dropzone-icon">📄</div>
                                    <p>Drag & drop PDF race charts here</p>
                                    <p class="dropzone-subtitle">or click to select files (up to 50 PDFs)</p>
                                </div>
                            </div>
                            <div id="selectedFilesList" class="selected-files-list"></div>
                            <div class="race-upload-actions">
                                <button id="processRacePdfs" class="race-btn race-btn-primary" disabled>
                                    Process PDFs
                                </button>
                            </div>
                        </div>

                        <!-- Step 2: Processing -->
                        <div id="raceUploadStep2" class="race-upload-step" style="display: none;">
                            <div class="processing-status">
                                <div class="processing-spinner"></div>
                                <p id="processingText">Processing PDFs...</p>
                                <div class="progress-bar">
                                    <div id="progressFill" class="progress-fill"></div>
                                </div>
                                <p id="progressText">0 of 0</p>
                            </div>
                        </div>

                        <!-- Step 3: Review -->
                        <div id="raceUploadStep3" class="race-upload-step" style="display: none;">
                            <div class="review-summary" id="reviewSummary"></div>
                            <div class="review-list" id="reviewList"></div>
                            <div class="race-upload-actions">
                                <button onclick="raceUploader.goBackToUpload()" class="race-btn race-btn-secondary">
                                    Upload More
                                </button>
                                <button id="saveAllRaces" class="race-btn race-btn-primary" disabled>
                                    Save All Races
                                </button>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;

        document.body.insertAdjacentHTML('beforeend', modalHTML);
        this.modal = document.getElementById('raceUploadModal');
    }

    bindEvents() {
        const dropzone = document.getElementById('raceDropzone');
        const fileInput = document.getElementById('racePdfInput');
        const processBtn = document.getElementById('processRacePdfs');

        // Dropzone click
        dropzone.addEventListener('click', () => fileInput.click());

        // File input change
        fileInput.addEventListener('change', (e) => this.handleFileSelect(e.target.files));

        // Drag and drop
        dropzone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropzone.classList.add('dragover');
        });

        dropzone.addEventListener('dragleave', () => {
            dropzone.classList.remove('dragover');
        });

        dropzone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropzone.classList.remove('dragover');
            const files = Array.from(e.dataTransfer.files).filter(f => f.type === 'application/pdf');
            this.handleFileSelect(files);
        });

        // Process button
        processBtn.addEventListener('click', () => this.processPdfs());

        // Save button
        document.getElementById('saveAllRaces').addEventListener('click', () => this.saveAllRaces());
    }

    openModal() {
        this.modal.style.display = 'flex';
        this.resetModal();
    }

    closeModal() {
        this.modal.style.display = 'none';
        this.resetModal();
    }

    resetModal() {
        document.getElementById('raceUploadStep1').style.display = 'block';
        document.getElementById('raceUploadStep2').style.display = 'none';
        document.getElementById('raceUploadStep3').style.display = 'none';
        document.getElementById('selectedFilesList').innerHTML = '';
        document.getElementById('racePdfInput').value = '';
        document.getElementById('processRacePdfs').disabled = true;
        this.results = [];
    }

    handleFileSelect(files) {
        const filesList = document.getElementById('selectedFilesList');
        const processBtn = document.getElementById('processRacePdfs');

        if (!files || files.length === 0) return;

        const pdfFiles = Array.from(files).filter(f =>
            f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
        );

        if (pdfFiles.length === 0) {
            alert('Please select PDF files only');
            return;
        }

        filesList.innerHTML = `
            <div class="files-header">
                <span>${pdfFiles.length} PDF${pdfFiles.length !== 1 ? 's' : ''} selected</span>
                <button class="clear-files-btn" onclick="raceUploader.clearFiles()">Clear All</button>
            </div>
            <div class="files-grid">
                ${pdfFiles.map((f, i) => `
                    <div class="file-item">
                        <span class="file-icon">📄</span>
                        <span class="file-name" title="${f.name}">${this.truncateFileName(f.name)}</span>
                    </div>
                `).join('')}
            </div>
        `;

        // Store files for processing
        this.selectedFiles = pdfFiles;
        processBtn.disabled = false;
    }

    truncateFileName(name) {
        if (name.length <= 25) return name;
        return name.substring(0, 22) + '...';
    }

    clearFiles() {
        document.getElementById('selectedFilesList').innerHTML = '';
        document.getElementById('racePdfInput').value = '';
        document.getElementById('processRacePdfs').disabled = true;
        this.selectedFiles = [];
    }

    async processPdfs() {
        if (!this.selectedFiles || this.selectedFiles.length === 0) return;

        // Show processing step
        document.getElementById('raceUploadStep1').style.display = 'none';
        document.getElementById('raceUploadStep2').style.display = 'block';

        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const processingText = document.getElementById('processingText');

        progressText.textContent = `0 of ${this.selectedFiles.length}`;

        try {
            // Create form data with all PDFs
            const formData = new FormData();
            for (const file of this.selectedFiles) {
                formData.append('pdfs', file);
            }

            processingText.textContent = 'Uploading and parsing PDFs...';
            progressFill.style.width = '10%';

            const response = await fetch('/api/upload/race-charts', {
                method: 'POST',
                body: formData
            });

            progressFill.style.width = '90%';

            if (!response.ok) {
                throw new Error(`Server error: ${response.status}`);
            }

            const data = await response.json();

            progressFill.style.width = '100%';
            processingText.textContent = 'Processing complete!';

            this.results = data.results;
            this.existingHorses = data.existingHorses || [];

            setTimeout(() => this.showReviewStep(data), 500);

        } catch (error) {
            console.error('Error processing PDFs:', error);
            alert('Error processing PDFs: ' + error.message);
            this.goBackToUpload();
        }
    }

    showReviewStep(data) {
        document.getElementById('raceUploadStep2').style.display = 'none';
        document.getElementById('raceUploadStep3').style.display = 'block';

        // Show summary
        const summary = document.getElementById('reviewSummary');
        summary.innerHTML = `
            <div class="summary-stats">
                <div class="stat stat-success">
                    <span class="stat-number">${data.successfulMatches}</span>
                    <span class="stat-label">Ready to Save</span>
                </div>
                <div class="stat stat-warning">
                    <span class="stat-number">${data.needsVerification}</span>
                    <span class="stat-label">Needs Review</span>
                </div>
                <div class="stat stat-info">
                    <span class="stat-number">${data.duplicates}</span>
                    <span class="stat-label">Duplicates (will skip)</span>
                </div>
            </div>
        `;

        // Build review list
        const reviewList = document.getElementById('reviewList');
        reviewList.innerHTML = this.results.map((result, index) => this.buildReviewCard(result, index)).join('');

        // Add event listeners for checkboxes to update save button count
        document.querySelectorAll('.include-race').forEach(checkbox => {
            checkbox.addEventListener('change', () => this.updateSaveButton());
        });

        // Update save button
        this.updateSaveButton();
    }

    buildReviewCard(result, index) {
        const statusClass = result.isDuplicate ? 'duplicate' :
                          result.success && !result.needsVerification ? 'success' :
                          result.needsVerification ? 'warning' : 'error';

        const statusIcon = result.isDuplicate ? '⏭️' :
                          result.success && !result.needsVerification ? '✅' :
                          result.needsVerification ? '⚠️' : '❌';

        const statusText = result.isDuplicate ? 'Duplicate - Will Skip' :
                          result.success && !result.needsVerification ? 'Ready to Save' :
                          result.needsVerification ? 'Needs Verification' : 'Parse Error';

        let cardContent = '';

        if (result.data) {
            const d = result.data;
            cardContent = `
                <div class="review-card-data">
                    <div class="data-row">
                        <div class="data-field">
                            <label>Horse</label>
                            ${this.buildHorseSelector(result, index)}
                        </div>
                        <div class="data-field">
                            <label>Date</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="date"
                                   value="${this.formatDisplayDate(d.raceDate)}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>Track</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="track"
                                   value="${d.track || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>Surface</label>
                            <select class="review-input" data-index="${index}" data-field="surface" ${result.isDuplicate ? 'disabled' : ''}>
                                <option value="D" ${d.surface === 'D' ? 'selected' : ''}>Dirt</option>
                                <option value="T" ${d.surface === 'T' ? 'selected' : ''}>Turf</option>
                                <option value="AWT" ${d.surface === 'AWT' ? 'selected' : ''}>AWT</option>
                            </select>
                        </div>
                    </div>
                    <div class="data-row">
                        <div class="data-field">
                            <label>Distance</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="distance"
                                   value="${d.distance || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>Race Type</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="raceType"
                                   value="${d.raceType || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>Final Time</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="finalTime"
                                   value="${d.finalTime || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>Avg Speed</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="avgSpeed"
                                   value="${d.avgSpeedMph ? d.avgSpeedMph.toFixed(1) : ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                    </div>
                    <div class="data-row">
                        <div class="data-field">
                            <label>Finish</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="finishPosition"
                                   value="${d.posFin || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field">
                            <label>5F Reduction</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="fiveFReduction"
                                   value="${d.fiveFReductionTime || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                    </div>
                    <div class="data-row fractionals-row">
                        <div class="data-field small">
                            <label>1/4 Pos</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="pos1_4"
                                   value="${d.pos1_4 || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field small">
                            <label>1/4 Time</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="f1Time"
                                   value="${d.f1Time || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field small">
                            <label>1/2 Pos</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="pos1_2"
                                   value="${d.pos1_2 || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field small">
                            <label>1/2 Time</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="f2Time"
                                   value="${d.f2Time || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field small">
                            <label>3/4 Pos</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="pos3_4"
                                   value="${d.pos3_4 || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                        <div class="data-field small">
                            <label>3/4 Time</label>
                            <input type="text" class="review-input small" data-index="${index}" data-field="f3Time"
                                   value="${d.f3Time || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                    </div>
                    <div class="data-row">
                        <div class="data-field full-width">
                            <label>Comments</label>
                            <input type="text" class="review-input" data-index="${index}" data-field="comments"
                                   value="${d.comment || ''}" ${result.isDuplicate ? 'disabled' : ''}>
                        </div>
                    </div>
                </div>
            `;
        } else if (result.error) {
            // Manual entry form for parse errors
            cardContent = `
                <div class="review-card-error">
                    <p class="error-message">${result.error}</p>
                    <div class="manual-entry-toggle">
                        <button class="race-btn race-btn-small" onclick="raceUploader.showManualEntry(${index})">
                            Enter Data Manually
                        </button>
                    </div>
                    <div id="manualEntry${index}" class="manual-entry-form" style="display: none;">
                        ${this.buildManualEntryForm(index)}
                    </div>
                </div>
            `;
        }

        return `
            <div class="review-card ${statusClass}" data-index="${index}">
                <div class="review-card-header">
                    <div class="card-file-info">
                        <span class="status-icon">${statusIcon}</span>
                        <span class="file-name">${result.fileName}</span>
                        <span class="status-badge ${statusClass}">${statusText}</span>
                    </div>
                    ${!result.isDuplicate ? `
                        <label class="include-checkbox">
                            <input type="checkbox" class="include-race" data-index="${index}" checked>
                            Include
                        </label>
                    ` : ''}
                </div>
                ${cardContent}
            </div>
        `;
    }

    buildHorseSelector(result, index) {
        // Determine which horse to select - prioritize fuzzy match from server
        const matchedName = result.matchedHorse?.name || '';
        const chartHorse = result.selectedHorse || result.data?.horseName || '';
        const confidence = result.matchConfidence || 0;
        const isVerified = confidence >= 0.9;

        // Use horses found in the chart (already filtered by server)
        const horsesInChart = result.horsesFound || [];

        // Build a map of chart horses to their system equivalents (if any)
        const chartToSystemMap = new Map();
        for (const chartName of horsesInChart) {
            const systemMatch = this.existingHorses.find(h =>
                h.name.toLowerCase() === chartName.toLowerCase()
            );
            if (systemMatch) {
                chartToSystemMap.set(chartName, systemMatch.name);
            }
        }

        // Determine the best horse to auto-select:
        // 1. If we have a fuzzy match from the system, use that (the system horse name)
        // 2. Otherwise use the chart horse name (or its system equivalent)
        let autoSelectHorse = '';
        if (matchedName && confidence >= 0.6) {
            autoSelectHorse = matchedName;
        } else if (chartHorse) {
            // Use system name if available, otherwise chart name
            autoSelectHorse = chartToSystemMap.get(chartHorse) || chartHorse;
        }

        let options = '';

        // First add horses found in this chart (use system names when available)
        if (horsesInChart.length > 0) {
            options += '<optgroup label="Horses in this chart">';
            options += horsesInChart.map(chartName => {
                // Use system horse name if it exists, otherwise use chart name
                const valueName = chartToSystemMap.get(chartName) || chartName;
                const displayName = chartToSystemMap.has(chartName)
                    ? `${valueName} (from chart: ${chartName})`
                    : chartName;
                const isSelected = valueName.toLowerCase() === autoSelectHorse.toLowerCase();
                return `<option value="${valueName}" ${isSelected ? 'selected' : ''}>${displayName}</option>`;
            }).join('');
            options += '</optgroup>';
        }

        // Then add existing horses from the system (if not already matched to chart)
        if (this.existingHorses.length > 0) {
            const existingNotInChart = this.existingHorses.filter(h =>
                !horsesInChart.some(c => c.toLowerCase() === h.name.toLowerCase())
            );
            if (existingNotInChart.length > 0) {
                options += '<optgroup label="Other horses in system">';
                options += existingNotInChart.map(h => {
                    const isSelected = h.name.toLowerCase() === autoSelectHorse.toLowerCase();
                    return `<option value="${h.name}" ${isSelected ? 'selected' : ''}>${h.displayName || h.name}</option>`;
                }).join('');
                options += '</optgroup>';
            }
        }

        // Add "Create New" option
        options += `<option value="__NEW__">+ Create New Horse</option>`;

        return `
            <div class="horse-selector ${isVerified ? 'verified' : 'unverified'}">
                <select class="review-input horse-select" data-index="${index}" data-field="horseName"
                        onchange="raceUploader.handleHorseChange(this, ${index})" ${result.isDuplicate ? 'disabled' : ''}>
                    ${options}
                </select>
                ${confidence > 0 && confidence < 0.9 ?
                    `<span class="confidence-badge">${Math.round(confidence * 100)}% match</span>` : ''}
                <div id="newHorseFields${index}" class="new-horse-fields" style="display: none;">
                    <input type="text" placeholder="Horse Name" class="review-input"
                           data-index="${index}" data-field="newHorseName">
                    <input type="text" placeholder="Owner" class="review-input"
                           data-index="${index}" data-field="newHorseOwner">
                    <input type="text" placeholder="Country" class="review-input"
                           data-index="${index}" data-field="newHorseCountry">
                </div>
            </div>
        `;
    }

    buildManualEntryForm(index) {
        return `
            <div class="review-card-data">
                <div class="data-row">
                    <div class="data-field">
                        <label>Horse</label>
                        ${this.buildHorseSelector({ matchedHorse: null, matchConfidence: 0, data: {} }, index)}
                    </div>
                    <div class="data-field">
                        <label>Date</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="date" placeholder="MM/DD/YYYY">
                    </div>
                    <div class="data-field">
                        <label>Track</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="track" placeholder="KEE, SAR, etc.">
                    </div>
                    <div class="data-field">
                        <label>Surface</label>
                        <select class="review-input" data-index="${index}" data-field="surface">
                            <option value="D">Dirt</option>
                            <option value="T">Turf</option>
                            <option value="AWT">AWT</option>
                        </select>
                    </div>
                </div>
                <div class="data-row">
                    <div class="data-field">
                        <label>Distance</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="distance" placeholder="6F">
                    </div>
                    <div class="data-field">
                        <label>Race Type</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="raceType" placeholder="MSW, Stakes, etc.">
                    </div>
                    <div class="data-field">
                        <label>Final Time</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="finalTime" placeholder="1:12.45">
                    </div>
                    <div class="data-field">
                        <label>Finish Position</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="finishPosition" placeholder="1st, 3rd, etc.">
                    </div>
                </div>
                <div class="data-row">
                    <div class="data-field full-width">
                        <label>Comments</label>
                        <input type="text" class="review-input" data-index="${index}" data-field="comments" placeholder="Race comments...">
                    </div>
                </div>
            </div>
        `;
    }

    handleHorseChange(select, index) {
        const newHorseFields = document.getElementById(`newHorseFields${index}`);
        if (select.value === '__NEW__') {
            newHorseFields.style.display = 'block';
        } else {
            newHorseFields.style.display = 'none';

            // Update form fields with the selected horse's data
            const result = this.results[index];
            const selectedHorseName = select.value;

            // Check if we have data for this horse
            if (result.allHorseData && result.allHorseData[selectedHorseName]) {
                const horseData = result.allHorseData[selectedHorseName];
                this.updateFormFields(index, horseData);
            }
        }
        this.updateSaveButton();
    }

    updateFormFields(index, data) {
        const setValue = (field, value) => {
            const input = document.querySelector(`.review-input[data-index="${index}"][data-field="${field}"]`);
            if (input) {
                input.value = value || '';
            }
        };

        setValue('date', this.formatDisplayDate(data.raceDate));
        setValue('track', data.track);
        setValue('distance', data.distance);
        setValue('raceType', data.raceType);
        setValue('finalTime', data.finalTime);
        setValue('avgSpeed', data.avgSpeedMph ? data.avgSpeedMph.toFixed(1) : '');
        setValue('finishPosition', data.posFin);
        setValue('fiveFReduction', data.fiveFReductionTime);
        setValue('pos1_4', data.pos1_4);
        setValue('f1Time', data.f1Time);
        setValue('pos1_2', data.pos1_2);
        setValue('f2Time', data.f2Time);
        setValue('pos3_4', data.pos3_4);
        setValue('f3Time', data.f3Time);
        setValue('comments', data.comment);

        // Update surface select
        const surfaceSelect = document.querySelector(`.review-input[data-index="${index}"][data-field="surface"]`);
        if (surfaceSelect && data.surface) {
            surfaceSelect.value = data.surface;
        }
    }

    showManualEntry(index) {
        const manualEntry = document.getElementById(`manualEntry${index}`);
        manualEntry.style.display = 'block';

        // Mark this result as having manual entry
        this.results[index].hasManualEntry = true;

        // Check the include checkbox
        const checkbox = document.querySelector(`.include-race[data-index="${index}"]`);
        if (checkbox) checkbox.checked = true;

        this.updateSaveButton();
    }

    formatDisplayDate(dateStr) {
        if (!dateStr || dateStr === 'Unknown Date') return '';
        try {
            const date = new Date(dateStr);
            if (isNaN(date.getTime())) return dateStr;
            return `${date.getMonth() + 1}/${date.getDate()}/${date.getFullYear()}`;
        } catch (e) {
            return dateStr;
        }
    }

    updateSaveButton() {
        const checkboxes = document.querySelectorAll('.include-race:checked');
        const saveBtn = document.getElementById('saveAllRaces');
        saveBtn.disabled = checkboxes.length === 0;
        saveBtn.textContent = `Save ${checkboxes.length} Race${checkboxes.length !== 1 ? 's' : ''}`;
    }

    goBackToUpload() {
        document.getElementById('raceUploadStep3').style.display = 'none';
        document.getElementById('raceUploadStep2').style.display = 'none';
        document.getElementById('raceUploadStep1').style.display = 'block';
    }

    collectRaceData() {
        const races = [];
        const checkboxes = document.querySelectorAll('.include-race:checked');

        checkboxes.forEach(checkbox => {
            const index = parseInt(checkbox.dataset.index);
            const result = this.results[index];

            // Get all field values from inputs
            const getValue = (field) => {
                const input = document.querySelector(`.review-input[data-index="${index}"][data-field="${field}"]`);
                return input ? input.value : '';
            };

            const horseName = getValue('horseName');
            let finalHorseName = horseName;
            let isNewHorse = false;
            let owner = '';
            let country = '';

            // Check if creating new horse
            if (horseName === '__NEW__') {
                finalHorseName = getValue('newHorseName');
                owner = getValue('newHorseOwner');
                country = getValue('newHorseCountry');
                isNewHorse = true;

                if (!finalHorseName) {
                    alert(`Please enter a name for the new horse in ${result.fileName}`);
                    return;
                }
            }

            races.push({
                horseName: finalHorseName,
                date: getValue('date'),
                track: getValue('track'),
                surface: getValue('surface'),
                distance: getValue('distance'),
                raceType: getValue('raceType'),
                finalTime: getValue('finalTime'),
                avgSpeed: getValue('avgSpeed'),
                finishPosition: getValue('finishPosition'),
                fiveFReduction: getValue('fiveFReduction'),
                pos1_4: getValue('pos1_4'),
                f1Time: getValue('f1Time'),
                pos1_2: getValue('pos1_2'),
                f2Time: getValue('f2Time'),
                pos3_4: getValue('pos3_4'),
                f3Time: getValue('f3Time'),
                comments: getValue('comments'),
                isNewHorse,
                owner,
                country
            });
        });

        return races;
    }

    async saveAllRaces() {
        const races = this.collectRaceData();

        if (races.length === 0) {
            alert('No races selected to save');
            return;
        }

        const saveBtn = document.getElementById('saveAllRaces');
        saveBtn.disabled = true;
        saveBtn.textContent = 'Saving...';

        try {
            const response = await fetch('/api/race-charts/save', {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify({ races })
            });

            if (!response.ok) {
                throw new Error(`Server error: ${response.status}`);
            }

            const data = await response.json();

            if (data.success) {
                let message = `Successfully saved ${data.savedCount} race${data.savedCount !== 1 ? 's' : ''}!`;
                if (data.skippedCount > 0) {
                    message += `\n\nSkipped ${data.skippedCount} duplicate${data.skippedCount !== 1 ? 's' : ''}.`;
                }
                alert(message);

                // Close modal and refresh data
                this.closeModal();

                // Trigger data refresh
                if (typeof loadLatestSession === 'function') {
                    loadLatestSession();
                }
            } else {
                throw new Error(data.error || 'Unknown error');
            }

        } catch (error) {
            console.error('Error saving races:', error);
            alert('Error saving races: ' + error.message);
            saveBtn.disabled = false;
            saveBtn.textContent = `Save ${races.length} Races`;
        }
    }
}

// Initialize the uploader
let raceUploader;
document.addEventListener('DOMContentLoaded', () => {
    raceUploader = new RaceChartUploader();
});

// Global function to open the modal (called from button)
function openRaceUploadModal() {
    if (raceUploader) {
        raceUploader.openModal();
    }
}
