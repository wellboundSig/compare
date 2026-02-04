/**
 * DiffMaker - Wellbound Difference Report Tool
 * Advanced row-aware, column-aware diff comparison
 */

// ========================================
// Global State
// ========================================

const state = {
    file1: null,
    file2: null,
    data1: null,
    data2: null,
    headers: [],
    primaryKeys: [],
    diffResult: null,
    currentPage: 1,
    rowsPerPage: 50,
    currentFilter: 'all',
    currentDensity: 'standard',
    currentGranularity: 'summary',
    chart: null,
    columnFiltersVisible: false,
    columnFilters: {},  // { columnName: filterValue }
    showMovedRows: false,  // Whether to display moved/reordered rows
    // Highlight view specific state
    highlightSyncScroll: false,  // Whether to sync scroll both sides
    highlightColumnFilter: [],   // Array of column names to filter changes by (empty = all)
    // Color settings for persistence
    highlightColors: {
        modified: '#f59e0b',
        added: '#22c55e',
        removed: '#ef4444',
        unchanged: '#6b7280'
    },
    // Type filters for highlight view
    highlightTypeFilters: {
        modified: true,
        added: true,
        removed: true,
        unchanged: true
    }
};

// ========================================
// Initialization
// ========================================

document.addEventListener('DOMContentLoaded', () => {
    initializeNavigation();
    initializeFileUploads();
    initializeConfigControls();
    initializeExportControls();
    initializeViewerControls();
    initializeAnalyticsControls();
    initializeModals();
});

// ========================================
// Navigation
// ========================================

function initializeNavigation() {
    const navTabs = document.querySelectorAll('.nav-tab');
    navTabs.forEach(tab => {
        tab.addEventListener('click', () => {
            const pageId = tab.dataset.page;
            switchPage(pageId);
            navTabs.forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // If switching to viewer, update the content
            if (pageId === 'viewer') {
                updateViewerContent();
            }
        });
    });

    document.getElementById('goToGenerator')?.addEventListener('click', () => {
        switchPage('generator');
        document.querySelectorAll('.nav-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.page === 'generator');
        });
    });

    document.getElementById('openViewer')?.addEventListener('click', () => {
        switchPage('viewer');
        document.querySelectorAll('.nav-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.page === 'viewer');
        });
        updateViewerContent();
    });
}

function switchPage(pageId) {
    document.querySelectorAll('.page').forEach(page => {
        page.classList.remove('active');
    });
    document.getElementById(`${pageId}-page`)?.classList.add('active');
}

// ========================================
// File Upload Handling - Unified
// ========================================

function initializeFileUploads() {
    const dropzone = document.getElementById('unified-dropzone');
    const fileInput = document.getElementById('file-input');

    // File input change
    fileInput.addEventListener('change', handleFilesSelect);

    // Drag and drop
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });

    dropzone.addEventListener('dragleave', (e) => {
        // Only remove if leaving the dropzone entirely
        if (!dropzone.contains(e.relatedTarget)) {
            dropzone.classList.remove('dragover');
        }
    });

    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const files = Array.from(e.dataTransfer.files);
        processMultipleFiles(files);
    });

    // Click to upload (only in empty state)
    dropzone.addEventListener('click', (e) => {
        if (!e.target.closest('button') && !e.target.closest('.file-card-remove')) {
            const isEmpty = !state.file1 && !state.file2;
            if (isEmpty) {
                fileInput.click();
            }
        }
    });

    // Remove file buttons
    document.getElementById('remove-file1')?.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(1);
    });
    document.getElementById('remove-file2')?.addEventListener('click', (e) => {
        e.stopPropagation();
        removeFile(2);
    });
    
    // Swap files button
    document.getElementById('swap-files')?.addEventListener('click', (e) => {
        e.stopPropagation();
        swapFiles();
    });
}

function handleFilesSelect(e) {
    const files = Array.from(e.target.files);
    processMultipleFiles(files);
    // Reset input so the same files can be selected again
    e.target.value = '';
}

async function processMultipleFiles(files) {
    if (files.length === 0) return;

    showLoading(true);

    try {
        // If we have 2+ files, use first two - but auto-detect which is older
        if (files.length >= 2) {
            const file1 = files[0];
            const file2 = files[1];
            
            // Check lastModified to determine which is original (older)
            let originalFile, updatedFile;
            
            if (file1.lastModified <= file2.lastModified) {
                // file1 is older or same age, treat as original
                originalFile = file1;
                updatedFile = file2;
            } else {
                // file2 is older, swap them
                originalFile = file2;
                updatedFile = file1;
            }
            
            await processFile(originalFile, 1);
            await processFile(updatedFile, 2);
        } else if (files.length === 1) {
            // Single file - assign to first empty slot
            if (!state.file1) {
                await processFile(files[0], 1);
            } else if (!state.file2) {
                await processFile(files[0], 2);
            } else {
                // Both filled, replace file1
                await processFile(files[0], 1);
            }
        }
    } catch (error) {
        alert(`Error processing files: ${error.message}`);
    }

    showLoading(false);
}

function swapFiles() {
    // Swap file objects
    const tempFile = state.file1;
    const tempData = state.data1;
    
    state.file1 = state.file2;
    state.data1 = state.data2;
    
    state.file2 = tempFile;
    state.data2 = tempData;
    
    // Update display with animation
    const card1 = document.getElementById('file1-card');
    const card2 = document.getElementById('file2-card');
    
    card1.style.animation = 'swapLeft 0.3s ease';
    card2.style.animation = 'swapRight 0.3s ease';
    
    setTimeout(() => {
        updateFileDisplay();
        card1.style.animation = '';
        card2.style.animation = '';
    }, 300);
    
    // Hide config and results since files changed
    document.getElementById('config-section').hidden = true;
    document.getElementById('results-section').hidden = true;
    
    // Re-show config if both files present
    if (state.data1 && state.data2) {
        setTimeout(() => showConfigSection(), 350);
    }
}

async function processFile(file, fileNum) {
    try {
        const data = await parseFile(file);
        
        if (fileNum === 1) {
            state.file1 = file;
            state.data1 = data;
        } else {
            state.file2 = file;
            state.data2 = data;
        }

        updateFileDisplay();
        
        // If both files are loaded, show config
        if (state.data1 && state.data2) {
            showConfigSection();
        }
    } catch (error) {
        throw error;
    }
}

async function parseFile(file) {
    const extension = file.name.split('.').pop().toLowerCase();
    
    if (extension === 'csv') {
        try {
            return await parseCSV(file);
        } catch (error) {
            // Fallback: try reading as text and cleaning up
            console.warn('Standard CSV parsing failed, trying fallback:', error.message);
            return await parseCSVFallback(file);
        }
    } else if (['xlsx', 'xls'].includes(extension)) {
        return await parseExcel(file);
    } else {
        throw new Error('Unsupported file format. Please use CSV or Excel files.');
    }
}

async function parseCSVFallback(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                let text = e.target.result;
                
                // Clean up common issues
                // Remove BOM if present
                text = text.replace(/^\uFEFF/, '');
                // Normalize line endings
                text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
                // Fix trailing quotes issues by removing problematic quotes
                text = text.replace(/"([^"]*)"(?=[^,\n])/g, '$1');
                
                Papa.parse(text, {
                    header: true,
                    skipEmptyLines: true,
                    dynamicTyping: false,
                    complete: (results) => {
                        if (results.data.length === 0) {
                            reject(new Error('No data found in file after cleanup'));
                        } else {
                            const cleanData = results.data.filter(row => {
                                const values = Object.values(row);
                                return values.some(v => v !== null && v !== undefined && String(v).trim() !== '');
                            });
                            resolve(cleanData);
                        }
                    },
                    error: (error) => reject(error)
                });
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsText(file);
    });
}

function parseCSV(file) {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            dynamicTyping: false,
            delimitersToGuess: [',', '\t', '|', ';'],
            complete: (results) => {
                // Only reject on critical errors, not warnings
                const criticalErrors = results.errors.filter(e => 
                    e.type === 'FieldMismatch' || 
                    (e.type === 'Quotes' && results.data.length === 0)
                );
                
                if (criticalErrors.length > 0 && results.data.length === 0) {
                    reject(new Error(criticalErrors[0].message));
                } else if (results.data.length === 0) {
                    reject(new Error('No data found in file'));
                } else {
                    // Filter out any empty rows
                    const cleanData = results.data.filter(row => {
                        const values = Object.values(row);
                        return values.some(v => v !== null && v !== undefined && String(v).trim() !== '');
                    });
                    resolve(cleanData);
                }
            },
            error: (error) => reject(error)
        });
    });
}

function parseExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(e.target.result, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const data = XLSX.utils.sheet_to_json(worksheet);
                resolve(data);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

function updateFileDisplay() {
    const dropzone = document.getElementById('unified-dropzone');
    const emptyState = document.getElementById('upload-empty');
    const filesState = document.getElementById('upload-files');
    const bundleSuccess = document.getElementById('bundle-success');
    
    const hasAnyFile = state.file1 || state.file2;
    
    if (hasAnyFile) {
        dropzone.classList.add('has-files');
        emptyState.hidden = true;
        filesState.hidden = false;
        if (bundleSuccess) bundleSuccess.hidden = true;
        
        // Update file 1 card
        const file1Card = document.getElementById('file1-card');
        const file1Name = document.getElementById('file1-name');
        const file1Meta = document.getElementById('file1-meta');
        const removeBtn1 = document.getElementById('remove-file1');
        
        if (state.file1) {
            file1Card.classList.add('loaded');
            file1Name.textContent = state.file1.name;
            const date1 = formatFileDate(state.file1.lastModified);
            file1Meta.innerHTML = `<span class="rows">${state.data1.length} rows</span><span class="date">${date1}</span>`;
            removeBtn1.hidden = false;
        } else {
            file1Card.classList.remove('loaded');
            file1Name.textContent = 'No file selected';
            file1Meta.innerHTML = '';
            removeBtn1.hidden = true;
        }
        
        // Update file 2 card
        const file2Card = document.getElementById('file2-card');
        const file2Name = document.getElementById('file2-name');
        const file2Meta = document.getElementById('file2-meta');
        const removeBtn2 = document.getElementById('remove-file2');
        
        if (state.file2) {
            file2Card.classList.add('loaded');
            file2Name.textContent = state.file2.name;
            const date2 = formatFileDate(state.file2.lastModified);
            file2Meta.innerHTML = `<span class="rows">${state.data2.length} rows</span><span class="date">${date2}</span>`;
            removeBtn2.hidden = false;
        } else {
            file2Card.classList.remove('loaded');
            file2Name.textContent = 'No file selected';
            file2Meta.innerHTML = '';
            removeBtn2.hidden = true;
        }
        
        // Show/hide swap button based on whether both files are loaded
        const swapBtn = document.getElementById('swap-files');
        if (swapBtn) {
            swapBtn.style.display = (state.file1 && state.file2) ? 'flex' : 'none';
        }
    } else {
        dropzone.classList.remove('has-files');
        emptyState.hidden = false;
        filesState.hidden = true;
    }
}

function formatFileDate(timestamp) {
    const date = new Date(timestamp);
    const options = { month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' };
    return date.toLocaleDateString('en-US', options);
}

function removeFile(fileNum) {
    if (fileNum === 1) {
        state.file1 = null;
        state.data1 = null;
    } else {
        state.file2 = null;
        state.data2 = null;
    }
    
    updateFileDisplay();
    document.getElementById('config-section').hidden = true;
    document.getElementById('results-section').hidden = true;
}

// ========================================
// Configuration
// ========================================

function showConfigSection() {
    const configSection = document.getElementById('config-section');
    configSection.hidden = false;
    
    // Detect common headers
    const headers1 = Object.keys(state.data1[0] || {});
    const headers2 = Object.keys(state.data2[0] || {});
    state.headers = headers1.filter(h => headers2.includes(h));
    
    populateKeySelector();
}

function populateKeySelector() {
    const selector = document.getElementById('key-selector');
    selector.innerHTML = '';
    
    state.headers.forEach((header, index) => {
        const chip = document.createElement('button');
        chip.className = 'key-chip';
        chip.dataset.column = header;
        chip.innerHTML = `
            <span class="key-number" hidden>${index + 1}</span>
            ${escapeHtml(header)}
        `;
        chip.addEventListener('click', () => toggleKeyColumn(chip, header));
        selector.appendChild(chip);
    });
    
    // Auto-detect primary key
    if (document.getElementById('autoDetectKey').checked) {
        autoDetectPrimaryKey();
    }
}

function toggleKeyColumn(chip, column) {
    document.getElementById('autoDetectKey').checked = false;
    
    if (chip.classList.contains('selected')) {
        chip.classList.remove('selected');
        chip.querySelector('.key-number').hidden = true;
        state.primaryKeys = state.primaryKeys.filter(k => k !== column);
    } else {
        chip.classList.add('selected');
        state.primaryKeys.push(column);
    }
    
    // Update key numbers
    document.querySelectorAll('.key-chip.selected').forEach((c, i) => {
        const num = c.querySelector('.key-number');
        num.textContent = i + 1;
        num.hidden = false;
    });
}

function autoDetectPrimaryKey() {
    state.primaryKeys = [];
    
    // Look for common key column names
    const keyPatterns = ['id', 'key', 'code', 'number', 'identifier', 'uuid', 'guid'];
    
    for (const header of state.headers) {
        const lowerHeader = header.toLowerCase();
        if (keyPatterns.some(p => lowerHeader.includes(p))) {
            // Check if values are unique
            const values1 = state.data1.map(row => row[header]);
            const unique1 = new Set(values1);
            if (unique1.size === values1.length) {
                state.primaryKeys = [header];
                break;
            }
        }
    }
    
    // Fallback: use first column if no key found
    if (state.primaryKeys.length === 0 && state.headers.length > 0) {
        state.primaryKeys = [state.headers[0]];
    }
    
    // Update UI
    document.querySelectorAll('.key-chip').forEach(chip => {
        const isSelected = state.primaryKeys.includes(chip.dataset.column);
        chip.classList.toggle('selected', isSelected);
        const num = chip.querySelector('.key-number');
        if (isSelected) {
            num.textContent = '1';
            num.hidden = false;
        } else {
            num.hidden = true;
        }
    });
}

function initializeConfigControls() {
    document.getElementById('autoDetectKey').addEventListener('change', (e) => {
        if (e.target.checked) {
            autoDetectPrimaryKey();
        }
    });
    
    document.getElementById('compareBtn').addEventListener('click', runComparison);
}

// ========================================
// Comparison Engine
// ========================================

function runComparison() {
    if (state.primaryKeys.length === 0) {
        alert('Please select at least one column as primary key');
        return;
    }
    
    showLoading(true);
    
    setTimeout(() => {
        try {
            const options = {
                ignoreCase: document.getElementById('ignoreCase').checked,
                ignoreWhitespace: document.getElementById('ignoreWhitespace').checked,
                treatReorderAsSame: document.getElementById('treatReorderAsSame').checked,
                typeAware: document.getElementById('typeAware').checked
            };
            
            // Update state for moved rows visibility
            state.showMovedRows = document.getElementById('showMovedRows').checked;
            
            state.diffResult = performDiff(state.data1, state.data2, state.primaryKeys, options);
            displayResults();
        } catch (error) {
            alert(`Comparison error: ${error.message}`);
        }
        
        showLoading(false);
    }, 100);
}

function performDiff(data1, data2, primaryKeys, options) {
    const result = {
        unchanged: [],
        modified: [],
        added: [],
        removed: [],
        moved: [],
        columnChanges: {},
        // New: Found Status and Status Match tracking
        foundStatus: {
            found: 0,
            notFound: 0
        },
        statusMatch: {
            matched: 0,
            changed: 0
        },
        meta: {
            file1Name: state.file1.name,
            file2Name: state.file2.name,
            comparedAt: new Date().toISOString(),
            primaryKeys,
            options,
            originalRowCount: data1.length,
            updatedRowCount: data2.length
        }
    };
    
    // Initialize column changes tracking
    state.headers.forEach(h => {
        result.columnChanges[h] = 0;
    });
    
    // Create lookup maps using composite keys
    const map1 = createKeyMap(data1, primaryKeys);
    const map2 = createKeyMap(data2, primaryKeys);
    
    // Track positions for reorder detection
    const positions1 = new Map();
    const positions2 = new Map();
    
    data1.forEach((row, idx) => {
        positions1.set(getCompositeKey(row, primaryKeys), idx);
    });
    
    data2.forEach((row, idx) => {
        positions2.set(getCompositeKey(row, primaryKeys), idx);
    });
    
    // Find removed and modified rows
    for (const [key, row1] of map1) {
        if (map2.has(key)) {
            // Record is FOUND in both files
            result.foundStatus.found++;
            
            const row2 = map2.get(key);
            const changes = compareRows(row1, row2, state.headers, options);
            
            if (changes.length > 0) {
                // Has changes = Changed
                result.statusMatch.changed++;
                result.modified.push({
                    key,
                    original: row1,
                    updated: row2,
                    changes,
                    originalPosition: positions1.get(key),
                    updatedPosition: positions2.get(key)
                });
                
                // Track column changes
                changes.forEach(c => {
                    result.columnChanges[c.column]++;
                });
            } else {
                // Check if position changed
                const pos1 = positions1.get(key);
                const pos2 = positions2.get(key);
                
                if (options.treatReorderAsSame && pos1 !== pos2) {
                    // Moved but no data changes = Matched
                    result.statusMatch.matched++;
                    result.moved.push({
                        key,
                        data: row1,
                        originalPosition: pos1,
                        updatedPosition: pos2
                    });
                } else if (!options.treatReorderAsSame && pos1 !== pos2) {
                    // If not treating reorder as same, it counts as modified
                    result.modified.push({
                        key,
                        original: row1,
                        updated: row2,
                        changes: [{ column: '_position', oldValue: pos1, newValue: pos2 }],
                        originalPosition: pos1,
                        updatedPosition: pos2
                    });
                } else {
                    // No changes = Matched
                    result.statusMatch.matched++;
                    result.unchanged.push({
                        key,
                        data: row1,
                        position: pos1
                    });
                }
            }
        } else {
            // Record NOT FOUND in second file
            result.foundStatus.notFound++;
            result.removed.push({
                key,
                data: row1,
                position: positions1.get(key)
            });
        }
    }
    
    // Find added rows (these are in file2 but not in file1)
    for (const [key, row2] of map2) {
        if (!map1.has(key)) {
            result.added.push({
                key,
                data: row2,
                position: positions2.get(key)
            });
        }
    }
    
    return result;
}

// Normalize status values: 6 and blank/empty both mean "inactive"
function createKeyMap(data, primaryKeys) {
    const map = new Map();
    data.forEach(row => {
        const key = getCompositeKey(row, primaryKeys);
        map.set(key, row);
    });
    return map;
}

function getCompositeKey(row, primaryKeys) {
    return primaryKeys.map(k => String(row[k] || '')).join('|||');
}

function compareRows(row1, row2, headers, options) {
    const changes = [];
    
    headers.forEach(header => {
        let val1 = row1[header];
        let val2 = row2[header];
        
        // Apply comparison options
        if (options.ignoreWhitespace) {
            val1 = String(val1 || '').trim();
            val2 = String(val2 || '').trim();
        }
        
        if (options.ignoreCase) {
            val1 = String(val1 || '').toLowerCase();
            val2 = String(val2 || '').toLowerCase();
        }
        
        let isEqual = false;
        
        if (options.typeAware) {
            isEqual = typeAwareCompare(val1, val2);
        } else {
            isEqual = String(val1) === String(val2);
        }
        
        if (!isEqual) {
            changes.push({
                column: header,
                oldValue: row1[header],
                newValue: row2[header]
            });
        }
    });
    
    return changes;
}

function typeAwareCompare(val1, val2) {
    // Handle null/undefined
    if (val1 == null && val2 == null) return true;
    if (val1 == null || val2 == null) return false;
    
    // Try numeric comparison
    const num1 = parseFloat(val1);
    const num2 = parseFloat(val2);
    if (!isNaN(num1) && !isNaN(num2)) {
        return Math.abs(num1 - num2) < 0.0001;
    }
    
    // Try date comparison
    const date1 = new Date(val1);
    const date2 = new Date(val2);
    if (!isNaN(date1.getTime()) && !isNaN(date2.getTime())) {
        return date1.getTime() === date2.getTime();
    }
    
    // String comparison
    return String(val1) === String(val2);
}

// ========================================
// Results Display
// ========================================

function displayResults() {
    const resultsSection = document.getElementById('results-section');
    resultsSection.hidden = false;
    
    // Update Key Metrics - Found Status
    const found = state.diffResult.foundStatus.found;
    const notFound = state.diffResult.foundStatus.notFound;
    const totalForFound = found + notFound;
    
    document.getElementById('found-count').textContent = found;
    document.getElementById('not-found-count').textContent = notFound;
    
    const foundPercent = totalForFound > 0 ? (found / totalForFound * 100) : 0;
    document.getElementById('found-bar').style.width = `${foundPercent}%`;
    
    // Update Key Metrics - Status Match
    const matched = state.diffResult.statusMatch.matched;
    const changed = state.diffResult.statusMatch.changed;
    const totalForMatch = matched + changed;
    
    document.getElementById('status-match-count').textContent = matched;
    document.getElementById('status-mismatch-count').textContent = changed;
    
    const matchPercent = totalForMatch > 0 ? (matched / totalForMatch * 100) : 0;
    document.getElementById('match-bar').style.width = `${matchPercent}%`;
    
    // Update summary counts
    // When showMovedRows is off, add moved count to unchanged
    const displayUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + state.diffResult.moved.length;
    
    document.getElementById('unchanged-count').textContent = displayUnchangedCount;
    document.getElementById('modified-count').textContent = state.diffResult.modified.length;
    document.getElementById('added-count').textContent = state.diffResult.added.length;
    document.getElementById('removed-count').textContent = state.diffResult.removed.length;
    document.getElementById('moved-count').textContent = state.diffResult.moved.length;
    
    // Show/hide moved summary card based on setting
    const movedCard = document.querySelector('.summary-moved');
    if (movedCard) {
        movedCard.style.display = state.showMovedRows ? '' : 'none';
    }
    
    // Show/hide "Moved Only" option in preview filter
    const previewFilter = document.getElementById('previewFilter');
    const movedOption = previewFilter?.querySelector('option[value="moved"]');
    if (movedOption) {
        movedOption.style.display = state.showMovedRows ? '' : 'none';
        // If currently filtering by moved but moved is hidden, reset to all
        if (!state.showMovedRows && previewFilter.value === 'moved') {
            previewFilter.value = 'all';
        }
    }
    
    // Update preview
    updatePreviewTable();
    
    // Scroll to results
    resultsSection.scrollIntoView({ behavior: 'smooth' });
}

function updatePreviewTable() {
    const filter = document.getElementById('previewFilter').value;
    const table = document.getElementById('preview-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    
    // Build header
    thead.innerHTML = `
        <tr>
            <th>Status</th>
            ${state.headers.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
        </tr>
    `;
    
    // Collect rows based on filter
    let rows = [];
    
    if (filter === 'all' || filter === 'modified') {
        state.diffResult.modified.forEach(item => {
            rows.push({ status: 'modified', data: item.updated, changes: item.changes, original: item.original });
        });
    }
    
    if (filter === 'all' || filter === 'added') {
        state.diffResult.added.forEach(item => {
            rows.push({ status: 'added', data: item.data });
        });
    }
    
    if (filter === 'all' || filter === 'removed') {
        state.diffResult.removed.forEach(item => {
            rows.push({ status: 'removed', data: item.data });
        });
    }
    
    // Only show moved rows if the setting is enabled
    if (state.showMovedRows && (filter === 'all' || filter === 'moved')) {
        state.diffResult.moved.forEach(item => {
            rows.push({ status: 'moved', data: item.data, originalPos: item.originalPosition, newPos: item.updatedPosition });
        });
    }
    
    // Build body (limit to first 100 for preview)
    tbody.innerHTML = rows.slice(0, 100).map(row => {
        const changedColumns = row.changes ? row.changes.map(c => c.column) : [];
        
        return `
            <tr>
                <td><span class="row-status ${row.status}">${row.status}</span></td>
                ${state.headers.map(h => {
                    const value = row.data[h] ?? '';
                    if (changedColumns.includes(h)) {
                        const change = row.changes.find(c => c.column === h);
                        return `<td class="cell-changed">
                            <span class="cell-old">${escapeHtml(String(change.oldValue ?? ''))}</span>
                            <span class="cell-new">${escapeHtml(String(change.newValue ?? ''))}</span>
                        </td>`;
                    }
                    return `<td>${escapeHtml(String(value))}</td>`;
                }).join('')}
            </tr>
        `;
    }).join('');
    
    if (rows.length > 100) {
        tbody.innerHTML += `
            <tr>
                <td colspan="${state.headers.length + 1}" style="text-align: center; color: var(--text-light);">
                    ... and ${rows.length - 100} more rows
                </td>
            </tr>
        `;
    }
}

// ========================================
// Export Functions
// ========================================

function initializeExportControls() {
    // Format buttons
    document.querySelectorAll('.format-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.format-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
        });
    });
    
    // Preview filter
    document.getElementById('previewFilter')?.addEventListener('change', updatePreviewTable);
    
    // Download button
    document.getElementById('downloadReport')?.addEventListener('click', downloadReport);
    
    // Export diff view file
    document.getElementById('exportDiffView')?.addEventListener('click', exportDiffViewFile);
}

async function downloadReport() {
    const contentType = document.querySelector('input[name="exportContent"]:checked').value;
    const format = document.querySelector('.format-btn.active').dataset.format;
    
    showLoading(true);
    
    try {
        let data = [];
        let title = '';
        
        switch (contentType) {
            case 'changed':
                data = prepareChangedRowsData();
                title = 'Changed Rows Report';
                break;
            case 'all':
                data = prepareAllRowsData();
                title = 'Full Comparison Report';
                break;
            case 'unchanged':
                data = prepareUnchangedRowsData();
                title = 'Unchanged Rows Report';
                break;
        }
        
        switch (format) {
            case 'csv':
                downloadCSV(data, title);
                break;
            case 'xlsx':
                downloadExcel(data, title);
                break;
            case 'pdf':
                await downloadPDF(data, title);
                break;
        }
    } catch (error) {
        alert(`Export error: ${error.message}`);
    }
    
    showLoading(false);
}

function prepareChangedRowsData() {
    const rows = [];
    
    state.diffResult.modified.forEach(item => {
        const row = { _status: 'MODIFIED', ...item.updated };
        item.changes.forEach(c => {
            row[`_old_${c.column}`] = c.oldValue;
        });
        rows.push(row);
    });
    
    state.diffResult.added.forEach(item => {
        rows.push({ _status: 'ADDED', ...item.data });
    });
    
    state.diffResult.removed.forEach(item => {
        rows.push({ _status: 'REMOVED', ...item.data });
    });
    
    // Only include moved rows if the setting is enabled
    if (state.showMovedRows) {
        state.diffResult.moved.forEach(item => {
            rows.push({ 
                _status: 'MOVED', 
                _from_position: item.originalPosition + 1,
                _to_position: item.updatedPosition + 1,
                ...item.data 
            });
        });
    }
    
    return rows;
}

function prepareAllRowsData() {
    const rows = [];
    
    state.diffResult.unchanged.forEach(item => {
        rows.push({ _status: 'UNCHANGED', ...item.data });
    });
    
    // When showMovedRows is off, treat moved as unchanged
    if (!state.showMovedRows) {
        state.diffResult.moved.forEach(item => {
            rows.push({ _status: 'UNCHANGED', ...item.data });
        });
    }
    
    state.diffResult.modified.forEach(item => {
        const row = { _status: 'MODIFIED', ...item.updated };
        item.changes.forEach(c => {
            row[`_old_${c.column}`] = c.oldValue;
        });
        rows.push(row);
    });
    
    state.diffResult.added.forEach(item => {
        rows.push({ _status: 'ADDED', ...item.data });
    });
    
    state.diffResult.removed.forEach(item => {
        rows.push({ _status: 'REMOVED', ...item.data });
    });
    
    // Only include moved rows with MOVED status if setting is enabled
    if (state.showMovedRows) {
        state.diffResult.moved.forEach(item => {
            rows.push({ 
                _status: 'MOVED', 
                ...item.data 
            });
        });
    }
    
    return rows;
}

function prepareUnchangedRowsData() {
    const rows = state.diffResult.unchanged.map(item => item.data);
    
    // When showMovedRows is off, include moved rows as unchanged
    if (!state.showMovedRows) {
        state.diffResult.moved.forEach(item => {
            rows.push(item.data);
        });
    }
    
    return rows;
}

function downloadCSV(data, title) {
    const csv = Papa.unparse(data);
    const header = generateReportHeader(title);
    const content = header + '\n\n' + csv;
    
    downloadFile(content, `${sanitizeFilename(title)}.csv`, 'text/csv');
}

function downloadExcel(data, title) {
    const wb = XLSX.utils.book_new();
    
    // Create header sheet - adjust counts based on showMovedRows setting
    const displayUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + state.diffResult.moved.length;
    
    const headerData = [
        ['WELLBOUND DIFFERENCE REPORT'],
        [''],
        [`Report Type: ${title}`],
        [`Generated: ${new Date().toLocaleString()}`],
        [`Original File: ${state.diffResult.meta.file1Name}`],
        [`Updated File: ${state.diffResult.meta.file2Name}`],
        [`Primary Key(s): ${state.diffResult.meta.primaryKeys.join(', ')}`],
        [''],
        ['SUMMARY'],
        [`Unchanged Rows: ${displayUnchangedCount}`],
        [`Modified Rows: ${state.diffResult.modified.length}`],
        [`Added Rows: ${state.diffResult.added.length}`],
        [`Removed Rows: ${state.diffResult.removed.length}`]
    ];
    
    // Only include moved row count if the setting is enabled
    if (state.showMovedRows) {
        headerData.push([`Moved Rows: ${state.diffResult.moved.length}`]);
    }
    
    const wsHeader = XLSX.utils.aoa_to_sheet(headerData);
    XLSX.utils.book_append_sheet(wb, wsHeader, 'Summary');
    
    // Create data sheet
    const wsData = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, wsData, 'Data');
    
    XLSX.writeFile(wb, `${sanitizeFilename(title)}.xlsx`);
}

async function downloadPDF(data, title) {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('landscape');
    
    // Header
    doc.setFontSize(20);
    doc.setTextColor(37, 99, 235);
    doc.text('WELLBOUND', 14, 20);
    
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text('Certified Home Health Agency', 14, 26);
    
    doc.setFontSize(16);
    doc.setTextColor(0);
    doc.text(title, 14, 40);
    
    doc.setFontSize(10);
    doc.setTextColor(100);
    doc.text(`Generated: ${new Date().toLocaleString()}`, 14, 48);
    doc.text(`Original: ${state.diffResult.meta.file1Name} | Updated: ${state.diffResult.meta.file2Name}`, 14, 54);
    
    // Summary - adjust counts based on showMovedRows setting
    doc.setFontSize(12);
    doc.setTextColor(0);
    doc.text('Summary', 14, 66);
    
    const pdfUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + state.diffResult.moved.length;
    
    doc.setFontSize(10);
    const summaryText = state.showMovedRows 
        ? `Unchanged: ${state.diffResult.unchanged.length} | Modified: ${state.diffResult.modified.length} | Added: ${state.diffResult.added.length} | Removed: ${state.diffResult.removed.length} | Moved: ${state.diffResult.moved.length}`
        : `Unchanged: ${pdfUnchangedCount} | Modified: ${state.diffResult.modified.length} | Added: ${state.diffResult.added.length} | Removed: ${state.diffResult.removed.length}`;
    doc.text(summaryText, 14, 74);
    
    // Table - ALL rows, no limit
    const headers = Object.keys(data[0] || {});
    const tableData = data.map(row => headers.map(h => {
        const val = row[h];
        // Truncate very long cell values to prevent layout issues
        const strVal = String(val ?? '');
        return strVal.length > 50 ? strVal.substring(0, 47) + '...' : strVal;
    }));
    
    doc.autoTable({
        head: [headers],
        body: tableData,
        startY: 82,
        styles: { 
            fontSize: 7, 
            cellPadding: 2,
            overflow: 'linebreak',
            cellWidth: 'wrap'
        },
        headStyles: { 
            fillColor: [37, 99, 235],
            fontSize: 7,
            fontStyle: 'bold'
        },
        alternateRowStyles: { fillColor: [248, 250, 252] },
        // Add page numbers and header on each page
        didDrawPage: function(data) {
            // Footer with page number
            doc.setFontSize(8);
            doc.setTextColor(150);
            doc.text(
                `Page ${doc.internal.getNumberOfPages()}`,
                doc.internal.pageSize.width / 2,
                doc.internal.pageSize.height - 10,
                { align: 'center' }
            );
        },
        margin: { top: 20, bottom: 20 }
    });
    
    // Final page count
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
        doc.setPage(i);
        doc.setFontSize(8);
        doc.setTextColor(150);
        doc.text(
            `Page ${i} of ${totalPages}`,
            doc.internal.pageSize.width / 2,
            doc.internal.pageSize.height - 10,
            { align: 'center' }
        );
    }
    
    doc.save(`${sanitizeFilename(title)}.pdf`);
}

async function exportHighlightPDF() {
    if (!state.diffResult) {
        alert('No comparison data to export');
        return;
    }
    
    showLoading(true);
    
    try {
        const { jsPDF } = window.jspdf;
        // Landscape orientation for side-by-side view
        const doc = new jsPDF('landscape', 'mm', 'a4');
        
        const pageWidth = doc.internal.pageSize.getWidth();
        const pageHeight = doc.internal.pageSize.getHeight();
        const margin = 10;
        const halfWidth = (pageWidth - margin * 3) / 2;
        const rowHeight = 8;
        const headerHeight = 25;
        const footerHeight = 15;
        const contentHeight = pageHeight - margin - headerHeight - footerHeight;
        const rowsPerPage = Math.floor(contentHeight / rowHeight);
        
        // Get current colors from the UI
        const colors = {
            modified: document.getElementById('color-modified')?.value || '#f59e0b',
            added: document.getElementById('color-added')?.value || '#22c55e',
            removed: document.getElementById('color-removed')?.value || '#ef4444',
            unchanged: document.getElementById('color-unchanged')?.value || '#6b7280'
        };
        
        // Get current filter state
        const filters = {};
        document.querySelectorAll('#highlight-panel .legend-item .toggle-switch input').forEach(input => {
            if (input.dataset.filter) {
                filters[input.dataset.filter] = input.checked;
            }
        });
        
        // Get primary key column
        const keyCol = state.diffResult.meta?.primaryKeys?.[0] || state.headers[0];
        
        // Check column filter
        const hasColumnFilter = state.highlightColumnFilter.length > 0;
        
        // Build row data matching the display order
        const rows = [];
        
        // Unchanged rows
        if (filters.unchanged !== false) {
            state.diffResult.unchanged.forEach(item => {
                const keyVal = item.data[keyCol] || '';
                const preview = getRowPreviewForPDF(item.data);
                rows.push({
                    type: 'unchanged',
                    color: colors.unchanged,
                    original: { key: keyVal, data: preview },
                    updated: { key: keyVal, data: preview }
                });
            });
            
            // Moved rows (if hidden, treat as unchanged)
            if (!state.showMovedRows) {
                state.diffResult.moved?.forEach(item => {
                    const keyVal = item.data[keyCol] || '';
                    const preview = getRowPreviewForPDF(item.data);
                    rows.push({
                        type: 'unchanged',
                        color: colors.unchanged,
                        original: { key: keyVal, data: preview },
                        updated: { key: keyVal, data: preview }
                    });
                });
            }
        }
        
        // Modified rows excluded by column filter (shown as unchanged)
        if (hasColumnFilter && filters.unchanged !== false) {
            state.diffResult.modified
                .filter(item => !item.changes.some(c => state.highlightColumnFilter.includes(c.column)))
                .forEach(item => {
                    const keyVal = item.original[keyCol] || '';
                    const preview = getRowPreviewForPDF(item.updated);
                    rows.push({
                        type: 'unchanged',
                        color: colors.unchanged,
                        original: { key: keyVal, data: preview },
                        updated: { key: keyVal, data: preview }
                    });
                });
        }
        
        // Modified rows
        if (filters.modified !== false) {
            const filteredModified = hasColumnFilter 
                ? state.diffResult.modified.filter(item => 
                    item.changes.some(c => state.highlightColumnFilter.includes(c.column)))
                : state.diffResult.modified;
            
            filteredModified.forEach(item => {
                const keyVal = item.original[keyCol] || '';
                const originalPreview = getRowPreviewForPDF(item.original);
                const updatedPreview = getRowPreviewForPDF(item.updated);
                rows.push({
                    type: 'modified',
                    color: colors.modified,
                    badge: 'Modified',
                    original: { key: keyVal, data: originalPreview },
                    updated: { key: keyVal, data: updatedPreview }
                });
            });
        }
        
        // Removed rows
        if (filters.removed !== false) {
            state.diffResult.removed.forEach(item => {
                const keyVal = item.data[keyCol] || '';
                const preview = getRowPreviewForPDF(item.data);
                rows.push({
                    type: 'removed',
                    color: colors.removed,
                    badge: 'Removed',
                    original: { key: keyVal, data: preview },
                    updated: { key: '—', data: 'Record removed', italic: true }
                });
            });
        }
        
        // Added rows
        if (filters.added !== false) {
            state.diffResult.added.forEach(item => {
                const keyVal = item.data[keyCol] || '';
                const preview = getRowPreviewForPDF(item.data);
                rows.push({
                    type: 'added',
                    color: colors.added,
                    badge: 'Added',
                    original: { key: '—', data: 'New record', italic: true },
                    updated: { key: keyVal, data: preview }
                });
            });
        }
        
        if (rows.length === 0) {
            alert('No rows to export with current filters');
            showLoading(false);
            return;
        }
        
        // Calculate total pages
        const totalPages = Math.ceil(rows.length / rowsPerPage);
        
        // Helper to convert hex to RGB
        const hexToRGB = (hex) => {
            const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
            return result ? {
                r: parseInt(result[1], 16),
                g: parseInt(result[2], 16),
                b: parseInt(result[3], 16)
            } : { r: 0, g: 0, b: 0 };
        };
        
        // Draw each page
        for (let page = 0; page < totalPages; page++) {
            if (page > 0) doc.addPage();
            
            // Header
            doc.setFontSize(14);
            doc.setTextColor(37, 99, 235);
            doc.text('WELLBOUND - Side-by-Side Comparison', margin, margin + 5);
            
            doc.setFontSize(9);
            doc.setTextColor(100);
            const originalName = state.file1?.name || 'Original';
            const updatedName = state.file2?.name || 'Updated';
            doc.text(`${originalName} vs ${updatedName}`, margin, margin + 11);
            doc.text(`Page ${page + 1} of ${totalPages}`, pageWidth - margin - 30, margin + 5);
            
            // Column headers
            const headerY = margin + headerHeight - 5;
            
            // Original side header
            doc.setFillColor(248, 250, 252);
            doc.rect(margin, headerY - 5, halfWidth, 8, 'F');
            doc.setFontSize(10);
            doc.setTextColor(50);
            doc.setFont(undefined, 'bold');
            doc.text('Original', margin + 5, headerY);
            
            // Updated side header
            doc.rect(margin * 2 + halfWidth, headerY - 5, halfWidth, 8, 'F');
            doc.text('Updated', margin * 2 + halfWidth + 5, headerY);
            doc.setFont(undefined, 'normal');
            
            // Draw rows for this page
            const startRow = page * rowsPerPage;
            const endRow = Math.min(startRow + rowsPerPage, rows.length);
            
            for (let i = startRow; i < endRow; i++) {
                const row = rows[i];
                const rowIndex = i - startRow;
                const y = margin + headerHeight + (rowIndex * rowHeight);
                
                const rgb = hexToRGB(row.color);
                
                // Row background (light tint for changed rows)
                if (row.type !== 'unchanged') {
                    doc.setFillColor(rgb.r, rgb.g, rgb.b, 0.1);
                    // Left side
                    doc.rect(margin, y, halfWidth, rowHeight - 1, 'F');
                    // Right side
                    doc.rect(margin * 2 + halfWidth, y, halfWidth, rowHeight - 1, 'F');
                }
                
                // Left border indicator
                doc.setFillColor(rgb.r, rgb.g, rgb.b);
                doc.rect(margin, y, 2, rowHeight - 1, 'F');
                doc.rect(margin * 2 + halfWidth, y, 2, rowHeight - 1, 'F');
                
                // Row number
                doc.setFontSize(7);
                doc.setTextColor(150);
                const rowNum = String(i + 1);
                doc.text(rowNum, margin + 5, y + 5);
                
                // Original side content
                doc.setFontSize(8);
                if (row.original.italic) {
                    doc.setTextColor(150);
                    doc.setFont(undefined, 'italic');
                } else {
                    doc.setTextColor(50);
                    doc.setFont(undefined, 'normal');
                }
                
                const origKeyText = truncateText(String(row.original.key), 15);
                const origDataText = truncateText(row.original.data, 50);
                doc.text(origKeyText, margin + 15, y + 5);
                doc.text(origDataText, margin + 45, y + 5);
                
                // Badge on original side for removed
                if (row.type === 'removed') {
                    doc.setFillColor(rgb.r, rgb.g, rgb.b);
                    doc.roundedRect(margin + halfWidth - 22, y + 1, 18, 5, 1, 1, 'F');
                    doc.setFontSize(5);
                    doc.setTextColor(255);
                    doc.text('REMOVED', margin + halfWidth - 20, y + 4.5);
                }
                
                // Updated side content
                doc.setFontSize(8);
                if (row.updated.italic) {
                    doc.setTextColor(150);
                    doc.setFont(undefined, 'italic');
                } else {
                    doc.setTextColor(50);
                    doc.setFont(undefined, 'normal');
                }
                
                const updKeyText = truncateText(String(row.updated.key), 15);
                const updDataText = truncateText(row.updated.data, 50);
                doc.text(updKeyText, margin * 2 + halfWidth + 5, y + 5);
                doc.text(updDataText, margin * 2 + halfWidth + 35, y + 5);
                
                // Badge on updated side for modified/added
                if (row.type === 'modified' || row.type === 'added') {
                    doc.setFillColor(rgb.r, rgb.g, rgb.b);
                    const badgeText = row.type === 'modified' ? 'MODIFIED' : 'ADDED';
                    const badgeWidth = row.type === 'modified' ? 20 : 15;
                    doc.roundedRect(margin * 2 + halfWidth + halfWidth - badgeWidth - 4, y + 1, badgeWidth, 5, 1, 1, 'F');
                    doc.setFontSize(5);
                    doc.setTextColor(255);
                    doc.text(badgeText, margin * 2 + halfWidth + halfWidth - badgeWidth - 2, y + 4.5);
                }
                
                // Divider line
                doc.setDrawColor(230);
                doc.line(margin, y + rowHeight - 1, margin + halfWidth, y + rowHeight - 1);
                doc.line(margin * 2 + halfWidth, y + rowHeight - 1, margin * 2 + halfWidth * 2, y + rowHeight - 1);
            }
            
            // Footer with stats
            const footerY = pageHeight - footerHeight + 5;
            doc.setFontSize(7);
            doc.setTextColor(120);
            
            const statsText = `Unchanged: ${state.diffResult.unchanged.length} | Modified: ${state.diffResult.modified.length} | Added: ${state.diffResult.added.length} | Removed: ${state.diffResult.removed.length}`;
            doc.text(statsText, margin, footerY);
            doc.text(`Generated: ${new Date().toLocaleString()}`, pageWidth - margin - 50, footerY);
            
            // Center divider line
            doc.setDrawColor(200);
            doc.line(margin + halfWidth + margin/2, margin + headerHeight - 5, margin + halfWidth + margin/2, pageHeight - footerHeight);
        }
        
        // Save the PDF
        const timestamp = new Date().toISOString().slice(0, 10);
        doc.save(`highlight_comparison_${timestamp}.pdf`);
        
    } catch (error) {
        console.error('PDF export error:', error);
        alert(`PDF export error: ${error.message}`);
    }
    
    showLoading(false);
}

// Helper function to get row preview for PDF (simpler than HTML version)
function getRowPreviewForPDF(data) {
    if (!data) return '';
    const values = Object.values(data).slice(0, 5);
    return values.map(v => String(v ?? '')).join(' | ');
}

// Helper function to truncate text for PDF
function truncateText(text, maxLen) {
    if (!text) return '';
    const str = String(text);
    return str.length > maxLen ? str.substring(0, maxLen - 2) + '...' : str;
}

function exportDiffViewFile() {
    // Capture current viewer state for bundle
    const viewerState = {
        currentFilter: state.currentFilter,
        currentDensity: state.currentDensity,
        currentGranularity: state.currentGranularity,
        currentPage: state.currentPage,
        columnFilters: state.columnFilters,
        columnFiltersVisible: state.columnFiltersVisible,
        showMovedRows: state.showMovedRows,
        // Highlight view state
        highlightSyncScroll: state.highlightSyncScroll,
        highlightColumnFilter: state.highlightColumnFilter,
        highlightColors: state.highlightColors,
        highlightTypeFilters: state.highlightTypeFilters
    };
    
    const diffData = {
        version: '1.1',  // Bumped version for new state format
        type: 'wellbound-diff',
        ...state.diffResult,
        headers: state.headers,
        viewerState  // Include all viewer state in the bundle
    };
    
    const content = JSON.stringify(diffData, null, 2);
    const timestamp = new Date().toISOString().slice(0, 10);
    const filename = `comparison_bundle_${timestamp}.wbdiff`;
    
    downloadFile(content, filename, 'application/json');
}

function generateReportHeader(title) {
    const headerUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + state.diffResult.moved.length;
    
    let header = `# WELLBOUND DIFFERENCE REPORT
# ${title}
# Generated: ${new Date().toLocaleString()}
# Original File: ${state.diffResult.meta.file1Name}
# Updated File: ${state.diffResult.meta.file2Name}
# Primary Key(s): ${state.diffResult.meta.primaryKeys.join(', ')}
#
# Summary:
# - Unchanged Rows: ${headerUnchangedCount}
# - Modified Rows: ${state.diffResult.modified.length}
# - Added Rows: ${state.diffResult.added.length}
# - Removed Rows: ${state.diffResult.removed.length}`;
    
    if (state.showMovedRows) {
        header += `
# - Moved Rows: ${state.diffResult.moved.length}`;
    }
    
    return header;
}

function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function sanitizeFilename(name) {
    return name.replace(/[^a-z0-9]/gi, '_').substring(0, 50);
}

// ========================================
// Viewer Page
// ========================================

function initializeViewerControls() {
    // Import diff file
    document.getElementById('importDiffBtn')?.addEventListener('click', () => {
        document.getElementById('importDiffInput').click();
    });
    
    document.getElementById('importDiffInput')?.addEventListener('change', handleDiffImport);
    
    document.getElementById('importNewDiff')?.addEventListener('click', () => {
        document.getElementById('importDiffInput').click();
    });
    
    // Import bundle from main page
    document.getElementById('importBundleMain')?.addEventListener('click', () => {
        document.getElementById('importDiffInput').click();
    });
    
    // Open comparison view from bundle success state
    document.getElementById('openComparisonFromBundle')?.addEventListener('click', () => {
        switchPage('viewer');
        document.querySelectorAll('.nav-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.page === 'viewer');
        });
        updateViewerContent();
    });
    
    // Load different bundle (reset to initial state)
    document.getElementById('loadDifferentBundle')?.addEventListener('click', () => {
        // Hide success state, show empty state
        document.getElementById('bundle-success').hidden = true;
        document.getElementById('upload-empty').hidden = false;
        // Trigger import dialog
        document.getElementById('importDiffInput').click();
    });
    
    // Export bundle from viewer
    document.getElementById('exportBundleFromViewer')?.addEventListener('click', exportDiffViewFile);
    
    // Clear comparison button
    document.getElementById('clearComparison')?.addEventListener('click', clearComparison);
    
    // Granularity tabs
    document.querySelectorAll('.granularity-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('.granularity-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            state.currentGranularity = tab.dataset.level;
            updateViewerPanel();
        });
    });
    
    // Density buttons
    document.querySelectorAll('.density-btn').forEach(btn => {
        btn.addEventListener('click', () => {
            document.querySelectorAll('.density-btn').forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            state.currentDensity = btn.dataset.density;
            updateViewerPanel();
        });
    });
    
    // Filter chips
    document.querySelectorAll('.filter-chip').forEach(chip => {
        chip.addEventListener('click', () => {
            document.querySelectorAll('.filter-chip').forEach(c => c.classList.remove('active'));
            chip.classList.add('active');
            state.currentFilter = chip.dataset.filter;
            updateRowsPanel();
        });
    });
    
    // Pagination
    document.getElementById('prevPage')?.addEventListener('click', () => {
        if (state.currentPage > 1) {
            state.currentPage--;
            updateRowsPanel();
        }
    });
    
    document.getElementById('nextPage')?.addEventListener('click', () => {
        state.currentPage++;
        updateRowsPanel();
    });
    
    // Search
    document.getElementById('rowSearch')?.addEventListener('input', (e) => {
        state.currentPage = 1;
        updateRowsPanel();
    });
    
    // Column filters toggle
    document.getElementById('toggleColumnFilters')?.addEventListener('click', () => {
        state.columnFiltersVisible = !state.columnFiltersVisible;
        const btn = document.getElementById('toggleColumnFilters');
        const btnText = document.getElementById('filterBtnText');
        const clearBtn = document.getElementById('clearColumnFilters');
        
        if (state.columnFiltersVisible) {
            btn.classList.add('active');
            btnText.textContent = 'Hide Filters';
            clearBtn.hidden = false;
        } else {
            btn.classList.remove('active');
            btnText.textContent = 'Add Filters';
            clearBtn.hidden = true;
        }
        updateRowsPanel();
    });
    
    // Clear column filters
    document.getElementById('clearColumnFilters')?.addEventListener('click', () => {
        state.columnFilters = {};
        state.currentPage = 1;
        updateRowsPanel();
        updateFilterBadge();
    });
    
    // Export from viewer
    document.getElementById('exportFromViewer')?.addEventListener('click', () => {
        switchPage('generator');
        document.querySelectorAll('.nav-tab').forEach(t => {
            t.classList.toggle('active', t.dataset.page === 'generator');
        });
        document.getElementById('results-section').scrollIntoView({ behavior: 'smooth' });
    });
    
    // Highlight panel type filter controls (in legend items)
    document.querySelectorAll('#highlight-panel .legend-item .toggle-switch input').forEach(toggle => {
        toggle.addEventListener('change', () => {
            updateHighlightPanel();
        });
    });
    
    // Color pickers
    document.querySelectorAll('#highlight-panel input[type="color"]').forEach(picker => {
        picker.addEventListener('input', () => {
            updateHighlightPanel();
        });
    });
    
    // Reset colors button
    document.getElementById('resetColors')?.addEventListener('click', () => {
        document.getElementById('color-modified').value = '#f59e0b';
        document.getElementById('color-added').value = '#22c55e';
        document.getElementById('color-removed').value = '#ef4444';
        document.getElementById('color-unchanged').value = '#6b7280';
        state.highlightColors = {
            modified: '#f59e0b',
            added: '#22c55e',
            removed: '#ef4444',
            unchanged: '#6b7280'
        };
        updateHighlightPanel();
    });
    
    // Export highlight view as PDF
    document.getElementById('exportHighlightPDF')?.addEventListener('click', exportHighlightPDF);
    
    // Track color changes for state persistence
    document.querySelectorAll('#highlight-panel input[type="color"]').forEach(picker => {
        picker.addEventListener('change', () => {
            state.highlightColors = {
                modified: document.getElementById('color-modified')?.value || '#f59e0b',
                added: document.getElementById('color-added')?.value || '#22c55e',
                removed: document.getElementById('color-removed')?.value || '#ef4444',
                unchanged: document.getElementById('color-unchanged')?.value || '#6b7280'
            };
        });
    });
    
    // Track type filter changes for state persistence
    document.querySelectorAll('#highlight-panel .legend-item .toggle-switch input').forEach(toggle => {
        toggle.addEventListener('change', () => {
            const filterType = toggle.dataset.filter;
            if (filterType) {
                state.highlightTypeFilters[filterType] = toggle.checked;
            }
        });
    });
    
    // Sync scroll toggle
    document.getElementById('syncScrollToggle')?.addEventListener('change', (e) => {
        state.highlightSyncScroll = e.target.checked;
        updateSyncScrollState();
    });
    
    // Column filter dropdown
    initializeColumnFilterDropdown();
    
    // Cell level navigation
    document.getElementById('prevChange')?.addEventListener('click', navigateToPrevChange);
    document.getElementById('nextChange')?.addEventListener('click', navigateToNextChange);
}

// ========================================
// Sync Scroll Functionality
// ========================================

function updateSyncScrollState() {
    const container = document.querySelector('.highlight-container');
    const originalContent = document.getElementById('highlight-original');
    const updatedContent = document.getElementById('highlight-updated');
    
    if (!container || !originalContent || !updatedContent) return;
    
    if (state.highlightSyncScroll) {
        container.classList.add('sync-scroll');
        
        // Add scroll sync listeners
        originalContent.removeEventListener('scroll', syncScrollOriginal);
        updatedContent.removeEventListener('scroll', syncScrollUpdated);
        originalContent.addEventListener('scroll', syncScrollOriginal);
        updatedContent.addEventListener('scroll', syncScrollUpdated);
    } else {
        container.classList.remove('sync-scroll');
        
        // Remove scroll sync listeners
        originalContent.removeEventListener('scroll', syncScrollOriginal);
        updatedContent.removeEventListener('scroll', syncScrollUpdated);
    }
}

let isSyncingScroll = false;

function syncScrollOriginal() {
    if (isSyncingScroll) return;
    isSyncingScroll = true;
    
    const originalContent = document.getElementById('highlight-original');
    const updatedContent = document.getElementById('highlight-updated');
    
    if (originalContent && updatedContent) {
        updatedContent.scrollTop = originalContent.scrollTop;
    }
    
    requestAnimationFrame(() => {
        isSyncingScroll = false;
    });
}

function syncScrollUpdated() {
    if (isSyncingScroll) return;
    isSyncingScroll = true;
    
    const originalContent = document.getElementById('highlight-original');
    const updatedContent = document.getElementById('highlight-updated');
    
    if (originalContent && updatedContent) {
        originalContent.scrollTop = updatedContent.scrollTop;
    }
    
    requestAnimationFrame(() => {
        isSyncingScroll = false;
    });
}

// ========================================
// Column Filter Dropdown
// ========================================

function initializeColumnFilterDropdown() {
    const dropdownBtn = document.getElementById('columnFilterBtn');
    const dropdownMenu = document.getElementById('columnFilterMenu');
    const selectAllCheckbox = document.getElementById('selectAllColumns');
    const applyBtn = document.getElementById('applyColumnFilter');
    
    if (!dropdownBtn || !dropdownMenu) return;
    
    // Toggle dropdown
    dropdownBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const wasHidden = dropdownMenu.hidden;
        dropdownMenu.hidden = !wasHidden;
        dropdownBtn.classList.toggle('open', wasHidden);
        
        if (wasHidden) {
            // Populate columns when opening (was hidden, now showing)
            populateColumnFilterDropdown();
        }
    });
    
    // Close dropdown when clicking outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.column-filter-dropdown')) {
            dropdownMenu.hidden = true;
            dropdownBtn.classList.remove('open');
        }
    });
    
    // Select all checkbox
    selectAllCheckbox?.addEventListener('change', (e) => {
        const checkboxes = document.querySelectorAll('#columnFilterItems input[type="checkbox"]');
        checkboxes.forEach(cb => {
            cb.checked = e.target.checked;
        });
    });
    
    // Apply button
    applyBtn?.addEventListener('click', () => {
        applyColumnFilter();
        dropdownMenu.hidden = true;
        dropdownBtn.classList.remove('open');
    });
}

function populateColumnFilterDropdown() {
    const container = document.getElementById('columnFilterItems');
    if (!container || !state.headers) return;
    
    // Get all columns that have changes
    const columnsWithChanges = new Set();
    
    state.diffResult?.modified?.forEach(item => {
        item.changes?.forEach(change => {
            columnsWithChanges.add(change.column);
        });
    });
    
    // Build checkbox list
    container.innerHTML = state.headers.map(header => {
        const hasChanges = columnsWithChanges.has(header);
        const isSelected = state.highlightColumnFilter.length === 0 || 
                          state.highlightColumnFilter.includes(header);
        
        return `
            <label class="checkbox-label">
                <input type="checkbox" value="${escapeHtml(header)}" ${isSelected ? 'checked' : ''}>
                <span>${escapeHtml(header)}</span>
                ${hasChanges ? '<span class="column-change-badge">has changes</span>' : ''}
            </label>
        `;
    }).join('');
    
    // Update select all checkbox state
    updateSelectAllState();
    
    // Add change listeners to update select all
    container.querySelectorAll('input[type="checkbox"]').forEach(cb => {
        cb.addEventListener('change', updateSelectAllState);
    });
}

function updateSelectAllState() {
    const selectAllCheckbox = document.getElementById('selectAllColumns');
    const checkboxes = document.querySelectorAll('#columnFilterItems input[type="checkbox"]');
    
    if (!selectAllCheckbox || checkboxes.length === 0) return;
    
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    const someChecked = Array.from(checkboxes).some(cb => cb.checked);
    
    selectAllCheckbox.checked = allChecked;
    selectAllCheckbox.indeterminate = someChecked && !allChecked;
}

function applyColumnFilter() {
    const checkboxes = document.querySelectorAll('#columnFilterItems input[type="checkbox"]:checked');
    const selectedColumns = Array.from(checkboxes).map(cb => cb.value);
    
    // If all columns selected, store empty array (means "all")
    if (selectedColumns.length === state.headers.length) {
        state.highlightColumnFilter = [];
    } else {
        state.highlightColumnFilter = selectedColumns;
    }
    
    // Update label
    updateColumnFilterLabel();
    
    // Refresh highlight panel
    updateHighlightPanel();
}

function updateColumnFilterLabel() {
    const label = document.getElementById('columnFilterLabel');
    if (!label) return;
    
    if (state.highlightColumnFilter.length === 0) {
        label.textContent = 'All Columns';
    } else if (state.highlightColumnFilter.length === 1) {
        label.textContent = state.highlightColumnFilter[0];
    } else {
        label.textContent = `${state.highlightColumnFilter.length} Columns`;
    }
}

function handleDiffImport(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    // Check if we're on the generator page (main page)
    const isOnMainPage = document.getElementById('generator-page')?.classList.contains('active');
    
    const reader = new FileReader();
    reader.onload = (event) => {
        try {
            const data = JSON.parse(event.target.result);
            
            if (data.type !== 'wellbound-diff') {
                throw new Error('Invalid diff file format');
            }
            
            state.diffResult = data;
            state.headers = data.headers || [];
            
            // Restore viewer state from bundle (support both old and new format)
            if (data.viewerState) {
                // New format with full viewer state
                state.currentFilter = data.viewerState.currentFilter || 'all';
                state.currentDensity = data.viewerState.currentDensity || 'standard';
                state.currentGranularity = data.viewerState.currentGranularity || 'summary';
                state.currentPage = data.viewerState.currentPage || 1;
                state.columnFilters = data.viewerState.columnFilters || {};
                state.columnFiltersVisible = data.viewerState.columnFiltersVisible || false;
                state.showMovedRows = data.viewerState.showMovedRows ?? false;
                state.highlightSyncScroll = data.viewerState.highlightSyncScroll ?? false;
                state.highlightColumnFilter = data.viewerState.highlightColumnFilter || [];
                state.highlightColors = data.viewerState.highlightColors || {
                    modified: '#f59e0b',
                    added: '#22c55e',
                    removed: '#ef4444',
                    unchanged: '#6b7280'
                };
                state.highlightTypeFilters = data.viewerState.highlightTypeFilters || {
                    modified: true,
                    added: true,
                    removed: true,
                    unchanged: true
                };
            } else {
                // Old format - just restore showMovedRows
                state.showMovedRows = data.showMovedRows ?? false;
            }
            
            // Mock file objects for display
            state.file1 = { name: data.meta?.file1Name || 'Original' };
            state.file2 = { name: data.meta?.file2Name || 'Updated' };
            
            // Apply restored state to UI
            applyRestoredViewerState();
            
            // If on main page, show success state with animation
            if (isOnMainPage) {
                showBundleSuccessState(data);
            } else {
                // Already on viewer page, just update content
                updateViewerContent();
            }
        } catch (error) {
            alert(`Error importing diff file: ${error.message}`);
        }
    };
    reader.readAsText(file);
    
    // Reset the input so the same file can be selected again
    e.target.value = '';
}

function showBundleSuccessState(data) {
    // Hide other states
    document.getElementById('upload-empty').hidden = true;
    document.getElementById('upload-files').hidden = true;
    document.getElementById('config-section').hidden = true;
    document.getElementById('results-section').hidden = true;
    
    // Show bundle success state
    const bundleSuccess = document.getElementById('bundle-success');
    bundleSuccess.hidden = false;
    
    // Update the info
    const file1Name = data.meta?.file1Name || 'Original';
    const file2Name = data.meta?.file2Name || 'Updated';
    document.getElementById('bundle-file-info').textContent = 
        `${file1Name} vs ${file2Name}`;
    
    // Update stats
    document.getElementById('bundle-modified').textContent = data.modified?.length || 0;
    document.getElementById('bundle-added').textContent = data.added?.length || 0;
    document.getElementById('bundle-removed').textContent = data.removed?.length || 0;
    
    // Trigger re-animation by removing and re-adding the element
    bundleSuccess.style.animation = 'none';
    bundleSuccess.offsetHeight; // Trigger reflow
    bundleSuccess.style.animation = '';
}

function applyRestoredViewerState() {
    // Apply granularity tab
    document.querySelectorAll('.granularity-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.level === state.currentGranularity);
    });
    
    // Apply density button
    document.querySelectorAll('.density-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.density === state.currentDensity);
    });
    
    // Apply filter chips
    document.querySelectorAll('.filter-chip').forEach(chip => {
        chip.classList.toggle('active', chip.dataset.filter === state.currentFilter);
    });
    
    // Apply column filters visible state for rows panel
    const toggleFiltersBtn = document.getElementById('toggleColumnFilters');
    const filterBtnText = document.getElementById('filterBtnText');
    const clearFiltersBtn = document.getElementById('clearColumnFilters');
    
    if (toggleFiltersBtn && filterBtnText && clearFiltersBtn) {
        if (state.columnFiltersVisible) {
            toggleFiltersBtn.classList.add('active');
            filterBtnText.textContent = 'Hide Filters';
            clearFiltersBtn.hidden = false;
        } else {
            toggleFiltersBtn.classList.remove('active');
            filterBtnText.textContent = 'Add Filters';
            clearFiltersBtn.hidden = true;
        }
    }
    
    // Apply highlight colors
    if (state.highlightColors) {
        const colorModified = document.getElementById('color-modified');
        const colorAdded = document.getElementById('color-added');
        const colorRemoved = document.getElementById('color-removed');
        const colorUnchanged = document.getElementById('color-unchanged');
        
        if (colorModified) colorModified.value = state.highlightColors.modified;
        if (colorAdded) colorAdded.value = state.highlightColors.added;
        if (colorRemoved) colorRemoved.value = state.highlightColors.removed;
        if (colorUnchanged) colorUnchanged.value = state.highlightColors.unchanged;
    }
    
    // Apply highlight type filters
    if (state.highlightTypeFilters) {
        document.querySelectorAll('#highlight-panel .legend-item .toggle-switch input[data-filter]').forEach(input => {
            const filterType = input.dataset.filter;
            if (state.highlightTypeFilters.hasOwnProperty(filterType)) {
                input.checked = state.highlightTypeFilters[filterType];
            }
        });
    }
    
    // Apply sync scroll toggle
    const syncScrollToggle = document.getElementById('syncScrollToggle');
    if (syncScrollToggle) {
        syncScrollToggle.checked = state.highlightSyncScroll;
    }
    
    // Update column filter label (will be applied when highlight panel is shown)
    updateColumnFilterLabel();
}

function clearComparison() {
    // Reset comparison-related state
    state.diffResult = null;
    state.headers = [];
    state.currentPage = 1;
    state.currentFilter = 'all';
    state.columnFilters = {};
    state.columnFiltersVisible = false;
    state.highlightColumnFilter = [];
    state.highlightSyncScroll = false;
    
    // Reset highlight type filters to defaults
    state.highlightTypeFilters = {
        modified: true,
        added: true,
        removed: true,
        unchanged: true
    };
    
    // Reset highlight colors to defaults
    state.highlightColors = {
        modified: '#f59e0b',
        added: '#22c55e',
        removed: '#ef4444',
        unchanged: '#6b7280'
    };
    
    // Show empty state, hide viewer content
    document.getElementById('viewer-empty').hidden = false;
    document.getElementById('viewer-content').hidden = true;
    
    // Also reset the main page bundle success state if visible
    const bundleSuccess = document.getElementById('bundle-success');
    if (bundleSuccess && !bundleSuccess.hidden) {
        bundleSuccess.hidden = true;
        document.getElementById('upload-empty').hidden = false;
    }
    
    // Reset UI controls to defaults
    document.querySelectorAll('.filter-chip').forEach(chip => {
        chip.classList.toggle('active', chip.dataset.filter === 'all');
    });
    
    document.querySelectorAll('.granularity-tab').forEach(tab => {
        tab.classList.toggle('active', tab.dataset.level === 'summary');
    });
    state.currentGranularity = 'summary';
    
    // Reset sync scroll toggle
    const syncScrollToggle = document.getElementById('syncScrollToggle');
    if (syncScrollToggle) syncScrollToggle.checked = false;
    
    // Reset column filter label
    updateColumnFilterLabel();
}

function updateViewerContent() {
    if (!state.diffResult) {
        document.getElementById('viewer-empty').hidden = false;
        document.getElementById('viewer-content').hidden = true;
        return;
    }
    
    document.getElementById('viewer-empty').hidden = true;
    document.getElementById('viewer-content').hidden = false;
    
    // Update subtitle
    document.getElementById('viewer-subtitle').textContent = 
        `Comparing: ${state.file1?.name || 'File 1'} vs ${state.file2?.name || 'File 2'}`;
    
    // Update filenames in cell view
    document.getElementById('original-filename').textContent = state.file1?.name || 'Original';
    document.getElementById('updated-filename').textContent = state.file2?.name || 'Updated';
    
    // Show/hide moved filter chip based on setting
    const movedFilterChip = document.querySelector('.filter-chip[data-filter="moved"]');
    if (movedFilterChip) {
        movedFilterChip.style.display = state.showMovedRows ? '' : 'none';
        // If currently filtering by moved but moved is hidden, reset to all
        if (!state.showMovedRows && state.currentFilter === 'moved') {
            state.currentFilter = 'all';
            document.querySelectorAll('.filter-chip').forEach(chip => {
                chip.classList.toggle('active', chip.dataset.filter === 'all');
            });
        }
    }
    
    updateViewerPanel();
}

function updateViewerPanel() {
    const panels = ['summary', 'rows', 'highlight', 'cells'];
    panels.forEach(p => {
        const panel = document.getElementById(`${p}-panel`);
        if (panel) panel.hidden = p !== state.currentGranularity;
    });
    
    switch (state.currentGranularity) {
        case 'summary':
            updateSummaryPanel();
            break;
        case 'rows':
            updateRowsPanel();
            break;
        case 'highlight':
            updateHighlightPanel();
            break;
        case 'cells':
            updateCellsPanel();
            break;
    }
}

function updateSummaryPanel() {
    // Update chart
    updateDiffChart();
    
    // Update column changes
    updateColumnChanges();
    
    // Update stats
    document.getElementById('stat-original').textContent = state.diffResult.meta?.originalRowCount || state.diffResult.unchanged.length + state.diffResult.modified.length + state.diffResult.removed.length;
    document.getElementById('stat-updated').textContent = state.diffResult.meta?.updatedRowCount || state.diffResult.unchanged.length + state.diffResult.modified.length + state.diffResult.added.length;
    
    const totalChanges = state.diffResult.modified.length + state.diffResult.added.length + state.diffResult.removed.length;
    const totalRows = Math.max(state.diffResult.meta?.originalRowCount || 1, state.diffResult.meta?.updatedRowCount || 1);
    const changeRate = ((totalChanges / totalRows) * 100).toFixed(1);
    document.getElementById('stat-rate').textContent = `${changeRate}%`;
    
    // Find most changed column
    const columnChanges = state.diffResult.columnChanges || {};
    const sortedColumns = Object.entries(columnChanges).sort((a, b) => b[1] - a[1]);
    document.getElementById('stat-column').textContent = sortedColumns[0]?.[0] || '-';
    
    // Initialize expanded analytics if it's visible
    if (!document.getElementById('expandedAnalytics').hidden) {
        initializeExpandedAnalytics();
    }
}

// ========================================
// Expanded Analytics Dashboard
// ========================================

// Store chart instances for cleanup
const analyticsCharts = {};

function initializeAnalyticsControls() {
    // Expand/collapse button
    document.getElementById('expandAnalyticsBtn')?.addEventListener('click', toggleExpandedAnalytics);
    
    // Chart style selector
    document.getElementById('chartStyleSelect')?.addEventListener('change', () => {
        updateAllAnalyticsCharts();
    });
    
    // Labels toggle
    document.getElementById('showLabelsToggle')?.addEventListener('change', () => {
        updateAllAnalyticsCharts();
    });
    
    // Animation toggle
    document.getElementById('animationToggle')?.addEventListener('change', () => {
        updateAllAnalyticsCharts();
    });
    
    // Refresh button
    document.getElementById('refreshAnalytics')?.addEventListener('click', () => {
        initializeExpandedAnalytics();
    });
    
    // Value distribution controls
    document.getElementById('distColumnSelect')?.addEventListener('change', updateValueDistributionChart);
    document.getElementById('distChartType')?.addEventListener('change', updateValueDistributionChart);
    
    // Change breakdown chart type
    document.getElementById('breakdownChartType')?.addEventListener('change', updateChangeBreakdownChart);
    
    // Top columns count
    document.getElementById('topColumnsCount')?.addEventListener('change', updateTopColumnsChart);
    
    // Before/after comparison
    document.getElementById('compareColumnSelect')?.addEventListener('change', updateBeforeAfterChart);
    document.getElementById('swapCompareView')?.addEventListener('click', () => {
        state.beforeAfterSwapped = !state.beforeAfterSwapped;
        updateBeforeAfterChart();
    });
    
    // Status distribution
    document.getElementById('statusColumnSelect')?.addEventListener('change', updateStatusDistChart);
    
    // Numeric analysis
    document.getElementById('numericColumnSelect')?.addEventListener('change', updateNumericStats);
    
    // Changed records table
    document.getElementById('changedRecordsFilter')?.addEventListener('change', updateChangedRecordsTable);
    document.getElementById('changedRecordsCount')?.addEventListener('change', updateChangedRecordsTable);
    
    // Heatmap controls
    document.getElementById('heatmapShowValues')?.addEventListener('change', updateColumnHeatmap);
    document.getElementById('heatmapColorScale')?.addEventListener('change', updateColumnHeatmap);
}

function toggleExpandedAnalytics() {
    const panel = document.getElementById('expandedAnalytics');
    const btn = document.getElementById('expandAnalyticsBtn');
    const icon = btn.querySelector('.expand-icon');
    
    if (panel.hidden) {
        panel.hidden = false;
        btn.classList.add('expanded');
        icon.classList.remove('fa-chevron-down');
        icon.classList.add('fa-chevron-up');
        btn.querySelector('span').textContent = 'Collapse Analytics Dashboard';
        initializeExpandedAnalytics();
    } else {
        panel.hidden = true;
        btn.classList.remove('expanded');
        icon.classList.remove('fa-chevron-up');
        icon.classList.add('fa-chevron-down');
        btn.querySelector('span').textContent = 'Expand Analytics Dashboard';
    }
}

function initializeExpandedAnalytics() {
    if (!state.diffResult) return;
    
    // Populate column selectors
    populateColumnSelectors();
    
    // Update all analytics components
    updateMetricsCards();
    updateValueDistributionChart();
    updateChangeBreakdownChart();
    updateTopColumnsChart();
    updateBeforeAfterChart();
    updateCompletenessChart();
    updateStatusDistChart();
    updateNumericStats();
    updateChangedRecordsTable();
    updateColumnHeatmap();
    createSparklines();
}

function getChartColors() {
    const style = document.getElementById('chartStyleSelect')?.value || 'default';
    
    const colorSchemes = {
        default: ['#3b82f6', '#22c55e', '#f59e0b', '#ef4444', '#8b5cf6', '#06b6d4', '#ec4899', '#84cc16'],
        vibrant: ['#ff6384', '#36a2eb', '#ffce56', '#4bc0c0', '#9966ff', '#ff9f40', '#ff6384', '#c9cbcf'],
        pastel: ['#aec6cf', '#ffb3ba', '#baffc9', '#bae1ff', '#ffffba', '#ffdfba', '#e0bbff', '#c9c9ff'],
        monochrome: ['#2563eb', '#3b82f6', '#60a5fa', '#93c5fd', '#bfdbfe', '#dbeafe', '#1d4ed8', '#1e40af']
    };
    
    return colorSchemes[style] || colorSchemes.default;
}

function getChartOptions(type = 'bar') {
    const showLabels = document.getElementById('showLabelsToggle')?.checked ?? true;
    const animate = document.getElementById('animationToggle')?.checked ?? true;
    
    const baseOptions = {
        responsive: true,
        maintainAspectRatio: false,
        animation: animate ? { duration: 750 } : false,
        plugins: {
            legend: {
                display: showLabels,
                position: 'bottom',
                labels: {
                    usePointStyle: true,
                    padding: 15,
                    font: { size: 11 }
                }
            }
        }
    };
    
    if (type === 'bar' || type === 'horizontalBar') {
        baseOptions.scales = {
            y: { beginAtZero: true, grid: { color: 'rgba(0,0,0,0.05)' } },
            x: { grid: { display: false } }
        };
        if (type === 'horizontalBar') {
            baseOptions.indexAxis = 'y';
        }
    }
    
    return baseOptions;
}

function populateColumnSelectors() {
    const headers = state.headers || [];
    
    // Find columns with changes
    const changedColumns = Object.entries(state.diffResult.columnChanges || {})
        .filter(([, count]) => count > 0)
        .map(([col]) => col);
    
    // Populate distribution column selector
    const distSelect = document.getElementById('distColumnSelect');
    if (distSelect) {
        distSelect.innerHTML = headers.map(h => 
            `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`
        ).join('');
    }
    
    // Populate compare column selector (prioritize changed columns)
    const compareSelect = document.getElementById('compareColumnSelect');
    if (compareSelect) {
        const sortedHeaders = [...changedColumns, ...headers.filter(h => !changedColumns.includes(h))];
        compareSelect.innerHTML = sortedHeaders.map(h => 
            `<option value="${escapeHtml(h)}">${escapeHtml(h)}${changedColumns.includes(h) ? ' *' : ''}</option>`
        ).join('');
    }
    
    // Populate status column selector (look for status-like columns)
    const statusSelect = document.getElementById('statusColumnSelect');
    if (statusSelect) {
        const statusColumns = headers.filter(h => 
            h.toLowerCase().includes('status') || 
            h.toLowerCase().includes('state') || 
            h.toLowerCase().includes('type') ||
            h.toLowerCase().includes('category')
        );
        const options = statusColumns.length > 0 ? statusColumns : headers.slice(0, 5);
        statusSelect.innerHTML = options.map(h => 
            `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`
        ).join('');
    }
    
    // Populate numeric column selector
    const numericSelect = document.getElementById('numericColumnSelect');
    if (numericSelect) {
        // Detect numeric columns from data
        const numericColumns = headers.filter(h => {
            const sampleData = getAllData().slice(0, 100);
            const numericCount = sampleData.filter(row => {
                const val = row[h];
                return val !== null && val !== '' && !isNaN(parseFloat(val));
            }).length;
            return numericCount > sampleData.length * 0.5;
        });
        
        const options = numericColumns.length > 0 ? numericColumns : headers.slice(0, 5);
        numericSelect.innerHTML = options.map(h => 
            `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`
        ).join('');
    }
}

function getAllData() {
    if (!state.diffResult) return [];
    
    const allData = [];
    state.diffResult.unchanged?.forEach(item => allData.push(item.data));
    state.diffResult.modified?.forEach(item => allData.push(item.updated));
    state.diffResult.added?.forEach(item => allData.push(item.data));
    // Don't include removed since they're not in the updated data
    return allData;
}

function getOriginalData() {
    if (!state.diffResult) return [];
    
    const data = [];
    state.diffResult.unchanged?.forEach(item => data.push(item.data));
    state.diffResult.modified?.forEach(item => data.push(item.original));
    state.diffResult.removed?.forEach(item => data.push(item.data));
    return data;
}

function updateMetricsCards() {
    if (!state.diffResult) return;
    
    const found = state.diffResult.foundStatus?.found || 0;
    const notFound = state.diffResult.foundStatus?.notFound || 0;
    const totalFound = found + notFound;
    const foundRate = totalFound > 0 ? ((found / totalFound) * 100).toFixed(1) : 0;
    
    const matched = state.diffResult.statusMatch?.matched || 0;
    const changed = state.diffResult.statusMatch?.changed || 0;
    const totalMatch = matched + changed;
    const matchRate = totalMatch > 0 ? ((matched / totalMatch) * 100).toFixed(1) : 0;
    
    const modified = state.diffResult.modified?.length || 0;
    const totalRecords = state.diffResult.meta?.originalRowCount || 
        (state.diffResult.unchanged?.length || 0) + modified + (state.diffResult.removed?.length || 0);
    const modifiedRate = totalRecords > 0 ? ((modified / totalRecords) * 100).toFixed(1) : 0;
    
    // Update metrics
    document.getElementById('analytics-found-rate').textContent = `${foundRate}%`;
    document.getElementById('analytics-match-rate').textContent = `${matchRate}%`;
    document.getElementById('analytics-modified-rate').textContent = `${modifiedRate}%`;
    document.getElementById('analytics-total-records').textContent = totalRecords.toLocaleString();
    
    // Update trends
    const foundTrend = document.getElementById('analytics-found-trend');
    if (foundTrend) {
        foundTrend.querySelector('span').textContent = found.toLocaleString();
        foundTrend.className = `metric-trend ${foundRate >= 80 ? 'positive' : foundRate >= 50 ? 'neutral' : 'negative'}`;
    }
    
    const matchTrend = document.getElementById('analytics-match-trend');
    if (matchTrend) {
        matchTrend.querySelector('span').textContent = matched.toLocaleString();
        matchTrend.className = `metric-trend ${matchRate >= 80 ? 'positive' : matchRate >= 50 ? 'neutral' : 'warning'}`;
    }
    
    const modifiedTrend = document.getElementById('analytics-modified-trend');
    if (modifiedTrend) {
        modifiedTrend.querySelector('span').textContent = modified.toLocaleString();
        modifiedTrend.className = `metric-trend ${modifiedRate <= 20 ? 'positive' : modifiedRate <= 50 ? 'warning' : 'negative'}`;
    }
    
    const recordsTrend = document.getElementById('analytics-records-trend');
    if (recordsTrend) {
        recordsTrend.querySelector('span').textContent = (state.headers?.length || 0).toLocaleString();
    }
}

function createSparklines() {
    // Simple sparkline for found rate
    createMiniChart('foundSparkline', [30, 45, 60, 75, 85, 90, 95], '#22c55e');
    createMiniChart('matchSparkline', [40, 55, 65, 70, 80, 85, 88], '#3b82f6');
    createMiniChart('modifiedSparkline', [25, 20, 18, 15, 12, 10, 8], '#f59e0b');
    createMiniChart('recordsSparkline', [80, 85, 90, 95, 100, 105, 110], '#8b5cf6');
}

function createMiniChart(canvasId, data, color) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    
    const ctx = canvas.getContext('2d');
    
    // Destroy existing chart
    if (analyticsCharts[canvasId]) {
        analyticsCharts[canvasId].destroy();
    }
    
    analyticsCharts[canvasId] = new Chart(ctx, {
        type: 'line',
        data: {
            labels: data.map((_, i) => i),
            datasets: [{
                data: data,
                borderColor: color,
                backgroundColor: color + '20',
                fill: true,
                tension: 0.4,
                pointRadius: 0,
                borderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: {
                x: { display: false },
                y: { display: false }
            },
            animation: false
        }
    });
}

function updateValueDistributionChart() {
    const column = document.getElementById('distColumnSelect')?.value;
    const chartType = document.getElementById('distChartType')?.value || 'bar';
    
    if (!column || !state.diffResult) return;
    
    const data = getAllData();
    const valueCounts = {};
    
    data.forEach(row => {
        const val = String(row[column] ?? '(empty)').trim() || '(empty)';
        valueCounts[val] = (valueCounts[val] || 0) + 1;
    });
    
    // Sort by count and take top 15
    const sorted = Object.entries(valueCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    const labels = sorted.map(([label]) => label.length > 20 ? label.substring(0, 17) + '...' : label);
    const values = sorted.map(([, count]) => count);
    const colors = getChartColors();
    
    const canvas = document.getElementById('valueDistChart');
    if (!canvas) return;
    
    if (analyticsCharts.valueDistChart) {
        analyticsCharts.valueDistChart.destroy();
    }
    
    const actualChartType = chartType === 'horizontalBar' ? 'bar' : chartType;
    const options = getChartOptions(chartType);
    
    analyticsCharts.valueDistChart = new Chart(canvas.getContext('2d'), {
        type: actualChartType,
        data: {
            labels,
            datasets: [{
                label: `Values in "${column}"`,
                data: values,
                backgroundColor: chartType === 'bar' || chartType === 'horizontalBar' 
                    ? colors[0] + 'cc'
                    : colors.slice(0, values.length),
                borderColor: chartType === 'bar' || chartType === 'horizontalBar'
                    ? colors[0]
                    : colors.slice(0, values.length),
                borderWidth: 1
            }]
        },
        options
    });
}

function updateChangeBreakdownChart() {
    const chartType = document.getElementById('breakdownChartType')?.value || 'doughnut';
    
    if (!state.diffResult) return;
    
    // When showMovedRows is off, add moved count to unchanged
    const displayUnchangedCount = state.showMovedRows 
        ? (state.diffResult.unchanged?.length || 0)
        : (state.diffResult.unchanged?.length || 0) + (state.diffResult.moved?.length || 0);
    
    // Build data and labels based on showMovedRows setting
    const labels = state.showMovedRows 
        ? ['Unchanged', 'Modified', 'Added', 'Removed', 'Moved']
        : ['Unchanged', 'Modified', 'Added', 'Removed'];
    
    const data = state.showMovedRows 
        ? [
            state.diffResult.unchanged?.length || 0,
            state.diffResult.modified?.length || 0,
            state.diffResult.added?.length || 0,
            state.diffResult.removed?.length || 0,
            state.diffResult.moved?.length || 0
        ]
        : [
            displayUnchangedCount,
            state.diffResult.modified?.length || 0,
            state.diffResult.added?.length || 0,
            state.diffResult.removed?.length || 0
        ];
    
    const colors = state.showMovedRows 
        ? ['#6b7280', '#f59e0b', '#22c55e', '#ef4444', '#8b5cf6']
        : ['#6b7280', '#f59e0b', '#22c55e', '#ef4444'];
    
    const canvas = document.getElementById('changeBreakdownChart');
    if (!canvas) return;
    
    if (analyticsCharts.changeBreakdownChart) {
        analyticsCharts.changeBreakdownChart.destroy();
    }
    
    analyticsCharts.changeBreakdownChart = new Chart(canvas.getContext('2d'), {
        type: chartType,
        data: {
            labels,
            datasets: [{
                data,
                backgroundColor: colors,
                borderWidth: 0
            }]
        },
        options: getChartOptions(chartType)
    });
}

function updateTopColumnsChart() {
    const countSetting = document.getElementById('topColumnsCount')?.value || '10';
    
    if (!state.diffResult) return;
    
    const columnChanges = state.diffResult.columnChanges || {};
    let sorted = Object.entries(columnChanges)
        .filter(([, count]) => count > 0)
        .sort((a, b) => b[1] - a[1]);
    
    if (countSetting !== 'all') {
        sorted = sorted.slice(0, parseInt(countSetting));
    }
    
    const labels = sorted.map(([col]) => col.length > 15 ? col.substring(0, 12) + '...' : col);
    const values = sorted.map(([, count]) => count);
    const colors = getChartColors();
    
    const canvas = document.getElementById('topColumnsChart');
    if (!canvas) return;
    
    if (analyticsCharts.topColumnsChart) {
        analyticsCharts.topColumnsChart.destroy();
    }
    
    analyticsCharts.topColumnsChart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels,
            datasets: [{
                label: 'Number of Changes',
                data: values,
                backgroundColor: colors.map(c => c + 'cc'),
                borderColor: colors,
                borderWidth: 1
            }]
        },
        options: {
            ...getChartOptions('bar'),
            indexAxis: 'y'
        }
    });
}

function updateBeforeAfterChart() {
    const column = document.getElementById('compareColumnSelect')?.value;
    
    if (!column || !state.diffResult) return;
    
    // Get before and after values for modified rows
    const modifiedData = state.diffResult.modified || [];
    const relevantChanges = modifiedData.filter(item => 
        item.changes.some(c => c.column === column)
    );
    
    // Aggregate before/after values
    const beforeCounts = {};
    const afterCounts = {};
    
    relevantChanges.forEach(item => {
        const change = item.changes.find(c => c.column === column);
        if (change) {
            const oldVal = String(change.oldValue ?? '(empty)').trim() || '(empty)';
            const newVal = String(change.newValue ?? '(empty)').trim() || '(empty)';
            beforeCounts[oldVal] = (beforeCounts[oldVal] || 0) + 1;
            afterCounts[newVal] = (afterCounts[newVal] || 0) + 1;
        }
    });
    
    // Get all unique values
    const allValues = [...new Set([...Object.keys(beforeCounts), ...Object.keys(afterCounts)])];
    const topValues = allValues
        .sort((a, b) => (beforeCounts[b] || 0) + (afterCounts[b] || 0) - (beforeCounts[a] || 0) - (afterCounts[a] || 0))
        .slice(0, 10);
    
    const labels = topValues.map(v => v.length > 15 ? v.substring(0, 12) + '...' : v);
    const beforeData = topValues.map(v => beforeCounts[v] || 0);
    const afterData = topValues.map(v => afterCounts[v] || 0);
    
    const canvas = document.getElementById('beforeAfterChart');
    if (!canvas) return;
    
    if (analyticsCharts.beforeAfterChart) {
        analyticsCharts.beforeAfterChart.destroy();
    }
    
    const swapped = state.beforeAfterSwapped;
    
    analyticsCharts.beforeAfterChart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels,
            datasets: [
                {
                    label: swapped ? 'After' : 'Before',
                    data: swapped ? afterData : beforeData,
                    backgroundColor: swapped ? '#22c55e99' : '#ef444499',
                    borderColor: swapped ? '#22c55e' : '#ef4444',
                    borderWidth: 1
                },
                {
                    label: swapped ? 'Before' : 'After',
                    data: swapped ? beforeData : afterData,
                    backgroundColor: swapped ? '#ef444499' : '#22c55e99',
                    borderColor: swapped ? '#ef4444' : '#22c55e',
                    borderWidth: 1
                }
            ]
        },
        options: getChartOptions('bar')
    });
}

function updateCompletenessChart() {
    if (!state.diffResult) return;
    
    const data = getAllData();
    const headers = state.headers || [];
    
    // Calculate completeness per column
    const completeness = headers.map(col => {
        const filled = data.filter(row => {
            const val = row[col];
            return val !== null && val !== undefined && String(val).trim() !== '';
        }).length;
        return {
            column: col,
            rate: data.length > 0 ? (filled / data.length) * 100 : 0
        };
    });
    
    // Sort by completeness rate
    completeness.sort((a, b) => b.rate - a.rate);
    const top10 = completeness.slice(0, 10);
    
    const canvas = document.getElementById('completenessChart');
    if (!canvas) return;
    
    if (analyticsCharts.completenessChart) {
        analyticsCharts.completenessChart.destroy();
    }
    
    analyticsCharts.completenessChart = new Chart(canvas.getContext('2d'), {
        type: 'bar',
        data: {
            labels: top10.map(c => c.column.length > 12 ? c.column.substring(0, 9) + '...' : c.column),
            datasets: [{
                label: 'Completeness %',
                data: top10.map(c => c.rate.toFixed(1)),
                backgroundColor: top10.map(c => {
                    if (c.rate >= 90) return '#22c55e99';
                    if (c.rate >= 70) return '#f59e0b99';
                    return '#ef444499';
                }),
                borderWidth: 0
            }]
        },
        options: {
            ...getChartOptions('bar'),
            indexAxis: 'y',
            scales: {
                x: { max: 100, beginAtZero: true }
            }
        }
    });
}

function updateStatusDistChart() {
    const column = document.getElementById('statusColumnSelect')?.value;
    
    if (!column || !state.diffResult) return;
    
    const data = getAllData();
    const valueCounts = {};
    
    data.forEach(row => {
        const val = String(row[column] ?? '(empty)').trim() || '(empty)';
        valueCounts[val] = (valueCounts[val] || 0) + 1;
    });
    
    const sorted = Object.entries(valueCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8);
    
    const canvas = document.getElementById('statusDistChart');
    if (!canvas) return;
    
    if (analyticsCharts.statusDistChart) {
        analyticsCharts.statusDistChart.destroy();
    }
    
    const colors = getChartColors();
    
    analyticsCharts.statusDistChart = new Chart(canvas.getContext('2d'), {
        type: 'doughnut',
        data: {
            labels: sorted.map(([label]) => label.length > 15 ? label.substring(0, 12) + '...' : label),
            datasets: [{
                data: sorted.map(([, count]) => count),
                backgroundColor: colors.slice(0, sorted.length),
                borderWidth: 0
            }]
        },
        options: {
            ...getChartOptions('doughnut'),
            cutout: '60%'
        }
    });
}

function updateNumericStats() {
    const column = document.getElementById('numericColumnSelect')?.value;
    const container = document.getElementById('numericStats');
    
    if (!column || !container || !state.diffResult) return;
    
    const data = getAllData();
    const values = data
        .map(row => parseFloat(row[column]))
        .filter(v => !isNaN(v));
    
    if (values.length === 0) {
        container.innerHTML = '<p class="no-data">No numeric data available for this column</p>';
        return;
    }
    
    const sum = values.reduce((a, b) => a + b, 0);
    const avg = sum / values.length;
    const min = Math.min(...values);
    const max = Math.max(...values);
    const sorted = [...values].sort((a, b) => a - b);
    const median = sorted.length % 2 === 0 
        ? (sorted[sorted.length/2 - 1] + sorted[sorted.length/2]) / 2 
        : sorted[Math.floor(sorted.length/2)];
    
    container.innerHTML = `
        <div class="numeric-stat-grid">
            <div class="numeric-stat">
                <span class="stat-value">${values.length.toLocaleString()}</span>
                <span class="stat-label">Count</span>
            </div>
            <div class="numeric-stat">
                <span class="stat-value">${sum.toLocaleString(undefined, {maximumFractionDigits: 2})}</span>
                <span class="stat-label">Sum</span>
            </div>
            <div class="numeric-stat">
                <span class="stat-value">${avg.toLocaleString(undefined, {maximumFractionDigits: 2})}</span>
                <span class="stat-label">Average</span>
            </div>
            <div class="numeric-stat">
                <span class="stat-value">${median.toLocaleString(undefined, {maximumFractionDigits: 2})}</span>
                <span class="stat-label">Median</span>
            </div>
            <div class="numeric-stat">
                <span class="stat-value">${min.toLocaleString(undefined, {maximumFractionDigits: 2})}</span>
                <span class="stat-label">Min</span>
            </div>
            <div class="numeric-stat">
                <span class="stat-value">${max.toLocaleString(undefined, {maximumFractionDigits: 2})}</span>
                <span class="stat-label">Max</span>
            </div>
        </div>
    `;
}

function updateChangedRecordsTable() {
    const filter = document.getElementById('changedRecordsFilter')?.value || 'all';
    const count = parseInt(document.getElementById('changedRecordsCount')?.value || '10');
    
    const table = document.getElementById('changedRecordsTable');
    if (!table || !state.diffResult) return;
    
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    
    // Collect records based on filter
    let records = [];
    
    if (filter === 'all' || filter === 'modified') {
        state.diffResult.modified?.forEach(item => {
            records.push({ type: 'modified', data: item.updated, changes: item.changes, original: item.original });
        });
    }
    
    if (filter === 'all' || filter === 'added') {
        state.diffResult.added?.forEach(item => {
            records.push({ type: 'added', data: item.data });
        });
    }
    
    if (filter === 'all' || filter === 'removed') {
        state.diffResult.removed?.forEach(item => {
            records.push({ type: 'removed', data: item.data });
        });
    }
    
    // Take top N
    records = records.slice(0, count);
    
    // Get display columns (prioritize changed columns)
    const changedCols = new Set();
    records.forEach(r => {
        r.changes?.forEach(c => changedCols.add(c.column));
    });
    
    const displayCols = state.headers.slice(0, 6);
    
    // Build header
    thead.innerHTML = `
        <tr>
            <th>Type</th>
            ${displayCols.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
        </tr>
    `;
    
    // Build body
    tbody.innerHTML = records.map(record => {
        const changedColsForRow = record.changes?.map(c => c.column) || [];
        
        return `
            <tr class="${record.type}">
                <td><span class="row-status ${record.type}">${record.type}</span></td>
                ${displayCols.map(col => {
                    const value = record.data[col] ?? '';
                    const isChanged = changedColsForRow.includes(col);
                    
                    if (isChanged && record.original) {
                        const change = record.changes.find(c => c.column === col);
                        return `<td class="cell-changed" title="Changed from: ${escapeHtml(String(change?.oldValue ?? ''))}">
                            <span class="cell-value">${escapeHtml(String(value))}</span>
                            <i class="fas fa-exchange-alt change-indicator"></i>
                        </td>`;
                    }
                    
                    return `<td>${escapeHtml(String(value))}</td>`;
                }).join('')}
            </tr>
        `;
    }).join('');
    
    if (records.length === 0) {
        tbody.innerHTML = `<tr><td colspan="${displayCols.length + 1}" class="no-data">No records match the current filter</td></tr>`;
    }
}

function updateColumnHeatmap() {
    const showValues = document.getElementById('heatmapShowValues')?.checked ?? true;
    const colorScale = document.getElementById('heatmapColorScale')?.value || 'blue';
    
    const container = document.getElementById('columnHeatmap');
    if (!container || !state.diffResult) return;
    
    const columnChanges = state.diffResult.columnChanges || {};
    const maxChanges = Math.max(...Object.values(columnChanges), 1);
    
    const colorScales = {
        blue: (intensity) => `rgba(59, 130, 246, ${intensity})`,
        heat: (intensity) => `rgba(${Math.round(239 * intensity + 16)}, ${Math.round(68 * (1 - intensity) + 68)}, 68, ${Math.max(0.1, intensity)})`,
        green: (intensity) => `rgba(34, 197, 94, ${intensity})`
    };
    
    const getColor = colorScales[colorScale] || colorScales.blue;
    
    const sortedColumns = Object.entries(columnChanges)
        .sort((a, b) => b[1] - a[1]);
    
    container.innerHTML = `
        <div class="heatmap-grid">
            ${sortedColumns.map(([col, count]) => {
                const intensity = count / maxChanges;
                const bgColor = getColor(intensity);
                const textColor = intensity > 0.5 ? '#fff' : '#374151';
                
                return `
                    <div class="heatmap-cell" style="background: ${bgColor}; color: ${textColor}" title="${col}: ${count} changes">
                        <span class="heatmap-label">${col.length > 10 ? col.substring(0, 8) + '...' : col}</span>
                        ${showValues ? `<span class="heatmap-value">${count}</span>` : ''}
                    </div>
                `;
            }).join('')}
        </div>
        <div class="heatmap-legend">
            <span class="legend-label">Low</span>
            <div class="legend-gradient" style="background: linear-gradient(to right, ${getColor(0.1)}, ${getColor(1)})"></div>
            <span class="legend-label">High</span>
        </div>
    `;
}

function updateAllAnalyticsCharts() {
    if (document.getElementById('expandedAnalytics').hidden) return;
    initializeExpandedAnalytics();
}

function updateDiffChart() {
    const ctx = document.getElementById('diffChart').getContext('2d');
    
    if (state.chart) {
        state.chart.destroy();
    }
    
    // When showMovedRows is off, add moved count to unchanged
    const displayUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + state.diffResult.moved.length;
    
    // Build labels and data based on showMovedRows setting
    const labels = state.showMovedRows 
        ? ['Unchanged', 'Modified', 'Added', 'Removed', 'Moved']
        : ['Unchanged', 'Modified', 'Added', 'Removed'];
    
    const data = state.showMovedRows 
        ? [
            state.diffResult.unchanged.length,
            state.diffResult.modified.length,
            state.diffResult.added.length,
            state.diffResult.removed.length,
            state.diffResult.moved.length
        ]
        : [
            displayUnchangedCount,
            state.diffResult.modified.length,
            state.diffResult.added.length,
            state.diffResult.removed.length
        ];
    
    const colors = state.showMovedRows 
        ? ['#6b7280', '#f59e0b', '#22c55e', '#ef4444', '#8b5cf6']
        : ['#6b7280', '#f59e0b', '#22c55e', '#ef4444'];
    
    state.chart = new Chart(ctx, {
        type: 'doughnut',
        data: {
            labels,
            datasets: [{
                data,
                backgroundColor: colors,
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        usePointStyle: true,
                        padding: 20
                    }
                }
            }
        }
    });
}

function updateColumnChanges() {
    const container = document.getElementById('column-changes');
    const columnChanges = state.diffResult.columnChanges || {};
    const maxChanges = Math.max(...Object.values(columnChanges), 1);
    
    container.innerHTML = Object.entries(columnChanges)
        .sort((a, b) => b[1] - a[1])
        .map(([column, count]) => `
            <div class="column-change-item">
                <span class="column-name">${escapeHtml(column)}</span>
                <div class="column-bar">
                    <div class="column-bar-fill" style="width: ${(count / maxChanges) * 100}%"></div>
                </div>
                <span class="column-count">${count}</span>
            </div>
        `).join('');
}

function updateRowsPanel() {
    const table = document.getElementById('row-table');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const searchTerm = document.getElementById('rowSearch')?.value?.toLowerCase() || '';
    
    // Collect rows based on status filter
    let rows = [];
    
    if (state.currentFilter === 'all' || state.currentFilter === 'unchanged') {
        state.diffResult.unchanged.forEach(item => {
            rows.push({ status: 'unchanged', data: item.data, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'modified') {
        state.diffResult.modified.forEach(item => {
            rows.push({ status: 'modified', data: item.updated, original: item.original, changes: item.changes, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'added') {
        state.diffResult.added.forEach(item => {
            rows.push({ status: 'added', data: item.data, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'removed') {
        state.diffResult.removed.forEach(item => {
            rows.push({ status: 'removed', data: item.data, key: item.key });
        });
    }
    
    // Only include moved rows if the setting is enabled
    if (state.showMovedRows && (state.currentFilter === 'all' || state.currentFilter === 'moved')) {
        state.diffResult.moved.forEach(item => {
            rows.push({ status: 'moved', data: item.data, key: item.key, originalPos: item.originalPosition, newPos: item.updatedPosition });
        });
    }
    
    // Apply search filter
    if (searchTerm) {
        rows = rows.filter(row => {
            const values = Object.values(row.data).map(v => String(v).toLowerCase());
            return values.some(v => v.includes(searchTerm));
        });
    }
    
    // Apply column filters
    const activeColumnFilters = Object.entries(state.columnFilters).filter(([k, v]) => v && v.trim() !== '');
    if (activeColumnFilters.length > 0) {
        rows = rows.filter(row => {
            return activeColumnFilters.every(([column, filterValue]) => {
                const cellValue = String(row.data[column] ?? '').toLowerCase();
                const filter = filterValue.toLowerCase();
                return cellValue.includes(filter);
            });
        });
    }
    
    // Build header
    const displayHeaders = state.currentDensity === 'minimal' 
        ? state.headers.slice(0, 5)
        : state.headers;
    
    // Header row
    let headerHTML = `
        <tr>
            <th>Status</th>
            ${displayHeaders.map(h => `<th>${escapeHtml(h)}</th>`).join('')}
            ${state.currentDensity === 'minimal' && state.headers.length > 5 ? '<th>...</th>' : ''}
        </tr>
    `;
    
    // Filter row (if visible)
    if (state.columnFiltersVisible) {
        headerHTML += `
            <tr class="column-filter-row">
                <th></th>
                ${displayHeaders.map(h => {
                    const currentValue = state.columnFilters[h] || '';
                    const hasValue = currentValue.trim() !== '';
                    return `<th>
                        <div class="column-filter-wrapper">
                            <input type="text" 
                                class="column-filter ${hasValue ? 'has-value' : ''}" 
                                data-column="${escapeHtml(h)}" 
                                placeholder="Filter..."
                                value="${escapeHtml(currentValue)}">
                            <i class="fas fa-filter filter-icon"></i>
                        </div>
                    </th>`;
                }).join('')}
                ${state.currentDensity === 'minimal' && state.headers.length > 5 ? '<th></th>' : ''}
            </tr>
        `;
    }
    
    thead.innerHTML = headerHTML;
    
    // Attach filter input listeners
    if (state.columnFiltersVisible) {
        thead.querySelectorAll('.column-filter').forEach(input => {
            input.addEventListener('input', (e) => {
                const column = e.target.dataset.column;
                const cursorPos = e.target.selectionStart;
                state.columnFilters[column] = e.target.value;
                state.currentPage = 1;
                
                // Update only the body, not the header (to keep focus)
                updateRowsPanelBody();
                updateFilterBadge();
                
                // Update the has-value class on this input
                e.target.classList.toggle('has-value', e.target.value.trim() !== '');
            });
        });
    }
    
    // Pagination
    const totalPages = Math.ceil(rows.length / state.rowsPerPage);
    state.currentPage = Math.min(state.currentPage, totalPages || 1);
    const startIdx = (state.currentPage - 1) * state.rowsPerPage;
    const pageRows = rows.slice(startIdx, startIdx + state.rowsPerPage);
    
    // Build body
    tbody.innerHTML = pageRows.map(row => {
        const changedColumns = row.changes ? row.changes.map(c => c.column) : [];
        
        return `
            <tr class="${row.status}">
                <td><span class="row-status ${row.status}">${row.status}</span></td>
                ${displayHeaders.map(h => {
                    const value = row.data[h] ?? '';
                    
                    if (state.currentDensity === 'verbose' && changedColumns.includes(h)) {
                        const change = row.changes.find(c => c.column === h);
                        return `<td class="cell-changed">
                            <span class="cell-old">${escapeHtml(String(change.oldValue ?? ''))}</span>
                            <span class="cell-new">${escapeHtml(String(change.newValue ?? ''))}</span>
                        </td>`;
                    }
                    
                    if (changedColumns.includes(h)) {
                        return `<td class="cell-changed">${escapeHtml(String(value))}</td>`;
                    }
                    
                    return `<td>${escapeHtml(String(value))}</td>`;
                }).join('')}
                ${state.currentDensity === 'minimal' && state.headers.length > 5 ? '<td>...</td>' : ''}
            </tr>
        `;
    }).join('');
    
    // Update pagination info
    document.getElementById('pageInfo').textContent = `Page ${state.currentPage} of ${totalPages || 1} (${rows.length} rows)`;
    
    // Update filter badge
    updateFilterBadge();
}

function updateFilterBadge() {
    const activeFilters = Object.values(state.columnFilters).filter(v => v && v.trim() !== '').length;
    const btnText = document.getElementById('filterBtnText');
    const btn = document.getElementById('toggleColumnFilters');
    
    if (activeFilters > 0) {
        btnText.innerHTML = state.columnFiltersVisible 
            ? `Hide Filters <span class="filter-badge">${activeFilters}</span>`
            : `Add Filters <span class="filter-badge">${activeFilters}</span>`;
    } else {
        btnText.textContent = state.columnFiltersVisible ? 'Hide Filters' : 'Add Filters';
    }
}

// Separate function to update only the table body (preserves filter input focus)
function updateRowsPanelBody() {
    const table = document.getElementById('row-table');
    const tbody = table.querySelector('tbody');
    const searchTerm = document.getElementById('rowSearch')?.value?.toLowerCase() || '';
    
    // Collect rows based on status filter
    let rows = [];
    
    if (state.currentFilter === 'all' || state.currentFilter === 'unchanged') {
        state.diffResult.unchanged.forEach(item => {
            rows.push({ status: 'unchanged', data: item.data, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'modified') {
        state.diffResult.modified.forEach(item => {
            rows.push({ status: 'modified', data: item.updated, original: item.original, changes: item.changes, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'added') {
        state.diffResult.added.forEach(item => {
            rows.push({ status: 'added', data: item.data, key: item.key });
        });
    }
    
    if (state.currentFilter === 'all' || state.currentFilter === 'removed') {
        state.diffResult.removed.forEach(item => {
            rows.push({ status: 'removed', data: item.data, key: item.key });
        });
    }
    
    // Only include moved rows if the setting is enabled
    if (state.showMovedRows && (state.currentFilter === 'all' || state.currentFilter === 'moved')) {
        state.diffResult.moved.forEach(item => {
            rows.push({ status: 'moved', data: item.data, key: item.key, originalPos: item.originalPosition, newPos: item.updatedPosition });
        });
    }
    
    // Apply search filter
    if (searchTerm) {
        rows = rows.filter(row => {
            const values = Object.values(row.data).map(v => String(v).toLowerCase());
            return values.some(v => v.includes(searchTerm));
        });
    }
    
    // Apply column filters
    const activeColumnFilters = Object.entries(state.columnFilters).filter(([k, v]) => v && v.trim() !== '');
    if (activeColumnFilters.length > 0) {
        rows = rows.filter(row => {
            return activeColumnFilters.every(([column, filterValue]) => {
                const cellValue = String(row.data[column] ?? '').toLowerCase();
                const filter = filterValue.toLowerCase();
                return cellValue.includes(filter);
            });
        });
    }
    
    const displayHeaders = state.currentDensity === 'minimal' 
        ? state.headers.slice(0, 5)
        : state.headers;
    
    // Pagination
    const totalPages = Math.ceil(rows.length / state.rowsPerPage);
    state.currentPage = Math.min(state.currentPage, totalPages || 1);
    const startIdx = (state.currentPage - 1) * state.rowsPerPage;
    const pageRows = rows.slice(startIdx, startIdx + state.rowsPerPage);
    
    // Build body only
    tbody.innerHTML = pageRows.map(row => {
        const changedColumns = row.changes ? row.changes.map(c => c.column) : [];
        
        return `
            <tr class="${row.status}">
                <td><span class="row-status ${row.status}">${row.status}</span></td>
                ${displayHeaders.map(h => {
                    const value = row.data[h] ?? '';
                    
                    if (state.currentDensity === 'verbose' && changedColumns.includes(h)) {
                        const change = row.changes.find(c => c.column === h);
                        return `<td class="cell-changed">
                            <span class="cell-old">${escapeHtml(String(change.oldValue ?? ''))}</span>
                            <span class="cell-new">${escapeHtml(String(change.newValue ?? ''))}</span>
                        </td>`;
                    }
                    
                    if (changedColumns.includes(h)) {
                        return `<td class="cell-changed">${escapeHtml(String(value))}</td>`;
                    }
                    
                    return `<td>${escapeHtml(String(value))}</td>`;
                }).join('')}
                ${state.currentDensity === 'minimal' && state.headers.length > 5 ? '<td>...</td>' : ''}
            </tr>
        `;
    }).join('');
    
    // Update pagination info
    document.getElementById('pageInfo').textContent = `Page ${state.currentPage} of ${totalPages || 1} (${rows.length} rows)`;
}

function updateHighlightPanel() {
    if (!state.diffResult) return;
    
    // Update filenames
    document.getElementById('highlight-original-name').textContent = state.file1?.name || 'Original';
    document.getElementById('highlight-updated-name').textContent = state.file2?.name || 'Updated';
    
    // Update column filter label
    updateColumnFilterLabel();
    
    // Get current filter state from toggles
    const filters = {};
    document.querySelectorAll('#highlight-panel .legend-item .toggle-switch input').forEach(input => {
        if (input.dataset.filter) {
            filters[input.dataset.filter] = input.checked;
        }
    });
    
    // Get custom colors
    const colors = {
        modified: document.getElementById('color-modified')?.value || '#f59e0b',
        added: document.getElementById('color-added')?.value || '#22c55e',
        removed: document.getElementById('color-removed')?.value || '#ef4444',
        unchanged: document.getElementById('color-unchanged')?.value || '#6b7280'
    };
    
    // Build original side
    const originalContent = document.getElementById('highlight-original');
    const updatedContent = document.getElementById('highlight-updated');
    
    let originalHTML = '';
    let updatedHTML = '';
    let lineNum = 1;
    
    // Get primary key column for display
    const keyCol = state.diffResult.meta?.primaryKeys?.[0] || state.headers[0];
    
    // Check if column filter is active
    const hasColumnFilter = state.highlightColumnFilter.length > 0;
    
    // Filter modified rows based on column filter
    const filteredModified = hasColumnFilter 
        ? state.diffResult.modified.filter(item => {
            // Check if any of the changes are in the selected columns
            return item.changes.some(change => 
                state.highlightColumnFilter.includes(change.column)
            );
        })
        : state.diffResult.modified;
    
    // When column filter is active, modified rows that don't match become "unchanged"
    const columnFilterExcludedModified = hasColumnFilter 
        ? state.diffResult.modified.filter(item => {
            return !item.changes.some(change => 
                state.highlightColumnFilter.includes(change.column)
            );
        })
        : [];
    
    // Process unchanged rows
    state.diffResult.unchanged.forEach(item => {
        const hidden = !filters.unchanged ? 'hidden' : '';
        const keyVal = item.data[keyCol] || '';
        const dataPreview = getRowPreview(item.data);
        
        originalHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
        </div>`;
        
        updatedHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
        </div>`;
        
        lineNum++;
    });
    
    // When showMovedRows is off, treat moved rows as unchanged
    if (!state.showMovedRows) {
        state.diffResult.moved?.forEach(item => {
            const hidden = !filters.unchanged ? 'hidden' : '';
            const keyVal = item.data[keyCol] || '';
            const dataPreview = getRowPreview(item.data);
            
            originalHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
                <span class="highlight-row-number">${lineNum}</span>
                <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
                <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
            </div>`;
            
            updatedHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
                <span class="highlight-row-number">${lineNum}</span>
                <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
                <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
            </div>`;
            
            lineNum++;
        });
    }
    
    // Process modified rows that were excluded by column filter as unchanged
    columnFilterExcludedModified.forEach(item => {
        const hidden = !filters.unchanged ? 'hidden' : '';
        const keyVal = item.original[keyCol] || '';
        const dataPreview = getRowPreview(item.updated);
        
        originalHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
        </div>`;
        
        updatedHTML += `<div class="highlight-row unchanged ${hidden}" style="border-left-color: ${colors.unchanged}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
        </div>`;
        
        lineNum++;
    });
    
    // Process filtered modified rows
    filteredModified.forEach(item => {
        const hidden = !filters.modified ? 'hidden' : '';
        const keyVal = item.original[keyCol] || '';
        const originalPreview = getRowPreview(item.original);
        
        // Filter changes to only show selected columns
        const filteredChanges = hasColumnFilter 
            ? item.changes.filter(c => state.highlightColumnFilter.includes(c.column))
            : item.changes;
        
        const updatedPreview = getRowPreviewWithChanges(item.updated, filteredChanges);
        
        originalHTML += `<div class="highlight-row modified ${hidden}" style="border-left-color: ${colors.modified}; background: ${hexToRgba(colors.modified, 0.1)}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(originalPreview)}</span>
            <span class="highlight-row-badge modified" style="background: ${colors.modified}">Modified</span>
        </div>`;
        
        updatedHTML += `<div class="highlight-row modified ${hidden}" style="border-left-color: ${colors.modified}; background: ${hexToRgba(colors.modified, 0.1)}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${updatedPreview}</span>
            <span class="highlight-row-badge modified" style="background: ${colors.modified}">Modified</span>
        </div>`;
        
        lineNum++;
    });
    
    // Process removed rows (only in original)
    state.diffResult.removed.forEach(item => {
        const hidden = !filters.removed ? 'hidden' : '';
        const keyVal = item.data[keyCol] || '';
        const dataPreview = getRowPreview(item.data);
        
        originalHTML += `<div class="highlight-row removed ${hidden}" style="border-left-color: ${colors.removed}; background: ${hexToRgba(colors.removed, 0.1)}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
            <span class="highlight-row-badge removed" style="background: ${colors.removed}">Removed</span>
        </div>`;
        
        updatedHTML += `<div class="highlight-row removed ${hidden}" style="border-left-color: ${colors.removed}; background: ${hexToRgba(colors.removed, 0.1)}; opacity: 0.4">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">—</span>
            <span class="highlight-row-data" style="font-style: italic; color: var(--text-muted)">Record removed</span>
        </div>`;
        
        lineNum++;
    });
    
    // Process added rows (only in updated)
    state.diffResult.added.forEach(item => {
        const hidden = !filters.added ? 'hidden' : '';
        const keyVal = item.data[keyCol] || '';
        const dataPreview = getRowPreview(item.data);
        
        originalHTML += `<div class="highlight-row added ${hidden}" style="border-left-color: ${colors.added}; background: ${hexToRgba(colors.added, 0.1)}; opacity: 0.4">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">—</span>
            <span class="highlight-row-data" style="font-style: italic; color: var(--text-muted)">New record</span>
        </div>`;
        
        updatedHTML += `<div class="highlight-row added ${hidden}" style="border-left-color: ${colors.added}; background: ${hexToRgba(colors.added, 0.1)}">
            <span class="highlight-row-number">${lineNum}</span>
            <span class="highlight-row-key">${escapeHtml(String(keyVal))}</span>
            <span class="highlight-row-data">${escapeHtml(dataPreview)}</span>
            <span class="highlight-row-badge added" style="background: ${colors.added}">Added</span>
        </div>`;
        
        lineNum++;
    });
    
    originalContent.innerHTML = originalHTML;
    updatedContent.innerHTML = updatedHTML;
    
    // Update stats - account for column filter
    const displayedModifiedCount = filteredModified.length;
    document.getElementById('hl-modified-count').textContent = displayedModifiedCount;
    document.getElementById('hl-added-count').textContent = state.diffResult.added.length;
    document.getElementById('hl-removed-count').textContent = state.diffResult.removed.length;
    
    // When showMovedRows is off, count moved rows as unchanged
    // Also add rows excluded by column filter
    let hlUnchangedCount = state.showMovedRows 
        ? state.diffResult.unchanged.length 
        : state.diffResult.unchanged.length + (state.diffResult.moved?.length || 0);
    hlUnchangedCount += columnFilterExcludedModified.length;
    document.getElementById('hl-unchanged-count').textContent = hlUnchangedCount;
    
    // Update sync scroll state
    updateSyncScrollState();
}

function getRowPreview(row) {
    const values = state.headers.slice(0, 4).map(h => row[h] ?? '');
    let preview = values.join(' | ');
    if (state.headers.length > 4) preview += ' ...';
    return preview;
}

function getRowPreviewWithChanges(row, changes) {
    const changedCols = changes.map(c => c.column);
    const parts = state.headers.slice(0, 4).map(h => {
        const val = row[h] ?? '';
        if (changedCols.includes(h)) {
            return `<span class="changed-cell">${escapeHtml(String(val))}</span>`;
        }
        return escapeHtml(String(val));
    });
    let preview = parts.join(' | ');
    if (state.headers.length > 4) preview += ' ...';
    return preview;
}

function hexToRgba(hex, alpha) {
    const r = parseInt(hex.slice(1, 3), 16);
    const g = parseInt(hex.slice(3, 5), 16);
    const b = parseInt(hex.slice(5, 7), 16);
    return `rgba(${r}, ${g}, ${b}, ${alpha})`;
}

function updateCellsPanel() {
    const originalContent = document.getElementById('original-content');
    const updatedContent = document.getElementById('updated-content');
    
    // Build unified list of all rows for side-by-side display
    let allRows = [];
    let lineNum = 1;
    
    // Add unchanged rows
    state.diffResult.unchanged.forEach((item) => {
        allRows.push({ 
            type: 'unchanged', 
            originalData: item.data, 
            updatedData: item.data, 
            line: lineNum++,
            key: item.key
        });
    });
    
    // When showMovedRows is off, treat moved rows as unchanged
    if (!state.showMovedRows) {
        state.diffResult.moved?.forEach((item) => {
            allRows.push({ 
                type: 'unchanged', 
                originalData: item.data, 
                updatedData: item.data, 
                line: lineNum++,
                key: item.key
            });
        });
    }
    
    // Add modified rows
    state.diffResult.modified.forEach((item) => {
        allRows.push({ 
            type: 'modified', 
            originalData: item.original, 
            updatedData: item.updated, 
            changes: item.changes,
            line: lineNum++,
            key: item.key
        });
    });
    
    // Add removed rows
    state.diffResult.removed.forEach((item) => {
        allRows.push({ 
            type: 'removed', 
            originalData: item.data, 
            updatedData: null, 
            line: lineNum++,
            key: item.key
        });
    });
    
    // Add added rows
    state.diffResult.added.forEach((item) => {
        allRows.push({ 
            type: 'added', 
            originalData: null, 
            updatedData: item.data, 
            line: lineNum++,
            key: item.key
        });
    });
    
    // Store for navigation
    state.cellViewRows = allRows;
    state.cellViewChangeIndices = allRows
        .map((row, idx) => row.type !== 'unchanged' ? idx : -1)
        .filter(idx => idx !== -1);
    
    if (state.currentChangeIndex === undefined) {
        state.currentChangeIndex = 0;
    }
    
    // Build original side HTML
    originalContent.innerHTML = allRows.map((row, idx) => {
        const isChange = row.type !== 'unchanged';
        const dataId = `orig-row-${idx}`;
        
        if (row.originalData) {
            const rowStr = formatRowForCellView(row.originalData, row.changes, 'original');
            return `<div class="diff-line ${row.type}" id="${dataId}" data-index="${idx}">
                <span class="line-number">${row.line}</span>
                <span class="diff-line-content">${rowStr}</span>
                ${isChange ? `<span class="diff-line-badge ${row.type}">${row.type}</span>` : ''}
            </div>`;
        } else {
            return `<div class="diff-line ${row.type} empty-placeholder" id="${dataId}" data-index="${idx}">
                <span class="line-number">${row.line}</span>
                <span class="diff-line-content empty">— New record in updated file —</span>
            </div>`;
        }
    }).join('');
    
    // Build updated side HTML
    updatedContent.innerHTML = allRows.map((row, idx) => {
        const isChange = row.type !== 'unchanged';
        const dataId = `updated-row-${idx}`;
        
        if (row.updatedData) {
            const rowStr = formatRowForCellView(row.updatedData, row.changes, 'updated');
            return `<div class="diff-line ${row.type}" id="${dataId}" data-index="${idx}">
                <span class="line-number">${row.line}</span>
                <span class="diff-line-content">${rowStr}</span>
                ${isChange ? `<span class="diff-line-badge ${row.type}">${row.type}</span>` : ''}
            </div>`;
        } else {
            return `<div class="diff-line ${row.type} empty-placeholder" id="${dataId}" data-index="${idx}">
                <span class="line-number">${row.line}</span>
                <span class="diff-line-content empty">— Record removed —</span>
            </div>`;
        }
    }).join('');
    
    // Update change counter
    const totalChanges = state.cellViewChangeIndices.length;
    updateChangeCounter();
    
    // Scroll to first change if there is one
    if (totalChanges > 0 && state.currentChangeIndex === 0) {
        scrollToChange(0);
    }
}

function formatRowForCellView(data, changes, side) {
    const changedCols = changes ? changes.map(c => c.column) : [];
    const keyCol = state.diffResult.meta?.primaryKeys?.[0] || state.headers[0];
    
    // Show key and a few important columns
    const parts = [];
    
    // Always show key first
    const keyVal = data[keyCol] ?? '';
    parts.push(`<strong>${escapeHtml(String(keyVal))}</strong>`);
    
    // Show other columns (limit to prevent overflow)
    const otherHeaders = state.headers.filter(h => h !== keyCol).slice(0, 5);
    otherHeaders.forEach(h => {
        const val = data[h] ?? '';
        const isChanged = changedCols.includes(h);
        if (isChanged) {
            parts.push(`<span class="cell-highlight">${escapeHtml(h)}: ${escapeHtml(String(val))}</span>`);
        } else {
            parts.push(`${escapeHtml(h)}: ${escapeHtml(String(val))}`);
        }
    });
    
    if (state.headers.length > 6) {
        parts.push('...');
    }
    
    return parts.join(' | ');
}

function updateChangeCounter() {
    const total = state.cellViewChangeIndices?.length || 0;
    const current = total > 0 ? state.currentChangeIndex + 1 : 0;
    document.getElementById('changeCounter').textContent = `Change ${current} of ${total}`;
}

function scrollToChange(changeIndex) {
    if (!state.cellViewChangeIndices || state.cellViewChangeIndices.length === 0) return;
    
    // Clamp index
    changeIndex = Math.max(0, Math.min(changeIndex, state.cellViewChangeIndices.length - 1));
    state.currentChangeIndex = changeIndex;
    
    const rowIndex = state.cellViewChangeIndices[changeIndex];
    
    // Remove highlight from all rows
    document.querySelectorAll('.diff-line.current-change').forEach(el => {
        el.classList.remove('current-change');
    });
    
    // Highlight and scroll to the row in both panels
    const origRow = document.getElementById(`orig-row-${rowIndex}`);
    const updatedRow = document.getElementById(`updated-row-${rowIndex}`);
    
    if (origRow) {
        origRow.classList.add('current-change');
        origRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
    if (updatedRow) {
        updatedRow.classList.add('current-change');
    }
    
    updateChangeCounter();
}

function navigateToNextChange() {
    if (!state.cellViewChangeIndices || state.cellViewChangeIndices.length === 0) return;
    const nextIndex = (state.currentChangeIndex + 1) % state.cellViewChangeIndices.length;
    scrollToChange(nextIndex);
}

function navigateToPrevChange() {
    if (!state.cellViewChangeIndices || state.cellViewChangeIndices.length === 0) return;
    const prevIndex = state.currentChangeIndex - 1;
    scrollToChange(prevIndex < 0 ? state.cellViewChangeIndices.length - 1 : prevIndex);
}

// ========================================
// Modals
// ========================================

function initializeModals() {
    // Help modal
    document.getElementById('helpBtn')?.addEventListener('click', () => {
        document.getElementById('helpModal').hidden = false;
    });
    
    document.getElementById('closeHelp')?.addEventListener('click', () => {
        document.getElementById('helpModal').hidden = true;
    });
    
    // Close modals on backdrop click
    document.querySelectorAll('.modal-backdrop').forEach(backdrop => {
        backdrop.addEventListener('click', () => {
            backdrop.closest('.modal').hidden = true;
        });
    });
    
    // Close modal on Escape key
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            document.querySelectorAll('.modal').forEach(modal => {
                modal.hidden = true;
            });
        }
    });
}

// ========================================
// Utility Functions
// ========================================

function showLoading(show) {
    document.getElementById('loading').hidden = !show;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
