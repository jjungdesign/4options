class Spreadsheet {
    constructor() {
        this.rows = 20;
        this.columns = 8;
        this.selectedCell = null;
        this.selectedCells = [];
        this.data = {};
        this.columnWidths = {};
        this.cellData = {};
        this.appColumns = new Set();
        this.originalUploadedData = null;
        this.originalFullData = null; 
        this.uploadedHeaders = [];
        this.columnCount = 0; // Track number of columns for credit calculation
        // Load saved version or default to 1
        this.currentVersion = parseInt(localStorage.getItem('selectedVersion') || '1');
        // Load test mode state (default to true for Option 2, and start with toggle ON)
        this.isTestMode = localStorage.getItem('testMode') === 'false' ? false : true;
        // Load selected row range from localStorage
        this.selectedRowRange = localStorage.getItem('selectedRowRange') || null;
        
                       this.init();
               this.setupVersionToggle(); // Add version toggle setup
               this.setupTestModeToggle(); // Add test mode toggle setup
               this.loadSavedVersion(); // Apply saved version on load
               this.setupCreditPillDropdown(); // Setup credit pill dropdown
               this.showTestModeToastOnLoad(); // Show test mode toast on load
    }

    init() {
        this.generateRows();
        this.setupEventListeners();
        this.setupKeyboardNavigation();
        this.setupSearch();
        this.setupActionButtons();
        this.setupStateSaving();
    }

    generateRows() {
        const tbody = document.getElementById('spreadsheet-body');
        tbody.innerHTML = '';

        // Update header row with correct number of columns
        this.updateHeaderRow();

        // Generate exactly 20 empty rows for v2
        for (let row = 1; row <= 20; row++) {
            const tr = document.createElement('tr');
            tr.className = 'data-row';
            tr.dataset.row = row;

            // Row header
            const th = document.createElement('th');
            th.className = 'row-header';
            th.textContent = row;
            tr.appendChild(th);

            // Data cells - always generate exactly 8 columns with empty values
            for (let col = 0; col < 8; col++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.dataset.row = row;
                td.dataset.col = String.fromCharCode(65 + col); // A, B, C, etc.
                
                const input = document.createElement('input');
                input.type = 'text';
                input.placeholder = '';
                input.value = ''; // Always empty for v2
                
                td.appendChild(input);
                tr.appendChild(td);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

            tbody.appendChild(tr);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // No test mode row for v2
    }



    updateHeaderRow() {
        const headerRow = document.querySelector('.header-row');
        if (!headerRow) return;

        // Keep the corner cell
        const cornerCell = headerRow.querySelector('.corner-cell');
        headerRow.innerHTML = '';
        headerRow.appendChild(cornerCell);

        // Always add exactly 8 column headers - all start with A, B, C, D, E, F, G, H for v2
        for (let col = 0; col < 8; col++) {
            const th = document.createElement('th');
            th.className = 'column-header';
            th.dataset.col = String.fromCharCode(65 + col);
            th.textContent = String.fromCharCode(65 + col); // A, B, C, D, E, F, G, H
            
            headerRow.appendChild(th);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    setupEventListeners() {
        // Cell click events
        document.addEventListener('click', (e) => {
            if (e.target.closest('.cell')) {
                const cell = e.target.closest('.cell');
                this.selectCell(cell);
                
                // Special handling for Option 2 Scale Mode: fill rows 11-20 when any cell in rows 11-20 is clicked
                const rowNum = parseInt(cell.dataset.row);
                if (this.currentVersion === 2 && !this.isTestMode && rowNum >= 11 && rowNum <= 20) {
                    this.fillRows11To20();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (e.target.closest('.column-header')) {
                const header = e.target.closest('.column-header');
                this.selectColumn(header);
            } else if (e.target.closest('.row-header')) {
                const header = e.target.closest('.row-header');
                this.selectRow(header);
            } else {
                this.clearSelection();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Cell input events
        document.addEventListener('input', (e) => {
            if (e.target.closest('.cell input')) {
                const input = e.target;
                const cell = input.closest('.cell');
                const cellKey = `${cell.dataset.row}-${cell.dataset.col}`;
                this.data[cellKey] = input.value;
                
                // Save state when data changes
                this.saveState();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Double click to edit
        document.addEventListener('dblclick', (e) => {
            if (e.target.closest('.cell')) {
                const cell = e.target.closest('.cell');
                const input = cell.querySelector('input');
                if (input) {
                    input.focus();
                    input.select();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Context menu
        document.addEventListener('contextmenu', (e) => {
            e.preventDefault();
            if (e.target.closest('.cell')) {
                const cell = e.target.closest('.cell');
                this.selectCell(cell);
                this.showContextMenu(e);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }

    setupKeyboardNavigation() {
        document.addEventListener('keydown', (e) => {
            if (!this.currentCell) return;

            const currentRow = parseInt(this.currentCell.dataset.row);
            const currentCol = this.currentCell.dataset.col;
            let nextCell = null;

            switch (e.key) {
                case 'ArrowUp':
                    if (currentRow > 1) {
                        nextCell = document.querySelector(`[data-row="${currentRow - 1}"][data-col="${currentCol}"]`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
                case 'ArrowDown':
                    if (currentRow < this.rows) {
                        nextCell = document.querySelector(`[data-row="${currentRow + 1}"][data-col="${currentCol}"]`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
                case 'ArrowLeft':
                    if (currentCol.charCodeAt(0) > 65) { // A = 65
                        const prevCol = String.fromCharCode(currentCol.charCodeAt(0) - 1);
                        nextCell = document.querySelector(`[data-row="${currentRow}"][data-col="${prevCol}"]`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
                case 'ArrowRight':
                    if (currentCol.charCodeAt(0) < 65 + this.columns - 1) {
                        const nextCol = String.fromCharCode(currentCol.charCodeAt(0) + 1);
                        nextCell = document.querySelector(`[data-row="${currentRow}"][data-col="${nextCol}"]`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
                case 'Enter':
                    e.preventDefault();
                    if (e.shiftKey) {
                        // Move up
                        if (currentRow > 1) {
                            nextCell = document.querySelector(`[data-row="${currentRow - 1}"][data-col="${currentCol}"]`);
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    } else {
                        // Move down
                        if (currentRow < this.rows) {
                            nextCell = document.querySelector(`[data-row="${currentRow + 1}"][data-col="${currentCol}"]`);
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
                case 'Tab':
                    e.preventDefault();
                    if (e.shiftKey) {
                        // Move left
                        if (currentCol.charCodeAt(0) > 65) {
                            const prevCol = String.fromCharCode(currentCol.charCodeAt(0) - 1);
                            nextCell = document.querySelector(`[data-row="${currentRow}"][data-col="${prevCol}"]`);
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    } else {
                        // Move right
                        if (currentCol.charCodeAt(0) < 65 + this.columns - 1) {
                            const nextCol = String.fromCharCode(currentCol.charCodeAt(0) + 1);
                            nextCell = document.querySelector(`[data-row="${currentRow}"][data-col="${nextCol}"]`);
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    break;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

            if (nextCell) {
                this.selectCell(nextCell);
                const input = nextCell.querySelector('input');
                if (input) {
                    input.focus();
                    input.select();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }

    setupSearch() {
        const searchInput = document.querySelector('.search-input');
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase();
            this.searchCells(searchTerm);
        });
    }

    setupActionButtons() {
        // Free rows toggle is now static - no event listener needed

        // Add column button
        const addColumnBtn = document.querySelector('.add-column-btn');
        addColumnBtn.addEventListener('click', () => {
            this.showAddColumnPopover();
        });

        // Run all button
        const runAllBtn = document.getElementById('run-all-btn');
        runAllBtn.addEventListener('click', () => {
            this.runAll();
        });

        // Confirm output button
        const confirmOutputBtn = document.getElementById('confirm-output-btn');
        confirmOutputBtn.addEventListener('click', () => {
            this.confirmOutput();
        });

        // Floating action buttons
        const floatingActions = document.querySelectorAll('.action-bar .action-btn');
        floatingActions.forEach((action, index) => {
            action.addEventListener('click', () => {
                this.handleFloatingAction(index);
            });
        });

        // Testing action buttons
        const approveBtn = document.getElementById('approve-btn');
        const rejectBtn = document.getElementById('reject-btn');
        
        if (approveBtn) {
            approveBtn.addEventListener('click', () => {
                this.approveTestRun();
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        if (rejectBtn) {
            rejectBtn.addEventListener('click', () => {
                this.rejectTestRun();
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Import All button setup is now handled in setupImportAllButton() after button creation
    }

    setupImportAllButton() {
        // Set up Import All button event listener (called after button is created)
        setTimeout(() => {
            const importAllBtn = document.getElementById('import-all-btn');
            if (importAllBtn) {
                console.log('✅ Import All button found after creation:', importAllBtn);
                console.log('📊 Import All button disabled state:', importAllBtn.disabled);
                
                // Remove any existing event listeners
                importAllBtn.onclick = null;
                
                // Add click event listener with error handling
                const clickHandler = (e) => {
                    console.log('🖱️ Import All button clicked!', e);
                    console.log('📊 Button disabled state on click:', importAllBtn.disabled);
                    
                    // Prevent default and stop propagation
                    e.preventDefault();
                    e.stopPropagation();
                    
                    if (importAllBtn.disabled) {
                        console.log('⚠️ Button is disabled, cannot proceed');
                        return false;
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    
                    console.log('🚀 Proceeding with import...');
                    this.importAllData();
                    return false;
                };
                
                importAllBtn.addEventListener('click', clickHandler);
                
                // Also add onclick as backup
                importAllBtn.onclick = clickHandler;
                
                console.log('✅ Event listeners attached to Import All button');
            } else {
                console.log('❌ Import All button not found after creation!');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        }, 100);
    }

    setupStateSaving() {
        // Save state before page unload (refresh, close, etc.)
        window.addEventListener('beforeunload', () => {
            this.saveState();
        });

        // Save state periodically (every 30 seconds)
        setInterval(() => {
            this.saveState();
        }, 30000);

        // Save state when test mode changes
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            const observer = new MutationObserver(() => {
                this.saveState();
            });
            observer.observe(testModeContainer, { attributes: true, attributeFilter: ['style'] });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    selectCell(cell) {
        this.clearSelection();
        this.currentCell = cell;
        cell.classList.add('selected');
        this.selectedCells = [cell];
    }

    selectColumn(header) {
        this.clearSelection();
        const col = header.dataset.col;
        
        // Special functionality for v2: clicking A, B, C headers transforms them and fills data
        if (col === 'A' || col === 'B' || col === 'C') {
            this.handleColumnHeaderClick(header);
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        const cells = document.querySelectorAll(`[data-col="${col}"]`);
        cells.forEach(cell => {
            cell.classList.add('selected');
            this.selectedCells.push(cell);
        });
        
        // Show column popover menu
        this.showColumnPopover(header);
    }

    handleColumnHeaderClick(header) {
        const col = header.dataset.col;
        
        // Transform header text and fill column with sample data
        if (col === 'A') {
            header.textContent = 'Email';
            this.fillColumnWithSampleData('A', 'Email');
        } else if (col === 'B') {
            header.textContent = 'Contact';
            this.fillColumnWithSampleData('B', 'Contact');
        } else if (col === 'C') {
            header.textContent = 'Company';
            this.fillColumnWithSampleData('C', 'Company');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    fillColumnWithSampleData(column, type) {
        // In Option 2 test mode, only fill rows 1-10; otherwise fill all 20 rows
        const maxRows = (this.currentVersion === 2 && this.isTestMode) ? 10 : 20;
        
        for (let row = 1; row <= maxRows; row++) {
            const cellKey = `${row}-${column}`;
            let sampleData = '';
            
            if (type === 'Email') {
                sampleData = `sample${row}@email.com`;
            } else if (type === 'Contact') {
                sampleData = `Sample User ${row}`;
            } else if (type === 'Company') {
                sampleData = `Sample Company ${row}`;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Update data store
            this.data[cellKey] = sampleData;
            
            // Update UI
            const cell = document.querySelector(`[data-row="${row}"][data-col="${column}"]`);
            if (cell) {
                const input = cell.querySelector('input');
                if (input) {
                    input.value = sampleData;
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Save state after filling data
        this.saveState();
    }

    selectRow(header) {
        this.clearSelection();
        const row = header.dataset.row;
        const cells = document.querySelectorAll(`[data-row="${row}"]`);
        cells.forEach(cell => {
            cell.classList.add('selected');
            this.selectedCells.push(cell);
        });
        

    }

    clearSelection() {
        if (this.currentCell) {
            this.currentCell.classList.remove('selected');
            this.currentCell = null;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        this.selectedCells.forEach(cell => {
            cell.classList.remove('selected');
        });
        this.selectedCells = [];
    }

    fillRows11To20() {
        // First, identify which columns have data in rows 1-10
        const columnsWithData = new Set();
        const appColumns = new Set();
        
        // Check each column to see if it has data in rows 1-10
        for (let col = 0; col < 8; col++) {
            const colLetter = String.fromCharCode(65 + col);
            let hasData = false;
            let isAppColumn = false;
            
            // Check if this column has any data in rows 1-10
            for (let row = 1; row <= 10; row++) {
                const cell = document.querySelector(`[data-row="${row}"][data-col="${colLetter}"]`);
                if (cell) {
                    const input = cell.querySelector('input');
                    if (input && input.value.trim() !== '') {
                        hasData = true;
                        break;
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    
                    // Check if it's an app column (has app-cell-content)
                    const appContent = cell.querySelector('.app-cell-content');
                    if (appContent) {
                        isAppColumn = true;
                        hasData = true;
                        break;
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            if (hasData) {
                columnsWithData.add(colLetter);
                if (isAppColumn) {
                    appColumns.add(colLetter);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Now fill rows 11-20 only for columns that have data
        for (let row = 11; row <= 20; row++) {
            columnsWithData.forEach(colLetter => {
                const cell = document.querySelector(`[data-row="${row}"][data-col="${colLetter}"]`);
                if (cell) {
                    if (appColumns.has(colLetter)) {
                        // This is an app column - create blue "Click to run" cell
                        cell.innerHTML = `
                            <div class="app-cell-content">
                                <span>Click to run</span>
                            </div>
                        `;
                        cell.classList.add('app-column-cell');
                    } else {
                        // This is a regular column - fill with sample data
                        const input = cell.querySelector('input');
                        if (input) {
                            // Get sample data from row 1 to determine the pattern
                            const sampleCell = document.querySelector(`[data-row="1"][data-col="${colLetter}"]`);
                            if (sampleCell) {
                                const sampleInput = sampleCell.querySelector('input');
                                if (sampleInput && sampleInput.value) {
                                    const sampleValue = sampleInput.value;
                                    
                                    // Generate similar data based on the pattern
                                    if (sampleValue.includes('@email.com')) {
                                        input.value = `sample${row}@email.com`;
                                    } else if (sampleValue.includes('Sample User')) {
                                        input.value = `Sample User ${row}`;
                                    } else if (sampleValue.includes('Sample Company')) {
                                        input.value = `Sample Company ${row}`;
                                    } else {
                                        // Generic pattern - replace the number
                                        input.value = sampleValue.replace(/\d+/, row.toString());
                                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Update credit system to show 1-20 rows: 300 credits
        const option2Text = document.getElementById('credit-text-option2');
        if (option2Text) {
            option2Text.innerHTML = '<strong>1-20 rows:</strong> 300 credits';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    searchCells(searchTerm) {
        const cells = document.querySelectorAll('.cell');
        cells.forEach(cell => {
            const input = cell.querySelector('input');
            const cellValue = input.value.toLowerCase();
            const cellKey = `${cell.dataset.row}-${cell.dataset.col}`;
            
            if (searchTerm === '' || cellValue.includes(searchTerm)) {
                cell.style.backgroundColor = '#ffffff';
            } else {
                cell.style.backgroundColor = '#fff3cd';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }

    addColumn() {
        this.columns++;
        const newCol = String.fromCharCode(65 + this.columns - 1);
        
        // Add column header
        const headerRow = document.querySelector('.header-row');
        const newHeader = document.createElement('th');
        newHeader.className = 'column-header';
        newHeader.dataset.col = newCol;
        newHeader.textContent = newCol;
        headerRow.appendChild(newHeader);

        // Add cells to each row
        const rows = document.querySelectorAll('.data-row');
        rows.forEach(row => {
            const newCell = document.createElement('td');
            newCell.className = 'cell';
            newCell.dataset.row = row.dataset.row;
            newCell.dataset.col = newCol;
            
            const input = document.createElement('input');
            input.type = 'text';
            input.placeholder = '';
            newCell.appendChild(input);
            
            row.appendChild(newCell);
        });
    }

    runAll() {
        console.log('Running all operations...');
        
        // Add output pills to all app columns
        this.appColumns.forEach(column => {
            this.addOutputPillsToColumn(column);
        });
        
        // For Option 1, only process first 10 rows
        if (this.currentVersion === 1) {
            console.log('Option 1: Processing first 10 rows only...');
        } else {
            console.log('Processing started...');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    addOutputPillsToColumn(column) {
        const cells = document.querySelectorAll(`[data-col="${column}"].app-column-cell`);
        let processedCount = 0;
        let totalToProcess = 0;
        
        // Determine max rows based on version, test mode, and selected row range
        let maxRows = 20;
        if (this.currentVersion === 1) {
            maxRows = 10; // Option 1: only process first 10 rows
        } else if (this.currentVersion === 2 && this.isTestMode) {
            maxRows = 10; // Option 2 test mode: only process first 10 rows
        } else if (this.currentVersion === 3 && this.selectedRowRange === '1-10') {
            maxRows = 10; // Option 3: process only first 10 rows when "1-10 rows" is selected
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Only count cells that need processing (empty cells without output pills)
        cells.forEach((cell, index) => {
            const row = parseInt(cell.dataset.row);
            if (row <= maxRows) {
                // Check if cell already has output pill
                const hasOutputPill = cell.querySelector('.output-pill');
                if (!hasOutputPill) {
                    totalToProcess++;
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
        
        cells.forEach((cell, index) => {
            const row = parseInt(cell.dataset.row);
            
            // Only process cells within row limit
            if (row <= maxRows) {
                // Check if cell already has output pill - skip if it does
                const hasOutputPill = cell.querySelector('.output-pill');
                if (hasOutputPill) {
                    console.log(`⏭️ Skipping row ${row} - already has output`);
                    return;
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Find the existing app-cell-content span and change its text
                const existingSpan = cell.querySelector('.app-cell-content span');
                if (existingSpan) {
                    // Change to "Generating..." without recreating elements
                    existingSpan.textContent = 'Generating...';
                    
                    // After 5 seconds, replace with output pill
                    setTimeout(() => {
                        // Replace the span content with output pill structure
                        existingSpan.outerHTML = `
                            <div class="output-pill">
                                <span class="output-text">Output</span>
                            </div>
                        `;
                        
                        processedCount++;
                        console.log(`🎯 Output generated for row ${row}! (${processedCount}/${totalToProcess})`);
                        
                        // Check if all app columns across all columns have outputs generated
                        if (processedCount === totalToProcess) {
                            if (this.currentVersion === 1) {
                                console.log('🎯 All first 10 rows completed! Updating credit text for Option 1...');
                                this.updateCreditTextAfterProcessing();
                            } else {
                                console.log('🎯 All processing completed!');
                                this.checkAllAppColumnsComplete();
                            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    }, 5000);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }





    updateCreditTextAfterProcessing() {
        const runAllText = document.querySelector('.run-all-text');
        const runAllIcon = document.querySelector('.run-all-info-icon');
        
        if (runAllText) {
            console.log('💰 Updating credit text to show remaining rows');
            runAllText.textContent = '11-20 rows: 300 credits';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        if (runAllIcon) {
            console.log('💰 Hiding info icon after processing first 10 rows');
            runAllIcon.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Show toast notification
        this.showCreditToast();
    }

    checkAllAppColumnsComplete() {
        // Check if all app column cells have outputs generated
        const allAppCells = document.querySelectorAll('.app-column-cell');
        let allComplete = true;
        let hasAppCells = false;
        
        allAppCells.forEach(cell => {
            hasAppCells = true;
            const outputPill = cell.querySelector('.output-pill');
            if (!outputPill) {
                allComplete = false;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
        
        // If all app cells have outputs, update credit system to indicate no empty rows
        if (hasAppCells && allComplete) {
            console.log('🎯 All app columns have outputs! Updating credit system...');
            
            // Update credit text based on version
            if (this.currentVersion === 2) {
                const option2Text = document.getElementById('credit-text-option2');
                if (option2Text) {
                    option2Text.innerHTML = '<strong>All rows processed</strong>';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (this.currentVersion === 3) {
                const option3Text = document.getElementById('credit-text-option3');
                if (option3Text) {
                    const creditPill = option3Text.querySelector('.credit-pill');
                    if (creditPill) {
                        creditPill.innerHTML = '<strong>All rows processed</strong> <i class="fas fa-chevron-down"></i>';
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Hide credit cost display when "All rows processed" is shown
            this.hideCreditDisplay();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    updateCreditTextForScaleMode() {
        const totalCredits = this.columnCount * 10;
        
        if (this.currentVersion === 3) {
            // Update Option 3 text with pill
            const option3Text = document.getElementById('credit-text-option3');
            if (option3Text) {
                // Show actual credits only when 1-10 rows is selected, otherwise show "- credits"
                const creditText = this.selectedRowRange === '1-10' ? `${totalCredits} credits` : '- credits';
                option3Text.innerHTML = `<span class="credit-pill" id="credit-pill"><strong>1-10 rows</strong> <i class="fas fa-chevron-down"></i></span> ${creditText}`;
                this.setupCreditPillDropdown();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (this.currentVersion === 2) {
            // Update Option 2 text without pill
            const option2Text = document.getElementById('credit-text-option2');
            if (option2Text) {
                option2Text.innerHTML = `<strong>1-10 rows:</strong> ${totalCredits} credits`;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (this.currentVersion === 1) {
            // Update Option 1 text
            const option1Text = document.getElementById('credit-text-option1');
            if (option1Text) {
                option1Text.innerHTML = `<strong>1-10 rows:</strong> ${totalCredits} credits`;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    setupOption3CreditPill() {
        const option3Text = document.getElementById('credit-text-option3');
        
        if (option3Text) {
            this.setupCreditPillDropdown();
            this.updateTestModeIndicator();
            this.updatePillState();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    updateTestModeIndicator() {
        const testModeIndicator = document.getElementById('test-mode-indicator-option3');
        if (!testModeIndicator) return;
        
        // Don't show test mode indicator for Option 3
        if (this.currentVersion === 3) {
            testModeIndicator.style.display = 'none';
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Only show test mode indicator when 1-10 rows is explicitly selected (not default state)
        if (this.selectedRowRange === '1-10') {
            testModeIndicator.style.display = 'inline-flex';
        } else {
            testModeIndicator.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

        updatePillState() {
        const creditPill = document.getElementById('credit-pill');
        if (!creditPill) return;
        
        // In Option 3, the pill should be enabled only when Run All button is enabled
        if (this.currentVersion === 3) {
            const runAllBtn = document.getElementById('run-all-btn');
            const isRunAllEnabled = runAllBtn && !runAllBtn.disabled;
            
            if (isRunAllEnabled) {
                creditPill.style.pointerEvents = 'auto';
                creditPill.style.opacity = '1';
                creditPill.style.cursor = 'pointer';
                
                // Auto-select "1-10 rows (Test)" when pill becomes active
                this.selectRowRange('1-10');
            } else {
                creditPill.style.pointerEvents = 'none';
                creditPill.style.opacity = '0.5';
                creditPill.style.cursor = 'not-allowed';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else {
            // For other options, only enable when app columns exist
            if (this.appColumns.size > 0) {
                creditPill.style.pointerEvents = 'auto';
                creditPill.style.opacity = '1';
                creditPill.style.cursor = 'pointer';
            } else {
                creditPill.style.pointerEvents = 'none';
                creditPill.style.opacity = '0.5';
                creditPill.style.cursor = 'not-allowed';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    setupCreditPillDropdown() {
        // Only setup dropdown for Option 3
        if (this.currentVersion !== 3) return;
        
        const creditPill = document.getElementById('credit-pill');
        if (!creditPill) return;

        // Remove existing event listeners
        creditPill.removeEventListener('click', this.handleCreditPillClick);
        
        // Only add click event listener if pill is enabled (Run All button is enabled)
        const runAllBtn = document.getElementById('run-all-btn');
        const isRunAllEnabled = runAllBtn && !runAllBtn.disabled;
        
        if (isRunAllEnabled) {
            creditPill.addEventListener('click', this.handleCreditPillClick.bind(this));
            // Ensure the pill text is correct when it becomes active
            if (this.selectedRowRange === '1-10' || !this.selectedRowRange) {
                creditPill.innerHTML = '<strong>1-10 rows (Test)</strong> <i class="fas fa-chevron-down"></i>';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    handleCreditPillClick(event) {
        event.stopPropagation();
        
        const creditPill = document.getElementById('credit-pill');
        const existingDropdown = document.querySelector('.credit-dropdown');
        
        // If dropdown is already open, close it
        if (existingDropdown && existingDropdown.classList.contains('show')) {
            this.closeCreditDropdown();
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Close any existing dropdown
        if (existingDropdown) {
            existingDropdown.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Create dropdown
        const dropdown = document.createElement('div');
        dropdown.className = 'credit-dropdown';
        dropdown.innerHTML = `
            <div class="credit-dropdown-item" data-rows="1-10">
                <div class="dropdown-main-text">1-10 rows (Test) ${this.selectedRowRange === '1-10' ? '<i class="fas fa-check"></i>' : ''}</div>
                <div class="dropdown-subtext">These rows allow you to test inputs without credit charges</div>
            </div>
            <div class="credit-dropdown-item" data-rows="all">All rows ${this.selectedRowRange === 'all' ? '<i class="fas fa-check"></i>' : ''}</div>
        `;
        
        // Position dropdown
        const rect = creditPill.getBoundingClientRect();
        dropdown.style.position = 'fixed';
        dropdown.style.top = `${rect.bottom + 4}px`;
        dropdown.style.left = `${rect.left}px`;
        
        // Add to body
        document.body.appendChild(dropdown);
        
        // Show dropdown
        setTimeout(() => {
            dropdown.classList.add('show');
            creditPill.classList.add('open');
        }, 10);
        
        // Add click handlers for dropdown items
        dropdown.querySelectorAll('.credit-dropdown-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                const rows = item.dataset.rows;
                this.selectRowRange(rows);
                this.closeCreditDropdown();
            });
        });
        
        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', this.closeCreditDropdown.bind(this), { once: true });
        }, 100);
    }

    closeCreditDropdown() {
        const dropdown = document.querySelector('.credit-dropdown');
        const creditPill = document.getElementById('credit-pill');
        
        if (dropdown) {
            dropdown.classList.remove('show');
            setTimeout(() => {
                if (dropdown.parentNode) {
                    dropdown.parentNode.removeChild(dropdown);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 200);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        if (creditPill) {
            creditPill.classList.remove('open');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    selectRowRange(rows) {
        const creditPill = document.getElementById('credit-pill');
        if (!creditPill) return;
        
        if (rows === 'all') {
            creditPill.innerHTML = '<strong>All rows</strong> <i class="fas fa-chevron-down"></i>';
        } else {
            creditPill.innerHTML = '<strong>1-10 rows (Test)</strong> <i class="fas fa-chevron-down"></i>';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Store the selection
        this.selectedRowRange = rows;
        localStorage.setItem('selectedRowRange', rows);
        
        // Update the test mode indicator outside the pill (will be ignored for Option 3)
        this.updateTestModeIndicator();
        
        // Update credit text based on selection
        this.updateCreditTextForScaleMode();
        
        // Update credit display for Option 3
        if (this.currentVersion === 3) {
            const creditText = document.getElementById('credit-text-option3');
            if (creditText) {
                if (rows === 'all') {
                    creditText.innerHTML = '<span class="credit-pill" id="credit-pill"><strong>All rows</strong> <i class="fas fa-chevron-down"></i></span> 500 credits';
                } else {
                    creditText.innerHTML = '<span class="credit-pill" id="credit-pill"><strong>1-10 rows (Test)</strong> <i class="fas fa-chevron-down"></i></span> 0 credits';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                // Re-setup the dropdown after updating the HTML
                this.setupCreditPillDropdown();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showCreditToast() {
        // Create toast element
        const toast = document.createElement('div');
        toast.className = 'credit-toast';
        toast.innerHTML = `
            <div class="toast-title">0 credits used for the first 10 cells</div>
            <div class="toast-body">Review outputs before running the rest. Credits apply after the first 10</div>
        `;
        
        // Add to body
        document.body.appendChild(toast);
        
        // Trigger animation
        setTimeout(() => {
            toast.classList.add('show');
        }, 100);
        
        // Auto hide after 10 seconds
        setTimeout(() => {
            toast.classList.remove('show');
            setTimeout(() => {
                if (toast.parentNode) {
                    toast.parentNode.removeChild(toast);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 300);
        }, 10000);
        
        console.log('🔔 Showing credit toast notification');
    }

    checkRunAllButton() {
        const runAllBtn = document.getElementById('run-all-btn');
        const runAllInfoOption1 = document.getElementById('run-all-info-option1');
        const runAllInfoOption3 = document.getElementById('run-all-info-option3');
        
        if (runAllBtn) {
            if (this.appColumns.size > 0) {
                runAllBtn.disabled = false;
                runAllBtn.style.opacity = '1';
                runAllBtn.style.cursor = 'pointer';
                // Show credit info when button is enabled
                if (this.currentVersion === 3) {
                    if (runAllInfoOption3) {
                        runAllInfoOption3.style.display = 'flex';
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    this.updatePillState();
                    this.setupCreditPillDropdown();
                } else if (this.currentVersion === 1) {
                    if (runAllInfoOption1) {
                        runAllInfoOption1.style.display = 'flex';
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                // Option 2 doesn't show credit info
            } else {
                runAllBtn.disabled = true;
                runAllBtn.style.opacity = '0.6';
                runAllBtn.style.cursor = 'not-allowed';
                // Hide credit info when button is disabled
                if (this.currentVersion === 3) {
                    if (runAllInfoOption3) {
                        runAllInfoOption3.style.display = 'none';
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                    this.updatePillState();
                    this.setupCreditPillDropdown();
                } else if (this.currentVersion === 1) {
                    if (runAllInfoOption1) {
                        runAllInfoOption1.style.display = 'none';
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                // Option 2 doesn't show credit info
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    startFreeRowsTest() {
        console.log('Starting free rows test on first 10 rows...');
        
        // Update UI to testing mode
        this.enterFreeRowsMode();
        
        console.log('Running free rows test on first 10 rows...');
        
        // Simulate test processing
        setTimeout(() => {
            this.completeFreeRowsTest();
        }, 1500);
    }

    enterFreeRowsMode() {
        // Update toggle text to show testing
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            toggleText.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Testing...';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Disable the checkbox during testing
        const freeRowsCheckbox = document.getElementById('free-rows-checkbox');
        if (freeRowsCheckbox) {
            freeRowsCheckbox.disabled = true;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Show test mode banner
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            testModeContainer.style.display = 'block';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    completeFreeRowsTest() {
        // Update toggle text to show completed
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            toggleText.innerHTML = '<i class="fas fa-check"></i> Test Complete';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Re-enable the checkbox
        const freeRowsCheckbox = document.getElementById('free-rows-checkbox');
        if (freeRowsCheckbox) {
            freeRowsCheckbox.disabled = false;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Show confirm output button
        const confirmOutputBtn = document.getElementById('confirm-output-btn');
        if (confirmOutputBtn) {
            confirmOutputBtn.style.display = 'inline-flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Update test mode message
        const testModeMessage = document.getElementById('test-mode-message');
        if (testModeMessage) {
            testModeMessage.innerHTML = '<strong>Free rows test completed!</strong> Review the output and confirm to proceed with importing all data.';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Enable Import All button
        this.enableImportAllButton();
        
        console.log('Free rows test completed! Review results and confirm output.');
    }

    resetFreeRowsTest() {
        // Reset toggle text
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            toggleText.textContent = '10 test rows';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Hide confirm output button
        const confirmOutputBtn = document.getElementById('confirm-output-btn');
        if (confirmOutputBtn) {
            confirmOutputBtn.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Hide test mode banner
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            testModeContainer.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Disable Import All button
        this.disableImportAllButton();
        
        console.log('Test rows reset.');
    }

    confirmOutput() {
        console.log('Output confirmed, enabling full import...');
        
        // Hide confirm output button
        const confirmOutputBtn = document.getElementById('confirm-output-btn');
        if (confirmOutputBtn) {
            confirmOutputBtn.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Reset toggle text
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            toggleText.textContent = '10 test rows';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Keep checkbox checked and disabled
        const freeRowsCheckbox = document.getElementById('free-rows-checkbox');
        if (freeRowsCheckbox) {
            freeRowsCheckbox.checked = true;
            freeRowsCheckbox.disabled = true;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Enable run all button
        const runAllBtn = document.getElementById('run-all-btn');
        if (runAllBtn) {
            runAllBtn.disabled = false;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Hide test mode banner
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            testModeContainer.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log('Output confirmed! You can now run all operations or import all data.');
    }

    enterTestingMode() {
        // Add testing mode class to body
        document.body.classList.add('testing-mode');
        
        // Show testing container
        const testingContainer = document.getElementById('testing-container');
        if (testingContainer) {
            testingContainer.style.display = 'block';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Update test run button
        const testRunBtn = document.querySelector('.test-run-btn');
        if (testRunBtn) {
            testRunBtn.classList.add('testing');
            testRunBtn.innerHTML = '<i class="fas fa-check"></i> Test Complete';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Disable run all button during testing
        const runAllBtn = document.querySelector('.run-all-btn');
        if (runAllBtn) {
            runAllBtn.disabled = true;
            runAllBtn.style.opacity = '0.5';
            runAllBtn.style.cursor = 'not-allowed';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    completeTestRun() {
        // Update testing message
        const testingMessage = document.getElementById('testing-message');
        if (testingMessage) {
            testingMessage.innerHTML = `
                <strong>Test completed successfully!</strong><br>
                Results from first 5 rows look good. Review the output and decide whether to proceed with all data.
            `;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Enable Import All button if in preview mode
        this.enableImportAllButton();
        
        console.log('Test run completed! Review results and approve to continue.');
    }

    approveTestRun() {
        console.log('Test approved, running full operation...');
        
        // Exit testing mode
        this.exitTestingMode();
        
        // Run the full operation
        this.runAll();
        
        console.log('Test approved! Running full operation on all data...');
    }

    rejectTestRun() {
        console.log('Test rejected, returning to edit mode...');
        
        // Exit testing mode
        this.exitTestingMode();
        
        console.log('Test rejected. You can make changes and test again.');
    }

    exitTestingMode() {
        // Remove testing mode class from body
        document.body.classList.remove('testing-mode');
        
        // Hide testing container
        const testingContainer = document.getElementById('testing-container');
        if (testingContainer) {
            testingContainer.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Reset test run button
        const testRunBtn = document.querySelector('.test-run-btn');
        if (testRunBtn) {
            testRunBtn.classList.remove('testing');
            testRunBtn.innerHTML = '<i class="fas fa-flask"></i> Test Run (First 5 rows)';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Re-enable run all button
        const runAllBtn = document.querySelector('.run-all-btn');
        if (runAllBtn) {
            runAllBtn.disabled = false;
            runAllBtn.style.opacity = '1';
            runAllBtn.style.cursor = 'pointer';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    loadSampleData() {
        // Add sample data to first 5 rows for testing
        const sampleData = {
            '1-A': 'Product A',
            '1-B': 'Marketing Campaign',
            '1-C': '1000',
            '1-D': 'High',
            '1-E': 'Active',
            '1-F': '2024-01-15',
            '1-G': 'John Smith',
            '1-H': 'Pending',
            
            '2-A': 'Product B',
            '2-B': 'Social Media',
            '2-C': '2500',
            '2-D': 'Medium',
            '2-E': 'Active',
            '2-F': '2024-01-20',
            '2-G': 'Jane Doe',
            '2-H': 'Approved',
            
            '3-A': 'Product C',
            '3-B': 'Email Campaign',
            '3-C': '800',
            '3-D': 'Low',
            '3-E': 'Inactive',
            '3-F': '2024-01-10',
            '3-G': 'Mike Johnson',
            '3-H': 'Rejected',
            
            '4-A': 'Product D',
            '4-B': 'Content Marketing',
            '4-C': '3000',
            '4-D': 'High',
            '4-E': 'Active',
            '4-F': '2024-01-25',
            '4-G': 'Sarah Wilson',
            '4-H': 'Pending',
            
            '5-A': 'Product E',
            '5-B': 'SEO Campaign',
            '5-C': '1500',
            '5-D': 'Medium',
            '5-E': 'Active',
            '5-F': '2024-01-18',
            '5-G': 'Tom Brown',
            '5-H': 'Approved'
        };

        // Load sample data
        Object.assign(this.data, sampleData);
        
        // Update the UI with sample data
        setTimeout(() => {
            this.updateUIWithData();
        }, 100);
    }

    updateUIWithData() {
        // Update all cells with current data
        console.log('🔄 Updating UI with data:', Object.keys(this.data).length, 'cells');
        Object.keys(this.data).forEach(cellKey => {
            const [row, col] = cellKey.split('-');
            const cell = document.querySelector(`[data-row="${row}"][data-col="${col}"]`);
            if (cell) {
                const input = cell.querySelector('input');
                if (input) {
                    input.value = this.data[cellKey];
                    console.log(`✅ Updated cell ${cellKey} with value: "${this.data[cellKey]}"`);
                } else {
                    console.log(`❌ No input found in cell ${cellKey}`);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else {
                console.log(`❌ Cell ${cellKey} not found in DOM`);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }

    checkUploadedFile() {
        const uploadedFile = localStorage.getItem('uploadedFile');
        const savedState = localStorage.getItem('spreadsheetState');
        
        console.log('🔍 Checking for uploaded file...', { uploadedFile: !!uploadedFile, savedState: !!savedState });
        
        if (uploadedFile) {
            try {
                const fileInfo = JSON.parse(uploadedFile);
                console.log('📁 Found uploaded file info:', fileInfo);
                
                // Update the file name in the navigation
                const fileNameElement = document.getElementById('file-name');
                if (fileNameElement && fileInfo.name) {
                    fileNameElement.textContent = fileInfo.name.replace(/\.[^/.]+$/, ''); // Remove file extension
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Check if this is preview mode with real data
                if (fileInfo.previewMode && fileInfo.data) {
                    console.log('📊 Loading preview mode with real data');
                    this.showPreviewMode(fileInfo);
                } else {
                    // Load sample data without showing notification
                    console.log('📝 Loading sample data');
                    this.loadSampleData();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Clear the uploaded file info from localStorage
                localStorage.removeItem('uploadedFile');
                
            } catch (error) {
                console.error('Error parsing uploaded file info:', error);
                this.loadSampleData();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (savedState) {
            // Restore saved state
            console.log('🔄 Restoring saved state');
            this.restoreState(savedState);
        } else {
            // No saved state, load sample data
            console.log('🆕 No data found, loading sample data');
            this.loadSampleData();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Import All button should remain disabled until output is generated
        console.log('💡 Import All button will be enabled after running output generation');
    }

    saveState() {
        const state = {
            data: this.data,
            rows: this.rows,
            columns: this.columns,
            fileName: document.getElementById('file-name')?.textContent || 'Untitled sheet',
            previewMode: document.getElementById('test-mode-container')?.style.display !== 'none',
            timestamp: new Date().toISOString()
        };
        
        localStorage.setItem('spreadsheetState', JSON.stringify(state));
    }

    restoreState(savedState) {
        try {
            const state = JSON.parse(savedState);
            
            // Restore basic properties
            this.data = state.data || {};
            this.rows = state.rows || 18;
            this.columns = state.columns || 8;
            
            // Update file name
            const fileNameElement = document.getElementById('file-name');
            if (fileNameElement && state.fileName) {
                fileNameElement.textContent = state.fileName;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Regenerate spreadsheet with saved data
            this.generateRows();
            this.updateUIWithData();
            
            // Restore preview mode if it was active
            if (state.previewMode) {
                const testModeContainer = document.getElementById('test-mode-container');
                if (testModeContainer) {
                    testModeContainer.style.display = 'block';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            console.log('State restored successfully');
            
        } catch (error) {
            console.error('Error restoring state:', error);
            this.loadSampleData();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showPreviewMode(fileInfo) {
        // Show test mode container
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            testModeContainer.style.display = 'block';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Load the real data from the Excel file (only first 10 rows for preview)
        this.loadRealData(fileInfo.data);
        
        // Store the full original data for later import
        this.originalFullData = fileInfo.data;
        
        // Keep Import All button disabled until output is generated
        console.log('📊 Preview mode: keeping Import All button disabled until output is generated');
        this.disableImportAllButton();
        
        // Save state after loading preview mode
        this.saveState();
    }

    enableImportAllButton() {
        const importAllBtn = document.getElementById('import-all-btn');
        if (importAllBtn) {
            console.log('🔓 Enabling Import All button');
            importAllBtn.disabled = false;
            importAllBtn.style.opacity = '1';
            importAllBtn.style.cursor = 'pointer';
            console.log('✅ Import All button enabled successfully');
        } else {
            console.log('❌ Import All button not found when trying to enable');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    disableImportAllButton() {
        const importAllBtn = document.getElementById('import-all-btn');
        if (importAllBtn) {
            importAllBtn.disabled = true;
            importAllBtn.style.opacity = '0.6';
            importAllBtn.style.cursor = 'not-allowed';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    importAllData() {
        console.log('🚀 importAllData() called - Starting import all data...');
        
        // Show loading state in the 11th row
        console.log('📋 Showing loading state...');
        this.showImportLoadingState();
        
        // Simulate import process with loading
        setTimeout(() => {
            console.log('⏰ Timeout completed, generating bulk data...');
            
            // Remove the "1-10 rows" copy from toggle
            this.removeRowCountCopy();
            
            // Generate and display bulk data
            this.generateBulkData();
            
            // Hide the test mode container
            this.hideTestModeRow();
            
            // Update Run All button with cost information
            this.updateRunAllButtonWithCost();
            
            console.log('✅ All data imported successfully!');
        }, 2000);
    }

    showImportLoadingState() {
        const testModeContainer = document.getElementById('test-mode-container');
        if (testModeContainer) {
            testModeContainer.innerHTML = `
                <div class="loading-state">
                    <div class="loading-text">
                        <div class="loading-title">Importing Data...</div>
                        <div class="loading-message">Please wait while we import all your data</div>
                    </div>
                    <div class="progress-bar-container">
                        <div class="progress-bar">
                            <div class="progress-bar-fill"></div>
                        </div>
                    </div>
                </div>
            `;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    hideTestModeRow() {
        const testModeRow = document.getElementById('test-mode-row');
        if (testModeRow) {
            testModeRow.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    removeRowCountCopy() {
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            console.log('🗑️ Removing "1-10 rows" copy from toggle');
            toggleText.textContent = ''; // Remove the text
            toggleText.style.display = 'none'; // Hide the element
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showCostDisplay() {
        if (this.currentVersion === 3) {
            const runAllInfoOption3 = document.getElementById('run-all-info-option3');
            if (runAllInfoOption3 && runAllInfoOption3.style.display === 'none') {
                console.log('💰 Showing cost display for Option 3');
                runAllInfoOption3.style.display = 'flex';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else {
            const runAllInfo = document.getElementById('run-all-info-option12');
            if (runAllInfo && runAllInfo.style.display === 'none') {
                console.log('💰 Showing cost display for Options 1/2');
                runAllInfo.style.display = 'flex';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    updateRunAllButtonWithCost() {
        const toggleText = document.querySelector('.toggle-text');
        if (toggleText) {
            console.log('💰 Updating toggle area with cost information');
            
            // Show the toggle text element and update with cost
            toggleText.style.display = 'flex';
            
            // Calculate the row range (11 to total rows after import)
            // First 10 rows were already processed in preview mode
            const totalRows = this.rows;
            toggleText.textContent = `11-${totalRows} rows: 800 credits`;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    removeColumnsEFGH() {
        console.log('🗑️ Removing columns E, F, G, H');
        
        // Remove column headers E, F, G, H
        const columnsToRemove = ['E', 'F', 'G', 'H'];
        columnsToRemove.forEach(col => {
            const header = document.querySelector(`[data-col="${col}"]`);
            if (header) {
                header.remove();
                console.log(`❌ Removed header for column ${col}`);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
        
        // Remove cells from all existing rows (including the first 10 rows)
        const allRows = document.querySelectorAll('.data-row');
        allRows.forEach(row => {
            columnsToRemove.forEach(col => {
                const cell = row.querySelector(`[data-col="${col}"]`);
                if (cell) {
                    cell.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            });
        });
        
        // Update the columns count to 4
        this.columns = 4;
        
        // Remove columns E, F, G, H from appColumns set if they exist
        columnsToRemove.forEach(col => {
            if (this.appColumns.has(col)) {
                this.appColumns.delete(col);
                console.log(`📱 Removed column ${col} from appColumns`);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
        
        // Clean up data for removed columns
        Object.keys(this.data).forEach(cellKey => {
            const [row, col] = cellKey.split('-');
            if (columnsToRemove.includes(col)) {
                delete this.data[cellKey];
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
        
        console.log('✅ Columns E, F, G, H removed successfully');
    }

    isColumnConfiguredWithApp(columnLetter) {
        // Check if this column has an app configured by looking at existing cells
        const existingAppCell = document.querySelector(`[data-col="${columnLetter}"].app-column-cell`);
        if (existingAppCell) {
            console.log(`📱 Column ${columnLetter} has app configured`);
            return true;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Also check if column is in the appColumns set
        if (this.appColumns.has(columnLetter)) {
            console.log(`📱 Column ${columnLetter} found in appColumns set`);
            return true;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        return false;
    }

    createAppCellForImportedRow(td, columnLetter, row) {
        // Get the app name from existing cells in this column
        const existingAppCell = document.querySelector(`[data-col="${columnLetter}"].app-column-cell`);
        let appName = 'App'; // Default fallback
        
        if (existingAppCell) {
            const headerElement = document.querySelector(`[data-col="${columnLetter}"]`);
            if (headerElement && headerElement.dataset.customName) {
                appName = headerElement.dataset.customName;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log(`🔧 Creating app cell for row ${row}, column ${columnLetter}, app: ${appName}`);
        
        // Create app display with same structure as existing app cells
        const appDisplay = document.createElement('div');
        appDisplay.className = 'app-cell-display';
        appDisplay.innerHTML = `
            <div class="app-cell-content">
                <span>Click to run</span>
            </div>
        `;
        
        // Add click handler
        appDisplay.addEventListener('click', () => {
            this.runAppInCell(td, appName);
        });
        
        td.appendChild(appDisplay);
        td.classList.add('app-column-cell');
        
        console.log(`✅ App cell created for ${columnLetter}${row}`);
    }

    generateBulkData() {
        // Remove the test mode row completely
        this.hideTestModeRow();
        
        // Get the real uploaded data from localStorage
        const uploadedFileData = this.getRealUploadedData();
        
        if (uploadedFileData && uploadedFileData.length > 0) {
            // Use the real uploaded data
            this.updateSpreadsheetForBulkData(uploadedFileData);
        } else {
            // Fallback to sample data if no real data is available
            console.log('No uploaded data found, using sample data as fallback');
            const bulkData = this.createBulkSampleData();
            this.updateSpreadsheetForBulkData(bulkData);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Add scrollbar to the spreadsheet container
        this.enableSpreadsheetScrolling();
    }

    getRealUploadedData() {
        try {
            // First check if we have the original full data stored in memory
            if (this.originalFullData && Array.isArray(this.originalFullData)) {
                console.log(`📊 Using stored original data with ${this.originalFullData.length} rows (importing remaining rows)`);
                // Skip the header row (index 0) and the first 10 data rows (indices 1-10)
                // So we want rows starting from index 11 (which is the 11th data row, 12th total row)
                const remainingData = this.originalFullData.slice(11);
                console.log(`📈 Importing ${remainingData.length} remaining rows (skipping header + first 10 data rows)`);
                
                // Create a temporary array with header + remaining data for proper conversion
                const dataWithHeader = [this.originalFullData[0], ...remainingData];
                return this.convertUploadedDataToSpreadsheetFormat(dataWithHeader);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Check for uploaded file data in localStorage
            const uploadedFile = localStorage.getItem('uploadedFile');
            if (uploadedFile) {
                const fileInfo = JSON.parse(uploadedFile);
                console.log('Found uploaded file info:', fileInfo);
                
                // Return the remaining data (beyond header + first 10 data rows)
                if (fileInfo.data && Array.isArray(fileInfo.data)) {
                    const remainingData = fileInfo.data.slice(11); // Skip header + first 10 data rows
                    console.log(`Using remaining uploaded data: ${remainingData.length} rows (total was ${fileInfo.data.length})`);
                    
                    // Create array with header + remaining data for proper conversion
                    const dataWithHeader = [fileInfo.data[0], ...remainingData];
                    return this.convertUploadedDataToSpreadsheetFormat(dataWithHeader);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Also check for any other data storage formats
            const spreadsheetState = localStorage.getItem('spreadsheetState');
            if (spreadsheetState) {
                const state = JSON.parse(spreadsheetState);
                if (state.originalUploadedData) {
                    console.log('Found original uploaded data in state');
                    const remainingData = state.originalUploadedData.slice(11); // Skip header + first 10 data rows
                    
                    // Create array with header + remaining data for proper conversion
                    const dataWithHeader = [state.originalUploadedData[0], ...remainingData];
                    return this.convertUploadedDataToSpreadsheetFormat(dataWithHeader);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            return null;
        } catch (error) {
            console.error('Error retrieving uploaded data:', error);
            return null;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    convertUploadedDataToSpreadsheetFormat(rawData) {
        if (!rawData || rawData.length === 0) return [];
        
        // Skip the first row (headers) and convert the rest
        const dataRows = rawData.slice(1);
        
        console.log('🔄 Converting uploaded data format, skipping header row');
        console.log('📊 Converting', dataRows.length, 'data rows');
        
        // Convert the raw uploaded data to the format expected by the spreadsheet
        return dataRows.map((row, index) => {
            // Ensure we have at least 8 columns of data
            const formattedRow = {};
            const columnKeys = ['email', 'contact', 'company', 'phone', 'status', 'priority', 'department', 'notes'];
            
            // If rawData is an array of arrays (CSV-like format)
            if (Array.isArray(row)) {
                columnKeys.forEach((key, colIndex) => {
                    formattedRow[key] = row[colIndex] || '';
                });
            } 
            // If rawData is an array of objects
            else if (typeof row === 'object' && row !== null) {
                // Try to map object properties to our expected format
                formattedRow.email = row.email || row.Email || row['Email Address'] || row[Object.keys(row)[0]] || '';
                formattedRow.contact = row.contact || row.Contact || row.Name || row.name || row[Object.keys(row)[1]] || '';
                formattedRow.company = row.company || row.Company || row.Organization || row.organization || row[Object.keys(row)[2]] || '';
                formattedRow.phone = row.phone || row.Phone || row['Phone Number'] || row.phoneNumber || row[Object.keys(row)[3]] || '';
                formattedRow.status = row.status || row.Status || row[Object.keys(row)[4]] || '';
                formattedRow.priority = row.priority || row.Priority || row[Object.keys(row)[5]] || '';
                formattedRow.department = row.department || row.Department || row[Object.keys(row)[6]] || '';
                formattedRow.notes = row.notes || row.Notes || row.Comments || row.comments || row[Object.keys(row)[7]] || '';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            return formattedRow;
        });
    }

    createBulkSampleData() {
        const companies = ['TechCorp', 'DataSolutions', 'CloudWorks', 'InnovateTech', 'DigitalFlow', 'SmartSystems', 'NextGen Labs', 'ProTech Inc', 'CyberSoft', 'WebDynamics'];
        const domains = ['gmail.com', 'company.com', 'business.org', 'tech.io', 'startup.co'];
        const firstNames = ['John', 'Jane', 'Michael', 'Sarah', 'David', 'Emily', 'Robert', 'Lisa', 'James', 'Maria'];
        const lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones', 'Garcia', 'Miller', 'Davis', 'Rodriguez', 'Martinez'];
        
        const bulkData = [];
        
        // Generate 100 rows of data
        for (let i = 1; i <= 100; i++) {
            const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
            const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
            const company = companies[Math.floor(Math.random() * companies.length)];
            const domain = domains[Math.floor(Math.random() * domains.length)];
            const email = `${firstName.toLowerCase()}.${lastName.toLowerCase()}@${domain}`;
            
            bulkData.push({
                email: email,
                contact: `${firstName} ${lastName}`,
                company: company,
                phone: `+1-${Math.floor(Math.random() * 900 + 100)}-${Math.floor(Math.random() * 900 + 100)}-${Math.floor(Math.random() * 9000 + 1000)}`,
                status: ['Active', 'Pending', 'Inactive'][Math.floor(Math.random() * 3)],
                priority: ['High', 'Medium', 'Low'][Math.floor(Math.random() * 3)],
                department: ['Sales', 'Marketing', 'Support', 'Engineering'][Math.floor(Math.random() * 4)],
                notes: `Contact notes for ${firstName} ${lastName}`
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        return bulkData;
    }

    updateSpreadsheetForBulkData(bulkData) {
        const tbody = document.getElementById('spreadsheet-body');
        
        // Store original uploaded data for future reference
        this.originalUploadedData = bulkData;
        
        // Append the remaining data to existing rows (starting from row 11)
        const startingRow = this.rows + 1; // Continue from where we left off
        
        // Generate rows with bulk data
        bulkData.forEach((rowData, rowIndex) => {
            const row = startingRow + rowIndex;
            const tr = document.createElement('tr');
            tr.className = 'data-row';
            tr.dataset.row = row;

            // Row header
            const th = document.createElement('th');
            th.className = 'row-header';
            th.textContent = row;
            tr.appendChild(th);

            // Data cells - generate all 8 columns, but E-H will be empty
            const cellValues = [
                rowData.email || '',
                rowData.contact || '',
                rowData.company || '',
                rowData.phone || '',
                '', // Column E - empty
                '', // Column F - empty
                '', // Column G - empty
                ''  // Column H - empty
            ];

            for (let col = 0; col < 8; col++) {
                const td = document.createElement('td');
                td.className = 'cell';
                td.dataset.row = row;
                td.dataset.col = String.fromCharCode(65 + col); // A, B, C, etc.
                
                // Check if this column has an app configured
                const columnLetter = String.fromCharCode(65 + col);
                const isAppColumn = this.isColumnConfiguredWithApp(columnLetter);
                
                if (isAppColumn) {
                    // Create app cell for imported rows
                    this.createAppCellForImportedRow(td, columnLetter, row);
                } else {
                    // Create regular input cell
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.placeholder = '';
                    input.value = cellValues[col];
                    
                    // Store in data
                    const cellKey = `${row}-${columnLetter}`;
                    this.data[cellKey] = input.value;
                    
                    td.appendChild(input);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                tr.appendChild(td);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

            tbody.appendChild(tr);
        });
        
        // Update total rows count
        this.rows = startingRow + bulkData.length - 1;
        
        console.log(`📊 Added ${bulkData.length} rows. Total rows now: ${this.rows}`);
        
        // Keep all 8 columns, no header update needed
        
        // Save state with the updated data
        this.saveStateWithUploadedData();
    }

    updateHeaderRowForBulkData(bulkData = null) {
        const headerRow = document.querySelector('.header-row');
        if (!headerRow) return;

        // Keep the corner cell
        const cornerCell = headerRow.querySelector('.corner-cell');
        headerRow.innerHTML = '';
        headerRow.appendChild(cornerCell);

        // Try to determine column names from the actual data structure
        let columnNames = ['Email', 'Contact', 'Company', 'Phone', 'Status', 'Priority', 'Department', 'Notes'];
        
        if (bulkData && bulkData.length > 0) {
            const firstRow = bulkData[0];
            if (typeof firstRow === 'object' && firstRow !== null) {
                // Use the actual keys from the uploaded data for more accurate headers
                const dataKeys = Object.keys(firstRow);
                if (dataKeys.length > 0) {
                    columnNames = dataKeys.slice(0, 8).map(key => {
                        // Capitalize and clean up the key names
                        return key.charAt(0).toUpperCase() + key.slice(1).replace(/([A-Z])/g, ' $1').trim();
                    });
                    
                    // Pad with default names if we don't have 8 columns
                    while (columnNames.length < 8) {
                        columnNames.push(`Column ${columnNames.length + 1}`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        for (let col = 0; col < 8; col++) {
            const th = document.createElement('th');
            th.className = 'column-header';
            th.dataset.col = String.fromCharCode(65 + col);
            th.textContent = columnNames[col];
            headerRow.appendChild(th);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    updateHeaderRowForReducedColumns() {
        const headerRow = document.querySelector('.header-row');
        if (!headerRow) return;

        console.log('📋 Updating header row for 4-column layout');

        // Keep the corner cell
        const cornerCell = headerRow.querySelector('.corner-cell');
        headerRow.innerHTML = '';
        headerRow.appendChild(cornerCell);

        // Only create headers for columns A, B, C, D
        const columnNames = ['Email', 'Contact', 'Company', 'Phone'];
        
        for (let col = 0; col < 4; col++) {
            const th = document.createElement('th');
            th.className = 'column-header';
            th.dataset.col = String.fromCharCode(65 + col); // A, B, C, D
            th.textContent = columnNames[col];
            headerRow.appendChild(th);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log('✅ Header row updated for 4 columns');
    }

    updateHeaderRowWithUploadedHeaders() {
        const headerRow = document.querySelector('.header-row');
        if (!headerRow) return;

        console.log('📋 Updating headers with uploaded file headers');

        // Keep the corner cell
        const cornerCell = headerRow.querySelector('.corner-cell');
        headerRow.innerHTML = '';
        headerRow.appendChild(cornerCell);

        // Use uploaded headers or fallback to default names
        const columnNames = this.uploadedHeaders || ['Email', 'Contact', 'Company', 'Phone', 'Status', 'Priority', 'Department', 'Notes'];
        
        for (let col = 0; col < 8; col++) {
            const th = document.createElement('th');
            th.className = 'column-header';
            th.dataset.col = String.fromCharCode(65 + col);
            
            // Use uploaded header name or fallback
            const headerName = columnNames[col] || `Column ${String.fromCharCode(65 + col)}`;
            th.textContent = headerName;
            
            headerRow.appendChild(th);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log('✅ Headers updated with uploaded data:', columnNames);
    }

    saveStateWithUploadedData() {
        const state = {
            data: this.data,
            rows: this.rows,
            columns: this.columns,
            fileName: document.getElementById('file-name')?.textContent || 'Imported data',
            originalUploadedData: this.originalUploadedData,
            timestamp: new Date().toISOString()
        };
        
        localStorage.setItem('spreadsheetState', JSON.stringify(state));
        console.log('State saved with uploaded data');
    }

    enableSpreadsheetScrolling() {
        const spreadsheetContainer = document.querySelector('.spreadsheet-container');
        if (spreadsheetContainer) {
            // Ensure the container has proper scrolling styles
            spreadsheetContainer.style.maxHeight = 'calc(100vh - 48px - 60px - 40px)'; // Account for header, search bar, and some padding
            spreadsheetContainer.style.overflowY = 'auto';
            spreadsheetContainer.style.overflowX = 'auto';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Also ensure the spreadsheet wrapper can handle the overflow
        const spreadsheetWrapper = document.querySelector('.spreadsheet-wrapper');
        if (spreadsheetWrapper) {
            spreadsheetWrapper.style.overflow = 'visible';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    loadRealData(data) {
        // Clear existing data
        this.data = {};
        
        // Always use 8 columns for display
        this.columns = 8;
        
        if (data.length === 0) return;
        
        // First row becomes headers, rest becomes data
        const headerRow = data[0];
        const dataRows = data.slice(1); // Skip first row (headers)
        
        console.log('📋 Using first row as headers:', headerRow);
        console.log('📊 Total data rows available:', dataRows.length);
        console.log('📊 First 10 data rows:', dataRows.slice(0, 10).map((row, i) => `Row ${i+1}: [${row.slice(0, 3).join(', ')}...]`));
        
        // Store headers for later use
        this.uploadedHeaders = headerRow.slice(0, 8); // Take first 8 columns as headers
        
        // Determine number of rows from the data (max 10 for preview, excluding header row)
        // We want to show 10 data rows, so we need at least 10 data rows available
        const numRows = Math.min(dataRows.length, 10);
        this.rows = 10; // Always show 10 rows in preview mode
        
        // Regenerate the spreadsheet with 8 columns
        this.generateRows();
        
        // Load the data into the spreadsheet (only first 8 columns, starting from data row 2)
        // Fill up to 10 rows with available data
        console.log(`📊 Loading data for ${Math.min(dataRows.length, 10)} rows from uploaded file`);
        console.log(`📊 Available data rows:`, dataRows.length);
        
        for (let rowIndex = 0; rowIndex < 10; rowIndex++) {
            const row = dataRows[rowIndex];
            if (row && row.length > 0) {
                console.log(`📝 Loading row ${rowIndex + 1} with data:`, row.slice(0, 8));
                row.forEach((cellValue, colIndex) => {
                    if (colIndex < 8) { // Limit to 8 columns
                        const cellKey = `${rowIndex + 1}-${String.fromCharCode(65 + colIndex)}`;
                        this.data[cellKey] = cellValue ? String(cellValue) : '';
                        console.log(`   Cell ${cellKey} = "${this.data[cellKey]}"`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                });
            } else {
                console.log(`❌ No data available for row ${rowIndex + 1} - filling with sample data`);
                // Fill with sample data if no real data is available
                const sampleData = [
                    `sample${rowIndex + 1}@email.com`,
                    `Sample User ${rowIndex + 1}`,
                    `Sample Company ${rowIndex + 1}`
                ];
                
                for (let colIndex = 0; colIndex < 3; colIndex++) {
                    const cellKey = `${rowIndex + 1}-${String.fromCharCode(65 + colIndex)}`;
                    this.data[cellKey] = sampleData[colIndex];
                    console.log(`   Sample Cell ${cellKey} = "${this.data[cellKey]}"`);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Update headers with uploaded data
        this.updateHeaderRowWithUploadedHeaders();
        
        // Update the UI with the loaded data
        this.updateUIWithData();
        
        // Force update row 10 specifically to ensure it's filled
        console.log('🔧 Force checking row 10...');
        const row10Data = ['10-A', '10-B', '10-C'].map(key => this.data[key]).filter(val => val);
        console.log('🔧 Row 10 data in memory:', row10Data);
        
        if (row10Data.length === 0) {
            console.log('🔧 Row 10 is empty, forcing sample data...');
            // Force fill row 10 with sample data
            const sampleData = ['sample10@email.com', 'Sample User 10', 'Sample Company 10'];
            for (let colIndex = 0; colIndex < 3; colIndex++) {
                const cellKey = `10-${String.fromCharCode(65 + colIndex)}`;
                this.data[cellKey] = sampleData[colIndex];
                
                // Directly update the UI element
                const cell = document.querySelector(`[data-row="10"][data-col="${String.fromCharCode(65 + colIndex)}"]`);
                if (cell) {
                    const input = cell.querySelector('input');
                    if (input) {
                        input.value = sampleData[colIndex];
                        console.log(`🔧 Force updated cell ${cellKey} with: "${sampleData[colIndex]}"`);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    handleFloatingAction(index) {
        const actions = ['comment', 'package', 'search', 'add'];
        const action = actions[index];
        
        switch (action) {
            case 'comment':
                console.log('Comment feature clicked');
                break;
            case 'package':
                console.log('Package feature clicked');
                break;
            case 'search':
                document.querySelector('.search-input').focus();
                break;
            case 'add':
                console.log('Add feature clicked');
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showColumnPopover(header) {
        // Remove existing popovers
        const existingPopover = document.querySelector('.column-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create column popover
        const popover = document.createElement('div');
        popover.className = 'column-popover';
        popover.innerHTML = `
            <div class="popover-section">
                <div class="section-title">Input</div>
                <div class="popover-item" data-action="text">
                    <span>Text</span>
                </div>
                <div class="popover-item" data-action="dropdown">
                    <span>Dropdown</span>
                </div>
                <div class="popover-item" data-action="file-upload">
                    <span>File upload</span>
                </div>
                <div class="popover-item has-submenu" data-action="jasper-iq">
                    <span>Jasper IQ</span>
                    <i class="fas fa-chevron-right"></i>
                </div>
                <div class="popover-item" data-action="import-data">
                    <span>Import data</span>
                </div>
            </div>
            <div class="popover-divider"></div>
            <div class="popover-section">
                <div class="section-title">Output</div>
                <div class="popover-item" data-action="agent">
                    <span>∞ Agent</span>
                </div>
                <div class="popover-item" data-action="app">
                    <i class="fas fa-cube"></i>
                    <span>App</span>
                </div>
                <div class="popover-item" data-action="prompt">
                    <i class="fas fa-desktop"></i>
                    <span>Prompt</span>
                </div>
            </div>
        `;

        // Position the popover below the column header
        const headerRect = header.getBoundingClientRect();
        popover.style.cssText = `
            position: fixed;
            top: ${headerRect.bottom + 5}px;
            left: ${headerRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            min-width: 200px;
            padding: 16px 0;
        `;

        // Add event listeners to popover items
        popover.addEventListener('click', (e) => {
            const item = e.target.closest('.popover-item');
            if (item) {
                const action = item.dataset.action;
                if (action !== 'jasper-iq') {
                    this.handleColumnPopoverAction(action, header.dataset.col);
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Add hover events for Jasper IQ submenu
        const jasperItem = popover.querySelector('[data-action="jasper-iq"]');
        if (jasperItem) {
            jasperItem.addEventListener('mouseenter', () => {
                this.showJasperIQSubmenu(jasperItem, header.dataset.col);
            });
            
            jasperItem.addEventListener('mouseleave', () => {
                // Hide submenu when mouse leaves Jasper IQ item
                setTimeout(() => {
                    const submenu = document.querySelector('.jasper-iq-submenu');
                    if (submenu && !submenu.matches(':hover')) {
                        submenu.remove();
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                }, 100);
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        document.body.appendChild(popover);

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', () => {
                popover.remove();
            }, { once: true });
        }, 100);
    }

    showJasperIQSubmenu(jasperItem, column) {
        // Remove existing submenus
        const existingSubmenu = document.querySelector('.jasper-iq-submenu');
        if (existingSubmenu) {
            existingSubmenu.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create Jasper IQ submenu
        const submenu = document.createElement('div');
        submenu.className = 'jasper-iq-submenu';
        submenu.innerHTML = `
            <div class="submenu-header">
                <h4>Jasper IQ</h4>
            </div>
            <div class="submenu-item" data-action="brand-voice">Brand Voice</div>
            <div class="submenu-item" data-action="audience">Audience</div>
            <div class="submenu-item" data-action="knowledge-base">Knowledge Base</div>
            <div class="submenu-item" data-action="style-guide">Style Guide</div>
            <div class="submenu-item" data-action="visual-guidelines">Visual Guidelines</div>
        `;

        // Position the submenu to the right of the main popover
        const jasperRect = jasperItem.getBoundingClientRect();
        const mainPopover = jasperItem.closest('.column-popover');
        const popoverRect = mainPopover.getBoundingClientRect();

        submenu.style.cssText = `
            position: fixed;
            top: ${jasperRect.top}px;
            left: ${popoverRect.right + 5}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1001;
            min-width: 180px;
            padding: 12px 0;
        `;

        // Add event listeners to submenu items
        submenu.addEventListener('click', (e) => {
            const item = e.target.closest('.submenu-item');
            if (item) {
                const action = item.dataset.action;
                this.handleJasperIQAction(action, column);
                submenu.remove();
                // Also close the main popover
                const mainPopover = document.querySelector('.column-popover');
                if (mainPopover) {
                    mainPopover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Add hover events to keep submenu open
        submenu.addEventListener('mouseenter', () => {
            // Keep submenu open when hovering over it
        });

        submenu.addEventListener('mouseleave', () => {
            // Close submenu when mouse leaves
            setTimeout(() => {
                if (submenu.parentNode) {
                    submenu.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 100);
        });

        document.body.appendChild(submenu);
    }

    showContextMenu(e) {
        // Remove existing context menu
        const existingMenu = document.querySelector('.context-menu');
        if (existingMenu) {
            existingMenu.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create context menu
        const menu = document.createElement('div');
        menu.className = 'context-menu';
        menu.style.cssText = `
            position: fixed;
            top: ${e.clientY}px;
            left: ${e.clientX}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            min-width: 150px;
        `;

        const menuItems = [
            { text: 'Copy', icon: 'fas fa-copy' },
            { text: 'Paste', icon: 'fas fa-paste' },
            { text: 'Cut', icon: 'fas fa-cut' },
            { text: 'Delete', icon: 'fas fa-trash' },
            { text: 'Format', icon: 'fas fa-paint-brush' }
        ];

        menuItems.forEach(item => {
            const menuItem = document.createElement('div');
            menuItem.style.cssText = `
                padding: 8px 16px;
                cursor: pointer;
                display: flex;
                align-items: center;
                gap: 8px;
                transition: background-color 0.2s;
            `;
            menuItem.innerHTML = `<i class="${item.icon}"></i>${item.text}`;
            
            menuItem.addEventListener('mouseenter', () => {
                menuItem.style.backgroundColor = '#f5f5f5';
            });
            
            menuItem.addEventListener('mouseleave', () => {
                menuItem.style.backgroundColor = 'transparent';
            });
            
            menuItem.addEventListener('click', () => {
                this.handleContextMenuAction(item.text.toLowerCase());
                menu.remove();
            });
            
            menu.appendChild(menuItem);
        });

        document.body.appendChild(menu);

        // Close menu when clicking outside
        setTimeout(() => {
            document.addEventListener('click', () => {
                menu.remove();
            }, { once: true });
        }, 100);
    }

    showTextColumnConfig(column) {
        // Remove existing popovers
        const existingPopover = document.querySelector('.column-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create text column config popover
        const popover = document.createElement('div');
        popover.className = 'text-column-config';
        popover.innerHTML = `
            <div class="config-header">
                <h3>Text column</h3>
            </div>
            
            <div class="config-section">
                <label for="column-name">Column name</label>
                <input type="text" id="column-name" placeholder="Enter name" value="${column}">
            </div>
            
            <div class="config-divider"></div>
            
            <div class="config-section">
                <label class="checkbox-label">
                    <input type="checkbox" id="required-checkbox">
                    <span class="checkmark"></span>
                    Make this column required
                </label>
            </div>
            
            <div class="config-actions">
                <button class="btn btn-secondary" id="cancel-btn">Cancel</button>
                <button class="btn btn-primary" id="add-btn">Add</button>
            </div>
        `;

        // Position the popover below the column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        const headerRect = headerElement.getBoundingClientRect();
        popover.style.cssText = `
            position: fixed;
            top: ${headerRect.bottom + 5}px;
            left: ${headerRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            min-width: 300px;
            padding: 20px;
        `;

        // Add event listeners
        const cancelBtn = popover.querySelector('#cancel-btn');
        const addBtn = popover.querySelector('#add-btn');
        const columnNameInput = popover.querySelector('#column-name');

        cancelBtn.addEventListener('click', () => {
            popover.remove();
        });

        addBtn.addEventListener('click', () => {
            const newColumnName = columnNameInput.value.trim();
            if (newColumnName) {
                this.configureTextColumn(column, newColumnName);
                popover.remove();
            } else {
                columnNameInput.focus();
                columnNameInput.style.borderColor = '#ea4335';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Auto-focus the input
        setTimeout(() => {
            columnNameInput.focus();
            columnNameInput.select();
        }, 100);

        document.body.appendChild(popover);

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!popover.contains(e.target)) {
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    configureTextColumn(column, newColumnName) {
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = newColumnName;
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = newColumnName;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to accept text input
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                // Ensure the cell has an input field
                if (!cell.querySelector('input')) {
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.placeholder = '';
                    cell.appendChild(input);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Add a visual indicator that this is a text column
                cell.classList.add('text-column-cell');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        console.log(`Column ${column} configured as text column: "${newColumnName}"`);;
    }

    showDropdownColumnConfig(column) {
        // Remove existing popovers
        const existingPopover = document.querySelector('.column-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create dropdown column config popover
        const popover = document.createElement('div');
        popover.className = 'dropdown-column-config';
        popover.innerHTML = `
            <div class="config-header">
                <h3>Dropdown column</h3>
            </div>
            
            <div class="config-section">
                <label for="dropdown-column-name">Column name</label>
                <input type="text" id="dropdown-column-name" placeholder="Enter name" value="${column}">
            </div>
            
            <div class="config-section">
                <label>Dropdown options</label>
                <div id="dropdown-options-container">
                    <div class="option-row">
                        <input type="text" class="option-input" placeholder="Enter option">
                        <div class="option-actions">
                            <i class="fas fa-ellipsis-v"></i>
                        </div>
                    </div>
                    <div class="option-row">
                        <input type="text" class="option-input" placeholder="Enter option">
                        <div class="option-actions">
                            <i class="fas fa-ellipsis-v"></i>
                        </div>
                    </div>
                </div>
                <button class="btn btn-add-option" id="add-option-btn">
                    <i class="fas fa-plus"></i>
                    Add option
                </button>
            </div>
            
            <div class="config-divider"></div>
            
            <div class="config-section">
                <label class="checkbox-label">
                    <input type="checkbox" id="multiple-selections-checkbox">
                    <span class="checkmark"></span>
                    Allow multiple selections
                </label>
                <label class="checkbox-label">
                    <input type="checkbox" id="required-checkbox">
                    <span class="checkmark"></span>
                    Make this column required
                </label>
            </div>
            
            <div class="config-actions">
                <button class="btn btn-secondary" id="cancel-btn">Cancel</button>
                <button class="btn btn-primary" id="add-btn">Add</button>
            </div>
        `;

        // Position the popover below the column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        const headerRect = headerElement.getBoundingClientRect();
        popover.style.cssText = `
            position: fixed;
            top: ${headerRect.bottom + 5}px;
            left: ${headerRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            min-width: 350px;
            padding: 20px;
        `;

        // Add event listeners
        const cancelBtn = popover.querySelector('#cancel-btn');
        const addBtn = popover.querySelector('#add-btn');
        const addOptionBtn = popover.querySelector('#add-option-btn');
        const columnNameInput = popover.querySelector('#dropdown-column-name');
        const optionsContainer = popover.querySelector('#dropdown-options-container');

        // Add option functionality
        addOptionBtn.addEventListener('click', () => {
            const optionRow = document.createElement('div');
            optionRow.className = 'option-row';
            optionRow.innerHTML = `
                <input type="text" class="option-input" placeholder="Enter option">
                <div class="option-actions">
                    <i class="fas fa-ellipsis-v"></i>
                </div>
            `;
            optionsContainer.appendChild(optionRow);
        });

        // Remove option functionality (click on ellipsis)
        optionsContainer.addEventListener('click', (e) => {
            if (e.target.classList.contains('fa-ellipsis-v')) {
                const optionRow = e.target.closest('.option-row');
                if (optionsContainer.children.length > 1) {
                    optionRow.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        cancelBtn.addEventListener('click', () => {
            popover.remove();
        });

        addBtn.addEventListener('click', () => {
            console.log('Add button clicked');
            const newColumnName = columnNameInput.value.trim();
            const optionInputs = popover.querySelectorAll('.option-input');
            const options = Array.from(optionInputs)
                .map(input => input.value.trim())
                .filter(option => option.length > 0);
            const allowMultiple = popover.querySelector('#multiple-selections-checkbox').checked;
            const isRequired = popover.querySelector('#required-checkbox').checked;

            console.log('Column name:', newColumnName);
            console.log('Options:', options);
            console.log('Allow multiple:', allowMultiple);
            console.log('Required:', isRequired);

            if (newColumnName && options.length > 0) {
                console.log('Configuring dropdown column...');
                this.configureDropdownColumn(column, newColumnName, options, allowMultiple, isRequired);
                popover.remove();
            } else {
                if (!newColumnName) {
                    columnNameInput.focus();
                    columnNameInput.style.borderColor = '#ea4335';
                    console.log('No column name provided');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                if (options.length === 0) {
                    console.log('Please add at least one dropdown option');
                    console.log('No options provided');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Auto-focus the first input
        setTimeout(() => {
            columnNameInput.focus();
            columnNameInput.select();
        }, 100);

        document.body.appendChild(popover);

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!popover.contains(e.target)) {
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    configureDropdownColumn(column, newColumnName, options, allowMultiple, isRequired) {
        console.log('configureDropdownColumn called with:', { column, newColumnName, options, allowMultiple, isRequired });
        
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = newColumnName;
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = newColumnName;
            headerElement.dataset.columnType = 'dropdown';
            headerElement.dataset.dropdownOptions = JSON.stringify(options);
            headerElement.dataset.allowMultiple = allowMultiple;
            headerElement.dataset.isRequired = isRequired;
            console.log('Column header updated');
        } else {
            console.error('Header element not found for column:', column);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to have dropdowns
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        console.log(`Found ${cells.length} cells for column ${column}`);
        
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                // Remove existing input if any
                const existingInput = cell.querySelector('input');
                if (existingInput) {
                    existingInput.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

                // Create dropdown container
                const dropdownContainer = document.createElement('div');
                dropdownContainer.className = 'dropdown-cell-container';
                const row = cell.dataset.row;
                dropdownContainer.dataset.cellId = `${row}-${column}`;
                dropdownContainer.innerHTML = `
                    <div class="dropdown-display" data-column="${column}">
                        <span class="dropdown-placeholder">Select option</span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                `;

                // Add click event to show dropdown
                const dropdownDisplay = dropdownContainer.querySelector('.dropdown-display');
                dropdownDisplay.addEventListener('click', (e) => {
                    this.showDropdownOptions(e, column, options, allowMultiple);
                });

                cell.appendChild(dropdownContainer);
                cell.classList.add('dropdown-column-cell');
                console.log(`Updated cell ${row}-${column}`);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        console.log(`Column ${column} configured as dropdown: "${newColumnName}" with ${options.length} options`);
    }

    showDropdownOptions(e, column, options, allowMultiple) {
        // Check if dropdown is already open for this specific cell
        const cell = e.target.closest('.dropdown-cell-container');
        const existingDropdown = document.querySelector('.dropdown-options-popup');
        
        // If dropdown is already open for this cell, close it and return
        if (existingDropdown && existingDropdown.dataset.cellId === cell.dataset.cellId) {
            existingDropdown.remove();
            cell.classList.remove('open');
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Remove any other existing dropdowns
        if (existingDropdown) {
            existingDropdown.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Add open class to current cell
        cell.classList.add('open');
        const cellRect = cell.getBoundingClientRect();

        // Create dropdown options popup
        const dropdownPopup = document.createElement('div');
        dropdownPopup.className = 'dropdown-options-popup';
        dropdownPopup.dataset.cellId = cell.dataset.cellId;
        
        if (allowMultiple) {
            // Multiple selection with checkboxes
            dropdownPopup.innerHTML = `
                <div class="dropdown-options-list">
                    ${options.map(option => `
                        <label class="dropdown-option-checkbox">
                            <input type="checkbox" value="${option}">
                            <span>${option}</span>
                        </label>
                    `).join('')}
                </div>
            `;
        } else {
            // Single selection
            dropdownPopup.innerHTML = `
                <div class="dropdown-options-list">
                    ${options.map(option => `
                        <div class="dropdown-option" data-value="${option}">
                            ${option}
                        </div>
                    `).join('')}
                </div>
            `;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Position the dropdown
        dropdownPopup.style.cssText = `
            position: fixed;
            top: ${cellRect.bottom + 5}px;
            left: ${cellRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1001;
            min-width: ${cellRect.width}px;
            max-height: 200px;
            overflow-y: auto;
        `;

        // Add event listeners
        if (allowMultiple) {
            // Handle checkbox changes
            const checkboxes = dropdownPopup.querySelectorAll('input[type="checkbox"]');
            checkboxes.forEach(checkbox => {
                checkbox.addEventListener('change', () => {
                    const selectedOptions = Array.from(checkboxes)
                        .filter(cb => cb.checked)
                        .map(cb => cb.value);
                    
                    const display = cell.querySelector('.dropdown-placeholder');
                    if (selectedOptions.length > 0) {
                        display.textContent = selectedOptions.join(', ');
                        display.classList.add('has-selection');
                    } else {
                        display.textContent = 'Select option';
                        display.classList.remove('has-selection');
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                });
            });
        } else {
            // Handle single selection
            const optionElements = dropdownPopup.querySelectorAll('.dropdown-option');
            optionElements.forEach(option => {
                option.addEventListener('click', () => {
                    const value = option.dataset.value;
                    const display = cell.querySelector('.dropdown-placeholder');
                    display.textContent = value;
                    display.classList.add('has-selection');
                    dropdownPopup.remove();
                    cell.classList.remove('open');
                });
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        document.body.appendChild(dropdownPopup);

        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!dropdownPopup.contains(e.target) && !cell.contains(e.target)) {
                    dropdownPopup.remove();
                    cell.classList.remove('open');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    handleColumnPopoverAction(action, column) {
        switch (action) {
            case 'text':
                this.showTextColumnConfig(column);
                break;
            case 'dropdown':
                this.showDropdownColumnConfig(column);
                break;
            case 'file-upload':
                console.log(`File upload configured for column ${column}`);
                break;
            case 'jasper-iq':
                console.log(`Jasper IQ configured for column ${column}`);
                break;
            case 'import-data':
                console.log(`Import data configured for column ${column}`);
                break;
            case 'agent':
                this.showAgentLibraryModal(column);
                break;
            case 'app':
                this.showAppLibraryModal(column);
                break;
            case 'prompt':
                this.showPromptConfigPopover(column);
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    handleJasperIQAction(action, column) {
        let columnName = '';
        let dropdownOptions = [];

        switch (action) {
            case 'brand-voice':
                columnName = 'Brand Voice';
                dropdownOptions = [
                    'Professional & Corporate',
                    'Friendly & Approachable',
                    'Creative & Innovative',
                    'Authoritative & Expert',
                    'Casual & Conversational',
                    'Luxury & Premium',
                    'Technical & Detailed'
                ];
                break;
            case 'audience':
                columnName = 'Audience';
                dropdownOptions = [
                    'B2B Decision Makers',
                    'Small Business Owners',
                    'Enterprise Executives',
                    'Marketing Professionals',
                    'Sales Teams',
                    'Product Managers',
                    'C-Suite Leaders',
                    'Industry Specialists'
                ];
                break;
            case 'knowledge-base':
                columnName = 'Knowledge Base';
                dropdownOptions = [
                    'Product Documentation',
                    'Best Practices',
                    'Case Studies',
                    'Industry Research',
                    'Competitive Analysis',
                    'Market Trends',
                    'Technical Specs',
                    'User Guides'
                ];
                break;
            case 'style-guide':
                columnName = 'Style Guide';
                dropdownOptions = [
                    'AP Style',
                    'Chicago Manual',
                    'Company Brand',
                    'Academic',
                    'Journalistic',
                    'Technical Writing',
                    'Creative Copy',
                    'Legal Documents'
                ];
                break;
            case 'visual-guidelines':
                columnName = 'Visual Guidelines';
                dropdownOptions = [
                    'Minimalist Design',
                    'Bold & Colorful',
                    'Professional Layout',
                    'Creative Graphics',
                    'Data Visualization',
                    'Icon-Based',
                    'Typography Focused',
                    'Image-Heavy'
                ];
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        if (columnName && dropdownOptions.length > 0) {
            this.configureJasperIQColumn(column, columnName, dropdownOptions);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    configureJasperIQColumn(column, columnName, dropdownOptions) {
        console.log('configureJasperIQColumn called with:', { column, columnName, dropdownOptions });
        
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = columnName;
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = columnName;
            headerElement.dataset.columnType = 'jasper-iq';
            headerElement.dataset.dropdownOptions = JSON.stringify(dropdownOptions);
            console.log('Jasper IQ column header updated');
        } else {
            console.error('Header element not found for Jasper IQ column:', column);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to have dropdowns
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        console.log(`Found ${cells.length} cells for Jasper IQ column ${column}`);
        
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                // Remove existing input if any
                const existingInput = cell.querySelector('input');
                if (existingInput) {
                    existingInput.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

                // Remove existing dropdown if any
                const existingDropdown = cell.querySelector('.dropdown-cell-container');
                if (existingDropdown) {
                    existingDropdown.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

                // Create dropdown container
                const dropdownContainer = document.createElement('div');
                dropdownContainer.className = 'dropdown-cell-container jasper-iq-dropdown';
                const row = cell.dataset.row;
                dropdownContainer.dataset.cellId = `${row}-${column}`;
                dropdownContainer.innerHTML = `
                    <div class="dropdown-display" data-column="${column}">
                        <span class="dropdown-placeholder">Select ${columnName.toLowerCase()}</span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                `;

                // Add click event to show dropdown
                const dropdownDisplay = dropdownContainer.querySelector('.dropdown-display');
                dropdownDisplay.addEventListener('click', (e) => {
                    this.showJasperIQDropdownOptions(e, column, dropdownOptions);
                });

                cell.appendChild(dropdownContainer);
                cell.classList.add('jasper-iq-column-cell');
                console.log(`Updated Jasper IQ cell ${row}-${column}`);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        console.log(`Column ${column} configured as ${columnName} with ${dropdownOptions.length} options`);
    }

    showJasperIQDropdownOptions(e, column, options) {
        // Check if dropdown is already open for this specific cell
        const cell = e.target.closest('.dropdown-cell-container');
        const existingDropdown = document.querySelector('.jasper-iq-dropdown-popup');
        
        // If dropdown is already open for this cell, close it and return
        if (existingDropdown && existingDropdown.dataset.cellId === cell.dataset.cellId) {
            existingDropdown.remove();
            cell.classList.remove('open');
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Remove any other existing dropdowns
        if (existingDropdown) {
            existingDropdown.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Add open class to current cell
        cell.classList.add('open');
        const cellRect = cell.getBoundingClientRect();

        // Create dropdown options popup
        const dropdownPopup = document.createElement('div');
        dropdownPopup.className = 'jasper-iq-dropdown-popup';
        dropdownPopup.dataset.cellId = cell.dataset.cellId;
        dropdownPopup.innerHTML = `
            <div class="dropdown-options-list">
                ${options.map(option => `
                    <div class="dropdown-option" data-value="${option}">
                        ${option}
                    </div>
                `).join('')}
            </div>
        `;

        // Position the dropdown
        dropdownPopup.style.cssText = `
            position: fixed;
            top: ${cellRect.bottom + 5}px;
            left: ${cellRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1001;
            min-width: ${Math.max(cellRect.width, 250)}px;
            max-height: 200px;
            overflow-y: auto;
        `;

        // Add event listeners for single selection
        const optionElements = dropdownPopup.querySelectorAll('.dropdown-option');
        optionElements.forEach(option => {
            option.addEventListener('click', () => {
                const value = option.dataset.value;
                const display = cell.querySelector('.dropdown-placeholder');
                display.textContent = value;
                display.classList.add('has-selection');
                dropdownPopup.remove();
                cell.classList.remove('open');
            });
        });

        document.body.appendChild(dropdownPopup);

        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!dropdownPopup.contains(e.target) && !cell.contains(e.target)) {
                    dropdownPopup.remove();
                    cell.classList.remove('open');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    showAppLibraryModal(column) {
        // Remove existing modals
        const existingModal = document.querySelector('.app-library-modal');
        if (existingModal) {
            existingModal.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create App Library modal
        const modal = document.createElement('div');
        modal.className = 'app-library-modal';
        modal.innerHTML = `
            <div class="modal-overlay"></div>
            <div class="modal-content">
                <div class="modal-header">
                    <h2>App Library</h2>
                    <button class="close-btn" id="close-app-modal">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                
                <div class="modal-body">
                    <div class="sidebar">
                        <div class="search-section">
                            <div class="search-box">
                                <i class="fas fa-search"></i>
                                <input type="text" placeholder="Search" class="search-input">
                            </div>
                        </div>
                        
                        <nav class="main-nav">
                            <div class="nav-item active" data-category="all">
                                <i class="fas fa-th"></i>
                                <span>All</span>
                            </div>
                            <div class="nav-item" data-category="discover">
                                <i class="fas fa-home"></i>
                                <span>Discover</span>
                            </div>
                            <div class="nav-item" data-category="workspace">
                                <i class="fas fa-gem"></i>
                                <span>Workspace Apps</span>
                            </div>
                            <div class="nav-item" data-category="my-apps">
                                <i class="fas fa-user"></i>
                                <span>My Apps</span>
                            </div>
                            <div class="nav-item" data-category="favorites">
                                <i class="fas fa-star"></i>
                                <span>Favorites</span>
                            </div>
                        </nav>
                        
                        <div class="filters">
                            <div class="filter-section expanded">
                                <div class="filter-header" data-filter="marketing">
                                    <span>Marketing Function</span>
                                    <i class="fas fa-chevron-up"></i>
                                </div>
                                <div class="filter-options">
                                    <div class="filter-option" data-value="product-marketing">
                                        <i class="fas fa-box"></i>
                                        <span>Product Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="social-media">
                                        <i class="fas fa-user"></i>
                                        <span>Social Media Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="performance">
                                        <i class="fas fa-chart-bar"></i>
                                        <span>Performance Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="brand">
                                        <i class="fas fa-sparkles"></i>
                                        <span>Brand Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="content">
                                        <i class="fas fa-file-alt"></i>
                                        <span>Content Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="field">
                                        <i class="fas fa-globe"></i>
                                        <span>Field Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="lifecycle">
                                        <i class="fas fa-sync-alt"></i>
                                        <span>Lifecycle Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="partner">
                                        <i class="fas fa-briefcase"></i>
                                        <span>Partner Marketing</span>
                                    </div>
                                    <div class="filter-option" data-value="pr">
                                        <i class="fas fa-microphone"></i>
                                        <span>PR & Comms</span>
                                    </div>
                                </div>
                            </div>
                            
                            <div class="filter-section collapsed">
                                <div class="filter-header" data-filter="content-type">
                                    <span>Content Type</span>
                                    <i class="fas fa-chevron-down"></i>
                                </div>
                            </div>
                            
                            <div class="filter-section collapsed">
                                <div class="filter-header" data-filter="funnel-stage">
                                    <span>Funnel Stage</span>
                                    <i class="fas fa-chevron-down"></i>
                                </div>
                            </div>
                        </div>
                    </div>
                    
                    <div class="main-content">
                        <div class="apps-grid">
                            <div class="app-card" data-app="blog-post">
                                <div class="app-icon">
                                    <i class="fas fa-pen"></i>
                                </div>
                                <h3>Blog Post</h3>
                                <p>Write long-form content that provides value, drives traffic, and establishes thought leadership.</p>
                            </div>
                            
                            <div class="app-card" data-app="product-description">
                                <div class="app-icon">
                                    <i class="fas fa-cube"></i>
                                </div>
                                <h3>Product Description</h3>
                                <p>Compose detailed descriptions that highlight the benefits and features of your products.</p>
                            </div>
                            
                            <div class="app-card" data-app="facebook-post">
                                <div class="app-icon">
                                    <i class="fas fa-thumbs-up"></i>
                                </div>
                                <h3>Facebook Post</h3>
                                <p>Foster engagement and amplify reach using engaging Facebook updates and content.</p>
                            </div>
                            
                            <div class="app-card" data-app="content-rewriter">
                                <div class="app-icon">
                                    <i class="fas fa-sync-alt"></i>
                                </div>
                                <h3>Content Rewriter</h3>
                                <p>Transform your content to meet specific goals, including altering its tone and style.</p>
                            </div>
                            
                            <div class="app-card" data-app="social-campaign">
                                <div class="app-icon">
                                    <i class="fas fa-heart"></i>
                                </div>
                                <h3>Social Media Campaign</h3>
                                <p>Amplify your brand and engage followers with a cohesive social media strategy.</p>
                            </div>
                            
                            <div class="app-card" data-app="instagram-caption">
                                <div class="app-icon">
                                    <i class="fas fa-camera"></i>
                                </div>
                                <h3>Instagram Caption</h3>
                                <p>Boost engagement with captions that perfectly accompany your Instagram visuals.</p>
                            </div>
                            
                            <div class="app-card" data-app="email-sequence">
                                <div class="app-icon">
                                    <i class="fas fa-envelope"></i>
                                </div>
                                <h3>Email Sequence</h3>
                                <p>Guide customer journeys and boost conversions with a tailored email sequence.</p>
                            </div>
                            
                            <div class="app-card" data-app="meta-seo">
                                <div class="app-icon">
                                    <i class="fas fa-bookmark"></i>
                                </div>
                                <h3>Meta Title and Description</h3>
                                <p>Improve your webpage's visibility with SEO-friendly meta titles and descriptions.</p>
                            </div>
                            
                            <div class="app-card" data-app="linkedin-post">
                                <div class="app-icon">
                                    <i class="fas fa-briefcase"></i>
                                </div>
                                <h3>LinkedIn Post</h3>
                                <p>Enhance professional engagement on LinkedIn by sharing insights and industry knowledge.</p>
                            </div>
                            
                            <div class="app-card" data-app="ad-campaign">
                                <div class="app-icon">
                                    <i class="fas fa-bullhorn"></i>
                                </div>
                                <h3>Ad Campaign</h3>
                                <p>Target audiences on Meta, Google, and more with cohesive digital advertising campaigns.</p>
                            </div>
                            
                            <div class="app-card" data-app="multi-channel">
                                <div class="app-icon">
                                    <i class="fas fa-th"></i>
                                </div>
                                <h3>Multi-Channel Campaign</h3>
                                <p>Reach and resonate with audiences through an integrated marketing approach.</p>
                            </div>
                            
                            <div class="app-card" data-app="instructional">
                                <div class="app-icon">
                                    <i class="fas fa-bullseye"></i>
                                </div>
                                <h3>Instructional Post</h3>
                                <p>Guide readers through a detailed process to accomplish a specific goal or task.</p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        `;

        // Add event listeners
        const closeBtn = modal.querySelector('#close-app-modal');
        const navItems = modal.querySelectorAll('.nav-item');
        const filterHeaders = modal.querySelectorAll('.filter-header');
        const filterOptions = modal.querySelectorAll('.filter-option');
        const appCards = modal.querySelectorAll('.app-card');

        // Close modal
        closeBtn.addEventListener('click', () => {
            modal.remove();
        });

        // Navigation items
        navItems.forEach(item => {
            item.addEventListener('click', () => {
                navItems.forEach(nav => nav.classList.remove('active'));
                item.classList.add('active');
            });
        });

        // Filter headers (expand/collapse)
        filterHeaders.forEach(header => {
            header.addEventListener('click', () => {
                const section = header.closest('.filter-section');
                const isExpanded = section.classList.contains('expanded');
                
                if (isExpanded) {
                    section.classList.remove('expanded');
                    section.classList.add('collapsed');
                } else {
                    section.classList.remove('collapsed');
                    section.classList.add('expanded');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            });
        });

        // Filter options
        filterOptions.forEach(option => {
            option.addEventListener('click', () => {
                option.classList.toggle('selected');
            });
        });

        // App cards
        appCards.forEach(card => {
            card.addEventListener('click', () => {
                const appName = card.querySelector('h3').textContent;
                this.selectAppForColumn(column, appName);
                modal.remove();
            });
        });

        // Close modal when clicking overlay
        const overlay = modal.querySelector('.modal-overlay');
        overlay.addEventListener('click', () => {
            modal.remove();
        });

        document.body.appendChild(modal);
    }

    selectAppForColumn(column, appName) {
        console.log('selectAppForColumn called with:', { column, appName });
        
        // Check if this is the Blog Post app
        if (appName === 'Blog Post') {
            console.log('Blog Post detected, showing config popover...');
            this.showBlogPostConfigPopover(column);
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log('Not Blog Post, proceeding with default app configuration...');

        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = appName;
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = appName;
            headerElement.dataset.columnType = 'app';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update cells in the column to show app selection
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                const row = parseInt(cell.dataset.row);
                
                // In Option 2, only create app cells for rows 1-10 by default
                // In Option 1 and 3, create app cells for all rows
                if (this.currentVersion === 2 && row > 10) {
                    // Rows 11-20 remain as regular cells in Option 2 (will be greyed out in test mode)
                    return;
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Remove existing content
                cell.innerHTML = '';
                
                // Create app display
                const appDisplay = document.createElement('div');
                appDisplay.className = 'app-cell-display';
                appDisplay.innerHTML = `
                    <div class="app-cell-content">
                        <span>Click to run</span>
                    </div>
                `;
                
                // Add click handler
                appDisplay.addEventListener('click', () => {
                    this.runAppInCell(cell, appName);
                });
                
                cell.appendChild(appDisplay);
                cell.classList.add('app-column-cell');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Track this column as having an app and check Run All button
        this.appColumns.add(column);
        this.checkRunAllButton();
        
        // Show cost display when first app is added
        this.showCostDisplay();
        
        // Update credit count for Scale Mode (Option 2, not in test mode)
        if (this.currentVersion === 2 && !this.isTestMode) {
            this.columnCount++;
            this.updateCreditTextForScaleMode();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        console.log(`Column ${column} configured with app: ${appName}`);
    }

    runAppInCell(cell, appName) {
        console.log(`Running app: ${appName} in cell`);
        
        // Show loading state
        const appContent = cell.querySelector('.app-cell-content span');
        if (appContent) {
            appContent.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Running...';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Simulate app execution
        setTimeout(() => {
            // Show completion state
            if (appContent) {
                appContent.innerHTML = '<i class="fas fa-check"></i> Complete';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            console.log(`${appName} executed successfully!`);
            
            // Reset to clickable state after 2 seconds
            setTimeout(() => {
                if (appContent) {
                    appContent.textContent = 'Click to run';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 2000);
        }, 1500);
    }

    showAgentLibraryModal(column) {
        // Remove existing modals
        const existingModal = document.querySelector('.agent-library-modal');
        if (existingModal) {
            existingModal.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create Agent Library modal
        const modal = document.createElement('div');
        modal.className = 'agent-library-modal';
        
        // Add inline styles to ensure proper display
        modal.style.cssText = `
            position: fixed !important;
            top: 0 !important;
            left: 0 !important;
            width: 100% !important;
            height: 100% !important;
            z-index: 99999 !important;
            display: block !important;
            visibility: visible !important;
        `;
        
        modal.innerHTML = `
            <div class="modal-overlay" style="position: fixed !important; top: 0 !important; left: 0 !important; width: 100% !important; height: 100% !important; background: rgba(0, 0, 0, 0.5) !important; z-index: 99998 !important;"></div>
            <div class="modal-content" style="position: fixed !important; top: 50% !important; left: 50% !important; transform: translate(-50%, -50%) !important; width: 90% !important; max-width: 1000px !important; height: 90% !important; max-height: 700px !important; background: white !important; border-radius: 12px !important; box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3) !important; z-index: 99999 !important; overflow: hidden !important;">
                <div class="modal-header" style="padding: 24px 32px !important; border-bottom: 1px solid #e0e0e0 !important; display: flex !important; justify-content: space-between !important; align-items: center !important; background: white !important;">
                    <h2 style="margin: 0 !important; font-size: 24px !important; font-weight: 600 !important; color: #333 !important;">Agent Library</h2>
                    <button class="close-btn" id="close-agent-modal" style="background: none !important; border: none !important; font-size: 20px !important; color: #666 !important; cursor: pointer !important; padding: 8px !important; border-radius: 6px !important;">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                
                <div class="modal-body" style="padding: 32px !important; height: calc(100% - 80px) !important; overflow-y: auto !important; display: flex !important; flex-direction: column !important;">
                    <div class="search-filter-section" style="display: flex !important; justify-content: space-between !important; align-items: flex-end !important; margin-bottom: 32px !important; gap: 24px !important; flex-shrink: 0 !important;">
                        <div class="search-box" style="flex: 1 !important; max-width: 400px !important;">
                            <label for="agent-search" style="display: block !important; margin-bottom: 8px !important; font-size: 14px !important; font-weight: 500 !important; color: #333 !important;">Search</label>
                            <input type="text" id="agent-search" placeholder="Search agents..." style="width: 100% !important; padding: 12px 16px !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; font-size: 14px !important; outline: none !important;">
                        </div>
                        
                        <div class="filter-sort-buttons" style="display: flex !important; gap: 12px !important;">
                            <button class="btn btn-filter" style="display: flex !important; align-items: center !important; gap: 8px !important; padding: 12px 16px !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; background: white !important; color: #666 !important; font-size: 14px !important; font-weight: 500 !important; cursor: pointer !important; min-width: 100px !important; justify-content: center !important;">
                                <i class="fas fa-filter"></i>
                                Filter
                            </button>
                            <button class="btn btn-sort" style="display: flex !important; align-items: center !important; gap: 8px !important; padding: 12px 16px !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; background: white !important; color: #666 !important; font-size: 14px !important; font-weight: 500 !important; cursor: pointer !important; min-width: 100px !important; justify-content: center !important;">
                                <i class="fas fa-sort"></i>
                                Sort
                            </button>
                        </div>
                    </div>
                    
                    <div class="agents-grid" style="display: grid !important; grid-template-columns: repeat(2, 1fr) !important; gap: 24px !important; width: 100% !important; flex: 1 !important;">
                        <div class="agent-card" data-agent="optimization" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon optimization" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-mouse-pointer" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Optimization Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Uncover keywords, optimize content, and drive organic traffic</p>
                        </div>
                        
                        <div class="agent-card" data-agent="personalization" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon personalization" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-envelope" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Personalization Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Generate high-value content for target accounts that converts</p>
                        </div>
                        
                        <div class="agent-card" data-agent="research" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon research" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-search" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Research Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Research on a specific topic</p>
                        </div>
                        
                        <div class="agent-card" data-agent="keyword" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon keyword" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-key" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Keyword Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Optimize your website's SEO with a curated list of high-performing keywords</p>
                        </div>
                        
                        <div class="agent-card" data-agent="audit" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon audit" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-clipboard-list" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Audit Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Review and analyze websites or marketing content to assess quality and performance</p>
                        </div>
                        
                        <div class="agent-card" data-agent="rewrite" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon rewrite" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-pen" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Rewrite Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Transform your content to meet specific goals, including altering its tone and style</p>
                        </div>
                        
                        <div class="agent-card" data-agent="competitor" style="background: white !important; border: 1px solid #e0e0e0 !important; border-radius: 8px !important; padding: 20px !important; display: flex !important; flex-direction: column !important; box-shadow: 0 2px 4px rgba(0, 0, 0, 0.05) !important; min-height: 160px !important;">
                            <div class="agent-icon competitor" style="width: 48px !important; height: 48px !important; border-radius: 50% !important; display: flex !important; align-items: center !important; justify-content: center !important; margin-bottom: 16px !important; font-size: 20px !important; background: #f0f7ff !important; color: #1a73e8 !important; border: 1px solid #e0e0e0 !important;">
                                <i class="fas fa-globe" style="display: block !important; font-size: 20px !important; color: #1a73e8 !important;"></i>
                            </div>
                            <h3 style="margin: 0 0 8px 0 !important; font-size: 16px !important; font-weight: 600 !important; color: #333 !important; line-height: 1.2 !important;">Competitor URL Agent</h3>
                            <p style="margin: 0 !important; font-size: 14px !important; color: #666 !important; line-height: 1.5 !important;">Monitor competitor URLs and get actionable insights on messaging, positioning, and strategy</p>
                        </div>
                    </div>
                </div>
            </div>
        `;

        // Add event listeners
        const closeBtn = modal.querySelector('#close-agent-modal');
        const searchInput = modal.querySelector('#agent-search');
        const filterBtn = modal.querySelector('.btn-filter');
        const sortBtn = modal.querySelector('.btn-sort');

        // Close modal
        closeBtn.addEventListener('click', () => {
            modal.remove();
        });

        // Search functionality
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.toLowerCase();
            this.filterAgentCards(searchTerm);
        });

        // Filter button (placeholder for future functionality)
        filterBtn.addEventListener('click', () => {
            console.log('Filter functionality coming soon!');
        });

        // Sort button (placeholder for future functionality)
        sortBtn.addEventListener('click', () => {
            console.log('Sort functionality coming soon!');
        });

        // Close modal when clicking overlay
        const overlay = modal.querySelector('.modal-overlay');
        overlay.addEventListener('click', () => {
            modal.remove();
        });

        document.body.appendChild(modal);

        // Debug: Log modal creation
        console.log('Agent Library modal created and appended to DOM');
        console.log('Modal element:', modal);
        console.log('Modal classes:', modal.className);
        console.log('Modal styles:', modal.style.cssText);
        console.log('Modal HTML:', modal.innerHTML);

        // Auto-focus the search input
        setTimeout(() => {
            searchInput.focus();
        }, 100);
    }

    filterAgentCards(searchTerm) {
        const agentCards = document.querySelectorAll('.agent-card');
        agentCards.forEach(card => {
            const title = card.querySelector('h3').textContent.toLowerCase();
            const description = card.querySelector('p').textContent.toLowerCase();
            
            if (title.includes(searchTerm) || description.includes(searchTerm)) {
                card.style.display = 'block';
            } else {
                card.style.display = 'none';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });
    }

    selectAgentForColumn(column, agentName) {
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = agentName.charAt(0).toUpperCase() + agentName.slice(1) + ' Agent';
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = agentName + ' Agent';
            headerElement.dataset.columnType = 'agent';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to show agent selection
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                // Remove existing content
                cell.innerHTML = '';
                
                // Create agent display
                const agentDisplay = document.createElement('div');
                agentDisplay.className = 'agent-cell-display';
                agentDisplay.innerHTML = `
                    <div class="agent-cell-content">
                        <i class="fas fa-infinity"></i>
                        <span>${agentName.charAt(0).toUpperCase() + agentName.slice(1)} Agent</span>
                    </div>
                `;
                
                cell.appendChild(agentDisplay);
                cell.classList.add('agent-column-cell');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        console.log(`Column ${column} configured with agent: ${agentName.charAt(0).toUpperCase() + agentName.slice(1)} Agent`);
    }

    showAddColumnPopover() {
        // Remove existing popovers
        const existingPopover = document.querySelector('.column-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create column popover for new column
        const popover = document.createElement('div');
        popover.className = 'column-popover add-column-popover';
        popover.innerHTML = `
            <div class="popover-section">
                <div class="section-title">Input</div>
                <div class="popover-item" data-action="text">
                    <span>Text</span>
                </div>
                <div class="popover-item" data-action="dropdown">
                    <span>Dropdown</span>
                </div>
                <div class="popover-item" data-action="file-upload">
                    <span>File upload</span>
                </div>
                <div class="popover-item has-submenu" data-action="jasper-iq">
                    <span>Jasper IQ</span>
                    <i class="fas fa-chevron-right"></i>
                </div>
                <div class="popover-item" data-action="import-data">
                    <span>Import data</span>
                </div>
            </div>
            <div class="popover-divider"></div>
            <div class="popover-section">
                <div class="section-title">Output</div>
                <div class="popover-item" data-action="agent">
                    <span>∞ Agent</span>
                </div>
                <div class="popover-item" data-action="app">
                    <i class="fas fa-cube"></i>
                    <span>App</span>
                </div>
                <div class="popover-item" data-action="prompt">
                    <i class="fas fa-desktop"></i>
                    <span>Prompt</span>
                </div>
            </div>
        `;

        // Position the popover below the "Add column" button
        const addColumnBtn = document.querySelector('.add-column-btn');
        const btnRect = addColumnBtn.getBoundingClientRect();
        
        // Calculate position to ensure menu is fully visible
        let leftPosition = btnRect.left;
        let topPosition = btnRect.bottom + 5;
        
        // Check if menu would be cut off on the right
        const menuWidth = 200; // min-width of the popover
        const rightEdge = leftPosition + menuWidth;
        const viewportWidth = window.innerWidth;
        
        if (rightEdge > viewportWidth - 20) {
            // Position menu to the left of the button if it would be cut off
            leftPosition = btnRect.right - menuWidth;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Check if menu would be cut off on the bottom
        const menuHeight = 200; // approximate height of the popover
        const bottomEdge = topPosition + menuHeight;
        const viewportHeight = window.innerHeight;
        
        if (bottomEdge > viewportHeight - 20) {
            // Position menu above the button if it would be cut off
            topPosition = btnRect.top - menuHeight - 5;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        popover.style.cssText = `
            position: fixed;
            top: ${topPosition}px;
            left: ${leftPosition}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1000;
            min-width: ${menuWidth}px;
            padding: 16px 0;
        `;

        // Add event listeners to popover items
        popover.addEventListener('click', (e) => {
            const item = e.target.closest('.popover-item');
            if (item) {
                const action = item.dataset.action;
                if (action !== 'jasper-iq') {
                    this.handleAddColumnAction(action);
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Add hover events for Jasper IQ submenu
        const jasperItem = popover.querySelector('[data-action="jasper-iq"]');
        if (jasperItem) {
            jasperItem.addEventListener('mouseenter', () => {
                this.showJasperIQSubmenuForNewColumn(jasperItem);
            });
            
            jasperItem.addEventListener('mouseleave', () => {
                // Hide submenu when mouse leaves Jasper IQ item
                setTimeout(() => {
                    const submenu = document.querySelector('.jasper-iq-submenu');
                    if (submenu && !submenu.matches(':hover')) {
                        submenu.remove();
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                }, 100);
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        document.body.appendChild(popover);

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', () => {
                popover.remove();
            }, { once: true });
        }, 100);
    }

    showJasperIQSubmenuForNewColumn(jasperItem) {
        // Remove existing submenus
        const existingSubmenu = document.querySelector('.jasper-iq-submenu');
        if (existingSubmenu) {
            existingSubmenu.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create Jasper IQ submenu for new column
        const submenu = document.createElement('div');
        submenu.className = 'jasper-iq-submenu';
        submenu.innerHTML = `
            <div class="submenu-header">
                <h4>Jasper IQ</h4>
            </div>
            <div class="submenu-item" data-action="brand-voice">Brand Voice</div>
            <div class="submenu-item" data-action="audience">Audience</div>
            <div class="submenu-item" data-action="knowledge-base">Knowledge Base</div>
            <div class="submenu-item" data-action="style-guide">Style Guide</div>
            <div class="submenu-item" data-action="visual-guidelines">Visual Guidelines</div>
        `;

        // Position the submenu to the left of the main popover for "Add column" version
        const jasperRect = jasperItem.getBoundingClientRect();
        const mainPopover = jasperItem.closest('.column-popover');
        const popoverRect = mainPopover.getBoundingClientRect();
        
        // For "Add column" popover, always position submenu to the left
        const submenuLeft = popoverRect.left - 185; // 180px width + 5px gap

        submenu.style.cssText = `
            position: fixed;
            top: ${jasperRect.top}px;
            left: ${submenuLeft}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
            z-index: 1001;
            min-width: 180px;
            padding: 12px 0;
        `;

        // Add event listeners to submenu items
        submenu.addEventListener('click', (e) => {
            const item = e.target.closest('.submenu-item');
            if (item) {
                const action = item.dataset.action;
                this.handleAddColumnJasperIQAction(action);
                submenu.remove();
                // Also close the main popover
                const mainPopover = document.querySelector('.add-column-popover');
                if (mainPopover) {
                    mainPopover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Add hover events to keep submenu open
        submenu.addEventListener('mouseenter', () => {
            // Keep submenu open when hovering over it
        });

        submenu.addEventListener('mouseleave', () => {
            // Close submenu when mouse leaves
            setTimeout(() => {
                if (submenu.parentNode) {
                    submenu.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 100);
        });

        document.body.appendChild(submenu);
    }

    handleAddColumnAction(action) {
        // First add the column, then configure it
        this.addColumn();
        const newColumn = String.fromCharCode(65 + this.columns - 1);
        
        switch (action) {
            case 'text':
                this.showTextColumnConfig(newColumn);
                break;
            case 'dropdown':
                this.showDropdownColumnConfig(newColumn);
                break;
            case 'file-upload':
                console.log(`File upload configured for new column ${newColumn}`);
                break;
            case 'import-data':
                console.log(`Import data configured for new column ${newColumn}`);
                break;
            case 'agent':
                console.log(`∞ Agent output configured for new column ${newColumn}`);
                break;
            case 'app':
                this.showAppLibraryModal(newColumn);
                break;
            case 'prompt':
                this.showPromptConfigPopover(newColumn);
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    handleAddColumnJasperIQAction(action) {
        // First add the column, then configure it with Jasper IQ
        this.addColumn();
        const newColumn = String.fromCharCode(65 + this.columns - 1);
        
        let columnName = '';
        let dropdownOptions = [];

        switch (action) {
            case 'brand-voice':
                columnName = 'Brand Voice';
                dropdownOptions = [
                    'Professional & Corporate',
                    'Friendly & Approachable',
                    'Creative & Innovative',
                    'Authoritative & Expert',
                    'Casual & Conversational',
                    'Luxury & Premium',
                    'Technical & Detailed'
                ];
                break;
            case 'audience':
                columnName = 'Audience';
                dropdownOptions = [
                    'B2B Decision Makers',
                    'Small Business Owners',
                    'Enterprise Executives',
                    'Marketing Professionals',
                    'Sales Teams',
                    'Product Managers',
                    'C-Suite Leaders',
                    'Industry Specialists'
                ];
                break;
            case 'knowledge-base':
                columnName = 'Knowledge Base';
                dropdownOptions = [
                    'Product Documentation',
                    'Best Practices',
                    'Case Studies',
                    'Industry Research',
                    'Competitive Analysis',
                    'Market Trends',
                    'Technical Specs',
                    'User Guides'
                ];
                break;
            case 'style-guide':
                columnName = 'Style Guide';
                dropdownOptions = [
                    'AP Style',
                    'Chicago Manual',
                    'Company Brand',
                    'Academic',
                    'Journalistic',
                    'Technical Writing',
                    'Creative Copy',
                    'Legal Documents'
                ];
                break;
            case 'visual-guidelines':
                columnName = 'Visual Guidelines';
                dropdownOptions = [
                    'Minimalist Design',
                    'Bold & Colorful',
                    'Professional Layout',
                    'Creative Graphics',
                    'Data Visualization',
                    'Icon-Based',
                    'Typography Focused',
                    'Image-Heavy'
                ];
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        if (columnName && dropdownOptions.length > 0) {
            this.configureJasperIQColumn(newColumn, columnName, dropdownOptions);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    handleContextMenuAction(action) {
        switch (action) {
            case 'copy':
                console.log('Copied to clipboard');
                break;
            case 'paste':
                console.log('Pasted from clipboard');
                break;
            case 'cut':
                console.log('Cut to clipboard');
                break;
            case 'delete':
                if (this.currentCell) {
                    const input = this.currentCell.querySelector('input');
                    if (input) {
                        input.value = '';
                        const cellKey = `${this.currentCell.dataset.row}-${this.currentCell.dataset.col}`;
                        delete this.data[cellKey];
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                console.log('Cell cleared');
                break;
            case 'format':
                console.log('Format options opened');
                break;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showPromptConfigPopover(column) {
        // Remove existing prompt config popovers
        const existingPopover = document.querySelector('.prompt-config-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Get existing columns for context dropdown and @ button
        const existingColumns = this.getExistingColumns(column);
        
        // Create prompt config popover
        const popover = document.createElement('div');
        popover.className = 'prompt-config-popover';
        popover.innerHTML = `
            <div class="popover-header">
                <i class="fas fa-clipboard-arrow-right"></i>
                <span>Prompt</span>
            </div>
            
            <div class="prompt-input-section">
                <textarea 
                    class="prompt-textarea" 
                    placeholder="Describe the workflow or content you want to scale..."
                    rows="4"
                ></textarea>
                <div class="prompt-toolbar">
                    <button class="toolbar-btn" data-action="add">
                        <i class="fas fa-plus"></i>
                    </button>
                    <button class="toolbar-btn" data-action="clipboard">
                        <i class="fas fa-clipboard"></i>
                    </button>
                    <button class="toolbar-btn" data-action="at-symbol">
                        <i class="fas fa-at"></i>
                    </button>
                </div>
            </div>
            
            <div class="context-section">
                <label>Add column for context</label>
                <div class="context-dropdown">
                    <select class="context-select">
                        <option value="">Select column(s)</option>
                        ${existingColumns.map(col => `<option value="${col}">${col}</option>`).join('')}
                    </select>
                    <i class="fas fa-chevron-down"></i>
                </div>
            </div>
            
            <div class="popover-actions">
                <button class="btn btn-cancel">Cancel</button>
                <button class="btn btn-add">Add</button>
            </div>
        `;

        // Position the popover
        const columnHeader = document.querySelector(`[data-col="${column}"]`);
        if (columnHeader) {
            const rect = columnHeader.getBoundingClientRect();
            popover.style.cssText = `
                position: fixed;
                top: ${rect.bottom + 10}px;
                left: ${rect.left}px;
                background: white;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15);
                z-index: 1000;
                width: 400px;
                padding: 20px;
            `;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Add event listeners
        const cancelBtn = popover.querySelector('.btn-cancel');
        const addBtn = popover.querySelector('.btn-add');
        const atButton = popover.querySelector('[data-action="at-symbol"]');
        const contextSelect = popover.querySelector('.context-select');

        // Cancel button
        cancelBtn.addEventListener('click', () => {
            popover.remove();
        });

        // Add button
        addBtn.addEventListener('click', () => {
            const promptText = popover.querySelector('.prompt-textarea').value;
            const contextColumn = contextSelect.value;
            
            if (promptText.trim()) {
                this.configurePromptColumn(column, promptText, contextColumn);
                popover.remove();
            } else {
                console.log('Please enter a prompt description');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // @ button - show column suggestions
        atButton.addEventListener('click', () => {
            this.showColumnSuggestions(atButton, existingColumns);
        });

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!popover.contains(e.target)) {
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);

        document.body.appendChild(popover);
    }

    getExistingColumns(currentColumn) {
        // Get all existing columns that are to the left of the current column
        const columns = [];
        const currentColIndex = currentColumn.charCodeAt(0) - 65; // Convert A, B, C to 0, 1, 2
        
        console.log('🔍 getExistingColumns called with:', currentColumn, 'index:', currentColIndex);
        
        for (let i = 0; i < currentColIndex; i++) {
            const colLetter = String.fromCharCode(65 + i);
            const headerElement = document.querySelector(`[data-col="${colLetter}"]`);
            console.log('🔍 Looking for column', colLetter, 'found:', headerElement);
            if (headerElement) {
                const columnName = headerElement.textContent || colLetter;
                columns.push(columnName);
                console.log('✅ Added column:', columnName);
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        console.log('📋 Final columns array:', columns);
        return columns;
    }

    showColumnSuggestions(button, columns) {
        // Remove existing suggestions
        const existingSuggestions = document.querySelector('.column-suggestions');
        if (existingSuggestions) {
            existingSuggestions.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        if (columns.length === 0) {
            console.log('No columns available for context');
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create suggestions popup
        const suggestions = document.createElement('div');
        suggestions.className = 'column-suggestions';
        suggestions.innerHTML = `
            <div class="suggestions-header">Available Columns</div>
            ${columns.map(col => `<div class="suggestion-item" data-column="${col}">@${col}</div>`).join('')}
        `;

        // Position suggestions near the @ button
        const buttonRect = button.getBoundingClientRect();
        suggestions.style.cssText = `
            position: fixed;
            top: ${buttonRect.bottom + 5}px;
            left: ${buttonRect.left}px;
            background: white;
            border: 1px solid #e0e0e0;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
            z-index: 1001;
            min-width: 150px;
            max-height: 200px;
            overflow-y: auto;
        `;

        // Add click events to suggestions
        suggestions.addEventListener('click', (e) => {
            const item = e.target.closest('.suggestion-item');
            if (item) {
                const columnName = item.dataset.column;
                // Look for textarea in both prompt and blog post popovers
                const textarea = button.closest('.prompt-config-popover, .blog-post-config-popover').querySelector('.prompt-textarea, .blog-textarea');
                if (textarea) {
                    const cursorPos = textarea.selectionStart;
                    const textBefore = textarea.value.substring(0, cursorPos);
                    const textAfter = textarea.value.substring(cursorPos);
                    
                    // Insert @column at cursor position
                    textarea.value = textBefore + '@' + columnName + ' ' + textAfter;
                    
                    // Set cursor position after the inserted text
                    const newPos = cursorPos + columnName.length + 2; // +2 for @ and space
                    textarea.setSelectionRange(newPos, newPos);
                    textarea.focus();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                suggestions.remove();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Close suggestions when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!suggestions.contains(e.target) && !button.contains(e.target)) {
                    suggestions.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);

        document.body.appendChild(suggestions);
    }

    configurePromptColumn(column, promptText, contextColumn) {
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = 'Prompt';
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = 'Prompt';
            headerElement.dataset.columnType = 'prompt';
            headerElement.dataset.promptText = promptText;
            if (contextColumn) {
                headerElement.dataset.contextColumn = contextColumn;
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to show prompt output
        for (let row = 1; row <= this.rows; row++) {
            const cellKey = `${row}-${column}`;
            const cell = document.querySelector(`[data-row="${row}"][data-col="${column}"]`);
            if (cell) {
                const input = cell.querySelector('input');
                if (input) {
                    input.value = `Prompt: ${promptText}`;
                    input.readOnly = true;
                    input.style.backgroundColor = '#f8f9fa';
                    input.style.color = '#6c757d';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        console.log(`Prompt column "${column}" configured successfully`);
    }

    showBlogPostConfigPopover(column) {
        // Remove existing popovers
        const existingPopover = document.querySelector('.blog-post-config-popover');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create blog post config popover
        const popover = document.createElement('div');
        popover.className = 'blog-post-config-popover';
        popover.innerHTML = `
            <div class="popover-header">
                <div class="header-content">
                    <div class="header-icon">
                        <i class="fas fa-pen"></i>
                    </div>
                    <h3>Blog Post</h3>
                </div>
            </div>
            
            <div class="config-section">
                <label>App Settings</label>
                <div class="app-selector">
                    <i class="fas fa-bullhorn"></i>
                    <span>Continental Studios • English...</span>
                    <i class="fas fa-chevron-down"></i>
                </div>
            </div>
            
            <div class="config-section">
                <h4>Details</h4>
                <div class="input-group">
                    <label>What is the topic of your post?</label>
                    <div class="textarea-container">
                        <textarea class="blog-textarea" placeholder="Enter your blog post topic..."></textarea>
                        <button class="toolbar-btn" title="Add context">
                            <i class="fas fa-at"></i>
                        </button>
                    </div>
                </div>
                
                <div class="input-group">
                    <label>How long would you like the blog post to be?</label>
                    <div class="dropdown-selector">
                        <span class="dropdown-placeholder">Select length</span>
                        <i class="fas fa-chevron-down"></i>
                    </div>
                </div>
                
                <div class="input-group">
                    <label>What is the outline of your post?</label>
                    <div class="textarea-container">
                        <textarea class="blog-textarea" placeholder="Enter your blog post outline..."></textarea>
                        <button class="toolbar-btn" title="Add context">
                            <i class="fas fa-at"></i>
                        </button>
                    </div>
                </div>
            </div>
            
            <div class="config-section">
                <h4>Add additional context</h4>
                <div class="dropdown-selector" id="context-dropdown">
                    <span class="dropdown-placeholder">Select column(s)</span>
                    <i class="fas fa-chevron-down"></i>
                </div>
            </div>
            
            <div class="config-actions">
                <button class="btn btn-secondary" id="cancel-blog-btn">Cancel</button>
                <button class="btn btn-primary" id="add-blog-btn">Add</button>
            </div>
        `;

        // Position the popover below the column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        const headerRect = headerElement.getBoundingClientRect();
        
        // Add some basic inline styles to ensure visibility
        popover.style.cssText = `
            position: fixed !important;
            top: ${headerRect.bottom + 5}px !important;
            left: ${headerRect.left}px !important;
            background: white !important;
            border: 1px solid #e0e0e0 !important;
            border-radius: 8px !important;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.15) !important;
            z-index: 10000 !important;
            min-width: 400px !important;
            padding: 20px !important;
            display: block !important;
            visibility: visible !important;
            opacity: 1 !important;
        `;

        // Add event listeners
        const cancelBtn = popover.querySelector('#cancel-blog-btn');
        const addBtn = popover.querySelector('#add-blog-btn');
        
        // Debug: Check all elements in the popover
        console.log('🔍 Popover element:', popover);
        console.log('🔍 Popover HTML:', popover.innerHTML);
        
        const atButtons = popover.querySelectorAll('.toolbar-btn');
        console.log('🔍 Found @ buttons:', atButtons.length, atButtons);
        
        // Debug: Check if buttons exist with different selectors
        const allButtons = popover.querySelectorAll('button');
        console.log('🔍 All buttons found:', allButtons.length, allButtons);
        allButtons.forEach((btn, i) => {
            console.log(`🔍 Button ${i}:`, btn.className, btn.innerHTML);
        });

        // Add event listeners for @ buttons
        atButtons.forEach((button, index) => {
            console.log(`🔍 Adding event listener to @ button ${index}:`, button);
            button.addEventListener('click', () => {
                console.log('🔘 @ button clicked! Column:', column);
                const columns = this.getExistingColumns(column);
                console.log('🔘 Available columns from getExistingColumns:', columns);
                if (columns.length > 0) {
                    this.showColumnSuggestions(button, columns);
                } else {
                    console.log('No columns available to the left');
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            });
        });

        // Populate the context dropdown with available columns
        const contextDropdown = popover.querySelector('#context-dropdown');
        const availableColumns = this.getExistingColumns(column);
        
        if (availableColumns.length > 0) {
            contextDropdown.addEventListener('click', () => {
                this.showContextDropdownOptions(contextDropdown, availableColumns);
            });
        } else {
            contextDropdown.style.opacity = '0.5';
            contextDropdown.style.cursor = 'not-allowed';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        cancelBtn.addEventListener('click', () => {
            popover.remove();
        });

        addBtn.addEventListener('click', () => {
            const topicText = popover.querySelector('.blog-textarea').value.trim();
            const outlineText = popover.querySelectorAll('.blog-textarea')[1].value.trim();
            
            if (topicText && outlineText) {
                this.configureBlogPostColumn(column, topicText, outlineText);
                popover.remove();
            } else {
                console.log('Please fill in both topic and outline fields');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        // Auto-focus the first textarea
        setTimeout(() => {
            const firstTextarea = popover.querySelector('.blog-textarea');
            firstTextarea.focus();
        }, 100);

        document.body.appendChild(popover);
        


        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!popover.contains(e.target)) {
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    configureBlogPostColumn(column, topicText, outlineText) {
        // Update column header
        const headerElement = document.querySelector(`[data-col="${column}"]`);
        if (headerElement) {
            headerElement.textContent = 'Blog Post';
            headerElement.dataset.originalCol = column;
            headerElement.dataset.customName = 'Blog Post';
            headerElement.dataset.columnType = 'blog-post';
            headerElement.dataset.topic = topicText;
            headerElement.dataset.outline = outlineText;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Update all cells in the column to show blog post configuration
        const cells = document.querySelectorAll(`[data-col="${column}"]`);
        cells.forEach(cell => {
            if (cell.classList.contains('cell')) {
                // Remove existing content
                const existingInput = cell.querySelector('input');
                if (existingInput) {
                    existingInput.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

                // Create blog post display
                const blogDisplay = document.createElement('div');
                blogDisplay.className = 'blog-post-cell-display';
                blogDisplay.innerHTML = `
                    <div class="blog-post-content">
                        <i class="fas fa-pen"></i>
                        <span>Blog Post: ${topicText.substring(0, 30)}${topicText.length > 30 ? '...' : ''}</span>
                    </div>
                `;

                cell.appendChild(blogDisplay);
                cell.classList.add('blog-post-column-cell');
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        });

        console.log('Blog Post column configured successfully');
    }

    showContextDropdownOptions(dropdownElement, columns) {
        // Remove any existing dropdown options
        const existingOptions = document.querySelector('.context-dropdown-options');
        if (existingOptions) {
            existingOptions.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create dropdown options popup
        const optionsPopup = document.createElement('div');
        optionsPopup.className = 'context-dropdown-options';
        optionsPopup.innerHTML = `
            <div class="dropdown-options-list">
                ${columns.map(col => `
                    <div class="dropdown-option" data-value="${col}">
                        ${col}
                    </div>
                `).join('')}
            </div>
        `;

        // Position the dropdown below the context dropdown
        const dropdownRect = dropdownElement.getBoundingClientRect();
        optionsPopup.style.cssText = `
            position: fixed !important;
            top: ${dropdownRect.bottom + 5}px !important;
            left: ${dropdownRect.left}px !important;
            background: white !important;
            border: 1px solid #e0e0e0 !important;
            border-radius: 6px !important;
            box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15) !important;
            z-index: 10001 !important;
            min-width: ${dropdownRect.width}px !important;
            max-height: 200px !important;
            overflow-y: auto !important;
        `;

        // Add event listeners for options
        const optionElements = optionsPopup.querySelectorAll('.dropdown-option');
        optionElements.forEach(option => {
            option.addEventListener('click', () => {
                const value = option.dataset.value;
                const placeholder = dropdownElement.querySelector('.dropdown-placeholder');
                placeholder.textContent = value;
                placeholder.style.color = '#333';
                optionsPopup.remove();
            });
        });

        document.body.appendChild(optionsPopup);

        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                if (!optionsPopup.contains(e.target) && !dropdownElement.contains(e.target)) {
                    optionsPopup.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);
    }

    setupVersionToggle() {
        const versionBtns = document.querySelectorAll('.version-btn');
        versionBtns.forEach(btn => {
            btn.addEventListener('click', () => {
                const version = parseInt(btn.dataset.version);
                this.switchVersion(version);
            });
        });
    }

    loadSavedVersion() {
        // Set the correct active button
        document.querySelectorAll('.version-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-version="${this.currentVersion}"]`).classList.add('active');
        
        // Apply the saved version
        if (this.currentVersion === 1) {
            this.enableAllRows();
            this.showOption1CreditDisplay();
            this.hideTestModeToggle();
            this.hideTestModePillOption4();
        } else if (this.currentVersion === 2) {
            // Option 2 always starts in test mode
            this.isTestMode = true;
            localStorage.setItem('testMode', 'true');
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
        } else if (this.currentVersion === 3) {
            this.hideTestModeToggle();
            this.hideTestModePillOption4();
            this.showOption3CreditDisplay();
            this.enableAllRows();
            // Setup the credit pill for Option 3
            this.setupOption3CreditPill();
        } else if (this.currentVersion === 4) {
            this.hideTestModeToggle();
            this.showTestModePillOption4();
            this.setupTestModePillDropdownOption4();
            if (this.isTestMode) {
                this.greyOutRows11To20();
            } else {
                this.enableAllRows();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    switchVersion(version) {
        // Update active button
        document.querySelectorAll('.version-btn').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[data-version="${version}"]`).classList.add('active');
        
        this.currentVersion = version;
        
        // Save to localStorage
        localStorage.setItem('selectedVersion', version.toString());
        
        if (version === 1) {
            this.enableAllRows();
            this.showOption1CreditDisplay();
            this.hideTestModeToggle();
            this.hideTestModePillOption4();
        } else if (version === 2) {
            // Option 2 always starts in test mode
            this.isTestMode = true;
            localStorage.setItem('testMode', 'true');
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 2
        } else if (version === 3) {
            // Option 3: All rows unlocked by default, no test mode toggle
            this.hideTestModeToggle();
            this.hideTestModePillOption4();
            this.showOption3CreditDisplay();
            this.enableAllRows();
            // Setup the credit pill for Option 3
            this.setupOption3CreditPill();
        } else if (version === 4) {
            this.hideTestModeToggle();
            this.showTestModePillOption4();
            this.setupTestModePillDropdownOption4();
            if (this.isTestMode) {
                this.greyOutRows11To20();
            } else {
                this.enableAllRows();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    enableAllRows() {
        // Remove greyed out styling from all rows
        document.querySelectorAll('.data-row').forEach(row => {
            row.classList.remove('row-greyed-out');
            row.style.pointerEvents = 'auto';
            
            // Remove the tooltip event listener by cloning and replacing the row
            const newRow = row.cloneNode(true);
            row.parentNode.replaceChild(newRow, row);
        });
        
        // Re-enable all inputs
        document.querySelectorAll('.data-row input').forEach(input => {
            input.disabled = false;
            input.readOnly = false;
            input.style.pointerEvents = 'auto';
        });
        
        // Re-enable all cells
        document.querySelectorAll('.data-row td, .data-row th').forEach(cell => {
            cell.style.pointerEvents = 'auto';
        });
        
        // Remove any existing tooltips
        const existingTooltip = document.querySelector('.greyed-out-tooltip');
        if (existingTooltip) {
            existingTooltip.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Re-setup event listeners for the new rows
        this.setupEventListeners();
    }

    greyOutRows11To20() {
        // Grey out rows 11-20
        for (let row = 11; row <= 20; row++) {
            const rowElement = document.querySelector(`tr[data-row="${row}"]`);
            if (rowElement) {
                rowElement.classList.add('row-greyed-out');
                
                // Disable inputs in greyed out rows
                const inputs = rowElement.querySelectorAll('input');
                inputs.forEach(input => {
                    input.disabled = true;
                    input.readOnly = true;
                    input.style.pointerEvents = 'none';
                });
                
                // Disable all cells in the row
                const cells = rowElement.querySelectorAll('td, th');
                cells.forEach(cell => {
                    cell.style.pointerEvents = 'none';
                });
                
                // Add click handler to the entire row with higher priority
                rowElement.style.pointerEvents = 'auto';
                rowElement.addEventListener('click', (e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    this.showGreyedOutTooltip(e);
                }, true); // Use capture phase
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showGreyedOutTooltip(event) {
        // Remove existing tooltip
        const existingTooltip = document.querySelector('.greyed-out-tooltip');
        if (existingTooltip) {
            existingTooltip.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Create tooltip
        const tooltip = document.createElement('div');
        tooltip.className = 'greyed-out-tooltip';
        tooltip.textContent = 'Test 1-10 rows first before accessing additional rows';
        tooltip.style.cssText = `
            position: absolute;
            background: #333;
            color: white;
            padding: 8px 12px;
            border-radius: 6px;
            font-size: 12px;
            z-index: 10000;
            pointer-events: none;
            white-space: nowrap;
            left: ${event.pageX + 10}px;
            top: ${event.pageY - 30}px;
        `;
        
        document.body.appendChild(tooltip);
        
        // Remove tooltip after 3 seconds
        setTimeout(() => {
            if (tooltip.parentNode) {
                tooltip.remove();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        }, 3000);
    }

    showCreditDisplay() {
        const runAllInfo = document.getElementById('run-all-info-option12');
        if (runAllInfo) {
            runAllInfo.style.display = 'flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showOption1CreditDisplay() {
        const runAllInfoOption1 = document.getElementById('run-all-info-option1');
        if (runAllInfoOption1) {
            runAllInfoOption1.style.display = 'flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showOption3CreditDisplay() {
        const runAllInfoOption3 = document.getElementById('run-all-info-option3');
        if (runAllInfoOption3) {
            runAllInfoOption3.style.display = 'flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    hideCreditDisplay() {
        const runAllInfo = document.getElementById('run-all-info-option12');
        const runAllInfoOption3 = document.getElementById('run-all-info-option3');
        if (runAllInfo) {
            runAllInfo.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        if (runAllInfoOption3) {
            runAllInfoOption3.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    setupTestModeToggle() {
        const testModeToggle = document.getElementById('test-mode-toggle');
        if (testModeToggle) {
            testModeToggle.addEventListener('click', () => {
                this.toggleTestMode();
            });
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    toggleTestMode() {
        // If switching from Test Mode to Real Mode, show confirmation
        if (this.isTestMode) {
            this.showRealModeConfirmation();
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // If switching from Real Mode to Test Mode, do it immediately
        this.isTestMode = true;
        localStorage.setItem('testMode', this.isTestMode.toString());
        
        this.updateTestModeUI();
        
        // Apply row restrictions based on test mode
        if (this.currentVersion === 2) {
            this.greyOutRows11To20();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showRealModeConfirmation() {
        // Remove existing confirmation popover
        const existingPopover = document.querySelector('.real-mode-confirmation');
        if (existingPopover) {
            existingPopover.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Create confirmation popover
        const popover = document.createElement('div');
        popover.className = 'real-mode-confirmation';
                       popover.innerHTML = `
                   <div class="confirmation-content">
                       <div class="confirmation-title">Switch to Scale Mode?</div>
                       <div class="confirmation-message">Ending Test Mode will start using credits. You won't be able to switch back.</div>
                       <div class="confirmation-buttons">
                           <button class="btn-cancel" id="cancel-real-mode">Cancel</button>
                           <button class="btn-confirm" id="confirm-real-mode">Confirm</button>
                       </div>
                   </div>
               `;

        // Position the popover near the test mode toggle
        const testModeToggle = document.querySelector('.test-mode-toggle-container');
        if (testModeToggle) {
            const rect = testModeToggle.getBoundingClientRect();
            popover.style.cssText = `
                position: fixed;
                top: ${rect.bottom + 8}px;
                left: ${rect.left}px;
                z-index: 1000;
            `;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }

        // Add event listeners
        popover.querySelector('#cancel-real-mode').addEventListener('click', () => {
            popover.remove();
        });

        popover.querySelector('#confirm-real-mode').addEventListener('click', () => {
            this.confirmRealMode();
            popover.remove();
        });

        // Close popover when clicking outside
        setTimeout(() => {
            document.addEventListener('click', (e) => {
                const testModeToggle = document.querySelector('.test-mode-toggle-container');
                if (!popover.contains(e.target) && !testModeToggle.contains(e.target)) {
                    popover.remove();
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, { once: true });
        }, 100);

        document.body.appendChild(popover);
    }

    confirmRealMode() {
        this.isTestMode = false;
        localStorage.setItem('testMode', this.isTestMode.toString());
        
        // Enable all rows in real mode
        if (this.currentVersion === 2) {
            this.enableAllRows();
            
            // Check if all app columns already have outputs generated
            const allAppCells = document.querySelectorAll('.app-column-cell');
            let allComplete = true;
            let hasAppCells = false;
            
            allAppCells.forEach(cell => {
                hasAppCells = true;
                const outputPill = cell.querySelector('.output-pill');
                if (!outputPill) {
                    allComplete = false;
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            });
            
            // If all app cells have outputs, show "All rows processed" instead of credit system
            if (hasAppCells && allComplete) {
                console.log('🎯 All app columns have outputs! Showing "All rows processed" message...');
                const option2Text = document.getElementById('credit-text-option2');
                if (option2Text) {
                    option2Text.innerHTML = '<strong>All rows processed</strong>';
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                
                // Hide credit cost display when "All rows processed" is shown
                this.hideCreditDisplay();
            } else {
                // Reset column count to 0 when switching to Scale Mode
                this.columnCount = 0;
                this.updateCreditTextForScaleMode();
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Show Option 2's credit system
            const runAllInfoOption2 = document.getElementById('run-all-info-option2');
            if (runAllInfoOption2) {
                runAllInfoOption2.style.display = 'flex';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            
            // Remove the test mode toggle entirely
            const inlineTestMode = document.querySelector('.inline-test-mode');
            if (inlineTestMode) {
                inlineTestMode.style.display = 'none';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    updateTestModeUI() {
        const inlineTestMode = document.querySelector('.inline-test-mode');
        const inlineTestModeOption4 = document.querySelector('.inline-test-mode-option4');
        const toggleSwitch = document.getElementById('test-mode-toggle');
        const modeText = document.getElementById('mode-text');
        const testModePill = document.getElementById('test-mode-pill-option4');
        
        if (this.currentVersion === 2) {
            inlineTestMode.style.display = 'flex';
            
            if (this.isTestMode) {
                toggleSwitch.classList.add('active');
                modeText.textContent = 'Test mode';
            } else {
                toggleSwitch.classList.remove('active');
                modeText.textContent = 'Real mode';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else if (this.currentVersion === 4) {
            inlineTestModeOption4.style.display = 'flex';
            
            if (this.isTestMode) {
                testModePill.innerHTML = '<span class="test-dot"></span> Test mode<i class="fas fa-chevron-down"></i>';
            } else {
                testModePill.innerHTML = '<span class="test-dot"></span> Scale mode<i class="fas fa-chevron-down"></i>';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        } else {
            inlineTestMode.style.display = 'none';
            inlineTestModeOption4.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showTestModeToggle() {
        const inlineTestMode = document.querySelector('.inline-test-mode');
        if (inlineTestMode) {
            inlineTestMode.style.display = 'flex';
            this.updateTestModeUI();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    hideTestModeToggle() {
        const inlineTestMode = document.querySelector('.inline-test-mode');
        if (inlineTestMode) {
            inlineTestMode.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showTestModePill() {
        const inlineTestMode = document.querySelector('.inline-test-mode');
        if (inlineTestMode) {
            inlineTestMode.style.display = 'flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showTestModePillOption4() {
        const inlineTestModeOption4 = document.querySelector('.inline-test-mode-option4');
        if (inlineTestModeOption4) {
            inlineTestModeOption4.style.display = 'flex';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    hideTestModePillOption4() {
        const inlineTestModeOption4 = document.querySelector('.inline-test-mode-option4');
        if (inlineTestModeOption4) {
            inlineTestModeOption4.style.display = 'none';
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    setupTestModePillDropdown() {
        // Only setup dropdown for Option 2
        if (this.currentVersion !== 2) return;
        
        const testModePill = document.getElementById('test-mode-pill');
        if (!testModePill) return;

        // Remove existing event listeners
        testModePill.removeEventListener('click', this.handleTestModePillClick);
        
        // Add click event listener
        testModePill.addEventListener('click', this.handleTestModePillClick.bind(this));
    }

    setupTestModePillDropdownOption4() {
        // Only setup dropdown for Option 4
        if (this.currentVersion !== 4) return;
        
        const testModePill = document.getElementById('test-mode-pill-option4');
        if (!testModePill) return;

        // Remove existing event listeners
        testModePill.removeEventListener('click', this.handleTestModePillClickOption4);
        
        // Add click event listener
        testModePill.addEventListener('click', this.handleTestModePillClickOption4.bind(this));
    }

    handleTestModePillClick(event) {
        event.stopPropagation();
        
        const testModePill = document.getElementById('test-mode-pill');
        const existingDropdown = document.querySelector('.test-mode-dropdown');
        
        // If dropdown is already open, close it
        if (existingDropdown && existingDropdown.classList.contains('show')) {
            this.closeTestModeDropdown();
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Close any existing dropdown
        if (existingDropdown) {
            existingDropdown.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Create dropdown
        const dropdown = document.createElement('div');
        dropdown.className = 'test-mode-dropdown';
        dropdown.innerHTML = `
            <div class="test-mode-dropdown-item" data-mode="test">Test mode</div>
            <div class="test-mode-dropdown-item" data-mode="real">Scale mode</div>
        `;
        
        // Position dropdown
        const rect = testModePill.getBoundingClientRect();
        dropdown.style.position = 'fixed';
        dropdown.style.top = `${rect.bottom + 4}px`;
        dropdown.style.left = `${rect.left}px`;
        
        // Add to body
        document.body.appendChild(dropdown);
        
        // Show dropdown
        setTimeout(() => {
            dropdown.classList.add('show');
            testModePill.classList.add('open');
        }, 10);
        
        // Add click handlers for dropdown items
        dropdown.querySelectorAll('.test-mode-dropdown-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                const mode = item.dataset.mode;
                this.selectTestMode(mode);
                this.closeTestModeDropdown();
            });
        });
        
        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', this.closeTestModeDropdown.bind(this), { once: true });
        }, 100);
    }

    handleTestModePillClickOption4(event) {
        event.stopPropagation();
        
        const testModePill = document.getElementById('test-mode-pill-option4');
        const existingDropdown = document.querySelector('.test-mode-dropdown');
        
        // If dropdown is already open, close it
        if (existingDropdown && existingDropdown.classList.contains('show')) {
            this.closeTestModeDropdown();
            return;
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Close any existing dropdown
        if (existingDropdown) {
            existingDropdown.remove();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        // Create dropdown
        const dropdown = document.createElement('div');
        dropdown.className = 'test-mode-dropdown';
        dropdown.innerHTML = `
            <div class="test-mode-dropdown-item" data-mode="test">Test mode</div>
            <div class="test-mode-dropdown-item" data-mode="real">Scale mode</div>
        `;
        
        // Position dropdown
        const rect = testModePill.getBoundingClientRect();
        dropdown.style.position = 'fixed';
        dropdown.style.top = `${rect.bottom + 4}px`;
        dropdown.style.left = `${rect.left}px`;
        
        // Add to body
        document.body.appendChild(dropdown);
        
        // Show dropdown
        setTimeout(() => {
            dropdown.classList.add('show');
            testModePill.classList.add('open');
        }, 10);
        
        // Add click handlers for dropdown items
        dropdown.querySelectorAll('.test-mode-dropdown-item').forEach(item => {
            item.addEventListener('click', (e) => {
                e.stopPropagation();
                const mode = item.dataset.mode;
                this.selectTestMode(mode);
                this.closeTestModeDropdown();
            });
        });
        
        // Close dropdown when clicking outside
        setTimeout(() => {
            document.addEventListener('click', this.closeTestModeDropdown.bind(this), { once: true });
        }, 100);
    }

    closeTestModeDropdown() {
        const dropdown = document.querySelector('.test-mode-dropdown');
        const testModePill = document.getElementById('test-mode-pill');
        
        if (dropdown) {
            dropdown.classList.remove('show');
            setTimeout(() => {
                if (dropdown.parentNode) {
                    dropdown.parentNode.removeChild(dropdown);
                } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            }, 200);
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
        
        if (testModePill) {
            testModePill.classList.remove('open');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    selectTestMode(mode) {
        if (mode === 'test') {
            this.isTestMode = true;
            this.greyOutRows11To20();
            // Update pill text to show current mode
            const testModePill = document.getElementById('test-mode-pill');
            if (testModePill) {
                testModePill.innerHTML = '<span class="test-dot"></span> Test mode<i class="fas fa-chevron-down"></i>';
            } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
            // Save to localStorage
            localStorage.setItem('testMode', this.isTestMode.toString());
        } else if (mode === 'real') {
            // Show confirmation popover
            this.showRealModeConfirmation();
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }

    showTestModeToastOnLoad() {
        // Only show toast for Option 2 and when in test mode
        if (this.currentVersion === 2 && this.isTestMode) {
            // Create toast element using the same pattern as showCreditToast
            const toast = document.createElement('div');
            toast.className = 'credit-toast';
            toast.innerHTML = `
                <div class="toast-title">
                    <i class="fas fa-eye"></i>
                    Test Mode
                </div>
                <div class="toast-body">You're in Test mode. Explore and test outputs without using credits. When you're ready, switch off Test mode to start scaling.</div>
            `;
            
            // Add to body
            document.body.appendChild(toast);
            
            // Show toast with slide-up animation
            setTimeout(() => {
                toast.classList.add('show');
            }, 100);
            
            // Hide toast after 9 seconds
            setTimeout(() => {
                toast.classList.remove('show');
                setTimeout(() => {
                    if (toast.parentNode) {
                        toast.parentNode.removeChild(toast);
                    } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
                }, 300);
            }, 9000);
            
            console.log('👁️ Showing test mode toast notification');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    }
}

// Initialize the spreadsheet when the page loads
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheet = new Spreadsheet();
    
    // Add debug function to window for testing
    window.debugImportButton = () => {
        console.log('🐛 Debug: Import All button state');
        const btn = document.getElementById('import-all-btn');
        if (btn) {
            console.log('Button found:', btn);
            console.log('Disabled:', btn.disabled);
            console.log('Style:', btn.style.cssText);
            console.log('Classes:', btn.className);
            console.log('Event listeners:', btn.onclick);
            
            // Force enable
            console.log('Force enabling...');
            window.spreadsheet.enableImportAllButton();
            
            // Test click
            console.log('Testing click...');
            btn.click();
        } else {
            console.log('Button not found!');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    };
    
    // Add direct import test function
    window.testImportDirectly = () => {
        console.log('🧪 Testing import function directly');
        if (window.spreadsheet) {
            window.spreadsheet.importAllData();
        } else {
            console.log('Spreadsheet not found!');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    };
    
    // Add button state checker
    window.checkButtonState = () => {
        const btn = document.getElementById('import-all-btn');
        if (btn) {
            console.log('📋 Button State:');
            console.log('- Exists:', true);
            console.log('- Disabled:', btn.disabled);
            console.log('- Opacity:', btn.style.opacity);
            console.log('- Cursor:', btn.style.cursor);
            console.log('- Display:', getComputedStyle(btn).display);
            console.log('- Visibility:', getComputedStyle(btn).visibility);
            
            // Check if test mode container is visible
            const testContainer = document.getElementById('test-mode-container');
            console.log('- Test container visible:', testContainer ? testContainer.style.display !== 'none' : false);
        } else {
            console.log('❌ Import All button not found');
        } else if (version === 5) {
            // Option 5: Duplicate of Option 2 - always starts in test mode
            this.isTestMode = true;
            localStorage.setItem("testMode", "true");
            this.showTestModeToggle();
            this.setupTestModeToggle();
            this.hideTestModePillOption4();
            this.greyOutRows11To20();
            this.showTestModeToastOnLoad(); // Show toast when switching to Option 5
        }
    };
    
    // Sheet starts empty - no sample data
});
