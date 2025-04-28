// Function to apply text alignment
function applyAlignment(alignment) {
    console.log('Applying alignment:', alignment);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.textAlign = alignment;
        }
    });
    
    saveSheetData();
    showToast(`Text alignment set to: ${alignment}`, 'success');
}

// Function to apply header background color
function applyHeaderColor(color) {
    console.log('Applying header color:', color);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one header cell', 'error');
        return;
    }
    
    let headerCount = 0;
    selectedCells.forEach(cell => {
        if (cell.tagName === 'TH' || cell.classList.contains('sheet-header')) {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
                headerCount++;
            }
        }
    });
    
    if (headerCount > 0) {
        saveSheetData();
        showToast(`Applied color to ${headerCount} header(s)`, 'success');
    } else {
        showToast('Please select a header cell', 'error');
    }
}

// Function to apply font size to selected cells
function applyFontSize(fontSize) {
    console.log('Applying font size:', fontSize);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.fontSize = `${fontSize}px`;
        }
    });
    
    saveSheetData();
    showToast(`Font size set to: ${fontSize}px`, 'success');
}

// Function to apply font family to selected cells
function applyFontFamily(fontFamily) {
    console.log('Applying font family:', fontFamily);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.fontFamily = fontFamily;
        }
    });
    
    saveSheetData();
    showToast(`Font family set to: ${fontFamily.split(',')[0].replace(/['"]/g, '')}`, 'success');
}

// Function to apply text color to selected cells
function applyTextColor(color) {
    console.log('Applying text color:', color);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.color = color;
        }
    });
    
    saveSheetData();
    showToast('Text color applied', 'success');
}

// Function to apply background color to selected cells
function applyBackgroundColor(color) {
    console.log('Applying background color:', color);
    
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.backgroundColor = color;
        }
    });
    
    saveSheetData();
    showToast('Background color applied', 'success');
}

// Function to apply bold formatting to selected cells
function applyBoldFormatting() {
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            const currentWeight = window.getComputedStyle(editableDiv).fontWeight;
            editableDiv.style.fontWeight = (currentWeight === '700' || currentWeight === 'bold') ? 'normal' : 'bold';
        }
    });
    
    saveSheetData();
    showToast('Bold formatting toggled', 'success');
}

// Function to apply italic formatting to selected cells
function applyItalicFormatting() {
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            const currentStyle = window.getComputedStyle(editableDiv).fontStyle;
            editableDiv.style.fontStyle = (currentStyle === 'italic') ? 'normal' : 'italic';
        }
    });
    
    saveSheetData();
    showToast('Italic formatting toggled', 'success');
}

// Function to merge selected cells
function mergeCells() {
    console.log('Merging cells:', selectedCells);
    
    if (!selectedCells || selectedCells.length < 2) {
        showToast('Please select at least 2 cells to merge', 'error');
        return;
    }
    
    // Find the bounds of selection
    let minRow = Infinity;
    let maxRow = -1;
    let minCol = Infinity;
    let maxCol = -1;
    
    // Check if any header is selected
    const headerSelected = selectedCells.some(cell => 
        cell.tagName === 'TH' || cell.classList.contains('sheet-header')
    );
    
    if (headerSelected) {
        showToast('Cannot merge header cells', 'error');
        return;
    }
    
    // Get the cell boundaries
    selectedCells.forEach(cell => {
        const rowIndex = parseInt(cell.getAttribute('data-row'));
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        if (!isNaN(rowIndex) && !isNaN(colIndex)) {
            minRow = Math.min(minRow, rowIndex);
            maxRow = Math.max(maxRow, rowIndex);
            minCol = Math.min(minCol, colIndex);
            maxCol = Math.max(maxCol, colIndex);
        }
    });
    
    // Get the main cell (top-left)
    const mainCell = document.querySelector(`td[data-row="${minRow}"][data-col="${minCol}"]`);
    if (!mainCell) {
        showToast('Could not find the top-left cell for merging', 'error');
        return;
    }
    
    // Get content from all cells
    let content = '';
    selectedCells.forEach(cell => {
        const cellContent = cell.querySelector('.editable-cell')?.textContent || '';
        if (cellContent.trim()) {
            content += (content ? ' ' : '') + cellContent;
        }
    });
    
    // Apply to main cell
    const mainCellDiv = mainCell.querySelector('.editable-cell');
    if (mainCellDiv) {
        mainCellDiv.textContent = content;
    }
    
    // Setup merge attributes
    mainCell.setAttribute('data-rowspan', maxRow - minRow + 1);
    mainCell.setAttribute('data-colspan', maxCol - minCol + 1);
    mainCell.classList.add('merged-cell');
    
    // Size calculation
    const heightPx = (maxRow - minRow + 1) * 25;
    const widthPx = (maxCol - minCol + 1) * 80;
    
    // Apply style to main cell
    mainCell.style.height = `${heightPx}px`;
    mainCell.style.width = `${widthPx}px`;
    mainCell.style.position = 'relative';
    mainCell.style.zIndex = '5';
    
    // Hide other cells
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            if (r === minRow && c === minCol) continue;
            
            const cell = document.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
            if (cell) {
                cell.style.display = 'none';
                cell.setAttribute('data-hidden-by-merge', `${minRow},${minCol}`);
            }
        }
    }
    
    // Clear selection
    clearCellSelection();
    
    showToast('Cells merged successfully', 'success');
    saveSheetData();
}

// Function for cell division
function promptForCellDivision() {
    if (selectedCells.length !== 1) {
        showToast('Please select exactly one cell to divide', 'error');
        return;
    }
    
    const cell = selectedCells[0];
    
    // Create dialog HTML
    const dialogHTML = `
        <div class="cell-division-dialog" style="background: white; padding: 20px; border-radius: 8px; box-shadow: 0 0 10px rgba(0,0,0,0.2); max-width: 400px;">
            <h4>Divide Cell</h4>
            <div style="margin-bottom: 15px;">
                <label>Division Type:</label>
                <select id="divisionType" class="form-select">
                    <option value="horizontal">Horizontal (1:n)</option>
                    <option value="vertical">Vertical (n:1)</option>
                    <option value="grid">Grid (n:m)</option>
                </select>
            </div>
            
            <div id="horizontalSections">
                <label>Number of horizontal sections:</label>
                <input type="number" id="horizontalCount" class="form-control" min="2" value="2">
            </div>
            
            <div id="verticalSections" style="display:none; margin-top: 15px;">
                <label>Number of vertical sections:</label>
                <input type="number" id="verticalCount" class="form-control" min="2" value="2">
            </div>
            
            <div style="margin-top: 20px; text-align: right;">
                <button id="cancelDivide" class="btn btn-secondary">Cancel</button>
                <button id="confirmDivide" class="btn btn-primary">Divide</button>
            </div>
        </div>
    `;
    
    // Create overlay
    const overlay = document.createElement('div');
    overlay.style.position = 'fixed';
    overlay.style.top = '0';
    overlay.style.left = '0';
    overlay.style.width = '100%';
    overlay.style.height = '100%';
    overlay.style.backgroundColor = 'rgba(0,0,0,0.5)';
    overlay.style.zIndex = '9999';
    overlay.style.display = 'flex';
    overlay.style.justifyContent = 'center';
    overlay.style.alignItems = 'center';
    overlay.innerHTML = dialogHTML;
    
    document.body.appendChild(overlay);
    
    // Setup event handlers
    const divisionType = document.getElementById('divisionType');
    const horizontalSections = document.getElementById('horizontalSections');
    const verticalSections = document.getElementById('verticalSections');
    
    divisionType.addEventListener('change', function() {
        if (this.value === 'horizontal') {
            horizontalSections.style.display = 'block';
            verticalSections.style.display = 'none';
        } else if (this.value === 'vertical') {
            horizontalSections.style.display = 'none';
            verticalSections.style.display = 'block';
        } else {
            horizontalSections.style.display = 'block';
            verticalSections.style.display = 'block';
        }
    });
    
    document.getElementById('cancelDivide').addEventListener('click', function() {
        document.body.removeChild(overlay);
    });
    
    document.getElementById('confirmDivide').addEventListener('click', function() {
        const type = document.getElementById('divisionType').value;
        const hCount = parseInt(document.getElementById('horizontalCount').value) || 2;
        const vCount = parseInt(document.getElementById('verticalCount').value) || 2;
        
        document.body.removeChild(overlay);
        
        // Perform cell division
        divideCell(cell, type, hCount, vCount);
    });
}

// Function to toggle cell selection
function toggleCellSelection(cell, multiSelect = false, event = null) {
    // Check if Ctrl key is pressed
    const ctrlPressed = (event && (event.ctrlKey || event.metaKey)) || multiSelect;
    
    console.log('Selection with ctrl:', ctrlPressed);
    
    if (ctrlPressed) {
        const index = selectedCells.indexOf(cell);
        
        if (index !== -1) {
            // Deselect
            cell.classList.remove('selected');
            selectedCells.splice(index, 1);
        } else {
            // Add to selection
            cell.classList.add('selected');
            selectedCells.push(cell);
        }
    } else {
        // Single selection
        clearCellSelection();
        cell.classList.add('selected');
        selectedCells.push(cell);
    }
    
    // Update UI state
    updateFormatButtonStates();
}

// Function to clear cell selection
function clearCellSelection() {
    if (!selectedCells) return;
    
    selectedCells.forEach(cell => {
        if (cell) cell.classList.remove('selected');
    });
    
    selectedCells = [];
}

// Function to update format button states based on current selection
function updateFormatButtonStates() {
    if (!selectedCells || selectedCells.length === 0) return;
    
    // Get the first selected cell for reference
    const referenceCell = selectedCells[0].querySelector('.editable-cell');
    if (!referenceCell) return;
    
    const computedStyle = window.getComputedStyle(referenceCell);
    
    // Update font size select if it exists
    const fontSizeSelect = document.getElementById('fontSizeSelect');
    if (fontSizeSelect) {
        const fontSize = parseInt(computedStyle.fontSize);
        // Find the closest option or default to current value
        const options = Array.from(fontSizeSelect.options).map(opt => parseInt(opt.value));
        const closestSize = options.reduce((prev, curr) => 
            Math.abs(curr - fontSize) < Math.abs(prev - fontSize) ? curr : prev, options[0]);
        fontSizeSelect.value = closestSize;
    }
    
    // Update font family select if it exists
    const fontFamilySelect = document.getElementById('fontFamilySelect');
    if (fontFamilySelect) {
        const fontFamily = computedStyle.fontFamily;
        // Try to find matching option
        for (const option of fontFamilySelect.options) {
            if (fontFamily.includes(option.text)) {
                fontFamilySelect.value = option.value;
                break;
            }
        }
    }
    
    // Update text alignment buttons
    const textAlign = computedStyle.textAlign;
    document.querySelectorAll('#alignLeftBtn, #alignCenterBtn, #alignRightBtn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    if (textAlign === 'left') {
        document.getElementById('alignLeftBtn')?.classList.add('active');
    } else if (textAlign === 'center') {
        document.getElementById('alignCenterBtn')?.classList.add('active');
    } else if (textAlign === 'right') {
        document.getElementById('alignRightBtn')?.classList.add('active');
    }
    
    // Update bold button
    const boldBtn = document.getElementById('boldBtn');
    if (boldBtn) {
        if (computedStyle.fontWeight === '700' || computedStyle.fontWeight === 'bold') {
            boldBtn.classList.add('active');
        } else {
            boldBtn.classList.remove('active');
        }
    }
    
    // Update italic button
    const italicBtn = document.getElementById('italicBtn');
    if (italicBtn) {
        if (computedStyle.fontStyle === 'italic') {
            italicBtn.classList.add('active');
        } else {
            italicBtn.classList.remove('active');
        }
    }
}

// Initialize the buttons when the DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    console.log('Initializing Excel buttons');
    
    // Initialize selectedCells array if it doesn't exist
    window.selectedCells = window.selectedCells || [];
    
    // Alignment buttons
    const alignLeftBtn = document.getElementById('alignLeftBtn');
    const alignCenterBtn = document.getElementById('alignCenterBtn');
    const alignRightBtn = document.getElementById('alignRightBtn');
    
    if (alignLeftBtn) {
        alignLeftBtn.addEventListener('click', function() {
            applyAlignment('left');
        });
    }
    
    if (alignCenterBtn) {
        alignCenterBtn.addEventListener('click', function() {
            applyAlignment('center');
        });
    }
    
    if (alignRightBtn) {
        alignRightBtn.addEventListener('click', function() {
            applyAlignment('right');
        });
    }
    
    // Font size select
    const fontSizeSelect = document.getElementById('fontSizeSelect');
    if (fontSizeSelect) {
        fontSizeSelect.addEventListener('change', function() {
            applyFontSize(this.value);
        });
    }
    
    // Font family select
    const fontFamilySelect = document.getElementById('fontFamilySelect');
    if (fontFamilySelect) {
        fontFamilySelect.addEventListener('change', function() {
            applyFontFamily(this.value);
        });
    }
    
    // Bold button
    const boldBtn = document.getElementById('boldBtn');
    if (boldBtn) {
        boldBtn.addEventListener('click', function() {
            applyBoldFormatting();
        });
    }
    
    // Italic button
    const italicBtn = document.getElementById('italicBtn');
    if (italicBtn) {
        italicBtn.addEventListener('click', function() {
            applyItalicFormatting();
        });
    }
    
    // Text color button and picker
    const textColorBtn = document.getElementById('textColorBtn');
    const textColorPicker = document.getElementById('textColorPicker');
    
    if (textColorBtn) {
        textColorBtn.addEventListener('click', function() {
            if (textColorPicker) {
                textColorPicker.click();
            } else {
                // Create color picker if it doesn't exist
                const picker = document.createElement('input');
                picker.type = 'color';
                picker.id = 'textColorPicker';
                picker.style.display = 'none';
                document.body.appendChild(picker);
                
                picker.addEventListener('input', function() {
                    applyTextColor(this.value);
                });
                
                picker.click();
            }
        });
    }
    
    if (textColorPicker) {
        textColorPicker.addEventListener('input', function() {
            applyTextColor(this.value);
        });
    }
    
    // Fill color button and picker
    const fillColorBtn = document.getElementById('fillColorBtn');
    const fillColorPicker = document.getElementById('fillColorPicker');
    
    if (fillColorBtn) {
        fillColorBtn.addEventListener('click', function() {
            if (fillColorPicker) {
                fillColorPicker.click();
            } else {
                // Create color picker if it doesn't exist
                const picker = document.createElement('input');
                picker.type = 'color';
                picker.id = 'fillColorPicker';
                picker.style.display = 'none';
                document.body.appendChild(picker);
                
                picker.addEventListener('input', function() {
                    applyBackgroundColor(this.value);
                });
                
                picker.click();
            }
        });
    }
    
    if (fillColorPicker) {
        fillColorPicker.addEventListener('input', function() {
            applyBackgroundColor(this.value);
        });
    }
    
    // Header color button
    const headerColorBtn = document.getElementById('headerColorBtn');
    const headerColorPicker = document.getElementById('headerColorPicker');
    
    if (headerColorBtn && !headerColorPicker) {
        const picker = document.createElement('input');
        picker.type = 'color';
        picker.id = 'headerColorPicker';
        picker.style.display = 'none';
        document.body.appendChild(picker);
        
        headerColorBtn.addEventListener('click', function() {
            document.getElementById('headerColorPicker').click();
        });
        
        picker.addEventListener('change', function() {
            applyHeaderColor(this.value);
        });
    }
    
    // Add row and column buttons
    const addRowBtn = document.getElementById('addRowBtn');
    const addColumnBtn = document.getElementById('addColumnBtn');
    const deleteRowBtn = document.getElementById('deleteRowBtn');
    const deleteColumnBtn = document.getElementById('deleteColumnBtn');
    const saveBtn = document.getElementById('saveSheetBtn');
    
    if (addRowBtn) addRowBtn.addEventListener('click', addRow);
    if (addColumnBtn) addColumnBtn.addEventListener('click', addColumn);
    if (deleteRowBtn) deleteRowBtn.addEventListener('click', deleteSelectedRow);
    if (deleteColumnBtn) deleteColumnBtn.addEventListener('click', deleteSelectedColumn);
    if (saveBtn) saveBtn.addEventListener('click', saveSheetData);
    
    // Merge cells button
    const mergeCellsBtn = document.getElementById('mergeCellsBtn');
    if (mergeCellsBtn) {
        mergeCellsBtn.addEventListener('click', function() {
            mergeCells();
        });
    }
    
    // Direct merge button
    const directMergeBtn = document.getElementById('directMergeBtn');
    if (directMergeBtn) {
        directMergeBtn.addEventListener('click', function() {
            mergeCells();
        });
    }
    
    // Divide cell button
    const divideCellBtn = document.getElementById('divideCellBtn');
    if (divideCellBtn) {
        divideCellBtn.addEventListener('click', function() {
            promptForCellDivision();
        });
    }
    
    // Add keyboard shortcuts
    document.addEventListener('keydown', function(e) {
        // Ctrl+S to save
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            saveSheetData();
        }
        
        // Ctrl+I to insert row
        if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
            e.preventDefault();
            addRow();
        }
        
        // Ctrl+Shift+I to insert column
        if ((e.ctrlKey || e.metaKey) && e.shiftKey && e.key === 'I') {
            e.preventDefault();
            addColumn();
        }
    });
    
    // Set up cell selection on sheet cells
    const setupCellSelectionHandlers = () => {
        document.querySelectorAll('.sheet-cell').forEach(cell => {
            cell.addEventListener('click', function(event) {
                toggleCellSelection(this, false, event);
            });
        });
    };
    
    // Run setup after a short delay to ensure DOM is fully loaded
    setTimeout(setupCellSelectionHandlers, 500);
    
    console.log('Excel buttons initialized');
});

/**
 * Add a new row to the table
 */
function addRow() {
    const table = document.getElementById('sheetTable');
    if (!table) return;
    
    const tbody = table.querySelector('tbody');
    if (!tbody) return;
    
    const rows = tbody.querySelectorAll('tr');
    const lastRowIndex = rows.length > 0 ? parseInt(rows[rows.length - 1].getAttribute('data-row')) : -1;
    const newRowIndex = lastRowIndex + 1;
    
    // Get number of columns from header
    const headerCells = table.querySelectorAll('thead th');
    const numColumns = headerCells.length;
    
    // Create new row
    const newRow = document.createElement('tr');
    newRow.setAttribute('data-row', newRowIndex);
    
    // Create cells for the row
    for (let col = 0; col < numColumns; col++) {
        const cell = document.createElement('td');
        cell.className = 'sheet-cell';
        cell.setAttribute('data-row', newRowIndex);
        cell.setAttribute('data-col', col);
        
        const editableDiv = document.createElement('div');
        editableDiv.className = 'editable-cell';
        editableDiv.contentEditable = true;
        
        // Add event listeners to make editing work
        editableDiv.addEventListener('click', function(e) {
            e.stopPropagation();
            this.focus();
            
            // Update selection and formula bar
            const parentCell = this.closest('.sheet-cell');
            if (parentCell) {
                clearSelection();
                selectCell(parentCell);
                updateSelectionInfo();
                
                const formulaBar = document.getElementById('formulaBar');
                if (formulaBar) {
                    formulaBar.value = this.textContent;
                }
            }
        });
        
        editableDiv.addEventListener('input', function() {
            const formulaBar = document.getElementById('formulaBar');
            if (formulaBar) {
                formulaBar.value = this.textContent;
            }
        });
        
        cell.appendChild(editableDiv);
        
        // Add cell selection event handler
        cell.addEventListener('mousedown', function(e) {
            if (e.target === this) {
                if (!(e.ctrlKey || e.metaKey)) {
                    clearSelection();
                }
                selectCell(this);
                updateSelectionInfo();
            }
        });
        
        newRow.appendChild(cell);
    }
    
    // Append the row
    tbody.appendChild(newRow);
    
    // Show notification
    showToast('Row added successfully', 'success');
}

/**
 * Add a new column to the table
 */
function addColumn() {
    const table = document.getElementById('sheetTable');
    if (!table) return;
    
    // Get header row and all data rows
    const headerRow = table.querySelector('thead tr');
    const dataRows = table.querySelectorAll('tbody tr');
    
    if (!headerRow) return;
    
    const headerCells = headerRow.querySelectorAll('th');
    const newColumnIndex = headerCells.length;
    
    // Create new header cell
    const newHeaderCell = document.createElement('th');
    newHeaderCell.className = 'sheet-header';
    newHeaderCell.setAttribute('data-col', newColumnIndex);
    newHeaderCell.style.width = '80px';
    
    // Create editable div for header
    const headerEditableDiv = document.createElement('div');
    headerEditableDiv.className = 'editable-cell';
    headerEditableDiv.contentEditable = true;
    
    // Add event listeners for header editing
    headerEditableDiv.addEventListener('click', function(e) {
        e.stopPropagation();
        this.focus();
        
        const parentCell = this.closest('.sheet-header');
        if (parentCell) {
            clearSelection();
            selectCell(parentCell);
            updateSelectionInfo();
            
            const formulaBar = document.getElementById('formulaBar');
            if (formulaBar) {
                formulaBar.value = this.textContent;
            }
        }
    });
    
    headerEditableDiv.addEventListener('input', function() {
        const formulaBar = document.getElementById('formulaBar');
        if (formulaBar) {
            formulaBar.value = this.textContent;
        }
    });
    
    newHeaderCell.appendChild(headerEditableDiv);
    
    // Add header cell selection handler
    newHeaderCell.addEventListener('mousedown', function(e) {
        if (e.target === this) {
            if (!(e.ctrlKey || e.metaKey)) {
                clearSelection();
            }
            selectCell(this);
            updateSelectionInfo();
        }
    });
    
    // Add the header cell
    headerRow.appendChild(newHeaderCell);
    
    // Add new column letter in column headers container
    const columnHeadersContainer = document.getElementById('columnHeaders');
    if (columnHeadersContainer) {
        const colHeader = document.createElement('div');
        colHeader.className = 'excel-column-header';
        colHeader.setAttribute('data-col', newColumnIndex);
        colHeader.textContent = columnIndexToLetter(newColumnIndex);
        
        // Create resize handle
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'col-resize-handle';
        resizeHandle.setAttribute('data-col', newColumnIndex);
        colHeader.appendChild(resizeHandle);
        
        columnHeadersContainer.appendChild(colHeader);
    }
    
    // Add cell to each row
    dataRows.forEach(row => {
        const rowIndex = parseInt(row.getAttribute('data-row'));
        
        // Create new cell
        const newCell = document.createElement('td');
        newCell.className = 'sheet-cell';
        newCell.setAttribute('data-row', rowIndex);
        newCell.setAttribute('data-col', newColumnIndex);
        
        // Create editable div
        const editableDiv = document.createElement('div');
        editableDiv.className = 'editable-cell';
        editableDiv.contentEditable = true;
        
        // Add event listeners
        editableDiv.addEventListener('click', function(e) {
            e.stopPropagation();
            this.focus();
            
            const parentCell = this.closest('.sheet-cell');
            if (parentCell) {
                clearSelection();
                selectCell(parentCell);
                updateSelectionInfo();
                
                const formulaBar = document.getElementById('formulaBar');
                if (formulaBar) {
                    formulaBar.value = this.textContent;
                }
            }
        });
        
        editableDiv.addEventListener('input', function() {
            const formulaBar = document.getElementById('formulaBar');
            if (formulaBar) {
                formulaBar.value = this.textContent;
            }
        });
        
        newCell.appendChild(editableDiv);
        
        // Add cell selection
        newCell.addEventListener('mousedown', function(e) {
            if (e.target === this) {
                if (!(e.ctrlKey || e.metaKey)) {
                    clearSelection();
                }
                selectCell(this);
                updateSelectionInfo();
            }
        });
        
        row.appendChild(newCell);
    });
    
    // Show notification
    showToast('Column added successfully', 'success');
}

/**
 * Delete selected row(s)
 */
function deleteSelectedRow() {
    if (window.selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    // Get all unique row indexes
    const rowIndices = [...new Set(
        window.selectedCells.map(cell => 
            parseInt(cell.getAttribute('data-row'))
        ).filter(row => !isNaN(row))
    )].sort((a, b) => b - a); // Sort in descending order to delete from bottom up
    
    if (rowIndices.length === 0) {
        showToast('No valid rows selected', 'error');
        return;
    }
    
    // Confirm deletion
    if (!confirm(`Are you sure you want to delete ${rowIndices.length > 1 ? 'these rows' : 'this row'}?`)) {
        return;
    }
    
    // Delete rows
    const table = document.getElementById('sheetTable');
    const tbody = table.querySelector('tbody');
    
    rowIndices.forEach(rowIndex => {
        const row = tbody.querySelector(`tr[data-row="${rowIndex}"]`);
        if (row) {
            row.remove();
        }
    });
    
    // Renumber remaining rows
    const remainingRows = tbody.querySelectorAll('tr');
    remainingRows.forEach((row, index) => {
        row.setAttribute('data-row', index);
        
        // Update row cells
        row.querySelectorAll('td').forEach(cell => {
            cell.setAttribute('data-row', index);
        });
    });
    
    // Clear selection
    clearSelection();
    showToast('Row(s) deleted successfully', 'success');
}

/**
 * Delete selected column(s)
 */
function deleteSelectedColumn() {
    if (window.selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    // Get all unique column indexes
    const colIndices = [...new Set(
        window.selectedCells.map(cell => 
            parseInt(cell.getAttribute('data-col'))
        ).filter(col => !isNaN(col))
    )].sort((a, b) => b - a); // Sort in descending order to delete from right to left
    
    if (colIndices.length === 0) {
        showToast('No valid columns selected', 'error');
        return;
    }
    
    // Confirm deletion
    if (!confirm(`Are you sure you want to delete ${colIndices.length > 1 ? 'these columns' : 'this column'}?`)) {
        return;
    }
    
    const table = document.getElementById('sheetTable');
    const thead = table.querySelector('thead');
    const columnHeaders = document.getElementById('columnHeaders');
    
    colIndices.forEach(colIndex => {
        // Remove column header
        const th = thead.querySelector(`th[data-col="${colIndex}"]`);
        if (th) th.remove();
        
        // Remove column from column headers container
        if (columnHeaders) {
            const colHeader = columnHeaders.querySelector(`.excel-column-header[data-col="${colIndex}"]`);
            if (colHeader) colHeader.remove();
        }
        
        // Remove column cells
        table.querySelectorAll(`td[data-col="${colIndex}"]`).forEach(cell => {
            cell.remove();
        });
    });
    
    // Renumber remaining columns
    const headCells = thead.querySelectorAll('th');
    headCells.forEach((cell, index) => {
        cell.setAttribute('data-col', index);
    });
    
    // Update column header labels
    if (columnHeaders) {
        const colHeaders = columnHeaders.querySelectorAll('.excel-column-header');
        colHeaders.forEach((header, index) => {
            header.setAttribute('data-col', index);
            header.textContent = columnIndexToLetter(index);
            
            // Update resize handle
            const resizeHandle = header.querySelector('.col-resize-handle');
            if (resizeHandle) {
                resizeHandle.setAttribute('data-col', index);
            }
        });
    }
    
    // Update all data cells
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        cells.forEach((cell, index) => {
            cell.setAttribute('data-col', index);
        });
    });
    
    // Clear selection
    clearSelection();
    showToast('Column(s) deleted successfully', 'success');
}

/**
 * Convert column index to letter (A, B, C, ... Z, AA, AB, etc.)
 */
function columnIndexToLetter(index) {
    let letter = '';
    index++;
    
    while (index > 0) {
        const modulo = (index - 1) % 26;
        letter = String.fromCharCode(65 + modulo) + letter;
        index = Math.floor((index - modulo) / 26);
    }
    
    return letter;
}

/**
 * Save sheet data
 */
function saveSheetData() {
    const sheetName = document.getElementById('sheetTable').getAttribute('data-sheet-name');
    const rows = document.querySelectorAll('#sheetTable tbody tr');
    const headers = document.querySelectorAll('#sheetTable thead th');
    
    // Create data array
    const data = [];
    
    // Add headers
    const headerRow = [];
    headers.forEach(header => {
        const editableDiv = header.querySelector('.editable-cell');
        headerRow.push(editableDiv ? editableDiv.textContent : '');
    });
    data.push(headerRow);
    
    // Add data rows
    rows.forEach(row => {
        const dataRow = [];
        const cells = row.querySelectorAll('td');
        cells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            dataRow.push(editableDiv ? editableDiv.textContent : '');
        });
        data.push(dataRow);
    });
    
    // Send data to server
    fetch(`/update_sheet/${sheetName}`, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': document.querySelector('meta[name="csrf-token"]').getAttribute('content')
        },
        body: JSON.stringify({ data: data })
    })
    .then(response => response.json())
    .then(data => {
        if (data.success) {
            showToast('Sheet saved successfully', 'success');
        } else {
            showToast('Error saving sheet: ' + data.message, 'error');
        }
    })
    .catch(error => {
        showToast('Error saving sheet', 'error');
        console.error('Error:', error);
    });
}

// Helper functions shared with other scripts
window.saveSheetData = saveSheetData;
window.showToast = showToast;
window.clearSelection = clearSelection;
window.selectCell = selectCell;
window.updateSelectionInfo = updateSelectionInfo;

/**
 * Show toast notification
 */
function showToast(message, type = 'info') {
    const toastContainer = document.querySelector('.toast-container');
    if (!toastContainer) return;
    
    const toast = document.createElement('div');
    toast.className = `toast show mb-2`;
    toast.role = 'alert';
    toast.setAttribute('aria-live', 'assertive');
    toast.setAttribute('aria-atomic', 'true');
    
    const bgClass = type === 'error' ? 'bg-danger' : 
                  type === 'success' ? 'bg-success' : 'bg-info';
    
    toast.innerHTML = `
        <div class="toast-header ${bgClass} text-white">
            <strong class="me-auto">Excel Generator</strong>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="toast" aria-label="Close"></button>
        </div>
        <div class="toast-body bg-light">
            ${message}
        </div>
    `;
    
    toastContainer.appendChild(toast);
    
    toast.querySelector('.btn-close').addEventListener('click', function() {
        toast.remove();
    });
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

/**
 * Select a cell for operations
 */
function selectCell(cell) {
    if (!cell.classList.contains('selected')) {
        cell.classList.add('selected');
        if (!window.selectedCells) window.selectedCells = [];
        window.selectedCells.push(cell);
    }
}

/**
 * Clear all selected cells
 */
function clearSelection() {
    document.querySelectorAll('.selected').forEach(selected => {
        selected.classList.remove('selected');
    });
    window.selectedCells = [];
}

/**
 * Update selection info display
 */
function updateSelectionInfo() {
    if (!window.selectedCells || window.selectedCells.length === 0) return;
    
    const cell = window.selectedCells[0];
    const row = cell.getAttribute('data-row');
    const col = cell.getAttribute('data-col');
    
    if (row !== null && col !== null) {
        const colLetter = columnIndexToLetter(parseInt(col));
        const cellAddr = `${colLetter}${parseInt(row) + 1}`;
        const cellReference = document.getElementById('cellReference');
        const selectedCellInfo = document.getElementById('selectedCellInfo');
        
        if (cellReference) cellReference.value = cellAddr;
        if (selectedCellInfo) selectedCellInfo.textContent = cellAddr;
    }
}