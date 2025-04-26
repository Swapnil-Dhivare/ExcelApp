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

// Function to update format button states
function updateFormatButtonStates() {
    // This would update the active state of formatting buttons
    // based on the currently selected cell(s)
    console.log('Updating format button states');
}

// Initialize the buttons when the DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
    console.log('Initializing Excel buttons');
    
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
    
    // Add row and column buttons
    const addRowBtn = document.getElementById('addRowBtn');
    if (addRowBtn) {
        addRowBtn.addEventListener('click', function() {
            if (typeof addNewRow === 'function') {
                addNewRow();
            } else {
                console.error('addNewRow function not found');
            }
        });
    }
    
    const addColumnBtn = document.getElementById('addColumnBtn');
    if (addColumnBtn) {
        addColumnBtn.addEventListener('click', function() {
            if (typeof addNewColumn === 'function') {
                addNewColumn();
            } else {
                console.error('addNewColumn function not found');
            }
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
    
    console.log('Excel buttons initialized');
});