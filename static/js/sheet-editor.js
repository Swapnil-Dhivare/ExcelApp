// Excel-like Sheet Editor
// Global variables
let selectedCells = [];
let sheetTable = null;
let currentSheetName = '';
let columnHeadersContainer = null;

// Initialize on document ready
document.addEventListener('DOMContentLoaded', function() {
    // Expose selectedCells to window scope so other scripts can access it
    window.selectedCells = selectedCells;
    
    // Make sure buttons are connected to their functions
    document.getElementById('addRowBtn')?.addEventListener('click', addNewRow);
    document.getElementById('addColumnBtn')?.addEventListener('click', addNewColumn);
    document.getElementById('deleteRowBtn')?.addEventListener('click', deleteSelectedRow);
    document.getElementById('deleteColumnBtn')?.addEventListener('click', deleteSelectedColumn);
    document.getElementById('mergeCellsBtn')?.addEventListener('click', mergeCells);
    document.getElementById('divideCellBtn')?.addEventListener('click', divideCellInPlace);
    document.getElementById('columnFormulaBtn')?.addEventListener('click', promptForColumnFormula);
    document.getElementById('rowFormulaBtn')?.addEventListener('click', promptForRowFormula);

    // Font dropdowns
    document.getElementById('fontFamilySelect')?.addEventListener('change', changeFontFamily);
    document.getElementById('fontSizeSelect')?.addEventListener('change', changeFontSize);

    // Initialize all necessary components
    sheetTable = document.getElementById('sheetTable');
    columnHeadersContainer = document.querySelector('.column-headers');

    if (!sheetTable) {
        console.error('Sheet table not found');
        return;
    }
    
    // Get sheet name
    currentSheetName = sheetTable.getAttribute('data-sheet-name');
    console.log('Current sheet name:', currentSheetName);
    
    // Attach event listeners to all cells
    initializeSheetCells();
    
    // Add column headers
    addColumnHeaders();
    
    // Add sheet styles
    addSheetStyles();
    
    // Initialize toolbar buttons
    initAllToolbarButtons();
    
    // Direct button connections
    document.getElementById('addRowBtn')?.addEventListener('click', addNewRow);
    document.getElementById('addColumnBtn')?.addEventListener('click', addNewColumn);
    document.getElementById('deleteRowBtn')?.addEventListener('click', deleteSelectedRow);
    document.getElementById('deleteColumnBtn')?.addEventListener('click', deleteSelectedColumn);
    document.getElementById('saveSheetBtn')?.addEventListener('click', saveSheetData);
    
    // Direct connect format buttons
    const directDivideCellBtn = document.getElementById('divideCellBtn');
    if (directDivideCellBtn) {
        directDivideCellBtn.addEventListener('click', function() {
            divideCellInPlace();
        });
    }
    
    const directColumnFormulaBtn = document.getElementById('columnFormulaBtn');
    if (directColumnFormulaBtn) {
        directColumnFormulaBtn.addEventListener('click', function() {
            promptForColumnFormula();
        });
    }
    
    const directRowFormulaBtn = document.getElementById('rowFormulaBtn');
    if (directRowFormulaBtn) {
        directRowFormulaBtn.addEventListener('click', function() {
            promptForRowFormula();
        });
    }
    
    const directMergeBtn = document.getElementById('directMergeBtn') || document.getElementById('mergeCellsBtn');
    if (directMergeBtn) {
        directMergeBtn.addEventListener('click', function() {
            mergeCells();
        });
    }
    
    const alignLeftBtn = document.getElementById('alignLeftBtn');
    if (alignLeftBtn) {
        alignLeftBtn.addEventListener('click', function() {
            applyAlignment('left');
        });
    }
    
    const alignCenterBtn = document.getElementById('alignCenterBtn');
    if (alignCenterBtn) {
        alignCenterBtn.addEventListener('click', function() {
            applyAlignment('center');
        });
    }
    
    const alignRightBtn = document.getElementById('alignRightBtn');
    if (alignRightBtn) {
        alignRightBtn.addEventListener('click', function() {
            applyAlignment('right');
        });
    }
    
    const headerColorBtn = document.getElementById('headerColorBtn');
    if (headerColorBtn) {
        headerColorBtn.addEventListener('click', function() {
            promptForHeaderColor();
        });
    }
    
    // Load and apply saved formatting
    loadSavedFormatting();
    
    // Add keyboard shortcuts
    document.addEventListener('keydown', handleKeyboardShortcuts);

    // Clean up event listeners to avoid duplicates
    processedElements = new Set(); // Initialize the set to store processed elements
    cleanupEventListeners();

    // Initialize column resizing
    initializeColumnResizing();

    // Restore divided cells
    setTimeout(restoreDividedCells, 500); // Slight delay to ensure table is loaded

    // Initialize color picker buttons
    initializeColorPickers();
});

// Font family changing function
function changeFontFamily(e) {
    if (selectedCells.length === 0) {
        showToast('Please select at least one cell first', 'error');
        return;
    }
    
    const fontFamily = e.target.value;
    
    // Apply to all selected cells
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.fontFamily = fontFamily;
        }
    });
    
    showToast(`Font family changed to ${fontFamily.split(',')[0].replace(/['"]/g, '')}`, 'success');
}

// Font size changing function
function changeFontSize(e) {
    if (selectedCells.length === 0) {
        showToast('Please select at least one cell first', 'error');
        return;
    }
    
    const fontSize = e.target.value;
    
    // Apply to all selected cells
    selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.fontSize = `${fontSize}px`;
        }
    });
    
    showToast(`Font size changed to ${fontSize}px`, 'success');
}

// Function to improve cell width adjustment by making column resize handles easier to use
function initializeColumnResizing() {
    const resizeHandles = document.querySelectorAll('.column-resize-handle');
    
    resizeHandles.forEach(handle => {
        // Increase the width of the resize handle to make it easier to grab
        handle.style.width = '8px';
        
        // Add a hover effect
        handle.addEventListener('mouseover', function() {
            this.style.backgroundColor = '#217346';
            this.style.opacity = '0.5';
        });
        
        handle.addEventListener('mouseout', function() {
            this.style.backgroundColor = '';
            this.style.opacity = '';
        });
        
        // Make sure the handle extends the full height
        handle.style.height = '100%';
    });
}

// Column resize functionality - This was missing and causing the error
function startColumnResize(e) {
    e.preventDefault();
    e.stopPropagation();
    
    const colIndex = parseInt(this.getAttribute('data-col'));
    const startX = e.pageX;
    
    // Find all cells in this column
    const headerCell = document.querySelector(`th[data-col="${colIndex}"]`);
    const startWidth = headerCell.offsetWidth;
    
    // Create visual indicator
    const indicator = document.createElement('div');
    indicator.style.position = 'absolute';
    indicator.style.top = '0';
    indicator.style.bottom = '0';
    indicator.style.width = '2px';
    indicator.style.backgroundColor = '#0d6efd';
    indicator.style.zIndex = '999';
    indicator.style.left = (e.pageX) + 'px';
    document.body.appendChild(indicator);
    
    function handleMouseMove(e) {
        const width = Math.max(30, startWidth + (e.pageX - startX)); // Minimum width: 30px
        
        // Update indicator position
        indicator.style.left = (e.pageX) + 'px';
        
        // Update column width
        document.querySelectorAll(`[data-col="${colIndex}"]`).forEach(cell => {
            cell.style.width = `${width}px`;
        });
    }
    
    function handleMouseUp() {
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
        
        // Remove indicator
        document.body.removeChild(indicator);
    }
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
}

// Row resize functionality - Adding this for completeness
function startRowResize(e) {
    e.preventDefault();
    e.stopPropagation();
    
    const rowIndex = parseInt(this.getAttribute('data-row'));
    const startY = e.pageY;
    
    // Find the row
    const row = document.querySelector(`tr[data-row="${rowIndex}"]`);
    const startHeight = row.offsetHeight;
    
    // Create visual indicator
    const indicator = document.createElement('div');
    indicator.style.position = 'absolute';
    indicator.style.left = '0';
    indicator.style.right = '0';
    indicator.style.height = '2px';
    indicator.style.backgroundColor = '#0d6efd';
    indicator.style.zIndex = '999';
    indicator.style.top = (e.pageY) + 'px';
    document.body.appendChild(indicator);
    
    function handleMouseMove(e) {
        const height = Math.max(20, startHeight + (e.pageY - startY)); // Minimum height: 20px
        
        // Update indicator position
        indicator.style.top = (e.pageY) + 'px';
        
        // Update row height
        row.style.height = `${height}px`;
    }
    
    function handleMouseUp() {
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
        
        // Remove indicator
        document.body.removeChild(indicator);
    }
    
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
}

// Add column and row functionality
document.addEventListener('DOMContentLoaded', function() {
    // Get buttons
    const addRowBtn = document.getElementById('addRowBtn');
    const addColumnBtn = document.getElementById('addColumnBtn');
    
    // Add event listeners
    if (addRowBtn) {
        addRowBtn.addEventListener('click', addNewRow);
    }
    
    if (addColumnBtn) {
        addColumnBtn.addEventListener('click', addNewColumn);
    }
});

// Function to add a new row
function addNewRow() {
    console.log("Adding new row");
    
    // Get table and rows
    const table = document.getElementById('sheetTable');
    const tbody = table.querySelector('tbody');
    const rows = tbody.querySelectorAll('tr');
    
    // Get the last row index
    const lastRowIndex = rows.length > 0 ? 
        parseInt(rows[rows.length - 1].getAttribute('data-row')) : -1;
    const newRowIndex = lastRowIndex + 1;
    
    // Get the number of columns from header row
    const headerCells = table.querySelectorAll('thead th');
    const columnCount = headerCells.length;
    
    // Create new row
    const newRow = document.createElement('tr');
    newRow.setAttribute('data-row', newRowIndex);
    newRow.style.height = '25px';
    
    // Add cells to row
    for (let i = 0; i < columnCount; i++) {
        const cell = document.createElement('td');
        cell.className = 'sheet-cell';
        cell.setAttribute('data-row', newRowIndex);
        cell.setAttribute('data-col', i);
        
        // Add editable div
        const editableDiv = document.createElement('div');
        editableDiv.className = 'editable-cell';
        editableDiv.contentEditable = true;
        cell.appendChild(editableDiv);
        
        // Add cell click event handler
        cell.addEventListener('click', function(e) {
            // Skip if clicked on resize handle
            if (e.target.classList.contains('row-resize-handle') ||
                e.target.classList.contains('column-resize-handle')) {
                return;
            }
            
            // Check if Ctrl key is pressed for multi-selection
            if (e.ctrlKey || e.metaKey) {
                const index = selectedCells.indexOf(this);
                if (index !== -1) {
                    // Deselect
                    this.classList.remove('selected');
                    selectedCells.splice(index, 1);
                } else {
                    // Add to selection
                    this.classList.add('selected');
                    selectedCells.push(this);
                }
            } else {
                // Single selection - clear existing selection
                selectedCells.forEach(c => c.classList.remove('selected'));
                selectedCells = [this];
                this.classList.add('selected');
            }
            
            // Update button states
            updateButtonStates();
            e.stopPropagation();
        });
        
        // Add row resize handle to last cell
        if (i === columnCount - 1) {
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.setAttribute('data-row', newRowIndex);
            resizeHandle.addEventListener('mousedown', startRowResize);
            cell.appendChild(resizeHandle);
        }
        
        newRow.appendChild(cell);
    }
    
    // Append to table
    tbody.appendChild(newRow);
    
    showToast('New row added successfully', 'success');
    return newRow;
}

// Function to add a new column
function addNewColumn() {
    console.log("Adding new column");
    
    // Get the table
    const table = document.getElementById('sheetTable');
    
    // Get the current column count
    const headerRow = table.querySelector('thead tr');
    const headerCells = headerRow.querySelectorAll('th');
    const newColIndex = headerCells.length;
    
    // Create new header cell
    const newHeader = document.createElement('th');
    newHeader.className = 'sheet-header';
    newHeader.setAttribute('data-col', newColIndex);
    newHeader.style.width = '80px';
    
    // Add column letter
    const letterSpan = document.createElement('span');
    letterSpan.className = 'column-letter';
    letterSpan.textContent = getColumnLetter(newColIndex);
    newHeader.appendChild(letterSpan);
    
    // Add editable div to header
    const headerEditableDiv = document.createElement('div');
    headerEditableDiv.className = 'editable-cell';
    headerEditableDiv.contentEditable = true;
    headerEditableDiv.textContent = '';
    newHeader.appendChild(headerEditableDiv);
    
    // Add column resize handle to header
    const resizeHandle = document.createElement('div');
    resizeHandle.className = 'column-resize-handle';
    resizeHandle.setAttribute('data-col', newColIndex);
    resizeHandle.addEventListener('mousedown', startColumnResize);
    newHeader.appendChild(resizeHandle);
    
    // Add header click event
    newHeader.addEventListener('click', function(e) {
        // Skip if clicked on resize handle
        if (e.target.classList.contains('column-resize-handle')) {
            return;
        }
        
        // Check if Ctrl key is pressed for multi-selection
        if (e.ctrlKey || e.metaKey) {
            const index = selectedCells.indexOf(this);
            if (index !== -1) {
                // Deselect
                this.classList.remove('selected');
                selectedCells.splice(index, 1);
            } else {
                // Add to selection
                this.classList.add('selected');
                selectedCells.push(this);
            }
        } else {
            // Single selection - clear existing selection
            selectedCells.forEach(c => c.classList.remove('selected'));
            selectedCells = [this];
            this.classList.add('selected');
        }
        
        // Update button states
        updateButtonStates();
        e.stopPropagation();
    });
    
    // Add the header cell to the header row
    headerRow.appendChild(newHeader);
    
    // Add cells to each data row
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        const rowIndex = parseInt(row.getAttribute('data-row'));
        
        // Create new cell
        const newCell = document.createElement('td');
        newCell.className = 'sheet-cell';
        newCell.setAttribute('data-row', rowIndex);
        newCell.setAttribute('data-col', newColIndex);
        
        // Add editable div
        const editableDiv = document.createElement('div');
        editableDiv.className = 'editable-cell';
        editableDiv.contentEditable = true;
        newCell.appendChild(editableDiv);
        
        // Add click event
        newCell.addEventListener('click', function(e) {
            // Skip if clicked on resize handle
            if (e.target.classList.contains('row-resize-handle')) {
                return;
            }
            
            // Check if Ctrl key is pressed for multi-selection
            if (e.ctrlKey || e.metaKey) {
                const index = selectedCells.indexOf(this);
                if (index !== -1) {
                    // Deselect
                    this.classList.remove('selected');
                    selectedCells.splice(index, 1);
                } else {
                    // Add to selection
                    this.classList.add('selected');
                    selectedCells.push(this);
                }
            } else {
                // Single selection - clear existing selection
                selectedCells.forEach(c => c.classList.remove('selected'));
                selectedCells = [this];
                this.classList.add('selected');
            }
            
            // Update button states
            updateButtonStates();
            e.stopPropagation();
        });
        
        // Move row resize handle if this is the new last cell
        const lastCell = row.querySelector('.row-resize-handle');
        if (lastCell) {
            const oldHandle = lastCell.parentElement.removeChild(lastCell);
            newCell.appendChild(oldHandle);
        }
        
        // Add the cell to the row
        row.appendChild(newCell);
    });
    
    showToast('New column added successfully', 'success');
}

// Helper function to convert column index to letter (A, B, C..., Z, AA, AB, etc.)
function getColumnLetter(columnIndex) {
    let columnName = '';
    
    if (columnIndex >= 26) {
        columnName += String.fromCharCode(65 + Math.floor(columnIndex / 26) - 1);
        columnName += String.fromCharCode(65 + (columnIndex % 26));
    } else {
        columnName = String.fromCharCode(65 + columnIndex);
    }
    
    return columnName;
}

// Helper function to show a notification toast
function showToast(message, type = 'info') {
    const toastContainer = document.querySelector('.toast-container');
    if (!toastContainer) return;
    
    // Create toast element
    const toast = document.createElement('div');
    toast.className = `toast show mb-2`;
    toast.role = 'alert';
    toast.setAttribute('aria-live', 'assertive');
    toast.setAttribute('aria-atomic', 'true');
    
    // Set background color based on type
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
    
    // Add to container
    toastContainer.appendChild(toast);
    
    // Add close button functionality
    toast.querySelector('.btn-close').addEventListener('click', function() {
        toast.remove();
    });
    
    // Auto-remove after 3 seconds
    setTimeout(() => {
        toast.remove();
    }, 3000);
}

/**
 * Update states of action buttons based on current selection
 */
function updateButtonStates() {
    // Get references to all the buttons we need to manage
    const deleteRowBtn = document.getElementById('deleteRowBtn');
    const deleteColBtn = document.getElementById('deleteColumnBtn');
    const moveRowUpBtn = document.getElementById('moveRowUpBtn');
    const moveRowDownBtn = document.getElementById('moveRowDownBtn');
    const moveColLeftBtn = document.getElementById('moveColLeftBtn');
    const moveColRightBtn = document.getElementById('moveColRightBtn');
    
    // Determine if we have a selection
    const hasSelection = selectedCells.length > 0;
    
    // Default state is disabled
    let canDeleteRow = false;
    let canDeleteCol = false;
    let canMoveRowUp = false;
    let canMoveRowDown = false;
    let canMoveColLeft = false;
    let canMoveColRight = false;
    
    if (hasSelection) {
        // Get selected row and column indices
        const rowIndices = [...new Set(selectedCells.map(cell => 
            parseInt(cell.getAttribute('data-row'))))];
        const colIndices = [...new Set(selectedCells.map(cell => 
            parseInt(cell.getAttribute('data-col'))))];
        
        const totalRows = document.querySelectorAll('#sheetTable tbody tr').length;
        const totalCols = getTableColumnCount();
        
        // Always allow row deletion if there's more than one row
        canDeleteRow = totalRows > 1;
        
        // Always allow column deletion if there's more than one column
        canDeleteCol = totalCols > 1;
        
        // Can move rows up if they're not the first row
        canMoveRowUp = !rowIndices.includes(0);
        
        // Can move rows down if they're not the last row
        canMoveRowDown = !rowIndices.includes(totalRows - 1);
        
        // Can move columns left if they're not the first column
        canMoveColLeft = !colIndices.includes(0);
        
        // Can move columns right if they're not the last column
        canMoveColRight = !colIndices.includes(totalCols - 1);
    }
    
    // Update button states based on the above logic
    if (deleteRowBtn) {
        deleteRowBtn.disabled = !hasSelection || !canDeleteRow;
    }
    
    if (deleteColBtn) {
        deleteColBtn.disabled = !hasSelection || !canDeleteCol;
    }
    
    if (moveRowUpBtn) {
        moveRowUpBtn.disabled = !hasSelection || !canMoveRowUp;
    }
    
    if (moveRowDownBtn) {
        moveRowDownBtn.disabled = !hasSelection || !canMoveRowDown;
    }
    
    if (moveColLeftBtn) {
        moveColLeftBtn.disabled = !hasSelection || !canMoveColLeft;
    }
    
    if (moveColRightBtn) {
        moveColRightBtn.disabled = !hasSelection || !canMoveColRight;
    }
}

// Helper function to get the number of columns in the table
function getTableColumnCount() {
    const headerRow = document.querySelector('#sheetTable thead tr');
    return headerRow ? headerRow.children.length : 0;
}

// Function to delete selected row
function deleteSelectedRow() {
    // Check if we have a selection
    if (selectedCells.length === 0) {
        showToast('Please select a cell in the row you want to delete', 'error');
        return;
    }
    
    // Get unique row indices from selected cells
    const rowIndices = [...new Set(
        selectedCells.map(cell => parseInt(cell.getAttribute('data-row')))
    )].sort((a, b) => b - a); // Sort in descending order to delete from bottom up
    
    if (rowIndices.length === 0) {
        showToast('No valid rows selected for deletion', 'error');
        return;
    }
    
    // Confirm deletion
    if (rowIndices.length > 1) {
        if (!confirm(`Are you sure you want to delete ${rowIndices.length} rows?`)) {
            return;
        }
    } else {
        if (!confirm(`Are you sure you want to delete row ${rowIndices[0] + 1}?`)) {
            return;
        }
    }
    
    // Delete each row
    const table = document.getElementById('sheetTable');
    const tbody = table.querySelector('tbody');
    
    rowIndices.forEach(rowIndex => {
        const row = tbody.querySelector(`tr[data-row="${rowIndex}"]`);
        if (row) {
            row.remove();
        }
    });
    
    // Clear selection
    selectedCells = [];
    
    // Update row indices for remaining rows
    renumberRows();
    
    showToast(`${rowIndices.length} row(s) deleted`, 'success');
}

// Function to delete selected column
function deleteSelectedColumn() {
    // Check if we have a selection
    if (selectedCells.length === 0) {
        showToast('Please select a cell in the column you want to delete', 'error');
        return;
    }
    
    // Get unique column indices from selected cells
    const colIndices = [...new Set(
        selectedCells.map(cell => parseInt(cell.getAttribute('data-col')))
    )].sort((a, b) => b - a); // Sort in descending order to delete from right to left
    
    if (colIndices.length === 0) {
        showToast('No valid columns selected for deletion', 'error');
        return;
    }
    
    // Confirm deletion
    if (colIndices.length > 1) {
        if (!confirm(`Are you sure you want to delete ${colIndices.length} columns?`)) {
            return;
        }
    } else {
        if (!confirm(`Are you sure you want to delete column ${getColumnLetter(colIndices[0])}?`)) {
            return;
        }
    }
    
    // Delete each column
    const table = document.getElementById('sheetTable');
    
    colIndices.forEach(colIndex => {
        // Delete header cell
        const headerCell = table.querySelector(`th[data-col="${colIndex}"]`);
        if (headerCell) {
            headerCell.remove();
        }
        
        // Delete all cells in this column
        table.querySelectorAll(`td[data-col="${colIndex}"]`).forEach(cell => {
            cell.remove();
        });
    });
    
    // Clear selection
    selectedCells = [];
    
    // Update column indices for remaining columns
    renumberColumns();
    
    showToast(`${colIndices.length} column(s) deleted`, 'success');
}

// Helper function to renumber rows after deletion
function renumberRows() {
    const table = document.getElementById('sheetTable');
    const rows = table.querySelectorAll('tbody tr');
    
    rows.forEach((row, index) => {
        // Update row data attribute
        row.setAttribute('data-row', index);
        
        // Update row cells
        row.querySelectorAll('td').forEach(cell => {
            cell.setAttribute('data-row', index);
        });
        
        // Update row resize handle
        const resizeHandle = row.querySelector('.row-resize-handle');
        if (resizeHandle) {
            resizeHandle.setAttribute('data-row', index);
        }
    });
}

// Helper function to renumber columns after deletion
function renumberColumns() {
    const table = document.getElementById('sheetTable');
    
    // Get all header cells
    const headerCells = table.querySelectorAll('th');
    
    // Update header cells
    headerCells.forEach((cell, index) => {
        cell.setAttribute('data-col', index);
        
        // Update column letter
        const letterSpan = cell.querySelector('.column-letter');
        if (letterSpan) {
            letterSpan.textContent = getColumnLetter(index);
        }
        
        // Update resize handle
        const resizeHandle = cell.querySelector('.column-resize-handle');
        if (resizeHandle) {
            resizeHandle.setAttribute('data-col', index);
        }
    });
    
    // Update all body cells
    const rows = table.querySelectorAll('tbody tr');
    rows.forEach(row => {
        // Get all cells in this row
        const cells = row.querySelectorAll('td');
        
        // Update cell data attributes
        cells.forEach((cell, index) => {
            cell.setAttribute('data-col', index);
        });
    });
}

/**
 * Move selected row up
 */
function moveRowUp() {
    if (selectedCells.length === 0) return;
    
    // Get unique row indices that are selected
    const rowIndices = [...new Set(selectedCells.map(cell => 
        parseInt(cell.getAttribute('data-row'))))].sort();
    
    // Can't move if first row is selected
    if (rowIndices.includes(0)) {
        console.log("Can't move: first row is selected");
        return;
    }
    
    const table = document.getElementById('sheetTable');
    const tbody = table.querySelector('tbody');
    const rows = tbody.querySelectorAll('tr');
    
    // Process rows from top to bottom to avoid issues
    rowIndices.forEach(rowIdx => {
        const currentRow = rows[rowIdx];
        const prevRow = rows[rowIdx - 1];
        
        // Swap rows in the DOM
        tbody.insertBefore(currentRow, prevRow);
        
        // Update data attributes for the swapped rows
        updateRowAttributes(currentRow, rowIdx - 1);
        updateRowAttributes(prevRow, rowIdx);
    });
    
    // Clear and update selection to follow the moved rows
    clearSelection();
    
    // Select the moved rows at their new positions
    rowIndices.forEach(rowIdx => {
        const newRowIdx = rowIdx - 1;
        const row = tbody.querySelectorAll('tr')[newRowIdx];
        const cells = row.querySelectorAll('td');
        cells.forEach(cell => addToSelection(cell));
    });
    
    updateButtonStates();
    markSheetAsModified();
}

/**
 * Move selected row down
 */
function moveRowDown() {
    if (selectedCells.length === 0) return;
    
    // Get unique row indices that are selected
    const rowIndices = [...new Set(selectedCells.map(cell => 
        parseInt(cell.getAttribute('data-row'))))].sort().reverse();
    
    const table = document.getElementById('sheetTable');
    const tbody = table.querySelector('tbody');
    const rows = tbody.querySelectorAll('tr');
    
    // Can't move if last row is selected
    if (rowIndices.includes(rows.length - 1)) {
        console.log("Can't move: last row is selected");
        return;
    }
    
    // Process rows from bottom to top to avoid issues
    rowIndices.forEach(rowIdx => {
        const currentRow = rows[rowIdx];
        const nextRow = rows[rowIdx + 1];
        
        // Note: insertBefore inserts before the reference node
        // If nextRow.nextSibling is null, it will append to the end
        tbody.insertBefore(currentRow, nextRow.nextSibling);
        
        // Update data attributes for the swapped rows
        updateRowAttributes(currentRow, rowIdx + 1);
        updateRowAttributes(nextRow, rowIdx);
    });
    
    // Clear and update selection to follow the moved rows
    clearSelection();
    
    // Select the moved rows at their new positions
    rowIndices.forEach(rowIdx => {
        const newRowIdx = rowIdx + 1;
        const row = tbody.querySelectorAll('tr')[newRowIdx];
        const cells = row.querySelectorAll('td');
        cells.forEach(cell => addToSelection(cell));
    });
    
    updateButtonStates();
    markSheetAsModified();
}

/**
 * Move selected column left
 */
function moveColumnLeft() {
    if (selectedCells.length === 0) return;
    
    // Get unique column indices that are selected
    const colIndices = [...new Set(selectedCells.map(cell => 
        parseInt(cell.getAttribute('data-col'))))].sort();
    
    // Can't move if first column is selected
    if (colIndices.includes(0)) {
        console.log("Can't move: first column is selected");
        return;
    }
    
    const table = document.getElementById('sheetTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const headerRow = thead.querySelector('tr');
    const bodyRows = tbody.querySelectorAll('tr');
    
    // Process columns from left to right
    colIndices.forEach(colIdx => {
        // Move header cell
        const currentHeaderCell = headerRow.children[colIdx];
        const prevHeaderCell = headerRow.children[colIdx - 1];
        headerRow.insertBefore(currentHeaderCell, prevHeaderCell);
        
        // Move cells in each row
        bodyRows.forEach(row => {
            const currentCell = row.children[colIdx];
            const prevCell = row.children[colIdx - 1];
            row.insertBefore(currentCell, prevCell);
        });
        
        // Update data attributes for all cells in these columns
        updateColumnAttributes(colIdx - 1);
        updateColumnAttributes(colIdx);
    });
    
    // Clear and update selection to follow the moved columns
    clearSelection();
    
    // Select the moved columns at their new positions
    colIndices.forEach(colIdx => {
        const newColIdx = colIdx - 1;
        bodyRows.forEach(row => {
            const cell = row.children[newColIdx];
            if (cell) addToSelection(cell);
        });
    });
    
    updateButtonStates();
    markSheetAsModified();
}

/**
 * Move selected column right
 */
function moveColumnRight() {
    if (selectedCells.length === 0) return;
    
    // Get unique column indices that are selected
    const colIndices = [...new Set(selectedCells.map(cell => 
        parseInt(cell.getAttribute('data-col'))))].sort().reverse();
    
    const table = document.getElementById('sheetTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');
    const headerRow = thead.querySelector('tr');
    const bodyRows = tbody.querySelectorAll('tr');
    
    // Can't move if last column is selected
    if (colIndices.includes(headerRow.children.length - 1)) {
        console.log("Can't move: last column is selected");
        return;
    }
    
    // Process columns from right to left
    colIndices.forEach(colIdx => {
        // Move header cell (need to handle the case when moving the last cell)
        const currentHeaderCell = headerRow.children[colIdx];
        const nextHeaderCell = headerRow.children[colIdx + 1];
        
        // Insert after next cell (if next cell has a next sibling)
        if (nextHeaderCell.nextSibling) {
            headerRow.insertBefore(currentHeaderCell, nextHeaderCell.nextSibling);
        } else {
            headerRow.appendChild(currentHeaderCell);
        }
        
        // Move cells in each row
        bodyRows.forEach(row => {
            const currentCell = row.children[colIdx];
            const nextCell = row.children[colIdx + 1];
            
            if (nextCell.nextSibling) {
                row.insertBefore(currentCell, nextCell.nextSibling);
            } else {
                row.appendChild(currentCell);
            }
        });
        
        // Update data attributes for all cells in these columns
        updateColumnAttributes(colIdx);
        updateColumnAttributes(colIdx + 1);
    });
    
    // Clear and update selection to follow the moved columns
    clearSelection();
    
    // Select the moved columns at their new positions
    colIndices.forEach(colIdx => {
        const newColIdx = colIdx + 1;
        bodyRows.forEach(row => {
            const cell = row.children[newColIdx];
            if (cell) addToSelection(cell);
        });
    });
    
    updateButtonStates();
    markSheetAsModified();
}

/**
 * Helper function to update row data attributes after row movement
 */
function updateRowAttributes(row, newRowIdx) {
    const cells = row.querySelectorAll('td');
    cells.forEach(cell => {
        cell.setAttribute('data-row', newRowIdx);
        
        // Update cell ID which might use row index
        const colIdx = cell.getAttribute('data-col');
        cell.id = `cell_${newRowIdx}_${colIdx}`;
    });
}

/**
 * Helper function to update column data attributes after column movement
 */
function updateColumnAttributes(colIdx) {
    const table = document.getElementById('sheetTable');
    const headerCell = table.querySelector(`thead tr th:nth-child(${colIdx + 1})`);
    const bodyCells = table.querySelectorAll(`tbody tr td:nth-child(${colIdx + 1})`);
    
    // Update header
    if (headerCell) {
        headerCell.setAttribute('data-col', colIdx);
    }
    
    // Update body cells
    bodyCells.forEach(cell => {
        cell.setAttribute('data-col', colIdx);
        
        // Update cell ID which might use column index
        const rowIdx = cell.getAttribute('data-row');
        cell.id = `cell_${rowIdx}_${colIdx}`;
    });
}

// Function to initialize all color picker buttons
function initializeColorPickers() {
    // Cell background color
    const bgColorBtn = document.getElementById('bgColorBtn');
    const bgColorPicker = document.getElementById('bgColorPicker');
    
    if (bgColorBtn && bgColorPicker) {
        bgColorBtn.addEventListener('click', function() {
            bgColorPicker.click();
        });
        
        bgColorPicker.addEventListener('input', function() {
            applyCellBackgroundColor(this.value);
        });
    }
    
    // Header background color
    const headerColorBtn = document.getElementById('headerColorBtn');
    const headerColorPicker = document.getElementById('headerColorPicker');
    
    if (headerColorBtn && headerColorPicker) {
        headerColorBtn.addEventListener('click', function() {
            headerColorPicker.click();
        });
        
        headerColorPicker.addEventListener('input', function() {
            applyHeaderBackgroundColor(this.value);
        });
    }
    
    // Text color
    const textColorBtn = document.getElementById('textColorBtn');
    const textColorPicker = document.getElementById('textColorPicker');
    
    if (textColorBtn && textColorPicker) {
        textColorBtn.addEventListener('click', function() {
            textColorPicker.click();
        });
        
        textColorPicker.addEventListener('input', function() {
            applyTextColor(this.value);
        });
    }
    
    // Row background color
    const rowColorBtn = document.getElementById('rowColorBtn');
    const rowColorPicker = document.getElementById('rowColorPicker');
    
    if (rowColorBtn && rowColorPicker) {
        rowColorBtn.addEventListener('click', function() {
            rowColorPicker.click();
        });
        
        rowColorPicker.addEventListener('input', function() {
            applyRowBackgroundColor(this.value);
        });
    }
    
    // Column background color
    const columnColorBtn = document.getElementById('columnColorBtn');
    const columnColorPicker = document.getElementById('columnColorPicker');
    
    if (columnColorBtn && columnColorPicker) {
        columnColorBtn.addEventListener('click', function() {
            columnColorPicker.click();
        });
        
        columnColorPicker.addEventListener('input', function() {
            applyColumnBackgroundColor(this.value);
        });
    }
}

// Function to apply background color to selected cells
function applyCellBackgroundColor(color) {
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

// Function to apply header background color
function applyHeaderBackgroundColor(color) {
    // Get all selected header cells or apply to all headers if none selected
    let headerCells = selectedCells.filter(cell => 
        cell.tagName === 'TH' || cell.classList.contains('sheet-header')
    );
    
    // If no headers are selected, show error
    if (headerCells.length === 0) {
        showToast('Please select at least one header cell', 'error');
        return;
    }
    
    // Apply color to selected headers
    headerCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.backgroundColor = color;
        }
    });
    
    saveSheetData();
    showToast(`Applied color to ${headerCells.length} header(s)`, 'success');
}

// Function to apply text color to selected cells
function applyTextColor(color) {
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

// Function to apply background color to an entire row
function applyRowBackgroundColor(color) {
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select a cell in the row', 'error');
        return;
    }
    
    // Get unique row indices from selected cells
    const rowIndices = [...new Set(
        selectedCells.map(cell => parseInt(cell.getAttribute('data-row')))
    )];
    
    if (rowIndices.length === 0) {
        showToast('Could not identify row', 'error');
        return;
    }
    
    // Apply color to all cells in each selected row
    rowIndices.forEach(rowIndex => {
        const rowCells = document.querySelectorAll(`td[data-row="${rowIndex}"]`);
        rowCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        });
    });
    
    saveSheetData();
    showToast(`Applied color to ${rowIndices.length} row(s)`, 'success');
}

// Function to apply background color to an entire column
function applyColumnBackgroundColor(color) {
    if (!selectedCells || selectedCells.length === 0) {
        showToast('Please select a cell in the column', 'error');
        return;
    }
    
    // Get unique column indices from selected cells
    const colIndices = [...new Set(
        selectedCells.map(cell => parseInt(cell.getAttribute('data-col')))
    )];
    
    if (colIndices.length === 0) {
        showToast('Could not identify column', 'error');
        return;
    }
    
    // Apply color to all cells in each selected column
    colIndices.forEach(colIndex => {
        // Apply to header
        const headerCell = document.querySelector(`th[data-col="${colIndex}"]`);
        if (headerCell) {
            const editableDiv = headerCell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        }
        
        // Apply to data cells
        const colCells = document.querySelectorAll(`td[data-col="${colIndex}"]`);
        colCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        });
    });
    
    saveSheetData();
    showToast(`Applied color to ${colIndices.length} column(s)`, 'success');
}