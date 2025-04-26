// Add this at the top with other global variables
let selectedCells = []; 
let sheetTable;
let currentSheetName = ''; // Declare this globally

// Add this function at the top level (outside any other function)
function showToast(message, type) {
    // Remove any existing toasts
    const existingToasts = document.querySelectorAll('.excel-toast');
    existingToasts.forEach(toast => {
        if (document.body.contains(toast)) {
            document.body.removeChild(toast);
        }
    });
    
    // Create new toast
    const toast = document.createElement('div');
    toast.classList.add('excel-toast');
    toast.style.position = 'fixed';
    toast.style.bottom = '20px';
    toast.style.right = '20px';
    toast.style.maxWidth = '300px';
    toast.style.backgroundColor = type === 'success' ? '#d4edda' : '#f8d7da';
    toast.style.color = type === 'success' ? '#155724' : '#721c24';
    toast.style.padding = '10px 15px';
    toast.style.borderRadius = '4px';
    toast.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
    toast.style.zIndex = '9999';
    toast.style.fontSize = '14px';
    toast.style.opacity = '0.9';
    
    toast.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: center">
            <span>${message}</span>
            <button style="background: none; border: none; cursor: pointer; font-size: 16px; margin-left: 10px;">&times;</button>
        </div>
    `;
    
    // Add close button handler
    toast.querySelector('button').addEventListener('click', () => {
        document.body.removeChild(toast);
    });
    
    document.body.appendChild(toast);
    
    // Auto-hide after 2 seconds
    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.5s';
        setTimeout(() => {
            if (document.body.contains(toast)) {
                document.body.removeChild(toast);
            }
        }, 500);
    }, 2000);
}

// Move this function outside the DOMContentLoaded event
function getCsrfToken() {
    const meta = document.querySelector('meta[name="csrf-token"]');
    return meta ? meta.getAttribute('content') : '';
}

// Add this function at the global scope
function addColumnHeaders() {
    console.log('Adding column headers...');
    
    // Check if headers already exist
    if (document.querySelector('.column-headers')) {
        console.log('Column headers already exist');
        return;
    }
    
    // Get the table element
    const table = document.getElementById('sheetTable');
    if (!table) {
        console.error('Sheet table not found');
        return;
    }
    
    // Get the number of columns
    const headerRow = table.querySelector('thead tr');
    if (!headerRow) {
        console.error('Header row not found');
        return;
    }
    
    const columnCount = headerRow.children.length;
    console.log(`Creating headers for ${columnCount} columns`);
    
    // Create column headers container
    const columnHeadersContainer = document.createElement('div');
    columnHeadersContainer.className = 'column-headers d-flex';
    columnHeadersContainer.style.position = 'sticky';
    columnHeadersContainer.style.top = '0';
    columnHeadersContainer.style.left = '0';
    columnHeadersContainer.style.zIndex = '15';
    columnHeadersContainer.style.backgroundColor = '#f8f9fa';
    columnHeadersContainer.style.borderBottom = '1px solid #dee2e6';
    columnHeadersContainer.style.marginBottom = '4px';
    
    // Create empty space for row headers column
    const emptySpace = document.createElement('div');
    emptySpace.style.width = '40px';
    emptySpace.style.flexShrink = '0';
    columnHeadersContainer.appendChild(emptySpace);
    
    // Create column headers with resize handles
    for (let i = 0; i < columnCount; i++) {
        const columnHeaderWrapper = document.createElement('div');
        columnHeaderWrapper.className = 'column-header-wrapper';
        columnHeaderWrapper.style.position = 'relative';
        columnHeaderWrapper.style.width = '80px'; // Default width
        columnHeaderWrapper.style.flexShrink = '0';
        
        const columnHeader = document.createElement('div');
        columnHeader.className = 'column-header';
        columnHeader.setAttribute('data-col', i);
        columnHeader.textContent = getColumnLetter(i);
        columnHeader.style.width = '100%';
        columnHeader.style.textAlign = 'center';
        columnHeader.style.fontWeight = 'bold';
        columnHeader.style.padding = '4px';
        columnHeader.style.backgroundColor = '#e9ecef';
        columnHeader.style.border = '1px solid #dee2e6';
        columnHeader.style.cursor = 'pointer';
        
        // Add click handler to select the column
        columnHeader.addEventListener('click', function(e) {
            selectEntireColumn(i);
            e.stopPropagation();
        });
        
        // Add resize handle
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'column-resize-handle';
        resizeHandle.style.position = 'absolute';
        resizeHandle.style.top = '0';
        resizeHandle.style.right = '0';
        resizeHandle.style.width = '5px';
        resizeHandle.style.height = '100%';
        resizeHandle.style.cursor = 'col-resize';
        resizeHandle.style.backgroundColor = 'transparent';
        resizeHandle.setAttribute('data-col', i);
        
        // Add resize functionality
        resizeHandle.addEventListener('mousedown', startColumnResize);
        
        columnHeaderWrapper.appendChild(columnHeader);
        columnHeaderWrapper.appendChild(resizeHandle);
        columnHeadersContainer.appendChild(columnHeaderWrapper);
    }
    
    // Insert at the top of the table container
    const tableContainer = table.parentElement;
    tableContainer.insertBefore(columnHeadersContainer, tableContainer.firstChild);
    
    console.log('Column headers added successfully');
}

// Helper function for column letters
function getColumnLetter(index) {
    let letter = '';
    
    // For columns beyond Z (26th)
    if (index >= 26) {
        letter += String.fromCharCode(65 + Math.floor(index / 26) - 1);
        letter += String.fromCharCode(65 + (index % 26));
    } else {
        letter = String.fromCharCode(65 + index);
    }
    
    return letter;
}

// Function to select an entire column
function selectEntireColumn(colIndex) {
    // Clear current selection
    clearCellSelection();
    
    // Select all cells in this column
    const cells = document.querySelectorAll(`td[data-col="${colIndex}"]`);
    cells.forEach(cell => {
        cell.classList.add('selected');
        selectedCells.push(cell);
    });
    
    // Also select the header
    const header = document.querySelector(`th[data-col="${colIndex}"]`);
    if (header) {
        header.classList.add('selected');
        selectedCells.push(header);
    }
}

// Column resize functionality
function startColumnResize(e) {
    e.preventDefault();
    e.stopPropagation();
    
    // Get column index from the resize handle
    const colIndex = parseInt(this.getAttribute('data-col'));
    
    // Get all cells in this column (including header)
    const columnCells = document.querySelectorAll(`[data-col="${colIndex}"]`);
    const columnHeader = document.querySelector(`.column-header[data-col="${colIndex}"]`);
    
    // Starting width and mouse position
    const startX = e.clientX;
    const startWidth = columnHeader ? columnHeader.parentElement.offsetWidth : 80;
    
    // Store the elements being resized
    const elementsToResize = [...columnCells, columnHeader];
    
    // Show resize indicator
    const resizeIndicator = document.createElement('div');
    resizeIndicator.className = 'column-resize-indicator';
    resizeIndicator.style.position = 'absolute';
    resizeIndicator.style.top = '0';
    resizeIndicator.style.bottom = '0';
    resizeIndicator.style.width = '2px';
    resizeIndicator.style.backgroundColor = '#1a73e8';
    resizeIndicator.style.zIndex = '1000';
    resizeIndicator.style.left = (e.clientX) + 'px';
    document.body.appendChild(resizeIndicator);
    
    // Mouse move handler
    function handleMouseMove(e) {
        // Calculate new width
        const width = Math.max(40, startWidth + (e.clientX - startX));
        
        // Update resize indicator position
        resizeIndicator.style.left = (e.clientX) + 'px';
        
        // Preview width change
        if (columnHeader && columnHeader.parentElement) {
            columnHeader.parentElement.style.width = `${width}px`;
        }
        
        // Adjust the width of all cells in this column
        columnCells.forEach(cell => {
            if (cell.tagName === 'TH') {
                cell.style.width = `${width}px`;
            }
        });
    }
    
    // Mouse up handler
    function handleMouseUp(e) {
        // Remove event listeners
        document.removeEventListener('mousemove', handleMouseMove);
        document.removeEventListener('mouseup', handleMouseUp);
        
        // Remove resize indicator
        if (document.body.contains(resizeIndicator)) {
            document.body.removeChild(resizeIndicator);
        }
        
        // Calculate final width
        const finalWidth = Math.max(40, startWidth + (e.clientX - startX));
        
        // Apply final width to all cells
        columnCells.forEach(cell => {
            cell.style.width = `${finalWidth}px`;
        });
        
        // Update column header width
        if (columnHeader && columnHeader.parentElement) {
            columnHeader.parentElement.style.width = `${finalWidth}px`;
        }
        
        // Save the changes
        if (typeof saveSheetData === 'function') {
            saveSheetData();
        }
    }
    
    // Add temporary event listeners
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
}

document.addEventListener('DOMContentLoaded', function() {
    sheetTable = document.getElementById('sheetTable'); // Assign sheetTable here

    // Replace this line:
    // let currentSheetName = getCurrentSheetName();
    // With this:
    currentSheetName = getCurrentSheetName();
    
    console.log('Current sheet name:', currentSheetName);
    
    // Replace your existing click event listener with this one:
    if (sheetTable) {
        sheetTable.addEventListener('click', function(event) {
            const targetCell = event.target.closest('.sheet-cell, .sheet-header');
            if (targetCell) {
                // Pass the event object to toggleCellSelection
                toggleCellSelection(targetCell, false, event);
            } else {
                // Clicked outside cells, clear selection
                clearCellSelection();
            }
        });
    }

    const addRowBtn = document.getElementById('addRowBtn');
    const addColumnBtn = document.getElementById('addColumnBtn');
    const saveSheetBtn = document.getElementById('saveSheetBtn');
    const deleteRowBtn = document.getElementById('deleteRowBtn');
    const deleteColumnBtn = document.getElementById('deleteColumnBtn');
    
    console.log('Sheet editor initialized with buttons:', {
        addRowBtn, addColumnBtn, saveSheetBtn, deleteRowBtn, deleteColumnBtn
    });
    
    // Ensure this function exists to get the current sheet name
    function getCurrentSheetName() {
        if (sheetTable) {
            return sheetTable.getAttribute('data-sheet-name') || '';
        }
        return '';
    }

    // Add event listener to the table for cell clicks
    if (sheetTable) {
        sheetTable.addEventListener('keydown', function(event) {
            // Shift+Arrow keys for extending selection
            if (event.shiftKey && ['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
                // Prevent default to avoid scrolling
                event.preventDefault();
                
                // Make sure we have an active cell
                if (selectedCells.length === 0) return;
                
                const lastCell = selectedCells[selectedCells.length - 1];
                let row = parseInt(lastCell.getAttribute('data-row'));
                let col = parseInt(lastCell.getAttribute('data-col'));
                
                // Adjust row/col based on arrow key
                switch (event.key) {
                    case 'ArrowUp': row = Math.max(0, row - 1); break;
                    case 'ArrowDown': row++; break;
                    case 'ArrowLeft': col = Math.max(0, col - 1); break;
                    case 'ArrowRight': col++; break;
                }
                
                // Find the cell at the new position
                const nextCell = document.querySelector(`.sheet-cell[data-row="${row}"][data-col="${col}"]`);
                if (nextCell) {
                    toggleCellSelection(nextCell, true);  // Pass true for multiSelect
                }
            }
        });
    }

    // Add Row functionality
    if (addRowBtn) {
        addRowBtn.addEventListener('click', function() {
            addNewRow();
        });
    }
    
    // Add Column functionality
    if (addColumnBtn) {
        addColumnBtn.addEventListener('click', function() {
            addNewColumn();
        });
    }
    
    // Delete Row functionality
    if (deleteRowBtn) {
        deleteRowBtn.addEventListener('click', function() {
            deleteSelectedRow();
        });
    }
    
    // Delete Column functionality
    if (deleteColumnBtn) {
        deleteColumnBtn.addEventListener('click', function() {
            deleteSelectedColumn();
        });
    }

    // Save functionality
    if (saveSheetBtn) {
        saveSheetBtn.addEventListener('click', function() {
            saveSheetData();
        });
    }
    
    // Function to add a new row
    function addNewRow() {
        if (!sheetTable) return;
        
        const tbody = sheetTable.querySelector('tbody');
        if (!tbody) return;
        
        const rowCount = tbody.querySelectorAll('tr').length;
        const columnCount = sheetTable.querySelector('thead tr').children.length;
        
        // Create new row
        const newRow = document.createElement('tr');
        newRow.setAttribute('data-row', rowCount);
        newRow.style.height = '25px';
        
        for (let i = 0; i < columnCount; i++) {
            const cell = document.createElement('td');
            cell.classList.add('sheet-cell');
            cell.setAttribute('data-row', rowCount);
            cell.setAttribute('data-col', i);
            
            const editableDiv = document.createElement('div');
            editableDiv.classList.add('editable-cell');
            editableDiv.setAttribute('contenteditable', 'true');
            
            // Check if this column has a formula
            const headerCell = sheetTable.querySelector(`th[data-col="${i}"]`);
            if (headerCell && headerCell.hasAttribute('data-column-formula')) {
                const formula = headerCell.getAttribute('data-column-formula');
                const rowNum = rowCount + 2; // +2 for header row offset
                const cellFormula = formula.replace(/{n}/g, rowNum);
                
                // Apply formula
                cell.classList.add('formula-cell');
                cell.setAttribute('data-formula', cellFormula);
                editableDiv.textContent = cellFormula;
            }
            
            cell.appendChild(editableDiv);
            newRow.appendChild(cell);
        }
        
        tbody.appendChild(newRow);
        showToast('Row added successfully', 'success');
    }
    
    // Function to add a new column
    function addNewColumn() {
        if (!sheetTable) return;
        
        const headerRow = sheetTable.querySelector('thead tr');
        if (!headerRow) return;
        
        const columnCount = headerRow.children.length;
        
        // Add header cell
        const headerCell = document.createElement('th');
        headerCell.classList.add('sheet-header');
        headerCell.setAttribute('data-col', columnCount);
        headerCell.style.width = '80px';
        
        const headerEditableDiv = document.createElement('div');
        headerEditableDiv.classList.add('editable-cell');
        headerEditableDiv.setAttribute('contenteditable', 'true');
        headerEditableDiv.textContent = `Column ${columnCount + 1}`;
        
        headerCell.appendChild(headerEditableDiv);
        headerRow.appendChild(headerCell);
        
        // Add cells to each row
        const rows = sheetTable.querySelectorAll('tbody tr');
        rows.forEach((row, rowIndex) => {
            const cell = document.createElement('td');
            cell.classList.add('sheet-cell');
            cell.setAttribute('data-row', rowIndex);
            cell.setAttribute('data-col', columnCount);
            
            const editableDiv = document.createElement('div');
            editableDiv.classList.add('editable-cell');
            editableDiv.setAttribute('contenteditable', 'true');
            
            cell.appendChild(editableDiv);
            row.appendChild(cell);
        });
        
        showToast('Column added successfully', 'success');
    }
    
    // Function to delete a row
    function deleteSelectedRow() {
        if (!sheetTable || selectedCells.length === 0) {
            showToast('Please select a cell in the row you want to delete', 'error');
            return;
        }
        
        // Find all selected rows (by row index)
        const rowIndices = new Set();
        selectedCells.forEach(cell => {
            if (cell.hasAttribute('data-row')) {
                rowIndices.add(parseInt(cell.getAttribute('data-row')));
            }
        });
        
        if (rowIndices.size === 0) {
            showToast('Please select a cell in a data row to delete', 'error');
            return;
        }
        
        // Confirm deletion
        if (!confirm(`Are you sure you want to delete ${rowIndices.size} row(s)?`)) {
            return;
        }
        
        // Convert to array and sort in reverse order (to avoid index shifting issues)
        const sortedRowIndices = Array.from(rowIndices).sort((a, b) => b - a);
        
        // Remove rows
        const tbody = sheetTable.querySelector('tbody');
        sortedRowIndices.forEach(rowIndex => {
            const row = tbody.querySelector(`tr[data-row="${rowIndex}"]`);
            if (row) {
                row.remove();
            }
        });
        
        // Re-index remaining rows
        const remainingRows = tbody.querySelectorAll('tr');
        remainingRows.forEach((row, newIndex) => {
            row.setAttribute('data-row', newIndex);
            
            // Update cell data-row attributes
            const cells = row.querySelectorAll('.sheet-cell');
            cells.forEach(cell => {
                cell.setAttribute('data-row', newIndex);
            });
        });
        
        // Clear selection
        clearCellSelection();
        showToast(`${sortedRowIndices.length} row(s) deleted`, 'success');
    }
    
    // Function to delete a column
    function deleteSelectedColumn() {
        if (!sheetTable || selectedCells.length === 0) {
            showToast('Please select a cell in the column you want to delete', 'error');
            return;
        }
        
        // Find all selected columns (by column index)
        const colIndices = new Set();
        selectedCells.forEach(cell => {
            if (cell.hasAttribute('data-col')) {
                colIndices.add(parseInt(cell.getAttribute('data-col')));
            }
        });
        
        if (colIndices.size === 0) {
            showToast('Please select a valid column to delete', 'error');
            return;
        }
        
        // Confirm deletion
        if (!confirm(`Are you sure you want to delete ${colIndices.size} column(s)?`)) {
            return;
        }
        
        // Convert to array and sort in reverse order (to avoid index shifting issues)
        const sortedColIndices = Array.from(colIndices).sort((a, b) => b - a);
        
        // Remove column headers
        const headerRow = sheetTable.querySelector('thead tr');
        sortedColIndices.forEach(colIndex => {
            const header = headerRow.querySelector(`th[data-col="${colIndex}"]`);
            if (header) {
                header.remove();
            }
        });
        
        // Remove cells in each row
        const rows = sheetTable.querySelectorAll('tbody tr');
        rows.forEach(row => {
            sortedColIndices.forEach(colIndex => {
                const cell = row.querySelector(`td[data-col="${colIndex}"]`);
                if (cell) {
                    cell.remove();
                }
            });
        });
        
        // Re-index remaining columns
        const headerCells = headerRow.querySelectorAll('th');
        headerCells.forEach((header, newIndex) => {
            header.setAttribute('data-col', newIndex);
        });
        
        // Re-index cells
        rows.forEach((row, rowIndex) => {
            const cells = row.querySelectorAll('td');
            cells.forEach((cell, cellIndex) => {
                cell.setAttribute('data-col', cellIndex);
            });
        });
        
        // Clear selection
        clearCellSelection();
        showToast(`${sortedColIndices.length} column(s) deleted`, 'success');
    }
    
    // Update saveSheetData function to include merge data and column formulas
    function saveSheetData() {
        if (!currentSheetName) {
            showToast('Error: Sheet name not found', 'error');
            return;
        }
        
        // Extract sheet data
        const headers = [];
        const headerCells = sheetTable.querySelectorAll('thead th .editable-cell');
        headerCells.forEach(cell => {
            headers.push(cell.textContent.trim());
        });
        
        const rows = [];
        const tableRows = sheetTable.querySelectorAll('tbody tr');
        tableRows.forEach(row => {
            const rowData = [];
            const cells = row.querySelectorAll('.sheet-cell');
            cells.forEach(cell => {
                // Skip hidden merged cells
                if (cell.style.display === 'none') {
                    rowData.push('');
                    return;
                }
                
                let cellContent = cell.querySelector('.editable-cell')?.textContent.trim() || '';
                
                // Handle formulas
                if (cell.classList.contains('formula-cell')) {
                    cellContent = cell.getAttribute('data-formula') || cellContent;
                }
                
                rowData.push(cellContent);
            });
            rows.push(rowData);
        });
        
        const sheetData = [headers, ...rows];
        
        // Collect cell dimensions
        const columnWidths = {};
        sheetTable.querySelectorAll('thead th').forEach(header => {
            const colIndex = header.getAttribute('data-col');
            const width = header.style.width;
            if (width) {
                columnWidths[colIndex] = parseInt(width);
            }
        });
        
        const rowHeights = {};
        sheetTable.querySelectorAll('tbody tr').forEach(row => {
            const rowIndex = row.getAttribute('data-row');
            const height = row.style.height;
            if (height) {
                rowHeights[rowIndex] = parseInt(height);
            }
        });
        
        // Collect merged cells info
        const mergedCells = [];
        sheetTable.querySelectorAll('[data-rowspan], [data-colspan]').forEach(cell => {
            const rowIndex = parseInt(cell.getAttribute('data-row'));
            const colIndex = parseInt(cell.getAttribute('data-col'));
            const rowSpan = parseInt(cell.getAttribute('data-rowspan')) || 1;
            const colSpan = parseInt(cell.getAttribute('data-colspan')) || 1;
            
            if (rowSpan > 1 || colSpan > 1) {
                mergedCells.push({
                    first_row: rowIndex,
                    last_row: rowIndex + rowSpan - 1,
                    first_col: colIndex,
                    last_col: colIndex + colSpan - 1
                });
            }
        });
        
        // Collect header colors
        const headerColors = {};
        sheetTable.querySelectorAll('th.sheet-header .editable-cell').forEach(headerCell => {
            const header = headerCell.closest('th');
            if (header) {
                const colIndex = header.getAttribute('data-col');
                const bgColor = getComputedStyle(headerCell).backgroundColor;
                if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)') {
                    headerColors[colIndex] = rgbToHex(bgColor);
                }
            }
        });
        
        // Collect column formulas
        const columnFormulas = {};
        sheetTable.querySelectorAll('th.sheet-header').forEach(header => {
            const colIndex = header.getAttribute('data-col');
            const formula = header.getAttribute('data-column-formula');
            if (formula) {
                columnFormulas[colIndex] = formula;
            }
        });
        
        // Send data to server
        fetch('/save_sheet_data', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'X-CSRFToken': getCsrfToken()
            },
            body: JSON.stringify({
                sheet_name: currentSheetName,
                data: sheetData,
                column_widths: columnWidths,
                row_heights: rowHeights,
                merged_cells: mergedCells,
                header_colors: headerColors,
                column_formulas: columnFormulas
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                showToast('Sheet saved successfully', 'success');
            } else {
                showToast('Error saving sheet: ' + (data.error || 'Unknown error'), 'error');
            }
        })
        .catch(error => {
            console.error('Error:', error);
            showToast('Error saving sheet', 'error');
        });
    }
    
    // Helper function to convert RGB to HEX
    function rgbToHex(rgb) {
        if (rgb.startsWith('#')) return rgb;
        
        // Extract RGB values
        const rgbMatch = rgb.match(/rgb\((\d+),\s*(\d+),\s*(\d+)\)/i);
        if (!rgbMatch) return '#F0E68C'; // Default if can't parse
        
        const r = parseInt(rgbMatch[1]);
        const g = parseInt(rgbMatch[2]);
        const b = parseInt(rgbMatch[3]);
        
        return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
    }
    
    // Helper functions
    function clearCellSelection() {
        selectedCells.forEach(c => c.classList.remove('selected'));
        selectedCells = []; // Reset the array
    }
    
    // Update your toggleCellSelection function to accept the event parameter
    function toggleCellSelection(cell, multiSelect = false, event = null) {
        // Detect Ctrl/Cmd key from the event object
        const ctrlKeyPressed = (event && (event.ctrlKey || event.metaKey)) || multiSelect;
        
        // Log for debugging
        console.log('Selection with ctrl:', ctrlKeyPressed);
        
        if (ctrlKeyPressed) {
            const index = selectedCells.indexOf(cell);
            
            if (index !== -1) {
                // Deselect if already selected
                cell.classList.remove('selected');
                selectedCells.splice(index, 1);
            } else {
                // Add to selection
                cell.classList.add('selected');
                selectedCells.push(cell);
            }
        } else {
            // Single selection behavior (clear previous selection)
            clearCellSelection();
            cell.classList.add('selected');
            selectedCells.push(cell);
        }
        
        // Update formatting buttons for the last selected cell
        updateFormatButtonStates(cell);
        
        // Prevent default browser behavior if Ctrl is pressed
        if (event && (event.ctrlKey || event.metaKey)) {
            event.preventDefault();
        }
    }

    // Update formatting buttons based on selected cell
    function updateFormatButtonStates(cell) {
        const editableDiv = cell.querySelector('.editable-cell');
        if (!editableDiv) return;
        
        // Bold button
        const boldBtn = document.getElementById('boldBtn');
        if (boldBtn) {
            if (editableDiv.style.fontWeight === 'bold') {
                boldBtn.classList.add('btn-active');
            } else {
                boldBtn.classList.remove('btn-active');
            }
        }
        
        // Italic button
        const italicBtn = document.getElementById('italicBtn');
        if (italicBtn) {
            if (editableDiv.style.fontStyle === 'italic') {
                italicBtn.classList.add('btn-active');
            } else {
                italicBtn.classList.remove('btn-active');
            }
        }
        
        // Underline button
        const underlineBtn = document.getElementById('underlineBtn');
        if (underlineBtn) {
            if (editableDiv.style.textDecoration === 'underline') {
                underlineBtn.classList.add('btn-active');
            } else {
                underlineBtn.classList.remove('btn-active');
            }
        }
        
        // Text alignment buttons
        const alignLeftBtn = document.getElementById('alignLeftBtn');
        const alignCenterBtn = document.getElementById('alignCenterBtn');
        const alignRightBtn = document.getElementById('alignRightBtn');
        
        if (alignLeftBtn) alignLeftBtn.classList.remove('btn-active');
        if (alignCenterBtn) alignCenterBtn.classList.remove('btn-active');
        if (alignRightBtn) alignRightBtn.classList.remove('btn-active');
        
        const textAlign = editableDiv.style.textAlign;
        if (textAlign === 'left' && alignLeftBtn) {
            alignLeftBtn.classList.add('btn-active');
        } else if (textAlign === 'center' && alignCenterBtn) {
            alignCenterBtn.classList.add('btn-active');
        } else if (textAlign === 'right' && alignRightBtn) {
            alignRightBtn.classList.add('btn-active');
        }
    }
    
    // Enhance the insertColumnAt function to support row-specific column insertion
    function insertColumnAt(colIndex, startRowIndex = null) {
        console.log(`Inserting column at index ${colIndex}, starting from row ${startRowIndex || 'all'}`);
        
        if (!sheetTable) {
            showToast('Sheet table not found', 'error');
            return;
        }
        
        // Get header row
        const headerRow = sheetTable.querySelector('thead tr');
        if (!headerRow) {
            showToast('Header row not found', 'error');
            return;
        }
        
        // Get all rows including headers
        let rows = [headerRow, ...sheetTable.querySelectorAll('tbody tr')];
        
        // Filter rows if we're starting from a specific row index
        if (startRowIndex !== null) {
            rows = [
                headerRow, 
                ...Array.from(sheetTable.querySelectorAll('tbody tr'))
                    .filter((row, idx) => idx >= startRowIndex - 1) 
            ];
        }
        
        // Shift data-col attributes for cells at or after colIndex
        // but only for the rows we're modifying
        rows.forEach((row, rowIndex) => {
            // Skip this logic for header row if we're doing row-specific insertion
            if (startRowIndex !== null && rowIndex === 0) return;
            
            const cells = [...row.querySelectorAll('th, td')];
            
            // First update colIndex and higher cells
            for (let i = cells.length - 1; i >= 0; i--) {
                const cell = cells[i];
                const currentColIndex = parseInt(cell.getAttribute('data-col'));
                
                if (currentColIndex >= colIndex) {
                    // Update column index
                    cell.setAttribute('data-col', currentColIndex + 1);
                    
                    // Update formula references if needed
                    if (cell.classList.contains('formula-cell')) {
                        updateFormulaReferences(cell, colIndex);
                    }
                }
            }
        });
        
        // Now insert new cells at colIndex for the filtered rows
        rows.forEach((row, rowIndex) => {
            // Skip header row if we're doing row-specific insertion
            if (startRowIndex !== null && rowIndex === 0) return;
            
            // Create new cell
            let newCell;
            if (rowIndex === 0) {  // It's the header row
                // Only create a header if we're doing a full column insertion
                if (startRowIndex === null) {
                    newCell = document.createElement('th');
                    newCell.className = 'sheet-header';
                    newCell.setAttribute('data-col', colIndex);
                    
                    const headerEditableDiv = document.createElement('div');
                    headerEditableDiv.className = 'editable-cell';
                    headerEditableDiv.setAttribute('contenteditable', 'true');
                    headerEditableDiv.textContent = `Column ${colIndex + 1}`;
                    
                    newCell.appendChild(headerEditableDiv);
                    
                    // Find the element to insert before
                    const nextCell = row.querySelector(`[data-col="${colIndex + 1}"]`);
                    if (nextCell) {
                        row.insertBefore(newCell, nextCell);
                    } else {
                        row.appendChild(newCell);
                    }
                }
            } else {  // Regular data row
                // Determine if this is a row we should modify
                const actualRowIndex = rowIndex - 1;  // -1 because header is row 0
                if (startRowIndex === null || actualRowIndex >= startRowIndex - 1) {
                    newCell = document.createElement('td');
                    newCell.className = 'sheet-cell';
                    newCell.setAttribute('data-row', actualRowIndex);
                    newCell.setAttribute('data-col', colIndex);
                    
                    const editableDiv = document.createElement('div');
                    editableDiv.className = 'editable-cell';
                    editableDiv.setAttribute('contenteditable', 'true');
                    
                    newCell.appendChild(editableDiv);
                    
                    // Find the element to insert before
                    const nextCell = row.querySelector(`[data-col="${colIndex + 1}"]`);
                    if (nextCell) {
                        row.insertBefore(newCell, nextCell);
                    } else {
                        row.appendChild(newCell);
                    }
                }
            }
        });
        
        // Update server
        saveSheetData();
        
        showToast(`Column inserted at position ${colIndex + 1}`, 'success');
    }
    
    // Enhance the context menu to include row-specific column insertion
    function enhanceContextMenuForRowSpecificColumns() {
        // Extend the existing context menu
        const originalShowContextMenu = window.showContextMenu;
        if (typeof originalShowContextMenu === 'function') {
            window.showContextMenu = function(x, y) {
                // Call the original function
                originalShowContextMenu(x, y);
                
                // Get the context menu
                const menu = document.querySelector('.excel-context-menu');
                if (!menu) return;
                
                // Add insert column at specific row option
                const cell = selectedCells[0];
                if (cell && cell.hasAttribute('data-row')) {
                    const colIndex = parseInt(cell.getAttribute('data-col'));
                    const rowIndex = parseInt(cell.getAttribute('data-row')) + 1; // +1 for 1-based display
                    
                    // Create menu items for inserting column at this row
                    const insertColAtRowBeforeItem = document.createElement('div');
                    insertColAtRowBeforeItem.classList.add('excel-context-menu-item');
                    insertColAtRowBeforeItem.textContent = `Insert Column Before (From Row ${rowIndex})`;
                    insertColAtRowBeforeItem.addEventListener('click', function() {
                        insertColumnAt(colIndex, rowIndex);
                        hideContextMenu();
                    });
                    
                    const insertColAtRowAfterItem = document.createElement('div');
                    insertColAtRowAfterItem.classList.add('excel-context-menu-item');
                    insertColAtRowAfterItem.textContent = `Insert Column After (From Row ${rowIndex})`;
                    insertColAtRowAfterItem.addEventListener('click', function() {
                        insertColumnAt(colIndex + 1, rowIndex);
                        hideContextMenu();
                    });
                    
                    // Add separator and menu items
                    menu.appendChild(document.createElement('div')).classList.add('excel-context-menu-separator');
                    menu.appendChild(insertColAtRowBeforeItem);
                    menu.appendChild(insertColAtRowAfterItem);
                }
            };
        }
    }

    // Call this function during initialization
    enhanceContextMenuForRowSpecificColumns();

    // Function to insert multiple columns at a specific row position
    function insertMultipleColumnsAtRow() {
        // Check if we have a selection
        if (selectedCells.length === 0) {
            showToast('Please select a cell where you want to insert columns', 'error');
            return;
        }
        
        // Get the column and row index of the selected cell
        const cell = selectedCells[0];
        const colIndex = parseInt(cell.getAttribute('data-col'));
        const rowIndex = parseInt(cell.getAttribute('data-row')) + 1; // +1 for 1-based display
        
        // Prompt for number of columns
        const numColumns = prompt('How many columns do you want to insert?', '2');
        if (!numColumns || isNaN(parseInt(numColumns)) || parseInt(numColumns) <= 0) {
            return;
        }
        
        // Insert the specified number of columns
        const count = parseInt(numColumns);
        for (let i = 0; i < count; i++) {
            // Insert columns in reverse order to maintain position
            insertColumnAt(colIndex, rowIndex);
        }
        
        showToast(`${count} columns inserted at row ${rowIndex}`, 'success');
    }

    // Add toolbar button for this function
    function addInsertColumnsAtRowButton() {
        const toolbar = document.querySelector('.card-header .btn-toolbar');
        if (toolbar) {
            const buttonGroup = document.createElement('div');
            buttonGroup.className = 'btn-group me-2';
            
            const button = document.createElement('button');
            button.type = 'button';
            button.className = 'btn btn-sm btn-outline-secondary';
            button.innerHTML = '<i class="bi bi-grid-3x3-gap"></i> Insert Columns at Row';
            button.title = 'Insert multiple columns starting at a specific row';
            button.addEventListener('click', insertMultipleColumnsAtRow);
            
            buttonGroup.appendChild(button);
            toolbar.appendChild(buttonGroup);
        }
    }

    // Call this during initialization
    setTimeout(addInsertColumnsAtRowButton, 500);

    // Initialize color picker for headers
    initHeaderColorPicker();

    // Initialize text alignment controls
    initTextAlignment();

    // Add column headers
    setTimeout(addColumnHeaders, 200);

    // Add merge cells button
    initMergeCellsButton();
    
    // Add styles for merged cells
    addMergedCellsStyles();

    // Initialize header color picker
    initHeaderColorPicker();
});

function initTextAlignment() {
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
}

function applyAlignment(alignment) {
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
    showToast(`Text alignment set to: ${alignment}`, 'success');
}

// Replace the initMergeCellsButton function with this more robust version
function initMergeCellsButton() {
    // Try different selectors to find the toolbar
    const possibleToolbars = [
        '.card-header .btn-toolbar',
        '.btn-toolbar',
        '.sheet-toolbar',
        '.toolbar',
        '.card-header'
    ];
    
    let toolbar = null;
    for (const selector of possibleToolbars) {
        toolbar = document.querySelector(selector);
        if (toolbar) break;
    }
    
    if (!toolbar) {
        console.error('Could not find any suitable toolbar for Merge Cells button');
        
        // Create a simple toolbar if none exists
        const sheetContainer = document.querySelector('.card-body') || document.querySelector('.container') || document.body;
        if (sheetContainer) {
            toolbar = document.createElement('div');
            toolbar.className = 'btn-toolbar mb-2';
            toolbar.style.marginTop = '10px';
            sheetContainer.prepend(toolbar);
            console.log('Created a new toolbar');
        } else {
            console.error('Could not create toolbar, no suitable container found');
            return;
        }
    }
    
    // Check if the merge button already exists
    if (document.getElementById('mergeCellsBtn')) {
        console.log('Merge button already exists');
        return;
    }
    
    // Create the button
    const buttonGroup = document.createElement('div');
    buttonGroup.className = 'btn-group me-2';
    
    const mergeBtn = document.createElement('button');
    mergeBtn.type = 'button';
    mergeBtn.className = 'btn btn-sm btn-outline-secondary';
    mergeBtn.id = 'mergeCellsBtn';
    mergeBtn.innerHTML = '<i class="bi bi-grid-3x3-gap"></i> Merge';
    mergeBtn.title = 'Merge selected cells (Ctrl+click to select multiple)';
    
    // Add debug for click event
    mergeBtn.addEventListener('click', function(e) {
        console.log('Merge button clicked!');
        e.preventDefault(); // Prevent any default behavior
        mergeCells(); // Call the merge function
    });
    
    // Add to toolbar
    buttonGroup.appendChild(mergeBtn);
    toolbar.appendChild(buttonGroup);
    
    console.log('Merge Cells button added successfully');
}

// Replace the existing mergeCells function with this more forgiving version
function mergeCells() {
    console.log('mergeCells function called');
    console.log('Selected cells:', selectedCells);
    
    // Check if at least 2 cells are selected
    if (!selectedCells || selectedCells.length < 2) {
        showToast('Please select at least 2 cells to merge', 'error');
        return;
    }
    
    // Find the bounds of selection
    let minRow = Infinity;
    let maxRow = -1;
    let minCol = Infinity;
    let maxCol = -1;
    let headerCellSelected = false;
    
    // Determine selection boundaries
    selectedCells.forEach(cell => {
        const rowIndex = parseInt(cell.getAttribute('data-row'));
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // Check if this is a header cell
        if (cell.classList.contains('sheet-header')) {
            headerCellSelected = true;
        }
        
        console.log(`Cell at row ${rowIndex}, col ${colIndex}`);
        
        if (!isNaN(rowIndex) && !isNaN(colIndex)) {
            minRow = Math.min(minRow, rowIndex);
            maxRow = Math.max(maxRow, rowIndex);
            minCol = Math.min(minCol, colIndex);
            maxCol = Math.max(maxCol, colIndex);
        }
    });
    
    console.log('Selection bounds:', { minRow, maxRow, minCol, maxCol });
    
    // If we've selected a header cell, don't allow merging
    if (headerCellSelected) {
        showToast('Cannot merge header cells', 'error');
        return;
    }
    
    // Get all cells in the rectangle
    let allCellsInRectangle = [];
    for (let r = minRow; r <= maxRow; r++) {
        for (let c = minCol; c <= maxCol; c++) {
            const cell = sheetTable.querySelector(`td[data-row="${r}"][data-col="${c}"]`);
            if (cell) {
                allCellsInRectangle.push(cell);
            } else {
                console.warn(`Cell not found at row ${r}, col ${c}`);
            }
        }
    }
    
    // Just merge the selected cells even if they don't form a perfect rectangle
    console.log('Will merge all cells in rectangle:', allCellsInRectangle.length);
    
    // Find the main cell (top-left)
    const mainCell = sheetTable.querySelector(`td[data-row="${minRow}"][data-col="${minCol}"]`);
    console.log('Main cell found:', !!mainCell);
    
    if (!mainCell) {
        showToast('Could not find top-left cell for merging', 'error');
        return;
    }
    
    // Get content from all selected cells
    let content = '';
    selectedCells.forEach(cell => {
        const cellContent = cell.querySelector('.editable-cell')?.textContent || '';
        if (cellContent.trim()) {
            content += (content ? ' ' : '') + cellContent;
        }
    });
    
    // Apply content to main cell
    const mainCellDiv = mainCell.querySelector('.editable-cell');
    if (mainCellDiv) {
        mainCellDiv.textContent = content;
    }
    
    // Apply visual merge with CSS
    const spanHeight = maxRow - minRow + 1;
    const spanWidth = maxCol - minCol + 1;
    
    // Apply merge attributes
    mainCell.setAttribute('data-rowspan', spanHeight);
    mainCell.setAttribute('data-colspan', spanWidth);
    mainCell.classList.add('merged-cell');
    
    // Calculate new size
    const heightPx = spanHeight * 25; // 25px is default row height
    const widthPx = spanWidth * 80;  // 80px is default column width
    
    // Apply styles for merged cell
    mainCell.style.height = heightPx + 'px';
    mainCell.style.width = widthPx + 'px';
    mainCell.style.position = 'relative';
    mainCell.style.zIndex = '5';
    
    // Make other cells in the merge range invisible
    allCellsInRectangle.forEach(cell => {
        if (cell === mainCell) return; // Skip main cell
        
        cell.style.display = 'none';
        cell.setAttribute('data-hidden-by-merge', `${minRow},${minCol}`);
    });
    
    // Save merge info to server
    saveMergeCells(minRow, maxRow, minCol, maxCol, content);
    
    // Clear selection
    clearCellSelection();
    
    showToast('Cells merged successfully', 'success');
}

// Add function to save merge cells
function saveMergeCells(minRow, maxRow, minCol, maxCol, content) {
    fetch('/merge_cells', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'X-CSRFToken': getCsrfToken()
        },
        body: JSON.stringify({
            sheet_name: currentSheetName,
            min_row: minRow,
            max_row: maxRow,
            min_col: minCol, 
            max_col: maxCol,
            content: content
        })
    })
    .then(response => response.json())
    .then(data => {
        if (!data.success) {
            console.error('Error saving merge cells:', data.error);
        }
    })
    .catch(error => {
        console.error('Error saving merge cells:', error);
    });
}

// Add function to add styles for merged cells
function addMergedCellsStyles() {
    // Check if styles are already added
    if (document.getElementById('merged-cell-styles')) {
        return;
    }
    
    const style = document.createElement('style');
    style.id = 'merged-cell-styles';
    style.textContent = `
        .merged-cell {
            position: relative;
            box-shadow: 0 0 0 1px #1a73e8;
        }
        
        .merged-cell .editable-cell {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            width: 100%;
            height: 100%;
            overflow: auto;
        }
        
        .sheet-cell.selected, .sheet-header.selected {
            outline: 2px solid #1a73e8;
            z-index: 5;
        }
    `;
    
    document.head.appendChild(style);
}

// Add this at the bottom of your file
window.testMergeCells = function() {
    console.log('Test merge cells called');
    console.log('Selected cells:', selectedCells);
    console.log('Current sheet name:', currentSheetName);
    console.log('Sheet table found:', !!sheetTable);
    
    if (typeof mergeCells === 'function') {
        mergeCells();
    } else {
        console.error('mergeCells function is not defined');
    }
};

// Replace the initHeaderColorPicker function with this improved version
function initHeaderColorPicker() {
    // Try different selectors to find the toolbar
    const possibleToolbars = [
        '.card-header .btn-toolbar',
        '.btn-toolbar',
        '.sheet-toolbar',
        '.toolbar',
        '.card-header'
    ];
    
    let toolbar = null;
    for (const selector of possibleToolbars) {
        toolbar = document.querySelector(selector);
        if (toolbar) break;
    }
    
    if (!toolbar) {
        console.error('Could not find any suitable toolbar for Header Color button');
        return;
    }
    
    // Check if button already exists
    if (document.getElementById('headerColorBtn')) {
        return;
    }
    
    // Create button group
    const buttonGroup = document.createElement('div');
    buttonGroup.className = 'btn-group me-2';
    
    // Create header color button
    const colorBtn = document.createElement('button');
    colorBtn.type = 'button';
    colorBtn.className = 'btn btn-sm btn-outline-secondary';
    colorBtn.id = 'headerColorBtn';
    colorBtn.innerHTML = '<i class="bi bi-palette"></i> Header Color';
    colorBtn.title = 'Set header color';
    
    // Create color picker
    const colorPicker = document.createElement('input');
    colorPicker.type = 'color';
    colorPicker.id = 'headerColorPicker';
    colorPicker.style.position = 'absolute';
    colorPicker.style.visibility = 'hidden';
    colorPicker.value = '#f0e68c'; // Default color
    
    // Add color picker event
    colorPicker.addEventListener('change', function() {
        applyHeaderColor(this.value);
    });
    
    // Add button click handler
    colorBtn.addEventListener('click', function() {
        colorPicker.click();
    });
    
    // Add to DOM
    buttonGroup.appendChild(colorBtn);
    buttonGroup.appendChild(colorPicker);
    toolbar.appendChild(buttonGroup);
    
    console.log('Header color button added successfully');
}

// Fix applyHeaderColor function to handle different header types
function applyHeaderColor(color) {
    if (selectedCells.length === 0) {
        showToast('Please select at least one header cell', 'error');
        return;
    }
    
    // Find header cells - be more flexible in what we consider a header
    const headersApplied = selectedCells.filter(cell => {
        // The header could be a th element or have sheet-header class
        if (cell.tagName === 'TH' || cell.classList.contains('sheet-header')) {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
                return true;
            }
        }
        return false;
    });
    
    if (headersApplied.length === 0) {
        showToast('Please select a header cell', 'error');
        return;
    }
    
    // Save header colors
    saveSheetData();
    
    showToast(`Applied color to ${headersApplied.length} header(s)`, 'success');
}