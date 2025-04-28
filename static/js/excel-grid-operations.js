/**
 * Excel grid operations for adding/removing rows and columns
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize grid operation handlers
    initGridOperations();
    
    /**
     * Initialize grid operation handlers
     */
    function initGridOperations() {
        // Add event listeners for buttons
        document.getElementById('addRowBtn')?.addEventListener('click', addRow);
        document.getElementById('addColumnBtn')?.addEventListener('click', addColumn);
        document.getElementById('deleteRowBtn')?.addEventListener('click', deleteSelectedRow);
        document.getElementById('deleteColumnBtn')?.addEventListener('click', deleteSelectedColumn);
        
        // Add context menu items if context menu exists
        addContextMenuItems();
    }
    
    /**
     * Add row and column options to context menu if available
     */
    function addContextMenuItems() {
        // Check if context menu already exists
        const contextMenu = document.querySelector('.context-menu');
        if (contextMenu) {
            // Check if items already exist
            if (!contextMenu.querySelector('#cm-add-row')) {
                const addRowItem = document.createElement('div');
                addRowItem.id = 'cm-add-row';
                addRowItem.className = 'context-menu-item';
                addRowItem.innerHTML = '<i class="bi bi-plus-square me-2"></i> Add Row';
                addRowItem.addEventListener('click', addRow);
                
                const addColItem = document.createElement('div');
                addColItem.id = 'cm-add-col';
                addColItem.className = 'context-menu-item';
                addColItem.innerHTML = '<i class="bi bi-plus-square-dotted me-2"></i> Add Column';
                addColItem.addEventListener('click', addColumn);
                
                const deleteRowItem = document.createElement('div');
                deleteRowItem.id = 'cm-delete-row';
                deleteRowItem.className = 'context-menu-item';
                deleteRowItem.innerHTML = '<i class="bi bi-dash-square me-2"></i> Delete Row';
                deleteRowItem.addEventListener('click', deleteSelectedRow);
                
                const deleteColItem = document.createElement('div');
                deleteColItem.id = 'cm-delete-col';
                deleteColItem.className = 'context-menu-item';
                deleteColItem.innerHTML = '<i class="bi bi-dash-square-dotted me-2"></i> Delete Column';
                deleteColItem.addEventListener('click', deleteSelectedColumn);
                
                // Add items in the correct position
                const referenceNode = contextMenu.querySelector('.context-menu-divider');
                if (referenceNode) {
                    contextMenu.insertBefore(deleteColItem, referenceNode);
                    contextMenu.insertBefore(deleteRowItem, referenceNode);
                    contextMenu.insertBefore(addColItem, referenceNode);
                    contextMenu.insertBefore(addRowItem, referenceNode);
                } else {
                    contextMenu.appendChild(addRowItem);
                    contextMenu.appendChild(addColItem);
                    contextMenu.appendChild(deleteRowItem);
                    contextMenu.appendChild(deleteColItem);
                }
            }
        }
    }
    
    /**
     * Add a new row to the sheet
     */
    function addRow() {
        const table = document.getElementById('sheetTable');
        if (!table) return;
        
        // Get table body
        const tbody = table.querySelector('tbody');
        if (!tbody) return;
        
        // Get all existing rows
        const rows = tbody.querySelectorAll('tr');
        const lastRow = rows[rows.length - 1];
        
        if (!lastRow) return;
        
        // Get the number of columns based on header row
        const headerCells = table.querySelectorAll('thead th');
        const columnCount = headerCells.length;
        
        // Determine next row index
        const nextRowIndex = parseInt(lastRow.getAttribute('data-row')) + 1;
        
        // Create new row
        const newRow = document.createElement('tr');
        newRow.setAttribute('data-row', nextRowIndex);
        
        // Add cells to the row
        for (let i = 0; i < columnCount; i++) {
            const cell = document.createElement('td');
            cell.className = 'sheet-cell';
            cell.setAttribute('data-row', nextRowIndex);
            cell.setAttribute('data-col', i);
            
            // Add editable div inside
            const editableDiv = document.createElement('div');
            editableDiv.className = 'editable-cell';
            editableDiv.contentEditable = true;
            
            // Add event listeners
            editableDiv.addEventListener('focus', function() {
                // Ensure this cell is selected when focused
                const cellElement = this.parentElement;
                if (window.selectCell && !cellElement.classList.contains('selected')) {
                    window.selectCell(cellElement);
                }
            });
            
            // Add the editable div to the cell
            cell.appendChild(editableDiv);
            
            // Add resize handle if needed
            if (i === columnCount - 1) {
                const rowResizeHandle = document.createElement('div');
                rowResizeHandle.className = 'row-resize-handle';
                rowResizeHandle.setAttribute('data-row', nextRowIndex);
                cell.appendChild(rowResizeHandle);
            }
            
            newRow.appendChild(cell);
        }
        
        // Add the row to the table
        tbody.appendChild(newRow);
        
        // Update any selection related functionality
        if (window.updateSelectionInfo) {
            window.updateSelectionInfo();
        }
        
        // Initialize cell events for the new row
        initCellEvents(newRow);
        
        // Show success message
        showToast('Row added successfully', 'success');
        
        // Re-initialize resize handlers if available
        if (window.columnResize && typeof window.columnResize.init === 'function') {
            window.columnResize.init();
        }
    }
    
    /**
     * Add a new column to the sheet
     */
    function addColumn() {
        const table = document.getElementById('sheetTable');
        if (!table) return;
        
        // Get header row
        const headerRow = table.querySelector('thead tr');
        if (!headerRow) return;
        
        // Get all data rows
        const dataRows = table.querySelectorAll('tbody tr');
        
        // Determine next column index
        const headerCells = headerRow.querySelectorAll('th');
        const nextColIndex = headerCells.length;
        
        // Add column header
        const headerCell = document.createElement('th');
        headerCell.className = 'sheet-header';
        headerCell.setAttribute('data-col', nextColIndex);
        headerCell.style.width = '80px';
        
        // Add editable div inside
        const headerEditableDiv = document.createElement('div');
        headerEditableDiv.className = 'editable-cell';
        headerEditableDiv.contentEditable = true;
        headerCell.appendChild(headerEditableDiv);
        
        // Add resize handle
        const colResizeHandle = document.createElement('div');
        colResizeHandle.className = 'col-resize-handle';
        colResizeHandle.setAttribute('data-col', nextColIndex);
        headerCell.appendChild(colResizeHandle);
        
        // Add header cell to the header row
        headerRow.appendChild(headerCell);
        
        // Add column cells to each data row
        dataRows.forEach(row => {
            const rowIndex = parseInt(row.getAttribute('data-row'));
            
            // Create data cell
            const dataCell = document.createElement('td');
            dataCell.className = 'sheet-cell';
            dataCell.setAttribute('data-row', rowIndex);
            dataCell.setAttribute('data-col', nextColIndex);
            
            // Add editable div inside
            const editableDiv = document.createElement('div');
            editableDiv.className = 'editable-cell';
            editableDiv.contentEditable = true;
            
            // Add event listeners
            editableDiv.addEventListener('focus', function() {
                // Ensure this cell is selected when focused
                const cellElement = this.parentElement;
                if (window.selectCell && !cellElement.classList.contains('selected')) {
                    window.selectCell(cellElement);
                }
            });
            
            dataCell.appendChild(editableDiv);
            row.appendChild(dataCell);
        });
        
        // Update column headers in the fixed header row if present
        const columnHeadersContainer = document.getElementById('columnHeaders');
        if (columnHeadersContainer) {
            const columnHeader = document.createElement('div');
            columnHeader.className = 'excel-column-header';
            columnHeader.setAttribute('data-col', nextColIndex);
            columnHeader.textContent = getColumnLetter(nextColIndex);
            
            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'col-resize-handle';
            resizeHandle.setAttribute('data-col', nextColIndex);
            columnHeader.appendChild(resizeHandle);
            
            columnHeadersContainer.appendChild(columnHeader);
        }
        
        // Update any selection related functionality
        if (window.updateSelectionInfo) {
            window.updateSelectionInfo();
        }
        
        // Initialize cell events for the new column
        initCellEvents(null, nextColIndex);
        
        // Show success message
        showToast('Column added successfully', 'success');
        
        // Re-initialize resize handlers if available
        if (window.columnResize && typeof window.columnResize.init === 'function') {
            window.columnResize.init();
        }
    }
    
    /**
     * Delete selected row
     */
    function deleteSelectedRow() {
        // Check if we have a selection
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('Please select a cell in the row you want to delete', 'warning');
            return;
        }
        
        // Get row index from first selected cell
        const cell = window.selectedCells[0];
        const rowIndex = parseInt(cell.getAttribute('data-row'));
        
        // Confirm deletion
        if (!confirm(`Are you sure you want to delete row ${rowIndex + 1}?`)) {
            return;
        }
        
        // Find the row in the table
        const table = document.getElementById('sheetTable');
        if (!table) return;
        
        const tbody = table.querySelector('tbody');
        if (!tbody) return;
        
        const row = tbody.querySelector(`tr[data-row="${rowIndex}"]`);
        if (!row) {
            showToast('Row not found', 'error');
            return;
        }
        
        // Remove the row
        tbody.removeChild(row);
        
        // Update data-row attributes for subsequent rows
        const remainingRows = tbody.querySelectorAll('tr');
        for (let i = 0; i < remainingRows.length; i++) {
            const currentRow = remainingRows[i];
            const currentRowIndex = parseInt(currentRow.getAttribute('data-row'));
            
            if (currentRowIndex > rowIndex) {
                // Update row index
                currentRow.setAttribute('data-row', currentRowIndex - 1);
                
                // Update cells in the row
                const cells = currentRow.querySelectorAll('td');
                cells.forEach(cell => {
                    cell.setAttribute('data-row', currentRowIndex - 1);
                });
            }
        }
        
        // Clear selection
        if (window.clearCellSelection) {
            window.clearCellSelection();
        } else {
            window.selectedCells = [];
            document.querySelectorAll('.selected').forEach(el => el.classList.remove('selected'));
        }
        
        // Show success message
        showToast('Row deleted successfully', 'success');
        
        // Signal that the sheet has been modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    /**
     * Delete selected column
     */
    function deleteSelectedColumn() {
        // Check if we have a selection
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('Please select a cell in the column you want to delete', 'warning');
            return;
        }
        
        // Get column index from first selected cell
        const cell = window.selectedCells[0];
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // Confirm deletion
        if (!confirm(`Are you sure you want to delete column ${getColumnLetter(colIndex)}?`)) {
            return;
        }
        
        // Find the table
        const table = document.getElementById('sheetTable');
        if (!table) return;
        
        // Remove header cell
        const headerRow = table.querySelector('thead tr');
        if (headerRow) {
            const headerCell = headerRow.querySelector(`th[data-col="${colIndex}"]`);
            if (headerCell) {
                headerRow.removeChild(headerCell);
            }
        }
        
        // Remove column cells from body rows
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cell = row.querySelector(`td[data-col="${colIndex}"]`);
            if (cell) {
                row.removeChild(cell);
            }
        });
        
        // Remove column header from fixed header row if it exists
        const columnHeadersContainer = document.getElementById('columnHeaders');
        if (columnHeadersContainer) {
            const columnHeader = columnHeadersContainer.querySelector(`.excel-column-header[data-col="${colIndex}"]`);
            if (columnHeader) {
                columnHeadersContainer.removeChild(columnHeader);
            }
        }
        
        // Update data-col attributes for subsequent columns
        updateColumnIndices(colIndex);
        
        // Clear selection
        if (window.clearCellSelection) {
            window.clearCellSelection();
        } else {
            window.selectedCells = [];
            document.querySelectorAll('.selected').forEach(el => el.classList.remove('selected'));
        }
        
        // Show success message
        showToast('Column deleted successfully', 'success');
        
        // Signal that the sheet has been modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    /**
     * Update column indices after deletion
     */
    function updateColumnIndices(deletedColIndex) {
        // Update header cells
        document.querySelectorAll('#sheetTable th[data-col]').forEach(headerCell => {
            const colIndex = parseInt(headerCell.getAttribute('data-col'));
            if (colIndex > deletedColIndex) {
                headerCell.setAttribute('data-col', colIndex - 1);
            }
        });
        
        // Update data cells
        document.querySelectorAll('#sheetTable td[data-col]').forEach(cell => {
            const colIndex = parseInt(cell.getAttribute('data-col'));
            if (colIndex > deletedColIndex) {
                cell.setAttribute('data-col', colIndex - 1);
            }
        });
        
        // Update fixed column headers
        document.querySelectorAll('.excel-column-header[data-col]').forEach(header => {
            const colIndex = parseInt(header.getAttribute('data-col'));
            if (colIndex > deletedColIndex) {
                header.setAttribute('data-col', colIndex - 1);
                header.textContent = getColumnLetter(colIndex - 1);
                
                // Update resize handle
                const resizeHandle = header.querySelector('.col-resize-handle');
                if (resizeHandle) {
                    resizeHandle.setAttribute('data-col', colIndex - 1);
                }
            }
        });
    }
    
    /**
     * Initialize events for new cells
     */
    function initCellEvents(row, colIndex) {
        // Initialize cell selection
        const cells = row ? 
            row.querySelectorAll('.sheet-cell') : 
            document.querySelectorAll(`.sheet-cell[data-col="${colIndex}"]`);
            
        cells.forEach(cell => {
            cell.addEventListener('click', function(e) {
                if (window.selectCell) {
                    window.selectCell(this, e.ctrlKey || e.metaKey, e.shiftKey);
                } else {
                    this.classList.add('selected');
                }
            });
            
            // Add double-click to edit
            cell.addEventListener('dblclick', function() {
                const editableDiv = this.querySelector('.editable-cell');
                if (editableDiv) {
                    editableDiv.focus();
                }
            });
        });
    }
    
    /**
     * Get Excel-style column letter from index (0 = A, 25 = Z, 26 = AA, etc.)
     */
    function getColumnLetter(index) {
        let div = index, result = '';
        while (div >= 0) {
            const mod = div % 26;
            result = String.fromCharCode(65 + mod) + result;
            div = Math.floor(div / 26) - 1;
            if (div < 0) break;
        }
        return result;
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type = 'info') {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            console.log(`${type}: ${message}`);
        }
    }
    
    /**
     * Mark sheet as modified to signal unsaved changes
     */
    window.markSheetModified = function() {
        // Add asterisk to sheet name in the UI
        const sheetNameElement = document.querySelector('.excel-header h5');
        if (sheetNameElement && !sheetNameElement.textContent.includes('*')) {
            sheetNameElement.textContent += ' *';
        }
        
        // Enable save button if it was disabled
        const saveBtn = document.getElementById('saveSheetBtn');
        if (saveBtn) {
            saveBtn.disabled = false;
        }
    };
    
    // Expose methods globally
    window.gridOperations = {
        addRow,
        addColumn,
        deleteSelectedRow,
        deleteSelectedColumn
    };
});
