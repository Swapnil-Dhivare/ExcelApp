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
                
                // Add items in the correct position
                const referenceNode = contextMenu.querySelector('.context-menu-divider');
                if (referenceNode) {
                    contextMenu.insertBefore(addColItem, referenceNode);
                    contextMenu.insertBefore(addRowItem, referenceNode);
                } else {
                    contextMenu.appendChild(addRowItem);
                    contextMenu.appendChild(addColItem);
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
    
    // Expose methods globally
    window.gridOperations = {
        addRow,
        addColumn
    };
});
