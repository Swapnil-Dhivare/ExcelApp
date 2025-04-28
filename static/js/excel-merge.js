/**
 * Excel-like cell merging functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize cell merging functionality
    initCellMerging();
    
    /**
     * Initialize cell merging functionality
     */
    function initCellMerging() {
        // Add event listeners to merge/unmerge buttons
        const mergeBtn = document.getElementById('mergeCellsBtn');
        if (mergeBtn) {
            mergeBtn.addEventListener('click', mergeCells);
        }
        
        const unmergeBtn = document.getElementById('divideCellBtn');
        if (unmergeBtn) {
            unmergeBtn.addEventListener('click', unmergeCells);
        }
        
        // Load any existing merged cells
        loadMergedCells();
    }
    
    /**
     * Merge selected cells
     */
    function mergeCells() {
        // Check if we have selected cells
        if (!window.selectedCells || window.selectedCells.length <= 1) {
            showToast('Please select multiple cells to merge', 'warning');
            return;
        }
        
        // Get selected cell boundaries
        const boundaries = getSelectionBoundaries();
        
        // Check if selection forms a rectangle
        if (!isRectangularSelection(boundaries)) {
            showToast('Can only merge rectangular cell ranges', 'warning');
            return;
        }
        
        // Find the top-left cell
        const mainCell = document.querySelector(`.sheet-cell[data-row="${boundaries.minRow}"][data-col="${boundaries.minCol}"]`) || 
                        document.querySelector(`.sheet-header[data-row="${boundaries.minRow}"][data-col="${boundaries.minCol}"]`);
        
        if (!mainCell) {
            showToast('Could not find main cell for merge', 'error');
            return;
        }
        
        // Keep only the content from the top-left cell
        const mainContent = mainCell.querySelector('.editable-cell')?.innerHTML || '';
        
        // Calculate colspan and rowspan
        const colspan = boundaries.maxCol - boundaries.minCol + 1;
        const rowspan = boundaries.maxRow - boundaries.minRow + 1;
        
        // Set rowspan and colspan attributes
        mainCell.setAttribute('rowspan', rowspan);
        mainCell.setAttribute('colspan', colspan);
        mainCell.classList.add('merged-cell');
        
        // Set data attributes for saving state
        mainCell.setAttribute('data-merged', 'true');
        mainCell.setAttribute('data-merge-range', `${boundaries.minRow},${boundaries.minCol},${boundaries.maxRow},${boundaries.maxCol}`);
        
        // Set the actual DOM rowSpan and colSpan properties (this is required for HTML tables)
        mainCell.rowSpan = rowspan;
        mainCell.colSpan = colspan;
        
        // Hide all other cells in the merged range
        for (let r = boundaries.minRow; r <= boundaries.maxRow; r++) {
            for (let c = boundaries.minCol; c <= boundaries.maxCol; c++) {
                // Skip the main cell
                if (r === boundaries.minRow && c === boundaries.minCol) continue;
                
                const cell = document.querySelector(`.sheet-cell[data-row="${r}"][data-col="${c}"]`) || 
                            document.querySelector(`.sheet-header[data-row="${r}"][data-col="${c}"]`);
                
                if (cell && cell.parentNode) {
                    // Mark as part of a merge and store reference to the main cell
                    cell.classList.add('hidden-in-merge');
                    cell.setAttribute('data-merged-into', `${boundaries.minRow},${boundaries.minCol}`);
                    
                    // Remove from DOM to properly merge visually
                    cell.parentNode.removeChild(cell);
                }
            }
        }
        
        // Ensure the main cell has proper styling
        const mainEditableDiv = mainCell.querySelector('.editable-cell');
        if (mainEditableDiv) {
            // Center content in merged cell for better appearance
            mainEditableDiv.style.display = 'flex';
            mainEditableDiv.style.justifyContent = 'center';
            mainEditableDiv.style.alignItems = 'center';
            mainEditableDiv.style.height = '100%';
            mainEditableDiv.style.width = '100%';
            
            // Keep the original content
            mainEditableDiv.innerHTML = mainContent;
        }
        
        // Save merged state
        saveMergedCells();
        
        // Show confirmation
        showToast('Cells merged successfully', 'success');
        
        // Clear selection
        if (typeof window.clearCellSelection === 'function') {
            window.clearCellSelection();
        } else {
            document.querySelectorAll('.selected').forEach(cell => {
                cell.classList.remove('selected');
            });
            window.selectedCells = [];
        }
        
        // Select the merged cell
        mainCell.classList.add('selected');
        window.selectedCells = [mainCell];
        
        // Signal that the sheet has been modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    /**
     * Unmerge a merged cell
     */
    function unmergeCells() {
        // Check if we have a selected cell
        if (!window.selectedCells || window.selectedCells.length !== 1) {
            showToast('Please select a merged cell to unmerge', 'warning');
            return;
        }
        
        const cell = window.selectedCells[0];
        
        // Check if cell is merged
        if (!cell.hasAttribute('colspan') && !cell.hasAttribute('rowspan') && !cell.getAttribute('data-merged')) {
            showToast('Selected cell is not merged', 'warning');
            return;
        }
        
        // Get merge dimensions
        const colspan = parseInt(cell.getAttribute('colspan') || cell.colSpan || '1');
        const rowspan = parseInt(cell.getAttribute('rowspan') || cell.rowSpan || '1');
        
        // Check if it's actually merged
        if (colspan === 1 && rowspan === 1 && cell.getAttribute('data-merged') !== 'true') {
            showToast('Selected cell is not merged', 'warning');
            return;
        }
        
        // Get merge range
        let minRow, minCol, maxRow, maxCol;
        
        if (cell.getAttribute('data-merge-range')) {
            const range = cell.getAttribute('data-merge-range').split(',').map(Number);
            minRow = range[0];
            minCol = range[1];
            maxRow = range[2];
            maxCol = range[3];
        } else {
            // Fall back to current cell coordinates
            minRow = parseInt(cell.getAttribute('data-row'));
            minCol = parseInt(cell.getAttribute('data-col'));
            maxRow = minRow + rowspan - 1;
            maxCol = minCol + colspan - 1;
        }
        
        // Get the content from the merged cell
        const content = cell.querySelector('.editable-cell')?.innerHTML || '';
        
        // Reset the main cell
        cell.removeAttribute('colspan');
        cell.removeAttribute('rowspan');
        cell.removeAttribute('colSpan'); // DOM property
        cell.removeAttribute('rowSpan'); // DOM property
        cell.classList.remove('merged-cell');
        cell.removeAttribute('data-merged');
        cell.removeAttribute('data-merge-range');
        
        // Reset the actual HTML attributes
        cell.colSpan = 1;
        cell.rowSpan = 1;
        
        // Reset the editable div styling
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.display = '';
            editableDiv.style.justifyContent = '';
            editableDiv.style.alignItems = '';
            editableDiv.style.height = '';
            editableDiv.style.width = '';
        }
        
        // Get table reference
        const table = document.getElementById('sheetTable');
        if (!table) return;
        
        const tbody = table.querySelector('tbody');
        const thead = table.querySelector('thead');
        if (!tbody || !thead) return;
        
        // Recreate all cells that were part of the merge
        for (let r = minRow; r <= maxRow; r++) {
            // Get or create row
            let row;
            if (r === 0) {
                // Header row
                row = thead.querySelector('tr');
            } else {
                row = tbody.querySelector(`tr[data-row="${r-1}"]`);
            }
            
            if (!row) continue;
            
            // Find existing cells to determine where to insert new cells
            const existingCells = Array.from(row.cells);
            
            // Add cells at the right positions
            for (let c = minCol; c <= maxCol; c++) {
                // Skip the main cell
                if (r === minRow && c === minCol) continue;
                
                // Find position to insert
                let insertPosition = null;
                for (let i = 0; i < existingCells.length; i++) {
                    const cellCol = parseInt(existingCells[i].getAttribute('data-col'));
                    if (cellCol > c) {
                        insertPosition = existingCells[i];
                        break;
                    }
                }
                
                // Create the appropriate cell type
                let newCell;
                if (r === 0) {
                    // Create header cell
                    newCell = document.createElement('th');
                    newCell.className = 'sheet-header';
                } else {
                    // Create data cell
                    newCell = document.createElement('td');
                    newCell.className = 'sheet-cell';
                    newCell.setAttribute('data-row', r - 1);
                }
                
                newCell.setAttribute('data-col', c);
                
                // Add editable div
                const newEditableDiv = document.createElement('div');
                newEditableDiv.className = 'editable-cell';
                newEditableDiv.contentEditable = true;
                newEditableDiv.innerHTML = ''; // Start with empty content
                
                newCell.appendChild(newEditableDiv);
                
                // Insert at correct position
                if (insertPosition) {
                    row.insertBefore(newCell, insertPosition);
                } else {
                    row.appendChild(newCell);
                }
                
                // Setup event listeners
                setupCellEventListeners(newCell);
            }
        }
        
        // Save merged state
        saveMergedCells();
        
        // Show confirmation
        showToast('Cells unmerged successfully', 'success');
        
        // Signal that the sheet has been modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    /**
     * Setup event listeners for a newly created cell
     */
    function setupCellEventListeners(cell) {
        // Add click handler for selection
        cell.addEventListener('click', function(e) {
            if (typeof window.selectCell === 'function') {
                window.selectCell(this, e.ctrlKey || e.metaKey, e.shiftKey);
            } else {
                // Simple selection
                if (!e.ctrlKey && !e.shiftKey) {
                    document.querySelectorAll('.selected').forEach(selectedCell => {
                        selectedCell.classList.remove('selected');
                    });
                    window.selectedCells = [];
                }
                
                this.classList.add('selected');
                if (!window.selectedCells) window.selectedCells = [];
                window.selectedCells.push(this);
            }
        });
        
        // Add double-click listener for editing
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.addEventListener('dblclick', function() {
                this.focus();
            });
        }
    }
    
    /**
     * Check if selection forms a rectangle
     */
    function isRectangularSelection(boundaries) {
        if (!window.selectedCells) return false;
        
        // Calculate expected count based on boundaries
        const rows = boundaries.maxRow - boundaries.minRow + 1;
        const cols = boundaries.maxCol - boundaries.minCol + 1;
        const expectedCount = rows * cols;
        
        // Check if we have the right number of cells selected
        if (window.selectedCells.length !== expectedCount) {
            return false;
        }
        
        // Check if all cells in the rectangle are selected
        for (let r = boundaries.minRow; r <= boundaries.maxRow; r++) {
            for (let c = boundaries.minCol; c <= boundaries.maxCol; c++) {
                // Try to find this cell in the selection
                const found = window.selectedCells.some(cell => {
                    const cellRow = parseInt(cell.getAttribute('data-row'));
                    const cellCol = parseInt(cell.getAttribute('data-col'));
                    return cellRow === r && cellCol === c;
                });
                
                if (!found) return false;
            }
        }
        
        return true;
    }
    
    /**
     * Get the boundaries of the current cell selection
     */
    function getSelectionBoundaries() {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            return { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 };
        }
        
        const rows = [];
        const cols = [];
        
        window.selectedCells.forEach(cell => {
            const row = parseInt(cell.getAttribute('data-row'));
            const col = parseInt(cell.getAttribute('data-col'));
            
            if (!isNaN(row)) rows.push(row);
            if (!isNaN(col)) cols.push(col);
        });
        
        return {
            minRow: Math.min(...rows),
            minCol: Math.min(...cols),
            maxRow: Math.max(...rows),
            maxCol: Math.max(...cols)
        };
    }
    
    /**
     * Save merged cells state to localStorage
     */
    function saveMergedCells() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        const mergedCells = [];
        
        // Find all merged cells
        document.querySelectorAll('.merged-cell').forEach(cell => {
            const row = parseInt(cell.getAttribute('data-row'));
            const col = parseInt(cell.getAttribute('data-col'));
            const rowspan = parseInt(cell.getAttribute('rowspan') || cell.rowSpan || '1');
            const colspan = parseInt(cell.getAttribute('colspan') || cell.colSpan || '1');
            const content = cell.querySelector('.editable-cell')?.innerHTML || '';
            
            mergedCells.push({
                row,
                col,
                rowspan,
                colspan,
                content,
                mergeRange: cell.getAttribute('data-merge-range')
            });
        });
        
        try {
            localStorage.setItem(`excel_merged_cells_${sheetName}`, JSON.stringify(mergedCells));
        } catch (e) {
            console.error('Error saving merged cells:', e);
        }
    }
    
    /**
     * Load merged cells from localStorage
     */
    function loadMergedCells() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        try {
            const savedMergedCells = localStorage.getItem(`excel_merged_cells_${sheetName}`);
            if (!savedMergedCells) return;
            
            const mergedCells = JSON.parse(savedMergedCells);
            
            // Apply merged cells
            mergedCells.forEach(mergeData => {
                const cell = document.querySelector(`.sheet-cell[data-row="${mergeData.row}"][data-col="${mergeData.col}"]`) || 
                            document.querySelector(`.sheet-header[data-row="${mergeData.row}"][data-col="${mergeData.col}"]`);
                
                if (cell) {
                    // Set content
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) {
                        editableDiv.innerHTML = mergeData.content;
                        
                        // Add flex styling
                        editableDiv.style.display = 'flex';
                        editableDiv.style.justifyContent = 'center';
                        editableDiv.style.alignItems = 'center';
                        editableDiv.style.height = '100%';
                        editableDiv.style.width = '100%';
                    }
                    
                    // Set merge attributes
                    cell.setAttribute('rowspan', mergeData.rowspan);
                    cell.setAttribute('colspan', mergeData.colspan);
                    cell.classList.add('merged-cell');
                    cell.setAttribute('data-merged', 'true');
                    
                    // Set DOM properties for actual merging
                    cell.rowSpan = mergeData.rowspan;
                    cell.colSpan = mergeData.colspan;
                    
                    // Set merge range
                    if (mergeData.mergeRange) {
                        cell.setAttribute('data-merge-range', mergeData.mergeRange);
                    } else {
                        // Calculate merge range
                        const minRow = mergeData.row;
                        const minCol = mergeData.col;
                        const maxRow = minRow + mergeData.rowspan - 1;
                        const maxCol = minCol + mergeData.colspan - 1;
                        
                        cell.setAttribute('data-merge-range', `${minRow},${minCol},${maxRow},${maxCol}`);
                    }
                    
                    // Hide other cells in the merge range
                    const mergeRange = cell.getAttribute('data-merge-range').split(',').map(Number);
                    const minRow = mergeRange[0];
                    const minCol = mergeRange[1];
                    const maxRow = mergeRange[2];
                    const maxCol = mergeRange[3];
                    
                    for (let r = minRow; r <= maxRow; r++) {
                        for (let c = minCol; c <= maxCol; c++) {
                            // Skip the main cell
                            if (r === minRow && c === minCol) continue;
                            
                            const hiddenCell = document.querySelector(`.sheet-cell[data-row="${r}"][data-col="${c}"]`) || 
                                             document.querySelector(`.sheet-header[data-row="${r}"][data-col="${c}"]`);
                            
                            if (hiddenCell && hiddenCell.parentNode) {
                                // Mark and remove
                                hiddenCell.classList.add('hidden-in-merge');
                                hiddenCell.setAttribute('data-merged-into', `${minRow},${minCol}`);
                                hiddenCell.parentNode.removeChild(hiddenCell);
                            }
                        }
                    }
                }
            });
        } catch (e) {
            console.error('Error loading merged cells:', e);
            showToast('Error loading merged cells', 'error');
        }
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type) {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            console.log(`${type}: ${message}`);
        }
    }
    
    // Expose API for external use
    window.excelMerge = {
        mergeCells,
        unmergeCells,
        loadMergedCells
    };
});
