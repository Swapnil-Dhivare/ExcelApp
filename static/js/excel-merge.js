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
            showToast('Please select multiple cells to merge', 'error');
            return;
        }
        
        // Get selected cell boundaries
        const boundaries = getSelectionBoundaries();
        
        // Check if selection forms a rectangle
        if (!isRectangularSelection(boundaries)) {
            showToast('Can only merge rectangular cell ranges', 'error');
            return;
        }
        
        // Find the top-left cell
        const mainCell = document.querySelector(`.sheet-cell[data-row="${boundaries.minRow}"][data-col="${boundaries.minCol}"]`) || 
                        document.querySelector(`.sheet-header[data-row="${boundaries.minRow}"][data-col="${boundaries.minCol}"]`);
        
        if (!mainCell) {
            showToast('Could not find main cell for merge', 'error');
            return;
        }
        
        // Collect content from all cells
        let mergedContent = '';
        window.selectedCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv && editableDiv.textContent.trim()) {
                if (mergedContent) mergedContent += ' ';
                mergedContent += editableDiv.textContent.trim();
            }
        });
        
        // Set content to main cell
        const mainEditableDiv = mainCell.querySelector('.editable-cell');
        if (mainEditableDiv) {
            mainEditableDiv.textContent = mergedContent;
        }
        
        // Calculate colspan and rowspan
        const colspan = boundaries.maxCol - boundaries.minCol + 1;
        const rowspan = boundaries.maxRow - boundaries.minRow + 1;
        
        // Set attributes on the main cell
        mainCell.setAttribute('colspan', colspan);
        mainCell.setAttribute('rowspan', rowspan);
        mainCell.classList.add('merged-cell');
        
        // Set data attributes for saving state
        mainCell.setAttribute('data-merged', 'true');
        mainCell.setAttribute('data-merge-range', `${boundaries.minRow},${boundaries.minCol},${boundaries.maxRow},${boundaries.maxCol}`);
        
        // Hide all other cells in the merged range
        for (let r = boundaries.minRow; r <= boundaries.maxRow; r++) {
            for (let c = boundaries.minCol; c <= boundaries.maxCol; c++) {
                // Skip the main cell
                if (r === boundaries.minRow && c === boundaries.minCol) continue;
                
                const cell = document.querySelector(`.sheet-cell[data-row="${r}"][data-col="${c}"]`) || 
                            document.querySelector(`.sheet-header[data-row="${r}"][data-col="${c}"]`);
                
                if (cell) {
                    cell.style.display = 'none';
                    cell.classList.add('hidden-in-merge');
                    cell.setAttribute('data-merged-into', `${boundaries.minRow},${boundaries.minCol}`);
                }
            }
        }
        
        // Update the table DOM to reflect merge in HTML
        updateTableDOMForMerge(mainCell, rowspan, colspan);
        
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
        
        // Update UI
        if (typeof window.updateSelectionInfo === 'function') {
            window.updateSelectionInfo();
        }
        
        // Save sheet data
        if (typeof window.saveSheetData === 'function') {
            window.saveSheetData();
        }
    }
    
    /**
     * Unmerge a merged cell
     */
    function unmergeCells() {
        // Check if we have a selected cell
        if (!window.selectedCells || window.selectedCells.length !== 1) {
            showToast('Please select a merged cell to unmerge', 'error');
            return;
        }
        
        const cell = window.selectedCells[0];
        
        // Check if cell is merged
        if (!cell.hasAttribute('colspan') && !cell.hasAttribute('rowspan') && !cell.getAttribute('data-merged')) {
            showToast('Selected cell is not merged', 'error');
            return;
        }
        
        // Get merge dimensions
        const colspan = parseInt(cell.getAttribute('colspan') || '1');
        const rowspan = parseInt(cell.getAttribute('rowspan') || '1');
        
        // Check if it's actually merged
        if (colspan === 1 && rowspan === 1 && cell.getAttribute('data-merged') !== 'true') {
            showToast('Selected cell is not merged', 'error');
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
        
        // Remove colspan and rowspan
        cell.removeAttribute('colspan');
        cell.removeAttribute('rowspan');
        cell.classList.remove('merged-cell');
        cell.removeAttribute('data-merged');
        cell.removeAttribute('data-merge-range');
        
        // Reset the main cell content (optional)
        const mainContent = cell.querySelector('.editable-cell')?.textContent || '';
        
        // Show all hidden cells
        document.querySelectorAll(`.hidden-in-merge[data-merged-into="${minRow},${minCol}"]`).forEach(hiddenCell => {
            hiddenCell.style.display = '';
            hiddenCell.classList.remove('hidden-in-merge');
            hiddenCell.removeAttribute('data-merged-into');
            
            // Clear content in previously hidden cells
            const editableDiv = hiddenCell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.textContent = '';
            }
        });
        
        // Also check by coordinates for cells that might not have the attribute
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                // Skip the main cell
                if (r === minRow && c === minCol) continue;
                
                const hiddenCell = document.querySelector(`.sheet-cell[data-row="${r}"][data-col="${c}"], .sheet-header[data-row="${r}"][data-col="${c}"]`);
                
                if (hiddenCell) {
                    hiddenCell.style.display = '';
                    hiddenCell.classList.remove('hidden-in-merge');
                    hiddenCell.removeAttribute('data-merged-into');
                    
                    // Clear content
                    const editableDiv = hiddenCell.querySelector('.editable-cell');
                    if (editableDiv) {
                        editableDiv.textContent = '';
                    }
                }
            }
        }
        
        // Fix the table DOM
        updateTableDOMForUnmerge(cell);
        
        // Save merged state
        saveMergedCells();
        
        // Show confirmation
        showToast('Cells unmerged successfully', 'success');
        
        // Save sheet data
        if (typeof window.saveSheetData === 'function') {
            window.saveSheetData();
        }
    }
    
    /**
     * Update the table DOM to properly reflect merged cells in HTML
     */
    function updateTableDOMForMerge(cell, rowspan, colspan) {
        // For proper HTML table merging, we need to set the actual rowspan and colspan attributes
        const row = cell.closest('tr');
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // For actual HTML table cell merging
        const actualCell = row.cells[colIndex];
        if (actualCell) {
            actualCell.rowSpan = rowspan;
            actualCell.colSpan = colspan;
        }
    }
    
    /**
     * Update the table DOM when unmerging cells
     */
    function updateTableDOMForUnmerge(cell) {
        // Reset the HTML rowspan and colspan
        const row = cell.closest('tr');
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // For actual HTML table cell unmerging
        const actualCell = row.cells[colIndex];
        if (actualCell) {
            actualCell.rowSpan = 1;
            actualCell.colSpan = 1;
        }
    }
    
    /**
     * Check if selection forms a rectangle
     */
    function isRectangularSelection(boundaries) {
        if (!window.selectedCells) return false;
        
        const selectedCount = window.selectedCells.length;
        const expectedCount = (boundaries.maxRow - boundaries.minRow + 1) * (boundaries.maxCol - boundaries.minCol + 1);
        
        return selectedCount === expectedCount;
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
            const rowspan = parseInt(cell.getAttribute('rowspan') || '1');
            const colspan = parseInt(cell.getAttribute('colspan') || '1');
            
            mergedCells.push({
                row,
                col,
                rowspan,
                colspan,
                content: cell.querySelector('.editable-cell')?.textContent || ''
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
                        editableDiv.textContent = mergeData.content;
                    }
                    
                    // Set merge attributes
                    cell.setAttribute('rowspan', mergeData.rowspan);
                    cell.setAttribute('colspan', mergeData.colspan);
                    cell.classList.add('merged-cell');
                    cell.setAttribute('data-merged', 'true');
                    
                    // Calculate merge range
                    const minRow = mergeData.row;
                    const minCol = mergeData.col;
                    const maxRow = minRow + mergeData.rowspan - 1;
                    const maxCol = minCol + mergeData.colspan - 1;
                    
                    cell.setAttribute('data-merge-range', `${minRow},${minCol},${maxRow},${maxCol}`);
                    
                    // Hide other cells in the merge range
                    for (let r = minRow; r <= maxRow; r++) {
                        for (let c = minCol; c <= maxCol; c++) {
                            // Skip the main cell
                            if (r === minRow && c === minCol) continue;
                            
                            const hiddenCell = document.querySelector(`.sheet-cell[data-row="${r}"][data-col="${c}"]`) || 
                                              document.querySelector(`.sheet-header[data-row="${r}"][data-col="${c}"]`);
                            
                            if (hiddenCell) {
                                hiddenCell.style.display = 'none';
                                hiddenCell.classList.add('hidden-in-merge');
                                hiddenCell.setAttribute('data-merged-into', `${minRow},${minCol}`);
                            }
                        }
                    }
                    
                    // Update the DOM
                    updateTableDOMForMerge(cell, mergeData.rowspan, mergeData.colspan);
                }
            });
        } catch (e) {
            console.error('Error loading merged cells:', e);
        }
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type) {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            console.log(message);
        }
    }
    
    // Expose API for external use
    window.excelMerge = {
        mergeCells,
        unmergeCells,
        loadMergedCells
    };
});
