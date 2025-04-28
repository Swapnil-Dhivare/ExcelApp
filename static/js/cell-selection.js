/**
 * Cell selection functionality for Excel Generator
 */

document.addEventListener('DOMContentLoaded', function() {
    // Initialize selection array if it doesn't exist
    window.selectedCells = window.selectedCells || [];
    
    // Selection state
    let isSelecting = false;
    let startCell = null;
    let lastSelectedCell = null;
    let selectionMode = 'cell'; // 'cell', 'row', 'column'
    
    // All selectable cells (including headers)
    const cells = document.querySelectorAll('.sheet-cell, .sheet-header');
    const cellReference = document.getElementById('cellReference');
    const selectedCellInfo = document.getElementById('selectedCellInfo');
    
    // Clear all cell selections
    function clearSelections() {
        document.querySelectorAll('.selected').forEach(selected => {
            selected.classList.remove('selected');
        });
        window.selectedCells = [];
        updateSelectionInfo();
    }
    
    // Add a cell to the selection
    function addToSelection(cell) {
        if (!cell.classList.contains('selected')) {
            cell.classList.add('selected');
            window.selectedCells.push(cell);
        }
    }
    
    // Remove a cell from the selection
    function removeFromSelection(cell) {
        if (cell.classList.contains('selected')) {
            cell.classList.remove('selected');
            const index = window.selectedCells.indexOf(cell);
            if (index !== -1) {
                window.selectedCells.splice(index, 1);
            }
        }
    }
    
    // Toggle a cell's selection state
    function toggleSelection(cell) {
        if (cell.classList.contains('selected')) {
            removeFromSelection(cell);
        } else {
            addToSelection(cell);
        }
    }
    
    // Update cell reference and selection info
    function updateSelectionInfo() {
        if (window.selectedCells.length === 1) {
            const cell = window.selectedCells[0];
            const row = cell.getAttribute('data-row');
            const col = cell.getAttribute('data-col');
            
            if (row !== null && col !== null) {
                const colLetter = String.fromCharCode(65 + parseInt(col));
                const cellAddr = `${colLetter}${parseInt(row) + 1}`;
                if (cellReference) cellReference.value = cellAddr;
                if (selectedCellInfo) selectedCellInfo.textContent = cellAddr;
            }
        } else if (window.selectedCells.length > 1) {
            // Show range info
            const cellCount = window.selectedCells.length;
            if (selectedCellInfo) selectedCellInfo.textContent = `${cellCount} cells selected`;
        } else {
            // No selection
            if (cellReference) cellReference.value = '';
            if (selectedCellInfo) selectedCellInfo.textContent = '';
        }
    }
    
    // Get cell at given row and column
    function getCellAt(row, col) {
        return document.querySelector(`.sheet-cell[data-row="${row}"][data-col="${col}"]`) || 
               document.querySelector(`.sheet-header[data-row="${row}"][data-col="${col}"]`);
    }
    
    // Select range of cells between startCell and endCell
    function selectCellRange(startCell, endCell) {
        // Get the boundaries of the selection
        const startRow = parseInt(startCell.getAttribute('data-row') || '0');
        const startCol = parseInt(startCell.getAttribute('data-col') || '0');
        const endRow = parseInt(endCell.getAttribute('data-row') || '0');
        const endCol = parseInt(endCell.getAttribute('data-col') || '0');
        
        // Calculate the range
        const minRow = Math.min(startRow, endRow);
        const maxRow = Math.max(startRow, endRow);
        const minCol = Math.min(startCol, endCol);
        const maxCol = Math.max(startCol, endCol);
        
        // Clear previous selection
        clearSelections();
        
        // Select all cells in the range
        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                const cell = getCellAt(row, col);
                if (cell) {
                    addToSelection(cell);
                }
            }
        }
        
        updateSelectionInfo();
    }
    
    // Handle mousedown on cells
    function handleCellMouseDown(e) {
        // If not left-click, do nothing
        if (e.button !== 0) return;
        
        // Start cell selection
        isSelecting = true;
        const cell = e.target.closest('.sheet-cell, .sheet-header');
        
        if (!cell) return;
        
        startCell = cell;
        
        // If shift key is pressed, extend selection from last selected cell
        if (e.shiftKey && lastSelectedCell) {
            selectCellRange(lastSelectedCell, cell);
        } else {
            // If ctrl/cmd key is not pressed, clear selection
            if (!(e.ctrlKey || e.metaKey)) {
                clearSelections();
            }
            
            // Toggle this cell's selection
            toggleSelection(cell);
            lastSelectedCell = cell;
        }
        
        updateSelectionInfo();
        e.preventDefault(); // Prevent text selection
    }
    
    // Handle mousemove during selection
    function handleMouseMove(e) {
        if (!isSelecting || !startCell) return;
        
        // Get the cell under cursor
        const cellUnderCursor = document.elementFromPoint(e.clientX, e.clientY);
        if (!cellUnderCursor) return;
        
        const cell = cellUnderCursor.closest('.sheet-cell, .sheet-header');
        if (!cell) return;
        
        // Select range from start cell to current cell
        selectCellRange(startCell, cell);
        lastSelectedCell = cell;
    }
    
    // Handle mouseup to end selection
    function handleMouseUp() {
        isSelecting = false;
    }
    
    // Apply event handlers to cells
    cells.forEach(cell => {
        cell.addEventListener('mousedown', handleCellMouseDown);
    });
    
    // Global mouse events for drag selection
    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
    
    // Keyboard shortcuts for selection
    document.addEventListener('keydown', function(e) {
        // If no cells are selected, do nothing
        if (window.selectedCells.length === 0) return;
        
        // Get the last selected cell
        const cell = window.selectedCells[window.selectedCells.length - 1];
        if (!cell) return;
        
        const row = parseInt(cell.getAttribute('data-row') || '0');
        const col = parseInt(cell.getAttribute('data-col') || '0');
        
        // Handle arrow keys
        if (e.key === 'ArrowUp' || e.key === 'ArrowDown' || 
            e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
            
            let nextRow = row;
            let nextCol = col;
            
            // Calculate next position
            if (e.key === 'ArrowUp') nextRow = Math.max(0, row - 1);
            if (e.key === 'ArrowDown') nextRow = row + 1;
            if (e.key === 'ArrowLeft') nextCol = Math.max(0, col - 1);
            if (e.key === 'ArrowRight') nextCol = col + 1;
            
            const nextCell = getCellAt(nextRow, nextCol);
            if (nextCell) {
                // If shift key is pressed, extend selection
                if (e.shiftKey) {
                    const currentRange = window.selectedCells.slice();
                    selectCellRange(startCell || cell, nextCell);
                } else {
                    // Otherwise, move selection
                    clearSelections();
                    addToSelection(nextCell);
                    startCell = nextCell;
                    lastSelectedCell = nextCell;
                }
                updateSelectionInfo();
                e.preventDefault();
            }
        }
        
        // Ctrl+A to select all cells
        if ((e.ctrlKey || e.metaKey) && e.key === 'a') {
            clearSelections();
            cells.forEach(cell => addToSelection(cell));
            updateSelectionInfo();
            e.preventDefault();
        }
    });
    
    // Initialize the selection
    updateSelectionInfo();
});
