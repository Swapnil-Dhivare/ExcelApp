/**
 * Excel-like keyboard shortcuts and functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Track clipboard data
    let clipboardData = [];
    
    // Initialize keyboard shortcuts
    initKeyboardShortcuts();
    initCopyPaste();
    
    /**
     * Initialize keyboard shortcuts
     */
    function initKeyboardShortcuts() {
        document.addEventListener('keydown', function(e) {
            // Skip if inside an input field or contentEditable element that's focused
            if (document.activeElement.tagName === 'INPUT' || 
                document.activeElement.tagName === 'TEXTAREA') {
                return;
            }
            
            // Escape key - cancel current operation or selection
            if (e.key === 'Escape') {
                clearCellSelection();
                return;
            }
            
            // Ctrl+S - Save
            if ((e.ctrlKey || e.metaKey) && e.key === 's') {
                e.preventDefault();
                saveSheet();
                return;
            }
            
            // Ctrl+Z - Undo
            if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
                e.preventDefault();
                undoLastAction();
                return;
            }
            
            // Ctrl+Y - Redo
            if ((e.ctrlKey || e.metaKey) && e.key === 'y') {
                e.preventDefault();
                redoLastAction();
                return;
            }
            
            // Ctrl+B - Bold
            if ((e.ctrlKey || e.metaKey) && e.key === 'b') {
                e.preventDefault();
                applyFormatting('bold');
                return;
            }
            
            // Ctrl+I - Italic
            if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
                e.preventDefault();
                applyFormatting('italic');
                return;
            }
            
            // Ctrl+U - Underline
            if ((e.ctrlKey || e.metaKey) && e.key === 'u') {
                e.preventDefault();
                applyFormatting('underline');
                return;
            }
            
            // Delete key - Clear selected cells
            if (e.key === 'Delete') {
                e.preventDefault();
                clearSelectedCells();
                return;
            }
            
            // Arrow keys for navigation
            if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
                handleArrowKeys(e);
                return;
            }
            
            // F2 key - Edit cell
            if (e.key === 'F2') {
                e.preventDefault();
                editActiveCell();
                return;
            }
            
            // Tab key - Move to next cell
            if (e.key === 'Tab') {
                handleTabNavigation(e);
                return;
            }
            
            // Enter key - Move down or edit cell
            if (e.key === 'Enter') {
                handleEnterKey(e);
                return;
            }
            
            // If a regular key is pressed (letter, number), start editing the active cell
            if (e.key.length === 1 && !e.ctrlKey && !e.altKey && !e.metaKey) {
                startEditingWithKey(e);
            }
        });
    }
    
    /**
     * Initialize copy/paste functionality
     */
    function initCopyPaste() {
        // Copy
        document.addEventListener('copy', function(e) {
            if (window.selectedCells && window.selectedCells.length > 0) {
                e.preventDefault();
                copySelectedCells();
            }
        });
        
        // Cut
        document.addEventListener('cut', function(e) {
            if (window.selectedCells && window.selectedCells.length > 0) {
                e.preventDefault();
                cutSelectedCells();
            }
        });
        
        // Paste
        document.addEventListener('paste', function(e) {
            if (window.selectedCells && window.selectedCells.length > 0) {
                e.preventDefault();
                pasteToSelectedCells(e.clipboardData.getData('text'));
            }
        });
    }
    
    /**
     * Copy selected cells to clipboard
     */
    function copySelectedCells() {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        // Clear clipboard data
        clipboardData = [];
        
        // Get row and column boundaries
        const boundaries = getSelectionBoundaries();
        
        // Build clipboard data as 2D array
        for (let row = boundaries.minRow; row <= boundaries.maxRow; row++) {
            const rowData = [];
            
            for (let col = boundaries.minCol; col <= boundaries.maxCol; col++) {
                const cell = getCellAt(row, col);
                if (cell) {
                    const editableDiv = cell.querySelector('.editable-cell');
                    rowData.push(editableDiv ? editableDiv.textContent : '');
                } else {
                    rowData.push('');
                }
            }
            
            clipboardData.push(rowData);
        }
        
        // Convert to tab-delimited text for system clipboard
        const clipboardText = clipboardData.map(row => row.join('\t')).join('\n');
        
        // Copy to system clipboard
        navigator.clipboard.writeText(clipboardText).then(function() {
            showToast(`Copied ${clipboardData.length} × ${clipboardData[0].length} cells`, 'success');
        }).catch(function() {
            showToast('Failed to copy to clipboard', 'error');
        });
    }
    
    /**
     * Cut selected cells
     */
    function cutSelectedCells() {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        // Copy first
        copySelectedCells();
        
        // Then clear
        clearSelectedCells();
        
        showToast('Cut cells to clipboard', 'success');
    }
    
    /**
     * Paste data to selected cells
     */
    function pasteToSelectedCells(text) {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        // Parse clipboard text
        const rows = text.split(/\r\n|\n|\r/);
        const data = rows.map(row => row.split(/\t/));
        
        // Get the top-left cell of selection as starting point
        const boundaries = getSelectionBoundaries();
        const startRow = boundaries.minRow;
        const startCol = boundaries.minCol;
        
        // Paste data
        for (let i = 0; i < data.length; i++) {
            const rowData = data[i];
            
            for (let j = 0; j < rowData.length; j++) {
                const targetRow = startRow + i;
                const targetCol = startCol + j;
                
                const cell = getCellAt(targetRow, targetCol);
                if (cell) {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) {
                        editableDiv.textContent = rowData[j];
                    }
                }
            }
        }
        
        showToast('Pasted data', 'success');
    }
    
    /**
     * Get cell at specified row and column
     */
    function getCellAt(row, col) {
        return document.querySelector(`.sheet-cell[data-row="${row}"][data-col="${col}"]`);
    }
    
    /**
     * Get the boundaries of selected cells
     */
    function getSelectionBoundaries() {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            return { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 };
        }
        
        // Get row and column indices
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
     * Clear selected cells
     */
    function clearSelectedCells() {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        window.selectedCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.textContent = '';
            }
        });
        
        showToast('Cleared cells', 'success');
    }
    
    /**
     * Clear cell selection
     */
    function clearCellSelection() {
        if (typeof window.clearSelection === 'function') {
            window.clearSelection();
        } else {
            document.querySelectorAll('.selected').forEach(cell => {
                cell.classList.remove('selected');
            });
            window.selectedCells = [];
        }
    }
    
    /**
     * Save the current sheet
     */
    function saveSheet() {
        if (typeof window.saveSheetData === 'function') {
            window.saveSheetData();
        }
    }
    
    /**
     * Handle arrow key navigation
     */
    function handleArrowKeys(e) {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        e.preventDefault();
        
        // Get the active cell
        const lastSelectedCell = window.selectedCells[window.selectedCells.length - 1];
        const row = parseInt(lastSelectedCell.getAttribute('data-row'));
        const col = parseInt(lastSelectedCell.getAttribute('data-col'));
        
        let targetRow = row;
        let targetCol = col;
        
        // Calculate target cell based on arrow key
        switch (e.key) {
            case 'ArrowUp':
                targetRow = Math.max(0, row - 1);
                break;
            case 'ArrowDown':
                targetRow = row + 1;
                break;
            case 'ArrowLeft':
                targetCol = Math.max(0, col - 1);
                break;
            case 'ArrowRight':
                targetCol = col + 1;
                break;
        }
        
        // Get the target cell
        const targetCell = getCellAt(targetRow, targetCol);
        if (targetCell) {
            // Clear current selection unless Shift key is pressed
            if (!e.shiftKey) {
                clearCellSelection();
            }
            
            // Select the target cell
            targetCell.classList.add('selected');
            if (!window.selectedCells) window.selectedCells = [];
            window.selectedCells.push(targetCell);
            
            // Scroll into view if needed
            targetCell.scrollIntoView({ behavior: 'auto', block: 'nearest' });
            
            // Update selection info
            if (typeof window.updateSelectionInfo === 'function') {
                window.updateSelectionInfo();
            }
        }
    }
    
    /**
     * Edit the active cell
     */
    function editActiveCell() {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        const cell = window.selectedCells[window.selectedCells.length - 1];
        const editableDiv = cell.querySelector('.editable-cell');
        
        if (editableDiv) {
            editableDiv.focus();
            
            // Position cursor at the end of text
            const range = document.createRange();
            const selection = window.getSelection();
            range.selectNodeContents(editableDiv);
            range.collapse(false); // false = collapse to end
            selection.removeAllRanges();
            selection.addRange(range);
        }
    }
    
    /**
     * Handle Tab key navigation
     */
    function handleTabNavigation(e) {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        e.preventDefault();
        
        // Get current cell
        const currentCell = window.selectedCells[window.selectedCells.length - 1];
        const row = parseInt(currentCell.getAttribute('data-row'));
        const col = parseInt(currentCell.getAttribute('data-col'));
        
        // Calculate target cell (next cell or previous cell if Shift is pressed)
        let targetRow = row;
        let targetCol = e.shiftKey ? col - 1 : col + 1;
        
        // Handle row wrapping
        if (targetCol < 0) {
            targetRow = Math.max(0, targetRow - 1);
            // Find the last column in the previous row
            const lastCol = document.querySelectorAll(`[data-row="${targetRow}"]`).length - 1;
            targetCol = Math.max(0, lastCol);
        } else if (!getCellAt(targetRow, targetCol)) {
            // If next cell doesn't exist, go to next row
            targetRow++;
            targetCol = 0;
        }
        
        // Get the target cell
        const targetCell = getCellAt(targetRow, targetCol);
        if (targetCell) {
            // Clear current selection
            clearCellSelection();
            
            // Select the target cell
            targetCell.classList.add('selected');
            window.selectedCells = [targetCell];
            
            // Update selection info
            if (typeof window.updateSelectionInfo === 'function') {
                window.updateSelectionInfo();
            }
            
            // Scroll into view if needed
            targetCell.scrollIntoView({ behavior: 'auto', block: 'nearest' });
        }
    }
    
    /**
     * Handle Enter key
     */
    function handleEnterKey(e) {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        e.preventDefault();
        
        if (e.shiftKey) {
            // Shift+Enter - edit cell
            editActiveCell();
        } else {
            // Enter - move to cell below
            const currentCell = window.selectedCells[window.selectedCells.length - 1];
            const row = parseInt(currentCell.getAttribute('data-row'));
            const col = parseInt(currentCell.getAttribute('data-col'));
            
            const targetCell = getCellAt(row + 1, col);
            if (targetCell) {
                // Clear current selection
                clearCellSelection();
                
                // Select the target cell
                targetCell.classList.add('selected');
                window.selectedCells = [targetCell];
                
                // Update selection info
                if (typeof window.updateSelectionInfo === 'function') {
                    window.updateSelectionInfo();
                }
                
                // Scroll into view if needed
                targetCell.scrollIntoView({ behavior: 'auto', block: 'nearest' });
            }
        }
    }
    
    /**
     * Start editing cell with the pressed key
     */
    function startEditingWithKey(e) {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        const cell = window.selectedCells[window.selectedCells.length - 1];
        const editableDiv = cell.querySelector('.editable-cell');
        
        if (editableDiv) {
            // Focus on the editable div
            editableDiv.focus();
            
            // Clear existing content and set to the pressed key
            editableDiv.textContent = e.key;
            
            // Position cursor at the end
            const range = document.createRange();
            const selection = window.getSelection();
            range.selectNodeContents(editableDiv);
            range.collapse(false); // false = collapse to end
            selection.removeAllRanges();
            selection.addRange(range);
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
    
    /**
     * Apply formatting to selected cells
     */
    function applyFormatting(format) {
        if (typeof window.applyFormatting === 'function') {
            window.applyFormatting(format);
        }
    }
    
    /**
     * Undo last action
     */
    function undoLastAction() {
        if (typeof window.undoAction === 'function') {
            window.undoAction();
        } else {
            showToast('Undo functionality not implemented yet', 'info');
        }
    }
    
    /**
     * Redo last action
     */
    function redoLastAction() {
        if (typeof window.redoAction === 'function') {
            window.redoAction();
        } else {
            showToast('Redo functionality not implemented yet', 'info');
        }
    }
    
    // Expose some functions globally for use by other scripts
    window.excelKeyboard = {
        copySelectedCells,
        cutSelectedCells,
        pasteToSelectedCells,
        clearSelectedCells,
        editActiveCell
    };
});
