/**
 * Excel-like cell selection with mouse and keyboard support
 */
document.addEventListener('DOMContentLoaded', function() {
    // Cell selection state
    let isSelecting = false;
    let selectionStart = null;
    let selectionEnd = null;
    let selectedRange = { minRow: 0, maxRow: 0, minCol: 0, maxCol: 0 };
    let selectionOverlay = null;
    let lastSelectedCell = null;
    
    // Initialize global array for selected cells
    window.selectedCells = [];
    
    // Create selection overlay element
    createSelectionOverlay();
    
    // Add event listeners to table cells
    initializeCellSelection();
    
    /**
     * Create a selection overlay element for visual feedback
     */
    function createSelectionOverlay() {
        if (!document.getElementById('selection-overlay')) {
            selectionOverlay = document.createElement('div');
            selectionOverlay.id = 'selection-overlay';
            selectionOverlay.className = 'selection-overlay';
            document.body.appendChild(selectionOverlay);
        } else {
            selectionOverlay = document.getElementById('selection-overlay');
        }
    }
    
    /**
     * Initialize cell selection functionality
     */
    function initializeCellSelection() {
        // Get all table cells (header cells and data cells)
        const cells = document.querySelectorAll('#sheetTable th, #sheetTable td');
        
        cells.forEach(cell => {
            // Ignore cells that already have selection handlers
            if (cell.hasAttribute('data-selection-initialized')) return;
            
            // Add mouse down event for starting selection
            cell.addEventListener('mousedown', startCellSelection);
            
            // Add touch start event for mobile
            cell.addEventListener('touchstart', startCellSelectionMobile, { passive: true });
            
            // Mark the cell as initialized
            cell.setAttribute('data-selection-initialized', 'true');
        });
        
        // Add mouseover event for extending selection
        document.querySelectorAll('#sheetTable').forEach(table => {
            table.addEventListener('mouseover', extendCellSelection);
        });
        
        // Add mouse up event for ending selection
        document.addEventListener('mouseup', endCellSelection);
        document.addEventListener('touchend', endCellSelectionMobile);
        
        // Add keyboard support for cell navigation and selection
        document.addEventListener('keydown', handleKeyboardSelection);
        
        // Add click outside handler to clear selection
        document.addEventListener('click', function(e) {
            if (!e.target.closest('#sheetTable')) {
                clearCellSelection();
            }
        });
    }
    
    /**
     * Start cell selection
     */
    function startCellSelection(e) {
        // Ignore if click was on a resize handle or another interactive element
        if (e.target.classList.contains('col-resize-handle') ||
            e.target.classList.contains('row-resize-handle') ||
            e.target.closest('.interactive-control')) {
            return;
        }
        
        // Target the td/th element
        const cell = e.target.closest('td, th');
        if (!cell) return;
        
        // Get row and column indices
        const rowIndex = parseInt(cell.getAttribute('data-row'));
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // Check if indices are valid
        if (isNaN(rowIndex) && isNaN(colIndex)) return;
        
        // Start selecting
        isSelecting = true;
        
        // Update selection start
        selectionStart = { row: rowIndex, col: colIndex };
        selectionEnd = { row: rowIndex, col: colIndex };
        
        // Handle Ctrl/Cmd key for multiple selections
        if (!e.ctrlKey && !e.metaKey) {
            clearCellSelection();
        }
        
        // Add this cell to selection
        addCellToSelection(cell);
        
        // Update selection info
        updateSelectionInfo();
        
        // Prevent default to avoid text selection
        e.preventDefault();
    }
    
    /**
     * Start cell selection on mobile
     */
    function startCellSelectionMobile(e) {
        // Touch selection is simpler - just select the touched cell
        if (!e.touches || !e.touches[0]) return;
        
        // Ignore if touch was on a resize handle or another interactive element
        if (e.target.classList.contains('col-resize-handle') ||
            e.target.classList.contains('row-resize-handle') ||
            e.target.closest('.interactive-control')) {
            return;
        }
        
        // Target the td/th element
        const cell = e.target.closest('td, th');
        if (!cell) return;
        
        // Clear any existing selection and just select this cell
        clearCellSelection();
        addCellToSelection(cell);
        
        // Update selection info
        updateSelectionInfo();
    }
    
    /**
     * Extend cell selection while dragging
     */
    function extendCellSelection(e) {
        if (!isSelecting) return;
        
        // Target the td/th element
        const cell = e.target.closest('td, th');
        if (!cell) return;
        
        // Get row and column indices
        const rowIndex = parseInt(cell.getAttribute('data-row'));
        const colIndex = parseInt(cell.getAttribute('data-col'));
        
        // Check if indices are valid
        if ((isNaN(rowIndex) && cell.tagName !== 'TH') || isNaN(colIndex)) return;
        
        // If the selection hasn't changed, do nothing
        if (selectionEnd && selectionEnd.row === rowIndex && selectionEnd.col === colIndex) {
            return;
        }
        
        // Update selection end
        selectionEnd = { 
            row: cell.tagName === 'TH' ? null : rowIndex, 
            col: colIndex 
        };
        
        // Update selection visualization
        updateSelection();
        
        // Update selection info
        updateSelectionInfo();
    }
    
    /**
     * End cell selection
     */
    function endCellSelection() {
        if (!isSelecting) return;
        
        // End selection mode
        isSelecting = false;
        
        // Create clipboard data for the selection if needed
        if (window.selectedCells.length > 0) {
            lastSelectedCell = window.selectedCells[window.selectedCells.length - 1];
        }
        
        // Show selection feedback
        showSelectionOverlay();
        
        // Check if multiple cells are selected for merge cell button state
        updateMergeCellButton();
        
        // Update formatting buttons to reflect the last selected cell's formatting
        updateFormattingButtons();
        
        // Update formula bar if available
        updateFormulaBar();
    }
    
    /**
     * End cell selection on mobile
     */
    function endCellSelectionMobile() {
        // For mobile we just make sure the selection is properly visualized
        if (window.selectedCells.length > 0) {
            lastSelectedCell = window.selectedCells[window.selectedCells.length - 1];
        }
        
        // Update the UI elements
        updateMergeCellButton();
        updateFormattingButtons();
        updateFormulaBar();
    }
    
    /**
     * Handle keyboard navigation and selection
     */
    function handleKeyboardSelection(e) {
        // Skip if we're editing a cell or focus is in an input
        if (document.activeElement.tagName === 'INPUT' || 
            document.activeElement.tagName === 'TEXTAREA' ||
            document.activeElement.isContentEditable) {
            return;
        }
        
        // Skip if no cells are selected
        if (window.selectedCells.length === 0) return;
        
        // Get the last selected cell
        const lastCell = lastSelectedCell || window.selectedCells[window.selectedCells.length - 1];
        if (!lastCell) return;
        
        // Get the current row and column
        const rowIndex = parseInt(lastCell.getAttribute('data-row'));
        const colIndex = parseInt(lastCell.getAttribute('data-col'));
        
        // Skip if indices are invalid
        if ((isNaN(rowIndex) && lastCell.tagName !== 'TH') || isNaN(colIndex)) return;
        
        // Calculate target cell based on key press
        let targetRow = rowIndex;
        let targetCol = colIndex;
        
        switch(e.key) {
            case 'ArrowUp':
                targetRow = Math.max(0, rowIndex - 1);
                e.preventDefault();
                break;
                
            case 'ArrowDown':
                targetRow = rowIndex + 1;
                e.preventDefault();
                break;
                
            case 'ArrowLeft':
                targetCol = Math.max(0, colIndex - 1);
                e.preventDefault();
                break;
                
            case 'ArrowRight':
                targetCol = colIndex + 1;
                e.preventDefault();
                break;
                
            case 'Tab':
                if (e.shiftKey) {
                    targetCol = Math.max(0, colIndex - 1);
                } else {
                    targetCol = colIndex + 1;
                }
                e.preventDefault();
                break;
                
            case 'Home':
                targetCol = 0;
                if (e.ctrlKey || e.metaKey) {
                    targetRow = 0;
                }
                e.preventDefault();
                break;
                
            case 'End':
                const lastCol = getLastColumnIndex();
                targetCol = lastCol;
                if (e.ctrlKey || e.metaKey) {
                    targetRow = getLastRowIndex();
                }
                e.preventDefault();
                break;
                
            default:
                return; // Exit for keys we don't handle
        }
        
        // Find the target cell
        let targetCell;
        if (lastCell.tagName === 'TH' && targetRow === rowIndex) {
            // If we're moving horizontally from a header cell, stay in the header row
            targetCell = document.querySelector(`th[data-col="${targetCol}"]`);
        } else {
            targetCell = document.querySelector(`[data-row="${targetRow}"][data-col="${targetCol}"]`);
        }
        
        if (!targetCell) {
            // If we went beyond table bounds, try to find the cell at the edge
            if (targetRow > rowIndex) {
                // Trying to go down beyond the table
                targetCell = document.querySelector(`[data-row="${getLastRowIndex()}"][data-col="${targetCol}"]`);
            } else if (targetCol > colIndex) {
                // Trying to go right beyond the table
                targetCell = document.querySelector(`[data-row="${targetRow}"][data-col="${getLastColumnIndex()}"]`);
            }
            
            // If still no target cell, exit
            if (!targetCell) return;
        }
        
        // Handle shift key for extending selection
        if (e.shiftKey) {
            // Extend selection from the original start point to this new cell
            if (!selectionStart) {
                selectionStart = { row: rowIndex, col: colIndex };
            }
            selectionEnd = { 
                row: parseInt(targetCell.getAttribute('data-row')), 
                col: parseInt(targetCell.getAttribute('data-col'))
            };
            
            // Update selection
            updateSelection();
        } else {
            // Just move to the target cell
            clearCellSelection();
            addCellToSelection(targetCell);
            
            // Update selection start and end
            selectionStart = { 
                row: parseInt(targetCell.getAttribute('data-row')), 
                col: parseInt(targetCell.getAttribute('data-col'))
            };
            selectionEnd = { ...selectionStart };
        }
        
        // Ensure the target cell is visible
        targetCell.scrollIntoView({
            behavior: 'auto',
            block: 'nearest',
            inline: 'nearest'
        });
        
        // Update selection UI
        showSelectionOverlay();
        updateSelectionInfo();
        updateMergeCellButton();
        updateFormattingButtons();
        updateFormulaBar();
        
        // Store the last selected cell
        lastSelectedCell = targetCell;
    }
    
    /**
     * Add a cell to the selection
     */
    function addCellToSelection(cell) {
        // Skip if already selected
        if (cell.classList.contains('selected')) return;
        
        // Add selected class
        cell.classList.add('selected');
        
        // Add to global array
        window.selectedCells.push(cell);
        
        // Update last selected cell
        lastSelectedCell = cell;
    }
    
    /**
     * Clear cell selection
     */
    function clearCellSelection() {
        // Clear selected class from cells
        window.selectedCells.forEach(cell => {
            cell.classList.remove('selected');
        });
        
        // Clear global array
        window.selectedCells = [];
        
        // Clear selection overlay
        hideSelectionOverlay();
        
        // Reset selection state
        selectionStart = null;
        selectionEnd = null;
        lastSelectedCell = null;
        
        // Update UI
        updateSelectionInfo();
    }
    
    /**
     * Update the selection based on selection start and end
     */
    function updateSelection() {
        // Skip if either start or end is missing
        if (!selectionStart || !selectionEnd) return;
        
        // Clear current selection
        window.selectedCells.forEach(cell => {
            cell.classList.remove('selected');
        });
        window.selectedCells = [];
        
        // Calculate the rectangular range
        const minRow = Math.min(
            selectionStart.row !== null ? selectionStart.row : 0,
            selectionEnd.row !== null ? selectionEnd.row : 0
        );
        const maxRow = Math.max(
            selectionStart.row !== null ? selectionStart.row : getLastRowIndex(),
            selectionEnd.row !== null ? selectionEnd.row : getLastRowIndex()
        );
        const minCol = Math.min(selectionStart.col, selectionEnd.col);
        const maxCol = Math.max(selectionStart.col, selectionEnd.col);
        
        // Remember the selection range
        selectedRange = { minRow, maxRow, minCol, maxCol };
        
        // Select all cells in the range
        for (let r = minRow; r <= maxRow; r++) {
            for (let c = minCol; c <= maxCol; c++) {
                let cell;
                
                if (r === null) {
                    // Header row
                    cell = document.querySelector(`th[data-col="${c}"]`);
                } else {
                    cell = document.querySelector(`[data-row="${r}"][data-col="${c}"]`);
                }
                
                if (cell) {
                    addCellToSelection(cell);
                }
            }
        }
        
        // Show selection overlay
        showSelectionOverlay();
    }
    
    /**
     * Show a visual overlay for the current selection
     */
    function showSelectionOverlay() {
        if (!selectionOverlay) return;
        
        if (window.selectedCells.length <= 1) {
            // Hide overlay for single cell selection
            selectionOverlay.style.display = 'none';
            return;
        }
        
        // Find the position and size for the overlay
        const firstCell = window.selectedCells[0];
        const lastCell = window.selectedCells[window.selectedCells.length - 1];
        
        if (!firstCell || !lastCell) return;
        
        // Get positions
        const firstRect = firstCell.getBoundingClientRect();
        const lastRect = lastCell.getBoundingClientRect();
        
        // Calculate position and size
        const top = Math.min(firstRect.top, lastRect.top);
        const left = Math.min(firstRect.left, lastRect.left);
        const right = Math.max(firstRect.right, lastRect.right);
        const bottom = Math.max(firstRect.bottom, lastRect.bottom);
        
        // Set overlay styles
        selectionOverlay.style.display = 'block';
        selectionOverlay.style.top = (top + window.pageYOffset) + 'px';
        selectionOverlay.style.left = (left + window.pageXOffset) + 'px';
        selectionOverlay.style.width = (right - left) + 'px';
        selectionOverlay.style.height = (bottom - top) + 'px';
        
        // Add draggable handle if multiple cells are selected
        if (window.selectedCells.length > 1) {
            addDragHandleToOverlay();
        }
    }
    
    /**
     * Hide the selection overlay
     */
    function hideSelectionOverlay() {
        if (!selectionOverlay) return;
        selectionOverlay.style.display = 'none';
    }
    
    /**
     * Add a drag handle to the selection overlay for drag-fill
     */
    function addDragHandleToOverlay() {
        // First remove any existing handle
        const existingHandle = selectionOverlay.querySelector('.selection-drag-handle');
        if (existingHandle) existingHandle.remove();
        
        // Create drag handle element
        const dragHandle = document.createElement('div');
        dragHandle.className = 'selection-drag-handle';
        selectionOverlay.appendChild(dragHandle);
        
        // Add drag events
        dragHandle.addEventListener('mousedown', startDragFill);
        dragHandle.addEventListener('touchstart', startDragFillMobile, { passive: false });
    }
    
    /**
     * Start drag-fill operation
     */
    function startDragFill(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // Implement later - used for filling content to adjacent cells
        // by dragging the selection handle
    }
    
    /**
     * Start drag-fill on mobile
     */
    function startDragFillMobile(e) {
        e.preventDefault();
        
        // Implement later - mobile version of drag-fill
    }
    
    /**
     * Update selection info display
     */
    function updateSelectionInfo() {
        const infoElement = document.getElementById('selectedCellInfo');
        if (!infoElement) return;
        
        if (window.selectedCells.length === 0) {
            infoElement.textContent = '';
            return;
        }
        
        let displayText = '';
        
        if (window.selectedCells.length === 1) {
            // Single cell selected - show cell reference
            const cell = window.selectedCells[0];
            const row = parseInt(cell.getAttribute('data-row'));
            const col = parseInt(cell.getAttribute('data-col'));
            
            if (!isNaN(col)) {
                const colLetter = getColumnLetter(col);
                if (!isNaN(row)) {
                    // Normal cell
                    displayText = `${colLetter}${row + 1}`;
                } else {
                    // Header cell
                    displayText = `${colLetter}`;
                }
            }
        } else {
            // Multiple cells - show range and count
            const bounds = getSelectionBoundaries();
            const topLeft = `${getColumnLetter(bounds.minCol)}${bounds.minRow + 1}`;
            const bottomRight = `${getColumnLetter(bounds.maxCol)}${bounds.maxRow + 1}`;
            const count = (bounds.maxRow - bounds.minRow + 1) * (bounds.maxCol - bounds.minCol + 1);
            
            displayText = `${topLeft}:${bottomRight} (${count} cells)`;
        }
        
        infoElement.textContent = displayText;
    }
    
    /**
     * Get the letter representation of a column index (0=A, 1=B, etc.)
     */
    function getColumnLetter(index) {
        let letter = '';
        
        while (index >= 0) {
            letter = String.fromCharCode(65 + (index % 26)) + letter;
            index = Math.floor(index / 26) - 1;
        }
        
        return letter;
    }
    
    /**
     * Get the boundaries of the current selection
     */
    function getSelectionBoundaries() {
        if (window.selectedCells.length === 0) {
            return { minRow: 0, minCol: 0, maxRow: 0, maxCol: 0 };
        }
        
        // If we have a stored range, use it
        if (selectedRange.minRow !== undefined) {
            return selectedRange;
        }
        
        // Otherwise calculate from selected cells
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
     * Get the index of the last row in the table
     */
    function getLastRowIndex() {
        const rows = document.querySelectorAll('#sheetTable tbody tr');
        if (rows.length === 0) return 0;
        
        const lastRow = rows[rows.length - 1];
        return parseInt(lastRow.getAttribute('data-row')) || 0;
    }
    
    /**
     * Get the index of the last column in the table
     */
    function getLastColumnIndex() {
        const headerCells = document.querySelectorAll('#sheetTable th');
        if (headerCells.length === 0) return 0;
        
        const lastCol = headerCells[headerCells.length - 1];
        return parseInt(lastCol.getAttribute('data-col')) || 0;
    }
    
    /**
     * Update the merge cells button state
     */
    function updateMergeCellButton() {
        const mergeCellsBtn = document.getElementById('mergeCellsBtn');
        if (!mergeCellsBtn) return;
        
        // Enable if multiple cells are selected
        if (window.selectedCells.length > 1) {
            mergeCellsBtn.classList.remove('disabled');
            mergeCellsBtn.removeAttribute('disabled');
        } else {
            mergeCellsBtn.classList.add('disabled');
            mergeCellsBtn.setAttribute('disabled', 'disabled');
        }
    }
    
    /**
     * Update formatting buttons based on the last selected cell
     */
    function updateFormattingButtons() {
        // Will be implemented to sync button states with the selected cell's formatting
        if (typeof window.updateFormattingUI === 'function') {
            window.updateFormattingUI();
        }
    }
    
    /**
     * Update formula bar with active cell content
     */
    function updateFormulaBar() {
        const formulaBar = document.getElementById('formulaBar') || document.getElementById('activeFormulaBar');
        if (!formulaBar) return;
        
        const cellRef = document.getElementById('cellReference') || document.getElementById('cellRef');
        
        if (window.selectedCells.length === 1) {
            const cell = window.selectedCells[0];
            const editableDiv = cell.querySelector('.editable-cell');
            
            if (editableDiv) {
                formulaBar.value = editableDiv.textContent || '';
            }
            
            // Update cell reference
            if (cellRef) {
                const row = parseInt(cell.getAttribute('data-row'));
                const col = parseInt(cell.getAttribute('data-col'));
                
                if (!isNaN(col)) {
                    const colLetter = getColumnLetter(col);
                    if (!isNaN(row)) {
                        cellRef.value = `${colLetter}${row + 1}`;
                    } else {
                        cellRef.value = colLetter;
                    }
                }
            }
        } else {
            // For multiple selection, clear the formula bar
            formulaBar.value = '';
            
            // Update cell reference to show selection range
            if (cellRef && window.selectedCells.length > 0) {
                const bounds = getSelectionBoundaries();
                const topLeft = `${getColumnLetter(bounds.minCol)}${bounds.minRow + 1}`;
                const bottomRight = `${getColumnLetter(bounds.maxCol)}${bounds.maxRow + 1}`;
                
                cellRef.value = `${topLeft}:${bottomRight}`;
            }
        }
    }
    
    // Expose functions globally
    window.excelSelection = {
        clearCellSelection,
        addCellToSelection,
        getSelectionBoundaries,
        getColumnLetter,
        updateSelectionInfo,
        initializeCellSelection
    };
});
