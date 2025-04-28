/**
 * Performance optimizations for Excel-like functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Optimize event handlers and animations to improve performance
    optimizeEventHandlers();
    optimizeAnimations();
    enableHeaderSelection();
    
    /**
     * Optimize event handlers by using event delegation and throttling
     */
    function optimizeEventHandlers() {
        // Use event delegation instead of attaching handlers to each cell
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        // Throttle function to limit how often a function can execute
        function throttle(func, delay) {
            let lastCall = 0;
            return function(...args) {
                const now = Date.now();
                if (now - lastCall >= delay) {
                    lastCall = now;
                    func.apply(this, args);
                }
            };
        }
        
        // Throttle resize handlers
        const originalMousemoveHandler = window.handleResize;
        if (typeof originalMousemoveHandler === 'function') {
            window.handleResize = throttle(originalMousemoveHandler, 16); // ~60fps
        }
        
        // Throttle cell selection handling
        const originalExtendSelection = window.extendCellSelection;
        if (typeof originalExtendSelection === 'function') {
            window.extendCellSelection = throttle(originalExtendSelection, 16);
        }
        
        // Use passive event listeners where possible for better performance
        document.addEventListener('scroll', function() {
            // Performance optimized scroll handling
        }, { passive: true });
    }
    
    /**
     * Optimize animations by reducing reflows and repaints
     */
    function optimizeAnimations() {
        // Use transform and opacity for animations instead of top/left/width/height
        // This avoids layout recalculations (reflows) and performs better
        const style = document.createElement('style');
        style.textContent = `
            /* Use GPU-accelerated properties for better performance */
            .resize-preview, .row-resize-preview, .size-indicator {
                transform: translateZ(0);
                will-change: transform, opacity;
            }
            
            /* Reduce layout thrashing by avoiding properties that cause reflow */
            .editable-cell:focus, .cell-selection-indicator {
                transform: translateZ(0);
            }
            
            /* Optimize transitions */
            .selection-overlay {
                transition: transform 0.1s, opacity 0.1s;
                transform: translateZ(0);
                will-change: transform;
            }
        `;
        document.head.appendChild(style);
    }
    
    /**
     * Enable selection of header cells
     */
    function enableHeaderSelection() {
        // Make header cells selectable
        document.querySelectorAll('#sheetTable th').forEach(headerCell => {
            // Remove any click handlers that might interfere with selection
            const oldClone = headerCell.cloneNode(true);
            headerCell.parentNode.replaceChild(oldClone, headerCell);
            
            // Add selection capability
            oldClone.classList.add('selectable-header');
            oldClone.addEventListener('mousedown', function(e) {
                // Only if not clicking on a resize handle
                if (!e.target.classList.contains('col-resize-handle') && 
                    !e.target.classList.contains('row-resize-handle')) {
                    handleHeaderSelection(e, oldClone);
                }
            });
        });
        
        // Add the handler to column headers as well
        document.querySelectorAll('.excel-column-header').forEach(header => {
            if (!header.classList.contains('selectable-header')) {
                header.classList.add('selectable-header');
                header.addEventListener('mousedown', function(e) {
                    // Only if not clicking on a resize handle
                    if (!e.target.classList.contains('col-resize-handle')) {
                        handleHeaderSelection(e, header);
                    }
                });
            }
        });
    }
    
    /**
     * Handle header cell selection
     */
    function handleHeaderSelection(e, headerCell) {
        // If selection functionality exists
        if (typeof window.selectCell === 'function') {
            // If not holding Ctrl/Cmd, clear selection first
            if (!e.ctrlKey && !e.metaKey && typeof window.clearCellSelection === 'function') {
                window.clearCellSelection();
            }
            
            window.selectCell(headerCell);
        } else {
            // Basic selection
            if (!e.ctrlKey && !e.metaKey) {
                document.querySelectorAll('.selected').forEach(cell => {
                    cell.classList.remove('selected');
                });
                
                if (!window.selectedCells) window.selectedCells = [];
                window.selectedCells = [];
            }
            
            headerCell.classList.add('selected');
            if (!window.selectedCells) window.selectedCells = [];
            window.selectedCells.push(headerCell);
        }
        
        // Update UI after selection
        if (typeof window.updateSelectionInfo === 'function') {
            window.updateSelectionInfo();
        }
        
        // Don't propagate event to prevent text selection
        e.preventDefault();
        e.stopPropagation();
    }
    
    /**
     * Optimize selection rendering for better performance
     */
    window.optimizeSelectionRendering = function() {
        // Check if we should batch selection updates
        const selectionCount = window.selectedCells ? window.selectedCells.length : 0;
        
        // For large selections, use a more efficient approach
        if (selectionCount > 50) {
            // Instead of adding/removing classes, use a selection overlay
            const table = document.getElementById('sheetTable');
            if (!table) return;
            
            // Get bounds of the selection
            const bounds = window.getSelectionBoundaries ? 
                window.getSelectionBoundaries() : 
                calculateSelectionBounds();
                
            // Calculate physical bounds
            const minCell = document.querySelector(`[data-row="${bounds.minRow}"][data-col="${bounds.minCol}"]`);
            const maxCell = document.querySelector(`[data-row="${bounds.maxRow}"][data-col="${bounds.maxCol}"]`);
            
            if (minCell && maxCell) {
                const minRect = minCell.getBoundingClientRect();
                const maxRect = maxCell.getBoundingClientRect();
                
                // Create or update overlay
                let overlay = document.getElementById('batch-selection-overlay');
                if (!overlay) {
                    overlay = document.createElement('div');
                    overlay.id = 'batch-selection-overlay';
                    overlay.className = 'selection-overlay';
                    document.body.appendChild(overlay);
                }
                
                // Position overlay
                overlay.style.position = 'absolute';
                overlay.style.top = (minRect.top + window.scrollY) + 'px';
                overlay.style.left = (minRect.left + window.scrollX) + 'px';
                overlay.style.width = (maxRect.right - minRect.left) + 'px';
                overlay.style.height = (maxRect.bottom - minRect.top) + 'px';
                overlay.style.backgroundColor = 'rgba(33, 115, 70, 0.1)';
                overlay.style.border = '2px solid #217346';
                overlay.style.zIndex = '100';
                overlay.style.pointerEvents = 'none';
                overlay.style.display = 'block';
            }
        } else if (selectionCount <= 50) {
            // Hide any batch overlay for smaller selections
            const overlay = document.getElementById('batch-selection-overlay');
            if (overlay) {
                overlay.style.display = 'none';
            }
        }
    };
    
    /**
     * Calculate selection bounds manually if needed
     */
    function calculateSelectionBounds() {
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
});
