/**
 * Excel-like column resizing functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize column resize functionality
    initColumnResize();
    
    function initColumnResize() {
        // Add resize handles to all column headers
        addResizeHandles();
        
        // Load saved column widths
        loadColumnWidths();
    }
    
    /**
     * Add resize handles to column headers
     */
    function addResizeHandles() {
        // Get all column headers
        const headers = document.querySelectorAll('#sheetTable th, .excel-column-header');
        
        // Add resize handles
        headers.forEach(header => {
            // Check if handle already exists
            if (!header.querySelector('.col-resize-handle')) {
                const handle = document.createElement('div');
                handle.className = 'col-resize-handle';
                handle.setAttribute('data-col', header.getAttribute('data-col'));
                header.appendChild(handle);
                
                // Add event listener
                handle.addEventListener('mousedown', startResize);
            }
        });
    }
    
    /**
     * Start column resize
     */
    function startResize(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // Get column index
        const colIndex = this.getAttribute('data-col');
        
        // Get starting width and position
        const startX = e.clientX;
        const headCell = document.querySelector(`#sheetTable th[data-col="${colIndex}"]`);
        if (!headCell) return;
        
        const startWidth = headCell.offsetWidth;
        
        // Show resize guidance
        const guideLine = document.createElement('div');
        guideLine.className = 'resize-guide';
        guideLine.style.height = '9999px';
        guideLine.style.width = '2px';
        guideLine.style.backgroundColor = '#0d6efd';
        guideLine.style.position = 'absolute';
        guideLine.style.top = '0';
        guideLine.style.left = (e.clientX) + 'px';
        guideLine.style.zIndex = '9999';
        guideLine.style.pointerEvents = 'none';
        document.body.appendChild(guideLine);
        
        // Add mousemove and mouseup handlers
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
        
        // Handle mouse move
        function handleMouseMove(e) {
            // Calculate new width
            const diff = e.clientX - startX;
            const newWidth = Math.max(30, startWidth + diff);
            
            // Update guide line
            guideLine.style.left = (e.clientX) + 'px';
            
            // Show width indicator
            let indicator = document.getElementById('width-indicator');
            if (!indicator) {
                indicator = document.createElement('div');
                indicator.id = 'width-indicator';
                indicator.style.position = 'fixed';
                indicator.style.backgroundColor = '#212529';
                indicator.style.color = 'white';
                indicator.style.padding = '3px 6px';
                indicator.style.borderRadius = '4px';
                indicator.style.fontSize = '12px';
                indicator.style.zIndex = '9999';
                indicator.style.pointerEvents = 'none';
                document.body.appendChild(indicator);
            }
            
            // Position indicator
            indicator.textContent = `${Math.round(newWidth)}px`;
            indicator.style.left = (e.clientX + 10) + 'px';
            indicator.style.top = (e.clientY - 30) + 'px';
        }
        
        // Handle mouse up
        function handleMouseUp(e) {
            // Remove guide line
            if (guideLine.parentNode) {
                guideLine.parentNode.removeChild(guideLine);
            }
            
            // Remove indicator
            const indicator = document.getElementById('width-indicator');
            if (indicator && indicator.parentNode) {
                indicator.parentNode.removeChild(indicator);
            }
            
            // Calculate new width
            const diff = e.clientX - startX;
            const newWidth = Math.max(30, startWidth + diff);
            
            // Apply new width
            applyColumnWidth(colIndex, newWidth);
            
            // Remove event listeners
            document.removeEventListener('mousemove', handleMouseMove);
            document.removeEventListener('mouseup', handleMouseUp);
        }
    }
    
    /**
     * Apply column width
     */
    function applyColumnWidth(colIndex, width) {
        // Update header cells
        const headerCells = document.querySelectorAll(`#sheetTable th[data-col="${colIndex}"]`);
        headerCells.forEach(cell => {
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
        });
        
        // Update column headers
        const columnHeaders = document.querySelectorAll(`.excel-column-header[data-col="${colIndex}"]`);
        columnHeaders.forEach(header => {
            header.style.width = `${width}px`;
            header.style.minWidth = `${width}px`;
        });
        
        // Update data cells
        const dataCells = document.querySelectorAll(`#sheetTable td[data-col="${colIndex}"]`);
        dataCells.forEach(cell => {
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
        });
        
        // Save column widths
        saveColumnWidths();
    }
    
    /**
     * Save column widths to localStorage
     */
    function saveColumnWidths() {
        try {
            const table = document.getElementById('sheetTable');
            if (!table) return;
            
            const sheetName = table.getAttribute('data-sheet-name');
            if (!sheetName) return;
            
            const headerCells = document.querySelectorAll('#sheetTable th');
            const columnWidths = {};
            
            headerCells.forEach(cell => {
                const colIndex = cell.getAttribute('data-col');
                const width = cell.style.width;
                if (colIndex && width) {
                    columnWidths[colIndex] = width;
                }
            });
            
            localStorage.setItem(`excel_column_widths_${sheetName}`, JSON.stringify(columnWidths));
        } catch (error) {
            console.error('Error saving column widths:', error);
        }
    }
    
    /**
     * Load column widths from localStorage
     */
    function loadColumnWidths() {
        try {
            const table = document.getElementById('sheetTable');
            if (!table) return;
            
            const sheetName = table.getAttribute('data-sheet-name');
            if (!sheetName) return;
            
            const savedWidths = localStorage.getItem(`excel_column_widths_${sheetName}`);
            if (!savedWidths) return;
            
            const columnWidths = JSON.parse(savedWidths);
            
            // Apply saved widths
            Object.keys(columnWidths).forEach(colIndex => {
                const width = columnWidths[colIndex];
                if (width) {
                    const numericWidth = parseInt(width);
                    if (!isNaN(numericWidth)) {
                        applyColumnWidth(colIndex, numericWidth);
                    }
                }
            });
        } catch (error) {
            console.error('Error loading column widths:', error);
        }
    }
    
    // Initialize on load
    window.addEventListener('load', function() {
        // Re-initialize after a short delay to ensure the DOM is ready
        setTimeout(function() {
            initColumnResize();
        }, 100);
    });
    
    // Expose functions globally
    window.columnResize = {
        init: initColumnResize,
        applyColumnWidth: applyColumnWidth
    };
});
