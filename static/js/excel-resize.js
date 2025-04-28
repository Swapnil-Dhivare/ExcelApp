/**
 * Excel-like column and row resizing functionality
 */

document.addEventListener('DOMContentLoaded', function() {
    // State for resize operations
    let isResizing = false;
    let resizeTarget = null;
    let startX = 0;
    let startY = 0;
    let startWidth = 0;
    let startHeight = 0;
    let resizeType = null; // 'column' or 'row'
    let currentCol = null;
    let resizePreview = null;
    
    // Initialize all resize functionality
    initializeResizeHandlers();
    
    function initializeResizeHandlers() {
        // Add resize preview element
        addResizePreview();
        
        // Add column resize handles to all headers
        addAllColumnResizeHandles();
        
        // Add row resize handles
        addAllRowResizeHandles();
        
        // Load any saved dimensions
        loadColumnWidths();
        loadRowHeights();
    }
    
    /**
     * Add resize preview element
     */
    function addResizePreview() {
        if (!document.getElementById('resize-preview')) {
            resizePreview = document.createElement('div');
            resizePreview.id = 'resize-preview';
            resizePreview.className = 'resize-preview';
            document.body.appendChild(resizePreview);
        }
    }
    
    /**
     * Add resize handles to all column headers
     */
    function addAllColumnResizeHandles() {
        // Add to table header cells
        document.querySelectorAll('#sheetTable th, .excel-column-header').forEach(cell => {
            if (!cell.querySelector('.col-resize-handle')) {
                const handle = document.createElement('div');
                handle.className = 'col-resize-handle';
                handle.setAttribute('data-col', cell.getAttribute('data-col'));
                cell.appendChild(handle);
                
                // Add events
                handle.addEventListener('mousedown', startColumnResize);
                handle.addEventListener('touchstart', startColumnResizeMobile, { passive: false });
                handle.addEventListener('dblclick', autoSizeColumn);
            }
        });
    }
    
    /**
     * Add resize handles to all rows
     */
    function addAllRowResizeHandles() {
        // Add to table rows
        document.querySelectorAll('#sheetTable tr[data-row]').forEach(row => {
            if (!row.querySelector('.row-resize-handle')) {
                const handle = document.createElement('div');
                handle.className = 'row-resize-handle';
                handle.setAttribute('data-row', row.getAttribute('data-row'));
                row.appendChild(handle);
                
                // Add events
                handle.addEventListener('mousedown', startRowResize);
                handle.addEventListener('touchstart', startRowResizeMobile, { passive: false });
            }
        });
    }
    
    /**
     * Start column resize operation
     */
    function startColumnResize(e) {
        e.preventDefault();
        e.stopPropagation();
        
        isResizing = true;
        resizeType = 'column';
        resizeTarget = e.target;
        currentCol = parseInt(resizeTarget.getAttribute('data-col'));
        
        // Add resize class to body for cursor styling
        document.body.classList.add('resizing', 'column-resize');
        
        // Find the header cell
        const headerCell = document.querySelector(`#sheetTable th[data-col="${currentCol}"]`) || 
                          document.querySelector(`.excel-column-header[data-col="${currentCol}"]`);
        if (!headerCell) return;
        
        // Store starting position and width
        startX = e.clientX;
        startWidth = headerCell.offsetWidth;
        
        // Show resize preview
        showResizeLine(e.clientX, 'vertical');
        showSizeIndicator(e.clientX, e.clientY, `${startWidth}px`);
        
        // Add event listeners for move and up
        document.addEventListener('mousemove', handleResize);
        document.addEventListener('mouseup', stopResize);
    }
    
    /**
     * Start row resize operation
     */
    function startRowResize(e) {
        e.preventDefault();
        e.stopPropagation();
        
        isResizing = true;
        resizeType = 'row';
        resizeTarget = e.target;
        const rowIndex = parseInt(resizeTarget.getAttribute('data-row'));
        
        // Add resize class to body for cursor styling
        document.body.classList.add('resizing', 'row-resize');
        
        // Find the row
        const row = document.querySelector(`#sheetTable tr[data-row="${rowIndex}"]`);
        if (!row) return;
        
        // Store starting position and height
        startY = e.clientY;
        startHeight = row.offsetHeight;
        
        // Show resize preview
        showResizeLine(e.clientY, 'horizontal');
        showSizeIndicator(e.clientX, e.clientY, `${startHeight}px`);
        
        // Add event listeners for move and up
        document.addEventListener('mousemove', handleResize);
        document.addEventListener('mouseup', stopResize);
    }
    
    /**
     * Handle touch start for column resize
     */
    function startColumnResizeMobile(e) {
        if (!e.touches || !e.touches[0]) return;
        
        e.preventDefault();
        
        isResizing = true;
        resizeType = 'column';
        resizeTarget = e.target;
        currentCol = parseInt(resizeTarget.getAttribute('data-col'));
        
        // Add resize class to body for cursor styling
        document.body.classList.add('resizing', 'column-resize');
        
        // Find the header cell
        const headerCell = document.querySelector(`#sheetTable th[data-col="${currentCol}"]`) || 
                          document.querySelector(`.excel-column-header[data-col="${currentCol}"]`);
        if (!headerCell) return;
        
        // Store starting position and width
        startX = e.touches[0].clientX;
        startWidth = headerCell.offsetWidth;
        
        // Show resize preview
        showResizeLine(e.touches[0].clientX, 'vertical');
        showSizeIndicator(e.touches[0].clientX, e.touches[0].clientY, `${startWidth}px`);
        
        // Add event listeners for move and up
        document.addEventListener('touchmove', handleResizeMobile, { passive: false });
        document.addEventListener('touchend', stopResizeMobile);
    }
    
    /**
     * Handle touch start for row resize
     */
    function startRowResizeMobile(e) {
        if (!e.touches || !e.touches[0]) return;
        
        e.preventDefault();
        
        isResizing = true;
        resizeType = 'row';
        resizeTarget = e.target;
        const rowIndex = parseInt(resizeTarget.getAttribute('data-row'));
        
        // Add resize class to body for cursor styling
        document.body.classList.add('resizing', 'row-resize');
        
        // Find the row
        const row = document.querySelector(`#sheetTable tr[data-row="${rowIndex}"]`);
        if (!row) return;
        
        // Store starting position and height
        startY = e.touches[0].clientY;
        startHeight = row.offsetHeight;
        
        // Show resize preview
        showResizeLine(e.touches[0].clientY, 'horizontal');
        showSizeIndicator(e.touches[0].clientX, e.touches[0].clientY, `${startHeight}px`);
        
        // Add event listeners for move and up
        document.addEventListener('touchmove', handleResizeMobile, { passive: false });
        document.addEventListener('touchend', stopResizeMobile);
    }
    
    /**
     * Handle resize during mouse movement
     */
    function handleResize(e) {
        if (!isResizing || !resizeTarget) return;
        
        if (resizeType === 'column') {
            const diff = e.clientX - startX;
            const newWidth = Math.max(30, startWidth + diff); // Minimum width of 30px
            
            // Update resize line position
            updateResizeLine(e.clientX);
            
            // Update the size indicator
            showSizeIndicator(e.clientX, e.clientY, `${newWidth}px`);
        } else if (resizeType === 'row') {
            const diff = e.clientY - startY;
            const newHeight = Math.max(20, startHeight + diff); // Minimum height of 20px
            
            // Update resize line position
            updateResizeLine(e.clientY);
            
            // Update the size indicator
            showSizeIndicator(e.clientX, e.clientY, `${newHeight}px`);
        }
    }
    
    /**
     * Handle resize for mobile touch events
     */
    function handleResizeMobile(e) {
        if (!isResizing || !resizeTarget || !e.touches || !e.touches[0]) return;
        
        e.preventDefault();
        
        if (resizeType === 'column') {
            const diff = e.touches[0].clientX - startX;
            const newWidth = Math.max(30, startWidth + diff);
            
            // Update resize line position
            updateResizeLine(e.touches[0].clientX);
            
            // Update the size indicator
            showSizeIndicator(e.touches[0].clientX, e.touches[0].clientY, `${newWidth}px`);
        } else if (resizeType === 'row') {
            const diff = e.touches[0].clientY - startY;
            const newHeight = Math.max(20, startHeight + diff);
            
            // Update resize line position
            updateResizeLine(e.touches[0].clientY);
            
            // Update the size indicator
            showSizeIndicator(e.touches[0].clientX, e.touches[0].clientY, `${newHeight}px`);
        }
    }
    
    /**
     * Stop resize operation
     */
    function stopResize(e) {
        if (!isResizing) return;
        
        if (resizeType === 'column' && currentCol !== null) {
            const diff = e.clientX - startX;
            const newWidth = Math.max(30, startWidth + diff);
            
            // Apply the new width with smooth animation
            applyColumnWidth(currentCol, newWidth);
        } else if (resizeType === 'row') {
            const rowIndex = parseInt(resizeTarget.getAttribute('data-row'));
            const diff = e.clientY - startY;
            const newHeight = Math.max(20, startHeight + diff);
            
            // Apply the new height with smooth animation
            applyRowHeight(rowIndex, newHeight);
        }
        
        // Clean up
        isResizing = false;
        resizeTarget = null;
        currentCol = null;
        
        // Remove event listeners
        document.removeEventListener('mousemove', handleResize);
        document.removeEventListener('mouseup', stopResize);
        
        // Remove resize styling
        document.body.classList.remove('resizing', 'column-resize', 'row-resize');
        
        // Hide resize visuals
        removeResizeLine();
        removeSizeIndicator();
        
        // Force cleanup any lingering indicators
        setTimeout(() => {
            document.querySelectorAll('.resize-line, .size-indicator').forEach(el => {
                if (el && el.parentNode) el.parentNode.removeChild(el);
            });
        }, 100);
    }
    
    /**
     * Stop resize for mobile
     */
    function stopResizeMobile() {
        if (!isResizing) return;
        
        if (resizeType === 'column' && currentCol !== null) {
            // Get the final width from the resize line position
            const resizeLine = document.querySelector('.resize-line');
            const newWidth = Math.max(30, startWidth + (parseInt(resizeLine?.style.left) - startX));
            
            // Apply the new width
            applyColumnWidth(currentCol, newWidth);
        } else if (resizeType === 'row') {
            const rowIndex = parseInt(resizeTarget.getAttribute('data-row'));
            const resizeLine = document.querySelector('.resize-line');
            const newHeight = Math.max(20, startHeight + (parseInt(resizeLine?.style.top) - startY));
            
            // Apply the new height
            applyRowHeight(rowIndex, newHeight);
        }
        
        // Clean up
        isResizing = false;
        resizeTarget = null;
        currentCol = null;
        
        // Remove event listeners
        document.removeEventListener('touchmove', handleResizeMobile);
        document.removeEventListener('touchend', stopResizeMobile);
        
        // Remove resize styling
        document.body.classList.remove('resizing', 'column-resize', 'row-resize');
        
        // Hide resize visuals
        removeResizeLine();
        removeSizeIndicator();
        
        // Force cleanup any lingering indicators
        setTimeout(() => {
            document.querySelectorAll('.resize-line, .size-indicator').forEach(el => {
                if (el && el.parentNode) el.parentNode.removeChild(el);
            });
        }, 100);
    }
    
    /**
     * Apply column width with smooth transition
     */
    function applyColumnWidth(colIndex, width) {
        // Update column header styles with transition
        document.querySelectorAll(`#sheetTable th[data-col="${colIndex}"], .excel-column-header[data-col="${colIndex}"]`).forEach(cell => {
            cell.style.transition = 'width 0.15s ease-in-out';
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
            cell.style.maxWidth = `${width}px`;
            
            // Remove transition after animation completes
            setTimeout(() => {
                cell.style.transition = '';
            }, 200);
        });
        
        // Update data cells
        document.querySelectorAll(`#sheetTable td[data-col="${colIndex}"]`).forEach(cell => {
            cell.style.transition = 'width 0.15s ease-in-out';
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
            cell.style.maxWidth = `${width}px`;
            
            // Remove transition after animation completes
            setTimeout(() => {
                cell.style.transition = '';
            }, 200);
        });
        
        // Save to localStorage
        saveColumnWidths();
    }
    
    /**
     * Apply row height with smooth transition
     */
    function applyRowHeight(rowIndex, height) {
        const row = document.querySelector(`#sheetTable tr[data-row="${rowIndex}"]`);
        if (!row) return;
        
        // Set row height with transition
        row.style.transition = 'height 0.15s ease-in-out';
        row.style.height = `${height}px`;
        
        // Update all cells in the row
        row.querySelectorAll('td, th').forEach(cell => {
            cell.style.transition = 'height 0.15s ease-in-out';
            cell.style.height = `${height}px`;
            
            // Remove transition after animation completes
            setTimeout(() => {
                cell.style.transition = '';
            }, 200);
        });
        
        // Remove transition after animation completes
        setTimeout(() => {
            row.style.transition = '';
        }, 200);
        
        // Save to localStorage
        saveRowHeights();
    }
    
    /**
     * Show resize guide line
     */
    function showResizeLine(position, orientation) {
        // Remove any existing line
        removeResizeLine();
        
        // Create new resize line
        const resizeLine = document.createElement('div');
        resizeLine.className = 'resize-line';
        
        if (orientation === 'vertical') {
            resizeLine.style.height = '100vh';
            resizeLine.style.top = '0';
            resizeLine.style.left = `${position}px`;
            resizeLine.style.width = '2px';
            resizeLine.style.cursor = 'col-resize';
        } else {
            resizeLine.style.width = '100vw';
            resizeLine.style.left = '0';
            resizeLine.style.top = `${position}px`;
            resizeLine.style.height = '2px';
            resizeLine.style.cursor = 'row-resize';
        }
        
        document.body.appendChild(resizeLine);
    }
    
    /**
     * Update resize line position
     */
    function updateResizeLine(position) {
        const resizeLine = document.querySelector('.resize-line');
        if (!resizeLine) return;
        
        if (resizeType === 'column') {
            resizeLine.style.left = `${position}px`;
        } else {
            resizeLine.style.top = `${position}px`;
        }
    }
    
    /**
     * Remove resize line
     */
    function removeResizeLine() {
        document.querySelectorAll('.resize-line').forEach(line => {
            if (line && line.parentNode) {
                line.parentNode.removeChild(line);
            }
        });
    }
    
    /**
     * Show size indicator during resize
     */
    function showSizeIndicator(x, y, text) {
        // Remove existing indicators
        removeSizeIndicator();
        
        // Create new indicator
        const indicator = document.createElement('div');
        indicator.className = 'size-indicator';
        indicator.textContent = text;
        
        // Position indicator
        indicator.style.top = `${y - 30}px`;
        indicator.style.left = `${x + 10}px`;
        
        document.body.appendChild(indicator);
    }
    
    /**
     * Remove size indicator
     */
    function removeSizeIndicator() {
        document.querySelectorAll('.size-indicator').forEach(indicator => {
            if (indicator && indicator.parentNode) {
                indicator.parentNode.removeChild(indicator);
            }
        });
    }
    
    /**
     * Auto-size column based on content
     */
    function autoSizeColumn(e) {
        if (e) {
            e.preventDefault();
            e.stopPropagation();
        }
        
        // Get column index
        let colIndex;
        if (e && e.target) {
            colIndex = parseInt(e.target.getAttribute('data-col'));
        } else if (currentCol !== null) {
            colIndex = currentCol;
        } else {
            return;
        }
        
        // Get all cells in this column
        const headerCells = document.querySelectorAll(`#sheetTable th[data-col="${colIndex}"]`);
        const dataCells = document.querySelectorAll(`#sheetTable td[data-col="${colIndex}"]`);
        
        // Calculate maximum content width
        let maxWidth = 0;
        const tempDiv = document.createElement('div');
        tempDiv.style.position = 'absolute';
        tempDiv.style.visibility = 'hidden';
        tempDiv.style.whiteSpace = 'nowrap';
        tempDiv.style.fontFamily = window.getComputedStyle(document.body).fontFamily;
        tempDiv.style.fontSize = window.getComputedStyle(document.body).fontSize;
        document.body.appendChild(tempDiv);
        
        // Check header width
        headerCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                tempDiv.textContent = editableDiv.textContent;
                const width = tempDiv.offsetWidth + 25;  // Add padding
                maxWidth = Math.max(maxWidth, width);
            }
        });
        
        // Check data cell widths
        dataCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                tempDiv.textContent = editableDiv.textContent;
                const width = tempDiv.offsetWidth + 25;  // Add padding
                maxWidth = Math.max(maxWidth, width);
            }
        });
        
        document.body.removeChild(tempDiv);
        
        // Set minimum width
        maxWidth = Math.max(maxWidth, 60);
        
        // Apply the calculated width
        applyColumnWidth(colIndex, maxWidth);
        
        // Show notification
        showToast(`Column auto-sized to ${maxWidth}px`, 'success');
    }
    
    /**
     * Save column widths to localStorage
     */
    function saveColumnWidths() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
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
        
        try {
            localStorage.setItem(`excel_column_widths_${sheetName}`, JSON.stringify(columnWidths));
        } catch (e) {
            console.error('Error saving column widths:', e);
        }
    }
    
    /**
     * Save row heights to localStorage
     */
    function saveRowHeights() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        const rows = document.querySelectorAll('#sheetTable tr[data-row]');
        const rowHeights = {};
        
        rows.forEach(row => {
            const rowIndex = row.getAttribute('data-row');
            const height = row.style.height;
            if (rowIndex && height) {
                rowHeights[rowIndex] = height;
            }
        });
        
        try {
            localStorage.setItem(`excel_row_heights_${sheetName}`, JSON.stringify(rowHeights));
        } catch (e) {
            console.error('Error saving row heights:', e);
        }
    }
    
    /**
     * Load column widths from localStorage
     */
    function loadColumnWidths() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        try {
            const savedWidths = localStorage.getItem(`excel_column_widths_${sheetName}`);
            if (!savedWidths) return;
            
            const columnWidths = JSON.parse(savedWidths);
            
            Object.keys(columnWidths).forEach(colIndex => {
                if (columnWidths[colIndex]) {
                    // Extract numeric width value from string like "100px"
                    const width = columnWidths[colIndex];
                    const numericWidth = parseInt(width);
                    if (!isNaN(numericWidth)) {
                        // Apply width without animation
                        applyColumnWidthImmediate(colIndex, numericWidth);
                    }
                }
            });
        } catch (e) {
            console.error('Error loading column widths:', e);
        }
    }
    
    /**
     * Load row heights from localStorage
     */
    function loadRowHeights() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        try {
            const savedHeights = localStorage.getItem(`excel_row_heights_${sheetName}`);
            if (!savedHeights) return;
            
            const rowHeights = JSON.parse(savedHeights);
            
            Object.keys(rowHeights).forEach(rowIndex => {
                if (rowHeights[rowIndex]) {
                    // Extract numeric height value from string like "30px"
                    const height = rowHeights[rowIndex];
                    const numericHeight = parseInt(height);
                    if (!isNaN(numericHeight)) {
                        // Apply height without animation
                        applyRowHeightImmediate(rowIndex, numericHeight);
                    }
                }
            });
        } catch (e) {
            console.error('Error loading row heights:', e);
        }
    }
    
    /**
     * Apply column width immediately (no transition)
     */
    function applyColumnWidthImmediate(colIndex, width) {
        // Update column header styles without transition
        document.querySelectorAll(`#sheetTable th[data-col="${colIndex}"], .excel-column-header[data-col="${colIndex}"]`).forEach(cell => {
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
            cell.style.maxWidth = `${width}px`;
        });
        
        // Update data cells
        document.querySelectorAll(`#sheetTable td[data-col="${colIndex}"]`).forEach(cell => {
            cell.style.width = `${width}px`;
            cell.style.minWidth = `${width}px`;
            cell.style.maxWidth = `${width}px`;
        });
    }
    
    /**
     * Apply row height immediately (no transition)
     */
    function applyRowHeightImmediate(rowIndex, height) {
        const row = document.querySelector(`#sheetTable tr[data-row="${rowIndex}"]`);
        if (!row) return;
        
        // Set row height without transition
        row.style.height = `${height}px`;
        
        // Update all cells in the row
        row.querySelectorAll('td, th').forEach(cell => {
            cell.style.height = `${height}px`;
        });
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
    
    // Expose APIs for other modules to use
    window.excelResize = {
        initializeResizeHandlers,
        applyColumnWidth,
        applyRowHeight,
        autoSizeColumn
    };
    
    // Re-initialize resize handlers when the DOM changes (like when cells are added)
    const observer = new MutationObserver(function(mutations) {
        initializeResizeHandlers();
    });
    
    const sheetTable = document.getElementById('sheetTable');
    if (sheetTable) {
        observer.observe(sheetTable, { childList: true, subtree: true });
    }
});
