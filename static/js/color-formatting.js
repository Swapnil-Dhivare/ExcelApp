/**
 * Color and Formatting Functionality for Excel Generator
 */

document.addEventListener('DOMContentLoaded', function() {
    // Initialize color indicators
    updateColorIndicators();
    
    // Attach click handlers to buttons if they exist
    attachColorPickerHandlers();
    
    // Enable header cell selection
    enableHeaderSelection();
});

/**
 * Updates the color indicators on buttons to match the current color
 */
function updateColorIndicators() {
    // Define the button and picker pairs
    const pairs = [
        { buttonId: 'bgColorBtn', pickerId: 'bgColorPicker', defaultColor: '#ffffff' },
        { buttonId: 'headerColorBtn', pickerId: 'headerColorPicker', defaultColor: '#f8f9fa' },
        { buttonId: 'textColorBtn', pickerId: 'textColorPicker', defaultColor: '#000000' },
        { buttonId: 'rowColorBtn', pickerId: 'rowColorPicker', defaultColor: '#e9ecef' },
        { buttonId: 'columnColorBtn', pickerId: 'columnColorPicker', defaultColor: '#e9ecef' }
    ];
    
    // Process each pair
    pairs.forEach(pair => {
        const button = document.getElementById(pair.buttonId);
        const picker = document.getElementById(pair.pickerId);
        
        if (button && picker) {
            // Get or create the color indicator
            let indicator = button.querySelector('.color-indicator');
            
            if (!indicator) {
                indicator = document.createElement('span');
                indicator.className = 'color-indicator';
                button.appendChild(indicator);
            }
            
            // Set the picker's default value
            picker.value = pair.defaultColor;
            
            // Update the indicator color
            indicator.style.backgroundColor = picker.value;
        }
    });
}

/**
 * Attaches event handlers to color picker buttons
 */
function attachColorPickerHandlers() {
    // Define the button and picker pairs
    const pairs = [
        { buttonId: 'bgColorBtn', pickerId: 'bgColorPicker', handler: applyCellBackgroundColor },
        { buttonId: 'headerColorBtn', pickerId: 'headerColorPicker', handler: applyHeaderBackgroundColor },
        { buttonId: 'textColorBtn', pickerId: 'textColorPicker', handler: applyTextColor },
        { buttonId: 'rowColorBtn', pickerId: 'rowColorPicker', handler: applyRowBackgroundColor },
        { buttonId: 'columnColorBtn', pickerId: 'columnColorPicker', handler: applyColumnBackgroundColor }
    ];
    
    // Process each pair
    pairs.forEach(pair => {
        const button = document.getElementById(pair.buttonId);
        const picker = document.getElementById(pair.pickerId);
        
        if (button && picker) {
            // Remove any existing event listeners (to avoid duplicates)
            const newButton = button.cloneNode(true);
            button.parentNode.replaceChild(newButton, button);
            
            // Add new event listener - important: use mousedown instead of click
            // This preserves cell selection
            newButton.addEventListener('mousedown', function(e) {
                e.preventDefault(); // Prevent focus change
                picker.click();
            });
            
            // Input event for color picker
            picker.addEventListener('input', function() {
                pair.handler(this.value);
                
                // Update color indicator
                const indicator = newButton.querySelector('.color-indicator') || 
                                  document.createElement('span');
                
                if (!newButton.querySelector('.color-indicator')) {
                    indicator.className = 'color-indicator';
                    newButton.appendChild(indicator);
                }
                
                indicator.style.backgroundColor = this.value;
            });
        }
    });
}

/**
 * Enable selection of header cells
 */
function enableHeaderSelection() {
    // Get all header cells
    const headerCells = document.querySelectorAll('th.sheet-header');
    
    headerCells.forEach(header => {
        // Add the click event listener for selection
        header.addEventListener('click', function(e) {
            // Skip if clicked on a child element that handles its own events
            if (e.target !== this && 
                !e.target.classList.contains('editable-cell')) {
                return;
            }
            
            // Check if Ctrl key is pressed for multi-selection
            if (e.ctrlKey || e.metaKey) {
                // Multi-select
                if (window.selectedCells && window.selectedCells.includes(this)) {
                    // Deselect if already selected
                    this.classList.remove('selected');
                    const index = window.selectedCells.indexOf(this);
                    if (index !== -1) {
                        window.selectedCells.splice(index, 1);
                    }
                } else {
                    // Add to selection
                    this.classList.add('selected');
                    if (window.selectedCells) {
                        window.selectedCells.push(this);
                    } else {
                        window.selectedCells = [this];
                    }
                }
            } else {
                // Single selection - clear existing selection
                if (window.selectedCells) {
                    window.selectedCells.forEach(cell => {
                        if (cell) cell.classList.remove('selected');
                    });
                }
                
                // Select this cell
                this.classList.add('selected');
                window.selectedCells = [this];
            }
            
            e.stopPropagation(); // Prevent bubbling
        });
        
        // Make editable-cell inside header also trigger selection
        const editableCell = header.querySelector('.editable-cell');
        if (editableCell) {
            editableCell.addEventListener('click', function(e) {
                const headerCell = this.closest('th');
                
                // If Ctrl is pressed, toggle in selection
                if (e.ctrlKey || e.metaKey) {
                    if (window.selectedCells && window.selectedCells.includes(headerCell)) {
                        // Deselect
                        headerCell.classList.remove('selected');
                        const index = window.selectedCells.indexOf(headerCell);
                        if (index !== -1) {
                            window.selectedCells.splice(index, 1);
                        }
                    } else {
                        // Add to selection
                        headerCell.classList.add('selected');
                        if (window.selectedCells) {
                            window.selectedCells.push(headerCell);
                        } else {
                            window.selectedCells = [headerCell];
                        }
                    }
                } else {
                    // Single selection
                    if (window.selectedCells) {
                        window.selectedCells.forEach(cell => {
                            if (cell) cell.classList.remove('selected');
                        });
                    }
                    
                    headerCell.classList.add('selected');
                    window.selectedCells = [headerCell];
                }
                
                // Don't stop propagation to allow editing, but prevent default behavior
                // e.stopPropagation();
            });
        }
    });
}

/**
 * Applies background color to selected cells
 */
function applyCellBackgroundColor(color) {
    if (!window.selectedCells || window.selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    window.selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.backgroundColor = color;
        }
    });
    
    if (typeof window.saveSheetData === 'function') {
        window.saveSheetData();
    }
    
    showToast('Background color applied', 'success');
}

/**
 * Applies background color to header cells
 */
function applyHeaderBackgroundColor(color) {
    if (!window.selectedCells || window.selectedCells.length === 0) {
        showToast('Please select at least one header cell', 'error');
        return;
    }
    
    let headerCount = 0;
    window.selectedCells.forEach(cell => {
        // Apply to any cell - not just header cells
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.backgroundColor = color;
            headerCount++;
        }
    });
    
    if (typeof window.saveSheetData === 'function') {
        window.saveSheetData();
    }
    
    showToast(`Applied color to ${headerCount} cell(s)`, 'success');
}

/**
 * Applies text color to selected cells
 */
function applyTextColor(color) {
    if (!window.selectedCells || window.selectedCells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    window.selectedCells.forEach(cell => {
        const editableDiv = cell.querySelector('.editable-cell');
        if (editableDiv) {
            editableDiv.style.color = color;
        }
    });
    
    if (typeof window.saveSheetData === 'function') {
        window.saveSheetData();
    }
    
    showToast('Text color applied', 'success');
}

/**
 * Applies background color to entire rows
 */
function applyRowBackgroundColor(color) {
    if (!window.selectedCells || window.selectedCells.length === 0) {
        showToast('Please select a cell in the row', 'error');
        return;
    }
    
    const rowIndices = [...new Set(
        window.selectedCells
            .filter(cell => cell.hasAttribute('data-row'))
            .map(cell => parseInt(cell.getAttribute('data-row')))
    )];
    
    if (rowIndices.length === 0) {
        showToast('Could not identify row', 'error');
        return;
    }
    
    rowIndices.forEach(rowIndex => {
        const rowCells = document.querySelectorAll(`td[data-row="${rowIndex}"]`);
        rowCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        });
    });
    
    if (typeof window.saveSheetData === 'function') {
        window.saveSheetData();
    }
    
    showToast(`Applied color to ${rowIndices.length} row(s)`, 'success');
}

/**
 * Applies background color to entire columns
 */
function applyColumnBackgroundColor(color) {
    if (!window.selectedCells || window.selectedCells.length === 0) {
        showToast('Please select a cell in the column', 'error');
        return;
    }
    
    const colIndices = [...new Set(
        window.selectedCells
            .filter(cell => cell.hasAttribute('data-col'))
            .map(cell => parseInt(cell.getAttribute('data-col')))
    )];
    
    if (colIndices.length === 0) {
        showToast('Could not identify column', 'error');
        return;
    }
    
    colIndices.forEach(colIndex => {
        // Apply to header
        const headerCell = document.querySelector(`th[data-col="${colIndex}"]`);
        if (headerCell) {
            const editableDiv = headerCell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        }
        
        // Apply to data cells
        const colCells = document.querySelectorAll(`td[data-col="${colIndex}"]`);
        colCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
            }
        });
    });
    
    if (typeof window.saveSheetData === 'function') {
        window.saveSheetData();
    }
    
    showToast(`Applied color to ${colIndices.length} column(s)`, 'success');
}

/**
 * Shows a toast notification
 */
function showToast(message, type = 'info') {
    // Check if there's an existing showToast function
    if (typeof window.showToast === 'function') {
        window.showToast(message, type);
        return;
    }
    
    const toastContainer = document.querySelector('.toast-container');
    if (!toastContainer) return;
    
    const toast = document.createElement('div');
    toast.className = `toast show mb-2`;
    toast.role = 'alert';
    toast.setAttribute('aria-live', 'assertive');
    toast.setAttribute('aria-atomic', 'true');
    
    const bgClass = type === 'error' ? 'bg-danger' : 
                  type === 'success' ? 'bg-success' : 'bg-info';
    
    toast.innerHTML = `
        <div class="toast-header ${bgClass} text-white">
            <strong class="me-auto">Excel Generator</strong>
            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="toast" aria-label="Close"></button>
        </div>
        <div class="toast-body bg-light">
            ${message}
        </div>
    `;
    
    toastContainer.appendChild(toast);
    
    toast.querySelector('.btn-close').addEventListener('click', function() {
        toast.remove();
    });
    
    setTimeout(() => {
        toast.remove();
    }, 3000);
}
