/**
 * Color picker functionality for Excel app
 * This file handles all color picker functionalities including initialization and events
 */

document.addEventListener('DOMContentLoaded', function() {
    // Initialize color picker functionality
    initColorPickers();
    
    // Handle format panel toggle
    const toggleFormatPanelBtn = document.getElementById('toggleFormatPanel');
    if (toggleFormatPanelBtn) {
        const formatPanel = document.querySelector('.formatting-panel');
        toggleFormatPanelBtn.addEventListener('click', function() {
            if (formatPanel) {
                formatPanel.style.display = formatPanel.style.display === 'none' ? 'block' : 'none';
            }
        });
    }
    
    // Debug selected cells
    setInterval(() => {
        if (window.selectedCells && window.selectedCells.length > 0) {
            console.log(`Selected cells count: ${window.selectedCells.length}`);
        }
    }, 5000);
});

// Initialize all color pickers in the application
function initColorPickers() {
    console.log('Initializing color pickers');
    
    // Background color picker functionality
    setupColorPicker('bgColorBtn', 'bgColorPicker', applyBackgroundColor);
    
    // Header color picker functionality
    setupColorPicker('headerColorBtn', 'headerColorPicker', applyHeaderColor);
    
    // Text color picker functionality
    setupColorPicker('textColorBtn', 'textColorPicker', applyTextColor);
    
    // Create color pickers if they don't exist
    createColorPickerIfNeeded('bgColorPicker', 'bgColorBtn');
    createColorPickerIfNeeded('headerColorPicker', 'headerColorBtn');
    createColorPickerIfNeeded('textColorPicker', 'textColorBtn');
}

// Helper function to create a color picker if it doesn't exist
function createColorPickerIfNeeded(pickerId, buttonId) {
    if (!document.getElementById(pickerId)) {
        const button = document.getElementById(buttonId);
        if (!button) return;
        
        console.log(`Creating color picker for ${buttonId}`);
        
        const picker = document.createElement('input');
        picker.type = 'color';
        picker.id = pickerId;
        picker.style.display = 'none';
        picker.value = '#ffffff'; // Default color
        
        // Add to document body
        document.body.appendChild(picker);
        
        console.log(`Created color picker: ${pickerId}`);
    }
}

// Setup a specific color picker button with its input and handler
function setupColorPicker(btnId, pickerId, handler) {
    const colorBtn = document.getElementById(btnId);
    
    if (colorBtn) {
        console.log(`Setting up color picker button: ${btnId}`);
        
        // Add click event to button to trigger the picker
        colorBtn.addEventListener('click', function(e) {
            console.log(`${btnId} clicked`);
            e.preventDefault();
            e.stopPropagation();
            
            // Get or create the picker
            let picker = document.getElementById(pickerId);
            if (!picker) {
                picker = document.createElement('input');
                picker.type = 'color';
                picker.id = pickerId;
                picker.style.display = 'none';
                document.body.appendChild(picker);
                
                // Add change event
                picker.addEventListener('input', function() {
                    handler(this.value);
                });
            }
            
            // Trigger the color picker
            picker.click();
        });
    }
    
    const colorPicker = document.getElementById(pickerId);
    if (colorPicker) {
        // Add the change event to the picker if it exists
        colorPicker.addEventListener('input', function() {
            handler(this.value);
        });
        
        console.log(`Set up color picker input: ${pickerId}`);
    }
}

// Function to apply background color to cells
function applyBackgroundColor(color) {
    console.log('Applying background color:', color);
    
    // Get selected cells from the global scope or local scope
    const cells = getSelectedCells();
    
    if (!cells || cells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    cells.forEach(cell => {
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

// Function to apply header background color
function applyHeaderColor(color) {
    console.log('Applying header color:', color);
    
    // Get selected cells from the global scope or local scope
    const cells = getSelectedCells();
    
    if (!cells || cells.length === 0) {
        showToast('Please select at least one header cell', 'error');
        return;
    }
    
    let headerCount = 0;
    cells.forEach(cell => {
        if (cell.tagName === 'TH' || cell.classList.contains('sheet-header')) {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.backgroundColor = color;
                headerCount++;
            }
        }
    });
    
    if (headerCount > 0) {
        if (typeof window.saveSheetData === 'function') {
            window.saveSheetData();
        }
        showToast(`Applied color to ${headerCount} header(s)`, 'success');
    } else {
        showToast('Please select a header cell', 'error');
    }
}

// Function to apply text color to cells
function applyTextColor(color) {
    console.log('Applying text color:', color);
    
    // Get selected cells from the global scope or local scope
    const cells = getSelectedCells();
    
    if (!cells || cells.length === 0) {
        showToast('Please select at least one cell', 'error');
        return;
    }
    
    cells.forEach(cell => {
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

// Helper function to get the selected cells from either window scope or global scope
function getSelectedCells() {
    // First try the window object
    if (window.selectedCells && window.selectedCells.length > 0) {
        return window.selectedCells;
    }
    
    // Try the local variable if it exists
    if (typeof selectedCells !== 'undefined' && selectedCells.length > 0) {
        return selectedCells;
    }
    
    // Last resort: query the DOM for selected cells
    const selected = document.querySelectorAll('.sheet-cell.selected, .sheet-header.selected');
    if (selected.length > 0) {
        return Array.from(selected);
    }
    
    return [];
}

// Helper to show toast notifications if not defined elsewhere
function showToast(message, type = 'info') {
    // Check if the function exists globally
    if (typeof window.showToast === 'function') {
        window.showToast(message, type);
        return;
    }
    
    // Fallback implementation
    const toastContainer = document.querySelector('.toast-container');
    if (!toastContainer) {
        // Create toast container if it doesn't exist
        const container = document.createElement('div');
        container.className = 'toast-container position-fixed top-0 end-0 p-3';
        document.body.appendChild(container);
    }
    
    const toast = document.createElement('div');
    toast.className = `toast show mb-2`;
    toast.role = 'alert';
    toast.setAttribute('aria-live', 'assertive');
    toast.setAttribute('aria-atomic', 'true');
    
    // Set background color based on type
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
    
    // Add to container
    const container = document.querySelector('.toast-container');
    container.appendChild(toast);
    
    // Add close button functionality
    toast.querySelector('.btn-close').addEventListener('click', function() {
        toast.remove();
    });
    
    // Auto-remove after 3 seconds
    setTimeout(() => {
        toast.remove();
    }, 3000);
}