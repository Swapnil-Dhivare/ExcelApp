/**
 * Preview Styling - Direct implementation for font formatting in preview modal
 */
document.addEventListener('DOMContentLoaded', function() {
    console.log('Preview styling module loaded');
    
    // Set up modal event handler to initialize formatting buttons when shown
    const previewModal = document.getElementById('previewModal');
    if (previewModal) {
        previewModal.addEventListener('shown.bs.modal', initializeFormatButtons);
    }
    
    // Function to initialize all formatting buttons
    function initializeFormatButtons() {
        console.log('Initializing format buttons in preview modal');
        
        // Font family selector
        const fontFamilySelect = document.getElementById('fontFamilySelect');
        if (fontFamilySelect) {
            fontFamilySelect.addEventListener('change', function() {
                const fontFamily = this.value;
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.fontFamily = fontFamily;
                });
            });
        }
        
        // Font size selector
        const fontSizeSelect = document.getElementById('fontSizeSelect');
        if (fontSizeSelect) {
            fontSizeSelect.addEventListener('change', function() {
                const fontSize = this.value;
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.fontSize = `${fontSize}px`;
                });
            });
        }
        
        // Bold button
        const boldBtn = document.getElementById('boldBtn');
        if (boldBtn) {
            boldBtn.addEventListener('click', function() {
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) {
                        const currentWeight = window.getComputedStyle(editableDiv).fontWeight;
                        editableDiv.style.fontWeight = (currentWeight === '700' || currentWeight === 'bold') ? 'normal' : 'bold';
                    }
                });
            });
        }
        
        // Italic button
        const italicBtn = document.getElementById('italicBtn');
        if (italicBtn) {
            italicBtn.addEventListener('click', function() {
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) {
                        const currentStyle = window.getComputedStyle(editableDiv).fontStyle;
                        editableDiv.style.fontStyle = (currentStyle === 'italic') ? 'normal' : 'italic';
                    }
                });
            });
        }
        
        // Text color button
        const textColorBtn = document.getElementById('textColorBtn');
        const textColorPicker = document.getElementById('textColorPicker');
        if (textColorBtn && textColorPicker) {
            textColorBtn.addEventListener('click', function() {
                textColorPicker.click();
            });
            
            textColorPicker.addEventListener('input', function() {
                const color = this.value;
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.color = color;
                });
            });
        }
        
        // Fill color button
        const fillColorBtn = document.getElementById('fillColorBtn');
        const fillColorPicker = document.getElementById('fillColorPicker');
        if (fillColorBtn && fillColorPicker) {
            fillColorBtn.addEventListener('click', function() {
                fillColorPicker.click();
            });
            
            fillColorPicker.addEventListener('input', function() {
                const color = this.value;
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.backgroundColor = color;
                });
            });
        }
        
        // Alignment buttons
        const alignLeftBtn = document.getElementById('alignLeftBtn');
        if (alignLeftBtn) {
            alignLeftBtn.addEventListener('click', function() {
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.textAlign = 'left';
                });
            });
        }
        
        const alignCenterBtn = document.getElementById('alignCenterBtn');
        if (alignCenterBtn) {
            alignCenterBtn.addEventListener('click', function() {
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.textAlign = 'center';
                });
            });
        }
        
        const alignRightBtn = document.getElementById('alignRightBtn');
        if (alignRightBtn) {
            alignRightBtn.addEventListener('click', function() {
                applyToSelectedCells(cell => {
                    const editableDiv = cell.querySelector('.editable-cell');
                    if (editableDiv) editableDiv.style.textAlign = 'right';
                });
            });
        }
        
        // Set up cell selection
        setupCellSelection();
    }
    
    // Function to apply formatting to selected cells
    function applyToSelectedCells(formatFn) {
        // Try to get cells from global selectedCells first
        let cells = window.selectedCells || [];
        
        // If no cells are selected, show an error message
        if (cells.length === 0) {
            cells = document.querySelectorAll('.sheet-cell.selected');
        }
        
        // If still no cells, show error message
        if (cells.length === 0) {
            console.warn('No cells selected. Please select at least one cell.');
            return;
        }
        
        // Apply the formatting function to each cell
        cells.forEach(cell => formatFn(cell));
    }
    
    // Function to set up cell selection in the preview modal
    function setupCellSelection() {
        // Initialize selectedCells if it doesn't exist globally
        window.selectedCells = window.selectedCells || [];
        
        // Find all cells in the preview table
        const cells = document.querySelectorAll('.sheet-table .sheet-cell');
        
        cells.forEach(cell => {
            // Remove any existing click handlers
            cell.removeEventListener('click', cellClickHandler);
            
            // Add the click handler
            cell.addEventListener('click', cellClickHandler);
        });
    }
    
    // Click handler for cell selection
    function cellClickHandler(event) {
        // Initialize selectedCells if needed
        window.selectedCells = window.selectedCells || [];
        
        // Handle Ctrl/Cmd key for multi-selection
        if (!(event.ctrlKey || event.metaKey)) {
            // Clear previous selection
            document.querySelectorAll('.sheet-cell.selected').forEach(cell => {
                cell.classList.remove('selected');
            });
            window.selectedCells = [];
        }
        
        // Toggle selection for this cell
        this.classList.toggle('selected');
        
        // Update selectedCells array
        if (this.classList.contains('selected')) {
            window.selectedCells.push(this);
        } else {
            window.selectedCells = window.selectedCells.filter(cell => cell !== this);
        }
        
        console.log(`Selected cells: ${window.selectedCells.length}`);
        event.stopPropagation();
    }
});
