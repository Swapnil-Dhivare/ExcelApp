/**
 * Excel format button functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize formatting buttons
    initFormatButtons();
    
    /**
     * Initialize all formatting buttons
     */
    function initFormatButtons() {
        // Bold button
        document.getElementById('boldBtn')?.addEventListener('click', function() {
            applyFormatting('bold');
        });
        
        // Italic button
        document.getElementById('italicBtn')?.addEventListener('click', function() {
            applyFormatting('italic');
        });
        
        // Underline button
        document.getElementById('underlineBtn')?.addEventListener('click', function() {
            applyFormatting('underline');
        });
        
        // Alignment buttons
        document.getElementById('alignLeftBtn')?.addEventListener('click', function() {
            applyTextAlign('left');
        });
        
        document.getElementById('alignCenterBtn')?.addEventListener('click', function() {
            applyTextAlign('center');
        });
        
        document.getElementById('alignRightBtn')?.addEventListener('click', function() {
            applyTextAlign('right');
        });
        
        // Background color button
        const bgColorBtn = document.getElementById('bgColorBtn');
        const bgColorPicker = document.getElementById('bgColorPicker');
        
        if (bgColorBtn && bgColorPicker) {
            bgColorBtn.addEventListener('click', function() {
                bgColorPicker.click();
            });
            
            bgColorPicker.addEventListener('input', function() {
                applyBackgroundColor(this.value);
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
                applyTextColor(this.value);
            });
        }
    }
    
    // Apply text formatting (bold, italic, underline)
    function applyFormatting(format) {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('No cells selected', 'info');
            return;
        }
        
        window.selectedCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                switch(format) {
                    case 'bold':
                        if (editableDiv.style.fontWeight === 'bold') {
                            editableDiv.style.fontWeight = 'normal';
                        } else {
                            editableDiv.style.fontWeight = 'bold';
                        }
                        break;
                    case 'italic':
                        if (editableDiv.style.fontStyle === 'italic') {
                            editableDiv.style.fontStyle = 'normal';
                        } else {
                            editableDiv.style.fontStyle = 'italic';
                        }
                        break;
                    case 'underline':
                        if (editableDiv.style.textDecoration === 'underline') {
                            editableDiv.style.textDecoration = 'none';
                        } else {
                            editableDiv.style.textDecoration = 'underline';
                        }
                        break;
                }
            }
        });
        
        showToast(`Applied ${format} formatting`, 'success');
        
        // Mark sheet as modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    // Apply text alignment
    function applyTextAlign(alignment) {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('No cells selected', 'info');
            return;
        }
        
        window.selectedCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.textAlign = alignment;
            }
        });
        
        showToast(`Applied ${alignment} alignment`, 'success');
        
        // Mark sheet as modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    // Apply background color
    function applyBackgroundColor(color) {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('No cells selected', 'info');
            return;
        }
        
        window.selectedCells.forEach(cell => {
            cell.style.backgroundColor = color;
        });
        
        showToast('Applied background color', 'success');
        
        // Mark sheet as modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    // Apply text color
    function applyTextColor(color) {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('No cells selected', 'info');
            return;
        }
        
        window.selectedCells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.style.color = color;
            }
        });
        
        showToast('Applied text color', 'success');
        
        // Mark sheet as modified
        if (window.markSheetModified) {
            window.markSheetModified();
        }
    }
    
    // Show toast notification
    function showToast(message, type) {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            // Fallback toast implementation
            console.log(`${type}: ${message}`);
        }
    }
    
    // Expose formatting functions globally
    window.excelFormatting = {
        applyFormatting,
        applyTextAlign,
        applyBackgroundColor,
        applyTextColor
    };
});
