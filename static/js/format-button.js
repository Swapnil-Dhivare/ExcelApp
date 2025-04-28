/**
 * Excel format button functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Get the format button and panel
    const formatButton = document.getElementById('toggleFormatPanel');
    const formatPanel = document.querySelector('.formatting-panel');
    
    if (formatButton && formatPanel) {
        // Set initial state
        formatPanel.style.display = 'none';
        
        // Add click handler
        formatButton.addEventListener('click', function(e) {
            e.preventDefault();
            e.stopPropagation();
            
            // Toggle format panel visibility
            if (formatPanel.style.display === 'none') {
                formatPanel.style.display = 'block';
                formatButton.classList.add('active');
            } else {
                formatPanel.style.display = 'none';
                formatButton.classList.remove('active');
            }
        });
        
        // Initialize all formatting buttons
        initFormatButtons();
    }
    
    // Initialize all formatting buttons
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
    }
    
    // Show toast notification
    function showToast(message, type) {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            // Fallback toast implementation
            console.log(`${type}: ${message}`);
            
            // Create toast container if it doesn't exist
            let toastContainer = document.querySelector('.toast-container');
            if (!toastContainer) {
                toastContainer = document.createElement('div');
                toastContainer.className = 'toast-container position-fixed bottom-0 end-0 p-3';
                document.body.appendChild(toastContainer);
            }
            
            // Create toast
            const toastEl = document.createElement('div');
            toastEl.className = `toast show`;
            
            // Set color based on type
            const bgClass = type === 'error' ? 'bg-danger' : 
                           type === 'success' ? 'bg-success' : 
                           type === 'warning' ? 'bg-warning' : 'bg-info';
            
            toastEl.innerHTML = `
                <div class="toast-header ${bgClass} text-white">
                    <strong class="me-auto">Excel Generator</strong>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="toast"></button>
                </div>
                <div class="toast-body">
                    ${message}
                </div>
            `;
            
            // Add to container
            toastContainer.appendChild(toastEl);
            
            // Auto remove after delay
            setTimeout(() => {
                toastEl.remove();
            }, 3000);
        }
    }
    
    // Expose formatting functions globally
    window.applyFormatting = applyFormatting;
    window.applyTextAlign = applyTextAlign;
    window.applyBackgroundColor = applyBackgroundColor;
    window.applyTextColor = applyTextColor;
});
