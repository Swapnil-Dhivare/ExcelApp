/**
 * Excel-like header formatting functionality
 * Enables collective styling of header cells
 */

document.addEventListener('DOMContentLoaded', function() {
    // Initialize header formatting
    initHeaderFormatting();
    
    /**
     * Initialize header formatting functionality
     */
    function initHeaderFormatting() {
        // Add a button to the formatting panel for header styling
        const formattingPanel = document.querySelector('.formatting-panel .row');
        if (formattingPanel) {
            const headerFormattingDiv = document.createElement('div');
            headerFormattingDiv.className = 'col-auto';
            headerFormattingDiv.innerHTML = `
                <span class="d-block text-muted small mb-1">Headers</span>
                <div class="btn-group" role="group" aria-label="Header formatting">
                    <button type="button" class="btn btn-sm btn-outline-secondary" id="formatHeadersBtn" title="Format Headers">
                        <i class="bi bi-type-h1"></i> Headers
                    </button>
                </div>
            `;
            
            formattingPanel.appendChild(headerFormattingDiv);
            
            // Add event listener
            document.getElementById('formatHeadersBtn').addEventListener('click', showHeaderFormattingDialog);
        }
    }
    
    /**
     * Show header formatting dialog
     */
    function showHeaderFormattingDialog() {
        // Create modal dialog for header formatting
        const dialog = document.createElement('div');
        dialog.className = 'modal fade';
        dialog.id = 'headerFormattingModal';
        dialog.setAttribute('tabindex', '-1');
        dialog.setAttribute('aria-labelledby', 'headerFormattingModalLabel');
        dialog.setAttribute('aria-hidden', 'true');
        
        dialog.innerHTML = `
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white">
                        <h5 class="modal-title" id="headerFormattingModalLabel">Format Headers</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <div class="mb-3">
                            <label for="header-bg-color" class="form-label">Background Color</label>
                            <input type="color" class="form-control form-control-color" id="header-bg-color" value="#f3f2f1">
                        </div>
                        <div class="mb-3">
                            <label for="header-text-color" class="form-label">Text Color</label>
                            <input type="color" class="form-control form-control-color" id="header-text-color" value="#000000">
                        </div>
                        <div class="mb-3">
                            <label for="header-font-weight" class="form-label">Font Weight</label>
                            <select class="form-select" id="header-font-weight">
                                <option value="normal">Normal</option>
                                <option value="bold" selected>Bold</option>
                            </select>
                        </div>
                        <div class="mb-3">
                            <label for="header-alignment" class="form-label">Text Alignment</label>
                            <div class="btn-group w-100" role="group">
                                <input type="radio" class="btn-check" name="header-align" id="header-align-left" value="left">
                                <label class="btn btn-outline-primary" for="header-align-left"><i class="bi bi-text-left"></i></label>
                                
                                <input type="radio" class="btn-check" name="header-align" id="header-align-center" value="center" checked>
                                <label class="btn btn-outline-primary" for="header-align-center"><i class="bi bi-text-center"></i></label>
                                
                                <input type="radio" class="btn-check" name="header-align" id="header-align-right" value="right">
                                <label class="btn btn-outline-primary" for="header-align-right"><i class="bi bi-text-right"></i></label>
                            </div>
                        </div>
                        <div class="form-check mb-3">
                            <input class="form-check-input" type="checkbox" id="apply-to-all-headers" checked>
                            <label class="form-check-label" for="apply-to-all-headers">
                                Apply to all headers (not just selected)
                            </label>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-primary" id="applyHeaderFormattingBtn">Apply</button>
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(dialog);
        
        // Initialize Bootstrap modal
        const headerModal = new bootstrap.Modal(document.getElementById('headerFormattingModal'));
        headerModal.show();
        
        // Set up event handler for apply button
        document.getElementById('applyHeaderFormattingBtn').addEventListener('click', function() {
            applyHeaderFormatting();
            headerModal.hide();
        });
        
        // Clean up modal when hidden
        dialog.addEventListener('hidden.bs.modal', function() {
            this.remove();
        });
    }
    
    /**
     * Apply formatting to headers
     */
    function applyHeaderFormatting() {
        // Get formatting options
        const bgColor = document.getElementById('header-bg-color').value;
        const textColor = document.getElementById('header-text-color').value;
        const fontWeight = document.getElementById('header-font-weight').value;
        const alignment = document.querySelector('input[name="header-align"]:checked').value;
        const applyToAll = document.getElementById('apply-to-all-headers').checked;
        
        // Get headers to format
        let headers;
        
        if (applyToAll) {
            // Apply to all headers
            headers = document.querySelectorAll('#sheetTable th, .excel-column-header');
        } else {
            // Apply only to selected headers
            headers = [];
            if (window.selectedCells && window.selectedCells.length > 0) {
                window.selectedCells.forEach(cell => {
                    if (cell.tagName === 'TH' || cell.classList.contains('excel-column-header')) {
                        headers.push(cell);
                    }
                });
            }
        }
        
        // Apply formatting to headers
        headers.forEach(header => {
            // Apply background color
            header.style.backgroundColor = bgColor;
            
            // Find the editable cell inside header
            const editableCell = header.querySelector('.editable-cell');
            if (editableCell) {
                editableCell.style.color = textColor;
                editableCell.style.fontWeight = fontWeight;
                editableCell.style.textAlign = alignment;
            } else {
                // If no editable cell, apply to header itself
                header.style.color = textColor;
                header.style.fontWeight = fontWeight;
                header.style.textAlign = alignment;
            }
        });
        
        // Store the header styling in localStorage
        saveHeaderStyling(bgColor, textColor, fontWeight, alignment);
        
        // Show success message
        if (typeof window.showToast === 'function') {
            window.showToast('Header formatting applied', 'success');
        }
    }
    
    /**
     * Save header styling to localStorage
     */
    function saveHeaderStyling(bgColor, textColor, fontWeight, alignment) {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        const headerStyle = {
            bgColor,
            textColor,
            fontWeight,
            alignment
        };
        
        try {
            localStorage.setItem(`excel_header_style_${sheetName}`, JSON.stringify(headerStyle));
        } catch (e) {
            console.error('Error saving header styling to localStorage:', e);
        }
    }
    
    /**
     * Load header styling from localStorage
     */
    function loadHeaderStyling() {
        const sheetTable = document.getElementById('sheetTable');
        if (!sheetTable) return;
        
        const sheetName = sheetTable.getAttribute('data-sheet-name');
        if (!sheetName) return;
        
        try {
            const savedStyle = localStorage.getItem(`excel_header_style_${sheetName}`);
            if (!savedStyle) return;
            
            const headerStyle = JSON.parse(savedStyle);
            
            // Apply the saved styling to all headers
            const headers = document.querySelectorAll('#sheetTable th, .excel-column-header');
            headers.forEach(header => {
                // Apply background color
                header.style.backgroundColor = headerStyle.bgColor;
                
                // Find the editable cell inside header
                const editableCell = header.querySelector('.editable-cell');
                if (editableCell) {
                    editableCell.style.color = headerStyle.textColor;
                    editableCell.style.fontWeight = headerStyle.fontWeight;
                    editableCell.style.textAlign = headerStyle.alignment;
                } else {
                    // If no editable cell, apply to header itself
                    header.style.color = headerStyle.textColor;
                    header.style.fontWeight = headerStyle.fontWeight;
                    header.style.textAlign = headerStyle.alignment;
                }
            });
        } catch (e) {
            console.error('Error loading header styling:', e);
        }
    }
    
    // Load saved header styling when page loads
    loadHeaderStyling();
});
