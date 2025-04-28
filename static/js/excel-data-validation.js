/**
 * Excel-like data validation functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize data validation menu
    initDataValidationContextMenu();
    
    /**
     * Initialize data validation context menu
     */
    function initDataValidationContextMenu() {
        const sheetTable = document.getElementById('sheetTable');
        if (sheetTable) {
            // Add data validation option to context menu
            document.addEventListener('contextmenu', function(e) {
                const cell = e.target.closest('.sheet-cell, .sheet-header');
                if (!cell) return;
                
                // Check if there's an existing context menu event handler
                const existingContextMenu = document.getElementById('resize-context-menu');
                if (existingContextMenu) {
                    // Add data validation option to existing menu
                    const validationItem = document.createElement('div');
                    validationItem.className = 'context-menu-item';
                    validationItem.id = 'data-validation-option';
                    validationItem.innerHTML = 'Data Validation...';
                    
                    const divider = document.createElement('div');
                    divider.className = 'context-menu-divider';
                    
                    existingContextMenu.appendChild(divider);
                    existingContextMenu.appendChild(validationItem);
                    
                    validationItem.addEventListener('click', function() {
                        showDataValidationDialog();
                        document.getElementById('resize-context-menu').remove();
                    });
                }
            });
        }
        
        // Add Data Validation button to formatting panel
        const formattingPanel = document.querySelector('.formatting-panel');
        if (formattingPanel) {
            const validationDiv = document.createElement('div');
            validationDiv.className = 'col-auto';
            validationDiv.innerHTML = `
                <span class="d-block text-muted small mb-1">Validation</span>
                <div class="btn-group" role="group" aria-label="Data validation">
                    <button type="button" class="btn btn-sm btn-outline-secondary" id="dataValidationBtn" title="Data Validation">
                        <i class="bi bi-check-circle"></i> Validate
                    </button>
                </div>
            `;
            
            formattingPanel.querySelector('.row').appendChild(validationDiv);
            
            // Add event listener
            document.getElementById('dataValidationBtn')?.addEventListener('click', showDataValidationDialog);
        }
    }
    
    /**
     * Show data validation dialog
     */
    function showDataValidationDialog() {
        if (!window.selectedCells || window.selectedCells.length === 0) {
            showToast('Please select at least one cell', 'error');
            return;
        }
        
        // Create modal dialog for data validation
        const modal = document.createElement('div');
        modal.className = 'modal fade';
        modal.id = 'dataValidationModal';
        modal.setAttribute('tabindex', '-1');
        modal.setAttribute('aria-labelledby', 'dataValidationModalLabel');
        modal.setAttribute('aria-hidden', 'true');
        
        modal.innerHTML = `
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white">
                        <h5 class="modal-title" id="dataValidationModalLabel">Data Validation</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <form id="validationForm">
                            <div class="mb-3">
                                <label class="form-label">Validation criteria</label>
                                <select class="form-select" id="validationType">
                                    <option value="any">Any value</option>
                                    <option value="number">Number</option>
                                    <option value="text">Text only</option>
                                    <option value="date">Date</option>
                                    <option value="list">List of items</option>
                                    <option value="decimal">Decimal</option>
                                    <option value="custom">Custom formula</option>
                                </select>
                            </div>
                            
                            <div class="validation-options" id="numberOptions" style="display: none;">
                                <div class="mb-3">
                                    <select class="form-select" id="numberComparison">
                                        <option value="between">between</option>
                                        <option value="notBetween">not between</option>
                                        <option value="equal">equal to</option>
                                        <option value="notEqual">not equal to</option>
                                        <option value="greater">greater than</option>
                                        <option value="less">less than</option>
                                        <option value="greaterEqual">greater than or equal to</option>
                                        <option value="lessEqual">less than or equal to</option>
                                    </select>
                                </div>
                                
                                <div class="mb-3 row g-2" id="numberBetweenInputs">
                                    <div class="col">
                                        <input type="number" class="form-control" id="minValue" placeholder="Minimum">
                                    </div>
                                    <div class="col-auto pt-2">and</div>
                                    <div class="col">
                                        <input type="number" class="form-control" id="maxValue" placeholder="Maximum">
                                    </div>
                                </div>
                                
                                <div class="mb-3" id="numberSingleInput" style="display: none;">
                                    <input type="number" class="form-control" id="singleValue" placeholder="Value">
                                </div>
                            </div>
                            
                            <div class="validation-options" id="listOptions" style="display: none;">
                                <div class="mb-3">
                                    <label class="form-label">Source (comma separated values)</label>
                                    <textarea class="form-control" id="listSource" rows="3" placeholder="Option 1, Option 2, Option 3"></textarea>
                                </div>
                                <div class="form-check">
                                    <input class="form-check-input" type="checkbox" id="inCellDropdown">
                                    <label class="form-check-label" for="inCellDropdown">
                                        Show dropdown in cell
                                    </label>
                                </div>
                            </div>
                            
                            <div class="validation-options" id="customOptions" style="display: none;">
                                <div class="mb-3">
                                    <label class="form-label">Custom formula (starts with =)</label>
                                    <input type="text" class="form-control" id="customFormula" placeholder="=CELL>0">
                                </div>
                            </div>
                            
                            <div class="form-check mb-3">
                                <input class="form-check-input" type="checkbox" id="ignoreBlank" checked>
                                <label class="form-check-label" for="ignoreBlank">
                                    Ignore blank cells
                                </label>
                            </div>
                            
                            <div class="mb-3">
                                <label class="form-label">Error message (displayed for invalid entries)</label>
                                <input type="text" class="form-control" id="errorMessage" placeholder="Invalid input">
                            </div>
                        </form>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Cancel</button>
                        <button type="button" class="btn btn-primary" id="applyValidationBtn">Apply</button>
                    </div>
                </div>
            </div>
        `;
        
        document.body.appendChild(modal);
        
        // Initialize Bootstrap modal
        const modalElement = new bootstrap.Modal(document.getElementById('dataValidationModal'));
        modalElement.show();
        
        // Set up validation type change handler
        const validationType = document.getElementById('validationType');
        validationType.addEventListener('change', function() {
            // Hide all option groups first
            document.querySelectorAll('.validation-options').forEach(el => {
                el.style.display = 'none';
            });
            
            // Show the appropriate options based on selected type
            switch(this.value) {
                case 'number':
                case 'decimal':
                    document.getElementById('numberOptions').style.display = 'block';
                    break;
                case 'list':
                    document.getElementById('listOptions').style.display = 'block';
                    break;
                case 'custom':
                    document.getElementById('customOptions').style.display = 'block';
                    break;
            }
        });
        
        // Set up comparison type change handler
        const numberComparison = document.getElementById('numberComparison');
        numberComparison.addEventListener('change', function() {
            if (this.value === 'between' || this.value === 'notBetween') {
                document.getElementById('numberBetweenInputs').style.display = 'flex';
                document.getElementById('numberSingleInput').style.display = 'none';
            } else {
                document.getElementById('numberBetweenInputs').style.display = 'none';
                document.getElementById('numberSingleInput').style.display = 'block';
            }
        });
        
        // Apply validation button handler
        document.getElementById('applyValidationBtn').addEventListener('click', function() {
            applyDataValidation();
            modalElement.hide();
        });
        
        // Clean up modal when hidden
        document.getElementById('dataValidationModal').addEventListener('hidden.bs.modal', function() {
            this.remove();
        });
    }
    
    /**
     * Apply data validation to selected cells
     */
    function applyDataValidation() {
        if (!window.selectedCells || window.selectedCells.length === 0) return;
        
        const validationType = document.getElementById('validationType').value;
        let validationRule = { type: validationType };
        
        // Build validation rule based on type
        switch(validationType) {
            case 'number':
            case 'decimal':
                const comparison = document.getElementById('numberComparison').value;
                validationRule.comparison = comparison;
                
                if (comparison === 'between' || comparison === 'notBetween') {
                    validationRule.min = document.getElementById('minValue').value;
                    validationRule.max = document.getElementById('maxValue').value;
                } else {
                    validationRule.value = document.getElementById('singleValue').value;
                }
                break;
                
            case 'list':
                const source = document.getElementById('listSource').value;
                validationRule.values = source.split(',').map(item => item.trim());
                validationRule.dropdown = document.getElementById('inCellDropdown').checked;
                break;
                
            case 'custom':
                validationRule.formula = document.getElementById('customFormula').value;
                break;
        }
        
        // Add common properties
        validationRule.ignoreBlank = document.getElementById('ignoreBlank').checked;
        validationRule.errorMessage = document.getElementById('errorMessage').value;
        
        // Apply validation to all selected cells
        window.selectedCells.forEach(cell => {
            // Store validation rule in data attribute
            cell.setAttribute('data-validation', JSON.stringify(validationRule));
            cell.classList.add('has-validation');
            
            // For list validation, add dropdown functionality if enabled
            if (validationType === 'list' && validationRule.dropdown) {
                addDropdownToCell(cell, validationRule.values);
            }
            
            // Add input validation event listener
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.addEventListener('blur', function() {
                    validateCell(cell);
                });
            }
        });
        
        showToast(`Data validation applied to ${window.selectedCells.length} ${window.selectedCells.length === 1 ? 'cell' : 'cells'}`, 'success');
    }
    
    /**
     * Add dropdown to cell for list validation
     */
    function addDropdownToCell(cell, values) {
        const editableDiv = cell.querySelector('.editable-cell');
        if (!editableDiv) return;
        
        // Add dropdown icon
        const dropdownIcon = document.createElement('span');
        dropdownIcon.className = 'dropdown-icon';
        dropdownIcon.innerHTML = '<i class="bi bi-caret-down-fill"></i>';
        cell.appendChild(dropdownIcon);
        
        // Add click handler for the dropdown
        dropdownIcon.addEventListener('click', function(e) {
            e.stopPropagation();
            showDropdownList(cell, values);
        });
        
        // Add dropdown functionality to the editable div as well
        editableDiv.addEventListener('click', function(e) {
            if (e.ctrlKey) {
                showDropdownList(cell, values);
            }
        });
    }
    
    /**
     * Show dropdown list for cell
     */
    function showDropdownList(cell, values) {
        // Remove any existing dropdown
        document.querySelectorAll('.cell-dropdown').forEach(el => el.remove());
        
        // Create dropdown
        const dropdown = document.createElement('div');
        dropdown.className = 'cell-dropdown';
        
        // Add list items
        values.forEach(value => {
            const item = document.createElement('div');
            item.className = 'cell-dropdown-item';
            item.textContent = value;
            
            item.addEventListener('click', function() {
                const editableDiv = cell.querySelector('.editable-cell');
                if (editableDiv) {
                    editableDiv.textContent = value;
                    validateCell(cell);
                }
                dropdown.remove();
            });
            
            dropdown.appendChild(item);
        });
        
        // Position dropdown
        const rect = cell.getBoundingClientRect();
        dropdown.style.top = (rect.bottom + window.scrollY) + 'px';
        dropdown.style.left = (rect.left + window.scrollX) + 'px';
        dropdown.style.minWidth = rect.width + 'px';
        
        document.body.appendChild(dropdown);
        
        // Close dropdown when clicking outside
        document.addEventListener('click', function closeDropdown() {
            dropdown.remove();
            document.removeEventListener('click', closeDropdown);
        });
    }
    
    /**
     * Validate cell based on validation rules
     */
    function validateCell(cell) {
        // Get validation rule
        const validationAttr = cell.getAttribute('data-validation');
        if (!validationAttr) return true;
        
        try {
            const rule = JSON.parse(validationAttr);
            const editableDiv = cell.querySelector('.editable-cell');
            if (!editableDiv) return true;
            
            const value = editableDiv.textContent.trim();
            
            // Check if blank and should be ignored
            if (rule.ignoreBlank && value === '') return true;
            
            // Validate based on rule type
            let isValid = true;
            
            switch(rule.type) {
                case 'any':
                    isValid = true;
                    break;
                    
                case 'number':
                    isValid = validateNumber(value, rule);
                    break;
                    
                case 'decimal':
                    isValid = validateDecimal(value, rule);
                    break;
                    
                case 'text':
                    isValid = isNaN(parseFloat(value));
                    break;
                    
                case 'date':
                    isValid = !isNaN(new Date(value).getTime());
                    break;
                    
                case 'list':
                    isValid = rule.values.includes(value);
                    break;
                    
                case 'custom':
                    isValid = validateCustomFormula(value, rule.formula, cell);
                    break;
            }
            
            if (!isValid) {
                markCellInvalid(cell, rule.errorMessage || 'Invalid input');
                return false;
            } else {
                markCellValid(cell);
                return true;
            }
        } catch (e) {
            console.error('Error validating cell:', e);
            return true; // Fail open
        }
    }
    
    /**
     * Validate numeric input
     */
    function validateNumber(value, rule) {
        const num = parseFloat(value);
        if (isNaN(num)) return false;
        
        switch(rule.comparison) {
            case 'between':
                return num >= parseFloat(rule.min) && num <= parseFloat(rule.max);
            case 'notBetween':
                return num < parseFloat(rule.min) || num > parseFloat(rule.max);
            case 'equal':
                return num === parseFloat(rule.value);
            case 'notEqual':
                return num !== parseFloat(rule.value);
            case 'greater':
                return num > parseFloat(rule.value);
            case 'less':
                return num < parseFloat(rule.value);
            case 'greaterEqual':
                return num >= parseFloat(rule.value);
            case 'lessEqual':
                return num <= parseFloat(rule.value);
            default:
                return !isNaN(num); // Is a number
        }
    }
    
    /**
     * Validate decimal input
     */
    function validateDecimal(value, rule) {
        // Check if it's a decimal
        if (!/^\-?[0-9]*\.?[0-9]+$/.test(value)) return false;
        
        return validateNumber(value, rule);
    }
    
    /**
     * Validate using custom formula
     * Simple implementation - could be expanded
     */
    function validateCustomFormula(value, formula, cell) {
        if (!formula.startsWith('=')) return true;
        
        // Remove the equals sign
        formula = formula.substring(1);
        
        // Very basic formula support - replace CELL with value
        formula = formula.replace(/CELL/g, value);
        
        try {
            return Function('"use strict"; return (' + formula + ')')();
        } catch (e) {
            console.error('Error evaluating formula:', e);
            return false;
        }
    }
    
    /**
     * Mark a cell as invalid
     */
    function markCellInvalid(cell, message) {
        cell.classList.add('invalid-cell');
        cell.setAttribute('title', message);
        
        // Add small red triangle in corner for error indicator
        if (!cell.querySelector('.validation-error')) {
            const errorIndicator = document.createElement('div');
            errorIndicator.className = 'validation-error';
            cell.appendChild(errorIndicator);
        }
        
        // Make cell shake briefly
        cell.classList.add('shake');
        setTimeout(() => {
            cell.classList.remove('shake');
        }, 500);
    }
    
    /**
     * Mark a cell as valid
     */
    function markCellValid(cell) {
        cell.classList.remove('invalid-cell');
        cell.removeAttribute('title');
        
        const errorIndicator = cell.querySelector('.validation-error');
        if (errorIndicator) {
            errorIndicator.remove();
        }
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
    
    // Add global API
    window.excelValidation = {
        showDataValidationDialog,
        validateCell
    };
});
