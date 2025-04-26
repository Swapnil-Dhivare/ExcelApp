document.addEventListener('DOMContentLoaded', function() {
    // Only run preview functionality if we're on a page with the preview elements
    const previewTable = document.getElementById('previewTable');
    if (!previewTable) {
        // We're not on a page with preview functionality - exit early
        return;
    }
    
    // Get elements
    const sheetNameInput = document.getElementById('sheet_name');
    const dataInput = document.getElementById('data');
    const formatControls = document.querySelectorAll('[data-preview]');
    
    // Preview text element for format testing
    const previewText = document.querySelector('.preview-text');
    
    // Initialize color palettes
    initColorPalettes();
    
    // Update data preview
    function updateDataPreview(data) {
        const tbody = previewTable.querySelector('tbody');
        tbody.innerHTML = '';

        const rows = data.split(';');
        rows.forEach(row => {
            if (row.trim()) {
                const tr = document.createElement('tr');
                const cells = row.split(',');
                
                cells.forEach(cell => {
                    const td = document.createElement('td');
                    td.textContent = cell.trim();
                    td.className = 'border px-2 py-1';
                    tr.appendChild(td);
                });
                
                tbody.appendChild(tr);
            }
        });
    }
    
    // Update format preview
    function updateFormatPreview() {
        const formatting = {
            'font-family': document.getElementById('font_name')?.value || 'Arial',
            'font-size': `${document.getElementById('font_size')?.value || 11}px`,
            'font-weight': document.getElementById('font_bold')?.checked ? 'bold' : 'normal',
            'font-style': document.getElementById('font_italic')?.checked ? 'italic' : 'normal',
            'color': document.getElementById('font_color')?.value || '#000000',
            'text-align': document.getElementById('cell_alignment')?.value || 'left',
            'vertical-align': document.getElementById('vertical_alignment')?.value || 'middle'
        };
        
        // Special handling for background color
        const bgColor = document.getElementById('cell_bg_color')?.value;
        if (bgColor === 'transparent') {
            if (previewText) {
                previewText.style.backgroundColor = '';
            }
        } else {
            formatting['background-color'] = bgColor || '#FFFFFF';
        }
        
        // Apply to preview text
        if (previewText) {
            Object.assign(previewText.style, formatting);
        }
        
        // Apply to preview table cells
        previewTable.querySelectorAll('td').forEach(cell => {
            applyFormatting(cell);
        });
    }
    
    // Apply formatting to an element
    function applyFormatting(element) {
        const formatting = {
            'font-family': document.getElementById('font_name')?.value || 'Arial',
            'font-size': `${document.getElementById('font_size')?.value || 11}px`,
            'font-weight': document.getElementById('font_bold')?.checked ? 'bold' : 'normal',
            'font-style': document.getElementById('font_italic')?.checked ? 'italic' : 'normal',
            'color': document.getElementById('font_color')?.value || '#000000',
            'text-align': document.getElementById('cell_alignment')?.value || 'left',
            'vertical-align': document.getElementById('vertical_alignment')?.value || 'middle',
            'border': document.getElementById('all_borders')?.checked ? '1px solid #000' : 'none',
            'padding': '4px'
        };
        
        // Special handling for background color
        const bgColor = document.getElementById('cell_bg_color')?.value;
        if (bgColor === 'transparent') {
            element.style.backgroundColor = ''; // Remove background color
        } else {
            formatting['background-color'] = bgColor || '#FFFFFF';
        }
        
        Object.assign(element.style, formatting);
    }
    
    // Initialize color palettes
    function initColorPalettes() {
        // Excel-like color palette
        const excelColors = [
            ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
            ['#F2F2F2', '#7F7F7F', '#D0CECE', '#D6DCE4', '#D9E1F2', '#FBE5D6', '#EDEDED', '#FFF2CC', '#DEEAF6', '#E2EFD9'],
            ['#D9D9D9', '#595959', '#AEAAAA', '#ADB9CA', '#B4C6E7', '#F7CBAC', '#DBDBDB', '#FFEA99', '#BDD7EE', '#C5E0B4'],
            ['#BFBFBF', '#404040', '#767171', '#8496B0', '#8EAADB', '#F4B183', '#C9C9C9', '#FFD966', '#9BC2E6', '#A9D08E'],
            ['#A6A6A6', '#262626', '#524F4F', '#5B6B95', '#698ED0', '#F19558', '#B7B7B7', '#FFC000', '#6FA1D9', '#6BB567'],
            ['#808080', '#0C0C0C', '#2E2B2B', '#323E56', '#2F5597', '#C45911', '#7B7B7B', '#BF8F00', '#2E74B5', '#538135']
        ];
        
        // Standard colors
        const standardColors = [
            '#FF0000', '#FF9900', '#FFFF00', '#00FF00', '#00FFFF', '#0000FF', '#9900FF', '#FF00FF', 
            '#C00000', '#E36C09', '#FFFF00', '#00B050', '#00B0F0', '#0070C0', '#7030A0', '#C00000'
        ];
        
        // Initialize all color pickers in the format panel
        renderColorPalettes(excelColors, standardColors);
        connectColorDropdowns(excelColors, standardColors);
    }
    
    // Render color palettes in all dropdown menus
    function renderColorPalettes(excelColors, standardColors) {
        const colorDropdowns = document.querySelectorAll('.dropdown-menu.color-palette-dropdown');
        
        colorDropdowns.forEach(dropdown => {
            // Create the theme colors section
            const themeColorsSection = document.createElement('div');
            themeColorsSection.className = 'color-section';
            
            const themeTitle = document.createElement('div');
            themeTitle.className = 'color-section-title';
            themeTitle.textContent = 'Theme Colors';
            themeColorsSection.appendChild(themeTitle);
            
            const themeColorsGrid = document.createElement('div');
            themeColorsGrid.className = 'excel-color-theme';
            
            // Add all theme colors
            excelColors.forEach(rowColors => {
                rowColors.forEach(color => {
                    const colorCell = document.createElement('div');
                    colorCell.className = 'color-cell';
                    colorCell.style.backgroundColor = color;
                    colorCell.setAttribute('data-color', color);
                    colorCell.title = color;
                    themeColorsGrid.appendChild(colorCell);
                });
            });
            
            themeColorsSection.appendChild(themeColorsGrid);
            dropdown.appendChild(themeColorsSection);
            
            // Create the standard colors section
            const standardColorsSection = document.createElement('div');
            standardColorsSection.className = 'color-section';
            
            const standardTitle = document.createElement('div');
            standardTitle.className = 'color-section-title';
            standardTitle.textContent = 'Standard Colors';
            standardColorsSection.appendChild(standardTitle);
            
            const standardColorsGrid = document.createElement('div');
            standardColorsGrid.className = 'standard-colors';
            
            // Add "No Color" option with improved styling
            const noColorCell = document.createElement('div');
            noColorCell.className = 'color-cell no-color-cell';
            noColorCell.setAttribute('data-color', 'transparent');
            noColorCell.title = 'No Color';
            
            // Add diagonal red line to visually indicate "No Color"
            const noColorLine = document.createElement('div');
            noColorLine.className = 'no-color-line';
            noColorCell.appendChild(noColorLine);
            
            standardColorsGrid.appendChild(noColorCell);
            
            // Add standard colors
            standardColors.forEach(color => {
                const colorCell = document.createElement('div');
                colorCell.className = 'color-cell';
                colorCell.style.backgroundColor = color;
                colorCell.setAttribute('data-color', color);
                colorCell.title = color;
                standardColorsGrid.appendChild(colorCell);
            });
            
            standardColorsSection.appendChild(standardColorsGrid);
            dropdown.appendChild(standardColorsSection);
            
            // Add custom color section
            const customSection = document.createElement('div');
            customSection.className = 'custom-color-section';
            
            const customLabel = document.createElement('label');
            customLabel.className = 'color-label';
            customLabel.textContent = 'Custom Color';
            customSection.appendChild(customLabel);
            
            const customColorInput = document.createElement('input');
            customColorInput.type = 'color';
            customColorInput.className = 'form-control form-control-sm';
            customColorInput.id = `custom${dropdown.id.replace('Dropdown', '')}Color`;
            customSection.appendChild(customColorInput);
            
            dropdown.appendChild(customSection);
        });
    }
    
    // Connect color dropdown functionality 
    function connectColorDropdowns(excelColors, standardColors) {
        // Handle all color dropdowns with color-palette-dropdown class
        const colorDropdowns = document.querySelectorAll('.dropdown-menu.color-palette-dropdown');
        
        colorDropdowns.forEach(dropdown => {
            const dropdownId = dropdown.getAttribute('aria-labelledby');
            const dropdownButton = document.getElementById(dropdownId);
            const colorPreviewSpan = dropdownButton?.querySelector('.color-preview');
            const hiddenInput = getAssociatedInput(dropdownId);
            
            // Set up all color cells in this dropdown
            const colorCells = dropdown.querySelectorAll('.color-cell');
            colorCells.forEach(cell => {
                const color = cell.getAttribute('data-color');
                
                // Update color on click
                cell.addEventListener('click', function(e) {
                    e.preventDefault();
                    e.stopPropagation();
                    
                    // Remove selected class from all cells
                    dropdown.querySelectorAll('.color-cell.selected').forEach(selected => {
                        selected.classList.remove('selected');
                    });
                    
                    // Add selected class to this cell
                    cell.classList.add('selected');
                    
                    // Special handling for "No Color" option
                    if (color === 'transparent') {
                        if (colorPreviewSpan) {
                            // Clear background color and add visual indicator for "No Color"
                            colorPreviewSpan.style.backgroundColor = '';
                            colorPreviewSpan.classList.add('no-color-cell');
                            
                            // Add diagonal line to preview
                            if (!colorPreviewSpan.querySelector('.no-color-line')) {
                                const noColorLine = document.createElement('div');
                                noColorLine.className = 'no-color-line';
                                colorPreviewSpan.appendChild(noColorLine);
                            }
                        }
                    } else {
                        // Update color preview in button for regular colors
                        if (colorPreviewSpan) {
                            colorPreviewSpan.style.backgroundColor = color;
                            colorPreviewSpan.classList.remove('no-color-cell');
                            
                            // Remove no-color line if it exists
                            const noColorLine = colorPreviewSpan.querySelector('.no-color-line');
                            if (noColorLine) {
                                noColorLine.remove();
                            }
                        }
                    }
                    
                    // Update hidden input value
                    if (hiddenInput) {
                        hiddenInput.value = color;
                        
                        // Trigger change event
                        const event = new Event('change', { bubbles: true });
                        hiddenInput.dispatchEvent(event);
                    }
                    
                    // Hide dropdown
                    const bsDropdown = bootstrap.Dropdown.getInstance(dropdownButton);
                    if (bsDropdown) {
                        bsDropdown.hide();
                    }
                    
                    // Refresh preview if available
                    updateFormatPreview();
                });
            });
            
            // Handle custom color inputs
            const customColorInput = dropdown.querySelector('input[type="color"]');
            if (customColorInput) {
                customColorInput.addEventListener('input', function() {
                    // Update color preview in button
                    if (colorPreviewSpan) {
                        colorPreviewSpan.style.backgroundColor = this.value;
                        colorPreviewSpan.classList.remove('no-color-cell');
                        
                        // Remove no-color line if it exists
                        const noColorLine = colorPreviewSpan.querySelector('.no-color-line');
                        if (noColorLine) {
                            noColorLine.remove();
                        }
                    }
                    
                    // Update hidden input value
                    if (hiddenInput) {
                        hiddenInput.value = this.value;
                        
                        // Trigger change event
                        const event = new Event('change', { bubbles: true });
                        hiddenInput.dispatchEvent(event);
                    }
                    
                    // Refresh preview if available
                    updateFormatPreview();
                });
                
                // Also handle the change event for when a color is selected
                customColorInput.addEventListener('change', function() {
                    // Remove selected class from all cells
                    dropdown.querySelectorAll('.color-cell.selected').forEach(selected => {
                        selected.classList.remove('selected');
                    });
                    
                    // Refresh preview if available
                    updateFormatPreview();
                });
            }
            
            // Initialize color preview with the current value
            if (colorPreviewSpan && hiddenInput && hiddenInput.value) {
                // Special handling for "transparent" value
                if (hiddenInput.value === 'transparent') {
                    colorPreviewSpan.classList.add('no-color-cell');
                    const noColorLine = document.createElement('div');
                    noColorLine.className = 'no-color-line';
                    colorPreviewSpan.appendChild(noColorLine);
                } else {
                    colorPreviewSpan.style.backgroundColor = hiddenInput.value;
                    colorPreviewSpan.classList.remove('no-color-cell');
                }
                
                // Mark the corresponding cell as selected if it exists
                const matchingCell = Array.from(dropdown.querySelectorAll('.color-cell')).find(
                    cell => cell.getAttribute('data-color') === hiddenInput.value
                );
                
                if (matchingCell) {
                    matchingCell.classList.add('selected');
                }
            }
        });
    }
    
    // Helper function to find the associated input for a color dropdown
    function getAssociatedInput(dropdownId) {
        switch(dropdownId) {
            case 'textColorDropdown':
                return document.getElementById('customTextColor') || document.getElementById('font_color');
            case 'fillColorDropdown':
                return document.getElementById('customFillColor') || document.getElementById('cell_bg_color');
            case 'borderColorDropdown':
                return document.getElementById('customBorderColor') || document.getElementById('border_color');
            default:
                return null;
        }
    }
    
    // Event listeners
    if (dataInput) {
        dataInput.addEventListener('input', function() {
            updateDataPreview(this.value);
        });
    }
    
    formatControls.forEach(control => {
        control.addEventListener('change', () => {
            updateFormatPreview();
            if (dataInput) {
                updateDataPreview(dataInput.value);
            }
        });
    });
    
    // Initial preview
    updateFormatPreview();
    if (dataInput) {
        updateDataPreview(dataInput.value);
    }
    
    // Make loadSheetPreview globally accessible
    window.loadSheetPreview = function(sheetName) {
        const sheetsDataInput = document.getElementById('sheets-data');
        if (!sheetsDataInput) {
            console.error('sheets-data element not found');
            return;
        }

        let sheets;
        try {
            sheets = JSON.parse(sheetsDataInput.value);
        } catch (e) {
            console.error('Failed to parse sheets data:', e);
            return;
        }

        const sheet = sheets.find(s => s.sheet_name === sheetName);
        if (!sheet || !sheet.data || !sheet.data.length) {
            console.error('Sheet or data not found');
            return;
        }
        
        // Get preview elements
        const previewModal = document.getElementById('previewModal');
        const previewTitle = document.getElementById('previewModalLabel');
        const previewContent = document.getElementById('previewContent');
        
        if (!previewModal || !previewTitle || !previewContent) {
            console.error('Preview modal elements not found');
            return;
        }
        
        // Set title
        previewTitle.textContent = `Preview: ${sheetName}`;
        
        // Show loading indicator
        previewContent.innerHTML = '<div class="d-flex justify-content-center"><div class="spinner-border text-primary" role="status"><span class="visually-hidden">Loading...</span></div></div>';
        
        // Create preview table
        const table = document.createElement('table');
        table.className = 'table table-bordered table-sm sheet-table';
        
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');
        
        // Process headers
        if (sheet.data.length > 0) {
            const headerRow = document.createElement('tr');
            
            // Add row header column
            const cornerCell = document.createElement('th');
            cornerCell.className = 'sheet-header bg-light';
            cornerCell.innerHTML = '<i class="bi bi-grid-3x3"></i>';
            headerRow.appendChild(cornerCell);
            
            // Add column headers (A, B, C, etc.)
            sheet.data[0].forEach((header, idx) => {
                const th = document.createElement('th');
                th.className = 'sheet-header bg-light';
                
                // Use column letter (A, B, C) instead of index
                const colLetter = String.fromCharCode(65 + idx);
                th.textContent = colLetter;
                
                headerRow.appendChild(th);
            });
            
            thead.appendChild(headerRow);
        }
        
        // Process data rows
        sheet.data.forEach((row, rowIdx) => {
            const tr = document.createElement('tr');
            
            // Add row header (1, 2, 3, etc.)
            const rowHeader = document.createElement('th');
            rowHeader.className = 'row-header bg-light text-center';
            rowHeader.textContent = rowIdx + 1;
            tr.appendChild(rowHeader);
            
            // Add cells
            row.forEach((cellData, colIdx) => {
                const td = document.createElement('td');
                td.className = 'sheet-cell';
                
                // Create editable cell div
                const editableDiv = document.createElement('div');
                editableDiv.className = 'editable-cell';
                
                // Check if it's a formula
                if (typeof cellData === 'string' && cellData.trim().startsWith('=')) {
                    td.classList.add('formula-cell');
                    editableDiv.textContent = cellData;
                    
                    // Add info about formula
                    td.setAttribute('data-formula', cellData);
                } else {
                    editableDiv.textContent = cellData;
                }
                
                // Apply cell formatting if available
                if (sheet.format_info && sheet.format_info.cells) {
                    const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx + 1}`;
                    const format = sheet.format_info.cells[cellRef];
                    
                    if (format) {
                        // Apply font
                        if (format.font) {
                            editableDiv.style.fontFamily = format.font.name || 'inherit';
                            editableDiv.style.fontSize = `${format.font.size || 11}px`;
                            editableDiv.style.fontWeight = format.font.bold ? 'bold' : 'normal';
                            editableDiv.style.fontStyle = format.font.italic ? 'italic' : 'normal';
                            
                            if (format.font.color) {
                                editableDiv.style.color = format.font.color;
                            }
                        }
                        
                        // Apply background color
                        if (format.fill && format.fill.color) {
                            editableDiv.style.backgroundColor = format.fill.color;
                        }
                        
                        // Apply alignment
                        if (format.alignment) {
                            editableDiv.style.textAlign = format.alignment.horizontal || 'left';
                            editableDiv.style.verticalAlign = format.alignment.vertical || 'middle';
                        }
                        
                        // Apply border
                        if (format.border) {
                            const borderStyle = format.border.double ? '2px double' : '1px solid';
                            editableDiv.style.border = `${borderStyle} #000`;
                        }
                    }
                }
                
                td.appendChild(editableDiv);
                tr.appendChild(td);
            });
            
            tbody.appendChild(tr);
        });
        
        table.appendChild(thead);
        table.appendChild(tbody);
        
        // Wrap table in responsive div
        const tableWrapper = document.createElement('div');
        tableWrapper.className = 'table-responsive';
        tableWrapper.appendChild(table);
        
        // Update preview content
        previewContent.innerHTML = '';
        previewContent.appendChild(tableWrapper);
        
        // Add Excel-like toolbar above the table
        const toolbar = document.createElement('div');
        toolbar.className = 'excel-toolbar mb-3';
        toolbar.innerHTML = `
            <div class="btn-toolbar" role="toolbar">
                <div class="btn-group me-2" role="group">
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Bold">
                        <i class="bi bi-type-bold"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Italic">
                        <i class="bi bi-type-italic"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Underline">
                        <i class="bi bi-type-underline"></i>
                    </button>
                </div>
                
                <div class="btn-group me-2" role="group">
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Align Left">
                        <i class="bi bi-text-left"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Align Center">
                        <i class="bi bi-text-center"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Align Right">
                        <i class="bi bi-text-right"></i>
                    </button>
                </div>
                
                <div class="btn-group me-2" role="group">
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Format as Number">
                        <i class="bi bi-123"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Format as Currency">
                        <i class="bi bi-currency-dollar"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Format as Percentage">
                        <i class="bi bi-percent"></i>
                    </button>
                </div>
                
                <div class="btn-group" role="group">
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Add Chart">
                        <i class="bi bi-bar-chart"></i>
                    </button>
                    <button type="button" class="btn btn-sm btn-outline-secondary" title="Add Function">
                        <i class="bi bi-calculator"></i>
                    </button>
                </div>
            </div>
        `;
        
        previewContent.insertBefore(toolbar, tableWrapper);
        
        // Add formula bar
        const formulaBar = document.createElement('div');
        formulaBar.className = 'formula-bar mb-2';
        formulaBar.innerHTML = `
            <div class="cell-address">A1</div>
            <input type="text" class="formula-input form-control form-control-sm" placeholder="fx">
        `;
        
        previewContent.insertBefore(formulaBar, tableWrapper);
        
        // Show the modal
        const modal = new bootstrap.Modal(previewModal);
        modal.show();
    };
});

