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
            'background-color': document.getElementById('cell_bg_color')?.value || '#FFFFFF',
            'text-align': document.getElementById('cell_alignment')?.value || 'left',
            'vertical-align': document.getElementById('vertical_alignment')?.value || 'middle'
        };
        
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
            'background-color': document.getElementById('cell_bg_color')?.value || '#FFFFFF',
            'text-align': document.getElementById('cell_alignment')?.value || 'left',
            'vertical-align': document.getElementById('vertical_alignment')?.value || 'middle',
            'border': document.getElementById('all_borders')?.checked ? '1px solid #000' : 'none',
            'padding': '4px'
        };
        
        Object.assign(element.style, formatting);
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

        const thead = previewTable.querySelector('thead');
        const tbody = previewTable.querySelector('tbody');
        
        // Clear existing content
        thead.innerHTML = '';
        tbody.innerHTML = '';

        // Add header row
        if (sheet.data.length > 0) {
            const headerRow = document.createElement('tr');
            sheet.data[0].forEach(header => {
                const th = document.createElement('th');
                th.className = 'bg-light fw-bold';
                th.textContent = header;
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);

            // Add data rows
            sheet.data.slice(1).forEach(row => {
                const tr = document.createElement('tr');
                row.forEach((cell, colIndex) => {
                    const td = document.createElement('td');
                    if (typeof cell === 'string' && cell.startsWith('=')) {
                        td.textContent = evaluateFormula(cell, row, sheet.data);
                        td.classList.add('formula-cell');
                    } else {
                        td.textContent = cell;
                    }
                    td.contentEditable = true;
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        // Store current sheet name for save functionality
        previewTable.dataset.sheetName = sheetName;
    };

    // Save functionality
    window.saveSheetChanges = function() {
        const thead = previewTable.querySelector('thead');
        const tbody = previewTable.querySelector('tbody');
        const sheetName = previewTable.dataset.sheetName;

        // Collect all data including headers
        const updatedData = [
            // Get headers
            Array.from(thead.querySelectorAll('th')).map(th => th.textContent),
            // Get row data
            ...Array.from(tbody.querySelectorAll('tr')).map(row => 
                Array.from(row.querySelectorAll('td')).map(cell => cell.textContent.trim())
            )
        ];

        // Send to server
        fetch('/update_sheet_data', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({
                sheet_name: sheetName,
                data: updatedData
            })
        })
        .then(response => response.json())
        .then(data => {
            if (data.success) {
                location.reload();
            } else {
                alert('Failed to update sheet: ' + (data.error || 'Unknown error'));
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert('Failed to update sheet');
        });
    };

    function evaluateFormula(cell, rowData, allData) {
        // Remove the = sign if present
        const formula = cell.startsWith('=') ? cell.substring(1) : cell;
        
        try {
            if (formula.includes('SUM')) {
                // For assignment total calculation
                const assignments = ['ASS1', 'ASS2', 'ASS3', 'ASS4', 'ASS5'];
                let total = 0;
                assignments.forEach(ass => {
                    const idx = allData[0].indexOf(ass); // Get column index
                    if (idx !== -1) {
                        const val = parseFloat(rowData[idx]) || 0;
                        total += val;
                    }
                });
                return total;
            }
            return '#ERROR!';
        } catch (e) {
            console.error('Formula error:', e);
            return '#ERROR!';
        }
    }

    function getColumnIndex(column) {
        let index = 0;
        for (let i = 0; i < column.length; i++) {
            index = index * 26 + column.charCodeAt(i) - 64;
        }
        return index - 1;
    }
});

