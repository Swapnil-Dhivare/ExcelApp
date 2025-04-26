// Core app functionality

document.addEventListener('DOMContentLoaded', function() {
    // Initialize tooltips
    const tooltips = document.querySelectorAll('[data-bs-toggle="tooltip"]');
    if (tooltips.length > 0) {
        tooltips.forEach(tooltip => new bootstrap.Tooltip(tooltip));
    }
    
    // Initialize popovers
    const popovers = document.querySelectorAll('[data-bs-toggle="popover"]');
    if (popovers.length > 0) {
        popovers.forEach(popover => new bootstrap.Popover(popover));
    }
    
    // Log button status - debug only (remove in production)
    if (window.location.pathname.includes('/edit_sheet')) {
        console.log('Sheet editor detected, buttons:');
        console.log('Add Row button:', document.getElementById('addRowBtn'));
        console.log('Add Column button:', document.getElementById('addColumnBtn'));
        console.log('Delete Row button:', document.getElementById('deleteRowBtn'));
        console.log('Delete Column button:', document.getElementById('deleteColumnBtn'));
        console.log('Save Sheet button:', document.getElementById('saveSheetBtn'));
    }
});

/**
 * Remove a sheet with confirmation
 * @param {string} sheetName - Name of sheet to remove
 */
function removeSheet(sheetName) {
    if (confirm(`Are you sure you want to delete the sheet "${sheetName}"?`)) {
        // Create a form to submit POST request
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = `/remove_sheet/${encodeURIComponent(sheetName)}`;
        
        // Add CSRF token if needed
        const csrfToken = document.querySelector('meta[name="csrf-token"]').content;
        if (csrfToken) {
            const csrfInput = document.createElement('input');
            csrfInput.type = 'hidden';
            csrfInput.name = 'csrf_token';
            csrfInput.value = csrfToken;
            form.appendChild(csrfInput);
        }
        
        // Append to body, submit, then remove
        document.body.appendChild(form);
        form.submit();
        document.body.removeChild(form);
    }
}

/**
 * Preview sheet content
 * @param {string} sheetName - Name of sheet to preview
 */
function loadSheetPreview(sheetName) {
    const previewModal = new bootstrap.Modal(document.getElementById('previewModal'));
    
    // Show loading state
    document.getElementById('previewModalLabel').textContent = `Preview: ${sheetName}`;
    document.getElementById('previewContent').innerHTML = '<div class="text-center py-5"><div class="spinner-border text-primary"></div><p class="mt-2">Loading preview...</p></div>';
    previewModal.show();
    
    // Load sheet data
    fetch(`/sheet_preview/${encodeURIComponent(sheetName)}`)
        .then(response => {
            if (!response.ok) {
                throw new Error('Failed to load preview');
            }
            return response.json();
        })
        .then(data => {
            if (data.success) {
                renderSheetPreview(data.data);
            } else {
                document.getElementById('previewContent').innerHTML = 
                    `<div class="alert alert-danger">Error: ${data.error}</div>`;
            }
        })
        .catch(error => {
            document.getElementById('previewContent').innerHTML = 
                `<div class="alert alert-danger">Error: ${error.message}</div>`;
        });
}

/**
 * Render sheet preview in modal
 * @param {Array} data - Sheet data to render
 */
function renderSheetPreview(data) {
    const previewContent = document.getElementById('previewContent');
    
    if (!data || data.length === 0) {
        previewContent.innerHTML = '<div class="alert alert-info">No data available for preview.</div>';
        return;
    }
    
    const tableWrapper = document.createElement('div');
    tableWrapper.className = 'table-responsive';
    const table = document.createElement('table');
    table.className = 'table table-bordered table-sm';
    
    // Create headers
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headerRow.className = 'table-light';
    
    data[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);
    
    // Create body rows
    const tbody = document.createElement('tbody');
    data.slice(1).forEach(rowData => {
        const row = document.createElement('tr');
        rowData.forEach(cell => {
            const td = document.createElement('td');
            // Format cell content
            if (cell && typeof cell === 'string' && cell.startsWith('=')) {
                td.className = 'text-primary';
                td.innerHTML = `<code>${cell}</code>`;
            } else {
                td.textContent = cell;
            }
            row.appendChild(td);
        });
        tbody.appendChild(row);
    });
    table.appendChild(tbody);
    
    tableWrapper.appendChild(table);
    previewContent.innerHTML = '';
    previewContent.appendChild(tableWrapper);
}