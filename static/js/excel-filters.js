/**
 * Excel-like filtering and sorting capabilities
 */

document.addEventListener('DOMContentLoaded', function() {
    // Add filter buttons to column headers
    const headerCells = document.querySelectorAll('.sheet-header');
    headerCells.forEach((header, index) => {
        addFilterButton(header, index);
    });
    
    /**
     * Add a filter button to a header cell
     */
    function addFilterButton(headerCell, colIndex) {
        // Create filter button
        const filterBtn = document.createElement('button');
        filterBtn.className = 'filter-btn';
        filterBtn.innerHTML = '<i class="bi bi-filter"></i>';
        filterBtn.title = 'Filter';
        
        // Add filter menu
        filterBtn.addEventListener('click', function(e) {
            e.stopPropagation();
            showFilterMenu(headerCell, colIndex);
        });
        
        // Add to header
        headerCell.appendChild(filterBtn);
    }
    
    /**
     * Show the filter menu for a column
     */
    function showFilterMenu(headerCell, colIndex) {
        // Remove any existing filter menu
        const existingMenu = document.querySelector('.filter-menu');
        if (existingMenu) {
            existingMenu.remove();
        }
        
        // Get all values in this column
        const uniqueValues = getUniqueColumnValues(colIndex);
        
        // Create filter menu
        const filterMenu = document.createElement('div');
        filterMenu.className = 'filter-menu';
        
        // Position menu below header
        const headerRect = headerCell.getBoundingClientRect();
        filterMenu.style.top = (headerRect.bottom + window.scrollY) + 'px';
        filterMenu.style.left = (headerRect.left + window.scrollX) + 'px';
        
        // Add sort options
        const sortingDiv = document.createElement('div');
        sortingDiv.className = 'filter-section';
        sortingDiv.innerHTML = `
            <div class="filter-header">Sort</div>
            <div class="filter-option" data-action="sort-asc">
                <i class="bi bi-sort-alpha-down"></i> Sort A to Z
            </div>
            <div class="filter-option" data-action="sort-desc">
                <i class="bi bi-sort-alpha-down-alt"></i> Sort Z to A
            </div>
            <div class="filter-option" data-action="sort-num-asc">
                <i class="bi bi-sort-numeric-down"></i> Sort Smallest to Largest
            </div>
            <div class="filter-option" data-action="sort-num-desc">
                <i class="bi bi-sort-numeric-down-alt"></i> Sort Largest to Smallest
            </div>
        `;
        
        // Add filter options
        const filteringDiv = document.createElement('div');
        filteringDiv.className = 'filter-section';
        filteringDiv.innerHTML = `
            <div class="filter-header">Filter</div>
            <div class="filter-search">
                <input type="text" placeholder="Search values" class="filter-search-input">
            </div>
            <div class="filter-values">
                <div class="filter-value-option">
                    <input type="checkbox" id="select-all" checked>
                    <label for="select-all">(Select All)</label>
                </div>
            </div>
            <div class="filter-actions">
                <button class="btn btn-sm btn-primary filter-apply">Apply</button>
                <button class="btn btn-sm btn-secondary filter-clear">Clear</button>
            </div>
        `;
        
        // Add all unique values to filter
        const filterValues = filteringDiv.querySelector('.filter-values');
        uniqueValues.forEach((value, i) => {
            const valueOption = document.createElement('div');
            valueOption.className = 'filter-value-option';
            valueOption.innerHTML = `
                <input type="checkbox" id="filter-value-${i}" value="${value}" checked>
                <label for="filter-value-${i}">${value}</label>
            `;
            filterValues.appendChild(valueOption);
        });
        
        // Add event listeners to sort options
        sortingDiv.querySelectorAll('.filter-option').forEach(option => {
            option.addEventListener('click', function() {
                const action = this.getAttribute('data-action');
                
                switch(action) {
                    case 'sort-asc':
                        sortColumn(colIndex, 'asc', false);
                        break;
                    case 'sort-desc':
                        sortColumn(colIndex, 'desc', false);
                        break;
                    case 'sort-num-asc':
                        sortColumn(colIndex, 'asc', true);
                        break;
                    case 'sort-num-desc':
                        sortColumn(colIndex, 'desc', true);
                        break;
                }
                
                filterMenu.remove();
            });
        });
        
        // Add event listener to select all checkbox
        const selectAll = filteringDiv.querySelector('#select-all');
        selectAll.addEventListener('change', function() {
            const checkboxes = filteringDiv.querySelectorAll('.filter-value-option input[type="checkbox"]:not(#select-all)');
            checkboxes.forEach(checkbox => {
                checkbox.checked = this.checked;
            });
        });
        
        // Add event listener to filter search
        const searchInput = filteringDiv.querySelector('.filter-search-input');
        searchInput.addEventListener('input', function() {
            const searchTerm = this.value.toLowerCase();
            const valueOptions = filteringDiv.querySelectorAll('.filter-value-option:not(:first-child)');
            
            valueOptions.forEach(option => {
                const label = option.querySelector('label').textContent.toLowerCase();
                if (label.includes(searchTerm)) {
                    option.style.display = '';
                } else {
                    option.style.display = 'none';
                }
            });
        });
        
        // Add event listener to apply filter
        const applyBtn = filteringDiv.querySelector('.filter-apply');
        applyBtn.addEventListener('click', function() {
            // Get selected values
            const selectedValues = [];
            const checkboxes = filteringDiv.querySelectorAll('.filter-value-option:not(:first-child) input[type="checkbox"]');
            
            checkboxes.forEach(checkbox => {
                if (checkbox.checked) {
                    selectedValues.push(checkbox.value);
                }
            });
            
            // Apply filter
            filterColumn(colIndex, selectedValues);
            filterMenu.remove();
        });
        
        // Add event listener to clear filter
        const clearBtn = filteringDiv.querySelector('.filter-clear');
        clearBtn.addEventListener('click', function() {
            clearFilter();
            filterMenu.remove();
        });
        
        // Add sections to menu
        filterMenu.appendChild(sortingDiv);
        filterMenu.appendChild(filteringDiv);
        
        // Add menu to document
        document.body.appendChild(filterMenu);
        
        // Close menu when clicking outside
        document.addEventListener('click', function closeFilterMenu(e) {
            if (!filterMenu.contains(e.target) && e.target !== filterBtn) {
                filterMenu.remove();
                document.removeEventListener('click', closeFilterMenu);
            }
        });
    }
    
    /**
     * Get all unique values in a column
     */
    function getUniqueColumnValues(colIndex) {
        const cells = document.querySelectorAll(`td[data-col="${colIndex}"]`);
        const values = new Set();
        
        cells.forEach(cell => {
            const editableDiv = cell.querySelector('.editable-cell');
            if (editableDiv) {
                values.add(editableDiv.textContent.trim());
            }
        });
        
        return Array.from(values).sort();
    }
    
    /**
     * Sort a column
     */
    function sortColumn(colIndex, direction = 'asc', numeric = false) {
        const table = document.getElementById('sheetTable');
        const tbody = table.querySelector('tbody');
        const rows = Array.from(tbody.querySelectorAll('tr'));
        
        rows.sort((rowA, rowB) => {
            // Get cell values
            const cellA = rowA.querySelector(`td[data-col="${colIndex}"] .editable-cell`);
            const cellB = rowB.querySelector(`td[data-col="${colIndex}"] .editable-cell`);
            
            if (!cellA || !cellB) return 0;
            
            let valA = cellA.textContent.trim();
            let valB = cellB.textContent.trim();
            
            // Convert to numbers if numeric sort
            if (numeric) {
                valA = parseFloat(valA) || 0;
                valB = parseFloat(valB) || 0;
            }
            
            // Compare values
            if (valA === valB) return 0;
            
            if (direction === 'asc') {
                return valA < valB ? -1 : 1;
            } else {
                return valA > valB ? -1 : 1;
            }
        });
        
        // Re-append rows in sorted order
        rows.forEach(row => {
            tbody.appendChild(row);
        });
        
        showToast(`Column sorted ${direction === 'asc' ? 'ascending' : 'descending'}`, 'success');
    }
    
    /**
     * Filter a column to show only specific values
     */
    function filterColumn(colIndex, allowedValues) {
        clearFilter(); // Clear any existing filters
        
        const table = document.getElementById('sheetTable');
        const rows = table.querySelectorAll('tbody tr');
        
        rows.forEach(row => {
            const cell = row.querySelector(`td[data-col="${colIndex}"] .editable-cell`);
            if (cell) {
                const value = cell.textContent.trim();
                
                if (!allowedValues.includes(value)) {
                    row.classList.add('filtered-out');
                }
            }
        });
        
        // Add filter indicator to header
        const header = document.querySelector(`.sheet-header[data-col="${colIndex}"]`);
        if (header) {
            header.classList.add('filtered');
            
            // Update filter button appearance
            const filterBtn = header.querySelector('.filter-btn');
            if (filterBtn) {
                filterBtn.classList.add('active');
            }
        }
        
        showToast(`Filter applied to column`, 'success');
    }
    
    /**
     * Clear all filters
     */
    function clearFilter() {
        // Show all rows
        const filteredRows = document.querySelectorAll('.filtered-out');
        filteredRows.forEach(row => {
            row.classList.remove('filtered-out');
        });
        
        // Remove filter indicators
        const filteredHeaders = document.querySelectorAll('.sheet-header.filtered');
        filteredHeaders.forEach(header => {
            header.classList.remove('filtered');
            
            // Reset filter button appearance
            const filterBtn = header.querySelector('.filter-btn');
            if (filterBtn) {
                filterBtn.classList.remove('active');
            }
        });
    }
});
