/**
 * Fix dropdown positioning to ensure they don't go off-screen
 */
document.addEventListener('DOMContentLoaded', function() {
    // Fix dropdowns that are created dynamically (like context menus)
    const observer = new MutationObserver(function(mutations) {
        mutations.forEach(function(mutation) {
            if (mutation.addedNodes) {
                mutation.addedNodes.forEach(function(node) {
                    if (node.classList && 
                        (node.classList.contains('dropdown-menu') || 
                         node.classList.contains('context-menu') ||
                         node.classList.contains('filter-menu'))) {
                        fixDropdownPosition(node);
                    }
                });
            }
        });
    });
    
    // Start observing for dynamically added dropdowns
    observer.observe(document.body, { childList: true, subtree: true });
    
    // Fix dropdown position
    function fixDropdownPosition(dropdown) {
        // Get viewport dimensions
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        
        // Get dropdown dimensions and position
        const rect = dropdown.getBoundingClientRect();
        
        // Check if dropdown goes off the right edge
        if (rect.right > viewportWidth) {
            dropdown.style.left = 'auto';
            dropdown.style.right = '10px';
        }
        
        // Check if dropdown goes off the bottom
        if (rect.bottom > viewportHeight) {
            // If the dropdown is taller than the viewport, position at top
            if (rect.height > viewportHeight * 0.8) {
                dropdown.style.top = '10px';
                dropdown.style.maxHeight = (viewportHeight - 20) + 'px';
            } else {
                // Otherwise, adjust position to be above the trigger
                dropdown.style.top = 'auto';
                dropdown.style.bottom = (viewportHeight - rect.top) + 'px';
            }
        }
        
        // Ensure action buttons stay visible
        const actionButtons = dropdown.querySelector('.filter-actions, .modal-footer, .btn-group');
        if (actionButtons) {
            actionButtons.style.position = 'sticky';
            actionButtons.style.bottom = '0';
            actionButtons.style.backgroundColor = 'white';
            actionButtons.style.padding = '8px 0';
            actionButtons.style.zIndex = '2';
            actionButtons.style.borderTop = '1px solid #eee';
        }
    }
    
    // Fix Bootstrap dropdowns
    const dropdownToggleButtons = document.querySelectorAll('.dropdown-toggle');
    dropdownToggleButtons.forEach(function(button) {
        button.addEventListener('shown.bs.dropdown', function() {
            const dropdown = this.nextElementSibling;
            if (dropdown && dropdown.classList.contains('dropdown-menu')) {
                fixDropdownPosition(dropdown);
            }
        });
    });
    
    // Fix context menus
    document.addEventListener('contextmenu', function(e) {
        // Wait a bit for the menu to be created
        setTimeout(function() {
            const contextMenus = document.querySelectorAll('.context-menu, .filter-menu');
            contextMenus.forEach(fixDropdownPosition);
        }, 10);
    });
});
