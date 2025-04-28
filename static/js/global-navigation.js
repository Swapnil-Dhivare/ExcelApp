/**
 * Global navigation functionality to ensure all links work properly
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize all dropdowns properly
    initializeDropdowns();
    
    // Fix all navigation links
    fixNavigationLinks();
    
    // Ensure active nav items are highlighted
    highlightActiveNavItems();
    
    // Improve export buttons
    improveExportButtons();
    
    /**
     * Initialize all Bootstrap dropdowns
     */
    function initializeDropdowns() {
        const dropdownToggleList = document.querySelectorAll('.dropdown-toggle');
        
        dropdownToggleList.forEach(dropdownToggle => {
            // Create new dropdown instance for each toggle
            const dropdown = new bootstrap.Dropdown(dropdownToggle, {
                boundary: 'viewport',
                reference: 'toggle',
                display: 'dynamic'
            });
            
            // Ensure dropdowns position correctly
            dropdownToggle.addEventListener('shown.bs.dropdown', function() {
                positionDropdown(this);
            });
            
            // Fix mobile dropdown click issue
            dropdownToggle.addEventListener('click', function(e) {
                // On mobile, we may need to manually toggle dropdown
                if (window.innerWidth < 992) {
                    e.preventDefault();
                    e.stopPropagation();
                    dropdown.toggle();
                }
            });
        });
        
        // Fix downloads dropdown specific issues
        const fileHistoryDropdown = document.getElementById('fileHistoryDropdown');
        if (fileHistoryDropdown) {
            fileHistoryDropdown.addEventListener('click', function(e) {
                // Prevent navigation to # and ensure dropdown opens
                e.preventDefault();
            });
        }
    }
    
    /**
     * Position dropdown menus correctly
     */
    function positionDropdown(toggleElement) {
        const dropdownMenu = toggleElement.nextElementSibling;
        if (!dropdownMenu || !dropdownMenu.classList.contains('dropdown-menu')) return;
        
        // Get viewport dimensions
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;
        
        // Get menu dimensions
        const menuRect = dropdownMenu.getBoundingClientRect();
        
        // Check if dropdown goes off screen horizontally
        if (menuRect.right > viewportWidth) {
            dropdownMenu.classList.add('dropdown-menu-end');
        }
        
        // Check if dropdown goes off screen vertically
        if (menuRect.bottom > viewportHeight) {
            // If there's enough room above the toggle, flip upward
            const toggleRect = toggleElement.getBoundingClientRect();
            if (toggleRect.top > menuRect.height) {
                dropdownMenu.classList.add('dropdown-menu-up');
            } else {
                // Otherwise just constrain the height
                dropdownMenu.style.maxHeight = `${viewportHeight - menuRect.top - 10}px`;
                dropdownMenu.style.overflowY = 'auto';
            }
        }
    }
    
    /**
     * Fix all navigation links to use absolute paths
     */
    function fixNavigationLinks() {
        document.querySelectorAll('a[href]:not([href^="#"]):not([href^="javascript"]):not([href^="http"]):not([href^="mailto"])').forEach(link => {
            // Make sure link is using an absolute path
            if (!link.href.startsWith('/') && !link.href.startsWith(window.location.origin)) {
                const href = link.getAttribute('href');
                
                // Check if href already starts with a slash
                if (!href.startsWith('/')) {
                    link.href = '/' + href;
                }
            }
        });
    }
    
    /**
     * Highlight active navigation items
     */
    function highlightActiveNavItems() {
        const currentPath = window.location.pathname;
        
        document.querySelectorAll('.navbar-nav .nav-link').forEach(navLink => {
            const href = navLink.getAttribute('href');
            
            // Remove any existing active class
            navLink.classList.remove('active');
            
            // Check if this link matches the current path
            if (href === currentPath || 
                (currentPath.startsWith(href) && href !== '/' && href.length > 1)) {
                navLink.classList.add('active');
                
                // If this is inside a dropdown, also highlight the dropdown toggle
                const dropdownParent = navLink.closest('.dropdown');
                if (dropdownParent) {
                    const dropdownToggle = dropdownParent.querySelector('.dropdown-toggle');
                    if (dropdownToggle) {
                        dropdownToggle.classList.add('active');
                    }
                }
            }
        });
    }
    
    /**
     * Improve export buttons functionality
     */
    function improveExportButtons() {
        // Get all export buttons
        const exportButtons = document.querySelectorAll('.dropdown-item[href*="export_sheet"]');
        
        exportButtons.forEach(button => {
            button.addEventListener('click', function(e) {
                e.preventDefault();
                
                const url = this.getAttribute('href');
                const format = url.split('format_type=')[1];
                
                // Show loading indicator
                const loadingToast = showLoadingToast(`Exporting as ${format.toUpperCase()}...`);
                
                // Perform fetch to server
                fetch(url)
                    .then(response => {
                        if (!response.ok) {
                            throw new Error('Export failed');
                        }
                        return response.blob();
                    })
                    .then(blob => {
                        // Create download link
                        const downloadUrl = window.URL.createObjectURL(blob);
                        const a = document.createElement('a');
                        
                        // Get filename from Content-Disposition header or create a default one
                        const filename = `excel_export_${new Date().toISOString().slice(0,10)}.${format}`;
                        
                        a.href = downloadUrl;
                        a.download = filename;
                        document.body.appendChild(a);
                        a.click();
                        window.URL.revokeObjectURL(downloadUrl);
                        a.remove();
                        
                        // Hide loading and show success
                        hideLoadingToast(loadingToast);
                        showToast(`Successfully exported as ${format.toUpperCase()}`, 'success');
                    })
                    .catch(error => {
                        console.error('Export error:', error);
                        hideLoadingToast(loadingToast);
                        showToast('Export failed. Please try again.', 'danger');
                    });
            });
        });
    }
    
    /**
     * Show loading toast notification
     */
    function showLoadingToast(message) {
        const toastContainer = document.querySelector('.toast-container');
        if (!toastContainer) return null;
        
        const toast = document.createElement('div');
        toast.className = 'toast show';
        toast.role = 'status';
        toast.setAttribute('aria-live', 'polite');
        
        toast.innerHTML = `
            <div class="toast-header bg-info text-white">
                <strong class="me-auto">Excel Generator</strong>
                <span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span>
                <span>Loading...</span>
            </div>
            <div class="toast-body">
                ${message}
            </div>
        `;
        
        toastContainer.appendChild(toast);
        return toast;
    }
    
    /**
     * Hide loading toast
     */
    function hideLoadingToast(toast) {
        if (toast && toast.parentNode) {
            toast.parentNode.removeChild(toast);
        }
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type = 'info') {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
            return;
        }
        
        const toastContainer = document.querySelector('.toast-container');
        if (!toastContainer) return;
        
        const toast = document.createElement('div');
        toast.className = `toast show mb-2`;
        toast.role = 'alert';
        toast.setAttribute('aria-live', 'assertive');
        toast.setAttribute('aria-atomic', 'true');
        
        const bgClass = type === 'danger' ? 'bg-danger' : 
                      type === 'success' ? 'bg-success' : 'bg-info';
        
        toast.innerHTML = `
            <div class="toast-header ${bgClass} text-white">
                <strong class="me-auto">Excel Generator</strong>
                <button type="button" class="btn-close btn-close-white" data-bs-dismiss="toast" aria-label="Close"></button>
            </div>
            <div class="toast-body bg-light">
                ${message}
            </div>
        `;
        
        toastContainer.appendChild(toast);
        
        toast.querySelector('.btn-close').addEventListener('click', function() {
            toast.remove();
        });
        
        setTimeout(() => {
            if (toast.parentNode) {
                toast.remove();
            }
        }, 3000);
    }
});
