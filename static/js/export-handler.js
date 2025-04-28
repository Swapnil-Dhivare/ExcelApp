/**
 * Enhanced export functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize export functionality
    initExportFunctionality();
    
    /**
     * Initialize export functionality
     */
    function initExportFunctionality() {
        // Get all export buttons
        document.querySelectorAll('.export-link').forEach(button => {
            button.addEventListener('click', handleExport);
        });
    }
    
    /**
     * Handle export action
     */
    function handleExport(e) {
        e.preventDefault();
        
        const url = this.getAttribute('href');
        const format = this.getAttribute('data-format') || this.getAttribute('href').split('format_type=')[1];
        
        // Show loading spinner in the dropdown item
        const originalHTML = this.innerHTML;
        this.innerHTML = `<span class="spinner-border spinner-border-sm me-2" role="status" aria-hidden="true"></span> Exporting...`;
        this.style.pointerEvents = 'none';
        
        // Close dropdown to prevent clutter
        const dropdown = this.closest('.dropdown-menu');
        if (dropdown) {
            const bsDropdown = bootstrap.Dropdown.getInstance(dropdown.previousElementSibling);
            if (bsDropdown) {
                setTimeout(() => {
                    bsDropdown.hide();
                }, 300);
            }
        }
        
        // Show loading toast
        showToast(`Preparing ${format.toUpperCase()} export...`, 'info');
        
        // Perform export
        fetch(url, {
            method: 'GET',
            headers: {
                'X-Requested-With': 'XMLHttpRequest'
            }
        })
        .then(response => {
            if (!response.ok) {
                throw new Error(`Export failed: ${response.statusText}`);
            }
            
            // Check if this is a regular download or a JSON response
            const contentType = response.headers.get('Content-Type');
            
            if (contentType && contentType.includes('application/json')) {
                return response.json().then(data => {
                    if (data.success) {
                        showToast(`${format.toUpperCase()} export successful!`, 'success');
                    } else {
                        throw new Error(data.message || 'Export failed');
                    }
                });
            } else {
                // It's a file download
                return response.blob().then(blob => {
                    // Create object URL
                    const url = window.URL.createObjectURL(blob);
                    
                    // Get filename from headers or create a default one
                    let filename;
                    const disposition = response.headers.get('Content-Disposition');
                    
                    if (disposition && disposition.indexOf('filename=') !== -1) {
                        filename = disposition.split('filename=')[1].replace(/"/g, '').replace(/'/g, '');
                    } else {
                        // Default filename
                        const date = new Date().toISOString().split('T')[0];
                        filename = `export_${date}.${format}`;
                    }
                    
                    // Create download link
                    const link = document.createElement('a');
                    link.href = url;
                    link.download = filename;
                    document.body.appendChild(link);
                    link.click();
                    
                    // Clean up
                    setTimeout(() => {
                        window.URL.revokeObjectURL(url);
                        document.body.removeChild(link);
                    }, 100);
                    
                    showToast(`${format.toUpperCase()} downloaded successfully!`, 'success');
                });
            }
        })
        .catch(error => {
            console.error('Export error:', error);
            showToast(`Export failed: ${error.message}`, 'danger');
        })
        .finally(() => {
            // Restore button
            this.innerHTML = originalHTML;
            this.style.pointerEvents = '';
        });
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
