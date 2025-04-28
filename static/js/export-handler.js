/**
 * Excel export functionality to ensure files download properly
 */
document.addEventListener('DOMContentLoaded', function() {
    // Attach event handlers to export links
    initializeExportLinks();
    
    /**
     * Initialize event handlers for all export links
     */
    function initializeExportLinks() {
        // Find all export links in the document
        document.querySelectorAll('a[href*="export_sheet"]').forEach(link => {
            link.addEventListener('click', handleExport);
        });
    }
    
    /**
     * Handle export link clicks
     */
    function handleExport(e) {
        e.preventDefault();
        
        // Get export URL and format
        const url = this.getAttribute('href');
        const format = url.split('format_type=')[1];
        
        // Show loading indicator
        showToast(`Preparing ${format.toUpperCase()} export...`, 'info');
        
        // Create a hidden iframe to handle the download
        const downloadFrame = document.createElement('iframe');
        downloadFrame.style.display = 'none';
        document.body.appendChild(downloadFrame);
        
        // Create download form for POST request
        const form = document.createElement('form');
        form.method = 'GET';
        form.action = url;
        form.target = downloadFrame.name;
        
        // Add form to document and submit
        document.body.appendChild(form);
        
        // Set unique name for iframe
        downloadFrame.name = 'download_frame_' + Date.now();
        
        // Add a notification for the user
        showToast('Your download will begin shortly. If it doesn\'t start automatically, check your browser\'s download settings.', 'info');
        
        // Submit the form to start the download
        form.submit();
        
        // Clean up after a delay
        setTimeout(() => {
            if (form.parentNode) form.parentNode.removeChild(form);
            if (downloadFrame.parentNode) downloadFrame.parentNode.removeChild(downloadFrame);
            showToast(`${format.toUpperCase()} export completed! Check your downloads folder.`, 'success');
        }, 3000);
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type) {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            // Fallback toast implementation
            console.log(`${type}: ${message}`);
            
            const toastContainer = document.querySelector('.toast-container') || 
                createToastContainer();
            
            const toast = document.createElement('div');
            toast.className = `toast show mb-2`;
            toast.setAttribute('role', 'alert');
            
            const bgClass = type === 'error' ? 'bg-danger' : 
                            type === 'success' ? 'bg-success' : 'bg-info';
            
            toast.innerHTML = `
                <div class="toast-header ${bgClass} text-white">
                    <strong class="me-auto">Excel Generator</strong>
                    <button type="button" class="btn-close btn-close-white" data-bs-dismiss="toast"></button>
                </div>
                <div class="toast-body bg-light">
                    ${message}
                </div>
            `;
            
            toastContainer.appendChild(toast);
            
            // Auto-remove toast after a delay
            setTimeout(() => {
                if (toast.parentNode) toast.parentNode.removeChild(toast);
            }, 5000);
        }
    }
    
    /**
     * Create toast container if it doesn't exist
     */
    function createToastContainer() {
        const container = document.createElement('div');
        container.className = 'toast-container position-fixed top-0 end-0 p-3';
        container.style.zIndex = '1050';
        document.body.appendChild(container);
        return container;
    }
});
