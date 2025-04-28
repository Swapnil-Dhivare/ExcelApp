/**
 * Cleanup function for Excel UI elements
 * Ensures no tooltips, indicators or overlays get stuck on screen
 */

document.addEventListener('DOMContentLoaded', function() {
    // Clean up any lingering UI elements
    cleanupUIElements();
    
    // Set periodic cleanup
    setInterval(cleanupUIElements, 10000); // Run every 10 seconds
    
    // Clean up when window gets focus (in case sizing happens while window is inactive)
    window.addEventListener('focus', cleanupUIElements);
    
    // Clean up when switching tabs
    document.addEventListener('visibilitychange', function() {
        if (!document.hidden) {
            cleanupUIElements();
        }
    });
    
    /**
     * Remove any UI elements that might get stuck
     */
    function cleanupUIElements() {
        // Remove size indicators
        document.querySelectorAll('.size-indicator').forEach(el => {
            if (el && el.parentNode) {
                el.parentNode.removeChild(el);
            }
        });
        
        // Hide resize previews
        const resizePreview = document.getElementById('resize-preview');
        if (resizePreview) {
            resizePreview.style.display = 'none';
        }
        
        const rowResizePreview = document.getElementById('row-resize-preview');
        if (rowResizePreview) {
            rowResizePreview.style.display = 'none';
        }
        
        // Hide resize overlay
        const resizeOverlay = document.getElementById('resize-overlay');
        if (resizeOverlay) {
            resizeOverlay.style.display = 'none';
        }
        
        // Remove any selection highlighting that might be stuck
        document.querySelectorAll('.col-highlight, .row-highlight, .resizing-column, .resizing-row').forEach(el => {
            el.classList.remove('col-highlight', 'row-highlight', 'resizing-column', 'resizing-row');
        });
        
        // Reset body classes
        document.body.classList.remove('resizing', 'column-resize', 'row-resize');
    }
    
    // Expose cleanup function globally
    window.cleanupUIElements = cleanupUIElements;
});
