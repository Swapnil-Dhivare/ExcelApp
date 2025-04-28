/**
 * Fix for formatting panel display toggle
 */
document.addEventListener('DOMContentLoaded', function() {
    // Get the format button and panel
    const formatButton = document.getElementById('toggleFormatPanel');
    const formatPanel = document.querySelector('.formatting-panel');
    
    if (formatButton && formatPanel) {
        // Make sure the format panel is initially hidden
        formatPanel.style.display = 'none';
        
        // Remove any existing click handlers
        formatButton.removeEventListener('click', toggleFormatPanel);
        
        // Add new click handler
        formatButton.addEventListener('click', toggleFormatPanel);
        
        // Define toggle function
        function toggleFormatPanel(e) {
            if (e) {
                e.preventDefault();
                e.stopPropagation();
            }
            
            // Toggle format panel visibility
            if (formatPanel.style.display === 'none' || !formatPanel.style.display) {
                formatPanel.style.display = 'block';
                formatButton.classList.add('active');
                console.log('Format panel shown');
            } else {
                formatPanel.style.display = 'none';
                formatButton.classList.remove('active');
                console.log('Format panel hidden');
            }
        }
    } else {
        console.error('Format button or panel not found');
    }
    
    // Define global toggleFormatPanel function to ensure it can be called from inline handlers
    window.toggleFormatPanel = function() {
        const panel = document.querySelector('.formatting-panel');
        const btn = document.getElementById('toggleFormatPanel');
        
        if (panel && btn) {
            if (panel.style.display === 'none' || !panel.style.display) {
                panel.style.display = 'block';
                btn.classList.add('active');
            } else {
                panel.style.display = 'none';
                btn.classList.remove('active');
            }
        }
    };
});
