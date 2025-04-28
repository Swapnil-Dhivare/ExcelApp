/**
 * Bootstrap dropdown positioning fix
 * Ensures dropdowns don't go off screen
 */
document.addEventListener('DOMContentLoaded', function() {
    // Fix dropdown positioning for all dropdowns
    document.querySelectorAll('.dropdown-toggle').forEach(dropdownToggle => {
        const dropdown = new bootstrap.Dropdown(dropdownToggle, {
            boundary: 'viewport',
            reference: 'toggle',
            display: 'dynamic'
        });
        
        // Manually position dropdown if it goes off screen
        dropdownToggle.addEventListener('shown.bs.dropdown', function() {
            const menu = this.nextElementSibling;
            if (!menu || !menu.classList.contains('dropdown-menu')) return;
            
            // Get position info
            const menuRect = menu.getBoundingClientRect();
            const toggleRect = this.getBoundingClientRect();
            const windowWidth = window.innerWidth;
            const windowHeight = window.innerHeight;
            
            // Fix horizontal positioning
            if (menuRect.right > windowWidth) {
                if (menuRect.width > windowWidth) {
                    // If menu is wider than window, center it and make it scrollable
                    menu.style.left = '10px';
                    menu.style.right = '10px';
                    menu.style.width = 'calc(100vw - 20px)';
                    menu.style.maxWidth = 'none';
                } else {
                    // Otherwise align to right
                    menu.style.left = 'auto';
                    menu.style.right = '0';
                }
            }
            
            // Fix vertical positioning
            if (menuRect.bottom > windowHeight) {
                if (menuRect.height > windowHeight) {
                    // If menu is taller than window, position at top and make it scrollable
                    menu.style.top = '10px';
                    menu.style.maxHeight = 'calc(100vh - 20px)';
                } else {
                    // Otherwise position above toggle
                    menu.style.top = 'auto';
                    menu.style.transform = `translateY(-${menuRect.height}px)`;
                }
            }
        });
    });
});
