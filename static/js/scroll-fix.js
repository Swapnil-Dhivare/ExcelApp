/**
 * Fix scrolling issues in the application
 */
document.addEventListener('DOMContentLoaded', function() {
    // Fix body scroll locking from dropdowns
    fixBodyScrollLock();
    
    // Fix touch scrolling on mobile
    fixTouchScrolling();
    
    // Ensure navbar doesn't block scrolling
    adjustNavbarScroll();
    
    /**
     * Prevent dropdowns from locking body scrolling
     */
    function fixBodyScrollLock() {
        // Detect when a dropdown is opened
        document.querySelectorAll('.dropdown-toggle').forEach(toggle => {
            toggle.addEventListener('shown.bs.dropdown', function() {
                if (!document.body.classList.contains('sheet-editor-page')) {
                    document.body.style.overflow = 'auto';
                }
            });
            
            toggle.addEventListener('hidden.bs.dropdown', function() {
                if (!document.body.classList.contains('sheet-editor-page')) {
                    document.body.style.overflow = 'auto';
                }
            });
        });
        
        // Make sure modals don't lock scrolling of body when closed
        document.addEventListener('hidden.bs.modal', function() {
            if (!document.body.classList.contains('sheet-editor-page')) {
                document.body.style.overflow = 'auto';
            }
        });
    }
    
    /**
     * Fix touch scrolling on mobile devices
     */
    function fixTouchScrolling() {
        // Enable momentum scrolling on iOS
        document.querySelectorAll('.table-responsive, .excel-workspace, .modal-body').forEach(elem => {
            elem.style.WebkitOverflowScrolling = 'touch';
        });
    }
    
    /**
     * Adjust padding when navbar is fixed
     */
    function adjustNavbarScroll() {
        const navbar = document.querySelector('.navbar');
        const main = document.querySelector('main');
        
        if (navbar && main) {
            // Only apply this to non-sheet editor pages
            if (!document.body.classList.contains('sheet-editor-page')) {
                // Add padding to main equal to navbar height
                main.style.paddingTop = navbar.offsetHeight + 'px';
            }
        }
    }
    
    // Add scroll to top button for long pages
    addScrollToTopButton();
    
    function addScrollToTopButton() {
        // Only add for non-sheet-editor pages
        if (document.body.classList.contains('sheet-editor-page')) {
            return;
        }
        
        // Create button
        const scrollTopBtn = document.createElement('button');
        scrollTopBtn.className = 'btn btn-primary scroll-to-top-btn';
        scrollTopBtn.innerHTML = '<i class="bi bi-arrow-up"></i>';
        scrollTopBtn.title = 'Scroll to top';
        scrollTopBtn.style.position = 'fixed';
        scrollTopBtn.style.bottom = '20px';
        scrollTopBtn.style.right = '20px';
        scrollTopBtn.style.opacity = '0';
        scrollTopBtn.style.visibility = 'hidden';
        scrollTopBtn.style.transition = 'opacity 0.3s, visibility 0.3s';
        scrollTopBtn.style.zIndex = '1000';
        
        // Add to body
        document.body.appendChild(scrollTopBtn);
        
        // Add event listener
        scrollTopBtn.addEventListener('click', function() {
            window.scrollTo({
                top: 0,
                behavior: 'smooth'
            });
        });
        
        // Show/hide based on scroll position
        window.addEventListener('scroll', function() {
            if (window.scrollY > 300) {
                scrollTopBtn.style.opacity = '1';
                scrollTopBtn.style.visibility = 'visible';
            } else {
                scrollTopBtn.style.opacity = '0';
                scrollTopBtn.style.visibility = 'hidden';
            }
        });
    }
});
