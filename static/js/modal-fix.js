/**
 * Bootstrap Modal Backdrop Fix
 * Resolves the "Cannot read properties of undefined (reading 'backdrop')" error
 */
document.addEventListener('DOMContentLoaded', function() {
    // Create a backup modal backdrop if needed
    const backupBackdrop = document.getElementById('backup-backdrop');
    
    // Fix for Bootstrap modal backdrop issues
    const fixModals = function() {
        // Get all modal trigger elements
        const modalTriggers = document.querySelectorAll('[data-bs-toggle="modal"]');
        
        modalTriggers.forEach(trigger => {
            trigger.addEventListener('click', function(e) {
                const targetModalId = this.getAttribute('data-bs-target');
                if (!targetModalId) return;
                
                const targetModal = document.querySelector(targetModalId);
                if (!targetModal) return;
                
                try {
                    // Try to initialize modal normally first
                    const modal = new bootstrap.Modal(targetModal);
                    modal.show();
                } catch (err) {
                    console.warn('Modal initialization error, attempting fallback: ', err);
                    
                    // Apply fallback method - make backdrop visible if needed
                    if (backupBackdrop) {
                        backupBackdrop.classList.remove('d-none');
                        backupBackdrop.classList.add('show');
                    }
                    
                    // Force show the modal with inline styles
                    targetModal.style.display = 'block';
                    targetModal.classList.add('show');
                    targetModal.setAttribute('aria-modal', 'true');
                    targetModal.setAttribute('role', 'dialog');
                    document.body.classList.add('modal-open');
                    
                    // Add event listener to close button
                    targetModal.querySelectorAll('[data-bs-dismiss="modal"]').forEach(closeBtn => {
                        closeBtn.addEventListener('click', function() {
                            targetModal.style.display = 'none';
                            targetModal.classList.remove('show');
                            if (backupBackdrop) {
                                backupBackdrop.classList.add('d-none');
                                backupBackdrop.classList.remove('show');
                            }
                            document.body.classList.remove('modal-open');
                        });
                    });
                }
            });
        });
    };
    
    // Run the fix
    fixModals();
});
