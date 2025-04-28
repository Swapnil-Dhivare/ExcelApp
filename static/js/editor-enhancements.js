/**
 * Enhanced UI functionality for the Excel sheet editor
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize enhancements
    initEditorEnhancements();
    
    /**
     * Initialize editor enhancements
     */
    function initEditorEnhancements() {
        updateSheetStats();
        setupToggleFormatPanel();
        improveSheetScrolling();
        enhanceButtonsTooltips();
    }
    
    /**
     * Update sheet statistics in the status bar
     */
    function updateSheetStats() {
        const rowColCount = document.getElementById('rowColCount');
        if (!rowColCount) return;
        
        // Count rows and columns
        const rows = document.querySelectorAll('#sheetTable tbody tr').length;
        const cols = document.querySelectorAll('#sheetTable thead th').length;
        
        rowColCount.textContent = `${rows} rows × ${cols} columns`;
    }
    
    /**
     * Setup toggle behavior for format panel
     */
    function setupToggleFormatPanel() {
        // Toggle format panel when button is clicked
        const toggleBtn = document.getElementById('toggleFormatPanel');
        const formatPanel = document.querySelector('.formatting-panel');
        
        if (toggleBtn && formatPanel) {
            toggleBtn.addEventListener('click', function() {
                if (formatPanel.style.display === 'none') {
                    formatPanel.style.display = 'block';
                    toggleBtn.classList.add('active');
                    // Save state
                    localStorage.setItem('format_panel_visible', 'true');
                } else {
                    formatPanel.style.display = 'none';
                    toggleBtn.classList.remove('active');
                    // Save state
                    localStorage.setItem('format_panel_visible', 'false');
                }
            });
            
            // Check saved state
            const savedState = localStorage.getItem('format_panel_visible');
            if (savedState === 'true') {
                formatPanel.style.display = 'block';
                toggleBtn.classList.add('active');
            } else {
                formatPanel.style.display = 'none';
                toggleBtn.classList.remove('active');
            }
        }
    }
    
    /**
     * Improve sheet scrolling behavior
     */
    function improveSheetScrolling() {
        const workspace = document.querySelector('.excel-workspace');
        if (!workspace) return;
        
        // Make scrolling smoother
        workspace.style.scrollBehavior = 'smooth';
        
        // Add keyboard navigation for scrolling
        document.addEventListener('keydown', function(e) {
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.contentEditable === 'true') {
                return;
            }
            
            const scrollDistance = 100;
            
            switch(e.key) {
                case 'PageDown':
                    workspace.scrollBy(0, workspace.clientHeight);
                    break;
                case 'PageUp':
                    workspace.scrollBy(0, -workspace.clientHeight);
                    break;
                case 'End':
                    if (e.ctrlKey) {
                        workspace.scrollTo(0, workspace.scrollHeight);
                    }
                    break;
                case 'Home':
                    if (e.ctrlKey) {
                        workspace.scrollTo(0, 0);
                    }
                    break;
                case 'ArrowRight':
                    if (e.ctrlKey) workspace.scrollBy(scrollDistance, 0);
                    break;
                case 'ArrowLeft':
                    if (e.ctrlKey) workspace.scrollBy(-scrollDistance, 0);
                    break;
                case 'ArrowDown':
                    if (e.ctrlKey) workspace.scrollBy(0, scrollDistance);
                    break;
                case 'ArrowUp':
                    if (e.ctrlKey) workspace.scrollBy(0, -scrollDistance);
                    break;
            }
        });
    }
    
    /**
     * Enhance buttons with better tooltips
     */
    function enhanceButtonsTooltips() {
        document.querySelectorAll('[title]').forEach(el => {
            // Format the title if it contains keyboard shortcut
            if (el.title.includes('(')) {
                const parts = el.title.split('(');
                const action = parts[0].trim();
                const shortcut = parts[1].replace(')', '').trim();
                
                el.setAttribute('data-action', action);
                el.setAttribute('data-shortcut', shortcut);
                
                // For better tooltip display, we'll handle this with CSS
                el.title = `${action} (${shortcut})`;
            }
        });
    }
    
    // Update status on selection change
    if (typeof window.updateSelectionInfo === 'function') {
        const originalUpdateSelectionInfo = window.updateSelectionInfo;
        window.updateSelectionInfo = function() {
            originalUpdateSelectionInfo.apply(window);
            updateSheetStats();
        };
    }
    
    // Update status when sheet is modified
    const originalMarkSheetModified = window.markSheetModified;
    window.markSheetModified = function() {
        if (originalMarkSheetModified) {
            originalMarkSheetModified.apply(window);
        }
        
        // Update status text
        const statusElement = document.getElementById('sheetStatus');
        if (statusElement) {
            statusElement.textContent = 'Modified';
            statusElement.classList.add('text-warning');
        }
        
        // Add unsaved class to sheet title
        const sheetTitle = document.querySelector('.excel-header h5');
        if (sheetTitle) {
            sheetTitle.classList.add('unsaved');
        }
        
        updateSheetStats();
    };
    
    // Create keyboard shortcut map
    const keyboardShortcuts = {
        'Ctrl+S': 'Save sheet',
        'Ctrl+Z': 'Undo',
        'Ctrl+Y': 'Redo',
        'Ctrl+B': 'Bold',
        'Ctrl+I': 'Italic',
        'Ctrl+U': 'Underline',
        'Delete': 'Clear cell content',
        'F2': 'Edit cell',
        'Enter': 'Confirm edit',
        'Escape': 'Cancel edit',
        'Tab': 'Move right',
        'Shift+Tab': 'Move left',
        'Arrow keys': 'Navigate cells'
    };
    
    // Initialize keyboard shortcuts modal
    const shortcutsBtn = document.getElementById('keyboardShortcutsBtn');
    if (shortcutsBtn) {
        shortcutsBtn.addEventListener('click', function() {
            let modalHTML = `
            <div class="modal fade" id="keyboardShortcutsModal" tabindex="-1">
                <div class="modal-dialog">
                    <div class="modal-content">
                        <div class="modal-header bg-primary text-white">
                            <h5 class="modal-title">Keyboard Shortcuts</h5>
                            <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
                        </div>
                        <div class="modal-body">
                            <table class="table table-striped">
                                <thead>
                                    <tr>
                                        <th>Shortcut</th>
                                        <th>Action</th>
                                    </tr>
                                </thead>
                                <tbody>
            `;
            
            for (const [key, action] of Object.entries(keyboardShortcuts)) {
                modalHTML += `
                    <tr>
                        <td><kbd>${key}</kbd></td>
                        <td>${action}</td>
                    </tr>
                `;
            }
            
            modalHTML += `
                                </tbody>
                            </table>
                        </div>
                        <div class="modal-footer">
                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Close</button>
                        </div>
                    </div>
                </div>
            </div>
            `;
            
            // Add modal to body
            const modalContainer = document.createElement('div');
            modalContainer.innerHTML = modalHTML;
            document.body.appendChild(modalContainer);
            
            // Show the modal
            const modal = new bootstrap.Modal(document.getElementById('keyboardShortcutsModal'));
            modal.show();
            
            // Remove modal from DOM when hidden
            document.getElementById('keyboardShortcutsModal').addEventListener('hidden.bs.modal', function() {
                if (modalContainer.parentNode) {
                    modalContainer.parentNode.removeChild(modalContainer);
                }
            });
        });
    }
});
