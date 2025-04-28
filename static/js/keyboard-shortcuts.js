/**
 * Handle keyboard shortcuts modal and functionality
 */
document.addEventListener('DOMContentLoaded', function() {
    // Initialize shortcuts button
    const shortcutsBtn = document.getElementById('keyboardShortcutsBtn');
    if (shortcutsBtn) {
        shortcutsBtn.addEventListener('click', showShortcutsModal);
    }
    
    /**
     * Show keyboard shortcuts modal
     */
    function showShortcutsModal() {
        // Create shortcuts modal if it doesn't exist
        if (!document.getElementById('keyboardShortcutsModal')) {
            createShortcutsModal();
        }
        
        // Show the modal
        const modal = new bootstrap.Modal(document.getElementById('keyboardShortcutsModal'));
        modal.show();
    }
    
    /**
     * Create keyboard shortcuts modal
     */
    function createShortcutsModal() {
        // Define keyboard shortcuts
        const shortcuts = [
            { key: 'Ctrl+S', action: 'Save sheet' },
            { key: 'Ctrl+Z', action: 'Undo (when implemented)' },
            { key: 'Ctrl+Y', action: 'Redo (when implemented)' },
            { key: 'Ctrl+B', action: 'Bold text' },
            { key: 'Ctrl+I', action: 'Italic text' },
            { key: 'Ctrl+U', action: 'Underline text' },
            { key: 'Delete', action: 'Clear cell content' },
            { key: 'Tab', action: 'Move to next cell' },
            { key: 'Shift+Tab', action: 'Move to previous cell' },
            { key: 'Enter', action: 'Move down/Confirm edit' },
            { key: 'Arrow keys', action: 'Navigate between cells' },
            { key: 'Ctrl+Arrow keys', action: 'Scroll the sheet' },
            { key: 'Shift+Arrow keys', action: 'Extend selection' },
            { key: 'Ctrl+A', action: 'Select all cells' },
            { key: 'F2', action: 'Edit cell' },
            { key: 'Escape', action: 'Cancel editing/Clear selection' },
        ];
        
        // Create modal HTML
        let modalHTML = `
        <div class="modal fade" id="keyboardShortcutsModal" tabindex="-1" aria-labelledby="shortcutsModalLabel" aria-hidden="true">
            <div class="modal-dialog">
                <div class="modal-content">
                    <div class="modal-header bg-primary text-white">
                        <h5 class="modal-title" id="shortcutsModalLabel">Keyboard Shortcuts</h5>
                        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal" aria-label="Close"></button>
                    </div>
                    <div class="modal-body">
                        <table class="table table-striped table-hover">
                            <thead class="table-light">
                                <tr>
                                    <th>Key Combination</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody>
        `;
        
        // Add each shortcut to the table
        shortcuts.forEach(shortcut => {
            modalHTML += `
                <tr>
                    <td><kbd>${shortcut.key}</kbd></td>
                    <td>${shortcut.action}</td>
                </tr>
            `;
        });
        
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
        
        // Append the modal to the document
        const modalContainer = document.createElement('div');
        modalContainer.innerHTML = modalHTML;
        document.body.appendChild(modalContainer);
    }
    
    // Add keyboard shortcuts
    document.addEventListener('keydown', function(e) {
        // Skip if the user is editing text
        if (e.target.isContentEditable || e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
            return;
        }
        
        // Implement keyboard shortcuts
        if (e.ctrlKey || e.metaKey) {
            switch (e.key.toLowerCase()) {
                case 'b':
                    e.preventDefault();
                    if (window.excelFormatting) {
                        window.excelFormatting.applyFormatting('bold');
                    }
                    break;
                case 'i':
                    e.preventDefault();
                    if (window.excelFormatting) {
                        window.excelFormatting.applyFormatting('italic');
                    }
                    break;
                case 'u':
                    e.preventDefault();
                    if (window.excelFormatting) {
                        window.excelFormatting.applyFormatting('underline');
                    }
                    break;
                case 'a':
                    e.preventDefault();
                    selectAllCells();
                    break;
            }
        } else if (e.key === 'F2') {
            e.preventDefault();
            editSelectedCell();
        }
    });
    
    /**
     * Select all cells in the sheet
     */
    function selectAllCells() {
        const cells = document.querySelectorAll('.sheet-cell');
        
        cells.forEach(cell => {
            cell.classList.add('selected');
        });
        
        window.selectedCells = Array.from(cells);
        
        // Update selection info if available
        if (typeof window.updateSelectionInfo === 'function') {
            window.updateSelectionInfo();
        }
        
        // Show notification
        showToast('All cells selected', 'info');
    }
    
    /**
     * Edit the currently selected cell
     */
    function editSelectedCell() {
        if (window.selectedCells && window.selectedCells.length === 1) {
            const editableDiv = window.selectedCells[0].querySelector('.editable-cell');
            if (editableDiv) {
                editableDiv.focus();
            }
        }
    }
    
    /**
     * Show toast notification
     */
    function showToast(message, type = 'info') {
        if (typeof window.showToast === 'function') {
            window.showToast(message, type);
        } else {
            console.log(`${type}: ${message}`);
        }
    }
});
