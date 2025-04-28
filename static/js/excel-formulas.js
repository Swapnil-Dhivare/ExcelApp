/**
 * Excel-like formula evaluation
 */

document.addEventListener('DOMContentLoaded', function() {
    // Get reference to formula input elements
    const formulaBar = document.getElementById('formulaBar');
    const formulaBtn = document.getElementById('formulaBtn');
    
    // Add event listeners
    if (formulaBar) {
        formulaBar.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                applyFormula(this.value);
            }
        });
    }
    
    if (formulaBtn) {
        formulaBtn.addEventListener('click', function() {
            if (formulaBar) {
                applyFormula(formulaBar.value);
            }
        });
    }
    
    // Get all editable cells
    const editableCells = document.querySelectorAll('.editable-cell');
    editableCells.forEach(cell => {
        cell.addEventListener('blur', function() {
            if (this.textContent.trim().startsWith('=')) {
                // This is a formula - try to evaluate it
                const cellElement = this.closest('.sheet-cell, .sheet-header');
                evaluateFormula(this.textContent, cellElement);
            }
        });
    });
    
    /**
     * Apply formula to the currently selected cell
     */
    function applyFormula(formula) {
        if (!window.selectedCells || window.selectedCells.length !== 1) {
            showToast('Please select a single cell to apply formula', 'error');
            return;
        }
        
        const cell = window.selectedCells[0];
        const editableDiv = cell.querySelector('.editable-cell');
        if (!editableDiv) return;
        
        editableDiv.textContent = formula;
        editableDiv.classList.add('has-formula');
        
        evaluateFormula(formula, cell);
    }
    
    /**
     * Evaluate formula and display result
     */
    function evaluateFormula(formula, cell) {
        if (!formula.trim().startsWith('=')) {
            // Not a formula
            return;
        }
        
        try {
            // Store original formula as data attribute
            cell.setAttribute('data-formula', formula);
            
            const editableDiv = cell.querySelector('.editable-cell');
            if (!editableDiv) return;
            
            // Get formula without the equals sign
            const formulaExpression = formula.substring(1).trim();
            
            // Handle SUM function
            if (formulaExpression.toUpperCase().startsWith('SUM(') && formulaExpression.endsWith(')')) {
                const rangeStr = formulaExpression.substring(4, formulaExpression.length - 1);
                const result = evaluateSumFunction(rangeStr);
                
                // Display result but keep original formula
                editableDiv.textContent = result;
                editableDiv.classList.add('formula-result');
                
                // Add tooltip to show formula
                editableDiv.setAttribute('title', formula);
            }
            // Handle AVERAGE function
            else if (formulaExpression.toUpperCase().startsWith('AVERAGE(') && formulaExpression.endsWith(')')) {
                const rangeStr = formulaExpression.substring(8, formulaExpression.length - 1);
                const result = evaluateAverageFunction(rangeStr);
                
                editableDiv.textContent = result;
                editableDiv.classList.add('formula-result');
                editableDiv.setAttribute('title', formula);
            }
            // Handle simple calculations (addition, subtraction, multiplication, division)
            else {
                const cellReferencesRegex = /([A-Z]+[0-9]+)/g;
                // Replace cell references with their values
                let calculableExpression = formulaExpression.replace(cellReferencesRegex, function(match) {
                    return getCellValue(match);
                });
                
                // Safely evaluate the expression
                const result = safeEval(calculableExpression);
                
                // Display result
                editableDiv.textContent = result;
                editableDiv.classList.add('formula-result');
                editableDiv.setAttribute('title', formula);
            }
        } catch (e) {
            console.error('Formula evaluation error:', e);
            
            const editableDiv = cell.querySelector('.editable-cell');
            editableDiv.classList.add('formula-error');
            editableDiv.setAttribute('title', `Error: ${e.message}`);
        }
    }
    
    /**
     * Evaluate SUM function
     * Example: SUM(A1:A5) or SUM(A1,B1,C1)
     */
    function evaluateSumFunction(rangeStr) {
        let sum = 0;
        
        // Check if it's a range (contains :)
        if (rangeStr.includes(':')) {
            const [startCell, endCell] = rangeStr.split(':').map(cell => cell.trim());
            const cellRange = getCellsInRange(startCell, endCell);
            
            // Sum all values in range
            cellRange.forEach(cellRef => {
                sum += parseFloat(getCellValue(cellRef)) || 0;
            });
        }
        // Comma-separated list of cells
        else {
            const cells = rangeStr.split(',').map(cell => cell.trim());
            cells.forEach(cellRef => {
                sum += parseFloat(getCellValue(cellRef)) || 0;
            });
        }
        
        return sum;
    }
    
    /**
     * Evaluate AVERAGE function
     * Example: AVERAGE(A1:A5) or AVERAGE(A1,B1,C1)
     */
    function evaluateAverageFunction(rangeStr) {
        let sum = 0;
        let count = 0;
        
        // Check if it's a range (contains :)
        if (rangeStr.includes(':')) {
            const [startCell, endCell] = rangeStr.split(':').map(cell => cell.trim());
            const cellRange = getCellsInRange(startCell, endCell);
            
            cellRange.forEach(cellRef => {
                const value = parseFloat(getCellValue(cellRef));
                if (!isNaN(value)) {
                    sum += value;
                    count++;
                }
            });
        }
        // Comma-separated list of cells
        else {
            const cells = rangeStr.split(',').map(cell => cell.trim());
            cells.forEach(cellRef => {
                const value = parseFloat(getCellValue(cellRef));
                if (!isNaN(value)) {
                    sum += value;
                    count++;
                }
            });
        }
        
        return count > 0 ? (sum / count).toFixed(2) : 0;
    }
    
    /**
     * Get all cells in a range
     * Example: from A1 to B3 returns [A1, A2, A3, B1, B2, B3]
     */
    function getCellsInRange(startCell, endCell) {
        const startCol = startCell.match(/[A-Z]+/)[0];
        const startRow = parseInt(startCell.match(/[0-9]+/)[0]);
        
        const endCol = endCell.match(/[A-Z]+/)[0];
        const endRow = parseInt(endCell.match(/[0-9]+/)[0]);
        
        // Convert column letters to indices
        const startColIndex = columnLetterToIndex(startCol);
        const endColIndex = columnLetterToIndex(endCol);
        
        // Generate cell references in the range
        const cellRefs = [];
        
        for (let col = startColIndex; col <= endColIndex; col++) {
            const colLetter = indexToColumnLetter(col);
            
            for (let row = startRow; row <= endRow; row++) {
                cellRefs.push(`${colLetter}${row}`);
            }
        }
        
        return cellRefs;
    }
    
    /**
     * Get the value of a cell by its reference (e.g., A1, B2)
     */
    function getCellValue(cellRef) {
        // Parse cell reference
        const colLetter = cellRef.match(/[A-Z]+/)[0];
        const rowNum = parseInt(cellRef.match(/[0-9]+/)[0]);
        
        // Convert to 0-based indices
        const colIndex = columnLetterToIndex(colLetter);
        const rowIndex = rowNum - 1;
        
        // Find the cell in the DOM
        const cell = document.querySelector(`.sheet-cell[data-row="${rowIndex}"][data-col="${colIndex}"]`);
        
        if (!cell) {
            throw new Error(`Cell ${cellRef} not found`);
        }
        
        const editableDiv = cell.querySelector('.editable-cell');
        if (!editableDiv) {
            throw new Error(`Editable div in cell ${cellRef} not found`);
        }
        
        // Get content
        const content = editableDiv.textContent.trim();
        
        // If content is a number, return it
        if (!isNaN(parseFloat(content))) {
            return parseFloat(content);
        }
        
        // If content is empty, return 0
        if (!content) {
            return 0;
        }
        
        // Otherwise return the content
        return content;
    }
    
    /**
     * Convert column letter to index (A=0, B=1, ..., Z=25, AA=26, etc.)
     */
    function columnLetterToIndex(column) {
        let result = 0;
        for (let i = 0; i < column.length; i++) {
            const char = column.charCodeAt(i) - 65; // 'A' is 65 in ASCII
            result = result * 26 + char;
        }
        return result;
    }
    
    /**
     * Convert index to column letter
     */
    function indexToColumnLetter(index) {
        let letter = '';
        index++;
        
        while (index > 0) {
            const modulo = (index - 1) % 26;
            letter = String.fromCharCode(65 + modulo) + letter;
            index = Math.floor((index - modulo) / 26);
        }
        
        return letter;
    }
    
    /**
     * Safely evaluate a mathematical expression
     */
    function safeEval(expression) {
        // Remove any potentially harmful code
        expression = expression.replace(/[^-()\d/*+.]/g, '');
        
        // Use Function instead of eval for better safety
        try {
            return Function('"use strict"; return (' + expression + ')')();
        } catch (e) {
            throw new Error('Invalid expression: ' + expression);
        }
    }
    
    // Re-evaluate all formulas on load
    function evaluateAllFormulas() {
        const cellsWithFormulas = document.querySelectorAll('.sheet-cell[data-formula], .sheet-header[data-formula]');
        cellsWithFormulas.forEach(cell => {
            const formula = cell.getAttribute('data-formula');
            if (formula && formula.trim().startsWith('=')) {
                evaluateFormula(formula, cell);
            }
        });
    }
    
    // Run initial formula evaluation
    setTimeout(evaluateAllFormulas, 500);
});
