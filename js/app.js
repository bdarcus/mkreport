// Main App Logic

import { parseExcel as defaultParser } from './parsers/default.js';

const PARSERS = {
    'default': defaultParser
};

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-upload');
    const parserSelect = document.getElementById('parser-select');
    const reportOutput = document.getElementById('report-output');
    const errorMessage = document.getElementById('error-message');
    const templateSource = document.getElementById('annual-report-template').innerHTML;

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove('hidden');
        reportOutput.innerHTML = '<p class="placeholder">Upload an Excel file to generate your report.</p>';
    }

    function clearError() {
        errorMessage.textContent = '';
        errorMessage.classList.add('hidden');
    }

    async function handleFileUpload(event) {
        const file = event.target.files[0];
        if (!file) return;

        clearError();
        reportOutput.innerHTML = '<p class="placeholder">Processing your report...</p>';

        try {
            // Read file as ArrayBuffer
            const buffer = await readFileAsArrayBuffer(file);
            
            // Get selected parser
            const parserKey = parserSelect.value;
            const parseFn = PARSERS[parserKey];
            
            if (!parseFn) {
                throw new Error(`Parser '${parserKey}' is not implemented.`);
            }

            // Parse data
            const data = parseFn(buffer);

            // Render Markdown using Mustache
            const markdownText = Mustache.render(templateSource, data);

            // Convert Markdown to HTML using Marked
            const htmlText = marked.parse(markdownText);

            // Display
            reportOutput.innerHTML = htmlText;

        } catch (error) {
            console.error("Error generating report:", error);
            showError(`Error processing report: ${error.message}`);
        }
    }

    function readFileAsArrayBuffer(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => resolve(e.target.result);
            reader.onerror = (e) => reject(new Error("Failed to read file"));
            reader.readAsArrayBuffer(file);
        });
    }

    // Attach event listeners
    fileInput.addEventListener('change', handleFileUpload);
});
