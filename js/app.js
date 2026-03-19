// Main App Logic (Non-Module Version)

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-upload');
    const parserSelect = document.getElementById('parser-select');
    const generateBtn = document.getElementById('generate-btn');
    const copyBtn = document.getElementById('copy-btn');
    const reportOutput = document.getElementById('report-output');
    const errorMessage = document.getElementById('error-message');
    const templateSource = document.getElementById('annual-report-template').innerHTML;

    function showError(message) {
        errorMessage.textContent = message;
        errorMessage.classList.remove('hidden');
        copyBtn.classList.add('hidden');
        reportOutput.innerHTML = '<p class="placeholder">Upload an Excel file and click "Generate Report" to see the results.</p>';
    }

    function clearError() {
        errorMessage.textContent = '';
        errorMessage.classList.add('hidden');
    }

    async function handleGenerate() {
        const file = fileInput.files[0];
        if (!file) {
            showError("Please select an Excel file first.");
            return;
        }

        clearError();
        reportOutput.innerHTML = '<p class="placeholder">Processing your report...</p>';
        generateBtn.disabled = true;
        copyBtn.classList.add('hidden');

        try {
            // Read file as ArrayBuffer
            const buffer = await readFileAsArrayBuffer(file);
            
            // Get selected parser from global window.Parsers
            const parserKey = parserSelect.value;
            const parseFn = window.Parsers[parserKey];
            
            if (!parseFn) {
                throw new Error(`Parser '${parserKey}' is not implemented.`);
            }

            // Parse data
            const data = parseFn(buffer);

            if (!data.evaluations || data.evaluations.length === 0) {
                throw new Error("No evaluation data found. Please check if you selected the correct 'Report Format'.");
            }

            // Render Markdown using Mustache
            const markdownText = Mustache.render(templateSource, data);

            // Convert Markdown to HTML using Marked
            const htmlText = marked.parse(markdownText);

            // Display
            reportOutput.innerHTML = htmlText;
            copyBtn.classList.remove('hidden');

        } catch (error) {
            console.error("Error generating report:", error);
            showError(`Error processing report: ${error.message}`);
        } finally {
            generateBtn.disabled = false;
        }
    }

    function handleCopy() {
        const html = reportOutput.innerHTML;
        navigator.clipboard.writeText(html).then(() => {
            const originalText = copyBtn.textContent;
            copyBtn.textContent = 'Copied!';
            setTimeout(() => {
                copyBtn.textContent = originalText;
            }, 2000);
        }).catch(err => {
            console.error('Could not copy text: ', err);
            alert('Failed to copy text to clipboard.');
        });
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
    generateBtn.addEventListener('click', handleGenerate);
    copyBtn.addEventListener('click', handleCopy);
    
    fileInput.addEventListener('change', () => {
        clearError();
        copyBtn.classList.add('hidden');
    });
    parserSelect.addEventListener('change', () => {
        clearError();
        copyBtn.classList.add('hidden');
    });
});
