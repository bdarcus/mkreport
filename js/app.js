// Main App Logic (Non-Module Version)

document.addEventListener('DOMContentLoaded', () => {
    const fileInput = document.getElementById('excel-upload');
    const parserSelect = document.getElementById('parser-select');
    const compactToggle = document.getElementById('compact-toggle');
    const groupSemesterToggle = document.getElementById('group-semester-toggle');
    const generateBtn = document.getElementById('generate-btn');
    const copyBtn = document.getElementById('copy-btn');
    const reportOutput = document.getElementById('report-output');
    const errorMessage = document.getElementById('error-message');

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
            const buffer = await readFileAsArrayBuffer(file);
            const parserKey = parserSelect.value;
            const parseFn = window.Parsers[parserKey];
            
            if (!parseFn) {
                throw new Error(`Parser '${parserKey}' is not implemented.`);
            }

            const rawData = parseFn(buffer);

            if (!rawData.evaluations || rawData.evaluations.length === 0) {
                throw new Error("No evaluation data found. Check your report format.");
            }

            // Transform Data based on options
            const data = transformData(rawData.evaluations);

            // Select Template
            const templateId = compactToggle.checked ? 'compact-report-template' : 'annual-report-template';
            const templateElement = document.getElementById(templateId);
            if (!templateElement) {
                throw new Error(`Template '${templateId}' not found.`);
            }
            const templateSource = templateElement.innerHTML;

            // Render Markdown using Mustache
            const markdownText = Mustache.render(templateSource, data);

            // Convert Markdown to HTML using Marked
            const htmlText = marked.parse(markdownText);

            // Display
            reportOutput.innerHTML = htmlText;
            copyBtn.classList.remove('hidden');

        } catch (error) {
            console.error("Error generating report:", error);
            showError(`Error: ${error.message}`);
        } finally {
            generateBtn.disabled = false;
        }
    }

    function transformData(evaluations) {
        // Prepare extra fields for compact mode
        evaluations.forEach(eval => {
            const key1 = eval.questions.find(q => q.questionText.includes("Upon reflection"));
            const key2 = eval.questions.find(q => q.questionText.includes("On a scale of 1-10"));
            eval.keyMean1 = key1 ? key1.mean : "-";
            eval.keyMean2 = key2 ? key2.mean : "-";
        });

        const isGrouped = groupSemesterToggle.checked;
        
        if (isGrouped) {
            const semestersMap = new Map();
            evaluations.forEach(eval => {
                if (!semestersMap.has(eval.term)) {
                    semestersMap.set(eval.term, []);
                }
                semestersMap.get(eval.term).push(eval);
            });

            const semesters = Array.from(semestersMap.entries()).map(([term, evals]) => ({
                term,
                evaluations: evals
            }));
            
            semesters.sort((a, b) => b.term.localeCompare(a.term));
            
            return { 
                semesters,
                isGrouped: true
            };
        } else {
            return {
                semesters: [{
                    term: "All Courses",
                    evaluations: evaluations
                }],
                isGrouped: false
            };
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

    generateBtn.addEventListener('click', handleGenerate);
    copyBtn.addEventListener('click', handleCopy);
    fileInput.addEventListener('change', clearError);
});
