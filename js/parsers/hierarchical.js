// Hierarchical Excel Parser (Global Version)

window.Parsers = window.Parsers || {};

window.Parsers.hierarchical = function parseExcel(dataBuffer) {
    const HIGHLIGHT_QUESTIONS = [
        "Upon reflection, this instructor is an effective teacher.",
        "On a scale of 1-10, how effective are the teaching methods of this faculty member?"
    ];

    try {
        const workbook = XLSX.read(dataBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        
        if (!rawData || rawData.length === 0) {
            throw new Error("Excel sheet is empty.");
        }

        const evaluations = [];
        let currentCourse = null;

        for (let i = 0; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;

            // Heuristic for Course row: [String, empty/null, String/Number]
            if (row.length >= 3 && typeof row[0] === 'string' && !row[1] && row[2]) {
                currentCourse = {
                    courseName: row[0].trim(),
                    term: row[2].toString(),
                    questions: [],
                    comments: []
                };
                evaluations.push(currentCourse);
                continue;
            }

            // Heuristic for Question row: [String, "Crs Mean", Number]
            if (currentCourse && row.length >= 3 && typeof row[0] === 'string' && row[1] === "Crs Mean") {
                const rawQuestion = row[0].trim();
                const meanScore = parseFloat(row[2]);

                if (isNaN(meanScore)) continue;

                let formattedQuestion = rawQuestion;
                const isHighlight = HIGHLIGHT_QUESTIONS.some(q => rawQuestion.includes(q));
                
                if (isHighlight) {
                    formattedQuestion = `<strong class="highlight-question">${rawQuestion}</strong>`;
                }

                currentCourse.questions.push({
                    questionText: formattedQuestion,
                    mean: meanScore.toFixed(2),
                    responses: "N/A"
                });
            }
        }

        return { evaluations };

    } catch (error) {
        console.error("Error parsing Hierarchical Excel:", error);
        throw error;
    }
};
