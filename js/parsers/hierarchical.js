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
                    comments: [],
                    responseRate: "N/A",
                    enrollments: "N/A"
                };
                evaluations.push(currentCourse);
                continue;
            }

            if (!currentCourse) continue;

            // Heuristic for Question row: [String, "Crs Mean", Number]
            if (row.length >= 3 && typeof row[0] === 'string' && row[1] === "Crs Mean") {
                const rawQuestion = row[0].trim();
                const meanScore = parseFloat(row[2]);

                if (isNaN(meanScore)) continue;

                let formattedQuestion = rawQuestion;
                const isHighlight = HIGHLIGHT_QUESTIONS.some(q => rawQuestion.includes(q));
                
                if (isHighlight) {
                    formattedQuestion = `<strong class="highlight-question">${rawQuestion}</strong>`;
                }

                const questionObj = {
                    questionText: formattedQuestion,
                    mean: meanScore.toFixed(2),
                    deptMean: "-",
                    schoolMean: "-",
                    univMean: "-"
                };

                // Look ahead for comparative means
                let j = i + 1;
                while (j < rawData.length && j < i + 5) {
                    const nextRow = rawData[j];
                    if (!nextRow || nextRow[0] !== null) break;
                    
                    const label = nextRow[1];
                    const value = parseFloat(nextRow[2]);

                    if (label === "Dept Mean") {
                        if (!isNaN(value)) questionObj.deptMean = value.toFixed(2);
                    } else if (label === "School Mean") {
                        if (!isNaN(value)) questionObj.schoolMean = value.toFixed(2);
                    } else if (label === "Univ Mean") {
                        if (!isNaN(value)) questionObj.univMean = value.toFixed(2);
                    } else {
                        break; // Not a mean label, don't consume it
                    }
                    j++;
                }

                currentCourse.questions.push(questionObj);
                i = j - 1; // Advance main loop to last consumed mean row
                continue;
            }

            // Capture response rate and enrollments
            if (row[1] === "Total Enrollments") {
                currentCourse.enrollments = row[2];
            } else if (row[1] === "Response Rate") {
                const rate = parseFloat(row[2]);
                currentCourse.responseRate = isNaN(rate) ? row[2] : (rate * 100).toFixed(0) + "%";
            }
        }

        return { evaluations };

    } catch (error) {
        console.error("Error parsing Hierarchical Excel:", error);
        throw error;
    }
};
