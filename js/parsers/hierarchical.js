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
            // Usually looks like: ["Course Name", null, "202610"]
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
            if (currentCourse && row.length >= 3 && typeof row[0] === 'string' && row[1] === "Crs Mean") {
                const rawQuestion = row[0].replace(/\n/g, ' ').trim();
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

                // Look ahead for comparative means (Dept, School, Univ)
                // These rows have the first cell empty (null/undefined)
                let j = i + 1;
                while (j < rawData.length && j < i + 10) { // Check up to 10 rows just in case
                    const nextRow = rawData[j];
                    if (!nextRow || nextRow.length < 2) {
                        j++;
                        continue;
                    }

                    // If we find a new question or a new course, stop looking
                    if (nextRow[0] && typeof nextRow[0] === 'string' && nextRow[0].trim() !== "") {
                        break; 
                    }

                    const label = nextRow[1] ? nextRow[1].toString().trim() : "";
                    const value = parseFloat(nextRow[2]);

                    if (label === "Dept Mean") {
                        if (!isNaN(value)) questionObj.deptMean = value.toFixed(2);
                    } else if (label === "School Mean") {
                        if (!isNaN(value)) questionObj.schoolMean = value.toFixed(2);
                    } else if (label === "Univ Mean") {
                        if (!isNaN(value)) questionObj.univMean = value.toFixed(2);
                    } else if (label === "Total Enrollments" || label === "Response Rate") {
                        // These are course-level, stop mean lookup
                        break;
                    } else if (label === "Crs Mean") {
                        // Found next question, stop
                        break;
                    }
                    
                    j++;
                }

                currentCourse.questions.push(questionObj);
                // We DON'T advance 'i' here because course-level metrics might be after the means 
                // and we want the main loop to see them. 
                // However, we want to skip the mean rows we just processed.
                // But wait, if we don't advance 'i', the main loop will hit Dept Mean rows.
                // We need to advance 'i' to the last consumed mean row.
                i = j - 1;
                continue;
            }

            // Capture response rate and enrollments (usually at the end of the course block)
            const label = row[1] ? row[1].toString().trim() : "";
            if (label === "Total Enrollments") {
                currentCourse.enrollments = row[2];
            } else if (label === "Response Rate") {
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
