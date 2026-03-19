// Default Excel Parser (Global Version)

window.Parsers = window.Parsers || {};

window.Parsers.default = function parseExcel(dataBuffer) {
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
            throw new Error("Excel sheet is empty or not formatted correctly.");
        }

        const evaluationsMap = new Map();
        let headerIdx = -1;
        for (let i = 0; i < rawData.length; i++) {
            if (rawData[i] && rawData[i].length >= 5) {
                headerIdx = i;
                break;
            }
        }

        if (headerIdx === -1 || headerIdx === rawData.length - 1) {
            throw new Error("Could not find a valid data header row.");
        }

        for (let i = headerIdx + 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (!row || row.length === 0) continue;

            const courseName = row[0] || "Unknown Course";
            const term = row[1] || "Unknown Term";
            const rawQuestion = row[2] || "";
            const meanScore = parseFloat(row[3]) || 0;
            const responsesCount = parseInt(row[4], 10) || 0;
            const comment = row[5] || null;

            if (!rawQuestion && !comment) continue;

            const evalKey = `${courseName}-${term}`;
            if (!evaluationsMap.has(evalKey)) {
                evaluationsMap.set(evalKey, { 
                    courseName, 
                    term, 
                    questions: [], 
                    comments: [],
                    enrollments: "-",
                    responseRate: responsesCount ? "N/A" : "-"
                });
            }

            const evalData = evaluationsMap.get(evalKey);

            if (rawQuestion) {
                let formattedQuestion = rawQuestion;
                const isHighlight = HIGHLIGHT_QUESTIONS.some(q => rawQuestion.includes(q));
                if (isHighlight) formattedQuestion = `<strong class="highlight-question">${rawQuestion}</strong>`;

                evalData.questions.push({
                    questionText: formattedQuestion,
                    mean: meanScore.toFixed(2),
                    responses: responsesCount,
                    deptMean: "-",
                    schoolMean: "-",
                    univMean: "-"
                });
            }

            if (comment) evalData.comments.push(comment);
        }

        return { evaluations: Array.from(evaluationsMap.values()) };

    } catch (error) {
        console.error("Error parsing Excel:", error);
        throw error;
    }
};
