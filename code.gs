/**
 * This Google Apps Script (GAS) file provides the backend for Rank Predictor Interface.html.
 * Paste this code into the Apps Script project linked to your Google Sheets.
 *
 * Important:
 * 1. Permissions: On first run, you need to grant permissions for Spreadsheet access.
 * 2. Logic: Modify the data processing and rank calculation logic in 'submitFormData'
 * and 'checkAndDisplayRank' functions as per your requirements.
 * 3. Device-based restriction: Each device can submit data only once, tracked via a device ID.
 * 4. Rank checking: Uses name and email for unique identification.
 * 5. Sheet structure: Headers: Timestamp, Device ID, Name, Category, Shift, Email, Attempted Question, Correct Question, Wrong Question, Raw Score, Overall Rank, Shift Rank, Category Rank
 * 6. Tie-breaking: Candidates with equal Raw Scores get the same rank, and the next rank is skipped.
 * 7. Tied count: Shows how many candidates share the same score for overall, shift, and category.
 * 8. Validation: Attempted Questions <= Correct Questions + Wrong Questions.
 */

// --- Loads the web app URL ---
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interface')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('4th Grade Rank Calculator');
}

/**
 * Receives form data and saves it directly to Google Sheet after checking device ID.
 * @param {Object} formData - Data received from the HTML form, including deviceId and email.
 * @returns {Object} Success/failure message.
 */
function submitFormData(formData) {
    try {
        const deviceId = formData.deviceId;
        const email = formData.email;
        if (!deviceId) {
            return { success: false, message: "Device ID is missing. Please enable JavaScript and try again." };
        }
        if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
            return { success: false, message: "Please enter a valid email address." };
        }

        // Validate Attempted Questions <= Correct Questions + Wrong Questions
        const attempted = parseInt(formData.attmptedQuestion, 10);
        const correct = parseInt(formData.correctScore, 10);
        const wrong = parseInt(formData.wrongQuestion, 10);
        if (attempted > (correct + wrong)) {
            return { success: false, message: "Attempted Questions cannot be greater than Correct Questions + Wrong Questions." };
        }
        if (attempted > 120) {
            return { success: false, message: "Attempted Questions cannot exceed 120." };
        }

        // Check if device ID already exists in the sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        // Index 1 is the 'Device ID' column
        const deviceIdExists = values.some(row => row[1] === deviceId);

        if (deviceIdExists) {
            return { success: false, message: "This device has already been used for a submission. Each device can only submit data once." };
        }

        // Ensure header row is complete (13 columns)
        if (sheet.getLastRow() === 0) {
            sheet.appendRow(["Timestamp", "Device ID", "Name", "Category", "Shift", "Email", "Attempted Question", "Correct Question", "Wrong Question", "Raw Score", "Overall Rank", "Shift Rank", "Category Rank"]);
        }

        // Raw Score Calculation (+1.666 for correct, -0.555 for wrong)
        const correctMarks = parseFloat(formData.correctScore) * 1.666;
        const negativeMarks = parseFloat(formData.wrongQuestion) * 0.555;
        const rawScore = Math.round((correctMarks - negativeMarks) * 100) / 100;

        // Append 13 columns, keeping Rank columns (11, 12, 13) empty
        sheet.appendRow([
            new Date(),                         // 1. Timestamp
            formData.deviceId,                  // 2. Device ID
            formData.name,                      // 3. Name
            formData.category,                  // 4. Category
            formData.shift,                     // 5. Shift
            formData.email,                     // 6. Email
            formData.attmptedQuestion,          // 7. Attempted Question
            formData.correctScore,              // 8. Correct Question
            formData.wrongQuestion,             // 9. Wrong Question
            rawScore,                           // 10. Raw Score
            "",                                 // 11. Overall Rank (Initial Empty Value)
            "",                                 // 12. Shift Rank (Initial Empty Value)
            ""                                  // 13. Category Rank (Initial Empty Value)
        ]);

        Logger.log(`Data submitted successfully for device ${deviceId} and email ${email}`);
        return { success: true, message: "✅ Data submitted successfully! You can check your rank immediately." };

    } catch (e) {
        Logger.log("Error in submitFormData: " + e.toString());
        return { success: false, message: "Failed to submit data. Error: " + e.toString() };
    }
}

/**
 * Calculates and displays the rank for the given name and email.
 * Writes the calculated ranks back to the sheet.
 * Includes count of candidates tied at the same score.
 * @param {Object} formData - User's name and email.
 * @returns {Object} Rank details, tied counts, or failure message.
 */
function checkAndDisplayRank(formData) {
    try {
        // Input validation
        if (!formData || !formData.email || !formData.name) {
            Logger.log("Invalid input: formData is missing email or name");
            return { success: false, message: "Please provide both email and name to check rank." };
        }
        formData.email = String(formData.email || '').trim();
        formData.name = String(formData.name || '').trim();
        if (!formData.email || !formData.name) {
            Logger.log("Invalid input: email or name is empty after trimming");
            return { success: false, message: "Email or name cannot be empty. Please check your input." };
        }

        // Fetch ALL data
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0]; 
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        // Check if sheet has data
        if (values.length <= 1) {
            Logger.log("No data found in sheet for ranking");
            return { success: false, message: "No data available in the sheet to calculate ranks. Please submit data first." };
        }

        // Pre-check if email exists in the sheet
        const emailExists = values.some(row => row[5] === formData.email);
        if (!emailExists) {
            return { success: false, message: "No submission found for this email. Please submit your data first." };
        }

        // Get user's submitted row data based on email and name
        let userRow = null;
        let userRowIndexInSheet = -1;
        for (let i = 1; i < values.length; i++) {
            const email = values[i][5] ? String(values[i][5]).trim() : '';
            const name = values[i][2] ? String(values[i][2]).trim() : '';
            if (email === formData.email && name === formData.name) {
                userRow = values[i];
                userRowIndexInSheet = i;
                break;
            }
        }

        if (!userRow) {
            return { success: false, message: "Email found, but the name does not match. Please check the name and try again." };
        }
        
        const userRawScore = userRow[9]; // Raw score is column 10 (index 9)
        const userShift = userRow[4] ? String(userRow[4]).trim() : '';
        const userCategory = userRow[3] ? String(userRow[3]).trim() : '';
        
        // Filter valid rows (skipping header row)
        const rankedData = values.slice(1).filter(row => {
            const score = row[9];
            return (score != null && score !== "" && !isNaN(Number(score)));
        });

        // Sort data by Raw Score (descending)
        rankedData.sort((a, b) => Number(b[9]) - Number(a[9])); 

        let overallRank = 0;
        let shiftRank = 0;
        let categoryRank = 0;
        let overallTiedCount = 0;
        let shiftTiedCount = 0;
        let categoryTiedCount = 0;
        let totalShiftCandidates = 0;
        let totalCategoryCandidates = 0;

        // Calculate Overall Rank and Tied Count
        let currentRank = 0;
        let lastScore = null;
        for (let i = 0; i < rankedData.length; i++) {
            const currentScore = Number(rankedData[i][9]);
            if (currentScore !== lastScore) {
                currentRank = i + 1; // New rank for new score
                lastScore = currentScore;
            }
            if (Number(rankedData[i][9]) === Number(userRawScore)) {
                overallTiedCount++;
            }
            if (String(rankedData[i][5]).trim() === formData.email && String(rankedData[i][2]).trim() === formData.name) {
                overallRank = currentRank;
            }
        }
        
        // Calculate Shift Rank and Tied Count
        const shiftCandidates = rankedData.filter(row => (String(row[4]).trim() === userShift));
        totalShiftCandidates = shiftCandidates.length;
        currentRank = 0;
        lastScore = null;
        for (let i = 0; i < shiftCandidates.length; i++) {
            const currentScore = Number(shiftCandidates[i][9]);
            if (currentScore !== lastScore) {
                currentRank = i + 1;
                lastScore = currentScore;
            }
            if (Number(shiftCandidates[i][9]) === Number(userRawScore)) {
                shiftTiedCount++;
            }
            if (String(shiftCandidates[i][5]).trim() === formData.email && String(shiftCandidates[i][2]).trim() === formData.name) {
                shiftRank = currentRank;
            }
        }

        // Calculate Category Rank and Tied Count
        const categoryCandidates = rankedData.filter(row => (String(row[3]).trim() === userCategory));
        totalCategoryCandidates = categoryCandidates.length;
        currentRank = 0;
        lastScore = null;
        for (let i = 0; i < categoryCandidates.length; i++) {
            const currentScore = Number(categoryCandidates[i][9]);
            if (currentScore !== lastScore) {
                currentRank = i + 1;
                lastScore = currentScore;
            }
            if (Number(categoryCandidates[i][9]) === Number(userRawScore)) {
                categoryTiedCount++;
            }
            if (String(categoryCandidates[i][5]).trim() === formData.email && String(categoryCandidates[i][2]).trim() === formData.name) {
                categoryRank = currentRank;
            }
        }

        // Write calculated ranks back to the Google Sheet
        if (userRowIndexInSheet !== -1) {
            try {
                const targetRow = userRowIndexInSheet + 1;
                sheet.getRange(targetRow, 11, 1, 3).setValues([[overallRank, shiftRank, categoryRank]]);
                Logger.log(`Successfully updated ranks for email ${formData.email} in row ${targetRow}.`);
            } catch (e) {
                Logger.log(`Failed to update ranks for email ${formData.email}: ${e.toString()}`);
                return { 
                    success: true, 
                    message: "Ranks calculated but failed to update sheet. Contact admin. Error: " + e.toString(),
                    name: userRow[2],
                    overallRank: overallRank,
                    totalCandidates: rankedData.length,
                    rawScore: userRawScore,
                    shiftRank: shiftRank,
                    totalShiftCandidates: totalShiftCandidates,
                    categoryRank: categoryRank,
                    totalCategoryCandidates: totalCategoryCandidates,
                    overallTiedCount: overallTiedCount,
                    shiftTiedCount: shiftTiedCount,
                    categoryTiedCount: categoryTiedCount,
                    shift: userShift,
                    category: userCategory
                };
            }
        }

        return { 
            success: true, 
            name: userRow[2],
            overallRank: overallRank,
            totalCandidates: rankedData.length,
            rawScore: userRawScore,
            shiftRank: shiftRank,
            totalShiftCandidates: totalShiftCandidates, 
            categoryRank: categoryRank,
            totalCategoryCandidates: totalCategoryCandidates,
            overallTiedCount: overallTiedCount,
            shiftTiedCount: shiftTiedCount,
            categoryTiedCount: categoryTiedCount,
            shift: userShift,
            category: userCategory
        };

    } catch (e) {
        Logger.log("Error in checkAndDisplayRank: " + e.toString());
        return { success: false, message: "An unexpected error occurred during rank check: " + e.toString() };
    }
}
