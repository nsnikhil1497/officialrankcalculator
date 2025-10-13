/**
 * This Google Apps Script (GAS) file provides the backend for Rank Predictor Interface.html.
 * Paste this code into the Apps Script project linked to your Google Sheets.
 *
 * Important:
 * 1. Permissions: On first run, you need to grant permissions for Gmail and Spreadsheet access.
 * 2. Logic: You must modify the data processing and rank calculation logic in 'checkAndDisplayRank'
 * and 'verifyOTPAndSubmit' functions as per your requirements.
 */

// --- FIX: Add doGet() function to load the web app URL ---
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Interface')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('4th Grade Rank Calculator');
}
// ----------------------------------------------------------------

// Use PropertiesService to store OTP and submission timestamp.
const OTP_STORAGE = PropertiesService.getScriptProperties();
const SUBMISSION_STORAGE = PropertiesService.getScriptProperties(); // NEW: Separate storage for submission timestamp

// NEW: Array of fallback email accounts (your provided emails)
const FALLBACK_EMAILS = ['nsnikhil1497gmail.com', 'rankcalculator2025@gmail.com', 'nikhilizzm96@gmail.com']; // Replace with your actual emails

/**
 * NEW: Helper function to send email using a specific sender (fallback logic).
 * Tries to send from each email in FALLBACK_EMAILS until success.
 * @param {string} to - Recipient email.
 * @param {string} subject - Email subject.
 * @param {string} htmlBody - HTML body.
 * @returns {boolean} True if sent successfully from any account.
 */
function sendEmailWithFallback(to, subject, htmlBody) {
  for (let senderEmail of FALLBACK_EMAILS) {
    try {
      // Note: MailApp uses the script owner's account (primary@gmail.com).
      // For true multi-account sending, use Gmail API with OAuth tokens (advanced).
      // This is a simplified fallback; upgrade to API for high volume.
      MailApp.sendEmail({
        to: to,
        subject: subject,
        htmlBody: htmlBody
      });
      
      Logger.log(`OTP sent successfully from ${senderEmail} to ${to}`);
      return true; // Success, break loop
    } catch (error) {
      Logger.log(`Failed to send from ${senderEmail} to ${to}: ${error.toString()}. Trying next...`);
      // If this is the last one, it will throw at the end.
    }
  }
  // If all fail
  throw new Error('All fallback email accounts failed to send OTP. Please check account limits.');
}

/**
 * Receives form data, generates an OTP, emails it,
 * and stores the OTP and submission timestamp.
 * @param {Object} formData - Data received from the HTML form.
 * @returns {Object} Success/failure message.
 */
function generateAndSendOTP(formData) {
    try {
        const email = formData.email;
        if (!email) {
            return { success: false, message: "Email address is missing. Please enter a valid email." };
        }

        // Check if email already has a pending OTP in PropertiesService
        const storedOTPString = OTP_STORAGE.getProperty(email);
        if (storedOTPString) {
            const storedSubmissionString = SUBMISSION_STORAGE.getProperty(`submission_${email}`);
            if (storedSubmissionString) {
                const stored = JSON.parse(storedSubmissionString);
                const timeDifference = new Date().getTime() - stored.timestamp;
                const FIVE_MINUTES = 5 * 60 * 1000;

                if (timeDifference > FIVE_MINUTES) {
                    // OTP expired, clear both OTP and submission data
                    OTP_STORAGE.deleteProperty(email);
                    SUBMISSION_STORAGE.deleteProperty(`submission_${email}`);
                    Logger.log(`Cleared expired OTP and submission for ${email}`);
                } else {
                    return { success: false, message: "A submission is already in progress for this email. Please complete OTP verification or try again after 5 minutes." };
                }
            } else {
                // No submission timestamp, clear stale OTP
                OTP_STORAGE.deleteProperty(email);
                Logger.log(`Cleared stale OTP for ${email} due to missing submission timestamp`);
            }
        }

        // Check if email already exists in the sheet
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        // Index 4 is the 'Email' column
        const emailExists = values.some(row => row[4] === email);

        if (emailExists) {
            return { success: false, message: "This email has already been used for a submission. Each email can only submit data once." };
        }

        // Generate 6-digit OTP
        const otp = Math.floor(100000 + Math.random() * 900000).toString();
        const submissionTimestamp = new Date().getTime();

        // Store OTP and submission timestamp
        OTP_STORAGE.setProperty(email, JSON.stringify({ otp: otp, data: formData }));
        SUBMISSION_STORAGE.setProperty(`submission_${email}`, JSON.stringify({ timestamp: submissionTimestamp }));

        // NEW: Log remaining email quota
        Logger.log("Remaining email quota: " + MailApp.getRemainingDailyQuota());

        // NEW: Use fallback email sending
        const subject = "Your Rank Predictor OTP";
        const htmlBody = `
            <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ccc; border-radius: 8px;">
                <h2 style="color: #1F3A65;">Rank Submission OTP</h2>
                <p>Below is your one-time password (OTP) to verify your Rank Predictor data submission:</p>
                <p style="font-size: 24px; font-weight: bold; color: #DAA520; background-color: #f0f4f8; padding: 10px; border-radius: 4px; display: inline-block;">
                    ${otp}
                </p>
                <p>This OTP is valid for the next 5 minutes. Please enter it in the app.</p>
                <p style="font-size: 12px; color: #777;">If you did not request this, please ignore this email.</p>
            </div>
        `;

        sendEmailWithFallback(email, subject, htmlBody);

        Logger.log("OTP Sent to " + email);
        return { success: true, message: "OTP sent successfully! Please check your email to finalize submission.", email: email };

    } catch (e) {
        Logger.log("Error in generateAndSendOTP: " + e.toString());
        return { success: false, message: "Failed to send OTP. Error: " + e.toString() };
    }
}

/**
 * Verifies the OTP provided by the user and, if successful, saves the data
 * to Google Sheet (or another DB).
 * @param {string} email - User's email.
 * @param {string} otp - OTP entered by the user.
 * @returns {Object} Success/failure message.
 */
function verifyOTPAndSubmit(email, otp) {
    const storedDataString = OTP_STORAGE.getProperty(email);

    if (!storedDataString) {
        return { success: false, message: "Verification failed. OTP not found or expired. Please resubmit the form." };
    }

    const stored = JSON.parse(storedDataString);
    const timeDifference = new Date().getTime() - JSON.parse(SUBMISSION_STORAGE.getProperty(`submission_${email}`)).timestamp;
    const FIVE_MINUTES = 5 * 60 * 1000;

    // Check for 5-minute expiry
    if (timeDifference > FIVE_MINUTES) {
        OTP_STORAGE.deleteProperty(email);
        return { success: false, message: "Verification failed. The OTP has expired. Please try submitting again." };
    }

    // Check OTP match
    if (stored.otp !== otp) {
        return { success: false, message: "Invalid OTP. Please check the code and try again." };
    }

    try {
        // --- 1. OTP Storage Clear ---
        OTP_STORAGE.deleteProperty(email); // Remove OTP after verification, keep submission timestamp in SUBMISSION_STORAGE
        
        // --- 2. Data Persistence ---
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0]; 
        
        // Ensure header row is complete (12 columns)
        if (sheet.getLastRow() === 0) {
            sheet.appendRow(["Timestamp", "Name", "Category", "Shift", "Email", "Attempted Question", "Correct question", "Wrong Question", "Raw score", "Overall Rank", "Shift Rank", "Category Rank"]);
        }
        
        // Raw Score Calculation (+1.666 for correct, -0.555 for wrong)
        const correctMarks = parseFloat(stored.data.correctScore) * 1.666;
        const negativeMarks = parseFloat(stored.data.wrongQuestion) * 0.555;
        const rawScore = Math.round((correctMarks - negativeMarks) * 100) / 100; 
        
        // Append 12 columns, keeping Rank columns (10, 11, 12) empty
        sheet.appendRow([
            new Date(),                         // 1. Timestamp
            stored.data.name,                   // 2. Name
            stored.data.category,               // 3. Category
            stored.data.shift,                  // 4. Shift
            stored.data.email,                  // 5. Email
            stored.data.attmptedQuestion,       // 6. Attempted Question
            stored.data.correctScore,           // 7. Correct question
            stored.data.wrongQuestion,          // 8. Wrong Question
            rawScore,                           // 9. Raw score
            "",                                 // 10. Overall Rank (Initial Empty Value)
            "",                                 // 11. Shift Rank (Initial Empty Value)
            ""                                  // 12. Category Rank (Initial Empty Value)
        ]);
        
        // --- 3. Return Success ---
        return { success: true, message: "✅ Data submitted and verified successfully! You can check your rank after 5 minutes." };

    } catch (e) {
        Logger.log("Error in verifyOTPAndSubmit: " + e.toString());
        return { success: false, message: "Verification successful, but failed to save data. Error: " + e.toString() };
    }
}

/**
 * Checks if the user's submitted data is still within the 5-minute wait period.
 * @param {string} email - User's email to check submission time.
 * @returns {Object} Status, remaining time, and rank check permission.
 */
function checkOTPExpiry(email) {
    try {
        const storedSubmissionString = SUBMISSION_STORAGE.getProperty(`submission_${email}`);
        if (!storedSubmissionString) {
            return { 
                success: true, 
                isExpired: true, 
                canCheckRank: true, 
                message: "No pending submission found. You can check your rank and score." 
            };
        }

        const stored = JSON.parse(storedSubmissionString);
        const timeDifference = new Date().getTime() - stored.timestamp;
        const FIVE_MINUTES = 5 * 60 * 1000;

        if (timeDifference > FIVE_MINUTES) {
            SUBMISSION_STORAGE.deleteProperty(`submission_${email}`);
            return { 
                success: true, 
                isExpired: true, 
                canCheckRank: true, 
                message: "The 5-minute processing period has ended. You can check your rank and score." 
            };
        }

        const remainingTime = Math.floor((FIVE_MINUTES - timeDifference) / 1000);
        return { 
            success: true, 
            isExpired: false, 
            canCheckRank: false, 
            remainingTime: remainingTime, 
            message: `Please wait ${remainingTime} seconds for your data to be fully processed before checking your rank or score.` 
        };
    } catch (e) {
        Logger.log("Error in checkOTPExpiry: " + e.toString());
        return { 
            success: false, 
            canCheckRank: false, 
            message: "Error checking submission status: " + e.toString() 
        };
    }
}

/**
 * Calculates and displays the rank for the given name and email.
 * Writes the calculated ranks back to the sheet.
 * @param {Object} formData - User's name and email.
 * @returns {Object} Rank details or failure message.
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

        // Check submission timestamp before proceeding
        const expiryCheck = checkOTPExpiry(formData.email);
        if (!expiryCheck.canCheckRank) {
            return {
                success: false,
                message: expiryCheck.message,
                remainingTime: expiryCheck.remainingTime
            };
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
        const emailExists = values.some(row => row[4] === formData.email);
        if (!emailExists) {
            return { success: false, message: "No submission found for this email. Please submit your data first." };
        }

        // Get user's submitted row data based on email and name
        let userRow = null;
        let userRowIndexInSheet = -1;
        for (let i = 1; i < values.length; i++) {
            const email = values[i][4] ? String(values[i][4]).trim() : '';
            const name = values[i][1] ? String(values[i][1]).trim() : '';
            if (email === formData.email && name === formData.name) {
                userRow = values[i];
                userRowIndexInSheet = i;
                break;
            }
        }

        if (!userRow) {
            return { success: false, message: "Email found, but the name does not match. Please check the name and try again." };
        }
        
        const userRawScore = userRow[8]; // Raw score is column 9 (index 8)
        const userShift = userRow[3] ? String(userRow[3]).trim() : '';
        const userCategory = userRow[2] ? String(userRow[2]).trim() : '';
        
        // Filter valid rows (skipping header row)
        const rankedData = values.slice(1).filter(row => {
            const score = row[8];
            return (score != null && score !== "" && !isNaN(Number(score)));
        });

        // Rank Calculation
        rankedData.sort((a, b) => Number(b[8]) - Number(a[8])); 

        let overallRank = 0;
        let shiftRank = 0;
        let categoryRank = 0;
        let totalShiftCandidates = 0;
        let totalCategoryCandidates = 0;
        
        // Calculate Overall Rank
        for (let i = 0; i < rankedData.length; i++) {
            if (String(rankedData[i][4]).trim() === formData.email) {
                overallRank = i + 1;
                break;
            }
        }
        
        // Calculate Shift Rank
        const shiftCandidates = rankedData.filter(row => (String(row[3]).trim() === userShift));
        totalShiftCandidates = shiftCandidates.length;
        for (let i = 0; i < shiftCandidates.length; i++) {
            if (String(shiftCandidates[i][4]).trim() === formData.email) {
                shiftRank = i + 1;
                break;
            }
        }

        // Calculate Category Rank
        const categoryCandidates = rankedData.filter(row => (String(row[2]).trim() === userCategory));
        totalCategoryCandidates = categoryCandidates.length;
        for (let i = 0; i < categoryCandidates.length; i++) {
            if (String(categoryCandidates[i][4]).trim() === formData.email) {
                categoryRank = i + 1;
                break;
            }
        }

        // Write calculated ranks back to the Google Sheet
        if (userRowIndexInSheet !== -1) {
            try {
                const targetRow = userRowIndexInSheet + 1;
                sheet.getRange(targetRow, 10, 1, 3).setValues([[overallRank, shiftRank, categoryRank]]);
                Logger.log(`Successfully updated ranks for ${formData.email} in row ${targetRow}.`);
                SUBMISSION_STORAGE.deleteProperty(`submission_${formData.email}`);
            } catch (e) {
                Logger.log(`Failed to update ranks for ${formData.email}: ${e.toString()}`);
                return { 
                    success: true, 
                    message: "Ranks calculated but failed to update sheet. Contact admin. Error: " + e.toString(),
                    name: userRow[1],
                    overallRank: overallRank,
                    totalCandidates: rankedData.length,
                    rawScore: userRawScore,
                    shiftRank: shiftRank,
                    totalShiftCandidates: totalShiftCandidates,
                    categoryRank: categoryRank,
                    totalCategoryCandidates: totalCategoryCandidates,
                    shift: userShift,
                    category: userCategory
                };
            }
        }

        return { 
            success: true, 
            name: userRow[1],
            overallRank: overallRank,
            totalCandidates: rankedData.length,
            rawScore: userRawScore,
            shiftRank: shiftRank,
            totalShiftCandidates: totalShiftCandidates, 
            categoryRank: categoryRank,
            totalCategoryCandidates: totalCategoryCandidates,
            shift: userShift,
            category: userCategory
        };

    } catch (e) {
        Logger.log("Error in checkAndDisplayRank: " + e.toString());
        return { success: false, message: "An unexpected error occurred during rank check: " + e.toString() };
    }
}

function checkQuota() {
  Logger.log("Remaining email quota: " + MailApp.getRemainingDailyQuota());
  Logger.log("Active user: " + Session.getActiveUser().getEmail());
}
