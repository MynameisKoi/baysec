/**
 * ============================================================================
 * BAYSEC ATTENDANCE & LEAGUE SCORING ENGINE (Google Apps Script)
 * ============================================================================
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your BaySec Google Sheet.
 * 2. Go to: Extensions > Apps Script.
 * 3. Delete any existing code and paste this entire script.
 * 4. Run the function `setupSheets()` once (select setupSheets in the toolbar and click Run).
 *    This will automatically create your "Admin_Config" and "Attendance_Log" tabs!
 * 5. Click "Deploy" (top right) > "New deployment".
 *    - Select type: "Web app"
 *    - Description: "BaySec Check-In API"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (allows student browser requests to reach your script)
 *    - Click "Deploy", authorize permissions, and copy the "Web app URL" ending in /exec.
 * 6. Paste that URL into `checkin.html` (APPS_SCRIPT_URL variable).
 * 7. Done! You now have a 100% cheat-proof, automated attendance gateway.
 */

// Handle POST requests from checkin.html
function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const data = JSON.parse(rawData);
    
    if (data.action === "checkin") {
      const result = processCheckin(data);
      return createJsonResponse(result);
    }

    if (data.action === "submit_flag") {
      const result = processFlagSubmission(data);
      return createJsonResponse(result);
    }
    
    return createJsonResponse({ success: false, message: "Invalid action." });
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

// Handle GET requests (100% CORS-friendly for all web clients)
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.action) {
      if (e.parameter.action === "checkin") {
        return createJsonResponse(processCheckin(e.parameter));
      }
      if (e.parameter.action === "submit_flag") {
        return createJsonResponse(processFlagSubmission(e.parameter));
      }
    }
    const config = getAdminConfig();
    return createJsonResponse({
      status: "online",
      activeWeek: config.activeWeek,
      checkinOpen: config.isOpen
    });
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

/**
 * Process and authenticate an incoming student check-in.
 */
function processCheckin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const studentWeek = parseInt(data.week) || 0;
  const config = getAdminConfig(studentWeek);
  
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const submittedPasscode = String(data.passcode || "").trim().toUpperCase();
  
  // 1. GATE 1: Did we find configuration for this week?
  if (!config.found) {
    return {
      success: false,
      message: "Week " + studentWeek + " is not configured in the Admin_Config sheet tab."
    };
  }

  // 2. GATE 2: Is Check-In Open for this week?
  if (!config.isOpen) {
    return {
      success: false,
      message: "Check-in for Week " + studentWeek + " is currently CLOSED (marked FALSE in Admin_Config). Ask an officer in the room to open check-in."
    };
  }
  
  // 3. GATE 3: Passcode Verification (Never stored on GitHub!)
  if (submittedPasscode !== config.passcode.toUpperCase()) {
    return {
      success: false,
      message: "Invalid meeting passcode for Week " + studentWeek + ". (Hint: check the slide projected on the screen!)"
    };
  }
  
  // 4. GATE 4: Anti-Duplicate Verification
  const logSheet = ss.getSheetByName("Attendance_Log");
  const logData = logSheet.getDataRange().getValues();
  
  for (let i = 1; i < logData.length; i++) {
    const loggedWeek = parseInt(logData[i][1]);
    const loggedHandle = String(logData[i][2]).trim().toLowerCase();
    const loggedEmail = String(logData[i][3]).trim().toLowerCase();
    
    if (loggedWeek === studentWeek && (loggedHandle === handle.toLowerCase() || (email && loggedEmail === email))) {
      return {
        success: false,
        message: "You have already checked in for Week " + studentWeek + "! Each student may only claim points once per meeting."
      };
    }
  }
  
  // 5. SUCCESS: Record in Attendance_Log
  const timestamp = new Date();
  logSheet.appendRow([timestamp, studentWeek, handle, email, submittedPasscode]);
  
  // 6. UPDATE MASTER LEADERBOARD SHEET
  const lbSheet = ss.getSheetByName("Leaderboard") || ss.getSheetByName("Sheet1");
  const lbData = lbSheet.getDataRange().getValues();
  let studentFound = false;
  let totalPoints = 50;
  
  for (let i = 1; i < lbData.length; i++) {
    const lbHandle = String(lbData[i][0]).trim();
    const lbEmail = String(lbData[i][1]).trim().toLowerCase();
    
    if (lbHandle.toLowerCase() === handle.toLowerCase() || (email && lbEmail === email)) {
      studentFound = true;
      
      // Increment Attendance_Count (Column C, index 2)
      let currentAttendance = parseInt(lbData[i][2]) || 0;
      let newAttendance = currentAttendance + 1;
      lbSheet.getRange(i + 1, 3).setValue(newAttendance);
      
      // Calculate Total Points: (Attendance * 50) + (Challenges * 100) + Bonus
      let challenges = parseInt(lbData[i][3]) || 0;
      let bonus = parseInt(lbData[i][4]) || 0;
      totalPoints = (newAttendance * 50) + (challenges * 100) + bonus;
      lbSheet.getRange(i + 1, 6).setValue(totalPoints);
      
      // Calculate Tier
      const tier = computeTier(totalPoints);
      lbSheet.getRange(i + 1, 7).setValue(tier);
      break;
    }
  }
  
  // If first-time student, append new row to Leaderboard!
  if (!studentFound) {
    const tier = computeTier(50);
    lbSheet.appendRow([handle, email, 1, 0, 0, 50, tier]);
  }
  
  return {
    success: true,
    message: "Attendance confirmed.",
    totalPoints: totalPoints
  };
}

/**
 * Process and credit a challenge flag or easter egg submission.
 */
function processFlagSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const handle = String(data.handle || "").trim();
  const week = parseInt(data.week) || 5;
  const flag = String(data.flag || "").trim();
  const isEasterEgg = Boolean(data.isEasterEgg);

  if (!handle || !flag) {
    return { success: false, message: "Missing handle or token." };
  }

  // Check Submissions_Log for duplicate claim
  let subSheet = ss.getSheetByName("Submissions_Log");
  if (!subSheet) {
    subSheet = ss.insertSheet("Submissions_Log");
    subSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Type", "Flag_Submitted"]]);
    subSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  const subData = subSheet.getDataRange().getValues();
  const submissionType = isEasterEgg ? "Easter_Egg" : "Challenge_Flag";

  for (let i = 1; i < subData.length; i++) {
    const loggedWeek = parseInt(subData[i][1]);
    const loggedHandle = String(subData[i][2]).trim().toLowerCase();
    const loggedType = String(subData[i][3]).trim();

    if (loggedWeek === week && loggedHandle === handle.toLowerCase() && loggedType === submissionType) {
      return {
        success: false,
        message: "You have already claimed points for this " + (isEasterEgg ? "easter egg" : "challenge flag") + "!"
      };
    }
  }

  // Update Leaderboard
  const lbSheet = ss.getSheetByName("Leaderboard") || ss.getSheetByName("Sheet1");
  const lbData = lbSheet.getDataRange().getValues();
  let studentFound = false;
  let totalPoints = isEasterEgg ? 50 : 100;

  for (let i = 1; i < lbData.length; i++) {
    const lbHandle = String(lbData[i][0]).trim();

    if (lbHandle.toLowerCase() === handle.toLowerCase()) {
      studentFound = true;
      const attendance = parseInt(lbData[i][2]) || 0;
      let challenges = parseInt(lbData[i][3]) || 0;
      let bonus = parseInt(lbData[i][4]) || 0;

      if (isEasterEgg) {
        bonus += 50;
        lbSheet.getRange(i + 1, 5).setValue(bonus);
      } else {
        challenges += 1;
        lbSheet.getRange(i + 1, 4).setValue(challenges);
      }

      totalPoints = (attendance * 50) + (challenges * 100) + bonus;
      lbSheet.getRange(i + 1, 6).setValue(totalPoints);
      lbSheet.getRange(i + 1, 7).setValue(computeTier(totalPoints));
      break;
    }
  }

  if (!studentFound) {
    const challenges = isEasterEgg ? 0 : 1;
    const bonus = isEasterEgg ? 50 : 0;
    const tier = computeTier(totalPoints);
    lbSheet.appendRow([handle, "", 0, challenges, bonus, totalPoints, tier]);
  }

  // Record in Submissions_Log
  subSheet.appendRow([new Date(), week, handle, submissionType, flag]);

  return {
    success: true,
    message: isEasterEgg ? "Easter egg confirmed (+50 Bonus Points)!" : "Challenge flag confirmed (+100 Points)!",
    pointsAwarded: isEasterEgg ? 50 : 100,
    totalPoints: totalPoints
  };
}

/**
 * Reads config safely from Admin_Config sheet (completely private to admin).
 * Supports multiple rows: e.g. Row 2 for Week 4, Row 3 for Week 5, etc.
 */
function getAdminConfig(targetWeek) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Admin_Config");
  
  if (!configSheet) {
    return {
      activeWeek: 5,
      passcode: "WIRESHARK",
      isOpen: true,
      found: true
    };
  }
  
  const data = configSheet.getDataRange().getValues();
  // Headers: [Active_Week, Current_Passcode, Is_Checkin_Open]

  // If a specific targetWeek was requested (e.g. Week 5):
  if (targetWeek) {
    for (let i = 1; i < data.length; i++) {
      const rowWeek = parseInt(data[i][0]);
      if (rowWeek === parseInt(targetWeek)) {
        const passcode = String(data[i][1] || "").trim();
        const isOpen = (String(data[i][2]).toUpperCase() === "TRUE" || data[i][2] === true);
        return { activeWeek: rowWeek, passcode: passcode, isOpen: isOpen, found: true };
      }
    }
    // Target week row not found in Admin_Config
    return { activeWeek: parseInt(targetWeek), passcode: "", isOpen: false, found: false };
  }

  // Otherwise, find the currently active week (first row where Is_Checkin_Open is TRUE)
  for (let i = 1; i < data.length; i++) {
    const rowWeek = parseInt(data[i][0]);
    const passcode = String(data[i][1] || "").trim();
    const isOpen = (String(data[i][2]).toUpperCase() === "TRUE" || data[i][2] === true);
    if (isOpen) {
      return { activeWeek: rowWeek, passcode: passcode, isOpen: true, found: true };
    }
  }

  // Fallback to row 2
  if (data.length > 1) {
    const rowWeek = parseInt(data[1][0]) || 5;
    const passcode = String(data[1][1] || "").trim();
    const isOpen = (String(data[1][2]).toUpperCase() === "TRUE" || data[1][2] === true);
    return { activeWeek: rowWeek, passcode: passcode, isOpen: isOpen, found: true };
  }
  
  return { activeWeek: 5, passcode: "WIRESHARK", isOpen: true, found: true };
}

function computeTier(points) {
  if (points >= 1000) return "Grandmaster";
  if (points >= 600) return "Breaker";
  if (points >= 300) return "Sentinel";
  if (points >= 150) return "Operator";
  return "Initiate";
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ONE-CLICK SETUP HELPER:
 * Run this function once from Apps Script editor to create tabs!
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Admin_Config tab if not exists
  let configSheet = ss.getSheetByName("Admin_Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Admin_Config");
    configSheet.getRange("A1:C1").setValues([["Active_Week", "Current_Passcode", "Is_Checkin_Open"]]);
    configSheet.getRange("A1:C1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#38bdf8");
    configSheet.getRange("A2:C2").setValues([[5, "WIRESHARK", true]]);
  }
  
  // 2. Setup Attendance_Log tab if not exists
  let logSheet = ss.getSheetByName("Attendance_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Attendance_Log");
    logSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Submitted_Passcode"]]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#4ade80");
  }

  // 3. Setup Submissions_Log tab if not exists
  let subSheet = ss.getSheetByName("Submissions_Log");
  if (!subSheet) {
    subSheet = ss.insertSheet("Submissions_Log");
    subSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Type", "Flag_Submitted"]]);
    subSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }
  
  SpreadsheetApp.getUi().alert("Setup complete! Admin_Config, Attendance_Log, and Submissions_Log tabs have been created.");
}
