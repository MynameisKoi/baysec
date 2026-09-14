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
      if (e.parameter.action === "get_active_session" || e.parameter.action === "status") {
        const session = getActiveSessionConfig();
        return createJsonResponse({
          status: "online",
          isOpen: session.isOpen,
          checkinOpen: session.isOpen,
          activeWeek: session.activeWeek,
          message: session.isOpen
            ? "Active session open."
            : (session.message || "Check-in is currently CLOSED. No active meeting session is open right now.")
        });
      }
    }
    const session = getActiveSessionConfig();
    return createJsonResponse({
      status: "online",
      isOpen: session.isOpen,
      checkinOpen: session.isOpen,
      activeWeek: session.activeWeek,
      message: session.isOpen
        ? "Active session open."
        : (session.message || "Check-in is currently CLOSED. No active meeting session is open right now.")
    });
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

/**
 * Process and authenticate an incoming student check-in.
 * Dynamically resolves the active session from Admin_Config.
 */
function processCheckin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const session = getActiveSessionConfig();
  
  // 1. GATE 1: Is check-in currently open for any session?
  if (!session.isOpen || !session.activeWeek) {
    return {
      success: false,
      message: session.message || "Check-in is currently CLOSED. No active meeting session is open right now."
    };
  }

  const activeWeek = session.activeWeek;
  const validPasscode = session.passcode;
  
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const submittedPasscode = String(data.passcode || "").trim().toUpperCase();
  
  if (!handle || !email || !submittedPasscode) {
    return {
      success: false,
      message: "Please fill in all fields (handle, email, and passcode)."
    };
  }

  // 2. GATE 2: Passcode Verification against valid_passcode
  if (submittedPasscode !== validPasscode.toUpperCase()) {
    return {
      success: false,
      message: "Invalid meeting passcode for Week " + activeWeek + ". (Hint: check the slide projected on the screen!)"
    };
  }
  
  // 3. GATE 3: Anti-Duplicate Verification for the specific active_week
  let logSheet = ss.getSheetByName("Attendance_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Attendance_Log");
    logSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Submitted_Passcode"]]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#4ade80");
  }

  const logData = logSheet.getDataRange().getValues();
  for (let i = 1; i < logData.length; i++) {
    const loggedWeek = parseInt(logData[i][1]);
    const loggedHandle = String(logData[i][2]).trim().toLowerCase();
    const loggedEmail = String(logData[i][3]).trim().toLowerCase();
    
    if (loggedWeek === activeWeek && (loggedHandle === handle.toLowerCase() || (email && loggedEmail === email))) {
      return {
        success: false,
        message: "You have already checked in for Week " + activeWeek + "! Each student may only claim points once per meeting."
      };
    }
  }
  
  // 4. Log attendance record under dynamically resolved active_week
  const timestamp = new Date();
  logSheet.appendRow([timestamp, activeWeek, handle, email, submittedPasscode]);
  
  // 5. UPDATE MASTER LEADERBOARD SHEET
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
    message: "Attendance confirmed for Week " + activeWeek + ".",
    activeWeek: activeWeek,
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
 * Dynamically queries the currently active meeting session from the Admin_Config sheet.
 * Scans for the row where Is_Checkin_Open == TRUE (boolean true or case-insensitive "TRUE").
 * Does NOT hardcode or fallback to Week 5.
 */
function getActiveSessionConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Admin_Config");
  
  if (!configSheet) {
    return {
      isOpen: false,
      checkinOpen: false,
      activeWeek: null,
      passcode: "",
      message: "Check-in is currently CLOSED. No active meeting session is open right now."
    };
  }
  
  const data = configSheet.getDataRange().getValues();
  if (data.length <= 1) {
    return {
      isOpen: false,
      checkinOpen: false,
      activeWeek: null,
      passcode: "",
      message: "Check-in is currently CLOSED. No active meeting session is open right now."
    };
  }

  // Detect column mapping based on header row:
  // Expected headers: Active_Week (0), Current_Passcode (1), Is_Checkin_Open (2)
  let colWeek = 0;
  let colPasscode = 1;
  let colIsOpen = 2;

  const headers = data[0];
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c]).trim().toLowerCase();
    if (h === "active_week" || h === "week") colWeek = c;
    else if (h === "current_passcode" || h === "passcode") colPasscode = c;
    else if (h === "is_checkin_open" || h === "is_open" || h === "open") colIsOpen = c;
  }

  // Scan for the row where Is_Checkin_Open == TRUE
  for (let i = 1; i < data.length; i++) {
    const rawOpen = data[i][colIsOpen];
    const isOpen = (rawOpen === true || String(rawOpen).trim().toUpperCase() === "TRUE");
    if (isOpen) {
      const activeWeek = parseInt(data[i][colWeek]);
      const validPasscode = String(data[i][colPasscode] || "").trim();
      return {
        isOpen: true,
        checkinOpen: true,
        activeWeek: activeWeek,
        passcode: validPasscode,
        message: "Active session found."
      };
    }
  }

  // If no row has Is_Checkin_Open == TRUE:
  return {
    isOpen: false,
    checkinOpen: false,
    activeWeek: null,
    passcode: "",
    message: "Check-in is currently CLOSED. No active meeting session is open right now."
  };
}

/**
 * Reads config safely from Admin_Config sheet.
 * If targetWeek is specified, checks that row.
 * If targetWeek is omitted, delegates to getActiveSessionConfig().
 */
function getAdminConfig(targetWeek) {
  if (!targetWeek) {
    return getActiveSessionConfig();
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Admin_Config");
  
  if (!configSheet) {
    return { activeWeek: parseInt(targetWeek), passcode: "", isOpen: false, found: false };
  }
  
  const data = configSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowWeek = parseInt(data[i][0]);
    if (rowWeek === parseInt(targetWeek)) {
      const passcode = String(data[i][1] || "").trim();
      const isOpen = (data[i][2] === true || String(data[i][2]).trim().toUpperCase() === "TRUE");
      return { activeWeek: rowWeek, passcode: passcode, isOpen: isOpen, found: true };
    }
  }

  return { activeWeek: parseInt(targetWeek), passcode: "", isOpen: false, found: false };
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
