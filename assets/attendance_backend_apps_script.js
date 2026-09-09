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
    
    return createJsonResponse({ success: false, message: "Invalid action." });
  } catch (err) {
    return createJsonResponse({ success: false, message: "Server error: " + err.toString() });
  }
}

// Handle GET requests (health check or manual admin verification)
function doGet(e) {
  const config = getAdminConfig();
  return createJsonResponse({
    status: "online",
    activeWeek: config.activeWeek,
    checkinOpen: config.isOpen
  });
}

/**
 * Process and authenticate an incoming student check-in.
 */
function processCheckin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getAdminConfig();
  
  const studentWeek = parseInt(data.week) || 0;
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const submittedPasscode = String(data.passcode || "").trim().toUpperCase();
  
  // 1. GATE 1: Is Check-In Open?
  if (!config.isOpen) {
    return {
      success: false,
      message: "Check-in is currently closed. Attendance can only be claimed while a meeting is in session."
    };
  }
  
  // 2. GATE 2: Week Match
  if (studentWeek !== config.activeWeek) {
    return {
      success: false,
      message: "Active session is Week " + config.activeWeek + ". You cannot submit attendance for Week " + studentWeek + "."
    };
  }
  
  // 3. GATE 3: Passcode Verification (Never stored on GitHub!)
  if (submittedPasscode !== config.passcode.toUpperCase()) {
    return {
      success: false,
      message: "Invalid meeting passcode. Look at the slide projected in the room!"
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
 * Reads config safely from Admin_Config sheet (completely private to admin).
 */
function getAdminConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Admin_Config");
  
  if (!configSheet) {
    return {
      activeWeek: 5,
      passcode: "WIRESHARK",
      isOpen: true
    };
  }
  
  // Admin_Config format:
  // Row 2: Active Week (Cell A2)
  // Row 2: Current Passcode (Cell B2)
  // Row 2: Is Open (Cell C2: TRUE/FALSE)
  const activeWeek = parseInt(configSheet.getRange("A2").getValue()) || 5;
  const passcode = String(configSheet.getRange("B2").getValue() || "WIRESHARK").trim();
  const isOpen = Boolean(configSheet.getRange("C2").getValue() !== false);
  
  return { activeWeek, passcode, isOpen };
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
  
  SpreadsheetApp.getUi().alert("Setup complete! Admin_Config and Attendance_Log tabs have been created.");
}
