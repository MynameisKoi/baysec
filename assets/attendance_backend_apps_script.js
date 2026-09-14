/**
 * ============================================================================
 * BAYSEC ATTENDANCE & LEAGUE SCORING ENGINE (Google Apps Script)
 * ============================================================================
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your BaySec Google Sheet.
 * 2. Go to: Extensions > Apps Script.
 * 3. Delete any existing code and paste this entire script.
 * 4. Run `setupSheets()` once to initialize tabs.
 * 5. Run `migrateVerifiedSolvers()` to reset production Leaderboard and seed welcome cipher solvers.
 * 6. Click "Deploy" > "Manage deployments" > Edit (pencil) > New version > Deploy.
 */

// Custom toolbar menu in Google Sheets
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🛡️ BaySec Admin")
    .addItem("Reset & Seed Production Leaderboard", "migrateVerifiedSolvers")
    .addItem("Setup Simulation / Staging Tables", "setupSimulationTables")
    .addItem("Initialize All Sheets", "setupSheets")
    .addToUi();
}

/**
 * Detects whether Simulation/Staging Mode is active in Admin_Config.
 * Checks for key-value row: Setting_Key: "Sim_Mode", Setting_Value: TRUE/FALSE.
 */
function isSimModeActive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Admin_Config");
  if (!configSheet) return false;

  const data = configSheet.getDataRange().getValues();
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellVal = String(data[r][c]).trim().toLowerCase();
      if (cellVal === "sim_mode" || cellVal === "simmode" || cellVal === "simulation_mode" || cellVal === "staging_mode") {
        const nextVal = (c + 1 < data[r].length) ? data[r][c + 1] : "";
        if (nextVal === true || String(nextVal).trim().toUpperCase() === "TRUE") {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Returns active table names based on Sim_Mode setting.
 */
function getTableNames() {
  const isSim = isSimModeActive();
  return {
    isSim: isSim,
    leaderboard: isSim ? "Leaderboard_Sim" : "Leaderboard",
    attendance: isSim ? "Attendance_Sim" : "Attendance_Log",
    submissions: isSim ? "Submissions_Sim" : "Submissions_Log"
  };
}

// Handle POST requests from web clients
function doPost(e) {
  try {
    const rawData = e.postData.contents;
    const data = JSON.parse(rawData);
    
    if (data.action === "checkin") {
      const result = processCheckin(data);
      return createJsonResponse(result);
    }

    if (data.action === "submit_flag" || data.action === "submit_challenge") {
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
    const tables = getTableNames();
    const isSim = tables.isSim;

    if (e && e.parameter && e.parameter.action) {
      if (e.parameter.action === "checkin") {
        return createJsonResponse(processCheckin(e.parameter));
      }
      if (e.parameter.action === "submit_flag" || e.parameter.action === "submit_challenge") {
        return createJsonResponse(processFlagSubmission(e.parameter));
      }
      if (e.parameter.action === "get_leaderboard") {
        return createJsonResponse({
          status: "online",
          simMode: isSim,
          table: tables.leaderboard,
          leaderboard: getLeaderboardData()
        });
      }
      if (e.parameter.action === "get_active_session" || e.parameter.action === "status") {
        const session = getActiveSessionConfig();
        return createJsonResponse({
          status: "online",
          simMode: isSim,
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
      simMode: isSim,
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
 * Returns formatted leaderboard array for client API consumption.
 */
function getLeaderboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = getTableNames();
  const lbSheet = ss.getSheetByName(tables.leaderboard) || ss.getSheetByName("Sheet1");
  if (!lbSheet) return [];

  const data = lbSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const handle = String(data[i][0] || "").trim();
    if (!handle) continue;
    const email = String(data[i][1] || "").trim();
    const attendance = parseInt(data[i][2]) || 0;
    const challenges = parseInt(data[i][3]) || 0;
    const bonus = parseInt(data[i][4]) || 0;
    const points = parseInt(data[i][5]) || (attendance * 50 + challenges * 100 + bonus);
    const tier = String(data[i][6] || computeTier(points)).trim();
    const aliasSet = (data[i][7] === true || String(data[i][7]).trim().toUpperCase() === "TRUE");

    results.push({
      handle: handle,
      email: email,
      attendance: attendance,
      challenges: challenges,
      bonus: bonus,
      points: points,
      tier: tier,
      aliasSet: aliasSet
    });
  }

  // Sort descending by points
  results.sort(function(a, b) { return b.points - a.points; });
  return results;
}

/**
 * Process and authenticate an incoming student check-in.
 * Dynamically resolves active session and handles sandbox routing + alias fallback.
 */
function processCheckin(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const session = getActiveSessionConfig();
  const tables = getTableNames();
  
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
  let logSheet = ss.getSheetByName(tables.attendance);
  if (!logSheet) {
    logSheet = ss.insertSheet(tables.attendance);
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
  
  // 5. UPDATE MASTER LEADERBOARD SHEET (with Alias Fallback Resolution)
  const lbSheet = ss.getSheetByName(tables.leaderboard) || ss.getSheetByName("Sheet1");
  const lbData = lbSheet.getDataRange().getValues();
  let studentFound = false;
  let totalPoints = 50;
  
  for (let i = 1; i < lbData.length; i++) {
    const lbHandle = String(lbData[i][0]).trim();
    const lbEmail = String(lbData[i][1]).trim().toLowerCase();
    
    if (lbHandle.toLowerCase() === handle.toLowerCase() || (email && lbEmail === email)) {
      studentFound = true;
      
      // Alias Fallback Resolution: If student was migrated without a handle, overwrite with submitted handle
      const aliasSetVal = lbData[i][7];
      const isAliasSet = (aliasSetVal === true || String(aliasSetVal).trim().toUpperCase() === "TRUE");
      if (!isAliasSet) {
        lbSheet.getRange(i + 1, 1).setValue(handle);
        lbSheet.getRange(i + 1, 8).setValue(true);
      }
      if (email && !lbEmail) {
        lbSheet.getRange(i + 1, 2).setValue(email);
      }

      // Increment Attendance_Count (Column C, index 2)
      let currentAttendance = parseInt(lbData[i][2]) || 0;
      let newAttendance = currentAttendance + 1;
      lbSheet.getRange(i + 1, 3).setValue(newAttendance);
      
      // Calculate Total Points: (Attendance * 50) + (Challenges * 100) + Bonus (preserves welcome bonus!)
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
    lbSheet.appendRow([handle, email, 1, 0, 0, 50, tier, true]);
  }
  
  return {
    success: true,
    message: "Attendance confirmed for Week " + activeWeek + "." + (tables.isSim ? " [SIMULATION MODE]" : ""),
    activeWeek: activeWeek,
    totalPoints: totalPoints,
    simMode: tables.isSim
  };
}

/**
 * Process and credit a challenge flag, mini-challenge, or easter egg submission.
 * Supports Challenge 0x01 (+25 Points), Easter Eggs (+50 Points), and Main Flags (+100 Points).
 */
function processFlagSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = getTableNames();
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const week = parseInt(data.week) || 4;
  const flag = String(data.flag || "").trim();
  const isEasterEgg = Boolean(data.isEasterEgg === true || data.isEasterEgg === "true" || String(data.isEasterEgg).toLowerCase() === "true");
  const category = String(data.category || data.type || (isEasterEgg ? "Easter_Egg" : "Challenge_Flag")).trim();

  // Determine points to award:
  let pointsAwarded = parseInt(data.points || data.pointsAwarded) || 0;
  if (!pointsAwarded) {
    if (isEasterEgg) {
      pointsAwarded = 50;
    } else if (category.includes("0x01") || category.toLowerCase().includes("dom leak")) {
      pointsAwarded = 25;
    } else {
      pointsAwarded = 100;
    }
  }

  if (!handle || !flag) {
    return { success: false, message: "Missing handle or token." };
  }

  // 1. Setup / Check Submissions tab
  let subSheet = ss.getSheetByName(tables.submissions);
  if (!subSheet) {
    subSheet = ss.insertSheet(tables.submissions);
    subSheet.getRange("A1:G1").setValues([["Timestamp", "Week", "Handle", "Email", "Category", "Points_Awarded", "Flag_Submitted"]]);
    subSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  const subData = subSheet.getDataRange().getValues();
  const headers = subData[0].map(h => String(h).toLowerCase().trim());
  let colTime = 0, colWeek = 1, colHandle = 2, colEmail = -1, colCategory = 3, colPoints = -1, colFlag = 4;

  for (let c = 0; c < headers.length; c++) {
    const h = headers[c];
    if (h.includes("time")) colTime = c;
    else if (h.includes("week")) colWeek = c;
    else if (h.includes("handle") || h.includes("alias")) colHandle = c;
    else if (h.includes("email")) colEmail = c;
    else if (h.includes("category") || h.includes("type")) colCategory = c;
    else if (h.includes("point")) colPoints = c;
    else if (h.includes("flag") || h.includes("token")) colFlag = c;
  }

  // 2. Anti-Duplicate Verification: Prevent duplicate point farming
  for (let i = 1; i < subData.length; i++) {
    const loggedHandle = String(subData[i][colHandle]).trim().toLowerCase();
    const loggedEmail = (colEmail !== -1) ? String(subData[i][colEmail] || "").trim().toLowerCase() : "";
    const loggedCat = (colCategory !== -1) ? String(subData[i][colCategory] || "").trim().toLowerCase() : "";
    const loggedFlag = (colFlag !== -1) ? String(subData[i][colFlag] || "").trim().toLowerCase() : "";

    const isMatchUser = (loggedHandle === handle.toLowerCase()) || (email && loggedEmail && loggedEmail === email);
    
    // Check if this submission is for the same challenge/category
    const isSameCategory = (loggedCat === category.toLowerCase()) ||
      (category.toLowerCase().includes("0x01") && (loggedCat.includes("0x01") || loggedCat.includes("dom leak") || loggedFlag === flag.toLowerCase())) ||
      (category.toLowerCase().includes("cipher") && loggedCat.includes("cipher"));

    if (isMatchUser && isSameCategory) {
      return {
        success: false,
        alreadyClaimed: true,
        message: "FLAG VERIFIED: You have already claimed points for this challenge."
      };
    }
  }

  // 3. Update Master Leaderboard Sheet (with Alias Fallback Resolution)
  const lbSheet = ss.getSheetByName(tables.leaderboard) || ss.getSheetByName("Sheet1");
  const lbData = lbSheet.getDataRange().getValues();
  let studentFound = false;
  let totalPoints = pointsAwarded;

  for (let i = 1; i < lbData.length; i++) {
    const lbHandle = String(lbData[i][0]).trim();
    const lbEmail = String(lbData[i][1]).trim().toLowerCase();

    if (lbHandle.toLowerCase() === handle.toLowerCase() || (email && lbEmail === email)) {
      studentFound = true;
      
      // Alias Fallback Resolution: If student was migrated without a handle, overwrite with submitted handle
      const aliasSetVal = lbData[i][7];
      const isAliasSet = (aliasSetVal === true || String(aliasSetVal).trim().toUpperCase() === "TRUE");
      if (!isAliasSet) {
        lbSheet.getRange(i + 1, 1).setValue(handle);
        lbSheet.getRange(i + 1, 8).setValue(true);
      }
      if (email && !lbEmail) {
        lbSheet.getRange(i + 1, 2).setValue(email);
      }

      const attendance = parseInt(lbData[i][2]) || 0;
      let challenges = parseInt(lbData[i][3]) || 0;
      let bonus = parseInt(lbData[i][4]) || 0;

      if (pointsAwarded === 100 && !isEasterEgg) {
        challenges += 1;
        lbSheet.getRange(i + 1, 4).setValue(challenges);
      } else {
        bonus += pointsAwarded;
        lbSheet.getRange(i + 1, 5).setValue(bonus);
      }

      totalPoints = (attendance * 50) + (challenges * 100) + bonus;
      lbSheet.getRange(i + 1, 6).setValue(totalPoints);
      lbSheet.getRange(i + 1, 7).setValue(computeTier(totalPoints));
      break;
    }
  }

  if (!studentFound) {
    let challenges = (pointsAwarded === 100 && !isEasterEgg) ? 1 : 0;
    let bonus = (pointsAwarded === 100 && !isEasterEgg) ? 0 : pointsAwarded;
    const tier = computeTier(pointsAwarded);
    lbSheet.appendRow([handle, email, 0, challenges, bonus, pointsAwarded, tier, true]);
  }

  // 4. Record submission in Submissions tab
  const timestamp = data.timestamp ? new Date(data.timestamp) : new Date();
  if (colEmail !== -1) {
    subSheet.appendRow([timestamp, week, handle, email, category, pointsAwarded, flag]);
  } else {
    subSheet.appendRow([timestamp, week, handle, category, flag]);
  }

  return {
    success: true,
    message: (category.toLowerCase().includes("0x01") || category.toLowerCase().includes("dom leak"))
      ? "ACCESS GRANTED. +25 Points added to your profile!"
      : (isEasterEgg ? "Easter egg confirmed (+50 Bonus Points)!" : "Challenge flag confirmed (+100 Points)!"),
    pointsAwarded: pointsAwarded,
    totalPoints: totalPoints,
    simMode: tables.isSim
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
 * PRODUCTION LEADERBOARD RESET & SOLVERS MIGRATION:
 * 1. Resets the live production "Leaderboard" tab (clearing legacy mock data).
 * 2. Scans 'Form Responses 1' (Membership Directory) for poster cipher solvers ("BaySec is here").
 * 3. Guaranteed verified seed solvers:
 *    - David Le (gle48724@student.sfbu.edu)
 *    - Rahim Ajmal Ikhlas (rikhlas4480@student.sfbu.edu)
 *    - Damir Mertl (dmertl25494@student.sfbu.edu)
 * 4. Seeds each solver with +25 bonus points, fallback alias from email prefix, and Alias_Set: FALSE.
 * 5. Seeds Submissions_Log with welcome bonus records to prevent duplicate claims.
 * 6. Preserves existing mock/test leaderboard inside Leaderboard_Sim for testing.
 */
function migrateVerifiedSolvers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Ensure staging/simulation tables preserve mock scores
  setupSimulationTables();

  // 2. Reset production Leaderboard tab
  let lbSheet = ss.getSheetByName("Leaderboard");
  if (!lbSheet) {
    lbSheet = ss.insertSheet("Leaderboard");
  } else {
    lbSheet.clear();
  }

  // Headers with Alias_Set tracking
  const headers = [["Handle", "Email", "Attendance_Count", "Challenges_Solved", "Bonus_Points", "Total_Points", "Tier", "Alias_Set"]];
  lbSheet.getRange("A1:H1").setValues(headers);
  lbSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#38bdf8");

  // 3. Verified Solvers List
  const verifiedSolvers = [
    { name: "David Le", email: "gle48724@student.sfbu.edu" },
    { name: "Rahim Ajmal Ikhlas", email: "rikhlas4480@student.sfbu.edu" },
    { name: "Damir Mertl", email: "dmertl25494@student.sfbu.edu" }
  ];

  // Dynamically inspect 'Form Responses 1' if present
  const formSheet = ss.getSheetByName("Form Responses 1") || ss.getSheetByName("Form Responses");
  if (formSheet) {
    const formData = formSheet.getDataRange().getValues();
    if (formData.length > 1) {
      let colEmail = -1, colCipher = -1, colName = -1;
      const formHeaders = formData[0];
      for (let c = 0; c < formHeaders.length; c++) {
        const h = String(formHeaders[c]).trim().toLowerCase();
        if (h.includes("email")) colEmail = c;
        else if (h.includes("cipher") || h.includes("poster") || h.includes("bonus")) colCipher = c;
        else if (h.includes("name") || h.includes("handle")) colName = c;
      }

      if (colEmail !== -1 && colCipher !== -1) {
        for (let r = 1; r < formData.length; r++) {
          const cipherAnswer = String(formData[r][colCipher] || "").trim().toLowerCase();
          const studentEmail = String(formData[r][colEmail] || "").trim().toLowerCase();
          const studentName = (colName !== -1) ? String(formData[r][colName] || "").trim() : "";

          // Solved phrase: "BaySec is here"
          if (cipherAnswer.includes("baysec is here") && studentEmail) {
            const alreadyInList = verifiedSolvers.some(s => s.email.toLowerCase() === studentEmail);
            if (!alreadyInList) {
              verifiedSolvers.push({ name: studentName, email: studentEmail });
            }
          }
        }
      }
    }
  }

  // 4. Ensure Submissions_Log exists
  let subSheet = ss.getSheetByName("Submissions_Log");
  if (!subSheet) {
    subSheet = ss.insertSheet("Submissions_Log");
    subSheet.getRange("A1:G1").setValues([["Timestamp", "Week", "Handle", "Email", "Category", "Points_Awarded", "Flag_Submitted"]]);
    subSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  const subData = subSheet.getDataRange().getValues();
  const timestamp = new Date();

  // 5. Populate Leaderboard with verified solvers
  verifiedSolvers.forEach(solver => {
    const emailPrefix = solver.email.split("@")[0].trim().toLowerCase();
    const fallbackHandle = emailPrefix;
    const tier = computeTier(25); // "Initiate"

    // [Handle, Email, Attendance_Count, Challenges_Solved, Bonus_Points, Total_Points, Tier, Alias_Set]
    lbSheet.appendRow([fallbackHandle, solver.email.toLowerCase(), 0, 0, 25, 25, tier, false]);

    // Record in Submissions_Log to prevent duplicate claiming
    let alreadyLogged = false;
    for (let i = 1; i < subData.length; i++) {
      const loggedEmail = String(subData[i][3] || "").trim().toLowerCase();
      const loggedCat = String(subData[i][4] || "").trim().toLowerCase();
      if (loggedEmail === solver.email.toLowerCase() && (loggedCat.includes("cipher") || loggedCat.includes("welcome"))) {
        alreadyLogged = true;
        break;
      }
    }

    if (!alreadyLogged) {
      subSheet.appendRow([timestamp, 4, fallbackHandle, solver.email.toLowerCase(), "Welcome Bonus - Poster Cipher", 25, "BaySec is here"]);
    }
  });

  const msg = "Production Leaderboard successfully reset and seeded with verified welcome cipher solvers!\n\n" +
    "- David Le (gle48724@student.sfbu.edu) -> 25 pts (Alias_Set: FALSE)\n" +
    "- Rahim Ajmal Ikhlas (rikhlas4480@student.sfbu.edu) -> 25 pts (Alias_Set: FALSE)\n" +
    "- Damir Mertl (dmertl25494@student.sfbu.edu) -> 25 pts (Alias_Set: FALSE)\n\n" +
    "Mock scores preserved in Leaderboard_Sim.";

  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(msg);
  }
}

/**
 * Sets up Leaderboard_Sim, Attendance_Sim, and Submissions_Sim preserving mock scores.
 */
function setupSimulationTables() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let simLbSheet = ss.getSheetByName("Leaderboard_Sim");
  if (!simLbSheet) {
    simLbSheet = ss.insertSheet("Leaderboard_Sim");
    simLbSheet.getRange("A1:H1").setValues([["Handle", "Email", "Attendance_Count", "Challenges_Solved", "Bonus_Points", "Total_Points", "Tier", "Alias_Set"]]);
    simLbSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#334155").setFontColor("#38bdf8");

    // Legacy mock standings for sandbox testing
    const mockMembers = [
      ["0xViper", "viper@student.sfbu.edu", 1, 1, 0, 150, "Operator", true],
      ["CyberKoi", "cyberkoi@student.sfbu.edu", 1, 1, 0, 150, "Operator", true],
      ["BitPhantom", "phantom@student.sfbu.edu", 1, 1, 0, 150, "Operator", true],
      ["NetStalker", "stalker@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["ZeroTrace", "zerotrace@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["MatrixRebel", "rebel@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["NullPointer", "null@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["EchoBreaker", "echo@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["ByteHawk", "hawk@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true],
      ["RootAdmin", "root@student.sfbu.edu", 1, 0, 0, 50, "Initiate", true]
    ];
    simLbSheet.getRange(2, 1, mockMembers.length, 8).setValues(mockMembers);
  }

  let simAttSheet = ss.getSheetByName("Attendance_Sim");
  if (!simAttSheet) {
    simAttSheet = ss.insertSheet("Attendance_Sim");
    simAttSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Submitted_Passcode"]]);
    simAttSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#334155").setFontColor("#4ade80");
  }

  let simSubSheet = ss.getSheetByName("Submissions_Sim");
  if (!simSubSheet) {
    simSubSheet = ss.insertSheet("Submissions_Sim");
    simSubSheet.getRange("A1:G1").setValues([["Timestamp", "Week", "Handle", "Email", "Category", "Points_Awarded", "Flag_Submitted"]]);
    simSubSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#334155").setFontColor("#f59e0b");
  }
}

/**
 * ONE-CLICK SETUP HELPER:
 * Initializes all required tabs including Admin_Config with Sim_Mode key.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Admin_Config tab
  let configSheet = ss.getSheetByName("Admin_Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Admin_Config");
    configSheet.getRange("A1:E1").setValues([["Active_Week", "Current_Passcode", "Is_Checkin_Open", "Setting_Key", "Setting_Value"]]);
    configSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#38bdf8");
    configSheet.getRange("A2:E2").setValues([[5, "WIRESHARK", true, "Sim_Mode", false]]);
    configSheet.getRange("A3:E3").setValues([[4, "WELCOME", false, "", ""]]);
  } else {
    // Ensure Sim_Mode setting exists in configSheet
    const data = configSheet.getDataRange().getValues();
    let simModeFound = false;
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        if (String(data[r][c]).trim().toLowerCase() === "sim_mode") {
          simModeFound = true;
          break;
        }
      }
    }
    if (!simModeFound) {
      const lastRow = configSheet.getLastRow();
      configSheet.getRange(lastRow + 1, 1, 1, 5).setValues([["", "", "", "Sim_Mode", false]]);
    }
  }
  
  // 2. Setup Attendance_Log tab
  let logSheet = ss.getSheetByName("Attendance_Log");
  if (!logSheet) {
    logSheet = ss.insertSheet("Attendance_Log");
    logSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Submitted_Passcode"]]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#4ade80");
  }

  // 3. Setup Submissions_Log tab
  let subSheet = ss.getSheetByName("Submissions_Log");
  if (!subSheet) {
    subSheet = ss.insertSheet("Submissions_Log");
    subSheet.getRange("A1:G1").setValues([["Timestamp", "Week", "Handle", "Email", "Category", "Points_Awarded", "Flag_Submitted"]]);
    subSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  // 4. Setup Simulation/Staging tabs
  setupSimulationTables();
  
  try {
    SpreadsheetApp.getUi().alert("Setup complete! Admin_Config, Attendance_Log, Submissions_Log, and Simulation tabs are ready.");
  } catch(e) {
    Logger.log("Setup complete!");
  }
}
