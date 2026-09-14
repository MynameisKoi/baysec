/**
 * ============================================================================
 * BAYSEC ATTENDANCE & LEAGUE SCORING ENGINE (Google Apps Script)
 * ============================================================================
 * 
 * FEATURES:
 * 1. Dynamic Staging/Simulation Mode Switch (Admin_Config -> Sim_Mode: TRUE/FALSE)
 * 2. Dynamic Tiered Scoring Engine (First Blood Bonus & 24h First Day Decay)
 * 3. Strict Deduplication & Unique Constraint on (user_email, challenge_id)
 * 4. Production Leaderboard Reset & Solvers Migration with Alias Fallback
 * 5. Secret Exposure Prevention & Anti-Cheat: Salted SHA-256 Hashing & Sanitized API Responses
 * 
 * SETUP INSTRUCTIONS:
 * 1. Open your BaySec Google Sheet.
 * 2. Go to: Extensions > Apps Script.
 * 3. Delete any existing code and paste this entire script.
 * 4. Run `setupSheets()` once to initialize tabs.
 * 5. Run `migrateVerifiedSolvers()` to reset production Leaderboard and seed welcome cipher solvers.
 * 6. Click "Deploy" > "Manage deployments" > Edit (pencil) > New version > Deploy.
 */

// Global Salt for SHA-256 Hashing (ensures anti-rainbow table & secret protection)
const DEFAULT_SALT = "baysec_salt_2026_#sfbu";

// Default Challenge Matrix with Salted Hashes, Release Timestamps & Tiered Bonuses
const CHALLENGES_CONFIG = {
  "challenge_0x01": {
    id: "challenge_0x01",
    title: "Challenge 0x01 - DOM Leak",
    week: 4,
    releaseTime: "2026-09-07T18:00:00Z",
    basePoints: 25,
    firstBloodBonus: 15, // 25 + 15 = 40 pts
    firstDayBonus: 5,    // 25 + 5 = 30 pts
    // Salted SHA-256 hash of "sfbu{w3lc0m3_t0_b4ys3c_2026}"
    saltedHashes: ["5ffb4d11576aec23e7d1e7650919e0825b4d608d19c91c0c7b3aba288ba18e0a"]
  },
  "challenge_0x02": {
    id: "challenge_0x02",
    title: "Challenge 0x02 - PCAP Packet Sniffing",
    week: 5,
    releaseTime: "2026-09-14T18:00:00Z",
    basePoints: 100,
    firstBloodBonus: 25, // 100 + 25 = 125 pts
    firstDayBonus: 10,   // 100 + 10 = 110 pts
    // Salted SHA-256 hash of "sfbu{w1r3sh4rk_p4ck3t_sn1ff3r_2026}"
    saltedHashes: ["2e42bce4a71a27918df6822849f42f459572e4b75c48ce4b77e023f0b194432f"]
  },
  "easter_egg_0x01": {
    id: "easter_egg_0x01",
    title: "Easter Egg 0x01 - HTTP Basic Auth Creds",
    week: 5,
    releaseTime: "2026-09-14T18:00:00Z",
    basePoints: 50,
    firstBloodBonus: 0,
    firstDayBonus: 0,
    // Supports any of: base64, plaintext creds, password, or sfbu{password}
    saltedHashes: [
      "236a5df4deaef042d87ba560fd1f715435c09bc9463deb968a410650398cc102", // c2ZidV9hZG1pbjpoNGNrM3JfcDRzc3cwcmQ=
      "f37c7d45de252cede4e085cd94a426724a2381c47ff49a185021d7e001f00b5e", // sfbu_admin:h4ck3r_p4ssw0rd
      "f71995635694550a2778c71a4744fcb0e1e5d209e0d6b77b89a82e3596c0e7e9", // h4ck3r_p4ssw0rd
      "d238ab5b3e0874f7babf36b177f0635edfdf5e4505bf44877b8483b7f43cf92a"  // sfbu{h4ck3r_p4ssw0rd}
    ]
  },
  "welcome_poster_cipher": {
    id: "welcome_poster_cipher",
    title: "Welcome Bonus - Poster Cipher",
    week: 4,
    releaseTime: "2026-09-01T00:00:00Z",
    basePoints: 25,
    firstBloodBonus: 0,
    firstDayBonus: 0,
    // Salted SHA-256 hash of "BaySec is here"
    saltedHashes: ["ed62383b618c53865e6dc0f289fa03ce45a7c1d1157cab51a4f265b788d03ef0"]
  }
};

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
 * Computes a salted SHA-256 hex string for a given secret token.
 * Prevents client-side secret leakage and defends against rainbow table lookups.
 */
function hashSecret(secret, salt) {
  salt = salt || DEFAULT_SALT;
  const normalized = salt + ":" + String(secret || "").trim().toLowerCase();
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized, Utilities.Charset.UTF_8);
  let hex = "";
  for (let i = 0; i < digest.length; i++) {
    let byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    let byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = "0" + byteHex;
    hex += byteHex;
  }
  return hex;
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
  let attSheet = isSim ? "Attendance_Sim" : "Attendance";
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!isSim && ss) {
      if (!ss.getSheetByName("Attendance") && ss.getSheetByName("Attendance_Log")) {
        attSheet = "Attendance_Log";
      }
    }
  } catch (e) {}

  return {
    isSim: isSim,
    leaderboard: isSim ? "Leaderboard_Sim" : "Leaderboard",
    attendance: attSheet,
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
        // SANITIZED RESPONSE: strictly returns active_week, isOpen, and simMode (never leaks passcodes or hashes)
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

    // Default status response (sanitized)
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
 * Includes firstBlood boolean badge tag and aliasSet status.
 */
function getLeaderboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = getTableNames();
  const lbSheet = ss.getSheetByName(tables.leaderboard) || ss.getSheetByName("Sheet1");
  if (!lbSheet) return [];

  const data = lbSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Identify solvers who earned FIRST_BLOOD in Submissions
  const firstBloodSolvers = new Set();
  const subSheet = ss.getSheetByName(tables.submissions);
  if (subSheet) {
    const subData = subSheet.getDataRange().getValues();
    if (subData.length > 1) {
      for (let i = 1; i < subData.length; i++) {
        const solveTier = String(subData[i][5] || "").trim().toUpperCase();
        if (solveTier === "FIRST_BLOOD") {
          const subHandle = String(subData[i][2] || "").trim().toLowerCase();
          const subEmail = String(subData[i][3] || "").trim().toLowerCase();
          if (subHandle) firstBloodSolvers.add(subHandle);
          if (subEmail) firstBloodSolvers.add(subEmail);
        }
      }
    }
  }

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

    const hasFirstBlood = firstBloodSolvers.has(handle.toLowerCase()) || (email && firstBloodSolvers.has(email.toLowerCase()));

    results.push({
      handle: handle,
      email: email,
      attendance: attendance,
      challenges: challenges,
      bonus: bonus,
      points: points,
      tier: tier,
      aliasSet: aliasSet,
      firstBlood: Boolean(hasFirstBlood)
    });
  }

  // Sort descending by points
  results.sort(function(a, b) { return b.points - a.points; });
  return results;
}

/**
 * Process and authenticate an incoming student check-in.
 * Validates against salted hash or plaintext passcode.
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
  const validPasscodeHash = session.passcodeHash;
  
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const submittedPasscode = String(data.passcode || "").trim();
  
  if (!handle || !email || !submittedPasscode) {
    return {
      success: false,
      message: "Please fill in all fields (handle, email, and passcode)."
    };
  }

  // 2. GATE 2: Passcode Verification against salted hash and plaintext
  const submittedHash = hashSecret(submittedPasscode);
  const isMatch = (validPasscodeHash && submittedHash === validPasscodeHash) ||
                  (validPasscode && submittedPasscode.toUpperCase() === validPasscode.toUpperCase()) ||
                  (validPasscode && submittedHash === hashSecret(validPasscode));

  if (!isMatch) {
    return {
      success: false,
      message: "Invalid meeting passcode for Week " + activeWeek + ". (Hint: check the slide projected on the screen!)"
    };
  }
  
  // 3. GATE 3: Anti-Duplicate Verification for the specific active_week
  let logSheet = ss.getSheetByName(tables.attendance);
  if (!logSheet) {
    logSheet = ss.insertSheet(tables.attendance);
    logSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Passcode_Hash"]]);
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
  
  // 4. Log attendance record under dynamically resolved active_week with salted hash
  const timestamp = new Date();
  logSheet.appendRow([timestamp, activeWeek, handle, email, submittedHash]);
  
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
      
      // Alias Fallback Resolution: If student was migrated without a custom handle (Alias_Set: FALSE), overwrite with submitted handle
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
    message: "Attendance confirmed for Week " + activeWeek + " (+50 Points)!" + (tables.isSim ? " [SIMULATION MODE]" : ""),
    activeWeek: activeWeek,
    totalPoints: totalPoints,
    simMode: tables.isSim
  };
}

/**
 * Dynamic Tiered Scoring Engine (First Blood & First Day Decay)
 * Evaluates submitted token against salted SHA-256 hashes.
 * Strictly enforces unique constraint on (user_email, challenge_id).
 */
function processFlagSubmission(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tables = getTableNames();
  const handle = String(data.handle || "").trim();
  const email = String(data.email || "").trim().toLowerCase();
  const submittedToken = String(data.flag || data.token || "").trim();
  const explicitChallengeId = String(data.challenge_id || data.challengeId || "").trim();
  const submittedTimestamp = data.timestamp ? new Date(data.timestamp) : new Date();

  if (!handle || !submittedToken) {
    return { success: false, message: "Missing handle or token." };
  }

  // 1. Match challenge via salted SHA-256 hash
  const submittedHash = hashSecret(submittedToken);
  let matchedChallenge = null;

  if (explicitChallengeId && CHALLENGES_CONFIG[explicitChallengeId]) {
    const ch = CHALLENGES_CONFIG[explicitChallengeId];
    if (ch.saltedHashes.indexOf(submittedHash) !== -1) {
      matchedChallenge = ch;
    }
  }

  if (!matchedChallenge) {
    // Scan all registered challenges & easter eggs
    for (const key in CHALLENGES_CONFIG) {
      const ch = CHALLENGES_CONFIG[key];
      if (ch.saltedHashes.indexOf(submittedHash) !== -1) {
        matchedChallenge = ch;
        break;
      }
    }
  }

  if (!matchedChallenge) {
    return {
      success: false,
      message: "ACCESS DENIED: Invalid flag format or incorrect payload token."
    };
  }

  const challengeId = matchedChallenge.id;
  const challengeTitle = matchedChallenge.title;
  const challengeWeek = matchedChallenge.week;
  const releaseTime = new Date(matchedChallenge.releaseTime);

  // 2. Setup / Check Submissions tab
  let subSheet = ss.getSheetByName(tables.submissions);
  if (!subSheet) {
    subSheet = ss.insertSheet(tables.submissions);
    subSheet.getRange("A1:H1").setValues([["Timestamp", "Week", "Handle", "Email", "Challenge_ID", "Solve_Tier", "Points_Awarded", "Flag_Hash"]]);
    subSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  const subData = subSheet.getDataRange().getValues();

  // 3. Strict Deduplication & Unique Constraint on (user_email, challenge_id) and (handle, challenge_id)
  let priorSolvesCount = 0;
  for (let i = 1; i < subData.length; i++) {
    const loggedHandle = String(subData[i][2] || "").trim().toLowerCase();
    const loggedEmail = String(subData[i][3] || "").trim().toLowerCase();
    const loggedChallenge = String(subData[i][4] || "").trim().toLowerCase();

    // Track total club solves for First Blood evaluation
    if (loggedChallenge === challengeId.toLowerCase()) {
      priorSolvesCount++;
    }

    // Check unique constraint per student
    const isUserMatch = (loggedHandle === handle.toLowerCase()) || (email && loggedEmail && loggedEmail === email);
    if (isUserMatch && (loggedChallenge === challengeId.toLowerCase())) {
      return {
        success: false,
        alreadyClaimed: true,
        message: "FLAG VERIFIED: You have already claimed points for this challenge."
      };
    }
  }

  // 4. Dynamic Tiered Point Calculation
  let solveTier = "STANDARD";
  let pointsAwarded = matchedChallenge.basePoints;

  if (matchedChallenge.firstBloodBonus > 0 && priorSolvesCount === 0) {
    // First Blood: Very first solver across the entire club
    solveTier = "FIRST_BLOOD";
    pointsAwarded = matchedChallenge.basePoints + matchedChallenge.firstBloodBonus;
  } else if (matchedChallenge.firstDayBonus > 0) {
    // First Day Decay: Solved within 24 hours of challenge release
    const diffMs = submittedTimestamp.getTime() - releaseTime.getTime();
    const isWithin24Hours = (diffMs >= 0 && diffMs <= 24 * 60 * 60 * 1000);
    if (isWithin24Hours) {
      solveTier = "FIRST_DAY";
      pointsAwarded = matchedChallenge.basePoints + matchedChallenge.firstDayBonus;
    }
  }

  // 5. Update Master Leaderboard Sheet (with Alias Fallback Resolution)
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

      if (matchedChallenge.basePoints >= 100) {
        challenges += 1;
        const extraBonus = pointsAwarded - matchedChallenge.basePoints;
        if (extraBonus > 0) bonus += extraBonus;
        lbSheet.getRange(i + 1, 4).setValue(challenges);
        lbSheet.getRange(i + 1, 5).setValue(bonus);
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
    let challenges = (matchedChallenge.basePoints >= 100) ? 1 : 0;
    let bonus = (matchedChallenge.basePoints >= 100) ? (pointsAwarded - matchedChallenge.basePoints) : pointsAwarded;
    const tier = computeTier(pointsAwarded);
    lbSheet.appendRow([handle, email, 0, challenges, bonus, pointsAwarded, tier, true]);
  }

  // 6. Record submission in Submissions tab
  subSheet.appendRow([submittedTimestamp, challengeWeek, handle, email, challengeId, solveTier, pointsAwarded, submittedHash]);

  // Construct celebratory response message
  let displayMessage = "";
  if (solveTier === "FIRST_BLOOD") {
    displayMessage = "ACCESS GRANTED. 🩸 FIRST BLOOD! +" + pointsAwarded + " Points added to your profile!";
  } else if (solveTier === "FIRST_DAY") {
    displayMessage = "ACCESS GRANTED. ⚡ Day-Of Solve Bonus! +" + pointsAwarded + " Points added to your profile!";
  } else if (challengeId === "easter_egg_0x01") {
    displayMessage = "🌟 SECRET EASTER EGG UNLOCKED! +" + pointsAwarded + " Bonus Points added to your profile!";
  } else {
    displayMessage = "ACCESS GRANTED. +" + pointsAwarded + " Points added to your profile!";
  }

  return {
    success: true,
    challengeId: challengeId,
    challengeTitle: challengeTitle,
    solveTier: solveTier,
    firstBlood: (solveTier === "FIRST_BLOOD"),
    firstDay: (solveTier === "FIRST_DAY"),
    pointsAwarded: pointsAwarded,
    totalPoints: totalPoints,
    message: displayMessage,
    simMode: tables.isSim
  };
}

/**
 * Dynamically queries the currently active meeting session from the Admin_Config sheet.
 * Scans for the row where Is_Checkin_Open == TRUE (boolean true or case-insensitive "TRUE").
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
      passcodeHash: "",
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
      passcodeHash: "",
      message: "Check-in is currently CLOSED. No active meeting session is open right now."
    };
  }

  let colWeek = 0;
  let colPasscode = 1;
  let colIsOpen = 2;

  const headers = data[0];
  for (let c = 0; c < headers.length; c++) {
    const h = String(headers[c]).trim().toLowerCase();
    if (h === "active_week" || h === "week") colWeek = c;
    else if (h.includes("passcode")) colPasscode = c;
    else if (h === "is_checkin_open" || h === "is_open" || h === "open") colIsOpen = c;
  }

  // Scan for the row where Is_Checkin_Open == TRUE
  for (let i = 1; i < data.length; i++) {
    const rawOpen = data[i][colIsOpen];
    const isOpen = (rawOpen === true || String(rawOpen).trim().toUpperCase() === "TRUE");
    if (isOpen) {
      const activeWeek = parseInt(data[i][colWeek]);
      const rawPasscode = String(data[i][colPasscode] || "").trim();
      // If passcode is stored as hash, keep it; if plaintext, compute hash
      let pHash = "";
      let pPlain = "";
      if (rawPasscode.length === 64 && /^[0-9a-fA-F]+$/.test(rawPasscode)) {
        pHash = rawPasscode.toLowerCase();
      } else {
        pPlain = rawPasscode;
        pHash = hashSecret(rawPasscode);
      }

      return {
        isOpen: true,
        checkinOpen: true,
        activeWeek: activeWeek,
        passcode: pPlain,
        passcodeHash: pHash,
        message: "Active session found."
      };
    }
  }

  return {
    isOpen: false,
    checkinOpen: false,
    activeWeek: null,
    passcode: "",
    passcodeHash: "",
    message: "Check-in is currently CLOSED. No active meeting session is open right now."
  };
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
 * 5. Seeds Submissions_Log with welcome bonus records to enforce unique constraint.
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
    subSheet.getRange("A1:H1").setValues([["Timestamp", "Week", "Handle", "Email", "Challenge_ID", "Solve_Tier", "Points_Awarded", "Flag_Hash"]]);
    subSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  const subData = subSheet.getDataRange().getValues();
  const timestamp = new Date();
  const cipherHash = hashSecret("BaySec is here");

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
      const loggedChId = String(subData[i][4] || "").trim().toLowerCase();
      if (loggedEmail === solver.email.toLowerCase() && (loggedChId === "welcome_poster_cipher" || loggedChId.includes("cipher"))) {
        alreadyLogged = true;
        break;
      }
    }

    if (!alreadyLogged) {
      subSheet.appendRow([timestamp, 4, fallbackHandle, solver.email.toLowerCase(), "welcome_poster_cipher", "MIGRATED_WELCOME", 25, cipherHash]);
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
    simAttSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Passcode_Hash"]]);
    simAttSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#334155").setFontColor("#4ade80");
  }

  let simSubSheet = ss.getSheetByName("Submissions_Sim");
  if (!simSubSheet) {
    simSubSheet = ss.insertSheet("Submissions_Sim");
    simSubSheet.getRange("A1:H1").setValues([["Timestamp", "Week", "Handle", "Email", "Challenge_ID", "Solve_Tier", "Points_Awarded", "Flag_Hash"]]);
    simSubSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#334155").setFontColor("#f59e0b");
  }
}

/**
 * ONE-CLICK SETUP HELPER:
 * Initializes all required tabs including Admin_Config with Sim_Mode key and salted hashes.
 */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Admin_Config tab with salted passcode hashes
  let configSheet = ss.getSheetByName("Admin_Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Admin_Config");
    configSheet.getRange("A1:E1").setValues([["Active_Week", "Current_Passcode_Hash", "Is_Checkin_Open", "Setting_Key", "Setting_Value"]]);
    configSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#38bdf8");
    // Week 5 (WIRESHARK) salted hash, Sim_Mode: false
    configSheet.getRange("A2:E2").setValues([[5, "2cf4b557a8f9f865e9fab937277185ce2756be1477a63f119042c347a086e320", true, "Sim_Mode", false]]);
    // Week 4 (WELCOME) salted hash
    configSheet.getRange("A3:E3").setValues([[4, "fb60441c3a0bb2dd02459b6395d9c2a0447a0d96e11092697ae64d91729b6780", false, "", ""]]);
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
    logSheet.getRange("A1:E1").setValues([["Timestamp", "Week", "Handle", "Email", "Passcode_Hash"]]);
    logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#4ade80");
  }

  // 3. Setup Submissions_Log tab
  let subSheet = ss.getSheetByName("Submissions_Log");
  if (!subSheet) {
    subSheet = ss.insertSheet("Submissions_Log");
    subSheet.getRange("A1:H1").setValues([["Timestamp", "Week", "Handle", "Email", "Challenge_ID", "Solve_Tier", "Points_Awarded", "Flag_Hash"]]);
    subSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#1e293b").setFontColor("#f59e0b");
  }

  // 4. Setup Simulation/Staging tabs
  setupSimulationTables();
  
  try {
    SpreadsheetApp.getUi().alert("Setup complete! Admin_Config, Attendance_Log, Submissions_Log, and Simulation tabs are ready.");
  } catch(e) {
    Logger.log("Setup complete!");
  }
}
