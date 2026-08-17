function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("app")
    .setTitle("SCCK ROTC Command Center")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ─── MASTER SETUP ENGINE ─────────────────────────────────
function INITIAL_SETUP() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("UPLOADS_FOLDER_ID")) {
    let folder = DriveApp.createFolder("SCCK_SYSTEM_UPLOADS");
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    props.setProperty("UPLOADS_FOLDER_ID", folder.getId());
  }
  const accSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Accounts");
  if(accSheet) {
    accSheet.appendRow(["commander", hashData("scck2026"), "Super Admin", "Active", "Granted"]);
  }
  logAction("SYSTEM", "Initialization", "System Fortress Initialized. Super Admin created.");
}

// 🔥 CHANGE THIS TO YOUR ACTUAL EMAIL TO BYPASS BLOCKS 🔥
const SYSTEM_OWNER_EMAIL = "your.email@gmail.com";

// ─── CRYPTOGRAPHIC ENGINE (SHA-256) ───────────────────────
function hashData(input) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

// ─── SHEET NAMES ─────────────────────────────────────────
const S_ATTENDANCE = "Attendance";
const S_ACCOUNTS   = "Accounts";
const S_MASTERLIST = "Masterlist";
const S_LEAVES     = "Leaves";
const S_AUDIT      = "AuditLog"; 
const S_LEG_CLASS  = "Legacy_Classes";
const S_LEG_OFF    = "Legacy_Officers";
const S_POSTS      = "Public_Posts";
const S_DISCIPLINE = "Discipline";
const S_COMMAND    = "Command_Staff";

// ─── HELPER FUNCTIONS ────────────────────────────────────
function getColIdx(headers, possibleNames) {
  for (let i = 0; i < headers.length; i++) {
    let h = headers[i].toString().toLowerCase().replace(/\s/g, '');
    for (let name of possibleNames) {
      if (h === name.toLowerCase().replace(/\s/g, '')) return i;
    }
  }
  return -1;
}

function getDynamicCol(headers, aliases, fallback) {
  let idx = getColIdx(headers, aliases);
  return idx !== -1 ? idx : fallback;
}

function getVirtualSchedule() {
  const schedJson = PropertiesService.getScriptProperties().getProperty('TRAINING_SCHEDULE') || "[]";
  const schedule = JSON.parse(schedJson);
  let virtualDays = [];
  
  for (let s of schedule) {
    if (s.type === "Double") {
      virtualDays.push({ rawDate: s.date, dbStr: s.date + " - AM" });
      virtualDays.push({ rawDate: s.date, dbStr: s.date + " - PM" });
    } else {
      virtualDays.push({ rawDate: s.date, dbStr: s.date });
    }
  }
  
  // Cap at 15 official training blocks
  return virtualDays.slice(0, 15);
}

// ─── API ROUTER ──────────────────────────────────────────
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    const action = params.action;
    let result = {};

    if (action === "systemLogin") result = systemLogin(params.user, params.pass);
    else if (action === "getAdminDashboardData") result = getAdminDashboardData(params.user, params.pass);
    else if (action === "markAttendance") result = markAttendance(params.scannedData, params.user, params.pass, params.offlineDate, params.offlineTime);
    else if (action === "bulkMarkAttendance") result = bulkMarkAttendance(params.scans, params.user, params.pass);
    else if (action === "updateStudentInfo") result = updateStudentInfo(params.id, params.company, params.contact, params.user, params.pass);
    else if (action === "deleteStudentData") result = deleteStudentData(params.id, params.user, params.pass);
    else if (action === "createAccount") result = { message: createAccount(params.newUser, params.newPass, params.role, params.user, params.pass) };
    else if (action === "updateAccountStatus") result = { message: updateAccountStatus(params.targetUser, params.newStatus, params.user, params.pass) };
    else if (action === "updateLegacyAccess") result = { message: updateLegacyAccess(params.targetUser, params.newStatus, params.user, params.pass) };
    else if (action === "toggleLegacyAccess") result = { message: toggleLegacyAccess(params.targetUser, params.user, params.pass) };
    
    // Leaves
    else if (action === "submitLeaveRequest") result = submitLeaveRequest(params.studentId, params.fromDate, params.toDate, params.reason, params.user, params.pass);
    else if (action === "getLeaveRequests") result = getLeaveRequests(params.user, params.pass);
    else if (action === "updateLeaveStatus") result = updateLeaveStatus(params.leaveId, params.status, params.user, params.pass);
    
    // System & Schedule
    else if (action === "getStudentsForIDCards") result = getStudentsForIDCards(params.user, params.pass);
    else if (action === "getAuditLogs") result = getAuditLogs(params.user, params.pass);
    else if (action === "getSettings") result = getSettings();
    else if (action === "saveSettings") result = saveSettings(params.logoData, params.bgData, params.navLogoData, params.clearBg, params.user, params.pass);
    else if (action === "getProfile") result = getProfile(params.targetUser, params.user, params.pass);
    else if (action === "saveProfile") result = saveProfile(params.profileData, params.user, params.pass);
    else if (action === "saveTrainingSchedule") result = saveTrainingSchedule(params.schedule, params.user, params.pass);
    else if (action === "getTrainingSchedule") result = getTrainingSchedule(params.user, params.pass);
    
    // Cadets & Grading
    else if (action === "getCadetData") result = getCadetData(params.user, params.pass);
    else if (action === "getGradebookData") result = processAndGetGradebook(params.user, params.pass); 
    
    // Discipline
    else if (action === "bulkAddDiscipline") result = bulkAddDiscipline(params.targetCompany, params.type, params.points, params.reason, params.user, params.pass);
    else if (action === "addDiscipline") result = addDiscipline(params.studentId, params.type, params.points, params.reason, params.user, params.pass);
    else if (action === "getDiscipline") result = getDiscipline(params.studentId, params.user, params.pass);
    
    // Legacy DB
    else if (action === "addLegacyClass") result = addLegacyClass(params.year, params.name, params.logoData, params.user, params.pass);
    else if (action === "getLegacyClasses") result = getLegacyClasses(params.user, params.pass); 
    else if (action === "addLegacyOfficer") result = addLegacyOfficer(params.classId, params.rank, params.name, params.designation, params.course, params.bio, params.imageData, params.user, params.pass);
    else if (action === "getAdminLegacyData") result = getAdminLegacyData(params.user, params.pass); 
    else if (action === "updateLegacyClassStatus") result = updateLegacyClassStatus(params.classId, params.status, params.user, params.pass);
    else if (action === "deleteLegacyClass") result = deleteLegacyClass(params.classId, params.user, params.pass);
    else if (action === "deleteLegacyOfficer") result = deleteLegacyOfficer(params.offId, params.user, params.pass);
    else if (action === "getLegacyData") result = getLegacyData();
    else if (action === "requestEditOTP") result = requestEditOTP(params.user, params.pass);
    else if (action === "submitOfficerEdit") result = submitOfficerEdit(params.offId, params.rank, params.name, params.desig, params.course, params.bio, params.otp, params.user, params.pass);
    
    // Command History
    else if (action === "addCommandRecord") result = addCommandRecord(params.parentId, params.role, params.rank, params.name, params.tenure, params.status, params.quote, params.bio, params.imageData, params.user, params.pass);
    else if (action === "getCommandHistory") result = getCommandHistory();
    else if (action === "editCommandRecord") result = editCommandRecord(params.logId, params.rank, params.name, params.tenure, params.status, params.quote, params.bio, params.imageData, params.user, params.pass);
    else if (action === "deleteCommandRecord") result = deleteCommandRecord(params.logId, params.user, params.pass);
    
    // Posts
    else if (action === "submitPost") result = submitPost(params.title, params.description, params.imageArray, params.videoUrl, params.user, params.pass);
    else if (action === "getPostsAdmin") result = getPostsAdmin(params.user, params.pass);
    else if (action === "updatePostStatus") result = updatePostStatus(params.postId, params.status, params.feedback, params.user, params.pass);
    else if (action === "resubmitPost") result = resubmitPost(params.postId, params.title, params.description, params.imageArray, params.videoUrl, params.user, params.pass);
    else if (action === "getApprovedPosts") result = getApprovedPosts();
    else if (action === "deletePost") result = deletePost(params.postId, params.user, params.pass);
    else if (action === "togglePinPost") result = togglePinPost(params.postId, params.user, params.pass);
    // --- NEW FEATURES API ROUTES ---
    else if (action === "getSystemCompanies") result = { success: true, companies: getSystemCompanies() };
    else if (action === "saveSystemCompanies") result = saveSystemCompanies(params.companies, params.user, params.pass);
    else if (action === "getAnyCadetData") result = getAnyCadetData(params.targetId, params.user, params.pass);
    else if (action === "requestPasswordOTP") result = requestPasswordOTP(params.user, params.pass);
    else if (action === "confirmPasswordChange") result = confirmPasswordChange(params.otp, params.newPass, params.user, params.pass);
    else if (action === "getAttendanceDates") result = getAttendanceDates(params.user, params.pass);
    else if (action === "getHistoricalAttendance") result = getHistoricalAttendance(params.date, params.coy, params.user, params.pass);
    else throw new Error("Unknown action requested.");

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: true, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── AUTHENTICATION ──────────────────────────────────────
function checkAuth(user, pass, requiredRole) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_ACCOUNTS);
  if (!sheet) throw new Error("Accounts sheet missing from database.");
  
  const data = sheet.getDataRange().getDisplayValues();
  const inputHash = hashData(pass);
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === user && data[i][1] === inputHash && data[i][3] === "Active") {
      const role = data[i][2];
      if (requiredRole === "Super Admin" && role !== "Super Admin") throw new Error("Super Admin privileges required.");
      if (requiredRole === "Admin" && (role !== "Super Admin" && role !== "Admin")) throw new Error("Admin privileges required.");
      return { role: role, legacyClearance: data[i][4] || "Denied" };
    }
  }
  throw new Error("Invalid credentials or inactive account.");
}

function checkLegacyAuth(user, pass) {
  const auth = checkAuth(user, pass, "Admin");
  if (auth.role !== "Super Admin" && auth.legacyClearance !== "Granted") {
      throw new Error("CLEARANCE DENIED: You lack authorization to alter historical records. Contact the Commander.");
  }
  return auth.role;
}

function systemLogin(user, pass) {
  try {
    const auth = checkAuth(user, pass, "Any");
    const todayDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
    logAction(user, "System Login", `User authenticated successfully.`);
    return { success: true, role: auth.role, day: todayDate };
  } catch(e) {
    logAction(user || "UNKNOWN", "Failed Login", "Attempt failed: " + e.message);
    return { error: true, message: e.message };
  }
}

// ─── BLOCKCHAIN & AUDIT ───────────────────────────────────
function logAction(user, action, details) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_AUDIT);
    if (!sheet) return; 
    
    const tz = Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy HH:mm:ss");
    const lastRow = sheet.getLastRow();
    
    let prevSignature = "GENESIS_BLOCK";
    if (lastRow > 1) {
      prevSignature = sheet.getRange(lastRow, 5).getValue() || "GENESIS_BLOCK";
    }
    const newSignature = hashData(prevSignature + timestamp + user + action + details);
    
    sheet.appendRow(["'" + timestamp, user, action, details, newSignature]);
  } catch (e) {}
}

function getAuditLogs(user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    logAction(user, "Audit Log Accessed", "Super Admin viewed the cryptographic system audit logs.");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_AUDIT);
    if (!sheet) return { logs: [] };
    
    const data = sheet.getDataRange().getDisplayValues();
    const logs = [];
    let tampered = false;
    let tamperDetails = "";
    let prevSignature = "GENESIS_BLOCK";
    
    for (let i = 1; i < data.length; i++) {
      let t = data[i][0], u = data[i][1], a = data[i][2], d = data[i][3], sig = data[i][4];
      let expectedSig = hashData(prevSignature + t + u + a + d);
      
      if (sig !== expectedSig) {
          tampered = true;
          if(!tamperDetails) tamperDetails = `Chain broken at entry: [${t}] Operator: ${u} | Action: ${a}`;
      }
      prevSignature = sig;
      logs.push({ timestamp: t, user: u, action: a, details: d, hash: sig });
    }
    return { logs: logs.reverse(), tampered: tampered, tamperDetails: tamperDetails };
  } catch(e) { 
    return { error: true, message: e.message }; 
  }
}

// ─── SCHEDULE ENGINE ──────────────────────────────────────
function saveTrainingSchedule(schedule, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    PropertiesService.getScriptProperties().setProperty('TRAINING_SCHEDULE', JSON.stringify(schedule));
    logAction(user, "Schedule Updated", "Updated the official 15-day training schedule & session parameters.");
    return { success: true, message: "Training schedule securely saved." };
  } catch(e) { return { error: true, message: e.message }; }
}

function getTrainingSchedule(user, pass) {
  try {
    checkAuth(user, pass, "Any");
    let schedJson = PropertiesService.getScriptProperties().getProperty('TRAINING_SCHEDULE') || "[]";
    return { success: true, schedule: JSON.parse(schedJson) };
  } catch(e) { return { error: true, message: e.message }; }
}

// ─── ACADEMIC ENGINE (GRADEBOOK) ──────────────────────
function processAndGetGradebook(user, pass) {
  try {
    checkAuth(user, pass, "Any");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName(S_MASTERLIST) || ss.getSheetByName("MasterList");
    const masterData = mSheet.getDataRange().getValues();
    const attendance = ss.getSheetByName(S_ATTENDANCE).getDataRange().getValues();
    const discipline = ss.getSheetByName(S_DISCIPLINE).getDataRange().getValues();
    
    let prelimData = [], midData = [], finalData = [];
    try { prelimData = ss.getSheetByName("Exam_Prelim") ? ss.getSheetByName("Exam_Prelim").getDataRange().getValues() : []; } catch(e){}
    try { midData = ss.getSheetByName("Exam_Midterm") ? ss.getSheetByName("Exam_Midterm").getDataRange().getValues() : []; } catch(e){}
    try { finalData = ss.getSheetByName("Exam_Finals") ? ss.getSheetByName("Exam_Finals").getDataRange().getValues() : []; } catch(e){}

    const mHeaders = masterData[0];
    const idCol = getColIdx(mHeaders, ["studentnumber", "idnumber", "studentid"]);
    const fnCol = getColIdx(mHeaders, ["firstname", "givenname"]);
    const lnCol = getColIdx(mHeaders, ["lastname", "surname"]);
    const coyCol = getColIdx(mHeaders, ["company", "coy"]);

    const tz = Session.getScriptTimeZone();
    const todayObj = new Date(Utilities.formatDate(new Date(), tz, "MM/dd/yyyy"));
    const virtualDays = getVirtualSchedule();
    
    let totalActiveDrillDays = 0;
    for (let vDay of virtualDays) {
      if (new Date(vDay.rawDate) <= todayObj) {
        totalActiveDrillDays++;
      }
    }
    if (totalActiveDrillDays === 0) totalActiveDrillDays = 1; 

    let gradebookOutput = [];
    let gradebookSheet = ss.getSheetByName("Gradebook");
    if(!gradebookSheet) gradebookSheet = ss.insertSheet("Gradebook");
    
    let newGradebookData = [["Cadet ID", "Full Name", "Prelim Score", "Midterm Score", "Finals Score", "Exam Average", "Attendance %", "Aptitude %", "Final Grade"]];

    const dHeaders = discipline[0];
    const dIdCol = getDynamicCol(dHeaders, ["cadetid", "studentid", "idnumber", "studentnumber"], 1);
    const dTypeCol = getDynamicCol(dHeaders, ["type"], 2);
    const dPtsCol = getDynamicCol(dHeaders, ["points", "pts"], 3);

    for(let i = 1; i < masterData.length; i++) {
      let cId = idCol >= 0 ? masterData[i][idCol].toString().trim() : (masterData[i][1] ? masterData[i][1].toString().trim() : "");
      if(!cId) continue;
      
      let fName = fnCol >= 0 ? masterData[i][fnCol].toString().trim() : (masterData[i][2] ? masterData[i][2].toString().trim() : "");
      let lName = lnCol >= 0 ? masterData[i][lnCol].toString().trim() : (masterData[i][4] ? masterData[i][4].toString().trim() : "");
      let cName = [fName, lName].filter(Boolean).join(" ");
      let cCompany = coyCol >= 0 ? masterData[i][coyCol] : "";
      
      let daysPresent = 0;
      for(let a = 1; a < attendance.length; a++) {
        if(attendance[a][0].toString().trim() === cId) {
          let stat = attendance[a][5];
          if(['Present', 'Late', 'Very Late'].includes(stat)) daysPresent++;
        }
      }
      let attPct = Math.min(100, Math.round((daysPresent / totalActiveDrillDays) * 100));

      let aptScore = 100;
      for(let d = 1; d < discipline.length; d++) {
        let dId = discipline[d][dIdCol].toString().trim();
        if(dId === cId) {
          let dType = discipline[d][dTypeCol] ? discipline[d][dTypeCol].toString().trim().toUpperCase() : "";
          let rawPts = parseInt(discipline[d][dPtsCol], 10) || 0;
          let absPts = Math.abs(rawPts);
          
          if (dType === "DEMERIT" || dType === "DEDUCTION") {
              aptScore -= absPts;
          } else if (dType === "MERIT") {
              aptScore += absPts;
          } else {
              aptScore += rawPts;
          }
        }
      }
      aptScore = Math.max(0, Math.min(100, aptScore));

      let prelim = extractScore(cId, prelimData), midterm = extractScore(cId, midData), finals = extractScore(cId, finalData);
      let examAvg = Math.round((prelim + midterm + finals) / 3);
      let finalGrade = ((attPct * 0.30) + (aptScore * 0.30) + (examAvg * 0.40)).toFixed(2);

      gradebookOutput.push({ id: cId, name: cName, company: cCompany, daysAttended: daysPresent, prelim: prelim, midterm: midterm, final: finals, examAvg: examAvg, attendancePct: attPct, aptitudePct: aptScore, finalGrade: finalGrade });
      newGradebookData.push([cId, cName, prelim, midterm, finals, examAvg, attPct, aptScore, finalGrade]);
    }

    gradebookSheet.clear();
    gradebookSheet.getRange(1, 1, newGradebookData.length, newGradebookData[0].length).setValues(newGradebookData);
    return { success: true, gradebook: gradebookOutput };
  } catch(e) { return { error: true, message: e.message }; }
}

function extractScore(targetId, formSheetData) {
  if(!formSheetData || formSheetData.length < 2) return 0;
  let idCol = -1, scoreCol = -1;
  for(let c = 0; c < formSheetData[0].length; c++) {
    let header = formSheetData[0][c].toString().toLowerCase().replace(/\s/g, '');
    if(header.includes("id") || header.includes("studentnumber")) idCol = c;
    if(header.includes("score")) scoreCol = c;
  }
  if(idCol === -1) idCol = 1; 
  if(scoreCol === -1) scoreCol = 2;

  for(let r = 1; r < formSheetData.length; r++) {
    if(formSheetData[r][idCol] && formSheetData[r][idCol].toString().trim() === targetId) {
      let rawScore = formSheetData[r][scoreCol];
      if(rawScore && rawScore.toString().includes('/')) return Math.round((parseInt(rawScore.toString().split('/')[0].trim()) / parseInt(rawScore.toString().split('/')[1].trim())) * 100);
      return parseInt(rawScore || 0);
    }
  }
  return 0;
}


// ─── CACHED ADMIN DASHBOARD ──────────────────────────────
function getAdminDashboardData(user, pass) {
  const cache = CacheService.getScriptCache();
  
  // REVERTED TO STANDARD KEY FOR INSTANT SYNCING
  const cacheKey = 'cache_admin_dashboard';
  
  const cachedData = cache.get(cacheKey);
  
  if (cachedData != null) {
    return JSON.parse(cachedData);
  }

  try {
    checkAuth(user, pass, "Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const todayDate = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");
    const todayObj = new Date(todayDate);
    
    const accSheet = ss.getSheetByName(S_ACCOUNTS);
    const masterSheet = ss.getSheetByName(S_MASTERLIST) || ss.getSheetByName("MasterList");
    const attSheet = ss.getSheetByName(S_ATTENDANCE);
    if (!accSheet || !masterSheet || !attSheet) throw new Error("A core database tab is missing or misspelled.");
    
    const accounts = accSheet.getDataRange().getDisplayValues();
    const masterData = masterSheet.getDataRange().getDisplayValues();
    const attData = attSheet.getDataRange().getDisplayValues();

    const mHeaders = masterData[0];
    const idCol = getColIdx(mHeaders, ["studentnumber", "idnumber", "studentid"]);
    const fnCol = getColIdx(mHeaders, ["firstname", "givenname"]);
    const mnCol = getColIdx(mHeaders, ["middlename", "m.i.", "middleinitial"]);
    const lnCol = getColIdx(mHeaders, ["lastname", "surname"]);
    const coyCol = getColIdx(mHeaders, ["company", "coy"]);
    const conCol = getColIdx(mHeaders, ["emergencycontact", "emergencycontactnumber", "contact"]);

    let discMap = {};
    const discSheet = ss.getSheetByName(S_DISCIPLINE);
    if (discSheet && discSheet.getLastRow() > 1) {
      const discData = discSheet.getDataRange().getDisplayValues();
      const dHeaders = discData[0];
      const dIdCol = getDynamicCol(dHeaders, ["cadetid", "studentid", "idnumber", "studentnumber"], 1);
      const dTypeCol = getDynamicCol(dHeaders, ["type"], 2);
      const dPtsCol = getDynamicCol(dHeaders, ["points", "pts"], 3);

      for (let k = 1; k < discData.length; k++) {
        let dId = discData[k][dIdCol].toString().trim();
        let dType = discData[k][dTypeCol].toString().trim().toUpperCase();
        let dPts = parseInt(discData[k][dPtsCol], 10) || 0;
        
        if (dId) {
          if (discMap[dId] === undefined) discMap[dId] = 100;
          
          let absPts = Math.abs(dPts);
          if (dType === "DEMERIT" || dType === "DEDUCTION") {
            discMap[dId] -= absPts;
          } else if (dType === "MERIT") {
            discMap[dId] += absPts;
          } else {
            discMap[dId] += dPts;
          }
        }
      }
    }

    const attMap = {};
    for (let j = 1; j < attData.length; j++) {
      const rowId = attData[j][0].replace(/'/g, "").trim(), rowDate = attData[j][2];
      const status = attData[j][5] || "Present";
      if (!attMap[rowId]) attMap[rowId] = {};
      attMap[rowId][rowDate] = { timeIn: attData[j][3] || "---", timeOut: attData[j][4] || "---", status: status };
    }

    const schedJson = PropertiesService.getScriptProperties().getProperty('TRAINING_SCHEDULE') || "[]";
    const schedule = JSON.parse(schedJson);
    let todaySchedObj = schedule.find(s => s.date === todayDate);
    let activeTodayString = todayDate;
    
    if (todaySchedObj && todaySchedObj.type === "Double") {
        const currentHour = parseInt(Utilities.formatDate(new Date(), tz, "H"), 10);
        activeTodayString += (currentHour < 12) ? " - AM" : " - PM";
    }

    let mergedList = [], rawMasterlist = [], todayPresent = 0, todayLate = 0;
    for (let i = 1; i < masterData.length; i++) {
      const studentId = idCol >= 0 ? masterData[i][idCol].toString().trim() : (masterData[i][1] ? masterData[i][1].toString().trim() : "");
      if (!studentId) continue;
      
      const fName = fnCol >= 0 ? masterData[i][fnCol].toString().trim() : (masterData[i][2] ? masterData[i][2].toString().trim() : "");
      let mName = mnCol >= 0 ? masterData[i][mnCol].toString().trim() : (masterData[i][3] ? masterData[i][3].toString().trim() : "");
      const lName = lnCol >= 0 ? masterData[i][lnCol].toString().trim() : (masterData[i][4] ? masterData[i][4].toString().trim() : "");
      
      if (mName.toUpperCase() === "N/A" || mName.toUpperCase() === "NONE" || mName === "-" || mName === "") mName = "";
      const fullName = [fName, mName, lName].filter(Boolean).join(" ");
      const company = coyCol >= 0 ? masterData[i][coyCol] : "";
      const contact = conCol >= 0 ? masterData[i][conCol] : "";
      
      let currentScore = discMap[studentId] !== undefined ? discMap[studentId] : 100;
      currentScore = Math.max(0, Math.min(100, currentScore));
      
      rawMasterlist.push({ id: studentId, name: fullName, company: company, contact: contact, score: currentScore });
      
      const todayRec = attMap[studentId] && attMap[studentId][activeTodayString];
      const timeIn = todayRec ? todayRec.timeIn : "---", timeOut = todayRec ? todayRec.timeOut : "---", status = todayRec ? todayRec.status : "Absent";

      if (status === "Present") todayPresent++; 
      if (status === "Late" || status === "Very Late") todayLate++;
      if (todayRec) mergedList.push({ id: studentId, name: fullName, in: timeIn, out: timeOut, status: status, course: company });
    }
    
    const virtualDays = getVirtualSchedule();
    let wLabels = [], wPresent = [], wLate = [], wAbsent = [];
    let totalPastDays = 0;

    for(let i=0; i<15; i++) {
      wLabels.push("D" + (i+1));
      if(i < virtualDays.length) {
        let vDay = virtualDays[i];
        let vDateObj = new Date(vDay.rawDate);
        if(vDateObj <= todayObj) {
           totalPastDays++;
           let pCount = 0, lCount = 0, aCount = 0;
           for(let m = 1; m < masterData.length; m++) {
              let sId = idCol >= 0 ? masterData[m][idCol].toString().trim() : (masterData[m][1] ? masterData[m][1].toString().trim() : "");
              if(!sId) continue;
              
              let rec = attMap[sId] ? attMap[sId][vDay.dbStr] : null;
              if(!rec) aCount++;
              else if(rec.status === 'Present') pCount++;
              else if(rec.status === 'Late' || rec.status === 'Very Late') lCount++;
              else aCount++;
           }
           wPresent.push(pCount); wLate.push(lCount); wAbsent.push(aCount);
        } else {
           wPresent.push(0); wLate.push(0); wAbsent.push(0);
        }
      } else {
         wPresent.push(0); wLate.push(0); wAbsent.push(0);
      }
    }

    const weeklyData = { labels: wLabels, present: wPresent, late: wLate, absent: wAbsent };

    let pendLv = 0, pendPst = 0, pendLeg = 0;
    const lvSheet = ss.getSheetByName(S_LEAVES);
    if(lvSheet && lvSheet.getLastRow() > 1) {
        const lvData = lvSheet.getDataRange().getValues();
        for(let i=1; i<lvData.length; i++) { if(lvData[i][6] === "Pending") pendLv++; }
    }
    const pstSheet = ss.getSheetByName(S_POSTS);
    if(pstSheet && pstSheet.getLastRow() > 1) {
        const pstData = pstSheet.getDataRange().getValues();
        for(let i=1; i<pstData.length; i++) { if(pstData[i][4] === "Pending" || pstData[i][4] === "Revise") pendPst++; }
    }
    const legSheet = ss.getSheetByName(S_LEG_CLASS);
    if(legSheet && legSheet.getLastRow() > 1) {
        const legData = legSheet.getDataRange().getValues();
        for(let i=1; i<legData.length; i++) { if(legData[i][4] === "Pending") pendLeg++; }
    }
    
    logAction(user, "Dashboard Accessed", "Accessed Main Command Dashboard.");
    
    const responseData = { 
      status: "ok", 
      accounts: accounts, 
      masterlist: rawMasterlist, 
      attendance: mergedList, 
      activeDay: activeTodayString, 
      weeklyData: weeklyData, 
      trainingDaysCompleted: totalPastDays, 
      stats: { total: rawMasterlist.length, present: todayPresent, late: todayLate, absent: rawMasterlist.length - (todayPresent + todayLate) }, 
      alerts: { leaves: pendLv, posts: pendPst, legacy: pendLeg } 
    };

    cache.put(cacheKey, JSON.stringify(responseData), 300);
    return responseData;

  } catch(e) { return { error: true, message: e.message }; }
}


// ─── CACHED CADET DATA (UPDATED FOR ADMIN VIEW) ───────────
function getCadetData(user, pass, targetIdOverride = null) {
  const cacheKey = targetIdOverride ? 'cache_admin_view_' + targetIdOverride : 'cache_cadet_' + user;
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get(cacheKey);
  
  if (cachedData != null) {
    return JSON.parse(cachedData);
  }

  try {
    let userId;
    if (targetIdOverride) {
      checkAuth(user, pass, "Admin");
      userId = targetIdOverride.toString().trim();
    } else {
      checkAuth(user, pass, "Any");
      userId = user.toString().trim();
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const attSheet = ss.getSheetByName(S_ATTENDANCE);
    const discSheet = ss.getSheetByName(S_DISCIPLINE);
    const lvSheet = ss.getSheetByName(S_LEAVES);

    let score = 100;
    let attList = [], discList = [], lvList = [];
    
    if (discSheet && discSheet.getLastRow() > 1) {
      const discData = discSheet.getDataRange().getDisplayValues();
      const dHeaders = discData[0];
      const dIdCol = getDynamicCol(dHeaders, ["cadetid", "studentid", "idnumber", "studentnumber"], 1);
      const dTypeCol = getDynamicCol(dHeaders, ["type"], 2);
      const dPtsCol = getDynamicCol(dHeaders, ["points", "pts"], 3);
      const dReasonCol = getDynamicCol(dHeaders, ["reason"], 4);
      const dDateCol = getDynamicCol(dHeaders, ["date", "timestamp"], 5);
      const dAdminCol = getDynamicCol(dHeaders, ["admin", "user"], 6);

      for (let i = 1; i < discData.length; i++) {
        if (discData[i][dIdCol].toString().trim() === userId) {
          let dType = discData[i][dTypeCol].toString().trim().toUpperCase();
          let rawPts = parseInt(discData[i][dPtsCol], 10) || 0;
          let absPts = Math.abs(rawPts);
          
          if (dType === "DEMERIT" || dType === "DEDUCTION") {
              score -= absPts;
              rawPts = -absPts; // Keep negative for frontend ledger presentation
          } else if (dType === "MERIT") {
              score += absPts;
              rawPts = absPts;
          } else {
              score += rawPts;
          }
          
          discList.push({ 
            date: discData[i][dDateCol], 
            type: discData[i][dTypeCol], 
            reason: discData[i][dReasonCol], 
            points: rawPts > 0 ? "+" + rawPts : rawPts, 
            admin: discData[i][dAdminCol] 
          });
        }
      }
    }
    score = Math.max(0, Math.min(100, score));

    if (lvSheet && lvSheet.getLastRow() > 1) {
      const lvData = lvSheet.getDataRange().getDisplayValues();
      for (let i = 1; i < lvData.length; i++) {
        if (lvData[i][1].toString().trim() === userId) {
          lvList.push({ 
             id: lvData[i][0], 
             from: lvData[i][3], 
             to: lvData[i][4], 
             reason: lvData[i][5], 
             status: lvData[i][6], 
             date: lvData[i][7], 
             admin: lvData[i][8] || "Pending" 
          });
        }
      }
    }

    const tz = Session.getScriptTimeZone();
    const todayObj = new Date(Utilities.formatDate(new Date(), tz, "MM/dd/yyyy"));
    const virtualDays = getVirtualSchedule();
    
    let attData = [];
    if (attSheet && attSheet.getLastRow() > 1) {
      attData = attSheet.getDataRange().getDisplayValues();
    }
    
    let dayTrackerArr = [], p = 0, a = 0;
    let totalActiveDrillDays = 0;

    for (let i = 0; i < 15; i++) {
      if (i < virtualDays.length) {
        let vDay = virtualDays[i];
        let vDateObj = new Date(vDay.rawDate);
        if (vDateObj > todayObj) {
          dayTrackerArr.push("Upcoming");
        } else {
          totalActiveDrillDays++;
          let foundStatus = "Absent";
          for (let j = 1; j < attData.length; j++) {
            if (attData[j][0].replace(/'/g, "").trim() === userId && attData[j][2].toString() === vDay.dbStr) {
              foundStatus = attData[j][5]; 
              break;
            }
          }
          dayTrackerArr.push(foundStatus);
          if (['Present', 'Late', 'Very Late'].includes(foundStatus)) p++; else a++;
        }
      } else {
        dayTrackerArr.push("-"); 
      }
    }

    for (let i = 1; i < attData.length; i++) {
      if (attData[i][0].replace(/'/g, "").trim() === userId) {
        attList.push({ date: attData[i][2], in: attData[i][3], out: attData[i][4], status: attData[i][5] });
      }
    }
    
    let attPct = 0;
    if (totalActiveDrillDays > 0) attPct = Math.min(100, Math.round((p / totalActiveDrillDays) * 100));

    let finalGrade = "0.00", examAvg = "0";
    let gradeData = processAndGetGradebook(user, pass);
    if (gradeData.success) {
      for(let g=0; g<gradeData.gradebook.length; g++) {
          if(gradeData.gradebook[g].id === userId) {
              finalGrade = gradeData.gradebook[g].finalGrade; 
              examAvg = gradeData.gradebook[g].examAvg; 
              break;
          }
      }
    }

    logAction(user, "Cadet Portal Accessed", `Viewed dossier for Cadet ID: ${userId}`);
    
    const responseData = { score: score, presentDays: p, absentDays: a, attendance: attList.reverse().slice(0,15), discipline: discList.reverse(), leaves: lvList.reverse(), tracker: dayTrackerArr, attendancePct: attPct, finalGrade: finalGrade, examAvg: examAvg };
    
    cache.put(cacheKey, JSON.stringify(responseData), 300);
    return responseData;

  } catch(e) { return { error: true, message: e.message }; }
}

// ─── SHIELDED ATTENDANCE ENGINE (LOCK SERVICE) ────────────
function bulkMarkAttendance(scans, user, pass) {
  try {
    checkAuth(user, pass, "Any");
    let results = [];
    for (let i = 0; i < scans.length; i++) {
      let res = markAttendance(scans[i].scannedData, user, pass, scans[i].offlineDate, scans[i].offlineTime);
      results.push(res);
    }
    logAction(user, "Bulk Offline Sync", `Successfully synced ${scans.length} cached offline scans to the central database.`);
    return { success: true, results: results };
  } catch(e) { return { error: true, message: e.message }; }
}

function markAttendance(scannedData, user, pass, offlineDate, offlineTime) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(3000); // Enforce a neat line for simultaneous scanners
  } catch (e) {
    return { status: "warn", msg: "Traffic jam. Please scan again." };
  }

  try {
    checkAuth(user, pass, "Any");
    const cleanId = scannedData.toString().replace(/^[=+\-@\t\r]/, "").trim();
    if (!cleanId) return { status: "warn", msg: "Invalid QR code format." };
    const ss = SpreadsheetApp.getActiveSpreadsheet(), attSheet = ss.getSheetByName(S_ATTENDANCE), masterSheet = ss.getSheetByName(S_MASTERLIST) || ss.getSheetByName("MasterList");
    const tz = Session.getScriptTimeZone();
    let studentName = "", studentFound = false, matchedId = cleanId, studentImage = "";
    
    if (masterSheet) {
      const mData = masterSheet.getDataRange().getDisplayValues();
      const mHeaders = mData[0];
      
      let imgColIdx = getColIdx(mHeaders, ["2x2formalphoto", "formalphoto", "photo", "image", "picture"]);
      if (imgColIdx === -1) imgColIdx = findImageColumnIndex(mHeaders);

      const idCol = getColIdx(mHeaders, ["studentnumber", "idnumber", "studentid"]);
      const fnCol = getColIdx(mHeaders, ["firstname", "givenname"]);
      const mnCol = getColIdx(mHeaders, ["middlename", "m.i.", "middleinitial"]);
      const lnCol = getColIdx(mHeaders, ["lastname", "surname"]);
      
      for (let i = 1; i < mData.length; i++) {
        const sheetId = idCol >= 0 ? mData[i][idCol].toString().trim() : mData[i][1].toString().trim();
        if (sheetId === cleanId || parseInt(sheetId, 10) === parseInt(cleanId, 10)) {
          const fName = fnCol >= 0 ? mData[i][fnCol] : mData[i][2];
          let mName = mnCol >= 0 ? mData[i][mnCol] : mData[i][3];
          const lName = lnCol >= 0 ? mData[i][lnCol] : mData[i][4];
          if (mName.toUpperCase() === "N/A" || mName.toUpperCase() === "NONE" || mName === "-" || mName === "") mName = "";
          studentName = [fName, mName, lName].filter(Boolean).join(" ");
          
          studentFound = true; matchedId = sheetId;
          let rawImage = imgColIdx >= 0 ? mData[i][imgColIdx] : "";
          studentImage = convertDriveLink(rawImage);
          break; 
        }
      }
    }

    if (!studentFound) {
       logAction(user, "Scan Rejected", `Unauthorized QR scan attempt. Scanned Data: ${cleanId}`);
       return { status: "warn", msg: "UNAUTHORIZED QR: Cadet not found in the Masterlist." };
    }

    const todayDate = offlineDate ? offlineDate : Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");
    const timestamp = offlineTime ? offlineTime : Utilities.formatDate(new Date(), tz, "hh:mm:ss a");
    let currentHour, currentMin;
    if (offlineTime) {
        const timeParts = offlineTime.split(/:| /);
        currentHour = parseInt(timeParts[0], 10);
        currentMin = parseInt(timeParts[1], 10);
        if (timeParts[3] === "PM" && currentHour !== 12) currentHour += 12;
        if (timeParts[3] === "AM" && currentHour === 12) currentHour = 0;
    } else {
        currentHour = parseInt(Utilities.formatDate(new Date(), tz, "H"), 10);
        currentMin = parseInt(Utilities.formatDate(new Date(), tz, "m"), 10);
    }

    const schedJson = PropertiesService.getScriptProperties().getProperty('TRAINING_SCHEDULE') || "[]";
    const schedule = JSON.parse(schedJson);
    const todaySchedObj = schedule.find(s => s.date === todayDate);
    
    if (!todaySchedObj && schedule.length > 0) {
        return { status: "warn", msg: "NO TRAINING SCHEDULED FOR TODAY." };
    }

    let activeDbDate = todayDate;
    let isAM = false;

    if (todaySchedObj && todaySchedObj.type === "Double") {
        if (currentHour < 12) {
            activeDbDate += " - AM";
            isAM = true;
        } else {
            activeDbDate += " - PM";
        }
    }

    let arrivalStatus = "Present";
    if (isAM) {
        if (currentHour < 7 || (currentHour === 7 && currentMin === 0)) arrivalStatus = "Present";
        else if (currentHour === 7 && currentMin >= 1 && currentMin <= 15) arrivalStatus = "Late";
        else arrivalStatus = "Very Late";
    } else {
        if (currentHour < 13 || (currentHour === 13 && currentMin === 0)) arrivalStatus = "Present";
        else if (currentHour === 13 && currentMin >= 1 && currentMin <= 15) arrivalStatus = "Late";
        else arrivalStatus = "Very Late";
    }
    
    const attData = attSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < attData.length; i++) {
      const rowId = attData[i][0].replace(/'/g, "").trim();
      if (rowId === matchedId && attData[i][2] === activeDbDate) {
        if (attData[i][3] && attData[i][3] !== "---" && (!attData[i][4] || attData[i][4] === "---")) {
          attSheet.getRange(i + 1, 5).setValue(timestamp);
          logAction(user, "Time Out Logged", `Recorded TIME OUT for Cadet ID: ${matchedId} (${studentName})`);
          
          // INVALIDATE CACHE
          CacheService.getScriptCache().remove('cache_admin_dashboard');
          CacheService.getScriptCache().remove('cache_cadet_' + cleanId);
          
          return { msg: "TIME OUT Recorded!", status: "ok", id: matchedId, name: studentName, image: studentImage, timeIn: attData[i][3], timeOut: timestamp, date: activeDbDate, action: "TIME OUT" };
        } else if (attData[i][3] !== "---" && attData[i][4] !== "---") {
          return { msg: "Already completed Time In & Out for this session!", status: "warn", name: studentName, image: studentImage };
        }
      }
    }

    attSheet.appendRow(["'" + matchedId, studentName, activeDbDate, timestamp, "---", arrivalStatus]);
    logAction(user, "Time In Logged", `Recorded ${arrivalStatus.toUpperCase()} TIME IN for Cadet ID: ${matchedId} (${studentName})`);
    
    // INVALIDATE CACHE
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    CacheService.getScriptCache().remove('cache_cadet_' + cleanId);

    return { msg: "TIME IN Recorded!", status: "ok", id: matchedId, name: studentName, image: studentImage, timeIn: timestamp, timeOut: "---", date: activeDbDate, action: arrivalStatus.toUpperCase() };
  } catch(e) { 
    return { status: "warn", msg: e.message }; 
  } finally {
    // ALWAYS RELEASE THE LOCK
    lock.releaseLock();
  }
}

// ─── BATCH DISCIPLINE ENGINE ──────────────────────────────
function bulkAddDiscipline(targetCompany, type, points, reason, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName(S_MASTERLIST) || ss.getSheetByName("MasterList");
    if (!masterSheet) throw new Error("Masterlist not found.");
    
    const mData = masterSheet.getDataRange().getDisplayValues();
    const mHeaders = mData[0];
    const idCol = getColIdx(mHeaders, ["studentnumber", "idnumber", "studentid"]);
    const coyCol = getColIdx(mHeaders, ["company", "coy"]);
    
    let targetIds = [];
    
    for (let i = 1; i < mData.length; i++) {
      let sId = idCol >= 0 ? mData[i][idCol].toString().trim() : (mData[i][1] ? mData[i][1].toString().trim() : "");
      let coy = coyCol >= 0 ? mData[i][coyCol].toString().trim() : (mData[i][5] ? mData[i][5].toString().trim() : "");
      
      if (sId) {
        if (targetCompany === "ALL" || coy.toUpperCase() === targetCompany.toUpperCase()) {
          targetIds.push(sId);
        }
      }
    }

    if (targetIds.length === 0) return { error: true, message: "No cadets found in that target group." };

    let discSheet = ss.getSheetByName(S_DISCIPLINE);
    if (!discSheet) {
      discSheet = ss.insertSheet(S_DISCIPLINE);
      discSheet.appendRow(["Log ID", "Cadet ID", "Type", "Points", "Reason", "Date", "Admin"]);
    }

    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");
    let rowsToAdd = [];
    
    targetIds.forEach((id, idx) => {
      const logId = "DSC_" + new Date().getTime() + "_" + idx;
      rowsToAdd.push([logId, id, type, points, reason, today, user]);
    });

    discSheet.getRange(discSheet.getLastRow() + 1, 1, rowsToAdd.length, 7).setValues(rowsToAdd);
    logAction(user, "Batch Discipline", `Issued ${type} (${points} pts) to ${targetCompany}. Target count: ${targetIds.length}. Reason: ${reason}`);
    
    // Invalidate dashboard cache
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    
    return { success: true, message: `Successfully issued ${type} to ${targetIds.length} cadets.` };
  } catch(e) { return { error: true, message: e.message }; }
}

function addDiscipline(studentId, type, points, reason, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(S_DISCIPLINE);
    if (!sheet) {
      sheet = ss.insertSheet(S_DISCIPLINE);
      sheet.appendRow(["Log ID", "Cadet ID", "Type", "Points", "Reason", "Date", "Admin"]);
    }
    const logId = "DSC_" + new Date().getTime();
    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, "MM/dd/yyyy");
    
    sheet.appendRow([logId, studentId, type, points, reason, today, user]);
    logAction(user, "Discipline Logged", `Issued ${type} (${points} pts) to Cadet ID: ${studentId}. Reason: ${reason}`);
    
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    CacheService.getScriptCache().remove('cache_cadet_' + studentId);
    
    return { success: true, message: "Discipline record updated." };
  } catch(e) { return { error: true, message: e.message }; }
}

function getDiscipline(studentId, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_DISCIPLINE);
    if (!sheet) return { records: [] };
    
    const data = sheet.getDataRange().getDisplayValues();
    const dHeaders = data[0];
    const dIdCol = getDynamicCol(dHeaders, ["cadetid", "studentid", "idnumber", "studentnumber"], 1);
    const dTypeCol = getDynamicCol(dHeaders, ["type"], 2);
    const dPtsCol = getDynamicCol(dHeaders, ["points", "pts"], 3);
    const dReasonCol = getDynamicCol(dHeaders, ["reason"], 4);
    const dDateCol = getDynamicCol(dHeaders, ["date", "timestamp"], 5);
    const dAdminCol = getDynamicCol(dHeaders, ["admin", "user"], 6);

    const records = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][dIdCol].toString().trim() === studentId.toString().trim()) {
        let dType = data[i][dTypeCol].toString().trim().toUpperCase();
        let rawPts = parseInt(data[i][dPtsCol], 10) || 0;
        let absPts = Math.abs(rawPts);
        
        if (dType === "DEMERIT" || dType === "DEDUCTION") {
            rawPts = -absPts;
        } else if (dType === "MERIT") {
            rawPts = absPts;
        }

        records.push({ 
          id: data[i][0], 
          type: data[i][dTypeCol], 
          points: rawPts, 
          reason: data[i][dReasonCol], 
          date: data[i][dDateCol], 
          admin: data[i][dAdminCol] 
        });
      }
    }
    return { records: records.reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

// ─── MASTERLIST & ACCOUNTS ────────────────────────────────
function updateStudentInfo(id, company, contact, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MASTERLIST) || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterList");
    const data  = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let idCol = getColIdx(headers, ["studentnumber", "idnumber", "studentid"]);
    if(idCol === -1) idCol = 1; 
    
    let coyCol = getColIdx(headers, ["company", "coy"]);
    if(coyCol === -1) {
      coyCol = headers.length; 
      sheet.getRange(1, coyCol + 1).setValue("Company");
      for(let j=1; j<data.length; j++) data[j].push(""); 
    }
    
    let conCol = getColIdx(headers, ["emergencycontact", "emergencycontactnumber", "contact"]);
    if(conCol === -1) conCol = 10; 
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol].toString().trim() === id.toString().trim()) {
        sheet.getRange(i + 1, coyCol + 1).setValue(company);
        sheet.getRange(i + 1, conCol + 1).setValue(contact); 
        logAction(user, "Cadet Dossier Updated", `Updated Unit Roster data for Cadet ID: ${id}. New Company: ${company} | Contact: ${contact}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true };
      }
    }
    return { error: "Student not found." };
  } catch(e) { return { error: true, message: e.message }; }
}

function deleteStudentData(id, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MASTERLIST) || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterList");
    const data  = sheet.getDataRange().getValues();
    const idCol = getColIdx(data[0], ["studentnumber", "idnumber", "studentid"]) >= 0 ? getColIdx(data[0], ["studentnumber", "idnumber", "studentid"]) : 1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idCol].toString().trim() === id.toString().trim()) {
        sheet.deleteRow(i + 1);
        logAction(user, "Cadet Discharged", `WARNING: Permanently discharged and deleted Cadet ID: ${id} from the Unit Masterlist.`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true };
      }
    }
    return { error: "Student not found." };
  } catch(e) { return { error: true, message: e.message }; }
}

function createAccount(newUser, newPass, role, user, pass) {
  checkAuth(user, pass, "Super Admin");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_ACCOUNTS);
  const data  = sheet.getDataRange().getDisplayValues();
  for (let i = 1; i < data.length; i++) { if (data[i][0] === newUser) return "Username already exists."; }
  let legAccess = (role === "Super Admin") ? "Granted" : "Denied";
  sheet.appendRow([newUser, hashData(newPass), role, "Active", legAccess]);
  logAction(user, "Account Created", `Provisioned new ${role} account for username: ${newUser}`);
  CacheService.getScriptCache().remove('cache_admin_dashboard');
  return "Account created successfully.";
}

function updateAccountStatus(targetUser, newStatus, user, pass) {
  checkAuth(user, pass, "Super Admin"); 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_ACCOUNTS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUser) { 
      sheet.getRange(i + 1, 4).setValue(newStatus);
      logAction(user, "Account Status Update", `Changed clearance status of ${targetUser} to ${newStatus}`);
      CacheService.getScriptCache().remove('cache_admin_dashboard');
      return "Status updated.";
  }
  }
  return "User not found.";
}

function toggleLegacyAccess(targetUser, user, pass) {
  checkAuth(user, pass, "Super Admin"); 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_ACCOUNTS);
  const data  = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUser) { 
      let current = data[i][4];
      let newStatus = (current === "Granted") ? "Denied" : "Granted";
      sheet.getRange(i + 1, 5).setValue(newStatus);
      logAction(user, "Legacy Clearance Toggle", `Changed Legacy DB clearance for ${targetUser} to ${newStatus}`);
      CacheService.getScriptCache().remove('cache_admin_dashboard');
      return "Legacy access updated.";
    }
  }
  return "User not found.";
}

// ─── LEAVE REQUESTS & UTILS ───────────────────────────────
function submitLeaveRequest(studentId, fromDate, toDate, reason, user, pass) {
  try {
    checkAuth(user, pass, "Any"); 
    const ss = SpreadsheetApp.getActiveSpreadsheet(), mData = (ss.getSheetByName(S_MASTERLIST) || ss.getSheetByName("MasterList")).getDataRange().getDisplayValues();
    
    const idCol = getColIdx(mData[0], ["studentnumber", "idnumber", "studentid"]);
    const fnCol = getColIdx(mData[0], ["firstname", "givenname"]);
    const mnCol = getColIdx(mData[0], ["middlename", "m.i.", "middleinitial"]);
    const lnCol = getColIdx(mData[0], ["lastname", "surname"]);
    
    let studentName = "Unknown";
    for (let i = 1; i < mData.length; i++) { 
      const sId = idCol >= 0 ? mData[i][idCol].toString().trim() : mData[i][1].toString().trim();
      if (sId === studentId.toString().trim()) { 
        const fName = fnCol >= 0 ? mData[i][fnCol] : mData[i][2];
        let mName = mnCol >= 0 ? mData[i][mnCol] : mData[i][3];
        const lName = lnCol >= 0 ? mData[i][lnCol] : mData[i][4];
        if (mName.toUpperCase() === "N/A" || mName.toUpperCase() === "NONE" || mName === "-" || mName === "") mName = "";
        studentName = [fName, mName, lName].filter(Boolean).join(" "); break;
      } 
    }
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy"), leaveId = "LV" + new Date().getTime();
    ss.getSheetByName(S_LEAVES).appendRow([leaveId, studentId, studentName, fromDate, toDate, reason, "Pending", today, ""]);
    logAction(user, "Leave Submitted", `Submitted leave request for Cadet ID: ${studentId} (${fromDate} to ${toDate})`);
    
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    CacheService.getScriptCache().remove('cache_cadet_' + studentId);
    
    return { success: true, message: "Leave request submitted successfully." };
  } catch(e) { return { success: false, message: e.message }; }
}

function getLeaveRequests(user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEAVES);
    if (!sheet) return { leaves: [] }; 
    const data = sheet.getDataRange().getDisplayValues(), leaves = [];
    for (let i = 1; i < data.length; i++) leaves.push({ id: data[i][0], studentId: data[i][1], name: data[i][2], from: data[i][3], to: data[i][4], reason: data[i][5], status: data[i][6] });
    return { leaves: leaves.reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

function updateLeaveStatus(leaveId, status, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEAVES);
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === leaveId) { 
        sheet.getRange(i + 1, 7).setValue(status);
        sheet.getRange(i + 1, 9).setValue(user); // Records Admin Name
        logAction(user, "Leave Processed", `Changed leave request ${leaveId} status to ${status}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        CacheService.getScriptCache().remove('cache_cadet_' + data[i][1]);
        return { success: true };
      }
    }
    return { success: false, message: "Leave record not found." };
  } catch(e) { return { success: false, message: e.message }; }
}

function getStudentsForIDCards(user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_MASTERLIST) || SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MasterList");
    if (!sheet) return { students: [] };
    const data = sheet.getDataRange().getDisplayValues(), students = [];
    const headers = data[0];
    
    let imgColIdx = getColIdx(headers, ["2x2formalphoto", "formalphoto", "photo", "image", "picture"]);
    if (imgColIdx === -1) imgColIdx = findImageColumnIndex(headers);
    
    const idCol = getColIdx(headers, ["studentnumber", "idnumber", "studentid"]);
    const fnCol = getColIdx(headers, ["firstname", "givenname"]);
    const mnCol = getColIdx(headers, ["middlename", "m.i.", "middleinitial"]);
    const lnCol = getColIdx(headers, ["lastname", "surname"]);
    const coyCol = getColIdx(headers, ["company", "coy"]);
    const courseCol = getColIdx(headers, ["course", "program"]); 
    
    for (let i = 1; i < data.length; i++) {
      const sId = idCol >= 0 ? data[i][idCol] : data[i][1];
      if (sId) {
        const fName = fnCol >= 0 ? data[i][fnCol] : data[i][2];
        let mName = mnCol >= 0 ? data[i][mnCol] : data[i][3];
        const lName = lnCol >= 0 ? data[i][lnCol] : data[i][4];
        if (mName.toUpperCase() === "N/A" || mName.toUpperCase() === "NONE" || mName === "-" || mName === "") mName = "";
        
        let rawImage = imgColIdx >= 0 ? data[i][imgColIdx] : "";
        let courseVal = courseCol >= 0 ? data[i][courseCol].toString().trim() : (data[i][7] ? data[i][7].toString().trim() : "");
        let compVal = coyCol >= 0 ? data[i][coyCol].toString().trim() : (data[i][5] ? data[i][5].toString().trim() : "");
        
        students.push({ id: sId, name: [fName, mName, lName].filter(Boolean).join(" "), course: courseVal, company: compVal, image: convertDriveLink(rawImage) });
      }
    }
    return { students: students };
  } catch(e) { return { error: true, message: e.message }; }
}

function uploadImageToDrive(base64Data, fileNamePrefix) {
  const props = PropertiesService.getScriptProperties();
  let folderId = props.getProperty("UPLOADS_FOLDER_ID");
  const cleanBase64 = base64Data.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, "");
  const blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), "image/png", fileNamePrefix + "_" + new Date().getTime() + ".png");
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  return "https://lh3.googleusercontent.com/d/" + file.getId();
}

function convertDriveLink(url) {
  if (!url) return "";
  let finalUrl = url.toString().split(',')[0].trim();
  if (finalUrl.includes("drive.google.com")) {
    let match = finalUrl.match(/id=([a-zA-Z0-9_-]+)/);
    if (!match) match = finalUrl.match(/d\/([a-zA-Z0-9_-]+)/);
    if (match && match[1]) return "https://lh3.googleusercontent.com/d/" + match[1];
  }
  return finalUrl;
}

function findImageColumnIndex(headers) {
  for (let i = headers.length - 1; i >= 0; i--) { 
    let h = headers[i].toString().toLowerCase().replace(/\s/g, '');
    if ((h.includes("image") && !h.includes("qr")) || h.includes("photo") || h.includes("formal")) return i;
  }
  return 18;
}

// ─── POST / NEWSFEED MANAGEMENT ───────────────────────────
function submitPost(title, description, imageArray, videoUrl, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    if(!sheet) throw new Error("Public_Posts tab missing in database.");
    const postId = "PST_" + new Date().getTime();
    let imgUrls = [];
    if (imageArray && imageArray.length > 0) {
      for (let i = 0; i < imageArray.length; i++) imgUrls.push(uploadImageToDrive(imageArray[i], "POST_" + new Date().getTime() + "_" + i));
    }
    const finalImgString = imgUrls.join(","); 
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy");
    sheet.appendRow([postId, title, description, finalImgString, "Pending", today, "", false, videoUrl || ""]);
    logAction(user, "Post Drafted", `Drafted public post: ${title}`);
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    return { success: true, message: "Post drafted and pending approval." };
  } catch(e) { return { error: true, message: e.message }; }
}

function resubmitPost(postId, title, description, imageArray, videoUrl, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === postId) {
        sheet.getRange(i + 1, 2).setValue(title);
        sheet.getRange(i + 1, 3).setValue(description);
        if (imageArray && imageArray.length > 0) {
          let imgUrls = [];
          for (let j = 0; j < imageArray.length; j++) imgUrls.push(uploadImageToDrive(imageArray[j], "POST_" + new Date().getTime() + "_" + j));
          sheet.getRange(i + 1, 4).setValue(imgUrls.join(","));
        }
        sheet.getRange(i + 1, 5).setValue("Pending"); 
        sheet.getRange(i + 1, 7).setValue("");
        sheet.getRange(i + 1, 9).setValue(videoUrl || ""); 
        logAction(user, "Post Revised", `Admin revised and resubmitted post: ${title}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true, message: "Post revised and sent to Commander for review." };
      }
    }
    return { success: false, message: "Original post not found." };
  } catch(e) { return { error: true, message: e.message }; }
}

function getPostsAdmin(user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    if (!sheet) return { posts: [] }; 
    const data = sheet.getDataRange().getDisplayValues();
    const posts = [];
    for (let i = 1; i < data.length; i++) {
      posts.push({ id: data[i][0], title: data[i][1], description: data[i][2], image: data[i][3], status: data[i][4], date: data[i][5], feedback: data[i][6] || "", pinned: (data[i][7] === "TRUE" || data[i][7] === "true" || data[i][7] === true), videoUrl: data[i][8] || "" });
    }
    return { success: true, posts: posts.reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

function updatePostStatus(postId, status, feedback, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === postId) {
        sheet.getRange(i + 1, 5).setValue(status);
        sheet.getRange(i + 1, 7).setValue(feedback || "");
        logAction(user, "Post Reviewed", `Changed post ${postId} status to ${status}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true };
      }
    }
    return { success: false, message: "Post not found." };
  } catch(e) { return { success: false, message: e.message }; }
}

function deletePost(postId, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    const data = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === postId) { 
        sheet.deleteRow(i + 1);
        logAction(user, "Post Deleted", `Permanently deleted public post ID: ${postId}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true, message: "Post permanently annihilated." };
      }
    }
    return { success: false, message: "Post not found." };
  } catch(e) { return { success: false, message: e.message }; }
}

function togglePinPost(postId, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === postId) { 
        let currentStatus = data[i][7];
        let newStatus = (currentStatus === true || currentStatus === "TRUE" || currentStatus === "true") ? false : true;
        sheet.getRange(i + 1, 8).setValue(newStatus);
        logAction(user, "Post Pin Toggled", `Toggled pin status for post ${postId} to ${newStatus}`);
        return { success: true, pinned: newStatus };
      }
    }
    return { success: false, message: "Post not found." };
  } catch(e) { return { success: false, message: e.message }; }
}

function getApprovedPosts() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_POSTS);
    if(!sheet) return { success: true, posts: [] }; 
    const data = sheet.getDataRange().getDisplayValues();
    const posts = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] === "Approved") {
        let isPinned = (data[i][7] === "TRUE" || data[i][7] === "true" || data[i][7] === true);
        posts.push({ id: data[i][0], title: data[i][1], description: data[i][2], image: data[i][3], date: data[i][5], pinned: isPinned, videoUrl: data[i][8] || "" });
      }
    }
    return { success: true, posts: posts.reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

// ─── BRANDING & PROFILE ───────────────────────────────────
function getSettings() {
  try {
    const props = PropertiesService.getScriptProperties().getProperties();
    return { logoUrl: props.logoUrl || "https://cdn-icons-png.flaticon.com/512/1055/1055664.png", navLogoUrl: props.navLogoUrl || props.logoUrl || "https://cdn-icons-png.flaticon.com/512/1055/1055664.png", bgUrl: props.bgUrl || "" };
  } catch(e) { return {}; }
}

function saveSettings(logoData, bgData, navLogoData, clearBg, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const props = PropertiesService.getScriptProperties();
    let updates = {};
    if (logoData) updates.logoUrl = uploadImageToDrive(logoData, "SYS_LOGO");
    if (bgData) updates.bgUrl = uploadImageToDrive(bgData, "SYS_BG");
    if (navLogoData) updates.navLogoUrl = uploadImageToDrive(navLogoData, "SYS_NAV_LOGO");
    if (clearBg) updates.bgUrl = ""; 
    if (Object.keys(updates).length > 0) props.setProperties(updates);
    logAction(user, "Settings Updated", "Updated system branding interface.");
    return { success: true, message: "System interface updated!" };
  } catch(e) { return { error: true, message: e.message }; }
}

function getProfile(targetUser, user, pass) {
  try {
    checkAuth(user, pass, "Any");
    const data = PropertiesService.getScriptProperties().getProperty('profile_' + targetUser);
    return data ? JSON.parse(data) : {};
  } catch(e) { return {}; }
}

function saveProfile(profileData, user, pass) {
  try {
    checkAuth(user, pass, "Any");
    const props = PropertiesService.getScriptProperties();
    const existing = JSON.parse(props.getProperty('profile_' + user) || "{}");
    if (profileData.imageData) {
      profileData.image = uploadImageToDrive(profileData.imageData, "PROF_" + user);
      delete profileData.imageData; 
    } else profileData.image = existing.image || "";
    props.setProperty('profile_' + user, JSON.stringify(profileData));
    logAction(user, "Profile Updated", "Updated personal account profile details.");
    return { success: true, message: "Profile saved successfully!" };
  } catch(e) { return { error: true, message: e.message }; }
}

// ─── LEGACY DATABASE ──────────────────────────────────────
function addLegacyClass(year, name, logoData, user, pass) {
  try {
    let role = checkLegacyAuth(user, pass); 
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(S_LEG_CLASS);
    if (!sheet) {
      sheet = ss.insertSheet(S_LEG_CLASS);
      sheet.appendRow(["Class ID", "Year", "Class Name", "Logo URL", "Status"]);
    }
    const classId = "CLS_" + new Date().getTime();
    let logoUrl = logoData ? uploadImageToDrive(logoData, "CLASS_" + year) : "";
    let status = (role === "Super Admin") ? "Approved" : "Pending";
    sheet.appendRow([classId, year, name, logoUrl, status]);
    logAction(user, "Legacy Class Drafted", `Requested Class: ${name} (${year})`);
    CacheService.getScriptCache().remove('cache_admin_dashboard');
    return { success: true, message: (role === "Super Admin" ? "Class officially added to archives." : "Class submitted for Commander approval!") };
  } catch(e) { return { error: true, message: e.message }; }
}

function updateLegacyClassStatus(classId, status, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEG_CLASS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === classId) { 
        sheet.getRange(i + 1, 5).setValue(status);
        logAction(user, "Legacy Class Reviewed", `Changed Class ${classId} status to ${status}`);
        CacheService.getScriptCache().remove('cache_admin_dashboard');
        return { success: true };
      }
    }
    return { success: false, message: "Class not found." };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteLegacyClass(classId, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEG_CLASS);
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
       if(data[i][0] === classId) {
          sheet.deleteRow(i+1);
          logAction(user, "Legacy Class Deleted", `Permanently Deleted Class ID: ${classId}`);
          CacheService.getScriptCache().remove('cache_admin_dashboard');
          return {success:true};
       }
    }
    return {success:false, message:"Not found."};
  } catch(e) { return {success:false, message:e.message}; }
}

function deleteLegacyOfficer(offId, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEG_OFF);
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) {
       if(data[i][0] === offId) {
          sheet.deleteRow(i+1);
          logAction(user, "Legacy Officer Deleted", `Permanently Deleted Officer ID: ${offId}`);
          return {success:true};
       }
    }
    return {success:false, message:"Not found."};
  } catch(e) { return {success:false, message:e.message}; }
}

function getLegacyClasses(user, pass) {
  try {
    checkAuth(user, pass, "Admin"); 
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEG_CLASS);
    if (!sheet) return { classes: [] };
    const data = sheet.getDataRange().getDisplayValues();
    const classes = [];
    for (let i = 1; i < data.length; i++) {
      let status = data[i][4] || "Approved";
      if (status === "Approved") classes.push({ id: data[i][0], year: data[i][1], name: data[i][2] });
    }
    return { classes: classes };
  } catch(e) { return { error: true, message: e.message }; }
}

function getAdminLegacyData(user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cSheet = ss.getSheetByName(S_LEG_CLASS);
    const oSheet = ss.getSheetByName(S_LEG_OFF);
    let classes = [];
    if(cSheet && cSheet.getLastRow() > 1) {
      const cData = cSheet.getDataRange().getDisplayValues();
      for(let i=1; i<cData.length; i++) {
        let status = cData[i][4] || "Approved"; 
        classes.push({ id: cData[i][0], year: cData[i][1], name: cData[i][2], logo: cData[i][3], status: status });
      }
    }
    let officers = [];
    if(oSheet && oSheet.getLastRow() > 1) {
      const oData = oSheet.getDataRange().getDisplayValues();
      for(let i=1; i<oData.length; i++) {
        officers.push({ id: oData[i][0], classId: oData[i][1], rank: oData[i][2], name: oData[i][3], designation: oData[i][4], image: oData[i][5], course: oData[i][6] || "", bio: oData[i][7] || "" });
      }
    }
    return { success: true, classes: classes.reverse(), officers: officers.reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

function addLegacyOfficer(classId, rank, name, designation, course, bio, imageData, user, pass) {
  try {
    checkLegacyAuth(user, pass);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(S_LEG_OFF);
    if (!sheet) {
      sheet = ss.insertSheet(S_LEG_OFF);
      sheet.appendRow(["Officer ID", "Class ID", "Rank", "Name", "Designation", "Image URL", "Course", "Bio"]);
    }
    const offId = "OFF_" + new Date().getTime();
    let imgUrl = imageData ? uploadImageToDrive(imageData, "OFFICER_" + name.replace(/\s+/g,"_")) : "";
    sheet.appendRow([offId, classId, rank, name, designation, imgUrl, course, bio]);
    logAction(user, "Legacy Officer Added", `Added ${rank} ${name} to Class ID: ${classId}`);
    return { success: true, message: "Officer permanently added to archives." };
  } catch(e) { return { error: true, message: e.message }; }
}

function requestEditOTP(user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    PropertiesService.getScriptProperties().setProperty("EDIT_OTP", otp);
    
    MailApp.sendEmail({
      to: SYSTEM_OWNER_EMAIL,
      subject: "SCCK ROTC - Legacy DB Edit Verification Code",
      body: `A Super Admin (${user}) is attempting to edit a Legacy Officer record.\n\nYour OTP is: ${otp}\n\nDo not share this code. If you did not authorize this, please check the system logs immediately.`
    });
    logAction(user, "OTP Requested", "Requested OTP to edit legacy officer.");
    return { success: true, message: "OTP sent to database owner." };
  } catch(e) { return { error: true, message: e.message }; }
}

function submitOfficerEdit(offId, rank, name, desig, course, bio, otp, user, pass) {
   try {
     checkAuth(user, pass, "Super Admin");
     const savedOtp = PropertiesService.getScriptProperties().getProperty("EDIT_OTP");
     if (!savedOtp || savedOtp !== otp) {
         return { error: true, message: "Invalid or expired OTP." };
     }
     PropertiesService.getScriptProperties().deleteProperty("EDIT_OTP");
     
     const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_LEG_OFF);
     if (!sheet) throw new Error("Legacy Officer tab missing.");
     const data = sheet.getDataRange().getValues();
     for (let i = 1; i < data.length; i++) {
       if (data[i][0] === offId) {
         sheet.getRange(i + 1, 3).setValue(rank);
         sheet.getRange(i + 1, 4).setValue(name);
         sheet.getRange(i + 1, 5).setValue(desig);
         sheet.getRange(i + 1, 7).setValue(course);
         sheet.getRange(i + 1, 8).setValue(bio);
         logAction(user, "Legacy Officer Edited", `Updated officer ID: ${offId} via OTP verification.`);
         return { success: true, message: "Officer record updated securely." };
       }
     }
     return { error: true, message: "Officer not found in database." };
   } catch(e) { return { error: true, message: e.message }; }
}

function getLegacyData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cSheet = ss.getSheetByName(S_LEG_CLASS);
    const oSheet = ss.getSheetByName(S_LEG_OFF);
    
    let classes = [];
    if (cSheet && cSheet.getLastRow() > 0) {
      const cData = cSheet.getDataRange().getDisplayValues();
      for(let i=1; i<cData.length; i++) {
         let status = cData[i][4] || "Approved"; 
         if (cData[i][0] && status === "Approved") {
             classes.push({ id: cData[i][0], year: cData[i][1], name: cData[i][2], logo: cData[i][3], officers: [] });
         }
      }
    }
    
    if (oSheet && oSheet.getLastRow() > 0) {
      const oData = oSheet.getDataRange().getDisplayValues();
      for(let i=1; i<oData.length; i++) {
        if (oData[i][0]) {
            const off = { id: oData[i][0], classId: oData[i][1], rank: oData[i][2], name: oData[i][3], designation: oData[i][4], image: oData[i][5], course: oData[i][6] || "", bio: oData[i][7] || "" };
            const targetClass = classes.find(c => c.id === off.classId);
            if (targetClass) targetClass.officers.push(off);
        }
      }
    }
    return { success: true, classes: classes, branding: getSettings() };
  } catch(e) { 
    return { success: false, error: true, message: e.message, classes: [], branding: getSettings() };
  }
}

// ─── CADET REGISTRATION ENGINE (GOOGLE FORMS TRIGGER) ─────────────
function onCadetRegistration(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const responses = e.namedValues;
    
    const masterSheet = ss.getSheetByName("Masterlist") || ss.getSheetByName("MasterList");
    if (!masterSheet) return;
    
    const headers = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
    let newRow = new Array(headers.length).fill("");
    
    function mapToRow(headerAliases, value) {
        let idx = getColIdx(headers, headerAliases);
        if (idx !== -1) newRow[idx] = value;
    }

    const studentNo = getFormField(responses, ["Student Number", "ID Number", "ID No."]);
    const sex = getFormField(responses, ["Sex", "Gender"]);
    if (!studentNo) return; 

    mapToRow(["Timestamp", "Time"], responses["Timestamp"] ? responses["Timestamp"][0] : "");
    mapToRow(["Email Address", "Email"], getFormField(responses, ["Email Address", "Email"]));
    mapToRow(["Student Number", "ID Number", "StudentID"], studentNo);
    mapToRow(["First Name", "Firstname", "Given Name", "First name "], getFormField(responses, ["First Name", "Firstname", "Given Name", "First name "]));
    mapToRow(["Middle Name", "Middle Initial", "M.I."], getFormField(responses, ["Middle Name", "Middle name", "M.I.", "Middle Initial"]));
    mapToRow(["Last Name", "Surname"], getFormField(responses, ["Last Name", "Last name", "Surname"]));
    mapToRow(["Course", "Program"], getFormField(responses, ["Course", "Program"]));
    mapToRow(["Sex", "Gender"], sex);
    mapToRow(["Date of Birth", "Birthdate", "DOB"], getFormField(responses, ["Date of Birth", "Birthdate", "DOB"]));
    mapToRow(["Emergency Contact Name", "Emergency Contact", "Guardian's Full Name", "Guardian Name"], getFormField(responses, ["Emergency Contact Name", "Emergency Contact", "Guardian's Full Name", "Guardian Name"]));
    mapToRow(["Emergency Contact Relationship", "Relationship"], getFormField(responses, ["Emergency Contact Relationship", "Relationship"]));
    mapToRow(["Emergency Contact Address", "Guardian Address"], getFormField(responses, ["Emergency Contact Address", "Guardian Address"]));
    mapToRow(["Emergency Contact Number", "Guardians Contact Number"], getFormField(responses, ["Emergency Contact Number", "Guardians Contact Number", "Guardians Contact Number "]));
    mapToRow(["Gmail Address", "Gmail"], getFormField(responses, ["Gmail Address", "Gmail"]));
    mapToRow(["Full Address", "Address"], getFormField(responses, ["Full Address", "Address"]));
    
    let rawImageUpload = getFormField(responses, ["2x2 Formal Photo", "Formal Photo", "Please send your formal image for your ID", "Photo", "Image"]);
    let imageUrl = rawImageUpload ? rawImageUpload.split(",")[0].trim() : "";
    mapToRow(["2x2 Formal Photo", "Image", "Photo", "Image URL", "Formal Photo"], imageUrl);
    
    const qrUrl = "https://quickchart.io/qr?size=300&margin=2&text=" + encodeURIComponent(studentNo);
    mapToRow(["QR", "QR Code", "QR Link"], qrUrl);

    masterSheet.appendRow(newRow);
    
    const genderRow = [
      responses["Timestamp"] ? responses["Timestamp"][0] : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"),
      getFormField(responses, ["Email Address", "Email"]),
      studentNo,
      getFormField(responses, ["First Name", "Firstname", "Given Name", "First name "]),
      getFormField(responses, ["Middle Name", "Middle name", "M.I.", "Middle Initial"]),
      getFormField(responses, ["Last Name", "Last name", "Surname"]),
      getFormField(responses, ["Course", "Program"]),
      imageUrl,
      getFormField(responses, ["Date of Birth", "Birthdate", "DOB"]),
      sex,
      getFormField(responses, ["Emergency Contact Name", "Guardian Name", "Guardian's Full Name"]),
      getFormField(responses, ["Emergency Contact Relationship", "Relationship"]),
      getFormField(responses, ["Emergency Contact Address", "Guardian Address"]),
      getFormField(responses, ["Emergency Contact Number", "Guardians Contact Number", "Guardians Contact Number "])
    ];

    if (sex.toLowerCase() === "male" && ss.getSheetByName("Male")) {
        ss.getSheetByName("Male").appendRow(genderRow);
    } else if (sex.toLowerCase() === "female" && ss.getSheetByName("Female")) {
        ss.getSheetByName("Female").appendRow(genderRow);
    }

    const accSheet = ss.getSheetByName(S_ACCOUNTS);
    if (accSheet) {
        const accData = accSheet.getDataRange().getValues();
        let accExists = false;
        for (let k = 1; k < accData.length; k++) {
            if (accData[k][0].toString() === studentNo.toString()) { accExists = true; break; }
        }
        if (!accExists) {
            accSheet.appendRow([studentNo, hashData(studentNo.toString()), "Cadet", "Active", "Denied"]);
        }
    }

    const email = getFormField(responses, ["Email Address", "Email"]);
    if (email) {
      const fName = getFormField(responses, ["First Name", "Firstname", "Given Name", "First name "]);
      const lName = getFormField(responses, ["Last Name", "Last name", "Surname"]);
      const fullName = [fName, lName].filter(Boolean).join(" ");
      const htmlBody = `<div style="font-family: 'Segoe UI', sans-serif; max-width: 500px; margin: 0 auto; text-align: center; border: 3px solid #1b2015; border-radius: 12px; padding: 30px; background-color: #f4f7f9;">
        <h2 style="color: #2d3821; margin-bottom: 5px; font-size: 24px;">SCCK ROTC COMMAND CENTER</h2>
        <p style="color: #d4af37; font-weight: bold; letter-spacing: 1px; margin-top: 0;">VISITOR & ATTENDANCE SYSTEM</p>
        <p style="font-size: 16px; color: #1a2c45; margin-top: 25px;">Welcome, <strong>${fullName}</strong>.</p>
        <p style="font-size: 14px; color: #6b7f96;">Your registration was successful. Below is your official assigned QR Code.</p>
        <div style="margin: 20px 0; background: #ffffff; padding: 20px; border-radius: 8px; display: inline-block; box-shadow: 0 4px 10px rgba(0,0,0,0.1);"><img src="${qrUrl}" alt="Your QR Code" style="width: 250px; height: 250px; display: block;"></div>
        <p style="font-size: 22px; font-weight: bold; color: #1b2015; margin: 10px 0;">ID: ${studentNo}</p>
        <div style="margin-top: 25px; padding: 15px; background: #e8ece1; border-radius: 8px; text-align: left;">
          <p style="font-size: 13px; color: #2d3821; margin: 0 0 5px 0;"><strong>YOUR CADET PORTAL ACCESS:</strong></p>
          <p style="font-size: 12px; color: #6b7f96; margin: 0;"><strong>Username:</strong> ${studentNo}</p>
          <p style="font-size: 12px; color: #6b7f96; margin: 0;"><strong>Password:</strong> ${studentNo}</p>
          <p style="font-size: 11px; color: #d32f2f; margin-top: 8px;"><em>Log in to track your attendance, discipline score, and submit leaves.</em></p>
        </div>
      </div>`;
      MailApp.sendEmail({ to: email, subject: "SCCK ROTC - Your Official QR ID & Portal Access", htmlBody: htmlBody });
    }
    
    logAction("SYSTEM", "New Cadet Registered", `Auto-segregated Cadet, generated QR, and created portal account for ID: ${studentNo}`);
  } catch (error) { logAction("SYSTEM", "Registration Automation Failed", `Failed routing data: ${error.message}`); }
}

function getFormField(responses, possibleKeys) {
  for (let key in responses) {
    let cleanKey = key.trim().toLowerCase();
    for (let name of possibleKeys) {
      if (cleanKey === name.toLowerCase()) { return responses[key][0] ? responses[key][0].trim() : ""; }
    }
  }
  return "";
}

// ─── COMMAND ERA & HALL OF FAME ENGINE ─────────────────────

function addCommandRecord(parentId, role, rank, name, tenure, status, quote, bio, imageData, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(S_COMMAND);
    if (!sheet) throw new Error("Command_Staff tab missing in database.");
    
    const logId = "CMD_" + new Date().getTime();
    let imgUrl = imageData ? uploadImageToDrive(imageData, "CMD_" + name.replace(/\s+/g,"_")) : "";
    sheet.appendRow([logId, parentId || "", role, rank, name, tenure, status, quote, bio, imgUrl]);
    
    logAction(user, "Command Staff Added", `Added ${role}: ${rank} ${name}`);
    return { success: true, message: "Leader permanently recorded in Unit History." };
  } catch(e) { return { error: true, message: e.message }; }
}

function getCommandHistory() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_COMMAND);
    if (!sheet) return { success: true, eras: [] };
    const data = sheet.getDataRange().getDisplayValues();

    let commandants = [];
    let staff = [];

    for (let i = 1; i < data.length; i++) {
      let record = { id: data[i][0], parentId: data[i][1], role: data[i][2], rank: data[i][3], name: data[i][4], tenure: data[i][5], status: data[i][6], quote: data[i][7], bio: data[i][8], image: data[i][9] };
      if (record.id) { if (record.role === "Commandant") commandants.push(record); else staff.push(record); }
    }

    let eras = commandants.map(cmd => { cmd.corpsCommanders = staff.filter(s => s.parentId === cmd.id); return cmd; });
    eras.sort((a, b) => (a.status === "Current" ? -1 : (b.status === "Current" ? 1 : 0)));

    return { success: true, eras: eras };
  } catch(e) { return { error: true, message: e.message }; }
}

function deleteCommandRecord(logId, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_COMMAND);
    const data = sheet.getDataRange().getValues();
    for(let i=1; i<data.length; i++) { if(data[i][0] === logId) { sheet.deleteRow(i+1); logAction(user, "Command Staff Deleted", `Removed record ID: ${logId}`); return {success:true, message: "Record removed."}; } }
    return {success:false, message:"Not found."};
  } catch(e) { return {success:false, message:e.message}; }
}

function editCommandRecord(logId, rank, name, tenure, status, quote, bio, imageData, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_COMMAND);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === logId) {
        let imgUrl = data[i][9]; 
        if (imageData) { 
          imgUrl = uploadImageToDrive(imageData, "CMD_" + name.replace(/\s+/g, "_"));
        }
        
        sheet.getRange(i + 1, 4).setValue(rank);
        sheet.getRange(i + 1, 5).setValue(name);
        sheet.getRange(i + 1, 6).setValue(tenure);
        sheet.getRange(i + 1, 7).setValue(status);
        sheet.getRange(i + 1, 8).setValue(quote);
        sheet.getRange(i + 1, 9).setValue(bio);
        if (imageData) sheet.getRange(i + 1, 10).setValue(imgUrl);
        
        logAction(user, "Command Staff Edited", `Updated record ID: ${logId}`);
        return { success: true, message: "Leader's dossier successfully updated." };
      }
    }
    throw new Error("Record not found.");
  } catch(e) { return { success: false, message: e.message }; }
}
// =========================================================================
// NEW FEATURES: COMPANIES, PASSWORD 2FA, AND ADMIN CADET VIEW
// =========================================================================

// --- 1. Dynamic Companies ---
function getSystemCompanies() {
  const props = PropertiesService.getScriptProperties().getProperty("OFFICIAL_COMPANIES");
  return props ? JSON.parse(props) : ["Alpha", "Bravo", "Charlie", "Headquarters"];
}

function saveSystemCompanies(companiesArr, user, pass) {
  try {
    checkAuth(user, pass, "Super Admin");
    PropertiesService.getScriptProperties().setProperty("OFFICIAL_COMPANIES", JSON.stringify(companiesArr));
    logAction(user, "Settings Updated", "Super Admin updated official company list.");
    return { success: true, message: "Companies updated." };
  } catch (e) { return { error: true, message: e.message }; }
}

/// --- 2. Admin View Any Cadet ---
function getAnyCadetData(targetId, user, pass) {
  try {
    checkAuth(user, pass, "Admin"); // Ensure the requester is an Admin
    // Pass the Admin user/pass, and inject the targetId as the 3rd parameter Override
    return getCadetData(user, pass, targetId); 
  } catch (e) { return { error: true, message: e.message }; }
}


// --- 3. Password OTP & 2FA ---
function requestPasswordOTP(user, pass) {
  try {
    checkAuth(user, pass, "Any");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName("Masterlist") || ss.getSheetByName("MasterList");
    const mData = masterSheet.getDataRange().getValues();
    const emailCol = getColIdx(mData[0], ["emailaddress", "email", "gmail", "gmailaddress"]);
    const idCol = getColIdx(mData[0], ["studentnumber", "idnumber", "studentid"]);
    
    let targetEmail = "";
    for (let i = 1; i < mData.length; i++) {
      if (mData[i][idCol].toString().trim() === user.toString().trim()) {
        targetEmail = mData[i][emailCol];
        break;
      }
    }
    
    if (!targetEmail) {
      // Fallback for Super Admins who aren't in the Masterlist
      if (user === "commander") targetEmail = SYSTEM_OWNER_EMAIL;
      else throw new Error("No email registered for this account.");
    }
    
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    PropertiesService.getScriptProperties().setProperty("PWD_OTP_" + user, otp);
    
    MailApp.sendEmail({
      to: targetEmail,
      subject: "SCCK ROTC - Password Change OTP",
      body: `Your One-Time Password (OTP) to change your account password is: ${otp}\n\nIf you did not request this, please notify the Command Administrator immediately.`
    });
    
    logAction(user, "OTP Requested", "User requested OTP for password change.");
    return { success: true, message: "OTP sent to your registered email." };
  } catch (e) { return { error: true, message: e.message }; }
}

function confirmPasswordChange(otp, newPass, user, pass) {
  try {
    checkAuth(user, pass, "Any");
    const savedOtp = PropertiesService.getScriptProperties().getProperty("PWD_OTP_" + user);
    if (!savedOtp || savedOtp !== otp.toString().trim()) {
      throw new Error("Invalid or expired OTP.");
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const accSheet = ss.getSheetByName("Accounts");
    const data = accSheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === user.toString()) {
        accSheet.getRange(i + 1, 2).setValue(hashData(newPass));
        PropertiesService.getScriptProperties().deleteProperty("PWD_OTP_" + user);
        logAction(user, "Password Changed", "User successfully changed their password via 2FA.");
        return { success: true, message: "Password updated successfully!" };
      }
    }
    throw new Error("Account not found.");
  } catch (e) { return { error: true, message: e.message }; }
}

// --- 4. HISTORICAL ATTENDANCE VIEWER ---
function getAttendanceDates(user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const attData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Attendance").getDataRange().getDisplayValues();
    let dates = new Set();
    for(let i=1; i<attData.length; i++) { if(attData[i][2]) dates.add(attData[i][2]); }
    return { success: true, dates: Array.from(dates).reverse() };
  } catch(e) { return { error: true, message: e.message }; }
}

function getHistoricalAttendance(targetDate, targetCoy, user, pass) {
  try {
    checkAuth(user, pass, "Admin");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const attData = ss.getSheetByName("Attendance").getDataRange().getDisplayValues();
    const mData = (ss.getSheetByName("Masterlist") || ss.getSheetByName("MasterList")).getDataRange().getDisplayValues();
    
    let coyCol = getColIdx(mData[0], ["company", "coy"]);
    let idCol = getColIdx(mData[0], ["studentnumber", "idnumber", "studentid"]);
    if(idCol === -1) idCol = 1;
    
    let coyMap = {};
    for(let i=1; i<mData.length; i++) {
        let id = mData[i][idCol].toString().trim();
        coyMap[id] = coyCol >= 0 ? mData[i][coyCol] : "N/A";
    }
    
    let results = [];
    for(let i=1; i<attData.length; i++) {
        if(attData[i][2] === targetDate) {
            let id = attData[i][0].replace(/'/g, "").trim();
            let coy = coyMap[id] || "Unknown";
            if (targetCoy === "ALL" || coy.toLowerCase() === targetCoy.toLowerCase()) {
                results.push({ id: id, name: attData[i][1], company: coy, in: attData[i][3], out: attData[i][4], status: attData[i][5] });
            }
        }
    }
    
    // Sort Alphabetically
    results.sort((a,b) => {
        let nameA = a.name.split(" ").pop() || "";
        let nameB = b.name.split(" ").pop() || "";
        return nameA.localeCompare(nameB);
    });
    
    return { success: true, records: results };
  } catch(e) { return { error: true, message: e.message }; }
}
function updateLegacyAccess(targetUser, newStatus, user, pass) {
  checkAuth(user, pass, "Super Admin"); 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(S_ACCOUNTS);
  if (!sheet) throw new Error("Accounts sheet missing.");
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === targetUser) { 
      // Updates Column 5 (Index 4) with the new access status (e.g., "ON" or "OFF")
      sheet.getRange(i + 1, 5).setValue(newStatus); 
      logAction(user, "Legacy Access Update", `Changed Legacy DB access of ${targetUser} to ${newStatus}`);
      return "Legacy access updated."; 
    }
  }
  return "User not found.";
}
