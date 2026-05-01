// ============================================================
// QA & D-SAT SYSTEM — Code.gs (Apps Script Backend)
// Built clean from scratch. Do not mix with prior versions.
// Sheet: https://docs.google.com/spreadsheets/d/1i7In8QeQ4s0BIujCGj_Ct0A3x-63YQovnYKxATo-poU
// ============================================================

var SHEET_ID = "1i7In8QeQ4s0BIujCGj_Ct0A3x-63YQovnYKxATo-poU";

var TABS = {
  agents:             "Agents",
  interactionReasons: "Interaction Reason",
  issueTypes:         "Agent Issue Type",
  qaTransactions:     "QATransactions",
  dsatImports:        "DSATImports",
  dsatValidations:    "DSATValidations",
  settings:           "AppSettings",
  users:              "AppUsers"
};

var USER_HEADERS = ["user_id","name","email","password","role","active","created_at"];

// QA Parameters — derived from Zainab scorecard. Hardcoded, NOT in a sheet.
var QA_PARAMETERS = [
  { id:"p1",  name:"Engaged Greeting & Confident Handling",                              max:4,  options:[0,2,4],   type:"regular", criticality:"Customer-Critical" },
  { id:"p2",  name:"Understood issue & asked relevant questions",                        max:20, options:[0,10,20], type:"regular", criticality:"Business-Supportive" },
  { id:"p3",  name:"Used tools/systems correctly & efficiently for investigation",       max:10, options:[0,5,10],  type:"regular", criticality:"Business-Supportive" },
  { id:"p4",  name:"Set expectations and going extra mile when required",                max:10, options:[0,5,10],  type:"regular", criticality:"Customer-Critical" },
  { id:"p5",  name:"Provide accurate/up-to-date, complete answer",                      max:20, options:[0,10,20], type:"regular", criticality:"Business-Critical" },
  { id:"p6",  name:"Professional communication & empathy during the transaction",        max:10, options:[0,5,10],  type:"regular", criticality:"Customer-Critical" },
  { id:"p7",  name:"Clear answer and guidance",                                         max:5,  options:[0,3,5],   type:"regular", criticality:"Customer-Critical" },
  { id:"p8",  name:"Reduced customer's time and effort (Dead air, hold, late response)",max:7,  options:[0,4,7],   type:"regular", criticality:"Customer-Critical" },
  { id:"p9",  name:"Proper grammar, communication format, tone clarity and smiley",     max:5,  options:[0,3,5],   type:"regular", criticality:"Business-Supportive" },
  { id:"p10", name:"Professional closing and close the call in a positive tone",        max:4,  options:[0,2,4],   type:"regular", criticality:"Customer-Critical" },
  { id:"p11", name:"Attached Order / Slack Reference for documentation",                max:5,  options:[0,5],     type:"binary",  criticality:"Business-Supportive" },
  { id:"f1",  name:"Gave incorrect information or didn't take promised action",         max:1,  options:[0,1],     type:"fatal",   criticality:"Business-Critical" },
  { id:"f2",  name:"Failed to verify customer identity",                                max:1,  options:[0,1],     type:"fatal",   criticality:"Business-Critical" },
  { id:"f3",  name:"Call back avoidance / Survey avoidance (calls only)",               max:1,  options:[0,1],     type:"fatal",   criticality:"Business-Critical" },
  { id:"f4",  name:"Negative attitude towards Platinumlist/products",                   max:1,  options:[0,1],     type:"fatal",   criticality:"Business-Critical" }
];

var QA_HEADERS = [
  "qa_id","agent_name","date","month","year","week_no","transaction_no",
  "channel","ticket_id","interaction_reason","ratings_json","notes_json",
  "total_without_fatal","final_score","rating_label","has_ko","saved_by","saved_at",
  "transaction_type"
];

var SETTINGS_HEADERS = ["key", "value", "updated_at", "updated_by"];

var DSAT_IMPORT_HEADERS = [
  "import_id","batch_id","date","month","year","agent_name","chat_id",
  "rating","transaction_type","imported_at"
];

var DSAT_VALIDATION_HEADERS = [
  "dsat_id","import_id","date","month","year","agent_name","chat_id","rating",
  "transaction_type","responsible","validation_status","agent_issue_type",
  "actionable_steps","additional_notes","done_by","coached_by","validated_by","validated_at"
];

// ============================================================
// ENTRY POINTS
// ============================================================

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("QA & D-SAT System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called from the frontend via google.script.run
// This is the correct approach — no CORS issues
function handleClientRequest(jsonStr) {
  try {
    var data    = JSON.parse(jsonStr);
    var action  = data.action;
    var payload = data.payload || {};
    var result;

    if      (action === "getAgents")              result = getAgents();
    else if (action === "getInteractionReasons")  result = getInteractionReasons();
    else if (action === "getIssueTypes")          result = getIssueTypes();
    else if (action === "getParameters")          result = QA_PARAMETERS;
    else if (action === "getQATransactions")      result = getQATransactions(payload);
    else if (action === "getDSATImports")         result = getDSATImports(payload);
    else if (action === "getDSATValidations")     result = getDSATValidations(payload);
    else if (action === "saveQATransaction")      result = saveQATransaction(payload);
    else if (action === "importDSAT")             result = importDSAT(payload);
    else if (action === "validateDSAT")           result = validateDSAT(payload);
    else if (action === "getAgentMonthlySummary") result = getAgentMonthlySummary(payload);
    else if (action === "getManagerDashboard")    result = getManagerDashboard(payload);
    else if (action === "generateCoachingEmail")  result = generateCoachingEmail(payload);
    else if (action === "setupSheet")             result = setupSheet();
    else if (action === "rebuildDSATImports")     result = rebuildDSATImports();
    else if (action === "rebuildDSATValidations") result = rebuildDSATValidations();
    else if (action === "getSettings")            result = getSettings();
    else if (action === "saveSettings")           result = saveSettings(payload);
    else if (action === "updateDSATValidation")   result = updateDSATValidation(payload);
    else if (action === "updateQATransaction")    result = updateQATransaction(payload);
    else if (action === "deleteQATransaction")    result = deleteQATransaction(payload);
    else if (action === "exportData")             result = exportData(payload);
    else if (action === "saveCsatData")           result = saveCsatData(payload);
    else if (action === "getCsatData")            result = getCsatData(payload);
    else if (action === "deleteDSATValidation")   result = deleteDSATValidation(payload);
    else if (action === "login")                  result = loginUser(payload);
    else if (action === "getUsers")               result = getUsers();
    else if (action === "addUser")                result = addUser(payload);
    else if (action === "deleteUser")             result = deleteUser(payload);
    else if (action === "changePassword")         result = changePassword(payload);
    else if (action === "saveAgent")              result = saveAgent(payload);
    else if (action === "syncAgentTransactions")  result = syncAgentTransactions(payload);
    else if (action === "createAgentMonthTab")       result = createAgentMonthTab(payload);
    else if (action === "getIntercomConversation")   result = getIntercomConversation(payload);
    else if (action === "getIntercomTags")           result = getIntercomTags();
    else if (action === "getIntercomAdmins")         result = getIntercomAdmins();
    else if (action === "tagIntercomConversation")   result = tagIntercomConversation(payload);
    else if (action === "testIntercomNoteWebApp")    result = testIntercomNoteWebApp(payload);
    else if (action === "analyzeConversation")        result = analyzeConversation(payload);
    else if (action === "getOpenAIKey")               result = getOpenAIKey();
    else if (action === "analyzeCallTranscript")      result = analyzeCallTranscript(payload);
    else if (action === "transcribeAudio")            result = transcribeAudio(payload);
    else if (action === "uploadAudioChunk")           result = uploadAudioChunk(payload);
    else if (action === "finalizeAudioTranscription") result = finalizeAudioTranscription(payload);
    else if (action === "createDriveUploadSession")   result = createDriveUploadSession(payload);
    else if (action === "transcribeFromDrive")        result = transcribeFromDrive(payload);
    else throw new Error("Unknown action: " + action);

    // google.script.run cannot serialize Date objects inside nested objects/arrays.
    // Serialize through JSON to convert all Dates to ISO strings and drop undefined values.
    var safe = JSON.parse(JSON.stringify(result === undefined ? null : result));
    return { success: true, data: safe };

  } catch (err) {
    // Capture error as string — err may be a plain string throw in Apps Script
    var msg = (err && (err.message || String(err))) || "Server error (no message)";
    return { success: false, error: msg };
  }
}

function doPost(e) {
  try {
    var data    = JSON.parse(e.postData.contents);
    var action  = data.action;
    var payload = data.payload || {};
    var result;

    if      (action === "getAgents")              result = getAgents();
    else if (action === "getInteractionReasons")  result = getInteractionReasons();
    else if (action === "getIssueTypes")          result = getIssueTypes();
    else if (action === "getParameters")          result = QA_PARAMETERS;
    else if (action === "getQATransactions")      result = getQATransactions(payload);
    else if (action === "getDSATImports")         result = getDSATImports(payload);
    else if (action === "getDSATValidations")     result = getDSATValidations(payload);
    else if (action === "saveQATransaction")      result = saveQATransaction(payload);
    else if (action === "importDSAT")             result = importDSAT(payload);
    else if (action === "validateDSAT")           result = validateDSAT(payload);
    else if (action === "getAgentMonthlySummary") result = getAgentMonthlySummary(payload);
    else if (action === "getManagerDashboard")    result = getManagerDashboard(payload);
    else if (action === "generateCoachingEmail")  result = generateCoachingEmail(payload);
    else if (action === "setupSheet")             result = setupSheet();
    else if (action === "rebuildDSATImports")     result = rebuildDSATImports();
    else if (action === "rebuildDSATValidations") result = rebuildDSATValidations();
    else if (action === "getSettings")            result = getSettings();
    else if (action === "saveSettings")           result = saveSettings(payload);
    else if (action === "updateDSATValidation")   result = updateDSATValidation(payload);
    else if (action === "updateQATransaction")    result = updateQATransaction(payload);
    else if (action === "deleteQATransaction")    result = deleteQATransaction(payload);
    else if (action === "exportData")             result = exportData(payload);
    else if (action === "saveCsatData")           result = saveCsatData(payload);
    else if (action === "getCsatData")            result = getCsatData(payload);
    else if (action === "deleteDSATValidation")   result = deleteDSATValidation(payload);
    else if (action === "login")                  result = loginUser(payload);
    else if (action === "getUsers")               result = getUsers();
    else if (action === "addUser")                result = addUser(payload);
    else if (action === "deleteUser")             result = deleteUser(payload);
    else if (action === "changePassword")         result = changePassword(payload);
    else if (action === "saveAgent")              result = saveAgent(payload);
    else if (action === "syncAgentTransactions")  result = syncAgentTransactions(payload);
    else if (action === "createAgentMonthTab")       result = createAgentMonthTab(payload);
    else if (action === "getIntercomConversation")   result = getIntercomConversation(payload);
    else if (action === "getIntercomTags")           result = getIntercomTags();
    else if (action === "getIntercomAdmins")         result = getIntercomAdmins();
    else if (action === "tagIntercomConversation")   result = tagIntercomConversation(payload);
    else if (action === "testIntercomNoteWebApp")    result = testIntercomNoteWebApp(payload);
    else if (action === "analyzeConversation")        result = analyzeConversation(payload);
    else if (action === "getOpenAIKey")               result = getOpenAIKey();
    else if (action === "analyzeCallTranscript")      result = analyzeCallTranscript(payload);
    else if (action === "transcribeAudio")            result = transcribeAudio(payload);
    else if (action === "uploadAudioChunk")           result = uploadAudioChunk(payload);
    else if (action === "finalizeAudioTranscription") result = finalizeAudioTranscription(payload);
    else if (action === "createDriveUploadSession")   result = createDriveUploadSession(payload);
    else if (action === "transcribeFromDrive")        result = transcribeFromDrive(payload);
    else throw new Error("Unknown action: " + action);

    var safe = JSON.parse(JSON.stringify(result === undefined ? null : result));
    return jsonOut({ success: true, data: safe });
  } catch (err) {
    var msg = (err && (err.message || String(err))) || "Server error (no message)";
    return jsonOut({ success: false, error: msg });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// SETUP — Run once manually after pasting
// ============================================================

function setupSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var created = [];

  function ensure(name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d9e1f2");
      sheet.setFrozenRows(1);
      created.push(name);
    }
    return sheet;
  }

  ensure(TABS.qaTransactions,  QA_HEADERS);
  ensure(TABS.dsatImports,     DSAT_IMPORT_HEADERS);
  ensure(TABS.dsatValidations, DSAT_VALIDATION_HEADERS);
  ensure(TABS.settings,        SETTINGS_HEADERS);
  ensure(TABS.users,           USER_HEADERS);

  return { message: "Setup complete.", created: created };
}

// Clears the DSATImports tab and resets headers to the correct format.
// Use this when the tab has wrong/old headers. All existing import rows will be erased.
// After running, re-import your CSV from the D-SAT Import tab.
function rebuildDSATImports() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.dsatImports);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.dsatImports);
  } else {
    sheet.clearContents();
  }
  sheet.getRange(1, 1, 1, DSAT_IMPORT_HEADERS.length).setValues([DSAT_IMPORT_HEADERS]);
  sheet.getRange(1, 1, 1, DSAT_IMPORT_HEADERS.length).setFontWeight("bold").setBackground("#fce5cd");
  sheet.setFrozenRows(1);
  return { message: "DSATImports tab cleared and headers fixed. Please re-import your CSV from the D-SAT Import tab." };
}

// Clears the DSATValidations tab and resets headers. Erases all validation records.
function rebuildDSATValidations() {
  var ss    = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.dsatValidations);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.dsatValidations);
  } else {
    sheet.clearContents();
  }
  sheet.getRange(1, 1, 1, DSAT_VALIDATION_HEADERS.length).setValues([DSAT_VALIDATION_HEADERS]);
  sheet.getRange(1, 1, 1, DSAT_VALIDATION_HEADERS.length).setFontWeight("bold").setBackground("#d9ead3");
  sheet.setFrozenRows(1);
  return { message: "DSATValidations tab cleared and headers fixed." };
}

// ============================================================
// HELPERS
// ============================================================

function getSheetData(tabName) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  // Always use canonical header arrays for our tabs so column position is reliable
  // even when the sheet's own header row has old/wrong values.
  var headers;
  if      (tabName === TABS.qaTransactions)  headers = QA_HEADERS;
  else if (tabName === TABS.dsatImports)     headers = DSAT_IMPORT_HEADERS;
  else if (tabName === TABS.dsatValidations) headers = DSAT_VALIDATION_HEADERS;
  else                                       headers = values[0].map(String);

  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var row = {};
    for (var j = 0; j < headers.length; j++) {
      var v = values[i][j];
      // Convert Date objects to ISO string — google.script.run cannot serialize Dates
      // inside nested objects. Convert everything to a safe primitive.
      if (v instanceof Date) {
        // Use localDateStr() — avoids UTC off-by-one in UTC+ timezones (e.g. Dubai UTC+4)
        row[headers[j]] = localDateStr(v);
      } else if (v === null || v === undefined) {
        row[headers[j]] = "";
      } else if (headers[j] === "chat_id" || headers[j] === "import_id" || headers[j] === "dsat_id") {
        // Force ID columns to string — Sheets can convert long numeric IDs to floats
        row[headers[j]] = String(v).trim();
      } else {
        row[headers[j]] = v;
      }
    }
    rows.push(row);
  }
  return rows;
}

function appendRow(tabName, headers, rowObj) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(tabName);
  if (!sheet) throw new Error("Tab not found: " + tabName);
  var rowArr = headers.map(function(h) {
    var v = rowObj[h];
    if (v === undefined || v === null) return "";
    // Force ID columns to plain text so Sheets doesn't convert long numbers to floats
    if (h === "chat_id" || h === "import_id" || h === "dsat_id" || h === "batch_id") return String(v).trim();
    return v;
  });
  var newRow = sheet.appendRow(rowArr);
  // Set chat_id column to plain text format to prevent numeric conversion
  var chatIdCol = headers.indexOf("chat_id");
  if (chatIdCol !== -1) {
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, chatIdCol + 1).setNumberFormat("@");
  }
}

function generateId(prefix) {
  return prefix + "_" + Date.now() + "_" + Math.random().toString(36).slice(2, 6);
}

function getMonthName(dateStr) {
  var months = ["January","February","March","April","May","June",
                "July","August","September","October","November","December"];
  var d = new Date(dateStr);
  return isNaN(d.getTime()) ? "" : months[d.getMonth()];
}

function getWeekOfMonth(dateStr) {
  var d = new Date(dateStr);
  return Math.ceil(d.getDate() / 7);
}

function nowISO() {
  return new Date().toISOString();
}

// Returns YYYY-MM-DD using LOCAL timezone (avoids UTC off-by-one in UTC+ zones like Dubai)
function localDateStr(d) {
  if (!d || isNaN(d.getTime())) return "";
  var y  = d.getFullYear();
  var m  = d.getMonth() + 1;
  var dy = d.getDate();
  return y + "-" + (m  < 10 ? "0"+m  : m)  + "-" + (dy < 10 ? "0"+dy : dy);
}

// ============================================================
// SCORE CALCULATION
// ============================================================

function calculateScores(ratingsObj) {
  var totalWithoutFatal = 0;
  var hasKo = false;

  QA_PARAMETERS.forEach(function(p) {
    var score = Number(ratingsObj[p.id]);
    if (isNaN(score)) score = 0;
    if (p.type === "regular" || p.type === "binary") {
      totalWithoutFatal += score;
    } else if (p.type === "fatal") {
      if (score === 0) hasKo = true;
    }
  });

  totalWithoutFatal = Math.round(totalWithoutFatal * 100) / 100;
  var finalScore = hasKo ? 0 : totalWithoutFatal;

  var ratingLabel;
  if (hasKo)                ratingLabel = "K.O";
  else if (finalScore >= 80) ratingLabel = "Nailed It!";
  else if (finalScore >= 70) ratingLabel = "Almost There!";
  else                       ratingLabel = "Not Quite There!";

  return {
    total_without_fatal: totalWithoutFatal,
    final_score:         finalScore,
    has_ko:              hasKo,
    rating_label:        ratingLabel
  };
}

// ============================================================
// WEAK AREA DETECTION
// ============================================================

function detectWeakAreas(transactions) {
  var paramScores = {};
  QA_PARAMETERS.filter(function(p){ return p.type !== "fatal"; }).forEach(function(p){
    paramScores[p.id] = [];
  });

  transactions.forEach(function(tx) {
    try {
      var ratings = JSON.parse(tx.ratings_json || "{}");
      Object.keys(paramScores).forEach(function(pid) {
        var s = Number(ratings[pid]);
        paramScores[pid].push(isNaN(s) ? 0 : s);
      });
    } catch(e) {}
  });

  var weak = [];
  QA_PARAMETERS.filter(function(p){ return p.type !== "fatal"; }).forEach(function(p) {
    var scores = paramScores[p.id];
    if (!scores || scores.length === 0) return;
    var avg = scores.reduce(function(a,b){ return a+b; }, 0) / scores.length;
    var threshold = p.max * 0.75;
    var lowCount = scores.filter(function(s){ return s < threshold; }).length;
    if (avg < threshold) {
      weak.push({
        parameter_id:   p.id,
        parameter_name: p.name,
        criticality:    p.criticality,
        max:            p.max,
        avg_score:      Math.round(avg * 100) / 100,
        avg_pct:        Math.round((avg / p.max) * 100),
        low_in:         lowCount,
        of_total:       scores.length,
        repetition_pct: Math.round((lowCount / scores.length) * 100)
      });
    }
  });

  return weak.sort(function(a, b){ return a.avg_pct - b.avg_pct; });
}

// ============================================================
// DATA READERS
// ============================================================

function getAgents() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.agents);
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  // Normalize headers: trim + lowercase for flexible matching
  var rawHeaders = values[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var rowObj = {};
    for (var j = 0; j < rawHeaders.length; j++) {
      var v = values[i][j];
      rowObj[rawHeaders[j]] = (v instanceof Date) ? localDateStr(v) : (v === null || v === undefined ? "" : String(v));
    }
    // Map to canonical names for the frontend
    var agent = {
      "Agent name":        rowObj["agent name"]        || rowObj["name"]  || "",
      "Email":             rowObj["email"]              || "",
      "online sheet link": rowObj["online sheet link"]  || rowObj["sheet link"] || rowObj["link"] || "",
      "active":            rowObj["active"]             || "TRUE",
      "intercom_admin_id": rowObj["intercom_admin_id"]  || rowObj["intercom admin id"] || ""
    };
    if (agent["Agent name"]) rows.push(agent);
  }
  return rows;
}

function saveAgent(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.agents);
  if (!sheet) throw new Error("Agents tab not found");
  var name       = String(payload.name       || "").trim();
  var email      = String(payload.email      || "").trim();
  var link       = String(payload.link       || "").trim();
  var intercomId = String(payload.intercom_admin_id || "").trim();
  if (!name) throw new Error("Agent name is required");

  var values = sheet.getDataRange().getValues();
  var rawHeaders = values[0].map(function(h){ return String(h).trim().toLowerCase(); });
  var nameCol       = rawHeaders.indexOf("agent name");
  var emailCol      = rawHeaders.indexOf("email");
  var linkCol       = rawHeaders.indexOf("online sheet link");
  var intercomCol   = rawHeaders.indexOf("intercom_admin_id");
  if (nameCol === -1) throw new Error("'Agent name' column not found in Agents tab");

  // If the intercom_admin_id column doesn't exist yet, add it
  if (intercomCol === -1 && intercomId) {
    var headerRow = sheet.getRange(1, 1, 1, rawHeaders.length + 1);
    sheet.getRange(1, rawHeaders.length + 1).setValue("intercom_admin_id");
    intercomCol = rawHeaders.length;
    rawHeaders.push("intercom_admin_id");
  }

  // Update existing row
  if (payload.original_name) {
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][nameCol]).trim() === String(payload.original_name).trim()) {
        sheet.getRange(i+1, nameCol+1).setValue(name);
        if (emailCol    !== -1) sheet.getRange(i+1, emailCol+1).setValue(email);
        if (linkCol     !== -1) sheet.getRange(i+1, linkCol+1).setValue(link);
        if (intercomCol !== -1) sheet.getRange(i+1, intercomCol+1).setValue(intercomId);
        return { saved: true, name: name };
      }
    }
    throw new Error("Agent not found: " + payload.original_name);
  }

  // Add new row
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][nameCol]).trim().toLowerCase() === name.toLowerCase()) throw new Error("Agent already exists: " + name);
  }
  var newRow = [];
  for (var j = 0; j < rawHeaders.length; j++) newRow.push("");
  if (nameCol       !== -1) newRow[nameCol]       = name;
  if (emailCol      !== -1) newRow[emailCol]      = email;
  if (linkCol       !== -1) newRow[linkCol]       = link;
  if (intercomCol   !== -1) newRow[intercomCol]   = intercomId;
  sheet.appendRow(newRow);
  return { saved: true, name: name, added: true };
}

function getInteractionReasons() {
  return getSheetData(TABS.interactionReasons);
}

function getIssueTypes() {
  return getSheetData(TABS.issueTypes);
}

function getQATransactions(filters) {
  var rows = getSheetData(TABS.qaTransactions);
  if (!filters) return rows;
  return rows.filter(function(r) {
    if (filters.agentName && r.agent_name !== filters.agentName) return false;
    if (filters.month && r.month !== filters.month) return false;
    if (filters.year  && String(r.year) !== String(filters.year)) return false;
    return true;
  });
}

function getDSATImports(filters) {
  var rows = getSheetData(TABS.dsatImports);
  if (!filters) return rows;
  return rows.filter(function(r) {
    if (filters.agentName && r.agent_name !== filters.agentName) return false;
    // If the stored month/year is empty (old import with no date), always include it
    if (filters.month && r.month && r.month !== filters.month) return false;
    if (filters.year  && r.year  && String(r.year) !== String(filters.year)) return false;
    return true;
  });
}

function getDSATValidations(filters) {
  var rows = getSheetData(TABS.dsatValidations);
  if (!filters) return rows;
  return rows.filter(function(r) {
    if (filters.agentName && r.agent_name !== filters.agentName) return false;
    // If the stored month/year is empty (old record with no date), always include it
    if (filters.month && r.month && r.month !== filters.month) return false;
    if (filters.year  && r.year  && String(r.year) !== String(filters.year)) return false;
    if (filters.status && r.validation_status !== filters.status) return false;
    return true;
  });
}

// ============================================================
// SAVE QA TRANSACTION
// ============================================================

function saveQATransaction(payload) {
  var agentName = payload.agent_name;
  var dateStr   = payload.date;
  var channel   = payload.channel;
  var ticketId  = payload.ticket_id  || "";
  var reason    = payload.interaction_reason;
  var ratings   = payload.ratings    || {};
  var notes     = payload.notes      || {};
  var savedBy   = payload.saved_by   || "";

  if (!agentName) throw new Error("agent_name is required");
  if (!dateStr)   throw new Error("date is required");

  var d         = new Date(dateStr);
  var monthName = getMonthName(dateStr);
  var year      = d.getFullYear();
  var weekNo    = getWeekOfMonth(dateStr);

  var existing = getQATransactions({ agentName: agentName, month: monthName, year: String(year) });
  var txNo = existing.length + 1;
  var maxTx = parseInt((getSettings()).max_transactions) || 8;
  if (txNo > maxTx) {
    throw new Error(agentName + " already has " + maxTx + " transactions for " + monthName + " " + year + ". Maximum is T" + maxTx + ".");
  }

  var scores = calculateScores(ratings);
  var qaId   = generateId("QA");

  var row = {
    qa_id:               qaId,
    agent_name:          agentName,
    date:                dateStr,
    month:               monthName,
    year:                year,
    week_no:             weekNo,
    transaction_no:      txNo,
    channel:             channel,
    ticket_id:           ticketId,
    interaction_reason:  reason,
    ratings_json:        JSON.stringify(ratings),
    notes_json:          JSON.stringify(notes),
    total_without_fatal: scores.total_without_fatal,
    final_score:         scores.final_score,
    rating_label:        scores.rating_label,
    has_ko:              scores.has_ko,
    saved_by:            savedBy,
    saved_at:            nowISO(),
    transaction_type:    payload.transaction_type || ""
  };

  appendRow(TABS.qaTransactions, QA_HEADERS, row);

  // Write to agent's personal scorecard sheet (best-effort, does not block save)
  try { writeTransactionToAgentSheet(row); } catch(e) { /* non-fatal */ }

  return {
    qa_id:          qaId,
    transaction_no: txNo,
    final_score:    scores.final_score,
    rating_label:   scores.rating_label,
    has_ko:         scores.has_ko
  };
}

// ============================================================
// AGENT SCORECARD SYNC
// ============================================================

function extractSheetIdFromUrl(url) {
  var m = String(url).match(/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/);
  return m ? m[1] : null;
}

function scoreToSheetRating(score, param) {
  var s = Number(score);
  if (param.type === "fatal")  return s === 0 ? "Misstep"  : "Nailed It!";
  if (param.type === "binary") return s === 0 ? "No"       : "Yes";
  if (s === 0)                 return "N/A";
  if (s >= param.max)          return "Excellent";
  return "Good";
}

// Write a single QA transaction row to the agent's personal Google Sheet
function writeTransactionToAgentSheet(tx) {
  var agents    = getAgents();
  var agentLink = "";
  for (var i = 0; i < agents.length; i++) {
    if ((agents[i]["Agent name"] || "").trim() === (tx.agent_name || "").trim()) {
      agentLink = agents[i]["online sheet link"] || "";
      break;
    }
  }
  if (!agentLink) return { skipped: true, reason: "No sheet link for " + tx.agent_name };

  var sheetId = extractSheetIdFromUrl(agentLink);
  if (!sheetId) return { skipped: true, reason: "Invalid sheet URL" };

  var ss;
  try { ss = SpreadsheetApp.openById(sheetId); }
  catch(e) { return { skipped: true, reason: "Cannot open: " + e.message }; }

  // Find the month tab (try exact name and year-suffixed variants)
  var month  = tx.month || getMonthName(tx.date);
  var year   = String(tx.year || new Date(tx.date).getFullYear());
  var sheet  = ss.getSheetByName(month);
  if (!sheet) sheet = ss.getSheetByName(month + " " + year);
  if (!sheet) {
    var tabs = ss.getSheets().map(function(s){ return s.getName(); }).join(", ");
    return { skipped: true, reason: "No tab '" + month + "' in agent sheet. Tabs: " + tabs };
  }

  var txNo   = parseInt(tx.transaction_no);
  var allVals = sheet.getDataRange().getValues();
  var numRows = allVals.length;
  var numCols = allVals[0] ? allVals[0].length : 10;

  // Find the row with "Transaction N" text
  var txRow = -1;
  for (var r = 0; r < numRows; r++) {
    var joined = allVals[r].join(" ").replace(/\s+/g," ").trim().toLowerCase();
    if (joined === "transaction " + txNo || joined.indexOf("transaction " + txNo) !== -1) {
      txRow = r;
      break;
    }
  }
  if (txRow === -1) return { skipped: true, reason: "Transaction " + txNo + " block not found in " + month + " tab" };

  var ratings = {};
  var notes   = {};
  try { ratings = JSON.parse(tx.ratings_json || "{}"); } catch(e){}
  try { notes   = JSON.parse(tx.notes_json   || "{}"); } catch(e){}

  // ── 1. Write metadata: scan block for Date/Channel/Ticket/Interaction labels ──
  var labelSearchEnd = Math.min(txRow + 25, numRows);
  for (var r = txRow + 1; r < labelSearchEnd; r++) {
    for (var c = 0; c < numCols; c++) {
      var cell = String(allVals[r][c] || "").trim().toLowerCase();
      var valueColOffset = 1; // write to c+1 by default
      // skip header rows
      if (cell === "date") {
        sheet.getRange(r + 1, c + 1 + valueColOffset).setValue(tx.date || "");
        break;
      } else if (cell === "channel") {
        sheet.getRange(r + 1, c + 1 + valueColOffset).setValue(tx.channel || "");
        break;
      } else if (cell === "ticket id / phone" || cell === "ticket id" || cell.indexOf("ticket") !== -1) {
        sheet.getRange(r + 1, c + 1 + valueColOffset).setValue(tx.ticket_id || "");
        break;
      } else if (cell === "interaction reason" || cell === "interaction" || cell.indexOf("reason") !== -1) {
        sheet.getRange(r + 1, c + 1 + valueColOffset).setValue(tx.interaction_reason || "");
        break;
      }
    }
  }

  // ── 2. Find parameters header row (contains "rating" column) ──
  var paramHeaderRow = -1;
  var ratingCol      = -1;
  var noteCol        = -1;
  for (var r = txRow + 1; r < labelSearchEnd; r++) {
    for (var c = 0; c < numCols; c++) {
      var cell = String(allVals[r][c] || "").trim().toLowerCase();
      if (cell === "rating") { paramHeaderRow = r; ratingCol = c; }
      if (cell === "notes" || cell === "note") { noteCol = c; }
    }
    if (paramHeaderRow !== -1) break;
  }
  if (paramHeaderRow === -1) return { written: "metadata only", reason: "Rating column header not found" };

  // Score column is one to the right of Rating column
  var scoreCol = ratingCol + 1;

  // ── 3. Write each parameter rating AND score by matching the parameter name ──
  var paramSearchEnd = Math.min(paramHeaderRow + 25, numRows);
  QA_PARAMETERS.forEach(function(param) {
    var shortName = param.name.toLowerCase().replace(/[^a-z0-9]/g,"").slice(0, 18);
    for (var r = paramHeaderRow + 1; r < paramSearchEnd; r++) {
      var rowText = allVals[r].join(" ").toLowerCase().replace(/[^a-z0-9 ]/g,"");
      if (rowText.replace(/\s+/g,"").indexOf(shortName) !== -1) {
        var score       = ratings[param.id] !== undefined ? Number(ratings[param.id]) : param.max;
        var ratingLabel = scoreToSheetRating(score, param);
        sheet.getRange(r + 1, ratingCol + 1).setValue(ratingLabel);
        // Write numeric score to Score column (next column after Rating)
        if (param.type !== "fatal") {
          sheet.getRange(r + 1, scoreCol + 1).setValue(score);
        }
        if (noteCol !== -1 && notes[param.id]) {
          sheet.getRange(r + 1, noteCol + 1).setValue(notes[param.id]);
        }
        break;
      }
    }
  });

  // ── 4. Write totals — scan for "TOTAL Without Fatal" and "Final Score" rows ──
  var totalSearchEnd = Math.min(paramHeaderRow + 35, numRows);
  for (var r = paramHeaderRow + 1; r < totalSearchEnd; r++) {
    var rowJoined = allVals[r].join(" ").toLowerCase();
    if (rowJoined.indexOf("total without fatal") !== -1 || rowJoined.indexOf("total without") !== -1) {
      sheet.getRange(r + 1, scoreCol + 1).setValue(Number(tx.total_without_fatal || 0));
    } else if (rowJoined.indexOf("final score") !== -1) {
      sheet.getRange(r + 1, scoreCol + 1).setValue(Number(tx.final_score || 0));
    }
  }

  return { written: true, agent: tx.agent_name, transaction: txNo, month: month };
}

// Sync all QA transactions for a given agent+month to their personal sheet
function syncAgentTransactions(payload) {
  var agentName = payload.agentName || payload.agent_name || "";
  var month     = payload.month || "";
  var year      = String(payload.year || "");
  var txs       = getQATransactions({ agentName: agentName, month: month, year: year });
  var results   = [];
  txs.forEach(function(tx) {
    try { results.push(writeTransactionToAgentSheet(tx)); }
    catch(e) { results.push({ error: e.message }); }
  });
  return { synced: results.length, results: results };
}

// Create the month tab with QA template in each agent's personal sheet
function createAgentMonthTab(payload) {
  var month     = String(payload.month || "");
  var year      = String(payload.year  || new Date().getFullYear());
  var agentName = String(payload.agentName || ""); // empty = all agents
  if (!month) throw new Error("month is required");

  var agents  = getAgents();
  var results = [];

  agents.forEach(function(agent) {
    var name = agent["Agent name"] || "";
    if (agentName && name !== agentName) return;
    var link = agent["online sheet link"] || "";
    if (!link) { results.push({ agent: name, skipped: true, reason: "No sheet link" }); return; }

    var sheetId = extractSheetIdFromUrl(link);
    if (!sheetId) { results.push({ agent: name, skipped: true, reason: "Invalid URL" }); return; }

    var ss;
    try { ss = SpreadsheetApp.openById(sheetId); }
    catch(e) { results.push({ agent: name, error: "Cannot open sheet: " + e.message }); return; }

    // Delete existing tab if present (force recreate with fresh template)
    var existing = ss.getSheetByName(month);
    if (existing) {
      try { ss.deleteSheet(existing); } catch(e) { /* ignore */ }
    }

    try {
      var sheet = ss.insertSheet(month);
      buildMonthTemplate(sheet);
      results.push({ agent: name, created: true });
    } catch(e) {
      results.push({ agent: name, error: e.message });
    }
  });

  return { month: month, results: results };
}

function buildMonthTemplate(sheet) {
  var regularParams = QA_PARAMETERS.filter(function(p){ return p.type !== "fatal"; }); // 11
  var fatalParams   = QA_PARAMETERS.filter(function(p){ return p.type === "fatal";  }); // 4

  // Each transaction block layout (row offsets from block start, 0-based):
  // 0  : "Transaction N" header
  // 1  : (empty)
  // 2  : Date      | value
  // 3  : Channel   | value
  // 4  : Ticket ID / Phone | value
  // 5  : Interaction Reason | value
  // 6  : (empty)
  // 7  : Weight | Parameters | Rating | Score | Score Range | Criticality | Notes
  // 8-18: 11 regular parameters
  // 19 : "Serious Slip-Ups and Missteps" header
  // 20-23: 4 fatal parameters
  // 24 : (empty)
  // 25 : TOTAL Without Fatal Errors
  // 26 : Final Score
  // 27 : (empty separator)
  // 28 : (empty separator)
  // Total = 29 rows per transaction

  var BLOCK = 29;
  var NUM_TX = 8;
  var NUM_COLS = 7;
  var totalRows = BLOCK * NUM_TX;

  var vals = [];
  for (var i = 0; i < totalRows; i++) vals.push(["","","","","","",""]);

  for (var t = 1; t <= NUM_TX; t++) {
    var b = (t - 1) * BLOCK;

    vals[b][0]   = "Transaction " + t;

    vals[b+2][0] = "Date";
    vals[b+3][0] = "Channel";
    vals[b+4][0] = "Ticket ID / Phone";
    vals[b+5][0] = "Interaction Reason";

    // Parameter table header — column C (index 2) = "Rating" (used by writeTransactionToAgentSheet)
    vals[b+7] = ["Weight", "Parameters", "Rating", "Score", "Score Range", "Criticality", "Notes"];

    regularParams.forEach(function(p, i) {
      vals[b+8+i] = [p.max, p.name, "", "", p.options.join("-"), p.criticality, ""];
    });

    vals[b+19][0] = ""; vals[b+19][1] = "Serious Slip-Ups and Missteps — Fatal Errors (K.O on any)";

    fatalParams.forEach(function(p, i) {
      vals[b+20+i] = ["", p.name, "", "", "0-1", p.criticality, ""];
    });

    vals[b+25] = ["", "TOTAL Without Fatal Errors", "", "", "", "", ""];
    vals[b+26] = ["", "Final Score", "", "", "", "", ""];
  }

  // Set Score Range column (col 5 = E) to plain text BEFORE writing
  // to prevent Google Sheets from auto-parsing "0-2-4" as a date
  sheet.getRange(1, 5, totalRows, 1).setNumberFormat("@");
  sheet.getRange(1, 1, totalRows, NUM_COLS).setValues(vals);

  // ── Formatting ──
  var brandPurple = "#6B00EE";
  var lightPurple = "#f3eeff";
  var grayHdr     = "#d9d9d9";
  var yellowSect  = "#fff2cc";
  var redLight    = "#fce8e6";
  var greenTot    = "#d9ead3";
  var white       = "#ffffff";

  for (var t = 1; t <= NUM_TX; t++) {
    var b = (t - 1) * BLOCK + 1; // 1-indexed sheet rows

    // Transaction header
    sheet.getRange(b, 1, 1, NUM_COLS)
      .merge().setBackground(grayHdr).setFontWeight("bold")
      .setFontSize(11).setHorizontalAlignment("center");

    // Metadata label cells (col A)
    sheet.getRange(b+2, 1, 4, 1).setFontWeight("bold").setBackground(lightPurple);
    // Metadata value cells (col B) — bordered
    sheet.getRange(b+2, 2, 4, 1).setBackground(white)
      .setBorder(true, true, true, true, null, null, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    // Table header
    sheet.getRange(b+7, 1, 1, NUM_COLS)
      .setBackground(brandPurple).setFontColor("#ffffff").setFontWeight("bold");

    // Regular parameter rows — alternate shading
    for (var i = 0; i < 11; i++) {
      var bg = (i % 2 === 0) ? "#fafafa" : white;
      sheet.getRange(b+8+i, 1, 1, NUM_COLS).setBackground(bg);
    }

    // Fatal section header
    sheet.getRange(b+19, 1, 1, NUM_COLS)
      .merge().setBackground(yellowSect).setFontWeight("bold").setFontStyle("italic");

    // Fatal rows
    sheet.getRange(b+20, 1, 4, NUM_COLS).setBackground(redLight);

    // Totals
    sheet.getRange(b+25, 1, 2, NUM_COLS)
      .setBackground(greenTot).setFontWeight("bold");
  }

  // Column widths
  sheet.setColumnWidth(1, 70);   // Weight / Label
  sheet.setColumnWidth(2, 300);  // Parameters / Value
  sheet.setColumnWidth(3, 110);  // Rating
  sheet.setColumnWidth(4, 70);   // Score
  sheet.setColumnWidth(5, 100);  // Score Range
  sheet.setColumnWidth(6, 130);  // Criticality
  sheet.setColumnWidth(7, 220);  // Notes

  sheet.setFrozenRows(0);
}

// ============================================================
// IMPORT D-SAT
// ============================================================

function importDSAT(payload) {
  var rows    = payload.rows    || [];
  var batchId = payload.batchId || generateId("BATCH");

  var existing = getSheetData(TABS.dsatImports);
  var existingChatIds = {};
  existing.forEach(function(r){ existingChatIds[String(r.chat_id)] = true; });

  var imported = 0;
  var skipped  = 0;
  var now = nowISO();

  rows.forEach(function(r) {
    var rating = Number(r.rating);
    if (rating < 1 || rating > 3) { skipped++; return; }

    var chatId = String(r.chat_id || "").trim();
    if (!chatId) { skipped++; return; }
    if (existingChatIds[chatId]) { skipped++; return; }

    var dateStr   = (r.date || "").trim();
    var d         = new Date(dateStr);
    var monthName = isNaN(d.getTime()) ? "" : getMonthName(dateStr);
    var year      = isNaN(d.getTime()) ? "" : d.getFullYear();

    var row = {
      import_id:        generateId("DIMP"),
      batch_id:         batchId,
      date:             dateStr,
      month:            monthName,
      year:             year,
      agent_name:       (r.agent_name || "").trim(),
      chat_id:          chatId,
      rating:           rating,
      transaction_type: (r.transaction_type || "").trim(),
      imported_at:      now
    };

    appendRow(TABS.dsatImports, DSAT_IMPORT_HEADERS, row);
    existingChatIds[chatId] = true;
    imported++;
  });

  return { imported: imported, skipped: skipped, batch_id: batchId };
}

// ============================================================
// VALIDATE D-SAT
// ============================================================

function validateDSAT(payload) {
  var responsible      = (payload.responsible || "").toLowerCase().trim();
  var validationStatus = (responsible === "agent") ? "valid" : "invalid";

  var row = {
    dsat_id:           generateId("DVAL"),
    import_id:         payload.import_id         || "",
    date:              payload.date              || "",
    month:             payload.month             || "",
    year:              payload.year              || "",
    agent_name:        payload.agent_name        || "",
    chat_id:           payload.chat_id           || "",
    rating:            payload.rating            || "",
    transaction_type:  payload.transaction_type  || "",
    responsible:       payload.responsible       || "",
    validation_status: validationStatus,
    agent_issue_type:  payload.agent_issue_type  || "",
    actionable_steps:  payload.actionable_steps  || "",
    additional_notes:  payload.additional_notes  || "",
    done_by:           payload.done_by           || "",
    coached_by:        payload.coached_by        || "",
    validated_by:      payload.validated_by      || "",
    validated_at:      nowISO()
  };

  appendRow(TABS.dsatValidations, DSAT_VALIDATION_HEADERS, row);

  // Post internal note to Intercom (best-effort — does not block save)
  var noteResult = "skipped";
  try {
    postIntercomNote(row.chat_id, row, validationStatus);
    noteResult = "ok";
  } catch(ne) {
    noteResult = "note_error: " + ne.message;
  }

  return { dsat_id: row.dsat_id, validation_status: validationStatus, note_result: noteResult };
}

// ============================================================
// INTERCOM API
// ============================================================

function callIntercomAPI(method, endpoint, body) {
  var token = (getSettings()).intercom_token || "";
  if (!token) throw new Error("Intercom API token not set. Add it in Settings → Intercom.");
  var options = {
    method: method.toLowerCase(),
    headers: {
      "Authorization":    "Bearer " + token,
      "Accept":           "application/json",
      "Intercom-Version": "2.10"
    },
    muteHttpExceptions: true
  };
  if (body) {
    options.contentType = "application/json";
    options.payload = JSON.stringify(body);
  }
  var res  = UrlFetchApp.fetch("https://api.intercom.io/" + endpoint, options);
  var code = res.getResponseCode();
  var text = res.getContentText();
  if (code >= 400) throw new Error("Intercom " + code + ": " + text.slice(0, 200));
  return JSON.parse(text);
}

function getIntercomConversation(payload) {
  var convId = String(payload.conversation_id || "").trim();
  if (!convId) throw new Error("conversation_id is required");

  var data = callIntercomAPI("GET", "conversations/" + convId + "?display_as=plaintext");

  // Customer info
  var customer = { name: "", email: "", phone: "", location: "" };
  if (data.contacts && data.contacts.contacts && data.contacts.contacts.length) {
    try {
      var c = callIntercomAPI("GET", "contacts/" + data.contacts.contacts[0].id);
      customer.name     = c.name  || "";
      customer.email    = c.email || "";
      customer.phone    = c.phone || "";
      customer.location = c.location ? [(c.location.city||""),(c.location.country||"")].filter(Boolean).join(", ") : "";
    } catch(e) {}
  }

  // Source info
  var sourceUrl  = (data.source && data.source.url)     ? data.source.url     : "";
  var browser    = (data.source && data.source.browser) ? data.source.browser : "";
  var os         = (data.source && data.source.os)      ? data.source.os      : "";
  var sourceType = (data.source && data.source.type)    ? data.source.type    : "";

  // Messages
  var messages = [];
  if (data.source && data.source.body) {
    messages.push({
      type: "customer", author: customer.name || "Customer",
      body: data.source.body, created_at: data.created_at || 0, part_type: "comment"
    });
  }
  if (data.conversation_parts && data.conversation_parts.conversation_parts) {
    data.conversation_parts.conversation_parts.forEach(function(p) {
      if (!p.body && p.part_type !== "note") return;
      var aType = p.author ? p.author.type : "";
      messages.push({
        type:       p.part_type === "note" ? "note" : (aType === "admin" || aType === "bot" ? "agent" : "customer"),
        author:     p.author ? (p.author.name || p.author.type) : "Unknown",
        body:       p.body || "",
        created_at: p.created_at || 0,
        part_type:  p.part_type || "comment"
      });
    });
  }

  // Tags
  var tags = [];
  if (data.tags && data.tags.tags) tags = data.tags.tags.map(function(t){ return t.name; });

  return {
    id: convId, customer: customer, source_url: sourceUrl,
    browser: browser, os: os, source_type: sourceType,
    messages: messages, tags: tags, state: data.state || ""
  };
}

function getIntercomTags() {
  var data = callIntercomAPI("GET", "tags");
  return (data.data || []).map(function(t){ return { id: t.id, name: t.name }; });
}

function testIntercomNoteWebApp(payload) {
  var s = getSettings();
  var adminId = String(s.intercom_admin_id || "").trim();
  var log = [];
  log.push("token_set: " + (s.intercom_token ? "yes" : "no"));
  log.push("admin_id_from_settings: [" + adminId + "]");
  if (!adminId) {
    try {
      var admins = callIntercomAPI("GET", "admins");
      var list = admins.admins || admins.data || [];
      adminId = list.length > 0 ? String(list[0].id) : "";
      log.push("admin_id_from_api: [" + adminId + "]");
    } catch(e) { log.push("admins_error: " + e.message); }
  }
  var convId = String(payload.conversation_id || "").trim();
  if (!convId) {
    try {
      var convs = callIntercomAPI("GET", "conversations?order=desc&per_page=1");
      convId = (convs.conversations||[])[0] ? String(convs.conversations[0].id) : "";
      log.push("test_conv_id: " + convId);
    } catch(e) { log.push("conv_error: " + e.message); }
  }
  if (adminId && convId) {
    try {
      var body = { message_type: "note", type: "admin", admin_id: adminId, body: "<b>[QA Test Note — safe to delete]</b>" };
      log.push("posting_body: " + JSON.stringify(body));
      callIntercomAPI("POST", "conversations/" + convId + "/reply", body);
      log.push("result: SUCCESS");
    } catch(e) { log.push("result: FAILED — " + e.message); }
  } else {
    log.push("result: SKIPPED — missing adminId or convId");
  }
  return { log: log };
}

function testIntercomNote() {
  var s = getSettings();
  Logger.log("=== INTERCOM DEBUG ===");
  Logger.log("Token set: " + (s.intercom_token ? "YES (" + s.intercom_token.slice(0,10) + "...)" : "NO"));
  Logger.log("Admin ID from settings: [" + (s.intercom_admin_id || "") + "]");
  try {
    var admins = callIntercomAPI("GET", "admins");
    var list = admins.admins || admins.data || [];
    Logger.log("Admin list length: " + list.length);
    // Check if saved admin ID exists in list
    var found = list.filter(function(a){ return String(a.id) === String(s.intercom_admin_id); });
    Logger.log("Saved admin ID in list: " + (found.length > 0 ? "YES — " + found[0].name : "NO — not found"));
  } catch(e) {
    Logger.log("Error calling /admins: " + e.message);
  }
  // Test actual note posting using first conversation
  try {
    var convs = callIntercomAPI("GET", "conversations?order=desc&sort=updated_at&per_page=1");
    var testConvId = (convs.conversations||[])[0] ? convs.conversations[0].id : null;
    Logger.log("Test conversation ID: " + testConvId);
    if (testConvId && s.intercom_admin_id) {
      var result = callIntercomAPI("POST", "conversations/" + testConvId + "/reply", {
        message_type: "note",
        type: "admin",
        admin_id: String(s.intercom_admin_id),
        body: "<b>[QA Tool Test Note — safe to delete]</b>"
      });
      Logger.log("Note POST result: " + JSON.stringify(result).slice(0, 200));
    }
  } catch(e) {
    Logger.log("Note POST error: " + e.message);
  }
}

// ============================================================
// AI CONVERSATION ANALYSIS (Claude API)
// ============================================================

function analyzeConversation(payload) {
  var s = getSettings();
  var apiKey = String(s.claude_api_key || "").trim();
  if (!apiKey) throw new Error("Claude API key not set. Add it in Settings → Claude API Key.");

  var messages = payload.messages || [];
  var agentName = payload.agent_name || "Agent";

  // Build plain-text conversation
  var convText = messages.map(function(m) {
    var body = String(m.body || m.text || "").replace(/<[^>]+>/g, "").trim();
    if (!body) return null;
    var role = m.type === "agent" ? ("Agent ("+agentName+")") : m.type === "note" ? "Internal Note" : "Customer";
    return role + ": " + body;
  }).filter(Boolean).join("\n\n");

  if (!convText) throw new Error("No conversation text to analyze.");

  var issueTypes = [
    "Not Reading Chat History","Premature Closure","Closing Pending Cases (PL Side)",
    "Not Raising Cases When/As Required","Not Checking/Verifying Before Responding",
    "Not Asking Clarifying Questions","Not/Missing Answering Customer's Actual Question",
    "Late Response Time - 1 day and above","Language Mismatch",
    "Wrong Team/Department Assigned","Lack of Empathy / Inappropriate Tone",
    "Not Informing Customer of Updates - Late & Ignore",
    "Not/Missing Collecting Supporting Documents","Wrong Action Taken",
    "Repetitive/Generic Responses","Incorrect Information Provided",
    "Not Sharing Correct Links/Hyperlinks","Not Sharing All Items/Documents",
    "Not Navigating System Properly","Over promise",
    "Not Educating Customer on Process","Not Providing Clear Guidance",
    "Ignore Customer Concern","Late Response Time - 1 hour and above","Failed verification"
  ];

  var prompt = 'You are a senior QA analyst at Platinumlist, a ticketing and events company in the GCC region.\n' +
    'Your job is to carefully read the customer support conversation below and extract structured information.\n\n' +

    'CRITICAL RULES — READ BEFORE ANALYZING:\n' +
    '• Read the ENTIRE conversation thoroughly before drawing any conclusions.\n' +
    '• For "issue_types": ONLY flag an issue if you have direct, undeniable evidence from the agent\'s actual messages. Be conservative. False positives are worse than missing an issue.\n' +
    '• If the agent DID provide information, details, or guidance — do NOT flag "Not Educating Customer on Process", "Not Providing Clear Guidance", or "Not/Missing Answering Customer\'s Actual Question".\n' +
    '• If the agent asked for details or clarified — do NOT flag "Not Asking Clarifying Questions".\n' +
    '• Do NOT flag an issue just because the conversation is short or the customer is unhappy — judge only the AGENT\'s behavior.\n' +
    '• If no issues are clearly present, return an empty array for issue_types.\n\n' +

    'Return ONLY a valid JSON object with exactly these 3 keys:\n\n' +

    '1. "order_numbers": array of strings — every order number, booking ID, reference code, or ticket number mentioned by anyone. Look for patterns like #, PL-, order, booking, ref, رقم الطلب. Return ["N/A"] if none found.\n\n' +

    '2. "keywords": array of objects {text, category} — extract ALL of the following if present:\n' +
    '   • Event/artist/concert names (e.g. "Doja Cat", "Layla Ahmadi Concert", "MDL Beast") — category: "event"\n' +
    '   • Event status (postponed, cancelled, rescheduled) — category: "postponed" or "cancelled"\n' +
    '   • Customer requests (refund, exchange, upgrade) — category: "refund" or "booking"\n' +
    '   • Complaints or escalation requests — category: "complaint" or "escalation"\n' +
    '   • Verification/identity checks — category: "verification"\n' +
    '   • Payment issues — category: "payment"\n' +
    '   • Venue or location names — category: "event"\n' +
    '   • Any other key topic — category: "other"\n' +
    '   Category must be exactly one of: refund, postponed, cancelled, complaint, escalation, event, verification, payment, account, booking, wrong_info, other\n\n' +

    '3. "issue_types": array of objects {name, reason} — ONLY include if clearly evidenced. Choose names EXACTLY from:\n' +
    issueTypes.map(function(t){ return '"'+t+'"'; }).join(', ') + '\n\n' +

    'For each issue_type you flag, the "reason" must quote or reference a specific thing the agent said (or failed to say) that proves the issue.\n\n' +

    'Conversation:\n"""\n' + convText + '\n"""\n\n' +
    'Return ONLY the JSON object. No explanation, no markdown, no code fences.';

  var response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5",
      max_tokens: 1500,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 400) throw new Error("Claude API " + code + ": " + text.slice(0, 300));

  var result = JSON.parse(text);
  var raw = (result.content && result.content[0]) ? result.content[0].text : "{}";

  // Strip markdown code fences if present
  raw = raw.replace(/```json\s*/gi, "").replace(/```\s*/g, "").trim();

  // Extract just the JSON object — find first { and last } to handle any trailing text
  var firstBrace = raw.indexOf("{");
  var lastBrace  = raw.lastIndexOf("}");
  if (firstBrace !== -1 && lastBrace !== -1 && lastBrace > firstBrace) {
    raw = raw.slice(firstBrace, lastBrace + 1);
  }

  // Try to parse; if it fails, return a safe fallback with the raw text for debugging
  try {
    return JSON.parse(raw);
  } catch (parseErr) {
    // Return safe empty result rather than crashing
    Logger.log("analyzeConversation JSON parse failed: " + parseErr + "\nRaw: " + raw.slice(0, 500));
    return {
      order_numbers: ["N/A"],
      keywords: [],
      issue_types: [],
      _parse_error: "AI response could not be parsed. Raw (first 200 chars): " + raw.slice(0, 200)
    };
  }
}

function uploadAudioChunk(payload) {
  var uploadId  = String(payload.upload_id  || "");
  var chunkIdx  = parseInt(payload.chunk_index) || 0;
  var chunkData = String(payload.chunk_data || "");
  var fileName  = String(payload.file_name  || "recording.mp3");
  if (!uploadId || !chunkData) throw new Error("upload_id and chunk_data required");

  // Decode base64 → binary → store as a Drive file
  var chunkName = uploadId + "_" + String(chunkIdx).padStart(5, "0");
  var bytes = Utilities.base64Decode(chunkData);
  var blob  = Utilities.newBlob(bytes, "application/octet-stream", chunkName);
  var folderName = "QA_Audio_Temp";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder  = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var chunkFile = folder.createFile(blob);
  return { ok: true, chunk: chunkIdx };
}

function finalizeAudioTranscription(payload) {
  var uploadId    = String(payload.upload_id    || "");
  var totalChunks = parseInt(payload.total_chunks) || 0;
  var mimeType    = String(payload.mime_type    || "audio/mpeg");
  var fileName    = String(payload.file_name    || "recording.mp3");
  if (!uploadId || !totalChunks) throw new Error("upload_id and total_chunks required");

  var s = getSettings();
  var apiKey = String(s.openai_api_key || "").trim();
  if (!apiKey) throw new Error("OpenAI API key not set. Add it in Settings → Call Transcription.");

  // Reassemble chunks from Drive
  var folderName = "QA_Audio_Temp";
  var folders = DriveApp.getFoldersByName(folderName);
  if (!folders.hasNext()) throw new Error("Temp folder not found in Drive.");
  var folder = folders.next();

  var allBytes = [];
  for (var i = 0; i < totalChunks; i++) {
    var chunkName  = uploadId + "_" + String(i).padStart(5, "0");
    var chunkFiles = folder.getFilesByName(chunkName);
    if (!chunkFiles.hasNext()) throw new Error("Missing chunk " + i + " of " + totalChunks + ". Please try again.");
    var chunkFile  = chunkFiles.next();
    var chunkBytes = chunkFile.getBlob().getBytes();
    // push in safe sub-batches to avoid "Maximum call stack size exceeded"
    var BATCH = 16384;
    for (var b = 0; b < chunkBytes.length; b += BATCH) {
      Array.prototype.push.apply(allBytes, chunkBytes.slice(b, b + BATCH));
    }
    try { chunkFile.setTrashed(true); } catch(e) {}
  }

  var audioBlob = Utilities.newBlob(allBytes, mimeType, fileName);

  // Send to Whisper
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: { "file": audioBlob, "model": "whisper-1", "response_format": "json" },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 400) throw new Error("Whisper API " + code + ": " + text.slice(0, 300));

  var result = JSON.parse(text);
  return { transcript: result.text || "" };
}

function createDriveUploadSession(payload) {
  var mimeType = payload.mime_type || "audio/mpeg";
  var fileName = payload.file_name || "recording.mp3";

  // Create (or reuse) a temp folder in Drive
  var folderName = "QA_Audio_Temp";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // Create an empty placeholder file so we have a file ID
  var placeholderBlob = Utilities.newBlob("", mimeType, fileName);
  var driveFile = folder.createFile(placeholderBlob);
  var fileId = driveFile.getId();

  // Make the file writable via resumable upload URL
  var token = ScriptApp.getOAuthToken();
  var initResp = UrlFetchApp.fetch(
    "https://www.googleapis.com/upload/drive/v3/files/" + fileId + "?uploadType=resumable",
    {
      method: "patch",
      headers: {
        "Authorization": "Bearer " + token,
        "Content-Type": "application/json",
        "X-Upload-Content-Type": mimeType
      },
      payload: JSON.stringify({ name: fileName, mimeType: mimeType }),
      muteHttpExceptions: true
    }
  );

  var uploadUrl = initResp.getHeaders()["Location"] || initResp.getHeaders()["location"] || "";
  if (!uploadUrl) throw new Error("Could not get Drive upload URL (status " + initResp.getResponseCode() + "): " + initResp.getContentText().slice(0,200));

  return { upload_url: uploadUrl, file_id: fileId };
}

function transcribeFromDrive(payload) {
  var s = getSettings();
  var apiKey = String(s.openai_api_key || "").trim();
  if (!apiKey) throw new Error("OpenAI API key not set. Add it in Settings → Call Transcription.");

  var fileId   = String(payload.file_id   || "").trim();
  var mimeType = String(payload.mime_type || "audio/mpeg").trim();
  var fileName = String(payload.file_name || "recording.mp3").trim();
  if (!fileId) throw new Error("No file_id provided.");

  // Read the file from Drive
  var driveFile = DriveApp.getFileById(fileId);
  var audioBlob = driveFile.getBlob();
  audioBlob.setContentType(mimeType);
  audioBlob.setName(fileName);

  // Call Whisper via UrlFetchApp (runs on Google's servers — no corporate firewall)
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: {
      "file":            audioBlob,
      "model":           "whisper-1",
      "response_format": "json"
    },
    muteHttpExceptions: true
  });

  // Clean up temp file from Drive regardless of Whisper result
  try { driveFile.setTrashed(true); } catch(e) { Logger.log("Could not trash temp file: " + e); }

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 400) throw new Error("Whisper API " + code + ": " + text.slice(0, 300));

  var result = JSON.parse(text);
  return { transcript: result.text || "" };
}

function getOpenAIKey() {
  var s = getSettings();
  var key = String(s.openai_api_key || "").trim();
  if (!key) throw new Error("OpenAI API key not set. Add it in Settings → OpenAI API Key.");
  return { key: key };
}

function transcribeAudio(payload) {
  var s = getSettings();
  var apiKey = String(s.openai_api_key || "").trim();
  if (!apiKey) throw new Error("OpenAI API key not set. Add it in Settings → Call Transcription.");

  var base64Data = payload.audio_base64 || "";
  var mimeType   = payload.mime_type   || "audio/mpeg";
  var fileName   = payload.file_name   || "recording.mp3";

  if (!base64Data) throw new Error("No audio data received.");

  // Decode base64 → binary blob
  var audioBytes = Utilities.base64Decode(base64Data);
  var audioBlob  = Utilities.newBlob(audioBytes, mimeType, fileName);

  // Call Whisper via Apps Script (bypasses corporate firewall)
  var response = UrlFetchApp.fetch("https://api.openai.com/v1/audio/transcriptions", {
    method: "post",
    headers: { "Authorization": "Bearer " + apiKey },
    payload: {
      "file":            audioBlob,
      "model":           "whisper-1",
      "response_format": "json"
    },
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 400) throw new Error("Whisper API " + code + ": " + text.slice(0, 300));

  var result = JSON.parse(text);
  return { transcript: result.text || "" };
}

function analyzeCallTranscript(payload) {
  var s = getSettings();
  var apiKey = String(s.claude_api_key || "").trim();
  if (!apiKey) throw new Error("Claude API key not set in Settings.");

  var transcript = String(payload.transcript || "").trim();
  var agentName  = String(payload.agent_name || "Agent").trim();
  if (!transcript) throw new Error("No transcript provided.");

  var issueTypes = [
    "Not Reading Chat History","Premature Closure","Closing Pending Cases (PL Side)",
    "Not Raising Cases When/As Required","Not Checking/Verifying Before Responding",
    "Not Asking Clarifying Questions","Not/Missing Answering Customer's Actual Question",
    "Late Response Time - 1 day and above","Language Mismatch",
    "Wrong Team/Department Assigned","Lack of Empathy / Inappropriate Tone",
    "Not Informing Customer of Updates - Late & Ignore",
    "Not/Missing Collecting Supporting Documents","Wrong Action Taken",
    "Repetitive/Generic Responses","Incorrect Information Provided",
    "Not Sharing Correct Links/Hyperlinks","Not Sharing All Items/Documents",
    "Not Navigating System Properly","Over promise",
    "Not Educating Customer on Process","Not Providing Clear Guidance",
    "Ignore Customer Concern","Late Response Time - 1 hour and above","Failed verification"
  ];

  var prompt =
    'You are a senior QA analyst at Platinumlist, a GCC ticketing and events company.\n' +
    'Analyze this call transcript between a customer support agent (' + agentName + ') and a customer.\n\n' +

    'CRITICAL RULES:\n' +
    '• Read the ENTIRE transcript before drawing any conclusions.\n' +
    '• agent_tone and customer_tone must reflect what is actually expressed — do not guess.\n' +
    '• Only flag issue_types if there is direct evidence in what the agent said or failed to say.\n' +
    '• If the agent provided information or guidance, do NOT flag "Not Providing Clear Guidance" or "Not Educating Customer on Process".\n' +
    '• If no issues are present, return empty array for issue_types.\n\n' +

    'Return ONLY a valid JSON object with exactly these keys:\n\n' +
    '{\n' +
    '  "order_numbers": ["array of order/booking/reference numbers mentioned, or N/A"],\n' +
    '  "summary": "2-3 sentence plain English summary of the call",\n' +
    '  "agent_tone": "one of: Professional, Empathetic, Neutral, Impatient, Unprofessional",\n' +
    '  "customer_tone": "one of: Calm, Frustrated, Angry, Satisfied, Confused",\n' +
    '  "agent_performance": "one of: Excellent, Good, Needs Improvement, Poor",\n' +
    '  "performance_notes": "2-3 specific observations about the agent\'s handling of the call",\n' +
    '  "keywords": [{"text": "keyword", "category": "one of: refund|postponed|cancelled|complaint|escalation|event|verification|payment|account|booking|wrong_info|other"}],\n' +
    '  "issue_types": [{"name": "exact name from list", "reason": "specific evidence from transcript"}]\n' +
    '}\n\n' +
    'issue_types names must be chosen EXACTLY from: ' + issueTypes.map(function(t){ return '"'+t+'"'; }).join(', ') + '\n\n' +
    'Transcript:\n"""\n' + transcript + '\n"""\n\n' +
    'Return ONLY the JSON. No markdown, no code fences, no explanation.';

  var response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5",
      max_tokens: 1500,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 400) throw new Error("Claude API " + code + ": " + text.slice(0, 300));

  var result = JSON.parse(text);
  var raw = (result.content && result.content[0]) ? result.content[0].text : "{}";
  raw = raw.replace(/```json\s*/gi,"").replace(/```\s*/g,"").trim();
  var firstBrace = raw.indexOf("{");
  var lastBrace  = raw.lastIndexOf("}");
  if (firstBrace !== -1 && lastBrace > firstBrace) raw = raw.slice(firstBrace, lastBrace + 1);
  try {
    return JSON.parse(raw);
  } catch(e) {
    return { order_numbers:["N/A"], summary:"", agent_tone:"", customer_tone:"", agent_performance:"", performance_notes:"", keywords:[], issue_types:[], _parse_error: raw.slice(0,200) };
  }
}

function getIntercomAdmins() {
  var data = callIntercomAPI("GET", "admins");
  var list = data.admins || data.data || [];
  return list.map(function(a){ return { id: a.id, name: a.name || "", email: a.email || "" }; });
}

function tagIntercomConversation(payload) {
  var convId = String(payload.conversation_id || "").trim();
  var tagId  = String(payload.tag_id          || "").trim();
  if (!convId || !tagId) throw new Error("conversation_id and tag_id required");
  return callIntercomAPI("POST", "conversations/" + convId + "/tags", { id: tagId });
}

function autoTagIntercom(chatId, validationStatus) {
  var s = getSettings();
  var tagId = validationStatus === "valid" ? (s.intercom_tag_valid || "") : (s.intercom_tag_invalid || "");
  if (!chatId || !tagId || !s.intercom_token) return;
  callIntercomAPI("POST", "conversations/" + chatId + "/tags", { id: tagId });
}

function postIntercomNote(chatId, row, validationStatus) {
  if (!chatId) return;
  var s = getSettings();
  if (!s.intercom_token) throw new Error("no_token");
  // Get admin ID — use saved setting first, else fetch from /admins
  var adminId = String(s.intercom_admin_id || "").trim();
  if (!adminId) {
    var admins = callIntercomAPI("GET", "admins");
    var adminList = admins.admins || admins.data || [];
    if (adminList.length === 0) throw new Error("no admins found in workspace");
    adminId = String(adminList[0].id || "");
  }
  if (!adminId) throw new Error("could not resolve admin_id");
  var isValid = validationStatus === "valid";
  var statusLabel = isValid ? "✅ Valid — Agent Fault" : "❌ Invalid — Not Agent Fault";

  // Build @mention tags for manager and agent (valid D-SATs only)
  var mentions = "";
  if (isValid) {
    var managerId = String(s.intercom_manager_id || "").trim();
    if (managerId) {
      // Look up manager's actual name from Intercom
      var managerName = "Manager";
      try {
        var allAdmins = callIntercomAPI("GET", "admins");
        var adminList = allAdmins.admins || allAdmins.data || [];
        var managerAdmin = adminList.filter(function(a){ return String(a.id) === managerId; })[0];
        if (managerAdmin && managerAdmin.name) managerName = managerAdmin.name;
      } catch(me) {}
      mentions += '<a class="intercom-mention" data-ref="admin_' + managerId + '">@' + managerName + '</a> ';
    }
    // Use stored intercom_admin_id from the Agents sheet (direct, no email-matching)
    try {
      var agents = getAgents();
      var agentRecord = agents.filter(function(a){ return (a["Agent name"]||"").toLowerCase() === (row.agent_name||"").toLowerCase(); })[0];
      var agentIntercomId = agentRecord ? String(agentRecord["intercom_admin_id"] || "").trim() : "";
      if (agentIntercomId) {
        mentions += '<a class="intercom-mention" data-ref="admin_' + agentIntercomId + '">@' + (row.agent_name||"Agent") + '</a> ';
      }
    } catch(me) {}
  }

  var lines = [
    "<b>D-SAT Validation</b>",
    "Status: <b>" + statusLabel + "</b>",
    "Responsible: " + (row.responsible || ""),
  ];
  if (row.agent_issue_type) lines.push("Issue Type: " + row.agent_issue_type);
  if (row.actionable_steps) lines.push("Actionable Steps: " + row.actionable_steps);
  if (row.additional_notes) lines.push("Notes: " + row.additional_notes);
  lines.push("Validated by: " + (row.done_by || row.validated_by || ""));
  lines.push("Date: " + (row.date || ""));
  if (mentions) lines.push("<br/>" + mentions);
  // Try to post note — if conversation is closed, open it first then re-close
  var adminIdInt = parseInt(adminId, 10);
  if (isNaN(adminIdInt)) throw new Error("invalid admin_id: [" + adminId + "]");
  var noteBody = { message_type: "note", type: "admin", admin_id: adminIdInt, body: lines.join("<br/>") };
  try {
    callIntercomAPI("POST", "conversations/" + chatId + "/reply", noteBody);
  } catch(e) {
    if (String(e.message).indexOf("closed") > -1 || String(e.message).indexOf("403") > -1 || String(e.message).indexOf("422") > -1) {
      // Reopen conversation, post note, close again
      try { callIntercomAPI("PUT", "conversations/" + chatId, { state: "open", admin_id: adminIdInt }); } catch(re) {}
      callIntercomAPI("POST", "conversations/" + chatId + "/reply", noteBody);
      try { callIntercomAPI("PUT", "conversations/" + chatId, { state: "closed", admin_id: adminIdInt }); } catch(ce) {}
    } else {
      throw e;
    }
  }
}

// ============================================================
// AGENT MONTHLY SUMMARY
// ============================================================

function getAgentMonthlySummary(payload) {
  var agentName = payload.agentName || payload.agent_name || "";
  var month     = payload.month || "";
  var year      = String(payload.year || "");

  var transactions = getQATransactions({ agentName: agentName, month: month, year: year });
  transactions.sort(function(a, b){ return Number(a.transaction_no) - Number(b.transaction_no); });

  var avgScore = 0;
  if (transactions.length > 0) {
    var total = transactions.reduce(function(s, tx){ return s + Number(tx.final_score || 0); }, 0);
    avgScore = Math.round((total / transactions.length) * 100) / 100;
  }

  var weakAreas    = detectWeakAreas(transactions);
  var allImp       = getDSATImports({ agentName: agentName, month: month, year: year });

  // Match validations by chat_id across ALL validations (no month/agent filter)
  // This correctly handles cases where a chat was validated under a different month or agent name spelling
  var allValsGlobal    = getDSATValidations({});
  var impChatIds       = allImp.map(function(i){ return String(i.chat_id).trim(); });
  var allVal           = allValsGlobal.filter(function(v){ return impChatIds.indexOf(String(v.chat_id).trim()) !== -1; });
  var validatedChatIds = allValsGlobal.map(function(v){ return String(v.chat_id).trim(); });
  var validDsats       = allVal.filter(function(r){ return r.validation_status === "valid"; });
  var invalidDsats     = allVal.filter(function(r){ return r.validation_status === "invalid"; });
  var pendingImports   = allImp.filter(function(imp){
    return !validatedChatIds.includes(String(imp.chat_id).trim());
  });

  var issueDist = {};
  validDsats.forEach(function(r) {
    var it = r.agent_issue_type || "Unspecified";
    issueDist[it] = (issueDist[it] || 0) + 1;
  });
  var issueDistArr = Object.keys(issueDist).map(function(k) {
    return { issue_type: k, count: issueDist[k], pct: validDsats.length ? Math.round((issueDist[k] / validDsats.length) * 100) : 0 };
  }).sort(function(a, b){ return b.count - a.count; });

  var agents    = getAgents();
  var agentRow  = null;
  for (var i = 0; i < agents.length; i++) {
    if (agents[i]["Agent name"] === agentName || agents[i]["agent_name"] === agentName) {
      agentRow = agents[i]; break;
    }
  }
  var sheetLink = agentRow ? (agentRow["online sheet link"] || "") : "";

  return {
    agent_name:          agentName,
    month:               month,
    year:                year,
    transactions:        transactions,
    avg_score:           avgScore,
    weak_areas:          weakAreas,
    parameter_breakdown: buildParameterBreakdown(transactions),
    dsat_total:          allImp.length,
    dsat_validated:      allVal.length,
    dsat_valid:          validDsats.length,
    dsat_invalid:        invalidDsats.length,
    dsat_pending:        pendingImports.length,
    valid_dsats:         validDsats,
    issue_distribution:  issueDistArr,
    agent_sheet_link:    sheetLink
  };
}

// ============================================================
// MANAGER DASHBOARD
// ============================================================

function getManagerDashboard(payload) {
  var month = payload.month || "";
  var year  = String(payload.year || "");

  var allTx  = getQATransactions({ month: month, year: year });
  var allImp = getDSATImports({ month: month, year: year });
  var allVal = getDSATValidations({ month: month, year: year });

  var agentQA = {};
  allTx.forEach(function(tx) {
    var n = tx.agent_name;
    if (!agentQA[n]) agentQA[n] = [];
    agentQA[n].push(Number(tx.final_score || 0));
  });

  var agentDsatRaw     = {};
  var agentDsatValid   = {};
  var agentDsatInvalid = {};
  allImp.forEach(function(r){ agentDsatRaw[r.agent_name] = (agentDsatRaw[r.agent_name] || 0) + 1; });
  allVal.forEach(function(r) {
    if (r.validation_status === "valid")
      agentDsatValid[r.agent_name] = (agentDsatValid[r.agent_name] || 0) + 1;
    else
      agentDsatInvalid[r.agent_name] = (agentDsatInvalid[r.agent_name] || 0) + 1;
  });

  var ranked = Object.keys(agentQA).map(function(name) {
    var scores = agentQA[name];
    var avg = Math.round(scores.reduce(function(s,v){ return s+v; }, 0) / scores.length * 100) / 100;
    var rl = avg >= 80 ? "Nailed It!" : avg >= 70 ? "Almost There!" : "Not Quite There!";
    return {
      agent_name:   name,
      avg_score:    avg,
      rating_label: rl,
      tx_count:     scores.length,
      dsat_raw:     agentDsatRaw[name]     || 0,
      dsat_valid:   agentDsatValid[name]   || 0,
      dsat_invalid: agentDsatInvalid[name] || 0
    };
  }).sort(function(a, b){ return b.avg_score - a.avg_score; });

  ranked.forEach(function(r, i){ r.rank = i + 1; });

  var teamAvg = ranked.length
    ? Math.round(ranked.reduce(function(s,r){ return s + r.avg_score; }, 0) / ranked.length * 100) / 100
    : 0;

  var allIssues = {};
  allVal.filter(function(r){ return r.validation_status === "valid"; }).forEach(function(r) {
    var it = r.agent_issue_type || "Unspecified";
    allIssues[it] = (allIssues[it] || 0) + 1;
  });
  var topIssue = Object.keys(allIssues).reduce(function(top, k){
    return (!top || allIssues[k] > allIssues[top]) ? k : top;
  }, null);

  var weakAreas = detectWeakAreas(allTx);
  var totalValidDsat = allVal.filter(function(r){ return r.validation_status === "valid"; }).length;

  return {
    month:            month,
    year:             year,
    agents:           ranked,
    team_avg:         teamAvg,
    agents_above_90:  ranked.filter(function(r){ return r.avg_score >= 80; }).length,
    total_valid_dsat: totalValidDsat,
    total_raw_dsat:   allImp.length,
    top_issue_type:   topIssue || "N/A",
    team_weak_areas:  weakAreas
  };
}

// ============================================================
// SETTINGS
// ============================================================

function getSettings() {
  var rows = [];
  try { rows = getSheetData(TABS.settings); } catch(e) {}
  var settings = {};
  rows.forEach(function(r) {
    if (r.key) settings[String(r.key)] = String(r.value || "");
  });
  if (!settings.max_transactions) settings.max_transactions = "8";
  if (!settings.coordinator_name) settings.coordinator_name = "Quality & Training Team";
  if (!settings.coordinator_title) settings.coordinator_title = "Regional Training and Quality Coordinator";
  if (!settings.done_by_names) settings.done_by_names = "";
  return settings;
}

function saveSettings(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.settings);
  if (!sheet) {
    sheet = ss.insertSheet(TABS.settings);
    sheet.getRange(1,1,1,4).setValues([SETTINGS_HEADERS]);
    sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#cfe2ff");
    sheet.setFrozenRows(1);
  }
  var values = sheet.getDataRange().getValues();
  var updatedBy = String(payload.updated_by || "");
  var now = nowISO();
  var keys = Object.keys(payload).filter(function(k){ return k !== "updated_by"; });
  keys.forEach(function(key) {
    var found = false;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === key) {
        sheet.getRange(i+1, 2, 1, 3).setValues([[String(payload[key]), now, updatedBy]]);
        values[i][1] = String(payload[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      var newRow = [key, String(payload[key]), now, updatedBy];
      sheet.appendRow(newRow);
      values.push(newRow);
    }
  });
  return { message: "Settings saved." };
}

// ============================================================
// UPDATE / DELETE QA TRANSACTION
// ============================================================

function updateQATransaction(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.qaTransactions);
  if (!sheet) throw new Error("QATransactions tab not found");
  var values = sheet.getDataRange().getValues();
  var qaId = String(payload.qa_id || "");
  if (!qaId) throw new Error("qa_id is required");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === qaId) {
      var dateStr = payload.date;
      var scores = calculateScores(payload.ratings || {});
      var updatedRow = [
        qaId,
        payload.agent_name,
        dateStr,
        getMonthName(dateStr),
        new Date(dateStr).getFullYear(),
        getWeekOfMonth(dateStr),
        values[i][6], // keep original transaction_no
        payload.channel,
        payload.ticket_id || "",
        payload.interaction_reason || "",
        JSON.stringify(payload.ratings || {}),
        JSON.stringify(payload.notes || {}),
        scores.total_without_fatal,
        scores.final_score,
        scores.rating_label,
        scores.has_ko,
        payload.saved_by || "",
        nowISO(),
        payload.transaction_type || ""
      ];
      sheet.getRange(i+1, 1, 1, updatedRow.length).setValues([updatedRow]);
      return { qa_id: qaId, final_score: scores.final_score, rating_label: scores.rating_label, has_ko: scores.has_ko };
    }
  }
  throw new Error("QA Transaction not found: " + qaId);
}

function deleteQATransaction(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.qaTransactions);
  if (!sheet) throw new Error("QATransactions tab not found");
  var values = sheet.getDataRange().getValues();
  var qaId = String(payload.qa_id || "");
  if (!qaId) throw new Error("qa_id is required");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === qaId) {
      sheet.deleteRow(i+1);
      return { deleted: true, qa_id: qaId };
    }
  }
  throw new Error("QA Transaction not found: " + qaId);
}

// ============================================================
// UPDATE D-SAT VALIDATION
// ============================================================

function updateDSATValidation(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.dsatValidations);
  if (!sheet) throw new Error("DSATValidations tab not found");
  var values = sheet.getDataRange().getValues();
  var dsatId = String(payload.dsat_id || "");
  if (!dsatId) throw new Error("dsat_id is required");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === dsatId) {
      var responsible = (payload.responsible || "").toLowerCase().trim();
      var validationStatus = (responsible === "agent") ? "valid" : "invalid";
      // Date change: recompute month + year if new date provided
      var newDate = String(payload.date || "").trim();
      var rowDate  = newDate || String(values[i][2] || "");
      var rowMonth = values[i][3];
      var rowYear  = values[i][4];
      if (newDate) {
        var dp = newDate.split("-");
        if (dp.length === 3) {
          rowMonth = getMonthName(parseInt(dp[1], 10));
          rowYear  = dp[0];
        }
      }
      var updatedRow = [
        dsatId,
        values[i][1], // import_id
        rowDate,
        rowMonth,
        rowYear,
        values[i][5], // agent_name
        values[i][6], // chat_id
        values[i][7], // rating
        values[i][8], // transaction_type
        payload.responsible || "",
        validationStatus,
        payload.agent_issue_type || "",
        payload.actionable_steps || "",
        payload.additional_notes || "",
        payload.done_by || "",
        payload.coached_by || "",
        payload.validated_by || payload.done_by || "",
        nowISO()
      ];
      sheet.getRange(i+1, 1, 1, updatedRow.length).setValues([updatedRow]);
      return { dsat_id: dsatId, validation_status: validationStatus };
    }
  }
  throw new Error("DSATValidation record not found: " + dsatId);
}

// ============================================================
// EXPORT DATA
// ============================================================

function exportData(payload) {
  var tabName = payload.tab || TABS.qaTransactions;
  var rows = getSheetData(tabName);
  return { tab: tabName, rows: rows, count: rows.length };
}

// ============================================================
// PARAMETER BREAKDOWN (for agent view charts)
// ============================================================

function buildParameterBreakdown(transactions) {
  var result = [];
  QA_PARAMETERS.filter(function(p){ return p.type !== "fatal"; }).forEach(function(p) {
    var scores = [];
    transactions.forEach(function(tx) {
      try {
        var r = JSON.parse(tx.ratings_json || "{}");
        var v = Number(r[p.id]);
        if (!isNaN(v)) scores.push(v);
      } catch(e) {}
    });
    if (scores.length === 0) return;
    var avg = scores.reduce(function(a,b){ return a+b; },0) / scores.length;
    result.push({
      id: p.id,
      name: p.name,
      max: p.max,
      avg: Math.round(avg * 100) / 100,
      pct: Math.round((avg / p.max) * 100),
      target_pct: 75,
      below_target: Math.round((avg / p.max) * 100) < 75,
      scores: scores,
      type: p.type,
      criticality: p.criticality
    });
  });
  return result;
}

// ============================================================
// COACHING EMAIL GENERATOR
// ============================================================

// ============================================================
// DELETE D-SAT VALIDATION (returns case to pending queue)
// ============================================================

function deleteDSATValidation(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.dsatValidations);
  if (!sheet) throw new Error("DSATValidations tab not found");
  var values = sheet.getDataRange().getValues();
  var dsatId = String(payload.dsat_id || "");
  if (!dsatId) throw new Error("dsat_id is required");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === dsatId) {
      sheet.deleteRow(i + 1);
      return { deleted: true, dsat_id: dsatId };
    }
  }
  throw new Error("D-SAT Validation record not found: " + dsatId);
}

// ============================================================
// C-SAT DATA — Manual input per agent per month
// ============================================================

function saveCsatData(payload) {
  var month = payload.month || "";
  var year  = String(payload.year || "");
  var data  = payload.data || {};
  var key   = "csat__" + month + "__" + year;
  var sp    = {};
  sp[key]   = JSON.stringify(data);
  sp.updated_by = payload.updated_by || "";
  return saveSettings(sp);
}

function getCsatData(payload) {
  var month    = payload.month || "";
  var year     = String(payload.year || "");
  var key      = "csat__" + month + "__" + year;
  var settings = getSettings();
  var raw      = settings[key];
  if (!raw) return {};
  try { return JSON.parse(raw); } catch(e) { return {}; }
}

// ============================================================
// USER AUTHENTICATION
// ============================================================

function loginUser(payload) {
  var email    = String(payload.email    || "").toLowerCase().trim();
  var password = String(payload.password || "");
  if (!email || !password) throw new Error("Email and password are required.");

  var users = getSheetData(TABS.users);
  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    if (String(u.email || "").toLowerCase().trim() === email) {
      if (String(u.active || "").toUpperCase() === "FALSE") throw new Error("Account is inactive. Contact your administrator.");
      if (String(u.password || "") !== password) throw new Error("Incorrect password.");
      return { user_id: u.user_id, name: u.name, email: u.email, role: u.role || "user" };
    }
  }
  throw new Error("No account found for that email address.");
}

function getUsers() {
  var users = getSheetData(TABS.users);
  // Return without password field
  return users.map(function(u) {
    return { user_id: u.user_id, name: u.name, email: u.email, role: u.role, active: u.active, created_at: u.created_at };
  });
}

function addUser(payload) {
  var name     = String(payload.name     || "").trim();
  var email    = String(payload.email    || "").toLowerCase().trim();
  var password = String(payload.password || "").trim();
  var role     = String(payload.role     || "user").trim();
  if (!name || !email || !password) throw new Error("Name, email, and password are required.");

  // Check duplicate
  var existing = getSheetData(TABS.users);
  for (var i = 0; i < existing.length; i++) {
    if (String(existing[i].email || "").toLowerCase().trim() === email) throw new Error("A user with that email already exists.");
  }

  var row = {
    user_id:    generateId("usr"),
    name:       name,
    email:      email,
    password:   password,
    role:       role,
    active:     "TRUE",
    created_at: nowISO()
  };
  appendRow(TABS.users, USER_HEADERS, row);
  return { user_id: row.user_id, name: row.name, email: row.email, role: row.role };
}

function deleteUser(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.users);
  if (!sheet) throw new Error("AppUsers tab not found");
  var values = sheet.getDataRange().getValues();
  var userId = String(payload.user_id || "");
  if (!userId) throw new Error("user_id is required");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === userId) {
      sheet.deleteRow(i + 1);
      return { deleted: true, user_id: userId };
    }
  }
  throw new Error("User not found: " + userId);
}

function changePassword(payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(TABS.users);
  if (!sheet) throw new Error("AppUsers tab not found");
  var values = sheet.getDataRange().getValues();
  var email       = String(payload.email       || "").toLowerCase().trim();
  var oldPassword = String(payload.oldPassword || "");
  var newPassword = String(payload.newPassword || "").trim();
  if (!email || !oldPassword || !newPassword) throw new Error("All fields are required.");
  if (newPassword.length < 6) throw new Error("New password must be at least 6 characters.");
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][2] || "").toLowerCase().trim() === email) {
      if (String(values[i][3] || "") !== oldPassword) throw new Error("Current password is incorrect.");
      sheet.getRange(i + 1, 4).setValue(newPassword);
      return { changed: true };
    }
  }
  throw new Error("User not found.");
}

function generateCoachingEmail(payload) {
  var summary    = getAgentMonthlySummary(payload);
  var settings   = getSettings();
  var agentName  = summary.agent_name;
  var monthLabel = summary.month + " " + summary.year;
  var coordName  = settings.coordinator_name || "Quality & Training Team";
  var coordTitle = settings.coordinator_title || "Regional Training and Quality Coordinator";

  var L = [];
  L.push("Hi " + agentName + ",");
  L.push("");
  L.push("Here is your quality review summary for " + monthLabel + ".");
  L.push("");
  L.push("-- QA PERFORMANCE -------------------------------------------------");
  L.push("Average Score : " + summary.avg_score + " / 100");
  var rl = summary.avg_score >= 80 ? "Nailed It!" : summary.avg_score >= 70 ? "Almost There!" : "Not Quite There!";
  L.push("Rating        : " + rl);
  L.push("Transactions  : " + summary.transactions.length + " of 8 evaluated");
  L.push("");

  if (summary.transactions.length > 0) {
    L.push("Transaction breakdown:");
    summary.transactions.forEach(function(tx) {
      var ko = (String(tx.has_ko) === "true" || String(tx.has_ko) === "TRUE") ? " [K.O]" : "";
      L.push("  T" + tx.transaction_no + " (" + tx.date + ") -- " + (tx.interaction_reason || "N/A") + " -- Score: " + tx.final_score + ko);
    });
    L.push("");
  }

  L.push("-- AREAS FOR DEVELOPMENT ------------------------------------------");
  if (summary.weak_areas && summary.weak_areas.length > 0) {
    L.push("The following parameters were consistently below target (75%):");
    L.push("");
    summary.weak_areas.forEach(function(w, i) {
      L.push("  " + (i+1) + ". " + w.parameter_name);
      L.push("     Average: " + w.avg_score + " / " + w.max + " (" + w.avg_pct + "%) -- Low in " + w.low_in + " of " + w.of_total + " transactions (" + w.repetition_pct + "%)");
    });

    // Include per-transaction scores for weak parameters
    L.push("");
    L.push("  Parameter scores per transaction:");
    summary.weak_areas.forEach(function(w) {
      var paramLine = "  " + w.parameter_name + ": ";
      var scores = [];
      summary.transactions.forEach(function(tx) {
        try {
          var ratings = JSON.parse(tx.ratings_json || "{}");
          var pid = w.parameter_id;
          scores.push("T" + tx.transaction_no + "=" + (ratings[pid] !== undefined ? ratings[pid] : "?"));
        } catch(e) {}
      });
      L.push(paramLine + scores.join(", "));
    });
  } else {
    L.push("All parameters met the target this month. Great work!");
  }
  L.push("");

  L.push("-- D-SAT REVIEW ---------------------------------------------------");
  L.push("D-SATs received   : " + summary.dsat_total);
  L.push("Validated         : " + summary.dsat_validated);
  L.push("Valid (agent)     : " + summary.dsat_valid);
  L.push("Invalid (not agent): " + summary.dsat_invalid);
  L.push("Pending           : " + summary.dsat_pending);
  L.push("");

  if (summary.valid_dsats && summary.valid_dsats.length > 0) {
    L.push("Agent-fault D-SATs this month: " + summary.dsat_valid);
    L.push("");
    if (summary.issue_distribution && summary.issue_distribution.length > 0) {
      L.push("Issue breakdown:");
      summary.issue_distribution.forEach(function(d) {
        L.push("  * " + d.issue_type + ": " + d.count + " case(s) (" + d.pct + "%)");
      });
    }
    var actionable = (summary.valid_dsats || [])
      .filter(function(r){ return r.actionable_steps && String(r.actionable_steps).trim(); })
      .map(function(r){ return String(r.actionable_steps).trim(); });
    if (actionable.length > 0) {
      L.push("");
      L.push("Actionable steps:");
      actionable.forEach(function(a){ L.push("  - " + a); });
    }
  } else {
    L.push("No agent-fault D-SATs recorded for this period.");
  }

  if (summary.dsat_invalid > 0) {
    L.push("");
    L.push("Note: " + summary.dsat_invalid + " D-SAT(s) were validated as not agent fault (customer / process / technology).");
  }

  L.push("");
  L.push("-------------------------------------------------------------------");
  if (summary.agent_sheet_link) {
    L.push("Your scorecard: " + summary.agent_sheet_link);
    L.push("");
  }
  L.push("Please feel free to reach out if you have any questions.");
  L.push("");
  L.push(coordName);
  L.push(coordTitle);

  var subject = "QA Coaching -- " + agentName + " -- " + monthLabel;
  return { subject: subject, body: L.join("\n"), summary: summary };
}

// ============================================================
// TICKET FINDABILITY ANALYSIS — for Product / Business Teams
// Run one of these from Apps Script editor:
//   runDecemberAnalysis()  → December 2025
//   runJanuaryAnalysis()   → January 2026
// Results written live to Google Sheet tabs per month.
// ============================================================

function runDecemberAnalysis() {
  _runTicketAnalysis(
    1764547200,              // Dec  1, 2025 00:00:00 UTC
    1767225599,              // Dec 31, 2025 23:59:59 UTC
    "TicketAnalysis_Dec2025",
    "ticketCursor_Dec2025"
  );
}

function runJanuaryAnalysis() {
  _runTicketAnalysis(
    1767225600,              // Jan  1, 2026 00:00:00 UTC
    1769903999,              // Jan 31, 2026 23:59:59 UTC
    "TicketAnalysis_Jan2026",
    "ticketCursor_Jan2026"
  );
}

// Keep old name working too (runs December)
function runTicketFindabilityAnalysis() { runDecemberAnalysis(); }

function _runTicketAnalysis(START_TS, END_TS, TAB_NAME, CHECKPOINT_KEY) {

  var MAX_PAGES     = 60;

  var settings      = getSettings();
  var intercomToken = settings.intercom_token || "";
  var claudeKey     = settings.claude_api_key  || "";
  if (!intercomToken) throw new Error("Intercom token not configured in Settings.");
  if (!claudeKey)     throw new Error("Claude API key not configured in Settings.");

  var ss  = SpreadsheetApp.openById(SHEET_ID);
  var tab = ss.getSheetByName(TAB_NAME);

  // ── Resume or fresh start ────────────────────────────────────
  var savedCursor  = PropertiesService.getScriptProperties().getProperty(CHECKPOINT_KEY);
  var savedCounts  = PropertiesService.getScriptProperties().getProperty(CHECKPOINT_KEY + "_counts");
  var resuming     = !!(savedCursor && tab && tab.getLastRow() > 1);

  var rowNum     = 0;
  var matchCount = 0;
  var errorCount = 0;

  if (resuming) {
    // Pick up from saved position
    var counts  = JSON.parse(savedCounts || '{"row":0,"match":0,"error":0}');
    rowNum     = counts.row;
    matchCount = counts.match;
    errorCount = counts.error;
    Logger.log("▶ RESUMING from row " + rowNum + " | cursor: " + savedCursor.slice(0,20) + "...");
  } else {
    // Fresh start — recreate sheet
    if (tab) ss.deleteSheet(tab);
    tab = ss.insertSheet(TAB_NAME);
    tab.appendRow(["#", "Conversation ID", "Date", "Customer Email", "Match (YES/NO)", "Reason", "First Customer Message"]);
    tab.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#4a0e8f").setFontColor("#ffffff");
    tab.setFrozenRows(1);
    PropertiesService.getScriptProperties().deleteProperty(CHECKPOINT_KEY);
    PropertiesService.getScriptProperties().deleteProperty(CHECKPOINT_KEY + "_counts");
    Logger.log("🆕 Fresh start — " + TAB_NAME);
  }

  var cursor    = resuming ? savedCursor : null;
  var pageCount = 0;
  var batchRows = [];

  Logger.log("Starting Ticket Findability Analysis — " + TAB_NAME + "...");

  do {
    // ── Fetch one page of search results ─────────────────────
    var searchBody = {
      query: {
        operator: "AND",
        value: [
          { field: "created_at", operator: ">", value: START_TS },
          { field: "created_at", operator: "<", value: END_TS   }
        ]
      },
      pagination: { per_page: 20 }  // reduced from 50 to lower data per call
    };
    if (cursor) searchBody.pagination.starting_after = cursor;

    var searchRes;
    var retries = 0;
    while (retries < 3) {
      try {
        searchRes = callIntercomAPI("POST", "conversations/search", searchBody);
        break; // success
      } catch(e) {
        retries++;
        if (e.message.indexOf("Bandwidth quota") !== -1 && retries < 3) {
          Logger.log("⚠️ Bandwidth quota hit on search — waiting 90 seconds before retry " + retries + "/3...");
          Utilities.sleep(90000);
        } else {
          Logger.log("Search failed after " + retries + " retries: " + e.message);
          // Save checkpoint so next run can resume
          if (cursor) {
            PropertiesService.getScriptProperties().setProperty(CHECKPOINT_KEY, cursor);
            PropertiesService.getScriptProperties().setProperty(CHECKPOINT_KEY + "_counts",
              JSON.stringify({ row: rowNum, match: matchCount, error: errorCount }));
          }
          // Flush any remaining rows before exiting
          if (batchRows.length > 0) {
            tab.getRange(tab.getLastRow() + 1, 1, batchRows.length, 7).setValues(batchRows);
            batchRows = [];
            SpreadsheetApp.flush();
          }
          Logger.log("⏸ Paused at row " + rowNum + ". Run again to resume from checkpoint.");
          return "PAUSED at row " + rowNum + " — quota hit. Run again in 1-2 hours to resume.";
        }
      }
    }
    if (!searchRes) break;

    var convList = searchRes.conversations || [];
    if (convList.length === 0) break;

    // ── Classify each conversation using source.body only ────
    // No secondary Intercom API call — zero bandwidth quota risk.
    for (var i = 0; i < convList.length; i++) {
      var conv = convList[i];
      rowNum++;

      var convDate  = conv.created_at
        ? Utilities.formatDate(new Date(conv.created_at * 1000), "UTC", "yyyy-MM-dd")
        : "";

      // Customer email — already in search result contacts list (no extra fetch)
      var custEmail = "";
      try {
        var ctcs = conv.contacts && conv.contacts.contacts;
        if (ctcs && ctcs.length && ctcs[0].email) custEmail = ctcs[0].email;
      } catch(e) {}

      // First customer message from source (already in search result)
      var firstMsg = "";
      try {
        if (conv.source && conv.source.body) {
          firstMsg = conv.source.body.replace(/<[^>]+>/g, "").trim().slice(0, 400);
        }
      } catch(e) {}

      var match  = "NO";
      var reason = "";

      if (firstMsg.length > 10) {
        // ── Step 1: keyword pre-filter ──────────────────────────
        // Only call Claude if message contains at least one relevant word.
        // Skips ~85-90% of conversations (refunds, event info, etc.)
        if (_hasTicketKeyword(firstMsg)) {
          try {
            var res = _classifyTicketFindability(firstMsg, claudeKey);
            match  = res.match;
            reason = res.reason;
            if (match === "YES") matchCount++;
          } catch(e) {
            errorCount++;
            reason = "Claude error: " + e.message.slice(0, 80);
          }
        } else {
          // Skipped — no ticket-related keyword found
          match  = "NO";
          reason = "Skipped — no ticket keyword";
        }
      } else {
        reason = "No message text";
      }

      batchRows.push([rowNum, conv.id, convDate, custEmail, match, reason, firstMsg.slice(0, 200)]);

      // Flush to sheet every 20 rows so data is safe if it times out
      if (batchRows.length >= 20) {
        tab.getRange(tab.getLastRow() + 1, 1, batchRows.length, 7).setValues(batchRows);
        // Highlight YES rows
        var sheetRow = tab.getLastRow() - batchRows.length + 1;
        batchRows.forEach(function(r, ri) {
          if (r[4] === "YES") tab.getRange(sheetRow + ri, 5).setBackground("#c6efce").setFontWeight("bold");
        });
        batchRows = [];
        SpreadsheetApp.flush();
      }
    }

    pageCount++;
    cursor = (searchRes.pages && searchRes.pages.next && searchRes.pages.next.starting_after)
             ? searchRes.pages.next.starting_after : null;

    // Save checkpoint so we can resume if stopped or timed out
    if (cursor) {
      PropertiesService.getScriptProperties().setProperty(CHECKPOINT_KEY, cursor);
      PropertiesService.getScriptProperties().setProperty(CHECKPOINT_KEY + "_counts",
        JSON.stringify({ row: rowNum, match: matchCount, error: errorCount }));
    }

    var claudeCalls = rowNum - errorCount; // approximate
    Logger.log("Page " + pageCount + " done | Scanned: " + rowNum + " | Sent to Claude: ~" +
               Math.round(rowNum * 0.15) + " (keyword filtered) | Matches: " + matchCount +
               (cursor ? " | Checkpoint saved ✓" : " | Last page ✅"));
    if (cursor) Utilities.sleep(800);

  } while (cursor && pageCount < MAX_PAGES);

  // Flush remaining rows
  if (batchRows.length > 0) {
    tab.getRange(tab.getLastRow() + 1, 1, batchRows.length, 7).setValues(batchRows);
    var sheetRow = tab.getLastRow() - batchRows.length + 1;
    batchRows.forEach(function(r, ri) {
      if (r[4] === "YES") tab.getRange(sheetRow + ri, 5).setBackground("#c6efce").setFontWeight("bold");
    });
  }

  // ── Summary ──────────────────────────────────────────────────
  tab.appendRow(["—", "—", "—", "—", "—", "—", "—"]);
  tab.appendRow(["SUMMARY", "", "", "Total Conversations:", rowNum, "Can't find ticket (YES):", matchCount]);
  tab.appendRow(["", "", "", "Errors / No text:", errorCount, "Match Rate:", rowNum > 0 ? (matchCount/rowNum*100).toFixed(1)+"%" : "0%"]);
  tab.getRange(tab.getLastRow()-1, 1, 2, 7).setFontWeight("bold").setBackground("#fff2cc");
  SpreadsheetApp.flush();

  // Clear checkpoint — analysis complete
  PropertiesService.getScriptProperties().deleteProperty(CHECKPOINT_KEY);
  PropertiesService.getScriptProperties().deleteProperty(CHECKPOINT_KEY + "_counts");

  var summary = "✅ COMPLETE! Analysed: " + rowNum + " | Matches: " + matchCount + " | Errors: " + errorCount;
  Logger.log(summary);
  return summary;
}

// Keyword pre-filter — returns true only if message likely relates to ticket findability
function _hasTicketKeyword(text) {
  var t = text.toLowerCase();
  var keywords = [
    // English
    "ticket", "booking", "e-ticket", "eticket",
    "can't find", "cannot find", "cant find",
    "didn't receive", "did not receive", "not received",
    "not showing", "not show", "don't see", "do not see", "cant see", "can't see",
    "where is my", "where are my",
    "no ticket", "missing ticket",
    "didn't get", "did not get",
    "not arrived", "didn't arrive",
    "can't access", "cannot access",
    "how do i get my", "how to get my",
    "never got", "never received",
    "purchase", "bought", "paid",
    // Arabic
    "تذكرة", "تذاكر",
    "ما لقيت", "ما وجدت", "ما أشوف", "ما شفت",
    "ما وصل", "ما وصلت", "ما جاني", "ما جتني",
    "وين تذكر", "فين تذكر",
    "ما عندي تذكرة", "ما عندي تذاكر",
    "كيف أحصل", "كيف اوصل",
    "دفعت", "اشتريت", "اشترينا",
    "ما ظهرت", "ما ظهر",
    "ما وصلني إيميل", "ما وصلني ايميل"
  ];
  for (var i = 0; i < keywords.length; i++) {
    if (t.indexOf(keywords[i]) !== -1) return true;
  }
  return false;
}

// Claude classifier — returns {match: "YES"/"NO", reason: string}
function _classifyTicketFindability(text, apiKey) {
  var prompt =
    'You are analysing a Platinumlist customer support conversation. The conversation may be in English, Arabic, or both.\n\n' +
    'Task: Decide if the customer is saying they CANNOT FIND, SEE, ACCESS or LOCATE their ticket(s) after completing a purchase.\n\n' +

    '✅ Examples that are YES (English):\n' +
    '- "I bought tickets but I cannot find them"\n' +
    '- "Where are my tickets?"\n' +
    '- "My tickets did not arrive"\n' +
    '- "I don\'t see the ticket in my account or email"\n' +
    '- "I purchased but I have no ticket"\n' +
    '- "How do I get my ticket?"\n' +
    '- "I can\'t access my booking"\n\n' +

    '✅ Examples that are YES (Arabic):\n' +
    '- "اشتريت تذاكر بس ما لقيتها" (bought tickets but can\'t find them)\n' +
    '- "وين تذاكري؟" (where are my tickets?)\n' +
    '- "ما وصلتني التذكرة" (ticket didn\'t arrive to me)\n' +
    '- "ما أشوف التذاكر في حسابي" (don\'t see tickets in my account)\n' +
    '- "دفعت بس ما عندي تذكرة" (paid but have no ticket)\n' +
    '- "تذاكري ما ظهرت" (my tickets didn\'t appear)\n' +
    '- "كيف أحصل على تذكرتي؟" (how do I get my ticket?)\n' +
    '- "ما جاني أي إيميل فيه التذكرة" (didn\'t receive any email with the ticket)\n\n' +

    '❌ Examples that are NO:\n' +
    '- Asking for a refund or cancellation\n' +
    '- Asking about event details, timing, or location\n' +
    '- General complaints not about finding tickets\n' +
    '- Complaints about wrong tickets or wrong event (different issue)\n' +
    '- Already resolved — agent already sent the ticket in the same conversation\n\n' +

    'Customer message:\n' + text + '\n\n' +
    'Reply in this exact format:\n' +
    'MATCH: YES or NO\n' +
    'REASON: one short sentence explaining why';

  var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "content-type": "application/json"
    },
    payload: JSON.stringify({
      model: "claude-haiku-4-5",
      max_tokens: 60,
      messages: [{ role: "user", content: prompt }]
    }),
    muteHttpExceptions: true
  });

  var data = JSON.parse(res.getContentText());
  if (data.error) return { match: "NO", reason: "Claude error: " + (data.error.message || "") };

  var raw    = (data.content && data.content[0] && data.content[0].text) || "";
  var mLine  = (raw.match(/MATCH:\s*(YES|NO)/i)  || [])[1] || "NO";
  var rLine  = (raw.match(/REASON:\s*(.+)/i)      || [])[1] || "";
  return { match: mLine.toUpperCase(), reason: rLine.trim().slice(0, 150) };
}
