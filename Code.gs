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
  settings:           "AppSettings"
};

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
        row[headers[j]] = isNaN(v.getTime()) ? "" : v.toISOString().slice(0, 10);
      } else if (v === null || v === undefined) {
        row[headers[j]] = "";
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
    return (v === undefined || v === null) ? "" : v;
  });
  sheet.appendRow(rowArr);
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
  else if (finalScore >= 90) ratingLabel = "Nailed It!";
  else if (finalScore >= 80) ratingLabel = "Almost There!";
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
  return getSheetData(TABS.agents);
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
  return {
    qa_id:          qaId,
    transaction_no: txNo,
    final_score:    scores.final_score,
    rating_label:   scores.rating_label,
    has_ko:         scores.has_ko
  };
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
  return { dsat_id: row.dsat_id, validation_status: validationStatus };
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
  var allVal       = getDSATValidations({ agentName: agentName, month: month, year: year });
  var allImp       = getDSATImports({ agentName: agentName, month: month, year: year });
  var validDsats   = allVal.filter(function(r){ return r.validation_status === "valid"; });
  var invalidDsats = allVal.filter(function(r){ return r.validation_status === "invalid"; });

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
    dsat_pending:        allImp.length - allVal.length,
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
    var rl = avg >= 90 ? "Nailed It!" : avg >= 80 ? "Almost There!" : "Not Quite There!";
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
    agents_above_90:  ranked.filter(function(r){ return r.avg_score >= 90; }).length,
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
      var updatedRow = [
        dsatId,
        values[i][1], // import_id
        values[i][2], // date
        values[i][3], // month
        values[i][4], // year
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
  var rl = summary.avg_score >= 90 ? "Nailed It!" : summary.avg_score >= 80 ? "Almost There!" : "Not Quite There!";
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
