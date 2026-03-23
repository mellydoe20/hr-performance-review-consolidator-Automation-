/**
 * ============================================================
 *  Performance Review → Master Sheet Automation
 *  Google Apps Script
 * ============================================================
 *
 *  HOW TO USE
 *  ----------
 *  1. Open Google Drive → New → More → Google Apps Script
 *  2. Paste this entire file into the editor (replace any existing code)
 *  3. Save (Ctrl+S) and name the project "PR Automation"
 *  4. Run setupMasterSheet() ONCE to create the master Google Sheet
 *  5. Upload your .xlsx review files into a Drive folder
 *  6. Run extractAllReviews() to process all files and populate the master sheet
 *
 *  OPTIONAL: Set a time-based trigger so it runs automatically
 *  → Run setupTrigger() once, and new files will be processed every night at midnight
 *
 * ============================================================
 */


// ── CONFIG ─────────────────────────────────────────────────
// Folder name in Google Drive where .xlsx review files will be dropped
const SOURCE_FOLDER_NAME   = "Performance Reviews - Inbox";

// Name of the master Google Sheet that will be created/updated
const MASTER_SHEET_NAME    = "Master Performance Reviews";

// Name of the tab inside the master sheet
const MASTER_TAB_NAME      = "All Reviews";

// Name of the log tab
const LOG_TAB_NAME         = "Processing Log";
// ───────────────────────────────────────────────────────────


/**
 * MASTER COLUMN HEADERS
 * These define every column in the master sheet.
 */
const HEADERS = [
  "Timestamp Extracted",
  "Source File Name",
  "Employee Name",
  "Position",
  "Manager Name",
  "Review Period",
  "Achievements to Date",
  "Yourself – Areas of Excellence",
  "Yourself – Challenges Faced",
  "Team – Areas of Excellence",
  "Team – Challenges Faced",
  "Manager – Areas of Excellence",
  "Manager – Challenges Faced",
  "Areas of Growth / Development",
  "Future Goals & Plan",
  "Individual Performance Score (out of 10)",
  "Individual Performance Score – Reason",
  "Employee Experience Score (out of 10)",
  "Employee Experience Score – Reason",
  "Manager Summary & Next Steps",
  "Additional HR Comments",
  "Employee Signature",
  "Employee Signature Date",
  "Manager Signature",
  "Manager Signature Date",
];


/**
 * Cell address mapping for each field in the .xlsx review template.
 * These correspond to the merged answer cells in the performance review form.
 *
 *  Row references (1-based, matching the xlsx):
 *   Row 9  → Employee Name (col D=4), Position (col I=9)
 *   Row 11 → Manager Name (col B=2),  Review Period (col G=7)
 *   Row 22 → Achievements
 *   Row 26 → Yourself Excellence
 *   Row 28 → Yourself Challenges
 *   Row 31 → Team Excellence
 *   Row 33 → Team Challenges
 *   Row 36 → Manager Excellence
 *   Row 38 → Manager Challenges
 *   Row 40 → Growth / Dev Areas
 *   Row 42 → Future Goals
 *   Row 44 → Individual Perf Score
 *   Row 46 → Individual Perf Reason
 *   Row 48 → Employee Exp Score
 *   Row 50 → Employee Exp Reason
 *   Row 54 → Manager Summary
 *   Row 64 → Employee Sig Date (col J=10)
 *   Row 71 → Manager Sig Date  (col J=10)
 *   Row 75 → HR Additional Comments
 */
const FIELD_MAP = {
  // ── EMPLOYEE DETAILS ──────────────────────────────────────
  employeeName:          { row: 8,  col: 4  },  // D8
  position:              { row: 8,  col: 9  },  // I8
  managerName:           { row: 10, col: 4  },  // D10
  reviewPeriod:          { row: 10, col: 9  },  // I10

  // ── PERFORMANCE REVIEW SECTIONS ───────────────────────────
  achievements:          { row: 22, col: 2  },  // B22
  yourselfExcellence:    { row: 26, col: 2  },  // B26
  yourselfChallenges:    { row: 28, col: 2  },  // B28
  teamExcellence:        { row: 31, col: 2  },  // B31
  teamChallenges:        { row: 33, col: 2  },  // B33
  managerExcellence:     { row: 36, col: 2  },  // B36
  managerChallenges:     { row: 38, col: 2  },  // B38
  growthAreas:           { row: 40, col: 2  },  // B40
  futureGoals:           { row: 42, col: 2  },  // B42
  indivPerfScore:        { row: 44, col: 2  },  // B44
  indivPerfReason:       { row: 46, col: 2  },  // B46
  empExpScore:           { row: 48, col: 2  },  // B48
  empExpReason:          { row: 50, col: 2  },  // B50
  managerSummary:        { row: 54, col: 2  },  // B54
  hrComments:            { row: 75, col: 2  },  // B75

  // ── SIGNATURES & DATES ────────────────────────────────────
  // Row 62: Employee signature at C62, date filled by employee at K62
  // Row 70: Manager signature at C70, date filled by manager at K70
  empSignature:          { row: 62, col: 3  },  // C62 — Employee Signature
  empSigDate:            { row: 62, col: 11 },  // K62 — Employee Date
  managerSignature:      { row: 70, col: 3  },  // C70 — Manager Signature
  managerSigDate:        { row: 70, col: 11 },  // K70 — Manager Date
};


// ── MAIN ENTRY POINT ────────────────────────────────────────

/**
 * Scans the source folder for unprocessed .xlsx review files,
 * extracts data from each, and appends rows to the master sheet.
 * Run this manually or via a time trigger.
 */
function extractAllReviews() {
  // Auto-run setup if the master sheet doesn't exist yet
  if (!findSpreadsheetByName(MASTER_SHEET_NAME)) {
    Logger.log("Master sheet not found — running setup automatically...");
    setupMasterSheet(true);
  }

  const masterSheet  = getOrCreateMasterSheet();
  const logSheet     = getOrCreateLogSheet(masterSheet.getParent());
  const sourceFolder = getOrCreateSourceFolder();

  const files = sourceFolder.getFilesByType(MimeType.MICROSOFT_EXCEL);
  const alreadyProcessed = getProcessedFileIds(logSheet);

  let processedCount = 0;
  let skippedCount   = 0;
  let errorCount     = 0;

  while (files.hasNext()) {
    const file = files.next();

    if (alreadyProcessed.has(file.getId())) {
      skippedCount++;
      continue;
    }

    try {
      const data = extractDataFromXlsx(file);
      appendRowToMaster(masterSheet, data, file.getName());
      logProcessed(logSheet, file, "SUCCESS");
      processedCount++;
    } catch (e) {
      logProcessed(logSheet, file, "ERROR: " + e.message);
      errorCount++;
    }
  }

  const summary = `Done. Processed: ${processedCount} | Skipped (already done): ${skippedCount} | Errors: ${errorCount}`;
  Logger.log(summary);
  try { SpreadsheetApp.getUi().alert("✅ Extraction Complete\n\n" + summary); } catch (e) { /* running via trigger — skip UI alert */ }
}


/**
 * Extracts one .xlsx file's review data by temporarily converting
 * it to a Google Sheet, reading the cells, then deleting the temp copy.
 */
function extractDataFromXlsx(file) {
  // Convert .xlsx → Google Sheet (temporary)
  const tempSheet = Drive.Files.copy(
    { title: "__temp_pr_" + file.getId(), mimeType: MimeType.GOOGLE_SHEETS },
    file.getId()
  );

  try {
    const ss  = SpreadsheetApp.openById(tempSheet.id);
    const tab = ss.getSheets()[0]; // First sheet of the review workbook

    const getValue = ({ row, col }) => {
      const val = tab.getRange(row, col).getValue();
      return val !== null && val !== undefined ? String(val).trim() : "";
    };

    return {
      employeeName:       getValue(FIELD_MAP.employeeName),
      position:           getValue(FIELD_MAP.position),
      managerName:        getValue(FIELD_MAP.managerName),
      reviewPeriod:       getValue(FIELD_MAP.reviewPeriod),
      achievements:       getValue(FIELD_MAP.achievements),
      yourselfExcellence: getValue(FIELD_MAP.yourselfExcellence),
      yourselfChallenges: getValue(FIELD_MAP.yourselfChallenges),
      teamExcellence:     getValue(FIELD_MAP.teamExcellence),
      teamChallenges:     getValue(FIELD_MAP.teamChallenges),
      managerExcellence:  getValue(FIELD_MAP.managerExcellence),
      managerChallenges:  getValue(FIELD_MAP.managerChallenges),
      growthAreas:        getValue(FIELD_MAP.growthAreas),
      futureGoals:        getValue(FIELD_MAP.futureGoals),
      indivPerfScore:     getValue(FIELD_MAP.indivPerfScore),
      indivPerfReason:    getValue(FIELD_MAP.indivPerfReason),
      empExpScore:        getValue(FIELD_MAP.empExpScore),
      empExpReason:       getValue(FIELD_MAP.empExpReason),
      managerSummary:     getValue(FIELD_MAP.managerSummary),
      hrComments:         getValue(FIELD_MAP.hrComments),
      empSignature:       getValue(FIELD_MAP.empSignature),
      empSigDate:         getValue(FIELD_MAP.empSigDate),
      managerSignature:   getValue(FIELD_MAP.managerSignature),
      managerSigDate:     getValue(FIELD_MAP.managerSigDate),
    };

  } finally {
    // Always delete the temp Google Sheet copy
    Drive.Files.remove(tempSheet.id);
  }
}


// ── MASTER SHEET HELPERS ────────────────────────────────────

/**
 * Appends a single extracted review as a new row in the master sheet.
 */
function appendRowToMaster(sheet, data, fileName) {
  const row = [
    new Date(),                  // Timestamp Extracted
    fileName,                    // Source File Name
    data.employeeName,
    data.position,
    data.managerName,
    data.reviewPeriod,
    data.achievements,
    data.yourselfExcellence,
    data.yourselfChallenges,
    data.teamExcellence,
    data.teamChallenges,
    data.managerExcellence,
    data.managerChallenges,
    data.growthAreas,
    data.futureGoals,
    data.indivPerfScore,
    data.indivPerfReason,
    data.empExpScore,
    data.empExpReason,
    data.managerSummary,
    data.hrComments,
    data.empSignature,
    data.empSigDate,
    data.managerSignature,
    data.managerSigDate,
  ];

  sheet.appendRow(row);

  // Wrap text for long-form fields (col 7 onwards)
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1, 1, row.length).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(lastRow, 21); // Keep rows compact by default
}


// ── SETUP FUNCTIONS ─────────────────────────────────────────

/**
 * Creates the Master Google Sheet with headers and formatting.
 * Run this ONCE before first use.
 */
function setupMasterSheet(silent = false) {
  const existing = findSpreadsheetByName(MASTER_SHEET_NAME);
  if (existing) {
    if (!silent) {
      try { SpreadsheetApp.getUi().alert(`⚠️ A sheet named "${MASTER_SHEET_NAME}" already exists.\n\nOpen it from your Drive. Run extractAllReviews() to process files.`); } catch(e) {}
    }
    return;
  }

  const ss  = SpreadsheetApp.create(MASTER_SHEET_NAME);
  const tab = ss.getActiveSheet().setName(MASTER_TAB_NAME);

  // Header row
  tab.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  // Header styling
  const headerRange = tab.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setFontFamily("Arial")
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#FFFFFF")
    .setBackground("#1a3c5e")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  tab.setRowHeight(1, 40);

  // Freeze header
  tab.setFrozenRows(1);
  tab.setFrozenColumns(3); // Freeze Timestamp, File, Employee Name

  // Column widths
  const colWidths = [
    160, // Timestamp
    220, // File Name
    160, // Employee Name
    160, // Position
    160, // Manager Name
    140, // Review Period
    300, // Achievements
    300, // Yourself Excellence
    300, // Yourself Challenges
    300, // Team Excellence
    300, // Team Challenges
    300, // Manager Excellence
    300, // Manager Challenges
    300, // Growth Areas
    300, // Future Goals
    80,  // Indiv Score
    300, // Indiv Reason
    80,  // Exp Score
    300, // Exp Reason
    300, // Manager Summary
    250, // HR Comments
    140, // Emp Sig Date
    140, // Manager Sig Date
  ];
  colWidths.forEach((w, i) => tab.setColumnWidth(i + 1, w));

  // Alternating row colour (applied via banding)
  const dataRange = tab.getRange(2, 1, 1000, HEADERS.length);
  const banding = dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  banding.setHeaderRowColor("#1a3c5e");

  // Log tab
  getOrCreateLogSheet(ss);

  // Create the source folder
  getOrCreateSourceFolder();

  const url = ss.getUrl();
  Logger.log('Setup complete! Sheet: ' + url);
  try {
    SpreadsheetApp.getUi().alert(
      `✅ Setup complete!\n\nMaster Sheet: ${MASTER_SHEET_NAME}\nSource Folder: "${SOURCE_FOLDER_NAME}" (created in My Drive)\n\nNext steps:\n1. Upload .xlsx review files into the "${SOURCE_FOLDER_NAME}" folder\n2. Run extractAllReviews() to process them\n\nSheet URL: ${url}`
    );
  } catch(e) { /* silent when called internally */ }
}


/**
 * Sets up a nightly trigger to auto-process new files at midnight.
 * Run this once after setupMasterSheet().
 */
function setupTrigger() {
  // Remove any existing triggers for this function
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "extractAllReviews")
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("extractAllReviews")
    .timeBased()
    .everyDays(1)
    .atHour(0)
    .create();

  Logger.log('Nightly trigger set for extractAllReviews()');
  try {
    SpreadsheetApp.getUi().alert("⏰ Nightly trigger set!\n\nextractAllReviews() will run automatically every night at midnight.");
  } catch(e) {}
}


// ── UTILITY FUNCTIONS ───────────────────────────────────────

function getOrCreateMasterSheet() {
  let ss = findSpreadsheetByName(MASTER_SHEET_NAME);
  if (!ss) {
    // Shouldn't reach here (extractAllReviews auto-setups), but just in case
    setupMasterSheet();
    ss = findSpreadsheetByName(MASTER_SHEET_NAME);
  }
  return ss.getSheetByName(MASTER_TAB_NAME) || ss.getSheets()[0];
}

function getOrCreateLogSheet(ss) {
  let log = ss.getSheetByName(LOG_TAB_NAME);
  if (!log) {
    log = ss.insertSheet(LOG_TAB_NAME);
    log.getRange(1, 1, 1, 4).setValues([["Timestamp", "File Name", "File ID", "Status"]]);
    log.getRange(1, 1, 1, 4)
      .setFontWeight("bold")
      .setBackground("#f0f0f0")
      .setFontFamily("Arial");
    log.setColumnWidth(1, 160);
    log.setColumnWidth(2, 260);
    log.setColumnWidth(3, 240);
    log.setColumnWidth(4, 300);
    log.setFrozenRows(1);
  }
  return log;
}

function getOrCreateSourceFolder() {
  const folders = DriveApp.getFoldersByName(SOURCE_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  const folder = DriveApp.createFolder(SOURCE_FOLDER_NAME);
  Logger.log("Created source folder: " + folder.getUrl());
  return folder;
}

function findSpreadsheetByName(name) {
  const files = DriveApp.getFilesByName(name);
  while (files.hasNext()) {
    const f = files.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) {
      return SpreadsheetApp.openById(f.getId());
    }
  }
  return null;
}

function getProcessedFileIds(logSheet) {
  const data = logSheet.getDataRange().getValues();
  const ids  = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][3] === "SUCCESS") ids.add(data[i][2]);
  }
  return ids;
}

function logProcessed(logSheet, file, status) {
  logSheet.appendRow([new Date(), file.getName(), file.getId(), status]);
}


// ── MENU ────────────────────────────────────────────────────

/**
 * Adds a custom menu to the Google Sheet UI.
 * Automatically appears when the master sheet is opened.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📋 PR Automation")
    .addItem("▶ Extract New Reviews", "extractAllReviews")
    .addSeparator()
    .addItem("⚙️ Setup Master Sheet (first time)", "setupMasterSheet")
    .addItem("⏰ Set Nightly Auto-Trigger",         "setupTrigger")
    .addToUi();
}
