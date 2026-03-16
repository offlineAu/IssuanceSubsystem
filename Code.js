const SPREADSHEET_ID = "1p865Qx9OSNj0E2yU_z4YYz3kJbnT50E_zH9Ll1euGPs";
const SHEET_NAME = "Database";
const HISTORY_SHEET_NAME = "IssuanceHistory";

const COLUMNS_TO_SHOW = [
  "ID",
  "Name of the Student",
  "HEI",

  // Verification
  "Verification: Date Received",
  "Verification: Date Accomplished",
  "Verification Status",
  "Verification: Date Forwarded",

  // Issuance
  "Issuance and Encoding of No.: Assigned Person",
  "Issuance and Encoding of No.: Date Received",
  "Issuance and Encoding of No.: Date Accomplished",
  "Issuance and Encoding of No.: Status",
  "Issuance and Encoding of No.: Remarks",
  "Issuance and Encoding of No.: Date Forwarded",
  "SO No.",
  "Series",
  "Date of SO Issuance"
];

const HISTORY_HEADERS = [
  "Timestamp",
  "Row",
  "ID",
  "Name of the Student",
  "Issuance and Encoding of No.: Date Received",
  "Issuance and Encoding of No.: Date Accomplished",
  "Issuance and Encoding of No.: Status",
  "Issuance and Encoding of No.: Remarks",
  "Issuance and Encoding of No.: Date Forwarded",
  "SO No.",
  "Series",
  "Date of SO Issuance",
  "Editor"
];

function formatDateForClient_(value) {
  if (!(value instanceof Date)) return value || "";
  return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseIsoDate_(value) {
  if (!value || typeof value !== "string") return "";
  const parts = value.split("-");
  if (parts.length !== 3) return "";

  const year = Number(parts[0]);
  const month = Number(parts[1]);
  const day = Number(parts[2]);
  if (!year || !month || !day) return "";

  return new Date(year, month - 1, day);
}

function getOrCreateHistorySheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(HISTORY_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(HISTORY_SHEET_NAME);
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HISTORY_HEADERS.length).setValues([HISTORY_HEADERS]);
  }
  return sheet;
}

function appendIssuanceHistory_(history) {
  const historySheet = getOrCreateHistorySheet_();
  const editor = Session.getActiveUser().getEmail();
  const row = [
    new Date(),
    history.row || "",
    history.id || "",
    history.name || "",
    parseIsoDate_(history.dateReceived) || "",
    parseIsoDate_(history.dateAccomplished) || "",
    history.status || "",
    history.remarks || "",
    parseIsoDate_(history.dateForwarded) || "",
    history.soNo || "",
    history.series || "",
    parseIsoDate_(history.dateSOIssuance) || "",
    editor
  ];
  historySheet.appendRow(row);
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Document Issuance Subsystem')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData(batchSize = 1000) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { data: [], total: 0 };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const columnIndexes = COLUMNS_TO_SHOW.map(h => {
    const idx = headers.indexOf(h);
    if (idx === -1) Logger.log("Warning: Column not found: " + h);
    return idx;
  });

  const numRows = Math.min(batchSize, lastRow - 1);
  const rows = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();

  const startRow = 2;

  return {
    data: rows.map((row, index) => {
      const obj = {};
      columnIndexes.forEach((colIndex, i) => {
        let val = row[colIndex];
        val = formatDateForClient_(val);
        obj[COLUMNS_TO_SHOW[i]] = val || "";
      });

      obj._row = startRow + index; // 🔥 store actual sheet row
      return obj;
    }),
    total: lastRow - 1
  };
}

function loadRemainingData(startRow = 1001, batchSize = 20000) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (startRow > lastRow) return { data: [] };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const columnIndexes = COLUMNS_TO_SHOW.map(h => {
    const idx = headers.indexOf(h);
    if (idx === -1) Logger.log("Warning: Column not found: " + h);
    return idx;
  });

  const numRows = Math.min(batchSize, lastRow - startRow + 1);
  const rows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  return {
    data: rows.map((row, index) => {
      const obj = {};
      columnIndexes.forEach((colIndex, i) => {
        let val = row[colIndex];
        val = formatDateForClient_(val);
        obj[COLUMNS_TO_SHOW[i]] = val || "";
      });

      obj._row = startRow + index; // 🔥 store correct row
      return obj;
    })
  };
}

function searchExactName(searchTerm) {
  if (!searchTerm) return [];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const nameIdx = headers.indexOf("Name of the Student");
  if (nameIdx === -1) return [];
  const columnIndexes = COLUMNS_TO_SHOW.map(h => headers.indexOf(h));

  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const matched = [];
  allData.forEach((row, idx) => {
    if (row[nameIdx] && row[nameIdx].toString().toLowerCase().includes(searchTerm.toLowerCase())) {
      matched.push({ row, rowNumber: idx + 2 });
    }
  });

  return matched.map(entry => {
    const obj = {};
    columnIndexes.forEach((colIndex, i) => {
      let val = entry.row[colIndex];
      val = formatDateForClient_(val);
      obj[COLUMNS_TO_SHOW[i]] = val || "";
    });
    obj._row = entry.rowNumber;
    return obj;
  });
}

// ---------------- UPDATE ISSUANCE ----------------
function updateIssuanceColumns(data) {
  Logger.log(data);
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0].map(h => h.toString().trim());

  const columnMapping = {
    dateReceived: "Issuance and Encoding of No.: Date Received",
    dateAccomplished: "Issuance and Encoding of No.: Date Accomplished",
    status: "Issuance and Encoding of No.: Status",
    remarks: "Issuance and Encoding of No.: Remarks",
    dateForwarded: "Issuance and Encoding of No.: Date Forwarded",
    soNo: "SO No.",
    series: "Series",
    dateSOIssuance: "Date of SO Issuance"
  };

  const convertDate = parseIsoDate_;

  const row = Number(data.row); // Ensure row is a number
  if (!row || isNaN(row)) {
    Logger.log("Error: Invalid row number provided: " + data.row);
    return;
  }

  const idIndex = headers.indexOf("ID");
  const nameIndex = headers.indexOf("Name of the Student");
  const sourceRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const studentId = idIndex > -1 ? String(sourceRow[idIndex] || "") : "";
  const studentName = nameIndex > -1 ? String(sourceRow[nameIndex] || "") : "";

  const dateFields = ["dateReceived", "dateAccomplished", "dateForwarded", "dateSOIssuance"];
  
  Object.keys(columnMapping).forEach(key => {
    const colIndex = headers.indexOf(columnMapping[key]);
    if (colIndex === -1) return;

    let value = data[key];
    if (dateFields.indexOf(key) !== -1) {
      value = convertDate(value);
    }

    sheet.getRange(row, colIndex + 1).setValue(value);
  });

  appendIssuanceHistory_({
    row: row,
    id: studentId,
    name: studentName,
    dateReceived: data.dateReceived || "",
    dateAccomplished: data.dateAccomplished || "",
    status: data.status || "",
    remarks: data.remarks || "",
    dateForwarded: data.dateForwarded || "",
    soNo: data.soNo || "",
    series: data.series || "",
    dateSOIssuance: data.dateSOIssuance || ""
  });
}

function getIssuanceHistory(studentId, studentName) {
  const idNeedle = String(studentId || "").trim().toLowerCase();
  const nameNeedle = String(studentName || "").trim().toLowerCase();

  const sheet = getOrCreateHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, HISTORY_HEADERS.length).getValues();
  const tz = Session.getScriptTimeZone();

  const mapped = values.map(function(r) {
    return {
      timestamp: formatDateForClient_(r[0]),
      row: String(r[1] || ""),
      id: String(r[2] || ""),
      name: String(r[3] || ""),
      dateReceived: formatDateForClient_(r[4]),
      dateAccomplished: formatDateForClient_(r[5]),
      status: String(r[6] || ""),
      remarks: String(r[7] || ""),
      dateForwarded: formatDateForClient_(r[8]),
      soNo: String(r[9] || ""),
      series: String(r[10] || ""),
      dateSOIssuance: formatDateForClient_(r[11]),
      dateSaved: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd HH:mm:ss") : ""
    };
  });

  const filtered = (!idNeedle && !nameNeedle)
    ? mapped
    : mapped.filter(function(item) {
        const idMatch = idNeedle ? item.id.toLowerCase() === idNeedle : false;
        const nameMatch = !idNeedle && nameNeedle ? item.name.toLowerCase() === nameNeedle : false;
        return idMatch || nameMatch;
      });

  filtered.sort(function(a, b) {
    return String(b.dateSaved || "").localeCompare(String(a.dateSaved || ""));
  });

  return filtered;
}

function getSavedSoPrefixes() {
  const raw = PropertiesService.getScriptProperties().getProperty("SO_PREFIXES_JSON");
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    Logger.log("Failed to parse SO prefixes: " + e);
    return [];
  }
}

function saveSoPrefix(prefix) {
  const value = String(prefix || "").trim();
  if (!value) return getSavedSoPrefixes();

  const existing = getSavedSoPrefixes().filter(p => String(p || "").trim());
  const map = {};
  const merged = [value].concat(existing).filter(p => {
    if (map[p]) return false;
    map[p] = true;
    return true;
  }).slice(0, 100);

  PropertiesService.getScriptProperties().setProperty("SO_PREFIXES_JSON", JSON.stringify(merged));
  return merged;
}
