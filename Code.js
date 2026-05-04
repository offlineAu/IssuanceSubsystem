const spreadsheetID = "1p865Qx9OSNj0E2yU_z4YYz3kJbnT50E_zH9Ll1euGPs";
const BIGQUERY_PROJECT_ID = "sodium-shard-459702-t9";
const BIGQUERY_LOCATION = "US";
const BIGQUERY_DATASET_ID = "executive_dataset";
const BIGQUERY_DATABASE_FLAT_TABLE = "`sodium-shard-459702-t9.executive_dataset.database_flat`";
const BIGQUERY_NATIVE_REPORTING_VIEW = "`sodium-shard-459702-t9.executive_dataset.reporting_flat_native_view`";
const BIGQUERY_HOLIDAYS_TABLE = "`sodium-shard-459702-t9.executive_dataset.ph_holidays`";
const databaseSheetName = "Database";
const historySheetName = "IssuanceHistory";
const saveFailureSheetName = "IssuanceSaveFailures";
const accountsSheetName = "Accounts";
const usersSheetName = "Users";
const userAccessSheetName = "User Access";
const dataTypesSheetName = "Data Types";
const navbarLogoFileId = "1BZo5fBsJ7zYW69ieDIZW1_-E3Xmqc2zn";

// ========== Executive Dashboard Constants ==========
const DASHBOARD_SPREADSHEET_ID = spreadsheetID;
const DASHBOARD_STATIC_SHEET = "Static";
const DASHBOARD_USERS_SHEET = usersSheetName;
const DASHBOARD_DATA_TYPES_SHEET = dataTypesSheetName;
const DASHBOARD_REPORTING_FLAT_SHEET = "Reporting_Flat";
const DASHBOARD_INITIAL_BATCH_SIZE = 1000;
const DASHBOARD_FLAT_REBUILD_INTERVAL_MINUTES = 30;
const DASHBOARD_FLAT_REBUILD_BATCH_SIZE = 2500;
const DASHBOARD_FLAT_REBUILD_EXECUTION_BUDGET_MS = 240000;
const DASHBOARD_FLAT_REBUILD_STATE_KEY = "dashboard_reporting_flat_rebuild_state";
const DASHBOARD_REFERENCE_CACHE_TTL_SECONDS = 1800; // 30 minutes
const DASHBOARD_SECTION_CACHE_TTL_SECONDS = 600; // 10 minutes
const BIGQUERY_CACHE_VERSION = "v2";
const DASHBOARD_DEBUG = false;

const DASHBOARD_STATUS_ORDER = [
  "For Evaluation",
  "Evaluated",
  "Deficiency",
  "Approved",
  "Verified",
  "For Signature",
  "For Release",
  "Released",
  "Cancelled"
];

const DASHBOARD_REPORTING_FLAT_HEADERS = [
  "_row",
  "recordId",
  "studentName",
  "program",
  "es2InCharge",
  "es2InChargeRaw",
  "applicationStatus",
  "statusOfApplication",
  "transactionType",
  "hei",
  "dateReceived",
  "dateReceivedDate",
  "targetReleaseDate",
  "targetReleaseDateDate",
  "evaluationStatus",
  "evaluationAssignedPeopleText",
  "evaluationDateReceived",
  "evaluationDateAccomplished",
  "verificationStatus",
  "verificationDateReceived",
  "verificationDateAccomplished",
  "issuanceStatus",
  "issuanceAssignedPeopleText",
  "issuanceDateReceived",
  "issuanceDateAccomplished",
  "deficiencyStatus",
  "deficiencyDateReceived",
  "deficiencyDateAccomplished",
  "signatureStatus",
  "signatureAssignedPeopleText",
  "signatureDateReceived",
  "signatureDateAccomplished",
  "releaseStatus",
  "releaseAssignedPeopleText",
  "releaseDateReceived",
  "releaseDateAccomplished",
  "currentLocation",
  "canonicalStatus",
  "verificationDateReceived",
  "verificationDateReceivedDate",
  "issuanceDateReceived",
  "issuanceDateReceivedDate",
  "deficiencyDateReceived",
  "deficiencyDateReceivedDate",
  "releaseDateReceived",
  "releaseDateReceivedDate",
  "verificationDateAccomplished",
  "verificationDateAccomplishedDate",
  "issuanceDateAccomplished",
  "issuanceDateAccomplishedDate",
  "deficiencyDateAccomplished",
  "deficiencyDateAccomplishedDate",
  "signatureDateAccomplished",
  "signatureDateAccomplishedDate",
  "releaseDateAccomplished",
  "releaseDateAccomplishedDate",
  "hasPicture",
  "lastName",
  "firstName",
  "middleName",
  "suffix"
];

const columnsToShow = [
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

const identityColumnsToShow = [
  "ID",
  "Name of the Student",
  "HEI"
];

const developerTableColumns = [
  "ID",
  "Name of the Student",
  "HEI",
  "Verification: Date Received",
  "Verification: Date Accomplished",
  "Verification Status",
  "Verification: Date Forwarded",
  "Issuance and Encoding of No.: Assigned Person"
];

const typingTableColumns = [
  "Date Received",
  "Target Release Date",
  "Name of the Student",
  "Transaction Type",
  "HEI",
  "HEI: Address",
  "Date of Graduation",
  "Program",
  "ES II In-charge of the Program",
  "OR No.",
  "OR: Date Issued",
  "Status of the Application",
  "DATE RETURN BY CHA CHA"
];

const historyHeaders = [
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

const saveFailureHeaders = [
  "Timestamp",
  "Debug ID",
  "Row",
  "ID",
  "Name of the Student",
  "Payload JSON",
  "Error Code",
  "Message",
  "Details",
  "Editor",
  "Status",
  "Retry Timestamp",
  "Retry Result",
  "Retry Debug ID"
];

const editorsHeaders = [
  "Email",
  "Name",
  "Permission"
];

const usersHeaders = [
  "ID",
  "Name",
  "Position",
  "Email",
  "Access Type"
];

const issuanceFieldConfig = {
  dateReceived: {
    columnName: "Issuance and Encoding of No.: Date Received",
    label: "Date Received",
    formFieldKey: "dateReceived",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  dateAccomplished: {
    columnName: "Issuance and Encoding of No.: Date Accomplished",
    label: "Date Accomplished",
    formFieldKey: "dateAccomplished",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  status: {
    columnName: "Issuance and Encoding of No.: Status",
    label: "Status",
    formFieldKey: "status",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  remarks: {
    columnName: "Issuance and Encoding of No.: Remarks",
    label: "Remarks",
    formFieldKey: "remarks",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  dateForwarded: {
    columnName: "Issuance and Encoding of No.: Date Forwarded",
    label: "Date Forwarded",
    formFieldKey: "dateForwarded",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  soNo: {
    columnName: "SO No.",
    label: "SO No.",
    formFieldKey: "soNo",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  series: {
    columnName: "Series",
    label: "Series",
    formFieldKey: "series",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  },
  dateSOIssuance: {
    columnName: "Date of SO Issuance",
    label: "Date of SO Issuance",
    formFieldKey: "dateSOIssuance",
    assignmentColumn: "Issuance and Encoding of No.: Assigned Person"
  }
};

const issuanceFieldKeys = Object.keys(issuanceFieldConfig);
const typingCreateFieldConfig = {
  dateReceived: {
    columnName: "Date Received",
    type: "date"
  },
  studentName: {
    columnName: "Name of the Student",
    type: "text"
  },
  hei: {
    columnName: "HEI",
    type: "text"
  },
  dateOfGraduation: {
    columnName: "Date of Graduation",
    type: "date"
  },
  program: {
    columnName: "Program",
    type: "text"
  }
};
const typingCreateFieldKeys = Object.keys(typingCreateFieldConfig);
const developerFallbackEmail = "janoplo.airljoriz@gmail.com";

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
  const ss = SpreadsheetApp.openById(spreadsheetID);
  let sheet = ss.getSheetByName(historySheetName);
  if (!sheet) {
    sheet = ss.insertSheet(historySheetName);
    sheet.getRange(1, 1, 1, historyHeaders.length).setValues( [historyHeaders]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, historyHeaders.length).setValues( [historyHeaders]);
  }
  return sheet;
}

function getOrCreateSaveFailureSheet_() {
  const ss = SpreadsheetApp.openById(spreadsheetID);
  let sheet = ss.getSheetByName(saveFailureSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(saveFailureSheetName);
    sheet.getRange(1, 1, 1, saveFailureHeaders.length).setValues([saveFailureHeaders]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, saveFailureHeaders.length).setValues([saveFailureHeaders]);
  }
  return sheet;
}

function getEditorEmail_() {
  const sessionState = getSessionEmailState_();
  return sessionState.resolvedEmail;
}

function getResolvedSessionEmail_() {
  let activeEmail = "";
  let effectiveEmail = "";
  try {
    activeEmail = Session.getActiveUser().getEmail() || "";
  } catch (e) {
    activeEmail = "";
  }
  if (activeEmail) {
    Logger.log("[AUTH OK] Active session email resolved: " + activeEmail);
    return {
      activeEmail: activeEmail,
      effectiveEmail: "",
      resolvedEmail: activeEmail,
      source: "active",
      usedDeveloperFallback: false
    };
  }

  try {
    effectiveEmail = Session.getEffectiveUser().getEmail() || "";
  } catch (e2) {
    effectiveEmail = "";
  }
  if (effectiveEmail) {
    return {
      activeEmail: "",
      effectiveEmail: effectiveEmail,
      resolvedEmail: effectiveEmail,
      source: "effective",
      usedDeveloperFallback: false
    };
  }

  Logger.log("[AUTH FALLBACK] Using developer email fallback: " + developerFallbackEmail);
  return {
    activeEmail: "",
    effectiveEmail: "",
    resolvedEmail: developerFallbackEmail,
    source: "developer_fallback",
    usedDeveloperFallback: true
  };
}

function getSessionEmailState_() {
  return getResolvedSessionEmail_();
}

function getEffectiveEmail_() {
  try {
    return Session.getEffectiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function createDebugId_() {
  return "DBG-" + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function toSafeString_(value) {
  return value == null ? "" : String(value);
}

function normalizeEmail_(value) {
  return String(value || "").trim().toLowerCase();
}

function isValidEmail_(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || "").trim());
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) {
    return String(h || "").trim();
  });
}

function normalizeHeaderKey_(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function findHeaderIndexByAliases_(headers, aliases) {
  const normalizedHeaders = (headers || []).map(normalizeHeaderKey_);
  for (let i = 0; i < aliases.length; i++) {
    const idx = normalizedHeaders.indexOf(normalizeHeaderKey_(aliases[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function getSheetDataByName_(sheetName) {
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    return {
      sheet: sheet,
      headers: [],
      rows: []
    };
  }
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function(value) { return String(value || "").trim(); });
  return {
    sheet: sheet,
    headers: headers,
    rows: values.slice(1)
  };
}

function buildRowObject_(headers, row) {
  const obj = {};
  (headers || []).forEach(function(header, index) {
    obj[header] = row[index];
  });
  return obj;
}

function parseOperationFlags_(value) {
  const normalized = String(value || "").trim().toUpperCase();
  return {
    create: normalized.indexOf("C") !== -1,
    read: normalized.indexOf("R") !== -1,
    update: normalized.indexOf("U") !== -1,
    "delete": normalized.indexOf("D") !== -1,
    raw: normalized
  };
}

function parseBooleanFlag_(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return ["true", "yes", "y", "1", "x", "checked"].indexOf(normalized) !== -1;
}

function normalizeScope_(value) {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "assigned") return "assigned";
  return "all";
}

function extractEmails_(value) {
  const matches = String(value || "").toLowerCase().match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/g);
  return /** @type {string[]} */ (matches || []);
}

function getStageLabelFromFieldName_(fieldName) {
  const raw = String(fieldName || "").trim();
  if (!raw) return "";
  const parts = raw.split(":");
  return parts.length > 1 ? String(parts[0] || "").trim() : raw;
}

function getStudentIdentityByRow_(sheet, headers, row) {
  const sourceRow = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idIndex = headers.indexOf("ID");
  const nameIndex = headers.indexOf("Name of the Student");
  return {
    id: idIndex > -1 ? String(sourceRow[idIndex] || "") : "",
    name: nameIndex > -1 ? String(sourceRow[nameIndex] || "") : ""
  };
}

function getProtectedEditorEmails_() {
  const protectedEmails = {};
  const file = DriveApp.getFileById(spreadsheetID);
  const owner = file.getOwner();
  const ownerEmail = normalizeEmail_(owner && owner.getEmail ? owner.getEmail() : "");
  if (ownerEmail) protectedEmails[ownerEmail] = true;

  const effectiveEmail = normalizeEmail_(getEffectiveEmail_());
  if (effectiveEmail) protectedEmails[effectiveEmail] = true;

  const activeEmail = normalizeEmail_(getEditorEmail_());
  if (activeEmail) protectedEmails[activeEmail] = true;

  const raw = PropertiesService.getScriptProperties().getProperty("PROTECTED_EDITOR_EMAILS");
  if (raw) {
    raw.split(",").forEach(function(item) {
      const email = normalizeEmail_(item);
      if (email) protectedEmails[email] = true;
    });
  }

  return protectedEmails;
}

function getUserDirectorySource_() {
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const usersSheet = ss.getSheetByName(usersSheetName);
  const accountsSheet = ss.getSheetByName(accountsSheetName);

  if (usersSheet && usersSheet.getLastColumn() > 0) {
    const usersHeadersRow = getHeaders_(usersSheet);
    if (
      usersHeadersRow.indexOf("Email") !== -1 &&
      usersHeadersRow.indexOf("Name") !== -1 &&
      usersHeadersRow.indexOf("Access Type") !== -1
    ) {
      return {
        type: "users",
        sheetName: usersSheetName,
        sheet: usersSheet,
        headers: usersHeadersRow
      };
    }
  }

  const accounts = getOrCreateAccountsSheet_();
  return {
    type: "accounts",
    sheetName: accountsSheetName,
    sheet: accounts,
    headers: accounts.getLastColumn() > 0 ? getHeaders_(accounts) : []
  };
}

function getManagedEditorEmailsFromDirectory_() {
  const source = getUserDirectorySource_();
  const sheet = source.sheet;
  const data = sheet.getDataRange().getValues();
  if (!data.length) {
    return {
      sourceType: source.type,
      sourceSheetName: source.sheetName,
      canEditEmails: [],
      listedEmails: [],
      invalidEntries: []
    };
  }

  const headers = source.headers && source.headers.length ? source.headers : data[0];
  const emailIdx = headers.indexOf("Email");
  const accessIdx = source.type === "users"
    ? headers.indexOf("Access Type")
    : headers.indexOf("Permission");
  if (emailIdx === -1 || accessIdx === -1) {
    return {
      sourceType: source.type,
      sourceSheetName: source.sheetName,
      canEditEmails: [],
      listedEmails: [],
      invalidEntries: ["Missing Email or access column in " + source.sheetName + " sheet."]
    };
  }

  const canEditMap = {};
  const listedMap = {};
  const invalidEntries = [];

  for (let i = 1; i < data.length; i++) {
    const rawEmail = String(data[i][emailIdx] || "").trim();
    const email = normalizeEmail_(rawEmail);
    const accessType = String(data[i][accessIdx] || "").trim().toLowerCase();
    if (!rawEmail) continue;
    if (!isValidEmail_(rawEmail)) {
      invalidEntries.push(rawEmail);
      continue;
    }
    listedMap[email] = true;
    if (source.type === "users") {
      const profile = resolveAccessProfile_(data[i][accessIdx]);
      const actions = profile ? profile.actions : null;
      if (actions && (actions.create || actions.update || actions["delete"])) {
        canEditMap[email] = true;
      }
      continue;
    }
    if (accessType === "can edit") {
      canEditMap[email] = true;
    }
  }

  return {
    sourceType: source.type,
    sourceSheetName: source.sheetName,
    canEditEmails: Object.keys(canEditMap),
    listedEmails: Object.keys(listedMap),
    invalidEntries: invalidEntries
  };
}

function getUserAccessDefinitions_() {
  const data = getSheetDataByName_(userAccessSheetName);
  const headers = data.headers || [];
  const rows = data.rows || [];
  const idIdx = findHeaderIndexByAliases_(headers, ["ID", "Access Type ID", "Access ID"]);
  const operationsIdx = findHeaderIndexByAliases_(headers, ["Operations", "Operation", "CRUD"]);
  const scopeIdx = findHeaderIndexByAliases_(headers, ["Access", "Scope"]);
  const detailIdx = findHeaderIndexByAliases_(headers, ["Detail", "Details", "Description"]);
  const privilegedIdx = findHeaderIndexByAliases_(headers, ["Is Privileged", "Privileged", "Bypass Rules", "Bypass Position", "Bypass Assignment"]);

  return rows.map(function(row) {
    const id = idIdx > -1 ? String(row[idIdx] || "").trim() : "";
    const operations = operationsIdx > -1 ? String(row[operationsIdx] || "").trim() : "";
    const scope = scopeIdx > -1 ? String(row[scopeIdx] || "").trim() : "";
    const detail = detailIdx > -1 ? String(row[detailIdx] || "").trim() : "";
    const flags = parseOperationFlags_(operations);
    return {
      id: id,
      operations: operations,
      scope: normalizeScope_(scope),
      detail: detail,
      isPrivileged: privilegedIdx > -1 ? parseBooleanFlag_(row[privilegedIdx]) : false,
      actions: {
        create: flags.create,
        read: flags.read,
        update: flags.update,
        "delete": flags["delete"]
      }
    };
  }).filter(function(profile) {
    return profile.id || profile.operations;
  });
}

function resolveAccessProfile_(accessTypeValue) {
  const needle = String(accessTypeValue || "").trim();
  if (!needle) return null;
  const profiles = getUserAccessDefinitions_();
  for (let i = 0; i < profiles.length; i++) {
    if (String(profiles[i].id || "").trim() === needle) return profiles[i];
  }
  return null;
}

function getUsersDirectoryData_() {
  const data = getSheetDataByName_(usersSheetName);
  const headers = data.headers || [];
  return {
    headers: headers,
    rows: data.rows || [],
    emailIdx: findHeaderIndexByAliases_(headers, ["Email", "E-mail"]),
    nameIdx: findHeaderIndexByAliases_(headers, ["Name", "Full Name"]),
    positionIdx: findHeaderIndexByAliases_(headers, ["Position", "Role Position"]),
    accessTypeIdx: findHeaderIndexByAliases_(headers, ["Access Type", "User Access", "Access Profile"])
  };
}

function getUserRowByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  const directory = getUsersDirectoryData_();
  if (!normalizedEmail || directory.emailIdx === -1) return null;
  for (let i = 0; i < directory.rows.length; i++) {
    const rowEmail = normalizeEmail_(directory.rows[i][directory.emailIdx]);
    if (rowEmail === normalizedEmail) {
      return {
        headers: directory.headers,
        row: directory.rows[i],
        rowObject: buildRowObject_(directory.headers, directory.rows[i]),
        name: directory.nameIdx > -1 ? String(directory.rows[i][directory.nameIdx] || "").trim() : "",
        position: directory.positionIdx > -1 ? String(directory.rows[i][directory.positionIdx] || "").trim() : "",
        accessType: directory.accessTypeIdx > -1 ? String(directory.rows[i][directory.accessTypeIdx] || "").trim() : ""
      };
    }
  }
  return null;
}

function getDataTypeFieldMap_() {
  const definitions = getDataTypeFieldDefinitions_();
  const map = {};
  definitions.forEach(function(item) {
    map[item.name] = item;
  });
  return map;
}

function getDataTypeFieldDefinitions_() {
  const data = getSheetDataByName_(dataTypesSheetName);
  const headers = data.headers || [];
  const nameIdx = findHeaderIndexByAliases_(headers, ["Name", "Field Name"]);
  const typeIdx = findHeaderIndexByAliases_(headers, ["Type"]);
  const optionsIdx = findHeaderIndexByAliases_(headers, ["Select Options (For Select and Multi- Select type only)", "Select Options", "Options"]);
  const assignedIdIdx = findHeaderIndexByAliases_(headers, ["Assigned ID", "AssignedID"]);
  const classIdx = findHeaderIndexByAliases_(headers, ["Class"]);
  const positionIdx = findHeaderIndexByAliases_(headers, ["Position"]);
  return (data.rows || []).map(function(row, index) {
    const name = nameIdx > -1 ? String(row[nameIdx] || "").trim() : "";
    if (!name) return null;
    const rawOptions = optionsIdx > -1 ? String(row[optionsIdx] || "").trim() : "";
    return {
      name: name,
      type: typeIdx > -1 ? String(row[typeIdx] || "").trim() : "",
      selectOptions: rawOptions ? rawOptions.split(",").map(function(item) { return String(item || "").trim(); }).filter(Boolean) : [],
      fieldId: assignedIdIdx > -1 ? String(row[assignedIdIdx] || "").trim() : "",
      className: classIdx > -1 ? String(row[classIdx] || "").trim() : "",
      position: positionIdx > -1 ? String(row[positionIdx] || "").trim() : "",
      order: index
    };
  }).filter(Boolean);
}

function isCurrentIssuanceFieldName_(fieldName) {
  const field = String(fieldName || "").trim();
  if (!field) return false;
  for (let i = 0; i < issuanceFieldKeys.length; i++) {
    const config = issuanceFieldConfig[issuanceFieldKeys[i]];
    if (config && config.columnName === field) return true;
  }
  return field === "Issuance and Encoding of No.: Assigned Person";
}

function getFieldCrudCapabilities_(actions, fieldMeta) {
  const safeActions = actions || { create: false, read: true, update: false, "delete": false };
  const normalizedType = String((fieldMeta && fieldMeta.type) || "").trim().toLowerCase();
  const isReadonlyType = normalizedType === "readonly";
  return {
    create: !!safeActions.create && !isReadonlyType,
    read: !!safeActions.read,
    update: !!safeActions.update && !isReadonlyType,
    "delete": !!safeActions["delete"] && !isReadonlyType
  };
}

function isFieldEditableForPosition_(fieldMeta, position) {
  if (!fieldMeta) return false;
  const fieldPosition = String(fieldMeta.position || "").trim().toLowerCase();
  if (!fieldPosition || fieldPosition === "all") return true;
  return fieldPosition === String(position || "").trim().toLowerCase();
}

function isUserAssignedToValue_(assignedValue, userEmail) {
  const normalizedUserEmail = normalizeEmail_(userEmail);
  if (!normalizedUserEmail) return false;
  return extractEmails_(assignedValue).indexOf(normalizedUserEmail) !== -1;
}

function resolveIssuanceAssignmentFieldMeta_(fieldMeta, fieldMap) {
  if (!fieldMeta || !fieldMap) return null;
  const normalizedFieldId = normalizeHeaderKey_(fieldMeta.fieldId || "");
  const stageLabel = getStageLabelFromFieldName_(fieldMeta.name);
  const candidates = Object.keys(fieldMap).map(function(key) { return fieldMap[key]; }).filter(function(item) {
    return item && String(item.type || "").trim().toLowerCase() === "assign";
  });

  if (normalizedFieldId.indexOf("issuanceencode") === 0) {
    for (let i = 0; i < candidates.length; i++) {
      if (normalizeHeaderKey_(candidates[i].fieldId || "") === "issuanceencodeassigned") return candidates[i];
    }
  }

  for (let j = 0; j < candidates.length; j++) {
    if (getStageLabelFromFieldName_(candidates[j].name) === stageLabel) return candidates[j];
  }

  return null;
}

function isPrivilegedAccess_(profile, position) {
  const normalizedPosition = String(position || "").trim().toLowerCase();
  if (normalizedPosition === "developer") return true;
  return !!(profile && profile.isPrivileged);
}

function resolveUserContext_(email) {
  const normalizedEmail = normalizeEmail_(email);
  const userRecord = getUserRowByEmail_(normalizedEmail);
  const profile = userRecord ? resolveAccessProfile_(userRecord.accessType) : null;
  const position = userRecord ? userRecord.position : "";
  const normalizedPosition = String(position || "").trim().toLowerCase();
  const sessionState = getSessionEmailState_();
  const isDeveloperFallbackUser = !!sessionState.usedDeveloperFallback && normalizedEmail === normalizeEmail_(developerFallbackEmail);
  const isDeveloper = normalizedPosition === "developer" || isDeveloperFallbackUser;
  const actions = isDeveloper
    ? { create: true, read: true, update: true, "delete": true }
    : (profile ? profile.actions : { create: false, read: true, update: false, "delete": false });
  return {
    email: normalizedEmail,
    knownUser: !!userRecord,
    name: userRecord ? userRecord.name : "",
    position: position,
    accessType: userRecord ? userRecord.accessType : "",
    profileId: isDeveloper ? (profile && profile.id ? profile.id : "DEVELOPER_PRIVILEGE") : (profile ? profile.id : ""),
    detail: profile ? profile.detail : "",
    scope: isDeveloper ? "all" : (profile ? profile.scope : "all"),
    actions: actions,
    isPrivileged: isDeveloperFallbackUser || isPrivilegedAccess_(profile, position),
    denyReason: !normalizedEmail
      ? "Unable to determine the current signed-in user."
      : !userRecord
        ? "Your email is not registered in the Users sheet."
        : (!profile && !isDeveloper)
          ? "Your access type is not configured in the User Access sheet."
          : ""
  };
}

function getCurrentUserContext_() {
  return resolveUserContext_(getEditorEmail_());
}

function getFieldSchemaForCurrentUser_() {
  const sessionState = getSessionEmailState_();
  const userContext = resolveUserContext_(sessionState.resolvedEmail);
  const actions = userContext.actions || { create: false, read: true, update: false, "delete": false };
  const position = String(userContext.position || "").trim().toLowerCase();
  const isPrivileged = !!userContext.isPrivileged;
  const fields = getDataTypeFieldDefinitions_().filter(function(fieldMeta) {
    if (!userContext.knownUser || !userContext.profileId) return false;
    if (isPrivileged) return true;
    const fieldPosition = String(fieldMeta.position || "").trim().toLowerCase();
    return !fieldPosition || fieldPosition === "all" || fieldPosition === position;
  }).map(function(fieldMeta) {
    const stageLabel = getStageLabelFromFieldName_(fieldMeta.name);
    const capabilities = getFieldCrudCapabilities_(actions, fieldMeta);
    return {
      name: fieldMeta.name,
      type: fieldMeta.type,
      selectOptions: fieldMeta.selectOptions || [],
      fieldId: fieldMeta.fieldId,
      className: fieldMeta.className,
      position: fieldMeta.position,
      stageLabel: stageLabel,
      operations: capabilities,
      operationCode: [
        capabilities.create ? "C" : "",
        capabilities.read ? "R" : "",
        capabilities.update ? "U" : "",
        capabilities["delete"] ? "D" : ""
      ].join(""),
      isCurrentIssuanceField: isCurrentIssuanceFieldName_(fieldMeta.name),
      isPlaceholderOnly: !isCurrentIssuanceFieldName_(fieldMeta.name),
      order: fieldMeta.order
    };
  });

  return {
    sessionEmail: sessionState.resolvedEmail,
    source: sessionState.source,
    position: userContext.position,
    accessType: userContext.accessType,
    profileId: userContext.profileId,
    knownUser: userContext.knownUser,
    isPrivileged: userContext.isPrivileged,
    fields: fields,
    placeholderFields: fields.filter(function(field) { return field.isPlaceholderOnly; }),
    liveIssuanceFields: fields.filter(function(field) { return field.isCurrentIssuanceField; }),
    denyReason: userContext.denyReason || ""
  };
}

function getDeveloperRecordViewConfig_() {
  return {
    mode: "developer",
    supportsIssuanceEditor: true,
    showVerificationFilters: true,
    dataColumns: columnsToShow.slice(),
    tableColumns: developerTableColumns.map(function(name) {
      return {
        name: name,
        label: name === "Name of the Student" ? "Name" : (name === "Issuance and Encoding of No.: Assigned Person" ? "Assigned Person" : name),
        isIdentity: identityColumnsToShow.indexOf(name) !== -1,
        operations: { create: false, read: true, update: true, "delete": false }
      };
    }),
    modalFields: issuanceFieldKeys.map(function(fieldKey) {
      const config = issuanceFieldConfig[fieldKey];
      return {
        key: fieldKey,
        name: config.columnName,
        label: config.label,
        operations: { create: true, read: true, update: true, "delete": false },
        isIdentity: false
      };
    })
  };
}

function getPositionRecordViewConfig_() {
  const schema = getFieldSchemaForCurrentUser_();
  const fieldNames = {};
  const visibleFields = (schema.fields || []).filter(function(field) {
    if (!field || !field.name) return false;
    if (identityColumnsToShow.indexOf(field.name) !== -1) return false;
    if (fieldNames[field.name]) return false;
    fieldNames[field.name] = true;
    return true;
  });
  const tableColumns = identityColumnsToShow.map(function(name) {
    return {
      name: name,
      label: name === "Name of the Student" ? "Name" : name,
      isIdentity: true,
      operations: { create: false, read: true, update: false, "delete": false }
    };
  }).concat(visibleFields.map(function(field) {
    return {
      name: field.name,
      label: field.name,
      type: field.type,
      position: field.position,
      fieldId: field.fieldId,
      operations: field.operations || { create: false, read: true, update: false, "delete": false },
      isIdentity: false
    };
  }));

  return {
    mode: "position",
    supportsIssuanceEditor: false,
    showVerificationFilters: false,
    dataColumns: tableColumns.map(function(column) { return column.name; }),
    tableColumns: tableColumns,
    modalFields: visibleFields.map(function(field) {
      return {
        key: field.fieldId || field.name,
        name: field.name,
        label: field.name,
        type: field.type,
        selectOptions: field.selectOptions || [],
        position: field.position,
        className: field.className,
        operations: field.operations || { create: false, read: true, update: false, "delete": false },
        isIdentity: false
      };
    }),
    position: schema.position || "",
    profileId: schema.profileId || "",
    accessType: schema.accessType || "",
    knownUser: !!schema.knownUser,
    denyReason: schema.denyReason || ""
  };
}

function getTypingRecordViewConfig_() {
  return {
    mode: "typing",
    supportsIssuanceEditor: false,
    showVerificationFilters: false,
    dataColumns: typingTableColumns.slice(),
    tableColumns: typingTableColumns.map(function(name) {
      return {
        name: name,
        label: name,
        isIdentity: identityColumnsToShow.indexOf(name) !== -1,
        operations: { create: false, read: true, update: false, "delete": false }
      };
    }),
    modalFields: [
      { key: "dateReceived", name: "Date Received", label: "Date Received", type: "date", operations: { create: true, read: true, update: true, "delete": false } },
      { key: "studentName", name: "Name of the Student", label: "Name of the Student", type: "text", operations: { create: true, read: true, update: true, "delete": false } },
      { key: "hei", name: "HEI", label: "HEI", type: "text", operations: { create: true, read: true, update: true, "delete": false } },
      { key: "program", name: "Program", label: "Program", type: "text", operations: { create: true, read: true, update: true, "delete": false } },
      { key: "dateGraduation", name: "Date of Graduation", label: "Date of Graduation", type: "date", operations: { create: true, read: true, update: true, "delete": false } }
    ]
  };
}

function getCurrentRecordViewConfig_() {
  const userContext = getCurrentUserContext_();
  const normalizedPosition = String(userContext.position || "").trim().toLowerCase();
  if (normalizedPosition === "developer") {
    return getDeveloperRecordViewConfig_();
  }
  if (normalizedPosition === "releasing/verifier") {
    return getTypingRecordViewConfig_();
  }
  return getPositionRecordViewConfig_();
}

function mapDatabaseRowsForView_(headers, rows, startRow, visibleColumns) {
  const columns = Array.isArray(visibleColumns) ? visibleColumns : [];
  const columnIndexes = columns.map(function(name) {
    const idx = headers.indexOf(name);
    if (idx === -1) Logger.log("Warning: Column not found: " + name);
    return idx;
  });

  return (rows || []).map(function(row, index) {
    const obj = {};
    columnIndexes.forEach(function(colIndex, i) {
      let val = colIndex > -1 ? row[colIndex] : "";
      val = formatDateForClient_(val);
      obj[columns[i]] = val || "";
    });
    obj._row = startRow + index;
    return obj;
  });
}

function getCurrentUserDebugInfo() {
  const sessionState = getSessionEmailState_();
  const directory = getUsersDirectoryData_();
  const normalizedResolvedEmail = normalizeEmail_(sessionState.resolvedEmail);
  const userRecord = getUserRowByEmail_(sessionState.resolvedEmail);
  const profile = userRecord ? resolveAccessProfile_(userRecord.accessType) : null;
  const availableEmails = (directory.rows || []).map(function(row) {
    return normalizeEmail_(directory.emailIdx > -1 ? row[directory.emailIdx] : "");
  }).filter(Boolean);

  return {
    activeEmail: sessionState.activeEmail,
    effectiveEmail: sessionState.effectiveEmail,
    resolvedEmail: sessionState.resolvedEmail,
    resolvedSource: sessionState.source,
    normalizedResolvedEmail: normalizedResolvedEmail,
    userFound: !!userRecord,
    matchedName: userRecord ? userRecord.name : "",
    matchedPosition: userRecord ? userRecord.position : "",
    matchedAccessType: userRecord ? userRecord.accessType : "",
    profileFound: !!profile,
    profileId: profile ? profile.id : "",
    profileScope: profile ? profile.scope : "",
    profileOperations: profile ? profile.operations : "",
    isPrivileged: !!(profile && profile.isPrivileged),
    usersSheetEmailCount: availableEmails.length,
    usersSheetContainsResolvedEmail: availableEmails.indexOf(normalizedResolvedEmail) !== -1,
    sampleEmails: availableEmails.slice(0, 10),
    denyReason: !sessionState.resolvedEmail
      ? "No session email was resolved."
      : !userRecord
        ? "Resolved email was not found in Users."
        : !profile
          ? "Resolved user matched Users but Access Type did not match any User Access ID."
          : ""
  };
}

// Time-driven trigger function to refresh cache
function refreshDashboardCache() {
  Logger.log("[CACHE_REFRESH] Starting scheduled cache refresh");
  try {
    // Fetch dashboard data with default filters
    var payload = { datePreset: "thisYear" };
    var data = dashboard_fetchData_(payload);
    if (data && data.ok) {
      // Store in both cache layers
      setDashboardPropertiesCache_(data);
      var cacheKey = getDashboardCacheKey_(payload);
      setCachedDashboardData_(cacheKey, data);
      Logger.log("[CACHE_REFRESH] Successfully refreshed dashboard cache");
      return true;
    }
    Logger.log("[CACHE_REFRESH] Failed to fetch dashboard data");
    return false;
  } catch (e) {
    Logger.log("[CACHE_REFRESH] Error: " + e.toString());
    return false;
  }
}


function getUserPositionFieldSchema() {
  return getFieldSchemaForCurrentUser_();
}

function getIssuancePermissionContext_(rowNumber) {
  const userContext = getCurrentUserContext_();
  const fieldMap = getDataTypeFieldMap_();
  const context = {
    email: userContext.email,
    knownUser: userContext.knownUser,
    name: userContext.name,
    position: userContext.position,
    accessType: userContext.accessType,
    profileId: userContext.profileId,
    scope: userContext.scope,
    actions: userContext.actions,
    isPrivileged: userContext.isPrivileged,
    editableFieldKeys: [],
    readOnlyFieldKeys: issuanceFieldKeys.slice(),
    canUpdateAny: false,
    canCreate: !!userContext.actions.create,
    canRead: !!userContext.actions.read,
    canDelete: !!userContext.actions["delete"],
    isAssigned: false,
    canEditRecord: false,
    assignmentColumn: "",
    assignmentFieldId: "",
    assignmentValue: "",
    denyReason: userContext.denyReason || ""
  };

  issuanceFieldKeys.forEach(function(fieldKey) {
    const config = issuanceFieldConfig[fieldKey];
    const fieldMeta = fieldMap[config.columnName];
    if (context.actions.update && (context.isPrivileged || isFieldEditableForPosition_(fieldMeta, context.position))) {
      context.editableFieldKeys.push(fieldKey);
    }
  });
  context.readOnlyFieldKeys = issuanceFieldKeys.filter(function(fieldKey) {
    return context.editableFieldKeys.indexOf(fieldKey) === -1;
  });
  context.canUpdateAny = context.editableFieldKeys.length > 0;
  const firstEditableFieldKey = context.editableFieldKeys.length ? context.editableFieldKeys[0] : issuanceFieldKeys[0];
  const firstEditableFieldMeta = fieldMap[issuanceFieldConfig[firstEditableFieldKey].columnName];
  const assignmentFieldMeta = resolveIssuanceAssignmentFieldMeta_(firstEditableFieldMeta, fieldMap);
  if (assignmentFieldMeta) {
    context.assignmentColumn = assignmentFieldMeta.name || "";
    context.assignmentFieldId = assignmentFieldMeta.fieldId || "";
  }

  if (!context.knownUser) {
    return context;
  }
  if (!context.profileId) {
    return context;
  }
  if (!context.actions.update) {
    context.denyReason = "Your access type does not allow updates.";
    return context;
  }
  if (!context.canUpdateAny) {
    context.denyReason = "Your position does not allow editing issuance fields.";
    return context;
  }

  if (!rowNumber || Number(rowNumber) < 2) {
    context.canEditRecord = context.isPrivileged || context.scope === "all";
    if (!context.canEditRecord && context.scope === "assigned") {
      context.denyReason = "This access type requires an assigned record.";
    }
    return context;
  }

  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(databaseSheetName);
  if (!sheet || Number(rowNumber) > sheet.getLastRow()) {
    context.denyReason = "The selected database row no longer exists.";
    return context;
  }
  const headers = getHeaders_(sheet);
  const values = sheet.getRange(Number(rowNumber), 1, 1, sheet.getLastColumn()).getValues()[0];
  const assignmentIdx = context.assignmentColumn ? headers.indexOf(context.assignmentColumn) : -1;
  context.assignmentValue = assignmentIdx > -1 ? String(values[assignmentIdx] || "").trim() : "";
  context.isAssigned = isUserAssignedToValue_(context.assignmentValue, context.email);
  if (context.scope === "assigned" && !context.assignmentColumn && !context.isPrivileged) {
    context.denyReason = "No assignment field mapping exists for issuance updates.";
    return context;
  }
  context.canEditRecord = context.isPrivileged ? true : (context.scope === "all" ? true : context.isAssigned);
  if (!context.canEditRecord) {
    context.denyReason = context.scope === "assigned"
      ? "This record is not assigned to your email."
      : "You do not have permission to edit this record.";
  }
  return context;
}

function getUserAccessContext() {
  return getIssuancePermissionContext_("");
}

function buildAuthFailureResult_(row, code, message) {
  return buildSaveResult_(false, {
    row: row || "",
    message: message,
    code: code,
    debugId: createDebugId_()
  });
}

function requireEditorAccessResult_(row) {
  const access = getIssuancePermissionContext_(row);
  if (!access.email) {
    return buildAuthFailureResult_(row, "AUTH_USER_UNAVAILABLE", "Unable to determine the current signed-in user.");
  }
  if (!access.knownUser) {
    return buildAuthFailureResult_(row, "USER_NOT_FOUND", access.denyReason || "Your email is not registered in the Users sheet.");
  }
  if (!access.actions.update) {
    return buildAuthFailureResult_(row, "FORBIDDEN", access.denyReason || "You do not have permission to edit issuance records.");
  }
  if (!access.canUpdateAny) {
    return buildAuthFailureResult_(row, "FIELD_ACCESS_DENIED", access.denyReason || "Your position does not allow editing issuance fields.");
  }
  if (row && !access.canEditRecord) {
    return buildAuthFailureResult_(row, "ASSIGNMENT_REQUIRED", access.denyReason || "This record is not assigned to your email.");
  }
  return null;
}

function requirePermissionSyncAccessResult_() {
  const userEmail = normalizeEmail_(getEditorEmail_());
  if (!userEmail) {
    return buildAuthFailureResult_("", "AUTH_USER_UNAVAILABLE", "Unable to determine the current signed-in user.");
  }
  const protectedEmails = getProtectedEditorEmails_();
  const access = getIssuancePermissionContext_("");
  if (protectedEmails[userEmail] || access.actions.update || access.actions.create || access.actions["delete"]) {
    return null;
  }
  return buildAuthFailureResult_("", "FORBIDDEN", "You do not have permission to synchronize spreadsheet editors.");
}

function sanitizePayloadForLog_(data, identity) {
  return {
    row: Number(data && data.row) || "",
    id: identity && identity.id ? identity.id : "",
    name: identity && identity.name ? identity.name : "",
    dateReceived: toSafeString_(data && data.dateReceived),
    dateAccomplished: toSafeString_(data && data.dateAccomplished),
    status: toSafeString_(data && data.status),
    remarks: toSafeString_(data && data.remarks),
    dateForwarded: toSafeString_(data && data.dateForwarded),
    soNo: toSafeString_(data && data.soNo),
    series: toSafeString_(data && data.series),
    dateSOIssuance: toSafeString_(data && data.dateSOIssuance)
  };
}

function appendSaveFailureLog_(failure) {
  const sheet = getOrCreateSaveFailureSheet_();
  const row = [
    new Date(),
    failure.debugId || "",
    failure.row || "",
    failure.id || "",
    failure.name || "",
    JSON.stringify(failure.payload || {}),
    failure.code || "",
    failure.message || "",
    failure.details || "",
    getEditorEmail_(),
    failure.status || "failed",
    "",
    "",
    ""
  ];
  sheet.appendRow(row);
  return sheet.getLastRow();
}

function updateSaveFailureRetryStatus_(logRow, result) {
  const sheet = getOrCreateSaveFailureSheet_();
  const targetRow = Number(logRow);
  if (!targetRow || targetRow < 2 || targetRow > sheet.getLastRow()) return;
  sheet.getRange(targetRow, 11, 1, 4).setValues([[
    result && result.ok ? "retried-success" : "retried-failed",
    new Date(),
    result && result.message ? result.message : "",
    result && result.debugId ? result.debugId : ""
  ]]);
}

function buildSaveResult_(ok, payload) {
  const base = {
    ok: !!ok,
    row: payload && payload.row ? payload.row : "",
    id: payload && payload.id ? payload.id : "",
    name: payload && payload.name ? payload.name : "",
    message: payload && payload.message ? payload.message : "",
    code: payload && payload.code ? payload.code : "",
    debugId: payload && payload.debugId ? payload.debugId : createDebugId_()
  };
  if (payload && payload.details) base.details = payload.details;
  if (payload && payload.logRow) base.logRow = payload.logRow;
  return base;
}

function logSaveFailureWithResult_(result, payload) {
  const logRow = appendSaveFailureLog_({
    debugId: result.debugId,
    row: result.row,
    id: result.id,
    name: result.name,
    payload: sanitizePayloadForLog_(payload, { id: result.id, name: result.name }),
    code: result.code,
    message: result.message,
    details: result.details || "",
    status: "failed"
  });
  result.logRow = logRow;
  return result;
}

function performIssuanceSave_(data, options) {
  options = options || {};
  const debugId = createDebugId_();
  const payload = data || {};

  try {
    const authFailure = requireEditorAccessResult_(payload && payload.row);
    if (authFailure) return authFailure;

    const row = Number(payload.row);
    if (!row || isNaN(row) || row < 2) {
      return buildSaveResult_(false, {
        row: payload.row,
        message: "Invalid row number.",
        code: "INVALID_ROW",
        debugId: debugId
      });
    }

    const ss = SpreadsheetApp.openById(spreadsheetID);
    const sheet = ss.getSheetByName(databaseSheetName);
    if (!sheet) {
      return buildSaveResult_(false, {
        row: row,
        message: "Database sheet not found.",
        code: "MISSING_DATABASE_SHEET",
        debugId: debugId
      });
    }

    if (row > sheet.getLastRow()) {
      return buildSaveResult_(false, {
        row: row,
        message: "The selected database row no longer exists.",
        code: "ROW_NOT_FOUND",
        debugId: debugId
      });
    }

    const headers = getHeaders_(sheet);
    const identity = getStudentIdentityByRow_(sheet, headers, row);
    const access = getIssuancePermissionContext_(row);
    if (!access.canEditRecord) {
      return buildSaveResult_(false, {
        row: row,
        id: identity.id,
        name: identity.name,
        message: access.denyReason || "This record is not assigned to your email.",
        code: "ASSIGNMENT_REQUIRED",
        debugId: debugId
      });
    }

    const columnMapping = {};
    issuanceFieldKeys.forEach(function(fieldKey) {
      columnMapping[fieldKey] = issuanceFieldConfig[fieldKey].columnName;
    });
    const dateFields = ["dateReceived", "dateAccomplished", "dateForwarded", "dateSOIssuance"];
    const disallowedKeys = issuanceFieldKeys.filter(function(key) {
      return access.editableFieldKeys.indexOf(key) === -1;
    }).filter(function(key) {
      return Object.prototype.hasOwnProperty.call(payload, key);
    });
    if (disallowedKeys.length) {
      const deniedLabels = disallowedKeys.map(function(key) {
        return issuanceFieldConfig[key] ? issuanceFieldConfig[key].label : key;
      });
      return buildSaveResult_(false, {
        row: row,
        id: identity.id,
        name: identity.name,
        message: "You do not have permission to edit: " + deniedLabels.join(", ") + ".",
        code: "FIELD_ACCESS_DENIED",
        debugId: debugId
      });
    }

    Object.keys(columnMapping).forEach(function(key) {
      const colIndex = headers.indexOf(columnMapping[key]);
      if (colIndex === -1) {
        throw new Error("Missing required column: " + columnMapping[key]);
      }
      let value = payload[key];
      if (dateFields.indexOf(key) !== -1) {
        value = parseIsoDate_(value);
      }
      sheet.getRange(row, colIndex + 1).setValue(value);
    });

    markDatabaseRowUpdated_(row);

    try {
      appendIssuanceHistory_({
        row: row,
        id: identity.id,
        name: identity.name,
        dateReceived: payload.dateReceived || "",
        dateAccomplished: payload.dateAccomplished || "",
        status: payload.status || "",
        remarks: payload.remarks || "",
        dateForwarded: payload.dateForwarded || "",
        soNo: payload.soNo || "",
        series: payload.series || "",
        dateSOIssuance: payload.dateSOIssuance || ""
      });
    } catch (historyError) {
      const historyResult = buildSaveResult_(false, {
        row: row,
        id: identity.id,
        name: identity.name,
        message: "Saved to the database, but history logging failed.",
        code: "HISTORY_APPEND_FAILED",
        debugId: debugId,
        details: historyError && historyError.message ? historyError.message : String(historyError)
      });
      Logger.log(JSON.stringify({
        event: "issuance_save_failure",
        debugId: historyResult.debugId,
        code: historyResult.code,
        row: row,
        id: identity.id,
        details: historyResult.details
      }));
      return historyResult;
    }

    const successResult = buildSaveResult_(true, {
      row: row,
      id: identity.id,
      name: identity.name,
      message: "Saved successfully.",
      code: "SAVE_OK",
      debugId: debugId
    });
    Logger.log(JSON.stringify({
      event: "issuance_save_success",
      debugId: successResult.debugId,
      row: row,
      id: identity.id
    }));
    return successResult;
  } catch (error) {
    const failureResult = buildSaveResult_(false, {
      row: payload && payload.row ? payload.row : "",
      message: "Unable to save this student right now.",
      code: "SAVE_WRITE_FAILED",
      debugId: debugId,
      details: error && error.message ? error.message : String(error)
    });
    Logger.log(JSON.stringify({
      event: "issuance_save_failure",
      debugId: failureResult.debugId,
      code: failureResult.code,
      row: failureResult.row,
      details: failureResult.details
    }));
    return failureResult;
  }
}

function appendIssuanceHistory_(history) {
  const historySheet = getOrCreateHistorySheet_();
  const editor = getResolvedSessionEmail_().resolvedEmail;
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

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    console.error("Include failed for " + filename + ": " + e.toString());
    return "<!-- Include error: " + filename + " -->";
  }
}

function doGet() {
  const email = getEditorEmail_();
  const userContext = resolveUserContext_(email);
  const position = userContext ? String(userContext.position || "").trim().toLowerCase() : "";
  
  let templateName = "IssuanceIndex"; // Default for Verifier/Developer
  let title = "Document Issuance Subsystem";

  if (position === "es ii") {
    templateName = "DashboardIndex";
    title = "Executive Dashboard";
  } else if (position === "releasing/verifier" || position === "verifier") {
    // Both Releasing/Verifier and Verifier (if distinct from base Issuance Verifier) might need Typing
    // The prompt says: Releasing/Verifier -> TypingIndex, Verifier -> IssuanceIndex
    if (position === "releasing/verifier") {
      templateName = "TypingIndex";
      title = "Typing Records";
    }
  }

  return HtmlService.createTemplateFromFile(templateName)
    .evaluate()
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function legacy_getInitialData_(batchSize) {
  if (batchSize == null) batchSize = 1000;
  const viewConfig = getCurrentRecordViewConfig_();
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(databaseSheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { data: [], total: 0, viewConfig: viewConfig };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const numRows = Math.min(batchSize, lastRow - 1);
  const rows = sheet.getRange(2, 1, numRows, sheet.getLastColumn()).getValues();
  const startRow = 2;

  return {
    data: mapDatabaseRowsForView_(headers, rows, startRow, viewConfig.dataColumns),
    viewConfig: viewConfig,
    total: lastRow - 1
  };
}

function unwrapRouteWrapperResult_(action, routeResult) {
  var normalized = normalizeIssuanceExecutionResult_(routeResult, "route", action);
  return normalized.data;
}

function getServerRequestContext_(argsLike, payload, defaultModule) {
  var payloadContext = payload && payload.__context && typeof payload.__context === "object" ? payload.__context : null;
  if (payloadContext) return payloadContext;
  var args = Array.prototype.slice.call(argsLike || []);
  for (var i = args.length - 1; i >= 0; i--) {
    var candidate = args[i];
    if (candidate && typeof candidate === "object" && Object.prototype.hasOwnProperty.call(candidate, "module") && Object.prototype.hasOwnProperty.call(candidate, "requestId")) {
      return candidate;
    }
  }
  return {
    module: String(defaultModule || ""),
    requestId: "",
    timestamp: "",
    source: ""
  };
}

function logServerRequestContext_(context) {
  var safe = context || {};
  Logger.log("[SERVER CONTEXT] module = " + String(safe.module || "") + ", requestId = " + String(safe.requestId || ""));
}

function getInitialData(batchSize = 1000) {
  var serverContext = getServerRequestContext_(arguments, null, "issuance");
  logServerRequestContext_(serverContext);
  Logger.log("[ROUTE WRAPPER] getInitialData -> issuance.fetch");
  return unwrapRouteWrapperResult_("fetch", routeModule("issuance", "fetch", { batchSize: batchSize }));
}

function legacy_loadRemainingData_(startRow, batchSize) {
  if (startRow == null) startRow = 1001;
  if (batchSize == null) batchSize = 20000;
  const viewConfig = getCurrentRecordViewConfig_();
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(databaseSheetName);
  const lastRow = sheet.getLastRow();
  if (startRow > lastRow) return { data: [] };

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const numRows = Math.min(batchSize, lastRow - startRow + 1);
  const rows = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();

  return {
    data: mapDatabaseRowsForView_(headers, rows, startRow, viewConfig.dataColumns)
  };
}

function loadRemainingData(startRow = 1001, batchSize = 20000) {
  Logger.log("[ROUTE WRAPPER] loadRemainingData -> issuance.fetch");
  return unwrapRouteWrapperResult_("fetch", routeModule("issuance", "fetch", { startRow: startRow, batchSize: batchSize }));
}

function legacy_searchExactName_(searchTerm) {
  if (!searchTerm) return [];
  const viewConfig = getCurrentRecordViewConfig_();
  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(databaseSheetName);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
  const nameIdx = headers.indexOf("Name of the Student");
  if (nameIdx === -1) return [];

  const allData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const matched = [];
  allData.forEach((row, idx) => {
    if (row[nameIdx] && row[nameIdx].toString().toLowerCase().includes(searchTerm.toLowerCase())) {
      matched.push({ row, rowNumber: idx + 2 });
    }
  });

  return mapDatabaseRowsForView_(headers, matched.map(function(entry) { return entry.row; }), matched.length ? matched[0].rowNumber : 2, viewConfig.dataColumns).map(function(obj, index) {
    obj._row = matched[index].rowNumber;
    return obj;
  });
}

function searchExactName(searchTerm) {
  var serverContext = getServerRequestContext_(arguments, null, "issuance");
  logServerRequestContext_(serverContext);
  Logger.log("[ROUTE WRAPPER] searchExactName -> issuance.search");
  return unwrapRouteWrapperResult_("search", routeModule("issuance", "search", { searchTerm: searchTerm }));
}

// ---------------- UPDATE ISSUANCE ----------------
function legacy_updateIssuanceColumns_(data) {
  const result = performIssuanceSave_(data);
  if (!result.ok) {
    return logSaveFailureWithResult_(result, data);
  }
  return result;
}

function updateIssuanceColumns(data) {
  var serverContext = getServerRequestContext_(arguments, data, "issuance");
  logServerRequestContext_(serverContext);
  Logger.log("[ROUTE WRAPPER] updateIssuanceColumns -> issuance.update");
  return unwrapRouteWrapperResult_("update", routeModule("issuance", "update", data));
}

function legacy_getIssuanceHistory_(studentId, studentName) {
  const idNeedle = String(studentId || "").trim().toLowerCase();
  const nameNeedle = String(studentName || "").trim().toLowerCase();

  const sheet = getOrCreateHistorySheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, historyHeaders.length).getValues();
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

function getIssuanceHistory(studentId, studentName) {
  Logger.log("[ROUTE WRAPPER] getIssuanceHistory -> issuance.history");
  return unwrapRouteWrapperResult_("history", routeModule("issuance", "history", { studentId: studentId, studentName: studentName }));
}

function legacy_getSaveFailureHistory_() {
  const sheet = getOrCreateSaveFailureSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, saveFailureHeaders.length).getValues();
  const tz = Session.getScriptTimeZone();

  return values.map(function(r, index) {
    return {
      logRow: index + 2,
      timestamp: formatDateForClient_(r[0]),
      debugId: String(r[1] || ""),
      row: String(r[2] || ""),
      id: String(r[3] || ""),
      name: String(r[4] || ""),
      payloadJson: String(r[5] || ""),
      code: String(r[6] || ""),
      message: String(r[7] || ""),
      details: String(r[8] || ""),
      editor: String(r[9] || ""),
      status: String(r[10] || "failed"),
      retryTimestamp: r[11] instanceof Date ? Utilities.formatDate(r[11], tz, "yyyy-MM-dd HH:mm:ss") : String(r[11] || ""),
      retryResult: String(r[12] || ""),
      retryDebugId: String(r[13] || ""),
      dateSaved: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, "yyyy-MM-dd HH:mm:ss") : ""
    };
  }).sort(function(a, b) {
    return String(b.dateSaved || "").localeCompare(String(a.dateSaved || ""));
  });
}

function getSaveFailureHistory() {
  Logger.log("[ROUTE WRAPPER] getSaveFailureHistory -> issuance.failures");
  return unwrapRouteWrapperResult_("failures", routeModule("issuance", "failures", {}));
}

function legacy_logClientSaveFailure_(data, clientMessage) {
  const row = Number(data && data.row) || "";
  let identity = { id: "", name: "" };
  try {
    const ss = SpreadsheetApp.openById(spreadsheetID);
    const sheet = ss.getSheetByName(databaseSheetName);
    if (sheet && row && row >= 2 && row <= sheet.getLastRow()) {
      identity = getStudentIdentityByRow_(sheet, getHeaders_(sheet), row);
    }
  } catch (e) {}

  const result = buildSaveResult_(false, {
    row: row,
    id: identity.id,
    name: identity.name,
    message: toSafeString_(clientMessage) || "The save request did not complete.",
    code: "CLIENT_RPC_FAILURE",
    debugId: createDebugId_(),
    details: toSafeString_(clientMessage)
  });
  return logSaveFailureWithResult_(result, data || {});
}

function logClientSaveFailure(data, clientMessage) {
  Logger.log("[ROUTE WRAPPER] logClientSaveFailure -> issuance.log_failure");
  return unwrapRouteWrapperResult_("log_failure", routeModule("issuance", "log_failure", { data: data, clientMessage: clientMessage }));
}

function legacy_retryFailedSave_(logRow) {
  const authFailure = requireEditorAccessResult_("");
  if (authFailure) return authFailure;

  const targetRow = Number(logRow);
  const sheet = getOrCreateSaveFailureSheet_();
  if (!targetRow || targetRow < 2 || targetRow > sheet.getLastRow()) {
    return buildSaveResult_(false, {
      row: "",
      message: "Failed save record not found.",
      code: "FAILURE_LOG_NOT_FOUND",
      debugId: createDebugId_()
    });
  }

  const rowValues = sheet.getRange(targetRow, 1, 1, saveFailureHeaders.length).getValues()[0];
  const payloadJson = String(rowValues[5] || "");
  if (!payloadJson) {
    const missingPayloadResult = buildSaveResult_(false, {
      row: rowValues[2] || "",
      id: rowValues[3] || "",
      name: rowValues[4] || "",
      message: "Failed save payload is missing.",
      code: "FAILURE_PAYLOAD_MISSING",
      debugId: createDebugId_()
    });
    updateSaveFailureRetryStatus_(targetRow, missingPayloadResult);
    return missingPayloadResult;
  }

  try {
    const payload = JSON.parse(payloadJson);
    const result = performIssuanceSave_(payload, { isRetry: true, failureLogRow: targetRow });
    updateSaveFailureRetryStatus_(targetRow, result);
    return result;
  } catch (error) {
    const parseResult = buildSaveResult_(false, {
      row: rowValues[2] || "",
      id: rowValues[3] || "",
      name: rowValues[4] || "",
      message: "Failed save payload could not be retried.",
      code: "FAILURE_PAYLOAD_INVALID",
      debugId: createDebugId_(),
      details: error && error.message ? error.message : String(error)
    });
    updateSaveFailureRetryStatus_(targetRow, parseResult);
    return parseResult;
  }
}

function retryFailedSave(logRow) {
  Logger.log("[ROUTE WRAPPER] retryFailedSave -> issuance.retry");
  return unwrapRouteWrapperResult_("retry", routeModule("issuance", "retry", { logRow: logRow }));
}

function syncAccountPermissions() {
  const debugId = createDebugId_();
  try {
    const authFailure = requirePermissionSyncAccessResult_();
    if (authFailure) {
      return {
        ok: false,
        summary: authFailure.message,
        added: [],
        removed: [],
        skipped: [],
        failures: [],
        debugId: authFailure.debugId,
        code: authFailure.code
      };
    }

    const accountsState = getManagedEditorEmailsFromDirectory_();
    const protectedEmails = getProtectedEditorEmails_();
    const file = DriveApp.getFileById(spreadsheetID);
    const currentEditors = file.getEditors().map(function(user) {
      return normalizeEmail_(user && user.getEmail ? user.getEmail() : "");
    }).filter(Boolean);

    const currentMap = {};
    currentEditors.forEach(function(email) { currentMap[email] = true; });

    const canEditMap = {};
    accountsState.canEditEmails.forEach(function(email) { canEditMap[email] = true; });
    const listedMap = {};
    accountsState.listedEmails.forEach(function(email) { listedMap[email] = true; });

    const added = [];
    const removed = [];
    const skipped = [];
    const failures = [];

    accountsState.invalidEntries.forEach(function(item) {
      skipped.push({ email: item, reason: "Invalid email in " + accountsState.sourceSheetName + " sheet" });
    });

    accountsState.canEditEmails.forEach(function(email) {
      if (protectedEmails[email]) {
        skipped.push({ email: email, reason: "Protected editor" });
        return;
      }
      if (currentMap[email]) return;
      try {
        file.addEditor(email);
        added.push(email);
      } catch (error) {
        failures.push({
          email: email,
          action: "add",
          reason: error && error.message ? error.message : String(error)
        });
      }
    });

    accountsState.listedEmails.forEach(function(email) {
      if (canEditMap[email]) return;
      if (!currentMap[email]) return;
      if (protectedEmails[email]) {
        skipped.push({ email: email, reason: "Protected editor" });
        return;
      }
      try {
        file.removeEditor(email);
        removed.push(email);
      } catch (error) {
        failures.push({
          email: email,
          action: "remove",
          reason: error && error.message ? error.message : String(error)
        });
      }
    });

    const summary = "Added " + added.length + ", removed " + removed.length + ", skipped " + skipped.length + ", failures " + failures.length + ".";
    Logger.log(JSON.stringify({
      event: "account_permission_sync",
      debugId: debugId,
      added: added,
      removed: removed,
      skipped: skipped,
      failures: failures
    }));

    return {
      ok: failures.length === 0,
      summary: summary,
      added: added,
      removed: removed,
      skipped: skipped,
      failures: failures,
      debugId: debugId
    };
  } catch (error) {
    Logger.log(JSON.stringify({
      event: "account_permission_sync_failure",
      debugId: debugId,
      details: error && error.message ? error.message : String(error)
    }));
    return {
      ok: false,
      summary: "Unable to synchronize spreadsheet editors right now.",
      added: [],
      removed: [],
      skipped: [],
      failures: [{
        email: "",
        action: "sync",
        reason: error && error.message ? error.message : String(error)
      }],
      debugId: debugId
    };
  }
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

function getOrCreateAccountsSheet_() {
  const ss = SpreadsheetApp.openById(spreadsheetID);
  let sheet = ss.getSheetByName(accountsSheetName);
  if (!sheet) {
    sheet = ss.insertSheet(accountsSheetName);
    sheet.getRange(1, 1, 1, editorsHeaders.length).setValues([editorsHeaders]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, editorsHeaders.length).setValues([editorsHeaders]);
  }
  return sheet;
}

function getUsersSheet_() {
  const ss = SpreadsheetApp.openById(spreadsheetID);
  return ss.getSheetByName(usersSheetName);
}

function getAccountName() {
  const context = getCurrentUserContext_();
  if (context.name) return context.name;
  return "Guest Viewer";
}

function getUserRole() {
  const access = getIssuancePermissionContext_("");
  return access.actions.update || access.actions.create || access.actions["delete"] ? "Editor" : "Viewer";
}

function markSheetUpdated_(sheetName, rowNumber) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("LAST_UPDATED_ROW", String(rowNumber || ""));
  props.setProperty("LAST_UPDATED_SHEET", String(sheetName || ""));
  props.setProperty("LAST_UPDATED_TIME", new Date().toISOString());
}

function markDatabaseRowUpdated_(rowNumber) {
  if (!rowNumber || rowNumber < 2) return;
  markSheetUpdated_(databaseSheetName, rowNumber);
}

function markAccountsSheetUpdated_(rowNumber) {
  if (!rowNumber || rowNumber < 2) return;
  markSheetUpdated_(accountsSheetName, rowNumber);
}

function markUsersSheetUpdated_(rowNumber) {
  if (!rowNumber || rowNumber < 2) return;
  markSheetUpdated_(usersSheetName, rowNumber);
}

function markUserAccessSheetUpdated_(rowNumber) {
  if (!rowNumber || rowNumber < 2) return;
  markSheetUpdated_(userAccessSheetName, rowNumber);
}

function markDataTypesSheetUpdated_(rowNumber) {
  if (!rowNumber || rowNumber < 2) return;
  markSheetUpdated_(dataTypesSheetName, rowNumber);
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() === databaseSheetName) {
    markDatabaseRowUpdated_(e.range.getRow());
    return;
  }
  if (sheet.getName() === accountsSheetName) {
    markAccountsSheetUpdated_(e.range.getRow());
    return;
  }
  if (sheet.getName() === usersSheetName) {
    markUsersSheetUpdated_(e.range.getRow());
    return;
  }
  if (sheet.getName() === userAccessSheetName) {
    markUserAccessSheetUpdated_(e.range.getRow());
    return;
  }
  if (sheet.getName() === dataTypesSheetName) {
    markDataTypesSheetUpdated_(e.range.getRow());
  }
}

function legacy_getDatabaseRowInfo_() {
  const props = PropertiesService.getScriptProperties();
  return {
    row: props.getProperty("LAST_UPDATED_ROW") || "",
    sheet: props.getProperty("LAST_UPDATED_SHEET") || "",
    updatedAt: props.getProperty("LAST_UPDATED_TIME") || "0"
  }
}

function getDatabaseRowInfo() {
  Logger.log("[ROUTE WRAPPER] getDatabaseRowInfo -> issuance.row_info");
  return unwrapRouteWrapperResult_("row_info", routeModule("issuance", "row_info", {}));
}

function legacy_getDatabaseRow_(rowNumber) {
  const viewConfig = getCurrentRecordViewConfig_();
  const row = Number(rowNumber);
  if (!row || row < 2) return null;

  const ss = SpreadsheetApp.openById(spreadsheetID);
  const sheet = ss.getSheetByName(databaseSheetName);

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0]
      .map(h => h.toString().trim());
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn())
      .getValues()[0];
  return mapDatabaseRowsForView_(headers, [values], row, viewConfig.dataColumns)[0] || null;
}

function getDatabaseRow(rowNumber) {
  Logger.log("[ROUTE WRAPPER] getDatabaseRow -> issuance.fetch");
  return unwrapRouteWrapperResult_("fetch", routeModule("issuance", "fetch", { row: rowNumber }));
}

function convertImageToBase64(fileId) {
  try {
    if (!fileId) return null;

    const file = DriveApp.getFileById(String(fileId).trim());
    const blob = file.getBlob();
    const contentType = file.getMimeType() || blob.getContentType();
    const base64 = Utilities.base64Encode(blob.getBytes());

    return `data:${contentType};base64,${base64}`;
  } catch (error) {
    console.error("Error converting image to base64:", error);
    return null;
  }
}

function getNavbarLogoData() {
  var dataUrl = convertImageToBase64(navbarLogoFileId);
  return {
    fileId: navbarLogoFileId,
    dataUrl: dataUrl
  };
}

function saveIssuanceData(payload) {
  return legacy_updateIssuanceColumns_(payload);
}

function updateIssuanceData(payload) {
  return legacy_updateIssuanceColumns_(payload);
}

function getIssuanceData(payload) {
  payload = payload || {};
  if (payload.row != null && payload.row !== "") {
    return legacy_getDatabaseRow_(payload.row);
  }
  if (payload.searchTerm) {
    return legacy_searchExactName_(payload.searchTerm);
  }
  if (payload.startRow != null && payload.startRow !== "") {
    return legacy_loadRemainingData_(payload.startRow, payload.batchSize);
  }
  return legacy_getInitialData_(payload.batchSize);
}

const ROUTING_AUTHORITY_MODE = false;
const ROUTE_AUDIT_COMPARE_MODE = true;

function buildHandledRouteResponse_(data) {
  return {
    handled: true,
    status: "SUCCESS",
    data: data
  };
}

function buildUnhandledRouteResponse_(module, action, reason) {
  return {
    handled: false,
    status: "NOT_IMPLEMENTED",
    module: module || "",
    action: action || "",
    reason: reason || ""
  };
}

function normalizeTypingRouteAction_(action) {
  var normalized = String(action || "").trim().toLowerCase();
  if (normalized === "createbatch") return "createBatch";
  if (normalized === "retryfailedrow") return "retryFailedRow";
  return "";
}

function isTypingPositionAllowed_(position) {
  var normalizedPosition = String(position || "").trim().toLowerCase();
  return normalizedPosition === "releasing/verifier" || normalizedPosition === "developer";
}

function getTypingBatchPropertyKey_(batchId) {
  return "TYPING_BATCH_" + String(batchId || "").trim().toUpperCase().replace(/[^A-Z0-9_]+/g, "_");
}

function isTypingBatchProcessed_(batchId) {
  return !!PropertiesService.getScriptProperties().getProperty(getTypingBatchPropertyKey_(batchId));
}

function markTypingBatchProcessed_(batchId, result) {
  PropertiesService.getScriptProperties().setProperty(getTypingBatchPropertyKey_(batchId), JSON.stringify({
    batchId: String(batchId || "").trim(),
    inserted: Number(result && result.inserted) || 0,
    failed: Number(result && result.failed) || 0,
    timestamp: new Date().toISOString()
  }));
}

function formatTypingDateValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var trimmed = String(value == null ? "" : value).trim();
  if (!trimmed) return "";
  var parsed = parseIsoDate_(trimmed);
  if (parsed instanceof Date && !isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return trimmed;
}

function normalizeTypingFieldValue_(fieldKey, value) {
  var trimmed = String(value == null ? "" : value).trim();
  if (fieldKey === "studentName") return trimmed.toUpperCase();
  if (typingCreateFieldConfig[fieldKey] && typingCreateFieldConfig[fieldKey].type === "date") {
    return formatTypingDateValue_(trimmed);
  }
  return trimmed;
}

function getTypingWritableFieldKeys_() {
  var definitions = getDataTypeFieldDefinitions_();
  var releasingVerifierFields = {};
  definitions.forEach(function(fieldMeta) {
    var normalizedPosition = String((fieldMeta && fieldMeta.position) || "").trim().toLowerCase();
    if (normalizedPosition === "releasing/verifier") {
      releasingVerifierFields[String(fieldMeta.name || "").trim()] = true;
    }
  });

  var writableFieldKeys = [];
  typingCreateFieldKeys.forEach(function(fieldKey) {
    var columnName = typingCreateFieldConfig[fieldKey].columnName;
    if (releasingVerifierFields[columnName]) {
      writableFieldKeys.push(fieldKey);
      return;
    }
    Logger.log("[TYPING VALIDATION WARNING] Missing Data Types mapping for position Releasing/Verifier: " + columnName);
  });

  return writableFieldKeys;
}

function getTypingFailurePayload_(batchId, student, retryCount) {
  var safeStudent = student || {};
  return {
    batchId: String(batchId || "").trim(),
    studentName: normalizeTypingFieldValue_("studentName", safeStudent.studentName),
    originalData: {
      dateReceived: normalizeTypingFieldValue_("dateReceived", safeStudent.dateReceived),
      studentName: normalizeTypingFieldValue_("studentName", safeStudent.studentName),
      hei: normalizeTypingFieldValue_("hei", safeStudent.hei),
      dateOfGraduation: normalizeTypingFieldValue_("dateOfGraduation", safeStudent.dateOfGraduation),
      program: normalizeTypingFieldValue_("program", safeStudent.program)
    },
    retryCount: Number(retryCount) || 0
  };
}

function getTypingFailureDetails_(batchId, retryCount) {
  return "batchId=" + String(batchId || "").trim() + ";retryCount=" + (Number(retryCount) || 0);
}

function findTypingFailureLogRow_(batchId, studentName) {
  var sheet = getOrCreateSaveFailureSheet_();
  if (sheet.getLastRow() < 2) return 0;
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var normalizedBatchId = String(batchId || "").trim();
  var normalizedName = normalizeTypingFieldValue_("studentName", studentName);

  for (var i = rows.length - 1; i >= 0; i--) {
    var code = String(rows[i][6] || "").trim();
    if (code !== "TYPING_CREATE_FAILED") continue;
    var payload = {};
    try {
      payload = JSON.parse(String(rows[i][5] || "{}"));
    } catch (e) {
      payload = {};
    }
    if (
      String(payload.batchId || "").trim() === normalizedBatchId &&
      normalizeTypingFieldValue_("studentName", payload.studentName || (payload.originalData && payload.originalData.studentName) || rows[i][4]) === normalizedName
    ) {
      return i + 2;
    }
  }
  return 0;
}

function logTypingCreateFailure_(batchId, student, error, retryCount, existingLogRow) {
  var failurePayload = getTypingFailurePayload_(batchId, student, retryCount);
  var failureMessage = error && error.message ? String(error.message) : String(error || "Unknown typing create failure.");
  var targetLogRow = Number(existingLogRow) || 0;
  if (targetLogRow > 1) {
    var targetSheet = getOrCreateSaveFailureSheet_();
    targetSheet.getRange(targetLogRow, 1, 1, targetSheet.getLastColumn()).setValues([[
      new Date(),
      createDebugId_(),
      "",
      "",
      failurePayload.studentName,
      JSON.stringify(failurePayload),
      "TYPING_CREATE_FAILED",
      failureMessage,
      getTypingFailureDetails_(batchId, retryCount),
      getEditorEmail_(),
      "failed",
      "",
      "",
      ""
    ]]);
    return targetLogRow;
  }
  return appendSaveFailureLog_({
    debugId: createDebugId_(),
    row: "",
    id: "",
    name: failurePayload.studentName,
    payload: failurePayload,
    code: "TYPING_CREATE_FAILED",
    message: failureMessage,
    details: getTypingFailureDetails_(batchId, retryCount),
    status: "failed"
  });
}

function parseTypingFailurePayload_(value) {
  try {
    return JSON.parse(String(value || "{}"));
  } catch (e) {
    return {};
  }
}

function getTypingFailureHistory(batchId) {
  var normalizedBatchId = String(batchId || "").trim();
  if (!normalizedBatchId) return [];
  var sheet = getOrCreateSaveFailureSheet_();
  if (sheet.getLastRow() < 2) return [];
  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  return rows.map(function(row, index) {
    var payload = parseTypingFailurePayload_(row[5]);
    return {
      logRow: index + 2,
      timestamp: row[0] || "",
      debugId: row[1] || "",
      code: row[6] || "",
      studentName: String(row[4] || payload.studentName || "").trim(),
      payload: payload,
      message: row[7] || "",
      details: row[8] || "",
      status: row[10] || "failed",
      retryTimestamp: row[11] || "",
      retryResult: row[12] || "",
      retryDebugId: row[13] || "",
      retryCount: Number(payload.retryCount) || 0
    };
  }).filter(function(entry) {
    return String(entry.code || "").trim() === "TYPING_CREATE_FAILED" &&
      String(entry.payload.batchId || "").trim() === normalizedBatchId &&
      entry.payload &&
      (String(entry.payload.studentName || "").trim() || (entry.payload.originalData && String(entry.payload.originalData.studentName || "").trim())) &&
      String(entry.payload.batchId || "").trim();
  });
}

function getTypingBatchColumnIndex_(headers) {
  var safeHeaders = Array.isArray(headers) ? headers : [];
  var index = safeHeaders.indexOf("Typing Batch ID");
  if (index === -1) {
    Logger.log("[TYPING INFO] Typing Batch ID column not found, running in legacy mode");
    Logger.log("[TYPING TRACKING DISABLED] Falling back to value-based detection");
    return null;
  }
  Logger.log("[TYPING TRACKING ENABLED] Using Typing Batch ID column");
  return index;
}

function legacy_getDashboardData_(filters) {
  return dashboard_fetchData_(filters || {});
}

function getDashboardData(filters) {
  var serverContext = getServerRequestContext_(arguments, filters, "dashboard");
  logServerRequestContext_(serverContext);
  return legacy_getDashboardData_(filters || {});
}

// ========== Executive Dashboard Migration: Core Utilities ==========

function toTrimmedString_(value) {
  return value == null ? "" : String(value).trim();
}

function buildSql_(lines) {
  return lines.filter(function(line) { return line != null && String(line).trim() !== ""; }).join("\n");
}

function formatDateInputValue_(date) {
  if (!date || isNaN(date.getTime())) return "";
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  var day = date.getDate();
  var monthStr = month < 10 ? "0" + month : String(month);
  var dayStr = day < 10 ? "0" + day : String(day);
  return year + "-" + monthStr + "-" + dayStr;
}

function offsetDateValue_(baseDate, offsetDays) {
  if (!baseDate || isNaN(baseDate.getTime())) return "";
  var result = new Date(baseDate);
  result.setDate(result.getDate() + offsetDays);
  return formatDateInputValue_(result);
}

function createDiagnostics_(label) {
  return {
    requestId: String(label || "") + "-" + new Date().getTime() + "-" + Math.floor(Math.random() * 100000),
    startedAt: new Date().getTime(),
    timings: {},
    cacheStatus: "bypass",
    warnings: []
  };
}

function startTiming_(diagnostics, name, meta) {
  diagnostics.timings[name] = {
    startedAt: new Date().getTime(),
    meta: meta || {}
  };
}

function endTiming_(diagnostics, name, meta) {
  var timing = diagnostics.timings[name];
  if (timing) {
    timing.endedAt = new Date().getTime();
    timing.durationMs = timing.endedAt - timing.startedAt;
    if (meta) {
      timing.meta = timing.meta || {};
      Object.keys(meta).forEach(function(key) {
        timing.meta[key] = meta[key];
      });
    }
  }
}

function finalizeDiagnostics_(diagnostics, status, meta) {
  diagnostics.status = status;
  diagnostics.endedAt = new Date().getTime();
  diagnostics.durationMs = diagnostics.endedAt - diagnostics.startedAt;
  if (meta) {
    diagnostics.meta = diagnostics.meta || {};
    Object.keys(meta).forEach(function(key) {
      diagnostics.meta[key] = meta[key];
    });
  }
}

// ========== Executive Dashboard Migration: Filter Normalization ==========

function normalizeDashboardStatus_(value) {
  var raw = toTrimmedString_(value).toLowerCase();
  if (!raw) return "";
  if (raw.indexOf("for evaluation") !== -1) return "For Evaluation";
  if (raw.indexOf("evaluated") !== -1 || raw.indexOf("for verification") !== -1) return "Evaluated";
  if (raw.indexOf("deficien") !== -1) return "Deficiency";
  if (raw.indexOf("verified") !== -1) return "Verified";
  if (raw.indexOf("for signature") !== -1 || raw.indexOf("for signing") !== -1 || raw === "signature") return "For Signature";
  if (raw.indexOf("for release") !== -1 || raw.indexOf("ready for release") !== -1 || raw.indexOf("pending release") !== -1 || raw.indexOf("releasing") !== -1) return "For Release";
  if (raw.indexOf("released") !== -1 || raw.indexOf("release complete") !== -1 || raw.indexOf("completed") !== -1) return "Released";
  if (raw.indexOf("approved") !== -1 || raw.indexOf("approeved") !== -1 || raw.indexOf("arpproved") !== -1) return "Approved";
  if (raw.indexOf("cancel") !== -1) return "Cancelled";
  return "";
}

function normalizeFilterValue_(value) {
  var trimmed = toTrimmedString_(value);
  if (!trimmed || trimmed.toUpperCase() === "ALL") return "";
  return trimmed;
}

function normalizeToken_(value) {
  return toTrimmedString_(value).toLowerCase().replace(/\s+/g, " ");
}

function normalizeDatePresetValue_(value) {
  var preset = toTrimmedString_(value) || "allTime";
  var validPresets = ["allTime", "thisYear", "yearToDate", "last30Days", "last90Days", "specificDate"];
  if (validPresets.indexOf(preset) !== -1) return preset;
  return "thisYear";
}

function normalizeDateValue_(value) {
  var trimmed = toTrimmedString_(value);
  if (!trimmed || !/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return "";
  return trimmed;
}

function resolveDashboardDatePresetRange_(datePreset, dateFrom, dateTo) {
  var anchorDateValue = getDashboardPresetAnchorDate_();
  var anchorDate = anchorDateValue ? new Date(anchorDateValue + "T00:00:00") : new Date();
  var effectiveDate = isNaN(anchorDate.getTime()) ? new Date() : anchorDate;
  var currentYear = effectiveDate.getFullYear();
  var todayValue = formatDateInputValue_(effectiveDate);
  
  // For last30Days and last90Days, use CURRENT DATE (not anchor date)
  var now = new Date();
  var currentDateValue = formatDateInputValue_(now);

  if (datePreset === "allTime") {
    return { dateFrom: "", dateTo: "" };
  }
  if (datePreset === "thisYear") {
    return {
      dateFrom: currentYear + "-01-01",
      dateTo: todayValue
    };
  }
  if (datePreset === "yearToDate") {
    var fiscalYearStart = new Date(Date.UTC(currentYear, 0, 1));
    return {
      dateFrom: formatDateInputValue_(fiscalYearStart),
      dateTo: todayValue
    };
  }
  if (datePreset === "last30Days") {
    // Use current date, not anchor date
    return {
      dateFrom: offsetDateValue_(now, -29),
      dateTo: currentDateValue
    };
  }
  if (datePreset === "last90Days") {
    // Use current date, not anchor date
    return {
      dateFrom: offsetDateValue_(now, -89),
      dateTo: currentDateValue
    };
  }
  if (datePreset === "specificDate") {
    return {
      dateFrom: normalizeDateValue_(dateFrom),
      dateTo: normalizeDateValue_(dateTo)
    };
  }


  return { dateFrom: "", dateTo: "" };
}

function normalizeFilters_(filters) {
  var datePreset = normalizeDatePresetValue_(filters && filters.datePreset);
  var resolvedDateRange = resolveDashboardDatePresetRange_(
    datePreset,
    normalizeDateValue_(filters && filters.dateFrom),
    normalizeDateValue_(filters && filters.dateTo)
  );
  return {
    program: normalizeFilterValue_(filters && filters.program),
    es2InCharge: normalizeFilterValue_(filters && filters.es2InCharge),
    datePreset: datePreset,
    dateFrom: resolvedDateRange.dateFrom,
    dateTo: resolvedDateRange.dateTo,
    applicationStatus: normalizeFilterValue_(filters && filters.applicationStatus),
    transactionType: normalizeFilterValue_(filters && filters.transactionType),
    hei: normalizeFilterValue_(filters && filters.hei)
  };
}

function runBigQueryQuery_(query, queryParameters) {
  Logger.log("[RBQ] runBigQueryQuery_ called - FIXED VERSION");
  var projectId = BIGQUERY_PROJECT_ID;
  var request = {
    query: query,
    useLegacySql: false,
    parameterMode: "NAMED",
    queryParameters: queryParameters || []
  };
  try {
    var queryResults = BigQuery.Jobs.query(request, projectId);
    var jobId = queryResults.jobReference.jobId;

    var attempts = 0;
    while (!queryResults.jobComplete && attempts < 10) {
      Utilities.sleep(500);
      queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
      attempts += 1;
    }

    var rows = queryResults.rows || [];
    while (queryResults.pageToken) {
      queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId, {
        pageToken: queryResults.pageToken
      });
      rows = rows.concat(queryResults.rows || []);
    }
    return rows;
  } catch (e) {
    var msg = e.toString();
    if (msg.indexOf("Not found") !== -1) throw new Error("404: Analytics dataset not found.");
    if (msg.indexOf("Access Denied") !== -1) throw new Error("403: No BigQuery access.");
    throw new Error("BigQuery Error: " + msg);
  }
}

function mapBigQueryRows_(rows, mapper) {
  if (!rows || !rows.length) return [];
  return rows.map(function(row) {
    var values = (row.f || []).map(function(field) {
      return field && field.v != null ? field.v : "";
    });
    return mapper(values);
  });
}

function buildBigQueryNamedParameter_(name, type, value) {
  var param = {
    name: name,
    parameterType: { type: type }
  };
  if (type === "DATE") {
    param.parameterValue = { value: value };
  } else if (type === "INT64" || type === "INTEGER") {
    param.parameterValue = { value: String(value) };
  } else if (type === "FLOAT" || type === "DOUBLE") {
    param.parameterValue = { value: String(value) };
  } else if (type === "BOOL" || type === "BOOLEAN") {
    param.parameterValue = { value: value ? "true" : "false" };
  } else {
    param.parameterValue = { value: String(value || "") };
  }
  return param;
}

function buildBigQueryWhereClause_(filters, options) {
  var config = options || {};
  var excludedFilters = {};
  var excludeFiltersArray = config.excludeFilters || [];
  for (var i = 0; i < excludeFiltersArray.length; i++) {
    excludedFilters[String(excludeFiltersArray[i] || "")] = true;
  }

  var normalizedFilters = normalizeFilters_(filters || {});
  var queryParts = ["WHERE 1=1"];
  var queryParameters = [];
  var canonicalStatusExpression = getDashboardCanonicalStatusExpression_();

  // ES2/Program alignment: ensure filters respect ownership mapping
  var programOwnershipMap = config.programOwnershipMap || {};
  var effectiveProgram = normalizedFilters.program;
  var effectiveEs2InCharge = normalizedFilters.es2InCharge;

  // If both program and es2InCharge are selected, verify alignment
  if (effectiveProgram && effectiveEs2InCharge && Object.keys(programOwnershipMap).length > 0) {
    var expectedOwner = programOwnershipMap[effectiveProgram];
    if (expectedOwner && expectedOwner !== effectiveEs2InCharge) {
      // Misalignment: selected ES2 doesn't own the selected program
      // Give precedence to program ownership (filter by program and its actual owner)
      effectiveEs2InCharge = expectedOwner;
    }
  }

  if (!excludedFilters.program && effectiveProgram) {
    queryParts.push("AND program = @program");
    queryParameters.push(buildBigQueryNamedParameter_("program", "STRING", effectiveProgram));
  }
  if (!excludedFilters.es2InCharge && effectiveEs2InCharge) {
    queryParts.push("AND es2InCharge = @es2InCharge");
    queryParameters.push(buildBigQueryNamedParameter_("es2InCharge", "STRING", effectiveEs2InCharge));
  }
  if (!excludedFilters.applicationStatus && normalizedFilters.applicationStatus) {
    var canonicalStatus = normalizeDashboardStatus_(normalizedFilters.applicationStatus) || normalizedFilters.applicationStatus;
    queryParts.push("AND (" + canonicalStatusExpression + ") = @canonicalStatus");
    queryParameters.push(buildBigQueryNamedParameter_("canonicalStatus", "STRING", canonicalStatus));
  }
  if (!excludedFilters.transactionType && normalizedFilters.transactionType) {
    queryParts.push("AND transactionType = @transactionType");
    queryParameters.push(buildBigQueryNamedParameter_("transactionType", "STRING", normalizedFilters.transactionType));
  }
  if (!excludedFilters.hei && normalizedFilters.hei) {
    queryParts.push("AND hei = @hei");
    queryParameters.push(buildBigQueryNamedParameter_("hei", "STRING", normalizedFilters.hei));
  }
  if (!excludedFilters.dateFrom && normalizedFilters.dateFrom) {
    queryParts.push("AND dateReceivedDate >= @dateFrom");
    queryParameters.push(buildBigQueryNamedParameter_("dateFrom", "DATE", normalizedFilters.dateFrom));
  }
  if (!excludedFilters.dateTo && normalizedFilters.dateTo) {
    queryParts.push("AND dateReceivedDate <= @dateTo");
    queryParameters.push(buildBigQueryNamedParameter_("dateTo", "DATE", normalizedFilters.dateTo));
  }

  return {
    whereClause: queryParts.join(" "),
    queryParameters: queryParameters,
    normalizedFilters: normalizedFilters
  };
}

function runFilteredReportingQuery_(filters, queryFactory, options) {
  var filterSpec = buildBigQueryWhereClause_(filters, options);
  var result = runBigQueryOverReportingSources_(function(source) {
    return queryFactory(source, filterSpec);
  }, filterSpec.queryParameters);
  return {
    result: result,
    filterSpec: filterSpec
  };
}

function runBigQueryOverReportingSources_(queryFactory, queryParameters) {
  var source = BIGQUERY_NATIVE_REPORTING_VIEW;
  try {
    return {
      source: source,
      rows: runBigQueryQuery_(queryFactory(source), queryParameters || [])
    };
  } catch (error) {
    Logger.log("[DASHBOARD QUERY ERROR] Failed to query " + source + ": " + error.toString());
    throw new Error("Unable to query dashboard source: " + error.message);
  }
}

function runDashboardBigQueryOverSources_(queryFactory) {
  // Use only the native reporting view (matches Executive Dashboard)
  // The database_flat table doesn't have dateReceivedDate as a proper DATE column
  var source = BIGQUERY_NATIVE_REPORTING_VIEW;
  try {
    return {
      source: source,
      rows: runBigQueryQuery_(queryFactory(source), [])
    };
  } catch (error) {
    Logger.log("[DASHBOARD QUERY ERROR] Failed to query " + source + ": " + error.toString());
    throw new Error("Unable to query dashboard source: " + error.message);
  }
}

function getDashboardCanonicalStatusExpression_() {
  return [
    "CASE",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%for evaluation%' THEN 'For Evaluation'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%evaluated%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%for verification%' THEN 'Evaluated'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%deficien%' THEN 'Deficiency'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%verified%' THEN 'Verified'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%for signature%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%for signing%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) = 'signature' THEN 'For Signature'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%for release%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%ready for release%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%pending release%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%releasing%' THEN 'For Release'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%released%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%release complete%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%completed%' THEN 'Released'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%approved%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%approeved%' OR LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%arpproved%' THEN 'Approved'",
    "  WHEN LOWER(TRIM(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), ''))) LIKE '%cancel%' THEN 'Cancelled'",
    "  ELSE ''",
    "END"
  ].join("\n");
}

function getDashboardPresetAnchorDate_() {
  try {
    var cache = CacheService.getScriptCache();
    var cacheKey = "dashboard:presetAnchorDate";
    var cached = cache.get(cacheKey);
    if (cached) return cached;
    
    var result = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT CAST(MAX(dateReceivedDate) AS STRING) AS maxDate",
        "FROM " + source,
        "WHERE dateReceivedDate IS NOT NULL"
      ].join("\n");
    });
    var firstRow = result.rows && result.rows[0] ? result.rows[0].f || [] : [];
    var anchorDate = String(firstRow[0] && firstRow[0].v || "").trim();
    if (anchorDate) {
      cache.put(cacheKey, anchorDate, 3600); // Cache for 1 hour
      return anchorDate;
    }
  } catch (e) {
    Logger.log("[ANCHOR DATE] Error: " + e.toString());
  }
  return Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
}

function buildDashboardWhereClause_(filters, options) {
  // New parameterized version with backward compatibility and filter exclusions
  var filterSpec = buildBigQueryWhereClause_(filters, options);
  // For backward compatibility, convert "WHERE 1=1 AND x = @x" to "1=1 AND x = 'value'"
  // This maintains compatibility with existing query functions that expect string interpolation
  var normalizedFilters = filterSpec.normalizedFilters;
  var clause = filterSpec.whereClause;

  // Extract aligned ES2 value from queryParameters (ES2/Program alignment may have modified it)
  var effectiveEs2InCharge = normalizedFilters.es2InCharge;
  for (var pi = 0; pi < filterSpec.queryParameters.length; pi++) {
    var p = filterSpec.queryParameters[pi];
    if (p.name === "es2InCharge" && p.parameterValue && p.parameterValue.value) {
      effectiveEs2InCharge = p.parameterValue.value;
      break;
    }
  }

  // Replace named parameters with actual values for backward compatibility
  if (normalizedFilters.program) {
    clause = clause.replace("@program", "'" + normalizedFilters.program.replace(/'/g, "\\'") + "'");
  }
  if (effectiveEs2InCharge) {
    clause = clause.replace("@es2InCharge", "'" + effectiveEs2InCharge.replace(/'/g, "\\'") + "'");
  }
  if (normalizedFilters.applicationStatus) {
    var status = normalizeDashboardStatus_(normalizedFilters.applicationStatus) || normalizedFilters.applicationStatus;
    clause = clause.replace("@canonicalStatus", "'" + status.replace(/'/g, "\\'") + "'");
  }
  if (normalizedFilters.transactionType) {
    clause = clause.replace("@transactionType", "'" + normalizedFilters.transactionType.replace(/'/g, "\\'") + "'");
  }
  if (normalizedFilters.hei) {
    clause = clause.replace("@hei", "'" + normalizedFilters.hei.replace(/'/g, "\\'") + "'");
  }
  if (normalizedFilters.dateFrom) {
    clause = clause.replace("@dateFrom", "DATE('" + normalizedFilters.dateFrom + "')");
  }
  if (normalizedFilters.dateTo) {
    clause = clause.replace("@dateTo", "DATE('" + normalizedFilters.dateTo + "')");
  }

  // Remove "WHERE " prefix to maintain backward compatibility with existing code
  return clause.replace(/^WHERE /, "");
}

function getDashboardSummary_(filters, options) {
  var canonicalStatusExpression = getDashboardCanonicalStatusExpression_();
  var whereClause = buildDashboardWhereClause_(filters, options);
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COUNT(1) AS totalCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'For Evaluation' THEN 1 ELSE 0 END) AS forEvaluationCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Evaluated' THEN 1 ELSE 0 END) AS evaluatedCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Deficiency' THEN 1 ELSE 0 END) AS deficiencyCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Approved' THEN 1 ELSE 0 END) AS approvedCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Verified' THEN 1 ELSE 0 END) AS verifiedCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'For Signature' THEN 1 ELSE 0 END) AS forSignatureCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'For Release' THEN 1 ELSE 0 END) AS forReleaseCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Released' THEN 1 ELSE 0 END) AS releasedCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Cancelled' THEN 1 ELSE 0 END) AS cancelledCount",
      "FROM " + source,
      "WHERE " + whereClause
    ].join("\n");
  });
  var rows = mapBigQueryRows_(result.rows, function(values) {
    return {
      totalCount: Number(values[0] || 0),
      forEvaluationCount: Number(values[1] || 0),
      evaluatedCount: Number(values[2] || 0),
      deficiencyCount: Number(values[3] || 0),
      approvedCount: Number(values[4] || 0),
      verifiedCount: Number(values[5] || 0),
      forSignatureCount: Number(values[6] || 0),
      forReleaseCount: Number(values[7] || 0),
      releasedCount: Number(values[8] || 0),
      cancelledCount: Number(values[9] || 0)
    };
  });
  return {
    source: result.source,
    values: rows[0] || { totalCount: 0, forEvaluationCount: 0, evaluatedCount: 0, deficiencyCount: 0, approvedCount: 0, verifiedCount: 0, forSignatureCount: 0, forReleaseCount: 0, releasedCount: 0, cancelledCount: 0 }
  };
}

function getDashboardStatusDistribution_(filters, options) {
  var canonicalStatusExpression = getDashboardCanonicalStatusExpression_();
  var whereClause = buildDashboardWhereClause_(filters, options);
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  " + canonicalStatusExpression + " AS statusLabel,",
      "  COUNT(1) AS statusCount",
      "FROM " + source,
      "WHERE " + whereClause,
      "GROUP BY statusLabel",
      "ORDER BY statusCount DESC"
    ].join("\n");
  });
  return {
    source: result.source,
    rows: mapBigQueryRows_(result.rows, function(values) {
      return {
        label: String(values[0] || "Unknown"),
        value: Number(values[1] || 0)
      };
    })
  };
}

function getDashboardWorkloadSummary_(filters, options) {
  var canonicalStatusExpression = getDashboardCanonicalStatusExpression_();
  var whereClause = buildDashboardWhereClause_(filters, options);
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(es2InCharge AS STRING), 'Unassigned') AS ownerLabel,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Approved' THEN 1 ELSE 0 END) AS approvedCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") = 'Deficiency' THEN 1 ELSE 0 END) AS deficiencyCount,",
      "  SUM(CASE WHEN (" + canonicalStatusExpression + ") NOT IN ('Approved', 'Released', 'Cancelled', 'Deficiency', '') THEN 1 ELSE 0 END) AS processingPendingCount,",
      "  COUNT(1) AS workloadCount",
      "FROM " + source,
      "WHERE " + whereClause,
      "GROUP BY ownerLabel",
      "ORDER BY workloadCount DESC",
      "LIMIT 12"
    ].join("\n");
  });
  return {
    source: result.source,
    rows: mapBigQueryRows_(result.rows, function(values) {
      return {
        es2: String(values[0] || "Unassigned"),
        approvedCount: Number(values[1] || 0),
        deficiencyCount: Number(values[2] || 0),
        processingPendingCount: Number(values[3] || 0),
        totalApplications: Number(values[4] || 0)
      };
    })
  };
}

function getDashboardTopPrograms_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  Logger.log("[TOP PROGRAMS] whereClause: " + whereClause);
  Logger.log("[TOP PROGRAMS] Filters received: " + JSON.stringify(filters));
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      // Match Executive Dashboard query format
      var query = [
        "SELECT",
        "  program,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "GROUP BY program",
        "ORDER BY total DESC, program ASC",
        "LIMIT 10"
      ].join("\n");
      Logger.log("[TOP PROGRAMS] Query: " + query);
      return query;
    });
    Logger.log("[TOP PROGRAMS] Returned " + (result.rows ? result.rows.length : 0) + " rows from " + result.source);
    
    return {
      source: result.source,
      rows: mapBigQueryRows_(result.rows, function(values) {
        var program = String(values[0] || "").trim();
        return {
          program: program || "Unspecified Program",
          total: Number(values[1] || 0),
          value: Number(values[1] || 0) // Keep value for backward compatibility
        };
      }).map(function(row, index) {
        row.rank = index + 1;
        return row;
      })
    };
  } catch (e) {
    Logger.log("[TOP PROGRAMS] ERROR: " + e.toString());
    return { source: "error", rows: [] };
  }
}

function getDashboardLocationLoad_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(currentLocation AS STRING), 'Unknown') AS loc,",
      "  COUNT(1) AS c",
      "FROM " + source,
      "WHERE " + whereClause,
      "GROUP BY loc",
      "ORDER BY c DESC"
    ].join("\n");
  });
  return {
    source: result.source,
    rows: mapBigQueryRows_(result.rows, function(values) {
      return {
        label: String(values[0] || "Unknown"),
        value: Number(values[1] || 0)
      };
    })
  };
}

function getDashboardMonthlyVolume_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  Logger.log("[MONTHLY] whereClause: " + whereClause);
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      var query = [
        "SELECT",
        "  FORMAT_DATE('%B %Y', dateReceivedDate) AS label,",
        "  FORMAT_DATE('%Y-%m', dateReceivedDate) AS monthKey,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "AND dateReceivedDate IS NOT NULL",
        "GROUP BY label, monthKey",
        "ORDER BY monthKey DESC"
      ].join("\n");
      Logger.log("[MONTHLY] Query: " + query.substring(0, 300));
      return query;
    });
    Logger.log("[MONTHLY] Returned " + (result.rows ? result.rows.length : 0) + " rows");
    return {
      source: result.source,
      rows: mapBigQueryRows_(result.rows, function(values) {
        return {
          label: String(values[0] || ""),
          value: Number(values[2] || 0)
        };
      })
    };
  } catch (e) {
    Logger.log("[MONTHLY] ERROR: " + e.toString());
    return { source: "error", rows: [] };
  }
}

function getDashboardMonthlyTrend_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  Logger.log("[TREND] whereClause: " + whereClause);
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      var query = [
        "SELECT",
        "  FORMAT_DATE('%Y-%m', dateReceivedDate) AS monthKey,",
        "  FORMAT_DATE('%b %Y', dateReceivedDate) AS monthLabel,",
        "  SUM(CASE WHEN canonicalStatus = 'For Evaluation' THEN 1 ELSE 0 END) AS forEvaluation,",
        "  SUM(CASE WHEN canonicalStatus = 'Evaluated' THEN 1 ELSE 0 END) AS evaluated,",
        "  SUM(CASE WHEN canonicalStatus = 'Deficiency' THEN 1 ELSE 0 END) AS deficiency,",
        "  SUM(CASE WHEN canonicalStatus = 'Approved' THEN 1 ELSE 0 END) AS approved,",
        "  SUM(CASE WHEN canonicalStatus = 'Verified' THEN 1 ELSE 0 END) AS verified,",
        "  SUM(CASE WHEN canonicalStatus = 'For Signature' THEN 1 ELSE 0 END) AS forSignature,",
        "  SUM(CASE WHEN canonicalStatus = 'For Release' THEN 1 ELSE 0 END) AS forRelease,",
        "  SUM(CASE WHEN canonicalStatus = 'Released' THEN 1 ELSE 0 END) AS released,",
        "  SUM(CASE WHEN canonicalStatus = 'Cancelled' THEN 1 ELSE 0 END) AS cancelled,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "AND dateReceivedDate IS NOT NULL",
        "GROUP BY monthKey, monthLabel",
        "ORDER BY monthKey ASC"
      ].join("\n");
      Logger.log("[TREND] Query: " + query.substring(0, 300));
      return query;
    });
    Logger.log("[TREND] Returned " + (result.rows ? result.rows.length : 0) + " rows");
    
    // Transform rows into status-keyed format for frontend
    var statusKeys = ["forEvaluation", "evaluated", "deficiency", "approved", "verified", "forSignature", "forRelease", "released", "cancelled", "total"];
    var monthlyData = {};
    
    statusKeys.forEach(function(key) {
      monthlyData[key] = result.rows.map(function(row) {
        var values = row.f || [];
        var keyIndex = statusKeys.indexOf(key) + 2; // Skip monthKey (0) and monthLabel (1)
        return Number(values[keyIndex] && values[keyIndex].v || 0);
      });
    });
    
    return {
      source: result.source,
      months: result.rows.map(function(row) {
        var values = row.f || [];
        return String(values[1] && values[1].v || "");
      }),
      data: monthlyData
    };
  } catch (e) {
    Logger.log("[TREND] ERROR: " + e.toString());
    return { source: "error", months: [], data: {} };
  }
}

function getDashboardSlaRows_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  Logger.log("[SLA] whereClause: " + whereClause);
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      var query = [
        "SELECT",
        "  SUM(CASE WHEN targetReleaseDate IS NOT NULL AND canonicalStatus NOT IN ('Released', 'Cancelled', '') AND targetReleaseDate < CURRENT_DATE('Asia/Singapore') THEN 1 ELSE 0 END) AS overdue,",
        "  SUM(CASE WHEN targetReleaseDate IS NOT NULL AND canonicalStatus NOT IN ('Released', 'Cancelled', '') AND DATE_DIFF(targetReleaseDate, CURRENT_DATE('Asia/Singapore'), DAY) BETWEEN 0 AND 3 THEN 1 ELSE 0 END) AS nearingDeadline",
        "FROM " + source,
        "WHERE " + whereClause
      ].join("\n");
      Logger.log("[SLA] Query: " + query.substring(0, 300));
      return query;
    });
    Logger.log("[SLA] Returned " + (result.rows ? result.rows.length : 0) + " rows");
    
    var firstRow = result.rows && result.rows[0] ? result.rows[0].f || [] : [];
    return {
      source: result.source,
      rows: [
        { label: "Nearing Deadline", value: Number(firstRow[1] && firstRow[1].v || 0) },
        { label: "Overdue", value: Number(firstRow[0] && firstRow[0].v || 0) }
      ]
    };
  } catch (e) {
    Logger.log("[SLA] ERROR: " + e.toString());
    return { source: "error", rows: [{ label: "Nearing Deadline", value: 0 }, { label: "Overdue", value: 0 }] };
  }
}

function getDashboardBacklogRows_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(currentStageLabel AS STRING), 'Unknown') AS stage,",
      "  COUNT(1) AS c",
      "FROM " + source,
      "WHERE " + whereClause,
      "GROUP BY stage",
      "ORDER BY c DESC"
    ].join("\n");
  });
  return {
    source: result.source,
    rows: mapBigQueryRows_(result.rows, function(values) {
      return {
        label: String(values[0] || "Unknown"),
        value: Number(values[1] || 0)
      };
    })
  };
}

function getDashboardProductivityRows_(filters, options) {
  var whereClause = buildDashboardWhereClause_(filters, options);
  Logger.log("[PRODUCTIVITY] whereClause: " + whereClause);
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      var query = [
        "SELECT",
        "  TRIM(person) AS person,",
        "  COUNT(DISTINCT t0.recordId) AS totalHandled,",
        "  SUM(CASE WHEN t0.canonicalStatus = 'Released' THEN 1 ELSE 0 END) AS completed,",
        "  SUM(CASE WHEN t0.canonicalStatus != 'Released' THEN 1 ELSE 0 END) AS pending",
        "FROM " + source + " AS t0,",
        "UNNEST(SPLIT(COALESCE(t0.responsiblePeopleText, ''), ',')) AS person",
        "WHERE " + whereClause,
        "AND TRIM(person) != ''",
        "GROUP BY TRIM(person)",
        "ORDER BY totalHandled DESC"
      ].join("\n");
      Logger.log("[PRODUCTIVITY] Query: " + query.substring(0, 300));
      return query;
    });
    Logger.log("[PRODUCTIVITY] Returned " + (result.rows ? result.rows.length : 0) + " rows");
    return {
      source: result.source,
      rows: mapBigQueryRows_(result.rows, function(values) {
        var person = String(values[0] || "").trim();
        var totalHandled = Number(values[1] || 0);
        var completed = Number(values[2] || 0);
        var pending = Number(values[3] || 0);
        return {
          person: person || "Unknown",
          totalHandled: totalHandled,
          completed: completed,
          pending: pending
        };
      })
    };
  } catch (e) {
    Logger.log("[PRODUCTIVITY] ERROR: " + e.toString());
    return { source: "error", rows: [] };
  }
}

function getDashboardTurnaroundAnalytics(filters, options) {
  var turnaroundSource = getDashboardTurnaroundSource_();
  var whereClause = buildDashboardWhereClause_(filters, options);
  var holidayContextSql = "SELECT ARRAY_AGG(holidayDate) AS holidayDates FROM " + BIGQUERY_HOLIDAYS_TABLE;
  
  var stageDateExpressions = turnaroundSource.stageDateExpressions || {
    dateReceivedDate: "dateReceivedDate",
    verificationDateReceived: "CAST(NULL AS DATE)",
    issuanceDateReceived: "CAST(NULL AS DATE)",
    deficiencyDateReceived: "CAST(NULL AS DATE)",
    releaseDateReceived: "CAST(NULL AS DATE)"
  };
  var evaluationStatusExpression = turnaroundSource.evaluationStatusExpression || "CAST(NULL AS STRING)";
  
  var transitions = [
    { key: "dateToVerification", label: "Date Received -> Verification" },
    { key: "verificationToIssuance", label: "Verification -> Issuance" },
    { key: "issuanceToDeficiency", label: "Issuance -> Deficiency" },
    { key: "deficiencyToRelease", label: "Deficiency -> Release" }
  ];
  
  var businessDayExpressions = {
    total: buildBusinessDayCountSql_(
      "CASE WHEN ARRAY_LENGTH(orderedDates) >= 2 THEN orderedDates[OFFSET(0)] ELSE NULL END",
      "CASE WHEN ARRAY_LENGTH(orderedDates) >= 2 THEN orderedDates[OFFSET(ARRAY_LENGTH(orderedDates) - 1)] ELSE NULL END",
      "holidayDates"
    ),
    dateToVerification: buildBusinessDayCountSql_("dateReceivedDate", "verificationDateReceived", "holidayDates"),
    verificationToIssuance: buildBusinessDayCountSql_("verificationDateReceived", "issuanceDateReceived", "holidayDates"),
    issuanceToDeficiency: buildBusinessDayCountSql_("issuanceDateReceived", "deficiencyDateReceived", "holidayDates"),
    deficiencyToRelease: buildBusinessDayCountSql_("deficiencyDateReceived", "releaseDateReceived", "holidayDates")
  };
  
  try {
    var effectiveSource = turnaroundSource.sourceRef || BIGQUERY_NATIVE_REPORTING_VIEW;

    var transitionSelectLines = transitions.map(function(t) {
      return "    " + businessDayExpressions[t.key] + " AS " + t.key;
    });
    
    var transitionSummaryLines = transitions.map(function(t) {
      return "    AVG(" + t.key + ") AS avg_" + t.key + ",\n    COUNT(" + t.key + ") AS count_" + t.key;
    });
    
    var transitionGroupLines = transitions.map(function(t) {
      return "    AVG(" + t.key + ") AS avg_" + t.key + ",\n    COUNT(" + t.key + ") AS count_" + t.key;
    });
    
    var transitionNullLines = transitions.map(function() {
      return "  CAST(NULL AS FLOAT64),\n  CAST(NULL AS INT64)";
    });
    
    var groupedTransitionQueries = transitions.map(function(t) {
      return [
        "SELECT",
        "  'grouped-transition' AS rowType,",
        "  evaluationStatus,",
        "  '" + t.key + "' AS metricKey,",
        "  CONCAT(evaluationStatus, ' - " + t.label.replace(/'/g, "''") + "') AS metricLabel,",
        "  CAST(NULL AS INT64) AS eligibleRecordCount,",
        "  CAST(NULL AS FLOAT64) AS averageTotalDays,",
        "  CAST(NULL AS FLOAT64) AS medianTotalDays,",
        transitionNullLines.join(",\n"),
        ", statusCount,",
        "  avg_" + t.key + " AS metricAverage,",
        "  count_" + t.key + " AS metricSampleCount",
        "FROM grouped"
      ].join("\n");
    });

    var turnaroundQuery = [
      "WITH holiday_context AS (",
      "  " + holidayContextSql,
      "), filtered AS (",
      "  SELECT",
      "    COALESCE(NULLIF(TRIM(INITCAP(LOWER(" + evaluationStatusExpression + "))), ''), 'Unspecified') AS evaluationStatus,",
      "    " + stageDateExpressions.dateReceivedDate + " AS dateReceivedDate,",
      "    " + stageDateExpressions.verificationDateReceived + " AS verificationDateReceived,",
      "    " + stageDateExpressions.issuanceDateReceived + " AS issuanceDateReceived,",
      "    " + stageDateExpressions.deficiencyDateReceived + " AS deficiencyDateReceived,",
      "    " + stageDateExpressions.releaseDateReceived + " AS releaseDateReceived,",
      "    holidayDates",
      "  FROM " + effectiveSource,
      "  CROSS JOIN holiday_context",
      "  WHERE " + whereClause,
      "), prepared AS (",
      "  SELECT",
      "    evaluationStatus,",
      "    orderedDates,",
      "    holidayDates,",
      "    ARRAY_LENGTH(orderedDates) AS nonNullStageCount,",
      "    " + businessDayExpressions.total + " AS totalTurnaroundDays,",
      transitionSelectLines.join(",\n"),
      "  FROM (",
      "    SELECT",
      "      evaluationStatus,",
      "      holidayDates,",
      "      ARRAY(",
      "        SELECT stageDate",
      "        FROM UNNEST([dateReceivedDate, verificationDateReceived, issuanceDateReceived, deficiencyDateReceived, releaseDateReceived]) AS stageDate",
      "        WHERE stageDate IS NOT NULL",
      "        ORDER BY stageDate",
      "      ) AS orderedDates,",
      "      dateReceivedDate,",
      "      verificationDateReceived,",
      "      issuanceDateReceived,",
      "      deficiencyDateReceived,",
      "      releaseDateReceived",
      "    FROM filtered",
      "  )",
      "), eligible AS (",
      "  SELECT *",
      "  FROM prepared",
      "  WHERE nonNullStageCount >= 2",
      "), summary AS (",
      "  SELECT",
      "    COUNT(*) AS eligibleRecordCount,",
      "    AVG(totalTurnaroundDays) AS averageTotalDays,",
      "    APPROX_QUANTILES(totalTurnaroundDays, 100)[OFFSET(50)] AS medianTotalDays,",
      transitionSummaryLines.join(",\n"),
      "  FROM eligible",
      "), grouped AS (",
      "  SELECT",
      "    evaluationStatus,",
      "    COUNT(*) AS statusCount,",
      "    AVG(totalTurnaroundDays) AS avgTotalDays,",
      transitionGroupLines.join(",\n"),
      "  FROM eligible",
      "  GROUP BY evaluationStatus",
      ")",
      "SELECT",
      "  'summary' AS rowType,",
      "  CAST(NULL AS STRING) AS evaluationStatus,",
      "  CAST(NULL AS STRING) AS metricKey,",
      "  CAST(NULL AS STRING) AS metricLabel,",
      "  eligibleRecordCount,",
      "  averageTotalDays,",
      "  medianTotalDays,",
      transitionNullLines.join(",\n"),
      ", CAST(NULL AS INT64) AS statusCount,",
      "  CAST(NULL AS FLOAT64) AS metricAverage,",
      "  CAST(NULL AS INT64) AS metricSampleCount",
      "FROM summary",
      "UNION ALL",
      "SELECT",
      "  'grouped-total' AS rowType,",
      "  evaluationStatus,",
      "  'total' AS metricKey,",
      "  CONCAT(evaluationStatus, ' - Total Turnaround') AS metricLabel,",
      "  CAST(NULL AS INT64) AS eligibleRecordCount,",
      "  CAST(NULL AS FLOAT64) AS averageTotalDays,",
      "  CAST(NULL AS FLOAT64) AS medianTotalDays,",
      transitionNullLines.join(",\n"),
      ", statusCount,",
      "  avgTotalDays AS metricAverage,",
      "  statusCount AS metricSampleCount",
      "FROM grouped",
      "UNION ALL",
      groupedTransitionQueries.join("\nUNION ALL\n")
    ].join("\n");

    Logger.log("[TURNAROUND] Query length: " + turnaroundQuery.length);

    var rawRows = runBigQueryJobAsync_(turnaroundQuery);

    Logger.log("[TURNAROUND] Returned " + rawRows.length + " rows");

    if (!rawRows.length) {
      return {
        source: effectiveSource,
        summary: { eligibleRecordCount: 0, averageTotalDays: 0, medianTotalDays: 0 },
        highlights: [],
        groups: []
      };
    }

    var parsedRows = rawRows.map(function(row) {
      var values = row.f || [];
      return {
        rowType: String(values[0] ? values[0].v : ""),
        evaluationStatus: String(values[1] ? values[1].v : ""),
        metricKey: String(values[2] ? values[2].v : ""),
        metricLabel: String(values[3] ? values[3].v : ""),
        eligibleRecordCount: Number(values[4] ? values[4].v : 0),
        averageTotalDays: Number(values[5] ? values[5].v : 0),
        medianTotalDays: Number(values[6] ? values[6].v : 0),
        avgDateToVerification: Number(values[7] ? values[7].v : 0),
        countDateToVerification: Number(values[8] ? values[8].v : 0),
        avgVerificationToIssuance: Number(values[9] ? values[9].v : 0),
        countVerificationToIssuance: Number(values[10] ? values[10].v : 0),
        avgIssuanceToDeficiency: Number(values[11] ? values[11].v : 0),
        countIssuanceToDeficiency: Number(values[12] ? values[12].v : 0),
        avgDeficiencyToRelease: Number(values[13] ? values[13].v : 0),
        countDeficiencyToRelease: Number(values[14] ? values[14].v : 0),
        statusCount: Number(values[15] ? values[15].v : 0),
        metricAverage: Number(values[16] ? values[16].v : 0),
        metricSampleCount: Number(values[17] ? values[17].v : 0)
      };
    });

    var summaryRow = { eligibleRecordCount: 0, averageTotalDays: 0, medianTotalDays: 0 };
    for (var i = 0; i < parsedRows.length; i++) {
      if (parsedRows[i].rowType === "summary") {
        summaryRow = parsedRows[i];
        break;
      }
    }

    var summary = {
      eligibleRecordCount: summaryRow.eligibleRecordCount,
      averageTotalDays: summaryRow.averageTotalDays,
      medianTotalDays: summaryRow.medianTotalDays,
    };

    var highlights = [];
    var transitionData = [
      { key: "dateToVerification", label: "Date Received -> Verification", avg: summaryRow.avgDateToVerification, count: summaryRow.countDateToVerification },
      { key: "verificationToIssuance", label: "Verification -> Issuance", avg: summaryRow.avgVerificationToIssuance, count: summaryRow.countVerificationToIssuance },
      { key: "issuanceToDeficiency", label: "Issuance -> Deficiency", avg: summaryRow.avgIssuanceToDeficiency, count: summaryRow.countIssuanceToDeficiency },
      { key: "deficiencyToRelease", label: "Deficiency -> Release", avg: summaryRow.avgDeficiencyToRelease, count: summaryRow.countDeficiencyToRelease }
    ];

    transitionData.forEach(function(t) {
      if (t.count > 0) {
        highlights.push({
          label: t.label,
          primary: Math.round(t.avg * 10) / 10 + " days",
          secondary: t.count + " records",
          avgDays: t.avg,
          sampleCount: t.count
        });
      }
    });

    var groupedMetrics = parsedRows.filter(function(r) {
      return r.rowType === "grouped-total" || r.rowType === "grouped-transition";
    });

    var groupsMap = {};
    groupedMetrics.forEach(function(row) {
      var status = row.evaluationStatus || "Unspecified";
      if (!groupsMap[status]) {
        groupsMap[status] = { label: status, sampleCount: 0, metrics: [] };
      }
      var metricAvg = row.metricAverage || 0;
      var metricCount = row.metricSampleCount || 0;
      if (metricCount <= 0 || metricAvg <= 0) return;
      if (row.metricKey === "total") {
        groupsMap[status].sampleCount = Math.max(groupsMap[status].sampleCount, metricCount);
        groupsMap[status].metrics.unshift({
          key: "total",
          label: "Total Turnaround",
          averageDays: metricAvg,
          sampleCount: metricCount,
          isTotal: true
        });
      } else {
        var transition = null;
        for (var tIdx = 0; tIdx < transitions.length; tIdx++) {
          if (transitions[tIdx].key === row.metricKey) {
            transition = transitions[tIdx];
            break;
          }
        }
        if (transition) {
          groupsMap[status].sampleCount = Math.max(groupsMap[status].sampleCount, metricCount);
          groupsMap[status].metrics.push({
            key: row.metricKey,
            label: transition.label,
            averageDays: metricAvg,
            sampleCount: metricCount,
            isTotal: false
          });
        }
      }
    });

    var groups = Object.keys(groupsMap).map(function(key) {
      return groupsMap[key];
    }).sort(function(a, b) {
      return b.sampleCount - a.sampleCount;
    });

    // Largest Workload Status
    var largestWorkloadGroup = { label: "", sampleCount: 0, metrics: [] };
    var foundLargest = false;
    for (var g = 0; g < groups.length; g++) {
      if (!foundLargest || groups[g].sampleCount > largestWorkloadGroup.sampleCount) {
        largestWorkloadGroup = groups[g];
        foundLargest = true;
      }
    }

    // 3. Slowest transition - find from highlights first, then fall back to group metrics
    var slowestTransition = { label: "", avg: 0, count: 0, key: "" };
    var foundSlowest = false;
    
    // Search in highlights first
    for (var hi = 0; hi < highlights.length; hi++) {
      var h = highlights[hi];
      if (h.avgDays > 0 && h.sampleCount > 0 && (!foundSlowest || h.avgDays > slowestTransition.avg)) {
        slowestTransition = { label: h.label, avg: h.avgDays, count: h.sampleCount, key: h.label };
        foundSlowest = true;
      }
    }
    
    // If not found in highlights, search in group metrics (transitions)
    if (!foundSlowest) {
      for (var gi = 0; gi < groups.length; gi++) {
        var grp = groups[gi];
        for (var mi = 0; mi < grp.metrics.length; mi++) {
          var m = grp.metrics[mi];
          if (m.averageDays > 0 && !m.isTotal && (!foundSlowest || m.averageDays > slowestTransition.avg)) {
            slowestTransition = { 
              label: grp.label + " → " + m.label, 
              avg: m.averageDays, 
              count: m.sampleCount, 
              key: m.key 
            };
            foundSlowest = true;
          }
        }
      }
    }

    // Patch summary with derived fields now that groups and transitions are resolved
    summary.largestWorkloadStatus = foundLargest ? largestWorkloadGroup.label : "—";
    summary.largestWorkloadCount = foundLargest ? largestWorkloadGroup.sampleCount : 0;
    summary.slowestTransitionLabel = foundSlowest ? slowestTransition.label : "—";
    summary.slowestTransitionAvg = foundSlowest ? slowestTransition.avg : 0;

    return {
      source: effectiveSource,
      summary: summary,
      highlights: highlights,
      groups: groups
    };

  } catch (e) {
    Logger.log("[TURNAROUND] ERROR: " + e.toString());
    return {
      source: "error",
      summary: { eligibleRecordCount: 0, averageTotalDays: 0, medianTotalDays: 0 },
      highlights: [],
      groups: []
    };
  }
}

function runBigQueryJobAsync_(query) {
  var jobResource = {
    configuration: {
      query: {
        query: query,
        useLegacySql: false
      }
    }
  };
  
  var job = BigQuery.Jobs.insert(jobResource, BIGQUERY_PROJECT_ID);
  var jobId = job.jobReference.jobId;
  var location = job.jobReference.location || BIGQUERY_LOCATION;
  
  Logger.log("[ASYNC JOB] Started jobId: " + jobId);
  
  var attempts = 0;
  var maxAttempts = 30;
  while (attempts < maxAttempts) {
    Utilities.sleep(1000);
    var status = BigQuery.Jobs.get(BIGQUERY_PROJECT_ID, jobId, { location: location });
    var state = status.status && status.status.state;
    Logger.log("[ASYNC JOB] Attempt " + attempts + " state: " + state);
    
    if (state === "DONE") {
      if (status.status.errorResult) {
        throw new Error("BigQuery job failed: " + JSON.stringify(status.status.errorResult));
      }
      break;
    }
    attempts += 1;
  }
  
  if (attempts >= maxAttempts) {
    throw new Error("BigQuery job timed out after " + maxAttempts + " seconds.");
  }
  
  // Collect all pages of results
  var rows = [];
  var pageToken = null;
  do {
    var queryParams = { location: location, maxResults: 10000 };
    if (pageToken) queryParams.pageToken = pageToken;
    var resultPage = BigQuery.Jobs.getQueryResults(BIGQUERY_PROJECT_ID, jobId, queryParams);
    if (resultPage.rows) {
      rows = rows.concat(resultPage.rows);
    }
    pageToken = resultPage.pageToken || null;
  } while (pageToken);
  
  Logger.log("[ASYNC JOB] Complete. Total rows: " + rows.length);
  return rows;
}

// Turnaround Analytics helper functions (from Executive Dashboard)
function bigQueryTableExists_(datasetId, tableName) {
  try {
    // Try to query the table to check if it exists
    var query = "SELECT 1 FROM `" + BIGQUERY_PROJECT_ID + "." + datasetId + "." + tableName + "` LIMIT 0";
    BigQuery.Jobs.query({ query: query, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    return true;
  } catch (e) {
    return false;
  }
}

function getBigQueryTableColumns_(datasetId, tableName) {
  try {
    var query = [
      "SELECT column_name, data_type",
      "FROM `" + BIGQUERY_PROJECT_ID + "." + datasetId + ".INFORMATION_SCHEMA.COLUMNS`",
      "WHERE table_name = '" + tableName + "'"
    ].join("\n");
    var result = BigQuery.Jobs.query({ query: query, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    var columns = {};
    if (result.rows) {
      result.rows.forEach(function(row) {
        columns[String(row.f[0].v)] = String(row.f[1].v);
      });
    }
    return columns;
  } catch (e) {
    Logger.log("[COLUMNS] Error getting columns: " + e.toString());
    return {};
  }
}

function buildBusinessDayCountSql_(startExpression, endExpression, holidayArrayExpression) {
  var holidayFilterSql = holidayArrayExpression
    ? "AND businessDate NOT IN UNNEST(COALESCE(" + holidayArrayExpression + ", CAST([] AS ARRAY<DATE>)))"
    : "";
  return [
    "(SELECT COUNT(*) FROM UNNEST(GENERATE_DATE_ARRAY(",
    "  CAST(" + startExpression + " AS DATE),",
    "  CAST(" + endExpression + " AS DATE),",
    "  INTERVAL 1 DAY",
    ")) AS businessDate",
    "WHERE EXTRACT(DAYOFWEEK FROM businessDate) BETWEEN 2 AND 6", // Monday=2 to Friday=6
    holidayFilterSql,
    ")"
  ].join("\n");
}

function getDashboardTurnaroundSource_() {
  var datasetRef = BIGQUERY_PROJECT_ID + "." + BIGQUERY_DATASET_ID;
  var candidates = [
    { tableName: "reporting_flat_native", sourceRef: "`" + datasetRef + ".reporting_flat_native`", sourceKind: "native" },
    { tableName: "reporting_flat_native_view", sourceRef: "`" + datasetRef + ".reporting_flat_native_view`", sourceKind: "nativeView" },
    { tableName: "reporting_flat_native_plain", sourceRef: "`" + datasetRef + ".reporting_flat_native_plain`", sourceKind: "plain" }
  ];
  var stageSpecs = [
    { sourceColumn: "dateReceivedDate", outputName: "dateReceivedDate" },
    { sourceColumn: "dateReceived", outputName: "dateReceivedDate" },
    { sourceColumn: "verificationDateReceived", outputName: "verificationDateReceived" },
    { sourceColumn: "issuanceDateReceived", outputName: "issuanceDateReceived" },
    { sourceColumn: "deficiencyDateReceived", outputName: "deficiencyDateReceived" },
    { sourceColumn: "releaseDateReceived", outputName: "releaseDateReceived" }
  ];

  for (var i = 0; i < candidates.length; i++) {
    var candidate = candidates[i];
    var columns = getBigQueryTableColumns_(BIGQUERY_DATASET_ID, candidate.tableName);
    var expressions = {
      dateReceivedDate: "CAST(NULL AS DATE)",
      verificationDateReceived: "CAST(NULL AS DATE)",
      issuanceDateReceived: "CAST(NULL AS DATE)",
      deficiencyDateReceived: "CAST(NULL AS DATE)",
      releaseDateReceived: "CAST(NULL AS DATE)"
    };

    stageSpecs.forEach(function(spec) {
      if (!columns[spec.sourceColumn]) return;
      expressions[spec.outputName] = "SAFE_CAST(" + spec.sourceColumn + " AS DATE)";
    });

    if (expressions.dateReceivedDate === "CAST(NULL AS DATE)") continue;

    return {
      sourceRef: candidate.sourceRef,
      sourceKind: candidate.sourceKind,
      stageDateExpressions: expressions,
      evaluationStatusExpression: columns.evaluationStatus ? "CAST(evaluationStatus AS STRING)" : "CAST(NULL AS STRING)"
    };
  }

  // Fallback to default view
  return {
    sourceRef: BIGQUERY_NATIVE_REPORTING_VIEW,
    sourceKind: "runtime",
    stageDateExpressions: {
      dateReceivedDate: "dateReceivedDate",
      verificationDateReceived: "CAST(NULL AS DATE)",
      issuanceDateReceived: "CAST(NULL AS DATE)",
      deficiencyDateReceived: "CAST(NULL AS DATE)",
      releaseDateReceived: "CAST(NULL AS DATE)"
    },
    evaluationStatusExpression: "CAST(NULL AS STRING)"
  };
}

// Dashboard cache constants
var DASHBOARD_CACHE_TTL_SECONDS = 120; // 2 minutes - fresher data
var DASHBOARD_CACHE_KEY_PREFIX = "dashboard:v1:";

function getDashboardCacheKey_(filters) {
  // Create a stable cache key from filter values
  var keyParts = [
    filters.datePreset || "allTime",
    filters.dateFrom || "",
    filters.dateTo || "",
    filters.program || "ALL",
    filters.es2InCharge || "ALL",
    filters.applicationStatus || "ALL",
    filters.transactionType || "ALL",
    filters.hei || "ALL"
  ];
  return DASHBOARD_CACHE_KEY_PREFIX + keyParts.join(":");
}

function getCachedDashboardData_(cacheKey) {
  try {
    var cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      Logger.log("[CACHE] Cache HIT for key: " + cacheKey);
      console.info("[CACHE] Cache HIT for key: " + cacheKey)
      return JSON.parse(cached);
    }
    Logger.log("[CACHE] Cache MISS for key: " + cacheKey);
    console.info("[CACHE] Cache MISS for key: " + cacheKey);
    return null;
  } catch (e) {
    Logger.log("[CACHE] Error reading cache: " + e.toString());
    console.info("[CACHE] Error reading cache: " + e.toString());
    return null;
  }
}

function setCachedDashboardData_(cacheKey, data) {
  try {
    var cache = CacheService.getScriptCache();
    // Ensure data is serializable
    var serialized = JSON.stringify(data);
    // Check size (CacheService has 100KB limit per key)
    if (serialized.length > 90000) {
      Logger.log("[CACHE] Data too large to cache (>90KB), skipping cache write");
      return false;
    }
    cache.put(cacheKey, serialized, DASHBOARD_CACHE_TTL_SECONDS);
    Logger.log("[CACHE] Stored data in cache for key: " + cacheKey + " (TTL: " + DASHBOARD_CACHE_TTL_SECONDS + "s)");
    return true;
  } catch (e) {
    Logger.log("[CACHE] Error writing cache: " + e.toString());
    return false;
  }
}

function clearDashboardCache_() {
  try {
    var cache = CacheService.getScriptCache();
    // Note: CacheService doesn't support wildcard deletion, so we just log this
    Logger.log("[CACHE] Cache clear requested. Individual keys will expire after TTL.");
    return true;
  } catch (e) {
    Logger.log("[CACHE] Error clearing cache: " + e.toString());
    return false;
  }
}

// PropertiesService Cache (500KB limit - 5x larger than CacheService)
var DASHBOARD_PROPERTIES_CACHE_KEY = "dashboard:v1:data";
var DASHBOARD_PROPERTIES_TIMESTAMP_KEY = "dashboard:v1:timestamp";
var DASHBOARD_PROPERTIES_CACHE_MAX_AGE_MS = 2 * 60 * 1000; // 2 minutes - fresher data

function compressDashboardData_(data) {
  // Remove empty arrays and redundant fields to reduce size
  return JSON.stringify(data, function(key, value) {
    if (Array.isArray(value) && value.length === 0) return undefined;
    if (value === null) return undefined;
    if (value === "") return undefined;
    // Remove internal meta fields that can be regenerated
    if (key === "loadedAt") return undefined;
    return value;
  });
}

function decompressDashboardData_(compressed) {
  return JSON.parse(compressed);
}

function setDashboardPropertiesCache_(data) {
  try {
    var props = PropertiesService.getScriptProperties();
    var compressed = compressDashboardData_(data);
    var size = compressed.length;
    // PropertiesService has 500KB limit per key
    if (size > 450000) {
      Logger.log("[PROPERTIES_CACHE] Data too large (>450KB), skipping cache write. Size: " + size);
      return false;
    }
    props.setProperty(DASHBOARD_PROPERTIES_CACHE_KEY, compressed);
    props.setProperty(DASHBOARD_PROPERTIES_TIMESTAMP_KEY, String(Date.now()));
    Logger.log("[PROPERTIES_CACHE] Stored " + size + " bytes (TTL: " + (DASHBOARD_PROPERTIES_CACHE_MAX_AGE_MS / 60000) + " min)");
    return true;
  } catch (e) {
    Logger.log("[PROPERTIES_CACHE] Error writing cache: " + e.toString());
    return false;
  }
}

function getDashboardPropertiesCache_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var compressed = props.getProperty(DASHBOARD_PROPERTIES_CACHE_KEY);
    var timestamp = props.getProperty(DASHBOARD_PROPERTIES_TIMESTAMP_KEY);
    if (!compressed || !timestamp) {
      Logger.log("[PROPERTIES_CACHE] Cache MISS");
      return null;
    }
    var age = Date.now() - Number(timestamp);
    if (age > DASHBOARD_PROPERTIES_CACHE_MAX_AGE_MS) {
      Logger.log("[PROPERTIES_CACHE] Cache EXPIRED (age: " + (age / 1000) + "s)");
      return null;
    }
    var data = decompressDashboardData_(compressed);
    Logger.log("[PROPERTIES_CACHE] Cache HIT (age: " + (age / 1000) + "s, size: " + compressed.length + " bytes)");
    return data;
  } catch (e) {
    Logger.log("[PROPERTIES_CACHE] Error reading cache: " + e.toString());
    return null;
  }
}

function clearDashboardPropertiesCache_() {
  try {
    var props = PropertiesService.getScriptProperties();
    props.deleteProperty(DASHBOARD_PROPERTIES_CACHE_KEY);
    props.deleteProperty(DASHBOARD_PROPERTIES_TIMESTAMP_KEY);
    Logger.log("[PROPERTIES_CACHE] Cleared");
    return true;
  } catch (e) {
    Logger.log("[PROPERTIES_CACHE] Error clearing cache: " + e.toString());
    return false;
  }
}

function testCacheLimits_() {
  var cache = CacheService.getScriptCache();
  var results = [];
  
  // Test 1: Cache size limits
  var testSizes = [10000, 50000, 90000, 95000, 100000, 150000];
  testSizes.forEach(function(size) {
    try {
      var testData = { records: new Array(Math.floor(size / 50)).join("x"), timestamp: new Date().toISOString() };
      var serialized = JSON.stringify(testData);
      var key = "test:size" + size;
      cache.put(key, serialized, 60);
      var retrieved = cache.get(key);
      var success = retrieved === serialized;
      results.push("Size " + size + " bytes: " + (success ? "PASS" : "FAIL") + " (actual: " + serialized.length + ")");
    } catch (e) {
      results.push("Size " + size + " bytes: ERROR - " + e.message);
    }
  });
  
  // Test 2: Cache performance
  var iterations = 100;
  var start = Date.now();
  for (var i = 0; i < iterations; i++) {
    var testKey = "perf:test" + (i % 10);
    cache.put(testKey, JSON.stringify({ index: i, data: "test" }), 60);
    cache.get(testKey);
  }
  var duration = Date.now() - start;
  results.push("Performance: " + iterations + " ops in " + duration + "ms (" + (duration / iterations).toFixed(2) + "ms/op)");
  
  // Test 3: Dashboard data caching
  var mockDashboardData = {
    ok: true,
    filters: { programs: [], es2Owners: [] },
    data: { records: [], availableYears: [], es2Owners: [], es2ProgramMap: {}, programOwnershipMap: {} },
    summary: { total: 100, forEvaluation: 20, evaluated: 30 },
    sections: {
      summary: { rows: [] },
      es2Ownership: { rows: (function() { var arr = []; for (var i = 0; i < 20; i++) arr.push({ es2: "Test", totalApplications: 50 }); return arr; })() },
      monthlyVolume: { rows: (function() { var arr = []; for (var i = 0; i < 24; i++) arr.push({ label: "Jan 2024", value: 100 }); return arr; })() },
      locationLoad: { rows: (function() { var arr = []; for (var i = 0; i < 10; i++) arr.push({ label: "Location", value: 50 }); return arr; })() },
      stageBacklog: { rows: [] },
      slaExposure: { rows: [] },
      turnaroundTime: { rows: [], summary: {} },
      productivity: { rows: (function() { var arr = []; for (var i = 0; i < 50; i++) arr.push({ person: "Staff", completed: 10, pending: 5 }); return arr; })() }
    },
    referenceData: { es2Owners: [], es2ProgramMap: {}, programOwnershipMap: {} },
    meta: { loadedAt: new Date().toISOString(), recordCount: 100, totalRecordCount: 100, availableYears: [] }
  };
  
  var dashboardSerialized = JSON.stringify(mockDashboardData);
  results.push("Mock dashboard data size: " + dashboardSerialized.length + " bytes");
  
  if (dashboardSerialized.length <= 90000) {
    try {
      var cacheKey = "test:dashboard";
      cache.put(cacheKey, dashboardSerialized, 60);
      var retrieved = cache.get(cacheKey);
      results.push("CacheService (100KB) test: " + (retrieved === dashboardSerialized ? "PASS" : "FAIL"));
    } catch (e) {
      results.push("CacheService (100KB) test: ERROR - " + e.message);
    }
  } else {
    results.push("CacheService (100KB) test: SKIPPED (too large)");
  }
  
  // Test 4: PropertiesService (500KB limit)
  if (dashboardSerialized.length <= 450000) {
    try {
      var props = PropertiesService.getScriptProperties();
      props.setProperty("test:dashboard", dashboardSerialized);
      props.setProperty("test:timestamp", String(Date.now()));
      var propRetrieved = props.getProperty("test:dashboard");
      results.push("PropertiesService (500KB) test: " + (propRetrieved === dashboardSerialized ? "PASS" : "FAIL") + " (size: " + dashboardSerialized.length + " bytes)");
      // Cleanup
      props.deleteProperty("test:dashboard");
      props.deleteProperty("test:timestamp");
    } catch (e) {
      results.push("PropertiesService (500KB) test: ERROR - " + e.message);
    }
  } else {
    results.push("PropertiesService (500KB) test: SKIPPED (too large)");
  }
  
  // Test 5: Compression efficiency
  var compressed = compressDashboardData_(mockDashboardData);
  results.push("Compression efficiency: " + dashboardSerialized.length + " -> " + compressed.length + " bytes (" + (100 - (compressed.length / dashboardSerialized.length * 100)).toFixed(1) + "% reduction)");
  
  Logger.log("[CACHE TEST]\n" + results.join("\n"));
  return results;
}

function dashboard_fetchData_(payload) {
  try {
    var rawFilters = payload || {};
    
    Logger.log("[FETCH] Raw payload received: " + JSON.stringify(rawFilters));
    
    // Extract programOwnershipMap from payload for ES2/Program alignment
    var programOwnershipMap = rawFilters.programOwnershipMap || {};
    Logger.log("[FETCH] programOwnershipMap entries: " + Object.keys(programOwnershipMap).length);
    
    // Normalize filters using Executive Dashboard logic (resolves date presets to actual dates)
    var filters = normalizeFilters_(rawFilters);
    
    // Build filter options with ownership map for alignment
    var filterOptions = { programOwnershipMap: programOwnershipMap };
    
    Logger.log("[FETCH] Normalized filters: " + JSON.stringify(filters));
    Logger.log("[FETCH] es2InCharge value: '" + filters.es2InCharge + "' (length: " + (filters.es2InCharge ? filters.es2InCharge.length : 0) + ")");
    
    // Check for live mode (bypass all caches)
    var liveMode = rawFilters._liveMode === true;
    if (liveMode) {
      Logger.log("[FETCH] Live mode enabled - bypassing all caches");
    }
    
    // Check cache first (use normalized filters for cache key)
    var cacheKey = getDashboardCacheKey_(filters);
    Logger.log("[FETCH] Cache key: " + cacheKey);
    
    // Check CacheService first (fastest, 100KB limit) - skip if live mode
    if (!liveMode) {
      var cachedData = getCachedDashboardData_(cacheKey);
      if (cachedData && cachedData.ok) {
        Logger.log("[FETCH] Returning cached data from CacheService for key: " + cacheKey);
        return cachedData;
      }
    }
    
    // Check PropertiesService for default filter (thisYear) - stores up to 500KB - skip if live mode
    if (!liveMode && filters.datePreset === "thisYear" && !filters.program && !filters.es2InCharge && 
        !filters.applicationStatus && !filters.transactionType && !filters.hei) {
      var propertiesCachedData = getDashboardPropertiesCache_();
      if (propertiesCachedData && propertiesCachedData.ok) {
        Logger.log("[FETCH] Returning cached data from PropertiesService (default filters)");
        return propertiesCachedData;
      }
    }
    
    Logger.log("[FETCH] Cache miss for key: " + cacheKey);
    
    var summary, workload, status, turnaround, topPrograms, locationLoad, monthlyVolume, monthlyTrend, slaExposure, stageBacklog, productivity;
    var startTime = Date.now();
    
    try {
      Logger.log("[FETCH] Starting getDashboardSummary_...");
      summary = getDashboardSummary_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardSummary_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardWorkloadSummary_...");
      workload = getDashboardWorkloadSummary_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardWorkloadSummary_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardStatusDistribution_...");
      status = getDashboardStatusDistribution_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardStatusDistribution_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardTurnaroundAnalytics...");
      turnaround = getDashboardTurnaroundAnalytics(filters, filterOptions);
      Logger.log("[FETCH] getDashboardTurnaroundAnalytics done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardTopPrograms_...");
      topPrograms = getDashboardTopPrograms_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardTopPrograms_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardLocationLoad_...");
      locationLoad = getDashboardLocationLoad_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardLocationLoad_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardMonthlyVolume_...");
      monthlyVolume = getDashboardMonthlyVolume_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardMonthlyVolume_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardMonthlyTrend_...");
      monthlyTrend = getDashboardMonthlyTrend_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardMonthlyTrend_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardSlaRows_...");
      slaExposure = getDashboardSlaRows_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardSlaRows_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardBacklogRows_...");
      stageBacklog = getDashboardBacklogRows_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardBacklogRows_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] Starting getDashboardProductivityRows_...");
      productivity = getDashboardProductivityRows_(filters, filterOptions);
      Logger.log("[FETCH] getDashboardProductivityRows_ done in " + (Date.now() - startTime) + "ms");
      
      Logger.log("[FETCH] All sub-fetches completed in " + (Date.now() - startTime) + "ms");
    } catch (e) {
      Logger.log("[FETCH] Critical error in sub-fetches after " + (Date.now() - startTime) + "ms: " + e.toString());
      // Continue with empty results rather than failing the whole request
    }

    // Build filter options from BigQuery (more reliable than Static sheet)
    var fetchedFilterOptions, availableYears;
    try {
      fetchedFilterOptions = getDashboardFilterOptionsFromBigQuery_();
      availableYears = getBigQueryAvailableYears_();
      Logger.log("[FETCH] Available years: " + availableYears.length);
    } catch (filterError) {
      Logger.log("[FETCH] Filter options fetch failed: " + filterError.toString());
      fetchedFilterOptions = { programs: [], es2Owners: [], es2ProgramMap: {}, programOwnershipMap: {} };
      availableYears = [];
    }
    
    Logger.log("[FETCH] Filter options from BigQuery: es2Owners=" + fetchedFilterOptions.es2Owners.length + ", programs=" + fetchedFilterOptions.programs.length);

    var response = {
    ok: true,
    filters: fetchedFilterOptions,
    data: {
      records: [],
      availableYears: availableYears || [],
      es2Owners: (fetchedFilterOptions && fetchedFilterOptions.es2Owners) || [],
      es2ProgramMap: (fetchedFilterOptions && fetchedFilterOptions.es2ProgramMap) || {},
      programOwnershipMap: (fetchedFilterOptions && fetchedFilterOptions.programOwnershipMap) || {}
    },
    workload: workload || [],
    status: status || [],
    turnaround: turnaround || {},
    topPrograms: topPrograms || [],
    locationLoad: locationLoad || [],
    monthlyVolume: monthlyVolume || [],
    monthlyTrend: monthlyTrend || [],
    slaExposure: slaExposure || [],
    stageBacklog: stageBacklog || [],
    productivity: productivity || [],
    summary: {
      totalRecords: Number((summary && summary.values && summary.values.totalCount) || 0),
      total: Number((summary && summary.values && summary.values.totalCount) || 0),
      forEvaluation: Number((summary && summary.values && summary.values.forEvaluationCount) || 0),
      evaluated: Number((summary && summary.values && summary.values.evaluatedCount) || 0),
      deficiency: Number((summary && summary.values && summary.values.deficiencyCount) || 0),
      approved: Number((summary && summary.values && summary.values.approvedCount) || 0),
      verified: Number((summary && summary.values && summary.values.verifiedCount) || 0),
      forSignature: Number((summary && summary.values && summary.values.forSignatureCount) || 0),
      forRelease: Number((summary && summary.values && summary.values.forReleaseCount) || 0),
      released: Number((summary && summary.values && summary.values.releasedCount) || 0),
      cancelled: Number((summary && summary.values && summary.values.cancelledCount) || 0)
    },
    sections: {
      es2Ownership: { title: "Workload by ES II Owner", rows: (workload && workload.rows) || [] },
      topPrograms: { title: "Top 10 Programs", rows: (topPrograms && topPrograms.rows) || [] },
      overallApplicationStatus: { title: "Overall Application Status", rows: (status && status.rows) || [] },
      statusProgression: { title: "Status Progression Overview", rows: (status && status.rows) || [] },
      turnaroundTime: turnaround || { title: "Turnaround Time", rows: [] },
      locationLoad: { title: "Document Location Load", rows: (locationLoad && locationLoad.rows) || [] },
      monthlyVolume: { title: "Monthly Application Volume", rows: (monthlyVolume && monthlyVolume.rows) || [] },
      slaExposure: { title: "SLA Exposure", rows: (slaExposure && slaExposure.rows) || [] },
      stageBacklog: { title: "Workflow Backlog", rows: (stageBacklog && stageBacklog.rows) || [] },
      productivity: { title: "Assigned Staff Activity", rows: (productivity && productivity.rows) || [] },
      monthlyTrend: { 
        title: "Application Status Trend", 
        months: (monthlyTrend && monthlyTrend.months) || [],
        data: (monthlyTrend && monthlyTrend.data) || {}
      }
    },
    referenceData: {
      availableYears: availableYears || [],
      es2Owners: (filterOptions && filterOptions.es2Owners) || [],
      es2ProgramMap: (filterOptions && filterOptions.es2ProgramMap) || {},
      programOwnershipMap: (filterOptions && filterOptions.programOwnershipMap) || {}
    },
    meta: {
      source: (summary && summary.source) || (workload && workload.source) || "",
      dataset: BIGQUERY_DATASET_ID,
      loadedAt: new Date().toISOString(),
      recordCount: Number((summary && summary.values && summary.values.totalCount) || 0),
      totalRecordCount: Number((summary && summary.values && summary.values.totalCount) || 0),
      appliedFilters: filters,
      availableYears: availableYears || []
    }
  };
  
  Logger.log("[FETCH] Summary total: " + (summary && summary.values && summary.values.totalCount) + ", Filtered es2InCharge: " + filters.es2InCharge);
  
  // Store in cache for subsequent requests (skip if live mode)
  if (!liveMode) {
    Logger.log("[FETCH] Storing in cache with key: " + cacheKey);
    setCachedDashboardData_(cacheKey, response);
    
    // Store in PropertiesService for default filters (supports up to 500KB)
    if (filters.datePreset === "thisYear" && !filters.program && !filters.es2InCharge && 
        !filters.applicationStatus && !filters.transactionType && !filters.hei) {
      setDashboardPropertiesCache_(response);
    }
  } else {
    Logger.log("[FETCH] Live mode - skipping cache storage");
  }
  
    // FINAL HARDENING: Ensure the object is clean and serializable for the GAS bridge
    try {
      Logger.log("[FETCH] About to serialize response, ok=" + response.ok);
      var serialized = JSON.stringify(response);
      Logger.log("[FETCH] Serialized length: " + serialized.length);
      var parsed = JSON.parse(serialized);
      Logger.log("[FETCH] Returning parsed response");
      return parsed;
    } catch (serializationError) {
      Logger.log("[FETCH] Serialization failure: " + serializationError.toString());
      return { ok: false, error: "Data serialization failed." };
    }
  } catch (topLevelError) {
    Logger.log("[FETCH] CRITICAL TOP-LEVEL ERROR: " + topLevelError.toString());
    return { ok: false, error: "Dashboard fetch failed: " + topLevelError.message };
  }
}

function dashboard_getAnalyticsData_(filters) {
  return dashboard_fetchData_(filters);
}

function dashboard_refreshDashboardCache_() {
  return dashboard_fetchData_({});
}

function getDashboardFilterOptionsFromBigQuery_() {
  try {
    // Get distinct ES2 Owners (without CAST to match Executive Dashboard)
    var es2Result = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT es2InCharge as value",
        "FROM " + source,
        "WHERE es2InCharge IS NOT NULL",
        "ORDER BY value"
      ].join("\n");
    });
    var es2Owners = mapBigQueryRows_(es2Result.rows, function(values) {
      return String(values[0] || "");
    }).filter(function(v) { return v !== ""; });

    // Get distinct Programs (without CAST)
    var programResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT program as value",
        "FROM " + source,
        "WHERE program IS NOT NULL",
        "ORDER BY value"
      ].join("\n");
    });
    var programs = mapBigQueryRows_(programResult.rows, function(values) {
      return String(values[0] || "");
    }).filter(function(v) { return v !== ""; });

    // Get distinct Application Statuses (without CAST)
    // Get distinct Application Statuses - canonicalized
    var statusResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT " + getDashboardCanonicalStatusExpression_() + " AS value",
        "FROM " + source,
        "WHERE applicationStatus IS NOT NULL",
        "ORDER BY value"
      ].join("\n");
    });
    var applicationStatuses = mapBigQueryRows_(statusResult.rows, function(values) {
      return String(values[0] || "");
    }).filter(function(v) { return v !== ""; });

    // Get Program-ES2 Ownership mapping for cascading filters
    var ownershipResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT program, es2InCharge",
        "FROM " + source,
        "WHERE program IS NOT NULL AND es2InCharge IS NOT NULL",
        "ORDER BY program, es2InCharge"
      ].join("\n");
    });
    var programOwnershipMap = {};
    var es2ProgramMap = {};
    mapBigQueryRows_(ownershipResult.rows, function(values) {
      var program = String(values[0] || "").trim();
      var es2 = String(values[1] || "").trim();
      if (program && es2) {
        programOwnershipMap[program] = es2;
        if (!es2ProgramMap[es2]) es2ProgramMap[es2] = [];
        es2ProgramMap[es2].push(program);
      }
      return null;
    });

    Logger.log("[FILTER_OPTIONS] Loaded from BigQuery: " + es2Owners.length + " ES2, " + programs.length + " Programs, " + applicationStatuses.length + " Statuses, " + Object.keys(programOwnershipMap).length + " ownership mappings");

    // Add this HEI query
    var heiResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT hei AS value",
        "FROM " + source,
        "WHERE hei IS NOT NULL AND TRIM(CAST(hei AS STRING)) != ''",
        "ORDER BY value"
      ].join("\n");
    });
    var heis = mapBigQueryRows_(heiResult.rows, function(values) {
      return String(values[0] || "").trim();
    }).filter(function(v) { return v !== ""; });

    Logger.log("[FILTER_OPTIONS] HEIs loaded: " + heis.length);

    return {
      es2Owners: es2Owners,
      programs: programs,
      applicationStatuses: applicationStatuses,
      transactionTypes: ["Walk-In", "Online"],
      heis: heis,
      dateRange: { min: "", max: "" },
      programOwnershipMap: programOwnershipMap,
      es2ProgramMap: es2ProgramMap
    };
  } catch (e) {
    Logger.log("[FILTER_OPTIONS] ERROR loading from BigQuery: " + e.toString());
    // Fallback to empty arrays
    return {
      es2Owners: [],
      programs: [],
      applicationStatuses: DASHBOARD_STATUS_ORDER || [],
      transactionTypes: ["Walk-In", "Online"],
      heis: [],
      dateRange: { min: "", max: "" },
      programOwnershipMap: {},
      es2ProgramMap: {}
    };
  }
}

function dashboard_getDashboardDetail_(kind, name, filters, page, pageSize) {
  var normalizedKind = String(kind || "").trim();
  var normalizedName = String(name || "").trim();
  var currentPage = Math.max(1, parseInt(page) || 1);
  var pageSizeValue = Math.max(1, Math.min(100, parseInt(pageSize) || 25));
  // Normalize filters using same logic as main dashboard
  var currentFilters = normalizeFilters_(filters || {});
  
  // Extract programOwnershipMap from filters for ES2/Program alignment
  var programOwnershipMap = (filters && filters.programOwnershipMap) || {};
  var filterOptions = { programOwnershipMap: programOwnershipMap };

  Logger.log("[MODAL] Fetching detail for kind=" + normalizedKind + ", name=" + normalizedName + ", page=" + currentPage);

  try {
    // Build WHERE clause based on kind (with ES2/Program alignment)
    var whereClause = buildDashboardWhereClause_(currentFilters, filterOptions);
    var canonicalStatusExpression = getDashboardCanonicalStatusExpression_();

    // Add specific filter for the modal type (match Executive Dashboard approach)
    var entityFilter = "";
    if (normalizedKind === "ES II Owner" || normalizedKind === "es2InCharge") {
      entityFilter = "es2InCharge = '" + normalizedName.replace(/'/g, "\\'") + "'";
    } else if (normalizedKind === "Application Status" || normalizedKind === "status") {
      if (normalizedName === "Cancelled") {
        entityFilter = "(" + canonicalStatusExpression + ") = 'Cancelled'";
      } else if (normalizedName && normalizedName !== "Total") {
        entityFilter = "(" + canonicalStatusExpression + ") = '" + normalizedName.replace(/'/g, "\\'") + "'";
      }
    } else if (normalizedKind === "Assigned Staff" || normalizedKind === "staff") {
      // Match Executive Dashboard: use EXISTS with UNNEST on responsiblePeopleText
      entityFilter = "EXISTS (SELECT 1 FROM UNNEST(SPLIT(COALESCE(responsiblePeopleText, ''), ',')) AS person WHERE TRIM(person) = '" + normalizedName.replace(/'/g, "\\'") + "')";
    }

    if (entityFilter) {
      whereClause = whereClause + " AND " + entityFilter;
    }

    Logger.log("[MODAL] Final whereClause: " + whereClause);

    // Get total count
    var countResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause
      ].join("\n");
    });
    var totalCount = countResult.rows && countResult.rows.length > 0 ?
      Number(countResult.rows[0].f[0].v || 0) : 0;

    var totalPages = Math.max(1, Math.ceil(totalCount / pageSizeValue));
    var offset = (currentPage - 1) * pageSizeValue;

    // Get records with pagination (match Executive Dashboard columns)
    var recordsResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT",
        "  _row,",
        "  recordId,",
        "  studentName,",
        "  program,",
        "  es2InCharge,",
        "  applicationStatus,",
        "  statusValue,",
        "  " + canonicalStatusExpression + " AS canonicalStatus,",
        "  transactionType,",
        "  hei,",
        "  CAST(dateReceivedDate AS STRING) AS dateReceived,",
        "  CAST(targetReleaseDate AS STRING) AS targetReleaseDate,",
        "  currentLocation,",
        "  currentStage,",
        "  currentStageLabel,",
        "  responsiblePeopleText",
        "FROM " + source,
        "WHERE " + whereClause,
        "ORDER BY dateReceivedDate DESC, _row DESC",
        "LIMIT " + pageSizeValue + " OFFSET " + offset
      ].join("\n");
    });

    var records = mapBigQueryRows_(recordsResult.rows, function(values) {
      var responsiblePeopleText = String(values[15] || "").trim();
      var responsiblePeople = responsiblePeopleText ? responsiblePeopleText.split(",").map(function(item) {
        return String(item || "").trim();
      }).filter(Boolean) : [];
      return {
        _row: Number(values[0] || 0),
        recordId: String(values[1] || ""),
        studentName: String(values[2] || ""),
        program: String(values[3] || ""),
        es2InCharge: String(values[4] || ""),
        applicationStatus: String(values[5] || ""),
        statusValue: String(values[6] || ""),
        canonicalStatus: String(values[7] || ""),
        transactionType: String(values[8] || ""),
        hei: String(values[9] || ""),
        dateReceived: String(values[10] || ""),
        targetReleaseDate: String(values[11] || ""),
        currentLocation: String(values[12] || ""),
        currentStage: String(values[13] || ""),
        currentStageLabel: String(values[14] || ""),
        responsiblePeopleText: responsiblePeopleText,
        responsiblePeople: responsiblePeople,
        assignedPeople: responsiblePeople
      };
    });

    // Program breakdown (match Executive Dashboard format)
    var summaryResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT",
        "  COALESCE(NULLIF(TRIM(program), ''), 'Unspecified Program') AS label,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "GROUP BY label",
        "ORDER BY total DESC, label ASC",
        "LIMIT 8"
      ].join("\n");
    });
    var programs = mapBigQueryRows_(summaryResult.rows, function(values) {
      return { label: String(values[0] || ""), value: Number(values[1] || 0) };
    }).filter(function(row) { return row.label; });

    // Transaction type breakdown
    var txResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT",
        "  COALESCE(NULLIF(TRIM(transactionType), ''), 'Unspecified') AS label,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "GROUP BY label",
        "ORDER BY total DESC, label ASC"
      ].join("\n");
    });
    var transactionTypes = mapBigQueryRows_(txResult.rows, function(values) {
      return { label: String(values[0] || ""), value: Number(values[1] || 0) };
    }).filter(function(row) { return row.label; });

    // Location breakdown
    var locResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT",
        "  COALESCE(NULLIF(TRIM(currentLocation), ''), 'Unassigned') AS label,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "GROUP BY label",
        "ORDER BY total DESC, label ASC",
        "LIMIT 8"
      ].join("\n");
    });
    var locations = mapBigQueryRows_(locResult.rows, function(values) {
      return { label: String(values[0] || ""), value: Number(values[1] || 0) };
    }).filter(function(row) { return row.label; });

    // Overall status summary
    var statusResult = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT",
        "  " + canonicalStatusExpression + " AS st,",
        "  COUNT(*) AS total",
        "FROM " + source,
        "WHERE " + whereClause,
        "GROUP BY st",
        "ORDER BY total DESC"
      ].join("\n");
    });
    var overallStatus = mapBigQueryRows_(statusResult.rows, function(values) {
      return { label: String(values[0] || ""), value: Number(values[1] || 0) };
    }).filter(function(row) { return row.label; });

    // Build summary object for Executive Dashboard compatibility
    var summary = {
      totalCount: totalCount,
      programCount: programs.length,
      locationCount: locations.length,
      statusTypeCount: overallStatus.length
    };
    
    // Determine if this is a compact status modal (for status/location modals)
    var compactStatusModal = normalizedKind === "Application Status" || normalizedKind === "status" || normalizedKind === "Location";
    
    // Application Status Trend for ES2 Owner modal
    var statusTrend = [];
    if (normalizedKind === "ES II Owner" || normalizedKind === "es2InCharge") {
      var statusTrendResult = runDashboardBigQueryOverSources_(function(source) {
        return [
          "SELECT",
          "  FORMAT_DATE('%Y-%m', dateReceivedDate) AS month,",
          "  applicationStatus AS status,",
          "  COUNT(*) AS total",
          "FROM " + source,
          "WHERE " + whereClause,
          "GROUP BY month, status",
          "ORDER BY month DESC, total DESC"
        ].join("\n");
      });
      statusTrend = mapBigQueryRows_(statusTrendResult.rows, function(values) {
        return { 
          month: String(values[0] || ""),
          label: String(values[1] || ""),
          value: Number(values[2] || 0)
        };
      });
    }

    var response = {
      ok: true,
      detail: {
        kind: normalizedKind,
        name: normalizedName,
        page: currentPage,
        pageSize: pageSizeValue,
        totalCount: totalCount,
        displayedRecordCount: records.length,
        pageCount: totalPages,
        compactStatusModal: compactStatusModal,
        summary: summary,
        rows: records,
        // Executive Dashboard compatible naming
        programRows: programs,
        transactionRows: transactionTypes,
        locationRows: locations,
        overallStatusRows: overallStatus,
        statusTrendRows: statusTrend,
        // Legacy naming for backward compatibility
        programs: programs,
        transactionTypes: transactionTypes,
        locations: locations,
        overallStatus: overallStatus,
        statusTrend: statusTrend
      }
    };

    // Serialize to ensure clean object
    try {
      return JSON.parse(JSON.stringify(response));
    } catch (e) {
      return response;
    }

  } catch (error) {
    Logger.log("[MODAL] Error fetching detail: " + error.toString());
    return {
      ok: false,
      error: "Failed to load detail: " + (error.message || error.toString()),
      detail: {
        kind: normalizedKind,
        name: normalizedName,
        page: currentPage,
        pageSize: pageSizeValue,
        totalCount: 0,
        pageCount: 1,
        rows: [],
        programs: [],
        transactionTypes: [],
        locations: [],
        overallStatus: []
      }
    };
  }
}

// ========== Executive Dashboard Migration: Reference Data & Advanced Functions ==========

function readDashboardReferenceData_(diagnostics) {
  var cachedReferenceData = getCachedDashboardReferenceData_();
  if (cachedReferenceData) {
    startTiming_(diagnostics, "referenceRead");
    endTiming_(diagnostics, "referenceRead", {
      cached: true,
      staticRows: 0,
      usersRows: 0,
      dataTypesRows: 0
    });
    diagnostics.warnings = (cachedReferenceData.warnings || []).slice();
    return cachedReferenceData;
  }

  startTiming_(diagnostics, "spreadsheetOpen");
  var spreadsheet = SpreadsheetApp.openById(DASHBOARD_SPREADSHEET_ID);
  var staticSheet = spreadsheet.getSheetByName(DASHBOARD_STATIC_SHEET);
  var usersSheet = spreadsheet.getSheetByName(DASHBOARD_USERS_SHEET);
  var dataTypesSheet = spreadsheet.getSheetByName(DASHBOARD_DATA_TYPES_SHEET);
  endTiming_(diagnostics, "spreadsheetOpen", {
    hasStaticSheet: !!staticSheet,
    hasUsersSheet: !!usersSheet,
    hasDataTypesSheet: !!dataTypesSheet
  });

  startTiming_(diagnostics, "referenceRead");
  var staticData = readWholeSheet_(staticSheet);
  var usersData = readWholeSheet_(usersSheet);
  var dataTypesData = readWholeSheet_(dataTypesSheet);
  var staticConfig = buildStaticConfig_(staticData);
  var userDirectory = buildUserDirectory_(usersData);
  var dataTypesConfig = buildDataTypesConfig_(dataTypesData);
  var warnings = [];
  if (!(staticConfig.applicationStatuses || []).length) warnings.push("Static application status list is empty or unavailable.");
  if (!(staticConfig.es2Owners || []).length) warnings.push("Static ES II ownership list is empty or unavailable.");
  if (!(userDirectory.names || []).length) warnings.push("Users sheet is empty or unavailable.");
  endTiming_(diagnostics, "referenceRead", {
    staticRows: Math.max(0, staticData.length - 1),
    usersRows: Math.max(0, usersData.length - 1),
    dataTypesRows: Math.max(0, dataTypesData.length - 1)
  });

  diagnostics.warnings = warnings.slice();
  var referenceData = {
    staticConfig: staticConfig,
    userDirectory: userDirectory,
    dataTypesConfig: dataTypesConfig,
    warnings: warnings
  };
  cacheDashboardReferenceData_(referenceData);
  return referenceData;
}

function getDashboardReferenceCacheKey_() {
  return "dashboard:reference-data:v1";
}

function getCachedDashboardReferenceData_() {
  var cache = CacheService.getScriptCache();
  var cachedValue = cache.get(getDashboardReferenceCacheKey_());
  if (!cachedValue) return null;
  try {
    return JSON.parse(cachedValue);
  } catch (error) {
    Logger.log("[REFERENCE CACHE] Parse failed: " + (error && error.message ? error.message : error));
    return null;
  }
}

function cacheDashboardReferenceData_(referenceData) {
  var cache = CacheService.getScriptCache();
  cache.put(getDashboardReferenceCacheKey_(), JSON.stringify(referenceData), DASHBOARD_REFERENCE_CACHE_TTL_SECONDS);
}

function buildStaticConfig_(staticData) {
  var config = {
    applicationStatuses: [],
    es2Owners: [],
    es2Programs: {},
    programOwners: {}
  };
  if (!staticData || staticData.length < 2) return config;

  var headers = staticData[0] || [];
  var statusIndex = headers.indexOf("Application Status");
  var es2Index = headers.indexOf("ES II Owner");
  var programIndex = headers.indexOf("Program");

  for (var i = 1; i < staticData.length; i++) {
    var row = staticData[i];
    if (statusIndex >= 0 && row[statusIndex]) config.applicationStatuses.push(String(row[statusIndex]).trim());
    if (es2Index >= 0 && row[es2Index]) {
      var es2Owner = String(row[es2Index]).trim();
      config.es2Owners.push(es2Owner);
      if (programIndex >= 0 && row[programIndex]) {
        var program = String(row[programIndex]).trim();
        config.es2Programs[program] = es2Owner;
        config.programOwners[program] = es2Owner;
      }
    }
  }

  // Remove duplicates
  config.applicationStatuses = config.applicationStatuses.filter(function(v, i, a) { return a.indexOf(v) === i; });
  config.es2Owners = config.es2Owners.filter(function(v, i, a) { return a.indexOf(v) === i; });

  return config;
}

function buildUserDirectory_(usersData) {
  var directory = { names: [], emailMap: {} };
  if (!usersData || usersData.length < 2) return directory;

  var headers = usersData[0] || [];
  var nameIndex = headers.indexOf("Name");
  var emailIndex = headers.indexOf("Email");

  for (var i = 1; i < usersData.length; i++) {
    var row = usersData[i];
    if (nameIndex >= 0 && row[nameIndex]) {
      var name = String(row[nameIndex]).trim();
      directory.names.push(name);
      if (emailIndex >= 0 && row[emailIndex]) {
        directory.emailMap[name] = String(row[emailIndex]).trim();
      }
    }
  }

  return directory;
}

function buildDataTypesConfig_(dataTypesData) {
  var config = { transactionTypes: ["Walk-In", "Online"], aliasMap: {} };
  if (!dataTypesData || dataTypesData.length < 2) return config;

  var headers = dataTypesData[0] || [];
  var typeIndex = headers.indexOf("Transaction Type");
  var aliasIndex = headers.indexOf("Alias");

  for (var i = 1; i < dataTypesData.length; i++) {
    var row = dataTypesData[i];
    if (typeIndex >= 0 && row[typeIndex]) {
      var typeName = String(row[typeIndex]).trim();
      if (config.transactionTypes.indexOf(typeName) === -1) {
        config.transactionTypes.push(typeName);
      }
      if (aliasIndex >= 0 && row[aliasIndex]) {
        config.aliasMap[String(row[aliasIndex]).trim().toLowerCase()] = typeName;
      }
    }
  }

  return config;
}

function readWholeSheet_(sheet) {
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  return sheet.getRange(1, 1, lastRow, lastCol).getValues();
}

function safelyGetBigQueryTotalCount_(filters, options) {
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT COUNT(1) AS total",
        "FROM " + source,
        "WHERE " + buildDashboardWhereClause_(filters, options)
      ].join("\n");
    });
    return result.rows && result.rows.length > 0 ? Number(result.rows[0].f[0].v || 0) : 0;
  } catch (error) {
    Logger.log("[SAFETY] Total count fallback: " + (error && error.message ? error.message : error));
    return 0;
  }
}

function safelyGetDashboardFilterOptions_(referenceData, filters) {
  try {
    return buildFilterOptionsFromReference_(referenceData, filters);
  } catch (error) {
    Logger.log("[SAFETY] Filter options fallback: " + (error && error.message ? error.message : error));
    return {
      programs: [],
      es2Owners: [],
      applicationStatuses: [],
      transactionTypes: ["Walk-In", "Online"],
      heis: [],
      dateRange: { min: "", max: "" },
      usersDirectory: []
    };
  }
}

function buildFilterOptionsFromReference_(referenceData, filters) {
  var staticConfig = referenceData.staticConfig || {};
  var userDirectory = referenceData.userDirectory || {};
  var dataTypesConfig = referenceData.dataTypesConfig || {};

  return {
    programs: Object.keys(staticConfig.programOwners || {}).sort(),
    es2Owners: (staticConfig.es2Owners || []).slice(),
    applicationStatuses: (staticConfig.applicationStatuses || []).slice(),
    transactionTypes: ["Walk-In", "Online"],
    heis: [],
    dateRange: { min: "", max: "" },
    usersDirectory: (userDirectory.names || []).slice()
  };
}

function buildFailureResponse_(error, diagnostics, mode) {
  finalizeDiagnostics_(diagnostics, "error", { mode: mode, message: error && error.message ? error.message : String(error) });
  Logger.log("[DASHBOARD ERROR] [" + diagnostics.requestId + "] " + (error && error.stack ? error.stack : error));

  return {
    ok: false,
    filters: {
      programs: [],
      es2Owners: [],
      applicationStatuses: [],
      transactionTypes: ["Walk-In", "Online"],
      heis: [],
      dateRange: { min: "", max: "" },
      usersDirectory: []
    },
    summary: {
      totalRecords: 0,
      total: 0,
      forEvaluation: 0,
      evaluated: 0,
      deficiency: 0,
      approved: 0,
      verified: 0,
      forSignature: 0,
      forRelease: 0,
      released: 0,
      cancelled: 0
    },
    sections: {
      es2Ownership: { title: "Workload by ES II Owner", rows: [] },
      topPrograms: { title: "Top 10 Programs", rows: [] },
      overallApplicationStatus: { title: "Overall Application Status", rows: [] },
      statusProgression: { title: "Status Progression Overview", rows: [] },
      turnaroundTime: { title: "Turnaround Time", rows: [] },
      locationLoad: { title: "Document Location Load", rows: [] },
      monthlyVolume: { title: "Monthly Application Volume", rows: [] },
      slaExposure: { title: "SLA Exposure", rows: [] },
      stageBacklog: { title: "Workflow Backlog", rows: [] },
      productivity: { title: "Assigned Staff Activity", rows: [] }
    },
    meta: {
      cachedAt: "",
      requestId: diagnostics.requestId,
      timings: diagnostics.timings,
      cacheStatus: "error",
      isPartial: false,
      nextStartRow: 0,
      recordCount: 0,
      loadedRecordCount: 0,
      totalRecordCount: 0,
      sourceTotalRecordCount: 0,
      appliedFilters: normalizeFilters_({}),
      warnings: diagnostics.warnings || []
    },
    error: {
      userMessage: "Dashboard loading failed.",
      message: error && error.message ? error.message : String(error),
      mode: mode,
      requestId: diagnostics.requestId
    }
  };
}

function getBigQueryAvailableYears_() {
  try {
    var result = runDashboardBigQueryOverSources_(function(source) {
      return [
        "SELECT DISTINCT EXTRACT(YEAR FROM dateReceivedDate) AS year",
        "FROM " + source,
        "WHERE dateReceivedDate IS NOT NULL",
        "ORDER BY year DESC"
      ].join("\n");
    });

    var years = [];
    if (result.rows) {
      for (var i = 0; i < result.rows.length; i++) {
        var year = Number(result.rows[i].f[0].v);
        if (!isNaN(year)) years.push(year);
      }
    }
    return years;
  } catch (error) {
    Logger.log("[AVAILABLE YEARS] Error: " + (error && error.message ? error.message : error));
    return [];
  }
}

function isBigQueryAvailable_() {
  try {
    return typeof BigQuery !== "undefined" && BigQuery && BigQuery.Jobs;
  } catch (e) {
    return false;
  }
}

function typing_getInitialData_(batchSize) {
  // Integration with Typing Records subsystem
  return { data: [], total: 0, viewConfig: getCurrentRecordViewConfig_() };
}

function typing_loadRemainingData_(startRow, batchSize) {
  return { data: [] };
}

function typing_searchExactName_(searchTerm) {
  return [];
}

function typing_getTempRecordPrefill_(rowNumber, dateReceived) {
  return null;
}

function typing_updateTempRecordEntries_(dataList, batchId) {
  return { results: [] };
}

function typing_getSaveBatchProgress_(batchId) {
  return null;
}

function buildTypingInsertContext_() {
  var ss = SpreadsheetApp.openById(spreadsheetID);
  var sheet = ss.getSheetByName(databaseSheetName);
  if (!sheet) {
    throw new Error("Database sheet not found.");
  }

  var headers = getHeaders_(sheet);
  var batchColumnIndex = getTypingBatchColumnIndex_(headers);
  var writableFieldKeys = getTypingWritableFieldKeys_();
  var columnIndexes = {};
  for (var i = 0; i < writableFieldKeys.length; i++) {
    var fieldKey = writableFieldKeys[i];
    var columnName = typingCreateFieldConfig[fieldKey].columnName;
    var headerIndex = headers.indexOf(columnName);
    if (headerIndex === -1) {
      Logger.log("[TYPING VALIDATION WARNING] Missing Database column for typing field: " + columnName);
      continue;
    }
    columnIndexes[fieldKey] = headerIndex;
  }

  return {
    sheet: sheet,
    headers: headers,
    batchColumnIndex: batchColumnIndex,
    writableFieldKeys: writableFieldKeys,
    columnIndexes: columnIndexes
  };
}

function normalizeTypingStudent_(student) {
  var safeStudent = student || {};
  return {
    dateReceived: normalizeTypingFieldValue_("dateReceived", safeStudent.dateReceived),
    studentName: normalizeTypingFieldValue_("studentName", safeStudent.studentName),
    hei: normalizeTypingFieldValue_("hei", safeStudent.hei),
    dateOfGraduation: normalizeTypingFieldValue_("dateOfGraduation", safeStudent.dateOfGraduation),
    program: normalizeTypingFieldValue_("program", safeStudent.program)
  };
}

function findExistingTypingRow_(context, student) {
  var sheet = context.sheet;
  if (sheet.getLastRow() < 2) return 0;
  var headers = context.headers;
  var columnIndexes = context.columnIndexes;
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  var batchColumnIndex = typeof context.batchColumnIndex === "number" ? context.batchColumnIndex : null;
  var targetBatchId = String(context.batchId || "").trim();
  var targetStudentName = normalizeTypingFieldValue_("studentName", student.studentName);

  if (batchColumnIndex != null && targetBatchId) {
    for (var batchRowIndex = 0; batchRowIndex < values.length; batchRowIndex++) {
      var existingBatchId = String(values[batchRowIndex][batchColumnIndex] || "").trim();
      if (existingBatchId !== targetBatchId) continue;
      var existingStudentName = Object.prototype.hasOwnProperty.call(columnIndexes, "studentName")
        ? normalizeTypingFieldValue_("studentName", values[batchRowIndex][columnIndexes.studentName])
        : "";
      if (existingStudentName === targetStudentName) {
        return batchRowIndex + 2;
      }
    }
  }

  for (var rowIndex = 0; rowIndex < values.length; rowIndex++) {
    var matches = true;
    for (var i = 0; i < typingCreateFieldKeys.length; i++) {
      var fieldKey = typingCreateFieldKeys[i];
      if (!Object.prototype.hasOwnProperty.call(columnIndexes, fieldKey)) continue;
      var existingValue = normalizeTypingFieldValue_(fieldKey, values[rowIndex][columnIndexes[fieldKey]]);
      var targetValue = normalizeTypingFieldValue_(fieldKey, student[fieldKey]);
      if (existingValue !== targetValue) {
        matches = false;
        break;
      }
    }
    if (matches) return rowIndex + 2;
  }
  return 0;
}

function insertTypingStudentRow_(context, student) {
  var normalizedStudent = normalizeTypingStudent_(student);
  var duplicateRow = findExistingTypingRow_(context, normalizedStudent);
  if (duplicateRow) {
    return {
      ok: true,
      duplicate: true,
      row: duplicateRow,
      student: normalizedStudent
    };
  }

  var row = [];
  for (var i = 0; i < context.headers.length; i++) {
    row.push("");
  }
  context.writableFieldKeys.forEach(function(fieldKey) {
    if (!Object.prototype.hasOwnProperty.call(context.columnIndexes, fieldKey)) return;
    row[context.columnIndexes[fieldKey]] = normalizedStudent[fieldKey];
  });
  if (typeof context.batchColumnIndex === "number") {
    row[context.batchColumnIndex] = String(context.batchId || "").trim();
  }
  var targetRow = context.sheet.getLastRow() + 1;
  context.sheet.getRange(targetRow, 1, 1, context.headers.length).setValues([row]);
  return {
    ok: true,
    duplicate: false,
    row: targetRow,
    student: normalizedStudent
  };
}

function updateTypingFailureRetryStatus_(logRow, result, retryCount) {
  var sheet = getOrCreateSaveFailureSheet_();
  var targetRow = Number(logRow);
  if (!targetRow || targetRow < 2 || targetRow > sheet.getLastRow()) return;
  sheet.getRange(targetRow, 11, 1, 4).setValues([[
    result && result.status ? String(result.status).toLowerCase() : "retried-failed",
    new Date(),
    result && result.message ? result.message : "",
    result && result.debugId ? result.debugId : ""
  ]]);
  var payload = parseTypingFailurePayload_(sheet.getRange(targetRow, 6).getValue());
  payload.retryCount = Number(retryCount) || 0;
  sheet.getRange(targetRow, 6, 1, 4).setValues([[
    JSON.stringify(payload),
    "TYPING_CREATE_FAILED",
    result && result.message ? result.message : "",
    getTypingFailureDetails_(payload.batchId || "", retryCount)
  ]]);
}

function retryTypingFailedRow_(payload) {
  var safePayload = payload || {};
  var batchId = String(safePayload.batchId || "").trim();
  var originalData = safePayload.originalData || safePayload.student || {};
  var normalizedStudent = normalizeTypingStudent_(originalData);
  var retryCount = Number(safePayload.retryCount) || 0;
  var logRow = Number(safePayload.logRow) || findTypingFailureLogRow_(batchId, safePayload.studentName || normalizedStudent.studentName);
  var sessionEmail = getEditorEmail_();
  var userRecord = getUserRowByEmail_(sessionEmail);
  var userContext = resolveUserContext_(sessionEmail);

  if (!sessionEmail || !userRecord || !isTypingPositionAllowed_(userContext.position)) {
    return {
      handled: true,
      status: "FAILED",
      data: {
        message: "UNAUTHORIZED"
      }
    };
  }

  if (!batchId || !normalizedStudent.studentName) {
    return {
      handled: true,
      status: "NOT_FOUND",
      data: {
        message: "Missing batchId or studentName."
      }
    };
  }

  var context;
  try {
    context = buildTypingInsertContext_();
    context.batchId = batchId;
  } catch (error) {
    logTypingCreateFailure_(batchId, normalizedStudent, error, retryCount + 1, logRow);
    return {
      handled: true,
      status: "FAILED",
      data: {
        batchId: batchId,
        studentName: normalizedStudent.studentName,
        message: String(error && error.message ? error.message : error)
      }
    };
  }

  try {
    var insertResult = insertTypingStudentRow_(context, normalizedStudent);
    var response = {
      status: "SUCCESS",
      batchId: batchId,
      studentName: normalizedStudent.studentName,
      inserted: insertResult.duplicate ? 0 : 1,
      duplicate: !!insertResult.duplicate,
      row: insertResult.row || ""
    };
    if (logRow) {
      updateTypingFailureRetryStatus_(logRow, {
        status: insertResult.duplicate ? "retried-success" : "retried-success",
        message: insertResult.duplicate ? "Retry skipped because matching row already exists." : "Retry inserted successfully.",
        debugId: createDebugId_()
      }, retryCount + 1);
    }
    return {
      handled: true,
      status: "SUCCESS",
      data: response
    };
  } catch (error2) {
    var failureLogRow = logTypingCreateFailure_(batchId, normalizedStudent, error2, retryCount + 1, logRow);
    updateTypingFailureRetryStatus_(failureLogRow, {
      status: "retried-failed",
      message: String(error2 && error2.message ? error2.message : error2),
      debugId: createDebugId_()
    }, retryCount + 1);
    return {
      handled: true,
      status: "FAILED",
      data: {
        batchId: batchId,
        studentName: normalizedStudent.studentName,
        logRow: failureLogRow,
        message: String(error2 && error2.message ? error2.message : error2)
      }
    };
  }
}

function typing_createBatch_(payload) {
  var safePayload = payload || {};
  logServerRequestContext_(getServerRequestContext_(arguments, safePayload, "typing"));
  var students = Array.isArray(safePayload.students) ? safePayload.students : [];
  var batchId = String(safePayload.batchId || "").trim();
  var sessionEmail = getEditorEmail_();
  var userRecord = getUserRowByEmail_(sessionEmail);
  var userContext = resolveUserContext_(sessionEmail);
  var writableFieldKeys = getTypingWritableFieldKeys_();

  if (!sessionEmail || !userRecord || !isTypingPositionAllowed_(userContext.position)) {
    return {
      handled: true,
      status: "UNAUTHORIZED"
    };
  }

  if (!batchId) {
    return {
      handled: true,
      status: "ERROR",
      data: {
        inserted: 0,
        skippedDuplicates: 0,
        failed: 0,
        message: "Missing batchId."
      }
    };
  }

  if (!students.length) {
    return buildHandledRouteResponse_({
      inserted: 0,
      skippedDuplicates: 0,
      failed: 0
    });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (isTypingBatchProcessed_(batchId)) {
      return {
        handled: true,
        status: "DUPLICATE_BATCH",
        data: {
          inserted: 0,
          skippedDuplicates: students.length,
          failed: 0
        }
      };
    }

    var context;
    try {
      context = buildTypingInsertContext_();
      context.batchId = batchId;
    } catch (contextError) {
      return {
        handled: true,
        status: "ERROR",
        data: {
          inserted: 0,
          skippedDuplicates: 0,
          failed: students.length,
          message: String(contextError && contextError.message ? contextError.message : contextError)
        }
      };
    }

    var inserted = 0;
    var failed = 0;

    Logger.log("[TYPING CREATE] batchId=" + batchId + " count=" + students.length);

    students.forEach(function(student) {
      var normalizedStudent = normalizeTypingStudent_(student);
      try {
        var insertResult = insertTypingStudentRow_(context, normalizedStudent);
        if (insertResult.duplicate) return;
        inserted += 1;
      } catch (error) {
        failed += 1;
        logTypingCreateFailure_(batchId, normalizedStudent, error, 0, 0);
      }
    });

    if (inserted > 0 || failed === 0) {
      markTypingBatchProcessed_(batchId, {
        inserted: inserted,
        failed: failed
      });
    }

    return buildHandledRouteResponse_({
      inserted: inserted,
      skippedDuplicates: 0,
      failed: failed
    });
  } finally {
    lock.releaseLock();
  }
}

function normalizeIssuanceRouteAction_(action, payload) {
  var normalized = String(action || "").trim().toLowerCase();
  if (
    normalized === "save" ||
    normalized === "update" ||
    normalized === "fetch" ||
    normalized === "retry" ||
    normalized === "search" ||
    normalized === "history" ||
    normalized === "failures" ||
    normalized === "log_failure" ||
    normalized === "row_info"
  ) {
    return normalized;
  }
  if (!normalized && payload && payload.searchTerm) return "search";
  return "";
}

function logRouteAudit_(module, action, handled) {
  Logger.log(JSON.stringify({
    event: "route_module_audit",
    module: String(module || ""),
    action: String(action || ""),
    handled: !!handled,
    timestamp: new Date().toISOString()
  }));
}

function normalizeIssuanceExecutionResult_(result, source, action) {
  var raw = result;
  var data = raw;
  var success = true;
  if (raw && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "handled")) {
    success = !!raw.handled && String(raw.status || "") === "SUCCESS";
    data = Object.prototype.hasOwnProperty.call(raw, "data") ? raw.data : raw;
  } else if (raw && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "ok")) {
    success = !!raw.ok;
    data = raw;
  } else if (raw && typeof raw === "object" && Object.prototype.hasOwnProperty.call(raw, "status")) {
    success = String(raw.status || "").toUpperCase() === "SUCCESS";
    data = raw;
  }
  return {
    success: !!success,
    data: data,
    source: source === "legacy" ? "legacy" : "route",
    action: String(action || "")
  };
}

function canAuditCompareIssuanceAction_(action) {
  return action === "fetch" || action === "search";
}

function executeLegacyIssuanceActionForAudit_(action, payload) {
  if (action === "search") {
    return legacy_searchExactName_(payload && payload.searchTerm);
  }
  if (action === "fetch") {
    return getIssuanceData(payload);
  }
  if (action === "update") {
    return legacy_updateIssuanceColumns_(payload);
  }
  if (action === "retry") {
    return legacy_retryFailedSave_(payload && payload.logRow);
  }
  return null;
}

function extractRowIdentifiers_(value) {
  if (Array.isArray(value)) {
    return value.slice(0, 20).map(function(item) {
      return item && typeof item === "object" ? String(item._row || item.row || item.id || "") : String(item || "");
    });
  }
  if (value && typeof value === "object") {
    return [String(value._row || value.row || value.id || "")];
  }
  return [String(value || "")];
}

function compareIssuanceExecutionContracts_(action, routedResult, legacyResult) {
  var routeContract = normalizeIssuanceExecutionResult_(routedResult, "route", action);
  var legacyContract = normalizeIssuanceExecutionResult_(legacyResult, "legacy", action);
  var routeData = routeContract.data;
  var legacyData = legacyContract.data;
  var routeShape = Array.isArray(routeData) ? "array" : typeof routeData;
  var legacyShape = Array.isArray(legacyData) ? "array" : typeof legacyData;
  var routeRows = extractRowIdentifiers_(routeData);
  var legacyRows = extractRowIdentifiers_(legacyData);
  var mismatch = routeContract.success !== legacyContract.success ||
    routeShape !== legacyShape ||
    JSON.stringify(routeRows) !== JSON.stringify(legacyRows);

  Logger.log(JSON.stringify({
    event: "route_module_compare",
    action: String(action || ""),
    mismatch: mismatch,
    route: {
      success: routeContract.success,
      shape: routeShape,
      rows: routeRows
    },
    legacy: {
      success: legacyContract.success,
      shape: legacyShape,
      rows: legacyRows
    },
    timestamp: new Date().toISOString()
  }));
}

function logRouteCompareSkipped_(module, action, reason) {
  Logger.log(JSON.stringify({
    event: "route_module_compare_skipped",
    module: String(module || ""),
    action: String(action || ""),
    reason: String(reason || ""),
    timestamp: new Date().toISOString()
  }));
}

function handleIssuanceRoute(action, payload) {
  var normalizedAction = normalizeIssuanceRouteAction_(action, payload);
  Logger.log("Issuance Action: " + normalizedAction);

  switch (normalizedAction) {
    case "save":
      return buildHandledRouteResponse_(saveIssuanceData(payload));

    case "update":
      return buildHandledRouteResponse_(updateIssuanceData(payload));

    case "fetch":
      return buildHandledRouteResponse_(getIssuanceData(payload));

    case "search":
      return buildHandledRouteResponse_(legacy_searchExactName_(payload && payload.searchTerm));

    case "retry":
      return buildHandledRouteResponse_(legacy_retryFailedSave_(payload && payload.logRow));

    case "history":
      return buildHandledRouteResponse_(legacy_getIssuanceHistory_(payload && payload.studentId, payload && payload.studentName));

    case "failures":
      return buildHandledRouteResponse_(legacy_getSaveFailureHistory_());

    case "log_failure":
      return buildHandledRouteResponse_(legacy_logClientSaveFailure_(payload && payload.data, payload && payload.clientMessage));

    case "row_info":
      return buildHandledRouteResponse_(legacy_getDatabaseRowInfo_());

    default:
      return buildUnhandledRouteResponse_("issuance", normalizedAction || String(action || ""), "Unknown issuance action.");
  }
}

function normalizeTypingRouteAction_(action) {
  var raw = String(action || "").trim().toLowerCase();
  var map = {
    "createbatch": "createBatch",
    "retryfailedrow": "retryFailedRow",
    "getinitialdata": "getInitialData",
    "loadremainingdata": "loadRemainingData",
    "searchexactname": "searchExactName",
    "gettemprecordprefill": "getTempRecordPrefill",
    "updatetemprecordentries": "updateTempRecordEntries",
    "getsavebatchprogress": "getSaveBatchProgress"
  };
  return map[raw] || raw;
}

function handleTypingRoute(action, payload) {
  var normalizedAction = normalizeTypingRouteAction_(action);
  Logger.log("Typing Action: " + normalizedAction);

  switch (normalizedAction) {
    case "createBatch":
      return typing_createBatch_(payload);

    case "retryFailedRow":
      return retryTypingFailedRow_(payload);

    case "getInitialData":
      return buildHandledRouteResponse_(typing_getInitialData_(payload.batchSize));

    case "loadRemainingData":
      return buildHandledRouteResponse_(typing_loadRemainingData_(payload.startRow, payload.batchSize));

    case "searchExactName":
      return buildHandledRouteResponse_(typing_searchExactName_(payload.searchTerm));

    case "getTempRecordPrefill":
      return buildHandledRouteResponse_(typing_getTempRecordPrefill_(payload.rowNumber, payload.dateReceived));

    case "updateTempRecordEntries":
      return buildHandledRouteResponse_(typing_updateTempRecordEntries_(payload.dataList, payload.batchId));

    case "getSaveBatchProgress":
      return buildHandledRouteResponse_(typing_getSaveBatchProgress_(payload.batchId));

    default:
      return buildUnhandledRouteResponse_("typing", normalizedAction || String(action || ""), "Unknown typing action.");
  }
}

function handleDashboardRoute(action, payload) {
  var rawAction = String(action || "").trim().toLowerCase();
  var map = {
    "fetch": "fetch",
    "refresh": "refresh",
    "getdashboarddata": "fetch",
    "getanalyticsdata": "getAnalyticsData",
    "refreshdashboardcache": "refreshDashboardCache",
    "getnavbarlogodata": "getNavbarLogoData",
    "getdashboarddetail": "getDashboardDetail"
  };
  var normalizedAction = map[rawAction] || rawAction;
  Logger.log("[BACKEND] Dashboard Routing Action: " + normalizedAction);

  switch (normalizedAction) {
    case "fetch":
    case "refresh":
    case "getDashboardData":
      var data = dashboard_fetchData_(payload || {});
      if (data && typeof data === 'object') {
        data['ok'] = true;
        var totalCount = data['summary'] ? data['summary']['totalCount'] : "N/A";
        Logger.log("[BACKEND] Dashboard data fetched, total records: " + totalCount);
      }
      return data;

    case "getAnalyticsData":
      Logger.log("[ROUTE] getAnalyticsData payload: " + JSON.stringify(payload));
      return dashboard_getAnalyticsData_(payload);

    case "refreshDashboardCache":
      return dashboard_refreshDashboardCache_();

    case "getNavbarLogoData":
      var logo = getNavbarLogoData();
      Logger.log("[BACKEND] Logo data generated: " + (logo && logo.dataUrl ? "YES" : "NO"));
      return logo;

    case "getDashboardDetail":
      return dashboard_getDashboardDetail_(payload.kind, payload.name, payload.filters, payload.page, payload.pageSize);

    default:
      Logger.log("[BACKEND] UNKNOWN DASHBOARD ACTION: " + normalizedAction);
      return { error: "Unknown dashboard action: " + normalizedAction };
  }
}

function routeModule(module, action, payload) {
  var normalizedModule = String(module || "").trim().toLowerCase();
  var normalizedAction = normalizedModule === "issuance"
    ? normalizeIssuanceRouteAction_(action, payload)
    : (normalizedModule === "typing"
      ? normalizeTypingRouteAction_(action)
      : String(action || "").trim().toLowerCase());
  Logger.log("Routing Module: " + normalizedModule);
  Logger.log("Action: " + normalizedAction);

  if (normalizedModule === "issuance") {
    var issuanceResult = handleIssuanceRoute(normalizedAction, payload);
    if (ROUTE_AUDIT_COMPARE_MODE) {
      if (canAuditCompareIssuanceAction_(normalizedAction)) {
        var legacyResult = executeLegacyIssuanceActionForAudit_(normalizedAction, payload || {});
        compareIssuanceExecutionContracts_(normalizedAction, issuanceResult, legacyResult);
      } else {
        logRouteCompareSkipped_(normalizedModule, normalizedAction, "Comparison skipped for mutating or unsupported issuance action.");
      }
    }
    logRouteAudit_(normalizedModule, normalizedAction, issuanceResult && issuanceResult.handled);
    return issuanceResult;
  }

  if (normalizedModule === "typing") {
    var typingResult = handleTypingRoute(normalizedAction, payload);
    logRouteAudit_(normalizedModule, normalizedAction, typingResult && typingResult.handled);
    return typingResult;
  }

  if (normalizedModule === "dashboard") {
    var dashboardResult = handleDashboardRoute(normalizedAction, payload);
    try {
      var isHandled = !!(dashboardResult && (dashboardResult['handled'] || dashboardResult['ok']));
      logRouteAudit_(normalizedModule, normalizedAction, isHandled);
    } catch (logError) {
      Logger.log("Audit log failed for dashboard: " + logError.toString());
    }
    return dashboardResult || { ok: false, error: "Dashboard engine returned null." };
  }

  var invalidResult = buildUnhandledRouteResponse_(normalizedModule, normalizedAction, "Invalid module.");
  logRouteAudit_(normalizedModule, normalizedAction, false);
  return invalidResult;
}

function getIssuanceTestConfig_() {
  var props = PropertiesService.getScriptProperties();
  return {
    batchSize: Number(props.getProperty("ISSUANCE_TEST_BATCH_SIZE") || 10),
    searchTerm: String(props.getProperty("ISSUANCE_TEST_SEARCH_TERM") || "").trim(),
    updatePayload: props.getProperty("ISSUANCE_TEST_UPDATE_PAYLOAD_JSON") || "",
    retryLogRow: String(props.getProperty("ISSUANCE_TEST_RETRY_LOG_ROW") || "").trim(),
    enableMutation: String(props.getProperty("ISSUANCE_TEST_ENABLE_MUTATION") || "").trim().toLowerCase() === "true"
  };
}

function logTestSection_(title) {
  console.log("");
  console.log("==================================================");
  console.log(title);
  console.log("==================================================");
}

function logPrettyJson_(label, value) {
  console.log(label);
  console.log(JSON.stringify(value, null, 2));
}

function parseTestUpdatePayload_(raw) {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    console.log("Failed to parse ISSUANCE_TEST_UPDATE_PAYLOAD_JSON");
    console.log(String(error && error.message ? error.message : error));
    return null;
  }
}

function isPlainObject_(value) {
  return !!value && Object.prototype.toString.call(value) === "[object Object]";
}

function cloneComparableValue_(value) {
  if (value == null) return null;
  if (Array.isArray(value)) {
    return value.map(cloneComparableValue_);
  }
  if (isPlainObject_(value)) {
    var out = {};
    Object.keys(value).forEach(function(key) {
      if ([
        "handled",
        "status",
        "source",
        "action",
        "timestamp",
        "module",
        "reason",
        "debugId"
      ].indexOf(key) !== -1) {
        return;
      }
      out[key] = cloneComparableValue_(value[key]);
    });
    return out;
  }
  return value;
}

function compareValuesRecursive_(routeValue, legacyValue, path, mismatchDetails) {
  var currentPath = path || "root";
  if (routeValue == null && legacyValue == null) return;

  if (Array.isArray(routeValue) || Array.isArray(legacyValue)) {
    if (!Array.isArray(routeValue) || !Array.isArray(legacyValue)) {
      mismatchDetails.push(currentPath + ": shape mismatch");
      return;
    }
    if (routeValue.length !== legacyValue.length) {
      mismatchDetails.push(currentPath + ": array length mismatch (" + routeValue.length + " vs " + legacyValue.length + ")");
    }
    var max = Math.max(routeValue.length, legacyValue.length);
    for (var i = 0; i < max; i++) {
      compareValuesRecursive_(routeValue[i], legacyValue[i], currentPath + "[" + i + "]", mismatchDetails);
    }
    return;
  }

  var routeIsObject = isPlainObject_(routeValue);
  var legacyIsObject = isPlainObject_(legacyValue);
  if (routeIsObject || legacyIsObject) {
    if (!routeIsObject || !legacyIsObject) {
      mismatchDetails.push(currentPath + ": shape mismatch");
      return;
    }
    var keyMap = {};
    Object.keys(routeValue).forEach(function(key) { keyMap[key] = true; });
    Object.keys(legacyValue).forEach(function(key) { keyMap[key] = true; });
    Object.keys(keyMap).sort().forEach(function(key) {
      var childPath = currentPath === "root" ? key : currentPath + "." + key;
      if (!Object.prototype.hasOwnProperty.call(routeValue, key)) {
        mismatchDetails.push(childPath + ": missing in route");
        return;
      }
      if (!Object.prototype.hasOwnProperty.call(legacyValue, key)) {
        mismatchDetails.push(childPath + ": missing in legacy");
        return;
      }
      compareValuesRecursive_(routeValue[key], legacyValue[key], childPath, mismatchDetails);
    });
    return;
  }

  if (String(routeValue) !== String(legacyValue)) {
    mismatchDetails.push(currentPath + ": value mismatch (" + JSON.stringify(routeValue) + " vs " + JSON.stringify(legacyValue) + ")");
  }
}

function compareRouteVsLegacy(routeResult, legacyResult, context) {
  var action = context && context.action ? String(context.action) : "";
  var routeNormalized = normalizeIssuanceExecutionResult_(routeResult, "route", action);
  var legacyNormalized = normalizeIssuanceExecutionResult_(legacyResult, "legacy", action);
  var mismatchDetails = [];

  if (!!routeNormalized.success !== !!legacyNormalized.success) {
    mismatchDetails.push("status mismatch: route success=" + routeNormalized.success + ", legacy success=" + legacyNormalized.success);
  }

  var routeRows = extractRowIdentifiers_(routeNormalized.data);
  var legacyRows = extractRowIdentifiers_(legacyNormalized.data);
  if (JSON.stringify(routeRows) !== JSON.stringify(legacyRows)) {
    mismatchDetails.push("row identifiers mismatch: route=" + JSON.stringify(routeRows) + ", legacy=" + JSON.stringify(legacyRows));
  }

  var comparableRoute = cloneComparableValue_(routeNormalized.data);
  var comparableLegacy = cloneComparableValue_(legacyNormalized.data);
  compareValuesRecursive_(comparableRoute, comparableLegacy, "root", mismatchDetails);

  var uniqueDetails = mismatchDetails.filter(function(item, index, arr) {
    return arr.indexOf(item) === index;
  });
  var isMatch = uniqueDetails.length === 0;
  var summaryScore = isMatch ? 1 : (uniqueDetails.length <= 3 ? 0.5 : 0);

  return {
    isMatch: isMatch,
    summaryScore: summaryScore,
    mismatchDetails: uniqueDetails,
    normalizedRoute: routeNormalized,
    normalizedLegacy: legacyNormalized
  };
}

function test_fetch_compare() {
  var config = getIssuanceTestConfig_();
  var payload = { batchSize: config.batchSize };

  logTestSection_("TEST: FETCH COMPARE");
  console.log("Payload:");
  console.log(JSON.stringify(payload, null, 2));

  var routeResult = routeModule("issuance", "fetch", payload);
  var legacyResult = legacy_getInitialData_(config.batchSize);
  var evaluation = compareRouteVsLegacy(routeResult, legacyResult, { action: "fetch" });

  logPrettyJson_("ROUTE RESULT", routeResult);
  logPrettyJson_("LEGACY RESULT", legacyResult);
  logPrettyJson_("COMPARISON RESULT", {
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails
  });
  return {
    testName: "fetch",
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails,
    routeResult: routeResult,
    legacyResult: legacyResult
  };
}

function test_search_compare() {
  var config = getIssuanceTestConfig_();
  var payload = { searchTerm: config.searchTerm };

  logTestSection_("TEST: SEARCH COMPARE");
  console.log("Payload:");
  console.log(JSON.stringify(payload, null, 2));

  if (!payload.searchTerm) {
    console.log("Search test skipped. Set ISSUANCE_TEST_SEARCH_TERM in Script Properties.");
    return {
      testName: "search",
      isMatch: false,
      summaryScore: 0,
      mismatchDetails: ["Missing ISSUANCE_TEST_SEARCH_TERM"],
      routeResult: null,
      legacyResult: null
    };
  }

  var routeResult = routeModule("issuance", "search", payload);
  var legacyResult = legacy_searchExactName_(payload.searchTerm);
  var evaluation = compareRouteVsLegacy(routeResult, legacyResult, { action: "search" });

  logPrettyJson_("ROUTE RESULT", routeResult);
  logPrettyJson_("LEGACY RESULT", legacyResult);
  logPrettyJson_("COMPARISON RESULT", {
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails
  });
  return {
    testName: "search",
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails,
    routeResult: routeResult,
    legacyResult: legacyResult
  };
}

function test_update_compare() {
  var config = getIssuanceTestConfig_();
  var payload = parseTestUpdatePayload_(config.updatePayload);

  logTestSection_("TEST: UPDATE COMPARE (MUTATION TEST)");
  console.log("WARNING: This test can write to the Database sheet.");

  if (!config.enableMutation) {
    console.log("Update test skipped. Set ISSUANCE_TEST_ENABLE_MUTATION=true to allow one-off mutation tests.");
    return {
      testName: "update",
      isMatch: false,
      summaryScore: 0,
      mismatchDetails: ["Mutation tests disabled"],
      routeResult: null,
      legacyResult: null
    };
  }

  if (!payload) {
    console.log("Update test skipped. Set ISSUANCE_TEST_UPDATE_PAYLOAD_JSON in Script Properties.");
    return {
      testName: "update",
      isMatch: false,
      summaryScore: 0,
      mismatchDetails: ["Missing or invalid ISSUANCE_TEST_UPDATE_PAYLOAD_JSON"],
      routeResult: null,
      legacyResult: null
    };
  }

  console.log("Payload:");
  console.log(JSON.stringify(payload, null, 2));

  var routeResult = routeModule("issuance", "update", payload);
  var legacyResult = legacy_updateIssuanceColumns_(payload);
  var evaluation = compareRouteVsLegacy(routeResult, legacyResult, { action: "update" });

  logPrettyJson_("ROUTE RESULT", routeResult);
  logPrettyJson_("LEGACY RESULT", legacyResult);
  logPrettyJson_("COMPARISON RESULT", {
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails
  });
  return {
    testName: "update",
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails,
    routeResult: routeResult,
    legacyResult: legacyResult
  };
}

function test_retry_compare() {
  var config = getIssuanceTestConfig_();
  var payload = { logRow: config.retryLogRow };

  logTestSection_("TEST: RETRY COMPARE (MUTATION TEST)");
  console.log("WARNING: This test can write retry results and may re-attempt a failed save.");
  console.log("Payload:");
  console.log(JSON.stringify(payload, null, 2));

  if (!config.enableMutation) {
    console.log("Retry test skipped. Set ISSUANCE_TEST_ENABLE_MUTATION=true to allow one-off mutation tests.");
    return {
      testName: "retry",
      isMatch: false,
      summaryScore: 0,
      mismatchDetails: ["Mutation tests disabled"],
      routeResult: null,
      legacyResult: null
    };
  }

  if (!payload.logRow) {
    console.log("Retry test skipped. Set ISSUANCE_TEST_RETRY_LOG_ROW in Script Properties.");
    return {
      testName: "retry",
      isMatch: false,
      summaryScore: 0,
      mismatchDetails: ["Missing ISSUANCE_TEST_RETRY_LOG_ROW"],
      routeResult: null,
      legacyResult: null
    };
  }

  var routeResult = routeModule("issuance", "retry", payload);
  var legacyResult = legacy_retryFailedSave_(payload.logRow);
  var evaluation = compareRouteVsLegacy(routeResult, legacyResult, { action: "retry" });

  logPrettyJson_("ROUTE RESULT", routeResult);
  logPrettyJson_("LEGACY RESULT", legacyResult);
  logPrettyJson_("COMPARISON RESULT", {
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails
  });
  return {
    testName: "retry",
    isMatch: evaluation.isMatch,
    summaryScore: evaluation.summaryScore,
    mismatchDetails: evaluation.mismatchDetails,
    routeResult: routeResult,
    legacyResult: legacyResult
  };
}

function run_all_issuance_tests() {
  logTestSection_("RUN ALL ISSUANCE TESTS");
  var results = [
    test_fetch_compare(),
    test_search_compare(),
    test_update_compare(),
    test_retry_compare()
  ];
  var totalScore = results.reduce(function(sum, item) {
    return sum + Number(item && item.summaryScore || 0);
  }, 0);
  var averageScore = results.length ? totalScore / results.length : 0;
  var routingHealth = averageScore >= 0.9 ? "GOOD" : (averageScore >= 0.6 ? "DEGRADED" : "BROKEN");

  logTestSection_("ISSUANCE TEST SUMMARY");
  console.log("testName | isMatch | summaryScore");
  results.forEach(function(item) {
    console.log([
      String(item.testName || ""),
      String(!!item.isMatch),
      String(Number(item.summaryScore || 0))
    ].join(" | "));
  });
  console.log("SYSTEM HEALTH SCORE: " + averageScore);
  console.log("ROUTING HEALTH: " + routingHealth);
  logPrettyJson_("SUMMARY RESULT", {
    results: results,
    averageScore: averageScore,
    routingHealth: routingHealth
  });
  return {
    results: results,
    averageScore: averageScore,
    routingHealth: routingHealth
  };
}

function getUserModuleAccess() {
  var email = normalizeEmail_(getEditorEmail_());
  if (!email) {
    Logger.log("[MODULE ACCESS] email=null position=null modules=[]");
    return {
      email: null,
      position: null,
      modules: []
    };
  }

  var userRow = getUserRowByEmail_(email);
  var context = resolveUserContext_(email);
  if (!userRow || !context || !context.knownUser) {
    Logger.log("[MODULE ACCESS] email=null position=null modules=[]");
    return {
      email: null,
      position: null,
      modules: []
    };
  }

  var position = String(context.position || "").trim();
  var normalizedPosition = position.toLowerCase();
  var modules = [];

  switch (normalizedPosition) {
    case "es ii":
      modules = ["dashboard"];
      break;
    case "verifier":
      modules = ["issuance"];
      break;
    case "releasing/verifier":
      modules = ["typing"];
      break;
    case "developer":
      modules = ["issuance", "typing", "dashboard"];
      break;
    default:
      modules = [];
      break;
  }

  Logger.log("[MODULE ACCESS] email=" + email + " position=" + position + " modules=[" + modules.join(", ") + "]");
  return {
    email: email,
    position: position,
    access: String(context.accessType || "").trim(),
    modules: modules
  };
}



//
// ========== Reporting_Flat Constants and Helpers (Ported from Executive Dashboard) ==========

const FIELD_ALIASES = {
  recordId: ["ID", "Id"],
  studentName: ["Name", "Student Name", "Name of Student", "Applicant Name", "Full Name", "Name of Applicant", "Student/Applicant Name", "Name of the Student", "Student's Name"],
  firstName: ["First Name", "Firstname", "Given Name"],
  middleName: ["Middle Name", "Middlename", "Middle Initial"],
  lastName: ["Last Name", "Lastname", "Surname", "Family Name"],
  program: ["Program"],
  es2InCharge: ["ES II In-charge of the Program", "ES II In-charge", "ES II In Charge", "ES II"],
  applicationStatus: ["Application Status", "Status of the Application"],
  statusOfApplication: ["Status of the Application", "Application Status"],
  evaluationStatus: ["Evaluation Status"],
  verificationStatus: ["Verification Status"],
  issuanceStatus: ["Issuance and Encoding Status", "Issuance and Encoding of No.: Status"],
  deficiencyStatus: ["Encoding of Deficiency Status", "Encoding of Deficiency: Status"],
  signatureStatus: ["RD/CAO Signature Status", "RD/CAO Signature: Status"],
  releaseStatus: ["Release Status"],
  dateReceived: ["Date Received"],
  verificationDateReceived: ["Verification: Date Received"],
  issuanceDateReceived: ["Issuance and Encoding of No.: Date Received"],
  deficiencyDateReceived: ["Encoding of Deficiency: Date Received"],
  releaseDateReceived: ["Release: Date Received"],
  targetReleaseDate: ["Target Release Date"],
  transactionType: ["Transaction Type"],
  hei: ["HEI"],
  currentLocation: ["Current Location of the Document"],
  assignedPerson: ["Assigned Person", "Assigned Person fields"],
  verificationAssignedPerson: ["Verification Assigned Person", "Verification: Assigned Person"],
  issuanceAssignedPerson: ["Issuance Assigned Person", "Issuance and Encoding Assigned Person", "Issuance and Encoding of No.: Assigned Person"],
  signatureAssignedPerson: ["RD/CAO Signature Assigned Person", "RD/CAO Signature: Assigned Person", "Signature Assigned Person"],
  releaseAssignedPerson: ["Release Assigned Person", "Release: Assigned Person"]
};

const STAGE_DEFINITIONS = [
  { key: "evaluation", label: "Evaluation", statusField: "evaluationStatus" },
  { key: "verification", label: "Verification", statusField: "verificationStatus" },
  { key: "issuance", label: "Issuance", statusField: "issuanceStatus" },
  { key: "deficiency", label: "Encoding of Deficiency", statusField: "deficiencyStatus" },
  { key: "signature", label: "RD/CAO Signature", statusField: "signatureStatus" },
  { key: "release", label: "Release", statusField: "releaseStatus" }
];

const USER_NAME_ALIASES = ["Name", "Full Name", "User Name", "Display Name"];
const STATIC_STATUS_ALIASES = ["Status of Application", "Application Status"];

// Helper functions with underscore suffix
function toTrimmedString_(value) {
  return value == null ? "" : String(value).trim();
}

function toCanonicalKey_(value) {
  return toTrimmedString_(value).toLowerCase();
}

function buildHeaderMap_(headers) {
  const map = {};
  headers.forEach(function(header, index) {
    map[toCanonicalKey_(header)] = index;
  });
  return map;
}

function resolveHeaderIndex_(headerMap, aliases) {
  for (var i = 0; i < aliases.length; i += 1) {
    const key = toCanonicalKey_(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(headerMap, key)) return headerMap[key];
  }
  return -1;
}

function buildFieldIndexMap_(headerMap) {
  const result = {};
  Object.keys(FIELD_ALIASES).forEach(function(fieldName) {
    result[fieldName] = resolveHeaderIndex_(headerMap, FIELD_ALIASES[fieldName]);
  });
  return result;
}

function getValueAtIndex_(row, index) {
  return index >= 0 && index < row.length ? toTrimmedString_(row[index]) : "";
}

function getValueAtField_(row, index) {
  return index === -1 ? "" : getValueAtIndex_(row, index);
}

function normalizeToken_(value) {
  return toTrimmedString_(value).toLowerCase().replace(/\s+/g, " ");
}

function readWholeSheet_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) return [];
  return sheet.getRange(1, 1, lastRow, lastColumn).getValues();
}

function buildStaticConfig_(staticData) {
  const config = {
    programOwners: {},
    applicationStatuses: [],
    es2Programs: {},
    es2Owners: []
  };
  if (!staticData || staticData.length < 2) return config;
  const headers = staticData[0].map(toTrimmedString_);
  const headerMap = buildHeaderMap_(headers);
  const programIndex = resolveHeaderIndex_(headerMap, FIELD_ALIASES.program);
  const es2Index = resolveHeaderIndex_(headerMap, FIELD_ALIASES.es2InCharge);
  if (programIndex !== -1 && es2Index !== -1) {
    staticData.slice(1).forEach(function(row) {
      const program = getValueAtIndex_(row, programIndex);
      const es2 = getValueAtIndex_(row, es2Index);
      if (program && es2 && !config.programOwners[program]) config.programOwners[program] = es2;
      if (program && es2) {
        if (!config.es2Programs[es2]) config.es2Programs[es2] = [];
        if (config.es2Programs[es2].indexOf(program) === -1) config.es2Programs[es2].push(program);
      }
    });
  }
  config.es2Owners = Object.keys(config.es2Programs).sort();
  Object.keys(config.es2Programs).forEach(function(es2) {
    config.es2Programs[es2].sort();
  });
  return config;
}

function buildUserDirectory_(usersData) {
  const directory = { names: [], byToken: {} };
  if (!usersData || usersData.length < 2) return directory;
  const headers = usersData[0].map(toTrimmedString_);
  const headerMap = buildHeaderMap_(headers);
  const nameIndex = resolveHeaderIndex_(headerMap, USER_NAME_ALIASES);
  if (nameIndex === -1) return directory;
  usersData.slice(1).forEach(function(row) {
    const name = getValueAtIndex_(row, nameIndex);
    const token = normalizeToken_(name);
    if (!name || !token || directory.byToken[token]) return;
    directory.byToken[token] = name;
    directory.names.push(name);
  });
  directory.names.sort();
  return directory;
}

function getDirectoryDisplayValue_(rawValue, userDirectory) {
  if (!rawValue) return "";
  const normalized = normalizeToken_(rawValue);
  return userDirectory.byToken[normalized] || rawValue;
}

function buildDataTypesConfig_(dataTypesData) {
  const config = {
    aliasMap: {
      "walk-in": "Walk-In", "walk in": "Walk-In", "walkin": "Walk-In", "onsite": "Walk-In",
      "on-site": "Walk-In", "in person": "Walk-In", "personal appearance": "Walk-In",
      "online": "Online", "on line": "Online", "portal": "Online", "email": "Online",
      "e-mail": "Online", "electronic": "Online"
    },
    transactionTypes: ["Walk-In", "Online"]
  };
  if (!dataTypesData || dataTypesData.length < 2) return config;
  dataTypesData.forEach(function(row) {
    const normalizedRow = row.map(normalizeToken_).filter(Boolean);
    if (!normalizedRow.length) return;
    const hasWalkIn = normalizedRow.some(function(value) { return /walk/.test(value); });
    const hasOnline = normalizedRow.some(function(value) { return /online/.test(value); });
    if (!hasWalkIn && !hasOnline) return;
    const canonical = hasWalkIn ? "Walk-In" : "Online";
    row.forEach(function(value) {
      const token = normalizeToken_(value);
      if (token) config.aliasMap[token] = canonical;
    });
  });
  return config;
}

function normalizeTransactionType_(value, dataTypesConfig) {
  const token = normalizeToken_(value);
  return dataTypesConfig.aliasMap[token] || value || "Unspecified";
}

function normalizeDateValue_(value) {
  if (!value) return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const parsed = new Date(toTrimmedString_(value));
  return isNaN(parsed.getTime()) ? "" : Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function buildUserList_(values, userDirectory) {
  const list = [];
  const seen = {};
  values.forEach(function(val) {
    if (!val) return;
    const normalized = normalizeToken_(val);
    const display = userDirectory.byToken[normalized] || val;
    if (display && !seen[display]) {
      seen[display] = true;
      list.push(display);
    }
  });
  return list;
}

function normalizeRecord_(row, fieldIndexes, staticConfig, userDirectory, dataTypesConfig, rowNumber) {
  const program = getValueAtField_(row, fieldIndexes.program);
  const recordId = getValueAtField_(row, fieldIndexes.recordId);
  const studentName = getValueAtField_(row, fieldIndexes.studentName);
  const applicationStatus = getValueAtField_(row, fieldIndexes.applicationStatus);
  const statusOfApplication = getValueAtField_(row, fieldIndexes.statusOfApplication);
  const stages = {};
  const ownerRaw = staticConfig.programOwners[program] || getValueAtField_(row, fieldIndexes.es2InCharge) || "";
  const ownerDisplay = getDirectoryDisplayValue_(ownerRaw, userDirectory);
  const transactionType = normalizeTransactionType_(getValueAtField_(row, fieldIndexes.transactionType), dataTypesConfig);
  
  const stageAssignments = {
    evaluation: buildUserList_([getValueAtField_(row, fieldIndexes.assignedPerson)], userDirectory),
    verification: buildUserList_([getValueAtField_(row, fieldIndexes.verificationAssignedPerson)], userDirectory),
    issuance: buildUserList_([getValueAtField_(row, fieldIndexes.issuanceAssignedPerson)], userDirectory),
    signature: buildUserList_([getValueAtField_(row, fieldIndexes.signatureAssignedPerson)], userDirectory),
    release: buildUserList_([getValueAtField_(row, fieldIndexes.releaseAssignedPerson)], userDirectory)
  };
  
  STAGE_DEFINITIONS.forEach(function(stage) {
    stages[stage.key] = {
      status: getValueAtField_(row, fieldIndexes[stage.statusField])
    };
  });
  
  const rawStatusValue = applicationStatus || statusOfApplication;
  
  return {
    _row: rowNumber,
    recordId: recordId,
    studentName: studentName,
    program: program,
    es2InCharge: ownerDisplay,
    es2InChargeRaw: ownerRaw,
    applicationStatus: applicationStatus,
    statusOfApplication: statusOfApplication,
    statusValue: rawStatusValue,
    transactionType: transactionType,
    hei: getValueAtField_(row, fieldIndexes.hei),
    dateReceived: normalizeDateValue_(getValueAtField_(row, fieldIndexes.dateReceived)),
    verificationDateReceived: normalizeDateValue_(getValueAtField_(row, fieldIndexes.verificationDateReceived)),
    issuanceDateReceived: normalizeDateValue_(getValueAtField_(row, fieldIndexes.issuanceDateReceived)),
    deficiencyDateReceived: normalizeDateValue_(getValueAtField_(row, fieldIndexes.deficiencyDateReceived)),
    releaseDateReceived: normalizeDateValue_(getValueAtField_(row, fieldIndexes.releaseDateReceived)),
    targetReleaseDate: normalizeDateValue_(getValueAtField_(row, fieldIndexes.targetReleaseDate)),
    evaluationAssignedPeople: stageAssignments.evaluation,
    evaluationAssignedPeopleText: stageAssignments.evaluation.join(", "),
    issuanceAssignedPeople: stageAssignments.issuance,
    issuanceAssignedPeopleText: stageAssignments.issuance.join(", "),
    signatureAssignedPeople: stageAssignments.signature,
    signatureAssignedPeopleText: stageAssignments.signature.join(", "),
    releaseAssignedPeople: stageAssignments.release,
    releaseAssignedPeopleText: stageAssignments.release.join(", "),
    stages: stages,
    tokens: {
      program: normalizeToken_(program),
      es2InCharge: normalizeToken_(ownerDisplay),
      transactionType: normalizeToken_(transactionType),
      hei: normalizeToken_(getValueAtField_(row, fieldIndexes.hei))
    }
  };
}

// ========== Reporting_Flat Rebuild Functions (Ported from Executive Dashboard) ==========

function checkDatabaseHeaders() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const databaseSheet = spreadsheet.getSheetByName(databaseSheetName);
  const headers = databaseSheet.getRange(1, 1, 1, 5).getValues()[0];
  Logger.log("Database headers (first 5): " + JSON.stringify(headers));
}

function forceRebuildReportingFlat() {
  // Delete and recreate sheet for truly fresh rebuild - prevents duplicates
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const reportingFlatSheet = spreadsheet.getSheetByName("Reporting_Flat");
  if (reportingFlatSheet) {
    spreadsheet.deleteSheet(reportingFlatSheet);
    Logger.log("[FORCE REBUILD] Deleted old Reporting_Flat sheet");
  }
  // Clear any stale rebuild state
  clearReportingFlatRebuildState_();
  // Rebuild from scratch (will create new sheet with headers)
  const result = rebuildReportingFlat();
  Logger.log("[FORCE REBUILD] " + JSON.stringify(result));
  return result;
}

function runCompleteRebuild() {
  // Nuclear option: delete and recreate Reporting_Flat completely
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  let reportingFlatSheet = spreadsheet.getSheetByName("Reporting_Flat");
  
  if (reportingFlatSheet) {
    spreadsheet.deleteSheet(reportingFlatSheet);
  }
  
  // Clear state
  clearReportingFlatRebuildState_();
  
  // Run fresh rebuild (will create new sheet)
  const result = rebuildReportingFlat();
  
  Logger.log("[COMPLETE REBUILD] " + JSON.stringify(result));
  return result;
}

function smartRebuildReportingFlat() {
  // Check current state and decide best approach
  const diagnosis = diagnoseMissingRows();
  
  if (diagnosis.extraRows > 1000) {
    // Significant duplicates - do complete rebuild
    Logger.log("[SMART REBUILD] Found " + diagnosis.extraRows + " extra rows. Running complete rebuild.");
    return runCompleteRebuild();
  } else if (diagnosis.missingRows > 1000) {
    // Missing significant data - resume or restart
    const state = getReportingFlatRebuildState_();
    if (state && state.nextSourceRow && state.nextSourceRow > 2) {
      Logger.log("[SMART REBUILD] Missing " + diagnosis.missingRows + " rows. Resuming from row " + state.nextSourceRow);
      return rebuildReportingFlat();
    } else {
      Logger.log("[SMART REBUILD] Missing " + diagnosis.missingRows + " rows. Starting fresh rebuild.");
      return forceRebuildReportingFlat();
    }
  } else if (diagnosis.missingRows === 0 && diagnosis.extraRows === 0) {
    Logger.log("[SMART REBUILD] Reporting_Flat is in sync with Database.");
    return { ok: true, message: "Already in sync", diagnosis: diagnosis };
  } else {
    // Small drift - use force rebuild to avoid stale state issues
    Logger.log("[SMART REBUILD] Small drift detected. Running force rebuild.");
    return forceRebuildReportingFlat();
  }
}

function diagnoseMissingRows() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const databaseSheet = mustGetSheet_(spreadsheet, databaseSheetName);
  const reportingFlatSheet = spreadsheet.getSheetByName("Reporting_Flat");
  
  const dbRows = databaseSheet.getLastRow();
  const rfRows = reportingFlatSheet ? reportingFlatSheet.getLastRow() : 0;
  const missing = Math.max(0, dbRows - rfRows);
  const extra = Math.max(0, rfRows - dbRows);
  
  Logger.log("[DIAGNOSE] Database rows: " + dbRows + ", Reporting_Flat rows: " + rfRows + ", Missing: " + missing + ", Extra: " + extra);
  
  return {
    databaseRows: dbRows,
    reportingFlatRows: rfRows,
    headerRow: 1,
    dataRowsDatabase: dbRows - 1,
    dataRowsReportingFlat: Math.max(0, rfRows - 1),
    missingRows: missing,
    extraRows: extra,
    percentComplete: dbRows > 1 ? Math.round(((rfRows - 1) / (dbRows - 1)) * 100) : 0
  };
}

function removeDuplicatesFromReportingFlat() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const reportingFlatSheet = spreadsheet.getSheetByName("Reporting_Flat");
  if (!reportingFlatSheet || reportingFlatSheet.getLastRow() <= 1) {
    return { ok: true, removed: 0, message: "Sheet empty or not found" };
  }
  
  const lastRow = reportingFlatSheet.getLastRow();
  const lastColumn = reportingFlatSheet.getLastColumn();
  
  // Get all data
  const allData = reportingFlatSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  
  // Find recordId column index (should be column 2, index 1)
  const headers = reportingFlatSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const recordIdIndex = headers.indexOf("recordId");
  
  if (recordIdIndex === -1) {
    return { ok: false, error: "recordId column not found" };
  }
  
  // Track seen recordIds and duplicates
  const seen = {};
  const duplicates = []; // Row numbers to delete (1-indexed in sheet, starting from 2)
  
  for (var i = 0; i < allData.length; i++) {
    const recordId = String(allData[i][recordIdIndex] || "").trim();
    if (!recordId) continue;
    
    if (seen[recordId]) {
      // Duplicate found - mark for deletion (actual sheet row = i + 2)
      duplicates.push(i + 2);
    } else {
      seen[recordId] = true;
    }
  }
  
  // Delete duplicates from bottom to top (to avoid index shifting)
  duplicates.reverse().forEach(function(rowNum) {
    reportingFlatSheet.deleteRow(rowNum);
  });
  
  Logger.log("[DEDUP] Removed " + duplicates.length + " duplicate rows");
  
  return {
    ok: true,
    removed: duplicates.length,
    remaining: Object.keys(seen).length,
    message: "Removed " + duplicates.length + " duplicates, kept " + Object.keys(seen).length + " unique records"
  };
}

function rebuildReportingFlatOnEdit(e) {
  // Lightweight version for onEdit trigger - just queues the rebuild
  // to avoid timeout during edit
  const sheet = e.range.getSheet();
  if (sheet.getName() !== databaseSheetName) return;
  
  // Check if we should rebuild (avoid rebuilding on every cell edit)
  const now = new Date().getTime();
  const lastRebuild = PropertiesService.getScriptProperties().getProperty("last_onedit_rebuild");
  const minInterval = 60000; // 1 minute minimum between rebuilds
  
  if (lastRebuild && (now - parseInt(lastRebuild)) < minInterval) {
    Logger.log("[onEdit] Skipping rebuild - too soon since last rebuild");
    return;
  }
  
  PropertiesService.getScriptProperties().setProperty("last_onedit_rebuild", now.toString());
  
  // For onEdit, use smart rebuild to avoid data loss
  Logger.log("[onEdit] Running smart rebuild after Database edit");
  smartRebuildReportingFlat();
}

function findMissingRows() {
  const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
  const databaseSheet = spreadsheet.getSheetByName(databaseSheetName);
  const reportingFlatSheet = spreadsheet.getSheetByName("Reporting_Flat");
  
  // Get all recordIds from Database (column A = recordId = index 1)
  const dbData = databaseSheet.getRange(2, 1, databaseSheet.getLastRow() - 1, 1).getValues();
  const dbRecordIds = dbData.map(function(row) { return String(row[0] || "").trim(); }).filter(function(id) { return id; });
  
  // Get all recordIds from Reporting_Flat (column B = recordId = index 2)
  // Note: Column A in Reporting_Flat is _row (line number), not recordId
  const rfData = reportingFlatSheet.getRange(2, 2, reportingFlatSheet.getLastRow() - 1, 1).getValues();
  const rfRecordIds = rfData.map(function(row) { return String(row[0] || "").trim(); }).filter(function(id) { return id; });
  
  // Find missing (using plain object instead of Set for GAS compatibility)
  // Convert all to strings for consistent comparison
  const rfLookup = {};
  for (var j = 0; j < rfRecordIds.length; j++) {
    rfLookup[String(rfRecordIds[j])] = true;
  }
  const missing = dbRecordIds.filter(function(id) { return !rfLookup[String(id)]; });
  
  Logger.log("Database recordIds: " + dbRecordIds.length);
  Logger.log("Reporting_Flat recordIds: " + rfRecordIds.length);
  Logger.log("Missing recordIds: " + missing.length);
  
  if (missing.length > 0) {
    Logger.log("Missing IDs: " + missing.join(", "));
    
    // Find the row numbers in Database
    missing.forEach(function(id) {
      for (var i = 0; i < dbData.length; i++) {
        if (String(dbData[i][0] || "").trim() === id) {
          Logger.log("Missing row in Database: Row " + (i + 2) + ", recordId: " + id);
          // Show what's in that row
          var rowData = databaseSheet.getRange(i + 2, 1, 1, 5).getValues()[0];
          Logger.log("Row " + (i + 2) + " data: " + JSON.stringify(rowData));
        }
      }
    });
  }
  
  return { missing: missing, dbCount: dbRecordIds.length, rfCount: rfRecordIds.length };
}

function rebuildReportingFlat() {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetID);
    const databaseSheet = mustGetSheet_(spreadsheet, databaseSheetName);
    const staticSheet = spreadsheet.getSheetByName("Static");
    const usersSheet = spreadsheet.getSheetByName(usersSheetName);
    const dataTypesSheet = spreadsheet.getSheetByName(dataTypesSheetName);
    
    const staticData = readWholeSheet_(staticSheet);
    const usersData = readWholeSheet_(usersSheet);
    const dataTypesData = readWholeSheet_(dataTypesSheet);
    
    const lastRow = databaseSheet.getLastRow();
    const lastColumn = databaseSheet.getLastColumn();
    const totalDataRows = Math.max(0, lastRow - 1);
    const headers = lastColumn ? databaseSheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    
    const rebuildResult = rebuildReportingFlatBatched_(spreadsheet, databaseSheet, headers, totalDataRows, staticData, usersData, dataTypesData);
    
    return {
      ok: true,
      rowsWritten: Number(rebuildResult.rowsWritten || 0),
      processedRows: Number(rebuildResult.processedRows || 0),
      totalDataRows: totalDataRows,
      complete: !!rebuildResult.complete,
      sheetName: "Reporting_Flat"
    };
  } catch (error) {
    Logger.log("[REBUILD] Error: " + error.toString());
    throw error;
  }
}

function rebuildReportingFlatBatched_(spreadsheet, databaseSheet, headers, totalDataRows, staticData, usersData, dataTypesData) {
  const startedAt = new Date().getTime();
  let state = getReportingFlatRebuildState_();
  const headerMap = buildHeaderMap_(headers.map(toTrimmedString_));
  const fieldIndexes = buildFieldIndexMap_(headerMap);
  const staticConfig = buildStaticConfig_(staticData);
  const userDirectory = buildUserDirectory_(usersData);
  const dataTypesConfig = buildDataTypesConfig_(dataTypesData);
  const sourceLastColumn = databaseSheet.getLastColumn();
  
  // Always reset state on fresh start (nextSourceRow would be 2 or state is null)
  let nextSourceRow = state && state.nextSourceRow ? Number(state.nextSourceRow) : 2;
  
  // Validate state - if we've somehow gotten more rows than expected, reset
  const existingSheet = spreadsheet.getSheetByName("Reporting_Flat");
  const existingDataRows = existingSheet ? Math.max(0, existingSheet.getLastRow() - 1) : 0;
  
  // If state nextSourceRow doesn't match roughly where we should be, invalidate state
  // (allows for some variance due to filtered/empty rows)
  if (state && existingDataRows > 0) {
    const expectedNextRow = existingDataRows + 2; // +1 for header, +1 for next row
    const variance = Math.abs(nextSourceRow - expectedNextRow);
    if (variance > DASHBOARD_FLAT_REBUILD_BATCH_SIZE * 2) {
      Logger.log("[REBUILD] State mismatch detected. State says row " + nextSourceRow + 
                 ", but Reporting_Flat has " + existingDataRows + " data rows. Resetting.");
      state = null;
      nextSourceRow = 2;
    }
  }
  
  let processedRows = 0;
  let rowsWritten = 0;
  let isFreshStart = !state || nextSourceRow <= 2;
  
  // Only reset sheet on fresh start
  ensureReportingFlatSheetInitialized_(spreadsheet, isFreshStart);
  
  Logger.log("[REBUILD] Starting from row " + nextSourceRow + 
             ", totalDataRows=" + totalDataRows + 
             ", isFreshStart=" + isFreshStart);
  
  while (nextSourceRow <= totalDataRows + 1 && (new Date().getTime() - startedAt) < DASHBOARD_FLAT_REBUILD_EXECUTION_BUDGET_MS) {
    const remainingRows = (totalDataRows + 1) - nextSourceRow + 1;
    const batchRowCount = Math.min(DASHBOARD_FLAT_REBUILD_BATCH_SIZE, Math.max(0, remainingRows));
    if (!batchRowCount) break;
    
    const batchValues = databaseSheet.getRange(nextSourceRow, 1, batchRowCount, sourceLastColumn).getValues();
    const flatRows = [];
    
    for (var i = 0; i < batchValues.length; i += 1) {
      const normalized = normalizeRecord_(batchValues[i], fieldIndexes, staticConfig, userDirectory, dataTypesConfig, nextSourceRow + i);
      if (!hasMeaningfulRecord_(normalized)) continue;
      flatRows.push(mapNormalizedRecordToFlatRow_(normalized));
    }
    
    appendReportingFlatRows_(spreadsheet, flatRows);
    rowsWritten += flatRows.length;
    processedRows += batchValues.length;
    nextSourceRow += batchValues.length;
    
    Logger.log("[REBUILD] Processed up to row " + nextSourceRow + 
               ", written=" + rowsWritten + 
               ", elapsed=" + (new Date().getTime() - startedAt) + "ms");
  }
  
  const complete = nextSourceRow > totalDataRows + 1;
  
  Logger.log("[REBUILD] Complete=" + complete + 
             ", final nextSourceRow=" + nextSourceRow + 
             ", totalWritten=" + rowsWritten);
  
  if (complete) {
    clearReportingFlatRebuildState_();
    sortReportingFlatSheet_(spreadsheet);
  } else {
    setReportingFlatRebuildState_({
      nextSourceRow: nextSourceRow,
      totalExpected: totalDataRows,
      lastRun: new Date().toISOString()
    });
  }
  
  return {
    complete: complete,
    rowsWritten: rowsWritten,
    processedRows: processedRows,
    nextSourceRow: complete ? 0 : nextSourceRow
  };
}

function ensureReportingFlatSheetInitialized_(spreadsheet, resetSheet) {
  let sheet = spreadsheet.getSheetByName("Reporting_Flat");
  if (!sheet) sheet = spreadsheet.insertSheet("Reporting_Flat");
  if (resetSheet) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, DASHBOARD_REPORTING_FLAT_HEADERS.length).setValues([DASHBOARD_REPORTING_FLAT_HEADERS]);
  } else if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, DASHBOARD_REPORTING_FLAT_HEADERS.length).setValues([DASHBOARD_REPORTING_FLAT_HEADERS]);
  }
  return sheet;
}

function appendReportingFlatRows_(spreadsheet, rows) {
  if (!rows || !rows.length) return;
  const sheet = ensureReportingFlatSheetInitialized_(spreadsheet, false);
  
  // Duplicate prevention: check existing recordIds
  const lastRow = sheet.getLastRow();
  if (lastRow > 1 && rows[0].recordId) {
    const existingData = sheet.getRange(2, 2, lastRow - 1, 1).getValues(); // Column B = recordId
    const existingIds = {};
    for (var i = 0; i < existingData.length; i++) {
      existingIds[String(existingData[i][0])] = true;
    }
    // Filter out duplicates
    const uniqueRows = rows.filter(function(row) {
      return !existingIds[String(row.recordId)];
    });
    if (uniqueRows.length < rows.length) {
      Logger.log("[APPEND] Skipped " + (rows.length - uniqueRows.length) + " duplicate rows");
    }
    rows = uniqueRows;
  }
  
  if (!rows.length) return; // All duplicates, nothing to append
  
  const values = rows.map(function(row) {
    return DASHBOARD_REPORTING_FLAT_HEADERS.map(function(header) {
      return Object.prototype.hasOwnProperty.call(row, header) ? row[header] : "";
    });
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, values.length, DASHBOARD_REPORTING_FLAT_HEADERS.length).setValues(values);
}

function sortReportingFlatSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName("Reporting_Flat");
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow <= 2 || !lastColumn) return;
  const headerMap = buildHeaderMap_(DASHBOARD_REPORTING_FLAT_HEADERS);
  const dateIndex = getFlatIndex_(headerMap, "dateReceived") + 1;
  const rowIndex = getFlatIndex_(headerMap, "_row") + 1;
  if (!dateIndex || !rowIndex) return;
  sheet.getRange(2, 1, lastRow - 1, lastColumn).sort([
    { column: dateIndex, ascending: false },
    { column: rowIndex, ascending: false }
  ]);
}

function getReportingFlatRebuildState_() {
  const value = PropertiesService.getScriptProperties().getProperty(DASHBOARD_FLAT_REBUILD_STATE_KEY);
  if (!value) return null;
  try {
    return JSON.parse(value);
  } catch (error) {
    return null;
  }
}

function setReportingFlatRebuildState_(state) {
  PropertiesService.getScriptProperties().setProperty(DASHBOARD_FLAT_REBUILD_STATE_KEY, JSON.stringify(state || {}));
}

function clearReportingFlatRebuildState_() {
  PropertiesService.getScriptProperties().deleteProperty(DASHBOARD_FLAT_REBUILD_STATE_KEY);
}

function resetReportingFlatRebuildState() {
  clearReportingFlatRebuildState_();
  return {
    ok: true,
    nextSourceRow: 2
  };
}

function installReportingFlatRefreshTrigger() {
  removeReportingFlatRefreshTrigger();
  ScriptApp.newTrigger("smartRebuildReportingFlat")
    .timeBased()
    .everyMinutes(DASHBOARD_FLAT_REBUILD_INTERVAL_MINUTES)
    .create();
  return {
    ok: true,
    message: "Trigger installed for every " + DASHBOARD_FLAT_REBUILD_INTERVAL_MINUTES + " minutes (using smart rebuild)"
  };
}

function removeReportingFlatRefreshTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  var removed = 0;
  triggers.forEach(function(trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === "rebuildReportingFlat" || handler === "smartRebuildReportingFlat") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return { ok: true, message: "Removed " + removed + " trigger(s)" };
}

// ========== Helper Functions for Reporting_Flat ==========

function mapNormalizedRecordToFlatRow_(normalized) {
  var row = {};
  DASHBOARD_REPORTING_FLAT_HEADERS.forEach(function(header) {
    if (header === "_row") {
      row[header] = normalized._row || "";
    } else if (normalized.hasOwnProperty(header)) {
      row[header] = normalized[header];
    } else if (normalized.tokens && normalized.tokens.hasOwnProperty(header)) {
      row[header] = normalized.tokens[header];
    } else if (normalized.stages && header.indexOf("Status") > -1) {
      var stageKey = header.replace("Status", "").toLowerCase();
      if (normalized.stages[stageKey]) {
        row[header] = normalized.stages[stageKey].status || "";
      } else {
        row[header] = "";
      }
    } else if (header.indexOf("DateReceived") > -1 && header !== "dateReceived") {
      var stageKey = header.replace("DateReceived", "").toLowerCase();
      if (normalized.stages && normalized.stages[stageKey]) {
        row[header] = normalized.stages[stageKey].dateReceived || "";
      } else {
        row[header] = "";
      }
    } else if (header.indexOf("DateAccomplished") > -1) {
      var stageKey = header.replace("DateAccomplished", "").toLowerCase();
      if (normalized.stages && normalized.stages[stageKey]) {
        row[header] = normalized.stages[stageKey].dateAccomplished || "";
      } else {
        row[header] = "";
      }
    } else if (header.indexOf("AssignedPeopleText") > -1) {
      var stageKey = header.replace("AssignedPeopleText", "").toLowerCase();
      if (stageKey === "evaluation" && normalized.evaluationAssignedPeopleText) {
        row[header] = normalized.evaluationAssignedPeopleText;
      } else if (stageKey === "issuance" && normalized.issuanceAssignedPeopleText) {
        row[header] = normalized.issuanceAssignedPeopleText;
      } else if (stageKey === "signature" && normalized.signatureAssignedPeopleText) {
        row[header] = normalized.signatureAssignedPeopleText;
      } else if (stageKey === "release" && normalized.releaseAssignedPeopleText) {
        row[header] = normalized.releaseAssignedPeopleText;
      } else if (stageKey === "verification" && normalized.verificationAssignedPeople) {
        row[header] = normalized.verificationAssignedPeople.join(", ");
      } else {
        row[header] = "";
      }
    } else {
      row[header] = "";
    }
  });
  return row;
}

function hasMeaningfulRecord_(normalized) {
  return normalized && normalized.recordId && normalized.studentName;
}

function getFlatIndex_(headerMap, fieldName) {
  return headerMap.hasOwnProperty(toCanonicalKey_(fieldName)) ? headerMap[toCanonicalKey_(fieldName)] : -1;
}

function mustGetSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('Required sheet "' + sheetName + '" was not found.');
  return sheet;
}

function testTurnaroundShape() {
  var result = getDashboardTurnaroundAnalytics({ datePreset: "thisYear" });
  Logger.log("Keys: " + Object.keys(result).join(", "));
  Logger.log("summary keys: " + Object.keys(result.summary || {}).join(", "));
  Logger.log("highlights count: " + (result.highlights || []).length);
  Logger.log("groups count: " + (result.groups || []).length);
  if (result.groups && result.groups[0]) {
    Logger.log("First group keys: " + Object.keys(result.groups[0]).join(", "));
    Logger.log("First group metrics count: " + (result.groups[0].metrics || []).length);
  }
}

function debugTurnaroundError() {
  var query = [
    "SELECT COUNT(*) as total",
    "FROM `sodium-shard-459702-t9.executive_dataset.reporting_flat_native_view`",
    "WHERE dateReceivedDate IS NOT NULL",
    "LIMIT 1"
  ].join("\n");
  
  try {
    var request = {
      query: query,
      useLegacySql: false,
      parameterMode: "NAMED",
      queryParameters: []
    };
    Logger.log("[DEBUG] Sending query...");
    var result = BigQuery.Jobs.query(request, BIGQUERY_PROJECT_ID);
    Logger.log("[DEBUG] jobComplete: " + result.jobComplete);
    Logger.log("[DEBUG] rows type: " + typeof result.rows);
    Logger.log("[DEBUG] rows value: " + JSON.stringify(result.rows));
    Logger.log("[DEBUG] result keys: " + Object.keys(result).join(", "));
  } catch(e) {
    Logger.log("[DEBUG] Error: " + e.toString());
    Logger.log("[DEBUG] Stack: " + e.stack);
  }
}

function debugTurnaroundQueryDirect() {
  var source = BIGQUERY_NATIVE_REPORTING_VIEW;
  
  // Test 1: Just the filtered CTE
  var query1 = [
    "SELECT COUNT(*) as total",
    "FROM " + source,
    "WHERE dateReceivedDate IS NOT NULL",
    "AND dateReceivedDate >= DATE('2025-01-01')"
  ].join("\n");
  
  try {
    var r1 = BigQuery.Jobs.query({ query: query1, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    Logger.log("[TEST1] jobComplete: " + r1.jobComplete + " rows: " + JSON.stringify(r1.rows));
  } catch(e) { Logger.log("[TEST1] ERROR: " + e.toString()); }

  // Test 2: Check if holiday table exists and is queryable
  var query2 = "SELECT COUNT(*) FROM " + BIGQUERY_HOLIDAYS_TABLE;
  try {
    var r2 = BigQuery.Jobs.query({ query: query2, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    Logger.log("[TEST2] Holiday table rows: " + JSON.stringify(r2.rows));
  } catch(e) { Logger.log("[TEST2] Holiday table ERROR: " + e.toString()); }

  // Test 3: Check if GENERATE_DATE_ARRAY works on the view
  var query3 = [
    "WITH sample AS (",
    "  SELECT dateReceivedDate",
    "  FROM " + source,
    "  WHERE dateReceivedDate IS NOT NULL",
    "  LIMIT 5",
    ")",
    "SELECT",
    "  (SELECT COUNT(*)",
    "   FROM UNNEST(GENERATE_DATE_ARRAY(",
    "     DATE_ADD(dateReceivedDate, INTERVAL 1 DAY),",
    "     dateReceivedDate,",
    "     INTERVAL 1 DAY",
    "   )) AS d",
    "   WHERE EXTRACT(DAYOFWEEK FROM d) BETWEEN 2 AND 6",
    "  ) AS business_days",
    "FROM sample"
  ].join("\n");
  try {
    var r3 = BigQuery.Jobs.query({ query: query3, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    Logger.log("[TEST3] GENERATE_DATE_ARRAY test: jobComplete=" + r3.jobComplete + " rows=" + JSON.stringify(r3.rows));
  } catch(e) { Logger.log("[TEST3] ERROR: " + e.toString()); }

  // Test 4: Check stage date columns exist
  var query4 = [
    "SELECT",
    "  COUNTIF(verificationDateReceived IS NOT NULL) as verif,",
    "  COUNTIF(issuanceDateReceived IS NOT NULL) as issuance,",
    "  COUNTIF(deficiencyDateReceived IS NOT NULL) as deficiency,",
    "  COUNTIF(releaseDateReceived IS NOT NULL) as release_count",
    "FROM " + source,
    "LIMIT 1"
  ].join("\n");
  try {
    var r4 = BigQuery.Jobs.query({ query: query4, useLegacySql: false }, BIGQUERY_PROJECT_ID);
    Logger.log("[TEST4] Stage columns: jobComplete=" + r4.jobComplete + " rows=" + JSON.stringify(r4.rows));
  } catch(e) { Logger.log("[TEST4] Stage columns ERROR: " + e.toString()); }
}