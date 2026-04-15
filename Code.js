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

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Document Issuance Subsystem')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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

function runBigQueryQuery_(query, queryParameters) {
  if (typeof BigQuery === "undefined" || !BigQuery || !BigQuery.Jobs) {
    throw new Error("BigQuery advanced service is not enabled in this Apps Script project.");
  }
  var request = {
    query: query,
    useLegacySql: false,
    location: BIGQUERY_LOCATION
  };
  if (queryParameters && queryParameters.length) {
    request.parameterMode = "NAMED";
    request.queryParameters = queryParameters;
  }
  var result = BigQuery.Jobs.query(request, BIGQUERY_PROJECT_ID);
  while (!result.jobComplete) {
    Utilities.sleep(500);
    result = BigQuery.Jobs.getQueryResults(BIGQUERY_PROJECT_ID, result.jobReference.jobId, {
      location: BIGQUERY_LOCATION
    });
  }
  return result.rows || [];
}

function mapBigQueryRows_(rows, mapper) {
  return (rows || []).map(function(row, index) {
    var values = (row.f || []).map(function(cell) { return cell && cell.v != null ? cell.v : ""; });
    return mapper(values, index);
  });
}

function runDashboardBigQueryOverSources_(queryFactory) {
  var sources = [BIGQUERY_NATIVE_REPORTING_VIEW, BIGQUERY_DATABASE_FLAT_TABLE];
  var lastError = null;
  for (var i = 0; i < sources.length; i++) {
    try {
      return {
        source: sources[i],
        rows: runBigQueryQuery_(queryFactory(sources[i]), [])
      };
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError || new Error("Unable to query dashboard sources.");
}

function getDashboardSummary_(filters) {
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COUNT(1) AS totalCount,",
      "  SUM(CASE WHEN LOWER(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), '')) = 'released' THEN 1 ELSE 0 END) AS releasedCount,",
      "  SUM(CASE WHEN LOWER(COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), '')) != 'released' THEN 1 ELSE 0 END) AS pendingCount,",
      "  AVG(SAFE_CAST(DATE_DIFF(CURRENT_DATE(), DATE(dateReceived), DAY) AS FLOAT64)) AS avgAgeDays",
      "FROM " + source
    ].join("\n");
  });
  var rows = mapBigQueryRows_(result.rows, function(values) {
    return {
      totalCount: Number(values[0] || 0),
      releasedCount: Number(values[1] || 0),
      pendingCount: Number(values[2] || 0),
      avgAgeDays: Number(values[3] || 0)
    };
  });
  return {
    source: result.source,
    values: rows[0] || { totalCount: 0, releasedCount: 0, pendingCount: 0, avgAgeDays: 0 }
  };
}

function getDashboardStatusDistribution_(filters) {
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(canonicalStatus AS STRING), CAST(applicationStatus AS STRING), CAST(statusOfApplication AS STRING), 'Unknown') AS statusLabel,",
      "  COUNT(1) AS statusCount",
      "FROM " + source,
      "GROUP BY statusLabel",
      "ORDER BY statusCount DESC",
      "LIMIT 8"
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

function getDashboardWorkloadSummary_(filters) {
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(es2InCharge AS STRING), 'Unassigned') AS ownerLabel,",
      "  COUNT(1) AS workloadCount",
      "FROM " + source,
      "GROUP BY ownerLabel",
      "ORDER BY workloadCount DESC",
      "LIMIT 6"
    ].join("\n");
  });
  return {
    source: result.source,
    rows: mapBigQueryRows_(result.rows, function(values) {
      return {
        label: String(values[0] || "Unassigned"),
        value: Number(values[1] || 0)
      };
    })
  };
}

function getDashboardTurnaroundAnalytics_(filters) {
  var result = runDashboardBigQueryOverSources_(function(source) {
    return [
      "SELECT",
      "  COALESCE(CAST(currentStageLabel AS STRING), CAST(currentStage AS STRING), 'Unknown Stage') AS stageLabel,",
      "  AVG(SAFE_CAST(DATE_DIFF(CURRENT_DATE(), DATE(dateReceived), DAY) AS FLOAT64)) AS avgDays,",
      "  MAX(SAFE_CAST(DATE_DIFF(CURRENT_DATE(), DATE(dateReceived), DAY) AS FLOAT64)) AS maxDays,",
      "  MIN(SAFE_CAST(DATE_DIFF(CURRENT_DATE(), DATE(dateReceived), DAY) AS FLOAT64)) AS minDays,",
      "  COUNT(1) AS recordCount",
      "FROM " + source,
      "WHERE dateReceived IS NOT NULL",
      "GROUP BY stageLabel",
      "ORDER BY avgDays DESC",
      "LIMIT 6"
    ].join("\n");
  });
  var rows = mapBigQueryRows_(result.rows, function(values) {
    return {
      label: String(values[0] || "Unknown Stage"),
      avgDays: Number(values[1] || 0),
      maxDays: Number(values[2] || 0),
      minDays: Number(values[3] || 0),
      count: Number(values[4] || 0)
    };
  });
  return {
    source: result.source,
    summary: rows.length ? {
      avgDays: rows.reduce(function(sum, row) { return sum + row.avgDays; }, 0) / rows.length,
      maxDays: rows.reduce(function(best, row) { return Math.max(best, row.maxDays); }, 0),
      minDays: rows.reduce(function(best, row) { return Math.min(best, row.minDays); }, rows[0].minDays),
      holidayTable: BIGQUERY_HOLIDAYS_TABLE
    } : {
      avgDays: 0,
      maxDays: 0,
      minDays: 0,
      holidayTable: BIGQUERY_HOLIDAYS_TABLE
    },
    rows: rows
  };
}

function dashboard_fetchData_(filters) {
  var summary = getDashboardSummary_(filters || {});
  var workload = getDashboardWorkloadSummary_(filters || {});
  var status = getDashboardStatusDistribution_(filters || {});
  var turnaround = getDashboardTurnaroundAnalytics_(filters || {});

  return {
    kpis: {
      totalRecords: Number(summary.values.totalCount || 0),
      releasedCount: Number(summary.values.releasedCount || 0),
      pendingCount: Number(summary.values.pendingCount || 0),
      averageAgeDays: Math.round(Number(summary.values.avgAgeDays || 0) * 10) / 10
    },
    workloadSummary: workload.rows || [],
    statusDistribution: status.rows || [],
    turnaroundAnalytics: turnaround,
    meta: {
      source: summary.source || workload.source || status.source || turnaround.source || "",
      dataset: BIGQUERY_DATASET_ID,
      loadedAt: new Date().toISOString()
    }
  };
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

  var row = new Array(context.headers.length).fill("");
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

function handleTypingRoute(action, payload) {
  var normalizedAction = normalizeTypingRouteAction_(action);
  Logger.log("Typing Action: " + normalizedAction);

  switch (normalizedAction) {
    case "createBatch":
      return typing_createBatch_(payload);

    case "retryFailedRow":
      return retryTypingFailedRow_(payload);

    default:
      return buildUnhandledRouteResponse_("typing", normalizedAction || String(action || ""), "Unknown typing action.");
  }
}

function handleDashboardRoute(action, payload) {
  var normalizedAction = String(action || "").trim().toLowerCase();
  Logger.log("Dashboard Action: " + normalizedAction);

  switch (normalizedAction) {
    case "fetch":
    case "refresh":
      return buildHandledRouteResponse_(dashboard_fetchData_(payload || {}));

    default:
      return buildUnhandledRouteResponse_("dashboard", normalizedAction || String(action || ""), "Unknown dashboard action.");
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
    logRouteAudit_(normalizedModule, normalizedAction, dashboardResult && dashboardResult.handled);
    return dashboardResult;
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
