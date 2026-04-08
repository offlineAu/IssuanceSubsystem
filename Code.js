const spreadsheetID = "1p865Qx9OSNj0E2yU_z4YYz3kJbnT50E_zH9Ll1euGPs";
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

function getSessionEmailState_() {
  let activeEmail = "";
  let effectiveEmail = "";
  try {
    activeEmail = Session.getActiveUser().getEmail() || "";
  } catch (e) {
    activeEmail = "";
  }
  try {
    effectiveEmail = Session.getEffectiveUser().getEmail() || "";
  } catch (e2) {
    effectiveEmail = "";
  }
  return {
    activeEmail: activeEmail,
    effectiveEmail: effectiveEmail,
    resolvedEmail: activeEmail || effectiveEmail || "",
    source: activeEmail ? "active" : (effectiveEmail ? "effective" : "unknown")
  };
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
  const isDeveloper = normalizedPosition === "developer";
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
    isPrivileged: isPrivilegedAccess_(profile, position),
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

function getCurrentRecordViewConfig_() {
  const userContext = getCurrentUserContext_();
  const normalizedPosition = String(userContext.position || "").trim().toLowerCase();
  if (normalizedPosition === "developer") {
    return getDeveloperRecordViewConfig_();
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

function loadRemainingData(startRow = 1001, batchSize = 20000) {
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

function searchExactName(searchTerm) {
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

// ---------------- UPDATE ISSUANCE ----------------
function updateIssuanceColumns(data) {
  const result = performIssuanceSave_(data);
  if (!result.ok) {
    return logSaveFailureWithResult_(result, data);
  }
  return result;
}

function getIssuanceHistory(studentId, studentName) {
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

function getSaveFailureHistory() {
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

function logClientSaveFailure(data, clientMessage) {
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

function retryFailedSave(logRow) {
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

function getDatabaseRowInfo() {
  const props = PropertiesService.getScriptProperties();
  return {
    row: props.getProperty("LAST_UPDATED_ROW") || "",
    sheet: props.getProperty("LAST_UPDATED_SHEET") || "",
    updatedAt: props.getProperty("LAST_UPDATED_TIME") || "0"
  }
}

function getDatabaseRow(rowNumber) {
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



//
