const BIJIRIS_DEFAULT_PASSWORD = "bijiris-admin";
const BIJIRIS_FOLDER_NAME = "Bijiris";
const BIJIRIS_UPLOADS_FOLDER_NAME = "Bijiris Uploads";
const BIJIRIS_BACKUPS_FOLDER_NAME = "Bijiris Backups";
const BIJIRIS_DATA_FILE_NAME = "bijiris-data.json";
const BIJIRIS_SESSIONS_KEY = "BIJIRIS_ADMIN_SESSIONS";
const BIJIRIS_DATA_FILE_ID_KEY = "BIJIRIS_DATA_FILE_ID";
const BIJIRIS_ROOT_FOLDER_ID_KEY = "BIJIRIS_ROOT_FOLDER_ID";
const BIJIRIS_UPLOADS_FOLDER_ID_KEY = "BIJIRIS_UPLOADS_FOLDER_ID";
const BIJIRIS_BACKUPS_FOLDER_ID_KEY = "BIJIRIS_BACKUPS_FOLDER_ID";
const BIJIRIS_SESSION_TTL_MS = 1000 * 60 * 60 * 24 * 30;
const MEASUREMENT_CATEGORIES = [
  "モニター",
  "回数券",
  "トライアル",
  "単発",
  "初回お試し",
  "乗り放題キャンペーン",
  "その他",
];
const TICKET_END_FORM_SLUG = "bijiris-ticket-end";
const TICKET_SHEET_FIELD_KEY = "ticket_sheet_number";

function doGet(e) {
  const action = String((e && e.parameter && e.parameter.action) || "").trim();
  if (action === "health") {
    return jsonOutput({ statusCode: 200, data: { ok: true, timestamp: nowIso() } });
  }
  return jsonOutput({ statusCode: 200, data: { ok: true, message: "Bijiris GAS backend" } });
}

function doPost(e) {
  try {
    const payload = parseRequestPayload(e);
    if (payload.action === "initializeData") {
      return jsonOutput(handleInitializeData(payload));
    }
    if (payload.action === "uploadLocalAsset") {
      return jsonOutput(handleUploadLocalAsset(payload));
    }
    if (payload.action !== "api") {
      return jsonOutput({ statusCode: 400, error: "不明な action です。" });
    }
    return jsonOutput(handleApiRequest(payload));
  } catch (error) {
    return jsonOutput({ statusCode: 500, error: error && error.message ? error.message : String(error) });
  }
}

function parseRequestPayload(e) {
  const contents = String((e && e.postData && e.postData.contents) || "").trim();
  if (!contents) {
    throw new Error("リクエスト本文が空です。");
  }
  return JSON.parse(contents);
}

function jsonOutput(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function nowIso() {
  return new Date().toISOString();
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}

function sha256Hex(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || ""), Utilities.Charset.UTF_8);
  return bytes
    .map(function (item) {
      const normalized = item < 0 ? item + 256 : item;
      return ("0" + normalized.toString(16)).slice(-2);
    })
    .join("");
}

function scriptProperties() {
  return PropertiesService.getScriptProperties();
}

function getOrCreateRootFolder() {
  const props = scriptProperties();
  const existingId = props.getProperty(BIJIRIS_ROOT_FOLDER_ID_KEY);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (_error) {
      props.deleteProperty(BIJIRIS_ROOT_FOLDER_ID_KEY);
    }
  }
  const folder = DriveApp.createFolder(BIJIRIS_FOLDER_NAME);
  props.setProperty(BIJIRIS_ROOT_FOLDER_ID_KEY, folder.getId());
  return folder;
}

function getOrCreateChildFolder(propertyKey, name) {
  const props = scriptProperties();
  const existingId = props.getProperty(propertyKey);
  if (existingId) {
    try {
      return DriveApp.getFolderById(existingId);
    } catch (_error) {
      props.deleteProperty(propertyKey);
    }
  }
  const root = getOrCreateRootFolder();
  const folders = root.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : root.createFolder(name);
  props.setProperty(propertyKey, folder.getId());
  return folder;
}

function getUploadsFolder() {
  return getOrCreateChildFolder(BIJIRIS_UPLOADS_FOLDER_ID_KEY, BIJIRIS_UPLOADS_FOLDER_NAME);
}

function getBackupsFolder() {
  return getOrCreateChildFolder(BIJIRIS_BACKUPS_FOLDER_ID_KEY, BIJIRIS_BACKUPS_FOLDER_NAME);
}

function blankData() {
  const passwordHash = sha256Hex(BIJIRIS_DEFAULT_PASSWORD);
  return {
    meta: {
      schemaVersion: 1,
      createdAt: nowIso(),
      updatedAt: nowIso(),
    },
    settings: {
      adminUsername: "admin",
      adminPasswordSha256: passwordHash,
      publicBaseUrl: "",
    },
    counters: {
      nextFormId: 1,
      nextFieldId: 1,
      nextResponseId: 1,
      nextResponseFileId: 1,
      nextProfileRecordId: 1,
      nextMeasurementId: 1,
    },
    forms: [],
    respondents: [],
    responses: [],
    profileRecords: [],
    measurements: [],
  };
}

function normalizeDataShape(data) {
  const normalized = data && typeof data === "object" ? data : blankData();
  normalized.meta = normalized.meta || {};
  normalized.settings = normalized.settings || {};
  normalized.counters = normalized.counters || {};
  normalized.forms = Array.isArray(normalized.forms) ? normalized.forms : [];
  normalized.respondents = Array.isArray(normalized.respondents) ? normalized.respondents : [];
  normalized.responses = Array.isArray(normalized.responses) ? normalized.responses : [];
  normalized.profileRecords = Array.isArray(normalized.profileRecords) ? normalized.profileRecords : [];
  normalized.measurements = Array.isArray(normalized.measurements) ? normalized.measurements : [];
  normalized.settings.adminUsername = normalized.settings.adminUsername || "admin";
  normalized.settings.adminPasswordSha256 =
    normalized.settings.adminPasswordSha256 || sha256Hex(BIJIRIS_DEFAULT_PASSWORD);
  normalized.settings.publicBaseUrl = String(normalized.settings.publicBaseUrl || "").trim();
  normalized.counters.nextFormId = Number(normalized.counters.nextFormId || 1);
  normalized.counters.nextFieldId = Number(normalized.counters.nextFieldId || 1);
  normalized.counters.nextResponseId = Number(normalized.counters.nextResponseId || 1);
  normalized.counters.nextResponseFileId = Number(normalized.counters.nextResponseFileId || 1);
  normalized.counters.nextProfileRecordId = Number(normalized.counters.nextProfileRecordId || 1);
  normalized.counters.nextMeasurementId = Number(normalized.counters.nextMeasurementId || 1);
  return normalized;
}

function getOrCreateDataFile() {
  const props = scriptProperties();
  const existingId = props.getProperty(BIJIRIS_DATA_FILE_ID_KEY);
  if (existingId) {
    try {
      return DriveApp.getFileById(existingId);
    } catch (_error) {
      props.deleteProperty(BIJIRIS_DATA_FILE_ID_KEY);
    }
  }
  const root = getOrCreateRootFolder();
  const files = root.getFilesByName(BIJIRIS_DATA_FILE_NAME);
  const file = files.hasNext()
    ? files.next()
    : root.createFile(BIJIRIS_DATA_FILE_NAME, JSON.stringify(blankData()), MimeType.PLAIN_TEXT);
  props.setProperty(BIJIRIS_DATA_FILE_ID_KEY, file.getId());
  return file;
}

function readDataStore() {
  const file = getOrCreateDataFile();
  const raw = file.getBlob().getDataAsString("UTF-8");
  if (!raw) {
    return blankData();
  }
  return normalizeDataShape(JSON.parse(raw));
}

function writeDataStore(data) {
  const file = getOrCreateDataFile();
  const nextData = normalizeDataShape(data);
  nextData.meta.updatedAt = nowIso();
  file.setContent(JSON.stringify(nextData));
  return nextData;
}

function withDataStore(callback, shouldSave) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = readDataStore();
    const result = callback(data);
    if (shouldSave) {
      writeDataStore(data);
    }
    return result;
  } finally {
    lock.releaseLock();
  }
}

function loadSessions() {
  const raw = String(scriptProperties().getProperty(BIJIRIS_SESSIONS_KEY) || "").trim();
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw);
  } catch (_error) {
    return {};
  }
}

function saveSessions(sessions) {
  scriptProperties().setProperty(BIJIRIS_SESSIONS_KEY, JSON.stringify(sessions || {}));
}

function pruneSessions(sessions) {
  const now = Date.now();
  const nextSessions = {};
  Object.keys(sessions || {}).forEach(function (token) {
    const item = sessions[token];
    if (!item || Number(item.expiresAt || 0) <= now) {
      return;
    }
    nextSessions[token] = item;
  });
  return nextSessions;
}

function createAdminSession() {
  const token = Utilities.getUuid().replace(/-/g, "");
  const sessions = pruneSessions(loadSessions());
  sessions[token] = {
    createdAt: nowIso(),
    expiresAt: Date.now() + BIJIRIS_SESSION_TTL_MS,
  };
  saveSessions(sessions);
  return token;
}

function revokeAdminSession(token) {
  if (!token) {
    return;
  }
  const sessions = pruneSessions(loadSessions());
  delete sessions[token];
  saveSessions(sessions);
}

function verifyAdminSession(token) {
  const sessions = pruneSessions(loadSessions());
  saveSessions(sessions);
  return !!(token && sessions[token]);
}

function normalizeRespondentName(value) {
  return String(value || "")
    .normalize("NFKC")
    .trim()
    .replace(/\s+/g, " ");
}

function respondentNameMatchKey(value) {
  return normalizeRespondentName(value)
    .replace(/\s+/g, "")
    .toLowerCase();
}

function respondentNameKey(value) {
  const key = respondentNameMatchKey(value).replace(/[^a-z0-9\u3040-\u30ff\u3400-\u9fff]/gi, "");
  return key || Utilities.getUuid().replace(/-/g, "");
}

function parsePathAndQuery(rawPath) {
  const value = String(rawPath || "").trim() || "/";
  const parts = value.split("?");
  const path = parts[0] || "/";
  const query = {};
  if (parts[1]) {
    parts[1].split("&").forEach(function (pair) {
      if (!pair) {
        return;
      }
      const segment = pair.split("=");
      const key = decodeURIComponent(segment[0] || "");
      const item = decodeURIComponent(segment.slice(1).join("=") || "");
      query[key] = item;
    });
  }
  return { path: path, query: query };
}

function normalizeOptionalTicketSheetValue(value) {
  const text = String(value || "")
    .normalize("NFKC")
    .trim();
  if (!text) {
    return "";
  }
  const match = text.match(/(\d{1,3})/);
  if (!match) {
    throw new Error("最新の回数券は数字で入力してください。");
  }
  return match[1] + "枚目";
}

function normalizeTicketBookType(value) {
  const text = normalizeRespondentName(value);
  if (!text) {
    return "";
  }
  if (text !== "6回券" && text !== "10回券") {
    throw new Error("回数券種別は 6回券 または 10回券 を選択してください。");
  }
  return text;
}

function ticketBookTypeMax(value) {
  if (value === "6回券") {
    return 5;
  }
  if (value === "10回券") {
    return 9;
  }
  return 0;
}

function normalizeTicketStampCount(value, ticketBookType) {
  const max = ticketBookTypeMax(ticketBookType);
  if (!max) {
    return 0;
  }
  const match = String(value || "")
    .normalize("NFKC")
    .match(/-?\d+/);
  const number = match ? Number(match[0]) : 0;
  if (!isFinite(number) || number <= 0) {
    return 0;
  }
  return Math.min(number, max);
}

function formatMeasurementValue(value) {
  const amount = Number(value || 0);
  if (!isFinite(amount)) {
    return "";
  }
  return Math.round(amount) === amount ? String(Math.round(amount)) : amount.toFixed(1);
}

function driveViewUrl(fileId) {
  return "https://drive.google.com/uc?export=view&id=" + encodeURIComponent(fileId);
}

function drivePreviewUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + encodeURIComponent(fileId) + "&sz=w1600";
}

function directImageUrl(url) {
  const raw = String(url || "").trim();
  if (!raw) {
    return "";
  }
  const match = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/) || raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return driveViewUrl(match[1]);
  }
  return raw;
}

function directPreviewUrl(url) {
  const raw = String(url || "").trim();
  if (!raw) {
    return "";
  }
  const match = raw.match(/[?&]id=([a-zA-Z0-9_-]+)/) || raw.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match) {
    return drivePreviewUrl(match[1]);
  }
  return raw;
}

function sanitizeImageObject(image) {
  if (!image) {
    return null;
  }
  const next = clone(image);
  next.url = directImageUrl(next.url || next.relativePath || "");
  next.previewUrl = directPreviewUrl(next.previewUrl || next.url || next.relativePath || "");
  return next;
}

function sanitizeFileRecord(file) {
  const next = clone(file);
  next.url = directImageUrl(next.url || next.relativePath || "");
  next.previewUrl = directPreviewUrl(next.previewUrl || next.url || next.relativePath || "");
  return next;
}

function formById(data, formId) {
  return data.forms.find(function (item) {
    return Number(item.id) === Number(formId);
  }) || null;
}

function formBySlug(data, slug, includeInactive) {
  return data.forms.find(function (item) {
    return item.slug === slug && (includeInactive || item.isActive);
  }) || null;
}

function respondentById(data, respondentId) {
  return data.respondents.find(function (item) {
    return item.respondentId === respondentId;
  }) || null;
}

function respondentByNameMatch(data, name) {
  const key = respondentNameMatchKey(name);
  if (!key) {
    return null;
  }
  return (
    data.respondents.find(function (item) {
      return respondentNameMatchKey(item.respondentName) === key;
    }) || null
  );
}

function ensureRespondentRegistry(data, respondentName, respondentId) {
  const normalizedName = normalizeRespondentName(respondentName);
  if (!normalizedName) {
    throw new Error("お名前は必須です。");
  }
  const existing = respondentId ? respondentById(data, respondentId) : respondentByNameMatch(data, normalizedName);
  if (existing) {
    existing.respondentName = normalizedName;
    existing.updatedAt = nowIso();
    return existing;
  }
  const item = {
    respondentId: respondentId || respondentNameKey(normalizedName),
    respondentName: normalizedName,
    createdAt: nowIso(),
    updatedAt: nowIso(),
    ticketSheetManualValue: "",
    currentTicketBookType: "",
    currentTicketStampCount: 0,
    currentTicketStampManualEnabled: false,
  };
  data.respondents.push(item);
  return item;
}

function fieldByKey(form, key) {
  return (form.fields || []).find(function (item) {
    return item.key === key;
  }) || null;
}

function ensureArray(value) {
  if (Array.isArray(value)) {
    return value;
  }
  if (value === undefined || value === null || value === "") {
    return [];
  }
  return [value];
}

function answerValueForField(fields, fieldKey) {
  const value = fields[fieldKey];
  return Array.isArray(value) ? value : value === undefined || value === null ? "" : String(value);
}

function normalizeCheckboxAnswer(formFields, field) {
  const values = ensureArray(formFields[field.key]).filter(function (item) {
    return String(item || "").trim();
  });
  const hasOther = values.indexOf("__other__") >= 0;
  const plainValues = values.filter(function (item) {
    return item !== "__other__";
  });
  const otherText = String(formFields[field.key + "__other_text"] || "").trim();
  if (hasOther && otherText) {
    plainValues.push("その他: " + otherText);
  }
  return plainValues.join(", ");
}

function buildResponseAnswers(form, fields, extraAnswers) {
  const answers = [];
  (form.fields || []).forEach(function (field) {
    if (field.type === "file") {
      return;
    }
    let value = "";
    if (field.type === "checkbox") {
      value = normalizeCheckboxAnswer(fields, field);
    } else {
      const raw = answerValueForField(fields, field.key);
      value = Array.isArray(raw) ? raw.join(", ") : String(raw || "").trim();
    }
    answers.push({
      id: 0,
      fieldKey: field.key,
      label: field.label,
      value: value,
    });
  });
  (extraAnswers || []).forEach(function (item) {
    answers.push(item);
  });
  return answers;
}

function validateResponseFields(form, fields, files) {
  if ((form.categoryOptions || []).length) {
    const category = String(fields.category || "").trim();
    if (!category) {
      throw new Error((form.categoryLabel || "分類") + "を選択してください。");
    }
  }
  (form.fields || []).forEach(function (field) {
    if (!field.required) {
      return;
    }
    if (field.type === "file") {
      const count = files.filter(function (item) {
        return item.fieldName === field.key;
      }).length;
      if (!count) {
        throw new Error(field.label + "は必須です。");
      }
      return;
    }
    if (field.type === "checkbox") {
      const value = normalizeCheckboxAnswer(fields, field);
      if (!value) {
        throw new Error(field.label + "は必須です。");
      }
      return;
    }
    const raw = answerValueForField(fields, field.key);
    if (!String(raw || "").trim()) {
      throw new Error(field.label + "は必須です。");
    }
  });
}

function uploadedFilePayload(fileInfo, fileRecordId) {
  const extension = getFileExtension(fileInfo.name);
  const storedName = "upload_" + fileRecordId + "_" + Utilities.getUuid().replace(/-/g, "") + extension;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(String(fileInfo.base64 || "")),
    String(fileInfo.type || "application/octet-stream"),
    storedName
  );
  const file = getUploadsFolder().createFile(blob);
  file.setName(storedName);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    id: fileRecordId,
    fieldKey: fileInfo.fieldName,
    label: "",
    originalName: fileInfo.name,
    storedName: storedName,
    mimeType: fileInfo.type || file.getMimeType(),
    size: Number(fileInfo.size || file.getSize() || 0),
    relativePath: driveViewUrl(file.getId()),
    url: driveViewUrl(file.getId()),
    previewUrl: drivePreviewUrl(file.getId()),
    driveFileId: file.getId(),
    annotation: {
      title: "",
      entryDate: "",
      memo: "",
      createdAt: nowIso(),
      updatedAt: nowIso(),
    },
  };
}

function getFileExtension(name) {
  const match = String(name || "").match(/(\.[A-Za-z0-9._-]+)$/);
  return match ? match[1] : "";
}

function publicFormsPayload(data) {
  return data.forms
    .filter(function (form) {
      return !!form.isActive;
    })
    .map(function (form) {
      return {
        id: form.id,
        title: form.title,
        slug: form.slug,
        description: form.description || "",
        questionCount: (form.fields || []).length,
      };
    });
}

function publicFormPayload(data, slug) {
  const form = formBySlug(data, slug, false);
  if (!form) {
    throw notFound("フォームが見つかりません。");
  }
  return clone(form);
}

function notFound(message) {
  const error = new Error(message);
  error.statusCode = 404;
  return error;
}

function badRequest(message) {
  const error = new Error(message);
  error.statusCode = 400;
  return error;
}

function getAutoTicketStampCount(data, respondentId, ticketBookType) {
  const max = ticketBookTypeMax(ticketBookType);
  if (!max) {
    return 0;
  }
  const count = data.responses
    .filter(function (response) {
      return response.respondentId === respondentId && String(response.formTitle || "").indexOf("施術後アンケート") >= 0;
    })
    .reduce(function (memo, response) {
      const targetAnswer = (response.answers || []).find(function (answer) {
        return String(answer.label || "").indexOf("何回目の施術") >= 0;
      });
      if (!targetAnswer) {
        return memo;
      }
      const match = String(targetAnswer.value || "").match(/(\d{1,3})/);
      const value = match ? Number(match[1]) : 0;
      return Math.max(memo, value);
    }, 0);
  return Math.min(count, max);
}

function respondentProfileRecords(data, respondentId) {
  const manualRecords = data.profileRecords
    .filter(function (item) {
      return item.respondentId === respondentId;
    })
    .map(function (item) {
      const record = clone(item);
      record.image = sanitizeImageObject(record.image);
      return record;
    });

  const responseRecords = [];
  data.responses.forEach(function (response) {
    if (response.respondentId !== respondentId) {
      return;
    }
    (response.files || []).forEach(function (file) {
      const annotation = file.annotation || {};
      responseRecords.push({
        id: "response:" + file.id,
        recordId: file.id,
        sourceType: "response",
        respondentId: respondentId,
        respondentName: response.respondentName,
        title: annotation.title || file.label || file.originalName || "アンケート画像",
        date: annotation.entryDate || String(response.createdAt || "").slice(0, 10),
        memo: annotation.memo || "",
        createdAt: annotation.createdAt || response.createdAt,
        updatedAt: annotation.updatedAt || response.createdAt,
        editable: true,
        deletable: false,
        sourceLabel: file.label || "アンケート画像",
        formTitle: response.formTitle || "",
        responseId: response.id,
        responseCreatedAt: response.createdAt,
        image: sanitizeFileRecord(file),
      });
    });
  });

  return manualRecords
    .concat(responseRecords)
    .sort(function (a, b) {
      return String(b.date || "").localeCompare(String(a.date || "")) || String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });
}

function attachMeasurementImageLinks(data, records) {
  const imageMaps = {};
  records.forEach(function (record) {
    const respondentId = record.respondentId;
    if (!imageMaps[respondentId]) {
      imageMaps[respondentId] = {};
      respondentProfileRecords(data, respondentId).forEach(function (imageRecord) {
        if (!imageRecord.date || !imageRecord.image || !imageRecord.image.url) {
          return;
        }
        if (!imageMaps[respondentId][imageRecord.date]) {
          imageMaps[respondentId][imageRecord.date] = [];
        }
        imageMaps[respondentId][imageRecord.date].push({
          label: imageRecord.title || imageRecord.sourceLabel || imageRecord.image.originalName || "画像",
          url: imageRecord.image.url,
          previewUrl: imageRecord.image.previewUrl || imageRecord.image.url,
          title: imageRecord.title || "",
          sourceLabel: imageRecord.sourceLabel || "",
          originalName: imageRecord.image.originalName || "",
        });
      });
    }
  });

  return records.map(function (record) {
    const next = clone(record);
    next.imageLinks = (imageMaps[next.respondentId] && imageMaps[next.respondentId][next.date]) || [];
    return next;
  });
}

function latestMeasurementSummary(data, respondentId) {
  const records = data.measurements
    .filter(function (item) {
      return item.respondentId === respondentId;
    })
    .map(function (item) {
      const next = clone(item);
      next.waistLabel = formatMeasurementValue(next.waist);
      next.hipLabel = formatMeasurementValue(next.hip);
      next.thighLabel = formatMeasurementValue(next.thigh);
      next.editable = true;
      next.deletable = true;
      return next;
    })
    .sort(function (a, b) {
      return String(b.date || "").localeCompare(String(a.date || "")) || String(b.updatedAt || "").localeCompare(String(a.updatedAt || ""));
    });
  if (!records.length) {
    return {
      measurementCount: 0,
      latestMeasurementDate: "",
      latestMeasurements: null,
    };
  }
  return {
    measurementCount: records.length,
    latestMeasurementDate: records[0].date,
    latestMeasurements: records[0],
  };
}

function respondentOverview(data, respondentId, formId) {
  const respondent = respondentById(data, respondentId);
  if (!respondent) {
    return null;
  }
  const responses = data.responses.filter(function (item) {
    return item.respondentId === respondentId && (!formId || Number(item.formId) === Number(formId));
  });
  const allResponses = data.responses.filter(function (item) {
    return item.respondentId === respondentId;
  });
  const profileRecords = respondentProfileRecords(data, respondentId);
  const profileSummary = profileRecords[0]
    ? {
        profileDate: profileRecords[0].date || "",
        profileTitle: profileRecords[0].title || "",
        profileMemo: profileRecords[0].memo || "",
        profileImage: profileRecords[0].image || null,
        profileRecordCount: profileRecords.length,
      }
    : {
        profileDate: "",
        profileTitle: "",
        profileMemo: "",
        profileImage: null,
        profileRecordCount: 0,
      };
  const measurementSummary = latestMeasurementSummary(data, respondentId);
  const ticketBookType = String(respondent.currentTicketBookType || "");
  const autoStamp = getAutoTicketStampCount(data, respondentId, ticketBookType);
  const manualEnabled = !!respondent.currentTicketStampManualEnabled;
  const currentStampCount = manualEnabled
    ? normalizeTicketStampCount(respondent.currentTicketStampCount, ticketBookType)
    : autoStamp;
  return {
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
    respondentEmail: "",
    responseCount: responses.length,
    lastResponseAt: allResponses.length ? allResponses[allResponses.length - 1].createdAt : "",
    latestTicketSheet: respondent.ticketSheetManualValue || "",
    latestTicketSheetManualValue: respondent.ticketSheetManualValue || "",
    currentTicketBookType: ticketBookType,
    currentTicketStampCount: currentStampCount,
    currentTicketStampManualValue: normalizeTicketStampCount(respondent.currentTicketStampCount, ticketBookType),
    currentTicketStampAutoValue: autoStamp,
    currentTicketStampManualEnabled: manualEnabled,
    currentTicketStampMax: ticketBookTypeMax(ticketBookType),
    currentTicketStampAt: "",
    profileDate: profileSummary.profileDate,
    profileTitle: profileSummary.profileTitle,
    profileMemo: profileSummary.profileMemo,
    profileImage: profileSummary.profileImage,
    profileRecordCount: profileSummary.profileRecordCount,
    measurementCount: measurementSummary.measurementCount,
    latestMeasurementDate: measurementSummary.latestMeasurementDate,
    latestMeasurements: measurementSummary.latestMeasurements,
  };
}

function respondentSummary(data, formId, query, limit) {
  const search = respondentNameMatchKey(query);
  return data.respondents
    .map(function (item) {
      return respondentOverview(data, item.respondentId, formId);
    })
    .filter(function (item) {
      if (!item) {
        return false;
      }
      if (search && respondentNameMatchKey(item.respondentName).indexOf(search) < 0) {
        return false;
      }
      if (formId && item.responseCount <= 0) {
        return false;
      }
      return true;
    })
    .sort(function (a, b) {
      return String(a.respondentName || "").localeCompare(String(b.respondentName || ""), "ja");
    })
    .slice(0, limit || 100);
}

function responseSummaryItem(item) {
  return {
    id: item.id,
    formId: item.formId,
    formTitle: item.formTitle || "",
    respondentId: item.respondentId,
    respondentName: item.respondentName || "",
    respondentEmail: item.respondentEmail || "",
    category: item.category || "",
    notes: item.notes || "",
    createdAt: item.createdAt,
    files: (item.files || []).map(sanitizeFileRecord),
  };
}

function listResponses(data, options) {
  const settings = options || {};
  const respondentQuery = respondentNameMatchKey(settings.respondentQuery || "");
  return data.responses
    .filter(function (item) {
      if (settings.formId && Number(item.formId) !== Number(settings.formId)) {
        return false;
      }
      if (settings.category && item.category !== settings.category) {
        return false;
      }
      if (respondentQuery && respondentNameMatchKey(item.respondentName).indexOf(respondentQuery) < 0) {
        return false;
      }
      return true;
    })
    .sort(function (a, b) {
      return String(b.createdAt || "").localeCompare(String(a.createdAt || "")) || Number(b.id) - Number(a.id);
    })
    .slice(0, settings.limit || 100)
    .map(responseSummaryItem);
}

function categorySummary(data, formId, respondentQuery) {
  const counts = {};
  listResponses(data, { formId: formId, respondentQuery: respondentQuery, limit: 10000 }).forEach(function (item) {
    const key = item.category || "未分類";
    counts[key] = (counts[key] || 0) + 1;
  });
  return Object.keys(counts)
    .sort(function (a, b) {
      return counts[b] - counts[a] || a.localeCompare(b, "ja");
    })
    .map(function (key) {
      return { category: key, count: counts[key] };
    });
}

function responseDetail(data, responseId) {
  const item = data.responses.find(function (response) {
    return Number(response.id) === Number(responseId);
  });
  if (!item) {
    return null;
  }
  return {
    response: responseSummaryItem(item),
    answers: clone(item.answers || []),
    files: (item.files || []).map(sanitizeFileRecord),
  };
}

function respondentHistory(data, respondentId, formId) {
  return data.responses
    .filter(function (response) {
      return response.respondentId === respondentId && (!formId || Number(response.formId) === Number(formId));
    })
    .sort(function (a, b) {
      return String(a.createdAt || "").localeCompare(String(b.createdAt || "")) || Number(a.id) - Number(b.id);
    })
    .map(function (item) {
      return {
        id: item.id,
        formId: item.formId,
        formTitle: item.formTitle || "",
        respondentId: item.respondentId,
        respondentName: item.respondentName || "",
        respondentEmail: item.respondentEmail || "",
        category: item.category || "",
        notes: item.notes || "",
        createdAt: item.createdAt,
        files: (item.files || []).map(sanitizeFileRecord),
        answers: clone(item.answers || []),
      };
    });
}

function measurementRecordPayload(item) {
  const next = clone(item);
  next.waistLabel = formatMeasurementValue(next.waist);
  next.hipLabel = formatMeasurementValue(next.hip);
  next.thighLabel = formatMeasurementValue(next.thigh);
  next.editable = true;
  next.deletable = true;
  return next;
}

function listMeasurementRecords(data, options) {
  const settings = options || {};
  const search = respondentNameMatchKey(settings.query || "");
  const respondentName = respondentNameMatchKey(settings.respondentName || "");
  let records = data.measurements
    .filter(function (item) {
      if (settings.respondentId && item.respondentId !== settings.respondentId) {
        return false;
      }
      if (respondentName && respondentNameMatchKey(item.respondentName).indexOf(respondentName) < 0) {
        return false;
      }
      if (search) {
        const haystack = [item.respondentName, item.category, item.date].join(" ");
        if (respondentNameMatchKey(haystack).indexOf(search) < 0) {
          return false;
        }
      }
      return true;
    })
    .sort(function (a, b) {
      return String(a.date || "").localeCompare(String(b.date || "")) || String(a.updatedAt || "").localeCompare(String(b.updatedAt || ""));
    })
    .slice(0, settings.limit || 500)
    .map(measurementRecordPayload);
  records = attachMeasurementImageLinks(data, records);
  return records;
}

function updateRespondentTicketStatus(data, respondentId, payload) {
  const respondent = respondentById(data, respondentId);
  if (!respondent) {
    throw notFound("回答者が見つかりません。");
  }
  if (Object.prototype.hasOwnProperty.call(payload, "ticketSheetManualValue")) {
    respondent.ticketSheetManualValue = normalizeOptionalTicketSheetValue(payload.ticketSheetManualValue);
  }
  if (Object.prototype.hasOwnProperty.call(payload, "currentTicketBookType")) {
    respondent.currentTicketBookType = normalizeTicketBookType(payload.currentTicketBookType);
  }
  if (Object.prototype.hasOwnProperty.call(payload, "currentTicketStampManualEnabled")) {
    respondent.currentTicketStampManualEnabled = !!payload.currentTicketStampManualEnabled;
  }
  if (Object.prototype.hasOwnProperty.call(payload, "currentTicketStampCount")) {
    respondent.currentTicketStampCount = normalizeTicketStampCount(
      payload.currentTicketStampCount,
      respondent.currentTicketBookType
    );
  }
  respondent.updatedAt = nowIso();
  return respondent;
}

function nextCounter(data, key) {
  const value = Number(data.counters[key] || 1);
  data.counters[key] = value + 1;
  return value;
}

function createPublicResponse(data, form, payload) {
  const formData = payload.formData || { fields: {}, files: [] };
  const fields = formData.fields || {};
  const files = Array.isArray(formData.files) ? formData.files : [];
  const respondentName = normalizeRespondentName(fields.respondent_name || "");
  if (!respondentName) {
    throw badRequest("お名前は必須です。");
  }
  validateResponseFields(form, fields, files);
  const respondent = respondentByNameMatch(data, respondentName) || ensureRespondentRegistry(data, respondentName);
  const responseId = nextCounter(data, "nextResponseId");
  const category = String(fields.category || "").trim();
  const extraAnswers = [];
  if (fields[TICKET_SHEET_FIELD_KEY]) {
    extraAnswers.push({
      id: 0,
      fieldKey: TICKET_SHEET_FIELD_KEY,
      label: "今回終了した回数券",
      value: normalizeOptionalTicketSheetValue(fields[TICKET_SHEET_FIELD_KEY]),
    });
  }
  const responseFiles = files.map(function (fileInfo) {
    const fileId = nextCounter(data, "nextResponseFileId");
    const record = uploadedFilePayload(fileInfo, fileId);
    const field = fieldByKey(form, fileInfo.fieldName);
    record.label = field ? field.label : fileInfo.fieldName;
    return record;
  });
  const response = {
    id: responseId,
    formId: form.id,
    formTitle: form.title,
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
    respondentEmail: "",
    category: category,
    notes: "",
    createdAt: nowIso(),
    ipAddress: "",
    userAgent: "",
    answers: buildResponseAnswers(form, fields, extraAnswers).map(function (item, index) {
      item.id = index + 1;
      return item;
    }),
    files: responseFiles,
  };
  data.responses.push(response);
  respondent.updatedAt = nowIso();
  return {
    ok: true,
    message: form.successMessage || "送信ありがとうございました。",
    responseId: responseId,
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
  };
}

function saveForm(data, payload, formId) {
  const title = normalizeRespondentName(payload.title || "");
  const slug = String(payload.slug || "")
    .trim()
    .replace(/[^\w-]/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .toLowerCase();
  if (!title || !slug) {
    throw badRequest("フォーム名と公開URLスラッグは必須です。");
  }
  const duplicate = data.forms.find(function (item) {
    return item.slug === slug && Number(item.id) !== Number(formId || 0);
  });
  if (duplicate) {
    throw badRequest("同じURLスラッグのフォームが既に存在します。");
  }
  const fields = (payload.fields || []).map(function (field, index) {
    return {
      id: Number(field.id || nextCounter(data, "nextFieldId")),
      label: String(field.label || "").trim(),
      key: String(field.key || "").trim(),
      type: String(field.type || "short_text").trim(),
      required: !!field.required,
      options: Array.isArray(field.options) ? field.options : [],
      placeholder: String(field.placeholder || "").trim(),
      helpText: String(field.helpText || "").trim(),
      visibilityFieldKey: String(field.visibilityFieldKey || "").trim(),
      visibilityValues: Array.isArray(field.visibilityValues) ? field.visibilityValues : [],
      accept: String(field.accept || "").trim(),
      allowMultiple: !!field.allowMultiple,
      allowOther: !!field.allowOther,
      sortOrder: index,
    };
  });
  let form = formId ? formById(data, formId) : null;
  if (!form) {
    form = {
      id: nextCounter(data, "nextFormId"),
      createdAt: nowIso(),
    };
    data.forms.push(form);
  }
  form.title = title;
  form.slug = slug;
  form.description = String(payload.description || "").trim();
  form.successMessage = String(payload.successMessage || "").trim();
  form.categoryLabel = String(payload.categoryLabel || "分類").trim() || "分類";
  form.categoryOptions = Array.isArray(payload.categoryOptions) ? payload.categoryOptions : [];
  form.isActive = payload.isActive !== false;
  form.updatedAt = nowIso();
  form.fields = fields;
  return clone(form);
}

function toggleForm(data, formId) {
  const form = formById(data, formId);
  if (!form) {
    throw notFound("フォームが見つかりません。");
  }
  form.isActive = !form.isActive;
  form.updatedAt = nowIso();
  return clone(form);
}

function renameRespondentEverywhere(data, respondentId, newName) {
  const normalizedName = normalizeRespondentName(newName);
  if (!normalizedName) {
    throw badRequest("お名前は必須です。");
  }
  const respondent = respondentById(data, respondentId);
  if (!respondent) {
    throw notFound("回答者が見つかりません。");
  }
  const nextRespondentId = respondentNameKey(normalizedName);
  respondent.respondentId = nextRespondentId;
  respondent.respondentName = normalizedName;
  respondent.updatedAt = nowIso();
  data.responses.forEach(function (item) {
    if (item.respondentId === respondentId) {
      item.respondentId = nextRespondentId;
      item.respondentName = normalizedName;
    }
  });
  data.profileRecords.forEach(function (item) {
    if (item.respondentId === respondentId) {
      item.respondentId = nextRespondentId;
      item.respondentName = normalizedName;
      item.updatedAt = nowIso();
    }
  });
  data.measurements.forEach(function (item) {
    if (item.respondentId === respondentId) {
      item.respondentId = nextRespondentId;
      item.respondentName = normalizedName;
      item.updatedAt = nowIso();
    }
  });
  return respondentById(data, nextRespondentId) || respondent;
}

function updateRespondentProfile(data, respondentId, payload) {
  let respondent = respondentById(data, respondentId);
  if (!respondent) {
    throw notFound("回答者が見つかりません。");
  }
  const nextName = normalizeRespondentName(payload.name || respondent.respondentName);
  if (nextName !== respondent.respondentName) {
    respondent = renameRespondentEverywhere(data, respondentId, nextName);
    respondentId = respondent.respondentId;
  }
  updateRespondentTicketStatus(data, respondentId, {
    ticketSheetManualValue: Object.prototype.hasOwnProperty.call(payload, "ticketSheet")
      ? payload.ticketSheet
      : respondent.ticketSheetManualValue,
    currentTicketBookType: Object.prototype.hasOwnProperty.call(payload, "ticketBookType")
      ? payload.ticketBookType
      : respondent.currentTicketBookType,
    currentTicketStampCount: Object.prototype.hasOwnProperty.call(payload, "ticketStampCount")
      ? payload.ticketStampCount
      : respondent.currentTicketStampCount,
    currentTicketStampManualEnabled: Object.prototype.hasOwnProperty.call(payload, "ticketStampManualEnabled")
      ? String(payload.ticketStampManualEnabled || "").trim().toLowerCase() === "1" ||
        String(payload.ticketStampManualEnabled || "").trim().toLowerCase() === "true"
      : respondent.currentTicketStampManualEnabled,
  });
  respondent = respondentById(data, respondentId);
  return {
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
    ticketSheet: respondent.ticketSheetManualValue || "",
    ticketBookType: respondent.currentTicketBookType || "",
    ticketStampCount: respondent.currentTicketStampCount || 0,
    ticketStampManualEnabled: !!respondent.currentTicketStampManualEnabled,
  };
}

function createProfileRecord(data, respondentId, payload) {
  const respondent = respondentById(data, respondentId);
  if (!respondent) {
    throw notFound("回答者が見つかりません。");
  }
  const formData = payload.formData || { fields: {}, files: [] };
  const title = String(formData.fields.title || "").trim();
  const entryDate = String(formData.fields.entry_date || "").trim();
  if (!title || !entryDate) {
    throw badRequest("タイトルと日付は必須です。");
  }
  const fileInfo = (formData.files || [])[0];
  if (!fileInfo) {
    throw badRequest("画像は必須です。");
  }
  const fileRecord = uploadedFilePayload(fileInfo, nextCounter(data, "nextResponseFileId"));
  const record = {
    id: nextCounter(data, "nextProfileRecordId"),
    recordId: 0,
    sourceType: "manual",
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
    title: title,
    date: entryDate,
    memo: String(formData.fields.memo || "").trim(),
    createdAt: nowIso(),
    updatedAt: nowIso(),
    editable: true,
    deletable: true,
    sourceLabel: "管理者追加",
    image: sanitizeImageObject(fileRecord),
  };
  record.recordId = record.id;
  data.profileRecords.push(record);
  return clone(record);
}

function findManualProfileRecord(data, respondentId, recordId) {
  return data.profileRecords.find(function (item) {
    return item.respondentId === respondentId && Number(item.id) === Number(recordId);
  }) || null;
}

function findResponseFileRecord(data, respondentId, recordId) {
  let match = null;
  data.responses.some(function (response) {
    if (response.respondentId !== respondentId) {
      return false;
    }
    return (response.files || []).some(function (file) {
      if (Number(file.id) !== Number(recordId)) {
        return false;
      }
      match = { response: response, file: file };
      return true;
    });
  });
  return match;
}

function updateProfileRecord(data, respondentId, sourceType, recordId, payload) {
  const title = String(payload.title || "").trim();
  const entryDate = String(payload.entryDate || "").trim();
  const memo = String(payload.memo || "").trim();
  if (!entryDate) {
    throw badRequest("日付は必須です。");
  }
  if (sourceType === "manual") {
    const record = findManualProfileRecord(data, respondentId, recordId);
    if (!record) {
      throw notFound("画像記録が見つかりません。");
    }
    record.title = title || record.title || "無題";
    record.date = entryDate;
    record.memo = memo;
    record.updatedAt = nowIso();
    return clone(record);
  }
  const responseFile = findResponseFileRecord(data, respondentId, recordId);
  if (!responseFile) {
    throw notFound("画像記録が見つかりません。");
  }
  responseFile.file.annotation = responseFile.file.annotation || {
    title: "",
    entryDate: "",
    memo: "",
    createdAt: nowIso(),
    updatedAt: nowIso(),
  };
  responseFile.file.annotation.title = title || responseFile.file.annotation.title || responseFile.file.label || "アンケート画像";
  responseFile.file.annotation.entryDate = entryDate;
  responseFile.file.annotation.memo = memo;
  responseFile.file.annotation.updatedAt = nowIso();
  return respondentProfileRecords(data, respondentId).find(function (item) {
    return item.sourceType === "response" && Number(item.recordId) === Number(recordId);
  });
}

function deleteProfileRecord(data, respondentId, recordId) {
  const index = data.profileRecords.findIndex(function (item) {
    return item.respondentId === respondentId && Number(item.id) === Number(recordId);
  });
  if (index < 0) {
    throw notFound("画像記録が見つかりません。");
  }
  const record = data.profileRecords[index];
  if (record.image && record.image.driveFileId) {
    try {
      DriveApp.getFileById(record.image.driveFileId).setTrashed(true);
    } catch (_error) {
      // noop
    }
  }
  data.profileRecords.splice(index, 1);
  return { deletedId: Number(recordId) };
}

function createMeasurement(data, respondentId, payload) {
  const respondent = respondentById(data, respondentId);
  if (!respondent) {
    throw notFound("回答者が見つかりません。");
  }
  const entryDate = String(payload.entryDate || "").trim();
  if (!entryDate) {
    throw badRequest("計測日は必須です。");
  }
  const category = payload.category ? normalizeRespondentName(payload.category) : "";
  if (category && MEASUREMENT_CATEGORIES.indexOf(category) < 0) {
    throw badRequest("カテゴリは既定の選択肢から選択してください。");
  }
  const record = {
    id: nextCounter(data, "nextMeasurementId"),
    recordId: 0,
    respondentId: respondent.respondentId,
    respondentName: respondent.respondentName,
    date: entryDate,
    category: category,
    waist: Number(payload.waist || 0),
    hip: Number(payload.hip || 0),
    thigh: Number(payload.thigh || 0),
    createdAt: nowIso(),
    updatedAt: nowIso(),
    editable: true,
    deletable: true,
  };
  record.recordId = record.id;
  if (!(record.waist > 0 && record.hip > 0 && record.thigh > 0)) {
    throw badRequest("ウエスト・ヒップ・太ももは0より大きい数値で入力してください。");
  }
  data.measurements.push(record);
  return measurementRecordPayload(record);
}

function updateMeasurement(data, respondentId, recordId, payload) {
  const record = data.measurements.find(function (item) {
    return item.respondentId === respondentId && Number(item.id) === Number(recordId);
  });
  if (!record) {
    throw notFound("計測記録が見つかりません。");
  }
  record.date = String(payload.entryDate || record.date).trim();
  record.category = payload.category ? normalizeRespondentName(payload.category) : "";
  record.waist = Number(payload.waist || record.waist);
  record.hip = Number(payload.hip || record.hip);
  record.thigh = Number(payload.thigh || record.thigh);
  record.updatedAt = nowIso();
  if (!(record.waist > 0 && record.hip > 0 && record.thigh > 0)) {
    throw badRequest("ウエスト・ヒップ・太ももは0より大きい数値で入力してください。");
  }
  return measurementRecordPayload(record);
}

function deleteMeasurement(data, respondentId, recordId) {
  const index = data.measurements.findIndex(function (item) {
    return item.respondentId === respondentId && Number(item.id) === Number(recordId);
  });
  if (index < 0) {
    throw notFound("計測記録が見つかりません。");
  }
  data.measurements.splice(index, 1);
  return { deletedId: Number(recordId) };
}

function publicRespondentHistory(data, respondentName) {
  const respondent = respondentByNameMatch(data, respondentName);
  if (!respondent) {
    throw notFound("一致するお名前の履歴が見つかりません。");
  }
  const overview = respondentOverview(data, respondent.respondentId, null);
  const history = respondentHistory(data, respondent.respondentId, null).map(function (item) {
    return {
      id: item.id,
      formId: item.formId,
      formTitle: item.formTitle,
      category: item.category,
      createdAt: item.createdAt,
      answers: (item.answers || []).map(function (answer) {
        return {
          label: answer.label || "",
          value: answer.value || "",
        };
      }),
      files: (item.files || []).map(function (file) {
        return {
          label: file.label || "",
          originalName: file.originalName || "",
          mimeType: file.mimeType || "",
          size: Number(file.size || 0),
        };
      }),
    };
  });
  return { respondent: overview, history: history };
}

function updatePassword(data, payload) {
  const current = sha256Hex(String(payload.currentPassword || ""));
  if (current !== data.settings.adminPasswordSha256) {
    throw badRequest("現在のパスワードが正しくありません。");
  }
  const nextPassword = String(payload.newPassword || "");
  if (!nextPassword || nextPassword.length < 8) {
    throw badRequest("新しいパスワードは8文字以上にしてください。");
  }
  if (nextPassword !== String(payload.confirmPassword || "")) {
    throw badRequest("確認用パスワードが一致しません。");
  }
  data.settings.adminPasswordSha256 = sha256Hex(nextPassword);
  saveSessions({});
  return { ok: true };
}

function createBackup(data) {
  const folder = getBackupsFolder();
  const name = "bijiris-backup-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + ".json";
  const file = folder.createFile(name, JSON.stringify(data), MimeType.PLAIN_TEXT);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    name: name,
    sizeBytes: file.getSize(),
    createdAt: nowIso(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(file.getId()),
  };
}

function listBackups() {
  const files = getBackupsFolder().getFiles();
  const items = [];
  while (files.hasNext()) {
    const file = files.next();
    items.push({
      name: file.getName(),
      sizeBytes: file.getSize(),
      createdAt: file.getDateCreated().toISOString(),
      downloadUrl: "https://drive.google.com/uc?export=download&id=" + encodeURIComponent(file.getId()),
    });
  }
  return items.sort(function (a, b) {
    return String(b.createdAt).localeCompare(String(a.createdAt));
  });
}

function operationsStatus(data) {
  const responseImageCount = data.responses.reduce(function (memo, item) {
    return memo + (item.files || []).length;
  }, 0);
  const profileImageCount = data.profileRecords.filter(function (item) {
    return item.image && item.image.url;
  }).length;
  return {
    localUrl: "-",
    publicUrl: data.settings.publicBaseUrl || "",
    publicBaseUrlSource: data.settings.publicBaseUrl ? "config" : "auto",
    configuredPublicBaseUrl: data.settings.publicBaseUrl || "",
    defaultPasswordInUse: data.settings.adminPasswordSha256 === sha256Hex(BIJIRIS_DEFAULT_PASSWORD),
    databasePath: "Google Drive / " + BIJIRIS_DATA_FILE_NAME,
    uploadsPath: "Google Drive / " + BIJIRIS_UPLOADS_FOLDER_NAME,
    backupsPath: "Google Drive / " + BIJIRIS_BACKUPS_FOLDER_NAME,
    databaseSizeBytes: getOrCreateDataFile().getSize(),
    uploadFileCount: responseImageCount + profileImageCount,
    localImageCount: 0,
    externalImageCount: 0,
    responseImageCount: responseImageCount,
    profileImageCount: profileImageCount,
    backupCount: listBackups().length,
    latestBackup: listBackups()[0] || null,
    publicUrlIsTemporary: false,
  };
}

function handleLogin(data, payload) {
  const password = String((payload.body || {}).password || "");
  if (sha256Hex(password) !== data.settings.adminPasswordSha256) {
    return { statusCode: 401, error: "パスワードが正しくありません。" };
  }
  const token = createAdminSession();
  return {
    statusCode: 200,
    data: {
      ok: true,
      authToken: token,
      username: data.settings.adminUsername || "admin",
    },
  };
}

function requireAdmin(data, payload) {
  const token = String(payload.authToken || "").trim();
  if (!verifyAdminSession(token)) {
    return { statusCode: 401, error: "認証が必要です。" };
  }
  return null;
}

function handleApiRequest(payload) {
  return withDataStore(function (data) {
    const method = String(payload.method || "GET").toUpperCase();
    const parsed = parsePathAndQuery(payload.path);
    const path = parsed.path;
    const query = parsed.query;

    if (method === "POST" && path === "/api/admin/login") {
      return handleLogin(data, payload);
    }
    if (method === "POST" && path === "/api/admin/logout") {
      revokeAdminSession(String(payload.authToken || "").trim());
      return { statusCode: 200, data: { ok: true } };
    }

    if (method === "GET" && path === "/api/public/forms") {
      return { statusCode: 200, data: { forms: publicFormsPayload(data) } };
    }
    const formGetMatch = path.match(/^\/api\/public\/forms\/([^/]+)$/);
    if (method === "GET" && formGetMatch) {
      return { statusCode: 200, data: { form: publicFormPayload(data, decodeURIComponent(formGetMatch[1])) } };
    }
    const submitMatch = path.match(/^\/api\/public\/forms\/([^/]+)\/submit$/);
    if (method === "POST" && submitMatch) {
      const form = formBySlug(data, decodeURIComponent(submitMatch[1]), false);
      if (!form) {
        return { statusCode: 404, error: "フォームが見つかりません。" };
      }
      return { statusCode: 200, data: createPublicResponse(data, form, payload) };
    }
    if (method === "GET" && path === "/api/public/respondents/history") {
      return { statusCode: 200, data: publicRespondentHistory(data, query.name || "") };
    }

    const adminError = requireAdmin(data, payload);
    if (adminError) {
      return adminError;
    }

    if (method === "GET" && path === "/api/admin/bootstrap") {
      const forms = data.forms
        .map(function (form) {
          const next = clone(form);
          next.responseCount = data.responses.filter(function (response) {
            return Number(response.formId) === Number(form.id);
          }).length;
          next.respondentCount = data.responses.filter(function (response) {
            return Number(response.formId) === Number(form.id);
          }).reduce(function (memo, response) {
            memo[response.respondentId] = true;
            return memo;
          }, {});
          next.respondentCount = Object.keys(next.respondentCount).length;
          return next;
        })
        .sort(function (a, b) {
          return String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")) || Number(b.id) - Number(a.id);
        });
      return {
        statusCode: 200,
        data: {
          forms: forms,
          stats: {
            formCount: forms.length,
            responseCount: data.responses.length,
            respondentCount: data.respondents.length,
          },
          recentResponses: listResponses(data, { limit: 8 }),
          publicBaseUrl: data.settings.publicBaseUrl || "",
          defaultPassword: data.settings.adminPasswordSha256 === sha256Hex(BIJIRIS_DEFAULT_PASSWORD) ? BIJIRIS_DEFAULT_PASSWORD : null,
          settings: {
            configuredPublicBaseUrl: data.settings.publicBaseUrl || "",
            publicBaseUrlSource: data.settings.publicBaseUrl ? "config" : "auto",
            publicBaseUrl: data.settings.publicBaseUrl || "",
            defaultPasswordInUse: data.settings.adminPasswordSha256 === sha256Hex(BIJIRIS_DEFAULT_PASSWORD),
          },
          operationsStatus: operationsStatus(data),
          backups: listBackups(),
        },
      };
    }

    if (method === "GET" && path === "/api/admin/operations/status") {
      return { statusCode: 200, data: { status: operationsStatus(data), backups: listBackups() } };
    }
    if (method === "POST" && path === "/api/admin/settings/public-base-url") {
      data.settings.publicBaseUrl = String((payload.body || {}).publicBaseUrl || "").trim();
      return {
        statusCode: 200,
        data: {
          ok: true,
          configuredPublicBaseUrl: data.settings.publicBaseUrl,
          publicBaseUrl: data.settings.publicBaseUrl,
          publicBaseUrlSource: data.settings.publicBaseUrl ? "config" : "auto",
        },
      };
    }
    if (method === "POST" && path === "/api/admin/settings/password") {
      return { statusCode: 200, data: updatePassword(data, payload.body || {}) };
    }
    if (method === "POST" && path === "/api/admin/backups/create") {
      const backup = createBackup(data);
      return { statusCode: 200, data: { ok: true, backup: backup, backups: listBackups() } };
    }
    if (method === "GET" && path === "/api/admin/forms") {
      return { statusCode: 200, data: { forms: data.forms.map(clone) } };
    }
    if (method === "POST" && path === "/api/admin/forms") {
      return { statusCode: 200, data: { ok: true, form: saveForm(data, payload.body || {}, null) } };
    }
    const formIdMatch = path.match(/^\/api\/admin\/forms\/(\d+)$/);
    if (method === "PUT" && formIdMatch) {
      return {
        statusCode: 200,
        data: { ok: true, form: saveForm(data, payload.body || {}, Number(formIdMatch[1])) },
      };
    }
    const formToggleMatch = path.match(/^\/api\/admin\/forms\/(\d+)\/toggle$/);
    if (method === "POST" && formToggleMatch) {
      return { statusCode: 200, data: { ok: true, form: toggleForm(data, Number(formToggleMatch[1])) } };
    }
    const formResponsesMatch = path.match(/^\/api\/admin\/forms\/(\d+)\/responses$/);
    if (method === "GET" && formResponsesMatch) {
      const formId = Number(formResponsesMatch[1]);
      return {
        statusCode: 200,
        data: {
          responses: listResponses(data, {
            formId: formId,
            respondentQuery: query.respondent || "",
            category: query.category || "",
            limit: 100,
          }),
          categorySummary: categorySummary(data, formId, query.respondent || ""),
        },
      };
    }
    const responseDetailMatch = path.match(/^\/api\/admin\/responses\/(\d+)$/);
    if (method === "GET" && responseDetailMatch) {
      const detail = responseDetail(data, Number(responseDetailMatch[1]));
      if (!detail) {
        return { statusCode: 404, error: "回答が見つかりません。" };
      }
      return { statusCode: 200, data: detail };
    }
    if (method === "POST" && path === "/api/admin/respondents/create") {
      const respondent = ensureRespondentRegistry(data, String((payload.body || {}).name || ""));
      return { statusCode: 200, data: { ok: true, respondent: respondentOverview(data, respondent.respondentId, null) } };
    }
    if (method === "GET" && path === "/api/admin/respondents") {
      const limit = Math.min(Math.max(Number(query.limit || 100), 1), 1000);
      return {
        statusCode: 200,
        data: {
          respondents: respondentSummary(data, query.form_id ? Number(query.form_id) : null, query.q || "", limit),
        },
      };
    }
    if (method === "GET" && path === "/api/admin/measurements") {
      return {
        statusCode: 200,
        data: {
          records: listMeasurementRecords(data, {
            respondentId: query.respondent_id || "",
            respondentName: query.respondent_name || "",
            query: query.q || "",
            limit: Math.min(Math.max(Number(query.limit || 500), 1), 2000),
          }),
        },
      };
    }
    const respondentHistoryMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/history$/);
    if (method === "GET" && respondentHistoryMatch) {
      const respondentId = decodeURIComponent(respondentHistoryMatch[1]);
      const respondent = respondentOverview(data, respondentId, query.form_id ? Number(query.form_id) : null);
      if (!respondent) {
        return { statusCode: 404, error: "回答者が見つかりません。" };
      }
      return {
        statusCode: 200,
        data: {
          respondent: respondent,
          history: respondentHistory(data, respondentId, query.form_id ? Number(query.form_id) : null),
          imageRecords: respondentProfileRecords(data, respondentId),
          profileRecords: respondentProfileRecords(data, respondentId),
          measurementRecords: listMeasurementRecords(data, {
            respondentId: respondentId,
            limit: 500,
          }),
        },
      };
    }
    const respondentProfileMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/profile$/);
    if (method === "POST" && respondentProfileMatch) {
      return {
        statusCode: 200,
        data: Object.assign({ ok: true }, updateRespondentProfile(data, decodeURIComponent(respondentProfileMatch[1]), payload.body || {})),
      };
    }
    const respondentDeleteMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/delete$/);
    if (method === "POST" && respondentDeleteMatch) {
      const respondentId = decodeURIComponent(respondentDeleteMatch[1]);
      const responsesBefore = data.responses.length;
      data.responses = data.responses.filter(function (item) {
        return item.respondentId !== respondentId;
      });
      data.profileRecords = data.profileRecords.filter(function (item) {
        return item.respondentId !== respondentId;
      });
      data.measurements = data.measurements.filter(function (item) {
        return item.respondentId !== respondentId;
      });
      data.respondents = data.respondents.filter(function (item) {
        return item.respondentId !== respondentId;
      });
      return {
        statusCode: 200,
        data: { ok: true, deletedCount: Math.max(0, responsesBefore - data.responses.length), registryDeleted: true },
      };
    }
    const profileCreateMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/profile-records$/);
    if (method === "POST" && profileCreateMatch) {
      return {
        statusCode: 200,
        data: { ok: true, record: createProfileRecord(data, decodeURIComponent(profileCreateMatch[1]), payload) },
      };
    }
    const profileUpdateMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/profile-records\/([^/]+)\/(\d+)\/update$/);
    if (method === "POST" && profileUpdateMatch) {
      return {
        statusCode: 200,
        data: {
          ok: true,
          record: updateProfileRecord(
            data,
            decodeURIComponent(profileUpdateMatch[1]),
            decodeURIComponent(profileUpdateMatch[2]),
            Number(profileUpdateMatch[3]),
            payload.body || {}
          ),
        },
      };
    }
    const profileDeleteMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/profile-records\/([^/]+)\/(\d+)\/delete$/);
    if (method === "POST" && profileDeleteMatch) {
      if (decodeURIComponent(profileDeleteMatch[2]) !== "manual") {
        return { statusCode: 400, error: "アンケート画像は削除できません。" };
      }
      return {
        statusCode: 200,
        data: Object.assign(
          { ok: true },
          deleteProfileRecord(data, decodeURIComponent(profileDeleteMatch[1]), Number(profileDeleteMatch[3]))
        ),
      };
    }
    const measurementCreateMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/measurements$/);
    if (method === "POST" && measurementCreateMatch) {
      return {
        statusCode: 200,
        data: { ok: true, record: createMeasurement(data, decodeURIComponent(measurementCreateMatch[1]), payload.body || {}) },
      };
    }
    const measurementImportMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/measurements\/import-sheet$/);
    if (method === "POST" && measurementImportMatch) {
      return {
        statusCode: 400,
        error: "Googleスプレッドシート取込はこのデプロイではまだ有効化していません。",
      };
    }
    const measurementUpdateMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/measurements\/(\d+)\/update$/);
    if (method === "POST" && measurementUpdateMatch) {
      return {
        statusCode: 200,
        data: {
          ok: true,
          record: updateMeasurement(
            data,
            decodeURIComponent(measurementUpdateMatch[1]),
            Number(measurementUpdateMatch[2]),
            payload.body || {}
          ),
        },
      };
    }
    const measurementDeleteMatch = path.match(/^\/api\/admin\/respondents\/([^/]+)\/measurements\/(\d+)\/delete$/);
    if (method === "POST" && measurementDeleteMatch) {
      return {
        statusCode: 200,
        data: {
          ok: true,
          deletedId: deleteMeasurement(
            data,
            decodeURIComponent(measurementDeleteMatch[1]),
            Number(measurementDeleteMatch[2])
          ).deletedId,
        },
      };
    }
    return { statusCode: 404, error: "API が見つかりません。" };
  }, shouldPersistRequest(payload));
}

function shouldPersistRequest(payload) {
  const method = String(payload.method || "GET").toUpperCase();
  const path = parsePathAndQuery(payload.path).path;
  if (method === "GET") {
    return false;
  }
  return path !== "/api/admin/login" && path !== "/api/admin/logout";
}

function handleInitializeData(payload) {
  const password = String(payload.password || "");
  return withDataStore(function (data) {
    if (sha256Hex(password) !== data.settings.adminPasswordSha256) {
      return { statusCode: 401, error: "パスワードが正しくありません。" };
    }
    const imported = normalizeDataShape(clone(payload.data || {}));
    writeDataStore(imported);
    return {
      statusCode: 200,
      data: {
        ok: true,
        forms: imported.forms.length,
        respondents: imported.respondents.length,
        responses: imported.responses.length,
      },
    };
  }, false);
}

function handleUploadLocalAsset(payload) {
  const password = String(payload.password || "");
  return withDataStore(function (data) {
    if (sha256Hex(password) !== data.settings.adminPasswordSha256) {
      return { statusCode: 401, error: "パスワードが正しくありません。" };
    }
    const fileInfo = payload.file || null;
    if (!fileInfo || !fileInfo.base64) {
      return { statusCode: 400, error: "ファイルデータがありません。" };
    }
    const uploaded = uploadedFilePayload(fileInfo, nextCounter(data, "nextResponseFileId"));
    if (payload.kind === "responseFile") {
      let applied = false;
      data.responses.some(function (response) {
        if (Number(response.id) !== Number(payload.responseId)) {
          return false;
        }
        return (response.files || []).some(function (file) {
          if (Number(file.id) !== Number(payload.fileId)) {
            return false;
          }
          file.originalName = uploaded.originalName;
          file.storedName = uploaded.storedName;
          file.mimeType = uploaded.mimeType;
          file.size = uploaded.size;
          file.relativePath = uploaded.relativePath;
          file.url = uploaded.url;
          file.previewUrl = uploaded.previewUrl;
          file.driveFileId = uploaded.driveFileId;
          applied = true;
          return true;
        });
      });
      if (!applied) {
        return { statusCode: 404, error: "対象の画像が見つかりません。" };
      }
      return { statusCode: 200, data: { ok: true, file: uploaded } };
    }
    if (payload.kind === "profileRecord") {
      const record = findManualProfileRecord(data, String(payload.respondentId || ""), Number(payload.recordId));
      if (!record) {
        return { statusCode: 404, error: "対象の画像記録が見つかりません。" };
      }
      record.image = sanitizeImageObject(uploaded);
      record.updatedAt = nowIso();
      return { statusCode: 200, data: { ok: true, record: clone(record) } };
    }
    return { statusCode: 400, error: "不明な画像種別です。" };
  }, true);
}
