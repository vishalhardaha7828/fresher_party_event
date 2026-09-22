/**
 * Fresher & Farewell Event Manager — Google Apps Script backend.
 */

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const MEDIA_FOLDER_ID = "15klBTOvfsAOlDw4xm9WlFa9peC5qOw_O";
const BILL_FOLDER_ID = "1KlzRbFcBaxNxb1ji9WLvI45keFn0NWXK";
const QR_FOLDER_ID = "1btQmTpV23_PH8jX3sligmVdmggOAsjF5";
const SONG_FOLDER_ID = "1_Od2xPjjCR3_dBbrGlGWRzb9r1cpGEns";

const SESSION_TTL_SECONDS = 6 * 60 * 60;
const FAILED_LOGIN_DELAY_MS = 400;
const MAX_TEXT_LENGTH = 500;
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const ALLOWED_UPLOAD_TYPES = /^(image|video)\//i;
const ALLOWED_SONG_TYPES = /^(audio\/mpeg|audio\/mp3)$/i;

const ACTION_ROLES = {
  getVolunteerDashboard: ["Admin", "Volunteer"],
  getFeedback: ["Admin"],
  deleteFeedback: ["Admin"],
  updateParticipantOrder: ["Admin"],
  completeParticipant: ["Admin"],
  approveVolunteer: ["Admin"],
  updateVolunteer: ["Admin"],
  deleteVolunteer: ["Admin"],
  addCollection: ["Admin", "Volunteer"],
  updateCollection: ["Admin"],
  deleteCollection: ["Admin"],
  settleVolunteerCollection: ["Admin"],
  addExpense: ["Admin"],
  updateExpense: ["Admin"],
  deleteExpense: ["Admin"],
  addTask: ["Admin"],
  updateTaskStatus: ["Admin"],
  addMeeting: ["Admin"],
  uploadMedia: ["Admin"]
};

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function ok(extra) {
  return json(Object.assign({ status: "success" }, extra || {}));
}

function fail(message) {
  return json({ status: "error", message: message });
}

function text(value, limit) {
  return String(value == null ? "" : value).trim().slice(0, limit || MAX_TEXT_LENGTH);
}

function phoneOf(value) {
  const digits = String(value == null ? "" : value).replace(/\D/g, "");
  return digits.length === 10 ? digits : "";
}

function amountOf(value) {
  const number = Number(value);
  return isFinite(number) && number >= 0 ? number : null;
}

function safeEquals(left, right) {
  const a = String(left);
  const b = String(right);
  if (a.length !== b.length) return false;
  let difference = 0;
  for (let i = 0; i < a.length; i++) difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return difference === 0;
}

function sessionKey(token) {
  return "session:" + token;
}

function createSession(payload) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put(sessionKey(token), JSON.stringify(payload), SESSION_TTL_SECONDS);
  return token;
}

function readSession(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get(sessionKey(String(token)));
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (error) {
    return null;
  }
}

function destroySession(token) {
  if (token) CacheService.getScriptCache().remove(sessionKey(String(token)));
}

function authorize(action, token) {
  const roles = ACTION_ROLES[action];
  const session = readSession(token);
  if (!roles) return { allowed: true, session: session };
  if (!session || roles.indexOf(session.role) < 0) {
    return { allowed: false, response: fail("Session expire ho gaya hai. Dobara login karein.") };
  }
  return { allowed: true, session: session };
}

function sheet(name, aliases) {
  const names = [name].concat(aliases || []);
  for (const item of names) {
    const found = SPREADSHEET.getSheetByName(item);
    if (found) return found;
  }
  return null;
}

function headersOf(current) {
  return current ? current.getRange(1, 1, 1, current.getLastColumn()).getValues()[0].map(String) : [];
}

function col(headers, names) {
  for (const name of names) {
    const wanted = name.toLowerCase().replace(/[\s_]+/g, "");
    const index = headers.findIndex(value => String(value).toLowerCase().replace(/[\s_]+/g, "") === wanted);
    if (index >= 0) return index;
  }
  return -1;
}

function valueAt(row, headers, names) {
  const index = col(headers, names);
  return index >= 0 ? row[index] : "";
}

function duplicateColumn(headers, name, occurrence) {
  const wanted = String(name).toLowerCase().replace(/[\s_\/]+/g, "");
  let found = 0;
  for (let index = 0; index < headers.length; index++) {
    const current = String(headers[index] || "").toLowerCase().replace(/[\s_\/]+/g, "");
    if (current === wanted) {
      found += 1;
      if (found === occurrence) return index;
    }
  }
  return -1;
}

function dataRows(name, aliases) {
  const current = sheet(name, aliases);
  if (!current || current.getLastRow() < 2) return [];
  return current.getDataRange().getValues().slice(1);
}

function addRow(name, values) {
  let current = sheet(name);
  if (!current) current = SPREADSHEET.insertSheet(name);
  current.appendRow(values);
  return ok();
}

function updateRow(name, rowIndex, values) {
  const current = sheet(name);
  const sheetRow = Number(rowIndex) + 1;
  if (!current || !rowIndex || sheetRow <= 1 || sheetRow > current.getLastRow()) return fail("Row not found");
  current.getRange(sheetRow, 1, 1, values.length).setValues([values]);
  return ok();
}

function existingCellValue(name, rowIndex, column) {
  const current = sheet(name);
  const sheetRow = Number(rowIndex) + 1;
  if (!current || sheetRow <= 1 || sheetRow > current.getLastRow() || column > current.getLastColumn()) return "";
  return current.getRange(sheetRow, column).getValue() || "";
}

function deleteRow(name, rowIndex) {
  const current = sheet(name);
  const sheetRow = Number(rowIndex) + 1;
  if (!current || sheetRow <= 1 || sheetRow > current.getLastRow()) return fail("Row not found");
  current.deleteRow(sheetRow);
  return ok();
}

function doGet(e) {
  try {
    const parameters = (e && e.parameter) || {};
    const action = String(parameters.action || "");
    const auth = authorize(action, parameters.token);
    if (!auth.allowed) return auth.response;

    switch (action) {
      case "getData": return getData();
      case "getMedia": return getMedia();
      case "getApprovedVolunteers": return getVolunteers("approved");
      case "getPendingVolunteers": return getVolunteers("pending");
      case "getFeedback": return getFeedback();
      case "checkStatus": return checkStatus(parameters.phone);
      case "getVolunteerDashboard": return getVolunteerDashboard(auth.session);
      default: return fail("Invalid GET action");
    }
  } catch (error) {
    return fail(error.toString());
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const action = String(data.action || "");

    if (action === "adminLogin") return adminLogin(data.pin);
    if (action === "volunteerLogin") return volunteerLogin(data.phone, data.password);
    if (action === "logout") { destroySession(data.token); return ok(); }

    const auth = authorize(action, data.token);
    if (!auth.allowed) return auth.response;
    const session = auth.session;

    switch (action) {
      case "registerVolunteer": return registerVolunteer(data);
      case "approveVolunteer": return approveVolunteer(data);
      case "updateVolunteer": return updateVolunteer(data);
      case "deleteVolunteer": return deleteVolunteer(data);
      case "submitFeedback": return submitFeedback(data);
      case "deleteFeedback": return deleteFeedback(data);
      case "addCollection": return saveCollection(data, session, false);
      case "updateCollection": return saveCollection(data, session, true);
      case "deleteCollection": return deleteRow("Collection", data.rowIndex);
      case "settleVolunteerCollection": return settleVolunteerCollection(data);
      case "addExpense": return saveExpense(data, false);
      case "updateExpense": return saveExpense(data, true);
      case "deleteExpense": return deleteRow("Expenses", data.rowIndex);
      case "addParticipant": return addParticipant(data);
      case "updateParticipantOrder": return updateParticipantOrder(data);
      case "completeParticipant": return completeParticipant(data);
      case "addTask": return addTask(data);
      case "updateTaskStatus": return updateTaskStatus(data.title);
      case "addMeeting": return addMeeting(data);
      case "uploadMedia": return uploadMedia(data);
      default: return fail("Invalid POST action");
    }
  } catch (error) {
    return fail(error.toString());
  }
}

function adminLogin(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  if (!expected) return fail("Server par ADMIN_PIN script property set nahi hai.");
  if (!safeEquals(expected, text(pin, 64))) {
    Utilities.sleep(FAILED_LOGIN_DELAY_MS);
    return fail("Galat Admin PIN!");
  }
  return ok({ role: "Admin", token: createSession({ role: "Admin" }) });
}

function volunteerLogin(phone, password) {
  const found = findVolunteer(phone);
  const supplied = String(password || "");
  if (!found) {
    Utilities.sleep(FAILED_LOGIN_DELAY_MS);
    return fail("Invalid credentials or approval pending");
  }

  const row = found.current.getDataRange().getValues()[found.row - 1];
  const passwordIndex = col(found.headers, ["Password"]);
  const status = String(valueAt(row, found.headers, ["Status", "Approval Status"])).trim().toLowerCase();

  // Plaintext password comparison
  if (passwordIndex < 0 || !safeEquals(String(row[passwordIndex]), supplied) || status !== "approved") {
    Utilities.sleep(FAILED_LOGIN_DELAY_MS);
    return fail("Invalid credentials or approval pending");
  }

  const name = valueAt(row, found.headers, ["Name", "Volunteer Name"]);
  const branch = valueAt(row, found.headers, ["Branch"]);
  const semester = valueAt(row, found.headers, ["Semester"]);
  const qrUrl = valueAt(row, found.headers, ["QRURL", "QR URL", "QR Code URL"]);
  const token = createSession({
    role: "Volunteer",
    name: String(name),
    phone: phoneOf(phone),
    branch: String(branch),
    semester: String(semester),
    qrUrl: String(qrUrl)
  });
  return ok({ role: "Volunteer", name: name, branch: branch, semester: semester, phone: phoneOf(phone), qrUrl: qrUrl, token: token });
}

function volunteerSheet() {
  return sheet("Volunteers", ["Volunteer", "Volunteers Data"]);
}

function normalizeVolunteerName(value) {
  return String(value || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function volunteerMetricsByPhone() {
  const totals = {};
  const volunteerNameMap = {};
  const volunteerFirstNameMap = {};
  const current = volunteerSheet();

  if (current && current.getLastRow() >= 2) {
    const headers = headersOf(current);
    current.getDataRange().getValues().slice(1).forEach(row => {
      const phone = phoneOf(valueAt(row, headers, ["Phone", "Phone Number"]));
      const name = text(valueAt(row, headers, ["Name", "Volunteer Name"]), 120);
      const normalizedName = normalizeVolunteerName(name);
      if (!normalizedName) return;

      volunteerNameMap[normalizedName] = phone || volunteerNameMap[normalizedName] || "";
      const tokens = normalizedName.split(/\s+/).filter(Boolean);
      tokens.forEach(token => {
        if (token) volunteerFirstNameMap[token] = phone || volunteerFirstNameMap[token] || "";
      });
      if (tokens.length > 1) {
        const compactName = tokens.join("");
        volunteerNameMap[compactName] = phone || volunteerNameMap[compactName] || "";
      }
    });
  }

  collectionRowsWithOwner().forEach(item => {
    if (item.settled) return;
    const ownedPhone = phoneOf(item.phone);
    const ownerName = text(item.name, 120);
    const normalizedOwnerName = normalizeVolunteerName(ownerName);
    const ownerTokens = normalizedOwnerName ? normalizedOwnerName.split(/\s+/).filter(Boolean) : [];

    let phone = ownedPhone || "";
    if (!phone && normalizedOwnerName) {
      const candidateNames = [normalizedOwnerName, ownerTokens.join(""), ...ownerTokens, ownerTokens[0] || ""];
      for (const candidate of candidateNames) {
        if (!candidate) continue;
        const matchedPhone = volunteerNameMap[candidate] || volunteerFirstNameMap[candidate];
        if (matchedPhone) {
          phone = matchedPhone;
          break;
        }
      }
    }

    if (!phone) return;
    if (!totals[phone]) totals[phone] = { total: 0, entries: 0 };
    totals[phone].total += Number(item.amount) || 0;
    totals[phone].entries += 1;
  });

  return totals;
}

function getVolunteers(status) {
  const current = volunteerSheet();
  if (!current || current.getLastRow() < 2) return ok({ volunteers: [] });
  const headers = headersOf(current);
  const statusIndex = col(headers, ["Status", "Approval Status"]);
  const totals = volunteerMetricsByPhone();
  const volunteers = current.getDataRange().getValues().slice(1).map((row, offset) => {
    const phone = phoneOf(valueAt(row, headers, ["Phone", "Phone Number"]));
    const stats = totals[phone] || { total: 0, entries: 0 };
    return {
      name: valueAt(row, headers, ["Name", "Volunteer Name"]),
      branch: valueAt(row, headers, ["Branch"]),
      semester: valueAt(row, headers, ["Semester"]),
      phone: phone,
      qrUrl: valueAt(row, headers, ["QRURL", "QR URL", "QR Code URL"]),
      status: statusIndex >= 0 ? row[statusIndex] : "",
      total: stats.total,
      entries: stats.entries,
      rowIndex: offset + 2
    };
  }).filter(item => String(item.status).trim().toLowerCase() === status)
    .sort((left, right) => (Number(right.total) || 0) - (Number(left.total) || 0) || (String(left.name || "").localeCompare(String(right.name || ""))));
  return ok({ volunteers: volunteers });
}

function findVolunteer(phone) {
  const current = volunteerSheet();
  const wanted = phoneOf(phone);
  if (!current || !wanted || current.getLastRow() < 2) return null;
  const headers = headersOf(current);
  const phoneIndex = col(headers, ["Phone", "Phone Number"]);
  if (phoneIndex < 0) return null;
  const rows = current.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (phoneOf(rows[i][phoneIndex]) === wanted) return { current: current, headers: headers, row: i + 1 };
  }
  return null;
}

function registerVolunteer(data) {
  const name = text(data.name, 120);
  const phone = phoneOf(data.phone);
  const password = String(data.password || "");

  if (!name) return fail("Naam zaroori hai.");
  if (!phone) return fail("Phone number exact 10 digits ka hona chahiye.");
  if (password.length < 6) return fail("Password kam se kam 6 characters ka hona chahiye.");

  let current = volunteerSheet();
  if (!current) {
    current = SPREADSHEET.insertSheet("Volunteers");
    current.appendRow(["Name", "Branch", "Semester", "Phone", "Password", "Status", "Time", "QR URL"]);
  }
  
  if (findVolunteer(phone)) return fail("Ye phone number pehle se registered hai.");

  let qrUrl = "";
  if (data.qrFileData) {
    qrUrl = saveDriveFile(data.qrFileData, data.qrFileName, data.qrFileMimeType, true, QR_FOLDER_ID || MEDIA_FOLDER_ID);
  }

  // Exact sequence match: Name, Branch, Semester, Phone, Password (plaintext), Status, Time, QR URL
  current.appendRow([
    name,
    text(data.branch, 40),
    text(data.semester, 20),
    phone,
    password, // Plaintext password saved directly in sheet
    "Pending",
    new Date(),
    qrUrl
  ]);
  
  return ok();
}

function approveVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return fail("Volunteer not found");
  const index = col(found.headers, ["Status", "Approval Status"]);
  if (index < 0) return fail("Status column not found");
  found.current.getRange(found.row, index + 1).setValue("Approved");
  return ok();
}

function updateVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return fail("Volunteer not found");
  const updates = [
    [["Name", "Volunteer Name"], text(data.name, 120)],
    [["Branch"], text(data.branch, 40)],
    [["Semester"], text(data.semester, 20)]
  ];
  updates.forEach(pair => {
    const index = col(found.headers, pair[0]);
    if (index >= 0) found.current.getRange(found.row, index + 1).setValue(pair[1]);
  });
  return ok();
}

function deleteVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return fail("Volunteer not found");
  found.current.deleteRow(found.row);
  return ok();
}

function getData() {
  const result = {};
  ensureParticipantSheet();
  ensureTaskSheet();
  ["Collection", "Expenses", "Participants", "Tasks", "Meetings", "Settings"].forEach(name => {
    const current = sheet(name);
    if (current) result[name] = current.getDataRange().getValues();
  });
  return json(result);
}

function getVolunteerDashboard(session) {
  const current = volunteerSheet();
  if (!current) return ok({ volunteer: null, stats: { total: 0, entries: 0 }, volunteers: [] });
  const headers = headersOf(current);
  const totals = volunteerMetricsByPhone();
  const settlements = settlementSummaryByPhone();
  const volunteers = current.getDataRange().getValues().slice(1).map(row => {
    const phone = phoneOf(valueAt(row, headers, ["Phone", "Phone Number"]));
    const stats = totals[phone] || { total: 0, entries: 0 };
    return {
      name: valueAt(row, headers, ["Name", "Volunteer Name"]),
      phone: phone,
      branch: valueAt(row, headers, ["Branch"]),
      semester: valueAt(row, headers, ["Semester"]),
      status: valueAt(row, headers, ["Status", "Approval Status"]),
      qrUrl: valueAt(row, headers, ["QRURL", "QR URL", "QR Code URL"]),
      total: stats.total,
      entries: stats.entries,
      lastSettlement: settlements[phone] || null
    };
  }).filter(item => String(item.status).toLowerCase() === "approved")
    .sort((left, right) => (Number(right.total) || 0) - (Number(left.total) || 0) || (String(left.name || "").localeCompare(String(right.name || ""))));
  const phone = session && session.role === "Volunteer" ? phoneOf(session.phone) : "";
  return ok({
    volunteer: phone ? volunteers.find(item => item.phone === phone) || null : null,
    stats: phone ? (totals[phone] || { total: 0, entries: 0 }) : { total: 0, entries: 0 },
    volunteers: session && session.role === "Admin" ? volunteers : [],
    settlement: phone ? (settlements[phone] || null) : null
  });
}

function settlementSummaryByPhone() {
  const current = sheet("Collection");
  if (!current || current.getLastRow() < 2) return {};
  const headers = headersOf(current);
  const phoneIndex = col(headers, ["Collected By Phone", "Volunteer Phone"]);
  const nameIndex = col(headers, ["Collected By", "Volunteer Name"]);
  const settledIndex = col(headers, ["Settled", "Settlement Status"]);
  const amountIndex = col(headers, ["Settlement Amount"]);
  const dateIndex = col(headers, ["Settlement Date"]);
  if (settledIndex < 0) return {};

  const result = {};
  current.getDataRange().getValues().slice(1).forEach(row => {
    if (String(row[settledIndex] || "").trim().toLowerCase() !== "yes") return;
    const phone = phoneOf(phoneIndex >= 0 ? row[phoneIndex] : "");
    const name = text(nameIndex >= 0 ? row[nameIndex] : "", 120);
    const key = phone || normalizeVolunteerName(name);
    if (!key) return;
    const amount = amountIndex >= 0 ? Number(row[amountIndex]) || 0 : 0;
    const date = dateIndex >= 0 ? row[dateIndex] : "";
    if (!result[key]) result[key] = { amount: 0, date: date || "", settled: true };
    result[key].amount = Math.max(result[key].amount, amount);
    if (date) result[key].date = date;
  });
  return result;
}

function normalizeCollectionOwnerColumns() {
  const current = sheet("Collection");
  if (!current || current.getLastRow() < 2) return current;

  const headers = headersOf(current);
  const phoneIndex = col(headers, ["Collected By Phone", "Volunteer Phone"]);
  const nameIndex = col(headers, ["Collected By", "Volunteer Name"]);
  if (phoneIndex < 0 || nameIndex < 0 || phoneIndex === nameIndex) return current;

  const rows = current.getDataRange().getValues();
  rows.slice(1).forEach((row, offset) => {
    const rawPhone = row[phoneIndex];
    const rawName = row[nameIndex];
    const hasPhoneValue = phoneOf(rawPhone) !== "";
    const hasNameValueAsPhone = phoneOf(rawName) !== "" && String(rawName).trim() !== "";
    if (!hasPhoneValue && hasNameValueAsPhone) {
      current.getRange(offset + 2, phoneIndex + 1).setValue(rawName);
      current.getRange(offset + 2, nameIndex + 1).setValue(rawPhone);
    }
  });

  return current;
}

function collectionRowsWithOwner() {
  const current = normalizeCollectionOwnerColumns();
  if (!current || current.getLastRow() < 2) return [];
  const headers = headersOf(current);
  return current.getDataRange().getValues().slice(1).map(row => {
    const rawPhone = valueAt(row, headers, ["Collected By Phone", "Volunteer Phone"]);
    const rawName = valueAt(row, headers, ["Collected By", "Volunteer Name"]);
    const fixedPhone = phoneOf(rawPhone) || (phoneOf(rawName) || "");
    const fixedName = text(rawName, 120) || text(rawPhone, 120);
    return {
      amount: valueAt(row, headers, ["Amount"]),
      phone: fixedPhone,
      name: fixedName,
      settled: String(valueAt(row, headers, ["Settled", "Settlement Status"])).trim().toLowerCase() === "yes"
    };
  });
}

function settleVolunteerCollection(data) {
  const targetPhone = phoneOf(data.phone);
  const targetName = text(data.name, 120);
  const current = sheet("Collection");
  if (!current || current.getLastRow() < 2) return ok({ cleared: 0 });

  ensureCollectionSheet();
  let headers = headersOf(current);
  let settledIndex = col(headers, ["Settled", "Settlement Status"]);
  let settledAtIndex = col(headers, ["Settled At", "Settlement Date"]);
  let settlementAmountIndex = col(headers, ["Settlement Amount"]);
  let settlementDateIndex = col(headers, ["Settlement Date"]);
  if (settledIndex < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settled");
    headers = headersOf(current);
    settledIndex = col(headers, ["Settled", "Settlement Status"]);
  }
  if (settledAtIndex < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settled At");
    headers = headersOf(current);
    settledAtIndex = col(headers, ["Settled At", "Settlement Date"]);
  }
  if (settlementAmountIndex < 0 || settlementDateIndex < 0) {
    ensureCollectionSheet();
    headers = headersOf(current);
    settlementAmountIndex = col(headers, ["Settlement Amount"]);
    settlementDateIndex = col(headers, ["Settlement Date"]);
  }
  const rows = current.getDataRange().getValues();
  let settledCount = 0;
  let settlementAmount = 0;
  const settlementDate = new Date();

  rows.slice(1).forEach((row, index) => {
    const alreadySettled = String(row[settledIndex] || "").trim().toLowerCase() === "yes";
    if (alreadySettled) return;
    const rowPhone = phoneOf(valueAt(row, headers, ["Collected By Phone", "Volunteer Phone"]));
    const rowName = text(valueAt(row, headers, ["Collected By", "Volunteer Name"]), 120);
    const samePhone = targetPhone && rowPhone === targetPhone;
    const sameName = !targetPhone && targetName && normalizeVolunteerName(rowName) === normalizeVolunteerName(targetName);
    if (samePhone || sameName) {
      settlementAmount += Number(valueAt(row, headers, ["Amount", "Collected Amount"])) || 0;
      current.getRange(index + 2, settledIndex + 1).setValue("Yes");
      if (settledAtIndex >= 0) current.getRange(index + 2, settledAtIndex + 1).setValue("Yes");
      if (settlementAmountIndex >= 0) current.getRange(index + 2, settlementAmountIndex + 1).setValue(0);
      if (settlementDateIndex >= 0) current.getRange(index + 2, settlementDateIndex + 1).setValue(settlementDate);
      settledCount += 1;
    }
  });

  if (settledCount && settlementAmountIndex >= 0) {
    rows.slice(1).forEach((row, index) => {
      const rowPhone = phoneOf(valueAt(row, headers, ["Collected By Phone", "Volunteer Phone"]));
      const rowName = text(valueAt(row, headers, ["Collected By", "Volunteer Name"]), 120);
      const samePhone = targetPhone && rowPhone === targetPhone;
      const sameName = !targetPhone && targetName && normalizeVolunteerName(rowName) === normalizeVolunteerName(targetName);
      if (samePhone || sameName) current.getRange(index + 2, settlementAmountIndex + 1).setValue(settlementAmount);
    });
  }

  return ok({ cleared: settledCount, amount: settlementAmount, date: settlementDate });
}

function ensureParticipantSheet() {
  let current = sheet("Participants");
  if (!current) return null;

  let headers = headersOf(current);
  let statusIndex = col(headers, ["Program Status", "Status"]);
  let timestampIndex = col(headers, ["Timestamp", "Created At", "Date"]);
  let orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);

  if (timestampIndex < 0) {
    const insertAt = statusIndex >= 0 ? statusIndex + 1 : current.getLastColumn() + 1;
    current.insertColumnBefore(insertAt);
    current.getRange(1, insertAt).setValue("Timestamp");
  }

  headers = headersOf(current);
  statusIndex = col(headers, ["Program Status", "Status"]);
  orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  if (orderIndex < 0) {
    const insertAt = statusIndex >= 0 ? statusIndex + 1 : current.getLastColumn() + 1;
    current.insertColumnBefore(insertAt);
    current.getRange(1, insertAt).setValue("Program Order");
    orderIndex = insertAt - 1;
    const rowCount = current.getLastRow() - 1;
    if (rowCount > 0) {
      current.getRange(2, insertAt, rowCount, 1)
        .setValues(Array.from({ length: rowCount }, (_, index) => [index + 1]));
    }
  }

  headers = headersOf(current);
  statusIndex = col(headers, ["Program Status", "Status"]);
  if (statusIndex < 0) {
    const insertAt = current.getLastColumn() + 1;
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, insertAt).setValue("Program Status");
    const rowCount = current.getLastRow() - 1;
    if (rowCount > 0) {
      current.getRange(2, insertAt, rowCount, 1)
        .setValues(Array.from({ length: rowCount }, () => ["Pending"]));
    }
  }
  headers = headersOf(current);
  if (col(headers, ["Song URL", "Dance Song", "Song"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Song URL");
  }
  normalizeParticipantOrder(current);
  return current;
}

function normalizeParticipantOrder(current) {
  const headers = headersOf(current);
  const orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  const statusIndex = col(headers, ["Program Status", "Status"]);
  if (orderIndex < 0 || current.getLastRow() < 2) return;

  const rows = current.getDataRange().getValues().slice(1)
    .map((row, index) => ({
      sheetRow: index + 2,
      order: Number(row[orderIndex]) || index + 1,
      status: statusIndex >= 0 ? String(row[statusIndex] || "Pending") : "Pending"
    }))
    .sort((left, right) => (left.status === "Completed") - (right.status === "Completed") || left.order - right.order);
  rows.forEach((item, index) => current.getRange(item.sheetRow, orderIndex + 1).setValue(index + 1));
}

function getMedia() {
  const current = sheet("Media");
  if (!current || current.getLastRow() < 2) return ok({ media: [] });
  const media = current.getDataRange().getValues().slice(1).filter(row => row[1]).map(row => ({
    title: row[0] || "Event media",
    fileUrl: row[1],
    type: row[2] || "Photo"
  }));
  return ok({ media: media });
}

function ensureTaskSheet() {
  let current = sheet("Tasks");
  if (!current) {
    current = SPREADSHEET.insertSheet("Tasks");
    current.appendRow(["Title", "Type", "Status", "Time", "Responsible Person", "Timestamp"]);
    return current;
  }

  const headers = headersOf(current);
  if (col(headers, ["Responsible Person", "Responsible", "Assigned To"]) < 0) {
    const timestampIndex = col(headers, ["Timestamp", "Created At", "Date"]);
    const secondTimestamp = String(headers[5] || "").toLowerCase().replace(/\s/g, "") === "timestamp";
    if (timestampIndex === 4 && secondTimestamp) {
      if (current.getLastRow() >= 2) {
        const rows = current.getRange(2, 5, current.getLastRow() - 1, 2).getValues();
        rows.forEach(row => {
          if (!row[1] && row[0]) row[1] = row[0];
          row[0] = "";
        });
        current.getRange(2, 5, rows.length, 2).setValues(rows);
      }
    } else if (timestampIndex === 4) {
      current.insertColumnBefore(5);
    } else if (current.getLastColumn() < 5) {
      current.insertColumnAfter(Math.max(current.getLastColumn(), 1));
    }
    current.getRange(1, 5).setValue("Responsible Person");
  }
  if (current.getLastColumn() < 6) current.insertColumnAfter(5);
  current.getRange(1, 6).setValue("Timestamp");
  while (current.getLastColumn() > 6) {
    const latestHeaders = headersOf(current);
    const lastHeader = String(latestHeaders[latestHeaders.length - 1] || "").toLowerCase().replace(/\s/g, "");
    if (lastHeader !== "timestamp") break;
    const lastRow = current.getLastRow();
    if (lastRow >= 2) {
      const rows = current.getRange(2, 6, lastRow - 1, 2).getValues();
      rows.forEach(row => {
        if (!row[0] && row[1]) row[0] = row[1];
        row[1] = "";
      });
      current.getRange(2, 6, rows.length, 2).setValues(rows);
    }
    current.deleteColumn(current.getLastColumn());
  }
  return current;
}

function getFeedback() {
  const feedbacks = dataRows("Feedback").map((row, index) => ({
    rowIndex: index + 2,
    name: row[0] || "",
    message: row[1] || "",
    timestamp: row[2] || ""
  }));
  return ok({ feedbacks: feedbacks });
}

function submitFeedback(data) {
  const name = text(data.name, 120);
  const message = text(data.message, 2000);
  if (!name || !message) return fail("Naam aur sujhav dono zaroori hain.");

  let current = sheet("Feedback");
  if (!current) {
    current = SPRESSHEET.insertSheet("Feedback");
    current.appendRow(["Name", "Message", "Timestamp"]);
  }
  current.appendRow([name, message, new Date()]);
  return ok();
}

function deleteFeedback(data) {
  const current = sheet("Feedback");
  if (!current) return fail("Feedback sheet not found");
  (data.rows || []).map(Number).filter(isFinite).sort((a, b) => b - a).forEach(row => {
    if (row > 1 && row <= current.getLastRow()) current.deleteRow(row);
  });
  return ok();
}

function saveCollection(data, session, updating) {
  const isVolunteer = session && session.role === "Volunteer";
  const branch = isVolunteer ? text(session.branch, 40) : text(data.branch, 40);
  const semester = isVolunteer ? text(session.semester, 20) : text(data.semester, 20);
  const category = text(data.category, 40) || "Regular";
  const name = text(data.name, 120);
  const amount = amountOf(data.amount);

  if (!name) return fail("Naam zaroori hai.");
  if (amount === null) return fail("Amount ek valid number hona chahiye.");
  if (isVolunteer && (category === "Faculty" || (semester === "5th" && category === "Lateral"))) {
    return fail("Volunteer Faculty entry nahi kar sakta; 5th semester volunteer Lateral Entry bhi nahi bhar sakta.");
  }

  const current = ensureCollectionSheet();
  const ownerPhone = isVolunteer ? phoneOf(session.phone) : phoneOf(data.collectedByPhone);
  const ownerName = isVolunteer ? text(session.name, 120) : text(data.collectedByName, 120);
  const headers = headersOf(current);
  const values = current.getLastColumn() > 0 ? Array(current.getLastColumn()).fill("") : [];
  const setValue = (names, value) => {
    const index = col(headers, names);
    if (index >= 0) values[index] = value;
  };

  setValue(["Name", "Student Name", "Participant Name"], name);
  setValue(["Branch"], branch);
  setValue(["Semester", "Sem"], semester);
  setValue(["Category"], category);
  setValue(["Amount", "Collected Amount"], amount);
  setValue(["Mode", "Payment Mode"], text(data.mode, 20) || "Cash");
  const dateIndex = col(headers, ["Timestamp", "Date", "Date / Time", "Created At"]) >= 0
      ? col(headers, ["Timestamp", "Date", "Date / Time", "Created At"])
      : duplicateColumn(headers, "Mode", 2);
  if (dateIndex >= 0) values[dateIndex] = new Date();
  setValue(["Collected By Phone", "Volunteer Phone"], ownerPhone);
  setValue(["Collected By", "Volunteer Name", "Collector"], ownerName);
  if (!updating) {
    setValue(["Settled", "Settlement Status"], "");
    setValue(["Settled At", "Settlement Date"], "");
  }

  if (updating) {
    const sheetRow = Number(data.rowIndex) + 1;
    if (!data.rowIndex || sheetRow <= 1 || sheetRow > current.getLastRow()) return fail("Row not found");
    current.getRange(sheetRow, 1, 1, values.length).setValues([values]);
    return ok();
  }
  current.appendRow(values);
  return ok();
}

function ensureCollectionSheet() {
  let current = sheet("Collection");
  if (!current) {
    current = SPREADSHEET.insertSheet("Collection");
    current.appendRow(["Name", "Branch", "Semester", "Category", "Amount", "Mode", "Timestamp", "Collected By Phone", "Collected By", "Settled", "Settled At"]);
    return current;
  }
  let headers = headersOf(current);
  if (col(headers, ["Collected By Phone", "Volunteer Phone"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Collected By Phone");
  }
  headers = headersOf(current);
  if (col(headers, ["Collected By", "Volunteer Name"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Collected By");
  }
  headers = headersOf(current);
  if (col(headers, ["Settled", "Settlement Status"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settled");
  }
  headers = headersOf(current);
  if (col(headers, ["Settled At", "Settlement Date"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settled At");
  }
  headers = headersOf(current);
  if (col(headers, ["Settlement Amount"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settlement Amount");
  }
  headers = headersOf(current);
  if (col(headers, ["Settlement Date"]) < 0) {
    current.insertColumnAfter(current.getLastColumn());
    current.getRange(1, current.getLastColumn()).setValue("Settlement Date");
  }
  return current;
}

function saveExpense(data, updating) {
  const item = text(data.item, 200);
  const amount = amountOf(data.amount);
  const workerPhone = phoneOf(data.workerPhone);

  if (!item) return fail("Item / purpose zaroori hai.");
  if (amount === null) return fail("Amount ek valid number hona chahiye.");
  if (!workerPhone) return fail("Worker ka phone number exact 10 digits ka hona chahiye.");

  let fileUrl = text(data.fileUrl, 500);
  if (data.fileData && data.fileName) {
    fileUrl = saveDriveFile(data.fileData, data.fileName, data.fileMimeType, false, BILL_FOLDER_ID);
  } else if (updating && !fileUrl) {
    fileUrl = existingCellValue("Expenses", data.rowIndex, 6);
  }

  const values = [
    item,
    amount,
    text(data.status, 20) || "Pending",
    text(data.workerName, 120),
    workerPhone,
    fileUrl,
    "Admin",
    new Date()
  ];
  return updating ? updateRow("Expenses", data.rowIndex, values) : addRow("Expenses", values);
}

function addParticipant(data) {
  const name = text(data.name, 120);
  const phone = phoneOf(data.phone);
  if (!name) return fail("Naam zaroori hai.");
  if (!phone) return fail("Phone number exact 10 digits ka hona chahiye.");

  const current = ensureParticipantSheet() || sheet("Participants");
  if (!current) {
    const created = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Participants");
    created.appendRow([
      "Name", "Branch", "Semester", "Event", "Phone", "Type", "Group Members",
      "Timestamp", "Program Order", "Program Status", "Song URL"
    ]);
    return addParticipant(data);
  }
  normalizeParticipantOrder(current);
  const headers = headersOf(current);
  const orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  const existingOrders = current.getLastRow() >= 2
      ? current.getRange(2, orderIndex + 1, current.getLastRow() - 1, 1).getValues().flat().map(Number).filter(isFinite)
      : [];
  const nextOrder = (existingOrders.length ? Math.max.apply(null, existingOrders) : 0) + 1;
  const eventName = text(data.event, 60);
  let songUrl = "";
  if (eventName.toLowerCase() === "dance") {
    if (!data.songData || !data.songName) return fail("Dance participant ke liye MP3 song zaroori hai.");
    if (!/\.mp3$/i.test(String(data.songName))) return fail("Sirf MP3 song upload karein.");
    const songName = text(name.replace(/[\\/:*?"<>|]/g, "_") + ".mp3", 200);
    songUrl = saveDriveFile(data.songData, songName, "audio/mpeg", false, SONG_FOLDER_ID || MEDIA_FOLDER_ID, ALLOWED_SONG_TYPES, null);
  }
  const participantValues = Array(current.getLastColumn()).fill("");
  const setParticipantValue = (names, value) => {
    const index = col(headers, names);
    if (index >= 0) participantValues[index] = value;
  };
  setParticipantValue(["Name", "Participant Name"], name);
  setParticipantValue(["Branch"], text(data.branch, 40));
  setParticipantValue(["Semester", "Sem"], text(data.semester, 20));
  setParticipantValue(["Event"], eventName);
  setParticipantValue(["Phone", "Phone Number"], phone);
  setParticipantValue(["Type"], text(data.type, 20) || "Solo");
  setParticipantValue(["Group Members"], text(data.groupMembers, 500));
  setParticipantValue(["Timestamp", "Created At", "Date"], new Date());
  setParticipantValue(["Program Order", "Order", "Sequence"], nextOrder);
  setParticipantValue(["Program Status", "Status"], "Pending");
  setParticipantValue(["Song URL", "Dance Song", "Song"], songUrl);
  current.getRange(current.getLastRow() + 1, 1, 1, participantValues.length).setValues([participantValues]);
  const savedRow = current.getRange(current.getLastRow(), 1, 1, current.getLastColumn()).getValues()[0];
  const savedSongUrl = valueAt(savedRow, headers, ["Song URL", "Dance Song", "Song"]);
  if (eventName.toLowerCase() === "dance" && songUrl && String(savedSongUrl) !== String(songUrl)) {
    throw new Error("Participant save ho gaya, lekin Song URL column me save nahi hua.");
  }
  return ok({ songUrl: songUrl });
}

function completeParticipant(data) {
  const current = ensureParticipantSheet();
  const requestedPhone = phoneOf(data.phone);
  if (!current || !requestedPhone) return fail("Participant phone number required");

  const headers = headersOf(current);
  const phoneIndex = col(headers, ["Phone", "Phone Number"]);
  const orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  const statusIndex = col(headers, ["Program Status", "Status"]);
  if (phoneIndex < 0 || orderIndex < 0 || statusIndex < 0) return fail("Participant columns not found");

  const rows = current.getDataRange().getValues();
  const rowNumber = rows.findIndex((row, index) => index > 0 && phoneOf(row[phoneIndex]) === requestedPhone) + 1;
  if (rowNumber < 2) return fail("Participant not found");
  current.getRange(rowNumber, statusIndex + 1).setValue("Completed");

  const orderedRows = current.getDataRange().getValues().slice(1)
    .map((row, index) => ({
      row,
      sheetRow: index + 2,
      order: Number(row[orderIndex]) || index + 1,
      status: String(row[statusIndex] || "Pending")
    }))
    .sort((left, right) => (left.status === "Completed") - (right.status === "Completed") || left.order - right.order);
  orderedRows.forEach((item, index) => current.getRange(item.sheetRow, orderIndex + 1).setValue(index + 1));
  return ok();
}

function updateParticipantOrder(data) {
  const current = ensureParticipantSheet();
  const requestedPhone = phoneOf(data.phone);
  const direction = String(data.direction || "");
  if (!current || !requestedPhone || !["up", "down"].includes(direction)) return fail("Invalid participant order request");

  const headers = headersOf(current);
  const orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  const phoneIndex = col(headers, ["Phone", "Phone Number"]);
  const statusIndex = col(headers, ["Program Status", "Status"]);
  if (orderIndex < 0 || phoneIndex < 0) return fail("Participant order columns not found");

  const rows = current.getDataRange().getValues().slice(1)
    .map((row, index) => ({ row, sheetRow: index + 2, order: Number(row[orderIndex]) || index + 1, status: String(row[statusIndex] || "Pending") }))
    .sort((left, right) => (left.status === "Completed") - (right.status === "Completed") || left.order - right.order);
  const currentPosition = rows.findIndex(item => phoneOf(item.row[phoneIndex]) === requestedPhone);
  if (currentPosition < 0) return fail("Participant not found");
  const targetPosition = direction === "up" ? currentPosition - 1 : currentPosition + 1;
  if (targetPosition < 0 || targetPosition >= rows.length) return ok();

  [rows[currentPosition], rows[targetPosition]] = [rows[targetPosition], rows[currentPosition]];
  rows.forEach((item, index) => current.getRange(item.sheetRow, orderIndex + 1).setValue(index + 1));
  return ok();
}

function addTask(data) {
  const title = text(data.title, 200);
  const responsible = text(data.responsible, 120);
  if (!title) return fail("Task title zaroori hai.");
  if (!responsible) return fail("Responsible person ka naam zaroori hai.");
  const current = ensureTaskSheet();
  current.appendRow([
    title,
    text(data.type, 40),
    text(data.status, 20) || "Pending",
    text(data.time, 40),
    responsible,
    new Date()
  ]);
  return ok();
}

function updateTaskStatus(title) {
  const current = sheet("Tasks");
  if (!current) return fail("Tasks sheet not found");
  const rows = current.getDataRange().getValues();
  const wanted = text(title, 200);
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === wanted) {
      current.getRange(i + 1, 3).setValue("Completed");
      return ok();
    }
  }
  return fail("Task not found");
}

function addMeeting(data) {
  const decision = text(data.decision, 2000);
  if (!decision) return fail("Meeting decision zaroori hai.");
  return addRow("Meetings", [
    text(data.meetingDate, 40),
    decision,
    text(data.nextDate, 40),
    new Date()
  ]);
}

function uploadMedia(data) {
  const title = text(data.title, 200);
  if (!title) return fail("Media title zaroori hai.");
  const url = saveDriveFile(data.fileData, data.fileName, data.fileMimeType, false, MEDIA_FOLDER_ID, null, null);
  return addRow("Media", [title, url, text(data.mediaType, 20) || "Photo", new Date()]);
}

function authorizeDriveAccess() {
  const folder = DriveApp.getFolderById(SONG_FOLDER_ID);
  Logger.log("Drive access OK: " + folder.getName());
}

function saveDriveFile(base64, fileName, mimeType, imageOnly, folderId, allowedTypes, maxBytes) {
  const type = text(mimeType, 100) || "application/octet-stream";
  const allowed = allowedTypes || (imageOnly ? /^image\//i : ALLOWED_UPLOAD_TYPES);
  if (!allowed.test(type)) {
    throw new Error(imageOnly ? "Sirf QR image upload karein." : allowedTypes ? "Sirf MP3 song upload karein." : "Sirf image ya video upload ki ja sakti hai.");
  }

  const bytes = Utilities.base64Decode(String(base64 || ""));
  if (maxBytes !== null && bytes.length > (maxBytes || MAX_UPLOAD_BYTES)) {
    throw new Error("File 5 MB se chhoti honi chahiye.");
  }

  const blob = Utilities.newBlob(bytes, type, text(fileName, 200) || "upload");
  const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (error) {
    // Workspace policy may block public sharing, but the owner can still use the file URL.
    console.warn("Drive sharing update skipped: " + error);
  }
  return file.getUrl();
}

function checkStatus(phone) {
  const wanted = phoneOf(phone);
  if (!wanted) return json({ found: false });
  const current = ensureParticipantSheet();
  if (!current || current.getLastRow() < 2) return json({ found: false });
  const headers = headersOf(current);
  const rows = current.getDataRange().getValues().slice(1)
    .map((row, index) => ({ row, sheetRow: index + 2, order: Number(valueAt(row, headers, ["Program Order", "Order", "Sequence"])) || index + 1, status: String(valueAt(row, headers, ["Program Status", "Status"]) || "Pending") }))
    .sort((left, right) => (left.status === "Completed") - (right.status === "Completed") || left.order - right.order);
  const found = rows.find(item => phoneOf(valueAt(item.row, headers, ["Phone", "Phone Number"])) === wanted);
  if (!found) return json({ found: false });

  const row = found.row;
  const position = rows.indexOf(found) + 1;
  return json({
    found: true,
    event: valueAt(row, headers, ["Event"]),
    name: valueAt(row, headers, ["Name", "Participant Name"]),
    type: valueAt(row, headers, ["Type"]),
    groupMembers: valueAt(row, headers, ["Group Members"]),
    status: String(valueAt(row, headers, ["Program Status", "Status"]) || "Pending"),
    position: position,
    totalParticipants: rows.length,
    isNext: position === 1,
    nextProgram: rows[0] ? valueAt(rows[0].row, headers, ["Event"]) : ""
  });
}