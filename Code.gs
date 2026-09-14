/**
 * Fresher & Farewell Event Manager — Google Apps Script backend.
 *
 * SETUP (ek baar karna hai):
 *   1. Project Settings > Script Properties me `ADMIN_PIN` add karein.
 *      Admin PIN ab kabhi bhi HTML me nahi rakha jata.
 *   2. Deploy > New deployment > Web app
 *        Execute as      : Me
 *        Who has access  : Anyone
 *   3. Nayi /exec URL copy karke har HTML page ke CONFIG block me paste karein.
 *
 * SECURITY MODEL:
 *   - Login par server ek session token deta hai (6 ghante valid).
 *   - Har privileged action se pehle `authorize()` token verify karta hai.
 *   - Browser me `role` sirf UI dikhane ke liye hai; asli permission yahan check hoti hai.
 *   - Volunteer passwords `sha256$<salt>$<hash>` format me store hote hain.
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const MEDIA_FOLDER_ID = "";

const SESSION_TTL_SECONDS = 6 * 60 * 60; // CacheService ki maximum limit
const FAILED_LOGIN_DELAY_MS = 400;       // brute force ko dheema karta hai
const MAX_TEXT_LENGTH = 500;
const MAX_UPLOAD_BYTES = 5 * 1024 * 1024;
const ALLOWED_UPLOAD_TYPES = /^(image|video)\//i;

/**
 * Kaun sa action kis role se allowed hai.
 * Yahan list na hone ka matlab: action public hai (registration, feedback, read-only views).
 */
const ACTION_ROLES = {
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
  addExpense: ["Admin"],
  updateExpense: ["Admin"],
  deleteExpense: ["Admin"],
  addTask: ["Admin"],
  updateTaskStatus: ["Admin"],
  addMeeting: ["Admin"],
  uploadMedia: ["Admin"]
};

// ---------------------------------------------------------------------------
// Response helpers
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Input validation
// ---------------------------------------------------------------------------

/** Trim karke length cap lagata hai, taki koi sheet me bahut bada blob na daal sake. */
function text(value, limit) {
  return String(value == null ? "" : value).trim().slice(0, limit || MAX_TEXT_LENGTH);
}

/** Sirf 10-digit phone number accept karta hai, warna khali string. */
function phoneOf(value) {
  const digits = String(value == null ? "" : value).replace(/\D/g, "");
  return digits.length === 10 ? digits : "";
}

/** Valid non-negative number ya `null`. */
function amountOf(value) {
  const number = Number(value);
  return isFinite(number) && number >= 0 ? number : null;
}

/** Do strings ko compare karta hai bina timing hint diye. */
function safeEquals(left, right) {
  const a = String(left);
  const b = String(right);
  if (a.length !== b.length) return false;
  let difference = 0;
  for (let i = 0; i < a.length; i++) difference |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return difference === 0;
}

// ---------------------------------------------------------------------------
// Password hashing
// ---------------------------------------------------------------------------

function sha256Hex(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(byte => ("0" + (byte & 0xff).toString(16)).slice(-2))
    .join("");
}

function makePasswordHash(password) {
  const salt = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  return "sha256$" + salt + "$" + sha256Hex(salt + ":" + password);
}

function isLegacyPassword(stored) {
  return String(stored || "").indexOf("sha256$") !== 0;
}

/**
 * Hash ya purana plaintext password — dono ke saath kaam karta hai.
 * Purane accounts login ke waqt apne aap hash me upgrade ho jate hain.
 */
function passwordMatches(stored, supplied) {
  const value = String(stored || "");
  if (!value) return false;
  if (isLegacyPassword(value)) return safeEquals(value, String(supplied));
  const parts = value.split("$");
  return parts.length === 3 && safeEquals(parts[2], sha256Hex(parts[1] + ":" + supplied));
}

// ---------------------------------------------------------------------------
// Sessions
// ---------------------------------------------------------------------------

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

/**
 * Action ke liye permission check karta hai.
 * Return: { allowed: true, session } ya { allowed: false, response }.
 */
function authorize(action, token) {
  const roles = ACTION_ROLES[action];
  const session = readSession(token);
  if (!roles) return { allowed: true, session: session };
  if (!session || roles.indexOf(session.role) < 0) {
    return { allowed: false, response: fail("Session expire ho gaya hai. Dobara login karein.") };
  }
  return { allowed: true, session: session };
}

// ---------------------------------------------------------------------------
// Sheet helpers
// ---------------------------------------------------------------------------

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
    const wanted = name.toLowerCase().replace(/\s/g, "");
    const index = headers.findIndex(value => String(value).toLowerCase().replace(/\s/g, "") === wanted);
    if (index >= 0) return index;
  }
  return -1;
}

function valueAt(row, headers, names) {
  const index = col(headers, names);
  return index >= 0 ? row[index] : "";
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

/** Update ke dauran kisi column ki maujooda value padhta hai (1-based column). */
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

// ---------------------------------------------------------------------------
// Request routing
// ---------------------------------------------------------------------------

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

    // Login/logout token se pehle chalte hain, isliye inhe alag handle karte hain.
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

// ---------------------------------------------------------------------------
// Authentication
// ---------------------------------------------------------------------------

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

  if (passwordIndex < 0 || !passwordMatches(row[passwordIndex], supplied) || status !== "approved") {
    Utilities.sleep(FAILED_LOGIN_DELAY_MS);
    return fail("Invalid credentials or approval pending");
  }

  // Purana plaintext password pehle successful login par hash me badal dete hain.
  if (isLegacyPassword(row[passwordIndex])) {
    found.current.getRange(found.row, passwordIndex + 1).setValue(makePasswordHash(supplied));
  }

  const branch = valueAt(row, found.headers, ["Branch"]);
  const semester = valueAt(row, found.headers, ["Semester"]);
  const token = createSession({
    role: "Volunteer",
    phone: phoneOf(phone),
    branch: String(branch),
    semester: String(semester)
  });
  return ok({ role: "Volunteer", branch: branch, semester: semester, token: token });
}

// ---------------------------------------------------------------------------
// Volunteers
// ---------------------------------------------------------------------------

function volunteerSheet() {
  return sheet("Volunteers", ["Volunteer", "Volunteers Data"]);
}

/** Volunteer records — password column kabhi bahar nahi bhejta. */
function getVolunteers(status) {
  const current = volunteerSheet();
  if (!current || current.getLastRow() < 2) return ok({ volunteers: [] });
  const headers = headersOf(current);
  const statusIndex = col(headers, ["Status", "Approval Status"]);
  const volunteers = current.getDataRange().getValues().slice(1).map((row, offset) => ({
    name: valueAt(row, headers, ["Name", "Volunteer Name"]),
    branch: valueAt(row, headers, ["Branch"]),
    semester: valueAt(row, headers, ["Semester"]),
    phone: valueAt(row, headers, ["Phone", "Phone Number"]),
    status: statusIndex >= 0 ? row[statusIndex] : "",
    rowIndex: offset + 2
  })).filter(item => String(item.status).trim().toLowerCase() === status);
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
    current.appendRow(["Name", "Branch", "Semester", "Phone", "Password", "Status", "Timestamp"]);
  }
  if (findVolunteer(phone)) return fail("Ye phone number pehle se registered hai.");

  current.appendRow([
    name,
    text(data.branch, 40),
    text(data.semester, 20),
    phone,
    makePasswordHash(password),
    "Pending",
    new Date()
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

// ---------------------------------------------------------------------------
// Dashboard data
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Feedback
// ---------------------------------------------------------------------------

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
    current = SPREADSHEET.insertSheet("Feedback");
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

// ---------------------------------------------------------------------------
// Collection
// ---------------------------------------------------------------------------

/**
 * Volunteer ke liye branch/semester session se liye jate hain, form se nahi —
 * taki koi request badal kar dusri branch ki entry na kar sake.
 */
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

  const values = [name, branch, semester, category, amount, text(data.mode, 20) || "Cash"];
  return updating ? updateRow("Collection", data.rowIndex, values) : addRow("Collection", values.concat([new Date()]));
}

// ---------------------------------------------------------------------------
// Expenses
// ---------------------------------------------------------------------------

function saveExpense(data, updating) {
  const item = text(data.item, 200);
  const amount = amountOf(data.amount);
  const workerPhone = phoneOf(data.workerPhone);

  if (!item) return fail("Item / purpose zaroori hai.");
  if (amount === null) return fail("Amount ek valid number hona chahiye.");
  if (!workerPhone) return fail("Worker ka phone number exact 10 digits ka hona chahiye.");

  // Nayi file di gayi ho to upload karte hain; warna edit karte waqt purana
  // bill link waise ka waisa rehta hai (pehle wo delete ho jata tha).
  let fileUrl = text(data.fileUrl, 500);
  if (data.fileData && data.fileName) {
    fileUrl = saveDriveFile(data.fileData, data.fileName, data.fileMimeType);
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

// ---------------------------------------------------------------------------
// Participants, tasks and meetings
// ---------------------------------------------------------------------------

function addParticipant(data) {
  const name = text(data.name, 120);
  const phone = phoneOf(data.phone);
  if (!name) return fail("Naam zaroori hai.");
  if (!phone) return fail("Phone number exact 10 digits ka hona chahiye.");

  const current = ensureParticipantSheet() || sheet("Participants");
  if (!current) {
    const created = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Participants");
    created.appendRow(["Name", "Branch", "Semester", "Event", "Phone", "Type", "Group Members", "Timestamp", "Program Order"]);
    return addParticipant(data);
  }
  normalizeParticipantOrder(current);
  const headers = headersOf(current);
  const orderIndex = col(headers, ["Program Order", "Order", "Sequence"]);
  const existingOrders = current.getLastRow() >= 2
      ? current.getRange(2, orderIndex + 1, current.getLastRow() - 1, 1).getValues().flat().map(Number).filter(isFinite)
      : [];
  const nextOrder = (existingOrders.length ? Math.max.apply(null, existingOrders) : 0) + 1;
  current.appendRow([
    name,
    text(data.branch, 40),
    text(data.semester, 20),
    text(data.event, 60),
    phone,
    text(data.type, 20) || "Solo",
    text(data.groupMembers, 500),
    new Date(),
    nextOrder,
    "Pending"
  ]);
  return ok();
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

// ---------------------------------------------------------------------------
// Media uploads
// ---------------------------------------------------------------------------

function uploadMedia(data) {
  const title = text(data.title, 200);
  if (!title) return fail("Media title zaroori hai.");
  const url = saveDriveFile(data.fileData, data.fileName, data.fileMimeType);
  return addRow("Media", [title, url, text(data.mediaType, 20) || "Photo", new Date()]);
}

/**
 * Base64 file ko Drive me save karta hai.
 * Note: file "anyone with the link" par share hoti hai, kyunki gallery aur bill
 * links public page par dikhte hain. Isliye private documents yahan upload na karein.
 */
function saveDriveFile(base64, fileName, mimeType) {
  const type = text(mimeType, 100) || "application/octet-stream";
  if (!ALLOWED_UPLOAD_TYPES.test(type)) throw new Error("Sirf image ya video upload ki ja sakti hai.");

  const bytes = Utilities.base64Decode(String(base64 || ""));
  if (bytes.length > MAX_UPLOAD_BYTES) throw new Error("File 5 MB se chhoti honi chahiye.");

  const blob = Utilities.newBlob(bytes, type, text(fileName, 200) || "upload");
  const folder = MEDIA_FOLDER_ID ? DriveApp.getFolderById(MEDIA_FOLDER_ID) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// ---------------------------------------------------------------------------
// Participant status lookup
// ---------------------------------------------------------------------------

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
