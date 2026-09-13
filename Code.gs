const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const MEDIA_FOLDER_ID = "";

function json(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name, aliases) {
  const names = [name].concat(aliases || []);
  for (const item of names) {
    const found = SPREADSHEET.getSheetByName(item);
    if (found) return found;
  }
  return null;
}

function rowsFrom(name, aliases) {
  const current = sheet(name, aliases);
  return current ? current.getDataRange().getValues() : [];
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

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";
    switch (action) {
      case "getData": return getData();
      case "getMedia": return getMedia();
      case "getApprovedVolunteers": return getVolunteers("approved");
      case "getPendingVolunteers": return getVolunteers("pending");
      case "getFeedback": return getFeedback();
      case "verifyVolunteer": return verifyVolunteer(e.parameter.phone, e.parameter.pass);
      case "checkStatus": return checkStatus(e.parameter.phone);
      default: return json({ status: "error", message: "Invalid GET action" });
    }
  } catch (error) {
    return json({ status: "error", message: error.toString() });
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    switch (data.action) {
      case "registerVolunteer": return registerVolunteer(data);
      case "approveVolunteer": return approveVolunteer(data);
      case "updateVolunteer": return updateVolunteer(data);
      case "deleteVolunteer": return deleteVolunteer(data);
      case "submitFeedback": return submitFeedback(data);
      case "deleteFeedback": return deleteFeedback(data);
      case "addCollection": return addRow("Collection", [data.name, data.branch, data.semester, data.category, data.amount, data.mode, new Date()]);
      case "updateCollection": return updateRow("Collection", data.rowIndex, [data.name, data.branch, data.semester, data.category, data.amount, data.mode]);
      case "deleteCollection": return deleteRow("Collection", data.rowIndex);
      case "addExpense": return saveExpense(data, false);
      case "updateExpense": return saveExpense(data, true);
      case "deleteExpense": return deleteRow("Expenses", data.rowIndex);
      case "addParticipant": return addRow("Participants", [data.name, data.branch, data.semester, data.event, data.phone, data.type, data.groupMembers, new Date()]);
      case "addTask": return addRow("Tasks", [data.title, data.type, data.status, data.time, new Date()]);
      case "updateTaskStatus": return updateTaskStatus(data.title);
      case "addMeeting": return addRow("Meetings", [data.meetingDate, data.decision, data.nextDate, new Date()]);
      case "uploadMedia": return uploadMedia(data);
      default: return json({ status: "error", message: "Invalid POST action" });
    }
  } catch (error) {
    return json({ status: "error", message: error.toString() });
  }
}

function getData() {
  const result = {};
  ["Collection", "Expenses", "Participants", "Tasks", "Meetings", "Settings"].forEach(name => {
    const current = sheet(name);
    if (current) result[name] = current.getDataRange().getValues();
  });
  return json(result);
}

function getMedia() {
  const current = sheet("Media");
  if (!current || current.getLastRow() < 2) return json({ status: "success", media: [] });
  const rows = current.getDataRange().getValues();
  const media = rows.slice(1).filter(row => row[1]).map(row => ({
    title: row[0] || "Event media",
    fileUrl: row[1],
    type: row[2] || "Photo"
  }));
  return json({ status: "success", media });
}

function volunteerSheet() {
  return sheet("Volunteers", ["Volunteer", "Volunteers Data"]);
}

function getVolunteers(status) {
  const current = volunteerSheet();
  if (!current || current.getLastRow() < 2) return json({ status: "success", volunteers: [] });
  const headers = headersOf(current);
  const rows = current.getDataRange().getValues();
  const statusIndex = col(headers, ["Status", "Approval Status"]);
  const volunteers = rows.slice(1).map((row, offset) => ({
    name: valueAt(row, headers, ["Name", "Volunteer Name"]),
    branch: valueAt(row, headers, ["Branch"]),
    semester: valueAt(row, headers, ["Semester"]),
    phone: valueAt(row, headers, ["Phone", "Phone Number"]),
    status: statusIndex >= 0 ? row[statusIndex] : "",
    rowIndex: offset + 2
  })).filter(item => String(item.status).trim().toLowerCase() === status);
  return json({ status: "success", volunteers });
}

function findVolunteer(phone) {
  const current = volunteerSheet();
  if (!current || current.getLastRow() < 2) return null;
  const headers = headersOf(current);
  const phoneIndex = col(headers, ["Phone", "Phone Number"]);
  if (phoneIndex < 0) return null;
  const rows = current.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][phoneIndex]).trim() === String(phone).trim()) return { current, headers, row: i + 1 };
  }
  return null;
}

function registerVolunteer(data) {
  let current = volunteerSheet();
  if (!current) {
    current = SPREADSHEET.insertSheet("Volunteers");
    current.appendRow(["Name", "Branch", "Semester", "Phone", "Password", "Status", "Timestamp"]);
  }
  current.appendRow([data.name, data.branch, data.semester, data.phone, data.password, "Pending", new Date()]);
  return json({ status: "success" });
}

function verifyVolunteer(phone, password) {
  const found = findVolunteer(phone);
  if (!found) return json({ status: "error", message: "Invalid credentials" });
  const rows = found.current.getDataRange().getValues();
  const row = rows[found.row - 1];
  const passwordIndex = col(found.headers, ["Password"]);
  const statusIndex = col(found.headers, ["Status", "Approval Status"]);
  if (passwordIndex < 0 || String(row[passwordIndex]) !== String(password) || String(row[statusIndex]).trim().toLowerCase() !== "approved") {
    return json({ status: "error", message: "Invalid credentials or approval pending" });
  }
  return json({ status: "success", branch: valueAt(row, found.headers, ["Branch"]), semester: valueAt(row, found.headers, ["Semester"]) });
}

function approveVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return json({ status: "error", message: "Volunteer not found" });
  const index = col(found.headers, ["Status", "Approval Status"]);
  if (index < 0) return json({ status: "error", message: "Status column not found" });
  found.current.getRange(found.row, index + 1).setValue("Approved");
  return json({ status: "success" });
}

function updateVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return json({ status: "error", message: "Volunteer not found" });
  [["Name", data.name], ["Volunteer Name", data.name], ["Branch", data.branch], ["Semester", data.semester]].forEach(pair => {
    const index = col(found.headers, [pair[0]]);
    if (index >= 0) found.current.getRange(found.row, index + 1).setValue(pair[1]);
  });
  return json({ status: "success" });
}

function deleteVolunteer(data) {
  const found = findVolunteer(data.phone);
  if (!found) return json({ status: "error", message: "Volunteer not found" });
  found.current.deleteRow(found.row);
  return json({ status: "success" });
}

function getFeedback() {
  const rows = dataRows("Feedback");
  const feedbacks = rows.map((row, index) => ({ rowIndex: index + 2, name: row[0] || "", message: row[1] || "", timestamp: row[2] || "" }));
  return json({ status: "success", feedbacks });
}

function submitFeedback(data) {
  let current = sheet("Feedback");
  if (!current) {
    current = SPREADSHEET.insertSheet("Feedback");
    current.appendRow(["Name", "Message", "Timestamp"]);
  }
  current.appendRow([data.name, data.message, new Date()]);
  return json({ status: "success" });
}

function deleteFeedback(data) {
  const current = sheet("Feedback");
  if (!current) return json({ status: "error", message: "Feedback sheet not found" });
  (data.rows || []).map(Number).sort((a, b) => b - a).forEach(row => {
    if (row > 1 && row <= current.getLastRow()) current.deleteRow(row);
  });
  return json({ status: "success" });
}

function checkStatus(phone) {
  const rows = dataRows("Participants");
  const found = rows.find(row => String(row[4]).trim() === String(phone).trim());
  return found ? json({ found: true, event: found[3], name: found[0] }) : json({ found: false });
}

function addRow(name, values) {
  let current = sheet(name);
  if (!current) current = SPREADSHEET.insertSheet(name);
  current.appendRow(values);
  return json({ status: "success" });
}

function updateRow(name, rowIndex, values) {
  const current = sheet(name);
  if (!current || !rowIndex || rowIndex > current.getLastRow()) return json({ status: "error", message: "Row not found" });
  current.getRange(Number(rowIndex) + 1, 1, 1, values.length).setValues([values]);
  return json({ status: "success" });
}

function deleteRow(name, rowIndex) {
  const current = sheet(name);
  const sheetRow = Number(rowIndex) + 1;
  if (!current || sheetRow <= 1 || sheetRow > current.getLastRow()) return json({ status: "error", message: "Row not found" });
  current.deleteRow(sheetRow);
  return json({ status: "success" });
}

function saveExpense(data, updating) {
  let fileUrl = data.fileUrl || "";
  if (data.fileData && data.fileName) fileUrl = saveDriveFile(data.fileData, data.fileName, data.fileMimeType);
  const values = [data.item, data.amount, data.status, data.workerName, data.workerPhone, fileUrl, data.addedBy || "Admin", new Date()];
  return updating ? updateRow("Expenses", data.rowIndex, values) : addRow("Expenses", values);
}

function updateTaskStatus(title) {
  const current = sheet("Tasks");
  if (!current) return json({ status: "error", message: "Tasks sheet not found" });
  const rows = current.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(title)) {
      current.getRange(i + 1, 3).setValue("Completed");
      return json({ status: "success" });
    }
  }
  return json({ status: "error", message: "Task not found" });
}

function uploadMedia(data) {
  const url = saveDriveFile(data.fileData, data.fileName, data.fileMimeType || "application/octet-stream");
  return addRow("Media", [data.title, url, data.mediaType || "Photo", new Date()]);
}

function saveDriveFile(base64, fileName, mimeType) {
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, fileName);
  const folder = MEDIA_FOLDER_ID ? DriveApp.getFolderById(MEDIA_FOLDER_ID) : DriveApp.getRootFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function getMediaData() {
  return getMedia();
}
