/**
 * Webinar Registration - Google Apps Script backend
 *
 * What it does:
 *   1. Receives form data (parent + student details) from index.html
 *   2. Appends a row to the "Registrations" sheet
 *   3. Reads the Google Meet link from the "Config" sheet (cell B1)
 *      and returns it so the form can redirect the user.
 *
 * The Meet link lives in a SHEET CELL, so any non-coder can update it
 * by simply editing that cell -- no code changes needed.
 *
 * Setup steps are in README.md.
 */

// Google Sheet ID to use for registration data
var SPREADSHEET_ID = "1EpRMjMNY2195-Cs4LevPTdQw1Ht3JBdyKZw6ItZ9CDw";

// Name of the tab where registrations are stored
var REG_SHEET   = "Registrations";
// Name of the tab that holds the editable Meet link
var CONFIG_SHEET = "Config";
// Cell in the Config tab that holds the Meet link
var MEET_CELL   = "B1";

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    var ss = getSpreadsheet();

    // --- 1. Save the registration ---
    var sheet = ss.getSheetByName(REG_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(REG_SHEET);
      sheet.appendRow([
        "Timestamp", "Parent Name", "Parent Number",
        "Student Name", "Student Number", "City"
      ]);
    }

    // Skip insert if this phone number is already registered
    var existing = sheet.getDataRange().getValues().slice(1);
    var isDuplicate = existing.some(function(row) {
      return String(row[2]).trim() === String(data.parentNumber || "").trim();
    });
    if (!isDuplicate) {
      sheet.appendRow([
        new Date(),
        data.parentName    || "",
        data.parentNumber  || "",
        data.studentName   || "",
        data.studentNumber || "",
        data.city          || ""
      ]);
    }

    // --- 2. Read the (editable) Meet link ---
    var meetLink = getMeetLink(ss);

    return json({ status: "success", meetLink: meetLink });

  } catch (err) {
    return json({ status: "error", message: String(err) });
  }
}

/** Reads the Meet link from Config!B1, creating the Config tab if missing. */
function getMeetLink(ss) {
  var config = ss.getSheetByName(CONFIG_SHEET);
  if (!config) {
    config = ss.insertSheet(CONFIG_SHEET);
    config.getRange("A1").setValue("Google Meet Link:");
    config.getRange(MEET_CELL).setValue("https://meet.google.com/your-code");
  }
  return config.getRange(MEET_CELL).getValue();
}

/** Returns the configured spreadsheet by ID. */
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

/** Optional: lets you test by opening the web app URL in a browser. */
function doGet() {
  var ss = getSpreadsheet();
  return json({ status: "ok", meetLink: getMeetLink(ss) });
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
