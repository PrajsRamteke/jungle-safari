/**
 * 🌲 Jungle Safari Feedback Form — Google Apps Script backend
 *
 * SETUP (one time):
 * 1. Create a new Google Sheet (or open an existing one).
 * 2. Extensions → Apps Script → delete any code there and paste this whole file.
 * 3. Save (Cmd+S), then click Deploy → New deployment.
 * 4. Type: "Web app"
 *      - Description: Jungle Safari Feedback
 *      - Execute as: Me
 *      - Who has access: Anyone   ← IMPORTANT
 * 5. Click Deploy, authorize the permissions, and copy the Web App URL
 *    (ends with /exec).
 * 6. Paste that URL into SCRIPT_URL at the top of the <script> in index.html.
 *
 * Note: after ANY code change here, you must Deploy → Manage deployments →
 * ✏️ Edit → Version: "New version" → Deploy, otherwise the old code keeps running.
 */

var SHEET_NAME = "Feedback Responses";

var HEADERS = [
  "Timestamp",
  "Visitor Name",
  "Occupation",
  "Nationality",
  "Address",
  "Phone / Mobile",
  "Safari Date",
  "Safari Time",
  "Vehicle No. / Guide Name",
  "Overall Experience",
  "Booking Process (1-5)",
  "Vehicle Condition & Comfort (1-5)",
  "Guide / Driver Knowledge & Behavior (1-5)",
  "Safety Measures (1-5)",
  "Cleanliness & Amenities (1-5)",
  "Wildlife Sightings",
  "Liked Most",
  "Suggestions",
  "Would Recommend",
  "Additional Comments"
];

/** Handles form submissions (POST from index.html). */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000); // avoid clashing rows on simultaneous submissions

  try {
    var sheet = getOrCreateSheet_();
    var p = (e && e.parameter) ? e.parameter : {};

    // ----- Server-side validation (never trust only the client) -----
    var errors = [];

    if (!clean_(p.address))       errors.push("Address is required.");
    if (!clean_(p.vehicleGuide))  errors.push("Vehicle number / guide name is required.");
    if (!clean_(p.safariDate))    errors.push("Safari date is required.");
    if (!clean_(p.safariTime))    errors.push("Safari time is required.");

    var phone = clean_(p.phone).replace(/[\s-]/g, "");
    if (!/^(\+91)?\d{10}$/.test(phone)) errors.push("A valid 10-digit phone number is required.");

    if (["Excellent", "Good", "Average", "Poor"].indexOf(clean_(p.overallExperience)) === -1)
      errors.push("Overall experience rating is required.");

    var ratingKeys = ["bookingProcess", "vehicleCondition", "guideKnowledge", "safetyMeasures", "cleanliness"];
    for (var i = 0; i < ratingKeys.length; i++) {
      var v = clean_(p[ratingKeys[i]]);
      if (["1", "2", "3", "4", "5"].indexOf(v) === -1) {
        errors.push("All five service ratings (1-5) are required.");
        break;
      }
    }

    if (!clean_(p.wildlifeSighting)) errors.push("Wildlife sighting answer is required.");
    if (!clean_(p.recommend))        errors.push("Recommendation answer is required.");

    if (errors.length > 0) {
      return jsonResponse_({ status: "error", message: errors.join(" ") });
    }

    // ----- Append the row -----
    sheet.appendRow([
      new Date(),
      clean_(p.visitorName) || "Anonymous",
      clean_(p.occupation),
      clean_(p.nationality),
      clean_(p.address),
      "'" + phone, // leading apostrophe keeps leading zeros / + intact in Sheets
      clean_(p.safariDate),
      clean_(p.safariTime),
      clean_(p.vehicleGuide),
      clean_(p.overallExperience),
      Number(p.bookingProcess),
      Number(p.vehicleCondition),
      Number(p.guideKnowledge),
      Number(p.safetyMeasures),
      Number(p.cleanliness),
      clean_(p.wildlifeSighting),
      clean_(p.likedMost),
      clean_(p.suggestions),
      clean_(p.recommend),
      clean_(p.comments)
    ]);

    return jsonResponse_({ status: "success", message: "Feedback recorded. Thank you!" });

  } catch (err) {
    return jsonResponse_({ status: "error", message: "Server error: " + err.message });
  } finally {
    lock.releaseLock();
  }
}

/** Simple health check — open the /exec URL in a browser to verify deployment. */
function doGet() {
  return jsonResponse_({
    status: "ok",
    message: "🌲 Jungle Safari Feedback API is running. Submit the form via POST."
  });
}

/** Returns the response sheet, creating and formatting it on first use. */
function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
    headerRange
      .setFontWeight("bold")
      .setBackground("#2d6a4f")
      .setFontColor("#ffffff")
      .setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, HEADERS.length, 160);
  }

  return sheet;
}

/** Trims a value and guards against undefined/null. Also caps length. */
function clean_(value) {
  return String(value == null ? "" : value).trim().substring(0, 2000);
}

/** Builds a JSON text response. */
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * OPTIONAL: run this once from the Apps Script editor (Run ▶ setupSheet)
 * to create the formatted sheet immediately and grant permissions upfront.
 */
function setupSheet() {
  getOrCreateSheet_();
}
