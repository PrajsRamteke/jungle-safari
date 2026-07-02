/**
 * Jungle Safari Feedback Form — Google Apps Script
 *
 * SETUP INSTRUCTIONS:
 * 1. Create a new Google Sheet
 * 2. Extensions → Apps Script → paste this entire file
 * 3. Run `setupSheet` once (authorize when prompted)
 * 4. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL into index1.html → APPS_SCRIPT_URL
 */

var SHEET_NAME = 'Feedback Responses';

var HEADERS = [
  'Timestamp',
  'Visitor Name',
  'Address',
  'Phone / Mobile',
  'Safari Date',
  'Safari Time',
  'Vehicle / Guide',
  'Overall Experience',
  'Booking Process (1-5)',
  'Vehicle Condition (1-5)',
  'Guide / Driver (1-5)',
  'Safety Measures (1-5)',
  'Cleanliness & Amenities (1-5)',
  'Wildlife Sightings',
  'Liked Most',
  'Suggestions',
  'Would Recommend',
  'Additional Comments'
];

/**
 * One-time setup: creates sheet with headers and formatting.
 * Run manually from the Apps Script editor.
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  } else {
    sheet.clear();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange
    .setBackground('#1b3d2f')
    .setFontColor('#f7f3ea')
    .setFontWeight('bold')
    .setWrap(true);

  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  // Set reasonable column widths for long text fields
  sheet.setColumnWidth(2, 140);  // Visitor Name
  sheet.setColumnWidth(3, 200);  // Address
  sheet.setColumnWidth(15, 250); // Liked Most
  sheet.setColumnWidth(16, 250); // Suggestions
  sheet.setColumnWidth(18, 250); // Additional Comments

  SpreadsheetApp.getUi().alert(
    'Setup complete!',
    'Sheet "' + SHEET_NAME + '" is ready.\n\nNext: Deploy as Web App and paste the URL into index1.html.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Handles POST requests from the HTML form.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse(false, 'No data received.');
    }

    var data = JSON.parse(e.postData.contents);
    var validationError = validatePayload(data);

    if (validationError) {
      return jsonResponse(false, validationError);
    }

    var sheet = getOrCreateSheet();
    ensureHeaders(sheet);

    var row = [
      data.timestamp || new Date().toISOString(),
      sanitize(data.visitorName),
      sanitize(data.address),
      sanitize(data.phone),
      sanitize(data.safariDate),
      sanitize(data.safariTime),
      sanitize(data.vehicleGuide),
      sanitize(data.overallExperience),
      sanitize(data.bookingProcess),
      sanitize(data.vehicleCondition),
      sanitize(data.guideBehavior),
      sanitize(data.safetyMeasures),
      sanitize(data.cleanlinessAmenities),
      sanitize(data.wildlifeSightings),
      sanitize(data.likedMost),
      sanitize(data.suggestions),
      sanitize(data.recommend),
      sanitize(data.additionalComments)
    ];

    sheet.appendRow(row);

    return jsonResponse(true, 'Feedback saved successfully.');
  } catch (error) {
    Logger.log('doPost error: ' + error.toString());
    return jsonResponse(false, 'Server error: ' + error.message);
  }
}

/**
 * Optional: serve a simple test page when visiting the Web App URL directly.
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      message: 'Jungle Safari Feedback API is running. Use POST to submit form data.'
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ensureHeaders(sheet) {
  var firstRow = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];

  if (!firstRow[0] || firstRow[0] !== HEADERS[0]) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function validatePayload(data) {
  if (!data) return 'Invalid payload.';

  if (!data.phone || String(data.phone).trim().length < 10) {
    return 'Valid phone number is required.';
  }

  if (!data.address || String(data.address).trim().length < 5) {
    return 'Address is required.';
  }

  if (!data.safariDate) return 'Safari date is required.';
  if (!data.safariTime) return 'Safari time is required.';
  if (!data.vehicleGuide || String(data.vehicleGuide).trim().length === 0) {
    return 'Vehicle number or guide name is required.';
  }

  if (!data.overallExperience) return 'Overall experience rating is required.';
  if (!data.wildlifeSightings) return 'Wildlife sightings response is required.';
  if (!data.recommend) return 'Recommendation response is required.';

  var ratingFields = [
    'bookingProcess',
    'vehicleCondition',
    'guideBehavior',
    'safetyMeasures',
    'cleanlinessAmenities'
  ];

  for (var i = 0; i < ratingFields.length; i++) {
    var val = data[ratingFields[i]];
    var num = parseInt(val, 10);
    if (isNaN(num) || num < 1 || num > 5) {
      return 'All service ratings (1–5) are required.';
    }
  }

  var validOverall = ['Excellent', 'Good', 'Average', 'Poor'];
  if (validOverall.indexOf(data.overallExperience) === -1) {
    return 'Invalid overall experience value.';
  }

  return null;
}

function sanitize(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().substring(0, 2000);
}

function jsonResponse(success, message) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: success,
      message: message
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
