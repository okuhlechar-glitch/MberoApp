// ── Version ─────────────────────────────────────────────────────────
// Bump this every time you paste new code into Apps Script and redeploy.
// Visit the deployed URL in a browser (GET) to confirm which version is live.
const SCRIPT_VERSION = '2026-06-02-v7';

// ── Config ──────────────────────────────────────────────────────────
const QUOTE_SHEET_NAME = 'Quote Requests';
const CONTRACTOR_SHEET_NAME = 'Contractor Registrations';
const NOTIFICATION_EMAIL = 'smpwecare@gmail.com';

// Brevo API key — stored securely in Project Settings → Script Properties
const BREVO_API_KEY = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
const BREVO_SENDER_EMAIL = 'emergesites@gmail.com';
const BREVO_SENDER_NAME = 'SmP – WE CARE';

// ── Web app entry points ────────────────────────────────────────────

function doGet() {
  return jsonResponse({
    success: true,
    message: 'SMP web app is running.',
    version: SCRIPT_VERSION
  });
}

function doPost(e) {
  try {
    console.log('[doPost] START — version ' + SCRIPT_VERSION);
    console.log('[doPost] raw postData.type = ' + (e && e.postData ? e.postData.type : 'NONE'));
    console.log('[doPost] raw postData.length = ' + (e && e.postData && e.postData.contents ? e.postData.contents.length : 0));

    var payload = getPayload_(e);
    console.log('[doPost] parsed payload keys: ' + Object.keys(payload).join(', '));
    console.log('[doPost] payload.name = ' + (payload.name || '(empty)'));
    console.log('[doPost] payload.email = ' + (payload.email || '(empty)'));
    console.log('[doPost] payload.formType = ' + (payload.formType || '(empty, defaulting to quote)'));

    var formType = payload.formType === 'contractor' ? 'contractor' : 'quote';
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!spreadsheet) {
      throw new Error('Open this Apps Script from the target Google Sheet before deploying the web app.');
    }

    if (formType === 'quote') {
      saveQuote_(spreadsheet, payload);
      console.log('[doPost] sheet write PASS');
      sendQuoteEmail_(payload);
      console.log('[doPost] email PASS');
    } else {
      saveContractor_(spreadsheet, payload);
      console.log('[doPost] sheet write PASS');
      sendContractorEmail_(payload);
      console.log('[doPost] email PASS');
    }

    console.log('[doPost] END — success');
    return jsonResponse({ success: true, version: SCRIPT_VERSION });
  } catch (error) {
    console.error('[doPost] FAIL: ' + error.message);
    console.error('[doPost] stack: ' + error.stack);
    return jsonResponse({ success: false, message: error.message, version: SCRIPT_VERSION });
  }
}

// ── Photo upload ────────────────────────────────────────────────────

function savePhotoToDrive_(payload) {
  if (!payload.photoBase64) return '';
  console.log('[savePhotoToDrive_] Saving photo: ' + (payload.photoName || 'photo.jpg'));

  var folder;
  var folders = DriveApp.getFoldersByName('SMP Quote Photos');
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder('SMP Quote Photos');
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  }

  var blob = Utilities.newBlob(
    Utilities.base64Decode(payload.photoBase64),
    payload.photoMime || 'image/jpeg',
    payload.photoName || 'photo.jpg'
  );

  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  console.log('[savePhotoToDrive_] PASS — ' + file.getUrl());
  return file.getUrl();
}

// ── Sheet writes ────────────────────────────────────────────────────

function saveQuote_(spreadsheet, payload) {
  var photoLink = savePhotoToDrive_(payload);

  var headers = [
    'Submitted At', 'Name', 'Phone', 'Email', 'Service',
    'Location', 'Preferred Date', 'Details', 'Photo Link'
  ];

  var row = [
    payload.submittedAt || new Date().toISOString(),
    payload.name || '',
    payload.phone || '',
    payload.email || '',
    payload.service || '',
    payload.location || '',
    payload.preferredDate || '',
    payload.details || '',
    photoLink || payload.photoLink || ''
  ];

  appendRow_(spreadsheet, QUOTE_SHEET_NAME, headers, row);
  payload.photoLink = photoLink;
  console.log('[saveQuote_] PASS');
}

function saveContractor_(spreadsheet, payload) {
  var headers = [
    'Submitted At', 'Name', 'Phone', 'Email',
    'Area', 'Service', 'Experience', 'Notes'
  ];

  var row = [
    payload.submittedAt || new Date().toISOString(),
    payload.name || '',
    payload.phone || '',
    payload.email || '',
    payload.area || '',
    payload.service || '',
    payload.experience || '',
    payload.notes || ''
  ];

  appendRow_(spreadsheet, CONTRACTOR_SHEET_NAME, headers, row);
  console.log('[saveContractor_] PASS');
}

function appendRow_(spreadsheet, sheetName, headers, row) {
  var sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  sheet.appendRow(row);
}

// ── Brevo email ─────────────────────────────────────────────────────

function sendViaBrevo_(to, replyTo, subject, textContent) {
  console.log('[sendViaBrevo_] to=' + to + ', replyTo=' + replyTo + ', subject=' + subject);
  console.log('[sendViaBrevo_] textContent preview: ' + textContent.substring(0, 200));
  console.log('[sendViaBrevo_] API key present: ' + (BREVO_API_KEY ? 'YES (length ' + BREVO_API_KEY.length + ')' : 'NO — MISSING!'));

  if (!BREVO_API_KEY) {
    throw new Error('BREVO_API_KEY is not set. Go to Project Settings -> Script Properties and add it.');
  }

  var emailPayload = {
    sender: { name: BREVO_SENDER_NAME, email: BREVO_SENDER_EMAIL },
    to: [{ email: to }],
    replyTo: { email: replyTo || to },
    subject: subject,
    textContent: textContent
  };

  console.log('[sendViaBrevo_] Calling Brevo API...');
  var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'api-key': BREVO_API_KEY },
    payload: JSON.stringify(emailPayload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  var body = response.getContentText();
  console.log('[sendViaBrevo_] HTTP ' + code + ': ' + body);

  if (code < 200 || code >= 300) {
    console.error('[sendViaBrevo_] FAIL — Brevo HTTP ' + code + ': ' + body);
    throw new Error('Email send failed (Brevo ' + code + '): ' + body);
  }

  console.log('[sendViaBrevo_] PASS');
}

function sendQuoteEmail_(payload) {
  var subject = 'New SMP quote request';
  var body = [
    'A new quote request was submitted.',
    '',
    'Name: ' + (payload.name || '(not provided)'),
    'Phone: ' + (payload.phone || '(not provided)'),
    'Email: ' + (payload.email || '(not provided)'),
    'Service: ' + (payload.service || '(not provided)'),
    'Location: ' + (payload.location || '(not provided)'),
    'Preferred Date: ' + (payload.preferredDate || '(not provided)'),
    'Details: ' + (payload.details || '(not provided)'),
    'Photo Link: ' + (payload.photoLink || '(none)')
  ].join('\n');

  console.log('[sendQuoteEmail_] Sending email with body:\n' + body);
  sendViaBrevo_(NOTIFICATION_EMAIL, payload.email, subject, body);
}

function sendContractorEmail_(payload) {
  var subject = 'New SMP contractor registration';
  var body = [
    'A new contractor registration was submitted.',
    '',
    'Name: ' + (payload.name || '(not provided)'),
    'Phone: ' + (payload.phone || '(not provided)'),
    'Email: ' + (payload.email || '(not provided)'),
    'Area: ' + (payload.area || '(not provided)'),
    'Service: ' + (payload.service || '(not provided)'),
    'Experience: ' + (payload.experience || '(not provided)') + ' years',
    'Notes: ' + (payload.notes || '(not provided)')
  ].join('\n');

  console.log('[sendContractorEmail_] Sending email with body:\n' + body);
  sendViaBrevo_(NOTIFICATION_EMAIL, payload.email, subject, body);
}

// ── Payload parsing ─────────────────────────────────────────────────

function getPayload_(e) {
  var rawBody = e && e.postData && e.postData.contents ? e.postData.contents : '';

  if (rawBody) {
    console.log('[getPayload_] Raw body length: ' + rawBody.length);
    console.log('[getPayload_] Body preview: ' + rawBody.substring(0, 300));

    // Always try JSON first — the front-end sends JSON with text/plain content type
    try {
      var parsed = JSON.parse(rawBody);
      console.log('[getPayload_] JSON parse PASS — keys: ' + Object.keys(parsed).join(', '));
      return parsed;
    } catch (jsonErr) {
      console.log('[getPayload_] JSON parse failed: ' + jsonErr.message);
    }

    // Fallback: URL-encoded form data
    if (rawBody.indexOf('=') !== -1) {
      var formParsed = parseUrlEncodedBody_(rawBody);
      console.log('[getPayload_] URL-encoded parse — keys: ' + Object.keys(formParsed).join(', '));
      return formParsed;
    }

    throw new Error('Unable to parse request body');
  }

  var params = e && e.parameter ? e.parameter : {};
  console.log('[getPayload_] No body, using URL params: ' + Object.keys(params).join(', '));
  return params;
}

function parseUrlEncodedBody_(rawBody) {
  var result = {};
  rawBody.split('&').forEach(function(pair) {
    if (!pair) return;
    var parts = pair.split('=');
    var key = decodeURIComponent(parts[0].replace(/\+/g, ' '));
    var value = parts.length > 1 ? decodeURIComponent(parts.slice(1).join('=').replace(/\+/g, ' ')) : '';
    result[key] = value;
  });
  return result;
}

function jsonResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Test helpers (run from the editor, not from the web app) ────────

// Test 1: Send a plain test email via Brevo
function testSendBrevo() {
  console.log('[testSendBrevo] Version: ' + SCRIPT_VERSION);
  sendViaBrevo_(
    NOTIFICATION_EMAIL,
    BREVO_SENDER_EMAIL,
    'Brevo test — ' + SCRIPT_VERSION,
    'This is a test message from version ' + SCRIPT_VERSION + '.'
  );
  console.log('[testSendBrevo] PASS');
}

// Test 2: Simulate a full quote form submission (sheet + email)
// Run this from the editor to test the whole flow without needing the website.
function testQuoteSubmission() {
  console.log('[testQuoteSubmission] Version: ' + SCRIPT_VERSION);
  var fakePayload = {
    formType: 'quote',
    name: 'Test User',
    phone: '081 467 3054',
    email: NOTIFICATION_EMAIL,
    service: 'Painting & Finishes',
    location: 'Johannesburg',
    preferredDate: '2026-07-01',
    details: 'This is a test submission from the Apps Script editor.',
    submittedAt: new Date().toISOString()
  };

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  saveQuote_(spreadsheet, fakePayload);
  console.log('[testQuoteSubmission] Sheet write PASS');
  sendQuoteEmail_(fakePayload);
  console.log('[testQuoteSubmission] Email PASS — check your inbox for the full form data');
}

// Test 3: Simulate a full contractor form submission (sheet + email)
function testContractorSubmission() {
  console.log('[testContractorSubmission] Version: ' + SCRIPT_VERSION);
  var fakePayload = {
    formType: 'contractor',
    name: 'Test Contractor',
    phone: '082 555 1234',
    email: NOTIFICATION_EMAIL,
    area: 'Pretoria',
    service: 'Plumbing & Water Systems',
    experience: '5',
    notes: 'This is a test contractor submission from the Apps Script editor.',
    submittedAt: new Date().toISOString()
  };

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  saveContractor_(spreadsheet, fakePayload);
  console.log('[testContractorSubmission] Sheet write PASS');
  sendContractorEmail_(fakePayload);
  console.log('[testContractorSubmission] Email PASS — check your inbox for the full form data');
}
