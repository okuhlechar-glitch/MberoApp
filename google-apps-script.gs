const QUOTE_SHEET_NAME = 'Quote Requests';
const CONTRACTOR_SHEET_NAME = 'Contractor Registrations';
const NOTIFICATION_EMAIL = 'smpwecare@gmail.com';

// Brevo API key — stored securely in Script Properties
const BREVO_API_KEY = PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY');
const BREVO_SENDER_EMAIL = 'emergesites@gmail.com';
const BREVO_SENDER_NAME = 'SmP – WE CARE';

function myFunction() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

function doGet() {
  console.log('[doGet] Health check called');
  return jsonResponse({ success: true, message: 'SMP web app is running.' });
}

function doPost(e) {
  console.log('[doPost] ====== START ======');
  try {
    console.log('[doPost] Parsing payload...');
    const payload = getPayload_(e);
    const formType = payload.formType === 'contractor' ? 'contractor' : 'quote';
    console.log('[doPost] Form type: ' + formType);
    console.log('[doPost] Payload keys: ' + Object.keys(payload).join(', '));

    console.log('[doPost] Getting active spreadsheet...');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (!spreadsheet) {
      console.error('[doPost] FAIL — No active spreadsheet found');
      throw new Error('Open this Apps Script from the target Google Sheet before deploying the web app.');
    }
    console.log('[doPost] PASS — Spreadsheet found: ' + spreadsheet.getName());

    if (formType === 'quote') {
      console.log('[doPost] Processing quote submission...');
      saveQuote_(spreadsheet, payload);
      console.log('[doPost] PASS — Quote saved to sheet');
      sendQuoteEmail_(payload);
      console.log('[doPost] PASS — Quote email sent');
    } else {
      console.log('[doPost] Processing contractor registration...');
      saveContractor_(spreadsheet, payload);
      console.log('[doPost] PASS — Contractor saved to sheet');
      sendContractorEmail_(payload);
      console.log('[doPost] PASS — Contractor email sent');
    }

    console.log('[doPost] ====== SUCCESS ======');
    return jsonResponse({ success: true });
  } catch (error) {
    console.error('[doPost] ====== FAILED ====== ' + error.message);
    console.error('[doPost] Stack: ' + error.stack);
    return jsonResponse({ success: false, message: error.message });
  }
}

function savePhotoToDrive_(payload) {
  console.log('[savePhotoToDrive_] Called');
  if (!payload.photoBase64) {
    console.log('[savePhotoToDrive_] No photo attached — skipping');
    return '';
  }

  console.log('[savePhotoToDrive_] Photo detected: ' + (payload.photoName || 'unnamed') + ' (' + (payload.photoMime || 'unknown mime') + ')');
  console.log('[savePhotoToDrive_] Base64 length: ' + payload.photoBase64.length);

  try {
    console.log('[savePhotoToDrive_] Looking for Drive folder "SMP Quote Photos"...');
    var folder;
    var folders = DriveApp.getFoldersByName('SMP Quote Photos');
    if (folders.hasNext()) {
      folder = folders.next();
      console.log('[savePhotoToDrive_] PASS — Found existing folder');
    } else {
      folder = DriveApp.createFolder('SMP Quote Photos');
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      console.log('[savePhotoToDrive_] PASS — Created new folder');
    }

    console.log('[savePhotoToDrive_] Decoding base64 and creating blob...');
    var blob = Utilities.newBlob(
      Utilities.base64Decode(payload.photoBase64),
      payload.photoMime || 'image/jpeg',
      payload.photoName || 'photo.jpg'
    );
    console.log('[savePhotoToDrive_] PASS — Blob created');

    console.log('[savePhotoToDrive_] Uploading file to Drive...');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var url = file.getUrl();
    console.log('[savePhotoToDrive_] PASS — File uploaded: ' + url);
    return url;
  } catch (err) {
    console.error('[savePhotoToDrive_] FAIL — ' + err.message);
    throw err;
  }
}

function saveQuote_(spreadsheet, payload) {
  console.log('[saveQuote_] Called');

  try {
    var photoLink = savePhotoToDrive_(payload);

    const headers = [
      'Submitted At',
      'Name',
      'Phone',
      'Email',
      'Service',
      'Location',
      'Preferred Date',
      'Details',
      'Photo Link'
    ];

    const row = [
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

    console.log('[saveQuote_] Writing row to sheet "' + QUOTE_SHEET_NAME + '"...');
    appendRow_(spreadsheet, QUOTE_SHEET_NAME, headers, row);
    console.log('[saveQuote_] PASS — Row written');

    payload.photoLink = photoLink;
  } catch (err) {
    console.error('[saveQuote_] FAIL — ' + err.message);
    throw err;
  }
}

function saveContractor_(spreadsheet, payload) {
  console.log('[saveContractor_] Called');

  try {
    const headers = [
      'Submitted At',
      'Name',
      'Phone',
      'Email',
      'Area',
      'Service',
      'Experience',
      'Notes'
    ];

    const row = [
      payload.submittedAt || new Date().toISOString(),
      payload.name || '',
      payload.phone || '',
      payload.email || '',
      payload.area || '',
      payload.service || '',
      payload.experience || '',
      payload.notes || ''
    ];

    console.log('[saveContractor_] Writing row to sheet "' + CONTRACTOR_SHEET_NAME + '"...');
    appendRow_(spreadsheet, CONTRACTOR_SHEET_NAME, headers, row);
    console.log('[saveContractor_] PASS — Row written');
  } catch (err) {
    console.error('[saveContractor_] FAIL — ' + err.message);
    throw err;
  }
}

function appendRow_(spreadsheet, sheetName, headers, row) {
  console.log('[appendRow_] Called for sheet: ' + sheetName);

  try {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      console.log('[appendRow_] Sheet not found — creating "' + sheetName + '"...');
      sheet = spreadsheet.insertSheet(sheetName);
      console.log('[appendRow_] PASS — Sheet created');
    } else {
      console.log('[appendRow_] PASS — Sheet found, last row: ' + sheet.getLastRow());
    }

    if (sheet.getLastRow() === 0) {
      console.log('[appendRow_] Writing headers...');
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      console.log('[appendRow_] PASS — Headers written');
    }

    sheet.appendRow(row);
    console.log('[appendRow_] PASS — Data row appended');
  } catch (err) {
    console.error('[appendRow_] FAIL — ' + err.message);
    throw err;
  }
}

function sendViaBrevo_(to, replyTo, subject, textContent) {
  console.log('[sendViaBrevo_] Called — sending to: ' + to);
  console.log('[sendViaBrevo_] Subject: ' + subject);
  console.log('[sendViaBrevo_] API key present: ' + (BREVO_API_KEY ? 'YES (' + BREVO_API_KEY.substring(0, 8) + '...)' : 'NO — MISSING!'));

  if (!BREVO_API_KEY) {
    console.error('[sendViaBrevo_] FAIL — BREVO_API_KEY is not set in Script Properties');
    throw new Error('BREVO_API_KEY is not set. Go to Project Settings → Script Properties and add it.');
  }

  try {
    var emailPayload = {
      sender: { name: BREVO_SENDER_NAME, email: BREVO_SENDER_EMAIL },
      to: [{ email: to }],
      replyTo: { email: replyTo || to },
      subject: subject,
      textContent: textContent
    };

    console.log('[sendViaBrevo_] Sending request to Brevo API...');
    var response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'api-key': BREVO_API_KEY },
      payload: JSON.stringify(emailPayload),
      muteHttpExceptions: true
    });

    var code = response.getResponseCode();
    var body = response.getContentText();
    console.log('[sendViaBrevo_] Response code: ' + code);
    console.log('[sendViaBrevo_] Response body: ' + body);

    if (code < 200 || code >= 300) {
      console.error('[sendViaBrevo_] FAIL — Brevo returned HTTP ' + code + ': ' + body);
      throw new Error('Email send failed (Brevo HTTP ' + code + '): ' + body);
    }

    console.log('[sendViaBrevo_] PASS — Email sent successfully');
  } catch (err) {
    console.error('[sendViaBrevo_] FAIL — ' + err.message);
    throw err;
  }
}

function sendQuoteEmail_(payload) {
  console.log('[sendQuoteEmail_] Called');
  const subject = 'New SmP – WE CARE quote request';
  const body = [
    'A new quote request was submitted.',
    '',
    'Name: ' + (payload.name || ''),
    'Phone: ' + (payload.phone || ''),
    'Email: ' + (payload.email || ''),
    'Service: ' + (payload.service || ''),
    'Location: ' + (payload.location || ''),
    'Preferred Date: ' + (payload.preferredDate || ''),
    'Details: ' + (payload.details || ''),
    'Photo Link: ' + (payload.photoLink || '')
  ].join('\n');

  sendViaBrevo_(NOTIFICATION_EMAIL, payload.email, subject, body);
  console.log('[sendQuoteEmail_] PASS — Done');
}

function sendContractorEmail_(payload) {
  console.log('[sendContractorEmail_] Called');
  const subject = 'New SmP – WE CARE contractor registration';
  const body = [
    'A new contractor registration was submitted.',
    '',
    'Name: ' + (payload.name || ''),
    'Phone: ' + (payload.phone || ''),
    'Email: ' + (payload.email || ''),
    'Area: ' + (payload.area || ''),
    'Service: ' + (payload.service || ''),
    'Experience: ' + (payload.experience || '') + ' years',
    'Notes: ' + (payload.notes || '')
  ].join('\n');

  sendViaBrevo_(NOTIFICATION_EMAIL, payload.email, subject, body);
  console.log('[sendContractorEmail_] PASS — Done');
}

function getPayload_(e) {
  console.log('[getPayload_] Called');

  try {
    var hasPostData = e && e.postData && e.postData.contents;
    console.log('[getPayload_] Has postData.contents: ' + !!hasPostData);

    if (hasPostData) {
      var rawBody = e.postData.contents;
      console.log('[getPayload_] Raw body length: ' + rawBody.length);
      console.log('[getPayload_] Content type: ' + (e.postData.type || 'unknown'));
      var parsed = JSON.parse(rawBody);
      console.log('[getPayload_] PASS — JSON parsed, keys: ' + Object.keys(parsed).join(', '));
      return parsed;
    }

    console.log('[getPayload_] No postData — falling back to e.parameter');
    var params = e && e.parameter ? e.parameter : {};
    console.log('[getPayload_] Parameter keys: ' + Object.keys(params).join(', '));
    return params;
  } catch (err) {
    console.error('[getPayload_] FAIL — ' + err.message);
    throw err;
  }
}

function jsonResponse(payload) {
  console.log('[jsonResponse] Returning: ' + JSON.stringify(payload));
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
