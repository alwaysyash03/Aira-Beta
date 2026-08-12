/**
 * AIRA WAITLIST — GOOGLE SHEETS BACKEND
 * ---------------------------------------------------
 * SETUP:
 * 1. Create a new Google Sheet. Rename the first tab to exactly: Signups
 * 2. In row 1, add headers: Timestamp | Role | Name | Phone | Area
 * 3. Extensions > Apps Script. Delete any starter code and paste this whole file in.
 * 4. Change ADMIN_PASSWORD below to something only you know.
 * 5. Click Deploy > New deployment > type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the Web App URL it gives you.
 *    - Paste it into APPS_SCRIPT_URL in Aira_Waitlist_Landing_Page.html
 *    - Paste it into APPS_SCRIPT_URL in admin.html
 * 7. Re-deploy (Deploy > Manage deployments > edit > new version) any time you change this file.
 */

const SHEET_NAME = 'Signups';
const ADMIN_PASSWORD = 'Nikkiii@123';

// Handles form submissions from the landing page.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      data.role || '',
      data.name || '',
      data.phone || '',
      data.area || ''
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Serves signup data to the admin dashboard — requires the correct password as a query param.
function doGet(e) {
  const password = e.parameter.password;
  if (password !== ADMIN_PASSWORD) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: 'unauthorized' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  values.shift(); // drop header row

  const rows = values
    .filter(r => r[0]) // skip blank rows
    .map(r => ({
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      role: r[1],
      name: r[2],
      phone: r[3],
      area: r[4]
    }));

  return ContentService
    .createTextOutput(JSON.stringify(rows))
    .setMimeType(ContentService.MimeType.JSON);
}
