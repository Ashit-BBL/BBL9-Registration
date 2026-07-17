// ═══════════════════════════════════════════════════════════════════
//  BBL Season 9 — Registration Form Backend
//  Paste this entire file into Google Apps Script (script.google.com)
//  Then deploy as a Web App (see instructions at bottom of this file)
// ═══════════════════════════════════════════════════════════════════

// ── CONFIGURATION ───────────────────────────────────────────────────
const SHEET_NAME      = 'BBL9 Registrations';   // Tab name in your Google Sheet
const DRIVE_FOLDER    = 'BBL9 Submissions';      // Folder name in your Google Drive
const ADMIN_EMAIL     = 'ahgandhi65@gmail.com';  // Email for new registration alerts
// ────────────────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);

    // 1. Write to Google Sheet
    const sheet = getOrCreateSheet();
    appendToSheet(sheet, payload);

    // 2. Save uploaded files to Google Drive
    const folder = getOrCreateFolder(payload.teamName);
    if (payload.clubIdFile)  saveFile(folder, payload.clubIdFile,  'ClubID_'   + payload.teamName);
    if (payload.paymentFile) saveFile(folder, payload.paymentFile, 'Payment_'  + payload.teamName);

    // 3. Send email notification to admin
    sendAdminEmail(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── SHEET ────────────────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    // Header row
    const headers = [
      'Timestamp', 'Email', 'Team Name', 'Division', 'Retained from BBL8', 'Alternate Home Club',
      'P1 Name (Captain)',    'P1 Date of Birth', 'P1 Mobile',
      'P2 Name (Vice Capt)', 'P2 Date of Birth', 'P2 Mobile',
      'P3 Name', 'P3 Date of Birth', 'P3 Mobile',
      'P4 Name', 'P4 Date of Birth', 'P4 Mobile',
      'P5 Name', 'P5 Date of Birth', 'P5 Mobile',
      'P6 Name', 'P6 Date of Birth', 'P6 Mobile',
      'P7 Name', 'P7 Date of Birth', 'P7 Mobile',
      'P8 Name', 'P8 Date of Birth', 'P8 Mobile',
      'Club ID File', 'Payment File', 'Notes'
    ];
    sheet.appendRow(headers);

    // Style header row
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#E65100').setFontColor('#FFFFFF').setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 160);  // Timestamp
    sheet.setColumnWidth(2, 200);  // Email
    sheet.setColumnWidth(3, 220);  // Team Name
  }

  return sheet;
}

function appendToSheet(sheet, d) {
  const row = [
    new Date(),
    d.email,
    d.teamName,
    d.division,
    d.retained,
    d.homeClub || '',
  ];

  // Add player columns (8 players × 3 fields)
  (d.players || []).forEach(p => {
    row.push(p.name || '', p.dob || '', p.mobile || '');
  });
  // Pad to 8 players if fewer were sent
  const filled = (d.players || []).length;
  for (let i = filled; i < 8; i++) row.push('', '', '');

  // File names + notes
  row.push(d.clubIdFile?.name  || '', d.paymentFile?.name || '', d.notes || '');

  sheet.appendRow(row);
}

// ── DRIVE ────────────────────────────────────────────────────────────
function getOrCreateFolder(teamName) {
  // Get or create top-level BBL9 folder
  let parentFolder;
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER);
  parentFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER);

  // Create sub-folder per team (sanitise name)
  const safeName = (teamName || 'Unknown').replace(/[^a-zA-Z0-9 _-]/g, '').trim();
  const subFolders = parentFolder.getFoldersByName(safeName);
  return subFolders.hasNext() ? subFolders.next() : parentFolder.createFolder(safeName);
}

function sendAdminEmail(d) {
  try {
    const players = (d.players || []).map((p, i) =>
      `  Player ${i+1}: ${p.name || '-'}  |  DOB: ${p.dob || '-'}  |  Mobile: ${p.mobile || '-'}`
    ).join('\n');

    const subject = `BBL9 Registration: ${d.teamName} (Division ${d.division})`;
    const body =
`New BBL Season 9 Registration Received!

Team Name  : ${d.teamName}
Division   : ${d.division}
Email      : ${d.email}
Retained   : ${d.retained} players from BBL8
Home Club  : ${d.homeClub || 'N/A'}
Submitted  : ${new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' })}

── PLAYERS ──────────────────────────────
${players}

── DOCUMENTS ────────────────────────────
Club ID File  : ${d.clubIdFile?.name  || 'Not uploaded'}
Payment File  : ${d.paymentFile?.name || 'Not uploaded'}

── NOTES ────────────────────────────────
${d.notes || '(none)'}

Files are also attached to this email and saved in Google Drive under "BBL9 Submissions/${d.teamName}".`;

    // Build attachments array from base64 data
    const attachments = [];
    if (d.clubIdFile?.data) {
      attachments.push({
        fileName: d.clubIdFile.name,
        mimeType: d.clubIdFile.mimeType || 'application/octet-stream',
        content:  Utilities.base64Decode(d.clubIdFile.data),
      });
    }
    if (d.paymentFile?.data) {
      attachments.push({
        fileName: d.paymentFile.name,
        mimeType: d.paymentFile.mimeType || 'application/octet-stream',
        content:  Utilities.base64Decode(d.paymentFile.data),
      });
    }

    GmailApp.sendEmail(ADMIN_EMAIL, subject, body, { attachments });
  } catch (err) {
    Logger.log('Email error: ' + err.message);
  }
}

function saveFile(folder, fileObj, prefix) {
  try {
    const ext      = fileObj.name.split('.').pop();
    const fileName = prefix + '_' + Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd_HHmmss') + '.' + ext;
    const blob     = Utilities.newBlob(
      Utilities.base64Decode(fileObj.data),
      fileObj.mimeType || 'application/octet-stream',
      fileName
    );
    folder.createFile(blob);
  } catch (err) {
    Logger.log('File save error: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════
//  DEPLOYMENT INSTRUCTIONS
// ═══════════════════════════════════════════════════════════════════
//
//  1. Go to https://script.google.com
//  2. Click "New project"
//  3. Delete all default code and paste this entire file
//  4. Click the floppy-disk icon to Save (name it "BBL9 Backend")
//
//  5. LINK TO YOUR GOOGLE SHEET:
//     - Open your Google Sheet (create one if needed: sheets.google.com)
//     - In Apps Script: click "Resources" is old UI — in new UI go to:
//       Project Settings (gear icon) → there is no direct link option
//     - Instead: In the Script editor click Extensions > Apps Script
//       from within your Sheet — this auto-links the script to that Sheet
//     - OR: In the script, replace SpreadsheetApp.getActiveSpreadsheet()
//       with SpreadsheetApp.openById('YOUR_SHEET_ID') and paste your Sheet ID
//
//  6. DEPLOY:
//     - Click "Deploy" (top right) → "New deployment"
//     - Type: "Web app"
//     - Description: "BBL9 Registration v1"
//     - Execute as: "Me (your Gmail)"
//     - Who has access: "Anyone"   ← allows the HTML form to post data
//     - Click "Deploy"
//     - COPY the Web App URL shown
//
//  7. PASTE THE URL into BBL_Season9_Registration.html:
//     Find this line near the top of the <script> section:
//       const APPS_SCRIPT_URL = 'YOUR_APPS_SCRIPT_WEB_APP_URL_HERE';
//     Replace the placeholder with your copied URL.
//
//  8. SHEET SHARING (admin only):
//     - Open your Google Sheet
//     - Share → add committee member Gmail addresses → Editor or Viewer
//     - The Drive folder "BBL9 Submissions" is also private to you by default
//
// ═══════════════════════════════════════════════════════════════════
