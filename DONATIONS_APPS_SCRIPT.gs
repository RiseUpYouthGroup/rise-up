/**
 * Rise Up Donations backend
 * 1) Create a Google Sheet named "Rise Up Donations".
 * 2) Create a PRIVATE Google Drive folder named "Rise Up Donation Receipts".
 * 3) In the Sheet: Extensions > Apps Script, paste this code.
 * 4) Replace FOLDER_ID with the private folder ID.
 * 5) Deploy > New deployment > Web app. Execute as: Me. Who has access: Anyone.
 * 6) Copy the Web App URL and paste it into DONATION_ENDPOINT in index.html.
 */
const FOLDER_ID = 'PASTE_PRIVATE_DRIVE_FOLDER_ID_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.familyName || !data.amount || !data.fileData) throw new Error('Missing required fields');
    const amount = Number(data.amount);
    if (!isFinite(amount) || amount <= 0) throw new Error('Invalid amount');

    const bytes = Utilities.base64Decode(data.fileData);
    const safeName = String(data.familyName).replace(/[^a-z0-9 _-]/gi,'');
    const blob = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream',
      new Date().toISOString().replace(/[:.]/g,'-') + '_' + safeName + '_' + data.fileName);
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['Timestamp','Family / Donor Name','Amount','Receipt File']);
    sheet.appendRow([new Date(), data.familyName, amount, file.getUrl()]);
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}
