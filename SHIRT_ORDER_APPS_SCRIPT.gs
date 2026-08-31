/**
 * Rise Up Shirt Orders backend
 * 1) Create a Google Sheet.
 * 2) Extensions > Apps Script.
 * 3) Paste this code.
 * 4) Replace FOLDER_ID with the ID of a private Google Drive folder for receipts.
 * 5) Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: Anyone
 * 6) Copy the Web App URL and paste it into SHIRT_ORDER_ENDPOINT in index.html.
 */
const FOLDER_ID = 'PASTE_PRIVATE_DRIVE_FOLDER_ID_HERE';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.teenName || !data.size || !data.fileData) throw new Error('Missing required fields');

    const allowedSizes = ['S','M','L','XL'];
    if (!allowedSizes.includes(data.size)) throw new Error('Invalid size');

    const bytes = Utilities.base64Decode(data.fileData);
    const blob = Utilities.newBlob(bytes, data.mimeType || 'application/octet-stream',
      new Date().toISOString().replace(/[:.]/g,'-') + '_' + data.teenName.replace(/[^a-z0-9 _-]/gi,'') + '_' + data.fileName);

    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp','Teen Name','Size','Receipt File']);
    }
    sheet.appendRow([new Date(), data.teenName, data.size, file.getUrl()]);

    return ContentService.createTextOutput(JSON.stringify({ok:true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
