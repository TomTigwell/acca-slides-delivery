import { google } from 'googleapis';

// One-off: trash superseded draft copies. Trashed, not permanently deleted,
// so still recoverable.
const SUPERSEDED = [
  '1zxIRpmX3CxEdLeSClUrQUoQk_jipVLP7_EmRYRNQpHI', // Product — throwaway round-trip check copy against the new template master
];

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

for (const fileId of SUPERSEDED) {
  await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
  console.log('trashed', fileId);
}
