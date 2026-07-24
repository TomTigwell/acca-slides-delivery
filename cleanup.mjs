import { google } from 'googleapis';

// One-off: trash superseded draft copies. Trashed, not permanently deleted,
// so still recoverable.
const SUPERSEDED = [
  '1D1De_g0TIyaia0nUUxrnwKyyT78cu4sjerwAbxWghfs', // Brand — failed generate-brand.mjs run (batchUpdate errored before any edits applied, so this is an untouched copy)
];

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

for (const fileId of SUPERSEDED) {
  await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
  console.log('trashed', fileId);
}
