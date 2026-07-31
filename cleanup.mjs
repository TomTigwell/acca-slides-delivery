import { google } from 'googleapis';

// One-off: trash superseded draft copies. Trashed, not permanently deleted,
// so still recoverable.
const SUPERSEDED = [
  '1D1De_g0TIyaia0nUUxrnwKyyT78cu4sjerwAbxWghfs', // Brand — failed generate-brand.mjs run (batchUpdate errored before any edits applied, so this is an untouched copy)
  '1RayC0dlMgOzLX0li_pL66r0TCGN2C7wkv4vEi1PrvoA', // B2B — accidental generate.mjs run (forgot to point workflow at inspect20.mjs)
  '1HWxxkauVRGEz83Kba2tQwP34IR4DZeqbfT_e0V0LX0M', // Product — same accidental run
  '19SUj67khHq8DhUftgJa9UXd02kFQ_R-8DcAkBwdgcmM', // Employer Brand — same accidental run
];

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

for (const fileId of SUPERSEDED) {
  await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
  console.log('trashed', fileId);
}
