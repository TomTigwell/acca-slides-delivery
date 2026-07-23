import { google } from 'googleapis';

// One-off: trash superseded draft copies from earlier iterations this
// session (numbers-only pass with no cover, then a cover pass with a text
// overflow bug). Trashed, not permanently deleted, so still recoverable.
const SUPERSEDED = [
  '1TSnYj5QyLo3x4RKgegkunhsxB3GyLg9apSAqJ6ROyqc', // Product — July-brief pass, had 2 text overflow collisions
  '16CVIP62Q96-YdSy6P4eW0b4vAPnOFxpnoj7oSnlvIPw', // Product — pre-July-brief (numbers-only) version
];

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

for (const fileId of SUPERSEDED) {
  await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
  console.log('trashed', fileId);
}
