import { google } from 'googleapis';

// One-off: trash superseded draft copies from earlier iterations this
// session (numbers-only pass with no cover, then a cover pass with a text
// overflow bug). Trashed, not permanently deleted, so still recoverable.
const SUPERSEDED = [
  '12S7eCclUfEcvWQ43uP83CIhTKaeydFU6Ii0b0KApW7g', // B2B v1 — numbers only, screenshot cover
  '12FUbbVjaOvi-zhTdvAxikvOUVEUcqj_hxxBeWZT9h_0', // Product v1
  '1cTchIJuhdiOezeMxl3kxbTdHyXf9MOvSqcqVZJ_dtVo', // Brand v1
  '1z5YkY84ZCgXjRyOXys74g0gC8epMVZ4FGY4l59PPV7Q', // B2B v2 — cover overflow bug
  '1B_N_N1bOnUY8Jqcxc_deu7D03Lfw0NphIu35bZVxJAA', // Product v2 — cover overflow bug
  '1NLPo_BBVQjlVfZsCpsD6l-VBzDxMD5WIOKLEzOctiWw', // Brand v2 — cover overflow bug
];

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });

for (const fileId of SUPERSEDED) {
  await drive.files.update({ fileId, supportsAllDrives: true, requestBody: { trashed: true } });
  console.log('trashed', fileId);
}
