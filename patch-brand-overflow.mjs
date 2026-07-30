import { google } from 'googleapis';

// One-off patch: the merged "need" item written for the column-mapping
// fix (patch-column-fix.mjs) was too long (187 chars vs ~113-131 for its
// neighbours) and overflowed into the item below it. Shortened to match.

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const requests = [
  { deleteText: { objectId: 'g3f483bfcec4_0_143', textRange: { type: 'ALL' } } },
  { insertText: { objectId: 'g3f483bfcec4_0_143', insertionIndex: 0, text: 'Confirm whether the APAC/Africa CPL gap reflects saturation vs. a creative/targeting difference — share July’s VCF leads to check.\n' } },
];

await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests } });
console.log('patched Brand deck overflow fix');
