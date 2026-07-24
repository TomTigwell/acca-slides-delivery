import { google } from 'googleapis';

// One-off patch: fix the Reach-slide font-overlap and Total-row comma
// formatting on the already-generated ACCA Brand — July 2026 deck, without
// regenerating the whole thing from scratch.

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function fontSizeFix(objectId, pt) {
  return [{ updateTextStyle: { objectId, style: { fontSize: { magnitude: pt, unit: 'PT' } }, textRange: { type: 'ALL' }, fields: 'fontSize' } }];
}

const requests = [
  ...fontSizeFix('p3_i17', 8),
  ...fontSizeFix('p3_i16', 5.5),
  ...fontSizeFix('p3_i20', 14),
  ...fontSizeFix('g3f59aaca611_0_8', 14),
  { deleteText: { objectId: 'g3f59aaca611_0_73', cellLocation: { rowIndex: 13, columnIndex: 2 }, textRange: { type: 'ALL' } } },
  { insertText: { objectId: 'g3f59aaca611_0_73', cellLocation: { rowIndex: 13, columnIndex: 2 }, insertionIndex: 0, text: '£3,958.95' } },
];

await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests } });
console.log('patched', BRAND_DECK);
