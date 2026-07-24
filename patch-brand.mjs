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

// North Star note text box still overlaps the big stat numbers above it —
// that fix landed in generate-brand.mjs's source but was never applied to
// this already-generated deck (this patch only covered the Reach slide).
// Delete and recreate it lower, matching the corrected source coordinates.
const PX = 6350;
const MUTED = { red: 0.9490, green: 0.6392, blue: 0.7098 };
const chartSlideId = 'g3f483bfcec4_0_95';
const xPx = 57, yPx = 260, wPx = 799, hPx = 494;
const noteBox = { x: xPx + 30, y: yPx + 210, w: wPx - 60, h: hPx - 240 };

const requests = [
  ...fontSizeFix('p3_i17', 8),
  ...fontSizeFix('p3_i16', 5.5),
  ...fontSizeFix('p3_i20', 14),
  ...fontSizeFix('g3f59aaca611_0_8', 14),
  { deleteText: { objectId: 'g3f59aaca611_0_73', cellLocation: { rowIndex: 13, columnIndex: 2 }, textRange: { type: 'ALL' } } },
  { insertText: { objectId: 'g3f59aaca611_0_73', cellLocation: { rowIndex: 13, columnIndex: 2 }, insertionIndex: 0, text: '£3,958.95' } },
  { deleteObject: { objectId: 'brand_ns_note' } },
  { createShape: { objectId: 'brand_ns_note2', shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: chartSlideId, size: { width: { magnitude: noteBox.w * PX, unit: 'EMU' }, height: { magnitude: noteBox.h * PX, unit: 'EMU' } }, transform: { scaleX: 1, scaleY: 1, translateX: noteBox.x * PX, translateY: noteBox.y * PX, unit: 'EMU' } } } },
  { insertText: { objectId: 'brand_ns_note2', insertionIndex: 0, text: 'VCF Africa + APAC Lead Generation, July 2026. Native stat block — a per-month Sheets-linked chart for Brand is still pending (needs the Sheets API enabled on the automation project).' } },
  { updateTextStyle: { objectId: 'brand_ns_note2', style: { fontFamily: 'DM Sans', fontSize: { magnitude: 10, unit: 'PT' }, foregroundColor: { opaqueColor: { rgbColor: MUTED } } }, textRange: { type: 'ALL' }, fields: 'fontFamily,fontSize,foregroundColor' } },
  { updateParagraphStyle: { objectId: 'brand_ns_note2', style: { alignment: 'START', lineSpacing: 130 }, textRange: { type: 'ALL' }, fields: 'alignment,lineSpacing' } },
  { updateShapeProperties: { objectId: 'brand_ns_note2', shapeProperties: { outline: { propertyState: 'NOT_RENDERED' } }, fields: 'outline.propertyState' } },
];

await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests } });
console.log('patched', BRAND_DECK);
