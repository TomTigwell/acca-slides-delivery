import { google } from 'googleapis';

// Read-only: find sheetsChart page elements (linked-from-Sheets charts) on
// the new Product master. inspect5.mjs only dumped shape/table/image
// elements and missed these entirely — a linked chart has no shape.text,
// so it looked like "no text content" when it's actually backed by a live
// Google Sheet (spreadsheetId + chartId) that the user edits directly.

const NEW_MASTER = '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

function dumpElement(el) {
  if (el.sheetsChart) {
    console.log(`  [SHEETSCHART ${el.objectId}] spreadsheetId=${el.sheetsChart.spreadsheetId} chartId=${el.sheetsChart.chartId}`);
  }
  if (el.elementGroup?.children) el.elementGroup.children.forEach(dumpElement);
}

const { data } = await slides.presentations.get({ presentationId: NEW_MASTER });
console.log(`=== Sheets-linked charts in NEW PRODUCT MASTER (${NEW_MASTER}) ===`);
data.slides.forEach((slide, i) => {
  console.log(`--- Slide ${i + 1} (${slide.objectId}) ---`);
  (slide.pageElements || []).forEach(dumpElement);
});
