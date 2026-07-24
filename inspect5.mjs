import { google } from 'googleapis';

// Read-only: full shape/table/image text+objectId dump for the NEW Product
// master (the client-edited, HTML-brief-matched deck saved as the new
// template). Same approach as inspect.mjs, pointed at the new file, plus
// image elements (the old inspect.mjs skipped those) so we can confirm
// nothing is a flattened screenshot in the new structure.

const NEW_MASTER = '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

function dumpElement(el) {
  if (el.shape?.text) {
    const t = el.shape.text.textElements.map((te) => te.textRun?.content || '').join('');
    if (t.trim()) console.log(`  [shape ${el.objectId}] ${JSON.stringify(t)}`);
  }
  if (el.table) {
    el.table.tableRows.forEach((row, r) => {
      row.tableCells.forEach((cell, c) => {
        const t = (cell.text?.textElements || []).map((te) => te.textRun?.content || '').join('');
        if (t.trim()) console.log(`  [table ${el.objectId} r${r}c${c}] ${JSON.stringify(t)}`);
      });
    });
  }
  if (el.image) {
    console.log(`  [IMAGE ${el.objectId}] size=${JSON.stringify(el.size)}`);
  }
  if (el.elementGroup?.children) el.elementGroup.children.forEach(dumpElement);
}

const { data } = await slides.presentations.get({ presentationId: NEW_MASTER });
console.log(`=== NEW PRODUCT MASTER (${NEW_MASTER}) — ${data.slides.length} slides ===`);
data.slides.forEach((slide, i) => {
  console.log(`--- Slide ${i + 1} (${slide.objectId}) ---`);
  (slide.pageElements || []).forEach(dumpElement);
});
