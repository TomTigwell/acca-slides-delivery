import { google } from 'googleapis';

// One-off patch: fix two real bugs found on the already-generated ACCA B2C
// deck without regenerating from scratch:
// 1. Cover card body overflowed into the footer (233 chars, too long) —
//    shortened to 194 chars, matching the range that worked for Brand.
// 2. Appendix table: "Cybersecurity"/"Data Analytics" in the narrow Layer
//    column wrapped to 2-3 lines, growing the table tall enough to push
//    the Cert-AI group row and Total row off the bottom of the slide.
//    Also, row 10 (meant to be blank — only 6 of 8 real rows were needed)
//    only had columns 0-1 cleared, leaking the master's original
//    "ProDipSust Sustainability Strategy" row data (£348/61,477/411/...)
//    in columns 2-7. Row 12 (the Cert-AI slot) was never touched at all
//    and likely still carries the master's original Hot-layer row.

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function cell(objectId, rowIndex, columnIndex, text) {
  const cellLocation = { rowIndex, columnIndex };
  const requests = [{ deleteText: { objectId, cellLocation, textRange: { type: 'ALL' } } }];
  if (text) requests.push({ insertText: { objectId, cellLocation, insertionIndex: 0, text } });
  return requests;
}

const TABLE = 'g3f59aaca611_0_73';
const requests = [
  { deleteText: { objectId: 'cov_cardbody', textRange: { type: 'ALL' } } },
  { insertText: { objectId: 'cov_cardbody', insertionIndex: 0, text: 'Six B2C campaigns drove 1,092 Registers from £1,478.42 spend (£1.35 blended CPA) across Cert-OT, Cybersecurity, and Data Analytics. Two requested Cert-AI campaigns weren’t found on this account.' } },

  ...cell(TABLE, 5, 0, 'Cyber'),
  ...cell(TABLE, 6, 0, 'Cyber'),
  ...cell(TABLE, 8, 0, 'DataAn'),
  ...cell(TABLE, 9, 0, 'DataAn'),

  // Row 10: unused 4th slot in this group — clear every column, not just 0-1.
  ...[0, 1, 2, 3, 4, 5, 6, 7].flatMap((c) => cell(TABLE, 10, c, '')),
  // Row 12: the Cert-AI slot has no real data — clear every column so
  // nothing from the master's original Hot-layer row survives.
  ...[0, 1, 2, 3, 4, 5, 6, 7].flatMap((c) => cell(TABLE, 12, c, '')),
];

await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
console.log('patched', B2C_DECK);
