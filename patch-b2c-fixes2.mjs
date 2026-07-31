import { google } from 'googleapis';

// Follow-up fixes after the July refresh:
//   1. Reach slide's own "JULY REGISTERS" tile (g3f59aaca611_0_8/9) — separate
//      object from the North Star slide's version, missed on the first pass.
//   2. g3f483bfcec4_0_283 (second experiment card's JULY ACTION cell) still
//      referenced the stale £2.44 Cybersecurity CPA — missed because only
//      the adjacent Learning cell (280) was in scope originally.
//   3. Appendix table (g3f59aaca611_0_73): filling rows 10/12 with real
//      Cert-AI text grew them from near-empty to full content height,
//      pushing the table to ~799px bottom on an 810px-tall slide — same
//      near-invisible-margin overflow class as the B2B Reach table. Row 0
//      ("Appendix" H1, built into the table) and row 1 (column headers) are
//      left alone; rows 2-13 (group labels + data) shrink from ~52-56px to
//      40px, giving a safe ~150px margin instead of ~10px.

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
const TABLE = 'g3f59aaca611_0_73';
const PX = 6350;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function edit(objectId, text) {
  return [
    { deleteText: { objectId, textRange: { type: 'ALL' } } },
    { insertText: { objectId, insertionIndex: 0, text } },
  ];
}

const requests = [
  ...edit('g3f59aaca611_0_8', '1,106\n'),
  ...edit('g3f59aaca611_0_9', 'Eight B2C campaigns\n'),
  ...edit('g3f483bfcec4_0_283', 'August EOM: test whether Data Analytics’ Single Image creative approach transfers to Cybersecurity to close its £2.52 CPA gap.\n\n'),
  {
    updateTableRowProperties: {
      objectId: TABLE,
      rowIndices: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
      tableRowProperties: { minRowHeight: { magnitude: 40 * PX, unit: 'EMU' } },
      fields: 'minRowHeight',
    },
  },
];

await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
console.log('B2C follow-up fixes applied: Reach tile, experiment action text, table row heights');
