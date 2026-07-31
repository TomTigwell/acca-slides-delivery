import { google } from 'googleapis';

// The last two PDF re-exports came back byte-identical despite the API
// confirming rows 2-13 shrank to 40px (bottom=651px, well inside the
// 810px slide) — looks like Drive's PDF export is serving a cached render
// that hasn't picked up the table relayout. Nudging the row height to a
// genuinely new value (38px, not just re-setting 40px) to force a real
// change and see whether the export reflects it.

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
const TABLE = 'g3f59aaca611_0_73';
const PX = 6350;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const requests = [
  {
    updateTableRowProperties: {
      objectId: TABLE,
      rowIndices: [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13],
      tableRowProperties: { minRowHeight: { magnitude: 38 * PX, unit: 'EMU' } },
      fields: 'minRowHeight',
    },
  },
];

await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
console.log('Row heights nudged to 38px to force relayout/cache invalidation');
