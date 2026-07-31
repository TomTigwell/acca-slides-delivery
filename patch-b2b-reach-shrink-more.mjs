import { google } from 'googleapis';

// The 42px minRowHeight still overflowed the slide in the rendered PDF export
// (confirmed visually — row 12 "UK — paused" clipped at the bottom edge)
// even though the Slides API echoed back rowHeight=42px per row, meaning the
// API's rowHeight readback doesn't reliably reflect the true rendered height
// (likely additional per-row padding/inset not captured in that field).
// Shrinking further, empirically, to 34px/row: 13 * 34 = 442px, table
// bottom at 251+442=693px vs the slide's 810px bottom — a much larger
// safety margin to absorb whatever the real per-row overhead turns out to be.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TABLE = 'p6_g11';
const PX = 6350;
const ROW_HEIGHT_PX = 34;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const requests = [
  {
    updateTableRowProperties: {
      objectId: TABLE,
      rowIndices: [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
      tableRowProperties: { minRowHeight: { magnitude: ROW_HEIGHT_PX * PX, unit: 'EMU' } },
      fields: 'minRowHeight',
    },
  },
];

await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
console.log('Shrunk Reach table rows to 34px each');
