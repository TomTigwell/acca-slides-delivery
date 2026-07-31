import { google } from 'googleapis';

// Follow-up fix: rows 11 (EMEA) and 12 (UK — paused) on the Reach table's
// column 0 (Month) were left blank to match the existing "blank month cell
// for non-first rows in a group" pattern (see April/May/June groups), but
// since they were never touched with a whiteBg request, they still carry
// the red background inherited from insertTableRows. Fix: set white bg
// on those two cells with no text change.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TABLE = 'p6_g11';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function whiteBg(rowIndex, columnIndex) {
  return {
    updateTableCellProperties: {
      objectId: TABLE,
      tableRange: { location: { rowIndex, columnIndex }, rowSpan: 1, columnSpan: 1 },
      tableCellProperties: { tableCellBackgroundFill: { solidFill: { color: { rgbColor: { red: 1, green: 1, blue: 1 } } } } },
      fields: 'tableCellBackgroundFill.solidFill.color',
    },
  };
}

const requests = [whiteBg(11, 0), whiteBg(12, 0)];

await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
console.log('Fixed col0 background for rows 11 and 12');
