import { google } from 'googleapis';

// Adds July as a 4th month group to the Reach — Employer Audience table
// (p6_g11) on the B2B deck, fixing the two bugs from the earlier attempt:
//   1. New rows from insertTableRows inherit the adjacent row's formatting
//      (here, the dark-red "Q1 Total" row) — every new cell gets an explicit
//      updateTableCellProperties (white bg) + updateTextStyle (real values
//      sampled from the existing April/May/June rows via inspect17.mjs, not
//      guessed) so it renders as a normal data row, not another total row.
//   2. The table's true rendered height is the SUM of its row heights
//      (311825 EMU = ~49.1px each for all 10 existing rows), not the stale
//      "size" field on the table element — confirmed via inspect17.mjs.
//      10 rows already use 491px of the slide's ~810px height, leaving only
//      ~68px of margin. 3 more rows at the same height would push the table
//      ~79px past the bottom of the slide (confirmed in the earlier reverted
//      attempt). Fix: shrink every row's minRowHeight to 42px (266700 EMU) —
//      well above the ~31px a single 9.5pt line actually needs — so all 13
//      rows fit in ~546px, ending at y≈796, comfortably inside the slide.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TABLE = 'p6_g11';
const PX = 6350;
const ROW_HEIGHT_PX = 42;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const BLACK = { red: 0, green: 0, blue: 0 };
const GREEN = { red: 0.08627451, green: 0.6392157, blue: 0.2901961 };
const RED = { red: 0.8627451, green: 0.14901961, blue: 0.14901961 };
const GREY = { red: 0.33333334, green: 0.33333334, blue: 0.33333334 };

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

function setCell(rowIndex, columnIndex, text, { bold = false, italic = false, color = BLACK } = {}) {
  const cellLocation = { rowIndex, columnIndex };
  return [
    whiteBg(rowIndex, columnIndex),
    { insertText: { objectId: TABLE, cellLocation, insertionIndex: 0, text } },
    {
      updateTextStyle: {
        objectId: TABLE,
        cellLocation,
        style: {
          bold, italic,
          foregroundColor: { opaqueColor: { rgbColor: color } },
          fontFamily: 'Calibri', fontSize: { magnitude: 9.5, unit: 'PT' },
        },
        textRange: { type: 'ALL' },
        fields: 'bold,italic,foregroundColor,fontFamily,fontSize',
      },
    },
  ];
}

const requests = [
  { insertTableRows: { tableObjectId: TABLE, cellLocation: { rowIndex: 9, columnIndex: 0 }, insertBelow: true, number: 3 } },

  ...setCell(10, 0, 'July\n', { bold: true }),
  ...setCell(10, 1, 'APAC\n'),
  ...setCell(10, 2, '£659\n'),
  ...setCell(10, 3, '37,757\n'),
  ...setCell(10, 4, '8\n'),
  ...setCell(10, 5, '£82\n', { color: GREEN }),

  ...setCell(11, 1, 'EMEA\n'),
  ...setCell(11, 2, '£636\n'),
  ...setCell(11, 3, '29,262\n'),
  ...setCell(11, 4, '2\n'),
  ...setCell(11, 5, '£318\n', { color: RED }),

  ...setCell(12, 1, 'UK — paused\n', { italic: true, color: GREY }),
  ...setCell(12, 2, '—\n', { color: GREY }),
  ...setCell(12, 3, '—\n', { color: GREY }),
  ...setCell(12, 4, '0\n', { color: GREY }),
  ...setCell(12, 5, '—\n', { color: GREY }),

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
console.log('Reach table updated with July rows, restyled and resized to fit');
