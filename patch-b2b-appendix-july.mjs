import { google } from 'googleapis';

// Adds a "Jul" column to the Appendix's Ad Set Level table
// (g3f393228d3a_0_160), between "Jun" and "Q1 Leads" — this table's axes are
// the opposite of the Reach slide's table (ad sets are rows, months are
// columns), so extending it means a new COLUMN, not a new row.
//
// "Q1 Leads"/"Q1 CPL" are deliberately left untouched: Q1 FY27 is Apr-Jun
// (closed, confirmed via this same table's own scoping last time), and July
// is Q2's first month. The per-month spend columns (Apr/May/Jun) are each
// independent monthly figures, not quarter-scoped, so a "Jul" spend column
// sits alongside them the same way — only the aggregate Leads/CPL columns
// carry the "Q1" label, and those stay as Apr-Jun sums.
//
// Column budget checked first (inspect19.mjs): existing 7 columns total
// 1208px, table starts at x=65 on a 1440px-wide slide — 167px of margin.
// An 8th ~122px-wide column fits with ~45px to spare, no shrinking needed
// (unlike the Reach table's row-height problem, this is a column insert and
// there's ample horizontal room).
//
// insertTableColumns inherits each new cell's style from the ADJACENT
// COLUMN cell in the SAME ROW (confirmed empirically) — since every row's
// existing cells are already internally consistent (header row = grey/bold,
// ad-set rows = white/normal, the Total row = red/bold/white), the new
// column should inherit correct per-row styling automatically. Verified via
// PDF re-export before considering this done.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TABLE = 'g3f393228d3a_0_160';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function insertCell(rowIndex, columnIndex, text) {
  return { insertText: { objectId: TABLE, cellLocation: { rowIndex, columnIndex }, insertionIndex: 0, text } };
}

const NEW_COL = 5; // between Jun(4) and Q1 Leads(old 5, now shifted to 6)

const requests = [
  { insertTableColumns: { tableObjectId: TABLE, cellLocation: { rowIndex: 0, columnIndex: 4 }, insertRight: true, number: 1 } },

  insertCell(0, NEW_COL, 'Jul\n'),
  insertCell(1, NEW_COL, '£659\n'),   // APAC
  insertCell(2, NEW_COL, '£0\n'),     // UK — paused
  insertCell(3, NEW_COL, '£636\n'),   // EMEA
  insertCell(4, NEW_COL, '£1,295\n'), // Q1 B2B Total row (sum of ad-set spend for July)
];

await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
console.log('Appendix table: Jul column added');
