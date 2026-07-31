import { google } from 'googleapis';

// Revert: appending July rows to the Reach slide's Month/Region table (p6_g11)
// was the wrong call. Two problems: (1) the new rows inherited the dark-red
// "Q1 Total" row's formatting instead of a normal data row's style, and more
// importantly (2) the table's fixed row height meant 3 extra rows pushed it
// past the bottom of the slide (the July/EMEA and July/UK rows render
// off-canvas — confirmed via PDF re-export). On reflection this table is a
// closed-quarter (Apr-Jun) historical breakdown, not a running monthly log —
// July is already correctly and prominently reported via this same slide's
// right-side stat tiles (labelled "July 2026", already verified current).
// Deleting the 3 rows restores the table to its original, correctly-fitting
// 10-row state.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TABLE = 'p6_g11';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const requests = [
  { deleteTableRow: { tableObjectId: TABLE, cellLocation: { rowIndex: 12, columnIndex: 0 } } },
  { deleteTableRow: { tableObjectId: TABLE, cellLocation: { rowIndex: 11, columnIndex: 0 } } },
  { deleteTableRow: { tableObjectId: TABLE, cellLocation: { rowIndex: 10, columnIndex: 0 } } },
];

await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
console.log('Reverted the 3 July rows on the Reach table');
