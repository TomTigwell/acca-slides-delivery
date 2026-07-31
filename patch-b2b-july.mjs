import { google } from 'googleapis';

// Updates the B2B report (1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA) with
// the latest July 2026 data, without changing the deck's format:
//   1. Reach — Employer Audience table (p6_g11): append July's 3 region rows
//      below the existing "Q1 Total" row, following the exact same
//      Month/Region/Spend/Impressions/Leads/CPL pattern already used for
//      April/May/June. Q1 FY27 (Apr-Jun) is a closed quarter per the deck's
//      own Appendix table — July is the first month of Q2 FY27, so it is
//      added as its own trailing month group, not folded into Q1 Total.
//   2. North Star slide (p4_i22): the narrative sentence is updated to
//      report July's actual vs the Q2 FY27 target, replacing the old
//      forward-looking "launching July" clause (July has now happened).
// Cover (cov_*) and Reach headline tiles (p6_i5/i9/i11/i13/i15) were checked
// against a fresh LinkedIn pull and are already exactly current — no change
// needed there. The Q1 FY27 Ad Set Level appendix table is intentionally
// left untouched: it's explicitly scoped to the closed Apr-Jun quarter.

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function cellReq(tableObjectId, rowIndex, columnIndex, text) {
  const cellLocation = { rowIndex, columnIndex };
  return [
    { deleteText: { objectId: tableObjectId, cellLocation, textRange: { type: 'ALL' } } },
    { insertText: { objectId: tableObjectId, cellLocation, insertionIndex: 0, text } },
  ];
}

const TABLE = 'p6_g11';

const requests = [
  // Insert 3 fresh rows below row 9 (the existing "Q1 Total" row).
  { insertTableRows: { tableObjectId: TABLE, cellLocation: { rowIndex: 9, columnIndex: 0 }, insertBelow: true, number: 3 } },

  ...cellReq(TABLE, 10, 0, 'July\n'),
  ...cellReq(TABLE, 10, 1, 'APAC\n'),
  ...cellReq(TABLE, 10, 2, '£659\n'),
  ...cellReq(TABLE, 10, 3, '37,757\n'),
  ...cellReq(TABLE, 10, 4, '8\n'),
  ...cellReq(TABLE, 10, 5, '£82\n'),

  ...cellReq(TABLE, 11, 1, 'EMEA\n'),
  ...cellReq(TABLE, 11, 2, '£636\n'),
  ...cellReq(TABLE, 11, 3, '29,262\n'),
  ...cellReq(TABLE, 11, 4, '2\n'),
  ...cellReq(TABLE, 11, 5, '£318\n'),

  ...cellReq(TABLE, 12, 1, 'UK — paused\n'),
  ...cellReq(TABLE, 12, 2, '—\n'),
  ...cellReq(TABLE, 12, 3, '—\n'),
  ...cellReq(TABLE, 12, 4, '0\n'),
  ...cellReq(TABLE, 12, 5, '—\n'),

  { deleteText: { objectId: 'p4_i22', textRange: { type: 'ALL' } } },
  { insertText: { objectId: 'p4_i22', insertionIndex: 0, text: 'Q1 closed at 21 leads vs a target of 30 — 70% of goal, but CPL improved 24% quarter-on-quarter (£291 → £221). July, the first month of Q2, delivered 10 leads at £130 CPL — 22% of the 45-lead Q2 target — as the persona-led creative and form architecture began rolling out.\n' } },
];

await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
console.log('B2B report updated with July data');
