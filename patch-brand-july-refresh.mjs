import { google } from 'googleapis';
import fs from 'fs';

// Updates the Brand report (1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw) with a
// full July data refresh, same format:
//   - VCF Africa grew from 25->27 leads as the month closed out further
//     (£1,181.47->£1,283.70 spend, 83,432->90,268 impr, 746->819 clicks,
//     10,368->10,925 reach).
//   - VCF APAC held at 10 leads but spend/impr/clicks also grew slightly
//     (£1,777.48->£1,789.08, 87,514->88,275, 613->616); reach unchanged
//     at 22,518.
//   - BOOST_VCF_APAC_JULY_2026_INHOUSE (organic) unchanged: £1,000.00 /
//     141,897 impr / 2,668 clicks.
//   - eZine's three static campaigns remained fully paused all month —
//     no change, no edits needed for those objects.
//   - "ACCA Learning Always On" still not found; account now has 30 total
//     campaigns (was 27) — cross-references to the campaign count updated.
//
// Fresh totals (LinkedIn Ads MCP pull, account 509501623, 01-31 Jul 2026):
//   TOTAL_LEADS=37  TOTAL_LEAD_SPEND=£3,072.78  BLENDED_CPL=£83.05
//   TOTAL_SPEND=£4,072.78  TOTAL_IMPR=320,440  TOTAL_CLICKS=4,103
//   Africa CPL=£47.54  APAC CPL=£178.91  (APAC is 3.8x Africa's rate)

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const TABLE = 'g3f59aaca611_0_73';
const PX = 6350;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

function edit(objectId, text) {
  return [
    { deleteText: { objectId, textRange: { type: 'ALL' } } },
    { insertText: { objectId, insertionIndex: 0, text } },
  ];
}
function cell(rowIndex, columnIndex, text) {
  const cellLocation = { rowIndex, columnIndex };
  return [
    { deleteText: { objectId: TABLE, cellLocation, textRange: { type: 'ALL' } } },
    { insertText: { objectId: TABLE, cellLocation, insertionIndex: 0, text } },
  ];
}

function imageBox(xPx, yPx, wPx, hPx) {
  return {
    size: { width: { magnitude: wPx * PX, unit: 'EMU' }, height: { magnitude: hPx * PX, unit: 'EMU' } },
    transform: { scaleX: 1, scaleY: 1, translateX: xPx * PX, translateY: yPx * PX, unit: 'EMU' },
  };
}

const requests = [
  // ---- Cover ----
  ...edit('cov_cardbody', 'VCF Lead Generation delivered 37 leads in July — 27 from Africa at £47.54 CPL and 10 from APAC at £178.91 CPL — while eZine’s three static campaigns stayed fully paused with no spend or delivery this month.\n'),

  // ---- Exec Summary ----
  ...edit('g3f483bfcec4_0_27', 'In July, VCF Lead Generation ran two active campaigns — Africa and APAC — delivering 37 leads combined from £3,072.78 spend, plus a £1,000 organic boost post reaching 141,897 impressions. eZine’s three static campaigns remained paused throughout the month with zero delivery. No activity was found under "ACCA Learning Always On" on this account.\n\n'),

  // ---- What We Did (items 1, 2, 4 — item 3/eZine unchanged) ----
  ...edit('g3f483bfcec4_0_35', 'The Africa campaign delivered 27 leads from £1,283.70 spend — a £47.54 CPL, less than a third of APAC’s cost per lead on the same format.\n\n\n'),
  ...edit('g3f483bfcec4_0_39', 'APAC delivered 10 leads from £1,789.08 spend — £178.91 CPL, 3.8x Africa’s rate on the same objective and creative approach.\n\n\n'),
  ...edit('g3f483bfcec4_0_47', 'No campaign matching that name was found across all 30 campaigns under account 509501623 for July — flagged for confirmation rather than reported on with invented figures.\n\n\n'),

  // ---- North Star narrative ----
  ...edit('g3f483bfcec4_0_106', 'July delivered 37 VCF leads at a blended £83.05 CPL — Africa’s £47.54 CPL is pulling the blend down against APAC’s £178.91. eZine remained fully paused all month; reactivating even one cluster would add incremental reach without touching the VCF budget.\n\n\n'),
  ...edit('g3f60d68de54_0_2', '37\n'), // North Star card "JULY LEADS" value — renamed since deck was first built; current live objectId confirmed via inspect22/23.mjs
  ...edit('g3f60d68de54_0_5', '£83.05\n'), // North Star card "BLENDED CPL" value

  // ---- What We Recommend Next ----
  ...edit('g3f483bfcec4_0_120', 'Running VCF Africa Lead Generation at current budget — £47.54 CPL is the most efficient result in the account this month.\n\n\n'),

  // ---- Reach ----
  ...edit('p3_i17', 'VCF’s Africa and APAC Lead Generation campaigns reached an estimated 10,925 and 22,518 unique members respectively in July, per LinkedIn’s approximate member reach metric. eZine added no reach this month.\n\n'),
  ...edit('p3_i20', '320,440\n\n'),
  ...edit('g3f59aaca611_0_8', '37\n\n'), // Reach slide's own "JULY LEADS" tile — separate object from the North Star slide's version

  // ---- Experiments ----
  ...edit('g3f483bfcec4_0_269', 'Confirmed a real gap, not comparable. Africa delivered 27 leads at £47.54 CPL; APAC delivered 10 leads at £178.91 CPL — a 3.8x difference on the same format and objective.\n\n\n'),
  ...edit('g3f483bfcec4_0_280', 'Confirmed. The boost delivered 141,897 impressions at £7.05 CPM — below both Lead Gen campaigns’ effective CPM (APAC £20.27, Africa £14.22) — though it drove website conversions, not native leads, a different objective.\n\n\n'),
  ...edit('g3f59aaca611_0_36', 'No campaign matching "ACCA Learning Always On" was found under account 509501623 for July, checked across all 30 campaigns on the account, including a substring search for "learning" and "always."\n\n\n'),

  // ---- Appendix table: refresh APAC/Africa rows + total ----
  ...cell(3, 2, '£1,789.08'), ...cell(3, 3, '88,275'), ...cell(3, 4, '616'), ...cell(3, 5, '0.70%'), ...cell(3, 6, '£20.27'), ...cell(3, 7, '£2.90'),
  ...cell(4, 2, '£1,283.70'), ...cell(4, 3, '90,268'), ...cell(4, 4, '819'), ...cell(4, 5, '0.91%'), ...cell(4, 6, '£14.22'), ...cell(4, 7, '£1.57'),
  ...cell(13, 2, '£4,072.78'), ...cell(13, 3, '320,440'), ...cell(13, 4, '4,103'),
];

console.log('applying', requests.length, 'requests...');
await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests } });
console.log('text + table updates applied');

// ---- Swap the North Star trend chart image for the refreshed July bar ----
const { data: uploaded } = await drive.files.create({
  requestBody: { name: 'brand_northstar_chart_v2.png', parents: [SHARED_DRIVE] },
  media: { mimeType: 'image/png', body: fs.createReadStream('assets/brand_chart_v2.png') },
  fields: 'id',
  supportsAllDrives: true,
});
await drive.permissions.create({
  fileId: uploaded.id,
  requestBody: { role: 'reader', type: 'anyone' },
  supportsAllDrives: true,
});
const chartUrl = `https://drive.google.com/uc?export=view&id=${uploaded.id}`;

await slides.presentations.batchUpdate({
  presentationId: BRAND_DECK,
  requestBody: {
    requests: [
      { deleteObject: { objectId: 'brand_ns_chart_img' } },
      {
        createImage: {
          objectId: 'brand_ns_chart_img',
          url: chartUrl,
          elementProperties: { pageObjectId: 'g3f483bfcec4_0_95', ...imageBox(140, 460, 634, 274) },
        },
      },
    ],
  },
});
console.log('North Star chart image refreshed');
