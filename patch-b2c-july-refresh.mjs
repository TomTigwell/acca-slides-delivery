import { google } from 'googleapis';
import fs from 'fs';

// Updates the B2C report (1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU) with a
// full July data refresh, same format:
//   1. The 2 previously-missing "Cert-AI Phase 3" campaigns (Traffic + Views)
//      have since launched (24 Jul per their start dates) and now have real
//      July data — added into the appendix table's two already-reserved,
//      already-empty rows (10 and 12), and the "not found" framing removed
//      everywhere it appeared (cover, exec summary, What We Did, Recommend
//      Next, appendix experiment card).
//   2. The original 6 campaigns' spend/impressions/clicks grew substantially
//      since this report was first built mid-month (e.g. Cert-OT Traffic:
//      £299.72->£351.04, 77,676->120,991 impr) even though Register counts
//      barely moved — every figure derived from these (CPAs, CTRs, CPMs,
//      CPCs, the totals, the North Star chart) is refreshed to match.
//   3. A genuine finding: the "817 vs 1,141" conversion-total discrepancy
//      flagged in the original report has NOT recurred on this fuller
//      8-campaign, full-month pull — every campaign's account-level
//      "website conversions" figure now matches its own per-rule breakdown
//      exactly (1,155 total either way). The discrepancy card is updated to
//      report this resolution rather than continuing to flag it as open.
//
// Fresh totals (LinkedIn Ads MCP pull, account 509501623, 01-31 Jul 2026):
//   TOTAL_REG=1,106  TOTAL_SPEND=£1,765.56  TOTAL_IMPR=265,920  TOTAL_CLICKS=4,502
//   BLENDED_CPA=£1.60  (ATC=21, Checkout=17 — unchanged from the original pull)

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
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
// Rows 10 and 12 are currently zero-length (cleared via deleteText+no-insert
// when the campaigns weren't found) — deleteText on an already-empty cell
// 400s ("startIndex 0 must be less than endIndex 0"), same bug hit on the
// B2B Reach table's inserted rows. insertText only, no delete.
function newCell(rowIndex, columnIndex, text) {
  return [{ insertText: { objectId: TABLE, cellLocation: { rowIndex, columnIndex }, insertionIndex: 0, text } }];
}

function imageBox(xPx, yPx, wPx, hPx) {
  return {
    size: { width: { magnitude: wPx * PX, unit: 'EMU' }, height: { magnitude: hPx * PX, unit: 'EMU' } },
    transform: { scaleX: 1, scaleY: 1, translateX: xPx * PX, translateY: yPx * PX, unit: 'EMU' },
  };
}

const requests = [
  // ---- Cover ----
  ...edit('cov_cardbody', 'Eight B2C campaigns drove 1,106 Registers from £1,765.56 spend (£1.60 blended CPA) across Cert-OT, Cybersecurity, Data Analytics, and Cert-AI — which launched July 24 and is still ramping.'),

  // ---- Exec Summary ----
  ...edit('g3f483bfcec4_0_27', 'In July, eight B2C campaigns ran across four certificate lines — Cert-OT Phase 3 engagement (Traffic + Views), Cybersecurity Cold Layer (Single Image + Video), Data Analytics Cold Layer (Single Image + Video), and Cert-AI Phase 3 (Traffic + Views, launched July 24) — driving 1,106 Register actions from £1,765.56 spend, a £1.60 blended CPA, plus 21 Add To Cart and 17 Proceed Checkout actions further down the funnel.\n'),

  // ---- What We Did (4 items) ----
  ...edit('g3f483bfcec4_0_35', 'The Single Image Traffic campaign delivered 345 Register actions from £351.04 spend — a £1.02 CPA, the most efficient result across all eight campaigns, at a 2.52% CTR.\n\n'),
  ...edit('g3f483bfcec4_0_39', 'The Views (Video) campaign delivered 247 Register actions from £405.93 spend — £1.64 CPA — a real but smaller gap than seen elsewhere in the account.\n\n'),
  ...edit('g3f483bfcec4_0_43', 'Across four Cold Layer campaigns, CPA ranged from £1.35 (Data Analytics Single Image) to £2.52 (Cybersecurity Single Image) — a 1.9x spread on the same Website Visits objective.\n\n'),
  ...edit('g3f483bfcec4_0_46', 'Cert-AI Phase 3 launched July 24 at a high early CPA.\n'),
  ...edit('g3f483bfcec4_0_47', 'Both campaigns went live with 8 days of delivery — £15.82 (Traffic) and £6.65 (Views) CPA, well above account average but expected for a fresh launch.\n\n'),

  // ---- North Star narrative ----
  ...edit('g3f483bfcec4_0_106', 'July drove 1,106 Register actions at a £1.60 blended CPA — Cert-OT Product Traffic’s £1.02 CPA is the strongest result, while Cert-AI Phase 3 Traffic’s £15.82 CPA (its first 8 days live) is the weakest. A smaller 21 Add To Cart and 17 Proceed Checkout actions sit further down the funnel.\n\n'),
  ...edit('b2c_ns_value2', '£1.60\n'),
  ...edit('g3f60cf2e19e_0_2', '1,106\n'), // "JULY REGISTERS" value — object was renamed/regrouped since the deck was first built; current live objectId confirmed via inspect20.mjs

  // ---- What We Recommend Next ----
  ...edit('g3f483bfcec4_0_120', 'Running Cert-OT Product Traffic (Single Image) at current budget — £1.02 CPA is the strongest result in the account this month.\n\n'),
  ...edit('g3f483bfcec4_0_126', 'Maintaining all four Cybersecurity and Data Analytics Cold Layer campaigns — CPA holding in a £1.35–£2.52 band.\n\n'),
  ...edit('g3f483bfcec4_0_129', 'Running Cert-AI Phase 3 (Traffic + Video) at initial budget since its July 24 launch — too early to scale.\n\n'),
  ...edit('g3f483bfcec4_0_134', 'Confirm whether Cert-AI’s early £15.82/£6.65 CPA is expected for a cold launch, or a signal to revisit creative before scaling spend.\n\n'),
  ...edit('g3f483bfcec4_0_137', 'Confirm whether Cybersecurity Single Image’s £2.52 CPA (the weakest of the six longer-running campaigns) should be paused in favour of the stronger Data Analytics Single Image result.\n\n'),
  ...edit('g3f483bfcec4_0_143', 'Share any planned August budget shift across the four cert lines so next month’s targets can be set accordingly.\n'),
  ...edit('g3f483bfcec4_0_146', 'Confirm the planned Cert-AI budget ramp for August now that it’s live and reporting.\n\n'),
  ...edit('g3f59aaca611_0_5', 'Confirm Register remains the primary KPI going forward — it now reconciles exactly with the account-level total this month (1,155 both ways; see appendix).\n\n'),

  // ---- Reach ----
  ...edit('p3_i17', 'Cert-OT reached an estimated 41,101 (Traffic) and 17,201 (Views) unique members in July; Cybersecurity and Data Analytics Cold Layer reached 10,929–15,091 each; Cert-AI added 1,557 and 831 since its 24 Jul launch, per LinkedIn’s approximate reach metric.\n'),
  ...edit('p3_i20', '265,920\n'),

  // ---- Experiments (Cert-OT vs Views + Cyber/DA cards) ----
  ...edit('g3f483bfcec4_0_269', 'Confirmed a real gap. Traffic (Single Image) delivered 345 Registers at £1.02 CPA; Views (Video) delivered 247 Registers at £1.64 CPA — a 1.6x difference favouring the static format.\n\n'),
  ...edit('g3f483bfcec4_0_280', 'Data Analytics held a lower CPA on Single Image (£1.35) than Cybersecurity (£2.52), while Video CPA was closer (£1.90 vs £1.65) — Single Image performance isn’t consistent across products.\n\n'),

  // ---- Cert-AI card (slide 8) — was "not found", now reports first results ----
  ...edit('g3f59aaca611_0_20', 'Cert-AI Phase 3 — Live From July 24\n'),
  ...edit('g3f59aaca611_0_22', 'This card reports the campaign pair’s first-week performance now that it’s live and reporting, rather than testing a hypothesis.\n'),
  ...edit('g3f59aaca611_0_25', 'Cert-AI Phase 3 Traffic and Views both launched 24 July — 8 days of delivery so far, driving 4 and 9 Register actions at £15.82 and £6.65 CPA respectively.\n\n'),
  ...edit('g3f59aaca611_0_28', 'Re-check CPA once a full month of delivery is available before deciding whether to scale or pause either format.\n\n'),

  // ---- Discrepancy card (slide 8) — the 817-vs-1,141 gap has not recurred ----
  ...edit('g3f59aaca611_0_31', 'Website Conversions — Two Totals, Now Reconciled\n'),
  ...edit('g3f59aaca611_0_33', 'LinkedIn’s account-level "website conversions" stat and a per-rule conversion breakdown should describe the same underlying activity — for July’s full 8-campaign dataset they now match exactly at 1,155 each.\n'),
  ...edit('g3f59aaca611_0_36', 'Every campaign’s account-level conversion figure matches its own Register + Add To Cart + Proceed Checkout + smaller-action breakdown exactly. The 817-vs-1,141 gap seen on an earlier partial-month pull hasn’t recurred.\n\n'),
  ...edit('g3f59aaca611_0_39', 'No further action needed here — continue using Register as the primary KPI given it names the specific action and now reconciles cleanly.\n\n'),

  // ---- Appendix table: refresh existing rows, fill Cert-AI's two reserved rows ----
  ...cell(3, 2, '£351.04'), ...cell(3, 3, '120,991'), ...cell(3, 4, '3,043'), ...cell(3, 5, '2.52%'), ...cell(3, 6, '£2.90'), ...cell(3, 7, '£0.12'),
  ...cell(4, 2, '£405.93'), ...cell(4, 3, '23,352'), ...cell(4, 4, '772'), ...cell(4, 5, '3.31%'), ...cell(4, 6, '£17.38'), ...cell(4, 7, '£0.53'),
  ...cell(5, 2, '£226.68'), ...cell(5, 3, '22,327'), ...cell(5, 4, '147'), ...cell(5, 5, '0.66%'), ...cell(5, 6, '£10.15'), ...cell(5, 7, '£1.54'),
  ...cell(6, 2, '£213.41'), ...cell(6, 3, '34,803'), ...cell(6, 4, '139'), ...cell(6, 5, '0.40%'), ...cell(6, 6, '£6.13'), ...cell(6, 7, '£1.54'),
  ...cell(8, 2, '£226.03'), ...cell(8, 3, '31,886'), ...cell(8, 4, '121'), ...cell(8, 5, '0.38%'), ...cell(8, 6, '£7.09'), ...cell(8, 7, '£1.87'),
  ...cell(9, 2, '£219.38'), ...cell(9, 3, '30,106'), ...cell(9, 4, '211'), ...cell(9, 5, '0.70%'), ...cell(9, 6, '£7.29'), ...cell(9, 7, '£1.04'),

  ...newCell(10, 0, 'Cert-AI\n'), ...newCell(10, 1, 'Product Traffic (Single Image)\n'), ...newCell(10, 2, '£63.28\n'), ...newCell(10, 3, '1,577\n'), ...newCell(10, 4, '37\n'), ...newCell(10, 5, '2.35%\n'), ...newCell(10, 6, '£40.13\n'), ...newCell(10, 7, '£1.71\n'),

  ...cell(11, 0, 'Cert-AI (Phase 3)\n'),

  ...newCell(12, 0, 'Cert-AI\n'), ...newCell(12, 1, 'Product Views (Video)\n'), ...newCell(12, 2, '£59.81\n'), ...newCell(12, 3, '878\n'), ...newCell(12, 4, '32\n'), ...newCell(12, 5, '3.64%\n'), ...newCell(12, 6, '£68.12\n'), ...newCell(12, 7, '£1.87\n'),

  ...cell(13, 2, '£1,765.56'), ...cell(13, 3, '265,920'), ...cell(13, 4, '4,502'),
];

console.log('applying', requests.length, 'requests...');
await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
console.log('text + table updates applied');

// ---- Swap the North Star trend chart image for the refreshed July bar ----
const { data: uploaded } = await drive.files.create({
  requestBody: { name: 'b2c_northstar_chart_v2.png', parents: [SHARED_DRIVE] },
  media: { mimeType: 'image/png', body: fs.createReadStream('assets/b2c_chart_v2.png') },
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
  presentationId: B2C_DECK,
  requestBody: {
    requests: [
      { deleteObject: { objectId: 'b2c_ns_chart_img' } },
      {
        createImage: {
          objectId: 'b2c_ns_chart_img',
          url: chartUrl,
          elementProperties: { pageObjectId: 'g3f483bfcec4_0_95', ...imageBox(140, 460, 634, 274) },
        },
      },
    ],
  },
});
console.log('North Star chart image refreshed');
