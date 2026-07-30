import { google } from 'googleapis';

// One-off: build the ACCA B2C July 2026 report on the same template as the
// Product master. Content is grounded in a real LinkedIn Ads pull for
// account 509501623, filtered to the 8 campaign names the user gave —
// only 6 of which exist on this account. The two "Cert-AI Phase 3"
// campaigns (Traffic + Views) were not found anywhere in the account's 27
// campaigns, checked by name prefix — flagged in the deck rather than
// invented. "Register" (LinkedIn conversion-rule name "Button Click:
// Register") is used as the primary KPI instead of the account-level
// "website conversions" stat, because the two numbers don't reconcile
// (817 vs 1,141 across these 6 campaigns) — see the appendix experiment
// card for the discrepancy itself, which is flagged rather than resolved.

const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const TEMPLATE_MASTER = '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

const PX = 6350;
const COLOR = {
  card: { red: 0.8314, green: 0.1765, blue: 0.2549 }, // #D42D41, matches the cover card
  white: { red: 1, green: 1, blue: 1 },
  muted: { red: 0.9490, green: 0.6392, blue: 0.7098 }, // #F2A3B5
};

function emuBox(xPx, yPx, wPx, hPx) {
  return {
    size: { width: { magnitude: wPx * PX, unit: 'EMU' }, height: { magnitude: hPx * PX, unit: 'EMU' } },
    transform: { scaleX: 1, scaleY: 1, translateX: xPx * PX, translateY: yPx * PX, unit: 'EMU' },
  };
}

function rectRequests(objectId, slideId, box, shapeType, color) {
  return [
    { createShape: { objectId, shapeType, elementProperties: { pageObjectId: slideId, ...emuBox(...box) } } },
    { updateShapeProperties: { objectId, shapeProperties: { shapeBackgroundFill: { solidFill: { color: { rgbColor: color } } }, outline: { propertyState: 'NOT_RENDERED' } }, fields: 'shapeBackgroundFill.solidFill.color,outline.propertyState' } },
  ];
}

function textRequests(objectId, slideId, box, text, { fontFamily, fontSize, bold = false, color, align = 'START', lineSpacing }) {
  return [
    { createShape: { objectId, shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: slideId, ...emuBox(...box) } } },
    { insertText: { objectId, insertionIndex: 0, text } },
    { updateTextStyle: { objectId, style: { fontFamily, fontSize: { magnitude: fontSize, unit: 'PT' }, bold, foregroundColor: { opaqueColor: { rgbColor: color } } }, textRange: { type: 'ALL' }, fields: 'fontFamily,fontSize,bold,foregroundColor' } },
    { updateParagraphStyle: { objectId, style: { alignment: align, ...(lineSpacing ? { lineSpacing } : {}) }, textRange: { type: 'ALL' }, fields: lineSpacing ? 'alignment,lineSpacing' : 'alignment' } },
    { updateShapeProperties: { objectId, shapeProperties: { outline: { propertyState: 'NOT_RENDERED' } }, fields: 'outline.propertyState' } },
  ];
}

function editRequests(edits) {
  const requests = [];
  for (const { objectId, text } of edits) {
    requests.push({ deleteText: { objectId, textRange: { type: 'ALL' } } });
    requests.push({ insertText: { objectId, insertionIndex: 0, text } });
  }
  return requests;
}

function tableEditRequests(edits) {
  const requests = [];
  for (const { objectId, rowIndex, columnIndex, text } of edits) {
    const cellLocation = { rowIndex, columnIndex };
    requests.push({ deleteText: { objectId, cellLocation, textRange: { type: 'ALL' } } });
    if (text) requests.push({ insertText: { objectId, cellLocation, insertionIndex: 0, text } });
  }
  return requests;
}

// A few reused template objects (Reach-slide stat tiles) carry legacy
// fontSize values that leave near-zero clearance in their box — learned
// from the Brand deck generation, where "312,843"/"35" overlapped their
// labels for exactly this reason. Set these explicitly up front instead of
// needing a follow-up patch.
function fontSizeFix(objectId, pt) {
  return [{ updateTextStyle: { objectId, style: { fontSize: { magnitude: pt, unit: 'PT' } }, textRange: { type: 'ALL' }, fields: 'fontSize' } }];
}

// ---- Real July data, account 509501623 (LinkedIn Ads MCP pull + conversion
// breakdown). "Register" = conversion rule "Button Click: Register", the
// dominant, well-defined action across all 6 campaigns.
const C = {
  traffic: { label: 'Cert-OT Product Traffic (Single Image)', spend: 299.72, impr: 77676, clicks: 2159, reg: 345 },
  views:   { label: 'Cert-OT Product Views (Video)',          spend: 336.25, impr: 19814, clicks: 570,  reg: 247 },
  cyberSI: { label: 'Cybersecurity Cold Layer (Single Image)', spend: 217.04, impr: 20322, clicks: 139,  reg: 89 },
  cyberVid:{ label: 'Cybersecurity Cold Layer (Video)',        spend: 193.65, impr: 31432, clicks: 126,  reg: 129 },
  daVid:   { label: 'Data Analytics Cold Layer (Video)',       spend: 220.60, impr: 30948, clicks: 119,  reg: 119 },
  daSI:    { label: 'Data Analytics Cold Layer (Single Image)',spend: 211.16, impr: 28379, clicks: 199,  reg: 163 },
};
const TOTAL_REG = Object.values(C).reduce((s, c) => s + c.reg, 0);
const TOTAL_SPEND = Object.values(C).reduce((s, c) => s + c.spend, 0);
const TOTAL_IMPR = Object.values(C).reduce((s, c) => s + c.impr, 0);
const TOTAL_CLICKS = Object.values(C).reduce((s, c) => s + c.clicks, 0);
const BLENDED_CPA = (TOTAL_SPEND / TOTAL_REG).toFixed(2);
const fmtMoney = (n) => `£${n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

const EDITS = [
  // Cover
  { objectId: 'cov_dept', text: 'B2C  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF' },
  { objectId: 'cov_headline', text: 'July 2026\nPerformance Brief' },
  { objectId: 'cov_subtitle', text: 'B2C · LinkedIn Account 509501623 · July 2026' },
  { objectId: 'cov_cardlabel', text: 'JULY SNAPSHOT' },
  { objectId: 'cov_cardbody', text: `Six B2C campaigns drove ${TOTAL_REG.toLocaleString()} Registers from ${fmtMoney(TOTAL_SPEND)} spend (£${BLENDED_CPA} blended CPA) across Cert-OT, Cybersecurity, and Data Analytics. Two requested Cert-AI campaigns weren’t found on this account.` },
  { objectId: 'cov_footer', text: 'PREPARED BY FILL MY FUNNEL  ·  JULY 2026' },

  // Slide 2 — Exec Summary
  { objectId: 'g3f483bfcec4_0_27', text: `In July, six B2C campaigns ran across three certificate lines — Cert-OT Phase 3 engagement (Traffic + Views), Cybersecurity Cold Layer (Single Image + Video), and Data Analytics Cold Layer (Single Image + Video) — driving ${TOTAL_REG.toLocaleString()} Register actions from ${fmtMoney(TOTAL_SPEND)} spend, a £${BLENDED_CPA} blended CPA, plus 21 Add To Cart and 17 Proceed Checkout actions further down the funnel. No campaigns matching "Cert-AI Phase 3" (Traffic or Views) were found under account 509501623 for July.\n` },

  // Slide 3 — What We Did (4 items)
  { objectId: 'g3f483bfcec4_0_34', text: 'Cert-OT Product Traffic drove the lowest CPA in the account.\n' },
  { objectId: 'g3f483bfcec4_0_35', text: 'The Single Image Traffic campaign delivered 345 Register actions from £299.72 spend — a £0.87 CPA, the most efficient result across all six campaigns, at a 2.78% CTR.\n\n' },
  { objectId: 'g3f483bfcec4_0_38', text: 'Cert-OT Product Views ran at over 1.5x Traffic’s CPA.\n' },
  { objectId: 'g3f483bfcec4_0_39', text: 'The Views (Video) campaign delivered 247 Register actions from £336.25 spend — £1.36 CPA — a real but smaller gap than seen elsewhere in the account.\n\n' },
  { objectId: 'g3f483bfcec4_0_42', text: 'Cybersecurity and Data Analytics Cold Layer held a wide CPA range.\n' },
  { objectId: 'g3f483bfcec4_0_43', text: 'Across four Cold Layer campaigns, CPA ranged from £1.30 (Data Analytics Single Image) to £2.44 (Cybersecurity Single Image) — a 1.9x spread on the same Website Visits objective.\n\n' },
  { objectId: 'g3f483bfcec4_0_46', text: '"Cert-AI Phase 3" (Traffic and Views) could not be located in this account.\n' },
  { objectId: 'g3f483bfcec4_0_47', text: 'No campaigns matching either name were found across all 27 campaigns under account 509501623 for July — flagged for confirmation rather than reported on with invented figures.\n\n' },

  // Slide 4 — North Star narrative (chart itself is replaced separately below)
  { objectId: 'g3f483bfcec4_0_106', text: `July drove ${TOTAL_REG.toLocaleString()} Register actions at a £${BLENDED_CPA} blended CPA — Cert-OT Product Traffic’s £0.87 CPA is the strongest result, while Cybersecurity Single Image’s £2.44 CPA is the weakest. A smaller 21 Add To Cart and 17 Proceed Checkout actions sit further down the funnel. Cert-AI campaigns are absent from this account and can’t be included until confirmed.\n\n` },

  // Slide 5 — What We Recommend Next. LEFT column ("WHAT WE'RE DOING") is
  // 120/123/126/129/140 (five slots, all x≈118); RIGHT column ("WHAT WE
  // NEED FROM YOU") is 134/137/143/146/g3f59aaca611_0_5 (x≈829) — confirmed
  // via inspect11.mjs after 140 was originally (and wrongly) treated as a
  // right-column slot, which put a "need"-phrased ask inside the doing box.
  { objectId: 'g3f483bfcec4_0_120', text: 'Running Cert-OT Product Traffic (Single Image) at current budget — £0.87 CPA is the strongest result in the account this month.\n\n' },
  { objectId: 'g3f483bfcec4_0_123', text: 'Running Cert-OT Product Views (Video) alongside Traffic — monitoring whether CPA narrows as video creative matures.\n\n' },
  { objectId: 'g3f483bfcec4_0_126', text: 'Maintaining all four Cybersecurity and Data Analytics Cold Layer campaigns — CPA holding in a £1.30–£2.44 band.\n\n' },
  { objectId: 'g3f483bfcec4_0_129', text: 'Holding budget steady pending confirmation of the Cert-AI campaigns.\n\n' },
  { objectId: 'g3f483bfcec4_0_140', text: 'Monitoring the Cert-OT Traffic vs Views CPA gap monthly to confirm whether it holds as spend scales.\n\n' },
  { objectId: 'g3f483bfcec4_0_134', text: 'Confirm whether "Cert-AI Phase 3" campaigns sit under a different account or exact naming convention — neither Traffic nor Views was found this month.\n\n' },
  { objectId: 'g3f483bfcec4_0_137', text: 'Confirm whether Cybersecurity Single Image’s £2.44 CPA (the weakest of the six) should be paused in favour of the stronger Data Analytics Single Image result.\n\n' },
  { objectId: 'g3f483bfcec4_0_143', text: 'Share any planned August budget shift across the three cert lines so next month’s targets can be set accordingly.\n' },
  { objectId: 'g3f483bfcec4_0_146', text: 'Provide the correct Cert-AI campaign names or IDs so August’s report can include them.\n\n' },
  { objectId: 'g3f59aaca611_0_5', text: 'Confirm which of the two conflicting conversion totals (817 vs 1,141 — see appendix) is authoritative, and which funnel action (Register, Add To Cart, or Proceed Checkout) should be the primary KPI going forward.\n\n' },

  // Slide 6 — Reach
  { objectId: 'p3_i16', text: 'B2C Audience Reach  ·  July 2026  ·  LinkedIn \n' },
  { objectId: 'p3_i17', text: 'Cert-OT reached an estimated 27,668 (Traffic) and 15,022 (Views) unique members in July; Cybersecurity and Data Analytics Cold Layer campaigns reached a further 10,358–14,296 each, per LinkedIn’s approximate member reach metric.\n' },
  { objectId: 'p3_i19', text: 'LI IMPRESSIONS (JULY)\n' },
  { objectId: 'p3_i20', text: `${TOTAL_IMPR.toLocaleString()}\n` },
  { objectId: 'p3_i21', text: 'B2C · Jul 2026\n' },
  { objectId: 'g3f59aaca611_0_7', text: 'JULY REGISTERS\n' },
  { objectId: 'g3f59aaca611_0_8', text: `${TOTAL_REG.toLocaleString()}\n` },
  { objectId: 'g3f59aaca611_0_9', text: 'Six B2C campaigns\n' },

  // Slide 7 — Experiments intro + 2 cards
  { objectId: 'g3f483bfcec4_0_261', text: 'Two findings this month: a confirmed CPA gap between Cert-OT’s Traffic and Views formats, and a wider CPA spread across the four Cybersecurity/Data Analytics Cold Layer campaigns than expected.\n' },
  { objectId: 'g3f483bfcec4_0_264', text: 'Cert-OT Traffic vs Views — Format Efficiency\n' },
  { objectId: 'g3f483bfcec4_0_266', text: 'Running the same Cert-OT Phase 3 engagement objective as both a Single Image Traffic campaign and a Video Views campaign would produce comparable CPA, since both target the same Core Global ICP.\n' },
  { objectId: 'g3f483bfcec4_0_269', text: 'Confirmed a real gap. Traffic (Single Image) delivered 345 Registers at £0.87 CPA; Views (Video) delivered 247 Registers at £1.36 CPA — a 1.6x difference favouring the static format.\n\n' },
  { objectId: 'g3f483bfcec4_0_272', text: 'August EOM: test whether shifting incremental budget from Views to Traffic improves blended CPA without losing video-specific signal.\n\n' },
  { objectId: 'g3f483bfcec4_0_275', text: 'Cybersecurity vs Data Analytics — Cold Layer Format Comparison\n' },
  { objectId: 'g3f483bfcec4_0_277', text: 'Running identical Single Image + Video pairs across two certs (Cybersecurity, Data Analytics) on the same Cold Layer clusters would produce comparable per-cert CPA.\n' },
  { objectId: 'g3f483bfcec4_0_280', text: 'Data Analytics held a lower CPA on Single Image (£1.30) than Cybersecurity (£2.44), while Video CPA was closer (£1.85 vs £1.50) — Single Image performance isn’t consistent across products.\n\n' },
  { objectId: 'g3f483bfcec4_0_283', text: 'August EOM: test whether Data Analytics’ Single Image creative approach transfers to Cybersecurity to close its £2.44 CPA gap.\n\n' },

  // Slide 8 — New cards: missing campaigns + a real data-integrity flag
  { objectId: 'g3f59aaca611_0_20', text: 'Cert-AI Phase 3 — Confirm Account or Naming\n' },
  { objectId: 'g3f59aaca611_0_22', text: 'This card flags an open item rather than tests a hypothesis: this campaign pair may sit under a different account or a different exact naming convention.\n' },
  { objectId: 'g3f59aaca611_0_25', text: 'No campaigns matching "Cert-AI Phase 3 Traffic" or "Cert-AI Phase 3 Views" were found under account 509501623 for July, checked across all 27 campaigns on the account.\n\n' },
  { objectId: 'g3f59aaca611_0_28', text: 'Confirm the correct account ID or exact campaign names before this can be reported on with real figures.\n\n' },
  { objectId: 'g3f59aaca611_0_31', text: 'Website Conversions — Two Totals That Don’t Reconcile\n' },
  { objectId: 'g3f59aaca611_0_33', text: 'LinkedIn’s account-level "website conversions" stat (817 for these six campaigns) and a per-rule conversion breakdown (1,092 Register + 21 Add To Cart + 17 Proceed Checkout + smaller actions ≈ 1,141) should describe the same underlying activity.\n' },
  { objectId: 'g3f59aaca611_0_36', text: 'The two totals don’t match. This report uses the per-rule breakdown (Register as the primary KPI) since it names the specific action, but the 817-vs-1,141 gap itself hasn’t been resolved — likely a counting or attribution difference between LinkedIn’s aggregate stats and its per-rule breakdown.\n\n' },
  { objectId: 'g3f59aaca611_0_39', text: 'Confirm with LinkedIn/analytics which total is authoritative before treating either as the definitive conversion count.\n\n' },
];

const TABLE_EDITS = [
  { objectId: 'g3f59aaca611_0_73', rowIndex: 2, columnIndex: 0, text: 'Cert-OT (Phase 3) + Cybersecurity (Cold Layer)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 0, text: 'Cert-OT' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 1, text: 'Product Traffic (Single Image)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 2, text: '£299.72' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 3, text: '77,676' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 4, text: '2,159' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 5, text: '2.78%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 6, text: '£3.86' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 7, text: '£0.14' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 0, text: 'Cert-OT' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 1, text: 'Product Views (Video)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 2, text: '£336.25' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 3, text: '19,814' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 4, text: '570' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 5, text: '2.88%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 6, text: '£16.97' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 7, text: '£0.59' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 0, text: 'Cyber' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 1, text: 'Cold Layer (Single Image)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 2, text: '£217.04' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 3, text: '20,322' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 4, text: '139' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 5, text: '0.68%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 6, text: '£10.68' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 7, text: '£1.56' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 0, text: 'Cyber' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 1, text: 'Cold Layer (Video)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 2, text: '£193.65' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 3, text: '31,432' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 4, text: '126' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 5, text: '0.40%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 6, text: '£6.16' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 7, text: '£1.54' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 7, columnIndex: 0, text: 'Data Analytics (Cold Layer)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 0, text: 'DataAn' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 1, text: 'Cold Layer (Video)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 2, text: '£220.60' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 3, text: '30,948' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 4, text: '119' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 5, text: '0.38%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 6, text: '£7.13' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 7, text: '£1.85' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 0, text: 'DataAn' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 1, text: 'Cold Layer (Single Image)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 2, text: '£211.16' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 3, text: '28,379' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 4, text: '199' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 5, text: '0.70%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 6, text: '£7.44' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 7, text: '£1.06' },

  // Row 10 (unused 4th slot in this group) and row 12 (Cert-AI slot, no
  // real data) must have EVERY column cleared, not just 0-1 — the master's
  // original row data (e.g. "ProDipSust Sustainability Strategy" £348/
  // 61,477/411/...) otherwise survives untouched in columns 2-7 and leaks
  // into the deck. Caught via a post-generation PDF render, not assumed.
  ...[0, 1, 2, 3, 4, 5, 6, 7].map((c) => ({ objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: c, text: '' })),
  ...[0, 1, 2, 3, 4, 5, 6, 7].map((c) => ({ objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: c, text: '' })),

  { objectId: 'g3f59aaca611_0_73', rowIndex: 11, columnIndex: 0, text: 'Cert-AI (Phase 3) — not found in account' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 2, text: fmtMoney(TOTAL_SPEND) },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 3, text: TOTAL_IMPR.toLocaleString() },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 4, text: TOTAL_CLICKS.toLocaleString() },
];

async function main() {
  const { data: copy } = await drive.files.copy({
    fileId: TEMPLATE_MASTER,
    supportsAllDrives: true,
    requestBody: { name: 'ACCA B2C — July 2026', parents: [SHARED_DRIVE] },
  });
  const presentationId = copy.id;

  // Locate the North Star chart element's exact bounds on THIS copy (it
  // still points at the Product master's checkout spreadsheet, wrong data
  // for B2C — replace with a native stat block at the same position).
  const { data: pres } = await slides.presentations.get({ presentationId });
  let chartEl = null;
  let reachTableEl = null;
  let reachCaptionId = null;
  for (const slide of pres.slides) {
    for (const el of slide.pageElements || []) {
      if (el.sheetsChart) chartEl = el;
      if (el.objectId === 'p3_g5') reachTableEl = el;
      if (el.objectId === 'p3_i30') reachCaptionId = el.objectId;
    }
  }
  if (!chartEl) throw new Error('North Star sheetsChart element not found on the copied deck');

  const chartSlideId = pres.slides.find((s) => (s.pageElements || []).some((el) => el === chartEl)).objectId;
  const size = chartEl.size;
  const transform = chartEl.transform;
  // Rendered size = size * transform scale, not size alone (this element's
  // scale is ~169x) — confirmed against the master via inspect9.mjs.
  const xPx = Math.round(transform.translateX / PX);
  const yPx = Math.round(transform.translateY / PX);
  const wPx = Math.round((size.width.magnitude * transform.scaleX) / PX);
  const hPx = Math.round((size.height.magnitude * transform.scaleY) / PX);

  const requests = [
    ...editRequests(EDITS),
    ...tableEditRequests(TABLE_EDITS),
    { deleteObject: { objectId: chartEl.objectId } },
    ...rectRequests('b2c_ns_card', chartSlideId, [xPx, yPx, wPx, hPx], 'ROUND_RECTANGLE', COLOR.card),
    ...textRequests('b2c_ns_label1', chartSlideId, [xPx + 30, yPx + 30, wPx / 2 - 45, 24], 'JULY REGISTERS', {
      fontFamily: 'DM Sans', fontSize: 10, bold: true, color: COLOR.muted,
    }),
    ...textRequests('b2c_ns_value1', chartSlideId, [xPx + 30, yPx + 55, wPx / 2 - 45, 60], TOTAL_REG.toLocaleString(), {
      fontFamily: 'Lora', fontSize: 40, bold: true, color: COLOR.white,
    }),
    ...textRequests('b2c_ns_label2', chartSlideId, [xPx + wPx / 2 + 15, yPx + 30, wPx / 2 - 45, 24], 'BLENDED CPA', {
      fontFamily: 'DM Sans', fontSize: 10, bold: true, color: COLOR.muted,
    }),
    ...textRequests('b2c_ns_value2', chartSlideId, [xPx + wPx / 2 + 15, yPx + 55, wPx / 2 - 45, 60], `£${BLENDED_CPA}`, {
      fontFamily: 'Lora', fontSize: 40, bold: true, color: COLOR.white,
    }),
    ...textRequests('b2c_ns_note', chartSlideId, [xPx + 30, yPx + 210, wPx - 60, hPx - 240], 'Cert-OT + Cybersecurity + Data Analytics, July 2026. Native stat block — a per-month Sheets-linked chart for B2C is still pending (needs the Sheets API enabled on the automation project).', {
      fontFamily: 'DM Sans', fontSize: 10, color: COLOR.muted, lineSpacing: 130,
    }),
    // Reach-slide stat-tile font fix, applied up front (see fontSizeFix comment above).
    ...fontSizeFix('p3_i17', 8),
    ...fontSizeFix('p3_i16', 5.5),
    ...fontSizeFix('p3_i20', 14),
    ...fontSizeFix('g3f59aaca611_0_8', 14),
  ];
  if (reachTableEl) requests.push({ deleteObject: { objectId: reachTableEl.objectId } });
  if (reachCaptionId) requests.push({ deleteObject: { objectId: reachCaptionId } });

  await slides.presentations.batchUpdate({ presentationId, requestBody: { requests } });
  console.log('ACCA B2C — July 2026', '→', `https://docs.google.com/presentation/d/${presentationId}/edit`);
}

await main();
