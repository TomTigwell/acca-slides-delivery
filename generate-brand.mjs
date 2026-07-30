import { google } from 'googleapis';

// One-off: build the ACCA Brand July 2026 report on the SAME template as the
// Product master (the user's manually-edited, now-canonical deck), rather
// than the old separate 12-slide Brand master. Content is grounded in a
// real LinkedIn Ads pull for account 509501623, filtered to campaign names
// containing "VCF" and "eZine" (the user's stated filter). No campaign
// matching "ACCA Learning Always On" was found anywhere in the account's 27
// campaigns — flagged in the deck itself rather than invented.

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

// A few objects on the Reach slide carry over legacy fontSize values that
// nearly exactly fill their box height (e.g. a 19.178pt value in a 19pt-tall
// box) — fine for the master's original short text, but with zero headroom
// it renders as visibly overlapping its neighbour once re-exported. Shrink
// just these to restore clearance without touching box geometry.
function fontSizeFix(objectId, pt) {
  return [{ updateTextStyle: { objectId, style: { fontSize: { magnitude: pt, unit: 'PT' } }, textRange: { type: 'ALL' }, fields: 'fontSize' } }];
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

// ---- Real July data, account 509501623 (pulled via LinkedIn Ads MCP) ------
const VCF_APAC = { name: 'APAC Lead Generation', spend: 1777.48, impr: 87514, clicks: 613, leads: 10 };
const VCF_AFRICA = { name: 'Africa Lead Generation', spend: 1181.47, impr: 83432, clicks: 746, leads: 25 };
const VCF_BOOST = { name: 'BOOST_VCF_APAC_JULY_2026_INHOUSE', spend: 1000.00, impr: 141897, clicks: 2668, leads: 0 };
const TOTAL_LEADS = VCF_APAC.leads + VCF_AFRICA.leads;
const TOTAL_LEAD_SPEND = VCF_APAC.spend + VCF_AFRICA.spend;
const BLENDED_CPL = (TOTAL_LEAD_SPEND / TOTAL_LEADS).toFixed(2);
const TOTAL_SPEND = (VCF_APAC.spend + VCF_AFRICA.spend + VCF_BOOST.spend).toFixed(2);
const TOTAL_IMPR = VCF_APAC.impr + VCF_AFRICA.impr + VCF_BOOST.impr;
const TOTAL_CLICKS = VCF_APAC.clicks + VCF_AFRICA.clicks + VCF_BOOST.clicks;
const TOTAL_CTR = ((TOTAL_CLICKS / TOTAL_IMPR) * 100).toFixed(2);
const TOTAL_CPM = ((parseFloat(TOTAL_SPEND) / TOTAL_IMPR) * 1000).toFixed(2);
const TOTAL_CPC = (parseFloat(TOTAL_SPEND) / TOTAL_CLICKS).toFixed(2);

const EDITS = [
  // Cover
  { objectId: 'cov_dept', text: 'BRAND & EVENTS  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF' },
  { objectId: 'cov_headline', text: 'July 2026\nPerformance Brief' },
  { objectId: 'cov_subtitle', text: 'Brand & Events · LinkedIn Account 509501623 · July 2026' },
  { objectId: 'cov_cardlabel', text: 'JULY SNAPSHOT' },
  { objectId: 'cov_cardbody', text: `VCF Lead Generation delivered ${TOTAL_LEADS} leads in July — 25 from Africa at £47.26 CPL and 10 from APAC at £177.75 CPL — while eZine's three static campaigns stayed fully paused with no spend or delivery this month.` },
  { objectId: 'cov_footer', text: 'PREPARED BY FILL MY FUNNEL  ·  JULY 2026' },

  // Slide 2 — Exec Summary
  { objectId: 'g3f483bfcec4_0_27', text: 'In July, VCF Lead Generation ran two active campaigns — Africa and APAC — delivering 35 leads combined from £2,958.95 spend, plus a £1,000 organic boost post reaching 141,897 impressions. eZine’s three static campaigns remained paused throughout the month with zero delivery. No activity was found under "ACCA Learning Always On" on this account.\n' },

  // Slide 3 — What We Did (4 items)
  { objectId: 'g3f483bfcec4_0_34', text: 'VCF Africa Lead Generation was the standout performer.\n' },
  { objectId: 'g3f483bfcec4_0_35', text: 'The Africa campaign delivered 25 leads from £1,181.47 spend — a £47.26 CPL, less than a third of APAC’s cost per lead on the same format.\n\n' },
  { objectId: 'g3f483bfcec4_0_38', text: 'VCF APAC Lead Generation ran at a materially higher CPL.\n' },
  { objectId: 'g3f483bfcec4_0_39', text: 'APAC delivered 10 leads from £1,777.48 spend — £177.75 CPL, 3.7x Africa’s rate on the same objective and creative approach.\n\n' },
  { objectId: 'g3f483bfcec4_0_42', text: 'eZine’s static campaigns stayed paused all month.\n' },
  { objectId: 'g3f483bfcec4_0_43', text: 'All three eZine campaigns (Clusters 2, 3, and 4) recorded zero impressions, clicks, or spend in July — no delivery to report.\n\n' },
  { objectId: 'g3f483bfcec4_0_46', text: '"ACCA Learning Always On" could not be located in this account.\n' },
  { objectId: 'g3f483bfcec4_0_47', text: 'No campaign matching that name was found across all 27 campaigns under account 509501623 for July — flagged for confirmation rather than reported on with invented figures.\n\n' },

  // Slide 4 — North Star narrative (chart itself is replaced separately below)
  { objectId: 'g3f483bfcec4_0_106', text: `July delivered ${TOTAL_LEADS} VCF leads at a blended £${BLENDED_CPL} CPL — Africa’s £47.26 CPL is pulling the blend down against APAC’s £177.75. eZine remained fully paused all month; reactivating even one cluster would add incremental reach without touching the VCF budget.\n\n` },

  // Slide 5 — What We Recommend Next. LEFT column ("WHAT WE'RE DOING") is
  // 120/123/126/129/140 (five slots, all x≈118); RIGHT column ("WHAT WE
  // NEED FROM YOU") is 134/137/143/146/g3f59aaca611_0_5 (x≈829) — confirmed
  // via inspect11.mjs after 140 was originally (and wrongly) treated as a
  // right-column slot, which put a "need"-phrased ask inside the doing box.
  { objectId: 'g3f483bfcec4_0_120', text: 'Running VCF Africa Lead Generation at current budget — £47.26 CPL is the most efficient result in the account this month.\n\n' },
  { objectId: 'g3f483bfcec4_0_123', text: 'Running VCF APAC Lead Generation alongside Africa — monitoring whether CPL narrows as the audience matures.\n\n' },
  { objectId: 'g3f483bfcec4_0_126', text: 'Maintaining the BOOST_VCF_APAC_JULY_2026_INHOUSE organic post — 141,897 impressions at £7.05 CPM, outside the lead-gen budget.\n\n' },
  { objectId: 'g3f483bfcec4_0_129', text: 'Holding eZine’s three static campaigns paused pending a reactivation decision.\n\n' },
  { objectId: 'g3f483bfcec4_0_140', text: 'Monitoring the Africa vs APAC CPL gap monthly to confirm whether it holds as spend scales.\n\n' },
  { objectId: 'g3f483bfcec4_0_134', text: 'Confirm whether "ACCA Learning Always On" sits under a different account or exact name — no matching campaign was found this month.\n\n' },
  { objectId: 'g3f483bfcec4_0_137', text: 'Decide whether to reactivate an eZine cluster as a test, or reallocate that budget toward Africa Lead Generation.\n\n' },
  { objectId: 'g3f483bfcec4_0_143', text: 'Confirm whether the APAC/Africa CPL gap reflects audience saturation or a creative/targeting difference worth testing — share July’s VCF lead list so we can check lead quality against it.\n' },
  { objectId: 'g3f483bfcec4_0_146', text: 'Provide any planned July/August budget shift between VCF and eZine so next month’s targets can be set accordingly.\n\n' },
  { objectId: 'g3f59aaca611_0_5', text: 'Confirm the "ACCA Learning Always On" account/name before it can be reported on in future briefs.\n\n' },

  // Slide 6 — Reach
  { objectId: 'p3_i16', text: 'Brand Audience Reach  ·  July 2026  ·  LinkedIn \n' },
  { objectId: 'p3_i17', text: 'VCF’s Africa and APAC Lead Generation campaigns reached an estimated 10,368 and 22,518 unique members respectively in July, per LinkedIn’s approximate member reach metric. eZine added no reach this month.\n' },
  { objectId: 'p3_i19', text: 'LI IMPRESSIONS (JULY)\n' },
  { objectId: 'p3_i20', text: `${TOTAL_IMPR.toLocaleString()}\n` },
  { objectId: 'p3_i21', text: 'VCF + eZine · Jul 2026\n' },
  { objectId: 'g3f59aaca611_0_7', text: 'JULY LEADS\n' },
  { objectId: 'g3f59aaca611_0_8', text: `${TOTAL_LEADS}\n` },
  { objectId: 'g3f59aaca611_0_9', text: 'VCF Africa + APAC\n' },

  // Slide 7 — Experiments intro + 2 cards
  { objectId: 'g3f483bfcec4_0_261', text: 'Two findings this month: a confirmed CPL gap between VCF’s Africa and APAC Lead Generation campaigns, and a confirmed read on the organic boost post’s reach efficiency versus paid lead-gen spend.\n' },
  { objectId: 'g3f483bfcec4_0_264', text: 'VCF Africa vs APAC Lead Generation — Cost Efficiency by Region\n' },
  { objectId: 'g3f483bfcec4_0_266', text: 'Running the same Lead Generation format and creative in Africa and APAC would produce comparable CPL, since both target similar HR/Finance personas on the same objective.\n' },
  { objectId: 'g3f483bfcec4_0_269', text: 'Confirmed a real gap, not comparable. Africa delivered 25 leads at £47.26 CPL; APAC delivered 10 leads at £177.75 CPL — a 3.7x difference on the same format and objective.\n\n' },
  { objectId: 'g3f483bfcec4_0_272', text: 'August EOM: test whether reallocating APAC budget toward Africa, or refining APAC’s audience/creative, narrows the gap.\n\n' },
  { objectId: 'g3f483bfcec4_0_275', text: 'BOOST_VCF_APAC_JULY_2026_INHOUSE — Organic Reach vs Lead Generation Spend\n' },
  { objectId: 'g3f483bfcec4_0_277', text: 'A £1,000 organic boost post could deliver meaningful incremental reach at a lower CPM than the paid Lead Generation campaigns, without competing for lead-gen budget.\n' },
  { objectId: 'g3f483bfcec4_0_280', text: 'Confirmed. The boost delivered 141,897 impressions at £7.05 CPM — below both Lead Gen campaigns’ effective CPM (APAC £20.31, Africa £14.16) — though it drove website conversions, not native leads, a different objective.\n\n' },
  { objectId: 'g3f483bfcec4_0_283', text: 'August EOM: evaluate a second boost post, timed with the Africa campaign, for incremental reach without cannibalising lead-gen spend.\n\n' },

  // Slide 8 — New cards: eZine + the missing-campaign flag
  { objectId: 'g3f59aaca611_0_20', text: 'eZine Reactivation — Is Paused Reach Recoverable?\n' },
  { objectId: 'g3f59aaca611_0_22', text: 'eZine’s three static campaigns (Clusters 2, 3, 4) still show residual website conversions (3, 1, and 11 respectively) despite zero July delivery — worth checking whether reactivation would recover meaningful reach.\n' },
  { objectId: 'g3f59aaca611_0_25', text: 'eZine has been fully paused all month — no impressions, clicks, or spend recorded. The 15 combined website conversions across the three clusters are attribution carryover from before they were paused, not July delivery.\n\n' },
  { objectId: 'g3f59aaca611_0_28', text: 'August EOM: decide whether to reactivate one eZine cluster as a test, or retire the format in favour of VCF Lead Generation.\n\n' },
  { objectId: 'g3f59aaca611_0_31', text: '"ACCA Learning Always On" — Confirm Account or Naming\n' },
  { objectId: 'g3f59aaca611_0_33', text: 'This card flags an open item rather than tests a hypothesis: this campaign name may sit under a different account or a different exact naming convention.\n' },
  { objectId: 'g3f59aaca611_0_36', text: 'No campaign matching "ACCA Learning Always On" was found under account 509501623 for July, checked across all 27 campaigns on the account, including a substring search for "learning" and "always."\n\n' },
  { objectId: 'g3f59aaca611_0_39', text: 'Confirm the correct account ID or exact campaign name before this can be reported on with real figures.\n\n' },
];

const TABLE_EDITS = [
  // Group: VCF — Lead Gen + Organic Boost (reusing the "Cold" 4-row slot; only 3 of 4 used)
  { objectId: 'g3f59aaca611_0_73', rowIndex: 2, columnIndex: 0, text: 'VCF — Lead Gen + Organic Boost' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 0, text: 'VCF' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 1, text: 'APAC Lead Generation' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 2, text: '£1,777.48' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 3, text: '87,514' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 4, text: '613' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 5, text: '0.70%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 6, text: '£20.31' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 7, text: '£2.90' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 0, text: 'VCF' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 1, text: 'Africa Lead Generation' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 2, text: '£1,181.47' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 3, text: '83,432' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 4, text: '746' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 5, text: '0.89%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 6, text: '£14.16' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 7, text: '£1.58' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 0, text: 'VCF' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 1, text: 'BOOST_VCF_APAC_JULY_2026_INHOUSE' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 2, text: '£1,000.00' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 3, text: '141,897' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 4, text: '2,668' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 5, text: '1.88%' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 6, text: '£7.05' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 7, text: '£0.37' },

  // r6 (4th Cold slot): unused this month — clear only, no insert
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 0, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 1, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 2, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 3, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 4, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 5, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 6, text: '' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 7, text: '' },

  // Group: eZine — Static (Paused, no July delivery) (reusing the "Warm" 3-row slot)
  { objectId: 'g3f59aaca611_0_73', rowIndex: 7, columnIndex: 0, text: 'eZine — Static (Paused, no July delivery)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 0, text: 'eZine' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 1, text: 'Cluster 2' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 2, text: '£0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 3, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 4, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 5, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 6, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 7, text: '—' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 0, text: 'eZine' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 1, text: 'Cluster 3 (APAC, S.E.)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 2, text: '£0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 3, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 4, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 5, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 6, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 7, text: '—' },

  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 0, text: 'eZine' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 1, text: 'Cluster 4 (MENA, SSA)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 2, text: '£0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 3, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 4, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 5, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 6, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 7, text: '—' },

  // Group: VCF — Pakistan (Paused, no July delivery) (reusing the "Hot" 1-row slot)
  { objectId: 'g3f59aaca611_0_73', rowIndex: 11, columnIndex: 0, text: 'VCF — Pakistan (Paused, no July delivery)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 0, text: 'VCF' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 1, text: 'Pakistan (2 campaigns)' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 2, text: '£0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 3, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 4, text: '0' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 5, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 6, text: '—' },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 7, text: '—' },

  // Total row
  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 2, text: `£${Number(TOTAL_SPEND).toLocaleString(undefined, { minimumFractionDigits: 2 })}` },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 3, text: TOTAL_IMPR.toLocaleString() },
  { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 4, text: TOTAL_CLICKS.toLocaleString() },
];

async function main() {
  const { data: copy } = await drive.files.copy({
    fileId: TEMPLATE_MASTER,
    supportsAllDrives: true,
    requestBody: { name: 'ACCA Brand — July 2026', parents: [SHARED_DRIVE] },
  });
  const presentationId = copy.id;

  // Locate the North Star chart element's exact bounds on THIS copy (it still
  // points at the Product master's checkout spreadsheet, which would be
  // wrong data for Brand — replace with a native stat block at the same
  // position instead of leaving a misleading embedded chart).
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
  // Rendered size = size * transform scale, not size alone — this element's
  // scale is ~169x (30000x18550 EMU base * 169.1658), so skipping the scale
  // factor produced a ~5x3px box and a "width must be > 0" API error.
  const xPx = Math.round(transform.translateX / PX);
  const yPx = Math.round(transform.translateY / PX);
  const wPx = Math.round((size.width.magnitude * transform.scaleX) / PX);
  const hPx = Math.round((size.height.magnitude * transform.scaleY) / PX);

  const requests = [
    ...editRequests(EDITS),
    ...tableEditRequests(TABLE_EDITS),
    { deleteObject: { objectId: chartEl.objectId } },
    ...rectRequests('brand_ns_card', chartSlideId, [xPx, yPx, wPx, hPx], 'ROUND_RECTANGLE', COLOR.card),
    ...textRequests('brand_ns_label1', chartSlideId, [xPx + 30, yPx + 30, wPx / 2 - 45, 24], 'JULY LEADS', {
      fontFamily: 'DM Sans', fontSize: 10, bold: true, color: COLOR.muted,
    }),
    ...textRequests('brand_ns_value1', chartSlideId, [xPx + 30, yPx + 55, wPx / 2 - 45, 60], `${TOTAL_LEADS}`, {
      fontFamily: 'Lora', fontSize: 40, bold: true, color: COLOR.white,
    }),
    ...textRequests('brand_ns_label2', chartSlideId, [xPx + wPx / 2 + 15, yPx + 30, wPx / 2 - 45, 24], 'BLENDED CPL', {
      fontFamily: 'DM Sans', fontSize: 10, bold: true, color: COLOR.muted,
    }),
    ...textRequests('brand_ns_value2', chartSlideId, [xPx + wPx / 2 + 15, yPx + 55, wPx / 2 - 45, 60], `£${BLENDED_CPL}`, {
      fontFamily: 'Lora', fontSize: 40, bold: true, color: COLOR.white,
    }),
    ...textRequests('brand_ns_note', chartSlideId, [xPx + 30, yPx + 210, wPx - 60, hPx - 240], 'VCF Africa + APAC Lead Generation, July 2026. Native stat block — a per-month Sheets-linked chart for Brand is still pending (needs the Sheets API enabled on the automation project).', {
      fontFamily: 'DM Sans', fontSize: 10, color: COLOR.muted, lineSpacing: 130,
    }),
    // Reach-slide overlap fix (see fontSizeFix comment above)
    ...fontSizeFix('p3_i17', 8),
    ...fontSizeFix('p3_i16', 5.5),
    ...fontSizeFix('p3_i20', 14),
    ...fontSizeFix('g3f59aaca611_0_8', 14),
  ];
  if (reachTableEl) requests.push({ deleteObject: { objectId: reachTableEl.objectId } });
  if (reachCaptionId) requests.push({ deleteObject: { objectId: reachCaptionId } });

  await slides.presentations.batchUpdate({ presentationId, requestBody: { requests } });
  console.log('ACCA Brand — July 2026', '→', `https://docs.google.com/presentation/d/${presentationId}/edit`);
}

await main();
