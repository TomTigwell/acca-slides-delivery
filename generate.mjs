import { google } from 'googleapis';

const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const MASTERS = {
  // [MASTER] ACCA Product — promoted from the client-edited, HTML-brief-
  // matched deck (the user's manual pass on top of the July delivery).
  // Native cov_* cover shapes and all object IDs below were confirmed to
  // survive drive.files.copy, so this is the canonical reusable template
  // going forward — not the original 15-slide generator output.
  product: '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks',
  b2b:     '1dZJV-dvJaRVUxTCYOoR2QQcyRa2N6LAIV780U8ezqHg',
  brand:   '12WS8qbsgSW8KrOLA9h2A9mjVLnf4aBAQuH4MY1s95P8',
};

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

// Numbers-only July fill. Each entry is {objectId, text} for a single shape
// whose ENTIRE text content is replaced outright — targeted by shape ID
// (via deleteText ALL + insertText), not replaceAllText, because several of
// these exact figures recur elsewhere in each deck (rollup tables, narrative
// sentences) that must NOT change. Object IDs were grounded against a
// read-only Slides API dump of the actual masters (inspect.mjs), not the
// reformatted Drive text export.
//
// Left untouched deliberately: exec-summary prose, Discussion-slide quotes,
// WE WILL/DECISION recommendation cards, North Star quarter-framed stat
// tiles and chart (B2B/Brand) and the Product checkout-initiations
// narrative — these present Q1/H1 as the current story and need a
// strategist's judgment call (real July commentary, or a redrawn chart)
// rather than a mechanical number swap; relabeling just the headline text
// while leaving the bars/target-line as Q4-vs-Q1 would make the slide
// actively wrong, not just incomplete. Company tables and per-region
// ad-set breakdowns are left as-is: no July-level granular breakdown was
// supplied for them, and their section headers ("Q1 FY27 · ...") still
// correctly describe the Q1 data they introduce.
const EDITS = {
  b2b: [
    // Slide 5 — Reach — Employer Audience stat tiles
    { objectId: 'p6_i3', text: 'July 2026' },
    { objectId: 'p6_i5', text: '67,019' },
    { objectId: 'p6_i9', text: '399' },
    { objectId: 'p6_i11', text: 'Avg CTR 0.60%' },
    { objectId: 'p6_i13', text: '10' },
    { objectId: 'p6_i15', text: '£130 blended CPL' },
  ],
  // Grounded against the NEW master (inspect5.mjs dump, run 30087794104):
  // the user's manually-restructured, HTML-brief-matched 9-slide deck. Every
  // entry below is the CURRENT (July 2026) text on that object ID — an
  // identity fill that proves the round-trip, and the array a future editor
  // swaps values in to populate next month's data. Object IDs prefixed
  // g3f483bfcec4_0_* are carried over from the original build; g3f59aaca611_0_*
  // are new content added during the user's manual edit. Left OUT deliberately
  // (structural template scaffolding, not monthly data): section headers
  // ("What We Did", "Hypothesis", "Learning", "WHAT WE'RE DOING", etc.),
  // numbered-item markers, experiment card titles/hypotheses (the experiment
  // itself doesn't change monthly, only its Learning/Action), and the two
  // static Q1 reference tables (p3_g5 company table, p3_i30 header) — no
  // July-level breakdown replaces those, same rationale as the original
  // numbers-only pass.
  product: [
    // ---- Cover ----
    { objectId: 'cov_dept', text: 'PRODUCT MARKETING  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF' },
    { objectId: 'cov_headline', text: 'July 2026\nPerformance Brief' },
    { objectId: 'cov_subtitle', text: 'Product Marketing · LinkedIn Account 509501623 · July 2026' },
    { objectId: 'cov_cardlabel', text: 'JULY SNAPSHOT' },
    { objectId: 'cov_cardbody', text: 'Q2 opens with CertOT live and the persona-creative build underway. July will show whether Apple-UI and the reactivated ProDipSust Sustainability Strategy ad set can hold June’s 127-checkout momentum.' },
    { objectId: 'cov_footer', text: 'PREPARED BY FILL MY FUNNEL  ·  JULY 2026' },

    // ---- Slide 2 — Exec Summary intro ----
    { objectId: 'g3f483bfcec4_0_27', text: 'In July we launched CertOT on LinkedIn, maintained the full product portfolio (ProDipSust, Data Analytics, Cybersecurity) with Apple-UI creatives, reactivated the paused ProDipSust Sustainability Strategy ad set, and began the CERT PFM groundwork — running across LinkedIn and Meta B2C simultaneously across APAC, UK, and EMEA. \n' },

    // ---- Slide 3 — What We Did (4 numbered items) ----
    { objectId: 'g3f483bfcec4_0_34', text: 'CertOT launched on LinkedIn.\n' },
    { objectId: 'g3f483bfcec4_0_35', text: 'Certificate in Organisational Transformation went live on LinkedIn in July per Rashid\'s July 06 EOM update. First week impressions and CTR data are pending — initial signals from Rashid confirm it is primarily a B2B proposition, meaning creative framing must target L&D and HR decision-makers, not individual learners.\n\n' },
    { objectId: 'g3f483bfcec4_0_38', text: 'ProDipSust Sustainability Strategy ad set reactivated.\n' },
    { objectId: 'g3f483bfcec4_0_39', text: 'This ad set was paused in Q1 despite a 0.77% CTR and £0.50 CPC. Reactivation in July recovers the cheapest clicks in the product portfolio. It should take the majority of the ProDipSust budget while testimonial video is tested alongside it.\n\n' },
    { objectId: 'g3f483bfcec4_0_42', text: 'The full-funnel spend model is now understood at senior level.\n' },
    { objectId: 'g3f483bfcec4_0_43', text: 'Trevor and Peter confirmed on the July 06 call that they understand awareness spend does not convert in the same month — they accepted the lagged attribution model and asked us to document the assumed lag. This clears the way to report spend, impressions, and checkouts without having to justify the funnel shape every month.\n\n' },
    { objectId: 'g3f483bfcec4_0_46', text: 'Checkout tracking is still incomplete — Alan\'s action is unblocked.\n' },
    { objectId: 'g3f483bfcec4_0_47', text: 'Trevor committed on the July 06 call to personally push Alan on the checkout-tracking implementation. Once in place, we can report actual sales attributed to paid ads and unlock Google Search as an additional channel. This remains the single most valuable technical action outstanding.\n\n' },

    // ---- Slide 4 — North Star narrative ----
    // NOTE: still literally says "July actuals pending" while the chart's
    // 4th bar (no text content, so not editable via this text-based pass)
    // visually reads ~87 — flagged to the user as an open contradiction.
    { objectId: 'g3f483bfcec4_0_106', text: 'Q1 close: 34 → 87 → 127. Q1 averaged 83 checkouts/month — already above the 80/month target. July actuals pending. New cert (CertOT) and reactivated ad set (ProDipSust Sustainability Strategy) should maintain or extend this trajectory. \n\n' },

    // ---- Slide 5 — What We Recommend Next (LinkedIn) ----
    { objectId: 'g3f483bfcec4_0_120', text: 'Running CertOT on LinkedIn with Apple-UI product creative — building impression base and first CTR read for August EOM.\n\n' },
    { objectId: 'g3f483bfcec4_0_123', text: 'Running the ProDipSust testimonial video test alongside the reactivated Sustainability Strategy static ad set — comparing checkout rates at next EOM.\n\n' },
    { objectId: 'g3f483bfcec4_0_126', text: 'Maintaining Data Analytics and Cybersecurity at current budget — Apple-UI on LinkedIn, video scaling on Meta B2C.\n\n' },
    { objectId: 'g3f483bfcec4_0_129', text: 'Preparing a full creative audit for Q2 direction, including a GIF-format test plan for August.\n\n' },
    { objectId: 'g3f483bfcec4_0_134', text: 'Checkout-tracking implementation — once live, we can report sales attributed to paid ads and unlock Google Search as a channel.\n\n' },
    { objectId: 'g3f483bfcec4_0_137', text: 'Rashid: confirm the CERT PFM timeline for paid social — we need 2–3 weeks to brief Hassan and build creative once confirmed.\n\n' },
    { objectId: 'g3f483bfcec4_0_140', text: 'Adding spend data to the North Star chart and a cert-level logbook to every future Product EOM report, as requested on the July 06 call.\n\n' },
    { objectId: 'g3f483bfcec4_0_143', text: 'Rashid: share the new B2B case study — we want to create snackable LinkedIn assets (document ad or carousel) from it.\n' },
    { objectId: 'g3f483bfcec4_0_146', text: 'Confirm ACCA Global Meta pixel coverage on cert product pages — this unlocks a warm retargeting pool beyond the LMS audience.\n\n' },
    { objectId: 'g3f59aaca611_0_5', text: 'Rashid: share the refreshed ProDipSust assets as soon as they are ready. New exam window is opening soon — we need to be live before it does.\n\n' },

    // ---- Slide 6 — Reach ----
    { objectId: 'p3_i16', text: 'Product Audience Reach  ·  July 2026  ·  LinkedIn \n' },
    { objectId: 'p3_i17', text: 'LinkedIn continues to deliver qualified professional reach into Big Four and global accounting firms. Deloitte and EY lead click volume in Q1. The new CertOT campaign in July extends reach into L&D and HR buying committees.\n' },
    { objectId: 'p3_i19', text: 'LI IMPRESSIONS (Q1)\n' },
    { objectId: 'p3_i20', text: '428K\n' },
    { objectId: 'p3_i21', text: 'All cert ad sets · Apr–Jun\n' },
    { objectId: 'g3f59aaca611_0_7', text: 'LI Q1 Clicks\n' },
    { objectId: 'g3f59aaca611_0_8', text: '2,673\n' },
    { objectId: 'g3f59aaca611_0_9', text: 'All cert ad sets · Apr–Jun\n' },

    // ---- Slide 7 — Experiments intro + 2 carried-forward cards ----
    { objectId: 'g3f483bfcec4_0_261', text: 'Three experiments running in July: ProDipSust testimonial video vs static, Meta video scaling on Cybersecurity, and CertOT first-impression test. Two experiments carry forward from Q1 as confirmed findings. The ProDipSust reactivation ad set (Sustainability Strategy) is the efficiency benchmark for the quarter — watch whether it holds £0.50 CPC at scale.\n' },
    { objectId: 'g3f483bfcec4_0_269', text: 'Confirmed. Data Analytics Single Image: 0.81% CTR. Cybersecurity Single Image: 0.82% CTR. Both 80%+ above the 0.45% benchmark. Rashid and Brinley endorsed this format on the June 11 call. Checkout activity confirmed from Cybersecurity.\n\n' },
    { objectId: 'g3f483bfcec4_0_271', text: 'JULY ACTION\n' },
    { objectId: 'g3f483bfcec4_0_272', text: 'Apple-UI is now the default format for all active certs. CertOT launches with Apple-UI creative in July. GIF-format test against Apple-UI static planned for Q2. Hassan briefed on next creative batch.\n\n' },
    { objectId: 'g3f483bfcec4_0_280', text: 'Rashid shared new ProDipSust testimonial videos on the June 11 call. Both ad formats are now live. Test began in July alongside the reactivated Sustainability Strategy ad set — three ProDipSust variants running simultaneously.\n\n' },
    { objectId: 'g3f483bfcec4_0_282', text: 'JULY ACTION\n' },
    { objectId: 'g3f483bfcec4_0_283', text: 'August EOM. We need minimum 4 weeks of data. Compare: video view rate, CTR, and checkout initiation rate between testimonial video and Apple-UI static. Winner takes majority of ProDipSust budget in Q2.\n\n' },

    // ---- Slide 8 — New experiment cards (CertOT, Hot Layer) ----
    { objectId: 'g3f59aaca611_0_25', text: 'CertOT went live in July per Rashid\'s July 06 update ("OT launched on Friday"). CertOT is primarily a B2B proposition. The Apple-UI format has not yet been tested in a B2B cert context — this is its first run. Audience: L&D, HR, and transformation leaders at target firms.\n\n' },
    { objectId: 'g3f59aaca611_0_27', text: 'JULY ACTION\n' },
    { objectId: 'g3f59aaca611_0_28', text: 'August EOM — first CTR and impression data. If CTR is below 0.45% benchmark, test a B2B-specific headline frame ("Your team\'s transformation starts here") alongside the product UI visual. If above benchmark, scale budget.\n\n' },
    { objectId: 'g3f59aaca611_0_36', text: 'Reactivating the hot layer with proper budget allocation will drive checkout initiations from warm audiences at the lowest CPL in the portfolio, because the custom audience is already the most commercially qualified pool we have.\n\n' },
    { objectId: 'g3f59aaca611_0_38', text: 'JULY ACTION\n' },
    { objectId: 'g3f59aaca611_0_39', text: 'Reactivate in August with a £300–500 test budget. Confirm the custom audience is still current. If CTR holds above 3%, scale aggressively. This is the Q2 reactivation priority — it should not wait another month.\n\n' },
  ],
  brand: [
    // Slide 4 — OKR chart label; June's combined LI+Meta total blanked
    // pending Meta (July LinkedIn-only figure isn't a like-for-like swap
    // for a tile explicitly captioned "Linkedin + Meta").
    { objectId: 'g3f393645d4a_0_71', text: 'July Total' },
    { objectId: 'g3f393645d4a_0_72', text: '—' },
  ],
};

// ---- Native cover rebuild (b2b / brand only) -------------------------------
// The b2b and brand masters' Slide 1 is still a single flattened screenshot
// image (confirmed via a raw Slides API element dump — one image element,
// no text). That means the month is baked into pixels and can never be
// updated without re-exporting a new screenshot from whatever design tool
// produced it. This replaces it with real Slides shapes/text so the month
// updates like any other field. Visual language (deep red field, darker
// diagonal wedge, serif headline, small-caps kicker, translucent stat card)
// is matched from the rendered master, not pixel-for-pixel — colours were
// sampled directly off the master's exported PDF.
//
// The product master no longer needs this: its cover was rebuilt as native
// cov_* shapes during the earlier "no screenshot cover" pass, those shapes
// survive drive.files.copy, and its cover text now lives as plain entries
// in EDITS.product above (see generate() below — buildCoverRequests only
// runs when COVER[deckKey] exists).

const PAGE = { w: 9144000, h: 5143500 }; // EMU, confirmed via presentations.get
const PX = 6350; // EMU per px at this page's native render scale

const COLOR = {
  bg: { red: 0.8118, green: 0.0784, blue: 0.1725 },     // #CF142C
  wedge: { red: 0.6392, green: 0.0549, blue: 0.1333 },   // #A30E22
  card: { red: 0.8314, green: 0.1765, blue: 0.2549 },    // #D42D41
  white: { red: 1, green: 1, blue: 1 },
  muted: { red: 0.9490, green: 0.6392, blue: 0.7098 },   // #F2A3B5
};

const COVER = {
  b2b: {
    slideId: 'g3f393228d3a_0_0',
    imageId: 'g3f393228d3a_0_4',
    deptLabel: 'B2B LEAD GENERATION  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF',
    subtitle: 'B2B Lead Generation · LinkedIn Account 509501623 · July 2026',
    cardLabel: 'JULY SNAPSHOT',
    cardBody: '£1,295 spend delivered 10 B2B leads at £130 CPL in July — a 69% improvement on June’s £415 CPL. APAC drove 8 of the 10 leads on £659 spend; EMEA added 2 on £636; UK stayed paused.',
    miniStats: null,
  },
  // product intentionally absent: its cover is native cov_* shapes already,
  // populated via plain text edits in EDITS.product instead.
  brand: {
    slideId: 'g3f393645d4a_0_0',
    imageId: 'g3f393645d4a_0_4',
    deptLabel: 'BRAND & EVENTS DIVISION  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF',
    subtitle: 'Brand & Events · LinkedIn Account 509501623 · July 2026',
    cardLabel: 'JULY SNAPSHOT',
    cardBody: 'LinkedIn paid reach hit 309,292 impressions in July at £12.4 CPM. VCF delivered 34 employer leads — Africa 24 at £44 CPL, APAC 10 at £177 — plus a 135,630-impression awareness boost.',
    miniStats: null,
  },
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
    {
      updateShapeProperties: {
        objectId,
        shapeProperties: { shapeBackgroundFill: { solidFill: { color: { rgbColor: color } } }, outline: { propertyState: 'NOT_RENDERED' } },
        fields: 'shapeBackgroundFill.solidFill.color,outline.propertyState',
      },
    },
  ];
}

function textRequests(objectId, slideId, box, text, { fontFamily, fontSize, bold = false, color, align = 'START', lineSpacing }) {
  const requests = [
    { createShape: { objectId, shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: slideId, ...emuBox(...box) } } },
    { insertText: { objectId, insertionIndex: 0, text } },
    {
      updateTextStyle: {
        objectId,
        style: { fontFamily, fontSize: { magnitude: fontSize, unit: 'PT' }, bold, foregroundColor: { opaqueColor: { rgbColor: color } } },
        textRange: { type: 'ALL' },
        fields: 'fontFamily,fontSize,bold,foregroundColor',
      },
    },
    {
      updateParagraphStyle: {
        objectId,
        style: { alignment: align, ...(lineSpacing ? { lineSpacing } : {}) },
        textRange: { type: 'ALL' },
        fields: lineSpacing ? 'alignment,lineSpacing' : 'alignment',
      },
    },
    { updateShapeProperties: { objectId, shapeProperties: { outline: { propertyState: 'NOT_RENDERED' } }, fields: 'outline.propertyState' } },
  ];
  return requests;
}

function buildCoverRequests(spec) {
  const { slideId, imageId } = spec;
  const requests = [];

  requests.push({ deleteObject: { objectId: imageId } });

  requests.push(...rectRequests('cov_bg', slideId, [0, 0, 1440, 810], 'RECTANGLE', COLOR.bg));
  requests.push(...rectRequests('cov_wedge', slideId, [760, 0, 680, 810], 'PARALLELOGRAM', COLOR.wedge));

  requests.push(...textRequests('cov_acca', slideId, [95, 28, 300, 48], 'ACCA', {
    fontFamily: 'Lora', fontSize: 24, bold: true, color: COLOR.white,
  }));
  requests.push(...textRequests('cov_dept', slideId, [95, 84, 900, 26], spec.deptLabel, {
    fontFamily: 'DM Sans', fontSize: 10, bold: true, color: COLOR.muted,
  }));
  requests.push(...rectRequests('cov_rule', slideId, [95, 162, 55, 3], 'RECTANGLE', COLOR.white));
  requests.push(...textRequests('cov_headline', slideId, [90, 192, 760, 150], 'July 2026\nPerformance Brief', {
    fontFamily: 'Lora', fontSize: 34, bold: true, color: COLOR.white, lineSpacing: 108,
  }));
  requests.push(...textRequests('cov_subtitle', slideId, [95, 352, 760, 60], spec.subtitle, {
    fontFamily: 'DM Sans', fontSize: 13, color: COLOR.muted,
  }));

  requests.push(...rectRequests('cov_card', slideId, [95, 458, 730, spec.miniStats ? 300 : 262], 'ROUND_RECTANGLE', COLOR.card));
  requests.push(...textRequests('cov_cardlabel', slideId, [130, 490, 600, 22], spec.cardLabel, {
    fontFamily: 'DM Sans', fontSize: 9, bold: true, color: COLOR.muted,
  }));
  requests.push(...textRequests('cov_cardbody', slideId, [130, 520, 655, 190], spec.cardBody, {
    fontFamily: 'Lora', fontSize: 14, bold: true, color: COLOR.white, lineSpacing: 125,
  }));

  if (spec.miniStats) {
    spec.miniStats.forEach(([label, value], i) => {
      const x = 130 + i * 210;
      requests.push(...textRequests(`cov_stat${i}_label`, slideId, [x, 695, 195, 18], label, {
        fontFamily: 'DM Sans', fontSize: 8, bold: true, color: COLOR.muted,
      }));
      requests.push(...textRequests(`cov_stat${i}_value`, slideId, [x, 714, 195, 26], value, {
        fontFamily: 'Lora', fontSize: 15, bold: true, color: COLOR.white,
      }));
    });
  }

  requests.push(...textRequests('cov_footer', slideId, [95, 766, 600, 26], 'PREPARED BY FILL MY FUNNEL  ·  JULY 2026', {
    fontFamily: 'DM Sans', fontSize: 9, bold: true, color: COLOR.muted,
  }));

  return requests;
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
    requests.push({ insertText: { objectId, cellLocation, insertionIndex: 0, text } });
  }
  return requests;
}

// Product's appendix table (slide 9, g3f59aaca611_0_73) replaced the old
// per-status table entirely — it's now a single Layer/Ad Set/Spend/Impr/
// Clicks/CTR/CPM/CPC breakdown with no Status column, so the old "Reactivated
// Jul"/"Reactivate Aug" cell edits no longer have anywhere to live. This is
// an identity fill of the CURRENT July numbers (grounded in the inspect5.mjs
// dump) — the cells a future month's editor swaps for new ad-set performance.
// NOTE: row 3 col 3 ("44.692") is dumped verbatim from the master and looks
// like a data-entry typo (every other Impr cell uses a comma, e.g. "46,769")
// — flagged, not silently corrected, since it's the user's own edit.
const TABLE_EDITS = {
  product: [
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 2, text: '£532' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 3, text: '44.692' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 4, text: '355' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 5, text: '0.81%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 6, text: '£12.77' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 3, columnIndex: 7, text: '£1.55' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 2, text: '£483' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 3, text: '46,769' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 4, text: '241' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 5, text: '0.52%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 6, text: '£10.48' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 4, columnIndex: 7, text: '£2.00' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 2, text: '£489' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 3, text: '33,705' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 4, text: '268' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 5, text: '0.82%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 6, text: '£15.32' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 5, columnIndex: 7, text: '£1.86' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 2, text: '£525' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 3, text: '66,840' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 4, text: '255' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 5, text: '0.38%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 6, text: '£7.92' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 6, columnIndex: 7, text: '£2.08' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 2, text: '£1,078' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 3, text: '57,528' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 4, text: '346' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 5, text: '0.60%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 6, text: '£20.51' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 8, columnIndex: 7, text: '£3.42' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 2, text: '£754' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 3, text: '117,179' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 4, text: '792' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 5, text: '0.69%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 6, text: '£6.43' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 9, columnIndex: 7, text: '£0.95' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 2, text: '£348' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 3, text: '61,477' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 4, text: '411' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 5, text: '0.77%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 6, text: '£3.48' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 10, columnIndex: 7, text: '£0.50' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 2, text: '£18' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 3, text: '155' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 4, text: '5' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 5, text: '3.23%' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 6, text: '£116' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 12, columnIndex: 7, text: '£3.60' },

    { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 2, text: '£4,229' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 3, text: '428,345' },
    { objectId: 'g3f59aaca611_0_73', rowIndex: 13, columnIndex: 4, text: '2,673' },
  ],
};

async function generate(deckKey, masterId, title, edits = []) {
  const { data } = await drive.files.copy({
    fileId: masterId,
    supportsAllDrives: true,
    requestBody: { name: title, parents: [SHARED_DRIVE] },
  });
  const requests = [
    ...editRequests(edits),
    ...tableEditRequests(TABLE_EDITS[deckKey] || []),
    ...(COVER[deckKey] ? buildCoverRequests(COVER[deckKey]) : []),
  ];
  await slides.presentations.batchUpdate({ presentationId: data.id, requestBody: { requests } });
  console.log(title, '→', `https://docs.google.com/presentation/d/${data.id}/edit`);
  return data.id;
}

// TEMP: verifying the new Product master round-trips correctly before
// re-enabling b2b/brand (unchanged, already delivered this cycle — no need
// to regenerate duplicate copies just to test Product's new template).
await generate('product', MASTERS.product, 'ACCA Product Marketing — July 2026 (template check)', EDITS.product);
// await generate('b2b', MASTERS.b2b, 'ACCA B2B — July 2026', EDITS.b2b);
// await generate('brand', MASTERS.brand, 'ACCA Employer Brand — July 2026', EDITS.brand);
