import { google } from 'googleapis';

const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const MASTERS = {
  product: '13RsJLhqAsTgowBdVnWqAtG460MSz49Ph7vcx3ecvnqo',
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
  product: [
    // ---- Slide 2 — Exec Summary intro (LinkedIn / Meta paragraphs) ----
    // Sourced from ACCA_Product_Monthly_Brief_July_2026.html (the client-
    // facing July EOM brief) — real July 06 call content, not invented.
    { objectId: 'g3f483bfcec4_0_27', text: 'Linkedin → In July we launched CertOT on LinkedIn, reactivated the paused ProDipSust Sustainability Strategy ad set — the most efficient in the portfolio at £0.50 CPC — and maintained Data Analytics and Cybersecurity on Apple-UI creative across APAC, UK, and EMEA. CERT PFM groundwork began, confirmed as a B2B-weighted proposition launching later this year.\n\nMeta → B2C Certificates continue on Data Analytics and Cybersecurity Apple-UI/video creative. June’s campaign spent £3,382 for 22.5M impressions and 7.84M unique users at £0.15 CPM. The Audience Network validation flag (76% of spend) stays open before scaling further into July.\n' },

    // ---- Slide 3 — What We Did (4) / What We Learned (2) ----
    { objectId: 'g3f483bfcec4_0_34', text: 'CertOT launched on LinkedIn.' },
    { objectId: 'g3f483bfcec4_0_35', text: 'Certificate in Organisational Transformation went live in July, targeting L&D and HR decision-makers as a B2B-weighted proposition. First CTR and impression data due at August EOM.\n' },
    { objectId: 'g3f483bfcec4_0_38', text: 'ProDipSust Sustainability Strategy ad set reactivated.' },
    { objectId: 'g3f483bfcec4_0_39', text: 'Paused in Q1 despite a 0.77% CTR and £0.50 CPC — the cheapest clicks in the portfolio. Reactivated in July, running alongside a new testimonial video test.\n' },
    { objectId: 'g3f483bfcec4_0_42', text: 'Apple-UI format holding across Data Analytics and Cybersecurity.' },
    { objectId: 'g3f483bfcec4_0_43', text: 'Both certs continue on Apple-UI creative across LinkedIn and Meta B2C, sustaining Q1’s CTR performance into July.\n' },
    { objectId: 'g3f483bfcec4_0_46', text: 'CERT PFM groundwork underway.' },
    { objectId: 'g3f483bfcec4_0_47', text: 'Rashid confirmed CERT PFM is B2B-weighted, launching later this year — budget split favors LinkedIn over Meta B2C.\n' },
    { objectId: 'g3f483bfcec4_0_57', text: 'Full-funnel spend timing is now understood at senior level.' },
    { objectId: 'g3f483bfcec4_0_58', text: 'Trevor and Peter confirmed on the July 06 call that awareness spend doesn’t convert in the same month and accepted the lagged attribution model — clearing the way to report spend, impressions, and checkouts without re-justifying the funnel shape each month.\n' },
    { objectId: 'g3f483bfcec4_0_61', text: 'Checkout tracking remains the highest-value open action.' },
    { objectId: 'g3f483bfcec4_0_62', text: 'Trevor committed to personally push Alan on the checkout-tracking implementation. Once live, we can report actual sales attributed to paid ads and unlock Google Search as a channel.\n' },

    // ---- Slide 4 — North Star narrative ----
    // HTML's own chart shows July as a dashed "TBC" placeholder and says
    // "July actuals pending" — matched here rather than asserting a number.
    { objectId: 'g3f483bfcec4_0_106', text: 'Q1 closed at 127 checkouts in June — 59% above the 80/month target, averaging 83/month across the quarter. July actuals are pending; CertOT’s launch and the reactivated ProDipSust Sustainability Strategy ad set are the two moves that should maintain or extend the trajectory.\n\nProduct Checkout Initiations\n' },

    // ---- Slide 5 — What We Recommend Next (LinkedIn) ----
    // This slide's content was a known copy-paste bug (B2B lead-form text
    // pasted into the Product master — flagged in fmf-skills' token-map.md).
    // Replaced with real Product content from the July brief.
    { objectId: 'g3f483bfcec4_0_120', text: 'Running CertOT on LinkedIn with Apple-UI product creative — building impression base and first CTR read for August EOM.\n' },
    { objectId: 'g3f483bfcec4_0_123', text: 'Running the ProDipSust testimonial video test alongside the reactivated Sustainability Strategy static ad set — comparing checkout rates at next EOM.\n' },
    { objectId: 'g3f483bfcec4_0_126', text: 'Maintaining Data Analytics and Cybersecurity at current budget — Apple-UI on LinkedIn, video scaling on Meta B2C.\n' },
    { objectId: 'g3f483bfcec4_0_129', text: 'Preparing a full creative audit for Q2 direction, including a GIF-format test plan for August.\n' },
    { objectId: 'g3f483bfcec4_0_134', text: 'Trevor: nudge Alan on checkout-tracking implementation — once live, we can report sales attributed to paid ads and unlock Google Search as a channel.\n' },
    { objectId: 'g3f483bfcec4_0_137', text: 'Rashid: confirm the CERT PFM timeline for paid social — we need 2–3 weeks to brief Hassan and build creative once confirmed.\n' },
    { objectId: 'g3f483bfcec4_0_139', text: 'DECISION\n' },
    { objectId: 'g3f483bfcec4_0_140', text: 'Rashid: share the refreshed ProDipSust assets before the new exam window opens — we need to be live before it does.\n' },
    { objectId: 'g3f483bfcec4_0_143', text: 'Rashid: share the new B2B case study — we want to create snackable LinkedIn assets (document ad or carousel) from it.\n' },
    { objectId: 'g3f483bfcec4_0_146', text: 'Confirm ACCA Global Meta pixel coverage on cert product pages — this unlocks a warm retargeting pool beyond the LMS audience.\n' },

    // ---- Slide 6 — What We Recommend Next (Meta) ----
    { objectId: 'g3f483bfcec4_0_167', text: 'Can we access GA4 checkout and session data to cross-reference against Meta’s Audience Network share (76% of spend) before scaling further into July — and confirm the 25–34 segment is converting at a comparable rate?\n' },

    // ---- Slide 7 — Reach — Employer & Buyer Audience ----
    // The July brief itself keeps this section framed as Q1 historical
    // (pending a fresh pull) rather than asserting new July reach figures —
    // matched here instead of the earlier (now superseded) 364K estimate.
    { objectId: 'p3_i16', text: 'EMPLOYER & BUYER AUDIENCE REACH  ·  JULY 2026  ·  PRODUCT PORTFOLIO' },
    { objectId: 'p3_i17', text: 'LinkedIn continues to deliver qualified reach into Big Four firms — Deloitte and EY lead clicks in Q1. CertOT extends reach into L&D and HR. Confirming ACCA Global Meta pixel remains the highest-leverage open action.\n' },
    { objectId: 'p3_i19', text: 'LI IMPRESSIONS (Q1)' },
    { objectId: 'p3_i20', text: '428K' },
    { objectId: 'p3_i21', text: 'All cert ad sets · Apr–Jun' },

    // ---- Slide 9 — Experiments intro ----
    { objectId: 'g3f483bfcec4_0_245', text: 'Five experiments this month: two Q1 findings now confirmed (Apple-UI format, Meta video), and three live in July — ProDipSust testimonial video vs static, CertOT’s first Apple-UI test in a B2B context, and the Hot Layer reactivation call for August. The reactivated ProDipSust Sustainability Strategy ad set is the efficiency benchmark for the quarter — watch whether it holds £0.50 CPC at scale.\n' },

    // ---- Slide 11 — Experiment cards ----
    { objectId: 'g3f483bfcec4_0_269', text: 'Confirmed. Data Analytics Single Image: 0.81% CTR. Cybersecurity Single Image: 0.82% CTR. Both 80%+ above the 0.45% benchmark. Checkout activity confirmed from Cybersecurity.\n' },
    { objectId: 'g3f483bfcec4_0_272', text: 'Apple-UI is now the default format for all active certs. CertOT launches with Apple-UI creative in July — its first run in a B2B cert context. A GIF-format test against Apple-UI static is planned for Q2.\n' },
    { objectId: 'g3f483bfcec4_0_280', text: 'Both ad formats are now live — testimonial video and Apple-UI static running alongside the reactivated Sustainability Strategy ad set, three ProDipSust variants running simultaneously.\n' },
    { objectId: 'g3f483bfcec4_0_283', text: 'Read date moved to August EOM — need a minimum 4 weeks of data. Compare video view rate, CTR, and checkout initiation rate; winner takes the majority of ProDipSust budget in Q2.\n' },

    // ---- Slide 15 — Appendix footer ----
    { objectId: 'p6_i56', text: 'ACCA · Product Marketing · Monthly Performance Brief · Template v2.0 · July 2026\n' },
  ],
  brand: [
    // Slide 4 — OKR chart label; June's combined LI+Meta total blanked
    // pending Meta (July LinkedIn-only figure isn't a like-for-like swap
    // for a tile explicitly captioned "Linkedin + Meta").
    { objectId: 'g3f393645d4a_0_71', text: 'July Total' },
    { objectId: 'g3f393645d4a_0_72', text: '—' },
  ],
};

// ---- Native cover rebuild -------------------------------------------------
// The master's Slide 1 is a single flattened screenshot image (confirmed via
// a raw Slides API element dump — one image element, no text). That means
// the month is baked into pixels and can never be updated without
// re-exporting a new screenshot from whatever design tool produced it. This
// replaces it with real Slides shapes/text so the month updates like any
// other field. Visual language (deep red field, darker diagonal wedge,
// serif headline, small-caps kicker, translucent stat card) is matched from
// the rendered master, not pixel-for-pixel — colours were sampled directly
// off the master's exported PDF.

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
  product: {
    slideId: 'g3f483bfcec4_0_0',
    imageId: 'g3f483bfcec4_0_4',
    deptLabel: 'PRODUCT MARKETING  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF',
    subtitle: 'Product Marketing · LinkedIn Account 509501623 · July 2026',
    cardLabel: 'JULY SNAPSHOT',
    cardBody: 'Q2 opens with CertOT live and the persona-creative build underway. July will show whether Apple-UI and the reactivated ProDipSust Sustainability Strategy ad set can hold June’s 127-checkout momentum.',
    miniStats: null,
  },
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

// Ad-set status cells the July brief explicitly updates (both previously
// "Paused") — grounded in ACCA_Product_Monthly_Brief_July_2026.html.
const TABLE_EDITS = {
  product: [
    { objectId: 'p6_g12', rowIndex: 7, columnIndex: 4, text: 'Reactivated Jul' },
    { objectId: 'p6_g12', rowIndex: 8, columnIndex: 4, text: 'Reactivate Aug' },
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
    ...buildCoverRequests(COVER[deckKey]),
  ];
  await slides.presentations.batchUpdate({ presentationId: data.id, requestBody: { requests } });
  console.log(title, '→', `https://docs.google.com/presentation/d/${data.id}/edit`);
  return data.id;
}

await generate('b2b', MASTERS.b2b, 'ACCA B2B — July 2026', EDITS.b2b);
await generate('product', MASTERS.product, 'ACCA Product Marketing — July 2026', EDITS.product);
await generate('brand', MASTERS.brand, 'ACCA Employer Brand — July 2026', EDITS.brand);
