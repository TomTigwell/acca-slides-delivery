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
    // Slide 7 — Reach — Employer & Buyer Audience: LinkedIn stat tiles
    { objectId: 'p3_i16', text: 'EMPLOYER & BUYER AUDIENCE REACH  ·  July 2026  ·  PRODUCT PORTFOLIO' },
    { objectId: 'p3_i19', text: 'LI IMPRESSIONS (July)' },
    { objectId: 'p3_i20', text: '364K' },
    { objectId: 'p3_i21', text: 'All cert ad sets · July' },
    // Meta-sourced tiles: blanked for the freelancer, not left showing June data
    { objectId: 'p3_i23', text: 'Meta Impr. (July)' },
    { objectId: 'p3_i24', text: '—' },
    { objectId: 'p3_i25', text: 'Pending Meta data' },
    { objectId: 'p3_i32', text: 'META B2C  ·  AUDIENCE INSIGHTS  ·  JULY 2026' },
    { objectId: 'p3_i38', text: '—' },
    { objectId: 'p3_i39', text: '—' },
    { objectId: 'p3_i43', text: '—' },
    { objectId: 'p3_i44', text: '—' },
    { objectId: 'p3_i48', text: '—' },
    { objectId: 'p3_i49', text: '—' },
    { objectId: 'p3_i53', text: '—' },
    { objectId: 'p3_i54', text: '—' },
    { objectId: 'p3_i58', text: '—' },
    { objectId: 'p3_i59', text: '—' },
    { objectId: 'p3_i62', text: 'Female —' },
    { objectId: 'p3_i64', text: 'Male —' },
    { objectId: 'p3_i65', text: '⚠ Pending Meta data — freelancer to confirm Audience Network share' },
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
    cardBody: 'Checkout initiations reached 86 in July so far — 8% above the 80/month target — at £25 cost per checkout, down from June’s 127 as the funnel narrows: 1,458 registrations, 112 add-to-carts, 86 checkouts.',
    miniStats: [
      ['JULY CHECKOUTS', '86 vs 80 target'],
      ['COST / CHECKOUT', '£25'],
      ['LINKEDIN REACH', '364,097 impr'],
    ],
  },
  brand: {
    slideId: 'g3f393645d4a_0_0',
    imageId: 'g3f393645d4a_0_4',
    deptLabel: 'BRAND & EVENTS DIVISION  ·  JULY 2026  ·  MONTHLY PERFORMANCE BRIEF',
    subtitle: 'Brand & Events · LinkedIn Account 509501623 · July 2026',
    cardLabel: 'JULY SNAPSHOT',
    cardBody: 'LinkedIn paid reach hit 309,292 impressions in July at £12.4 CPM. VCF delivered 34 employer leads — Africa 24 at £44 CPL, APAC 10 at £177 — plus a 135,630-impression APAC awareness boost. Meta figures pending from the freelancer.',
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

  requests.push(...rectRequests('cov_card', slideId, [95, 458, 730, spec.miniStats ? 300 : 250], 'ROUND_RECTANGLE', COLOR.card));
  requests.push(...textRequests('cov_cardlabel', slideId, [130, 492, 600, 22], spec.cardLabel, {
    fontFamily: 'DM Sans', fontSize: 9, bold: true, color: COLOR.muted,
  }));
  requests.push(...textRequests('cov_cardbody', slideId, [130, 522, 655, 170], spec.cardBody, {
    fontFamily: 'Lora', fontSize: 14.5, bold: true, color: COLOR.white, lineSpacing: 130,
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

async function generate(deckKey, masterId, title, edits = []) {
  const { data } = await drive.files.copy({
    fileId: masterId,
    supportsAllDrives: true,
    requestBody: { name: title, parents: [SHARED_DRIVE] },
  });
  const requests = [...editRequests(edits), ...buildCoverRequests(COVER[deckKey])];
  await slides.presentations.batchUpdate({ presentationId: data.id, requestBody: { requests } });
  console.log(title, '→', `https://docs.google.com/presentation/d/${data.id}/edit`);
  return data.id;
}

await generate('b2b', MASTERS.b2b, 'ACCA B2B — July 2026', EDITS.b2b);
await generate('product', MASTERS.product, 'ACCA Product Marketing — July 2026', EDITS.product);
await generate('brand', MASTERS.brand, 'ACCA Employer Brand — July 2026', EDITS.brand);
