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
// tiles (B2B/Brand) and the Product checkout-initiations narrative — these
// need a strategist's judgment call on how to reframe quarterly/H1 context
// into a mid-quarter monthly read, not a mechanical number swap. Company
// tables and per-region ad-set breakdowns are left as-is: no July-level
// granular breakdown was supplied for them. Charts are not touched (no
// chart-data API wired up here).
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

function editRequests(edits) {
  const requests = [];
  for (const { objectId, text } of edits) {
    requests.push({ deleteText: { objectId, textRange: { type: 'ALL' } } });
    requests.push({ insertText: { objectId, insertionIndex: 0, text } });
  }
  return requests;
}

async function generate(masterId, title, edits = []) {
  const { data } = await drive.files.copy({
    fileId: masterId,
    supportsAllDrives: true,
    requestBody: { name: title, parents: [SHARED_DRIVE] },
  });
  const requests = editRequests(edits);
  if (requests.length) {
    await slides.presentations.batchUpdate({ presentationId: data.id, requestBody: { requests } });
  }
  console.log(title, '→', `https://docs.google.com/presentation/d/${data.id}/edit`);
  return data.id;
}

await generate(MASTERS.b2b, 'ACCA B2B — July 2026', EDITS.b2b);
await generate(MASTERS.product, 'ACCA Product Marketing — July 2026', EDITS.product);
await generate(MASTERS.brand, 'ACCA Employer Brand — July 2026', EDITS.brand);
