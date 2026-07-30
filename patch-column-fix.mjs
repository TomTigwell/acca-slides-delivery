import { google } from 'googleapis';

// One-off patch: object g3f483bfcec4_0_140 on the "What We Recommend Next"
// slide is a genuine 5th LEFT-column ("WHAT WE'RE DOING") slot, not a
// right-column ("WHAT WE NEED FROM YOU") slot as assumed when building
// both the Brand and B2C decks — confirmed via inspect11.mjs (x=118,
// matching 120/123/126/129's x≈118-121; the five real right-column items
// are all at x=829). Both decks had a "need"-phrased ask sitting in the
// left box. Fix: give 140 a genuine "doing" statement, and fold the
// displaced ask into an existing right-column slot instead of losing it.

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function cell(objectId, text) {
  return [
    { deleteText: { objectId, textRange: { type: 'ALL' } } },
    { insertText: { objectId, insertionIndex: 0, text } },
  ];
}

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';

const brandRequests = [
  ...cell('g3f483bfcec4_0_140', 'Monitoring the Africa vs APAC CPL gap monthly to confirm whether it holds as spend scales.\n\n'),
  ...cell('g3f483bfcec4_0_143', 'Confirm whether the APAC/Africa CPL gap reflects audience saturation or a creative/targeting difference worth testing — share July’s VCF lead list so we can check lead quality against it.\n'),
];

const b2cRequests = [
  ...cell('g3f483bfcec4_0_140', 'Monitoring the Cert-OT Traffic vs Views CPA gap monthly to confirm whether it holds as spend scales.\n\n'),
  ...cell('g3f59aaca611_0_5', 'Confirm which of the two conflicting conversion totals (817 vs 1,141 — see appendix) is authoritative, and which funnel action (Register, Add To Cart, or Proceed Checkout) should be the primary KPI going forward.\n\n'),
];

await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests: brandRequests } });
console.log('patched Brand deck column fix');
await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests: b2cRequests } });
console.log('patched B2C deck column fix');
