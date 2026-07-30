import { google } from 'googleapis';

// Read-only: get exact positions of the "What We Recommend Next" text
// objects (slide 5) on the Product template master, to determine which
// ones actually fall inside the left "WHAT WE'RE DOING" box vs the right
// "WHAT WE NEED FROM YOU" box. Object 140's content has been rendering
// inside the LEFT box on both the Brand and B2C decks despite being
// written as a "need" item — confirming this empirically instead of
// guessing again.

const TEMPLATE_MASTER = '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks';
const PX = 6350;
const TARGETS = new Set(['g3f483bfcec4_0_120', 'g3f483bfcec4_0_123', 'g3f483bfcec4_0_126', 'g3f483bfcec4_0_129', 'g3f483bfcec4_0_134', 'g3f483bfcec4_0_137', 'g3f483bfcec4_0_140', 'g3f483bfcec4_0_143', 'g3f483bfcec4_0_146', 'g3f59aaca611_0_5']);

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

const { data } = await slides.presentations.get({ presentationId: TEMPLATE_MASTER });
const slide5 = data.slides[4];
console.log(`=== Slide 5 (${slide5.objectId}) ===`);
for (const el of slide5.pageElements || []) {
  if (!TARGETS.has(el.objectId)) continue;
  const x = Math.round(el.transform.translateX / PX);
  const y = Math.round(el.transform.translateY / PX);
  console.log(`  ${el.objectId}: x=${x} y=${y}`);
}
