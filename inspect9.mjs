import { google } from 'googleapis';

// Read-only: get the exact size/transform of the North Star sheetsChart
// element (and its neighbours) on the Product template master, so the
// Brand generator can size a replacement stat block correctly. The last
// generate-brand.mjs run failed with "createShape: width should be greater
// than zero" — meaning my assumed box math produced a non-positive width,
// so get the ground truth instead of guessing again.

const TEMPLATE_MASTER = '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks';
const PX = 6350;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

const { data } = await slides.presentations.get({ presentationId: TEMPLATE_MASTER });
const slide4 = data.slides[3]; // North Star is slide 4 (index 3)
console.log(`=== Slide 4 (${slide4.objectId}) page elements ===`);
for (const el of slide4.pageElements || []) {
  const w = el.size?.width?.magnitude;
  const h = el.size?.height?.magnitude;
  const x = el.transform?.translateX;
  const y = el.transform?.translateY;
  const wPx = w != null ? Math.round(w / PX) : null;
  const hPx = h != null ? Math.round(h / PX) : null;
  const xPx = x != null ? Math.round(x / PX) : null;
  const yPx = y != null ? Math.round(y / PX) : null;
  const kind = el.sheetsChart ? 'SHEETSCHART' : el.shape ? 'shape' : el.table ? 'table' : el.image ? 'image' : 'other';
  console.log(`  [${kind} ${el.objectId}] box=(${xPx},${yPx},${wPx},${hPx})px  size=${JSON.stringify(el.size)} transform=${JSON.stringify(el.transform)}`);
}
