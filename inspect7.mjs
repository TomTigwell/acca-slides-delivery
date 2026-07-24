import { google } from 'googleapis';

// Read-only: find sheetsChart page elements across all three masters.
// inspect.mjs/inspect2.mjs (the original b2b/brand inspection scripts)
// never checked for sheetsChart elements, so any linked charts on those
// masters' North Star slides were missed the same way the Product one was.

const MASTERS = {
  product: '1H6pbNQEjt7gBWzpUs3Ns06Y11mSSUsXXRpLA0QkgZks',
  b2b:     '1dZJV-dvJaRVUxTCYOoR2QQcyRa2N6LAIV780U8ezqHg',
  brand:   '12WS8qbsgSW8KrOLA9h2A9mjVLnf4aBAQuH4MY1s95P8',
};

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

function findCharts(el, out) {
  if (el.sheetsChart) out.push(el);
  if (el.elementGroup?.children) el.elementGroup.children.forEach((c) => findCharts(c, out));
}

for (const [key, id] of Object.entries(MASTERS)) {
  const { data } = await slides.presentations.get({ presentationId: id });
  console.log(`=== ${key} (${id}) — ${data.slides.length} slides ===`);
  data.slides.forEach((slide, i) => {
    const out = [];
    (slide.pageElements || []).forEach((el) => findCharts(el, out));
    if (out.length) {
      console.log(`--- Slide ${i + 1} (${slide.objectId}) ---`);
      out.forEach((el) => console.log(`  [SHEETSCHART ${el.objectId}] spreadsheetId=${el.sheetsChart.spreadsheetId} chartId=${el.sheetsChart.chartId}`));
    }
  });
}
