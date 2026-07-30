import { google } from 'googleapis';

// One-off patch: the May-Jul chart images inserted by apply-northstar-charts.mjs
// were placed at y=395, but the card's JULY LEADS / BLENDED CPL value text
// (brand_ns_value1/2, b2c_ns_value1/2) overflows its nominal 60px-tall box —
// confirmed via PDF text-bbox inspection: the "35"/"£84.54" glyphs render down
// to y~427 (40pt font in a 60px box, no autofit). Moving the chart down to
// y=460 and shrinking it to keep the same aspect ratio clears the overlap
// while staying inside the card (bottom edge at y=754).

const PX = 6350;
const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function imageBox(xPx, yPx, wPx, hPx) {
  return {
    size: { width: { magnitude: wPx * PX, unit: 'EMU' }, height: { magnitude: hPx * PX, unit: 'EMU' } },
    transform: { scaleX: 1, scaleY: 1, translateX: xPx * PX, translateY: yPx * PX, unit: 'EMU' },
  };
}

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';

const NEW_BOX = [140, 460, 634, 274];

async function reposition(deckId, objectId, url, label) {
  const requests = [
    { deleteObject: { objectId } },
    { createImage: { objectId, url, elementProperties: { pageObjectId: 'g3f483bfcec4_0_95', ...imageBox(...NEW_BOX) } } },
  ];
  await slides.presentations.batchUpdate({ presentationId: deckId, requestBody: { requests } });
  console.log(label, 'repositioned');
}

await reposition(BRAND_DECK, 'brand_ns_chart_img', 'https://drive.google.com/uc?export=view&id=19G7p-Eymo5dhYsDBQCGBRewh5bv96XQ2', 'Brand');
await reposition(B2C_DECK, 'b2c_ns_chart_img', 'https://drive.google.com/uc?export=view&id=1HX8frBR9QJCzm_OHL0O8vKV1iJgVdq63', 'B2C');
console.log('done');
