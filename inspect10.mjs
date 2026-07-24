import { google } from 'googleapis';

// Read-only: the generated Brand deck's Reach slide (6) shows header/label
// text overlapping its neighbour in a PDF export, but the equivalent
// Product slide (same master, unedited numbers) renders cleanly. Get the
// real box geometry + text style for the affected objects on the LIVE
// Brand deck (not the master) to determine whether this is a real layout
// break or a stale-export artifact.

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const TARGETS = new Set(['p3_i16', 'p3_i17', 'p3_i19', 'p3_i20', 'p3_i21', 'g3f59aaca611_0_7', 'g3f59aaca611_0_8', 'g3f59aaca611_0_9']);
const PX = 6350;

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

const { data } = await slides.presentations.get({ presentationId: BRAND_DECK });
for (const slide of data.slides) {
  for (const el of slide.pageElements || []) {
    if (!TARGETS.has(el.objectId)) continue;
    const w = el.size?.width?.magnitude;
    const h = el.size?.height?.magnitude;
    const sx = el.transform?.scaleX ?? 1;
    const sy = el.transform?.scaleY ?? 1;
    const x = el.transform?.translateX;
    const y = el.transform?.translateY;
    console.log(`--- ${el.objectId} ---`);
    console.log(`  box px: x=${Math.round(x / PX)} y=${Math.round(y / PX)} w=${Math.round((w * sx) / PX)} h=${Math.round((h * sy) / PX)}`);
    console.log(`  autofit: ${JSON.stringify(el.shape?.text?.autofit || el.shape?.shapeProperties?.autofit || 'none-reported')}`);
    const runs = el.shape?.text?.textElements || [];
    for (const te of runs) {
      if (te.textRun) {
        console.log(`  run: "${te.textRun.content}" style=${JSON.stringify(te.textRun.style)}`);
      }
    }
  }
}
