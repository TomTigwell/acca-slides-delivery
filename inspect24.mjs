import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const PX = 6350;

const { data } = await slides.presentations.get({ presentationId: BRAND_DECK });
const slide = data.slides.find(s => s.objectId === 'p3');

const targets = ['p3_i18', 'p3_i19', 'p3_i20', 'p3_i21', 'g3f59aaca611_0_6', 'g3f59aaca611_0_7', 'g3f59aaca611_0_8', 'g3f59aaca611_0_9'];
for (const el of slide.pageElements || []) {
  if (!targets.includes(el.objectId)) continue;
  const sz = el.size, tr = el.transform;
  const wPx = sz ? (sz.width.magnitude * (tr.scaleX ?? 1)) / PX : null;
  const hPx = sz ? (sz.height.magnitude * (tr.scaleY ?? 1)) / PX : null;
  const xPx = tr ? tr.translateX / PX : null;
  const yPx = tr ? tr.translateY / PX : null;
  const shape = el.shape;
  const autofit = shape?.text ? JSON.stringify(shape.autofit || {}) : '';
  const fontSize = shape?.text?.textElements?.find(te => te.textRun)?.textRun?.style?.fontSize;
  console.log(el.objectId, `pos=(${xPx?.toFixed(1)},${yPx?.toFixed(1)}) size=(${wPx?.toFixed(1)}x${hPx?.toFixed(1)})`, 'fontSize=', JSON.stringify(fontSize), 'autofit=', autofit);
}
