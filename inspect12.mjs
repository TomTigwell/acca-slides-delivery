import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const TARGETS = [
  { name: 'Brand', id: '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw', slideId: 'g3f483bfcec4_0_95' },
  { name: 'B2B', id: '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA', slideId: 'g3f393228d3a_0_85' },
  { name: 'B2C', id: '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU', slideId: 'g3f483bfcec4_0_0' },
];

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}

for (const t of TARGETS) {
  console.log('\n=====', t.name, t.id, '=====');
  const { data } = await slides.presentations.get({ presentationId: t.id });
  console.log('total slides:', data.slides.length);
  const idx = data.slides.findIndex(s => s.objectId === t.slideId);
  console.log('target slideId found at index:', idx);
  const slide = idx >= 0 ? data.slides[idx] : null;
  if (!slide) {
    console.log('NOT FOUND by id — dumping all slide IDs + first text for identification:');
    data.slides.forEach((s, i) => {
      const firstText = (s.pageElements || []).map(el => dumpText(el.shape)).filter(Boolean)[0] || '';
      console.log(i, s.objectId, JSON.stringify(firstText.slice(0, 60)));
    });
    continue;
  }
  console.log('--- elements on target slide ---');
  for (const el of slide.pageElements || []) {
    const kind = el.shape ? 'SHAPE:' + el.shape.shapeType : el.table ? 'TABLE' : el.sheetsChart ? 'SHEETSCHART' : el.image ? 'IMAGE' : el.line ? 'LINE' : Object.keys(el).find(k => !['objectId','size','transform'].includes(k));
    const w = el.size ? Math.round((el.size.width.magnitude * (el.transform?.scaleX ?? 1)) / 6350) : '?';
    const h = el.size ? Math.round((el.size.height.magnitude * (el.transform?.scaleY ?? 1)) / 6350) : '?';
    const x = el.transform ? Math.round(el.transform.translateX / 6350) : '?';
    const y = el.transform ? Math.round(el.transform.translateY / 6350) : '?';
    const text = el.shape ? dumpText(el.shape).slice(0, 100) : '';
    console.log(`  ${el.objectId} | ${kind} | box(${x},${y},${w},${h}) | "${text}"`);
  }
}
