import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const TARGETS = [
  { name: 'B2B', id: '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA', slideId: 'p4' },
  { name: 'B2C', id: '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU', slideId: 'g3f483bfcec4_0_95' },
];

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}

for (const t of TARGETS) {
  console.log('\n=====', t.name, t.id, t.slideId, '=====');
  const { data } = await slides.presentations.get({ presentationId: t.id });
  const slide = data.slides.find(s => s.objectId === t.slideId);
  if (!slide) { console.log('NOT FOUND'); continue; }
  for (const el of slide.pageElements || []) {
    const kind = el.shape ? 'SHAPE:' + el.shape.shapeType : el.table ? 'TABLE' : el.sheetsChart ? 'SHEETSCHART' : el.image ? 'IMAGE' : el.line ? 'LINE:' + el.line.lineType : Object.keys(el).find(k => !['objectId','size','transform'].includes(k));
    const w = el.size ? Math.round((el.size.width.magnitude * (el.transform?.scaleX ?? 1)) / 6350) : '?';
    const h = el.size ? Math.round((el.size.height.magnitude * (el.transform?.scaleY ?? 1)) / 6350) : '?';
    const x = el.transform ? Math.round(el.transform.translateX / 6350) : '?';
    const y = el.transform ? Math.round(el.transform.translateY / 6350) : '?';
    const text = el.shape ? dumpText(el.shape).slice(0, 100) : '';
    const fill = el.shape?.shapeProperties?.shapeBackgroundFill?.solidFill?.color?.rgbColor;
    console.log(`  ${el.objectId} | ${kind} | box(${x},${y},${w},${h}) | fill=${fill?JSON.stringify(fill):''} | "${text}"`);
  }
}
