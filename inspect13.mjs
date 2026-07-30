import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const TARGETS = [
  { name: 'Brand', id: '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw' },
  { name: 'B2B', id: '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA' },
  { name: 'B2C', id: '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU' },
];

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}

for (const t of TARGETS) {
  console.log('\n=====', t.name, t.id, '=====');
  const { data } = await slides.presentations.get({ presentationId: t.id });
  data.slides.forEach((s, i) => {
    const texts = (s.pageElements || []).map(el => dumpText(el.shape)).filter(Boolean);
    const heading = texts[0] || '(no text found)';
    const hasChart = (s.pageElements || []).some(el => el.sheetsChart);
    const hasLine = (s.pageElements || []).some(el => el.line);
    console.log(`  [${i}] ${s.objectId} | chart=${hasChart} line=${hasLine} | "${heading.slice(0, 70).replace(/\n/g, ' ')}"`);
  });
}
