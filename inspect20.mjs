import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
const PX = 6350;

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}
function dumpCellText(cell) {
  if (!cell?.text?.textElements) return '';
  return cell.text.textElements.map(te => te.textRun?.content || '').join('');
}

const { data } = await slides.presentations.get({ presentationId: B2C_DECK });

console.log('total slides:', data.slides.length);
data.slides.forEach((s, i) => {
  const texts = (s.pageElements || []).map(el => el.shape ? dumpText(el.shape) : '').filter(Boolean);
  console.log(i, s.objectId, JSON.stringify((texts[0] || '').slice(0, 50).replace(/\n/g,' ')));
});

for (const slideId of ['g3f483bfcec4_0_0', 'g3f483bfcec4_0_95', 'g3f59aaca611_0_64']) {
  const slide = data.slides.find(s => s.objectId === slideId);
  console.log('\n===== slide', slideId, '=====');
  if (!slide) { console.log('NOT FOUND'); continue; }
  for (const el of slide.pageElements || []) {
    if (el.table) {
      console.log(`  TABLE ${el.objectId} rows=${el.table.rows} cols=${el.table.columns}`);
      el.table.tableRows.forEach((row, ri) => {
        const cells = row.tableCells.map(c => dumpCellText(c).replace(/\n/g, '\\n')).join(' | ');
        console.log(`    [${ri}] ${cells}`);
      });
    } else if (el.shape) {
      const text = dumpText(el.shape).slice(0, 120).replace(/\n/g, ' ');
      console.log(`  ${el.objectId} | "${text}"`);
    } else if (el.image) {
      console.log(`  ${el.objectId} | IMAGE`);
    }
  }
}
