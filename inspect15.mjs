import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const TARGET_SLIDES = ['p6', 'g3f393228d3a_0_154', 'p4'];

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}

function dumpCellText(cell) {
  if (!cell?.text?.textElements) return '';
  return cell.text.textElements.map(te => te.textRun?.content || '').join('');
}

const { data } = await slides.presentations.get({ presentationId: B2B_DECK });

for (const slideId of TARGET_SLIDES) {
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
      const text = dumpText(el.shape).slice(0, 80).replace(/\n/g, ' ');
      console.log(`  ${el.objectId} | SHAPE:${el.shape.shapeType} | "${text}"`);
    } else if (el.image) {
      console.log(`  ${el.objectId} | IMAGE`);
    }
  }
}
