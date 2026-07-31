import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';

function dumpText(shape) {
  if (!shape?.text?.textElements) return '';
  return shape.text.textElements.map(te => te.textRun?.content || '').join('');
}
function dumpCellText(cell) {
  if (!cell?.text?.textElements) return '';
  return cell.text.textElements.map(te => te.textRun?.content || '').join('');
}

const { data } = await slides.presentations.get({ presentationId: BRAND_DECK });

console.log('total slides:', data.slides.length);

// Full dump of every slide's elements, so we catch any renamed objects.
data.slides.forEach((slide, i) => {
  console.log(`\n===== [${i}] slide ${slide.objectId} =====`);
  for (const el of slide.pageElements || []) {
    if (el.table) {
      console.log(`  TABLE ${el.objectId} rows=${el.table.rows} cols=${el.table.columns}`);
      el.table.tableRows.forEach((row, ri) => {
        const cells = row.tableCells.map(c => dumpCellText(c).replace(/\n/g, '\\n')).join(' | ');
        console.log(`    [${ri}] ${cells}`);
      });
    } else if (el.shape) {
      const text = dumpText(el.shape).replace(/\n/g, '\\n');
      console.log(`  ${el.objectId} | "${text}"`);
    } else if (el.image) {
      console.log(`  ${el.objectId} | IMAGE`);
    } else {
      console.log(`  ${el.objectId} | (other: ${Object.keys(el).filter(k => k !== 'objectId')})`);
    }
  }
});

// Also specifically flag any text containing key search terms, regardless of slide/object.
console.log('\n===== SEARCH: JULY LEADS / stat tiles =====');
for (const slide of data.slides) {
  for (const el of slide.pageElements || []) {
    const text = el.shape ? dumpText(el.shape) : '';
    if (/july|leads|cpl/i.test(text) && text.length < 60) {
      console.log(`  slide=${slide.objectId} obj=${el.objectId} text="${text.replace(/\n/g, '\\n')}"`);
    }
  }
}
