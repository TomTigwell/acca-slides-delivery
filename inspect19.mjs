import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const PX = 6350;

const { data } = await slides.presentations.get({ presentationId: B2B_DECK });
const slide = data.slides.find(s => s.objectId === 'g3f393228d3a_0_154');

for (const el of slide.pageElements || []) {
  const x = el.transform ? Math.round(el.transform.translateX / PX) : '?';
  const y = el.transform ? Math.round(el.transform.translateY / PX) : '?';
  console.log(`${el.objectId} @ (${x},${y})`);
}

const tableEl = slide.pageElements.find(el => el.objectId === 'g3f393228d3a_0_160');
const table = tableEl.table;
console.log('\ntable transform:', JSON.stringify(tableEl.transform));
console.log('num rows:', table.rows, 'num cols:', table.columns);
table.tableColumns.forEach((col, ci) => {
  console.log(`col ${ci} width:`, Math.round(col.columnWidth.magnitude / PX), 'px');
});
table.tableRows.forEach((row, ri) => {
  console.log(`row ${ri} height:`, Math.round(row.rowHeight.magnitude / PX), 'px');
});

function cellInfo(ri, ci) {
  const cell = table.tableRows[ri].tableCells[ci];
  const bg = cell.tableCellProperties?.tableCellBackgroundFill;
  const textEl = cell.text?.textElements?.find(te => te.textRun);
  const style = textEl?.textRun?.style;
  const content = cell.text?.textElements?.map(te => te.textRun?.content || '').join('') || '';
  return { content: content.replace(/\n/g, '\\n'), bg, style };
}

const samples = [
  ['row0 col2 (header "Apr")', 0, 2],
  ['row1 col2 (APAC £1,012 normal)', 1, 2],
  ['row1 col5 (13 Q1 Leads)', 1, 5],
  ['row1 col6 (£133 Q1 CPL)', 1, 6],
  ['row4 col0 (Q1 B2B Total label)', 4, 0],
  ['row4 col2 (£2,501 total)', 4, 2],
  ['row4 col5 (21 total leads)', 4, 5],
];
for (const [label, ri, ci] of samples) {
  console.log(`\n${label}:`, JSON.stringify(cellInfo(ri, ci)));
}
