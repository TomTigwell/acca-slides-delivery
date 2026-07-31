import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';

const { data } = await slides.presentations.get({ presentationId: B2B_DECK });
const slide = data.slides.find(s => s.objectId === 'p6');
const tableEl = slide.pageElements.find(el => el.objectId === 'p6_g11');
const table = tableEl.table;

console.log('table element size:', JSON.stringify(tableEl.size));
console.log('table element transform:', JSON.stringify(tableEl.transform));
console.log('num rows:', table.rows, 'num cols:', table.columns);
table.tableRows.forEach((row, ri) => {
  console.log(`row ${ri} rowHeight:`, JSON.stringify(row.rowHeight));
});
table.tableColumns.forEach((col, ci) => {
  console.log(`col ${ci} columnWidth:`, JSON.stringify(col.columnWidth));
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
  ['row1 col0 (April, bold month label)', 1, 0],
  ['row1 col1 (APAC region)', 1, 1],
  ['row1 col5 (£168 orange-ish CPL)', 1, 5],
  ['row3 col5 (£74 green CPL)', 3, 5],
  ['row5 col5 (£300 red CPL)', 5, 5],
  ['row7 col5 (No leads)', 7, 5],
  ['row8 col1 (UK — paused, grey/italic)', 8, 1],
  ['row8 col5 (— placeholder)', 8, 5],
  ['row9 col0 (Q1 Total, red bg)', 9, 0],
  ['row9 col5 (Q1 Total £221)', 9, 5],
];
for (const [label, ri, ci] of samples) {
  console.log(`\n${label}:`, JSON.stringify(cellInfo(ri, ci)));
}
