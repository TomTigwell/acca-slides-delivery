import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const PX = 6350;

const { data } = await slides.presentations.get({ presentationId: B2B_DECK });
const slide = data.slides.find(s => s.objectId === 'p6');
const tableEl = slide.pageElements.find(el => el.objectId === 'p6_g11');
const table = tableEl.table;

console.log('transform y (px):', Math.round(tableEl.transform.translateY / PX));
let total = 0;
table.tableRows.forEach((row, ri) => {
  const h = row.rowHeight.magnitude / PX;
  total += h;
  console.log(`row ${ri} height:`, Math.round(h), 'px');
});
console.log('total height px:', Math.round(total));
console.log('bottom y px:', Math.round(tableEl.transform.translateY / PX + total));

// check col0 background for rows 10,11,12
for (const ri of [10, 11, 12]) {
  const cell = table.tableRows[ri].tableCells[0];
  console.log(`row ${ri} col0 bg:`, JSON.stringify(cell.tableCellProperties?.tableCellBackgroundFill));
}
