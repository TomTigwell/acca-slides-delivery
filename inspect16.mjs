import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';
const PX = 6350;

const { data } = await slides.presentations.get({ presentationId: B2B_DECK });
const slide = data.slides.find(s => s.objectId === 'p6');

for (const el of slide.pageElements || []) {
  const x = el.transform ? Math.round(el.transform.translateX / PX) : '?';
  const y = el.transform ? Math.round(el.transform.translateY / PX) : '?';
  const w = el.size ? Math.round((el.size.width.magnitude * (el.transform?.scaleX ?? 1)) / PX) : '?';
  const h = el.size ? Math.round((el.size.height.magnitude * (el.transform?.scaleY ?? 1)) / PX) : '?';
  if (el.table) {
    console.log(`TABLE ${el.objectId} box(${x},${y},${w},${h}) rows=${el.table.rows} cols=${el.table.columns}`);
    el.table.tableRows.forEach((row, ri) => {
      console.log(`  row ${ri}: rowHeight=${Math.round(row.rowHeight / PX)}px`);
    });
    el.table.tableColumns.forEach((col, ci) => {
      console.log(`  col ${ci}: width=${Math.round(col.columnWidth / PX)}px`);
    });
  } else {
    console.log(`${el.objectId} box(${x},${y},${w},${h})`);
  }
}
