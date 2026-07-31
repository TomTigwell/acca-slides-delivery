import { google } from 'googleapis';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
const PX = 6350;

const { data } = await slides.presentations.get({ presentationId: B2C_DECK });
const slide = data.slides.find(s => s.objectId === 'g3f59aaca611_0_64');
const tableEl = slide.pageElements.find(el => el.objectId === 'g3f59aaca611_0_73');

console.log('table transform y (px):', Math.round(tableEl.transform.translateY / PX));
let total = 0;
tableEl.table.tableRows.forEach((row, ri) => {
  const h = row.rowHeight.magnitude / PX;
  total += h;
  console.log(`row ${ri} height:`, Math.round(h), 'px');
});
console.log('total height px:', Math.round(total));
console.log('bottom y px:', Math.round(tableEl.transform.translateY / PX + total));
