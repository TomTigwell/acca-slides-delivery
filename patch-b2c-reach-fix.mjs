import { google } from 'googleapis';

// Follow-up fix: the Reach slide has its OWN "JULY REGISTERS" stat tile
// (g3f59aaca611_0_7/8/9), separate from the North Star slide's version
// (g3f60cf2e19e_0_1/2) — missed on the first pass since generate-b2c.mjs
// defines them as two different object groups. Still showed the old
// 1,092 / "Six B2C campaigns" after the July refresh.

const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations'],
});
const slides = google.slides({ version: 'v1', auth });

function edit(objectId, text) {
  return [
    { deleteText: { objectId, textRange: { type: 'ALL' } } },
    { insertText: { objectId, insertionIndex: 0, text } },
  ];
}

const requests = [
  ...edit('g3f59aaca611_0_8', '1,106\n'),
  ...edit('g3f59aaca611_0_9', 'Eight B2C campaigns\n'),
];

await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
console.log('Reach slide JULY REGISTERS tile fixed');
