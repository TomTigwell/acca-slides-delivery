import { google } from 'googleapis';

// Read-only pass: dumps the exact text content of every shape/table cell in
// each master deck, so replaceAllText targets can be built from ground truth
// (the Drive "read as text" export reformats whitespace and won't do).

const MASTERS = {
  product: '13RsJLhqAsTgowBdVnWqAtG460MSz49Ph7vcx3ecvnqo',
  b2b:     '1dZJV-dvJaRVUxTCYOoR2QQcyRa2N6LAIV780U8ezqHg',
  brand:   '12WS8qbsgSW8KrOLA9h2A9mjVLnf4aBAQuH4MY1s95P8',
};

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly', 'https://www.googleapis.com/auth/drive.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

function textOf(textContent) {
  if (!textContent || !textContent.textElements) return '';
  return textContent.textElements.map((te) => te.textRun?.content || '').join('');
}

function dumpElement(el) {
  if (el.shape?.text) {
    const t = textOf(el.shape.text);
    if (t.trim()) console.log(`  [shape ${el.objectId}] ${JSON.stringify(t)}`);
  }
  if (el.table) {
    el.table.tableRows.forEach((row, r) => {
      row.tableCells.forEach((cell, c) => {
        const t = textOf(cell.text);
        if (t.trim()) console.log(`  [table ${el.objectId} r${r}c${c}] ${JSON.stringify(t)}`);
      });
    });
  }
  if (el.elementGroup?.children) {
    el.elementGroup.children.forEach(dumpElement);
  }
}

for (const [name, id] of Object.entries(MASTERS)) {
  const { data } = await slides.presentations.get({ presentationId: id });
  console.log(`\n=== ${name} (${id}) — ${data.slides.length} slides ===`);
  data.slides.forEach((slide, i) => {
    console.log(`--- Slide ${i + 1} (${slide.objectId}) ---`);
    (slide.pageElements || []).forEach(dumpElement);
  });
}
