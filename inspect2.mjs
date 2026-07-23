import { google } from 'googleapis';

// Read-only: page size + full pageElement inventory (including images, which
// inspect.mjs's shape/table-only dump skips) for slide 1 of each deck, plus
// the North Star / table slides needed to extend them with July rows.

const MASTERS = {
  product: '13RsJLhqAsTgowBdVnWqAtG460MSz49Ph7vcx3ecvnqo',
  b2b:     '1dZJV-dvJaRVUxTCYOoR2QQcyRa2N6LAIV780U8ezqHg',
  brand:   '12WS8qbsgSW8KrOLA9h2A9mjVLnf4aBAQuH4MY1s95P8',
};

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations.readonly'],
});
const slides = google.slides({ version: 'v1', auth });

function describeElement(el) {
  const size = el.size ? `${el.size.width?.magnitude}x${el.size.height?.magnitude}${el.size.width?.unit || ''}` : 'no-size';
  const t = el.transform ? `scale(${el.transform.scaleX},${el.transform.scaleY}) translate(${el.transform.translateX},${el.transform.translateY}) unit=${el.transform.unit}` : 'no-transform';
  let kind = 'unknown';
  if (el.shape) kind = `shape:${el.shape.shapeType}`;
  if (el.image) kind = 'image';
  if (el.table) kind = `table:${el.table.rows}x${el.table.columns}`;
  if (el.line) kind = 'line';
  if (el.elementGroup) kind = 'group';
  return `[${el.objectId}] ${kind} size=${size} ${t}`;
}

for (const [name, id] of Object.entries(MASTERS)) {
  const { data } = await slides.presentations.get({ presentationId: id });
  console.log(`\n=== ${name} (${id}) ===`);
  console.log('pageSize:', JSON.stringify(data.pageSize));
  const slide1 = data.slides[0];
  console.log(`Slide 1 (${slide1.objectId}) elements:`);
  (slide1.pageElements || []).forEach((el) => console.log(' ', describeElement(el)));
}
