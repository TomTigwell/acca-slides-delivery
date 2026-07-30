import { google } from 'googleapis';
import fs from 'fs';

const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const PX = 6350; // EMU per px

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

async function uploadPublicImage(localPath, name) {
  const { data } = await drive.files.create({
    requestBody: { name, parents: [SHARED_DRIVE] },
    media: { mimeType: 'image/png', body: fs.createReadStream(localPath) },
    fields: 'id',
    supportsAllDrives: true,
  });
  await drive.permissions.create({
    fileId: data.id,
    requestBody: { role: 'reader', type: 'anyone' },
    supportsAllDrives: true,
  });
  return `https://drive.google.com/uc?export=view&id=${data.id}`;
}

function imageBox(xPx, yPx, wPx, hPx) {
  return {
    size: { width: { magnitude: wPx * PX, unit: 'EMU' }, height: { magnitude: hPx * PX, unit: 'EMU' } },
    transform: { scaleX: 1, scaleY: 1, translateX: xPx * PX, translateY: yPx * PX, unit: 'EMU' },
  };
}

function textRequests(objectId, slideId, box, text, { fontFamily, fontSize, bold = false, color, align = 'START', lineSpacing }) {
  return [
    { createShape: { objectId, shapeType: 'TEXT_BOX', elementProperties: { pageObjectId: slideId, ...imageBox(...box) } } },
    { insertText: { objectId, insertionIndex: 0, text } },
    {
      updateTextStyle: {
        objectId,
        style: { fontFamily, fontSize: { magnitude: fontSize, unit: 'PT' }, bold, foregroundColor: { opaqueColor: { rgbColor: color } } },
        textRange: { type: 'ALL' },
        fields: 'fontFamily,fontSize,bold,foregroundColor',
      },
    },
    {
      updateParagraphStyle: {
        objectId,
        style: { alignment: align, ...(lineSpacing ? { lineSpacing } : {}) },
        textRange: { type: 'ALL' },
        fields: lineSpacing ? 'alignment,lineSpacing' : 'alignment',
      },
    },
    { updateShapeProperties: { objectId, shapeProperties: { outline: { propertyState: 'NOT_RENDERED' } }, fields: 'outline.propertyState' } },
  ];
}

const BRAND_DECK = '1l2Q2xRhwhxfghSX6CO4N8LvfW8d_ZhEa7snEHStp-Kw';
const B2C_DECK = '1vZNoUPU2_fDuwC85IXxarLxKMDOWnZW7CMQLdgbfVuU';
const B2B_DECK = '1j2Av4AjF84sLOc2_RkV6ffxymwj_2DYwbmFxMFzjIcA';

console.log('uploading chart images...');
const brandUrl = await uploadPublicImage('assets/brand_chart.png', 'brand_northstar_chart.png');
const b2cUrl = await uploadPublicImage('assets/b2c_chart.png', 'b2c_northstar_chart.png');
const b2bUrl = await uploadPublicImage('assets/b2b_chart.png', 'b2b_northstar_chart.png');
console.log('uploaded:', { brandUrl, b2cUrl, b2bUrl });

// ---- Brand: replace stale note with the May-Jul chart ----
{
  const requests = [
    { deleteObject: { objectId: 'brand_ns_note2' } },
    {
      createImage: {
        objectId: 'brand_ns_chart_img',
        url: brandUrl,
        elementProperties: { pageObjectId: 'g3f483bfcec4_0_95', ...imageBox(87, 395, 739, 320) },
      },
    },
  ];
  await slides.presentations.batchUpdate({ presentationId: BRAND_DECK, requestBody: { requests } });
  console.log('Brand: chart inserted');
}

// ---- B2C: replace stale note with the May-Jul chart ----
{
  const requests = [
    { deleteObject: { objectId: 'b2c_ns_note' } },
    {
      createImage: {
        objectId: 'b2c_ns_chart_img',
        url: b2cUrl,
        elementProperties: { pageObjectId: 'g3f483bfcec4_0_95', ...imageBox(87, 395, 739, 320) },
      },
    },
  ];
  await slides.presentations.batchUpdate({ presentationId: B2C_DECK, requestBody: { requests } });
  console.log('B2C: chart inserted');
}

// ---- B2B: existing North Star slide (p4) is a full quarterly chart with no
// room for a monthly breakdown — add a new slide immediately after it (index 3)
// with the May-Jul monthly trend instead of disturbing the user's existing
// quarterly chart/cards. ----
{
  const newSlideId = 'b2b_trend_slide';
  const requests = [
    { createSlide: { objectId: newSlideId, insertionIndex: 3 } },
    ...textRequests('b2b_trend_title', newSlideId, [65, 40, 1310, 70],
      'North Star — Monthly Trend (May–Jul 2026)',
      { fontFamily: 'Lora', fontSize: 28, bold: true, color: { red: 0.1, green: 0.1, blue: 0.1 } }),
    ...textRequests('b2b_trend_caption', newSlideId, [65, 118, 1310, 50],
      'ProDipSust B2B Lead Generation — APAC + EMEA + UK combined. The three months sum to 21 leads, matching the Q1 FY27 total on the previous slide.',
      { fontFamily: 'DM Sans', fontSize: 12, color: { red: 0.33, green: 0.33, blue: 0.33 } }),
    {
      createImage: {
        objectId: 'b2b_trend_chart_img',
        url: b2bUrl,
        elementProperties: { pageObjectId: newSlideId, ...imageBox(320, 190, 800, 487) },
      },
    },
  ];
  await slides.presentations.batchUpdate({ presentationId: B2B_DECK, requestBody: { requests } });
  console.log('B2B: new trend slide inserted');
}

console.log('done');
