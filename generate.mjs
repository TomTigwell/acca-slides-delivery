import { google } from 'googleapis';

const SHARED_DRIVE = '0AKpoLPF9OkBJUk9PVA';
const MASTERS = {
  product: '13RsJLhqAsTgowBdVnWqAtG460MSz49Ph7vcx3ecvnqo',
  b2b:     '1dZJV-dvJaRVUxTCYOoR2QQcyRa2N6LAIV780U8ezqHg',
  brand:   '12WS8qbsgSW8KrOLA9h2A9mjVLnf4aBAQuH4MY1s95P8',
};

const auth = new google.auth.GoogleAuth({
  scopes: ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive'],
});
const drive = google.drive({ version: 'v3', auth });
const slides = google.slides({ version: 'v1', auth });

async function generate(masterId, title, replacements = {}) {
  const { data } = await drive.files.copy({
    fileId: masterId,
    supportsAllDrives: true,
    requestBody: { name: title, parents: [SHARED_DRIVE] },
  });
  const requests = Object.entries(replacements).map(([t, v]) => ({
    replaceAllText: { containsText: { text: t, matchCase: true }, replaceText: v },
  }));
  if (requests.length) {
    await slides.presentations.batchUpdate({ presentationId: data.id, requestBody: { requests } });
  }
  console.log(title, '→', `https://docs.google.com/presentation/d/${data.id}/edit`);
}

// STEP 1 — prove the pipe: copy one master into the Shared Drive, no edits yet.
await generate(MASTERS.b2b, 'ACCA B2B — July 2026 (draft)');
