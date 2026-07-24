import { google } from 'googleapis';

// Read-only: check whether the service account can even see the North Star
// chart's backing spreadsheet (it lives outside the Shared Drive the rest
// of this pipeline operates in — parentId differs from SHARED_DRIVE), and
// if so, dump its exact sheet/tab name + data layout + embedded chart spec
// so a write can target the right range.

const SPREADSHEET_ID = '1EvMo7kOGVIcUss9B5TvJo5AjJsMTziWOWsTNu8pKZ5s';

const auth = new google.auth.GoogleAuth({
  scopes: [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/drive.readonly',
  ],
});
const sheets = google.sheets({ version: 'v4', auth });

try {
  const { data } = await sheets.spreadsheets.get({
    spreadsheetId: SPREADSHEET_ID,
    includeGridData: true,
  });
  console.log('ACCESS OK. Title:', data.properties?.title);
  for (const sheet of data.sheets || []) {
    console.log(`--- Tab: "${sheet.properties.title}" (sheetId=${sheet.properties.sheetId}) ---`);
    const grid = sheet.data?.[0]?.rowData || [];
    grid.forEach((row, r) => {
      const cells = (row.values || []).map((c) => c.formattedValue ?? '').join(' | ');
      console.log(`  row ${r}: ${cells}`);
    });
    if (sheet.charts) {
      sheet.charts.forEach((chart) => {
        console.log(`  [CHART chartId=${chart.chartId}] spec=${JSON.stringify(chart.spec).slice(0, 500)}`);
      });
    }
  }
} catch (err) {
  console.log('ACCESS FAILED:', err.message);
}
