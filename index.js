const GoogleSheetApi = require('./src/lib/googleSheetApi');
const credentials = require('./credentials.json');

(async () => {
  const googleSheetId = '<sheet_id>';
  const sheet = new GoogleSheetApi(credentials, googleSheetId);
  
  const dataHeader = ['Name', 'Phone', 'Value', 'Date'];
  const data = [
    ['Mary', '973673-392', 1, '1980-07-12T14:00:20.000Z'],
    ['John', '588745-543', 2, '2019-12-25']
  ]
  const sheetTabId = 0;

  try {
    await sheet.addHeader(sheetTabId, dataHeader);
    await sheet.addRows(sheetTabId, data);
    
    await sheet.changeFormat(sheetTabId, 2, ['BOLD','INTEGER']);
    await sheet.changeFormat(sheetTabId, 3, ['BR_DATE']);

  } catch(e) {
    console.log('ERROR', e);
  }

})();

