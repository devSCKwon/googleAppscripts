function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('위험처리방안 결정 양식')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveRiskTreatmentData(data) {
  const sheetName = '위험처리방안 결정 양식';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      '위험ID', '자산명', '위험등급', '위험도', '처리전략', '상세방안', 
      '구현일정', '담당자', '예산(만원)', '예상효과', '승인상태', '승인자', '승인일자'
    ];
    sheet.appendRow(headers);
  }

  // Clear existing data before appending new data (optional, but good for refresh)
  // If you want to append without clearing, comment out the following lines
  if (sheet.getLastRow() > 1) { // If there's data beyond headers
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }

  // Append new data
  if (data.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  }
  
  return '데이터가 스프레드시트에 성공적으로 저장되었습니다!';
}

function loadRiskTreatmentData() {
  const sheetName = '위험처리방안 결정 양식';
  Logger.log('Searching for sheet: %s', sheetName);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  Logger.log('Sheet found: %s', sheet ? true : false);
  if (!sheet) {
    Logger.log('Sheet not found: %s', sheetName);
    return JSON.stringify([]); // Return empty array as JSON string if sheet does not exist
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  Logger.log('Data range dimensions: %d rows, %d columns', values.length, values[0] ? values[0].length : 0);
  Logger.log('Returning data (first 5 rows): %s', JSON.stringify(values.slice(0, 5)));
  return JSON.stringify(values); // Return data as JSON string
}