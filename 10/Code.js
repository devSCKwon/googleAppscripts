function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('핵심문서통제이력관리양식')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function saveDocumentControlData(data) {
  const sheetName = '핵심문서통제이력관리양식';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    const headers = [
      '문서ID', '문서명', '문서유형', '보안등급', '중요도점수', '소유부서',
      '관리자', '핵심문서여부', '선정사유', '적용통제방안', '적용일자',
      '최종점검일', '상태', '특이사항'
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

function loadDocumentControlData() {
  const sheetName = '핵심문서통제이력관리양식';
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    return []; // Return empty array if sheet does not exist
  }

  const range = sheet.getDataRange();
  const values = range.getValues();
  return values;
}