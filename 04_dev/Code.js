/**
 * 웹 앱을 실행하기 위한 함수입니다.
 * @return {HtmlOutput} index.html 파일을 반환합니다.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('정보자산 위험 관리 절차서 - 자산 분류 기준표');
}

/**
 * HTML 템플릿의 스크립릿에서 호출될 수 있도록 파일을 가져오는 함수입니다.
 * @param {string} filename 불러올 파일의 이름 (예: 'index')
 * @return {string} 파일 내용을 반환합니다.
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/**
 * 정보자산 데이터를 스프레드시트에 저장합니다.
 * @param {Array<Array<string>>} data 저장할 자산 데이터 (자산번호, 자산유형, 자산분류, 자산명, 소유자, 관리자, 위치)
 * @return {string} 저장 성공 메시지를 반환합니다.
 */
function saveAssetData(data) {
  const sheetName = '자산분류기준표';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  // 시트가 없으면 새로 생성하고 헤더를 추가합니다.
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ['자산번호', '자산유형', '자산분류', '자산명', '소유자', '관리자', '위치'];
    sheet.appendRow(headers);
  } else {
    // 기존 데이터 삭제 (헤더 제외)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
    }
  }

  // 데이터 추가
  if (data.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, data.length, data[0].length).setValues(data);
  }
  
  return '자산 분류 데이터가 스프레드시트에 성공적으로 저장되었습니다.';
}

/**
 * 스프레드시트에서 정보자산 데이터를 불러옵니다.
 * @return {Array<Array<string>>} 스프레드시트의 모든 데이터를 반환합니다.
 */
function loadAssetData() {
  const sheetName = '자산분류기준표';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    return []; // 시트가 없으면 빈 배열 반환
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow <= 0 || lastColumn <= 0) {
    return []; // 데이터가 없으면 빈 배열 반환
  }

  // 헤더를 포함한 모든 데이터를 가져옵니다.
  return sheet.getRange(1, 1, lastRow, lastColumn).getValues();
}