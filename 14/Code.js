function doGet(e) {
  // 웹 앱으로 배포 시 index.html 파일을 서빙합니다.
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('문서관리 점검 체크리스트')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 필요에 따라 보안 설정 조정
}

function appendDataToSheet(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '문서관리점검체크리스트'; // 시트 이름 (원하는 이름으로 변경 가능)
  let sheet = ss.getSheetByName(sheetName);

  // 시트가 없으면 새로 생성하고 헤더를 추가합니다.
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    const headers = ['체크리스트 제목', '점검일', '점검자', '점검 구분', '점검 항목', '점검 기준', '점검 방법', '적합 여부', '미적합 여부', '개선사항', '저장일시'];
    sheet.appendRow(headers);
    sheet.setFrozenRows(1); // 첫 행을 고정
    
    // 헤더 스타일링 (선택 사항)
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground('#1A237E'); // 네이비 배경
    headerRange.setFontColor('#FFFFFF'); // 흰색 글자
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
  }

  const checkItems = data.checkItems;
  const timestamp = new Date();

  // 각 점검 항목을 시트에 한 줄씩 추가합니다.
  checkItems.forEach(item => {
    sheet.appendRow([
      data.documentTitle,
      data.checkDate,
      data.checkerName,
      item.checkDivision,
      item.checkItem,
      item.checkStandard,
      item.checkMethod,
      item.checkResult === '적합' ? 'O' : '', // 적합 여부
      item.checkResult === '미적합' ? 'O' : '', // 미적합 여부
      item.improvement,
      timestamp
    ]);
  });

  return '데이터가 스프레드시트에 성공적으로 저장되었습니다.';
}

// 이 함수는 외부에서 호출할 수 있도록 노출합니다.
// HTML 파일에서 google.script.run을 통해 호출될 것입니다.
function processFormSubmission(formData) {
  return appendDataToSheet(formData);
}