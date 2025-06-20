function doGet(e) {
  // 웹 앱으로 배포 시 index.html 파일을 서빙합니다.
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('문서관리 점검 체크리스트')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // 필요에 따라 보안 설정 조정
}

function appendDataToSheet(data) {
  try {
    if (!data.checkItems || data.checkItems.length === 0) {
      return '저장할 점검 항목 데이터가 없습니다. 점검 항목을 추가해주세요.';
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = '문서관리점검체크리스트';
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = ['체크리스트 제목', '점검일', '점검자', '점검 구분', '점검 항목', '점검 기준', '점검 방법', '적합 여부', '미적합 여부', '개선사항', '저장일시'];
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#1A237E').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
    }

    const checkItems = data.checkItems;
    const timestamp = new Date();
    let savedCount = 0;

    checkItems.forEach(item => {
      sheet.appendRow([
        data.documentTitle,
        data.checkDate,
        data.checkerName,
        item.checkDivision,
        item.checkItem,
        item.checkStandard,
        item.checkMethod,
        item.checkResult === '적합' ? 'O' : '',
        item.checkResult === '미적합' ? 'O' : '',
        item.improvement,
        timestamp
      ]);
      savedCount++;
    });

    if (savedCount > 0) {
      return `${savedCount}개의 점검 항목이 포함된 데이터가 스프레드시트에 성공적으로 저장되었습니다.`;
    } else {
      // 이 경우는 checkItems는 있었지만 forEach가 실행되지 않은 극히 드문 경우 대비
      return '저장할 유효한 점검 항목이 없어 실제 저장된 데이터는 없습니다.';
    }

  } catch (e) {
    // 오류 발생 시 로그 기록 (선택 사항)
    // console.error("appendDataToSheet Error: " + e.toString());
    return '데이터 저장 중 오류가 발생했습니다: ' + e.message;
  }
}

// 이 함수는 외부에서 호출할 수 있도록 노출합니다.
// HTML 파일에서 google.script.run을 통해 호출될 것입니다.
function processFormSubmission(formData) {
  return appendDataToSheet(formData);
}

function getDataFromSheet() {
  const sheetName = '문서관리점검체크리스트'; // 대상 시트 이름
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    // 시트가 없으면 빈 배열을 JSON 문자열로 반환
    return JSON.stringify([]);
  }

  // 시트에 데이터가 헤더만 있거나 아예 없는 경우 빈 배열 반환 처리
  if (sheet.getLastRow() <= 1) {
    return JSON.stringify([]);
  }

  // 데이터 범위에서 값을 가져옴 (헤더 제외)
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();

  // Apps Script 환경에서는 JavaScript 객체를 직접 반환하면 클라이언트 스크립트에서 객체로 받을 수 있습니다.
  // 그러나 일관성을 위해 10번 폴더처럼 JSON 문자열로 반환하겠습니다.
  // 또한, 키-값 쌍으로 매핑하여 반환해야 클라이언트에서 사용하기 용이합니다.
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = values.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });

  return JSON.stringify(data);
}