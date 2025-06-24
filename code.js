// Google Apps Script (Server-side logic)

/**
 * 웹 앱의 기본 GET 요청을 처리합니다.
 * @param {Object} e 이벤트 객체.
 * @return {HtmlOutput} HTML 서비스 출력.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('정보자산 위험 관리 절차서')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * HTML 템플릿 내에서 다른 파일을 포함시키기 위한 헬퍼 함수.
 * (index.html이 모든 JS/CSS를 포함하고 있다면 직접 사용되지 않을 수 있음)
 * @param {string} filename 포함할 파일의 이름.
 * @return {string} 파일 내용.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * 지정된 이름의 시트가 없으면 새로 생성하고, 결과를 반환합니다.
 * @param {string} sheetName 생성하거나 확인할 시트의 이름.
 * @return {object} 작업 결과 객체 (status, message).
 */
function createNewSheetIfNotExists(sheetName) {
  if (!sheetName) {
    return {status: "error", message: "시트 이름이 제공되지 않았습니다."};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      ss.insertSheet(sheetName);
      return {status: "created", message: "'" + sheetName + "' 시트가 새로 생성되었습니다."};
    } else {
      return {status: "exists", message: "'" + sheetName + "' 시트는 이미 존재합니다."};
    }
  } catch (e) {
    Logger.log('createNewSheetIfNotExists Error: ' + e.message);
    return {status: "error", message: "시트 생성/확인 중 오류 발생: " + e.message};
  }
}

/**
 * 지정된 이름의 시트 존재 여부를 확인합니다.
 * @param {string} sheetName 확인할 시트의 이름.
 * @return {object} 작업 결과 객체 (status, exists, message).
 */
function checkSheetExists(sheetName) {
  if (!sheetName) {
    return {status: "error", exists: false, message: "시트 이름이 제공되지 않았습니다."};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      return {status: "exists", exists: true, message: "'" + sheetName + "' 시트가 존재합니다."};
    } else {
      return {status: "not_exists", exists: false, message: "'" + sheetName + "' 시트가 존재하지 않습니다."};
    }
  } catch (e) {
    Logger.log('checkSheetExists Error: ' + e.message);
    return {status: "error", exists: false, message: "시트 존재 확인 중 오류 발생: " + e.message};
  }
}

/**
 * 지정된 시트에 데이터를 저장합니다. (기존 내용 삭제 후 저장)
 * @param {string} sheetName 데이터를 저장할 시트의 이름.
 * @param {Array<Array<String>>} dataToSave 저장할 2차원 배열 데이터 (첫 행은 헤더).
 * @return {object} 작업 결과 객체 (status, message).
 */
function saveAssetDataToSheet(sheetName, dataToSave) {
  if (!sheetName) {
    return {status: "error", message: "시트 이름이 제공되지 않았습니다."};
  }
  if (!dataToSave || dataToSave.length === 0) {
    return {status: "error", message: "저장할 데이터가 없습니다."};
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      // 시트가 없으면 새로 만들고 저장할 수도 있지만, 여기서는 클라이언트에서 checkSheetExists 후 호출하는 것을 가정.
      // 또는, createNewSheetIfNotExists를 내부적으로 호출할 수도 있음.
      // 현재 계획에서는 클라이언트가 시트 존재를 먼저 확인하거나, '절차서 생성하기'로 시트를 만들도록 유도.
      return {status: "error", message: "'" + sheetName + "' 시트를 찾을 수 없습니다. '절차서 생성하기'를 통해 먼저 시트를 만들어주세요."};
    }

    // 기존 내용 삭제 (헤더 포함 모든 데이터)
    sheet.clearContents();

    // 데이터 쓰기
    sheet.getRange(1, 1, dataToSave.length, dataToSave[0].length).setValues(dataToSave);

    return {status: "success", message: "'" + sheetName + "' 시트에 데이터가 성공적으로 저장되었습니다."};
  } catch (e) {
    Logger.log('saveAssetDataToSheet Error: ' + e.message + ' (Sheet: ' + sheetName + ')');
    return {status: "error", message: "데이터 저장 중 오류 발생: " + e.message};
  }
}

/**
 * 지정된 시트에서 데이터를 불러옵니다.
 * @param {string} sheetName 데이터를 불러올 시트의 이름.
 * @return {object} 작업 결과 객체 (status, message, data).
 */
function loadAssetDataFromSheet(sheetName) {
  if (!sheetName) {
    return {status: "error", message: "시트 이름이 제공되지 않았습니다."};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return {status: "no_sheet", message: "'" + sheetName + "' 시트를 찾을 수 없습니다.", data: [] };
    }

    // getLastRow()가 0을 반환하면 (완전히 빈 시트) 또는 1을 반환하면 (헤더만 있을 경우) 데이터가 없는 것으로 간주
    if (sheet.getLastRow() <= 1) {
      return {status: "no_data", message: "'" + sheetName + "' 시트에 저장된 데이터가 없습니다.", data: []};
    }

    var data = sheet.getDataRange().getValues();
    // 클라이언트에서 data.slice(1)을 사용하여 헤더를 제외하므로, 여기서는 전체 데이터를 전달합니다.
    // 또는 여기서 data.slice(1)을 해서 순수 데이터만 전달할 수도 있습니다. (클라이언트 로직과 일관성 유지)
    return {status: "success", message: "'" + sheetName + "' 시트에서 데이터를 성공적으로 불러왔습니다.", data: data};
  } catch (e) {
    Logger.log('loadAssetDataFromSheet Error: ' + e.message);
    return {status: "error", message: "데이터 불러오기 중 오류 발생: " + e.message, data: []};
  }
}
