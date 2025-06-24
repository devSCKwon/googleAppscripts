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
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === "") {
    Logger.log('createNewSheetIfNotExists: Invalid sheetName received - ' + sheetName);
    return {status: "error", message: "시트 이름이 유효하지 않습니다. (받은 값: " + sheetName + ")"};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    Logger.log("createNewSheetIfNotExists: Checking for sheet '" + sheetName + "'. Found: " + (sheet ? 'Yes' : 'No'));

    if (!sheet) {
      ss.insertSheet(sheetName);
      Logger.log("createNewSheetIfNotExists: Sheet '" + sheetName + "' created.");
      return {status: "created", message: "'" + sheetName + "' 시트가 새로 생성되었습니다."};
    } else {
      Logger.log("createNewSheetIfNotExists: Sheet '" + sheetName + "' already exists.");
      return {status: "exists", message: "'" + sheetName + "' 시트는 이미 존재합니다."};
    }
  } catch (e) {
    Logger.log('createNewSheetIfNotExists Error for sheet "' + sheetName + '": ' + e.toString() + "\nStack: " + e.stack);
    return {status: "error", message: "시트 생성/확인 중 오류 발생: " + e.message};
  }
}

/**
 * 지정된 이름의 시트 존재 여부를 확인합니다. (클라이언트에서 명시적으로 사용하지 않을 수 있음)
 * @param {string} sheetName 확인할 시트의 이름.
 * @return {object} 작업 결과 객체 (status, exists, message).
 */
function checkSheetExists(sheetName) {
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === "") {
    return {status: "error", exists: false, message: "시트 이름이 유효하지 않습니다."};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    Logger.log("checkSheetExists: Checking for sheet '" + sheetName + "'. Found: " + (sheet ? 'Yes' : 'No'));
    if (sheet) {
      return {status: "exists", exists: true, message: "'" + sheetName + "' 시트가 존재합니다."};
    } else {
      return {status: "not_exists", exists: false, message: "'" + sheetName + "' 시트가 존재하지 않습니다."};
    }
  } catch (e) {
    Logger.log('checkSheetExists Error for sheet "' + sheetName + '": ' + e.toString());
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
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === "") {
    return {status: "error", message: "시트 이름이 유효하지 않습니다."};
  }
  if (!dataToSave || !Array.isArray(dataToSave) || dataToSave.length === 0) {
    return {status: "error", message: "저장할 데이터가 없거나 형식이 올바르지 않습니다."};
  }
  // 데이터의 첫 번째 행이 배열인지, 그리고 그 배열이 비어있지 않은지 추가로 확인 (컬럼 수 확인을 위해)
  if (!Array.isArray(dataToSave[0]) || dataToSave[0].length === 0) {
    return {status: "error", message: "저장할 데이터의 헤더 정보가 올바르지 않습니다."};
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    Logger.log("saveAssetDataToSheet: Attempting to save to sheet '" + sheetName + "'. Found: " + (sheet ? 'Yes' : 'No'));

    if (!sheet) {
      Logger.log("saveAssetDataToSheet: Sheet '" + sheetName + "' not found.");
      return {status: "error", message: "'" + sheetName + "' 시트를 찾을 수 없습니다. '절차서 생성하기'를 통해 먼저 시트를 만들어주세요."};
    }

    sheet.clearContents();
    Logger.log("saveAssetDataToSheet: Cleared contents of sheet '" + sheetName + "'.");

    sheet.getRange(1, 1, dataToSave.length, dataToSave[0].length).setValues(dataToSave);
    Logger.log("saveAssetDataToSheet: Data saved to sheet '" + sheetName + "'. Rows: " + dataToSave.length);

    return {status: "success", message: "'" + sheetName + "' 시트에 데이터가 성공적으로 저장되었습니다."};
  } catch (e) {
    Logger.log('saveAssetDataToSheet Error for sheet "' + sheetName + '": ' + e.toString() + "\nStack: " + e.stack);
    return {status: "error", message: "데이터 저장 중 오류 발생: " + e.message};
  }
}

/**
 * 지정된 시트에서 데이터를 불러옵니다.
 * @param {string} sheetName 데이터를 불러올 시트의 이름.
 * @return {object} 작업 결과 객체 (status, message, data).
 */
function loadAssetDataFromSheet(sheetName) {
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === "") {
    return {status: "error", message: "시트 이름이 유효하지 않습니다.", data: []};
  }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    Logger.log("loadAssetDataFromSheet: Attempting to load from sheet '" + sheetName + "'. Found: " + (sheet ? 'Yes' : 'No'));

    if (!sheet) {
      Logger.log("loadAssetDataFromSheet: Sheet '" + sheetName + "' not found.");
      return {status: "no_sheet", message: "'" + sheetName + "' 시트를 찾을 수 없습니다.", data: [] };
    }

    if (sheet.getLastRow() <= 1 && sheet.getLastColumn() === 0) {
       Logger.log("loadAssetDataFromSheet: Sheet '" + sheetName + "' is empty or has only a header.");
      return {status: "no_data", message: "'" + sheetName + "' 시트에 저장된 데이터가 없습니다.", data: []};
    }

    var dataRange = sheet.getDataRange();
    if (!dataRange) {
        Logger.log("loadAssetDataFromSheet: No data range in sheet '" + sheetName + "'.");
        return {status: "no_data", message: "'" + sheetName + "' 시트에 저장된 데이터가 없습니다.", data: []};
    }
    var data = dataRange.getValues();

    if (data.length <= 1 && (data.length === 0 || (data.length === 1 && data[0].every(cell => cell === "")))) {
        Logger.log("loadAssetDataFromSheet: Sheet '" + sheetName + "' effectively has no data (or only empty header).");
        return {status: "no_data", message: "'" + sheetName + "' 시트에 저장된 데이터가 없습니다.", data: []};
    }

    Logger.log("loadAssetDataFromSheet: Data loaded from sheet '" + sheetName + "'. Rows: " + data.length);
    return {status: "success", message: "'" + sheetName + "' 시트에서 데이터를 성공적으로 불러왔습니다.", data: data};
  } catch (e) {
    Logger.log('loadAssetDataFromSheet Error for sheet "' + sheetName + '": ' + e.toString() + "\nStack: " + e.stack);
    return {status: "error", message: "데이터 불러오기 중 오류 발생: " + e.message, data: []};
  }
}
