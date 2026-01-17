/**
 * 초기 앱 실행 시 데이터 가져오기
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  initSheet(ss, SHEETS.RAW_MATERIALS, ['날짜', '품목명', '구분', '수량', '로트번호', '담당자', '내역']);
  initSheet(ss, SHEETS.PRODUCTION, ['날짜', '제품명', '생산량', '작업자', '비고']);
  initSheet(ss, SHEETS.SEMI_FINISHED, ['날짜', '반제품명', '입고', '사용', '재고']);
  initSheet(ss, SHEETS.PACKAGING, ['날짜', '제품명', '포장수량', '검수자', '상태']);
  
  // [UPDATED] 헤더에 '상태' 추가 (기존 시트가 있다면 덮어쓰지 않음, 신규 생성 시 적용)
  initSheet(ss, SHEETS.ITEMS, ['구분', '품목명', '단위', '상태']); 
  
  let settingSheet = ss.getSheetByName(SHEETS.SETTINGS);
  let companyName = "Woondooran Corp"; 
  
  if (!settingSheet) {
    settingSheet = ss.insertSheet(SHEETS.SETTINGS);
    settingSheet.appendRow(['Key', 'Value']);
    settingSheet.appendRow(['Company Name', companyName]);
    settingSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#E5E7EB');
  } else {
    const val = settingSheet.getRange(2, 2).getValue();
    if (val) companyName = val;
  }

  const itemSheet = ss.getSheetByName(SHEETS.ITEMS);
  let rawMaterials = [];
  if (itemSheet.getLastRow() > 1) {
    // [UPDATED] 4번째 열(상태)까지 데이터 읽기
    // 데이터 범위가 기존 3열에서 4열로 확장됨
    const lastCol = itemSheet.getLastColumn();
    // 안전하게 데이터를 가져오기 위해 열 개수 확인
    const numCols = lastCol < 4 ? 3 : 4;
    const data = itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, numCols).getValues();
    
    rawMaterials = data.filter(row => row[0] === '원재료').map(row => ({
      name: row[1],
      unit: row[2],
      // 4번째 열이 없거나 비어있으면 기본값 '사용'
      status: (row[3] || '사용').toString() 
    }));
  }

  return {
    companyName: companyName,
    currentUser: Session.getActiveUser().getEmail(),
    rawMaterials: rawMaterials
  };
}

/**
 * [품목관리] 새로운 품목 등록
 */
function addItem(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.ITEMS);
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === formObj.itemName && data[i][0] === formObj.category) {
      throw new Error("이미 등록된 품목명입니다.");
    }
  }

  // [UPDATED] 신규 등록 시 '상태' 컬럼에 '사용' 자동 입력
  sheet.appendRow([
    formObj.category,
    formObj.itemName,
    formObj.unit,
    '사용'
  ]);
  
  return "품목이 등록되었습니다.";
}

/**
 * [원재료] 데이터 불러오기
 */
function getMaterials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  if (sheet.getLastRow() < 2) {
    return { history: [], summary: [] };
  }
  
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  const history = values.map((row, index) => ({
    id: index + 2,
    date: formatDate(row[0]),
    item: row[1],
    type: row[2], 
    qty: Number(row[3]) || 0,
    lot: row[4] || '', 
    manager: row[5],
    desc: row[6]      
  })).reverse();

  const stockMap = {};
  values.forEach(row => {
    const item = row[1];
    const type = row[2];
    const qty = Number(row[3]) || 0;
    
    if (!stockMap[item]) stockMap[item] = 0;
    
    if (type === '입고') {
      stockMap[item] += qty;
    } else if (type === '출고') {
      stockMap[item] -= qty;
    }
  });

  const summary = Object.keys(stockMap).map(key => ({
    item: key,
    total: stockMap[key]
  }));

  return {
    history: history,
    summary: summary
  };
}

/**
 * [원재료] 데이터 저장
 */
function addMaterial(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  const rowData = [
    formObj.date,
    formObj.item,
    formObj.type,
    formObj.qty,
    formObj.lot, 
    Session.getActiveUser().getEmail(),
    formObj.desc 
  ];
  
  sheet.appendRow(rowData);
  return "원재료가 저장되었습니다.";
}

/**
 * [원재료] 데이터 수정
 */
function editRawMaterial(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  const rowIndex = parseInt(formObj.rowIndex);

  if (isNaN(rowIndex) || rowIndex < 2) {
    throw new Error("유효하지 않은 데이터 ID입니다.");
  }

  const rowData = [
    formObj.date,
    formObj.item,
    formObj.type,
    formObj.qty,
    formObj.lot,
    Session.getActiveUser().getEmail(),
    formObj.desc
  ];

  sheet.getRange(rowIndex, 1, 1, 7).setValues([rowData]);

  return "원재료 내역이 수정되었습니다.";
}

/**
 * [원재료] 데이터 삭제
 */
function deleteRawMaterial(rowIndex) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  const rIndex = parseInt(rowIndex);
  
  if (isNaN(rIndex) || rIndex < 2) {
    throw new Error("삭제할 수 없는 데이터입니다.");
  }
  
  sheet.deleteRow(rIndex);
  return "내역이 삭제되었습니다.";
}

/**
 * [생산일지] 데이터 불러오기
 */
function getProduction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTION);
  
  if (sheet.getLastRow() < 2) return [];
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return data.map(row => ({
    date: formatDate(row[0]),
    product: row[1],
    qty: row[2],
    worker: row[3],
    memo: row[4]
  })).reverse(); 
}

/**
 * [생산일지] 데이터 저장하기
 */
function addProduction(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.PRODUCTION);
  
  const rowData = [
    formObj.date,
    formObj.product,
    formObj.qty,
    Session.getActiveUser().getEmail(), 
    formObj.memo
  ];
  
  sheet.appendRow(rowData);
  return "생산실적이 등록되었습니다.";
}

function initSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f4f6');
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  }
  return sheet;
}

function formatDate(date) {
  if (!date) return "";
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}