/**
 * 초기 앱 실행 시 데이터 가져오기
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  initSheet(ss, SHEETS.RAW_MATERIALS, ['날짜', '품목명', '구분', '수량', '로트번호', '담당자', '내역']);
  initSheet(ss, SHEETS.PRODUCTION, ['날짜', '제품명', '생산량', '작업자', '비고']);
  initSheet(ss, SHEETS.SEMI_FINISHED, ['날짜', '반제품명', '입고', '사용', '재고']);
  initSheet(ss, SHEETS.PACKAGING, ['날짜', '제품명', '포장수량', '검수자', '상태']);
  
  // 헤더에 '상태' 추가
  initSheet(ss, SHEETS.ITEMS, ['구분', '품목명', '단위', '상태']); 
  
  // [NEW] 로트번호 관리 시트 초기화
  let lotSheet = ss.getSheetByName(SHEETS.LOT_SETTINGS);
  if (!lotSheet) {
    lotSheet = ss.insertSheet(SHEETS.LOT_SETTINGS);
    lotSheet.appendRow(['카테고리', '시작번호', '종료번호', '마지막사용번호']);
    lotSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#E5E7EB');
    // 기본값 세팅
    lotSheet.appendRow(['원재료', 100, 199, 99]);
    lotSheet.appendRow(['반제품', 200, 299, 199]);
    lotSheet.appendRow(['완제품', 300, 399, 299]);
    lotSheet.appendRow(['포장재', 400, 499, 399]);
  }

  // 로트 설정 데이터 읽기
  const lotDataRaw = lotSheet.getRange(2, 1, lotSheet.getLastRow() - 1, 4).getValues();
  const lotSettings = lotDataRaw.map(row => ({
    category: row[0],
    start: row[1],
    end: row[2],
    last: row[3]
  }));

  // 회사명 설정 읽기
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
    const lastCol = itemSheet.getLastColumn();
    const numCols = lastCol < 4 ? 3 : 4;
    const data = itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, numCols).getValues();
    
    rawMaterials = data.filter(row => row[0] === '원재료').map(row => ({
      name: row[1],
      unit: row[2],
      status: (row[3] || '사용').toString() 
    }));
  }

  return {
    companyName: companyName,
    currentUser: Session.getActiveUser().getEmail(),
    rawMaterials: rawMaterials,
    lotSettings: lotSettings
  };
}

/**
 * [NEW] 로트 설정 업데이트 (수정)
 */
function updateLotConfig(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LOT_SETTINGS);
  
  const category = formObj.lotCategory;
  const newStart = parseInt(formObj.lotStart);
  const newEnd = parseInt(formObj.lotEnd);

  if (isNaN(newStart) || isNaN(newEnd)) {
    throw new Error("시작번호와 종료번호는 숫자여야 합니다.");
  }
  if (newStart >= newEnd) {
    throw new Error("종료번호는 시작번호보다 커야 합니다.");
  }

  // [중복 검사] 수정하려는 카테고리(category)는 제외하고 검사
  checkLotOverlap(sheet, newStart, newEnd, category);

  const data = sheet.getDataRange().getValues();
  let found = false;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category) {
      // 시작번호(Col 2), 종료번호(Col 3) 업데이트
      sheet.getRange(i + 1, 2, 1, 2).setValues([[newStart, newEnd]]);
      found = true;
      break;
    }
  }

  if (!found) throw new Error("해당 카테고리를 찾을 수 없습니다.");
  
  return "로트번호 범위가 수정되었습니다.";
}

/**
 * [NEW] 로트 설정 추가 (등록)
 */
function addLotConfig(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LOT_SETTINGS);
  
  // 등록 시에는 입력 필드(lotCategoryInput)의 값을 사용
  const category = formObj.lotCategoryInput; 
  const start = parseInt(formObj.lotStart);
  const end = parseInt(formObj.lotEnd);

  if (!category || category.trim() === "") throw new Error("카테고리명을 입력해주세요.");
  if (isNaN(start) || isNaN(end)) throw new Error("숫자를 입력해주세요.");
  if (start >= end) throw new Error("종료 번호가 시작 번호보다 커야 합니다.");

  // [중복 검사] 신규 등록이므로 제외할 카테고리 없음
  checkLotOverlap(sheet, start, end);

  // 중복 확인 (이름)
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category) {
      throw new Error("이미 존재하는 카테고리입니다.");
    }
  }

  // 데이터 추가 (마지막 사용 번호는 start - 1 로 초기화)
  sheet.appendRow([category, start, end, start - 1]);

  return "새로운 로트 식별자가 등록되었습니다.";
}

/**
 * [Validation] 로트 범위 중복 검사 함수
 * @param {Sheet} sheet - 로트 설정 시트 객체
 * @param {number} newStart - 입력한 시작 번호
 * @param {number} newEnd - 입력한 종료 번호
 * @param {string|null} excludeCategory - 수정 시 본인 카테고리는 검사에서 제외 (등록 시 null)
 */
function checkLotOverlap(sheet, newStart, newEnd, excludeCategory = null) {
  const data = sheet.getDataRange().getValues();
  
  // 헤더(Row 0) 제외하고 반복
  for (let i = 1; i < data.length; i++) {
    const rowCategory = data[i][0];
    const rowStart = parseInt(data[i][1]);
    const rowEnd = parseInt(data[i][2]);
    
    // 수정 모드일 경우, 자기 자신과는 겹쳐도 상관없으므로 스킵
    if (excludeCategory && rowCategory === excludeCategory) {
      continue;
    }

    // 범위 겹침 알고리즘: (StartA <= EndB) && (StartB <= EndA)
    // 하나라도 겹치면 중복으로 판단
    if (newStart <= rowEnd && rowStart <= newEnd) {
      throw new Error(
        `범위가 중복됩니다!\n\n` +
        `[${rowCategory}]의 범위: ${rowStart} ~ ${rowEnd}\n` +
        `입력한 범위: ${newStart} ~ ${newEnd}\n\n` +
        `다른 범위를 설정해주세요.`
      );
    }
  }
}

/**
 * [NEW] 로트 설정 삭제
 */
function deleteLotConfig(category) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.LOT_SETTINGS);
  const data = sheet.getDataRange().getValues();
  
  let deleteIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === category) {
      deleteIndex = i + 1; // 행 번호 (1-based)
      break;
    }
  }
  
  if (deleteIndex !== -1) {
    sheet.deleteRow(deleteIndex);
    return "로트 식별자가 삭제되었습니다.";
  } else {
    throw new Error("삭제할 대상을 찾을 수 없습니다.");
  }
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

  sheet.appendRow([
    formObj.category,
    formObj.itemName,
    formObj.unit,
    formObj.status || '사용' // 기본값 사용
  ]);
  
  return "품목이 등록되었습니다.";
}

/**
 * [품목관리] 품목 수정 (연관 데이터 일괄 업데이트)
 */
function editItem(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName(SHEETS.ITEMS);
  const rawSheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);

  const originalName = formObj.originalItemName;
  const newName = formObj.itemName;
  const newUnit = formObj.unit;
  const newStatus = formObj.status;

  // 1. 품목 시트 업데이트
  const itemData = itemSheet.getDataRange().getValues();
  let itemFound = false;

  for (let i = 1; i < itemData.length; i++) {
    // 카테고리가 '원재료'이고 이름이 기존 이름과 같은 경우
    if (itemData[i][0] === '원재료' && itemData[i][1] === originalName) {
      // 이름, 단위, 상태 업데이트
      itemSheet.getRange(i + 1, 2, 1, 3).setValues([[newName, newUnit, newStatus]]);
      itemFound = true;
      break;
    }
  }

  if (!itemFound) {
    throw new Error("수정할 품목을 찾을 수 없습니다.");
  }

  // 2. 이름이 변경되었다면, '원재료수불부'에 있는 모든 내역도 변경해야 함
  if (originalName !== newName && rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getRange(2, 2, rawSheet.getLastRow() - 1, 1).getValues(); // 품목명 열(B열)만 가져옴
    const newRawValues = rawData.map(row => {
      // 기존 이름과 같으면 새 이름으로 변경, 아니면 유지
      return row[0] === originalName ? [newName] : [row[0]];
    });
    
    // 변경된 데이터 일괄 덮어쓰기 (속도 최적화)
    rawSheet.getRange(2, 2, rawSheet.getLastRow() - 1, 1).setValues(newRawValues);
  }

  return "품목 정보가 수정되었습니다.\n(관련된 입출고 내역의 품목명도 함께 변경되었습니다)";
}

/**
 * [품목관리] 품목 삭제 (사용 여부 체크)
 */
function deleteItem(itemName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const itemSheet = ss.getSheetByName(SHEETS.ITEMS);
  const rawSheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);

  // 1. 사용 여부 체크 (원재료수불부)
  if (rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getRange(2, 2, rawSheet.getLastRow() - 1, 1).getValues();
    // flat()으로 1차원 배열 변환 후 포함 여부 확인
    const isUsed = rawData.flat().includes(itemName);
    
    if (isUsed) {
      throw new Error(`'${itemName}'은(는) 이미 원재료 입출고 내역에 사용되고 있습니다.\n\n내역을 유지하려면 '미사용' 상태로 변경해주세요.`);
    }
  }

  // 2. 품목 시트에서 삭제
  const itemData = itemSheet.getDataRange().getValues();
  let deleteIndex = -1;

  for (let i = 1; i < itemData.length; i++) {
    if (itemData[i][0] === '원재료' && itemData[i][1] === itemName) {
      deleteIndex = i + 1; // 행 번호 (1-based)
      break;
    }
  }

  if (deleteIndex !== -1) {
    itemSheet.deleteRow(deleteIndex);
    return "품목이 삭제되었습니다.";
  } else {
    throw new Error("삭제할 품목을 찾을 수 없습니다.");
  }
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