/**
 * Woondooran ERP - Server Side Logic
 * 
 * 작성자: Vibe Coding Expert
 * 설명: ERP 시스템의 백엔드 로직입니다. 필요한 시트를 관리하고 데이터를 제공합니다.
 */

// 전역 상수: 사용할 시트 이름 정의
const SHEETS = {
  RAW_MATERIALS: '원재료수불부',
  PRODUCTION: '생산일지',
  SEMI_FINISHED: '반제품수불부',
  PACKAGING: '포장일지',
  SETTINGS: '설정'
};

/**
 * [필수] 웹 앱 접속 시 실행되는 함수 (GET 요청 처리)
 */
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
      .setTitle('Woondooran ERP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no') // 모바일 최적화
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 초기 앱 실행 시 데이터 가져오기
 * 1. 필요한 시트가 없으면 생성
 * 2. 회사명 등 설정 정보 반환
 */
function getInitialData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. 각 시트가 존재하는지 확인하고 없으면 생성
  // [수정] 원재료수불부 헤더 변경: 로트번호 추가, 비고 -> 내역
  initSheet(ss, SHEETS.RAW_MATERIALS, ['날짜', '품목명', '구분', '수량', '로트번호', '담당자', '내역']);
  initSheet(ss, SHEETS.PRODUCTION, ['날짜', '제품명', '생산량', '작업자', '비고']);
  initSheet(ss, SHEETS.SEMI_FINISHED, ['날짜', '반제품명', '입고', '사용', '재고']);
  initSheet(ss, SHEETS.PACKAGING, ['날짜', '제품명', '포장수량', '검수자', '상태']);
  
  // 설정 시트 확인 및 기본 회사명 설정
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

  return {
    companyName: companyName,
    currentUser: Session.getActiveUser().getEmail()
  };
}

/**
 * [원재료] 데이터 불러오기 (로트번호, 내역 추가)
 */
function getMaterials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  // 데이터가 없으면 빈 값 반환
  if (sheet.getLastRow() < 2) {
    return { history: [], summary: [] };
  }
  
  // 전체 데이터 가져오기
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  // 1. 히스토리 데이터 포맷팅 (컬럼 인덱스 변경됨)
  // 0:날짜, 1:품목, 2:구분, 3:수량, 4:로트번호, 5:담당자, 6:내역
  const history = values.map(row => ({
    date: formatDate(row[0]),
    item: row[1],
    type: row[2], 
    qty: Number(row[3]) || 0,
    lot: row[4] || '', // 로트번호 추가
    manager: row[5],
    desc: row[6]       // 비고 -> 내역
  })).reverse(); // 최신순 정렬

  // 2. 재고 합계 계산 (로직 동일)
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

  // Map을 배열로 변환
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
 * [원재료] 데이터 저장하기 (로트번호, 내역 반영)
 */
function addMaterial(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  const rowData = [
    formObj.date,
    formObj.item,
    formObj.type,
    formObj.qty,
    formObj.lot, // 로트번호 추가
    Session.getActiveUser().getEmail(),
    formObj.desc // 내역(구 비고)
  ];
  
  sheet.appendRow(rowData);
  return "원재료가 저장되었습니다.";
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

/**
 * 시트 초기화 헬퍼 함수
 */
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

/**
 * 날짜 데이터를 YYYY-MM-DD 문자열로 변환하는 유틸리티
 */
function formatDate(date) {
  if (!date) return "";
  if (typeof date === 'string') return date;
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}