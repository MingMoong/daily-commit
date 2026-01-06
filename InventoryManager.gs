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
  // CEO님 참고: 헤더 이름에 띄어쓰기가 있어도 전혀 문제 없습니다! (예: '품목 명', '입고 수량')
  initSheet(ss, SHEETS.RAW_MATERIALS, ['날짜', '품목명', '구분', '수량', '담당자', '비고']);
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
 * [원재료] 데이터 불러오기
 * 최신순으로 정렬하여 반환합니다.
 */
function getMaterials() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  // 데이터가 없으면 빈 배열 반환 (헤더 제외)
  if (sheet.getLastRow() < 2) return [];
  
  // 전체 데이터 가져오기 (2행부터 끝까지)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  // 배열을 객체(JSON) 형태로 변환하여 사용하기 편하게 가공
  // 최신 데이터가 위로 오도록 reverse() 사용
  return data.map(row => ({
    date: formatDate(row[0]), // 날짜 포맷팅
    item: row[1],
    type: row[2], // 입고 or 출고
    qty: row[3],
    manager: row[4],
    memo: row[5]
  })).reverse(); 
}

/**
 * [원재료] 데이터 저장하기
 * 앱에서 받은 데이터를 시트의 마지막 줄에 추가합니다.
 */
function addMaterial(formObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEETS.RAW_MATERIALS);
  
  // 입력된 값들 (순서: 날짜, 품목명, 구분, 수량, 담당자, 비고)
  const rowData = [
    formObj.date,
    formObj.item,
    formObj.type,
    formObj.qty,
    Session.getActiveUser().getEmail(), // 현재 로그인한 사용자 이메일 자동 기록
    formObj.memo
  ];
  
  sheet.appendRow(rowData);
  return "성공적으로 저장되었습니다.";
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
    // 날짜 열(A열) 서식 설정 (YYYY-MM-DD)
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  }
  return sheet;
}

/**
 * 날짜 데이터를 YYYY-MM-DD 문자열로 변환하는 유틸리티
 */
function formatDate(date) {
  if (!date) return "";
  if (typeof date === 'string') return date; // 이미 문자열이면 그대로
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}
