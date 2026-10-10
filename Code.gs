/**
 * 보강 알리미 (Google Apps Script Backend)
 */

function doGet(e) {
  var htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
  htmlOutput
    .setTitle('보강 알리미 | 온라인 교무실')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
  return htmlOutput;
}

// 기본 연결 구글 스프레드시트 ID
var DEFAULT_SPREADSHEET_ID = '1LsST_QqLkRIDQNvbeXw5EeCncJCEEOqAsC2duSngOcM';

/**
 * 데이터베이스 역할을 하는 구글 스프레드시트를 가져오거나 없으면 자동 생성합니다.
 */
function getDbSheet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ssId = scriptProperties.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  var ss;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) {
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('보강알리미_DB');
    scriptProperties.setProperty('SPREADSHEET_ID', ss.getId());
  }

  var sheet = ss.getSheetByName('보강내역');
  if (!sheet) {
    sheet = ss.insertSheet('보강내역');
    // 기본 시트 삭제 (Sheet1)
    var defaultSheet = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      try { ss.deleteSheet(defaultSheet); } catch (err) {}
    }
  }

  // 헤더 생성 또는 기존 헤더 갱신 (5번째 '보강학급' 열 추가 마이그레이션 포함, 총 13개 열)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '날짜', '교시', '교실', '보강학급', '보강교과', '원교사', '보강교사', '사유', '등록시각', '확인여부', '긴급여부', '삭제여부']);
    sheet.getRange(1, 1, 1, 13).setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  } else {
    // 기존 헤더가 '학급'인 경우 '교실'로 자동 갱신
    var col4Header = sheet.getRange(1, 4).getValue();
    if (col4Header === '학급') {
      sheet.getRange(1, 4).setValue('교실');
    }
    // 기존 DB 마이그레이션: 5번째 열이 '원교사'인 경우 '보강교과' 열 삽입
    var col5Header = String(sheet.getRange(1, 5).getValue() || '').trim();
    if (col5Header === '원교사') {
      sheet.insertColumnBefore(5);
      sheet.getRange(1, 5).setValue('보강교과').setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
      col5Header = '보강교과';
    }
    // 기존 DB 마이그레이션: 5번째 열이 '보강교과'인 경우 '보강학급' 열 자동 삽입
    if (col5Header === '보강교과') {
      sheet.insertColumnBefore(5);
      sheet.getRange(1, 5).setValue('보강학급').setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    }
    // 11번째 열 확인여부 헤더 추가 확인
    if (sheet.getLastColumn() < 11 || sheet.getRange(1, 11).getValue() === '') {
      sheet.getRange(1, 11).setValue('확인여부').setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    }
    // 12번째 열 긴급여부 헤더 추가 확인
    if (sheet.getLastColumn() < 12 || sheet.getRange(1, 12).getValue() === '') {
      sheet.getRange(1, 12).setValue('긴급여부').setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    }
    // 13번째 열 삭제여부 헤더 추가 확인
    if (sheet.getLastColumn() < 13 || sheet.getRange(1, 13).getValue() === '') {
      sheet.getRange(1, 13).setValue('삭제여부').setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    }
  }

  return sheet;
}

/**
 * 보강내역 행의 삭제여부(12열) 값에 따라 시트 행 전체(1~12열)에 취소선(line-through)을 적용하거나 해제(none)합니다.
 * @param {Sheet} sheet 구글 시트 개체
 * @param {number} rowNum 1-indexed 행 번호
 * @param {boolean} isDeleted 삭제 여부
 */
function applyRowStrikethrough(sheet, rowNum, isDeleted) {
  try {
    sheet.getRange(rowNum, 1, 1, 13).setFontLine(isDeleted ? 'line-through' : 'none');
  } catch (e) {
    Logger.log('Error in applyRowStrikethrough for row ' + rowNum + ': ' + e.toString());
  }
}

/**
 * 구글 시트에 직접 수동 입력 시(날짜, 교시, 교실, 보강교과, 원교사, 보강교사, 사유)
 * 동적으로 ID, 등록시각, 확인여부, 긴급여부, 삭제여부 기본값을 자동으로 채워주는 트리거 함수
 */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    if (sheet.getName() !== '보강내역') return;

    var startRow = e.range.getRow();
    var numRows = e.range.getNumRows();

    // 헤더 행 제외
    if (startRow <= 1) return;

    for (var r = startRow; r < startRow + numRows; r++) {
      var rowValues = sheet.getRange(r, 1, 1, 13).getValues()[0];
      var hasContent = rowValues[1] || rowValues[2] || rowValues[3] || rowValues[4] || rowValues[5] || rowValues[6] || rowValues[7] || rowValues[8];
      if (!hasContent) continue;

      // 1열 ID 자동 입력
      if (!rowValues[0]) {
        sheet.getRange(r, 1).setValue('SUB-MANUAL-' + new Date().getTime() + '-' + r);
      }
      // 10열 등록시각 자동 입력
      if (!rowValues[9]) {
        sheet.getRange(r, 10).setValue(new Date().toISOString());
      }
      // 11열 확인여부 기본값 (true)
      if (rowValues[10] === undefined || rowValues[10] === '') {
        sheet.getRange(r, 11).setValue(true);
      }
      // 12열 긴급여부 기본값 (false)
      if (rowValues[11] === undefined || rowValues[11] === '') {
        sheet.getRange(r, 12).setValue(false);
      }
      // 13열 삭제여부 기본값 (false)
      if (rowValues[12] === undefined || rowValues[12] === '') {
        sheet.getRange(r, 13).setValue(false);
        rowValues[12] = false;
      }

      // 삭제여부 (13열) 값에 따른 취소선 적용/해제
      var isDeleted = (rowValues[12] === true || String(rowValues[12]).toLowerCase() === 'true' || String(rowValues[12]) === 'y');
      applyRowStrikethrough(sheet, r, isDeleted);
    }
  } catch (err) {
    Logger.log('Error in onEdit: ' + err.toString());
  }
}

/**
 * 날짜 객체 또는 문자열을 YYYY-MM-DD 포맷으로 변환하는 헬퍼 함수
 */
function formatDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone() || 'GMT+9', 'yyyy-MM-dd');
  }
  var str = String(val).trim();
  if (str.indexOf('GMT') !== -1 || str.indexOf('한국 표준시') !== -1) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone() || 'GMT+9', 'yyyy-MM-dd');
    }
  }
  var match = str.match(/^(\d{4})[-.\/\s]+(\d{1,2})[-.\/\s]+(\d{1,2})/);
  if (match) {
    var y = match[1];
    var m = match[2].length === 1 ? '0' + match[2] : match[2];
    var d = match[3].length === 1 ? '0' + match[3] : match[3];
    return y + '-' + m + '-' + d;
  }
  var match2 = str.match(/^(\d{4})[-.\/]?(\d{2})[-.\/]?(\d{2})/);
  if (match2) {
    return match2[1] + '-' + match2[2] + '-' + match2[3];
  }
  return str;
}

/**
 * 보강 학급 값을 '1학년 1반' 포맷으로 정규화하는 헬퍼 함수
 * Google Sheets에서 '1-1'을 날짜(Date)로 자동 변환하는 현상을 포함하여 모든 입력 포맷을 일관되게 처리합니다.
 */
function formatSubClass(val) {
  if (!val) return '-';
  if (val instanceof Date) {
    var g = val.getMonth() + 1;
    var c = val.getDate();
    return g + '학년 ' + c + '반';
  }
  var str = String(val).trim();
  if (str === '' || str === '-') return '-';

  if (str.indexOf('GMT') !== -1 || str.indexOf('한국 표준시') !== -1) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      var g = d.getMonth() + 1;
      var c = d.getDate();
      return g + '학년 ' + c + '반';
    }
  }

  var match = str.match(/^(\d{1,2})[-.\s\/]+(\d{1,2})$/);
  if (match) {
    return match[1] + '학년 ' + match[2] + '반';
  }

  return str;
}

/**
 * 보강 내역을 조회합니다. (단일 날짜 또는 시작일~종료일 기간 검색 지원)
 * 수동 작성된 행(ID 미부여 행) 자동 보정 및 세팅 지원
 * @param {string} startDate 시작 날짜 (YYYY-MM-DD 또는 'ALL')
 * @param {string} endDate 종료 날짜 (YYYY-MM-DD, 옵션)
 */
function getSubstitutionRecords(startDate, endDate) {
  try {
    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var records = [];
    var needsFlush = false;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var hasData = row[1] || row[2] || row[3] || row[4] || row[5] || row[6] || row[7] || row[8];
      if (!hasData) continue; // 완전히 비어있는 행 스킵

      // ID가 없는 수동 입력 행 자동 보정
      var rowId = row[0] ? String(row[0]).trim() : '';
      if (!rowId) {
        rowId = 'SUB-MANUAL-' + new Date().getTime() + '-' + (i + 1);
        sheet.getRange(i + 1, 1).setValue(rowId);
        row[0] = rowId;
        needsFlush = true;
      }

      // 등록시각 (10열) 보정
      if (!row[9]) {
        var nowIso = new Date().toISOString();
        sheet.getRange(i + 1, 10).setValue(nowIso);
        row[9] = nowIso;
        needsFlush = true;
      }
      // 확인여부 (11열) 기본값 보정 (true)
      if (row[10] === undefined || row[10] === '') {
        sheet.getRange(i + 1, 11).setValue(true);
        row[10] = true;
        needsFlush = true;
      }
      // 긴급여부 (12열) 기본값 보정
      if (row[11] === undefined || row[11] === '') {
        sheet.getRange(i + 1, 12).setValue(false);
        row[11] = false;
        needsFlush = true;
      }
      // 삭제여부 (13열) 기본값 보정
      if (row[12] === undefined || row[12] === '') {
        sheet.getRange(i + 1, 13).setValue(false);
        row[12] = false;
        needsFlush = true;
      }

      var isDeleted = (row[12] === true || String(row[12]).toLowerCase() === 'true' || String(row[12]) === 'y');
      
      // 시트 행 취소선 적용/해제 동기화
      applyRowStrikethrough(sheet, i + 1, isDeleted);

      var rowDate = formatDateString(row[1]);

      // 날짜 필터링 (전체, 단일 날짜, 또는 기간 검색)
      if (startDate && startDate !== 'ALL') {
        if (endDate && endDate.trim() !== '') {
          if (rowDate < startDate || rowDate > endDate) continue;
        } else {
          if (rowDate !== startDate) continue;
        }
      }

      // 삭제 처리된 행은 웹 화면 조회에서 제외 (Soft Delete)
      if (isDeleted) continue;

      var isConf = (row[10] === true || String(row[10]).toLowerCase() === 'true' || String(row[10]) === '확인완료');
      var isUrgent = (row[11] === true || String(row[11]).toLowerCase() === 'true' || String(row[11]) === '긴급');

      records.push({
        id: String(row[0]),
        date: rowDate,
        period: String(row[2] || ''),
        className: String(row[3] || ''),
        subClass: formatSubClass(row[4]),
        subject: String(row[5] || ''),
        originalTeacher: String(row[6] || ''),
        substituteTeacher: String(row[7] || ''),
        reason: String(row[8] || ''),
        timestamp: row[9] ? String(row[9]) : '',
        confirmed: isConf,
        urgent: isUrgent
      });
    }

    if (needsFlush) {
      SpreadsheetApp.flush();
    }

    // 날짜 desc, 교시 asc 순으로 정렬
    records.sort(function(a, b) {
      if (a.date !== b.date) {
        return a.date > b.date ? -1 : 1;
      }
      var pA = parseInt(a.period) || 0;
      var pB = parseInt(b.period) || 0;
      return pA - pB;
    });

    return records;
  } catch (err) {
    Logger.log('Error in getSubstitutionRecords: ' + err.toString());
    throw new Error('보강 내역을 불러오는데 실패했습니다: ' + err.message);
  }
}

/**
 * 신규 보강 내역을 저장합니다. (날짜, 교시, 교실, 보강교과, 보강교사 필수)
 */
function addSubstitutionRecord(record) {
  try {
    if (!record.date || !record.period || !record.className || !record.subject || !record.substituteTeacher) {
      throw new Error('필수 입력 항목이 누락되었습니다.');
    }

    var sheet = getDbSheet();
    var newId = record.id || ('SUB-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000));
    var nowIso = new Date().toISOString();
    var formattedDate = formatDateString(record.date);
    var isConf = record.confirmed !== undefined ? (record.confirmed ? true : false) : true;
    var isUrgent = record.urgent ? true : false;

    sheet.appendRow([
      newId,
      formattedDate,
      record.period,
      record.className,
      formatSubClass(record.subClass),
      record.subject,
      record.originalTeacher || '',
      record.substituteTeacher,
      record.reason || '',
      nowIso,
      isConf,
      isUrgent,
      false // 삭제여부 (기본 false)
    ]);

    var lastRow = sheet.getLastRow();
    applyRowStrikethrough(sheet, lastRow, false);

    SpreadsheetApp.flush(); // 저장 즉시 적용

    return {
      success: true,
      id: newId,
      message: '보강 내역이 성공적으로 등록되었습니다.'
    };
  } catch (err) {
    Logger.log('Error in addSubstitutionRecord: ' + err.toString());
    return {
      success: false,
      message: err.message || '저장 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 기존 보강 내역을 수정합니다. (날짜, 교시, 교실, 보강교과, 보강교사 필수)
 */
function updateSubstitutionRecord(record) {
  try {
    if (!record.id || !record.date || !record.period || !record.className || !record.subject || !record.substituteTeacher) {
      throw new Error('필수 수정 정보가 누락되었습니다.');
    }

    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();
    var formattedDate = formatDateString(record.date);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(record.id)) {
        var rowNum = i + 1;
        var existingConf = (data[i][10] === true || String(data[i][10]).toLowerCase() === 'true');
        var isConf = record.confirmed !== undefined ? (record.confirmed ? true : false) : existingConf;

        var existingUrgent = (data[i][11] === true || String(data[i][11]).toLowerCase() === 'true');
        var isUrgent = record.urgent !== undefined ? (record.urgent ? true : false) : existingUrgent;

        var existingDeleted = (data[i][12] === true || String(data[i][12]).toLowerCase() === 'true');

        sheet.getRange(rowNum, 1, 1, 13).setValues([[
          String(record.id),
          formattedDate,
          record.period,
          record.className,
          formatSubClass(record.subClass),
          record.subject,
          record.originalTeacher || '',
          record.substituteTeacher,
          record.reason || '',
          new Date().toISOString(),
          isConf,
          isUrgent,
          existingDeleted
        ]]);

        applyRowStrikethrough(sheet, rowNum, existingDeleted);

        SpreadsheetApp.flush(); // 수정 즉시 적용

        return {
          success: true,
          message: '보강 내역이 성공적으로 수정되었습니다.'
        };
      }
    }
    return { success: false, message: '수정할 보강 내역을 찾을 수 없습니다.' };
  } catch (err) {
    Logger.log('Error in updateSubstitutionRecord: ' + err.toString());
    return {
      success: false,
      message: err.message || '수정 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 보강 교사 확인 상태 토글 저장 API
 */
function toggleSubstituteConfirm(id, confirmed) {
  try {
    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, 11).setValue(confirmed ? true : false);
        SpreadsheetApp.flush();
        return { success: true, confirmed: confirmed };
      }
    }
    return { success: false, message: '해당 내역을 찾을 수 없습니다.' };
  } catch (err) {
    Logger.log('Error in toggleSubstituteConfirm: ' + err.toString());
    return { success: false, message: err.message };
  }
}

/**
 * 보강 내역을 삭제 처리(Soft Delete)합니다.
 * - 실제 구글 시트의 행을 deleteRow 하지 않고 12번째 열 '삭제여부'를 true로 설정하여 웹 화면에서만 제외합니다.
 */
function deleteSubstitutionRecord(id) {
  try {
    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, 13).setValue(true);
        applyRowStrikethrough(sheet, rowNum, true);
        SpreadsheetApp.flush(); // 삭제 즉시 적용
        return { success: true, message: '보강 내역이 삭제되었습니다. (시트 데이터는 영구 보존됩니다)' };
      }
    }
    return { success: false, message: '해당 보강 내역을 찾을 수 없습니다.' };
  } catch (err) {
    Logger.log('Error in deleteSubstitutionRecord: ' + err.toString());
    return { success: false, message: err.message };
  }
}

/**
 * '태그관리' 시트에서 보강 교과 및 보강 유발 사유 태그 목록을 조회합니다.
 * (구글 시트 구조: 구분[SUBJECT/REASON] | 태그명 | 등록시각)
 */
function getTagsFromSheet() {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var ssId = scriptProperties.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
    var ss;
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    var sheet = ss.getSheetByName('태그관리');
    var nowIso = new Date().toISOString();
    var defaultSubjects = ['국어', '수학', '영어', '사회', '과학', '체육', '음악', '미술', '정보', '세계 문화와 영어A'];
    var defaultReasons = ['출장', '연가', '병가', '공가', '특별휴가', '조퇴', '외출', '지참'];

    if (!sheet) {
      sheet = ss.insertSheet('태그관리');
      sheet.appendRow(['구분', '태그명', '등록시각']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
      for (var s = 0; s < defaultSubjects.length; s++) {
        sheet.appendRow(['SUBJECT', defaultSubjects[s], nowIso]);
      }
      for (var r = 0; r < defaultReasons.length; r++) {
        sheet.appendRow(['REASON', defaultReasons[r], nowIso]);
      }
      SpreadsheetApp.flush();
    }
    
    var lastRow = sheet.getLastRow();
    var subjectTags = [];
    var reasonTags = [];
    
    if (lastRow > 1) {
      var data = sheet.getDataRange().getValues();
      var col1Header = String(data[0][0] || '').trim();
      var col2Header = String(data[0][1] || '').trim();
      
      for (var i = 1; i < data.length; i++) {
        var type = String(data[i][0] || '').trim().toUpperCase();
        var val = String(data[i][1] || '').trim();
        
        if (!type && !val) continue;
        
        if ((type === 'SUBJECT' || type.indexOf('교과') !== -1 || type.indexOf('SUBJECT') !== -1) && val) {
          if (subjectTags.indexOf(val) === -1) subjectTags.push(val);
        } else if ((type === 'REASON' || type.indexOf('사유') !== -1 || type.indexOf('REASON') !== -1) && val) {
          if (reasonTags.indexOf(val) === -1) reasonTags.push(val);
        } else if (col1Header.indexOf('보강교과') !== -1 || col2Header.indexOf('보강사유') !== -1) {
          var subj = String(data[i][0] || '').trim();
          var reas = String(data[i][1] || '').trim();
          if (subj && subjectTags.indexOf(subj) === -1) subjectTags.push(subj);
          if (reas && reasonTags.indexOf(reas) === -1) reasonTags.push(reas);
        }
      }
    }
    
    if (subjectTags.length === 0) {
      subjectTags = defaultSubjects;
    }
    if (reasonTags.length === 0) {
      reasonTags = defaultReasons;
    }
    
    return {
      success: true,
      subjectTags: subjectTags,
      reasonTags: reasonTags
    };
  } catch (e) {
    Logger.log('Error in getTagsFromSheet: ' + e.toString());
    return {
      success: false,
      error: e.toString(),
      subjectTags: ['국어', '수학', '영어', '사회', '과학', '체육', '음악', '미술', '정보', '세계 문화와 영어A'],
      reasonTags: ['출장', '연가', '병가', '공가', '특별휴가', '조퇴', '외출', '지참']
    };
  }
}

/**
 * '태그관리' 시트에 보강 교과 및 보강 유발 사유 태그 목록을 저장합니다.
 * (구조: 구분[SUBJECT/REASON] | 태그명 | 등록시각)
 */
function saveTagsToSheet(subjectTags, reasonTags) {
  try {
    var scriptProperties = PropertiesService.getScriptProperties();
    var ssId = scriptProperties.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
    var ss;
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    
    var sheet = ss.getSheetByName('태그관리');
    if (!sheet) {
      sheet = ss.insertSheet('태그관리');
    }
    sheet.clearContents();
    
    sheet.getRange(1, 1, 1, 3).setValues([['구분', '태그명', '등록시각']]).setFontWeight('bold').setBackground('#4f46e5').setFontColor('#ffffff');
    
    var sTags = Array.isArray(subjectTags) ? subjectTags : [];
    var rTags = Array.isArray(reasonTags) ? reasonTags : [];
    var nowIso = new Date().toISOString();
    var rows = [];
    
    for (var s = 0; s < sTags.length; s++) {
      var sVal = String(sTags[s] || '').trim();
      if (sVal) rows.push(['SUBJECT', sVal, nowIso]);
    }
    for (var r = 0; r < rTags.length; r++) {
      var rVal = String(rTags[r] || '').trim();
      if (rVal) rows.push(['REASON', rVal, nowIso]);
    }
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    }
    SpreadsheetApp.flush();
    return { success: true };
  } catch (e) {
    Logger.log('Error in saveTagsToSheet: ' + e.toString());
    return { success: false, error: e.toString() };
  }
}

// ============================================================================
// 📢 구글챗(Google Chat) 스페이스방 연동 및 보강 알림 시스템
// ============================================================================

/**
 * 구글 스프레드시트 내 '알림설정' 시트를 가져오거나 없으면 신규 생성 및 초기화합니다.
 */
function getNotificationSettingsSheet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ssId = scriptProperties.getProperty('SPREADSHEET_ID') || DEFAULT_SPREADSHEET_ID;
  var ss;
  try {
    ss = SpreadsheetApp.openById(ssId);
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  var sheet = ss.getSheetByName('알림설정');
  if (!sheet) {
    sheet = ss.insertSheet('알림설정');
    sheet.appendRow(['항목', '설정값', '설명 및 안내']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    sheet.setFrozenRows(1);

    var defaultRows = [
      ['구글챗_웹훅URL', '', '구글챗 교사 스페이스방 인커밍 웹훅 URL (https://chat.googleapis.com/... 형태)'],
      ['자동알림_사용여부', 'TRUE', '전날 낮 12시 익일 보강 자동 공지 실행 여부 (TRUE/FALSE)'],
      ['보강0건_알림생략', 'TRUE', '전날 12시 자동 공지 시 0건이면 발송 스킵 (TRUE/FALSE)'],
      ['금요일_월요일공지', 'TRUE', '금요일 발송 시 주말 건너뛰고 차주 월요일 공지 (TRUE/FALSE)'],
      ['웹앱_접속URL', '', '공지 메시지 하단에 표시될 보강 알리미 웹앱 URL']
    ];
    sheet.getRange(2, 1, defaultRows.length, 3).setValues(defaultRows);
    try {
      sheet.autoResizeColumns(1, 3);
    } catch (err) {}
    SpreadsheetApp.flush();
  }
  return sheet;
}

/**
 * '알림설정' 시트에서 현재 설정값들을 객체 형태로 반환합니다.
 */
function getNotificationSettings() {
  try {
    var sheet = getNotificationSettingsSheet();
    var data = sheet.getDataRange().getValues();
    var settings = {
      webhookUrl: '',
      autoNotifyEnabled: true,
      skipZero: true,
      fridayMonday: true,
      webAppUrl: ''
    };

    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      var val = String(data[i][1] || '').trim();

      if (key === '구글챗_웹훅URL') settings.webhookUrl = val;
      else if (key === '자동알림_사용여부') settings.autoNotifyEnabled = (val.toUpperCase() === 'TRUE' || val === '1');
      else if (key === '보강0건_알림생략') settings.skipZero = (val.toUpperCase() === 'TRUE' || val === '1');
      else if (key === '금요일_월요일공지') settings.fridayMonday = (val.toUpperCase() === 'TRUE' || val === '1');
      else if (key === '웹앱_접속URL') settings.webAppUrl = val;
    }
    return settings;
  } catch (err) {
    Logger.log('Error in getNotificationSettings: ' + err.toString());
    return {
      webhookUrl: '',
      autoNotifyEnabled: true,
      skipZero: true,
      fridayMonday: true,
      webAppUrl: ''
    };
  }
}

/**
 * '알림설정' 시트에 설정값을 저장하고, 자동 트리거 설정/해제를 동기화합니다.
 */
function saveNotificationSettings(newSettings) {
  try {
    var sheet = getNotificationSettingsSheet();
    var data = sheet.getDataRange().getValues();
    var map = {
      '구글챗_웹훅URL': newSettings.webhookUrl !== undefined ? String(newSettings.webhookUrl).trim() : null,
      '자동알림_사용여부': newSettings.autoNotifyEnabled !== undefined ? (newSettings.autoNotifyEnabled ? 'TRUE' : 'FALSE') : null,
      '보강0건_알림생략': newSettings.skipZero !== undefined ? (newSettings.skipZero ? 'TRUE' : 'FALSE') : null,
      '금요일_월요일공지': newSettings.fridayMonday !== undefined ? (newSettings.fridayMonday ? 'TRUE' : 'FALSE') : null,
      '웹앱_접속URL': newSettings.webAppUrl !== undefined ? String(newSettings.webAppUrl).trim() : null
    };

    for (var i = 1; i < data.length; i++) {
      var key = String(data[i][0] || '').trim();
      if (map[key] !== null && map[key] !== undefined) {
        sheet.getRange(i + 1, 2).setValue(map[key]);
      }
    }
    SpreadsheetApp.flush();

    // 자동 알림 트리거 설정 동기화 시도 (웹앱 권한 제한 환경에서도 시트 저장이 실패하지 않도록 안전 처리)
    try {
      if (newSettings.autoNotifyEnabled) {
        setupDaily12Trigger();
      } else {
        removeDaily12Trigger();
      }
    } catch (trigErr) {
      Logger.log('트리거 자동 동기화 건너뜀 (권한 필요): ' + trigErr.toString());
    }

    return { success: true, message: '알림 설정이 [알림설정] 시트에 성공적으로 저장되었습니다!' };
  } catch (err) {
    Logger.log('Error in saveNotificationSettings: ' + err.toString());
    return { success: false, message: '알림 설정 저장 실패: ' + err.message };
  }
}

/**
 * 텍스트의 시각적 너비를 계산하는 헬퍼 함수 (한글/특수문자 2칸, 영문/숫자 1칸)
 */
function getGasVisualWidth(str) {
  if (!str) return 0;
  var width = 0;
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    if (code >= 0xD800 && code <= 0xDBFF) {
      width += 2;
      i++;
      continue;
    }
    if (
      (code >= 0x1100 && code <= 0x11FF) ||
      (code >= 0x3130 && code <= 0x318F) ||
      (code >= 0xAC00 && code <= 0xD7A3) ||
      (code >= 0x2E80 && code <= 0x9FFF) ||
      (code >= 0xFF00 && code <= 0xFFEF) ||
      (code >= 0x2600 && code <= 0x27BF)
    ) {
      width += 2;
    } else {
      width += 1;
    }
  }
  return width;
}

function padRightVisualGas(str, targetWidth) {
  str = String(str || '');
  var currentWidth = getGasVisualWidth(str);
  if (currentWidth >= targetWidth) return str;
  return str + ' '.repeat(targetWidth - currentWidth);
}

function formatFormattedDateGas(dateStr) {
  if (!dateStr) return '';
  var parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  var d = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  var days = ['일', '월', '화', '수', '목', '금', '토'];
  return dateStr + ' (' + days[d.getDay()] + ')';
}

function formatPeriodTextGas(period) {
  if (!period) return '-교시';
  return String(period).indexOf('교시') !== -1 ? String(period) : (period + '교시');
}

function buildClassInfoTextGas(room, subClass) {
  if (room && subClass) return room + ' (' + subClass + ')';
  if (room) return room;
  if (subClass) return subClass;
  return '-';
}

/**
 * 모바일 및 PC 구글챗 화면에서 글자 쪼개짐과 줄바꿈 왜곡 없이
 * 최적의 가독성을 제공하는 2줄 카드형 개조식 보강 목록 텍스트 블록을 생성합니다.
 */
function buildRecordsTableBlockGas(records) {
  if (!records || records.length === 0) return '';

  var divider = '────────────────────────';
  var lines = [divider];

  records.forEach(function(r, idx) {
    var isUrgent = (r.urgent === true || String(r.urgent).toLowerCase() === 'true');
    var periodStr = formatPeriodTextGas(r.period);
    var classInfo = buildClassInfoTextGas(r.className, r.subClass);
    var subjectStr = r.subject || '-';
    var subTeacher = r.substituteTeacher ? ('[보강] ' + r.substituteTeacher) : '[보강] -';
    var origTeacher = r.originalTeacher ? ('[원교사] ' + r.originalTeacher) : '[원교사] -';

    // 1행: 핵심 수업 정보 (교시, 교실/학급, 교과)
    // 2행: 교사 인수인계 정보 (보강교사 ⬅️ 원교사)
    if (isUrgent) {
      lines.push('• 🚨 *[' + periodStr + ' 긴급]* ' + classInfo + ' | ' + subjectStr);
      lines.push('  👉 *' + subTeacher + '* ⬅️ ' + origTeacher);
    } else {
      lines.push('• *[' + periodStr + ']* ' + classInfo + ' | ' + subjectStr);
      lines.push('  👉 *' + subTeacher + '* ⬅️ ' + origTeacher);
    }

    if (idx < records.length - 1) {
      lines.push('');
    }
  });

  lines.push(divider);
  return lines.join('\n');
}

/**
 * [🚨 구글챗으로 긴급 공지] 발송용 메시지 생성:
 * 모바일과 PC 모두에서 줄바꿈 깨짐 없이 최상의 가독성으로 긴급 및 내일 보강을 결합합니다.
 */
function buildCombinedBroadcastMessage(urgentRecords, tomorrowRecords, tomorrowDateStr, webAppUrl) {
  var msgParts = [];

  // 1. 긴급 보강 섹션 (가시성 극대화 및 일시 명시)
  if (urgentRecords && urgentRecords.length > 0) {
    var urgentDates = [];
    urgentRecords.forEach(function(r) {
      if (r.date && urgentDates.indexOf(r.date) === -1) {
        urgentDates.push(r.date);
      }
    });
    var formattedUrgentDate = urgentDates.map(function(d) {
      return formatFormattedDateGas(d);
    }).join(' / ');
    if (!formattedUrgentDate && urgentRecords[0] && urgentRecords[0].date) {
      formattedUrgentDate = formatFormattedDateGas(urgentRecords[0].date);
    }

    msgParts.push('🚨🚨🚨 *[긴급 보강 알림]* 긴급 보강이 발생했습니다! 🚨🚨🚨');
    msgParts.push('⚠️ 해당 선생님께서는 변동된 일정을 최우선으로 즉시 확인해 주시기 바랍니다.\n');
    msgParts.push('🔥 *[🚨 긴급 보강 내역]* (총 ' + urgentRecords.length + '건)');
    if (formattedUrgentDate) {
      msgParts.push('📅 일시: ' + formattedUrgentDate);
    }
    msgParts.push(buildRecordsTableBlockGas(urgentRecords));
    msgParts.push('');
  } else {
    msgParts.push('ℹ️ [안내] 현재 등록된 긴급 보강 건은 없습니다.\n');
  }

  // 2. 기존 보강 섹션 ('[기존 보강 내역]'으로 명칭 통일)
  var formattedTomorrow = formatFormattedDateGas(tomorrowDateStr);
  msgParts.push('📢 *[기존 보강 내역]*');
  msgParts.push('📅 일시: ' + formattedTomorrow);

  if (tomorrowRecords && tomorrowRecords.length > 0) {
    msgParts.push(buildRecordsTableBlockGas(tomorrowRecords));
    msgParts.push('(기존 보강 총 ' + tomorrowRecords.length + '건)\n');
  } else {
    msgParts.push('────────────────────────');
    msgParts.push('배정된 기존 보강이 없습니다. 😊');
    msgParts.push('────────────────────────\n');
  }

  // 3. 하단 안내 및 웹앱 링크
  msgParts.push('👉 보강 교사께서는 웹앱에서 본인 확인 체크를 완료해 주세요.');
  if (webAppUrl) {
    msgParts.push('🌐 보강 알리미 바로가기: ' + webAppUrl);
  }

  return msgParts.join('\n');
}

/**
 * 구글챗 웹훅으로 메시지를 전송합니다.
 */
function sendGoogleChatMessage(textMessage, customWebhookUrl) {
  var settings = getNotificationSettings();
  var webhookUrl = customWebhookUrl || settings.webhookUrl;

  if (!webhookUrl || webhookUrl.indexOf('https://chat.googleapis.com/') !== 0) {
    throw new Error('올바른 구글챗 웹훅 URL이 설정되지 않았습니다. 스프레드시트의 [알림설정] 시트 또는 설정창에서 웹훅 URL을 입력해 주세요.');
  }

  var payload = {
    text: textMessage
  };

  var options = {
    method: 'post',
    contentType: 'application/json; charset=UTF-8',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(webhookUrl, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();

  if (responseCode !== 200) {
    throw new Error('구글챗 전송 실패 (상태 코드 ' + responseCode + '): ' + responseText);
  }

  return { success: true, responseText: responseText };
}

/**
 * 클라이언트 상단 툴바 [🚨 구글챗으로 즉시 공지]에서 호출하는 발송 API
 */
function sendInstantBroadcastFromClient(urgentRecords, tomorrowRecords, tomorrowDateStr) {
  try {
    var settings = getNotificationSettings();
    var webAppUrl = settings.webAppUrl;
    
    // 웹앱 URL이 비어있으면 현재 배포 URL 자동 추론
    if (!webAppUrl) {
      try {
        webAppUrl = ScriptApp.getService().getUrl();
      } catch (e) {
        webAppUrl = '';
      }
    }

    var message = buildCombinedBroadcastMessage(urgentRecords || [], tomorrowRecords || [], tomorrowDateStr, webAppUrl);
    sendGoogleChatMessage(message, settings.webhookUrl);

    return {
      success: true,
      message: '구글챗 교사 스페이스방에 긴급 및 내일 보강 공지가 성공적으로 발송되었습니다!'
    };
  } catch (err) {
    Logger.log('Error in sendInstantBroadcastFromClient: ' + err.toString());
    return {
      success: false,
      message: err.message || '구글챗 발송 중 오류가 발생했습니다.'
    };
  }
}

/**
 * 매일 낮 12:00 시간 기반 트리거로 실행되는 익일 보강 자동 공지 함수
 */
function sendScheduledTomorrowNotification() {
  try {
    var settings = getNotificationSettings();
    if (!settings.autoNotifyEnabled) {
      Logger.log('자동 알림이 비활성화되어 있어 실행을 건너뜁니다.');
      return { success: true, skipped: true, reason: 'disabled' };
    }

    if (!settings.webhookUrl) {
      Logger.log('구글챗 웹훅 URL이 등록되지 않았습니다.');
      return { success: false, reason: 'no_webhook' };
    }

    var now = new Date();
    var dayOfWeek = now.getDay(); // 0:일, 1:월, ..., 5:금, 6:토
    var daysToAdd = 1;

    // 금요일_월요일공지 설정 확인
    if (settings.fridayMonday) {
      if (dayOfWeek === 5) {
        daysToAdd = 3; // 금요일이면 월요일(+3)
      } else if (dayOfWeek === 6) {
        daysToAdd = 2; // 토요일이면 월요일(+2)
      }
    }

    var targetDateObj = new Date(now.getTime() + daysToAdd * 24 * 60 * 60 * 1000);
    var targetDateStr = formatDateString(targetDateObj);

    var records = getSubstitutionRecords(targetDateStr, targetDateStr);

    // 0건일 때 스킵 정책
    if ((!records || records.length === 0) && settings.skipZero) {
      Logger.log(targetDateStr + ' 보강 내역 0건으로 자동 공지를 생략합니다.');
      return { success: true, skipped: true, count: 0 };
    }

    var webAppUrl = settings.webAppUrl;
    if (!webAppUrl) {
      try { webAppUrl = ScriptApp.getService().getUrl(); } catch (e) { webAppUrl = ''; }
    }

    var formattedDate = formatFormattedDateGas(targetDateStr);
    var msgLines = [];
    msgLines.push('📢 *[보강 알리미] 내일 보강 현황 안내*');
    msgLines.push('📅 일시: ' + formattedDate);

    if (records && records.length > 0) {
      msgLines.push(buildRecordsTableBlockGas(records));
      msgLines.push('(내일 보강 총 ' + records.length + '건)\n');
    } else {
      msgLines.push('────────────────────────');
      msgLines.push('배정된 보강이 없습니다. 편안한 하루 되세요 😊');
      msgLines.push('────────────────────────\n');
    }

    msgLines.push('👉 보강 교사께서는 웹앱에서 본인 확인 체크를 완료해 주세요.');
    if (webAppUrl) {
      msgLines.push('🌐 보강 알리미 바로가기: ' + webAppUrl);
    }

    sendGoogleChatMessage(msgLines.join('\n'), settings.webhookUrl);
    Logger.log(targetDateStr + ' 보강 공지 완료 (총 ' + records.length + '건)');
    return { success: true, count: records.length };
  } catch (err) {
    Logger.log('Error in sendScheduledTomorrowNotification: ' + err.toString());
    return { success: false, error: err.message };
  }
}

/**
 * 매일 낮 12:00 시간 기반 트리거를 생성합니다. (중복 방지)
 */
function setupDaily12Trigger() {
  removeDaily12Trigger();
  ScriptApp.newTrigger('sendScheduledTomorrowNotification')
    .timeBased()
    .atHour(12)
    .everyDays(1)
    .create();
  Logger.log('매일 낮 12시 정기 알림 트리거가 등록되었습니다.');
}

/**
 * 기등록된 낮 12:00 정기 알림 트리거를 삭제합니다.
 */
function removeDaily12Trigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendScheduledTomorrowNotification') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  Logger.log('기존 낮 12시 정기 알림 트리거가 삭제되었습니다.');
}

/**
 * 현재 트리거 등록 상태를 반환합니다.
 */
function getTriggerStatus() {
  var triggers = ScriptApp.getProjectTriggers();
  var isScheduled = false;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendScheduledTomorrowNotification') {
      isScheduled = true;
      break;
    }
  }
  return { isScheduled: isScheduled };
}

/**
 * 구글챗 웹훅 연결 테스트 함수
 */
function testGoogleChatWebhook(customUrl) {
  try {
    var testMsg = '🔔 [홍익 보강 알리미] 구글챗 웹훅 연결 테스트 성공!\n교사 스페이스방과 보강 알리미 시스템이 정상적으로 연결되었습니다. 🎉';
    sendGoogleChatMessage(testMsg, customUrl);
    return { success: true, message: '테스트 메시지가 성공적으로 발송되었습니다.' };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 🔑 [최초 1회 실행 전용] 구글 필수 권한 승인 강제 호출 함수
 * try/catch 없이 직접 호출하여 구글의 권한 검토 팝업창을 발생시킵니다.
 */
function authorizePermissions() {
  Logger.log('1. UrlFetchApp(웹훅 전송) 권한 확인 중...');
  var res = UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true });
  Logger.log('UrlFetchApp 응답 코드: ' + res.getResponseCode());

  Logger.log('2. ScriptApp(낮 12시 트리거) 권한 확인 중...');
  var triggers = ScriptApp.getProjectTriggers();
  Logger.log('등록된 트리거 개수: ' + triggers.length);

  Logger.log('🎉 축하합니다! 모든 필수 권한(UrlFetchApp, ScriptApp)이 정상 승인되었습니다.');
}



