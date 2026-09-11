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
        subClass: String(row[4] || '-'),
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
      record.subClass || '-',
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
          record.subClass || '-',
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
