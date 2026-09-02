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

/**
 * 데이터베이스 역할을 하는 구글 스프레드시트를 가져오거나 없으면 자동 생성합니다.
 */
function getDbSheet() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var ssId = scriptProperties.getProperty('SPREADSHEET_ID');
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

  // 헤더 생성 확인
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID', '날짜', '교시', '학급', '원교사', '보강교사', '사유', '등록시각']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#006b67').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  return sheet;
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
  var match = str.match(/^(\d{4})[-.\/]?(\d{2})[-.\/]?(\d{2})/);
  if (match) {
    return match[1] + '-' + match[2] + '-' + match[3];
  }
  return str;
}

/**
 * 보강 내역을 조회합니다. (단일 날짜 또는 시작일~종료일 기간 검색 지원)
 * @param {string} startDate 시작 날짜 (YYYY-MM-DD 또는 'ALL')
 * @param {string} endDate 종료 날짜 (YYYY-MM-DD, 옵션)
 */
function getSubstitutionRecords(startDate, endDate) {
  try {
    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    var records = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0]) continue; // ID가 없는 행 스킵

      var rowDate = formatDateString(row[1]);

      // 날짜 필터링 (전체, 단일 날짜, 또는 기간 검색)
      if (startDate && startDate !== 'ALL') {
        if (endDate && endDate.trim() !== '') {
          if (rowDate < startDate || rowDate > endDate) continue;
        } else {
          if (rowDate !== startDate) continue;
        }
      }

      records.push({
        id: String(row[0]),
        date: rowDate,
        period: String(row[2]),
        className: String(row[3]),
        originalTeacher: String(row[4]),
        substituteTeacher: String(row[5]),
        reason: String(row[6]),
        timestamp: row[7] ? String(row[7]) : ''
      });
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
 * 신규 보강 내역을 저장합니다.
 */
function addSubstitutionRecord(record) {
  try {
    if (!record.date || !record.period || !record.className || !record.substituteTeacher || !record.originalTeacher) {
      throw new Error('필수 입력 항목이 누락되었습니다.');
    }

    var sheet = getDbSheet();
    var newId = 'SUB-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    var nowIso = new Date().toISOString();
    var formattedDate = formatDateString(record.date);

    sheet.appendRow([
      newId,
      formattedDate,
      record.period,
      record.className,
      record.originalTeacher,
      record.substituteTeacher,
      record.reason || '사유 없음',
      nowIso
    ]);

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
 * 기존 보강 내역을 수정합니다.
 */
function updateSubstitutionRecord(record) {
  try {
    if (!record.id || !record.date || !record.period || !record.className || !record.substituteTeacher || !record.originalTeacher) {
      throw new Error('필수 수정 정보가 누락되었습니다.');
    }

    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();
    var formattedDate = formatDateString(record.date);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(record.id)) {
        var rowNum = i + 1;
        sheet.getRange(rowNum, 2).setValue(formattedDate);
        sheet.getRange(rowNum, 3).setValue(record.period);
        sheet.getRange(rowNum, 4).setValue(record.className);
        sheet.getRange(rowNum, 5).setValue(record.originalTeacher);
        sheet.getRange(rowNum, 6).setValue(record.substituteTeacher);
        sheet.getRange(rowNum, 7).setValue(record.reason || '사유 없음');
        sheet.getRange(rowNum, 8).setValue(new Date().toISOString());

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
 * 보강 내역을 삭제합니다.
 */
function deleteSubstitutionRecord(id) {
  try {
    var sheet = getDbSheet();
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: '보강 내역이 삭제되었습니다.' };
      }
    }
    return { success: false, message: '해당 보강 내역을 찾을 수 없습니다.' };
  } catch (err) {
    Logger.log('Error in deleteSubstitutionRecord: ' + err.toString());
    return { success: false, message: err.message };
  }
}
