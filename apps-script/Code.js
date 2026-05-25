// ============================================================
// 한얼 스케줄 관리 시스템 - Google Apps Script
// Google Sheets ↔ Google Calendar 양방향 동기화
// ============================================================

const SPREADSHEET_ID = '1hMfhClmRJ5edl-fmtwWytk5B8eFWMyvsKx8WkbSIqiE';

const CALENDAR_ID_EVENT = 'haneol@somaandbody.com';
const CALENDAR_ID_ROUTINE = 'c_177b6987863f353fcd46f459f0c5f7f30cd6d1bb3ee3e727fbfdf7dff2185a2d@group.calendar.google.com';

const SHEET_ROUTINE = '루틴';
const SHEET_EVENT = '일정';
const SHEET_CARE = '케어';
const SHEET_TASTE = '취향';
const SHEET_HUMOR = '유머';
const SHEET_TODO_LIST = '할일리스트';
const SHEET_TODO = '할일';
const DATA_START_ROW = 4; // Row1=시스템, Row2=공백, Row3=헤더, Row4~=데이터

// 새 할일 리스트 생성 시 색상 자동 순환 배정
const TODO_LIST_COLORS = ['mint', 'sky', 'rose', 'gold', 'accent', 'lavender'];

// 기본 리스트 (시트 초기 생성 시 자동 추가)
const DEFAULT_TODO_LISTS = ['짓기', '살기', '아이디어'];

// ============================================================
// 웹 API 라우팅
// ============================================================

function doGet(e) {
  const action = e.parameter.action;

  // action이 없으면 HTML 페이지 서빙
  if (!action) {
    const template = HtmlService.createTemplateFromFile('Schedule');
    template.apiUrl = ScriptApp.getService().getUrl();
    return template.evaluate()
      .setTitle('한얼 스케줄')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // action이 있으면 JSON API
  const data = e.parameter.data ? JSON.parse(e.parameter.data) : null;
  const id = e.parameter.id || null;
  let result;

  try {
    switch (action) {
      case 'getRoutines':
        result = getRoutines();
        break;
      case 'getEvents':
        result = getEvents(e.parameter.startDate, e.parameter.endDate);
        break;
      case 'getWeekView':
        result = getWeekView(e.parameter.date);
        break;
      case 'addRoutine':
        result = addRoutine(data);
        break;
      case 'updateRoutine':
        result = updateRoutine(data);
        break;
      case 'deleteRoutine':
        result = deleteRoutine(id);
        break;
      case 'addEvent':
        result = addEvent(data);
        break;
      case 'updateEvent':
        result = updateEvent(data);
        break;
      case 'deleteEvent':
        result = deleteEvent(id);
        break;
      case 'getCares':
        result = getCares();
        break;
      case 'addCare':
        result = addCare(data);
        break;
      case 'updateCare':
        result = updateCare(data);
        break;
      case 'deleteCare':
        result = deleteCare(id);
        break;
      case 'syncCalendar':
        result = fullSync();
        break;
      case 'cleanupGarbage':
        result = cleanupAllGarbageExceptions();
        break;
      case 'getTastes':
        result = getTastes();
        break;
      case 'addTaste':
        result = addTaste(data);
        break;
      case 'updateTaste':
        result = updateTaste(data);
        break;
      case 'deleteTaste':
        result = deleteTaste(id);
        break;
      case 'getHumors':
        result = getHumors();
        break;
      case 'addHumor':
        result = addHumor(data);
        break;
      case 'updateHumor':
        result = updateHumor(data);
        break;
      case 'deleteHumor':
        result = deleteHumor(id);
        break;
      case 'getTodoLists':
        result = getTodoLists();
        break;
      case 'addTodoList':
        result = addTodoList(data);
        break;
      case 'updateTodoList':
        result = updateTodoList(data);
        break;
      case 'deleteTodoList':
        result = deleteTodoList(id);
        break;
      case 'getTodos':
        result = getTodos(e.parameter.listId);
        break;
      case 'addTodo':
        result = addTodo(data);
        break;
      case 'updateTodo':
        result = updateTodo(data);
        break;
      case 'deleteTodo':
        result = deleteTodo(id);
        break;
      case 'syncTodosToGoogleTasks':
        result = syncAllTodosToGoogleTasks();
        break;
      default:
        result = { error: 'Unknown action: ' + action };
    }
  } catch (err) {
    result = { error: err.message, stack: err.stack };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 헬퍼 함수
// ============================================================

function getSheet(name) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
}

function generateId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
}

function now() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm");
}

function toKSTString(date) {
  return Utilities.formatDate(date, 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm");
}

function formatTime(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Seoul', 'HH:mm');
  return String(val);
}

function formatDateVal(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, 'Asia/Seoul', 'yyyy-MM-dd');
  return String(val);
}

// 요일 매핑
const DAY_MAP = { '월': 1, '화': 2, '수': 3, '목': 4, '금': 5, '토': 6, '일': 0 };
const DAY_MAP_REVERSE = { 0: '일', 1: '월', 2: '화', 3: '수', 4: '목', 5: '금', 6: '토' };
const RRULE_DAY = { '월': 'MO', '화': 'TU', '수': 'WE', '목': 'TH', '금': 'FR', '토': 'SA', '일': 'SU' };

// ============================================================
// 루틴 CRUD
// ============================================================

function getRoutines() {
  const sheet = getSheet(SHEET_ROUTINE);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 11).getValues();
  return data.filter(row => row[0] !== '').map(row => ({
    id: row[0],
    title: row[1],
    dayOfWeek: String(row[2]),
    startTime: formatTime(row[3]),
    endTime: formatTime(row[4]),
    calendarEventId: row[5],
    lastModified: row[6],
    startDate: formatDateVal(row[7]),
    endDate: formatDateVal(row[8]),
    address: row[9] || '',
    memo: row[10] || ''
  }));
}

function addRoutine(data) {
  const sheet = getSheet(SHEET_ROUTINE);
  const id = generateId('routine');
  const timestamp = now();

  // Calendar에 recurring event 생성
  const calEventId = createRecurringCalendarEvent(data, id);

  sheet.appendRow([
    id,
    data.title,
    data.dayOfWeek,
    data.startTime,
    data.endTime,
    calEventId,
    timestamp,
    data.startDate || '',
    data.endDate || '',
    data.address || '',
    data.memo || ''
  ]);

  return { success: true, id: id, calendarEventId: calEventId };
}

function updateRoutine(data) {
  const sheet = getSheet(SHEET_ROUTINE);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      const oldCalId = sheet.getRange(row, 6).getValue();

      // Calendar 업데이트
      if (oldCalId) {
        try {
          CalendarApp.getCalendarById(CALENDAR_ID_ROUTINE).getEventById(oldCalId).deleteEvent();
        } catch (e) { /* 이미 삭제된 경우 무시 */ }
      }
      const newCalId = createRecurringCalendarEvent(data, data.id);

      sheet.getRange(row, 2, 1, 10).setValues([[
        data.title, data.dayOfWeek, data.startTime, data.endTime, newCalId, now(), data.startDate || '', data.endDate || '', data.address || '', data.memo || ''
      ]]);

      // 구 루틴 시리즈에서 자동 생성된 예외들 제거
      // (루틴 원본 변경 시 stale 예외가 주간뷰를 덮어쓰는 문제 해결)
      cleanupStaleRoutineExceptions(data.id);

      return { success: true };
    }
  }
  return { error: 'Routine not found' };
}

// source='calendar', status='modified'인 예외 중 지정한 routineId에 해당하는 것들 삭제
// 루틴 원본 변경 시 호출되어 stale 예외를 정리
function cleanupStaleRoutineExceptions(routineId) {
  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return 0;

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();
  let deleted = 0;
  // 뒤에서부터 순회해야 deleteRow 후 인덱스가 꼬이지 않음
  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (row[0] && row[8] === routineId && row[7] === 'calendar' && row[9] === 'modified') {
      sheet.deleteRow(DATA_START_ROW + i);
      deleted++;
    }
  }
  return deleted;
}

// 이벤트 시트 전체를 스캔해서 자동 생성된 garbage 예외 정리
// 1) 고아 예외 (routineId가 존재하지 않는 루틴 참조)
// 2) 오래된 시리즈 참조 (calendarEventId가 현재 루틴의 calId와 다름)
// 3) 현재 루틴 원본과 시간/제목이 완전히 동일한 예외 (진짜 변경이 아닌 자동 생성 garbage)
function cleanupAllGarbageExceptions() {
  const routines = getRoutines();
  const routineMap = {};
  routines.forEach(r => { routineMap[r.id] = r; });

  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { deleted: 0 };

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();
  let deleted = 0;

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    if (!row[0]) continue;
    // source='calendar' AND status='modified' AND routineId 있음
    if (row[7] !== 'calendar' || row[9] !== 'modified' || !row[8]) continue;

    const routine = routineMap[row[8]];

    // 1) 고아: 루틴이 사라짐
    if (!routine) {
      sheet.deleteRow(DATA_START_ROW + i);
      deleted++;
      continue;
    }

    // 2) stale: calendarEventId가 현재 루틴의 calId와 다름 (구 시리즈 참조)
    if (row[5] !== routine.calendarEventId) {
      sheet.deleteRow(DATA_START_ROW + i);
      deleted++;
      continue;
    }

    // 3) 원본과 동일: 자동 생성된 garbage
    const startTime = formatTime(row[3]);
    const endTime = formatTime(row[4]);
    if (row[1] === routine.title && startTime === routine.startTime && endTime === routine.endTime) {
      sheet.deleteRow(DATA_START_ROW + i);
      deleted++;
    }
  }

  return { deleted: deleted };
}

function deleteRoutine(id) {
  const sheet = getSheet(SHEET_ROUTINE);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      const row = DATA_START_ROW + i;
      const calId = sheet.getRange(row, 6).getValue();

      // Calendar에서 삭제
      if (calId) {
        try {
          CalendarApp.getCalendarById(CALENDAR_ID_ROUTINE).getEventById(calId).deleteEvent();
        } catch (e) { /* 무시 */ }
      }

      sheet.deleteRow(row);
      return { success: true };
    }
  }
  return { error: 'Routine not found' };
}

// ============================================================
// 일정 CRUD
// ============================================================

function getEvents(startDate, endDate) {
  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 12).getValues();
  return data.filter(row => {
    if (row[0] === '') return false;
    const rowDate = formatDateVal(row[2]);
    if (startDate && rowDate < startDate) return false;
    if (endDate && rowDate > endDate) return false;
    return true;
  }).map(row => ({
    id: row[0],
    title: row[1],
    date: formatDateVal(row[2]),
    startTime: formatTime(row[3]),
    endTime: formatTime(row[4]),
    calendarEventId: row[5],
    lastModified: row[6],
    source: row[7],
    routineId: row[8],
    status: row[9],
    address: row[10] || '',
    memo: row[11] || ''
  }));
}

function addEvent(data) {
  const sheet = getSheet(SHEET_EVENT);
  const id = generateId('event');
  const timestamp = now();

  // Calendar에 단일 이벤트 생성
  const calEventId = createSingleCalendarEvent(data);

  sheet.appendRow([
    id,
    data.title,
    data.date,
    data.startTime,
    data.endTime,
    calEventId,
    timestamp,
    'sheets',
    data.routineId || '',
    data.status || 'active',
    data.address || '',
    data.memo || ''
  ]);

  return { success: true, id: id, calendarEventId: calEventId };
}

function updateEvent(data) {
  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      const oldCalId = sheet.getRange(row, 6).getValue();

      // Calendar 업데이트
      if (oldCalId && data.status !== 'cancelled') {
        try {
          const cal = CalendarApp.getCalendarById(CALENDAR_ID_EVENT);
          const event = cal.getEventById(oldCalId);
          if (event) {
            const start = new Date(data.date + 'T' + data.startTime + ':00');
            const end = new Date(data.date + 'T' + data.endTime + ':00');
            event.setTime(start, end);
            event.setTitle(data.title);
          }
        } catch (e) { /* 무시 */ }
      }

      sheet.getRange(row, 2, 1, 11).setValues([[
        data.title, data.date, data.startTime, data.endTime,
        oldCalId, now(), data.source || 'sheets',
        data.routineId || '', data.status || 'active',
        data.address || '', data.memo || ''
      ]]);
      return { success: true };
    }
  }
  return { error: 'Event not found' };
}

function deleteEvent(id) {
  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      const row = DATA_START_ROW + i;
      const calId = sheet.getRange(row, 6).getValue();

      if (calId) {
        try {
          CalendarApp.getCalendarById(CALENDAR_ID_EVENT).getEventById(calId).deleteEvent();
        } catch (e) { /* 무시 */ }
      }

      sheet.deleteRow(row);
      return { success: true };
    }
  }
  return { error: 'Event not found' };
}

// ============================================================
// 케어 CRUD
// ============================================================

function getCares() {
  const sheet = getSheet(SHEET_CARE);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8).getValues();
  return data.filter(row => row[0] !== '').map(row => ({
    id: row[0],
    title: row[1],
    cycle: String(row[2]),
    duration: String(row[3] || 1),
    lastDone: formatDateVal(row[4]),
    nextDate: formatDateVal(row[5]),
    keywords: row[6] || '',
    memo: row[7] || ''
  }));
}

function addCare(data) {
  const sheet = getSheet(SHEET_CARE);
  if (!sheet) return { error: '케어 시트가 없습니다. initializeSheets()를 실행해주세요.' };
  const id = generateId('care');

  sheet.appendRow([
    id,
    data.title,
    parseInt(data.cycle),
    parseInt(data.duration) || 1,
    data.lastDone || '',
    data.nextDate || '',
    data.keywords || '',
    data.memo || ''
  ]);

  return { success: true, id: id };
}

function updateCare(data) {
  const sheet = getSheet(SHEET_CARE);
  if (!sheet) return { error: '케어 시트가 없습니다.' };
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      sheet.getRange(row, 2, 1, 7).setValues([[
        data.title,
        parseInt(data.cycle),
        parseInt(data.duration) || 1,
        data.lastDone || '',
        data.nextDate || '',
        data.keywords || '',
        data.memo || ''
      ]]);
      return { success: true };
    }
  }
  return { error: 'Care not found' };
}

function deleteCare(id) {
  const sheet = getSheet(SHEET_CARE);
  if (!sheet) return { error: '케어 시트가 없습니다.' };
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(DATA_START_ROW + i);
      return { success: true };
    }
  }
  return { error: 'Care not found' };
}

// ============================================================
// 취향 CRUD (영화/음악/전시/책 등)
// ============================================================

// 취향 시트가 없으면 자동 생성 (헤더 포함)
function ensureTasteSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TASTE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TASTE);
    sheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 취향');
    sheet.getRange(3, 1, 1, 7).setValues([[
      '아이디', '카테고리', '제목', '부가정보', '별점', '메모', '추가일시'
    ]]);
  }
  return sheet;
}

function getTastes() {
  const sheet = ensureTasteSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 7).getValues();
  return data.filter(row => row[0] !== '').map(row => ({
    id: row[0],
    category: row[1] || '',
    title: row[2] || '',
    subtitle: row[3] || '',
    rating: row[4] ? parseInt(row[4]) : 0,
    memo: row[5] || '',
    addedAt: row[6] ? formatDateVal(row[6]) : ''
  }));
}

function addTaste(data) {
  const sheet = ensureTasteSheet();
  const id = generateId('taste');
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  sheet.appendRow([
    id,
    data.category || '기타',
    data.title || '',
    data.subtitle || '',
    data.rating ? parseInt(data.rating) : '',
    data.memo || '',
    today
  ]);

  return { success: true, id: id };
}

function updateTaste(data) {
  const sheet = ensureTasteSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      // 추가일시(7열)는 그대로 두고 나머지만 업데이트
      sheet.getRange(row, 2, 1, 5).setValues([[
        data.category || '기타',
        data.title || '',
        data.subtitle || '',
        data.rating ? parseInt(data.rating) : '',
        data.memo || ''
      ]]);
      return { success: true };
    }
  }
  return { error: 'Taste not found' };
}

function deleteTaste(id) {
  const sheet = ensureTasteSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(DATA_START_ROW + i);
      return { success: true };
    }
  }
  return { error: 'Taste not found' };
}

// ============================================================
// 유머 CRUD (유머러스한 기억)
// ============================================================

function ensureHumorSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_HUMOR);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_HUMOR);
    sheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 유머');
    sheet.getRange(3, 1, 1, 6).setValues([[
      '아이디', '제목', '내용', '태그', '날짜', '추가일시'
    ]]);
  }
  return sheet;
}

function getHumors() {
  const sheet = ensureHumorSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  return data.filter(row => row[0] !== '').map(row => ({
    id: row[0],
    title: row[1] || '',
    content: row[2] || '',
    tags: row[3] || '',
    date: row[4] ? formatDateVal(row[4]) : '',
    addedAt: row[5] ? formatDateVal(row[5]) : ''
  }));
}

function addHumor(data) {
  const sheet = ensureHumorSheet();
  const id = generateId('humor');
  const today = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd');

  sheet.appendRow([
    id,
    data.title || '',
    data.content || '',
    data.tags || '',
    data.date || '',
    today
  ]);

  return { success: true, id: id };
}

function updateHumor(data) {
  const sheet = ensureHumorSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      // 추가일시(6열)는 보존, 나머지만 업데이트
      sheet.getRange(row, 2, 1, 4).setValues([[
        data.title || '',
        data.content || '',
        data.tags || '',
        data.date || ''
      ]]);
      return { success: true };
    }
  }
  return { error: 'Humor not found' };
}

function deleteHumor(id) {
  const sheet = ensureHumorSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      sheet.deleteRow(DATA_START_ROW + i);
      return { success: true };
    }
  }
  return { error: 'Humor not found' };
}

// ============================================================
// 할일 CRUD (리스트업 — 캘린더 연동 없음, 2단계 들여쓰기)
// ============================================================
// 구조: 할일리스트 시트(리스트 메타) + 할일 시트(항목)
// 항목은 parentId가 비어있으면 top-level, 값 있으면 하위 항목

// 할일리스트 시트 (6열): 아이디 | 이름 | 순서 | 색상 | 수정일시 | GoogleListId
function ensureTodoListSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TODO_LIST);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TODO_LIST);
    sheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 할일리스트');
    sheet.getRange(3, 1, 1, 6).setValues([[
      '아이디', '이름', '순서', '색상', '수정일시', 'GoogleListId'
    ]]);

    // 기본 리스트 자동 생성 (짓기, 살기, 아이디어)
    const timestamp = now();
    DEFAULT_TODO_LISTS.forEach((name, idx) => {
      const id = generateId('todolist');
      const googleListId = gtaskCreateList(name); // best-effort
      sheet.appendRow([
        id, name, idx, TODO_LIST_COLORS[idx % TODO_LIST_COLORS.length], timestamp, googleListId || ''
      ]);
    });
  } else {
    // 기존 시트인 경우 헤더가 5열이면 6열로 업그레이드
    const headerLen = sheet.getRange(3, 1, 1, 6).getValues()[0].filter(v => v !== '').length;
    if (headerLen < 6) {
      sheet.getRange(3, 6).setValue('GoogleListId');
    }
  }
  return sheet;
}

// 할일 시트 (8열): 아이디 | 리스트아이디 | 부모아이디 | 텍스트 | 완료여부 | 순서 | 수정일시 | GoogleTaskId
function ensureTodoSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_TODO);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TODO);
    sheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 할일');
    sheet.getRange(3, 1, 1, 8).setValues([[
      '아이디', '리스트아이디', '부모아이디', '텍스트', '완료여부', '순서', '수정일시', 'GoogleTaskId'
    ]]);
  } else {
    const headerLen = sheet.getRange(3, 1, 1, 8).getValues()[0].filter(v => v !== '').length;
    if (headerLen < 8) {
      sheet.getRange(3, 8).setValue('GoogleTaskId');
    }
  }
  return sheet;
}

// ── 할일 리스트 (카드) CRUD ──

function getTodoLists() {
  const sheet = ensureTodoListSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  return data.filter(row => row[0] !== '').map(row => ({
    id: row[0],
    name: row[1] || '',
    order: typeof row[2] === 'number' ? row[2] : parseInt(row[2]) || 0,
    color: row[3] || 'accent',
    lastModified: row[4],
    googleListId: row[5] || ''
  })).sort((a, b) => a.order - b.order);
}

function addTodoList(data) {
  const sheet = ensureTodoListSheet();
  const id = generateId('todolist');
  const timestamp = now();

  // 새 리스트 순서 = 현재 최대 순서 + 1
  const existing = getTodoLists();
  const nextOrder = existing.length > 0 ? Math.max.apply(null, existing.map(l => l.order)) + 1 : 0;
  const color = data.color || TODO_LIST_COLORS[existing.length % TODO_LIST_COLORS.length];
  const name = data.name || '새 리스트';

  // Google Tasks에 list 생성 (best-effort, 권한 없거나 실패해도 진행)
  const googleListId = gtaskCreateList(name);

  sheet.appendRow([
    id,
    name,
    typeof data.order === 'number' ? data.order : nextOrder,
    color,
    timestamp,
    googleListId || ''
  ]);

  return { success: true, id: id, googleListId: googleListId || '' };
}

function updateTodoList(data) {
  const sheet = ensureTodoListSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const ids = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      // 부분 업데이트: 보내지 않은 필드는 보존
      const current = sheet.getRange(row, 2, 1, 5).getValues()[0];
      const newName = data.name !== undefined ? data.name : current[0];
      const googleListId = current[4];

      // 이름 변경 시 Google Tasks list title 갱신
      if (data.name !== undefined && data.name !== current[0] && googleListId) {
        gtaskUpdateList(googleListId, newName);
      }

      sheet.getRange(row, 2, 1, 4).setValues([[
        newName,
        typeof data.order === 'number' ? data.order : current[1],
        data.color !== undefined ? data.color : current[2],
        now()
      ]]);
      return { success: true };
    }
  }
  return { error: 'TodoList not found' };
}

// 리스트 삭제 시 해당 리스트의 모든 항목도 함께 삭제
function deleteTodoList(id) {
  const listSheet = ensureTodoListSheet();
  const lastRow = listSheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  // 0. 삭제할 리스트의 googleListId 미리 조회
  let googleListId = '';
  const allListData = listSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 6).getValues();
  for (let i = 0; i < allListData.length; i++) {
    if (allListData[i][0] === id) {
      googleListId = allListData[i][5] || '';
      break;
    }
  }

  // 1. 해당 리스트의 항목들 모두 삭제
  const todoSheet = ensureTodoSheet();
  const todoLastRow = todoSheet.getLastRow();
  if (todoLastRow >= DATA_START_ROW) {
    const todoData = todoSheet.getRange(DATA_START_ROW, 1, todoLastRow - DATA_START_ROW + 1, 2).getValues();
    // 뒤에서부터 삭제 (인덱스 꼬임 방지)
    for (let i = todoData.length - 1; i >= 0; i--) {
      if (todoData[i][0] && todoData[i][1] === id) {
        todoSheet.deleteRow(DATA_START_ROW + i);
      }
    }
  }

  // 2. 리스트 메타 삭제
  const ids = listSheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === id) {
      listSheet.deleteRow(DATA_START_ROW + i);

      // 3. Google Tasks list 삭제 (best-effort)
      if (googleListId) gtaskDeleteList(googleListId);

      return { success: true };
    }
  }
  return { error: 'TodoList not found' };
}

// ── 할일 항목 CRUD ──

// listId 지정 시 해당 리스트만, 미지정 시 전체
function getTodos(listId) {
  const sheet = ensureTodoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return [];

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8).getValues();
  return data.filter(row => {
    if (row[0] === '') return false;
    if (listId && row[1] !== listId) return false;
    return true;
  }).map(row => ({
    id: row[0],
    listId: row[1] || '',
    parentId: row[2] || '',
    text: row[3] || '',
    done: row[4] === true || row[4] === 'TRUE' || row[4] === 'true',
    order: typeof row[5] === 'number' ? row[5] : parseInt(row[5]) || 0,
    lastModified: row[6],
    googleTaskId: row[7] || ''
  })).sort((a, b) => a.order - b.order);
}

// Sheets 리스트 id → googleListId 조회 (Google Tasks 호출용)
function getGoogleListIdForList(listId) {
  const lists = getTodoLists();
  const found = lists.find(l => l.id === listId);
  return found ? (found.googleListId || '') : '';
}

// Sheets 할일 id → googleTaskId 조회
function getGoogleTaskIdForTodo(todoId) {
  const todos = getTodos();
  const found = todos.find(t => t.id === todoId);
  return found ? (found.googleTaskId || '') : '';
}

function addTodo(data) {
  const sheet = ensureTodoSheet();
  if (!data.listId) return { error: 'listId is required' };

  const id = generateId('todo');
  const timestamp = now();

  // 같은 리스트 + 같은 부모 안에서 다음 순서
  let nextOrder = 0;
  if (typeof data.order !== 'number') {
    const siblings = getTodos(data.listId).filter(t => (t.parentId || '') === (data.parentId || ''));
    nextOrder = siblings.length > 0 ? Math.max.apply(null, siblings.map(t => t.order)) + 1 : 0;
  } else {
    nextOrder = data.order;
  }

  // Google Tasks에 task 생성 (best-effort)
  const googleListId = getGoogleListIdForList(data.listId);
  const parentGoogleTaskId = data.parentId ? getGoogleTaskIdForTodo(data.parentId) : '';
  const googleTaskId = googleListId
    ? gtaskCreate(googleListId, data.text || '', parentGoogleTaskId, data.done === true)
    : '';

  sheet.appendRow([
    id,
    data.listId,
    data.parentId || '',
    data.text || '',
    data.done === true,
    nextOrder,
    timestamp,
    googleTaskId || ''
  ]);

  return { success: true, id: id, googleTaskId: googleTaskId || '' };
}

function updateTodo(data) {
  const sheet = ensureTodoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const allData = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8).getValues();
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][0] === data.id) {
      const row = DATA_START_ROW + i;
      const current = {
        listId: allData[i][1],
        parentId: allData[i][2],
        text: allData[i][3],
        done: allData[i][4] === true || allData[i][4] === 'TRUE' || allData[i][4] === 'true',
        order: allData[i][5],
        googleTaskId: allData[i][7] || ''
      };

      // 부분 업데이트 적용 값
      const newText = data.text !== undefined ? data.text : current.text;
      const newDone = data.done !== undefined ? (data.done === true) : current.done;
      const newListId = data.listId !== undefined ? data.listId : current.listId;
      const newParentId = data.parentId !== undefined ? data.parentId : current.parentId;

      // Google Tasks sync (best-effort)
      const googleListId = getGoogleListIdForList(newListId);
      if (googleListId && current.googleTaskId) {
        // 같은 리스트 안에서 text/done 변경
        gtaskUpdate(googleListId, current.googleTaskId, { text: newText, done: newDone });
      }

      sheet.getRange(row, 2, 1, 6).setValues([[
        newListId,
        newParentId,
        newText,
        newDone,
        typeof data.order === 'number' ? data.order : current.order,
        now()
      ]]);
      return { success: true };
    }
  }
  return { error: 'Todo not found' };
}

// 항목 삭제 시 자식(부모아이디 === id)도 함께 삭제
function deleteTodo(id) {
  const sheet = ensureTodoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) return { error: 'No data' };

  const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 8).getValues();
  let targetFound = false;
  // 뒤에서부터 삭제: target과 자식 모두 제거 + Google Tasks 동기 삭제
  for (let i = data.length - 1; i >= 0; i--) {
    if (!data[i][0]) continue;
    if (data[i][0] === id || data[i][2] === id) {
      const rowListId = data[i][1];
      const rowGoogleTaskId = data[i][7] || '';
      if (rowGoogleTaskId) {
        const googleListId = getGoogleListIdForList(rowListId);
        if (googleListId) gtaskDelete(googleListId, rowGoogleTaskId);
      }
      sheet.deleteRow(DATA_START_ROW + i);
      if (data[i][0] === id) targetFound = true;
    }
  }
  return targetFound ? { success: true } : { error: 'Todo not found' };
}

// ============================================================
// Google Tasks API 헬퍼 (Phase 2 — Sheets → Google Tasks 단방향 sync)
// ============================================================
// 사전 조건:
//   1) appsscript.json에 Tasks Advanced Service 활성화 (이미 적용됨)
//   2) 한얼이 에디터에서 testTasksAuth() 1회 실행 → 권한 검토 → 승인
// 권한 승인 전이거나 실패 시: 모든 함수가 빈 ID 반환 + 로그 → Sheets 동작 영향 없음

function gtaskCreateList(name) {
  try {
    if (typeof Tasks === 'undefined') return '';
    const result = Tasks.Tasklists.insert({ title: name });
    return result.id || '';
  } catch (e) {
    Logger.log('gtaskCreateList error: ' + e.message);
    return '';
  }
}

function gtaskUpdateList(googleListId, name) {
  try {
    if (typeof Tasks === 'undefined' || !googleListId) return;
    Tasks.Tasklists.patch({ title: name }, googleListId);
  } catch (e) {
    Logger.log('gtaskUpdateList error: ' + e.message);
  }
}

function gtaskDeleteList(googleListId) {
  try {
    if (typeof Tasks === 'undefined' || !googleListId) return;
    Tasks.Tasklists.remove(googleListId);
  } catch (e) {
    Logger.log('gtaskDeleteList error: ' + e.message);
  }
}

function gtaskCreate(googleListId, text, parentGoogleTaskId, done) {
  try {
    if (typeof Tasks === 'undefined' || !googleListId) return '';
    const task = { title: text || '', status: done ? 'completed' : 'needsAction' };
    const options = parentGoogleTaskId ? { parent: parentGoogleTaskId } : {};
    const result = Tasks.Tasks.insert(task, googleListId, options);
    return result.id || '';
  } catch (e) {
    Logger.log('gtaskCreate error: ' + e.message);
    return '';
  }
}

function gtaskUpdate(googleListId, googleTaskId, fields) {
  try {
    if (typeof Tasks === 'undefined' || !googleListId || !googleTaskId) return;
    const updates = {};
    if (fields.text !== undefined) updates.title = fields.text;
    if (fields.done !== undefined) updates.status = fields.done ? 'completed' : 'needsAction';
    Tasks.Tasks.patch(updates, googleListId, googleTaskId);
  } catch (e) {
    Logger.log('gtaskUpdate error: ' + e.message);
  }
}

function gtaskDelete(googleListId, googleTaskId) {
  try {
    if (typeof Tasks === 'undefined' || !googleListId || !googleTaskId) return;
    Tasks.Tasks.remove(googleListId, googleTaskId);
  } catch (e) {
    Logger.log('gtaskDelete error: ' + e.message);
  }
}

// ── 권한 승인 트리거용 임시 함수 ──
// 한얼이 Apps Script 에디터에서 함수 드롭다운 → testTasksAuth 선택 → 실행
// → "권한 검토" → "고급" → "{프로젝트명}(으)로 이동" → "허용" → 다음부터 자동 sync 활성
function testTasksAuth() {
  const result = Tasks.Tasklists.list();
  const count = (result.items && result.items.length) || 0;
  Logger.log('Tasks API 연결 OK. 현재 Google Tasks 리스트 수: ' + count);
  return { success: true, googleTaskListCount: count };
}

// ── 일괄 sync: 기존 Sheets 데이터를 Google Tasks로 한 번에 push ──
// Phase 1에서 만들어진 리스트/항목들 (googleListId/googleTaskId 빈 상태) 일괄 처리
// 멱등 (이미 매핑된 건 skip)
function syncAllTodosToGoogleTasks() {
  if (typeof Tasks === 'undefined') {
    return { error: 'Tasks Advanced Service 미활성. appsscript.json + clasp push --force + 권한 승인 필요' };
  }

  const listSheet = ensureTodoListSheet();
  const todoSheet = ensureTodoSheet();
  let listsCreated = 0, listsUpdated = 0, tasksCreated = 0;

  // 1. 리스트 sync
  const listLastRow = listSheet.getLastRow();
  if (listLastRow >= DATA_START_ROW) {
    const listData = listSheet.getRange(DATA_START_ROW, 1, listLastRow - DATA_START_ROW + 1, 6).getValues();
    for (let i = 0; i < listData.length; i++) {
      if (!listData[i][0]) continue;
      if (listData[i][5]) continue; // 이미 매핑됨
      const googleListId = gtaskCreateList(listData[i][1]);
      if (googleListId) {
        listSheet.getRange(DATA_START_ROW + i, 6).setValue(googleListId);
        listsCreated++;
      }
    }
  }

  // 리스트 id → googleListId 맵 (방금 만든 것 포함)
  const lists = getTodoLists();
  const listIdMap = {};
  lists.forEach(l => { listIdMap[l.id] = l.googleListId; });

  // 2. 항목 sync — top-level 먼저 (parent 없는 것), 그 다음 children
  const todoLastRow = todoSheet.getLastRow();
  if (todoLastRow >= DATA_START_ROW) {
    const todoData = todoSheet.getRange(DATA_START_ROW, 1, todoLastRow - DATA_START_ROW + 1, 8).getValues();

    // sheetRowIdx → googleTaskId 매핑 (children sync 시 parent ID 찾기용)
    const todoIdToGoogleId = {};
    todoData.forEach((row, i) => {
      if (row[0] && row[7]) todoIdToGoogleId[row[0]] = row[7];
    });

    // Pass 1: top-level
    for (let i = 0; i < todoData.length; i++) {
      const row = todoData[i];
      if (!row[0] || row[2] || row[7]) continue; // 부모 있거나 이미 매핑됨
      const googleListId = listIdMap[row[1]];
      if (!googleListId) continue;
      const googleTaskId = gtaskCreate(googleListId, row[3], '', row[4] === true || row[4] === 'TRUE' || row[4] === 'true');
      if (googleTaskId) {
        todoSheet.getRange(DATA_START_ROW + i, 8).setValue(googleTaskId);
        todoIdToGoogleId[row[0]] = googleTaskId;
        tasksCreated++;
      }
    }

    // Pass 2: children
    for (let i = 0; i < todoData.length; i++) {
      const row = todoData[i];
      if (!row[0] || !row[2] || row[7]) continue; // top-level이거나 이미 매핑됨
      const googleListId = listIdMap[row[1]];
      const parentGoogleTaskId = todoIdToGoogleId[row[2]];
      if (!googleListId || !parentGoogleTaskId) continue;
      const googleTaskId = gtaskCreate(googleListId, row[3], parentGoogleTaskId, row[4] === true || row[4] === 'TRUE' || row[4] === 'true');
      if (googleTaskId) {
        todoSheet.getRange(DATA_START_ROW + i, 8).setValue(googleTaskId);
        tasksCreated++;
      }
    }
  }

  return {
    success: true,
    listsCreated: listsCreated,
    tasksCreated: tasksCreated,
    message: '리스트 ' + listsCreated + '개, 항목 ' + tasksCreated + '개 Google Tasks에 push 완료'
  };
}

// ============================================================
// Google Calendar 생성 함수
// ============================================================

function createRecurringCalendarEvent(data, routineId) {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_ROUTINE);

  // 간격 루틴 처리
  if (data.dayOfWeek.indexOf('간격:') === 0) {
    const interval = parseInt(data.dayOfWeek.split(':')[1]) || 1;
    const baseDate = new Date(data.startDate + 'T00:00:00');
    const startParts = data.startTime.split(':');
    const endParts = data.endTime.split(':');

    const eventStart = new Date(baseDate);
    eventStart.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
    const eventEnd = new Date(baseDate);
    eventEnd.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);

    const dailyRule = CalendarApp.newRecurrence().addDailyRule().interval(interval);
    const event = cal.createEventSeries(
      data.title, eventStart, eventEnd, dailyRule,
      { description: 'routineId:' + routineId }
    );
    return event.getId();
  }

  const days = data.dayOfWeek.split(',').map(d => d.trim());

  // 시작일이 설정되어 있으면 해당 날짜 사용, 없으면 다음 해당 요일
  let baseDate;
  if (data.startDate) {
    baseDate = new Date(data.startDate + 'T00:00:00');
  } else {
    const today = new Date();
    const targetDay = DAY_MAP[days[0]];
    const daysUntil = (targetDay - today.getDay() + 7) % 7 || 7;
    baseDate = new Date(today);
    baseDate.setDate(today.getDate() + daysUntil);
  }

  const startParts = data.startTime.split(':');
  const endParts = data.endTime.split(':');

  const eventStart = new Date(baseDate);
  eventStart.setHours(parseInt(startParts[0]), parseInt(startParts[1]), 0);
  const eventEnd = new Date(baseDate);
  eventEnd.setHours(parseInt(endParts[0]), parseInt(endParts[1]), 0);

  const weeklyRule = CalendarApp.newRecurrence()
    .addWeeklyRule()
    .onlyOnWeekdays(days.map(d => {
      const map = {
        '월': CalendarApp.Weekday.MONDAY,
        '화': CalendarApp.Weekday.TUESDAY,
        '수': CalendarApp.Weekday.WEDNESDAY,
        '목': CalendarApp.Weekday.THURSDAY,
        '금': CalendarApp.Weekday.FRIDAY,
        '토': CalendarApp.Weekday.SATURDAY,
        '일': CalendarApp.Weekday.SUNDAY
      };
      return map[d];
    }));

  // 기한이 설정된 경우 반복 종료일 지정
  if (data.endDate) {
    const untilDate = new Date(data.endDate + 'T23:59:59');
    weeklyRule.until(untilDate);
  }

  const recurrence = weeklyRule;

  const event = cal.createEventSeries(
    data.title,
    eventStart,
    eventEnd,
    recurrence,
    { description: 'routineId:' + routineId }
  );

  return event.getId();
}

function createSingleCalendarEvent(data) {
  const cal = CalendarApp.getCalendarById(CALENDAR_ID_EVENT);

  const start = new Date(data.date + 'T' + data.startTime + ':00');
  const end = new Date(data.date + 'T' + data.endTime + ':00');

  const event = cal.createEvent(data.title, start, end);
  return event.getId();
}

// ============================================================
// 양방향 동기화
// ============================================================

function fullSync() {
  const results = {
    sheetsToCalendar: syncSheetsToCalendar(),
    calendarToSheets: syncCalendarToSheets()
  };
  return { success: true, results: results };
}

// Sheets → Calendar: calendarEventId가 없는 행 처리
function syncSheetsToCalendar() {
  let synced = 0;

  // 루틴 동기화
  const routineSheet = getSheet(SHEET_ROUTINE);
  const routineLastRow = routineSheet.getLastRow();
  if (routineLastRow >= DATA_START_ROW) {
    const routineData = routineSheet.getRange(DATA_START_ROW, 1, routineLastRow - DATA_START_ROW + 1, 7).getValues();
    for (let i = 0; i < routineData.length; i++) {
      if (routineData[i][0] && !routineData[i][5]) {
        const data = {
          title: routineData[i][1],
          dayOfWeek: routineData[i][2],
          startTime: routineData[i][3],
          endTime: routineData[i][4]
        };
        const calId = createRecurringCalendarEvent(data, routineData[i][0]);
        routineSheet.getRange(DATA_START_ROW + i, 6).setValue(calId);
        routineSheet.getRange(DATA_START_ROW + i, 7).setValue(now());
        synced++;
      }
    }
  }

  // 일정 동기화
  const eventSheet = getSheet(SHEET_EVENT);
  const eventLastRow = eventSheet.getLastRow();
  if (eventLastRow >= DATA_START_ROW) {
    const eventData = eventSheet.getRange(DATA_START_ROW, 1, eventLastRow - DATA_START_ROW + 1, 10).getValues();
    for (let i = 0; i < eventData.length; i++) {
      if (eventData[i][0] && !eventData[i][5] && eventData[i][9] !== 'cancelled') {
        const data = {
          title: eventData[i][1],
          date: eventData[i][2],
          startTime: eventData[i][3],
          endTime: eventData[i][4]
        };
        const calId = createSingleCalendarEvent(data);
        eventSheet.getRange(DATA_START_ROW + i, 6).setValue(calId);
        eventSheet.getRange(DATA_START_ROW + i, 7).setValue(now());
        synced++;
      }
    }
  }

  return { synced: synced };
}

// Calendar → Sheets: 최근 변경분 가져오기
function syncCalendarToSheets() {
  let synced = 0;
  const props = PropertiesService.getScriptProperties();
  const lastSyncStr = props.getProperty('lastSync');
  const lastSync = lastSyncStr ? new Date(lastSyncStr) : new Date(Date.now() - 24 * 60 * 60 * 1000);

  // 일정 캘린더에서 변경분 가져오기
  synced += syncCalendarEventsToSheet(CALENDAR_ID_EVENT, lastSync, false);

  // 루틴 캘린더에서 변경분 가져오기 (개별 인스턴스 예외 처리)
  synced += syncRoutineCalendarToSheet(CALENDAR_ID_ROUTINE, lastSync);

  props.setProperty('lastSync', new Date().toISOString());
  return { synced: synced };
}

function syncCalendarEventsToSheet(calendarId, since, isRoutine) {
  let synced = 0;
  const cal = CalendarApp.getCalendarById(calendarId);
  const now_date = new Date();
  const futureDate = new Date(now_date.getTime() + 30 * 24 * 60 * 60 * 1000); // 30일 후

  const events = cal.getEvents(since, futureDate);
  const sheet = getSheet(SHEET_EVENT);
  const lastRow = sheet.getLastRow();

  // 기존 calendarEventId 목록
  let existingCalIds = {};
  if (lastRow >= DATA_START_ROW) {
    const data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, 10).getValues();
    data.forEach((row, i) => {
      if (row[5]) existingCalIds[row[5]] = DATA_START_ROW + i;
    });
  }

  events.forEach(event => {
    const eventId = event.getId();
    const eventStart = event.getStartTime();
    const eventEnd = event.getEndTime();
    const dateStr = Utilities.formatDate(eventStart, 'Asia/Seoul', 'yyyy-MM-dd');
    const startTimeStr = Utilities.formatDate(eventStart, 'Asia/Seoul', 'HH:mm');
    const endTimeStr = Utilities.formatDate(eventEnd, 'Asia/Seoul', 'HH:mm');

    if (existingCalIds[eventId]) {
      // 기존 행 업데이트 (Calendar 쪽이 더 최근이면)
      const row = existingCalIds[eventId];
      const sheetModified = new Date(sheet.getRange(row, 7).getValue());
      const calModified = event.getLastUpdated();

      if (calModified > sheetModified) {
        sheet.getRange(row, 2, 1, 9).setValues([[
          event.getTitle(), dateStr, startTimeStr, endTimeStr,
          eventId, toKSTString(calModified), 'calendar',
          sheet.getRange(row, 9).getValue(),
          sheet.getRange(row, 10).getValue()
        ]]);
        synced++;
      }
    } else {
      // 새 이벤트 → Sheets에 추가
      const id = generateId('event');
      sheet.appendRow([
        id, event.getTitle(), dateStr, startTimeStr, endTimeStr,
        eventId, toKSTString(event.getLastUpdated()), 'calendar', '', 'active'
      ]);
      synced++;
    }
  });

  return synced;
}

function syncRoutineCalendarToSheet(calendarId, since) {
  let synced = 0;
  const cal = CalendarApp.getCalendarById(calendarId);
  const now_date = new Date();
  const futureDate = new Date(now_date.getTime() + 30 * 24 * 60 * 60 * 1000);

  const events = cal.getEvents(since, futureDate);
  const eventSheet = getSheet(SHEET_EVENT);
  const routineSheet = getSheet(SHEET_ROUTINE);

  // 루틴 매핑: calEventId → { id, title, startTime, endTime }
  // (예외 감지 시 원본과 비교하기 위해 전체 정보 보관)
  const routineLastRow = routineSheet.getLastRow();
  let routineMap = {};
  if (routineLastRow >= DATA_START_ROW) {
    const routineData = routineSheet.getRange(DATA_START_ROW, 1, routineLastRow - DATA_START_ROW + 1, 7).getValues();
    routineData.forEach(row => {
      if (row[5]) {
        routineMap[row[5]] = {
          id: row[0],
          title: row[1],
          startTime: formatTime(row[3]),
          endTime: formatTime(row[4])
        };
      }
    });
  }

  // 기존 일정시트 calendarEventId 목록
  const eventLastRow = eventSheet.getLastRow();
  let existingCalIds = {};
  if (eventLastRow >= DATA_START_ROW) {
    const data = eventSheet.getRange(DATA_START_ROW, 1, eventLastRow - DATA_START_ROW + 1, 10).getValues();
    data.forEach((row, i) => {
      if (row[5]) existingCalIds[row[5]] = DATA_START_ROW + i;
    });
  }

  events.forEach(event => {
    const eventId = event.getId();
    const desc = event.getDescription() || '';
    const routineIdMatch = desc.match(/routineId:(\S+)/);

    // recurring event의 개별 인스턴스 변경 감지
    if (routineIdMatch || event.isRecurringEvent()) {
      const parentId = event.isRecurringEvent() ? event.getId() : null;
      const routineInfo = routineIdMatch ? null : routineMap[parentId];
      const routineId = routineIdMatch ? routineIdMatch[1] : (routineInfo ? routineInfo.id : '');

      // 이미 일정시트에 있는지 확인
      if (!existingCalIds[eventId] && routineId) {
        const eventStart = event.getStartTime();
        const eventEnd = event.getEndTime();
        const dateStr = Utilities.formatDate(eventStart, 'Asia/Seoul', 'yyyy-MM-dd');
        const startTimeStr = Utilities.formatDate(eventStart, 'Asia/Seoul', 'HH:mm');
        const endTimeStr = Utilities.formatDate(eventEnd, 'Asia/Seoul', 'HH:mm');

        // 루틴 원본과 제목/시간이 모두 일치하면 "진짜 예외"가 아니므로 skip
        // (이전 버그: 모든 반복 인스턴스를 modified 예외로 기록했던 문제 수정)
        if (routineInfo &&
            event.getTitle() === routineInfo.title &&
            startTimeStr === routineInfo.startTime &&
            endTimeStr === routineInfo.endTime) {
          return;
        }

        // 원본과 다른 인스턴스만 예외로 기록
        const id = generateId('event');
        eventSheet.appendRow([
          id, event.getTitle(), dateStr, startTimeStr, endTimeStr,
          eventId, toKSTString(event.getLastUpdated()), 'calendar',
          routineId, 'modified'
        ]);
        synced++;
      }
    }
  });

  // 삭제된 인스턴스 감지 (Calendar API 한계로, 주기적 폴링 시 처리)
  // Note: Google Apps Script의 CalendarApp은 삭제된 인스턴스를 직접 감지하기 어려움
  // Calendar Advanced Service를 사용하면 가능 (아래 함수 참고)

  return synced;
}

// Calendar Advanced Service를 이용한 삭제 감지 (선택적 활성화)
function detectDeletedInstances() {
  try {
    const routineSheet = getSheet(SHEET_ROUTINE);
    const eventSheet = getSheet(SHEET_EVENT);
    const routineLastRow = routineSheet.getLastRow();
    if (routineLastRow < DATA_START_ROW) return 0;

    const routineData = routineSheet.getRange(DATA_START_ROW, 1, routineLastRow - DATA_START_ROW + 1, 7).getValues();
    let synced = 0;

    routineData.forEach(routine => {
      if (!routine[5]) return;

      const calEventId = routine[5].replace('@google.com', '');
      const now_date = new Date();
      const futureDate = new Date(now_date.getTime() + 14 * 24 * 60 * 60 * 1000);

      try {
        const instances = Calendar.Events.instances(CALENDAR_ID_ROUTINE, calEventId, {
          timeMin: now_date.toISOString(),
          timeMax: futureDate.toISOString(),
          showDeleted: true
        });

        if (instances.items) {
          instances.items.forEach(inst => {
            if (inst.status === 'cancelled') {
              const dateStr = inst.originalStartTime.dateTime
                ? inst.originalStartTime.dateTime.substring(0, 10)
                : inst.originalStartTime.date;

              // 이미 기록되어 있는지 확인
              const existingEvents = getEvents(dateStr, dateStr);
              const alreadyRecorded = existingEvents.some(e =>
                e.routineId === routine[0] && e.date === dateStr && e.status === 'cancelled'
              );

              if (!alreadyRecorded) {
                const id = generateId('event');
                const days = routine[2].split(',').map(d => d.trim());
                eventSheet.appendRow([
                  id, routine[1], dateStr, routine[3], routine[4],
                  '', now(), 'calendar', routine[0], 'cancelled'
                ]);
                synced++;
              }
            }
          });
        }
      } catch (e) {
        // Calendar Advanced Service 미활성화 시 무시
        Logger.log('Advanced Calendar not available: ' + e.message);
      }
    });

    return synced;
  } catch (e) {
    Logger.log('detectDeletedInstances error: ' + e.message);
    return 0;
  }
}

// ============================================================
// 주간 뷰 생성
// ============================================================

function getWeekView(dateStr) {
  const targetDate = dateStr ? new Date(dateStr) : new Date();
  const dayOfWeek = targetDate.getDay();
  const monday = new Date(targetDate);
  monday.setDate(targetDate.getDate() - ((dayOfWeek + 6) % 7));

  const weekDays = [];
  for (let i = 0; i < 7; i++) {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    weekDays.push(Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd'));
  }

  const startDate = weekDays[0];
  const endDate = weekDays[6];

  // 루틴에서 이번 주 일정 생성
  const routines = getRoutines();
  const events = getEvents(startDate, endDate);

  const weekItems = [];

  // 루틴 → 요일별 매핑
  routines.forEach(routine => {
    // 간격 루틴 처리
    if (routine.dayOfWeek.indexOf('간격:') === 0) {
      const interval = parseInt(routine.dayOfWeek.split(':')[1]) || 1;
      if (!routine.startDate) return;
      const rStart = new Date(routine.startDate + 'T00:00:00');

      weekDays.forEach(dateStr => {
        if (dateStr < routine.startDate) return;
        const current = new Date(dateStr + 'T00:00:00');
        const diffDays = Math.round((current - rStart) / (1000 * 60 * 60 * 24));
        if (diffDays % interval !== 0) return;

        const exception = events.find(e => e.routineId === routine.id && e.date === dateStr);
        if (exception) {
          weekItems.push({ ...exception, isException: true, originalRoutine: routine });
        } else {
          weekItems.push({
            id: routine.id + '_' + dateStr, title: routine.title, date: dateStr,
            startTime: routine.startTime, endTime: routine.endTime,
            status: 'active', routineId: routine.id, isRoutine: true
          });
        }
      });
      return;
    }

    const days = routine.dayOfWeek.split(',').map(d => d.trim());
    days.forEach(day => {
      const dayIndex = DAY_MAP[day];
      const adjustedIndex = dayIndex === 0 ? 6 : dayIndex - 1;
      const dateForDay = weekDays[adjustedIndex];

      // 시작일 이전 또는 기한 이후의 루틴은 건너뛰기
      if (routine.startDate && dateForDay < routine.startDate) return;
      if (routine.endDate && dateForDay > routine.endDate) return;

      // 이 날짜에 예외가 있는지 확인
      const exception = events.find(e =>
        e.routineId === routine.id && e.date === dateForDay
      );

      if (exception) {
        weekItems.push({
          ...exception,
          isException: true,
          originalRoutine: routine
        });
      } else {
        weekItems.push({
          id: routine.id + '_' + dateForDay,
          title: routine.title,
          date: dateForDay,
          startTime: routine.startTime,
          endTime: routine.endTime,
          status: 'active',
          routineId: routine.id,
          isRoutine: true
        });
      }
    });
  });

  // 순수 개별 일정 추가
  events.filter(e => !e.routineId).forEach(e => {
    weekItems.push({ ...e, isEvent: true });
  });

  // 날짜 + 시간 순 정렬
  weekItems.sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return a.startTime.localeCompare(b.startTime);
  });

  return {
    weekStart: startDate,
    weekEnd: endDate,
    items: weekItems
  };
}

// ============================================================
// 시간 기반 자동 동기화 트리거 설정
// ============================================================

function setupSyncTrigger() {
  // 기존 트리거 제거
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'autoSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 10분마다 실행
  ScriptApp.newTrigger('autoSync')
    .timeBased()
    .everyMinutes(10)
    .create();

  return { success: true, message: '10분 간격 자동 동기화 트리거 설정됨' };
}

function autoSync() {
  try {
    fullSync();
    detectDeletedInstances();
  } catch (e) {
    Logger.log('Auto sync error: ' + e.message);
  }
}

// ============================================================
// 초기 설정: 시트 헤더 생성
// ============================================================

function initializeSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 루틴 시트
  let routineSheet = ss.getSheetByName(SHEET_ROUTINE);
  if (!routineSheet) {
    routineSheet = ss.insertSheet(SHEET_ROUTINE);
  }
  routineSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 루틴');
  routineSheet.getRange(3, 1, 1, 11).setValues([[
    '아이디', '이름', '요일', '시작시간', '종료시간', '캘린더ID', '수정일시', '시작일', '기한', '주소', '메모'
  ]]);

  // 일정 시트
  let eventSheet = ss.getSheetByName(SHEET_EVENT);
  if (!eventSheet) {
    eventSheet = ss.insertSheet(SHEET_EVENT);
  }
  eventSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 일정');
  eventSheet.getRange(3, 1, 1, 12).setValues([[
    '아이디', '이름', '날짜', '시작시간', '종료시간', '캘린더ID', '수정일시', '출처', '루틴ID', '상태', '주소', '메모'
  ]]);

  // 케어 시트
  let careSheet = ss.getSheetByName(SHEET_CARE);
  if (!careSheet) {
    careSheet = ss.insertSheet(SHEET_CARE);
  }
  careSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 케어');
  careSheet.getRange(3, 1, 1, 8).setValues([[
    '아이디', '이름', '주기(일)', '소요기간(일)', '마지막완료일', '예정일', '키워드', '메모'
  ]]);

  // 취향 시트
  let tasteSheet = ss.getSheetByName(SHEET_TASTE);
  if (!tasteSheet) {
    tasteSheet = ss.insertSheet(SHEET_TASTE);
  }
  tasteSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 취향');
  tasteSheet.getRange(3, 1, 1, 7).setValues([[
    '아이디', '카테고리', '제목', '부가정보', '별점', '메모', '추가일시'
  ]]);

  // 유머 시트
  let humorSheet = ss.getSheetByName(SHEET_HUMOR);
  if (!humorSheet) {
    humorSheet = ss.insertSheet(SHEET_HUMOR);
  }
  humorSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 유머');
  humorSheet.getRange(3, 1, 1, 6).setValues([[
    '아이디', '제목', '내용', '태그', '날짜', '추가일시'
  ]]);

  // 할일리스트 시트
  let todoListSheet = ss.getSheetByName(SHEET_TODO_LIST);
  const todoListIsNew = !todoListSheet;
  if (!todoListSheet) {
    todoListSheet = ss.insertSheet(SHEET_TODO_LIST);
  }
  todoListSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 할일리스트');
  todoListSheet.getRange(3, 1, 1, 6).setValues([[
    '아이디', '이름', '순서', '색상', '수정일시', 'GoogleListId'
  ]]);
  // 새로 생성된 경우에만 기본 리스트 추가 (재실행 시 중복 방지)
  if (todoListIsNew) {
    const timestamp = now();
    DEFAULT_TODO_LISTS.forEach((name, idx) => {
      const id = generateId('todolist');
      const googleListId = gtaskCreateList(name);
      todoListSheet.appendRow([
        id, name, idx, TODO_LIST_COLORS[idx % TODO_LIST_COLORS.length], timestamp, googleListId || ''
      ]);
    });
  }

  // 할일 시트
  let todoSheet = ss.getSheetByName(SHEET_TODO);
  if (!todoSheet) {
    todoSheet = ss.insertSheet(SHEET_TODO);
  }
  todoSheet.getRange(1, 1).setValue('한얼 스케줄 관리 시스템 - 할일');
  todoSheet.getRange(3, 1, 1, 8).setValues([[
    '아이디', '리스트아이디', '부모아이디', '텍스트', '완료여부', '순서', '수정일시', 'GoogleTaskId'
  ]]);

  return { success: true, message: '시트 초기화 완료' };
}