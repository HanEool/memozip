# 한얼 스케줄 관리 시스템

## 아키텍처

```
GitHub Pages (index.html)  ←→  Google Apps Script (API)  ←→  Google Sheets (저장소)
                                        ↕
                                Google Calendar (양방향 동기화)
```

- **프론트엔드**: 단일 HTML (vanilla JS), GitHub Pages에서 서빙
- **백엔드**: Google Apps Script 웹앱, 모든 요청 GET 방식
- **저장소**: Google Sheets (루틴 시트 + 일정 시트)
- **캘린더**: 양방향 동기화 (루틴 → recurring event, 일정 → single event)

## 핵심 ID/URL

- **스프레드시트 ID**: `1hMfhClmRJ5edl-fmtwWytk5B8eFWMyvsKx8WkbSIqiE`
- **Apps Script 웹앱 URL**: `https://script.google.com/a/macros/somaandbody.com/s/AKfycbz23iMsMa7QiETq50kVDTj8KUtRoSzBqsZZhkDCqC7ZZmlyFjv5IqJ6E_KAKyZT3IFI/exec`
- **GitHub Pages URL**: `https://haneool.github.io/memozip/` (2026-05-25 기준 — repo 이름 `haneol-schedule` → `memozip` 변경됨)
- **GitHub 저장소**: `https://github.com/HanEool/memozip`
- **일정 캘린더 ID**: `haneol@somaandbody.com`
- **루틴 캘린더 ID**: `c_177b6987863f353fcd46f459f0c5f7f30cd6d1bb3ee3e727fbfdf7dff2185a2d@group.calendar.google.com`

## Google Sheets 구조

**공통 규칙**: Row 1 = 시스템 정보, Row 2 = 공백, Row 3 = 헤더, Row 4~ = 데이터

### 루틴 시트 (11열)
아이디 | 이름 | 요일 | 시작시간 | 종료시간 | 캘린더ID | 수정일시 | 시작일 | 기한 | 주소 | 메모

- 요일: "월,화,목,금" 형태 (쉼표 구분)
- 시작시간/종료시간: "HH:mm" 문자열
- 시작일/기한: "yyyy-MM-dd" 또는 빈 값 (선택)
- Sheets가 시간을 Date 객체로 반환하므로 `formatTime()`, `formatDateVal()` 헬퍼로 변환 필수

### 일정 시트 (12열)
아이디 | 이름 | 날짜 | 시작시간 | 종료시간 | 캘린더ID | 수정일시 | 출처 | 루틴ID | 상태 | 주소 | 메모

- 출처: 'sheets' 또는 'calendar'
- 루틴ID: 빈 값이면 순수 일정, 값이 있으면 루틴에서 파생된 예외
- 상태: 'active', 'cancelled', 'modified'

### 케어 시트 (8열)
아이디 | 이름 | 주기(일) | 소요기간(일) | 마지막완료일 | 예정일 | 키워드 | 메모

- 주기(일): 숫자 (예: 365, 180)
- 소요기간(일): 숫자 (기본값 1, 예: 30이면 30일간 진행)
- 마지막완료일: "yyyy-MM-dd" 또는 빈 값
- 예정일: "yyyy-MM-dd" 또는 빈 값
- 키워드: 쉼표 구분 검색어 (일정 매칭용)

### 할일리스트 시트 (6열)
아이디 | 이름 | 순서 | 색상 | 수정일시 | GoogleListId

- 색상: mint/sky/rose/gold/accent/lavender 중 하나 (UI 카드 색상)
- GoogleListId: Google Tasks API의 task list ID (Phase 2 sync 후 채워짐)
- 기본 리스트 3개 자동 생성: 짓기, 살기, 아이디어

### 할일 시트 (8열)
아이디 | 리스트아이디 | 부모아이디 | 텍스트 | 완료여부 | 순서 | 수정일시 | GoogleTaskId

- 부모아이디: 비어있으면 top-level, 값 있으면 하위 항목 (2단계까지)
- 완료여부: boolean (체크박스 토글)
- GoogleTaskId: Google Tasks API의 task ID

## API 엔드포인트 (모두 GET)

모든 쓰기 작업도 GET으로 처리 (Apps Script POST의 CORS 문제 회피).
데이터는 `data` 파라미터에 JSON 문자열로 전달.

### 읽기
- `?action=getRoutines`
- `?action=getEvents&startDate=yyyy-MM-dd&endDate=yyyy-MM-dd`
- `?action=getWeekView&date=yyyy-MM-dd`
- `?action=getCares`

### 루틴
- `?action=addRoutine&data={JSON}`
- `?action=updateRoutine&data={JSON}`
- `?action=deleteRoutine&id=routine_xxx`

### 일정
- `?action=addEvent&data={JSON}`
- `?action=updateEvent&data={JSON}`
- `?action=deleteEvent&id=event_xxx`

### 케어
- `?action=getCares`
- `?action=addCare&data={JSON}`
- `?action=updateCare&data={JSON}`
- `?action=deleteCare&id=care_xxx`

### 할일 리스트 (카드)
- `?action=getTodoLists`
- `?action=addTodoList&data={JSON}`
- `?action=updateTodoList&data={JSON}`
- `?action=deleteTodoList&id=todolist_xxx` (해당 리스트의 모든 항목까지 함께 삭제)

### 할일 항목
- `?action=getTodos[&listId=todolist_xxx]`
- `?action=addTodo&data={JSON}` (`{listId, parentId, text, done, order}`)
- `?action=updateTodo&data={JSON}` (부분 업데이트 — 보낸 필드만 갱신)
- `?action=deleteTodo&id=todo_xxx` (자식 항목까지 함께 삭제)

### 동기화
- `?action=syncCalendar` (캘린더 양방향)
- `?action=syncTodosToGoogleTasks` (Sheets → Google Tasks 일괄 push, 멱등)

## 파일 구조

```
schedule/
├── index.html        ← GitHub Pages 메인 (PWA 메타태그 포함)
├── manifest.json     ← PWA manifest
├── app-icon.png      ← 180x180 앱 아이콘
└── CLAUDE.md         ← 이 파일
```

**Apps Script 측 (별도 관리)**:
- `Code.gs` — 백엔드 전체 로직
- `Schedule.html` — Apps Script에서 직접 서빙하는 HTML (GitHub Pages와 별개)

## 디자인 시스템

Cold pastel 기반, dark/bright 모드 전환 지원.

### 색상 팔레트
- accent (라벤더): bright `#9b8abf` / dark `#b0a0d0`
- mint (일정): bright `#7cb5a0` / dark `#8ccaae`
- sky (루틴): bright `#88aec8` / dark `#94b8d4`
- rose (취소): bright `#c4929c` / dark `#d4a4ae`
- gold (변경): bright `#c4b07a` / dark `#d4c48e`

### 타이포그래피
- 제목: Nanum Myeongjo (serif)
- 본문: Noto Sans KR (sans-serif)

### 배지 규칙
- 루틴 파생 항목 → 항상 "루틴" 배지 (변경되어도 동일)
- 순수 개별 일정 → "일정" 배지
- 취소된 항목만 → "취소" 배지

## 양방향 동기화 로직

### Sheets → Calendar
- calendarEventId가 없는 행을 감지해서 Calendar에 생성
- 루틴 → recurring event (루틴 캘린더), 일정 → single event (일정 캘린더)

### Calendar → Sheets
- `updatedMin`으로 마지막 동기화 이후 변경분만 폴링
- 충돌 시 lastModified 타임스탬프 비교, 더 최근 것 우선
- 루틴의 개별 인스턴스 수정 → 일정 시트에 예외(routineId 참조)로 기록
- 루틴의 개별 인스턴스 삭제 → 일정 시트에 status='cancelled'로 기록
- 삭제 감지는 Calendar Advanced Service 필요 (`detectDeletedInstances()`)

### Sheets → Google Tasks (할일 단방향 sync, 2026-05-25 추가)

**구조**: Sheets가 원본(source of truth) → addTodo/updateTodo/deleteTodo 등 CRUD 시 best-effort로 Google Tasks API 호출. Tasks API 실패해도 Sheets 동작은 정상. Google Tasks → Sheets 역방향은 **안 함**.

**Tasks Advanced Service**: `appsscript.json`의 `dependencies.enabledAdvancedServices`에 명시. `Tasks.Tasklists`, `Tasks.Tasks` 객체 사용.

**최초 1회 권한 승인** (글로벌 CLAUDE.md "MailApp 권한 승인 흐름"과 동일 패턴):
1. Apps Script 에디터에서 함수 드롭다운 → `testTasksAuth` 선택 → ▷ 실행
2. "권한 검토" → 계정 선택 → "안전하지 않습니다" → "고급" → "{프로젝트명}(으)로 이동(안전하지 않음)" → "허용"
3. 로그에 `Tasks API 연결 OK. 현재 Google Tasks 리스트 수: N` 보이면 성공

**기존 데이터 일괄 push**: Phase 1에서 만들어진 짓기/살기/아이디어 리스트는 `googleListId` 빈 상태. 한얼이 에디터에서 `syncAllTodosToGoogleTasks()` 1회 실행하면 일괄 push. 멱등 (이미 매핑된 건 skip).

**자동 sync 동작**: `addTodoList`/`addTodo` 등 호출 시 내부적으로 `gtaskCreateList`/`gtaskCreate` 등 호출. `googleListId`/`googleTaskId` 채워지면 이후 update/delete도 자동 동기. 권한 없으면 try/catch에서 잡고 빈 ID로 진행 (Sheets는 정상 동작).

## 주의사항

- 한국어 입력: React에서는 ref 기반 uncontrolled input 필요 (이 프로젝트는 vanilla JS라 해당 없음)
- 날짜/시간: UTC 변환 없이 로컬 시간 기준 (Asia/Seoul), `YYYY-MM-DDTHH:mm` 포맷
- Apps Script 배포: 코드 변경 시 "새 버전"으로 배포해야 반영됨
- Apps Script 접근 권한: "모든 사용자"로 설정 필요
- 폼 입력 font-size: iOS 자동 확대 방지를 위해 16px
- 저장 버튼: `isSaving` 플래그로 더블 클릭 방지
- Apps Script `initializeSheets()`: 헤더 컬럼 변경 시 실행 필요

## 현재 루틴 목록

| 이름 | 요일 | 시간 |
|------|------|------|
| 수어교실 | 화 | 19:30-21:30 |
| 화방 | 수 | 18:00-20:30 |
| 교보문고 | 월화수목금 | 09:30-12:30 |
| 스타벅스 | 월화수목금 | 07:00-09:00 |
| 힘의집 운동 | 월화목금 | 18:00-19:30 |
| 펠든크라이스 수련 | 월화목금 | 22:00-23:30 |
