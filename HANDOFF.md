# Sorter Data SQL Runner — 작업 인수인계 (v32)

> 이 문서는 Claude Code에서 작업을 이어가기 위한 컨텍스트 문서다.
> 현재 파일: `sql_runner_v32.py` (Tkinter + SQLAlchemy + pandas 단일 파일 GUI)

---

## 1. 프로그램 개요

반도체/태양전지 **소터(분류) 설비 생산 데이터**를 SQL Server에서 조회해 이상 항목을 패널별로 보여주는 데스크톱 GUI.

- **언어/스택**: Python 3, Tkinter(ttk), SQLAlchemy + pyodbc, pandas, openpyxl
- **실행**: `python sql_runner_v32.py` (실행 시 창 최대화)
- **DB**: SQL Server 2개 (`DCOLDB`=JC02, `DCOLDBJC1`=JC01), 읽기 전용 계정
- **핵심 흐름**:
  1. 사용자가 기간 + Site + Machine + Equipment 필터 선택 후 실행
  2. 백그라운드 스레드에서 SQL Server에 **temp 테이블**(`#BaseSorterResult`, `#LabelMissing` 등) 생성
  3. 5개 패널 SQL을 순차 실행해 각 이상 항목 조회
  4. 좌측 탭으로 패널 전환, Excel(.xlsx) 저장 지원

---

## 2. 화면 구조 (build_ui)

```
┌─────────────────────────────────────────────────────┐
│ [기간 시작][기간 종료]  [실행] [엑셀저장]   상태/시간   │  ← 상단 컨트롤
├─────────────────────────────────────────────────────┤
│ ┌Site─┐  ┌──Equipment(32개 체크박스)──┐ ┌─결과필터─┐ │  ← 필터 바 (1행)
│ │JC01 │  │전체 [JC01-#01][BG21]...    │ │전체/해제  │ │
│ │JC02 │  │기본                        │ │체크박스   │ │
│ ├Machine┤ │해제                       │ │          │ │
│ │NoATW │  └───────────────────────────┘ └──────────┘ │
│ │ATW   │                                              │
│ └JC2.. ┘                                              │
├──────────┬──────────────────────────────────────────┤
│ 조회 항목 │                                          │
│ ┌──────┐ │         (선택된 탭의 결과 Treeview)        │  ← 좌측 5탭 + 우측 패널
│ │혼입.. │ │                                          │
│ │wafer..│ │                                          │
│ │bincnt │ │                                          │
│ │0/null │ │                                          │
│ │30시간 │ │                                          │
│ └──────┘ │                                          │
├─────────────────────────────────────────────────────┤
│ 단계별 시간: Base temp 63.6s | └INSERT JC02 48.2s |..│  ← 하단 타이밍 표시줄
└─────────────────────────────────────────────────────┘
```

- 좌측 탭 리스트(`Listbox`) 클릭 → `_on_tab_select` → 해당 패널만 `show()`, 나머지 `hide()` (grid_remove 방식)
- 각 탭 아래에 행 수/상태 레이블 (`tab_status_labels`)
- **하단 로그 영역은 제거됨**. `self.log_text`는 코드 호환성 위해 숨겨진 채 유지.

---

## 3. 5개 패널 (PANELS)

| key | 한글 제목 | 의미 |
|---|---|---|
| `LABEL_MISSING` | 혼입으로 인한 Label 미발행 확인 | 한 Lot에 ClassGroup이 2개 이상 혼입 or Article mismatch |
| `WAFER_DUP` | waferID 중복 | 동일 WaferID가 2번 이상 |
| `BINCOUNTER_GAP` | bincounter 누락으로 인한 label 미발행 | Bin별 BinCounter 시퀀스에 결번 |
| `CLASS_0_120` | 0 or null | Class/Bin/BinCounter가 0 또는 null |
| `OVER_30H` | 30시간 넘어서 배출 된 소박스 | Lot+Bin의 첫~마지막 배출 간격 ≥ 1801분(30시간) |

`PANELS` 튜플의 `(row, col, colspan)` 값은 **현재 미사용** (과거 grid 배치 잔재). 탭 방식으로 전환되며 의미 없어짐 — 정리해도 됨.

---

## 4. Equipment 라벨 체계 (중요)

설비 식별자는 3가지 표기가 있다:

| 단계 | 예시 | 용도 |
|---|---|---|
| **EquipmentID** | `JC01-CS-01`, `JC02-CS-61` | DB 원본 값, SQL `IN` 절에 전달 |
| **2차 변경 라벨** | `JC01-#01`, `BG21`, `HM31` | **현재 UI 체크박스 표시값** |
| (1차 변경) | `BG01`, `HM21` 등 | 폐기됨 — 더 이상 사용 안 함 |

### 매핑 규칙 (EQUIPMENT_ROWS 테이블 기준)
- **Port 2개 장비**(CS-01~05, CS-01~07, CS-10, CS-61, CS-71): 같은 EquipmentID에 Port1/Port2 → **하나의 `JC0X-#YY` 라벨로 묶음**
  - 예: `JC01-CS-01`(Port1,2) → `JC01-#01` 체크박스 1개
- **Port 1개 장비**(ATW 계열 등): 기존 `BGxx`/`HMxx` 라벨 유지
  - 예: `JC01-CS-21` → `BG21`, `JC02-CS-31` → `HM31`

### 체크박스 ↔ SQL 변환 (핵심 구조)
- `equipment_vars`: **라벨이 키** (BooleanVar), 총 32개
- `label_to_eq`: `라벨 → {EquipmentID 집합}` (Port 2개면 EquipmentID 1개지만 set으로 보관)
- `selected_config()`에서 선택 라벨 → EquipmentID 집합 flatten → SQL `IN` 절 전달
- `eq_label_map`: `EquipmentID → 라벨` (결과 필터에서 역방향 표시용, Port1 우선)
- `label_to_machine`: `라벨 → Machine` (visible_equipment 필터용)

> ⚠️ 라벨/EquipmentID 변환은 4개 dict가 얽혀 있으니, EQUIPMENT_ROWS를 바꾸면 `__init__`의 dict 생성부도 같이 검증할 것.

### EQUIPMENT_ROWS 구조
```python
EQUIPMENT_ROWS = [
    # (EquipmentID, PortID, Machine, 2차변경라벨)
    ("JC01-CS-01", 1, "NoATW", "JC01-#01"), ("JC01-CS-01", 2, "NoATW", "JC01-#01"),
    ...
    ("JC01-CS-21", 1, "ATW", "BG21"),       # 단일 포트
    ...
]
```
- 총 라벨 32개 (BG/JC01 계열 20개 중 묶음 적용, HM/JC02 계열)
- `DEFAULT_EQUIPMENT`: 기본 체크되는 EquipmentID 집합 (ATW 계열 위주)
- `MACHINE_TO_EQUIPMENT`: Machine별 EquipmentID 집합 (필터 표시용)

---

## 5. SQL 파이프라인

### setup_temp_tables(conn, start, end, equipment_ids) → (count_df, subs)
순서대로 실행:
1. **DROP_CREATE_SQL**: `#BaseSorterResult`, `#BaseSorterLabel` 테이블 생성
2. **INSERT_JC02_SQL**: `DCOLDB.dbo.CT_SORTERRESULT`(SiteID='9012C')에서 기간+장비 필터로 INSERT
3. **INSERT_JC01_SQL**: `DCOLDBJC1.dbo.CT_SORTERRESULT`(SiteID='9011')에서 INSERT (JC02 SQL을 .replace로 파생)
4. **POST_BASE_SQL**: 인덱스 생성 → `#LotKeys` → `#BaseSorterLabel`(CT_SORTERLABEL 조인) → `#LabelMissing`(SELECT INTO) → 인덱스
5. **COUNT_SQL**: 행 수 집계

`subs` = `[(step명, 초, rows), ...]` — **v32에서 추가한 하위 단계별 타이밍** (병목 진단용).

### fetch_panel(conn, key) → DataFrame
`PANEL_SQL[key]` 실행. 모든 패널이 `#BaseSorterResult` 또는 `#LabelMissing`을 읽음.

### #LabelMissing 빌드 로직 (POST_BASE_SQL 핵심)
- `Class`/`ArtikelNummer` 문자열을 정규화(UPPER/TRIM/특수문자 제거)
- ClassGroup 대분류 분류: `A-2`, `A-1`, `U-L`, `GA`, `B0`, `EL`, `L-E`, `remeasure`, `RWM`, `Other`
- `MixClassGroup`: 혼입 판정용 (B0/EL/L-E/remeasure/RWM은 NULL 처리)
- `ArticleMismatch`: ArtikelNummer 대분류 ≠ Class 대분류면 1
- **v31에서 최적화**: 3중 CROSS APPLY → 단일 OUTER APPLY, REVERSE 제거, CASE 중복 통합

---

## 6. 데이터 타입 처리 (fix_types)

`fix_types(df)`:
- `TEXT_ID_COLUMNS`(`LotCounter`, `LamaID`, `WaferID`) → 문자열 ID로 정규화 (지수표기/`.0` 방지)
- `Counter`/`Count`/`HOUR`/`시간`/`분`/`Row_No` 포함 컬럼 → 정수면 `Int64`로
- **v30에서 최적화**: 컬럼 사전 분류 + `dropna()` 중복 호출 제거
- **v30**: `finish_run`의 중복 `fix_types` 호출 제거됨

`display_df(raw_df)`: `HIDDEN_COLUMNS` 제거 + fix_types + 행 수 제한(MAX_DISPLAY_ROWS=50000)

---

## 7. ⚠️ 열린 이슈 — base temp 성능 (최우선)

### 증상
`base temp` 단계 소요시간이 **23초 ~ 63초로 크게 출렁임**. 같은 쿼리/같은 필터인데 실행마다 3배 차이.

### 진단 (현재까지의 판단)
- 코드 로직 문제가 아님 (코드는 매번 동일).
- **SQL Server 측 I/O/부하 문제로 추정** — `CT_SORTERRESULT`는 분류 설비가 실시간 INSERT 중인 생산 원본 테이블. 다른 세션 부하·버퍼 캐시 상태에 따라 읽기 속도 변동.

### v32에서 한 일
`base temp`를 4개 하위 단계로 쪼개 타이밍 측정 (하단 표시줄에 표시):
```
└ DROP/CREATE
└ INSERT JC02(원본읽기)
└ INSERT JC01(원본읽기)
└ POST(라벨조인+#LabelMissing)
```

### 다음 할 일 (Claude Code에서)
**v32를 2~3회 실행**해서 어느 하위 단계가 흔들리는지 확인 후:

| 범인 단계 | 의미 | 대응 |
|---|---|---|
| **INSERT JC01/JC02** | 원본 테이블 읽기 I/O = 서버 부하 | **코드로 해결 불가**. `CT_SORTERRESULT`에 `[Date]`(+SiteID+EquipmentID) 인덱스 존재 여부 DBA 확인. 없으면 매번 풀스캔 |
| **POST** | `#LabelMissing` 빌드 = CPU | SQL 추가 최적화 가능 |
| **rows 수 변동** | 데이터 양 자체 변동 | 정상 (양에 비례) |

> 핵심 질문: **`CT_SORTERRESULT`에 Date 기반 인덱스가 있는가?** 없으면 인덱스 추가가 근본 해결책 (DBA 협의 필요).

---

## 8. ❌ 시도했다가 되돌린 것 (반복 금지)

### 패널 병렬화 (v28) — 실패, 되돌림
- 5개 패널 SQL을 `ThreadPoolExecutor`로 병렬 실행 시도
- local temp(`#`)를 global temp(`##`)로 바꿔 다중 connection에서 공유하려 함
- **결과: 데드락(lock 블로킹)으로 5분+ hang**
  - 원인: setup connection의 트랜잭션이 global temp에 락을 쥔 채 미커밋 → reader connection들이 무한 대기
- **결론**: 패널 쿼리는 애초에 병목이 아님(setup이 시간 대부분 차지). 병렬화 이득 없고 위험만 큼 → **sequential 단일 connection 유지가 정답.**

> 다시 병렬화를 시도하지 말 것. 진짜 병목은 setup의 원본 테이블 INSERT I/O다.

---

## 9. 코드 맵 (주요 함수 위치)

```
모듈 상수
  DB, PANELS, EQUIPMENT_ROWS, DEFAULT_EQUIPMENT, MACHINE_TO_EQUIPMENT
  DROP_CREATE_SQL, POST_BASE_SQL, COUNT_SQL, PANEL_SQL(5개)

유틸 함수
  build_engine()              SQLAlchemy 엔진 (pyodbc)
  mapping_df()                EQUIPMENT_ROWS → DataFrame
  machine_of_equipment()      EquipmentID → Machine
  split_by_site()             EquipmentID 목록 → (jc01, jc02)
  setup_temp_tables()         ★ temp 생성 + 하위타이밍 반환
  fetch_panel()               패널 SQL 실행
  normalize_id(), fix_types(), display_df(), filter_by_equipment()

class RunConfig               실행 파라미터 (start/end/sites/machines/equipment)

class ResultPanel             패널 1개 = Treeview + 스크롤 + show/hide
  load(), load_error(), autofit(), select_all(), copy_selection()

class SQLRunnerApp            메인 앱
  build_ui()                  ★ 전체 레이아웃
  build_filter_bar()          ★ 필터 바 (Site/Machine/Equipment/결과필터)
  visible_equipment()         라벨 필터링 (site+machine 기준)
  selected_config()           ★ 선택 라벨 → EquipmentID 집합 → RunConfig
  rebuild_result_filter()     결과 필터 체크박스 재생성
  apply_result_filter()       결과 필터 적용 → 패널 갱신
  _on_tab_select()            ★ 좌측 탭 클릭 → 패널 전환
  run_async() / run_worker()  ★ 백그라운드 실행 (sequential)
  apply_panel_now()           패널 즉시 갱신 (실행 중)
  finish_run()                완료 처리 + 타이밍 표시줄 채움
  export_xlsx()               Excel 저장
```

★ = 최근 작업으로 자주 건드린 함수

---

## 10. 버전 이력 요약

| 버전 | 변경 |
|---|---|
| v17 | 시작점 (업로드 원본) |
| v18 | UI 개편: 필터 컴팩트화 + 좌측 5탭 구조 |
| v19~v20 | 글자 크기 9pt, Machine을 Site 아래로, 결과필터 BG/HM 표기 |
| v21 | 전체/기본/해제 버튼 세로 배치, 결과필터 우측 이동 |
| v22~v23 | Equipment 라벨 BG/HM 공식화, HM99 등 항목 추가 |
| v24 | Port 방식 장비 라벨 누락 수정 (라벨 단위 체크박스로 전환) |
| v25 | **2차 변경명 적용** (`JC01-#XX`, 32개 라벨) |
| v26~v27 | OVER_30H 컬럼 순서/정렬 변경 |
| v28 | ❌ 패널 병렬화 (hang) → 되돌림 |
| v29 | sequential 복귀 + fix_types 벡터화 |
| v30 | 타이밍 표시줄 추가, fix_types 중복 제거, 타이틀 업데이트 |
| v31 | POST_BASE_SQL 최적화 (CROSS APPLY 평탄화 등) |
| **v32** | **base temp 하위 단계별 타이밍 측정 추가** ← 현재 |

---

## 11. 작업 규칙 (사용자 선호)

- 결론 먼저, 근거는 2~3가지로 간결하게
- 숫자/데이터 포함 설명 선호
- 파일 수정 시 **버전 번호 올리기** (`v32` → `v33`), 창 타이틀도 같이 업데이트
- 코드 수정 후 항상 구문 검증 (`python -m py_compile` 또는 `ast.parse`)
- LaTeX는 복잡한 수식에만, 일반 텍스트엔 쓰지 말 것
