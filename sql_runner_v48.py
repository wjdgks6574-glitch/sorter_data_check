# -*- coding: utf-8 -*-
"""
Sorter Data SQL Runner v48
- 소터(분류) 설비 생산 데이터를 SQL Server에서 조회해 이상 항목을 패널별로 표시
- UI: 상단 컴트롤 + 컴팩트 필터 바(Site/Machine/Equipment/결과필터)
      + 좌측 5탭 리스트 + 우측 단일 결과 패널
- 파이프라인: SQL Server #temp 테이블(#BaseSorterResult, #LabelMissing 등) 생성 후
      5개 패널 SQL 순차 조회 (원천 DataFrame 다운로드 없음)
- Excel(.xlsx) 저장, Ctrl+A/C 복사, 최대화 실행 유지
- 하단 단계별 소요시간 표시(병목 진단용)
"""

import csv
import io
import re
import threading
import time
import traceback
from dataclasses import dataclass
from datetime import datetime, timedelta
from tkinter import (
    Tk, Frame, Label, Entry, Button, Checkbutton, StringVar, BooleanVar,
    Text, Scrollbar, Listbox, RIGHT, Y, BOTH, X, END, LEFT, TOP, BOTTOM, N, S, E, W, SINGLE
)
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
from urllib.parse import quote_plus

if __name__ == "__main__":
    try:
        import ctypes
        _console_hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if _console_hwnd:
            ctypes.windll.user32.ShowWindow(_console_hwnd, 0)
    except Exception:
        pass

import pandas as pd

try:
    from sqlalchemy import create_engine, text, bindparam
    SQLALCHEMY_AVAILABLE = True
except Exception:
    SQLALCHEMY_AVAILABLE = False


# =============================================================================
# 기본 설정
# =============================================================================

DB = {
    "server": "172.23.10.58",
    "database": "DCOLDB",
    "user": "rouser",
    "password": "rodcol@12!",
    "driver": "ODBC Driver 17 for SQL Server",
}

PANELS = [
    ("LABEL_MISSING", "Label 미발행 Lot 이상 감지 (혼입·품번불일치)"),
    ("LABEL_ISSUED_ANOMALY", "label 발행 이상"),
    ("WAFER_DUP", "waferID 중복"),
    ("BINCOUNTER_GAP", "bincounter 누락으로 인한 label 미발행"),
    ("CLASS_0_120", "0 or null"),
    ("OVER_30H", "30시간 넘어서 배출 된 소박스"),
]
PANEL_KEYS = [p[0] for p in PANELS]
PANEL_TITLE = {p[0]: p[1] for p in PANELS}

SITE_OPTIONS = ["JC01", "JC02"]
MACHINE_OPTIONS = ["NoATW", "ATW", "JC2 6,7line"]
SMALL_FONT = ("맑은 고딕", 9)
SMALL_FONT_BOLD = ("맑은 고딕", 9, "bold")
MAX_DISPLAY_ROWS = 50000
TEXT_ID_COLUMNS = {"LotCounter", "LamaID", "WaferID"}
HIDDEN_COLUMNS = {"SourceDB", "Site", "SiteID", "EquipmentID", "PortID", "Machine", "Sorter"}

EQUIPMENT_ROWS = [
    # BG (JC01) - NoATW: Port 방식 → JC01-#XX 묶음 표시
    ("JC01-CS-01", 1, "NoATW", "JC01-#01"), ("JC01-CS-01", 2, "NoATW", "JC01-#01"),
    ("JC01-CS-02", 1, "NoATW", "JC01-#02"), ("JC01-CS-02", 2, "NoATW", "JC01-#02"),
    ("JC01-CS-03", 1, "NoATW", "JC01-#03"), ("JC01-CS-03", 2, "NoATW", "JC01-#03"),
    ("JC01-CS-04", 1, "NoATW", "JC01-#04"), ("JC01-CS-04", 2, "NoATW", "JC01-#04"),
    ("JC01-CS-05", 1, "NoATW", "JC01-#05"), ("JC01-CS-05", 2, "NoATW", "JC01-#05"),
    # BG (JC01) - ATW: 단일 Port → 기존 라벨 유지
    ("JC01-CS-21", 1, "ATW", "BG21"), ("JC01-CS-22", 1, "ATW", "BG22"),
    ("JC01-CS-31", 1, "ATW", "BG31"), ("JC01-CS-32", 1, "ATW", "BG32"),
    ("JC01-CS-41", 1, "ATW", "BG41"), ("JC01-CS-42", 1, "ATW", "BG42"),
    ("JC01-CS-51", 1, "ATW", "BG51"), ("JC01-CS-52", 1, "ATW", "BG52"),
    ("JC01-CS-61", 1, "ATW", "BG61"),
    # BG (JC01) - NoATW 기타
    ("JC01-CS-99", 1, "NoATW", "BG99"),
    # HM (JC02) - NoATW: Port 방식 → JC02-#XX 묶음 표시
    ("JC02-CS-01", 1, "NoATW", "JC02-#01"), ("JC02-CS-01", 2, "NoATW", "JC02-#01"),
    ("JC02-CS-02", 1, "NoATW", "JC02-#02"), ("JC02-CS-02", 2, "NoATW", "JC02-#02"),
    ("JC02-CS-03", 1, "NoATW", "JC02-#03"), ("JC02-CS-03", 2, "NoATW", "JC02-#03"),
    ("JC02-CS-04", 1, "NoATW", "JC02-#04"), ("JC02-CS-04", 2, "NoATW", "JC02-#04"),
    ("JC02-CS-05", 1, "NoATW", "JC02-#05"), ("JC02-CS-05", 2, "NoATW", "JC02-#05"),
    ("JC02-CS-06", 1, "NoATW", "JC02-#06"), ("JC02-CS-06", 2, "NoATW", "JC02-#06"),
    ("JC02-CS-07", 1, "NoATW", "JC02-#07"), ("JC02-CS-07", 2, "NoATW", "JC02-#07"),
    ("JC02-CS-10", 1, "NoATW", "JC02-#10"), ("JC02-CS-10", 2, "NoATW", "JC02-#10"),
    # HM (JC02) - ATW: 단일 Port → 기존 라벨 유지
    ("JC02-CS-31", 1, "ATW", "HM31"), ("JC02-CS-32", 1, "ATW", "HM32"),
    ("JC02-CS-51", 1, "ATW", "HM51"), ("JC02-CS-52", 1, "ATW", "HM52"),
    ("JC02-CS-60", 1, "ATW", "HM60"),
    # HM (JC02) - JC2 6,7line: Port 방식 → JC02-#61 / JC02-#71
    ("JC02-CS-61", 1, "JC2 6,7line", "JC02-#61"), ("JC02-CS-61", 2, "JC2 6,7line", "JC02-#61"),
    ("JC02-CS-71", 1, "JC2 6,7line", "JC02-#71"), ("JC02-CS-71", 2, "JC2 6,7line", "JC02-#71"),
    # HM (JC02) - NoATW 기타: 단일 Port → 기존 라벨 유지
    ("JC02-CS-96", 1, "NoATW", "HM96"),
    ("JC02-CS-99", 1, "NoATW", "HM99"),
]

DEFAULT_EQUIPMENT = {
    "JC01-CS-21", "JC01-CS-22", "JC01-CS-31", "JC01-CS-32", "JC01-CS-41",
    "JC01-CS-42", "JC01-CS-51", "JC01-CS-52", "JC01-CS-61",
    "JC02-CS-31", "JC02-CS-32", "JC02-CS-51", "JC02-CS-52", "JC02-CS-60",
}

MACHINE_TO_EQUIPMENT = {
    "NoATW": {
        "JC01-CS-01", "JC01-CS-02", "JC01-CS-03", "JC01-CS-04", "JC01-CS-05", "JC01-CS-99",
        "JC02-CS-01", "JC02-CS-02", "JC02-CS-03", "JC02-CS-04", "JC02-CS-05",
        "JC02-CS-06", "JC02-CS-07", "JC02-CS-10", "JC02-CS-96", "JC02-CS-99",
    },
    "ATW": {
        "JC01-CS-21", "JC01-CS-22", "JC01-CS-31", "JC01-CS-32", "JC01-CS-41", "JC01-CS-42",
        "JC01-CS-51", "JC01-CS-52", "JC01-CS-61",
        "JC02-CS-31", "JC02-CS-32", "JC02-CS-51", "JC02-CS-52", "JC02-CS-60",
    },
    "JC2 6,7line": {"JC02-CS-61", "JC02-CS-71"},
}


# =============================================================================
# SQL: setup + final-result queries
# =============================================================================

DROP_CREATE_SQL = """
SET NOCOUNT ON;
IF OBJECT_ID('tempdb..#ClassifiedBase') IS NOT NULL DROP TABLE #ClassifiedBase;
IF OBJECT_ID('tempdb..#LabelIssuedAnomaly') IS NOT NULL DROP TABLE #LabelIssuedAnomaly;
IF OBJECT_ID('tempdb..#LabelMissing') IS NOT NULL DROP TABLE #LabelMissing;
IF OBJECT_ID('tempdb..#BaseSorterLabel') IS NOT NULL DROP TABLE #BaseSorterLabel;
IF OBJECT_ID('tempdb..#LotKeys') IS NOT NULL DROP TABLE #LotKeys;
IF OBJECT_ID('tempdb..#BaseSorterResult') IS NOT NULL DROP TABLE #BaseSorterResult;

CREATE TABLE #BaseSorterResult (
    SourceDB varchar(20) NOT NULL,
    Site varchar(10) NOT NULL,
    SiteID varchar(20) NULL,
    [Date] datetime2(0) NULL,
    EquipmentID varchar(50) NULL,
    PortID int NULL,
    LotCounter bigint NULL,
    [Class] varchar(200) NULL,
    BinCounter bigint NULL,
    [Comment] bigint NULL,
    Comment2 varchar(200) NULL,
    ArtikelNummer varchar(100) NULL,
    WaferID varchar(100) NULL,
    [Bin] bigint NULL,
    NCell decimal(18, 6) NULL,
    LamaID varchar(50) NULL
);

CREATE TABLE #BaseSorterLabel (
    SourceDB varchar(20) NOT NULL,
    LotCounter bigint NOT NULL,
    Datum datetime2(0) NULL
);
"""

INSERT_JC02_SQL = """
SET NOCOUNT ON;
INSERT INTO #BaseSorterResult
SELECT
    'DCOLDB', 'JC02', CONVERT(varchar(20), sr.SiteID), sr.[Date], sr.EQUIPMENTID,
    TRY_CONVERT(int, sr.PortID), TRY_CONVERT(bigint, sr.LotCounter), CONVERT(varchar(200), sr.[Class]),
    TRY_CONVERT(bigint, sr.BinCounter), TRY_CONVERT(bigint, sr.[Comment]), CONVERT(varchar(200), sr.Comment2),
    CONVERT(varchar(100), sr.ArtikelNummer), CONVERT(varchar(100), sr.WaferID), TRY_CONVERT(bigint, sr.[Bin]),
    TRY_CONVERT(decimal(18, 6), sr.NCell), CONVERT(varchar(50), sr.LamaID)
FROM DCOLDB.dbo.CT_SORTERRESULT sr WITH (NOLOCK)
WHERE sr.[Date] >= :start_date
  AND sr.[Date] <= :end_date
  AND sr.SiteID = '9012C'
  AND sr.EQUIPMENTID IN :equipment_ids
OPTION (RECOMPILE);
"""

INSERT_JC01_SQL = INSERT_JC02_SQL.replace("'DCOLDB', 'JC02'", "'DCOLDBJC1', 'JC01'").replace(
    "FROM DCOLDB.dbo.CT_SORTERRESULT", "FROM DCOLDBJC1.dbo.CT_SORTERRESULT"
).replace("sr.SiteID = '9012C'", "sr.SiteID = '9011'")

POST_INDEX_SQL = """
SET NOCOUNT ON;
CREATE CLUSTERED INDEX IX_Base_LotBinCounter ON #BaseSorterResult (SourceDB, LotCounter, [Bin], BinCounter);
CREATE NONCLUSTERED INDEX IX_Base_Wafer ON #BaseSorterResult (WaferID) INCLUDE ([Date], EquipmentID, PortID, LotCounter, NCell, LamaID, [Comment]);
"""

POST_LOTKEYS_SQL = """
SET NOCOUNT ON;
SELECT DISTINCT SourceDB, LotCounter INTO #LotKeys
FROM #BaseSorterResult
WHERE LotCounter IS NOT NULL;
CREATE UNIQUE CLUSTERED INDEX IX_LotKeys ON #LotKeys (SourceDB, LotCounter);
"""

POST_LABEL_SQL = """
SET NOCOUNT ON;
INSERT INTO #BaseSorterLabel (SourceDB, LotCounter, Datum)
SELECT SourceDB, LotCounter, MAX(Datum)
FROM (
    SELECT 'DCOLDB' AS SourceDB, k.LotCounter, sl.Datum
    FROM DCOLDB.dbo.CT_SORTERLABEL sl WITH (NOLOCK)
    INNER JOIN #LotKeys k ON k.SourceDB = 'DCOLDB' AND sl.Lot = k.LotCounter
    UNION ALL
    SELECT 'DCOLDBJC1', k.LotCounter, sl.Datum
    FROM DCOLDBJC1.dbo.CT_SORTERLABEL sl WITH (NOLOCK)
    INNER JOIN #LotKeys k ON k.SourceDB = 'DCOLDBJC1' AND sl.Lot = k.LotCounter
) L
GROUP BY SourceDB, LotCounter;
CREATE UNIQUE CLUSTERED INDEX IX_Label_Lot ON #BaseSorterLabel (SourceDB, LotCounter);
"""

POST_CLASSIFY_SQL = """
SET NOCOUNT ON;
-- 분류 계산(OUTER APPLY)을 대상 전체 행에 대해 1회만 수행한 후,
-- 라벨 발행 여부(LabelIssued)로 분리해 #LabelMissing / #LabelIssuedAnomaly 생성
SELECT
    b.SourceDB, b.Site, b.[Date], b.EquipmentID, b.PortID, b.LotCounter,
    b.[Class], b.BinCounter, b.[Comment], b.ArtikelNummer, b.WaferID, b.[Bin],
    c.ClassRaw,
    c.ArtikelCode,
    c.ClassArticleCode,
    c.ClassOnlyGroup AS ClassGroup,
    CASE
        WHEN c.ClassOnlyGroup IN ('B0', 'EL', 'L-E') THEN NULL
        WHEN c.ArtikelCode IN ('REMEASURE', 'RWM')    THEN NULL
        WHEN c.ClassOnlyGroup IS NULL                  THEN NULL
        ELSE c.ClassOnlyGroup
    END AS MixClassGroup,
    CASE
        WHEN c.ClassOnlyGroup IN ('B0', 'EL', 'L-E') THEN 0
        -- REMEASURE/RWM는 등급 코드 자체가 아니므로 코드 비교 대상에서 제외
        -- (섞임 여부는 별도로 Lot 단위 RemeasureMixed 조건에서 판정)
        WHEN c.ArtikelCode IN ('REMEASURE', 'RWM')     THEN 0
        WHEN c.ArtikelGroup  IS NULL
          OR c.ClassOnlyGroup IS NULL                  THEN 0
        WHEN c.ArtikelGroup <> c.ClassOnlyGroup        THEN 1
        ELSE 0
    END AS ArticleMismatch,
    CASE WHEN lbl.LotCounter IS NULL THEN 0 ELSE 1 END AS LabelIssued,
    lbl.Datum AS LabelDatum
INTO #ClassifiedBase
FROM (
    SELECT b.*
    FROM #BaseSorterResult b
    WHERE b.[Comment] < 9000000
      AND b.LotCounter IS NOT NULL
) b
LEFT JOIN #BaseSorterLabel lbl
       ON lbl.SourceDB = b.SourceDB AND lbl.LotCounter = b.LotCounter
OUTER APPLY (
    SELECT
        ClassRaw,
        ArtikelRaw,
        NULLIF(REPLACE(REPLACE(REPLACE(ArtikelRaw, '-', ''), ' ', ''), '/', ''), '') AS ArtikelCode,
        -- 마지막 '/' 이후 추출: REVERSE는 위치 계산에만 짧게 사용(문자열 길이 짧아 성능 영향 미미)
        -- 기존 CHARINDEX 이중 계산 방식은 '/'가 3개 이상일 때 두 번째 '/' 기준으로 잘못 추출되는 버그가 있었음
        NULLIF(
            REPLACE(REPLACE(REPLACE(
                CASE
                    WHEN CHARINDEX('/', ClassRaw) = 0 THEN ClassRaw
                    ELSE SUBSTRING(ClassRaw, LEN(ClassRaw) - CHARINDEX('/', REVERSE(ClassRaw)) + 2, LEN(ClassRaw))
                END,
            '-', ''), ' ', ''), '/', ''),
        '') AS ClassArticleCode,
        -- 등급 판정: 대시/공백 제거한 ClassNorm 기준 (A2 == A-2 인식)
        CASE
            WHEN ClassRaw  IS NULL OR ClassRaw = ''   THEN 'class empty'
            WHEN ClassNorm LIKE '%B0'                 THEN 'B0'
            WHEN ClassNorm LIKE '%EL'                 THEN 'EL'
            WHEN ClassNorm LIKE '%LE'                 THEN 'L-E'
            WHEN ClassNorm LIKE '%A2'                 THEN 'A-2'
            WHEN ClassNorm LIKE '%A1'                 THEN 'A-1'
            WHEN ClassNorm LIKE '%UL'                 THEN 'U-L'
            WHEN ClassNorm LIKE '%GA'                 THEN 'GA'
            ELSE 'Other'
        END AS ClassOnlyGroup,
        -- REMEASURE/RWM는 원본 기준, 등급은 ArtikelNorm 기준 (대시/공백 무시)
        CASE
            WHEN NULLIF(REPLACE(REPLACE(REPLACE(ArtikelRaw, '-', ''), ' ', ''), '/', ''), '') IS NULL THEN NULL
            WHEN ArtikelRaw  LIKE '%REMEASURE%' THEN 'remeasure'
            WHEN ArtikelRaw  LIKE '%RWM%'       THEN 'RWM'
            WHEN ArtikelNorm LIKE '%B0'         THEN 'B0'
            WHEN ArtikelNorm LIKE '%EL'         THEN 'EL'
            WHEN ArtikelNorm LIKE '%LE'         THEN 'L-E'
            WHEN ArtikelNorm LIKE '%A2'         THEN 'A-2'
            WHEN ArtikelNorm LIKE '%A1'         THEN 'A-1'
            WHEN ArtikelNorm LIKE '%UL'         THEN 'U-L'
            WHEN ArtikelNorm LIKE '%GA'         THEN 'GA'
            ELSE 'Other'
        END AS ArtikelGroup
    FROM (
        SELECT
            ClassRaw, ArtikelRaw,
            REPLACE(REPLACE(ClassRaw,  '-', ''), ' ', '') AS ClassNorm,
            REPLACE(REPLACE(ArtikelRaw, '-', ''), ' ', '') AS ArtikelNorm
        FROM (
            SELECT
                UPPER(LTRIM(RTRIM(CONVERT(varchar(200), b.[Class])))) AS ClassRaw,
                UPPER(LTRIM(RTRIM(CONVERT(varchar(100), b.ArtikelNummer)))) AS ArtikelRaw
        ) r0
    ) raw
) c
OPTION (RECOMPILE);

CREATE NONCLUSTERED INDEX IX_ClassifiedBase_Label ON #ClassifiedBase (LabelIssued, SourceDB, LotCounter);

SELECT SourceDB, Site, [Date], EquipmentID, PortID, LotCounter, [Class], BinCounter, [Comment],
       ArtikelNummer, WaferID, [Bin], ClassRaw, ArtikelCode, ClassArticleCode, ClassGroup,
       MixClassGroup, ArticleMismatch
INTO #LabelMissing
FROM #ClassifiedBase
WHERE LabelIssued = 0;

CREATE CLUSTERED INDEX IX_LabelMissing_LotBin ON #LabelMissing (SourceDB, LotCounter, [Bin], BinCounter);
CREATE NONCLUSTERED INDEX IX_LabelMissing_Mix ON #LabelMissing (MixClassGroup, SourceDB, LotCounter)
    INCLUDE ([Date], BinCounter, EquipmentID, PortID, ArticleMismatch, ArtikelCode, ClassArticleCode);

SELECT SourceDB, Site, [Date], EquipmentID, PortID, LotCounter, [Class], BinCounter, [Comment],
       ArtikelNummer, WaferID, [Bin], ClassRaw, ArtikelCode, ClassArticleCode, ClassGroup,
       MixClassGroup, ArticleMismatch, LabelDatum
INTO #LabelIssuedAnomaly
FROM #ClassifiedBase
WHERE LabelIssued = 1;

CREATE CLUSTERED INDEX IX_LabelIssuedAnomaly_LotBin ON #LabelIssuedAnomaly (SourceDB, LotCounter, [Bin], BinCounter);
CREATE NONCLUSTERED INDEX IX_LabelIssuedAnomaly_Mix ON #LabelIssuedAnomaly (MixClassGroup, SourceDB, LotCounter)
    INCLUDE ([Date], BinCounter, EquipmentID, PortID, ArticleMismatch, ArtikelCode, ClassArticleCode, LabelDatum);

DROP TABLE #ClassifiedBase;
"""


COUNT_SQL = """
SELECT
    COUNT_BIG(*) AS BaseRows,
    (SELECT COUNT_BIG(*) FROM #LabelMissing) AS LabelMissingRows,
    (SELECT COUNT_BIG(*) FROM #LabelIssuedAnomaly) AS LabelIssuedAnomalyRows
FROM #BaseSorterResult;
"""

PANEL_SQL = {
    "LABEL_MISSING": """
;WITH ClassCounts AS (
    SELECT SourceDB, LotCounter, MixClassGroup AS ClassGroup, COUNT_BIG(*) AS ClassCount
    FROM #LabelMissing
    WHERE MixClassGroup IS NOT NULL
    GROUP BY SourceDB, LotCounter, MixClassGroup
), ClassGroupStats AS (
    SELECT SourceDB, LotCounter, COUNT_BIG(*) AS MixClassGroupCount
    FROM ClassCounts
    GROUP BY SourceDB, LotCounter
), ClassGroupList AS (
    -- Lot별 모든 등급을 "등급:건수" 형태로 나열 (건수 많은 순)
    SELECT cc.SourceDB, cc.LotCounter,
           STUFF((
               SELECT ', ' + cc2.ClassGroup + ':' + CONVERT(varchar(20), cc2.ClassCount)
               FROM ClassCounts cc2
               WHERE cc2.SourceDB = cc.SourceDB AND cc2.LotCounter = cc.LotCounter
               ORDER BY cc2.ClassCount DESC
               FOR XML PATH(''), TYPE
           ).value('.', 'varchar(max)'), 1, 2, '') AS ClassGroups
    FROM ClassCounts cc
    GROUP BY cc.SourceDB, cc.LotCounter
), TotalCounts AS (
    SELECT SourceDB, LotCounter,
           COUNT_BIG(*) AS TotalClassCount,
           COUNT(DISTINCT BinCounter) AS DistinctBinCount,
           MAX(BinCounter) AS MaxBinCounter
    FROM #LabelMissing
    GROUP BY SourceDB, LotCounter
), LatestDates AS (
    SELECT SourceDB, LotCounter, MAX([Date]) AS LatestDate
    FROM #LabelMissing
    GROUP BY SourceDB, LotCounter
), LotInfo AS (
    SELECT SourceDB, LotCounter, MIN(Site) AS Site, MIN(EquipmentID) AS EquipmentID, MIN(PortID) AS PortID
    FROM #LabelMissing
    GROUP BY SourceDB, LotCounter
), MismatchStats AS (
    SELECT SourceDB, LotCounter,
           SUM(CASE WHEN ArticleMismatch = 1 THEN 1 ELSE 0 END) AS MismatchCount,
           MIN(CASE WHEN ArticleMismatch = 1
                    THEN CONCAT('BinCounter=', CONVERT(varchar(30), BinCounter), ' ', ArtikelCode, '<>', ClassArticleCode)
               END) AS MismatchExample
    FROM #LabelMissing
    GROUP BY SourceDB, LotCounter
), RemeasureStats AS (
    -- REMEASURE/RWM 행과 그 외(정상 측정) 행이 같은 Lot에 공존하면 혼입으로 판정
    -- (Lot 전체가 REMEASURE/RWM뿐이면 비교 대상이 없어 정상 처리)
    SELECT SourceDB, LotCounter,
           SUM(CASE WHEN ArtikelCode IN ('REMEASURE', 'RWM') THEN 1 ELSE 0 END) AS RemeasureCount,
           SUM(CASE WHEN ArtikelCode NOT IN ('REMEASURE', 'RWM') OR ArtikelCode IS NULL THEN 1 ELSE 0 END) AS NonRemeasureCount
    FROM #LabelMissing
    GROUP BY SourceDB, LotCounter
)
SELECT li.Site, li.EquipmentID, li.PortID,
       REPLACE(li.EquipmentID, '-CS-', '-') AS Equipment,
       tc.LotCounter,
       CONVERT(varchar(19), ld.LatestDate, 120) AS LatestDate,
       CONVERT(varchar(10), ld.LatestDate, 120) AS WORKDAY,
       DATEPART(HOUR, ld.LatestDate) AS [HOUR],
       tc.MaxBinCounter AS [Bincounter 최대값], tc.TotalClassCount AS [데이터 갯수], tc.DistinctBinCount AS [서로 다른 bincounter 갯수]
FROM TotalCounts tc
LEFT JOIN ClassGroupStats cgs ON cgs.SourceDB = tc.SourceDB AND cgs.LotCounter = tc.LotCounter
LEFT JOIN ClassGroupList cgl ON cgl.SourceDB = tc.SourceDB AND cgl.LotCounter = tc.LotCounter
JOIN LatestDates ld ON ld.SourceDB = tc.SourceDB AND ld.LotCounter = tc.LotCounter
JOIN LotInfo li ON li.SourceDB = tc.SourceDB AND li.LotCounter = tc.LotCounter
JOIN MismatchStats ms ON ms.SourceDB = tc.SourceDB AND ms.LotCounter = tc.LotCounter
JOIN RemeasureStats rs ON rs.SourceDB = tc.SourceDB AND rs.LotCounter = tc.LotCounter
-- LEFT JOIN 사용: 혼입 등급이 0개라도(전부 REMEASURE 등) 품번불일치만으로 걸릴 수 있음
-- 데이터 누락 체크: MaxBinCounter=COUNT(1-based) 또는 MaxBinCounter+1=COUNT(0-based)
-- 둘 다 아니면 결번/초과로 판정 (라벨 미발행 원인 중 하나)
WHERE (ISNULL(cgs.MixClassGroupCount, 0) > 1
    OR ms.MismatchCount > 0
    OR (tc.MaxBinCounter <> tc.TotalClassCount AND tc.MaxBinCounter + 1 <> tc.TotalClassCount)
    OR (rs.RemeasureCount > 0 AND rs.NonRemeasureCount > 0))
ORDER BY MismatchCount DESC, Equipment, tc.LotCounter
OPTION (RECOMPILE);
""",
    "LABEL_ISSUED_ANOMALY": """
;WITH LotFlags AS (
    -- Lot 단위 Klasse/ArtikelNummer 토큰 존재 여부 (전부 대문자, 특수문자 제거 없이 원문 포함 여부로 판정)
    SELECT SourceDB, LotCounter,
           MAX(CASE WHEN ClassRaw LIKE '%A-2%' THEN 1 ELSE 0 END) AS L_A2,
           MAX(CASE WHEN ClassRaw LIKE '%A-1%' THEN 1 ELSE 0 END) AS L_A1,
           MAX(CASE WHEN ClassRaw LIKE '%U-L%' THEN 1 ELSE 0 END) AS L_UL,
           MAX(CASE WHEN ClassRaw LIKE '%B-0%' THEN 1 ELSE 0 END) AS L_B0,
           MAX(CASE WHEN ClassRaw LIKE '%L-E%' THEN 1 ELSE 0 END) AS L_LE,
           MAX(CASE WHEN ClassRaw LIKE '%/EL' THEN 1 ELSE 0 END) AS L_EL,
           MAX(CASE WHEN ClassRaw LIKE 'H-%' AND ArtikelCode = 'RWM' THEN 1 ELSE 0 END) AS L_H_RWM,
           MAX(CASE WHEN ClassRaw LIKE 'H-%' AND ArtikelCode = 'REMEASURE' THEN 1 ELSE 0 END) AS L_H_REMEASURE,
           MAX(CASE WHEN ArtikelCode LIKE '%UL%' THEN 1 ELSE 0 END) AS A_UL,
           MAX(CASE WHEN ArtikelCode LIKE '%A2%' THEN 1 ELSE 0 END) AS A_A2,
           MAX(CASE WHEN ArtikelCode LIKE '%A1%' THEN 1 ELSE 0 END) AS A_A1,
           MAX(CASE WHEN ArtikelCode LIKE '%B0%' THEN 1 ELSE 0 END) AS A_B0,
           MAX(CASE WHEN ArtikelCode LIKE '%RWM%' THEN 1 ELSE 0 END) AS A_RWM,
           MAX(CASE WHEN ArtikelCode LIKE '%LE%' THEN 1 ELSE 0 END) AS A_LE,
           MAX(CASE WHEN ArtikelCode LIKE '%REMEASURE%' THEN 1 ELSE 0 END) AS A_REMEASURE,
           MAX(CASE WHEN ArtikelCode LIKE '%GA%' THEN 1 ELSE 0 END) AS A_GA
    FROM #LabelIssuedAnomaly
    GROUP BY SourceDB, LotCounter
), LotAnomaly AS (
    SELECT SourceDB, LotCounter,
           CASE WHEN
                  (L_A2 = 1 AND (A_UL=1 OR A_A1=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_A1 = 1 AND (A_UL=1 OR A_A2=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_UL = 1 AND (A_A2=1 OR A_A1=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_B0 = 1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_LE = 1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_B0=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_EL = 1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_H_RWM = 1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_B0=1 OR A_REMEASURE=1 OR A_GA=1))
               OR (L_H_REMEASURE = 1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_B0=1 OR A_RWM=1 OR A_GA=1))
                THEN 1 ELSE 0
           END AS RuleViolation,
           -- 규칙 wrapper 없이, 실제로 섞여든 금지 토큰만 "TOKEN/TOKEN 혼입" 형태로 표기
           STUFF(
               CASE WHEN L_A2=1 AND (A_UL=1 OR A_A1=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_B0=1 THEN '/B0' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_LE=1 THEN '/LE' ELSE '' END + CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END +
                        CASE WHEN A_GA=1 THEN '/GA' ELSE '' END, 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_A1=1 AND (A_UL=1 OR A_A2=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_A2=1 THEN '/A2' ELSE '' END +
                        CASE WHEN A_B0=1 THEN '/B0' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_LE=1 THEN '/LE' ELSE '' END + CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END +
                        CASE WHEN A_GA=1 THEN '/GA' ELSE '' END, 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_UL=1 AND (A_A2=1 OR A_A1=1 OR A_B0=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_B0=1 THEN '/B0' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_LE=1 THEN '/LE' ELSE '' END + CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END +
                        CASE WHEN A_GA=1 THEN '/GA' ELSE '' END, 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_B0=1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_LE=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_LE=1 THEN '/LE' ELSE '' END + CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END +
                        CASE WHEN A_GA=1 THEN '/GA' ELSE '' END, 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_LE=1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_B0=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_B0=1 THEN '/B0' ELSE '' END + CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END +
                        CASE WHEN A_GA=1 THEN '/GA' ELSE '' END, 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_EL=1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_RWM=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END +
                        CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END + CASE WHEN A_GA=1 THEN '/GA' ELSE '' END
                    , 1, 1, '') + ' 혼입'
                    ELSE '' END +
               CASE WHEN L_H_RWM=1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_B0=1 OR A_REMEASURE=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_B0=1 THEN '/B0' ELSE '' END +
                        CASE WHEN A_REMEASURE=1 THEN '/REMEASURE' ELSE '' END + CASE WHEN A_GA=1 THEN '/GA' ELSE '' END
                    , 1, 1, '') + ' 혼입(RWM동반)'
                    ELSE '' END +
               CASE WHEN L_H_REMEASURE=1 AND (A_A2=1 OR A_A1=1 OR A_UL=1 OR A_B0=1 OR A_RWM=1 OR A_GA=1)
                    THEN ', ' + STUFF(
                        CASE WHEN A_A2=1 THEN '/A2' ELSE '' END + CASE WHEN A_A1=1 THEN '/A1' ELSE '' END +
                        CASE WHEN A_UL=1 THEN '/UL' ELSE '' END + CASE WHEN A_B0=1 THEN '/B0' ELSE '' END +
                        CASE WHEN A_RWM=1 THEN '/RWM' ELSE '' END + CASE WHEN A_GA=1 THEN '/GA' ELSE '' END
                    , 1, 1, '') + ' 혼입(REMEASURE동반)'
                    ELSE '' END
           , 1, 2, '') AS ViolationDetail
    FROM LotFlags
), QualifyingLots AS (
    -- 규칙 위반(RuleViolation=1) Lot만 대상. A-2/U-L/A-1/B0 등급이라도
    -- 위반이 없으면(예: 2320A2/2340A2처럼 같은 등급 내 하위값 혼재는 정상) 제외.
    -- 위반 Lot 중 A-2/U-L/A-1/B0 타입은 요구사항2에 따라 종류별로 breakdown되어 표시됨.
    SELECT lf.SourceDB, lf.LotCounter, la.RuleViolation, la.ViolationDetail
    FROM LotFlags lf
    JOIN LotAnomaly la ON la.SourceDB = lf.SourceDB AND la.LotCounter = lf.LotCounter
    WHERE la.RuleViolation = 1
), LabelInfo AS (
    SELECT SourceDB, LotCounter, MAX(LabelDatum) AS LabelDatum
    FROM #LabelIssuedAnomaly
    GROUP BY SourceDB, LotCounter
), LotInfo AS (
    SELECT SourceDB, LotCounter, MIN(Site) AS Site, MIN(EquipmentID) AS EquipmentID, MIN(PortID) AS PortID
    FROM #LabelIssuedAnomaly
    GROUP BY SourceDB, LotCounter
), ArtikelStats AS (
    SELECT SourceDB, LotCounter, ArtikelNummer, COUNT_BIG(*) AS ArtikelCount
    FROM #LabelIssuedAnomaly
    GROUP BY SourceDB, LotCounter, ArtikelNummer
), ArtikelBreakdownList AS (
    -- 요구사항 2: Lot당 1행으로 압축, ArtikelNummer 종류별 개수를 "종류:개수" 형태로 한 컴럼에 요약
    SELECT a.SourceDB, a.LotCounter,
           STUFF((
               SELECT ', ' + a2.ArtikelNummer + ':' + CONVERT(varchar(20), a2.ArtikelCount)
               FROM ArtikelStats a2
               WHERE a2.SourceDB = a.SourceDB AND a2.LotCounter = a.LotCounter
               ORDER BY a2.ArtikelCount DESC
               FOR XML PATH(''), TYPE
           ).value('.', 'varchar(max)'), 1, 2, '') AS ArtikelBreakdown
    FROM ArtikelStats a
    GROUP BY a.SourceDB, a.LotCounter
), LotBinStats AS (
    SELECT SourceDB, LotCounter,
           MIN([Bin]) AS BinMin, MAX([Bin]) AS BinMax,
           MIN(BinCounter) AS MinBinCounter, MAX(BinCounter) AS MaxBinCounter,
           MAX([Date]) AS LatestDate
    FROM #LabelIssuedAnomaly
    GROUP BY SourceDB, LotCounter
)
SELECT li.Site, li.EquipmentID, li.PortID,
       REPLACE(li.EquipmentID, '-CS-', '-') AS Equipment,
       ql.LotCounter,
       ab.ArtikelBreakdown,
       CONVERT(varchar(19), lb.LatestDate, 120) AS LatestDate,
       CONVERT(varchar(19), lf.LabelDatum, 120) AS LabelDatum,
       ql.ViolationDetail
FROM QualifyingLots ql
JOIN LotInfo li ON li.SourceDB = ql.SourceDB AND li.LotCounter = ql.LotCounter
JOIN ArtikelBreakdownList ab ON ab.SourceDB = ql.SourceDB AND ab.LotCounter = ql.LotCounter
JOIN LotBinStats lb ON lb.SourceDB = ql.SourceDB AND lb.LotCounter = ql.LotCounter
JOIN LabelInfo lf ON lf.SourceDB = ql.SourceDB AND lf.LotCounter = ql.LotCounter
ORDER BY Equipment, ql.LotCounter
OPTION (RECOMPILE);
""",
    "WAFER_DUP": """
;WITH Cleaned AS (
    SELECT b.[Date], b.Site, b.EquipmentID, b.PortID,
           REPLACE(b.EquipmentID, '-CS-', '-') AS Equipment,
           NULLIF(LTRIM(RTRIM(b.WaferID)), '') AS WaferID,
           b.LotCounter, b.NCell, b.LamaID, b.[Comment]
    FROM #BaseSorterResult b
    WHERE b.WaferID IS NOT NULL
), Dup AS (
    SELECT WaferID
    FROM Cleaned
    WHERE WaferID IS NOT NULL
    GROUP BY WaferID
    HAVING COUNT_BIG(*) > 1
)
SELECT CONVERT(varchar(19), c.[Date], 120) AS [Date],
       CONVERT(varchar(10), c.[Date], 120) AS WORKDAY,
       DATEPART(HOUR, c.[Date]) AS [HOUR],
       c.Site, c.EquipmentID, c.PortID, c.Equipment,
       c.WaferID, c.LotCounter, ROUND(c.NCell, 4) AS NCell, c.LamaID, c.[Comment]
FROM Cleaned c
JOIN Dup d ON d.WaferID = c.WaferID
ORDER BY c.WaferID, c.Equipment, c.[Date]
OPTION (RECOMPILE);
""",
    "BINCOUNTER_GAP": """
;WITH CandidateGroups AS (
    SELECT SourceDB, LotCounter, [Bin]
    FROM #LabelMissing
    WHERE BinCounter IS NOT NULL AND [Bin] IS NOT NULL
    GROUP BY SourceDB, LotCounter, [Bin]
    HAVING COUNT(DISTINCT BinCounter) < MAX(BinCounter) - MIN(BinCounter) + 1
), DistinctBinData AS (
    SELECT *
    FROM (
        SELECT b.SourceDB, b.[Date], b.Site, b.EquipmentID, b.PortID, b.LotCounter, b.[Bin], b.BinCounter, b.WaferID,
               ROW_NUMBER() OVER (PARTITION BY b.SourceDB, b.LotCounter, b.[Bin], b.BinCounter ORDER BY b.[Date], b.WaferID) AS rn
        FROM #BaseSorterResult b
        JOIN CandidateGroups c ON c.SourceDB = b.SourceDB AND c.LotCounter = b.LotCounter AND c.[Bin] = b.[Bin]
        WHERE b.BinCounter IS NOT NULL
    ) x
    WHERE rn = 1
), Windowed AS (
    SELECT d.*,
           LEAD(d.BinCounter) OVER (PARTITION BY d.SourceDB, d.LotCounter, d.[Bin] ORDER BY d.BinCounter) AS NextBinCounter,
           LAG(d.BinCounter) OVER (PARTITION BY d.SourceDB, d.LotCounter, d.[Bin] ORDER BY d.BinCounter) AS PrevBinCounter
    FROM DistinctBinData d
)
SELECT CONVERT(varchar(19), [Date], 120) AS [Date],
       CONVERT(varchar(10), [Date], 120) AS WORKDAY,
       DATEPART(HOUR, [Date]) AS [HOUR],
       Site, EquipmentID, PortID, REPLACE(EquipmentID, '-CS-', '-') AS Equipment,
       LotCounter, [Bin], BinCounter, WaferID
FROM Windowed
WHERE (NextBinCounter IS NOT NULL AND NextBinCounter > BinCounter + 1)
   OR (PrevBinCounter IS NOT NULL AND BinCounter > PrevBinCounter + 1)
ORDER BY LotCounter, [Bin], BinCounter, [Date]
OPTION (RECOMPILE);
""",
    "CLASS_0_120": """
SELECT CONVERT(varchar(19), b.[Date], 120) AS [Date],
       CONVERT(varchar(10), b.[Date], 120) AS WORKDAY,
       DATEPART(HOUR, b.[Date]) AS [HOUR],
       b.Site, b.EquipmentID, b.PortID, REPLACE(b.EquipmentID, '-CS-', '-') AS Equipment,
       b.WaferID, b.LamaID, b.[Bin], b.BinCounter, b.[Class]
FROM #BaseSorterResult b
WHERE b.[Class] IS NULL
   OR LTRIM(RTRIM(b.[Class])) = ''
   OR b.[Bin] IS NULL
   OR b.BinCounter IS NULL
   OR b.[Bin] IN (0, 120)
   OR b.BinCounter = 0
ORDER BY b.EquipmentID, b.[Date]
OPTION (RECOMPILE);
""",
    "OVER_30H": """
;WITH LotTimeDiff AS (
    SELECT SourceDB, LotCounter, [Bin], MIN([Date]) AS MinDate, MAX([Date]) AS MaxDate,
           DATEDIFF(MINUTE, MIN([Date]), MAX([Date])) AS DiffMinutes
    FROM #BaseSorterResult
    WHERE LotCounter IS NOT NULL AND [Bin] IS NOT NULL
    GROUP BY SourceDB, LotCounter, [Bin]
    HAVING DATEDIFF(MINUTE, MIN([Date]), MAX([Date])) >= 1801
), MaxBinInfo AS (
    SELECT *
    FROM (
        SELECT b.SourceDB, b.LotCounter, b.[Bin], b.Site, b.EquipmentID, b.PortID, b.BinCounter, b.[Class],
               ROW_NUMBER() OVER (PARTITION BY b.SourceDB, b.LotCounter, b.[Bin] ORDER BY b.BinCounter DESC, b.[Date] DESC) AS rn
        FROM #BaseSorterResult b
        JOIN LotTimeDiff l ON l.SourceDB = b.SourceDB AND l.LotCounter = b.LotCounter AND l.[Bin] = b.[Bin]
    ) x
    WHERE rn = 1
)
SELECT m.Site, m.EquipmentID, m.PortID, REPLACE(m.EquipmentID, '-CS-', '-') AS Equipment,
       CONVERT(varchar(10), l.MaxDate, 120) AS WORKDAY,
       DATEPART(HOUR, l.MaxDate) AS [HOUR],
       l.LotCounter, l.[Bin], m.BinCounter,
       CONVERT(varchar(19), l.MinDate, 120) AS MinDate,
       CONVERT(varchar(19), l.MaxDate, 120) AS MaxDate,
       l.DiffMinutes / 60 AS [시간], l.DiffMinutes % 60 AS [분],
       m.[Class], ISNULL(CONVERT(varchar(19), bl.Datum, 120), '미발행') AS Datum
FROM LotTimeDiff l
JOIN MaxBinInfo m ON m.SourceDB = l.SourceDB AND m.LotCounter = l.LotCounter AND m.[Bin] = l.[Bin]
LEFT JOIN #BaseSorterLabel bl ON bl.SourceDB = l.SourceDB AND bl.LotCounter = l.LotCounter
ORDER BY
    CASE m.[Class]
        WHEN 'A-2' THEN 1
        WHEN 'U-L' THEN 2
        WHEN 'A-1' THEN 3
        ELSE 4
    END,
    l.DiffMinutes DESC,
    m.EquipmentID
OPTION (RECOMPILE);
""",
}


# =============================================================================
# 유틸
# =============================================================================


def default_datetime_range():
    end = datetime.now().replace(hour=8, minute=0, second=0, microsecond=0)
    start = end - timedelta(days=3)
    return start.strftime("%Y-%m-%d %H:%M:%S"), end.strftime("%Y-%m-%d %H:%M:%S")


def build_engine():
    if not SQLALCHEMY_AVAILABLE:
        raise RuntimeError("필수 라이브러리가 없습니다. pip install pandas sqlalchemy pyodbc openpyxl")
    odbc = (
        f"DRIVER={{{DB['driver']}}};SERVER={DB['server']};DATABASE={DB['database']};"
        f"UID={DB['user']};PWD={DB['password']};TrustServerCertificate=yes;Encrypt=no;"
    )
    return create_engine("mssql+pyodbc:///?odbc_connect=" + quote_plus(odbc), pool_pre_ping=True, fast_executemany=True)


def split_by_site(equipment_ids):
    return (
        sorted([x for x in equipment_ids if x.startswith("JC01-")]),
        sorted([x for x in equipment_ids if x.startswith("JC02-")]),
    )


def setup_temp_tables(conn, start_date, end_date, equipment_ids):
    """base temp 생성. 하위 단계별 소요시간을 함께 반환해 병목 위치를 진단한다.
    반환: (count_df, sub_timings)  sub_timings = [(step, seconds, rows), ...]"""
    jc01, jc02 = split_by_site(equipment_ids)
    subs = []

    t = time.perf_counter()
    conn.exec_driver_sql(DROP_CREATE_SQL)
    subs.append(("  └ DROP/CREATE", time.perf_counter() - t, None))

    if jc02:
        t = time.perf_counter()
        stmt = text(INSERT_JC02_SQL).bindparams(bindparam("equipment_ids", expanding=True))
        res = conn.execute(stmt, {"start_date": start_date, "end_date": end_date, "equipment_ids": jc02})
        subs.append(("  └ INSERT JC02(원본읽기)", time.perf_counter() - t, res.rowcount))

    if jc01:
        t = time.perf_counter()
        stmt = text(INSERT_JC01_SQL).bindparams(bindparam("equipment_ids", expanding=True))
        res = conn.execute(stmt, {"start_date": start_date, "end_date": end_date, "equipment_ids": jc01})
        subs.append(("  └ INSERT JC01(원본읽기)", time.perf_counter() - t, res.rowcount))

    t = time.perf_counter()
    conn.exec_driver_sql(POST_INDEX_SQL)
    subs.append(("  └ POST-인덱스생성", time.perf_counter() - t, None))

    t = time.perf_counter()
    conn.exec_driver_sql(POST_LOTKEYS_SQL)
    subs.append(("  └ POST-LotKeys", time.perf_counter() - t, None))

    t = time.perf_counter()
    conn.exec_driver_sql(POST_LABEL_SQL)
    subs.append(("  └ POST-라벨조인", time.perf_counter() - t, None))

    t = time.perf_counter()
    conn.exec_driver_sql(POST_CLASSIFY_SQL)
    subs.append(("  └ POST-분류(미발행+발행완료)빌드", time.perf_counter() - t, None))

    count_df = pd.read_sql_query(text(COUNT_SQL), conn)
    return count_df, subs


def fetch_panel(conn, key):
    return pd.read_sql_query(text(PANEL_SQL[key]), conn)


def normalize_id(value):
    if value is None or pd.isna(value):
        return ""
    s = str(value).strip()
    if re.fullmatch(r"[-+]?\d+\.0", s):
        return s[:-2]
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return s


def fix_types(df):
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.copy()
    int_keywords = ("Counter", "Count", "HOUR", "시간", "분", "Row_No")
    cols = list(out.columns)
    text_id_cols = [c for c in cols if str(c) in TEXT_ID_COLUMNS]
    text_id_set = set(text_id_cols)
    int_cols = [c for c in cols if c not in text_id_set and any(k in str(c) for k in int_keywords)]
    # Text ID 컬럼 일괄 처리
    for col in text_id_cols:
        out[col] = out[col].map(normalize_id).astype("string")
    # Int 컬럼: dropna 1회만 호출
    for col in int_cols:
        num = pd.to_numeric(out[col], errors="coerce")
        non_null = num.dropna()
        if not non_null.empty and ((non_null % 1) == 0).all():
            out[col] = num.astype("Int64")
    return out


def display_df(raw_df):
    df = fix_types(raw_df)
    if df.empty:
        return df
    drop = [c for c in df.columns if str(c) in HIDDEN_COLUMNS]
    df = df.drop(columns=drop, errors="ignore")
    if "Equipment" in df.columns:
        df = df[["Equipment"] + [c for c in df.columns if c != "Equipment"]]
    return df


def filter_by_equipment(df, selected_equipment):
    if df is None or df.empty or "EquipmentID" not in df.columns:
        return pd.DataFrame() if selected_equipment == [] else df.copy()
    return df[df["EquipmentID"].astype(str).isin(set(selected_equipment))].copy()


@dataclass
class RunConfig:
    start_date: str
    end_date: str
    sites: list
    machines: list
    equipment: list


# =============================================================================
# Tree panel
# =============================================================================


class ResultPanel:
    def __init__(self, app, parent, key, title):
        self.app = app
        self.key = key
        self.title = title
        self.full_df = pd.DataFrame()
        self.status_var = StringVar(value="대기")

        self.box = Frame(parent, bd=1, relief="solid")
        # 탭 전환: 컴테이너의 0,0에 배치하되 숨김은 grid_remove로 처리
        self.box.grid(row=0, column=0, sticky=N + S + E + W, padx=4, pady=4)
        self.box.grid_columnconfigure(0, weight=1)
        self.box.grid_rowconfigure(1, weight=1)

        head = Frame(self.box)
        head.grid(row=0, column=0, sticky=E + W)
        head.grid_columnconfigure(0, weight=1)
        self.title_label = Label(head, text=title, anchor="w", font=("맑은 고딕", 10, "bold"), cursor="hand2")
        self.title_label.grid(row=0, column=0, sticky=E + W, padx=5, pady=(3, 2))
        Label(head, textvariable=self.status_var, anchor="e", fg="#606060").grid(row=0, column=1, sticky=E, padx=6)

        frame = Frame(self.box)
        frame.grid(row=1, column=0, sticky=N + S + E + W, padx=3, pady=(0, 3))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(frame, show="headings", selectmode="extended")
        ybar = Scrollbar(frame, orient="vertical", command=self.tree.yview)
        xbar = Scrollbar(frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.tree.grid(row=0, column=0, sticky=N + S + E + W)
        ybar.grid(row=0, column=1, sticky=N + S)
        xbar.grid(row=1, column=0, sticky=E + W)

        for seq in ("<FocusIn>", "<Button-1>"):
            self.tree.bind(seq, lambda e: self.app.set_active_panel(self))
        for seq in ("<Control-a>", "<Control-A>"):
            self.tree.bind(seq, lambda e: self.select_all())
        for seq in ("<Control-c>", "<Control-C>"):
            self.tree.bind(seq, lambda e: self.copy_selection())
        self.title_label.bind("<Button-1>", lambda e: self.select_all())
        self.load(pd.DataFrame())

    def show(self):
        self.box.grid(row=0, column=0, sticky=N + S + E + W, padx=4, pady=4)

    def hide(self):
        self.box.grid_remove()

    def clear(self, status="대기"):
        self.status_var.set(status)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = []
        self.full_df = pd.DataFrame()

    def load(self, df, status=None):
        self.tree.delete(*self.tree.get_children())
        self.full_df = pd.DataFrame() if df is None else df.copy()
        shown = self.full_df.head(MAX_DISPLAY_ROWS).copy()
        if shown.empty:
            shown = pd.DataFrame({"message": ["표시할 데이터가 없습니다."]})
            self.full_df = pd.DataFrame()
        shown = shown.astype(object).where(pd.notna(shown), "")
        cols = [str(c) for c in shown.columns]
        self.tree["columns"] = cols
        for c in cols:
            self.tree.heading(c, text=c)
        self.autofit(shown)
        for row in shown.itertuples(index=False, name=None):
            self.tree.insert("", END, values=[str(v) for v in row])
        if status is not None:
            self.status_var.set(status)

    def load_error(self, error_text):
        self.status_var.set("실패")
        lines = error_text.splitlines() or [error_text]
        self.load(pd.DataFrame({"error": lines}))

    def autofit(self, df):
        font = tkfont.nametofont("TkDefaultFont")
        for col in [str(c) for c in df.columns]:
            sample = [col] + df[col].astype(str).head(3000).tolist()
            width = max(font.measure(x) for x in sample) + 18
            self.tree.column(col, width=max(18, min(1200, width)), minwidth=8, anchor="w", stretch=False)

    def select_all(self):
        items = self.tree.get_children()
        if items:
            self.tree.selection_set(items)
            self.tree.focus(items[0])
        self.tree.focus_set()
        self.app.set_active_panel(self)
        return "break"

    def copy_selection(self):
        self.app.set_active_panel(self)
        cols = list(self.tree["columns"])
        if not cols:
            return "break"
        selected = list(self.tree.selection())
        all_items = list(self.tree.get_children())
        if not selected or len(selected) == len(all_items):
            data = self.full_df.to_csv(sep="\t", index=False, lineterminator="\n") if not self.full_df.empty else "\t".join(cols)
        else:
            out = io.StringIO()
            writer = csv.writer(out, delimiter="\t", lineterminator="\n")
            writer.writerow(cols)
            writer.writerows([self.tree.item(i, "values") for i in selected])
            data = out.getvalue()
        self.app.root.clipboard_clear()
        self.app.root.clipboard_append(data)
        self.app.root.update()
        return "break"


# =============================================================================
# Main App
# =============================================================================


class SQLRunnerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sorter Data SQL Runner v48")
        self.root.geometry("1680x980")
        self.root.minsize(1300, 780)
        self._maximize()
        ttk.Style().configure("Small.TLabelframe.Label", font=SMALL_FONT_BOLD)

        start, end = default_datetime_range()
        self.start_var, self.end_var = StringVar(value=start), StringVar(value=end)
        self.status_var, self.elapsed_var = StringVar(value="대기"), StringVar(value="")
        self.site_vars = {s: BooleanVar(value=True) for s in SITE_OPTIONS}
        self.machine_vars = {m: BooleanVar(value=True) for m in MACHINE_OPTIONS}
        # equipment_vars: 라벨(2차 변경명) 단위로 체크박스 생성
        all_labels_ordered = []
        seen = set()
        for eq, port, machine, label in EQUIPMENT_ROWS:
            if label not in seen:
                all_labels_ordered.append(label)
                seen.add(label)
        # 기본 선택: DEFAULT_EQUIPMENT에 속하는 라벨
        default_labels = {label for eq, port, machine, label in EQUIPMENT_ROWS if eq in DEFAULT_EQUIPMENT}
        self.equipment_vars = {label: BooleanVar(value=label in default_labels) for label in all_labels_ordered}
        # 라벨 → EquipmentID 집합 (같은 라벨에 2포트 장비 포함, SQL IN절 전달용)
        self.label_to_eq = {}
        for eq, port, machine, label in EQUIPMENT_ROWS:
            self.label_to_eq.setdefault(label, set()).add(eq)
        # 라벨 → Machine 매핑 (visible_equipment 필터용)
        self.label_to_machine = {label: machine for eq, port, machine, label in EQUIPMENT_ROWS}
        # EquipmentID → 표시 라벨 (Port1 기준, 결과필터용)
        self.eq_label_map = {}
        for eq, port, machine, label in EQUIPMENT_ROWS:
            if eq not in self.eq_label_map:
                self.eq_label_map[eq] = label
        self.result_equipment_vars = {}
        self.panels = {}
        self.results = {}
        self.panel_elapsed = {}
        self.combined_df = pd.DataFrame()
        self.running = False
        self.active_panel = None
        self.machine_cbs = {}
        self.equipment_cbs = {}
        self.equipment_pos = {}
        self.result_filter_frame = None
        self.result_filter_note = StringVar(value="조회 후 표시")

        self.build_ui()
        self.on_filter_change()

    def _maximize(self):
        try:
            self.root.state("zoomed")
        except Exception:
            try:
                self.root.attributes("-zoomed", True)
            except Exception:
                pass

    def build_ui(self):
        top = Frame(self.root)
        top.pack(side=TOP, fill=X, padx=10, pady=(8, 4))
        Label(top, text="Main 화면", font=("맑은 고딕", 13, "bold")).pack(side=LEFT, padx=(0, 18))
        Label(top, text="시작일시").pack(side=LEFT)
        Entry(top, textvariable=self.start_var, width=22).pack(side=LEFT, padx=6)
        Label(top, text="종료일시").pack(side=LEFT)
        Entry(top, textvariable=self.end_var, width=22).pack(side=LEFT, padx=6)
        self.run_button = Button(top, text="실행", command=self.run_async, width=10)
        self.run_button.pack(side=LEFT, padx=(10, 4))
        Button(top, text="통합 Excel 저장", command=self.export_xlsx, width=15).pack(side=LEFT, padx=4)
        Label(top, textvariable=self.status_var, fg="#404040").pack(side=LEFT, padx=(18, 8))
        Label(top, textvariable=self.elapsed_var, fg="#404040").pack(side=LEFT, padx=(6, 0))

        filters = Frame(self.root)
        filters.pack(side=TOP, fill=X, padx=10, pady=(0, 4))
        self.build_filter_bar(filters)

        # ── 바디: 좌측 탭 리스트 + 우측 단일 패널 ──
        body = Frame(self.root)
        body.pack(side=TOP, fill=BOTH, expand=True, padx=10, pady=(0, 4))
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        # 좌측 탭 리스트 패널
        tab_frame = Frame(body, bd=1, relief="solid", width=170)
        tab_frame.grid(row=0, column=0, sticky=N + S + W, padx=(0, 6))
        tab_frame.pack_propagate(False)
        Label(tab_frame, text="조회 항목", font=("맑은 고딕", 9, "bold"),
              bg="#2c3e50", fg="white", pady=6).pack(fill=X)
        self.tab_listbox = Listbox(tab_frame, selectmode=SINGLE, font=("맑은 고딕", 9),
                                   activestyle="none", relief="flat", bd=0,
                                   selectbackground="#2980b9", selectforeground="white",
                                   bg="#ecf0f1", fg="#2c3e50", highlightthickness=0)
        self.tab_listbox.pack(fill=BOTH, expand=True, pady=4, padx=4)
        for key, title in PANELS:
            self.tab_listbox.insert(END, f"  {title}")
        self.tab_listbox.bind("<<ListboxSelect>>", self._on_tab_select)

        # 각 탭 상태 표시 (탭 아래 상태 표시)
        self.tab_status_labels = {}
        status_frame = Frame(tab_frame, bg="#ecf0f1")
        status_frame.pack(fill=X, padx=4, pady=(0, 4))
        for key, title in PANELS:
            lbl = Label(status_frame, text="대기", font=("맑은 고딕", 7),
                        fg="#888888", bg="#ecf0f1", anchor="w")
            lbl.pack(fill=X)
            self.tab_status_labels[key] = lbl

        # 우측: 단일 결과 패널 컨테이너
        self.panel_container = Frame(body)
        self.panel_container.grid(row=0, column=1, sticky=N + S + E + W)
        self.panel_container.grid_rowconfigure(0, weight=1)
        self.panel_container.grid_columnconfigure(0, weight=1)

        # ResultPanel 생성 (숨김 상태로)
        for key, title in PANELS:
            self.panels[key] = ResultPanel(self, self.panel_container, key, title)
            self.panels[key].hide()

        # 첫 번째 탭 선택
        self.tab_listbox.selection_set(0)
        self._on_tab_select(None)

        # 하단: 단계별 소요시간 표시줄 (어디가 느린지 확인용)
        timing_bar = Frame(self.root)
        timing_bar.pack(side=BOTTOM, fill=X, padx=10, pady=(0, 6))
        Label(timing_bar, text="단계별 시간:", font=SMALL_FONT_BOLD).pack(side=LEFT, padx=(0, 6))
        self.timing_var = StringVar(value="(실행 후 표시)")
        Label(timing_bar, textvariable=self.timing_var, font=SMALL_FONT,
              fg="#2c3e50", anchor="w").pack(side=LEFT, fill=X, expand=True)

        # 로그 영역 제거 — 내부 로그는 숨겨진 Text에 유지 (코드 호환성)
        self.log_text = Text(self.root)
        # 화면에 표시하지 않음

    def build_filter_bar(self, parent):
        """컴팩트 필터 바: 1행 구조 (Site+Machine 세로 / Equipment / 결과필터)"""
        row1 = Frame(parent)
        row1.pack(side=TOP, fill=X, pady=(2, 2))

        # 좌측: Site 위 / Machine 아래
        left_col = Frame(row1)
        left_col.pack(side=LEFT, padx=(0, 8), pady=1, anchor=N)

        lf_site = ttk.LabelFrame(left_col, text="Site")
        lf_site.pack(side=TOP, fill=X, pady=(0, 4))
        for i, s in enumerate(SITE_OPTIONS):
            Checkbutton(lf_site, text=s, variable=self.site_vars[s],
                        command=self.on_filter_change).grid(row=0, column=i, sticky=W, padx=3, pady=2)

        lf_mach = ttk.LabelFrame(left_col, text="Machine")
        lf_mach.pack(side=TOP, fill=X)
        # NoATW / ATW → row 0, JC2 6,7line → row 1 (_MACH_GRID_POS 참조)
        for m in MACHINE_OPTIONS:
            r, c = self._MACH_GRID_POS[m]
            cb = Checkbutton(lf_mach, text=m, variable=self.machine_vars[m],
                             command=self.on_filter_change)
            cb.grid(row=r, column=c, sticky=W, padx=4, pady=2)
            self.machine_cbs[m] = cb

        # 중앙: Equipment (버튼 세로 3줄)
        lf_eq = ttk.LabelFrame(row1, text="Equipment")
        lf_eq.pack(side=LEFT, fill=BOTH, expand=True, pady=1)
        btn_frame = Frame(lf_eq)
        btn_frame.grid(row=0, column=0, sticky=N + W, padx=4, pady=2)
        for txt, cmd in [("전체", self.select_all_visible), ("기본", self.select_default_visible), ("해제", self.clear_visible)]:
            Button(btn_frame, text=txt, command=cmd, width=5,
                   font=SMALL_FONT).pack(side=TOP, pady=1)
        eq_grid = Frame(lf_eq)
        eq_grid.grid(row=0, column=1, sticky=E + W, padx=4, pady=2)
        lf_eq.grid_columnconfigure(1, weight=1)
        for i, label in enumerate(self.equipment_vars):
            r, c = divmod(i, 8)
            cb = Checkbutton(eq_grid, text=label, variable=self.equipment_vars[label],
                             font=SMALL_FONT)
            cb.grid(row=r, column=c, sticky=W, padx=2, pady=0)
            self.equipment_cbs[label] = cb
            self.equipment_pos[label] = (r, c)

        # 우측: 결과 Equipment 필터
        lf_rf = ttk.LabelFrame(row1, text="결과 Equipment 필터")
        lf_rf.pack(side=LEFT, fill=BOTH, pady=1, padx=(8, 0))
        btn_rf = Frame(lf_rf)
        btn_rf.grid(row=0, column=0, sticky=N + W, padx=4, pady=2)
        Button(btn_rf, text="전체", command=self.select_all_result,
               width=5, font=SMALL_FONT).pack(side=TOP, pady=1)
        Button(btn_rf, text="해제", command=self.clear_result,
               width=5, font=SMALL_FONT).pack(side=TOP, pady=1)
        Label(lf_rf, textvariable=self.result_filter_note,
              fg="#606060", font=SMALL_FONT).grid(row=0, column=1, sticky=N + W, padx=4, pady=2)
        self.result_filter_frame = Frame(lf_rf)
        self.result_filter_frame.grid(row=0, column=2, sticky=N + W, padx=2, pady=2)
        lf_rf.grid_columnconfigure(2, weight=1)
        self._lf_rf = lf_rf
        self._result_cb_widgets = []
        self._relayout_job = None
        lf_rf.bind("<Configure>", self._on_result_filter_resize)

    # -------------------------- filter --------------------------
    # Machine 체크박스 배치: NoATW/ATW → 0행, JC2 6,7line → 1행
    _MACH_GRID_POS = {"NoATW": (0, 0), "ATW": (0, 1), "JC2 6,7line": (1, 0)}

    def machine_visible(self, machine):
        return not (machine == "JC2 6,7line" and not self.site_vars["JC02"].get())

    def visible_machines(self):
        return [m for m in MACHINE_OPTIONS if self.machine_visible(m)]

    def visible_equipment(self):
        sites = [s for s, v in self.site_vars.items() if v.get()]
        machines = {m for m in self.visible_machines() if self.machine_vars[m].get()}
        result = []
        for label in self.equipment_vars:
            eq_set = self.label_to_eq.get(label, set())
            machine = self.label_to_machine.get(label, "")
            site_ok = any(eq.startswith(s) for eq in eq_set for s in sites)
            machine_ok = machine in machines
            if site_ok and machine_ok:
                result.append(label)
        return result

    def on_filter_change(self):
        for m in MACHINE_OPTIONS:
            cb = self.machine_cbs.get(m)
            if not cb:
                continue
            if self.machine_visible(m):
                r, c = self._MACH_GRID_POS[m]
                cb.grid(row=r, column=c, sticky=W, padx=4, pady=2)
            else:
                cb.grid_remove()
        visible = set(self.visible_equipment())
        for eq, cb in self.equipment_cbs.items():
            if eq in visible:
                r, c = self.equipment_pos[eq]
                cb.grid(row=r, column=c, sticky=W, padx=3, pady=1)
            else:
                cb.grid_remove()

    def select_all_visible(self):
        for eq in self.visible_equipment():
            self.equipment_vars[eq].set(True)

    def select_default_visible(self):
        for eq in self.visible_equipment():
            self.equipment_vars[eq].set(eq in DEFAULT_EQUIPMENT)

    def clear_visible(self):
        for eq in self.visible_equipment():
            self.equipment_vars[eq].set(False)

    def selected_config(self):
        sites = [s for s, v in self.site_vars.items() if v.get()]
        machines = [m for m in self.visible_machines() if self.machine_vars[m].get()]
        visible = set(self.visible_equipment())
        # 선택된 라벨 → EquipmentID 집합 flatten (중복 제거, SQL IN절 전달용)
        eq_set = set()
        for label, v in self.equipment_vars.items():
            if v.get() and label in visible:
                eq_set.update(self.label_to_eq.get(label, set()))
        equipment = sorted(eq_set)
        if not sites:
            raise ValueError("SiteID를 1개 이상 선택하세요.")
        if not machines:
            raise ValueError("Machine을 1개 이상 선택하세요.")
        if not equipment:
            raise ValueError("EquipmentID를 1개 이상 선택하세요.")
        return RunConfig(self.start_var.get().strip(), self.end_var.get().strip(), sites, machines, equipment)

    # -------------------------- result filter --------------------------
    def rebuild_result_filter(self):
        for child in self.result_filter_frame.winfo_children():
            child.destroy()
        eqs = sorted({str(x) for df in self.results.values() if df is not None and not df.empty and "EquipmentID" in df.columns for x in df["EquipmentID"].dropna().unique()})
        self.result_equipment_vars = {}
        if not eqs:
            self.result_filter_note.set("조회 결과 없음")
            self.apply_result_filter()
            return
        self.result_filter_note.set(f"{len(eqs)}개")
        self._result_cb_widgets = []
        for eq in eqs:
            var = BooleanVar(value=True)
            self.result_equipment_vars[eq] = var
            label = self.eq_label_map.get(eq, eq)
            cb = Checkbutton(self.result_filter_frame, text=label, variable=var,
                              command=self.apply_result_filter, font=SMALL_FONT)
            self._result_cb_widgets.append(cb)
        self._relayout_result_filter()
        self.apply_result_filter()

    def _on_result_filter_resize(self, event):
        if self._relayout_job:
            self.root.after_cancel(self._relayout_job)
        self._relayout_job = self.root.after(120, self._relayout_result_filter)

    def _relayout_result_filter(self):
        self._relayout_job = None
        widgets = self._result_cb_widgets
        if not widgets:
            return
        self._lf_rf.update_idletasks()
        avail_width = self._lf_rf.winfo_width() - 90
        max_w = max(w.winfo_reqwidth() for w in widgets)
        cols = (avail_width // max_w) if avail_width > 0 else 5
        cols = max(5, cols)
        for i, w in enumerate(widgets):
            r, c = divmod(i, cols)
            w.grid(row=r, column=c, sticky=W, padx=2, pady=0)

    def selected_result_equipment(self):
        return [eq for eq, var in self.result_equipment_vars.items() if var.get()]

    def select_all_result(self):
        for v in self.result_equipment_vars.values():
            v.set(True)
        self.apply_result_filter()

    def clear_result(self):
        for v in self.result_equipment_vars.values():
            v.set(False)
        self.apply_result_filter()

    def apply_result_filter(self):
        if not self.results:
            self.combined_df = pd.DataFrame()
            return
        selected = self.selected_result_equipment() if self.result_equipment_vars else None
        display_results = {}
        for key in PANEL_KEYS:
            raw = self.results.get(key, pd.DataFrame())
            raw_f = raw.copy() if selected is None else filter_by_equipment(raw, selected)
            shown = display_df(raw_f)
            display_results[key] = shown
            elapsed = self.panel_elapsed.get(key)
            status = f"{len(shown):,} / 원본 {0 if raw is None else len(raw):,} rows"
            if elapsed is not None:
                status += f" / {elapsed:.1f}초"
            self.panels[key].load(shown, status=status)
            self._set_tab_row_status(key, len(shown))
        self.combined_df = self.make_combined(display_results, display_ready=True)

    # -------------------------- UI helpers --------------------------
    def _on_tab_select(self, event):
        """좌측 탭 클릭 → 해당 패널만 표시"""
        sel = self.tab_listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        active_key = PANEL_KEYS[idx]
        for key, panel in self.panels.items():
            if key == active_key:
                panel.show()
                self.active_panel = panel
            else:
                panel.hide()

    def _update_tab_status(self, key, text, color="#888888"):
        """탭 목록 하단 상태 레이블 업데이트"""
        lbl = self.tab_status_labels.get(key)
        if lbl:
            lbl.config(text=text, fg=color)

    def _set_tab_row_status(self, key, row_count, suffix=""):
        """행 수 기준 탭 상태 갱신 (0건이면 회색, 있으면 녹색)"""
        color = "#27ae60" if row_count > 0 else "#888888"
        self._update_tab_status(key, f"{row_count:,} rows{suffix}", color)

    def set_active_panel(self, panel):
        self.active_panel = panel

    def log(self, msg):
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)
        self.root.update_idletasks()

    def ui(self, func, *args, **kwargs):
        self.root.after(0, lambda: func(*args, **kwargs))

    def set_running(self, running):
        self.running = running
        self.run_button.configure(state="disabled" if running else "normal")

    def clear_for_run(self):
        for p in self.panels.values():
            p.clear("대기")
        for key in PANEL_KEYS:
            self._update_tab_status(key, "대기", "#888888")
        self.results = {}
        self.panel_elapsed = {}
        self.combined_df = pd.DataFrame()
        self.log_text.delete("1.0", END)
        self.status_var.set("실행 준비")
        self.elapsed_var.set("")
        self.rebuild_result_filter()

    # -------------------------- run --------------------------
    def run_async(self):
        if self.running:
            return
        self.on_filter_change()
        try:
            cfg = self.selected_config()
            if not cfg.start_date or not cfg.end_date:
                raise ValueError("시작일시와 종료일시를 입력하세요.")
        except Exception as exc:
            messagebox.showwarning("필터 확인", str(exc))
            return
        self.clear_for_run()
        self.set_running(True)
        threading.Thread(target=self.run_worker, args=(cfg,), daemon=True).start()

    def run_worker(self, cfg: RunConfig):
        t0 = time.perf_counter()
        results, timing, errors = {}, [], {}
        try:
            self.ui(self.log, "실행 시작")
            self.ui(self.log, f"기간: {cfg.start_date} ~ {cfg.end_date}")
            self.ui(self.log, f"SiteID: {', '.join(cfg.sites)} / Machine: {', '.join(cfg.machines)} / EquipmentID: {len(cfg.equipment)}개")
            engine = build_engine()
            with engine.connect() as conn:
                self.ui(self.status_var.set, "SQL Server base 생성 중")
                t = time.perf_counter()
                count_df, subs = setup_temp_tables(conn, cfg.start_date, cfg.end_date, cfg.equipment)
                sec = time.perf_counter() - t
                row = count_df.iloc[0].to_dict() if not count_df.empty else {}
                self.ui(self.log, f"Base 생성 완료: {int(row.get('BaseRows', 0)):,} rows / label-missing {int(row.get('LabelMissingRows', 0)):,} rows / label-issued {int(row.get('LabelIssuedAnomalyRows', 0)):,} rows / {sec:.1f}초")
                timing.append({"Step": "Base temp", "Seconds": round(sec, 1), "Rows": int(row.get("BaseRows", 0))})
                # 하위 단계별 시간 (병목 진단용)
                for step, ssec, srows in subs:
                    timing.append({"Step": step, "Seconds": round(ssec, 1), "Rows": srows if srows is not None and srows >= 0 else 0})
                    self.ui(self.log, f"{step}: {ssec:.1f}초" + (f" / {srows:,} rows" if srows is not None and srows >= 0 else ""))
                for i, key in enumerate(PANEL_KEYS, 1):
                    title = PANEL_TITLE[key]
                    self.ui(self.status_var.set, f"{i}/5 {title} 조회 중")
                    self.ui(self.panels[key].status_var.set, "조회 중")
                    t = time.perf_counter()
                    try:
                        df = fetch_panel(conn, key)
                        sec = time.perf_counter() - t
                        results[key] = df
                        self.panel_elapsed[key] = sec
                        timing.append({"Step": title, "Seconds": round(sec, 1), "Rows": len(df)})
                        self.ui(self.log, f"{title} 완료: {len(df):,} rows / {sec:.1f}초")
                        self.ui(self.apply_panel_now, key, df, sec)
                    except Exception:
                        sec = time.perf_counter() - t
                        err = traceback.format_exc()
                        errors[key] = err
                        results[key] = pd.DataFrame()
                        timing.append({"Step": title, "Seconds": round(sec, 1), "Rows": 0})
                        self.ui(self.log, f"{title} 실패: {sec:.1f}초")
                        self.ui(self.log, err)
                        self.ui(self.panels[key].load_error, err)
                        self.ui(self._update_tab_status, key, "실패", "#e74c3c")
            total = time.perf_counter() - t0
            self.ui(self.finish_run, results, pd.DataFrame(timing), total, len(errors))
        except Exception:
            err = traceback.format_exc()
            self.ui(self.log, "실행 실패")
            self.ui(self.log, err)
            self.ui(self.status_var.set, "실패")
            for key in PANEL_KEYS:
                self.ui(self.panels[key].load_error, err)
        finally:
            self.ui(self.set_running, False)

    def apply_panel_now(self, key, raw, elapsed):
        shown = display_df(raw)
        self.panels[key].load(shown, status=f"{len(shown):,} rows / {elapsed:.1f}초")
        self._set_tab_row_status(key, len(shown), suffix=f" / {elapsed:.1f}초")

    def finish_run(self, results, timing_df, total_sec, fail_count):
        # results는 raw 그대로 저장 (apply_result_filter에서 display_df가 fix_types 처리)
        self.results = results
        self.rebuild_result_filter()
        self.status_var.set("완료" if fail_count == 0 else f"부분 완료: 실패 {fail_count}")
        self.elapsed_var.set(f"전체 {total_sec:.1f}초")
        # 단계별 소요시간 표시줄 업데이트
        try:
            parts = []
            if timing_df is not None and not timing_df.empty:
                for _, r in timing_df.iterrows():
                    parts.append(f"{r['Step']} {r['Seconds']:.1f}s")
            parts.append(f"[전체 {total_sec:.1f}s]")
            self.timing_var.set("  |  ".join(parts))
        except Exception:
            self.timing_var.set(f"전체 {total_sec:.1f}초")
        self.log(f"전체 완료: {total_sec:.1f}초 / 저장 가능 {len(self.combined_df):,} rows")

    def make_combined(self, results, display_ready=False):
        frames = []
        for key in PANEL_KEYS:
            df = results.get(key)
            if df is None or df.empty:
                continue
            out = df.copy() if display_ready else display_df(df)
            if out.empty:
                continue
            out.insert(0, "Table", PANEL_TITLE[key])
            out.insert(1, "Row_No", range(1, len(out) + 1))
            frames.append(out)
        return pd.concat(frames, ignore_index=True, sort=False) if frames else pd.DataFrame()

    # -------------------------- Excel --------------------------
    def export_xlsx(self):
        if self.combined_df is None or self.combined_df.empty:
            messagebox.showwarning("저장 불가", "저장할 통합 결과가 없습니다.")
            return
        try:
            from openpyxl.styles import Font, PatternFill, Alignment
            from openpyxl.utils import get_column_letter
        except Exception:
            messagebox.showerror("저장 실패", "xlsx 저장에는 openpyxl이 필요합니다.\n\npip install openpyxl")
            return
        if len(self.combined_df) + 1 > 1_048_576:
            messagebox.showerror("저장 실패", f"Excel 한 시트 최대 행 수 초과: {len(self.combined_df):,} rows")
            return
        path = filedialog.asksaveasfilename(
            title="통합 결과 Excel 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            initialfile="sorter_sql_combined_result.xlsx",
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        try:
            df = fix_types(self.combined_df.copy())
            text_cols = [i for i, c in enumerate(df.columns, 1) if str(c) in TEXT_ID_COLUMNS]
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="통합결과", index=False)
                ws = writer.sheets["통합결과"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                fill, font = PatternFill("solid", fgColor="D9EAF7"), Font(bold=True)
                for cell in ws[1]:
                    cell.fill, cell.font = fill, font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for col_idx in text_cols:
                    col_letter = get_column_letter(col_idx)
                    for row_idx in range(2, ws.max_row + 1):
                        cell = ws[f"{col_letter}{row_idx}"]
                        cell.value = normalize_id(cell.value)
                        cell.number_format = "@"
                sample = df.head(5000)
                for i, c in enumerate(df.columns, 1):
                    values = [str(c)] + sample[c].map(normalize_id if str(c) in TEXT_ID_COLUMNS else str).fillna("").tolist()
                    ws.column_dimensions[get_column_letter(i)].width = min(max(max(map(len, values)) + 2, 8), 60)
            messagebox.showinfo("저장 완료", f"통합 결과 Excel 저장 완료\n{path}")
        except Exception as exc:
            messagebox.showerror("저장 실패", f"Excel 저장 중 오류가 발생했습니다.\n\n{exc}")


# =============================================================================
# Entry point
# =============================================================================


def main():
    root = Tk()
    SQLRunnerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
