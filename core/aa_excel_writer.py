# -*- coding: utf-8 -*-
"""
아미노산 함량 분석 엑셀 라이터
시트 구성:
  1) 시스템적합성  — 나눔고딕 14 / 분리도 6런 + 판정
  2) 표준액 면적   — 맑은 고딕 14 / 6런 면적, 평균, %RSD + 판정  (AA당 2열)
  3) 검액_{lot}   — 맑은 고딕 14 / Sample 1/2 면적             (AA당 2열)
"""
import io
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from core.aa_pdf_parser import AA_ORDER

# ── 폰트 ──────────────────────────────────────────────────────────
FN_SST  = "나눔고딕"    # 시스템적합성
FN_DATA = "맑은 고딕"  # 표준액 면적 / 검액 면적
FS      = 14

# ── 채우기 ────────────────────────────────────────────────────────
FILL_HEADER = PatternFill("solid", fgColor="BDD7EE")   # 파란 헤더
FILL_GRAY   = PatternFill("solid", fgColor="D9D9D9")   # 행 레이블
FILL_DATA   = PatternFill("solid", fgColor="FFFFE1")   # 데이터 셀 (연노랑)
FILL_PASS   = PatternFill("solid", fgColor="C6EFCE")   # 적합 (초록)
FILL_FAIL   = PatternFill("solid", fgColor="FFC7CE")   # 부적합 (빨강)

# ── 테두리 ────────────────────────────────────────────────────────
_T = Side(style="thin")
_BORDER = Border(left=_T, right=_T, top=_T, bottom=_T)


def _c(ws, row, col, value=None, *,
       font_name=FN_DATA, fill=None, bold=False,
       align="center", border=True):
    """셀 하나 쓰기 헬퍼."""
    c = ws.cell(row=row, column=col, value=value)
    c.font      = Font(name=font_name, size=FS, bold=bold)
    c.alignment = Alignment(horizontal=align, vertical="center")
    if border:
        c.border = _BORDER
    if fill:
        c.fill = fill
    return c


def _merged_header(ws, row, c1, c2, value, fn):
    """
    두 열(c1, c2)을 병합한 헤더 셀 작성.
    openpyxl MergedCell은 border 설정이 파일에 반영되지 않으므로
    마스터 셀(c1)에 전체 테두리를 한 번에 적용한다.
    """
    ws.merge_cells(start_row=row, start_column=c1, end_row=row, end_column=c2)

    tl = ws.cell(row=row, column=c1, value=value)
    tl.font      = Font(name=fn, size=FS, bold=True)
    tl.alignment = Alignment(horizontal="center", vertical="center")
    tl.fill      = FILL_HEADER
    tl.border    = Border(left=_T, right=_T, top=_T, bottom=_T)


def _calc_rsd(values: list):
    if len(values) < 2:
        return None
    avg = sum(values) / len(values)
    if avg == 0:
        return None
    sd = math.sqrt(sum((x - avg) ** 2 for x in values) / (len(values) - 1))
    return sd / avg * 100


# ══════════════════════════════════════════════════════════════════
# 시트 1 : 시스템적합성  (나눔고딕 14)
# ══════════════════════════════════════════════════════════════════
def _sheet_sst(wb, resolutions: list[dict]):
    ws = wb.create_sheet("시스템적합성")

    ws.column_dimensions["A"].width = 12
    for i in range(2, 10):
        ws.column_dimensions[get_column_letter(i)].width = 9

    def cv(row, col, val=None, *, fill=None, bold=False):
        return _c(ws, row, col, val,
                  font_name=FN_SST, fill=fill, bold=bold)

    # 헤더
    cv(1, 1, "분리도", fill=FILL_HEADER, bold=True)
    for i in range(1, 7):
        cv(1, i + 1, str(i), fill=FILL_HEADER, bold=True)
    cv(1, 8, "판정", fill=FILL_HEADER, bold=True)

    for row_i, aa in enumerate(AA_ORDER):
        r = row_i + 2
        cv(r, 1, aa, fill=FILL_DATA)

        res_vals = [
            resolutions[ri].get(aa) if ri < len(resolutions) else None
            for ri in range(6)
        ]

        for col_i, rv in enumerate(res_vals):
            cv(r, col_i + 2,
               round(rv, 3) if rv is not None else "N/A",
               fill=FILL_DATA)

        valid = [v for v in res_vals if v is not None]
        if not valid:
            cv(r, 8, "N/A", fill=FILL_DATA)
        elif all(v >= 1.2 for v in valid):
            cv(r, 8, "적합",   fill=FILL_PASS, bold=True)
        else:
            cv(r, 8, "부적합", fill=FILL_FAIL, bold=True)

    ws.row_dimensions[1].height = 20
    return ws


# ══════════════════════════════════════════════════════════════════
# 시트 2 : 표준액 면적  (맑은 고딕 14 / AA당 2열)
# col A = 행 레이블
# col 2+i*2     = 값
# col 2+i*2+1   = %RSD행에서 판정, 나머지 행에서 빈칸
# ══════════════════════════════════════════════════════════════════
def _write_std_group(ws, start_row: int, group_aas: list, runs: list) -> int:
    """그룹 하나 작성. 다음 시작 행 반환."""

    # ─ 행 레이블 열(A) 헤더칸
    _c(ws, start_row, 1, "", fill=FILL_HEADER)

    # ─ AA 헤더 (2열 병합) — 내부선 없이
    for i, aa in enumerate(group_aas):
        c1 = 2 + i * 2
        _merged_header(ws, start_row, c1, c1 + 1, aa, FN_DATA)

    # ─ 런 1~6
    for ri, run_data in enumerate(runs[:6]):
        r = start_row + 1 + ri
        _c(ws, r, 1, str(ri + 1), fill=FILL_GRAY, bold=True)
        for i, aa in enumerate(group_aas):
            c1 = 2 + i * 2
            val = run_data.get(aa)
            _c(ws, r, c1,     round(val, 3) if val is not None else "", fill=FILL_DATA)
            _c(ws, r, c1 + 1, "",                                        fill=FILL_DATA)

    # ─ 평균
    avg_r = start_row + 7
    _c(ws, avg_r, 1, "평균", fill=FILL_GRAY, bold=True)
    for i, aa in enumerate(group_aas):
        c1 = 2 + i * 2
        vals = [run[aa] for run in runs if aa in run]
        avg  = round(sum(vals) / len(vals), 3) if vals else ""
        _c(ws, avg_r, c1,     avg, fill=FILL_DATA)
        _c(ws, avg_r, c1 + 1, "",  fill=FILL_DATA)

    # ─ %RSD
    rsd_r = start_row + 8
    _c(ws, rsd_r, 1, "%RSD", fill=FILL_GRAY, bold=True)
    for i, aa in enumerate(group_aas):
        c1 = 2 + i * 2
        vals = [run[aa] for run in runs if aa in run]
        rsd  = _calc_rsd(vals)
        _c(ws, rsd_r, c1, round(rsd, 1) if rsd is not None else "", fill=FILL_DATA)

        if rsd is None:
            _c(ws, rsd_r, c1 + 1, "N/A")
        elif rsd <= 2.0:
            _c(ws, rsd_r, c1 + 1, "적합",   fill=FILL_PASS, bold=True)
        else:
            _c(ws, rsd_r, c1 + 1, "부적합", fill=FILL_FAIL, bold=True)

    return rsd_r + 2   # 빈 행 1줄 두고 다음 그룹


def _sheet_std_area(wb, runs: list[dict]):
    ws = wb.create_sheet("표준액 면적")

    ws.column_dimensions["A"].width = 8
    for i in range(len(AA_ORDER)):
        ws.column_dimensions[get_column_letter(2 + i * 2)    ].width = 14  # 값 열
        ws.column_dimensions[get_column_letter(2 + i * 2 + 1)].width = 7   # 판정 열

    groups = [AA_ORDER[:5], AA_ORDER[5:10], AA_ORDER[10:]]
    cur = 1
    for grp in groups:
        cur = _write_std_group(ws, cur, grp, runs)

    return ws


# ══════════════════════════════════════════════════════════════════
# 시트 3+ : 검액 면적  (맑은 고딕 14 / AA당 2열)
# 표준액과 동일 2열 구조 / 두 번째 열은 빈칸
# ══════════════════════════════════════════════════════════════════
def _write_sp_group(ws, start_row: int, group_aas: list, lot_data: dict) -> int:
    """검액 그룹 하나 작성. 다음 시작 행 반환."""

    _c(ws, start_row, 1, "", fill=FILL_HEADER)

    # AA 헤더 (2열 병합) — 내부선 없이
    for i, aa in enumerate(group_aas):
        c1 = 2 + i * 2
        _merged_header(ws, start_row, c1, c1 + 1, aa, FN_DATA)

    # Sample 1 / Sample 2
    for s in (1, 2):
        r = start_row + s
        _c(ws, r, 1, f"Sample {s}", fill=FILL_GRAY, bold=True)
        for i, aa in enumerate(group_aas):
            c1 = 2 + i * 2
            val = lot_data.get(s, {}).get(aa)
            _c(ws, r, c1,     round(val, 3) if val is not None else "", fill=FILL_DATA)
            _c(ws, r, c1 + 1, "",                                        fill=FILL_DATA)

    return start_row + 4   # 데이터 3행 + 빈 1행


def _sheet_sp(wb, lot_id: str, lot_data: dict):
    ws = wb.create_sheet(f"검액_{lot_id}"[:31])

    ws.column_dimensions["A"].width = 12
    for i in range(len(AA_ORDER)):
        ws.column_dimensions[get_column_letter(2 + i * 2)    ].width = 14
        ws.column_dimensions[get_column_letter(2 + i * 2 + 1)].width = 7

    groups = [AA_ORDER[:5], AA_ORDER[5:10], AA_ORDER[10:]]
    cur = 1
    for grp in groups:
        cur = _write_sp_group(ws, cur, grp, lot_data)

    return ws


# ══════════════════════════════════════════════════════════════════
# 진입점
# ══════════════════════════════════════════════════════════════════
def write_aa_result(runs: list, resolutions: list, lots: dict) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    _sheet_sst(wb, resolutions)
    _sheet_std_area(wb, runs)
    for lot_id in sorted(lots.keys()):
        _sheet_sp(wb, lot_id, lots[lot_id])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
