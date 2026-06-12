# -*- coding: utf-8 -*-
"""
공진단 확인시험 엑셀 라이터
양식 이미지 기준:
  ● 시스템적합성: 그룹1(Morroniside, Loganin, Nodakenin) + 그룹2(Rg1, Rb1, Decursin)
  ● 결과 확인: 그룹A(Rg1,Rb1,Morroniside) + 그룹B(Loganin,Nodakenin,Decursin) + 그룹C(Decursin angelate)
  각 화합물당 3개 transition 열
"""

import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 색상/스타일 ─────────────────────────────────────────────────────────
YELLOW  = PatternFill("solid", fgColor="FFFFE1")
BLUE_H  = PatternFill("solid", fgColor="BDD7EE")
GRAY    = PatternFill("solid", fgColor="D9D9D9")
WHITE   = PatternFill("solid", fgColor="FFFFFF")
TITLE_F = PatternFill("solid", fgColor="2E75B6")

THIN   = Side(style="thin")
MEDIUM = Side(style="medium")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
AC = Alignment(horizontal="center", vertical="center", wrap_text=True)
AL = Alignment(horizontal="left",   vertical="center")
FS = 11

# ── 양식 고정 구조 ──────────────────────────────────────────────────────
# (기본명, [transition 레이블], CSV 컬럼명 후보)
COMPOUNDS = {
    "Morroniside":       (["465.3 -> 243.2", "465.3 -> 141.1", "465.3 -> 155.0"],
                          ["Morroniside (465.3 -> 243.2)", "Morroniside"]),
    "Loganin":           (["449.2 -> 227.0", "449.2 -> 101.0", "449.2 -> 127.0"],
                          ["Loganin (449.2 -> 227.0)", "Loganin"]),
    "Nodakenin":         (["409.2 -> 186.9", "409.2 -> 229.0", "409.2 -> 247.2"],
                          ["Nodakenin (409.2 -> 186.9)", "Nodakenin"]),
    "Ginsenoside Rg1":   (["859.6 -> 637.5", "859.6 -> 475.5", "859.6 -> 799.7"],
                          ["Ginsenoside Rg1 (859.6 -> 637.5)", "Ginsenoside_Rg1", "Ginsenoside Rg1"]),
    "Ginsenoside Rb1":   (["1107.6 -> 945.7","1107.6 -> 178.8","1107.6 -> 221.2"],
                          ["Ginsenoside Rb1 (1107.6 -> 945.7)", "Ginsenoside Rb1"]),
    "Decursin":          (["329.1 -> 247.2", "329.1 -> 115.0", "329.1 -> 128.0"],
                          ["Decursin (329.1 -> 247.2)", "Decursin"]),
    "Decursin angelate": (["329.1 -> 247.2", "329.1 -> 115.0", "329.1 -> 128.0"],
                          ["Decursin angelate (329.1 -> 247.2)", "Decursin Angelate", "Decursin angelate"]),
}

# 시스템적합성 그룹
SST_GROUP1 = ["Morroniside", "Loganin", "Nodakenin"]
SST_GROUP2 = ["Ginsenoside Rg1", "Ginsenoside Rb1", "Decursin"]

# 결과 확인 그룹
RESULT_GROUP1 = ["Ginsenoside Rg1", "Ginsenoside Rb1", "Morroniside"]
RESULT_GROUP2 = ["Loganin", "Nodakenin", "Decursin"]
RESULT_GROUP3 = ["Decursin angelate"]

N_TRANS   = 3
DATA_COL  = 2  # 구분 열: col A, 데이터 시작: col B


def _cell(ws, row, col, value="", fill=None, bold=False, num_fmt=None, fs=FS):
    from openpyxl.cell import MergedCell
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        return cell
    cell.value = value
    cell.font  = Font(name="맑은 고딕", size=fs, bold=bold)
    cell.alignment = AC
    cell.border    = BORDER
    if fill:    cell.fill = fill
    if num_fmt: cell.number_format = num_fmt
    return cell


def _merge(ws, r1, c1, r2, c2):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)


def _title_bar(ws, row, text, c1, c2, fs=FS):
    _merge(ws, row, c1, row, c2)
    cell = ws.cell(row=row, column=c1)
    cell.value     = text
    cell.font      = Font(name="맑은 고딕", size=fs+1, bold=True, color="FFFFFF")
    cell.fill      = TITLE_F
    cell.alignment = AL
    cell.border    = BORDER
    ws.row_dimensions[row].height = 20


def _outer_border(ws, r1, c1, r2, c2):
    M = Side(style="medium")
    for c in range(c1, c2 + 1):
        b = ws.cell(r1, c).border
        ws.cell(r1, c).border = Border(top=M, bottom=b.bottom, left=b.left, right=b.right)
        b = ws.cell(r2, c).border
        ws.cell(r2, c).border = Border(top=b.top, bottom=M, left=b.left, right=b.right)
    for r in range(r1, r2 + 1):
        b = ws.cell(r, c1).border
        ws.cell(r, c1).border = Border(top=b.top, bottom=b.bottom, left=M, right=b.right)
        b = ws.cell(r, c2).border
        ws.cell(r, c2).border = Border(top=b.top, bottom=b.bottom, left=b.left, right=M)


def _lookup_compound(parsed: dict, csv_candidates: list) -> list:
    """CSV에서 화합물 데이터 검색 (다양한 이름 형식 대응)."""
    for name in csv_candidates:
        if name in parsed:
            return parsed[name]
    # 부분 매칭
    for cname in parsed:
        for cand in csv_candidates:
            if cand.lower() in cname.lower() or cname.lower() in cand.lower():
                return parsed[cname]
    return {}


def _rt_list(comp_data, section="std") -> list:
    """해당 섹션의 RT 값 리스트 반환 (없으면 stability 시도)."""
    runs = comp_data.get(section, [])
    if not runs:
        runs = comp_data.get("stability", [])
    return [r.get("rt") for r in runs]


def _area_list(comp_data, section="sp") -> list:
    return [r.get("area") for r in comp_data.get(section, [])]


# ── 시스템적합성 섹션 ────────────────────────────────────────────────
def _write_sst_group(ws, row, group_names, parsed):
    """시스템적합성 3화합물 × 3 transition 블록 작성."""
    LABEL_COL = 1
    data_start_col = 2
    n_cols = len(group_names) * N_TRANS  # 9
    last_col = data_start_col + n_cols - 1

    # 화합물 헤더 행 (각 화합물 3열 병합)
    _cell(ws, row, LABEL_COL, "구분", fill=BLUE_H, bold=True)
    for gi, gname in enumerate(group_names):
        transitions, _ = COMPOUNDS[gname]
        for ti, trans in enumerate(transitions):
            col = data_start_col + gi * N_TRANS + ti
            label = f"{gname}\n({trans}) Results"
            _cell(ws, row, col, label, fill=BLUE_H, bold=True, fs=9)
    ws.row_dimensions[row].height = 35
    row += 1

    # "RT Value" 서브헤더
    _cell(ws, row, LABEL_COL, "RT Value", fill=BLUE_H, bold=True)
    for col in range(data_start_col, last_col + 1):
        _cell(ws, row, col, "", fill=BLUE_H)
    row += 1

    # 데이터 행 1~6
    data_row_start = row
    for run_idx in range(6):
        _cell(ws, row, LABEL_COL, run_idx + 1, fill=WHITE)
        for gi, gname in enumerate(group_names):
            transitions, csv_names = COMPOUNDS[gname]
            comp_data = _lookup_compound(parsed, csv_names)
            rt_vals = _rt_list(comp_data, "std")
            rt = rt_vals[run_idx] if run_idx < len(rt_vals) else None
            for ti in range(N_TRANS):
                col = data_start_col + gi * N_TRANS + ti
                _cell(ws, row, col, rt, fill=YELLOW, num_fmt="0.000")
        row += 1
    data_row_end = row - 1

    # 통계 행
    for label, func in [("평균", "AVERAGE"), ("표준편차", "STDEV"), ("유지시간 RSD %\n(2.0 % 이하)", None)]:
        _cell(ws, row, LABEL_COL, label, fill=GRAY, bold=True, fs=9)
        for col in range(data_start_col, last_col + 1):
            cl = get_column_letter(col)
            rng = f"{cl}{data_row_start}:{cl}{data_row_end}"
            if func:
                formula = f"=IFERROR({func}({rng}),\"\")"
                fmt = "0.000"
            else:
                avg_row = row - 2
                std_row = row - 1
                formula = f"=IFERROR(IF({cl}{avg_row}=0,\"\",{cl}{std_row}/{cl}{avg_row}*100),\"\")"
                fmt = "0.00"
            cell = ws.cell(row, col)
            cell.value = formula
            cell.fill = GRAY
            cell.border = BORDER
            cell.font = Font(name="맑은 고딕", size=FS)
            cell.alignment = AC
            cell.number_format = fmt
        row += 1

    _outer_border(ws, data_row_start, LABEL_COL, row - 1, last_col)
    return row


# ── 결과 확인 섹션 ────────────────────────────────────────────────
def _write_result_group(ws, row, group_names, parsed):
    """결과 확인 그룹 블록 (S/N비 + 결과)."""
    LABEL_COL = 1
    data_start_col = 2
    n_cols = len(group_names) * N_TRANS
    last_col = data_start_col + n_cols - 1

    # 헤더
    _cell(ws, row, LABEL_COL, "구분", fill=BLUE_H, bold=True)
    for gi, gname in enumerate(group_names):
        transitions, _ = COMPOUNDS[gname]
        for ti, trans in enumerate(transitions):
            col = data_start_col + gi * N_TRANS + ti
            label = f"{gname}\n({trans}) Results"
            _cell(ws, row, col, label, fill=BLUE_H, bold=True, fs=9)
    ws.row_dimensions[row].height = 35
    row += 1

    # 각 SP 런별 S/N비 행
    # sp 데이터 수집 (전체 화합물 중 가장 많은 런 수)
    max_runs = 0
    for gname in group_names:
        _, csv_names = COMPOUNDS[gname]
        comp_data = _lookup_compound(parsed, csv_names)
        max_runs = max(max_runs, len(comp_data.get("sp", [])))

    if max_runs == 0:
        max_runs = 1  # 빈 행 최소 1개

    sn_rows = []
    for run_idx in range(max_runs):
        sn_row = row
        sn_rows.append(sn_row)
        _cell(ws, row, LABEL_COL, "S/N비", fill=WHITE, bold=True)
        for gi, gname in enumerate(group_names):
            _, csv_names = COMPOUNDS[gname]
            comp_data = _lookup_compound(parsed, csv_names)
            sp_list = comp_data.get("sp", [])
            area = sp_list[run_idx].get("area") if run_idx < len(sp_list) else None
            for ti in range(N_TRANS):
                col = data_start_col + gi * N_TRANS + ti
                # 1번째 transition에만 값, 나머지는 비움(qualifier는 데이터 없음)
                val = area if ti == 0 else None
                _cell(ws, row, col, val, fill=YELLOW, num_fmt="0")
        row += 1

    # 결과 행
    _cell(ws, row, LABEL_COL, "결과", fill=GRAY, bold=True)
    for col in range(data_start_col, last_col + 1):
        cl = get_column_letter(col)
        # S/N비가 10 이상이면 "검출", 미만이면 공란 (단순 예시)
        if sn_rows:
            formula = f'=IFERROR(IF({cl}{sn_rows[0]}>=10,"검출","미검출"),"")'
        else:
            formula = ""
        _cell(ws, row, col, formula if sn_rows else "", fill=GRAY, fs=FS)
    row += 1

    _outer_border(ws, sn_rows[0] if sn_rows else row - 2, LABEL_COL, row - 1, last_col)
    return row


# ── 열 너비 설정 ──────────────────────────────────────────────────
def _set_widths(ws, n_comp_cols):
    ws.column_dimensions["A"].width = 16
    for ci in range(1, n_comp_cols + 1):
        ws.column_dimensions[get_column_letter(ci + 1)].width = 11


# ── 메인 엔트리 ──────────────────────────────────────────────────
def write_gj_id_result(parsed: dict) -> bytes:
    """
    parsed: {compound: {"std": [...], "sp": [...], ...}} from csv_utils.merge_parsed
    반환: Excel bytes
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("공진단_확인결과")
    _set_widths(ws, 9)

    row = 1

    # ── ● 시스템적합성 ────────────────────────────────────────────
    _title_bar(ws, row, "● 시스템적합성", 1, 10)
    row += 1
    _cell(ws, row, 1, "- 표준액 6 회 조작한 결과로 확인한다.", fill=None, bold=False, fs=9)
    _merge(ws, row, 1, row, 10)
    row += 1
    _cell(ws, row, 1, "- 유지시간 RSD% 2.0% 이하이다.", fill=None, bold=False, fs=9)
    _merge(ws, row, 1, row, 10)
    row += 1

    row = _write_sst_group(ws, row, SST_GROUP1, parsed)
    row += 1
    row = _write_sst_group(ws, row, SST_GROUP2, parsed)
    row += 1

    # ── ● 결과 확인 ───────────────────────────────────────────────
    _title_bar(ws, row, "● 결과 확인", 1, 10)
    row += 1

    row = _write_result_group(ws, row, RESULT_GROUP1, parsed)
    row += 1
    row = _write_result_group(ws, row, RESULT_GROUP2, parsed)
    row += 1
    row = _write_result_group(ws, row, RESULT_GROUP3, parsed)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
