# -*- coding: utf-8 -*-
"""
환제/현탁액 확인시험 CSV 파서
SP_A / SP_B / SP_C / SST CSV

공통 헤더 구조 (2줄):
  Row 0: Sample,,,,,,,, CompoundName Results,, ...
  Row 1: ,,Name,Data File,Type,Level,Acq. Date-Time,RT,S/N, RT,S/N, ...
  Row 2+: data

NAME_COL = 2, DATA_FILE_COL = 3
"""

import re
from core.csv_utils import read_rows, safe_float, detect_columns


DATA_FILE_COL = 3


def _read_rows(file_obj):
    return read_rows(file_obj)


def _safe_float(val):
    return safe_float(val)


def _parse_compound_headers(rows):
    """
    csv_utils.detect_columns 를 활용해 자동 감지 (S/N 우선).
    반환: [{"name": str, "rt_col": int, "sn_col": int}, ...]
    """
    _, compound_cols = detect_columns(rows, prefer="sn")
    return [{"name": name, "rt_col": rt, "sn_col": sn}
            for name, (rt, sn) in compound_cols.items()]


def _extract_lot(data_file: str) -> str:
    """
    Data File 이름에서 lot명 추출.
    'BYHyun_23026-C.d' → 'BYHyun_23026'
    'SST-1.d'           → 'SST-1'
    """
    stem = re.sub(r'\.d$', '', data_file.strip(), flags=re.IGNORECASE)
    stem = re.sub(r'-[A-Ca-c]$', '', stem)
    return stem


def parse_sp_csv(file_obj):
    """
    SP_A / SP_B / SP_C CSV 파싱.
    lot_name 및 sample_name은 Data File 컬럼에서 추출.
    반환: {
        "sample_names": [str, ...],
        "lot_name": str,
        "transitions": [...],       ← 첫 번째 샘플 (하위 호환)
        "all_transitions": [[...]], ← 모든 샘플 행
    }
    """
    rows = _read_rows(file_obj)
    if len(rows) < 3:
        return {}

    name_col, _ = detect_columns(rows, prefer="sn")
    compounds = _parse_compound_headers(rows)

    lot_name = None
    sample_names = []
    all_transitions = []

    for row in rows[2:]:
        name      = row[name_col].strip()      if name_col      < len(row) else ""
        data_file = row[DATA_FILE_COL].strip() if DATA_FILE_COL < len(row) else ""
        if not name and not data_file:
            continue

        transitions = []
        for c in compounds:
            rt = _safe_float(row[c["rt_col"]]) if c["rt_col"] < len(row) else None
            sn = _safe_float(row[c["sn_col"]]) if c["sn_col"] < len(row) else None
            transitions.append({"name": c["name"], "rt": rt, "sn": sn})

        lot = _extract_lot(data_file) if data_file else None
        # 환제 패턴 'ID_<lot>_<A|B|C>' 처리
        m = re.match(r'^ID_(.+?)_([ABC])$', name)
        if m:
            lot = m.group(1)

        if lot_name is None and lot:
            lot_name = lot

        label = data_file if data_file else name
        sample_names.append(label)
        all_transitions.append(transitions)

    if not all_transitions:
        return {}

    return {
        "sample_name":    sample_names[0],
        "sample_names":   sample_names,
        "lot_name":       lot_name or "",
        "transitions":    all_transitions[0],
        "all_transitions": all_transitions,
    }


def parse_sst_csv(file_obj):
    """
    SST / Stability CSV 파싱.
    Data File 컬럼으로 SST/Stability 분류:
      - 'STABILITY' in data_file → stability
      - 그 외 → sst
    반환: {
        "compound_names": [str, ...],
        "stability": [{comp_name: {"rt", "sn", "label"}}, ...],
        "sst":       [{comp_name: {"rt", "sn", "label"}}, ...],
    }
    """
    rows = _read_rows(file_obj)
    if len(rows) < 3:
        return {}

    name_col, _ = detect_columns(rows, prefer="sn")
    compounds = _parse_compound_headers(rows)
    comp_names = [c["name"] for c in compounds]

    stability_rows = []
    sst_rows = []

    for row in rows[2:]:
        name      = row[name_col].strip()      if name_col      < len(row) else ""
        data_file = row[DATA_FILE_COL].strip() if DATA_FILE_COL < len(row) else ""
        if not name and not data_file:
            continue

        row_data = {}
        for c in compounds:
            rt = _safe_float(row[c["rt_col"]]) if c["rt_col"] < len(row) else None
            sn = _safe_float(row[c["sn_col"]]) if c["sn_col"] < len(row) else None
            label = data_file if data_file else name
            row_data[c["name"]] = {"rt": rt, "sn": sn, "label": label}

        combined = (name + " " + data_file).upper()
        if "STABILITY" in combined:
            stability_rows.append(row_data)
        else:
            sst_rows.append(row_data)

    return {
        "compound_names": comp_names,
        "stability":      stability_rows,
        "sst":            sst_rows,
    }
