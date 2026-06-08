# -*- coding: utf-8 -*-
"""
Sykam ClarityAmino 크로마토그램 PDF 파서
- parse_std(file_obj) → (runs, resolutions)
- parse_sp(file_obj)  → lots dict
"""
import re
import pdfplumber

AA_ORDER = [
    "L-Asp", "L-Thr", "L-Ser", "L-Glu", "L-Pro",
    "Gly",   "L-Ala", "L-Val", "L-leu", "L-Lys", "L-Arg",
]

# ClarityAmino 영문명 → 표기명 매핑
COMPOUND_MAP = {
    "Aspartic acid": "L-Asp",
    "Threonine":     "L-Thr",
    "Serine":        "L-Ser",
    "Glutamic acid": "L-Glu",
    "Proline":       "L-Pro",
    "Glycine":       "Gly",
    "Alanine":       "L-Ala",
    "Valine":        "L-Val",
    "Leucine":       "L-leu",
    "Lysine":        "L-Lys",
    "Arginine":      "L-Arg",
}


def _extract_text(file_obj) -> str:
    with pdfplumber.open(file_obj) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def _parse_data_line(line: str):
    """
    크로마토그램 데이터 라인 파싱.
    형식: peak_no  signal  rt  area  height  area%  compound_name  [resolution]  symmetry
    반환: (aa_name, area, resolution) 또는 None
    """
    for eng_name, kor_name in COMPOUND_MAP.items():
        if eng_name not in line:
            continue
        tokens = line.split()
        try:
            area = float(tokens[3])
        except (IndexError, ValueError):
            return None

        name_tokens = eng_name.split()
        for i in range(len(tokens) - len(name_tokens) + 1):
            if tokens[i: i + len(name_tokens)] == name_tokens:
                after = tokens[i + len(name_tokens):]
                # 뒤에 숫자가 2개 이상 → 첫 번째가 Resolution, 그 다음이 Symmetry
                # 숫자가 1개만   → Symmetry만 (Resolution = N/A)
                try:
                    resolution = float(after[0]) if len(after) >= 2 else None
                except ValueError:
                    resolution = None
                return (kor_name, area, resolution)
    return None


def _table_sample_id(block_head: str) -> str | None:
    """'All Signals Result Table' 이후 텍스트에서 샘플 ID 추출."""
    # 경로 구분자가 있을 수도 없을 수도 있으므로 optional
    m = re.search(r'\(ESTD\s*-\s*(?:\S+[/\\])?([^)]+)\)', block_head)
    return m.group(1).strip() if m else None


def parse_std(file_obj):
    """
    STD PDF 파싱.

    Returns
    -------
    runs : list[dict]   길이 6, 각 dict = {aa_name: area}
    resolutions : list[dict]  길이 6, 각 dict = {aa_name: float|None}
    """
    text = _extract_text(file_obj)

    # sample_id별로 areas/resolutions 누적 (페이지 경계에서 블록이 쪼개지는 경우 병합)
    merged: dict = {}   # sample_id → {"areas": {...}, "ress": {...}}

    for block in re.split(r"All Signals Result Table", text)[1:]:
        sample_id = _table_sample_id(block)
        if not sample_id:
            continue
        if "_STD_" not in sample_id.upper() and "STD" not in sample_id.upper():
            continue

        entry = merged.setdefault(sample_id, {"areas": {}, "ress": {}})
        for line in block.splitlines():
            result = _parse_data_line(line)
            if result:
                aa, area, res = result
                entry["areas"].setdefault(aa, area)   # 먼저 나온 값 우선
                entry["ress"].setdefault(aa, res)

    runs, resolutions = [], []
    for entry in merged.values():
        if len(entry["areas"]) >= 5:
            runs.append(entry["areas"])
            resolutions.append(entry["ress"])

    return runs, resolutions


def parse_sp(file_obj, debug: bool = False):
    """
    SP(검액) PDF 파싱.

    Returns
    -------
    lots : dict
        { lot_id: { sample_num: {aa_name: area} } }
        예) {"26001A": {1: {...}, 2: {...}}, "26001B": {...}}
    debug_info : list[str]  (debug=True일 때만 반환, 아니면 빈 리스트)
    """
    text = _extract_text(file_obj)
    lots: dict = {}
    dbg: list[str] = []

    blocks = re.split(r"All Signals Result Table", text)
    if debug:
        dbg.append(f"블록 수(헤더 제외): {len(blocks)-1}")

    # (lot_id, sample_num) → areas 누적 dict (페이지 경계 블록 분리 대응)
    merged: dict = {}

    for bi, block in enumerate(blocks[1:], 1):
        sample_id = _table_sample_id(block)
        if debug:
            dbg.append(f"[블록{bi}] sample_id={repr(sample_id)}")
        if not sample_id:
            continue

        m = re.search(r"_(\d+[A-Za-z]*)-(\d+)[_\s]", sample_id)
        if not m:
            if debug:
                dbg.append(f"  → lot 패턴 불일치: {repr(sample_id)}")
            continue

        lot_id     = m.group(1)
        sample_num = int(m.group(2))
        key        = (lot_id, sample_num)

        entry = merged.setdefault(key, {})
        for line in block.splitlines():
            result = _parse_data_line(line)
            if result:
                aa, area, _ = result
                entry.setdefault(aa, area)   # 먼저 나온 값 우선

        if debug:
            dbg.append(f"  → lot={lot_id} sample={sample_num} 누적AA수={len(entry)}")

    for (lot_id, sample_num), areas in merged.items():
        if len(areas) >= 5:
            lots.setdefault(lot_id, {})[sample_num] = areas
        elif debug:
            dbg.append(f"  → {lot_id}-{sample_num} AA 5개 미만으로 스킵")

    return (lots, dbg) if debug else lots
