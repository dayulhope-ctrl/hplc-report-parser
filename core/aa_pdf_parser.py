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


def _to_float(tok: str):
    try:
        return float(tok)
    except (ValueError, TypeError):
        return None


def _parse_data_line(line: str):
    """
    크로마토그램 데이터 라인 파싱 (양식 독립적).

    ClarityAmino 리포트 양식이 수시로 바뀌어 컬럼 순서가 제각각이므로
    위치(index) 대신 값의 특성으로 파싱한다:
      1) 화합물명은 COMPOUND_MAP에서 탐색 (가장 긴 매칭 우선)
      2) 면적(Area/Response)은 라인 내 최대 수치 (peak area가 항상 가장 큼)
      3) 분리도(Resolution)는 화합물명 위치·주변 수치 배치로 추정
         - 화합물 앞에 큰 수치(>100)가 있으면(면적-먼저/Response-먼저 양식):
           화합물 뒤 첫 숫자 = Resolution (숫자 2개 이상일 때만)
         - 화합물 앞에 큰 수치가 없으면(화합물-먼저 양식):
           화합물 뒤 마지막 숫자 = Resolution (숫자 4개 이상일 때만)
    반환: (aa_name, area, resolution) 또는 None
    """
    # 가장 긴 화합물명 우선 매칭 (부분 매칭 방지)
    matched = None
    for eng_name, kor_name in COMPOUND_MAP.items():
        if eng_name in line and (matched is None or len(eng_name) > len(matched[0])):
            matched = (eng_name, kor_name)
    if not matched:
        return None
    eng_name, kor_name = matched

    tokens = line.split()
    nums = [v for v in (_to_float(t) for t in tokens) if v is not None]
    if len(nums) < 2:
        return None

    # 면적: 라인 내 최대 수치 (peak area/response가 항상 최대)
    area = max(nums)

    # 화합물명 토큰 위치
    name_tokens = eng_name.split()
    comp_idx = None
    for i in range(len(tokens) - len(name_tokens) + 1):
        if tokens[i: i + len(name_tokens)] == name_tokens:
            comp_idx = i
            break
    if comp_idx is None:
        return (kor_name, area, None)

    before_nums = [v for v in (_to_float(t) for t in tokens[:comp_idx]) if v is not None]
    after_nums  = [v for v in (_to_float(t) for t in tokens[comp_idx + len(name_tokens):]) if v is not None]

    big_before = any(v > 100 for v in before_nums)
    if big_before:
        # 면적/Response가 화합물 앞 → 뒤엔 [Resolution, Symmetry] 또는 [Symmetry]
        resolution = after_nums[0] if len(after_nums) >= 2 else None
    else:
        # 화합물-먼저 양식 → 뒤엔 [Response, Area%, Symmetry, (Resolution)]
        resolution = after_nums[-1] if len(after_nums) >= 4 else None

    return (kor_name, area, resolution)


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
