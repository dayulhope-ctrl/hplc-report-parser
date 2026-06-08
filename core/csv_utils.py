# -*- coding: utf-8 -*-
"""
CSV 파싱 공통 유틸리티

모든 MassHunter CSV 포맷에 대해 헤더를 자동 감지합니다.
- NAME_COL  : Row 1에서 "Name" 셀 위치 자동 탐색
- 화합물 컬럼: Row 0에서 "X Results" 위치 → Row 1로 RT/Area(S/N) 컬럼 확인
포맷이 변경돼도 헤더 키워드(Results, Name, RT, Area, S/N)만 유지되면 동작합니다.
"""

import csv, re, math
from collections import defaultdict


# ── 기본 I/O ─────────────────────────────────────────────────────────────────
def read_rows(file_obj) -> list:
    if hasattr(file_obj, "read"):
        content = file_obj.read()
        if isinstance(content, bytes):
            content = content.decode("utf-8-sig", errors="replace")
        lines = content.splitlines()
    else:
        with open(file_obj, encoding="utf-8-sig", errors="replace") as f:
            lines = f.read().splitlines()
    return list(csv.reader(lines))


def safe_float(val: str):
    try:
        v = val.strip()
        return float(v) if v else None
    except (ValueError, AttributeError):
        return None


# ── 헤더 자동 감지 ────────────────────────────────────────────────────────────
def detect_columns(rows: list, prefer: str = "area") -> tuple:
    """
    Row 0 (화합물 헤더), Row 1 (서브헤더)를 스캔해
    name_col 과 {compound: (rt_col, value_col)} 을 반환.

    prefer="area"  → Area 우선, 없으면 S/N (AS 함량용)
    prefer="sn"    → S/N 우선, 없으면 Area (ID 확인용)
    """
    if len(rows) < 2:
        return 0, {}

    header0 = rows[0]
    header1 = rows[1]

    # NAME_COL: row 1에서 "name" 텍스트 위치
    name_col = 0
    for ci, val in enumerate(header1):
        if val.strip().lower() == "name":
            name_col = ci
            break

    # 화합물 컬럼: row 0에서 "Results" 포함 셀 탐색
    compound_cols = {}
    for ci, val in enumerate(header0):
        if "Results" not in val:
            continue
        comp = val.replace("Results", "").strip()
        if not comp:
            continue

        rt_col = area_col = sn_col = None
        for offset in range(5):
            idx = ci + offset
            if idx >= len(header1):
                break
            sub = header1[idx].strip().upper()
            if sub == "RT" and rt_col is None:
                rt_col = idx
            elif sub == "AREA" and area_col is None:
                area_col = idx
            elif sub == "S/N" and sn_col is None:
                sn_col = idx

        if rt_col is None:
            continue

        # prefer에 따라 value 컬럼 결정
        if prefer == "area":
            val_col = area_col if area_col is not None else (sn_col if sn_col is not None else rt_col + 1)
        else:  # "sn"
            val_col = sn_col if sn_col is not None else (area_col if area_col is not None else rt_col + 1)

        compound_cols[comp] = (rt_col, val_col)

    return name_col, compound_cols


# ── 행 분류 ──────────────────────────────────────────────────────────────────
def classify(name: str) -> str:
    n = name.upper()
    if "SYSTEM" in n:     return "system_check"
    if "STABILITY" in n:  return "stability"
    if "STD" in n:        return "std"
    return "sp"


# ── 범용 CSV 파싱 ─────────────────────────────────────────────────────────────
def parse_csv(file_obj, prefer: str = "area") -> dict:
    """
    범용 파싱. 반환:
    {
      compound: {
        "std":          [{"name","rt","area"}, ...],
        "sp":           [...],
        "stability":    [...],
        "system_check": [...],
      }
    }
    """
    rows = read_rows(file_obj)
    if len(rows) < 3:
        return {}

    name_col, compound_cols = detect_columns(rows, prefer=prefer)
    if not compound_cols:
        return {}

    result = defaultdict(lambda: {"std": [], "sp": [], "stability": [], "system_check": []})

    for row in rows[2:]:
        if len(row) < 2:
            continue
        name = row[name_col].strip() if name_col < len(row) else ""
        if not name:
            continue

        section = classify(name)

        for comp, (rt_col, val_col) in compound_cols.items():
            rt   = safe_float(row[rt_col])  if rt_col  < len(row) else None
            area = safe_float(row[val_col]) if val_col < len(row) else None
            if rt is None and area is None:
                continue
            result[comp][section].append({"name": name, "rt": rt, "area": area})

    return dict(result)


# ── SP 런 lot 그룹핑 ──────────────────────────────────────────────────────────
def group_by_lot(parsed: dict) -> dict:
    """
    SP 런을 lot별·A/B 메서드별로 그룹핑.
    샘플명 패턴: <LotName>_<A|B>-<num>  (예: BSHwan_26006_B-1)
    반환: {lot_name: {"A": [{comp: area, ...}, ...], "B": [...]}}
    """
    run_map = {}
    for comp, data in parsed.items():
        for entry in data.get("sp", []):
            name = entry["name"]
            if name not in run_map:
                run_map[name] = {"name": name}
            run_map[name][comp] = entry.get("area")

    lots = {}
    for run_name, run_data in run_map.items():
        m = re.search(r'^(.+?)_([AB])-(\d+)$', run_name)
        if not m:
            continue
        lot_name = m.group(1)
        method   = m.group(2)
        run_num  = int(m.group(3))

        if lot_name not in lots:
            lots[lot_name] = {"A": [], "B": []}

        existing = lots[lot_name][method]
        while len(existing) < run_num:
            existing.append({})
        existing[run_num - 1] = run_data

    return lots


# ── 여러 파일 병합 ────────────────────────────────────────────────────────────
def merge_parsed(file_list, prefer: str = "area") -> dict:
    """
    여러 CSV 파일(또는 파일 객체 리스트)을 각각 파싱한 뒤 결과를 병합.
    파일이 1개면 그냥 parse_csv 결과를 반환.
    """
    from collections import defaultdict
    merged = defaultdict(lambda: {"std": [], "sp": [], "stability": [], "system_check": []})
    for f in file_list:
        partial = parse_csv(f, prefer=prefer)
        for comp, sections in partial.items():
            for sec, rows in sections.items():
                merged[comp][sec].extend(rows)
    return dict(merged)


# ── STD 통계 ─────────────────────────────────────────────────────────────────
def get_stats(compound_data: dict) -> dict:
    """STD 런 통계 (평균·표준편차·%RSD)"""
    std_rows = compound_data.get("std", [])
    if not std_rows:
        return {}
    areas = [r["area"] for r in std_rows if r["area"] is not None]
    rts   = [r["rt"]   for r in std_rows if r["rt"]   is not None]
    if not areas:
        return {}
    avg = sum(areas) / len(areas)
    sd  = math.sqrt(sum((x - avg) ** 2 for x in areas) / (len(areas) - 1)) if len(areas) > 1 else 0
    rsd = round(sd / avg * 100, 3) if avg else 0
    return {
        "count":     len(areas),
        "area_list": areas,
        "rt_list":   rts,
        "avg_area":  round(avg, 1),
        "avg_rt":    round(sum(rts) / len(rts), 3) if rts else None,
        "sd":        round(sd, 2),
        "rsd":       rsd,
    }
