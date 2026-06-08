# -*- coding: utf-8 -*-
"""
Agilent Quant Summary Report PDF 파서 - 함량(AS) 분석용
AS Summary PDF → STD 1~6, IS, 검액 데이터 추출
"""

import re
import pdfplumber
from collections import defaultdict


def _extract_text(pdf_path: str) -> str:
    """Summary 페이지만 추출 — 개별 Quant Sample Report 페이지는 건너뜀"""
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if not t:
                continue
            # 개별 샘플 크로마토그램 페이지는 제외
            if "Quant Sample Report" in t and "Quantitative Analysis Results Summary" not in t:
                continue
            parts.append(t)
    return "\n".join(parts)


# 페이지 헤더/푸터로 나타나는 무시할 패턴
_PAGE_HEADER_RE = re.compile(
    r'^(Quantitative Analysis|Batch Data Path|Analysis Time|Report Generation|'
    r'Calibration Last|Analyze Quant|Report Quant|Page \d|Generated at|'
    r'Sequence Table|Data File\s+Name|Acq\.|Vol\.|Level)'
)

# AS에서 나타나는 화합물 이름 패턴
_COMPOUND_RE = re.compile(
    r'^(Amygdalin|Paeoniflorin|Cinnamic acid|Ginsenoside Rb1|Glycyrrhizin|'
    r'Decursin|Cinnamic_acid|[A-Za-z][\w\-\s]+_\d+\.?\d*)$'
)


def _parse_as_blocks(text: str) -> dict:
    """
    화합물별 데이터 파싱
    반환: {compound: {"std": [...], "sp": [...], "stability": [...], "system_check": [...]}}
    """
    result = defaultdict(lambda: {"std": [], "sp": [], "stability": [], "system_check": []})
    current_compound = None
    in_data = False

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # 페이지 헤더/메타 라인 → 건너뜀 (in_data 유지)
        if _PAGE_HEADER_RE.match(line):
            continue

        # "Data File  Type  RT  Resp.  S/N  RRT" 헤더
        if line.startswith("Data File") and "RT" in line:
            in_data = True
            continue

        # 화합물 이름 라인
        if _COMPOUND_RE.match(line) and ".d" not in line:
            current_compound = line.strip()
            in_data = False
            continue

        # 데이터 행: filename.d 로 시작
        if current_compound and ".d" in line:
            parts = line.split()
            if len(parts) < 4:
                continue  # RT/Resp 없는 빈 행 건너뜀
            try:
                fname       = parts[0]
                sample_name = parts[1]
                rt          = float(parts[2])
                resp        = float(parts[3])
                sn          = float(parts[4]) if len(parts) > 4 else None
            except (ValueError, IndexError):
                continue

            entry = {
                "file": fname,
                "name": sample_name,
                "rt":   rt,
                "resp": resp,
                "sn":   sn,
            }

            # 파일명 기준으로 분류 (parts[1]은 항상 "Sample"이라 사용 불가)
            fname_up = fname.upper()
            if fname_up.startswith("STD"):
                result[current_compound]["std"].append(entry)
            elif fname_up.startswith("STABILITY"):
                result[current_compound]["stability"].append(entry)
            elif fname_up.startswith("SYSTEM"):
                result[current_compound]["system_check"].append(entry)
            else:
                result[current_compound]["sp"].append(entry)

    return dict(result)


def parse_as_summary(pdf_path: str) -> dict:
    text = _extract_text(pdf_path)
    return _parse_as_blocks(text)


def get_std_avg_response(compound_data: dict) -> dict:
    std_rows = compound_data.get("std", [])
    if not std_rows:
        return {}
    rts   = [r["rt"]   for r in std_rows]
    resps = [r["resp"] for r in std_rows]
    return {
        "count":     len(std_rows),
        "rt_avg":    round(sum(rts)   / len(rts),   3),
        "resp_avg":  round(sum(resps) / len(resps), 1),
        "resp_list": resps,
        "rt_list":   rts,
    }
