# -*- coding: utf-8 -*-
"""
Agilent Quant Summary Report PDF 파서 - 확인(ID) 분석용
SST PDF / SP_A/B/C PDF → 샘플별 화합물 데이터 추출
"""

import re
import pdfplumber
from collections import defaultdict


def _extract_text(pdf_path: str) -> str:
    """Summary 페이지만 읽음 — 'Quant Sample Report' 개별 샘플 페이지 나오면 중단"""
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if not t:
                continue
            # 개별 샘플 크로마토그램 페이지 시작 → 중단
            if "Quant Sample Report" in t and "Quantitative Analysis Results Summary" not in t:
                break
            parts.append(t)
    return "\n".join(parts)


def _parse_summary_blocks(text: str) -> dict:
    """
    Summary Report 텍스트에서 compound_transition별 샘플 데이터 추출
    반환: {compound_transition: [{name, rt, resp, sn}]}
    """
    result = {}
    current_compound = None
    in_data = False

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        # compound 헤더 감지: 숫자로 끝나는 라인 (e.g. "Platycodin-D_1223.7", "Amygdalin")
        # 헤더 다음 줄은 "Data File Type RT Resp. S/N RRT"
        if re.match(r'^[A-Za-z].*[A-Za-z0-9\-\_\s]+$', line) and not re.search(r'\d{4,}', line):
            # "Data File" 헤더 줄 건너뜀
            if line.startswith("Data File") or line.startswith("Batch") or line.startswith("Analysis") \
               or line.startswith("Report") or line.startswith("Calibration") or line.startswith("Acq.") \
               or line.startswith("Type") or line.startswith("Sequence") or line.startswith("Page") \
               or line.startswith("Generated") or line.startswith("Name") or line.startswith("Analyze"):
                in_data = False
                continue

            # compound 이름처럼 보이는 경우
            if re.search(r'(_\d+\.?\d*$)|^(Amygdalin|Ginsenoside|Gingerol|Bilirubin|Platycodin|BoRy|'
                         r'Prim-O|Saikosaponin|Atractylenolide|Ligustilide|Paeoniflorin|Baicalin|'
                         r'Glycyrrhizin|Decursin|SY|PH)', line):
                current_compound = line
                result[current_compound] = []
                in_data = False
                continue

        # "Data File Type RT ..." 헤더 줄
        if line.startswith("Data File") and "RT" in line:
            in_data = True
            continue

        # 데이터 행: filename.d Sample RT Resp S/N (RRT)
        if in_data and current_compound and ".d" in line:
            parts = line.split()
            if len(parts) >= 5:
                try:
                    # parts: [filename.d, Type/Name, RT, Resp, S/N, (RRT)]
                    # 파일명에 공백이 없으므로 parts[0]이 filename
                    fname = parts[0]
                    sample_name = parts[1]
                    rt = float(parts[2])
                    resp = float(parts[3])
                    sn = float(parts[4])
                    rrt = float(parts[5]) if len(parts) > 5 else None
                    result[current_compound].append({
                        "file": fname,
                        "name": sample_name,
                        "rt": rt,
                        "resp": resp,
                        "sn": sn,
                        "rrt": rrt,
                    })
                except (ValueError, IndexError):
                    pass

    return result


def parse_id_pdf(pdf_path: str) -> dict:
    """
    ID PDF (SST 또는 SP) 파싱
    반환: {sample_name: {compound_transition: {rt, resp, sn}}}
    """
    text = _extract_text(pdf_path)
    blocks = _parse_summary_blocks(text)

    # sample 중심으로 재구성
    samples = defaultdict(dict)
    for compound, rows in blocks.items():
        for row in rows:
            samples[row["name"]][compound] = {
                "rt": row["rt"],
                "resp": row["resp"],
                "sn": row["sn"],
                "file": row["file"],
            }

    return dict(samples)


def parse_sst_pdf(pdf_path: str) -> dict:
    """SST PDF 파싱 - SST-1/2/3 런만 반환"""
    data = parse_id_pdf(pdf_path)
    sst_runs = {k: v for k, v in data.items() if k.startswith("ID_SST")}
    return sst_runs


def compute_rrt(sample_data: dict, reference_compound_prefix: str) -> dict:
    """
    sample_data: {compound_transition: {rt, ...}}
    reference_compound_prefix: e.g. "Ginsenoside Rb1"
    → 각 compound_transition에 rrt 추가한 dict 반환
    """
    # 기준 compound의 RT 찾기 (첫번째 매칭 transition)
    ref_rt = None
    for key, val in sample_data.items():
        if key.startswith(reference_compound_prefix):
            ref_rt = val["rt"]
            break

    if ref_rt is None or ref_rt == 0:
        return sample_data

    result = {}
    for key, val in sample_data.items():
        result[key] = dict(val)
        result[key]["rrt_calc"] = round(val["rt"] / ref_rt, 3)

    return result


def judge_rrt(rrt_calc: float, rrt_expected: float, tolerance: float = 0.10) -> str:
    """RRT 적합 판정"""
    lower = rrt_expected * (1 - tolerance)
    upper = rrt_expected * (1 + tolerance)
    return "적합" if lower <= rrt_calc <= upper else "부적합"
