# -*- coding: utf-8 -*-
import sys, json
from pathlib import Path
import streamlit as st
import pandas as pd
import tempfile, os

sys.path.insert(0, str(Path(__file__).parent.parent))
from core.id_parser import parse_id_pdf, compute_rrt, judge_rrt
from core.excel_writer import write_id_result

CONFIG_DIR = Path(__file__).parent.parent / "config"

st.set_page_config(page_title="확인(ID) 분석", layout="wide")
st.title("확인 (ID) 분석")
st.markdown("---")

# ── 제품 선택 ──────────────────────────────────────────────────
product_name = st.selectbox("제품", ["우황청심원 환제", "우황청심원 현탁제"])

form_type = "환" if "환제" in product_name else "현"
config_file = CONFIG_DIR / ("csw_pill.json" if form_type == "환" else "csw_hyeon.json")
with open(config_file, encoding="utf-8") as f:
    config = json.load(f)

st.info(f"선택된 제품: **{product_name}** | RRT 허용범위: ±{int(config['id_rrt_tolerance']*100)}%")

# ── PDF 업로드 ──────────────────────────────────────────────────
st.subheader("PDF 파일 업로드")
col1, col2, col3, col4 = st.columns(4)
with col1:
    sst_file  = st.file_uploader("SST.pdf",  type="pdf", key="sst")
with col2:
    spa_file  = st.file_uploader("SP_A.pdf", type="pdf", key="spa")
with col3:
    spb_file  = st.file_uploader("SP_B.pdf", type="pdf", key="spb")
with col4:
    spc_file  = st.file_uploader("SP_C.pdf", type="pdf", key="spc")

if st.button("분석 실행", type="primary", disabled=not any([spa_file, spb_file, spc_file])):
    compounds = config["compounds_id"]
    ref_compound = config["id_reference_compound"]
    tolerance    = config["id_rrt_tolerance"]

    # 모든 SP 데이터 합치기
    all_sample_data = {}

    for label, upfile in [("A", spa_file), ("B", spb_file), ("C", spc_file)]:
        if upfile is None:
            continue
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(upfile.read())
            tmp_path = tmp.name
        try:
            parsed = parse_id_pdf(tmp_path)
            for sample_name, sample_data in parsed.items():
                enriched = compute_rrt(sample_data, ref_compound)
                all_sample_data[sample_name] = enriched
        finally:
            os.unlink(tmp_path)

    if not all_sample_data:
        st.error("PDF에서 데이터를 추출하지 못했습니다. PDF 형식을 확인해주세요.")
        st.stop()

    sample_names = list(all_sample_data.keys())
    st.success(f"데이터 추출 완료: {len(sample_names)}개 시료 — {', '.join(sample_names)}")

    # ── 결과 테이블 생성 ──────────────────────────────────────────
    result_rows = []

    for comp_cfg in compounds:
        comp_name   = comp_cfg["name"]
        transitions = comp_cfg["transitions"]
        rrt_exp     = comp_cfg["rrt_expected"]
        is_ref      = comp_cfg.get("is_reference", False)

        sample_results = {}
        for sn in sample_names:
            sdata = all_sample_data[sn]

            # 이 compound의 첫번째 transition 데이터 찾기
            matched_key = None
            for key in sdata:
                if key.startswith(comp_name):
                    matched_key = key
                    break

            if matched_key:
                entry = sdata[matched_key]
                rrt_c = entry.get("rrt_calc")
                result_str = judge_rrt(rrt_c, rrt_exp, tolerance) if rrt_c else "N/A"
                sample_results[sn] = {
                    "rt": entry.get("rt"),
                    "sn": entry.get("sn"),
                    "rrt_calc": rrt_c,
                    "result": "기준" if is_ref else result_str,
                }
            else:
                sample_results[sn] = {"rt": None, "sn": None, "rrt_calc": None, "result": "미검출"}

        result_rows.append({
            "compound": comp_name,
            "transitions": transitions,
            "rrt_expected": rrt_exp,
            "samples": sample_results,
        })

    # ── 화면 표시 ──────────────────────────────────────────────────
    st.subheader("분석 결과")

    for row in result_rows:
        cols_header = ["성분", "RRT기준"] + [f"{sn}\nRT | S/N | RRT | 결과" for sn in sample_names]
        display = {"성분": row["compound"], "RRT기준": row["rrt_expected"]}
        for sn in sample_names:
            s = row["samples"][sn]
            rt_v  = f"{s['rt']:.3f}"   if s["rt"]  else "-"
            sn_v  = f"{s['sn']:.1f}"   if s["sn"]  else "-"
            rrt_v = f"{s['rrt_calc']:.3f}" if s["rrt_calc"] else "-"
            display[sn] = f"RT:{rt_v} | S/N:{sn_v} | RRT:{rrt_v} | {s['result']}"

        df = pd.DataFrame([display])
        st.dataframe(df, use_container_width=True, hide_index=True)

    # ── 요약 테이블 (결과만) ───────────────────────────────────────
    st.subheader("적합/부적합 요약")
    summary_rows = []
    for row in result_rows:
        r = {"성분": row["compound"], "RRT 기준": row["rrt_expected"]}
        for sn in sample_names:
            r[sn] = row["samples"][sn].get("result", "-")
        summary_rows.append(r)

    df_summary = pd.DataFrame(summary_rows)

    def color_result(val):
        if val == "적합" or val == "기준":
            return "background-color: #C6EFCE"
        elif val == "부적합":
            return "background-color: #FFC7CE"
        return ""

    styled = df_summary.style.applymap(color_result, subset=sample_names)
    st.dataframe(styled, use_container_width=True, hide_index=True)

    # 부적합 경고
    fail_list = [
        f"{r['성분']} ({sn})"
        for r in summary_rows
        for sn in sample_names
        if r.get(sn) == "부적합"
    ]
    if fail_list:
        st.warning(f"부적합 항목: {', '.join(fail_list)}")
    else:
        st.success("전 항목 적합")

    # ── 엑셀 다운로드 ──────────────────────────────────────────────
    excel_bytes = write_id_result(result_rows, lot_info=product_name)
    st.download_button(
        label="결과 엑셀 다운로드",
        data=excel_bytes,
        file_name=f"{product_name}_확인결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
