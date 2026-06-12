# -*- coding: utf-8 -*-
import sys, json, os, tempfile
from pathlib import Path
import streamlit as st
import pandas as pd

BASE_DIR   = Path(__file__).parent
CONFIG_DIR = BASE_DIR / "config"

sys.path.insert(0, str(BASE_DIR))

st.set_page_config(page_title="면적값 자동파싱 시스템", layout="wide", initial_sidebar_state="collapsed")

# ══════════════════════════════════════════════════════════════════
# 분석 함수들
# ══════════════════════════════════════════════════════════════════

def _run_id_csv(product, sst_file, spa_file, spb_file, spc_file):
    from core.id_csv_parser import parse_sp_csv, parse_sst_csv
    from core.id_csv_excel_writer import write_id_csv_result

    with st.spinner("CSV 파싱 중..."):
        sst_data = parse_sst_csv(sst_file) if sst_file else {}
        sp_a     = parse_sp_csv(spa_file)  if spa_file else {}
        sp_b     = parse_sp_csv(spb_file)  if spb_file else {}
        sp_c     = parse_sp_csv(spc_file)  if spc_file else {}

    lot = sp_a.get("lot_name") or sp_b.get("lot_name") or sp_c.get("lot_name") or "결과"
    st.success(f"파싱 완료 — Lot: {lot}")

    from core.id_csv_excel_writer import HYEON_SP_ORDER
    order = HYEON_SP_ORDER if product == "현탁제" else None
    excel_bytes = write_id_csv_result(sst_data, sp_a, sp_b, sp_c, compound_order=order)
    st.download_button("📥 결과 엑셀 다운로드", data=excel_bytes,
                       file_name=f"{product}_확인결과_{lot}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)


def _run_as(product, as_files):
    from core.csv_utils import merge_parsed
    if product == "현탁제":
        from core.as_hyeon_csv_parser import group_hyeon_sp_by_lot as _group, get_hyeon_std_stats as _stats
        from core.as_hyeon_excel_writer import write_hyeon_result as _write
    else:
        from core.as_csv_parser import group_sp_by_lot as _group, get_std_stats as _stats
        from core.as_excel_writer import write_as_result as _write

    with st.spinner("CSV 파싱 중..."):
        parsed     = merge_parsed(as_files, prefer="area")
        lot_groups = _group(parsed)

    if not parsed:
        st.error("CSV에서 데이터를 추출하지 못했습니다."); return

    lot_names = list(lot_groups.keys())
    st.success(f"추출 완료 — STD 6런 / SP Lot: {', '.join(lot_names)}")

    preview = []
    for comp, cdata in parsed.items():
        stats = _stats(cdata)
        preview.append({
            "성분": comp,
            "STD 런": stats.get("count", 0),
            "STD 평균 Area": stats.get("avg_area"),
            "%RSD": stats.get("rsd"),
        })
    st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

    excel_bytes = _write(parsed, lot_groups)
    st.download_button("📥 결과 엑셀 다운로드", data=excel_bytes,
                       file_name=f"{product}_함량결과.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)


def _run_gj(gj_files, test_type):
    """공진단 확인 또는 함량 분석 (동일 CSV 파일 사용)."""
    from core.csv_utils import merge_parsed
    from core.gj_csv_parser import get_gj_std_stats

    with st.spinner("CSV 파싱 중..."):
        parsed = merge_parsed(gj_files, prefer="area")

    if not parsed:
        st.error("CSV에서 데이터를 추출하지 못했습니다."); return

    # 미리보기
    preview = []
    for comp, cdata in parsed.items():
        stats = get_gj_std_stats(cdata)
        preview.append({
            "성분": comp,
            "STD 런": stats.get("count", 0),
            "STD 평균 Area": stats.get("avg_area"),
            "%RSD": stats.get("rsd"),
            "SP 런": len(cdata.get("sp", [])),
        })
    st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

    # lot 이름 추출
    lot_names = set()
    for cdata in parsed.values():
        for entry in cdata.get("sp", []):
            import re
            m = re.search(r'^(.+?)_([AB])-', entry.get("name", ""))
            if m:
                lot_names.add(m.group(1))
    lot_name = list(lot_names)[0] if lot_names else "결과"
    st.info(f"Lot: {lot_name}")

    if test_type == "확인":
        from core.gj_id_excel_writer import write_gj_id_result
        with st.spinner("확인 결과 엑셀 생성 중..."):
            excel_bytes = write_gj_id_result(parsed)
        st.download_button("📥 확인결과 엑셀 다운로드", data=excel_bytes,
                           file_name=f"공진단_확인결과_{lot_name}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)
    else:
        from core.gj_as_excel_writer import write_gj_as_result
        with st.spinner("함량 결과 엑셀 생성 중..."):
            excel_bytes = write_gj_as_result(parsed)
        st.download_button("📥 함량결과 엑셀 다운로드", data=excel_bytes,
                           file_name=f"공진단_함량결과_{lot_name}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", use_container_width=True)


def _run_amino_acid(std_files, sp_files):
    from core.aa_pdf_parser import parse_std, parse_sp
    from core.aa_excel_writer import write_aa_result

    all_runs, all_res = [], []
    with st.spinner("STD PDF 파싱 중..."):
        for f in std_files:
            try:
                runs, ress = parse_std(f)
                all_runs.extend(runs)
                all_res.extend(ress)
            except Exception as e:
                st.error(f"STD 파싱 오류 ({f.name}): {e}"); return

    if not all_runs:
        st.error("STD PDF에서 데이터를 추출하지 못했습니다."); return

    st.success(f"STD {len(all_runs)}런 파싱 완료")

    all_lots: dict = {}
    with st.spinner("검액 PDF 파싱 중..."):
        for f in sp_files:
            try:
                lots, dbg = parse_sp(f, debug=True)
                for lot_id, lot_data in lots.items():
                    all_lots.setdefault(lot_id, {}).update(lot_data)
                if not lots:
                    with st.expander(f"🔍 파싱 디버그 ({f.name})", expanded=True):
                        st.code("\n".join(dbg) if dbg else "(블록 없음)")
            except Exception as e:
                st.error(f"검액 파싱 오류 ({f.name}): {e}"); return

    if not all_lots:
        st.error("검액 PDF에서 데이터를 추출하지 못했습니다."); return

    lot_names = sorted(all_lots.keys())
    st.success(f"검액 Lot: {', '.join(lot_names)}")

    with st.spinner("엑셀 생성 중..."):
        excel_bytes = write_aa_result(all_runs, all_res, all_lots)

    st.download_button("📥 결과 엑셀 다운로드", data=excel_bytes,
                       file_name="아미노산_함량_결과.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", use_container_width=True)


# ══════════════════════════════════════════════════════════════════
# 세션 상태 초기화
# ══════════════════════════════════════════════════════════════════
if "product"   not in st.session_state: st.session_state.product   = None
if "test_type" not in st.session_state: st.session_state.test_type = None

# ══════════════════════════════════════════════════════════════════
# UI — 헤더
# ══════════════════════════════════════════════════════════════════
st.markdown("""
<style>
div.stButton > button { height:55px; font-size:16px; font-weight:bold; border-radius:10px; }
</style>
""", unsafe_allow_html=True)

st.title("🔬 면적값 자동파싱 시스템")
st.caption("PDF를 업로드하면 자동으로 데이터를 추출하여 엑셀로 제공합니다")
st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 1 — 제품 선택
# ══════════════════════════════════════════════════════════════════
st.subheader("① 제품 선택")
col1, col2, col3, col4 = st.columns(4)

with col1:
    if st.button("💊  환제", use_container_width=True,
                 type="primary" if st.session_state.product=="환제" else "secondary"):
        st.session_state.product="환제"; st.session_state.test_type=None; st.rerun()
with col2:
    if st.button("🧪  현탁제", use_container_width=True,
                 type="primary" if st.session_state.product=="현탁제" else "secondary"):
        st.session_state.product="현탁제"; st.session_state.test_type=None; st.rerun()
with col3:
    if st.button("💫  공진단", use_container_width=True,
                 type="primary" if st.session_state.product=="공진단" else "secondary"):
        st.session_state.product="공진단"; st.session_state.test_type=None; st.rerun()
with col4:
    if st.button("🔬  파워라센 아미노산", use_container_width=True,
                 type="primary" if st.session_state.product=="아미노산" else "secondary"):
        st.session_state.product="아미노산"; st.session_state.test_type=None; st.rerun()

if not st.session_state.product:
    st.stop()

st.divider()

# ══════════════════════════════════════════════════════════════════
# STEP 2 — 시험 항목 선택 (아미노산 제외)
# ══════════════════════════════════════════════════════════════════
product = st.session_state.product

if product in ("환제", "현탁제"):
    st.subheader(f"② 시험 항목 선택  〔우황청심원 {product}〕")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📋  확인 (ID)", use_container_width=True,
                     type="primary" if st.session_state.test_type=="확인" else "secondary"):
            st.session_state.test_type="확인"; st.rerun()
    with col2:
        if st.button("⚗️  함량 (AS)", use_container_width=True,
                     type="primary" if st.session_state.test_type=="함량" else "secondary"):
            st.session_state.test_type="함량"; st.rerun()

    if not st.session_state.test_type:
        st.stop()
    st.divider()

elif product == "공진단":
    st.subheader("② 시험 항목 선택  〔공진단〕")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📋  확인 (ID)", use_container_width=True,
                     type="primary" if st.session_state.test_type=="확인" else "secondary"):
            st.session_state.test_type="확인"; st.rerun()
    with col2:
        if st.button("⚗️  함량 (AS)", use_container_width=True,
                     type="primary" if st.session_state.test_type=="함량" else "secondary"):
            st.session_state.test_type="함량"; st.rerun()

    if not st.session_state.test_type:
        st.stop()
    st.divider()

test_type = st.session_state.test_type

# ══════════════════════════════════════════════════════════════════
# STEP 3 — 파일 업로드
# ══════════════════════════════════════════════════════════════════

if product == "아미노산":
    st.subheader("② PDF 파일 업로드  〔아미노산 함량〕")
    st.caption("STD PDF + 검액 PDF(Lot별)를 한 번에 선택하세요 — 파일명에 STD 포함 여부로 자동 분류")
    uploaded_aa = st.file_uploader("PDF 파일 선택 (복수 선택 가능)",
                                   type="pdf", accept_multiple_files=True, key="aa_multi")
    std_files = [f for f in (uploaded_aa or []) if "STD" in f.name.upper()]
    sp_files  = [f for f in (uploaded_aa or []) if "STD" not in f.name.upper()]
    if uploaded_aa:
        col1, col2 = st.columns(2)
        with col1:
            st.caption("STD")
            for f in std_files: st.success(f"✓ {f.name}")
            if not std_files: st.warning("STD 파일 없음")
        with col2:
            st.caption("검액 (Lot)")
            for f in sp_files: st.success(f"✓ {f.name}")
            if not sp_files: st.warning("검액 파일 없음")
    if st.button("▶  분석 실행", type="primary",
                 disabled=not(std_files and sp_files), use_container_width=True):
        _run_amino_acid(std_files, sp_files)

elif product == "공진단":
    st.subheader(f"③ CSV 파일 업로드  〔공진단 {test_type}〕")
    st.caption("STD / SP-A / SP-B (/ SP-C) CSV 파일을 한 번에 선택하세요 — 파일명으로 자동 분류")
    uploaded = st.file_uploader("CSV 파일 선택 (복수 선택 가능)",
                                type="csv", accept_multiple_files=True, key="gj_csv")

    def _classify_gj(files):
        std_f, spa_f, spb_f = [], [], []
        for f in (files or []):
            n = f.name.upper().replace("-", "_")
            if "SP_B" in n or "SPB" in n:    spb_f.append(f)
            elif "SP_A" in n or "SPA" in n:  spa_f.append(f)
            elif "STD" in n:                 std_f.append(f)
            else:
                # 이름으로 판단 불가 시 SP_A/B 자동 구분 불가 → std로 추가
                std_f.append(f)
        return std_f, spa_f, spb_f

    gj_std, gj_spa, gj_spb = _classify_gj(uploaded)

    if uploaded:
        cols = st.columns(3)
        for col, label, flist in zip(cols,
            ["STD", "SP_A", "SP_B"],
            [gj_std, gj_spa, gj_spb]):
            if flist:
                for f in flist: col.success(f"✓ {f.name}")
            else:
                col.info(f"{label} 없음")
            col.caption(label)

    all_gj = (gj_std or []) + (gj_spa or []) + (gj_spb or [])
    if st.button("▶  분석 실행", type="primary",
                 disabled=not all_gj, use_container_width=True):
        try:
            _run_gj(all_gj, test_type)
        except Exception as e:
            st.error(f"오류: {e}")

elif test_type == "확인":
    st.subheader(f"③ CSV 파일 업로드  〔우황청심원 {product} 확인〕")
    st.caption("SST / SP_A / SP_B / SP_C 파일을 한 번에 선택하세요 (파일명에 SST, SP_A, SP_B, SP_C 포함)")
    uploaded = st.file_uploader("CSV 파일 선택 (복수 선택 가능)",
                                type="csv", accept_multiple_files=True,
                                key="id_csv_multi")
    sst_file = spa_file = spb_file = spc_file = None
    for f in (uploaded or []):
        n = f.name.upper()
        if "SST"  in n: sst_file = f
        elif "SP_A" in n or "SPA" in n: spa_file = f
        elif "SP_B" in n or "SPB" in n: spb_file = f
        elif "SP_C" in n or "SPC" in n: spc_file = f
    if uploaded:
        cols = st.columns(4)
        for col, label, f in zip(cols,
            ["SST", "SP_A", "SP_B", "SP_C"],
            [sst_file, spa_file, spb_file, spc_file]):
            col.success(f"✓ {f.name}" if f else "—")
            col.caption(label)
    if st.button("▶  분석 실행", type="primary",
                 disabled=not any([spa_file, spb_file, spc_file]),
                 use_container_width=True):
        _run_id_csv(product, sst_file, spa_file, spb_file, spc_file)

elif test_type == "함량":
    st.subheader(f"③ CSV 파일 업로드  〔우황청심원 {product} 함량〕")
    st.caption("통합 CSV 1개 또는 분리된 CSV 여러 개를 한 번에 선택하세요")
    as_files = st.file_uploader("AS 데이터 CSV (복수 선택 가능)",
                                type=["csv"], accept_multiple_files=True, key="as_csv")
    if as_files:
        for f in as_files:
            st.caption(f"✓ {f.name}")
    if st.button("▶  분석 실행", type="primary",
                 disabled=not as_files, use_container_width=True):
        _run_as(product, as_files)
