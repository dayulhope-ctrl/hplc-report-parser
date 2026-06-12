# -*- coding: utf-8 -*-
"""공진단 CSV 파서 — csv_utils 범용 엔진 사용"""
from core.csv_utils import parse_csv, merge_parsed, group_by_lot, get_stats

def parse_gj_csv(file_obj) -> dict:
    return parse_csv(file_obj, prefer="area")

def merge_gj_parsed(file_list) -> dict:
    return merge_parsed(file_list, prefer="area")

def group_gj_sp_by_lot(parsed: dict) -> dict:
    return group_by_lot(parsed)

def get_gj_std_stats(compound_data: dict) -> dict:
    return get_stats(compound_data)
