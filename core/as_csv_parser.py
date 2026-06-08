# -*- coding: utf-8 -*-
"""환제 AS 함량 CSV 파서 — csv_utils 범용 엔진 사용"""
from core.csv_utils import parse_csv, group_by_lot, get_stats

def parse_as_csv(file_obj) -> dict:
    return parse_csv(file_obj)

def group_sp_by_lot(parsed: dict) -> dict:
    return group_by_lot(parsed)

def get_std_stats(compound_data: dict) -> dict:
    return get_stats(compound_data)
