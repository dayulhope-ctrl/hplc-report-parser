# -*- coding: utf-8 -*-
"""함량 계산 로직"""


def calc_content(
    std_weight_mg: float,       # 표준품 채취량 (mg)
    std_purity_pct: float,      # 표준품 함량 (%)
    std_volume_mL: float,       # 표준액 부피 (mL)
    sp_volume_mL: float,        # 검액 부피 (mL)
    std_compound_resp: float,   # 표준액 중 해당성분 면적 (평균)
    std_is_resp: float,         # 표준액 중 IS 면적 (평균)
    sp_compound_resp: float,    # 검액 중 해당성분 면적
    sp_is_resp: float,          # 검액 중 IS 면적
    sp_weight_mg: float,        # 검체 채취량 (mg)
    avg_weight_mg: float,       # 평균 질량 (mg/환)
    unit_factor: float = 1.0,   # 1=mg, 1000=μg
) -> float:
    """
    함량(mg 또는 μg / 1환) 계산

    공통 공식:
      분자: 표준품채취량 × 표준액부피 × 검액성분면적 × 표준액IS면적 × unit_factor × 평균질량 × 표준품함량%
      분모: 검체채취량 × 검액부피 × 표준액성분면적 × 검액IS면적 × 1환 × 100
    """
    if std_compound_resp == 0 or sp_is_resp == 0:
        return None

    numerator = (std_weight_mg * std_volume_mL * sp_compound_resp
                 * std_is_resp * unit_factor * avg_weight_mg * std_purity_pct)
    denominator = (sp_weight_mg * sp_volume_mL * std_compound_resp
                   * sp_is_resp * 1 * 100)

    return round(numerator / denominator, 4)


def calc_std_is_ratio(std_compound_resp: float, std_is_resp: float) -> float:
    """표준액 STD/I.S 비율"""
    if std_is_resp == 0:
        return None
    return round(std_compound_resp / std_is_resp, 4)
