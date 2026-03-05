import numpy as np
import pandas as pd


def robust_zscore(values: np.ndarray) -> np.ndarray:
    """
    Robust Z-score (AAFCO 방식)
    Z = (x - median) / (1.4826 * MAD)
    MAD = median(|x - median(x)|)
    """
    values = np.array(values, dtype=float)
    median = np.median(values)
    mad = np.median(np.abs(values - median))
    if mad == 0:
        return np.where(values == median, 0.0, np.nan)
    return (values - median) / (1.4826 * mad)


def compute_zscores(df: pd.DataFrame, analyte_cols: list) -> pd.DataFrame:
    """
    각 분석 항목별 Robust Z-score 계산.
    반환: 동일 인덱스, 동일 컬럼 (Z-score 값으로 채워짐)
    """
    result = df[analyte_cols].copy().astype(float)
    for col in analyte_cols:
        vals = pd.to_numeric(df[col], errors="coerce")
        valid_mask = vals.notna()
        if valid_mask.sum() < 3:
            result[col] = np.nan
            continue
        zscores = robust_zscore(vals[valid_mask].values)
        result.loc[valid_mask, col] = zscores
        result.loc[~valid_mask, col] = np.nan
    return result


def zscore_flag(z: float) -> str:
    """Z-score 등급 판정"""
    if np.isnan(z):
        return "N/A"
    az = abs(z)
    if az <= 2.0:
        return "✅ 적합"
    elif az <= 3.0:
        return "⚠️ 경고"
    else:
        return "❌ 부적합"


def zscore_color(z: float) -> str:
    if np.isnan(z):
        return "#cccccc"
    az = abs(z)
    if az <= 2.0:
        return "#d4edda"
    elif az <= 3.0:
        return "#fff3cd"
    else:
        return "#f8d7da"
