import numpy as np
import pandas as pd


def robust_zscore(values: np.ndarray) -> np.ndarray:
    """
    Robust Z-score (AAFCO 방식)
    Z = (x - median) / (1.4826 * MAD)
    MAD=0인 경우 표준편차 기반 Z-score로 폴백.
    """
    values = np.array(values, dtype=float)
    median = np.median(values)
    mad = np.median(np.abs(values - median))
    if mad == 0:
        std = np.std(values)
        if std == 0:
            return np.zeros_like(values)
        return (values - np.mean(values)) / std
    return (values - median) / (1.4826 * mad)


def compute_zscores(df: pd.DataFrame, value_cols: list) -> pd.DataFrame:
    """각 값 컬럼별 전체 Robust Z-score 계산."""
    result = pd.DataFrame(index=df.index, dtype=float)
    for col in value_cols:
        vals = pd.to_numeric(df[col], errors="coerce")
        valid = vals.notna()
        z = pd.Series(np.nan, index=df.index)
        if valid.sum() > 5:
            z.loc[valid[valid].index] = robust_zscore(vals[valid].values)
        result[col] = z
    return result


def compute_zscores_by_method(
    df: pd.DataFrame, value_col: str, method_col: str
) -> pd.Series:
    """
    방법별 Robust Z-score 계산.
    동일 방법을 사용한 기관이 5개 미만이면 NaN 반환.
    """
    result = pd.Series(np.nan, index=df.index, dtype=float)
    if method_col not in df.columns:
        return result
    for method, grp in df.groupby(df[method_col].fillna("").astype(str)):
        if not method.strip():
            continue
        vals = pd.to_numeric(grp[value_col], errors="coerce")
        valid = vals.notna()
        if valid.sum() <= 5:
            continue
        z = robust_zscore(vals[valid].values)
        result.loc[vals[valid].index] = z
    return result


def zscore_flag(z: float) -> str:
    if np.isnan(z):
        return "N/A"
    az = abs(z)
    if az <= 2.0:
        return "적합"
    elif az <= 3.0:
        return "경고"
    else:
        return "부적합"


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
