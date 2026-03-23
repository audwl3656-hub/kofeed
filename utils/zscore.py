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
    """
    전체 Robust Z-score 계산.
    같은 base column(_N suffix)은 하나의 풀로 합산해 Z-score 계산.
    예: 조단백질_축우사료 + 조단백질_축우사료_2 → 동일 풀 사용.
    """
    from utils.config import get_base_col

    # base_col → [col, ...] 그룹핑
    groups: dict[str, list[str]] = {}
    for col in value_cols:
        base = get_base_col(col)
        groups.setdefault(base, []).append(col)

    result = pd.DataFrame(index=df.index, dtype=float)

    for base, cols in groups.items():
        # 풀링: 모든 suffix 컬럼 값을 하나로 모아 Z-score 계산
        pool_vals: list[float] = []
        pool_meta: list[tuple] = []  # (row_idx, col)

        for col in cols:
            vals = pd.to_numeric(df[col], errors="coerce")
            for idx in df.index[vals.notna()]:
                pool_vals.append(float(vals[idx]))
                pool_meta.append((idx, col))

        z_scores = (
            robust_zscore(np.array(pool_vals)).tolist()
            if len(pool_vals) > 5
            else [np.nan] * len(pool_vals)
        )

        # 컬럼별 Series 초기화 후 채우기
        col_series: dict[str, pd.Series] = {
            col: pd.Series(np.nan, index=df.index) for col in cols
        }
        for (idx, col), z in zip(pool_meta, z_scores):
            col_series[col][idx] = z

        for col in cols:
            result[col] = col_series[col]

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
