import numpy as np
import pandas as pd


def robust_zscore(values: np.ndarray) -> np.ndarray:
    """
    Robust Z-score (KS Q ISO 13528)
    Z = (x - Median) / ((Q3 - Q1) × 0.7413)
    IQR=0인 경우 표준편차 기반 Z-score로 폴백.
    """
    values = np.array(values, dtype=float)
    median = np.median(values)
    q1, q3 = np.percentile(values, [25, 75])
    niqr = (q3 - q1) * 0.7413
    if niqr == 0:
        std = np.std(values)
        if std == 0:
            return np.zeros_like(values)
        return (values - np.mean(values)) / std
    return (values - median) / niqr


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


def compute_zscores_by_method_multi(
    df: pd.DataFrame, value_cols: list, method_cols: list
) -> dict:
    """
    여러 suffix 컬럼(예: 조단백질_축우사료, 조단백질_축우사료_2)을
    하나의 풀로 묶어 방법별 Robust Z-score 계산.
    value_cols[i]와 method_cols[i]가 쌍을 이룸.
    동일 방법 사용 기관이 5개 이하이면 NaN.
    반환: {col: pd.Series}
    """
    from collections import defaultdict

    # (row_idx, col, value, method) 수집
    entries = []
    for vcol, mcol in zip(value_cols, method_cols):
        vals = (
            pd.to_numeric(df[vcol], errors="coerce")
            if vcol in df.columns
            else pd.Series(np.nan, index=df.index)
        )
        meths = (
            df[mcol].fillna("").astype(str)
            if mcol in df.columns
            else pd.Series("", index=df.index)
        )
        for idx in df.index:
            v = vals[idx]
            m = str(meths[idx]).strip()
            if pd.notna(v) and m:
                entries.append((idx, vcol, float(v), m))

    # 방법별 그룹핑
    method_groups: dict = defaultdict(list)
    for idx, vcol, v, m in entries:
        method_groups[m].append((idx, vcol, v))

    # 결과 Series 초기화
    result = {col: pd.Series(np.nan, index=df.index, dtype=float) for col in value_cols}

    # 방법별 z-score 계산 후 매핑
    for method, items in method_groups.items():
        if len(items) <= 5:
            continue
        vals_arr = np.array([v for _, _, v in items])
        z_arr = robust_zscore(vals_arr)
        for (idx, vcol, _), z in zip(items, z_arr):
            result[vcol].at[idx] = z

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
