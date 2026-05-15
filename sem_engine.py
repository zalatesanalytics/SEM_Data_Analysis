"""Core SEM analysis functions for the EMERGE+ Streamlit application."""
from __future__ import annotations

import io
from pathlib import Path
from typing import Dict, List, Tuple

import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import pearsonr

from sem_config import LATENT_VARIABLES, OBSERVED_OUTCOMES, STRUCTURAL_PATHS, MEDIATION_PATHS


def load_data(uploaded_file) -> pd.DataFrame:
    """Load CSV or Excel data from Streamlit uploader."""
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(uploaded_file)
    return pd.read_csv(uploaded_file)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Make column names Python-friendly while preserving analytical meaning."""
    clean = df.copy()
    clean.columns = (
        clean.columns.astype(str)
        .str.strip()
        .str.lower()
        .str.replace(r"[^0-9a-zA-Z]+", "_", regex=True)
        .str.strip("_")
    )
    return clean


def available_indicators(df: pd.DataFrame) -> Dict[str, List[str]]:
    return {k: [v for v in spec["indicators"] if v in df.columns] for k, spec in LATENT_VARIABLES.items()}


def variable_mapping(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for key, spec in LATENT_VARIABLES.items():
        for item in spec["indicators"]:
            rows.append({
                "Construct": key,
                "Construct name": spec["name"],
                "Expected variable": item,
                "Found in dataset": item in df.columns,
            })
    for out in OBSERVED_OUTCOMES:
        rows.append({"Construct": "Observed outcome", "Construct name": "Outcome", "Expected variable": out, "Found in dataset": out in df.columns})
    return pd.DataFrame(rows)


def coerce_analysis_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    vars_needed = set(OBSERVED_OUTCOMES)
    for spec in LATENT_VARIABLES.values():
        vars_needed.update(spec["indicators"])
    for col in vars_needed.intersection(out.columns):
        out[col] = pd.to_numeric(out[col], errors="coerce")
    return out


def descriptive_stats(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for key, spec in LATENT_VARIABLES.items():
        for ind in [x for x in spec["indicators"] if x in df.columns]:
            s = pd.to_numeric(df[ind], errors="coerce")
            v = s.dropna()
            rows.append({
                "Construct": key, "Construct name": spec["name"], "Variable": ind,
                "N": int(v.count()), "Missing": int(s.isna().sum()),
                "Mean": v.mean(), "SD": v.std(ddof=1), "Median": v.median(),
                "Min": v.min(), "Max": v.max(),
                "Skewness": stats.skew(v, bias=False) if len(v) > 2 else np.nan,
                "Kurtosis": stats.kurtosis(v, bias=False) if len(v) > 3 else np.nan,
            })
    return pd.DataFrame(rows).round(4)


def cronbach_alpha(df: pd.DataFrame, indicators: List[str]) -> float:
    data = df[indicators].apply(pd.to_numeric, errors="coerce").dropna()
    k = data.shape[1]
    if k < 2 or data.shape[0] < 3:
        return np.nan
    item_vars = data.var(axis=0, ddof=1)
    total_var = data.sum(axis=1).var(ddof=1)
    return np.nan if total_var == 0 else float((k / (k - 1)) * (1 - item_vars.sum() / total_var))


def factor_loadings_proxy(df: pd.DataFrame, indicators: List[str]) -> pd.Series:
    data = df[indicators].apply(pd.to_numeric, errors="coerce").dropna()
    if data.shape[0] < 5 or len(indicators) < 2:
        return pd.Series(dtype=float)
    total = data.sum(axis=1)
    vals = {}
    for ind in indicators:
        corrected = total - data[ind]
        if corrected.std() == 0 or data[ind].std() == 0:
            vals[ind] = np.nan
        else:
            vals[ind] = min(abs(pearsonr(data[ind], corrected)[0]), 0.99)
    return pd.Series(vals)


def reliability_table(df: pd.DataFrame) -> Tuple[pd.DataFrame, pd.DataFrame]:
    rel_rows, load_rows = [], []
    for key, spec in LATENT_VARIABLES.items():
        inds = [i for i in spec["indicators"] if i in df.columns]
        if len(inds) < 2:
            continue
        loads = factor_loadings_proxy(df, inds)
        alpha = cronbach_alpha(df, inds)
        arr = loads.dropna().to_numpy()
        cr = ((arr.sum() ** 2) / ((arr.sum() ** 2) + np.sum(1 - arr ** 2))) if len(arr) else np.nan
        ave = float(np.mean(arr ** 2)) if len(arr) else np.nan
        rel_rows.append({"Construct": key, "Name": spec["name"], "N items": len(inds), "Cronbach alpha": alpha, "Composite reliability": cr, "AVE": ave})
        for item, loading in loads.items():
            load_rows.append({"Construct": key, "Construct name": spec["name"], "Indicator": item, "Loading proxy": loading, "Communality": loading ** 2})
    return pd.DataFrame(rel_rows).round(4), pd.DataFrame(load_rows).round(4)


def construct_scores(df: pd.DataFrame) -> pd.DataFrame:
    scores = pd.DataFrame(index=df.index)
    for key, spec in LATENT_VARIABLES.items():
        inds = [i for i in spec["indicators"] if i in df.columns]
        if inds:
            scores[key] = df[inds].apply(pd.to_numeric, errors="coerce").mean(axis=1)
    for out in OBSERVED_OUTCOMES:
        if out in df.columns:
            scores[out] = pd.to_numeric(df[out], errors="coerce")
    return scores


def correlation_matrix(scores: pd.DataFrame) -> pd.DataFrame:
    return scores.dropna(axis=1, how="all").corr().round(3)


def path_coefficients(scores: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for src, tgt in STRUCTURAL_PATHS:
        if src not in scores.columns or tgt not in scores.columns:
            continue
        data = scores[[src, tgt]].dropna()
        if len(data) < 5 or data[src].std() == 0 or data[tgt].std() == 0:
            continue
        r, p = pearsonr(data[src], data[tgt])
        se = np.sqrt((1 - r ** 2) / (len(data) - 2))
        rows.append({"From": src, "To": tgt, "Beta/std r": r, "SE": se, "t-value": r / se if se else np.nan, "p-value": p, "Significant": p < 0.05})
    return pd.DataFrame(rows).round(4)


def mediation_analysis(scores: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for mp in MEDIATION_PATHS:
        iv, med, dv = mp["iv"], mp["mediator"], mp["dv"]
        if any(v not in scores.columns for v in (iv, med, dv)):
            continue
        data = scores[[iv, med, dv]].dropna()
        if len(data) < 10:
            continue
        a, pa = pearsonr(data[iv], data[med])
        b, pb = pearsonr(data[med], data[dv])
        c, pc = pearsonr(data[iv], data[dv])
        se_a = np.sqrt((1 - a ** 2) / (len(data) - 2))
        se_b = np.sqrt((1 - b ** 2) / (len(data) - 2))
        se_ab = np.sqrt((b ** 2 * se_a ** 2) + (a ** 2 * se_b ** 2))
        z = (a * b) / se_ab if se_ab else np.nan
        pz = 2 * (1 - stats.norm.cdf(abs(z))) if pd.notna(z) else np.nan
        rows.append({"Pathway": mp["name"], "a": a, "b": b, "c total": c, "Indirect a*b": a*b, "Direct approx": c-(a*b), "Sobel z": z, "Sobel p": pz, "Significant mediation": pz < 0.05 if pd.notna(pz) else False})
    return pd.DataFrame(rows).round(4)


def semopy_model(scores: pd.DataFrame):
    """Run a structural model on composite construct scores if semopy is available."""
    try:
        import semopy
    except Exception as exc:
        return None, pd.DataFrame({"Message": [f"semopy not available: {exc}"]}), pd.DataFrame()
    required = ["MHB", "CRB", "WD", "FP", "FR", "employment_confidence", "SC", "DR", "ME", "job_satisfaction", "financial_stability"]
    present = [c for c in required if c in scores.columns]
    data = scores[present].dropna()
    syntax_lines = []
    if set(["MHB", "CRB", "WD", "FP", "FR"]).issubset(data.columns):
        syntax_lines.append("MHB ~ CRB + WD + FP + FR")
    if set(["employment_confidence", "CRB", "SC", "DR", "MHB", "ME", "FR"]).issubset(data.columns):
        syntax_lines.append("employment_confidence ~ CRB + SC + DR + MHB + ME + FR")
    if set(["job_satisfaction", "WD", "SC", "MHB"]).issubset(data.columns):
        syntax_lines.append("job_satisfaction ~ WD + SC + MHB")
    if set(["financial_stability", "FP"]).issubset(data.columns):
        syntax_lines.append("financial_stability ~ FP")
    if set(["SC", "ME"]).issubset(data.columns):
        syntax_lines.append("SC ~ ME")
    if len(syntax_lines) < 2 or len(data) < 20:
        return None, pd.DataFrame({"Message": ["Not enough matched variables or complete cases for SEM."]}), pd.DataFrame()
    model = semopy.Model("\n".join(syntax_lines))
    model.fit(data)
    estimates = model.inspect()
    fit = semopy.calc_stats(model).T.reset_index().rename(columns={"index": "Fit statistic"})
    return model, estimates, fit


def write_excel(tables: Dict[str, pd.DataFrame], path: Path) -> Path:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, table in tables.items():
            safe = name[:31]
            table.to_excel(writer, sheet_name=safe, index=False)
    return path
