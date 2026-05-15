
from pathlib import Path
import re
import numpy as np
import matplotlib.pyplot as plt
import networkx as nx
import pandas as pd
import seaborn as sns
import streamlit as st
from docx import Document
from docx.shared import Inches

try:
    from scipy import stats
except Exception:
    stats = None

from sem_engine import (
    load_data, normalize_columns, coerce_analysis_columns, variable_mapping,
    descriptive_stats, reliability_table, construct_scores, correlation_matrix,
    path_coefficients, mediation_analysis, semopy_model, write_excel,
)
from sem_config import LATENT_VARIABLES

st.set_page_config(page_title="EMERGE+ SEM Analyzer", layout="wide")
st.title("EMERGE+ Newcomer SEM Analyzer")
st.caption(
    "Upload a CSV or Excel file, clean variables, estimate SEM pathways, "
    "visualize results, generate descriptive analysis, and produce a Word report."
)

OUTPUT_ROOT = Path("outputs")
for sub in ["tables", "figures", "processed_data", "model_outputs"]:
    (OUTPUT_ROOT / sub).mkdir(parents=True, exist_ok=True)


# -------------------------------------------------------------------
# Dashboard styling
# -------------------------------------------------------------------

st.markdown(
    """
    <style>
    .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }
    .dashboard-hero {
        padding: 1.2rem 1.4rem;
        border-radius: 18px;
        background: linear-gradient(135deg, #f7fbff 0%, #eef4ff 50%, #f7f7fb 100%);
        border: 1px solid #e7edf7;
        margin-bottom: 1rem;
    }
    .hero-title {
        font-size: 1.55rem;
        font-weight: 750;
        color: #1f2937;
        margin-bottom: 0.2rem;
    }
    .hero-text {
        font-size: 0.98rem;
        color: #4b5563;
    }
    .section-card {
        padding: 1rem;
        border-radius: 16px;
        border: 1px solid #e5e7eb;
        background-color: #ffffff;
        box-shadow: 0 1px 8px rgba(31, 41, 55, 0.06);
        margin-bottom: 1rem;
    }
    .insight-box {
        padding: 0.85rem 1rem;
        border-radius: 14px;
        background-color: #f9fafb;
        border-left: 5px solid #94a3b8;
        margin-bottom: 0.6rem;
        font-size: 0.95rem;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        padding: 0.8rem;
        border-radius: 16px;
        box-shadow: 0 1px 8px rgba(31, 41, 55, 0.05);
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -------------------------------------------------------------------
# General helper functions
# -------------------------------------------------------------------

def safe_filename(name: str) -> str:
    """Convert any variable name into a safe file name."""
    return re.sub(r"[^A-Za-z0-9_]+", "_", str(name))[:80]


def has_variance(series: pd.Series) -> bool:
    """Return True only when a variable has at least two distinct non-missing values."""
    return series.dropna().astype(str).nunique() > 1


def is_year_column(df: pd.DataFrame, col: str) -> bool:
    """Detect year-like variables so they are treated as time/categorical, not quantitative."""
    name = str(col).lower().strip()
    if "year" in name or name in ["yr", "years", "year_enrolled", "enrollment_year"]:
        return True

    values = pd.to_numeric(df[col], errors="coerce").dropna()
    if values.empty:
        return False

    # Treat mostly four-digit integer values in a reasonable year range as year/time variables.
    in_year_range = values.between(1900, 2100).mean() >= 0.90
    mostly_integer = (values % 1 == 0).mean() >= 0.90
    reasonable_unique = values.nunique() <= 80
    return bool(in_year_range and mostly_integer and reasonable_unique)


def get_no_variance_columns(df: pd.DataFrame, cols):
    """Identify variables that should not produce visuals because all values are the same."""
    results = []
    for col in cols:
        if col not in df.columns:
            continue
        non_missing = df[col].dropna()
        unique_count = non_missing.astype(str).nunique()
        if unique_count <= 1:
            single_value = "Missing only" if non_missing.empty else str(non_missing.iloc[0])
            results.append({
                "Variable": col,
                "Reason skipped": "No variance / all non-missing values are the same",
                "Single value detected": single_value,
                "Non-missing responses": int(len(non_missing)),
            })
    return pd.DataFrame(results)


def add_bar_labels(ax, fmt="{:.0f}"):
    """Add values on bar charts for easier interpretation."""
    for patch in ax.patches:
        width = patch.get_width()
        height = patch.get_height()
        if pd.isna(width) or width == 0:
            continue
        ax.text(
            width,
            patch.get_y() + height / 2,
            " " + fmt.format(width),
            va="center",
            fontsize=7,
        )


def add_docx_table(doc, df: pd.DataFrame):
    """Add a pandas DataFrame to a Word document."""
    if df is None or df.empty:
        doc.add_paragraph("No data available.")
        return

    display = df.copy().fillna("")
    table = doc.add_table(rows=1, cols=len(display.columns))
    table.style = "Table Grid"

    for i, col in enumerate(display.columns):
        table.rows[0].cells[i].text = str(col)

    for _, row in display.iterrows():
        cells = table.add_row().cells
        for i, val in enumerate(row):
            cells[i].text = str(val)[:120]


def detect_numeric_variables(df):
    """Detect quantitative variables, excluding year/time-like columns."""
    numeric = df.select_dtypes(include="number").columns.tolist()
    return [col for col in numeric if not is_year_column(df, col)]


def detect_year_columns(df):
    """Return columns that appear to represent years/time periods."""
    return [col for col in df.columns if is_year_column(df, col)]


def detect_likert_variables(df):
    """
    Detect Likert-scale variables coded 1 to 5.
    These are used to calculate percentage agreeing or strongly agreeing.
    """
    likert_cols = []
    for col in df.columns:
        values = df[col].dropna().unique()
        numeric_values = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
        if len(numeric_values) > 0 and numeric_values.between(1, 5).all():
            unique_count = numeric_values.nunique()
            if 2 <= unique_count <= 5:
                likert_cols.append(col)
    return likert_cols


def detect_yes_no_variables(df):
    """Detect Yes/No variables."""
    yes_no_cols = []
    yes_no_values = {"yes", "no", "1", "0", "true", "false"}

    for col in df.columns:
        values = df[col].dropna().astype(str).str.lower().str.strip().unique()
        if len(values) > 0 and set(values).issubset(yes_no_values):
            yes_no_cols.append(col)

    return yes_no_cols


def detect_categorical_variables(df, numeric_cols=None, max_unique=30):
    """
    Detect categorical variables for frequency distribution.
    Year/time-like columns are treated as categorical/time variables, not quantitative variables.
    """
    numeric_cols = numeric_cols or []
    categorical_cols = []

    for col in df.columns:
        unique_count = df[col].nunique(dropna=True)

        if unique_count == 0:
            continue

        if is_year_column(df, col):
            categorical_cols.append(col)
        elif col not in numeric_cols:
            if unique_count <= max_unique:
                categorical_cols.append(col)
        else:
            if unique_count <= 10:
                categorical_cols.append(col)

    return categorical_cols


def find_gender_column(df):
    """
    Find the most likely gender column.
    Priority is given to columns named gender, sex, respondent_gender, or participant_gender.
    """
    candidate_names = [
        "gender", "sex", "respondent_gender", "participant_gender",
        "gender_identity", "respondent_sex"
    ]

    lower_map = {col.lower().strip(): col for col in df.columns}

    for name in candidate_names:
        if name in lower_map:
            return lower_map[name]

    for col in df.columns:
        if "gender" in col.lower() or col.lower().strip() == "sex":
            return col

    return None


def standardize_gender_for_men_women(df, gender_col):
    """
    Keep gender disaggregation only for Men and Women.
    Any third option or other response is excluded from gender-disaggregated analysis.
    """
    if gender_col is None or gender_col not in df.columns:
        return df.copy(), None

    df_gender = df.copy()

    def recode_gender(value):
        if pd.isna(value):
            return None

        text = str(value).strip().lower()

        if text in ["male", "man", "men", "m", "1"]:
            return "Men"
        if text in ["female", "woman", "women", "f", "2"]:
            return "Women"

        return None

    df_gender["gender_group"] = df_gender[gender_col].apply(recode_gender)
    df_gender = df_gender[df_gender["gender_group"].isin(["Men", "Women"])].copy()

    return df_gender, "gender_group"


# -------------------------------------------------------------------
# Descriptive analysis functions
# -------------------------------------------------------------------

def quantitative_summary(df, numeric_cols):
    """Produce average and descriptive statistics for all quantitative variables."""
    if not numeric_cols:
        return pd.DataFrame()

    summary = df[numeric_cols].describe().T.reset_index().rename(columns={
        "index": "Variable",
        "count": "Valid responses",
        "mean": "Mean / Average",
        "std": "Standard deviation",
        "min": "Minimum",
        "25%": "25th percentile",
        "50%": "Median",
        "75%": "75th percentile",
        "max": "Maximum"
    })

    numeric_summary_cols = [
        "Mean / Average", "Standard deviation", "Minimum",
        "25th percentile", "Median", "75th percentile", "Maximum"
    ]

    for col in numeric_summary_cols:
        if col in summary.columns:
            summary[col] = summary[col].round(2)

    return summary


def categorical_frequency_summary(df, categorical_cols):
    """Create frequency and percentage distribution for categorical variables."""
    all_results = []

    for col in categorical_cols:
        counts = df[col].fillna("Missing").astype(str).value_counts(dropna=False)
        total = counts.sum()

        for category, count in counts.items():
            all_results.append({
                "Variable": col,
                "Category": category,
                "Frequency": int(count),
                "Percentage": round((count / total) * 100, 2) if total > 0 else 0
            })

    return pd.DataFrame(all_results)


def likert_agree_summary(df, likert_cols):
    """Calculate percentage Agree or Strongly Agree for Likert variables coded 1-5."""
    results = []

    for col in likert_cols:
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        total = len(series)

        if total == 0:
            continue

        agree_count = series[series >= 4].count()

        results.append({
            "Variable": col,
            "Valid responses": total,
            "Agree or Strongly Agree": agree_count,
            "Percentage Agree or Above": round((agree_count / total) * 100, 2)
        })

    return pd.DataFrame(results)


def yes_no_summary(df, yes_no_cols):
    """Calculate percentage Yes for Yes/No variables."""
    results = []

    for col in yes_no_cols:
        series = df[col].dropna().astype(str).str.lower().str.strip()
        total = len(series)

        if total == 0:
            continue

        yes_count = series.isin(["yes", "1", "true"]).sum()

        results.append({
            "Variable": col,
            "Valid responses": total,
            "Yes responses": int(yes_count),
            "Percentage Yes": round((yes_count / total) * 100, 2)
        })

    return pd.DataFrame(results)



def likert_full_distribution(df, likert_cols):
    """Show full Likert response distribution: Strongly Disagree to Strongly Agree."""
    if not likert_cols:
        return pd.DataFrame()

    labels = {
        1: "Strongly Disagree",
        2: "Disagree",
        3: "Neutral",
        4: "Agree",
        5: "Strongly Agree",
    }
    results = []

    for col in likert_cols:
        series = pd.to_numeric(df[col], errors="coerce").dropna()
        total = len(series)
        if total == 0:
            continue
        for score in [1, 2, 3, 4, 5]:
            count = int((series == score).sum())
            results.append({
                "Likert Variable": col,
                "Likert Value": score,
                "Response Label": labels[score],
                "Frequency": count,
                "Percentage": round((count / total) * 100, 2),
                "Valid responses": total,
            })

    return pd.DataFrame(results)


def likert_full_distribution_by_group(df, likert_cols, group_col):
    """Show full Likert distribution by selected grouping/category variable."""
    if group_col is None or group_col not in df.columns or not likert_cols:
        return pd.DataFrame()

    labels = {
        1: "Strongly Disagree",
        2: "Disagree",
        3: "Neutral",
        4: "Agree",
        5: "Strongly Agree",
    }
    results = []

    temp = df.copy()
    temp[group_col] = temp[group_col].fillna("Missing").astype(str)

    for group, gdf in temp.groupby(group_col):
        for col in likert_cols:
            series = pd.to_numeric(gdf[col], errors="coerce").dropna()
            total = len(series)
            if total == 0:
                continue
            for score in [1, 2, 3, 4, 5]:
                count = int((series == score).sum())
                results.append({
                    "Grouping Variable": group_col,
                    "Group / Category": group,
                    "Likert Variable": col,
                    "Likert Value": score,
                    "Response Label": labels[score],
                    "Frequency": count,
                    "Percentage": round((count / total) * 100, 2),
                    "Valid responses": total,
                })

    return pd.DataFrame(results)


def quantitative_summary_by_group(df, numeric_cols, group_col):
    """Mean/average quantitative analysis by gender or selected categorical group."""
    if group_col is None or group_col not in df.columns or not numeric_cols:
        return pd.DataFrame()

    results = []

    for group, gdf in df.groupby(group_col):
        for col in numeric_cols:
            series = pd.to_numeric(gdf[col], errors="coerce").dropna()
            if series.empty:
                continue

            results.append({
                "Group variable": group_col,
                "Group": group,
                "Variable": col,
                "Valid responses": len(series),
                "Mean / Average": round(series.mean(), 2),
                "Median": round(series.median(), 2),
                "Minimum": round(series.min(), 2),
                "Maximum": round(series.max(), 2),
                "Standard deviation": round(series.std(), 2) if len(series) > 1 else 0
            })

    return pd.DataFrame(results)


def likert_summary_by_group(df, likert_cols, group_col):
    """Likert Agree or Strongly Agree analysis by gender or selected categorical group."""
    if group_col is None or group_col not in df.columns or not likert_cols:
        return pd.DataFrame()

    results = []

    for group, gdf in df.groupby(group_col):
        for col in likert_cols:
            series = pd.to_numeric(gdf[col], errors="coerce").dropna()
            total = len(series)

            if total == 0:
                continue

            agree_count = series[series >= 4].count()

            results.append({
                "Group variable": group_col,
                "Group": group,
                "Variable": col,
                "Valid responses": total,
                "Agree or Strongly Agree": int(agree_count),
                "Percentage Agree or Above": round((agree_count / total) * 100, 2)
            })

    return pd.DataFrame(results)


def yes_no_summary_by_group(df, yes_no_cols, group_col):
    """Yes response percentage by gender or selected categorical group."""
    if group_col is None or group_col not in df.columns or not yes_no_cols:
        return pd.DataFrame()

    results = []

    for group, gdf in df.groupby(group_col):
        for col in yes_no_cols:
            series = gdf[col].dropna().astype(str).str.lower().str.strip()
            total = len(series)

            if total == 0:
                continue

            yes_count = series.isin(["yes", "1", "true"]).sum()

            results.append({
                "Group variable": group_col,
                "Group": group,
                "Variable": col,
                "Valid responses": total,
                "Yes responses": int(yes_count),
                "Percentage Yes": round((yes_count / total) * 100, 2)
            })

    return pd.DataFrame(results)


def categorical_frequency_by_group(df, categorical_cols, group_col):
    """Frequency distribution of categorical variables by gender or selected group."""
    if group_col is None or group_col not in df.columns or not categorical_cols:
        return pd.DataFrame()

    all_results = []

    for group, gdf in df.groupby(group_col):
        for col in categorical_cols:
            if col == group_col:
                continue

            counts = gdf[col].fillna("Missing").astype(str).value_counts(dropna=False)
            total = counts.sum()

            for category, count in counts.items():
                all_results.append({
                    "Group variable": group_col,
                    "Group": group,
                    "Variable": col,
                    "Category": category,
                    "Frequency": int(count),
                    "Percentage": round((count / total) * 100, 2) if total > 0 else 0
                })

    return pd.DataFrame(all_results)


def format_p_value(p_value):
    """Format p-values for reader-friendly reporting."""
    if pd.isna(p_value):
        return "Not available"
    if p_value < 0.001:
        return "<0.001"
    return round(float(p_value), 4)


def significance_label(p_value):
    """Create a simple interpretation label for p-values."""
    if pd.isna(p_value):
        return "Not tested"
    if p_value < 0.001:
        return "Highly significant"
    if p_value < 0.01:
        return "Significant at 1%"
    if p_value < 0.05:
        return "Significant at 5%"
    if p_value < 0.10:
        return "Marginally significant"
    return "Not significant"


def quantitative_gender_comparison_table(df_gender, numeric_cols, gender_group_col):
    """
    Create variable-by-variable quantitative comparison table:
    Variable | Men mean | Women mean | Difference | p-value.
    Difference is Men minus Women.
    """
    if gender_group_col is None or df_gender.empty or not numeric_cols:
        return pd.DataFrame()

    results = []
    for col in numeric_cols:
        men = pd.to_numeric(
            df_gender.loc[df_gender[gender_group_col] == "Men", col],
            errors="coerce"
        ).dropna()
        women = pd.to_numeric(
            df_gender.loc[df_gender[gender_group_col] == "Women", col],
            errors="coerce"
        ).dropna()

        if len(men) == 0 and len(women) == 0:
            continue

        men_mean = men.mean() if len(men) > 0 else np.nan
        women_mean = women.mean() if len(women) > 0 else np.nan
        difference = men_mean - women_mean if pd.notna(men_mean) and pd.notna(women_mean) else np.nan

        p_value = np.nan
        test_used = "Welch t-test"
        if stats is not None and len(men) >= 2 and len(women) >= 2:
            try:
                _, p_value = stats.ttest_ind(men, women, equal_var=False, nan_policy="omit")
            except Exception:
                p_value = np.nan

        results.append({
            "Variable / Indicator": col,
            "Men: average / mean": round(men_mean, 2) if pd.notna(men_mean) else "",
            "Women: average / mean": round(women_mean, 2) if pd.notna(women_mean) else "",
            "Difference (Men - Women)": round(difference, 2) if pd.notna(difference) else "",
            "p-value": format_p_value(p_value),
            "Significance Level": significance_label(p_value),
            "Men N": int(len(men)),
            "Women N": int(len(women)),
            "Test used": test_used if stats is not None else "Install scipy to calculate p-values"
        })

    return pd.DataFrame(results)


def quantitative_by_category_comparison(df, numeric_cols, group_col):
    """
    For categorical variables with multiple categories, show quantitative averages by category
    and include an ANOVA p-value when appropriate.
    """
    if group_col is None or group_col not in df.columns or not numeric_cols:
        return pd.DataFrame()

    results = []
    group_series = df[group_col].fillna("Missing").astype(str)
    group_values = [g for g in group_series.unique() if g != "Missing"]

    for col in numeric_cols:
        valid = pd.DataFrame({
            group_col: group_series,
            col: pd.to_numeric(df[col], errors="coerce")
        }).dropna(subset=[col])

        if valid.empty:
            continue

        grouped_values = []
        for group, gdf in valid.groupby(group_col):
            values = gdf[col].dropna()
            if len(values) > 0:
                grouped_values.append(values)

        p_value = np.nan
        test_used = "ANOVA"
        if stats is not None and len(grouped_values) >= 2 and all(len(x) >= 2 for x in grouped_values):
            try:
                _, p_value = stats.f_oneway(*grouped_values)
            except Exception:
                p_value = np.nan

        for group, gdf in valid.groupby(group_col):
            values = gdf[col].dropna()
            if len(values) == 0:
                continue
            results.append({
                "Grouping Variable": group_col,
                "Category": group,
                "Variable / Indicator": col,
                "N": int(len(values)),
                "Average / Mean": round(values.mean(), 2),
                "Median": round(values.median(), 2),
                "Standard deviation": round(values.std(), 2) if len(values) > 1 else 0,
                "Minimum": round(values.min(), 2),
                "Maximum": round(values.max(), 2),
                "Comparison p-value": format_p_value(p_value),
                "Significance Level": significance_label(p_value),
                "Test used": test_used if stats is not None else "Install scipy to calculate p-values"
            })

    return pd.DataFrame(results)


def categorical_proportion_by_category(df, categorical_cols, group_col):
    """
    For categorical outcomes, show proportions within each selected category/group.
    Chi-square p-value is added where a contingency table can be tested.
    """
    if group_col is None or group_col not in df.columns or not categorical_cols:
        return pd.DataFrame()

    results = []

    for col in categorical_cols:
        if col == group_col:
            continue

        temp = df[[group_col, col]].copy()
        temp[group_col] = temp[group_col].fillna("Missing").astype(str)
        temp[col] = temp[col].fillna("Missing").astype(str)

        p_value = np.nan
        test_used = "Chi-square"
        if stats is not None:
            try:
                ctab = pd.crosstab(temp[group_col], temp[col])
                if ctab.shape[0] >= 2 and ctab.shape[1] >= 2:
                    _, p_value, _, _ = stats.chi2_contingency(ctab)
            except Exception:
                p_value = np.nan

        for group, gdf in temp.groupby(group_col):
            total = len(gdf)
            counts = gdf[col].value_counts(dropna=False)
            for category, count in counts.items():
                results.append({
                    "Grouping Variable": group_col,
                    "Group / Category": group,
                    "Categorical Variable": col,
                    "Response Category": category,
                    "Frequency": int(count),
                    "Proportion / Percentage": round((count / total) * 100, 2) if total > 0 else 0,
                    "Comparison p-value": format_p_value(p_value),
                    "Significance Level": significance_label(p_value),
                    "Test used": test_used if stats is not None else "Install scipy to calculate p-values"
                })

    return pd.DataFrame(results)


def build_summary_insights(df, quant_summary, likert_summary, yes_no_results, gender_comparison):
    """Create short automated findings for the top of the dashboard."""
    insights = []
    insights.append(f"The uploaded dataset contains {df.shape[0]:,} records and {df.shape[1]:,} variables.")

    if not quant_summary.empty and "Mean / Average" in quant_summary.columns:
        top_numeric = quant_summary.sort_values("Valid responses", ascending=False).head(1)
        if not top_numeric.empty:
            row = top_numeric.iloc[0]
            insights.append(f"Most complete quantitative variable: {row['Variable']} with average {row['Mean / Average']} based on {int(row['Valid responses']):,} valid responses.")

    if not likert_summary.empty:
        top_barrier = likert_summary.sort_values("Percentage Agree or Above", ascending=False).head(1).iloc[0]
        insights.append(f"Highest agreement item: {top_barrier['Variable']} at {top_barrier['Percentage Agree or Above']}% Agree/Strongly Agree.")

    if not yes_no_results.empty:
        top_yes = yes_no_results.sort_values("Percentage Yes", ascending=False).head(1).iloc[0]
        insights.append(f"Highest Yes response: {top_yes['Variable']} at {top_yes['Percentage Yes']}% Yes.")

    if not gender_comparison.empty:
        sig = gender_comparison[gender_comparison["Significance Level"].astype(str).str.contains("Significant|Highly", case=False, na=False)]
        if not sig.empty:
            row = sig.iloc[0]
            insights.append(f"Gender comparison highlight: {row['Variable / Indicator']} shows a statistically meaningful Men/Women difference with p-value {row['p-value']}.")
        else:
            insights.append("Gender comparison table was generated for Men and Women; no statistically significant differences were detected at the 5% level based on available data.")

    return insights


# -------------------------------------------------------------------
# Visualization functions
# -------------------------------------------------------------------

def save_likert_bar_chart(summary_df, outpath):
    if summary_df.empty:
        return

    fig, ax = plt.subplots(figsize=(9, 5))
    plot_df = summary_df.sort_values("Percentage Agree or Above", ascending=True).tail(15)

    ax.barh(plot_df["Variable"], plot_df["Percentage Agree or Above"])
    add_bar_labels(ax, fmt="{:.1f}%")
    ax.set_xlabel("Percentage Agree or Strongly Agree")
    ax.set_title("Top Likert-Scale Barriers")
    ax.set_xlim(0, max(100, plot_df["Percentage Agree or Above"].max() * 1.18))
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def save_grouped_likert_chart(summary_df, outpath, title="Gender-Disaggregated Likert Analysis"):
    if summary_df.empty:
        return

    plot_df = summary_df.copy()

    # Keep top variables by average agreement for readability
    top_vars = (
        plot_df.groupby("Variable")["Percentage Agree or Above"]
        .mean()
        .sort_values(ascending=False)
        .head(10)
        .index
    )

    plot_df = plot_df[plot_df["Variable"].isin(top_vars)]

    fig, ax = plt.subplots(figsize=(10, 5))
    pivot = plot_df.pivot_table(
        index="Variable",
        columns="Group",
        values="Percentage Agree or Above",
        aggfunc="mean"
    ).fillna(0)

    pivot.plot(kind="barh", ax=ax)
    add_bar_labels(ax, fmt="{:.1f}%")
    ax.set_xlabel("Percentage Agree or Strongly Agree")
    ax.set_title(title)
    ax.legend(title="Group", loc="best")
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def save_numeric_histograms(df, numeric_cols, output_dir):
    saved_paths = []

    for col in numeric_cols[:12]:
        series = pd.to_numeric(df[col], errors="coerce").dropna()

        if series.empty or not has_variance(series):
            continue

        fig, ax = plt.subplots(figsize=(4.5, 3.2))
        counts, bins, patches = ax.hist(series, bins=15)
        for count, patch in zip(counts, patches):
            if count > 0:
                ax.text(
                    patch.get_x() + patch.get_width() / 2,
                    count,
                    str(int(count)),
                    ha="center",
                    va="bottom",
                    fontsize=7,
                )
        ax.set_title(f"{col}", fontsize=9)
        ax.set_xlabel(col, fontsize=8)
        ax.set_ylabel("Frequency", fontsize=8)
        ax.tick_params(axis="both", labelsize=7)
        fig.tight_layout()

        outpath = output_dir / f"histogram_{safe_filename(col)}.png"
        fig.savefig(outpath, dpi=220)
        plt.close(fig)
        saved_paths.append(outpath)

    return saved_paths


def save_yes_no_pie_charts(df, yes_no_cols, output_dir):
    saved_paths = []

    def autopct_with_counts(values):
        total = sum(values)
        def _inner(pct):
            count = int(round(pct * total / 100.0))
            return f"{pct:.1f}%\n(n={count})"
        return _inner

    for col in yes_no_cols[:8]:
        series = df[col].dropna().astype(str).str.lower().str.strip()
        yes_count = int(series.isin(["yes", "1", "true"]).sum())
        no_count = int(series.isin(["no", "0", "false"]).sum())
        values = [yes_count, no_count]

        if sum(values) == 0 or len([v for v in values if v > 0]) <= 1:
            continue

        fig, ax = plt.subplots(figsize=(4.2, 3.2))
        ax.pie(values, labels=["Yes", "No"], autopct=autopct_with_counts(values), textprops={"fontsize": 8})
        ax.set_title(f"{col}", fontsize=9)

        outpath = output_dir / f"pie_{safe_filename(col)}.png"
        fig.savefig(outpath, dpi=220)
        plt.close(fig)
        saved_paths.append(outpath)

    return saved_paths


def save_categorical_bar_charts(df, categorical_cols, output_dir):
    saved_paths = []

    for col in categorical_cols[:12]:
        counts = df[col].fillna("Missing").astype(str).value_counts().head(10)

        if counts.empty or counts.nunique() <= 1 and len(counts) <= 1:
            continue

        fig, ax = plt.subplots(figsize=(4.8, 3.2))
        counts_sorted = counts.sort_values()
        counts_sorted.plot(kind="barh", ax=ax)
        add_bar_labels(ax, fmt="{:.0f}")
        ax.set_title(f"{col}", fontsize=9)
        ax.set_xlabel("Frequency", fontsize=8)
        ax.tick_params(axis="both", labelsize=7)
        ax.set_xlim(0, max(1, counts_sorted.max() * 1.18))
        fig.tight_layout()

        outpath = output_dir / f"categorical_{safe_filename(col)}.png"
        fig.savefig(outpath, dpi=220)
        plt.close(fig)
        saved_paths.append(outpath)

    return saved_paths


def show_figures_grid(image_paths, columns_per_row=4):
    """Display smaller figures with at least four figures per row."""
    if not image_paths:
        st.info("No figures available for this section.")
        return

    for i in range(0, len(image_paths), columns_per_row):
        cols = st.columns(columns_per_row)
        for j, path in enumerate(image_paths[i:i + columns_per_row]):
            with cols[j]:
                st.image(str(path), use_container_width=True)



def display_variable_tables(table_df, variable_col, title_prefix=""):
    """Display one separate table per variable for reader-friendly review."""
    if table_df is None or table_df.empty:
        st.info("No data available for this analysis.")
        return
    if variable_col not in table_df.columns:
        st.dataframe(table_df, use_container_width=True)
        return

    variables = table_df[variable_col].dropna().astype(str).unique().tolist()
    for variable in variables:
        subset = table_df[table_df[variable_col].astype(str) == variable].copy()
        label = f"{title_prefix}{variable}" if title_prefix else variable
        with st.expander(label, expanded=False):
            st.dataframe(subset, use_container_width=True)


def save_corr_heatmap(corr: pd.DataFrame, outpath: Path):
    fig, ax = plt.subplots(figsize=(9, 6))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm", center=0, ax=ax)
    ax.set_title("Correlation Matrix")
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def save_path_diagram(paths: pd.DataFrame, outpath: Path):
    fig, ax = plt.subplots(figsize=(10, 7))
    g = nx.DiGraph()

    for _, row in paths.iterrows():
        g.add_edge(row["From"], row["To"], weight=row["Beta/std r"])

    if not g.nodes:
        ax.text(0.5, 0.5, "No valid path coefficients available", ha="center", va="center")
    else:
        pos = nx.spring_layout(g, seed=42, k=1.3)
        nx.draw_networkx_nodes(g, pos, node_size=2200, node_color="#f2f6ff", edgecolors="#333333", ax=ax)
        nx.draw_networkx_labels(g, pos, font_size=10, ax=ax)
        nx.draw_networkx_edges(g, pos, arrows=True, arrowstyle="-|>", arrowsize=18, width=1.8, ax=ax)

        labels = {
            (r["From"], r["To"]): f"{r['Beta/std r']:.2f}"
            for _, r in paths.iterrows()
        }
        nx.draw_networkx_edge_labels(g, pos, edge_labels=labels, font_size=9, ax=ax)

    ax.set_axis_off()
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)




def build_professional_summary_text(df, quant_summary, categorical_freq, likert_summary, likert_distribution, yes_no_results, gender_comparison):
    """Create a professional, reviewer-friendly automated summary and recommendations."""
    lines = []
    lines.append("# Automated Professional Summary Report")
    lines.append("")
    lines.append("## Executive Summary")
    lines.append(
        f"The uploaded dataset includes {df.shape[0]:,} respondent records and {df.shape[1]:,} variables. "
        "The dashboard produced descriptive, disaggregated, and SEM-ready outputs to support evidence review, program learning, and decision-making."
    )

    if not quant_summary.empty:
        lines.append("")
        lines.append("## Quantitative Profile")
        top_quant = quant_summary.sort_values("Valid responses", ascending=False).head(5)
        for _, row in top_quant.iterrows():
            lines.append(
                f"- {row['Variable']}: average/mean = {row['Mean / Average']}, median = {row['Median']}, "
                f"based on {int(row['Valid responses']):,} valid responses."
            )

    if not likert_summary.empty:
        lines.append("")
        lines.append("## Key Likert-Scale Findings")
        top_likert = likert_summary.sort_values("Percentage Agree or Above", ascending=False).head(5)
        for _, row in top_likert.iterrows():
            lines.append(
                f"- {row['Variable']}: {row['Percentage Agree or Above']}% of respondents selected Agree or Strongly Agree."
            )

    if not yes_no_results.empty:
        lines.append("")
        lines.append("## Key Yes/No Findings")
        top_yes = yes_no_results.sort_values("Percentage Yes", ascending=False).head(5)
        for _, row in top_yes.iterrows():
            lines.append(f"- {row['Variable']}: {row['Percentage Yes']}% answered Yes.")

    if not gender_comparison.empty:
        lines.append("")
        lines.append("## Gender-Disaggregated Findings")
        sig = gender_comparison[gender_comparison["Significance Level"].astype(str).str.contains("Significant|Highly", case=False, na=False)]
        if sig.empty:
            lines.append("- Men/Women comparison tables were generated. No statistically significant differences were detected at the 5% level among the tested quantitative variables.")
        else:
            for _, row in sig.head(5).iterrows():
                lines.append(
                    f"- {row['Variable / Indicator']}: Men mean = {row['Men: average / mean']}, "
                    f"Women mean = {row['Women: average / mean']}, Difference = {row['Difference (Men - Women)']}, "
                    f"p-value = {row['p-value']} ({row['Significance Level']})."
                )

    lines.append("")
    lines.append("## Recommended Interpretation for Reviewers")
    lines.append("- Focus first on variables with high agreement, high Yes responses, or statistically meaningful group differences.")
    lines.append("- Interpret p-values together with sample size, missing data, measurement quality, and program theory.")
    lines.append("- Use gender-disaggregated findings to identify whether barriers, confidence, costs, or outcomes differ between Men and Women.")
    lines.append("- Use optional category analysis for variables such as city, education, employment status, region, or country of origin to identify where support needs are concentrated.")

    lines.append("")
    lines.append("## Suggested Recommendations")
    lines.append("1. Prioritize the highest-ranked barriers in program design, referral pathways, and mentorship support.")
    lines.append("2. Use gender-disaggregated differences to design targeted supports where Men and Women experience different levels of cost, confidence, barriers, or outcomes.")
    lines.append("3. Review city/category-level patterns to target local partnerships, employer engagement, credential-navigation supports, and settlement services.")
    lines.append("4. Strengthen data quality by reviewing variables with low valid response counts, high missingness, or unclear coding before final reporting.")
    lines.append("5. Combine descriptive findings with SEM pathways to move from simple reporting to explanation of how barriers affect employment transition outcomes.")

    return "\n".join(lines)
# Updated histogram function
# Includes:
# - Absolute frequencies
# - Percentage labels
# - Data labels on bars

import matplotlib.pyplot as plt
import numpy as np

def save_quantitative_histogram(series, variable_name, outpath):
    series = series.dropna()

    fig, ax = plt.subplots(figsize=(7, 5))

    counts, bins, patches = ax.hist(series, bins=10)

    total = counts.sum()

    for count, patch in zip(counts, patches):
        if count > 0:
            percentage = (count / total) * 100
            x = patch.get_x() + patch.get_width() / 2
            y = patch.get_height()

            ax.text(
                x,
                y,
                f"{int(count)}\\n({percentage:.1f}%)",
                ha='center',
                va='bottom',
                fontsize=8
            )

    ax.set_title(f"Distribution of {variable_name}")
    ax.set_xlabel(variable_name)
    ax.set_ylabel("Frequency")

    plt.tight_layout()
    plt.savefig(outpath, dpi=300)
    plt.close()

def save_professional_summary_docx(summary_text, outpath: Path):
    """Save professional automated summary as a Word document."""
    doc = Document()
    for line in summary_text.split("\n"):
        clean = line.strip()
        if not clean:
            doc.add_paragraph("")
        elif clean.startswith("# "):
            doc.add_heading(clean.replace("# ", ""), level=0)
        elif clean.startswith("## "):
            doc.add_heading(clean.replace("## ", ""), level=1)
        elif clean.startswith("- "):
            doc.add_paragraph(clean.replace("- ", ""), style="List Bullet")
        elif re.match(r"^\d+\.\s", clean):
            doc.add_paragraph(re.sub(r"^\d+\.\s", "", clean), style="List Number")
        else:
            doc.add_paragraph(clean)
    doc.save(outpath)
    return outpath

# -------------------------------------------------------------------
# Word report
# -------------------------------------------------------------------

def create_word_report(tables, figures, outpath: Path):
    doc = Document()

    doc.add_heading("EMERGE+ Newcomer SEM Analysis Report", 0)

    doc.add_paragraph(
        "This automated report summarizes data readiness, descriptive statistics, "
        "categorical frequency distributions, gender-disaggregated findings, construct reliability, "
        "correlations, path coefficients, mediation tests, and SEM model fit where available."
    )

    doc.add_heading("1. Variable Mapping", level=1)
    add_docx_table(doc, tables["variable_mapping"].head(40))

    doc.add_heading("2. Descriptive Analysis", level=1)

    doc.add_heading("2.1 Quantitative Variables: Averages and Distribution", level=2)
    add_docx_table(doc, tables["quantitative_summary"].head(40))

    doc.add_heading("2.2 Categorical Variables: Frequency Distribution", level=2)
    add_docx_table(doc, tables["categorical_frequency"].head(60))

    doc.add_heading("2.3 Likert-Scale Agree or Above Analysis", level=2)
    add_docx_table(doc, tables["likert_agree_summary"].head(40))

    if figures.get("likert_bar") and figures["likert_bar"].exists():
        doc.add_picture(str(figures["likert_bar"]), width=Inches(6.5))

    doc.add_heading("2.4 Yes/No Variables", level=2)
    add_docx_table(doc, tables["yes_no_summary"].head(40))

    doc.add_heading("3. Gender-Disaggregated Analysis: Men and Women Only", level=1)
    doc.add_paragraph(
        "This section includes only respondents coded as Men or Women. Other gender responses "
        "are excluded from this specific disaggregation to meet the requested reporting format."
    )

    doc.add_heading("3.1 Quantitative Averages by Gender", level=2)
    add_docx_table(doc, tables["gender_quantitative"].head(60))

    doc.add_heading("3.2 Quantitative Men/Women Comparison with p-values", level=2)
    add_docx_table(doc, tables["gender_quantitative_comparison"].head(80))

    doc.add_heading("3.2 Likert Agreement by Gender", level=2)
    add_docx_table(doc, tables["gender_likert"].head(60))

    if figures.get("gender_likert") and figures["gender_likert"].exists():
        doc.add_picture(str(figures["gender_likert"]), width=Inches(6.5))

    doc.add_heading("3.3 Yes/No Responses by Gender", level=2)
    add_docx_table(doc, tables["gender_yes_no"].head(60))

    doc.add_heading("3.4 Categorical Frequency by Gender", level=2)
    add_docx_table(doc, tables["gender_categorical"].head(60))

    doc.add_heading("4. Reliability Summary", level=1)
    add_docx_table(doc, tables["reliability"])

    doc.add_heading("5. Correlation Heatmap", level=1)
    if figures.get("correlation") and figures["correlation"].exists():
        doc.add_picture(str(figures["correlation"]), width=Inches(6.5))

    doc.add_heading("6. SEM Path Diagram", level=1)
    if figures.get("paths") and figures["paths"].exists():
        doc.add_picture(str(figures["paths"]), width=Inches(6.5))

    doc.add_heading("7. Path Coefficients", level=1)
    add_docx_table(doc, tables["paths"].head(20))

    doc.add_heading("8. Mediation Analysis", level=1)
    add_docx_table(doc, tables["mediation"].head(20))

    doc.add_heading("9. Interpretation Notes", level=1)
    doc.add_paragraph(
        "Likert-scale results show the percentage of respondents who agreed or strongly agreed. "
        "Yes/No results show the share of respondents who answered Yes. Quantitative results show averages "
        "and other descriptive statistics. Gender-disaggregated results compare Men and Women only."
    )

    doc.save(outpath)
    return outpath


# -------------------------------------------------------------------
# Streamlit app
# -------------------------------------------------------------------

uploaded = st.file_uploader("Upload EMERGE newcomer dataset", type=["csv", "xlsx", "xls"])

if not uploaded:
    st.info(
        "Upload your dataset to begin. The app expects variables such as "
        "cred_req_difficulty, disc_hiring, rent_burden, professional_contacts, "
        "burnout, online_job_confidence, and employment_confidence."
    )
    st.stop()

raw = load_data(uploaded)
df = normalize_columns(raw)
df = coerce_analysis_columns(df)

mapping = variable_mapping(df)

numeric_cols = detect_numeric_variables(df)
year_cols = detect_year_columns(df)
likert_cols = detect_likert_variables(df)
yes_no_cols = detect_yes_no_variables(df)
categorical_cols = detect_categorical_variables(df, numeric_cols=numeric_cols)

# Variables with no variance are still shown in tables, but visuals are skipped to avoid unhelpful 100% charts.
skipped_visuals = get_no_variance_columns(df, list(dict.fromkeys(numeric_cols + categorical_cols + yes_no_cols + likert_cols)))

gender_col = find_gender_column(df)
df_gender, gender_group_col = standardize_gender_for_men_women(df, gender_col)

st.sidebar.header("Reader-Friendly Analysis Options")
st.sidebar.caption("Choose additional grouping variables if you want analysis beyond gender.")

available_group_vars = [
    col for col in categorical_cols
    if col != gender_col and df[col].nunique(dropna=True) <= 20
]

selected_group_vars = st.sidebar.multiselect(
    "Optional: choose other categorical variables for disaggregated analysis",
    options=available_group_vars,
    default=[]
)

# Original SEM workflow remains unchanged
scores = construct_scores(df)
desc = descriptive_stats(df)
rel, loadings = reliability_table(df)
corr = correlation_matrix(scores)
paths = path_coefficients(scores)
med = mediation_analysis(scores)
_, sem_estimates, sem_fit = semopy_model(scores)

# New descriptive summaries
quant_summary = quantitative_summary(df, numeric_cols)
categorical_freq = categorical_frequency_summary(df, categorical_cols)
likert_summary = likert_agree_summary(df, likert_cols)
yes_no_results = yes_no_summary(df, yes_no_cols)
likert_distribution = likert_full_distribution(df, likert_cols)

# Gender-disaggregated summaries: Men and Women only
gender_quant = quantitative_summary_by_group(df_gender, numeric_cols, gender_group_col)
gender_quant_comparison = quantitative_gender_comparison_table(df_gender, numeric_cols, gender_group_col)
gender_likert = likert_summary_by_group(df_gender, likert_cols, gender_group_col)
gender_likert_distribution = likert_full_distribution_by_group(df_gender, likert_cols, gender_group_col)
gender_yes_no = yes_no_summary_by_group(df_gender, yes_no_cols, gender_group_col)
gender_categorical = categorical_frequency_by_group(df_gender, categorical_cols, gender_group_col)

# Optional categorical disaggregation selected by client/user
optional_group_tables = {}
for group_col in selected_group_vars:
    optional_group_tables[f"{group_col}_quantitative"] = quantitative_by_category_comparison(df, numeric_cols, group_col)
    optional_group_tables[f"{group_col}_categorical_proportions"] = categorical_proportion_by_category(df, categorical_cols, group_col)
    optional_group_tables[f"{group_col}_likert"] = likert_summary_by_group(df, likert_cols, group_col)
    optional_group_tables[f"{group_col}_likert_full_distribution"] = likert_full_distribution_by_group(df, likert_cols, group_col)
    optional_group_tables[f"{group_col}_yes_no"] = yes_no_summary_by_group(df, yes_no_cols, group_col)
    optional_group_tables[f"{group_col}_categorical"] = categorical_frequency_by_group(df, categorical_cols, group_col)

# Save figures
fig_corr = OUTPUT_ROOT / "figures" / "correlation_heatmap.png"
fig_paths = OUTPUT_ROOT / "figures" / "sem_path_diagram.png"
fig_likert = OUTPUT_ROOT / "figures" / "likert_bar_chart.png"
fig_gender_likert = OUTPUT_ROOT / "figures" / "gender_likert_bar_chart.png"

if not corr.empty:
    save_corr_heatmap(corr, fig_corr)

save_path_diagram(paths, fig_paths)
save_likert_bar_chart(likert_summary, fig_likert)
save_grouped_likert_chart(gender_likert, fig_gender_likert)

histogram_paths = save_numeric_histograms(df, numeric_cols, OUTPUT_ROOT / "figures")
pie_paths = save_yes_no_pie_charts(df, yes_no_cols, OUTPUT_ROOT / "figures")
categorical_chart_paths = save_categorical_bar_charts(df, categorical_cols, OUTPUT_ROOT / "figures")

# Save processed outputs
processed_path = OUTPUT_ROOT / "processed_data" / "cleaned_emerge_dataset.csv"
scores_path = OUTPUT_ROOT / "processed_data" / "construct_scores.csv"

df.to_csv(processed_path, index=False)
scores.to_csv(scores_path, index=False)

all_tables = {
    "variable_mapping": mapping,
    "descriptive_stats": desc,
    "quantitative_summary": quant_summary,
    "categorical_frequency": categorical_freq,
    "likert_agree_summary": likert_summary,
    "likert_full_distribution": likert_distribution,
    "yes_no_summary": yes_no_results,
    "year_time_columns": pd.DataFrame({"Year / time-like columns treated as categorical": year_cols}),
    "visuals_skipped_no_variance": skipped_visuals,
    "gender_quantitative": gender_quant,
    "gender_quantitative_comparison": gender_quant_comparison,
    "gender_likert": gender_likert,
    "gender_likert_full_distribution": gender_likert_distribution,
    "gender_yes_no": gender_yes_no,
    "gender_categorical": gender_categorical,
    "reliability": rel,
    "factor_loadings": loadings,
    "correlation_matrix": corr.reset_index(),
    "paths": paths,
    "mediation": med,
    "sem_estimates": sem_estimates,
    "sem_fit": sem_fit,
}

# Add optional selected group analysis into Excel workbook
all_tables.update(optional_group_tables)

excel_path = OUTPUT_ROOT / "tables" / "emerge_sem_results.xlsx"
write_excel(all_tables, excel_path)

report_path = OUTPUT_ROOT / "model_outputs" / "emerge_sem_report.docx"
create_word_report(
    all_tables,
    {
        "correlation": fig_corr,
        "paths": fig_paths,
        "likert_bar": fig_likert,
        "gender_likert": fig_gender_likert,
    },
    report_path
)

# -------------------------------------------------------------------
# Display results
# -------------------------------------------------------------------

summary_insights = build_summary_insights(
    df=df,
    quant_summary=quant_summary,
    likert_summary=likert_summary,
    yes_no_results=yes_no_results,
    gender_comparison=gender_quant_comparison,
)

professional_summary_text = build_professional_summary_text(
    df=df,
    quant_summary=quant_summary,
    categorical_freq=categorical_freq,
    likert_summary=likert_summary,
    likert_distribution=likert_distribution,
    yes_no_results=yes_no_results,
    gender_comparison=gender_quant_comparison,
)
professional_report_path = OUTPUT_ROOT / "model_outputs" / "professional_automated_summary_report.docx"
save_professional_summary_docx(professional_summary_text, professional_report_path)

st.markdown(
    """
    <div class="dashboard-hero">
        <div class="hero-title">Analysis Results and Key Findings</div>
        <div class="hero-text">The most important findings are presented first, followed by detailed tables, visuals, SEM outputs, and downloadable files.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

metric_cols = st.columns(4)
metric_cols[0].metric("Records", f"{df.shape[0]:,}")
metric_cols[1].metric("Variables", f"{df.shape[1]:,}")
metric_cols[2].metric("Quantitative variables", f"{len(numeric_cols):,}")
metric_cols[3].metric("Categorical variables", f"{len(categorical_cols):,}")

if year_cols:
    st.info("Year/time-like columns were treated as categorical/time variables, not quantitative variables: " + ", ".join(map(str, year_cols)))
if not skipped_visuals.empty:
    st.warning(f"{len(skipped_visuals)} variable(s) had no variance and were skipped from visuals to avoid unhelpful 100% charts. See the Visual Dashboard note for details.")

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown("#### Automated Summary Insights")
st.caption("These automated insights help reviewers quickly identify the most important findings before reading detailed tables.")
for insight in summary_insights:
    st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)

with st.expander("Generate professional automated summary report and recommendations", expanded=False):
    st.markdown("The report below is automatically drafted for professional reviewers. It summarizes findings and provides practical recommendations.")
    st.markdown(professional_summary_text)
    with open(professional_report_path, "rb") as f:
        st.download_button(
            "Download professional automated summary report",
            f,
            file_name="professional_automated_summary_report.docx"
        )
st.markdown('</div>', unsafe_allow_html=True)

st.markdown("### Priority Analysis Tables")
priority_tabs = st.tabs([
    "Quantitative by Gender",
    "Averages for Quantitative Variables",
    "Categorical Frequencies",
    "Likert and Yes/No Summary",
    "Full Likert Distribution"
])

with priority_tabs[0]:
    st.markdown("#### Variable-by-variable quantitative analysis: Men and Women")
    st.caption("Difference is calculated as Men average minus Women average. P-values use Welch's t-test when scipy is installed and sufficient observations are available.")
    if gender_col is None:
        st.warning("No gender column was detected. Please include a gender, sex, respondent_gender, or participant_gender column.")
    elif gender_quant_comparison.empty:
        st.warning("No valid Men/Women comparison could be generated from the available data.")
    else:
        st.dataframe(gender_quant_comparison, use_container_width=True)

with priority_tabs[1]:
    st.markdown("#### Averages and descriptive statistics for all quantitative variables")
    st.dataframe(quant_summary, use_container_width=True)

with priority_tabs[2]:
    st.markdown("#### Frequency distribution for categorical variables")
    st.dataframe(categorical_freq, use_container_width=True)

with priority_tabs[3]:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("#### Likert: Agree or Strongly Agree")
        st.dataframe(likert_summary, use_container_width=True)
    with c2:
        st.markdown("#### Yes/No: Percentage Yes")
        st.dataframe(yes_no_results, use_container_width=True)

with priority_tabs[4]:
    st.markdown("#### Full Likert-scale distribution")
    st.caption("This table preserves the Likert response values as Strongly Disagree, Disagree, Neutral, Agree, and Strongly Agree, with frequency and percentage.")
    display_variable_tables(likert_distribution, "Likert Variable", title_prefix="Likert distribution: ")

st.markdown("### Detailed Results")

tabs = st.tabs([
    "Optional Category Analysis",
    "Visual Dashboard",
    "Dataset Overview",
    "Variable Mapping",
    "Gender Details",
    "Reliability",
    "Correlations",
    "Paths",
    "Mediation",
    "SEM fit"
])

with tabs[0]:
    st.subheader("Optional Disaggregated Analysis by Client-Selected Categorical Variables")
    st.caption("Use the sidebar to choose variables such as city, region, education level, employment status, country of origin, or age group. For variables with more than two categories, the app shows category-level averages/proportions and comparison statistics where appropriate.")

    if not selected_group_vars:
        st.info("No additional categorical variable selected. Use the sidebar to choose one or more variables for category-level analysis.")
    else:
        for group_col in selected_group_vars:
            st.markdown(f"### Analysis by {group_col}")

            st.markdown("#### Quantitative variables by category with comparison p-values")
            display_variable_tables(
                optional_group_tables.get(f"{group_col}_quantitative", pd.DataFrame()),
                "Variable / Indicator",
                title_prefix="Quantitative variable: "
            )

            st.markdown("#### Categorical proportions by category with chi-square p-values")
            display_variable_tables(
                optional_group_tables.get(f"{group_col}_categorical_proportions", pd.DataFrame()),
                "Categorical Variable",
                title_prefix="Categorical variable: "
            )

            st.markdown("#### Likert Agreement by Category")
            display_variable_tables(
                optional_group_tables.get(f"{group_col}_likert", pd.DataFrame()),
                "Variable",
                title_prefix="Likert agreement item: "
            )

            st.markdown("#### Full Likert Scale Distribution by Category")
            display_variable_tables(
                optional_group_tables.get(f"{group_col}_likert_full_distribution", pd.DataFrame()),
                "Likert Variable",
                title_prefix="Likert distribution item: "
            )

            st.markdown("#### Yes/No Responses by Category")
            display_variable_tables(
                optional_group_tables.get(f"{group_col}_yes_no", pd.DataFrame()),
                "Variable",
                title_prefix="Yes/No variable: "
            )

with tabs[1]:
    st.subheader("Reader-Friendly Visual Dashboard")
    st.caption("Figures are intentionally compact and displayed four per row for easier scanning. Values are displayed directly on charts where possible.")

    if not skipped_visuals.empty:
        st.info("The following variables were skipped from visuals because they have no variance/all values are the same. They remain available in the data tables and downloadable outputs.")
        st.dataframe(skipped_visuals, use_container_width=True)

    if year_cols:
        st.info("Year/time-like columns are treated as categorical/time variables and are not included in quantitative histograms: " + ", ".join(map(str, year_cols)))

    st.markdown("#### Key Summary Visuals")
    main_visuals = [p for p in [fig_likert, fig_gender_likert, fig_corr, fig_paths] if p.exists()]
    show_figures_grid(main_visuals, columns_per_row=4)

    st.markdown("#### Quantitative Histograms")
    show_figures_grid(histogram_paths, columns_per_row=4)

    st.markdown("#### Categorical Bar Charts")
    show_figures_grid(categorical_chart_paths, columns_per_row=4)

    st.markdown("#### Yes/No Pie Charts")
    show_figures_grid(pie_paths, columns_per_row=4)

with tabs[2]:
    st.subheader("Dataset overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Rows", f"{df.shape[0]:,}")
    col2.metric("Columns", f"{df.shape[1]:,}")
    col3.metric("Duplicated rows", int(df.duplicated().sum()))
    st.dataframe(df.head(), use_container_width=True)

with tabs[3]:
    st.subheader("Variable mapping")
    st.dataframe(mapping, use_container_width=True)

with tabs[4]:
    st.subheader("Gender-Disaggregated Details: Men and Women Only")
    if gender_col is None:
        st.warning("No gender column was detected.")
    elif df_gender.empty:
        st.warning("A gender column was detected, but no valid Men/Women records were found after recoding.")
    else:
        st.caption(f"Detected gender column: {gender_col}. Other gender responses are excluded only from this Men/Women comparison section.")
        st.markdown("#### Required comparison table: quantitative variables by gender")
        st.dataframe(gender_quant_comparison, use_container_width=True)
        st.markdown("#### Detailed quantitative averages by gender")
        st.dataframe(gender_quant, use_container_width=True)
        st.markdown("#### Likert agreement by gender")
        st.dataframe(gender_likert, use_container_width=True)
        st.markdown("#### Full Likert scale distribution by gender")
        display_variable_tables(gender_likert_distribution, "Likert Variable", title_prefix="Likert distribution by gender: ")
        if fig_gender_likert.exists():
            st.image(str(fig_gender_likert), use_container_width=True)
        st.markdown("#### Yes/No responses by gender")
        st.dataframe(gender_yes_no, use_container_width=True)
        st.markdown("#### Categorical frequencies by gender")
        st.dataframe(gender_categorical, use_container_width=True)

with tabs[5]:
    st.dataframe(rel, use_container_width=True)
    st.dataframe(loadings, use_container_width=True)

with tabs[6]:
    if fig_corr.exists():
        st.image(str(fig_corr), use_container_width=True)
    st.dataframe(corr, use_container_width=True)

with tabs[7]:
    if fig_paths.exists():
        st.image(str(fig_paths), use_container_width=True)
    st.dataframe(paths, use_container_width=True)

with tabs[8]:
    st.dataframe(med, use_container_width=True)

with tabs[9]:
    st.dataframe(sem_estimates, use_container_width=True)
    st.dataframe(sem_fit, use_container_width=True)

# -------------------------------------------------------------------
# Download outputs
# -------------------------------------------------------------------

st.subheader("Download outputs")

c1, c2, c3, c4, c5 = st.columns(5)

with open(excel_path, "rb") as f:
    c1.download_button(
        "Download Excel tables",
        f,
        file_name="emerge_sem_results.xlsx"
    )

with open(report_path, "rb") as f:
    c2.download_button(
        "Download Word report",
        f,
        file_name="emerge_sem_report.docx"
    )

with open(processed_path, "rb") as f:
    c3.download_button(
        "Download cleaned CSV",
        f,
        file_name="cleaned_emerge_dataset.csv"
    )

with open(scores_path, "rb") as f:
    c4.download_button(
        "Download construct scores",
        f,
        file_name="construct_scores.csv"
    )

with open(professional_report_path, "rb") as f:
    c5.download_button(
        "Download reviewer summary",
        f,
        file_name="professional_automated_summary_report.docx"
    )
