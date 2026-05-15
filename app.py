from pathlib import Path
import matplotlib.pyplot as plt
import networkx as nx
import pandas as pd
import seaborn as sns
import streamlit as st
from docx import Document
from docx.shared import Inches

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
    "visualize results, and generate a Word report."
)

OUTPUT_ROOT = Path("outputs")
for sub in ["tables", "figures", "processed_data", "model_outputs"]:
    (OUTPUT_ROOT / sub).mkdir(parents=True, exist_ok=True)


def add_docx_table(doc, df: pd.DataFrame):
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
    return df.select_dtypes(include="number").columns.tolist()


def detect_likert_variables(df):
    likert_cols = []
    for col in df.columns:
        values = df[col].dropna().unique()
        numeric_values = pd.to_numeric(pd.Series(values), errors="coerce").dropna()
        if len(numeric_values) > 0 and numeric_values.between(1, 5).all():
            likert_cols.append(col)
    return likert_cols


def detect_yes_no_variables(df):
    yes_no_cols = []
    for col in df.columns:
        values = df[col].dropna().astype(str).str.lower().str.strip().unique()
        if len(values) > 0 and set(values).issubset({"yes", "no", "1", "0", "true", "false"}):
            yes_no_cols.append(col)
    return yes_no_cols


def likert_agree_summary(df, likert_cols):
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
            "Yes responses": yes_count,
            "Percentage Yes": round((yes_count / total) * 100, 2)
        })

    return pd.DataFrame(results)


def quantitative_summary(df, numeric_cols):
    if not numeric_cols:
        return pd.DataFrame()

    return df[numeric_cols].describe().T.reset_index().rename(columns={
        "index": "Variable",
        "count": "Valid responses",
        "mean": "Mean",
        "std": "Standard deviation",
        "min": "Minimum",
        "25%": "25th percentile",
        "50%": "Median",
        "75%": "75th percentile",
        "max": "Maximum"
    })


def save_likert_bar_chart(summary_df, outpath):
    if summary_df.empty:
        return

    fig, ax = plt.subplots(figsize=(12, 7))
    plot_df = summary_df.sort_values("Percentage Agree or Above", ascending=True)

    ax.barh(plot_df["Variable"], plot_df["Percentage Agree or Above"])
    ax.set_xlabel("Percentage Agree or Strongly Agree")
    ax.set_title("Likert-Scale Barrier Analysis")
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def save_numeric_histograms(df, numeric_cols, output_dir):
    saved_paths = []

    for col in numeric_cols[:8]:
        series = pd.to_numeric(df[col], errors="coerce").dropna()

        if series.empty:
            continue

        fig, ax = plt.subplots(figsize=(8, 5))
        ax.hist(series, bins=20)
        ax.set_title(f"Distribution of {col}")
        ax.set_xlabel(col)
        ax.set_ylabel("Frequency")
        fig.tight_layout()

        outpath = output_dir / f"histogram_{col}.png"
        fig.savefig(outpath, dpi=300)
        plt.close(fig)
        saved_paths.append(outpath)

    return saved_paths


def save_yes_no_pie_charts(df, yes_no_cols, output_dir):
    saved_paths = []

    for col in yes_no_cols[:6]:
        series = df[col].dropna().astype(str).str.lower().str.strip()
        yes_count = series.isin(["yes", "1", "true"]).sum()
        no_count = series.isin(["no", "0", "false"]).sum()

        if yes_count + no_count == 0:
            continue

        fig, ax = plt.subplots(figsize=(6, 6))
        ax.pie([yes_count, no_count], labels=["Yes", "No"], autopct="%1.1f%%")
        ax.set_title(f"Yes/No Distribution: {col}")

        outpath = output_dir / f"pie_{col}.png"
        fig.savefig(outpath, dpi=300)
        plt.close(fig)
        saved_paths.append(outpath)

    return saved_paths


def save_corr_heatmap(corr: pd.DataFrame, outpath: Path):
    fig, ax = plt.subplots(figsize=(10, 7))
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm", center=0, ax=ax)
    ax.set_title("Correlation Matrix")
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def save_path_diagram(paths: pd.DataFrame, outpath: Path):
    fig, ax = plt.subplots(figsize=(11, 8))
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


def create_word_report(tables, figures, outpath: Path):
    doc = Document()

    doc.add_heading("EMERGE+ Newcomer SEM Analysis Report", 0)

    doc.add_paragraph(
        "This automated report summarizes data readiness, descriptive statistics, "
        "Likert-scale barrier analysis, Yes/No response patterns, construct reliability, "
        "correlations, path coefficients, mediation tests, and SEM model fit where available."
    )

    doc.add_heading("1. Variable Mapping", level=1)
    add_docx_table(doc, tables["variable_mapping"].head(40))

    doc.add_heading("2. Descriptive Analysis", level=1)

    doc.add_heading("2.1 Quantitative Variables", level=2)
    if not tables["quantitative_summary"].empty:
        add_docx_table(doc, tables["quantitative_summary"].head(30))
    else:
        doc.add_paragraph("No quantitative variables were detected.")

    doc.add_heading("2.2 Likert-Scale Agree or Above Analysis", level=2)
    if not tables["likert_agree_summary"].empty:
        add_docx_table(doc, tables["likert_agree_summary"].head(30))
    else:
        doc.add_paragraph("No Likert-scale variables were detected.")

    if figures.get("likert_bar") and figures["likert_bar"].exists():
        doc.add_picture(str(figures["likert_bar"]), width=Inches(6.5))

    doc.add_heading("2.3 Yes/No Variables", level=2)
    if not tables["yes_no_summary"].empty:
        add_docx_table(doc, tables["yes_no_summary"].head(30))
    else:
        doc.add_paragraph("No Yes/No variables were detected.")

    doc.add_heading("3. Reliability Summary", level=1)
    if not tables["reliability"].empty:
        add_docx_table(doc, tables["reliability"])
    else:
        doc.add_paragraph("No reliability table could be generated.")

    doc.add_heading("4. Correlation Heatmap", level=1)
    if figures.get("correlation") and figures["correlation"].exists():
        doc.add_picture(str(figures["correlation"]), width=Inches(6.5))

    doc.add_heading("5. SEM Path Diagram", level=1)
    if figures.get("paths") and figures["paths"].exists():
        doc.add_picture(str(figures["paths"]), width=Inches(6.5))

    doc.add_heading("6. Path Coefficients", level=1)
    if not tables["paths"].empty:
        add_docx_table(doc, tables["paths"].head(20))

    doc.add_heading("7. Mediation Analysis", level=1)
    if not tables["mediation"].empty:
        add_docx_table(doc, tables["mediation"].head(20))

    doc.add_heading("8. Interpretation Notes", level=1)
    doc.add_paragraph(
        "Likert-scale results show the percentage of respondents who agreed or strongly agreed "
        "that a specific issue is a barrier. Yes/No results show the share of respondents who answered Yes. "
        "Quantitative summaries describe central tendency and variation for variables such as age, income, and expenses."
    )

    doc.save(outpath)
    return outpath


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

st.subheader("Dataset overview")
col1, col2, col3 = st.columns(3)
col1.metric("Rows", f"{df.shape[0]:,}")
col2.metric("Columns", f"{df.shape[1]:,}")
col3.metric("Duplicated rows", int(df.duplicated().sum()))
st.dataframe(df.head(), use_container_width=True)

mapping = variable_mapping(df)
st.subheader("Variable mapping")
st.dataframe(mapping, use_container_width=True)

numeric_cols = detect_numeric_variables(df)
likert_cols = detect_likert_variables(df)
yes_no_cols = detect_yes_no_variables(df)

quant_summary = quantitative_summary(df, numeric_cols)
likert_summary = likert_agree_summary(df, likert_cols)
yes_no_results = yes_no_summary(df, yes_no_cols)

scores = construct_scores(df)
desc = descriptive_stats(df)
rel, loadings = reliability_table(df)
corr = correlation_matrix(scores)
paths = path_coefficients(scores)
med = mediation_analysis(scores)
_, sem_estimates, sem_fit = semopy_model(scores)

fig_corr = OUTPUT_ROOT / "figures" / "correlation_heatmap.png"
fig_paths = OUTPUT_ROOT / "figures" / "sem_path_diagram.png"
fig_likert = OUTPUT_ROOT / "figures" / "likert_bar_chart.png"

if not corr.empty:
    save_corr_heatmap(corr, fig_corr)

save_path_diagram(paths, fig_paths)
save_likert_bar_chart(likert_summary, fig_likert)

histogram_paths = save_numeric_histograms(df, numeric_cols, OUTPUT_ROOT / "figures")
pie_paths = save_yes_no_pie_charts(df, yes_no_cols, OUTPUT_ROOT / "figures")

processed_path = OUTPUT_ROOT / "processed_data" / "cleaned_emerge_dataset.csv"
scores_path = OUTPUT_ROOT / "processed_data" / "construct_scores.csv"

df.to_csv(processed_path, index=False)
scores.to_csv(scores_path, index=False)

all_tables = {
    "variable_mapping": mapping,
    "descriptive_stats": desc,
    "quantitative_summary": quant_summary,
    "likert_agree_summary": likert_summary,
    "yes_no_summary": yes_no_results,
    "reliability": rel,
    "factor_loadings": loadings,
    "correlation_matrix": corr.reset_index(),
    "paths": paths,
    "mediation": med,
    "sem_estimates": sem_estimates,
    "sem_fit": sem_fit,
}

excel_path = OUTPUT_ROOT / "tables" / "emerge_sem_results.xlsx"
write_excel(all_tables, excel_path)

report_path = OUTPUT_ROOT / "model_outputs" / "emerge_sem_report.docx"
create_word_report(
    all_tables,
    {
        "correlation": fig_corr,
        "paths": fig_paths,
        "likert_bar": fig_likert,
    },
    report_path
)

st.subheader("Analysis results")

tabs = st.tabs([
    "Descriptives",
    "Likert Barriers",
    "Yes/No Analysis",
    "Visuals",
    "Reliability",
    "Correlations",
    "Paths",
    "Mediation",
    "SEM fit"
])

with tabs[0]:
    st.subheader("Quantitative Descriptive Statistics")
    st.dataframe(quant_summary, use_container_width=True)

    st.subheader("Existing Descriptive Statistics")
    st.dataframe(desc, use_container_width=True)

with tabs[1]:
    st.subheader("Likert-Scale Agree or Strongly Agree Analysis")
    st.dataframe(likert_summary, use_container_width=True)

    if fig_likert.exists():
        st.image(str(fig_likert))

with tabs[2]:
    st.subheader("Yes/No Response Analysis")
    st.dataframe(yes_no_results, use_container_width=True)

with tabs[3]:
    st.subheader("Histograms for Quantitative Variables")
    for path in histogram_paths:
        st.image(str(path))

    st.subheader("Pie Charts for Yes/No Variables")
    for path in pie_paths:
        st.image(str(path))

with tabs[4]:
    st.dataframe(rel, use_container_width=True)
    st.dataframe(loadings, use_container_width=True)

with tabs[5]:
    if fig_corr.exists():
        st.image(str(fig_corr))
    st.dataframe(corr, use_container_width=True)

with tabs[6]:
    if fig_paths.exists():
        st.image(str(fig_paths))
    st.dataframe(paths, use_container_width=True)

with tabs[7]:
    st.dataframe(med, use_container_width=True)

with tabs[8]:
    st.dataframe(sem_estimates, use_container_width=True)
    st.dataframe(sem_fit, use_container_width=True)

st.subheader("Download outputs")

c1, c2, c3 = st.columns(3)

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
