from pathlib import Path
import shutil
import tempfile

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
st.caption("Upload a CSV or Excel file, clean variables, estimate SEM pathways, visualize results, and generate a Word report.")

OUTPUT_ROOT = Path("outputs")
for sub in ["tables", "figures", "processed_data", "model_outputs"]:
    (OUTPUT_ROOT / sub).mkdir(parents=True, exist_ok=True)


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
        labels = {(r["From"], r["To"]): f"{r['Beta/std r']:.2f}" for _, r in paths.iterrows()}
        nx.draw_networkx_edge_labels(g, pos, edge_labels=labels, font_size=9, ax=ax)
    ax.set_axis_off()
    fig.tight_layout()
    fig.savefig(outpath, dpi=300)
    plt.close(fig)


def create_word_report(tables, figures, outpath: Path):
    doc = Document()
    doc.add_heading("EMERGE+ Newcomer SEM Analysis Report", 0)
    doc.add_paragraph("This automated report summarizes data readiness, construct reliability, correlations, path coefficients, mediation tests, and SEM model fit where available.")

    doc.add_heading("1. Variable Mapping", level=1)
    doc.add_paragraph("The table below identifies expected SEM variables and whether they were found in the uploaded dataset.")
    add_docx_table(doc, tables["variable_mapping"].head(40))

    doc.add_heading("2. Reliability Summary", level=1)
    if not tables["reliability"].empty:
        add_docx_table(doc, tables["reliability"])
    else:
        doc.add_paragraph("No reliability table could be generated because insufficient construct indicators were found.")

    doc.add_heading("3. Correlation Heatmap", level=1)
    if figures.get("correlation") and figures["correlation"].exists():
        doc.add_picture(str(figures["correlation"]), width=Inches(6.5))

    doc.add_heading("4. SEM Path Diagram", level=1)
    if figures.get("paths") and figures["paths"].exists():
        doc.add_picture(str(figures["paths"]), width=Inches(6.5))

    doc.add_heading("5. Path Coefficients", level=1)
    if not tables["paths"].empty:
        add_docx_table(doc, tables["paths"].head(20))

    doc.add_heading("6. Mediation Analysis", level=1)
    if not tables["mediation"].empty:
        add_docx_table(doc, tables["mediation"].head(20))

    doc.add_heading("7. Interpretation Notes", level=1)
    doc.add_paragraph("Positive coefficients suggest reinforcing pathways; negative coefficients suggest constraining pathways. Statistical significance should be interpreted together with theory, sample size, measurement quality, and construct reliability.")
    doc.add_paragraph("For donor and MEAL reporting, prioritize pathways that are statistically meaningful, programmatically actionable, and aligned with the EMERGE+ learning questions.")
    doc.save(outpath)
    return outpath


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


uploaded = st.file_uploader("Upload EMERGE newcomer dataset", type=["csv", "xlsx", "xls"])

if not uploaded:
    st.info("Upload your dataset to begin. The app expects variables such as cred_req_difficulty, disc_hiring, rent_burden, professional_contacts, burnout, online_job_confidence, and employment_confidence.")
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

scores = construct_scores(df)
desc = descriptive_stats(df)
rel, loadings = reliability_table(df)
corr = correlation_matrix(scores)
paths = path_coefficients(scores)
med = mediation_analysis(scores)
_, sem_estimates, sem_fit = semopy_model(scores)

fig_corr = OUTPUT_ROOT / "figures" / "correlation_heatmap.png"
fig_paths = OUTPUT_ROOT / "figures" / "sem_path_diagram.png"
if not corr.empty:
    save_corr_heatmap(corr, fig_corr)
save_path_diagram(paths, fig_paths)

processed_path = OUTPUT_ROOT / "processed_data" / "cleaned_emerge_dataset.csv"
scores_path = OUTPUT_ROOT / "processed_data" / "construct_scores.csv"
df.to_csv(processed_path, index=False)
scores.to_csv(scores_path, index=False)

all_tables = {
    "variable_mapping": mapping,
    "descriptive_stats": desc,
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
create_word_report(all_tables, {"correlation": fig_corr, "paths": fig_paths}, report_path)

st.subheader("Analysis results")
tabs = st.tabs(["Descriptives", "Reliability", "Correlations", "Paths", "Mediation", "SEM fit"])
with tabs[0]: st.dataframe(desc, use_container_width=True)
with tabs[1]:
    st.dataframe(rel, use_container_width=True)
    st.dataframe(loadings, use_container_width=True)
with tabs[2]:
    if fig_corr.exists(): st.image(str(fig_corr))
    st.dataframe(corr, use_container_width=True)
with tabs[3]:
    if fig_paths.exists(): st.image(str(fig_paths))
    st.dataframe(paths, use_container_width=True)
with tabs[4]: st.dataframe(med, use_container_width=True)
with tabs[5]:
    st.dataframe(sem_estimates, use_container_width=True)
    st.dataframe(sem_fit, use_container_width=True)

st.subheader("Download outputs")
c1, c2, c3 = st.columns(3)
with open(excel_path, "rb") as f:
    c1.download_button("Download Excel tables", f, file_name="emerge_sem_results.xlsx")
with open(report_path, "rb") as f:
    c2.download_button("Download Word report", f, file_name="emerge_sem_report.docx")
with open(processed_path, "rb") as f:
    c3.download_button("Download cleaned CSV", f, file_name="cleaned_emerge_dataset.csv")
