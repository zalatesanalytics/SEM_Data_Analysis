# EMERGE+ SEM Streamlit Analyzer

This repository contains a GitHub-ready Streamlit app for the EMERGE+ Newcomer Professional Integration SEM analysis.

## What the app does

- Uploads CSV or Excel datasets.
- Normalizes column names for Python compatibility.
- Maps expected EMERGE+ SEM variables to uploaded dataset columns.
- Computes descriptive statistics, reliability, factor-loading proxies, construct scores, correlations, SEM path coefficients, mediation tests, and SEM model estimates when `semopy` can run.
- Displays results in Streamlit as tables and figures.
- Exports cleaned data, Excel tables, PNG figures, and an automated Word report.

## Folder structure

```text
.
├── app.py
├── sem_engine.py
├── sem_config.py
├── requirements.txt
├── README.md
├── .streamlit/config.toml
└── outputs/
    ├── tables/
    ├── figures/
    ├── processed_data/
    └── model_outputs/
```

## Local setup

```bash
python -m venv .venv
.venv\Scripts\activate   # Windows PowerShell
pip install -r requirements.txt
streamlit run app.py
```

For Mac/Linux:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## GitHub upload

```bash
git init
git add .
git commit -m "Initial EMERGE SEM Streamlit app"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```

## Important variable naming note

The app expects variables using normalized names such as:

- `cred_req_difficulty`
- `disc_hiring`
- `rent_burden`
- `professional_contacts`
- `burnout`
- `online_job_confidence`
- `mentor_sector_match`
- `employment_confidence`
- `job_satisfaction`
- `financial_stability`

If your uploaded dataset uses slightly different names, either rename the columns in your dataset or update `sem_config.py`.

## Streamlit Cloud deployment

1. Push this repository to GitHub.
2. Go to Streamlit Community Cloud.
3. Create a new app from your GitHub repository.
4. Select `app.py` as the entry point.
5. Deploy.

## Methodological note

This app provides a practical MEAL-oriented SEM workflow. Composite-score path coefficients and mediation tests are included for operational interpretation. `semopy` is used for a more formal SEM estimation where sufficient complete variables and sample size are available.
