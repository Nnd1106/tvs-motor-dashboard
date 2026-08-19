# TVS Motor Company — Equity Research Dashboard

![Python 3.11](https://img.shields.io/badge/python-3.11-blue) ![Streamlit](https://img.shields.io/badge/streamlit-1.35%2B-red) ![GitHub Actions](https://img.shields.io/badge/automation-GitHub%20Actions-blue)

A fully automated equity research dashboard for TVS Motor Company. A GitHub
Actions workflow refreshes live market data into the underlying Excel
valuation model every weekday morning, recalculates it headlessly, and
publishes the results as JSON that a Streamlit app renders as an interactive
SOTP, DCF, relative valuation, sensitivity, and Monte Carlo dashboard.

## Architecture

```
Excel Model → pipeline/refresh.py → data/outputs.json → Streamlit Dashboard
     ↑                                      ↑
yfinance / FBIL              GitHub Actions (daily, 9 AM IST)
```

## Setup

```bash
git clone <this-repo-url>
cd tvs-motor-dashboard
pip install -r requirements.txt
streamlit run dashboard/app.py
```

The pipeline needs the source workbook `TVS_Motor_Valuation_Model.xlsx` in the
project root (see the note on the Excel model below) and, for full
recalculation, a local install of LibreOffice (`soffice` on PATH). Without
LibreOffice, `pipeline/refresh.py` falls back to reading the workbook's last
cached values with a logged warning.

```bash
python pipeline/refresh.py
```

## Deployment

1. Push this repository to GitHub.
2. Go to [share.streamlit.io](https://share.streamlit.io), connect the repo.
3. Set the main file path to `dashboard/app.py`.
4. Deploy. The GitHub Actions workflow keeps `data/outputs.json` fresh daily;
   Streamlit Cloud redeploys automatically on every push to the tracked branch.

[Live Dashboard](YOUR_STREAMLIT_URL_HERE)

## Files

| File | Purpose |
|---|---|
| `TVS_Motor_Valuation_Model.xlsx` | Source valuation model (SOTP, DCF, comps, Monte Carlo) — see note below |
| `pipeline/refresh.py` | Fetches live market data, writes it into the model, recalculates, extracts outputs |
| `data/outputs.json` | Machine-readable snapshot of every dashboard input the Streamlit app reads |
| `dashboard/app.py` | The Streamlit dashboard itself |
| `.github/workflows/daily_refresh.yml` | Runs the pipeline on a weekday morning schedule and commits the refreshed JSON |
| `requirements.txt` | Python dependencies |

## A note on the Excel model and CI

`TVS_Motor_Valuation_Model.xlsx` is listed in `.gitignore` (as specified for
this project), which means a fresh `git clone` of this repository — including
the GitHub Actions runner — will **not** have the workbook available, and
`pipeline/refresh.py` will exit early with a clear "source workbook not
found" error until it's provided. Before relying on the scheduled workflow,
do one of the following:

- Force-add the workbook despite `.gitignore` (`git add -f
  TVS_Motor_Valuation_Model.xlsx`) if you're comfortable committing it, or
- Store it as a repository secret (base64-encoded) and add a decode step to
  `daily_refresh.yml` before the pipeline runs.

Locally, the workbook is already present in the project root, so
`streamlit run dashboard/app.py` and `python pipeline/refresh.py` work
immediately without any extra setup.

## Disclaimer

This dashboard is an academic and illustrative project. Nothing in it
constitutes investment advice.
