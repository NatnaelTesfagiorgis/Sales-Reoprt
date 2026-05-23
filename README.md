# Habesha Volume Performance Dashboard

This is the shareable Streamlit dashboard package.

## Required files for deployment

```text
app.py
requirements.txt
data/final_kpi_report.csv
```

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload `app.py`, `requirements.txt`, and the `data` folder.
3. Put your latest `final_kpi_report.csv` inside `data/`.
4. Deploy the repo on Streamlit Community Cloud and select `app.py` as the main file.
5. Share the generated URL.

## Important

If your GitHub repo is public, your data is effectively public. Use a private repo if the CSV contains confidential sales data.
