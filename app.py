import streamlit as st
from pathlib import Path
import streamlit.components.v1 as components
import json

st.set_page_config(
    page_title="Habesha Volume Dashboard",
    layout="wide"
)

html_path = Path("habesha_volume_dashboard.html")
csv_path = Path("final_kpi_report.csv")
#qqq
if not html_path.exists():
    st.error("habesha_volume_dashboard.html was not found.")
    st.stop()

if not csv_path.exists():
    st.error("final_kpi_report.csv was not found.")
    st.stop()

html_content = html_path.read_text(encoding="utf-8")
csv_content = csv_path.read_text(encoding="utf-8-sig")

# Inject CSV content directly into the HTML so the dashboard loads automatically
embedded_csv = json.dumps(csv_content)

replacement_code = f"""
Papa.parse({embedded_csv}, {{
    header: true,
    skipEmptyLines: true,
    transformHeader: normalizeHeader,
    complete: (results) => {{
        loadData(results.data || []);
    }},
    error: (err) => {{
        setStatus("Could not read embedded final_kpi_report.csv: " + err.message, "error");
    }}
}});
"""

html_content = html_content.replace("loadDefaultCsv();", replacement_code)

components.html(
    html_content,
    height=2600,
    scrolling=True
)
