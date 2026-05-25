import streamlit as st
from pathlib import Path
import streamlit.components.v1 as components

st.set_page_config(
    page_title="Habesha Volume Dashboard",
    layout="wide"
)

html_file = Path("habesha_volume_dashboard.html")

if not html_file.exists():
    st.error("habesha_volume_dashboard.html was not found.")
else:
    html_content = html_file.read_text(encoding="utf-8")
    components.html(html_content, height=2400, scrolling=True)
