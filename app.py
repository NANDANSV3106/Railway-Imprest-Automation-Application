"""
Railway Imprest Bill Automator — entrypoint.

This file only wires modules together: page config, global CSS, sidebar,
and the three tabs. All actual logic lives in core/, docx_engine/, and ui/.
"""
import streamlit as st

from ui.styles import inject_css
from ui.sidebar import render_sidebar
from ui.tab_expenses import render_tab1
from ui.tab_labour import render_tab2
from ui.tab_generate import render_tab3

st.set_page_config(page_title="Railway Imprest Automator", page_icon="🚆", layout="wide")
inject_css()

sel_month, sel_year, days_in_month, month_idx = render_sidebar()

# ══════════════════════════════════════════════════════════════
#  PAGE HEADER
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="page-header">
    <span class="ph-icon">🚆</span>
    <div>
        <p class="ph-title">Railway Imprest Bill Automator</p>
        <p class="ph-sub">Generate Annexure I &amp; Labour Annexure for SETTIHALLI Station</p>
    </div>
    <span class="ph-badge">📅 {sel_month} {sel_year}</span>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  TABS
# ══════════════════════════════════════════════════════════════
tab1, tab2, tab3 = st.tabs(["🛒  Consumables", "👷  Labour", "📊  Generate & Download"])

with tab1:
    tab1_data = render_tab1(sel_month, sel_year, month_idx, days_in_month)

with tab2:
    tab2_data = render_tab2(sel_month, sel_year, days_in_month)

with tab3:
    render_tab3(sel_month, sel_year, days_in_month, tab1_data, tab2_data)

# ══════════════════════════════════════════════════════════════
#  FOOTER
# ══════════════════════════════════════════════════════════════
st.markdown(f"""
<div class="page-footer">
    🚆 Railway Imprest Automator &nbsp;·&nbsp; SETTIHALLI Station &nbsp;·&nbsp;
    {sel_month} {sel_year} &nbsp;·&nbsp; Indian Railways
</div>
""", unsafe_allow_html=True)
