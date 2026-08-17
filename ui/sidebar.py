"""
Sidebar: brand header, month/year period selector, active-period card.
"""
import streamlit as st

from config import MONTHS
from core.dates import get_days_in_month


def render_sidebar():
    """Renders the sidebar and returns (sel_month, sel_year, days_in_month, month_idx)."""
    with st.sidebar:
        st.markdown("""
        <div class="sb-brand">
            <div class="sb-brand-icon">🚆</div>
            <p class="sb-brand-name">Imprest Automator</p>
            <p class="sb-brand-sub">Indian Railways · SETTIHALLI</p>
        </div>
        """, unsafe_allow_html=True)
    
        st.markdown('<span class="sb-section-label">Billing Period</span>', unsafe_allow_html=True)
        sel_month = st.selectbox("Month", MONTHS, index=6, label_visibility="collapsed")
        sel_year  = st.selectbox("Year", ["2025","2026","2027","2028"], index=1, label_visibility="collapsed")
    
        days_in_month, month_idx = get_days_in_month(sel_month, sel_year)
    
        st.markdown(f"""
        <div class="sb-period-card">
            <p class="sb-period-label-txt">Active Period</p>
            <p class="sb-period-full">{sel_month} {sel_year}</p>
            <div class="sb-period-row">
                <div class="sb-period-chip">
                    <p class="sb-chip-label">From</p>
                    <p class="sb-chip-val">01/{month_idx:02d}/{sel_year}</p>
                </div>
                <div class="sb-period-chip">
                    <p class="sb-chip-label">To</p>
                    <p class="sb-chip-val">{days_in_month}/{month_idx:02d}/{sel_year}</p>
                </div>
            </div>
            <div class="sb-period-row">
                <div class="sb-period-chip">
                    <p class="sb-chip-label">Days</p>
                    <p class="sb-chip-val">{days_in_month}</p>
                </div>
                <div class="sb-period-chip">
                    <p class="sb-chip-label">Sanctioned</p>
                    <p class="sb-chip-val">₹25,000</p>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
        st.markdown("""
        <div class="sb-footer">
            <p>Daily Wage · ₹380<br>Sanctioned Amount · ₹25,000</p>
        </div>
        """, unsafe_allow_html=True)
    

    return sel_month, sel_year, days_in_month, month_idx
