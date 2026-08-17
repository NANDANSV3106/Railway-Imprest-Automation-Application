"""
Small shared helper for explicitly seeding a Streamlit widget's default
value in session_state before the widget is created. Used instead of
relying on number_input's implicit `value=` handling, which can behave
inconsistently once a key has been touched elsewhere in the app.
"""
import streamlit as st


def init_default(key, default):
    if key not in st.session_state:
        st.session_state[key] = default
