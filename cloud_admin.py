from __future__ import annotations

import os
from typing import Any

import streamlit as st

from crm_admin import render_crm_admin


APP_VERSION = "15.1-dev.1"


def get_app_secret(name: str, default: Any = None) -> Any:
    try:
        value = st.secrets.get(name)
        if value is not None:
            return value
    except Exception:
        pass
    return os.environ.get(name, default)


st.set_page_config(
    page_title="Khalil Audit Admin",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root { color-scheme: light; }
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    .stApp {
        background: #f5f7fa !important;
        color: #111827 !important;
    }
    [data-testid="stHeader"] {
        background: rgba(245, 247, 250, 0.96) !important;
    }
    [data-testid="stTextInput"] input,
    [data-testid="stSelectbox"] div[data-baseweb="select"] > div,
    [data-testid="stNumberInput"] input,
    [data-testid="stTextArea"] textarea {
        background: #ffffff !important;
        color: #111827 !important;
        border-color: #cbd5e1 !important;
    }
    [data-testid="stForm"] {
        background: #ffffff !important;
        border-color: #d7dee8 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

render_crm_admin(APP_VERSION, get_app_secret)
