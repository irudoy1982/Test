from __future__ import annotations

import os
from typing import Any

import streamlit as st

from crm_admin import render_crm_admin


APP_VERSION = "15.1-dev.3"


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
    html, body, [data-testid="stAppViewContainer"] {
        color-scheme: light !important;
    }
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"],
    .stApp {
        background: #f5f7fa !important;
        color: #111827 !important;
    }
    [data-testid="stHeader"] {
        background: rgba(245, 247, 250, 0.96) !important;
    }
    .block-container {
        max-width: 1480px;
        padding-top: 2.2rem;
    }
    div[data-testid="stForm"] {
        background: #ffffff !important;
        border: 1px solid #d8dee8 !important;
        border-radius: 8px !important;
    }
    div[data-testid="stTextInput"] div[data-baseweb="input"],
    div[data-testid="stNumberInput"] div[data-baseweb="input"],
    div[data-testid="stTextArea"] div[data-baseweb="textarea"],
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div {
        min-height: 46px !important;
        background-color: #f3f6fa !important;
        border: 1px solid #b8c3d1 !important;
        border-radius: 6px !important;
        box-shadow: none !important;
    }
    div[data-testid="stTextInput"] div[data-baseweb="input"]:hover,
    div[data-testid="stNumberInput"] div[data-baseweb="input"]:hover,
    div[data-testid="stTextArea"] div[data-baseweb="textarea"]:hover,
    div[data-testid="stSelectbox"] div[data-baseweb="select"] > div:hover {
        border-color: #728096 !important;
    }
    div[data-testid="stTextInput"] div[data-baseweb="input"]:focus-within,
    div[data-testid="stNumberInput"] div[data-baseweb="input"]:focus-within,
    div[data-testid="stTextArea"] div[data-baseweb="textarea"]:focus-within,
    div[data-testid="stSelectbox"] div[data-baseweb="select"]:focus-within > div {
        border-color: #0f766e !important;
        box-shadow: 0 0 0 2px rgba(15, 118, 110, 0.14) !important;
    }
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input,
    div[data-testid="stTextArea"] textarea,
    div[data-testid="stSelectbox"] [role="combobox"] {
        color: #151922 !important;
        -webkit-text-fill-color: #151922 !important;
        caret-color: #0f766e !important;
        opacity: 1 !important;
    }
    div[data-testid="stTextInput"] input::placeholder,
    div[data-testid="stTextArea"] textarea::placeholder {
        color: #7b8797 !important;
        -webkit-text-fill-color: #7b8797 !important;
        opacity: 1 !important;
    }
    div[data-testid="stTextInput"] svg,
    div[data-testid="stSelectbox"] svg {
        color: #151922 !important;
        fill: currentColor !important;
    }
    button[kind="primary"] { font-weight: 700; }
    </style>
    """,
    unsafe_allow_html=True,
)

render_crm_admin(APP_VERSION, get_app_secret)
