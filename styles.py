import streamlit as st

_CSS = """
<style>
/* ---- Tarjetas de métricas (Tablero) ---- */
div[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid rgba(43,33,24,0.07);
    border-radius: 16px;
    padding: 1.1rem 1.3rem;
    box-shadow: 0 1px 2px rgba(43,33,24,0.04), 0 6px 16px rgba(43,33,24,0.06);
    transition: box-shadow .18s ease, transform .18s ease;
}
div[data-testid="stMetric"]:hover {
    box-shadow: 0 2px 4px rgba(43,33,24,0.06), 0 10px 24px rgba(43,33,24,0.10);
    transform: translateY(-2px);
}
div[data-testid="stMetricLabel"] p {
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: .05em;
    color: #9C8975;
    font-weight: 600;
}
div[data-testid="stMetricValue"] {
    font-size: 2.1rem;
    font-weight: 700;
    color: #2B2118;
}

/* ---- Gráficas ---- */
div[data-testid="stVegaLiteChart"] {
    background: #FFFFFF;
    border: 1px solid rgba(43,33,24,0.07);
    border-radius: 16px;
    padding: 1rem;
    box-shadow: 0 1px 2px rgba(43,33,24,0.04), 0 6px 16px rgba(43,33,24,0.06);
}

/* ---- Tablas ---- */
div[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    border: 1px solid rgba(43,33,24,0.07);
    box-shadow: 0 1px 2px rgba(43,33,24,0.04), 0 6px 16px rgba(43,33,24,0.06);
}

/* ---- Formularios y contenedores con borde ---- */
div[data-testid="stForm"] {
    background: #FFFFFF;
    border: 1px solid rgba(43,33,24,0.07) !important;
    border-radius: 16px !important;
    padding: 1.4rem 1.4rem 1rem 1.4rem !important;
    box-shadow: 0 1px 2px rgba(43,33,24,0.04), 0 6px 16px rgba(43,33,24,0.06);
}

/* ---- Expander (panel de respaldo) ---- */
div[data-testid="stExpander"] {
    border: 1px solid rgba(43,33,24,0.07) !important;
    border-radius: 14px !important;
    box-shadow: 0 1px 2px rgba(43,33,24,0.04), 0 4px 12px rgba(43,33,24,0.05);
    overflow: hidden;
}

/* ---- Botones ---- */
button[data-testid^="stBaseButton"] {
    border-radius: 10px !important;
    transition: box-shadow .15s ease, transform .15s ease;
}
button[data-testid^="stBaseButton"]:hover {
    box-shadow: 0 4px 12px rgba(43,33,24,0.12);
    transform: translateY(-1px);
}
button[kind="primary"], button[data-testid="stBaseButton-primary"] {
    box-shadow: 0 2px 8px rgba(232,135,30,0.35);
}

/* ---- Inputs de texto y número ---- */
div[data-testid="stTextInput"] input,
div[data-testid="stNumberInputField"],
div[data-baseweb="select"] > div {
    border-radius: 10px !important;
}

/* ---- Navegación lateral (radio como menú) ---- */
section[data-testid="stSidebar"] label[data-baseweb="radio"] {
    padding: .45rem .7rem;
    border-radius: 10px;
    margin-bottom: 2px;
    transition: background .15s ease;
}
section[data-testid="stSidebar"] label[data-baseweb="radio"]:hover {
    background: rgba(232,135,30,0.10);
}

/* ---- Encabezados con más aire ---- */
div[data-testid="stHeading"] h1 {
    letter-spacing: -0.01em;
}
div[data-testid="stHeading"] h2, div[data-testid="stHeading"] h3 {
    margin-top: .3rem;
}
</style>
"""


def inject():
    """Aplica el CSS de la app. Llamar una sola vez, al inicio del router (app.py)."""
    st.markdown(_CSS, unsafe_allow_html=True)
