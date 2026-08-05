import streamlit as st

st.set_page_config(page_title="ESPI — Pallets y conciliación", layout="wide")

with st.sidebar:
    st.header("Modo")
    modo = st.radio(
        "¿Qué quieres hacer?",
        ["🚚 Pallets y embarques", "📝 Captura de monitoreos", "📂 Reconciliar Excel existente"],
        label_visibility="collapsed",
    )
    st.divider()

if modo == "🚚 Pallets y embarques":
    import mode_pallets
    mode_pallets.render()
elif modo == "📝 Captura de monitoreos":
    import mode_captura
    mode_captura.render()
else:
    import mode_excel
    mode_excel.render()
