import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import core


def render():
    st.title("📂 Reconciliar un Excel de LIQUIDACIÓN ya existente")
    st.caption(
        "Modo de compatibilidad para archivos viejos con el formato de bloques por día. "
        "Para capturar datos nuevos, usa mejor la pestaña 'Captura de monitoreos'."
    )

    with st.sidebar:
        st.header("Archivos")
        liq_file = st.file_uploader("Archivo LIQUIDACIÓN (ej. LIQUIDACION_ATAULFOS_2026.xlsx)", type=["xlsx"])
        rcf_file = st.file_uploader("Archivo RCF LIQUIDACIONES (manifiestos)", type=["xlsx"])
        run_btn = st.button("Ejecutar conciliación", type="primary", disabled=not (liq_file and rcf_file))

    if "results" not in st.session_state:
        st.session_state["results"] = None
        st.session_state["output_bytes"] = None
        st.session_state["skipped"] = None

    if run_btn and liq_file and rcf_file:
        with tempfile.TemporaryDirectory() as tmp:
            liq_path = Path(tmp) / "liquidacion.xlsx"
            rcf_path = Path(tmp) / "rcf.xlsx"
            out_path = Path(tmp) / "liquidacion_conciliado.xlsx"
            liq_path.write_bytes(liq_file.getvalue())
            rcf_path.write_bytes(rcf_file.getvalue())

            with st.spinner("Leyendo manifiestos y cruzando pallets..."):
                results, skipped = core.run_reconciliation(str(liq_path), str(rcf_path), str(out_path))
                output_bytes = out_path.read_bytes()

            st.session_state["results"] = results
            st.session_state["output_bytes"] = output_bytes
            st.session_state["skipped"] = skipped

    results = st.session_state["results"]

    if results is None:
        st.info("Sube los dos archivos y presiona **Ejecutar conciliación** para empezar.")
        st.stop()

    skipped = st.session_state["skipped"]
    if skipped:
        st.warning(
            "No se pudo leer el encabezado de estas pestañas del archivo RCF "
            f"(se omitieron): {', '.join(skipped)}"
        )

    # ---------------------------------------------------------------------------
    # Resumen general
    # ---------------------------------------------------------------------------
    rows = []
    for sn, r in results.items():
        rows.append(dict(
            Hoja=sn, Lote=r["lote"], Productor=r["productor"],
            Listados=r["total_pallets_listados"],
            Encontrados=len(r["matched"]),
            **{"Con diferencia": len(r["mismatched"])},
            **{"No encontrados": len(r["not_found"])},
            Sobrantes=len(r["surplus"]),
        ))
    summary_df = pd.DataFrame(rows).sort_values("Hoja")

    st.subheader("Resumen por hoja / productor")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Pallets listados", int(summary_df["Listados"].sum()))
    c2.metric("Encontrados OK", int(summary_df["Encontrados"].sum()))
    c3.metric("Con diferencia", int(summary_df["Con diferencia"].sum()))
    c4.metric("No encontrados", int(summary_df["No encontrados"].sum()))
    c5.metric("Sobrantes en RCF", int(summary_df["Sobrantes"].sum()))

    st.download_button(
        "⬇️ Descargar Excel de LIQUIDACIÓN con la conciliación aplicada",
        data=st.session_state["output_bytes"],
        file_name="LIQUIDACION_CONCILIADO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    st.divider()

    # ---------------------------------------------------------------------------
    # Detalle por hoja
    # ---------------------------------------------------------------------------
    st.subheader("Detalle por hoja")
    sheet_choice = st.selectbox("Elige una hoja / productor", sorted(results.keys()))
    r = results[sheet_choice]

    st.markdown(f"**Lote:** {r['lote']}  |  **Productor:** {r['productor']}")

    cal_rows = []
    for cal, d in sorted(r["by_calibre"].items()):
        cal_rows.append(dict(
            Calibre=cal, **{"Pallets listados": d["listados"]},
            **{"Cajas listadas": d["cajas_listadas"]},
            **{"Pallets encontrados": d["encontrados"]},
            **{"Cajas confirmadas (RCF)": d["cajas_confirmadas"]},
            **{"Con diferencia": d["con_diferencia"]},
            **{"No encontrados": d["no_encontrados"]},
            **{"# pallets faltantes": ", ".join(str(x) for x in d["faltantes_pallets"]) or "—"},
        ))
    st.dataframe(pd.DataFrame(cal_rows), use_container_width=True, hide_index=True)

    tab1, tab2, tab3 = st.tabs(["No encontrados", "Con diferencia", "Sobrantes en RCF"])

    with tab1:
        if r["not_found"]:
            st.dataframe(pd.DataFrame([
                dict(Pallet=p["pallet"], Calibre=p["calibre"], Cajas=p["cajas"], Celda=p["coord"])
                for p in r["not_found"]
            ]), use_container_width=True, hide_index=True)
        else:
            st.success("Todos los pallets listados se encontraron en el RCF.")

    with tab2:
        if r["mismatched"]:
            st.dataframe(pd.DataFrame([
                dict(Pallet=p["pallet"], Calibre_hoja=p["calibre"], Cajas_hoja=p["cajas"],
                     Calibre_RCF=p["rcf_match"]["calibre"], Cajas_RCF=p["rcf_match"]["cajas"],
                     Manifiesto=p["rcf_match"]["manifiesto"], Celda=p["coord"])
                for p in r["mismatched"]
            ]), use_container_width=True, hide_index=True)
        else:
            st.success("Sin diferencias de calibre/cajas entre la hoja y el RCF.")

    with tab3:
        if r["surplus"]:
            st.dataframe(pd.DataFrame([
                dict(Pallet=x["pallet"], Calibre=x["calibre"], Cajas=x["cajas"],
                     Productor=x["productor"], Manifiesto=x["manifiesto"])
                for x in r["surplus"]
            ]), use_container_width=True, hide_index=True)
        else:
            st.success("No hay pallets de este lote en el RCF que falten en la hoja.")
