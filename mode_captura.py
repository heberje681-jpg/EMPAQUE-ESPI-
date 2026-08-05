import tempfile
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

import capture_core as cc


def render():
    st.title("📝 Captura de monitoreos — ESPI")
    st.caption(
        "Captura aquí cada productor, sus monitoreos (cortes) y la lista de pallets de "
        "cada uno. La app va contando pallets por calibre y, si subes el RCF, te dice "
        "en qué monitoreo está cada pallet y cuáles faltan — sin tocar Excel."
    )

    if "estado" not in st.session_state:
        st.session_state.estado = cc.nuevo_estado()
    if "rcf" not in st.session_state:
        st.session_state.rcf = None  # (idx, by_lote, skipped)

    estado = st.session_state.estado

    # ---------------------------------------------------------------------------
    # Barra lateral: guardar / continuar, y RCF
    # ---------------------------------------------------------------------------
    with st.sidebar:
        st.header("Guardar / continuar")
        if estado["productores"]:
            with tempfile.TemporaryDirectory() as tmp:
                out = Path(tmp) / "captura.xlsx"
                cc.exportar_captura(estado, str(out))
                st.download_button("💾 Descargar captura (para seguir después)",
                                    data=out.read_bytes(), file_name="captura_espi.xlsx",
                                    use_container_width=True)
        resumir = st.file_uploader("📂 Continuar una captura guardada", type=["xlsx"], key="resumir")
        if resumir is not None and st.button("Cargar captura", use_container_width=True):
            with tempfile.TemporaryDirectory() as tmp:
                p = Path(tmp) / "captura.xlsx"
                p.write_bytes(resumir.getvalue())
                st.session_state.estado = cc.importar_captura(str(p))
            st.rerun()

        st.divider()
        st.header("RCF (manifiestos)")
        rcf_file = st.file_uploader("Subir archivo RCF para cruzar", type=["xlsx"], key="rcf_upl")
        if rcf_file is not None and st.button("Cargar RCF", use_container_width=True):
            with tempfile.TemporaryDirectory() as tmp:
                p = Path(tmp) / "rcf.xlsx"
                p.write_bytes(rcf_file.getvalue())
                idx, by_lote, skipped = cc.cargar_rcf(str(p))
                st.session_state.rcf = (idx, by_lote, skipped)
            st.success("RCF cargado.")
        if st.session_state.rcf and st.session_state.rcf[2]:
            st.warning(f"Pestañas RCF omitidas: {', '.join(st.session_state.rcf[2])}")

    tab_productores, tab_captura, tab_resumen = st.tabs(
        ["👤 Productores", "📥 Capturar monitoreo", "📊 Resumen y conciliación"])

    # ---------------------------------------------------------------------------
    # Tab 1: productores
    # ---------------------------------------------------------------------------
    with tab_productores:
        st.subheader("Agregar productor")
        with st.form("form_productor", clear_on_submit=True):
            c1, c2 = st.columns(2)
            nombre = c1.text_input("Nombre del productor")
            lote = c2.number_input("Número de lote", min_value=0, step=1)
            if st.form_submit_button("Agregar productor") and nombre:
                cc.agregar_productor(estado, nombre, lote)
                st.success(f"Productor '{nombre}' (lote {lote}) agregado.")

        if estado["productores"]:
            st.subheader("Productores capturados")
            st.dataframe(pd.DataFrame(estado["productores"])[["nombre", "lote"]],
                         use_container_width=True, hide_index=True)
        else:
            st.info("Aún no hay productores. Agrega el primero arriba.")

    # ---------------------------------------------------------------------------
    # Tab 2: capturar monitoreo + pallets
    # ---------------------------------------------------------------------------
    with tab_captura:
        if not estado["productores"]:
            st.info("Primero agrega al menos un productor en la pestaña anterior.")
        else:
            nombres = {p["id"]: f'{p["nombre"]} (lote {p["lote"]})' for p in estado["productores"]}
            productor_id = st.selectbox("Productor", options=list(nombres.keys()),
                                         format_func=lambda i: nombres[i])

            st.subheader("Nuevo monitoreo (corte)")
            with st.form("form_monitoreo", clear_on_submit=True):
                c1, c2, c3 = st.columns(3)
                fecha = c1.date_input("Fecha", value=date.today())
                folio = c2.text_input("Folio")
                kilos_recibidos = c3.number_input("Kilos recibidos", min_value=0.0, step=1.0)
                c4, c5, c6 = st.columns(3)
                kilos_empacados = c4.number_input("Kilos empacados", min_value=0.0, step=1.0)
                kilos_merma = c5.number_input("Kilos / cajas de merma", min_value=0.0, step=1.0)
                kilos_jugo = c6.number_input("Kilos para jugo", min_value=0.0, step=1.0)
                crear = st.form_submit_button("Crear monitoreo")
                if crear:
                    mid = cc.agregar_monitoreo(estado, productor_id, str(fecha), folio,
                                                kilos_recibidos, kilos_empacados, kilos_merma, kilos_jugo)
                    st.session_state["monitoreo_activo"] = mid
                    st.success("Monitoreo creado. Agrega su lista de pallets abajo.")

            monitoreos = cc.monitoreos_de(estado, productor_id)
            if monitoreos:
                opciones = {m["id"]: f'{m["fecha"]} — folio {m["folio"] or "s/f"}' for m in monitoreos}
                default_id = st.session_state.get("monitoreo_activo", monitoreos[-1]["id"])
                if default_id not in opciones:
                    default_id = monitoreos[-1]["id"]
                monitoreo_id = st.selectbox("Monitoreo activo para capturar pallets",
                                             options=list(opciones.keys()),
                                             index=list(opciones.keys()).index(default_id),
                                             format_func=lambda i: opciones[i])

                st.subheader("Pallets de este monitoreo")
                st.caption("Escribe o pega (Ctrl+V) filas de # pallet, calibre y cajas. "
                           "Deja 'cajas' vacío si esta hoja no lo registra por pallet.")
                df_vacio = pd.DataFrame([{"pallet": None, "calibre": None, "cajas": None} for _ in range(8)])
                df_editado = st.data_editor(df_vacio, num_rows="dynamic", use_container_width=True,
                                             key=f"editor_{monitoreo_id}")
                if st.button("Agregar estas filas al monitoreo"):
                    filas = df_editado.dropna(subset=["pallet"]).to_dict("records")
                    agregados = cc.agregar_pallets(estado, monitoreo_id, filas)
                    st.success(f"{len(agregados)} pallets agregados.")
                    st.rerun()

                ya_capturados = cc.pallets_de_monitoreo(estado, monitoreo_id)
                if ya_capturados:
                    st.markdown(f"**Pallets ya capturados en este monitoreo: {len(ya_capturados)}**")
                    st.dataframe(pd.DataFrame(ya_capturados)[["pallet", "calibre", "cajas"]],
                                 use_container_width=True, hide_index=True)

    # ---------------------------------------------------------------------------
    # Tab 3: resumen y conciliación
    # ---------------------------------------------------------------------------
    with tab_resumen:
        if not estado["productores"]:
            st.info("Aún no hay datos capturados.")
        else:
            rows = []
            for p in estado["productores"]:
                pallets = cc.pallets_de_productor(estado, p["id"])
                n_mon = len(cc.monitoreos_de(estado, p["id"]))
                rows.append(dict(Productor=p["nombre"], Lote=p["lote"], Monitoreos=n_mon,
                                  **{"Pallets capturados": len(pallets)}))
            st.subheader("Captura hasta ahora")
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            if st.session_state.rcf is None:
                st.warning("Sube el archivo RCF en la barra lateral para ver la conciliación "
                           "(pallets encontrados, con diferencia, faltantes y sobrantes).")
            else:
                idx, by_lote, _ = st.session_state.rcf
                resultados = cc.reconciliar_todos(estado, idx, by_lote)

                resumen_rows = []
                for pid, r in resultados.items():
                    resumen_rows.append(dict(
                        Productor=r["productor"], Lote=r["lote"], Listados=r["total_listados"],
                        Encontrados=len(r["matched"]), **{"Con diferencia": len(r["mismatched"])},
                        **{"No encontrados": len(r["not_found"])}, Sobrantes=len(r["surplus"])))
                st.subheader("Conciliación contra RCF")
                st.dataframe(pd.DataFrame(resumen_rows), use_container_width=True, hide_index=True)

                with tempfile.TemporaryDirectory() as tmp:
                    out = Path(tmp) / "reporte.xlsx"
                    cc.exportar_reporte_conciliacion(estado, resultados, str(out))
                    st.download_button("⬇️ Descargar reporte de conciliación (Excel)",
                                        data=out.read_bytes(), file_name="conciliacion_espi.xlsx",
                                        type="primary")

                st.divider()
                nombres = {p["id"]: p["nombre"] for p in estado["productores"]}
                productor_sel = st.selectbox("Ver detalle de un productor", options=list(nombres.keys()),
                                              format_func=lambda i: nombres[i], key="sel_detalle")
                r = resultados[productor_sel]

                cal_rows = []
                for cal, d in sorted(r["by_calibre"].items()):
                    cal_rows.append(dict(
                        Calibre=cal, **{"Pallets listados": d["listados"]},
                        **{"Pallets encontrados": d["encontrados"]},
                        **{"Con diferencia": d["con_diferencia"]},
                        **{"No encontrados": d["no_encontrados"]},
                        **{"# pallets faltantes": ", ".join(str(x) for x in d["faltantes_pallets"]) or "—"}))
                st.dataframe(pd.DataFrame(cal_rows), use_container_width=True, hide_index=True)

                t1, t2, t3 = st.tabs(["No encontrados", "Con diferencia", "Sobrantes en RCF"])
                with t1:
                    if r["not_found"]:
                        st.dataframe(pd.DataFrame([
                            dict(Pallet=p["pallet"], Calibre=p["calibre"], Cajas=p["cajas"],
                                 Monitoreo=p["monitoreo"]["fecha"] if p.get("monitoreo") else "")
                            for p in r["not_found"]]), use_container_width=True, hide_index=True)
                    else:
                        st.success("Todos los pallets capturados se encontraron en el RCF.")
                with t2:
                    if r["mismatched"]:
                        st.dataframe(pd.DataFrame([
                            dict(Pallet=p["pallet"], Calibre_capturado=p["calibre"], Cajas_capturadas=p["cajas"],
                                 Calibre_RCF=p["rcf_match"]["calibre"], Cajas_RCF=p["rcf_match"]["cajas"],
                                 Manifiesto=p["rcf_match"]["manifiesto"],
                                 Monitoreo=p["monitoreo"]["fecha"] if p.get("monitoreo") else "")
                            for p in r["mismatched"]]), use_container_width=True, hide_index=True)
                    else:
                        st.success("Sin diferencias entre lo capturado y el RCF.")
                with t3:
                    if r["surplus"]:
                        st.dataframe(pd.DataFrame([
                            dict(Pallet=x["pallet"], Calibre=x["calibre"], Cajas=x["cajas"],
                                 Productor=x["productor"], Manifiesto=x["manifiesto"])
                            for x in r["surplus"]]), use_container_width=True, hide_index=True)
                    else:
                        st.success("No hay pallets de este productor en el RCF sin capturar.")
