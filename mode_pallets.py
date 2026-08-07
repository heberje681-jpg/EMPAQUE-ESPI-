from datetime import date
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import db


def render():
    st.title("📦 Pallets y embarques — ESPI")
    st.caption(
        "Base de datos compartida en tiempo real: todos los que entren aquí ven "
        "los mismos pallets y embarques, se actualiza al instante para todos."
    )

    with st.expander("💾 Respaldo (por si la app se reinicia y se borran los datos)"):
        st.caption(
            "Esta app no tiene una base de datos permanente todavía — si se reinicia "
            "(por inactividad o al subir un cambio de código) puede perder lo capturado. "
            "Descarga un respaldo seguido, y si un día abres la app y la ves vacía, "
            "sube aquí el último respaldo para recuperar todo."
        )
        c1, c2 = st.columns(2)
        with c1:
            with tempfile.TemporaryDirectory() as tmp:
                out = Path(tmp) / "respaldo_espi.xlsx"
                db.exportar_backup(str(out))
                st.download_button("⬇️ Descargar respaldo ahora", data=out.read_bytes(),
                                    file_name="respaldo_espi.xlsx", use_container_width=True)
        with c2:
            resp = st.file_uploader("Subir un respaldo para restaurar", type=["xlsx"], key="resp_upl")
            if resp is not None:
                st.warning("Esto REEMPLAZA todos los datos actuales con los del respaldo.")
                if st.button("Confirmar restauración", use_container_width=True):
                    with tempfile.TemporaryDirectory() as tmp:
                        p = Path(tmp) / "respaldo.xlsx"
                        p.write_bytes(resp.getvalue())
                        db.importar_backup(str(p))
                    st.success("Datos restaurados.")
                    st.rerun()

    sub = st.radio("Sección", ["👤 Productores", "🧊 Pallets", "🚚 Embarques"],
                    horizontal=True, label_visibility="collapsed")

    if sub == "👤 Productores":
        _render_productores()
    elif sub == "🧊 Pallets":
        _render_pallets()
    else:
        _render_embarques()


# ---------------------------------------------------------------------------
# Productores
# ---------------------------------------------------------------------------

def _render_productores():
    st.subheader("Agregar productor")
    with st.form("form_productor_db", clear_on_submit=True):
        c1, c2 = st.columns(2)
        nombre = c1.text_input("Nombre")
        lote = c2.number_input("Número de lote", min_value=0, step=1)
        c3, c4, c5 = st.columns(3)
        telefono = c3.text_input("Teléfono de contacto")
        empresa = c4.text_input("Empresa")
        huerta = c5.text_input("Huerta")
        if st.form_submit_button("Agregar productor") and nombre:
            db.crear_productor(nombre, lote, telefono, empresa, huerta)
            st.success(f"Productor '{nombre}' agregado.")

    productores = db.listar_productores()
    st.subheader(f"Productores registrados ({len(productores)})")
    if productores:
        st.dataframe(pd.DataFrame(productores)[["nombre", "lote", "telefono", "empresa", "huerta"]],
                     use_container_width=True, hide_index=True)
    else:
        st.info("Aún no hay productores registrados.")


def _selector_productor(label="Productor", key=None):
    productores = db.listar_productores()
    if not productores:
        st.warning("Primero registra al menos un productor en la sección 👤 Productores.")
        return None
    opciones = {p["id"]: f'{p["nombre"]} (lote {p["lote"]})' for p in productores}
    return st.selectbox(label, options=list(opciones.keys()), format_func=lambda i: opciones[i], key=key)


# ---------------------------------------------------------------------------
# Pallets
# ---------------------------------------------------------------------------

def _render_pallets():
    st.subheader("Registrar pallet nuevo")
    productor_id = _selector_productor(key="pallet_productor")
    if productor_id is None:
        return

    with st.form("form_pallet", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        numero = c1.number_input("# de pallet", min_value=0, step=1)
        variedad = c2.text_input("Variedad", value="ATAULFO")
        calibre = c3.selectbox("Calibre", options=db.CALIBRES)
        c4, c5 = st.columns(2)
        cajas_sugeridas = db.CALIBRE_CAJAS_SUGERIDAS.get(calibre, 225)
        cajas = c4.number_input("Cajas", min_value=0, step=1, value=cajas_sugeridas,
                                 help="Se sugiere según el calibre; ajústalo si es distinto (ej. doble línea).")
        organico = c5.checkbox("Orgánico")
        mixto_notas = st.text_input("Notas (ej. mixteado con otra variedad/productor)")
        if st.form_submit_button("Registrar pallet"):
            db.crear_pallet(numero, productor_id, variedad, calibre, cajas, organico, mixto_notas)
            st.success(f"Pallet #{numero} registrado.")
            st.rerun()

    st.divider()
    st.subheader("Pallets activos")
    c1, c2, c3 = st.columns(3)
    filtro_estado = c1.selectbox("Filtrar por estado", ["(todos)"] + db.ESTADOS_PALLET)
    filtro_prod = c2.selectbox("Filtrar por productor", ["(todos)"] +
                                [p["nombre"] for p in db.listar_productores()])
    solo_disponibles = c3.checkbox("Solo sin asignar a embarque")

    pallets = db.listar_pallets()
    if filtro_estado != "(todos)":
        pallets = [p for p in pallets if p["estado"] == filtro_estado]
    if filtro_prod != "(todos)":
        pallets = [p for p in pallets if p["productor_nombre"] == filtro_prod]
    if solo_disponibles:
        pallets = [p for p in pallets if p["embarque_id"] is None]

    if pallets:
        df = pd.DataFrame(pallets)[["numero", "productor_nombre", "variedad", "calibre", "cajas",
                                     "organico", "estado", "embarque_id", "mixto_notas"]]
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No hay pallets con ese filtro.")

    st.divider()
    st.subheader("Armar / deshacer pallets")
    t1, t2 = st.tabs(["🔀 Deshacer un pallet", "🔗 Armar (fusionar) pallets"])

    with t1:
        activos = db.listar_pallets()
        if activos:
            opciones = {p["id"]: f'#{p["numero"]} — {p["productor_nombre"]} — {p["cajas"]} cajas'
                        for p in activos}
            pid = st.selectbox("Pallet a deshacer", options=list(opciones.keys()),
                                format_func=lambda i: opciones[i], key="deshacer_sel")
            n_partes = st.number_input("¿En cuántas partes?", min_value=2, max_value=5, value=2, step=1)
            partes = []
            cols = st.columns(n_partes)
            for i in range(n_partes):
                with cols[i]:
                    st.markdown(f"**Parte {i+1}**")
                    num = st.number_input(f"# pallet {i+1}", min_value=0, step=1, key=f"dp_num_{i}")
                    caj = st.number_input(f"Cajas {i+1}", min_value=0, step=1, key=f"dp_caj_{i}")
                    partes.append(dict(numero=num, cajas=caj))
            if st.button("Confirmar deshacer"):
                nuevos = db.deshacer_pallet(pid, partes)
                st.success(f"Pallet dividido en {len(nuevos)} nuevos pallets.")
                st.rerun()
        else:
            st.info("No hay pallets activos para deshacer.")

    with t2:
        activos = db.listar_pallets()
        if activos:
            opciones = {p["id"]: f'#{p["numero"]} — {p["productor_nombre"]} — {p["cajas"]} cajas'
                        for p in activos}
            seleccion = st.multiselect("Pallets a fusionar (2 o más)", options=list(opciones.keys()),
                                        format_func=lambda i: opciones[i])
            c1, c2, c3 = st.columns(3)
            numero_resultante = c1.number_input("# del pallet resultante", min_value=0, step=1)
            productor_resultante = c2.selectbox(
                "Productor del pallet resultante",
                options=[p["id"] for p in db.listar_productores()],
                format_func=lambda i: db.productor(i)["nombre"])
            variedad_resultante = c3.text_input("Variedad/calibre resultante", value="MIXTO")
            if st.button("Confirmar armado") and len(seleccion) >= 2:
                nid = db.armar_pallet(seleccion, numero_resultante, productor_resultante,
                                       variedad_resultante, variedad_resultante)
                st.success(f"Pallets fusionados en el nuevo pallet #{numero_resultante}.")
                st.rerun()
        else:
            st.info("No hay pallets activos para armar.")


# ---------------------------------------------------------------------------
# Embarques
# ---------------------------------------------------------------------------

def _render_embarques():
    st.subheader("Nuevo embarque / viaje")
    with st.form("form_embarque", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        fecha = c1.date_input("Fecha", value=date.today())
        chofer = c2.text_input("Nombre del chofer")
        telefono_chofer = c3.text_input("Teléfono del chofer")
        c4, c5, c6 = st.columns(3)
        placas = c4.text_input("Placas")
        destino = c5.text_input("Destino (ej. Nogales, Tamaulipas)")
        contacto_llegada = c6.text_input("Contacto en destino")
        c7, c8 = st.columns(2)
        telefono_cliente = c7.text_input(
            "WhatsApp del cliente/destino (con código de país, ej. 5216681234567)",
            help="Para poder mandarle el aviso de embarque por WhatsApp. Puedes dejarlo vacío y agregarlo después.")
        notas = c8.text_input("Notas")
        if st.form_submit_button("Crear embarque"):
            eid = db.crear_embarque(str(fecha), chofer, telefono_chofer, placas, destino,
                                     contacto_llegada, telefono_cliente, notas)
            st.session_state["embarque_activo"] = eid
            st.success("Embarque creado.")
            st.rerun()

    embarques = db.listar_embarques()
    if not embarques:
        st.info("Aún no hay embarques.")
        return

    st.divider()
    opciones = {e["id"]: f'#{e["id"]} — {e["fecha"]} — {e["destino"] or "sin destino"} — {e["estado"]}'
                for e in embarques}
    default_id = st.session_state.get("embarque_activo", embarques[0]["id"])
    if default_id not in opciones:
        default_id = embarques[0]["id"]
    eid = st.selectbox("Selecciona un embarque", options=list(opciones.keys()),
                        index=list(opciones.keys()).index(default_id), format_func=lambda i: opciones[i])
    e = db.embarque(eid)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Estado", e["estado"])
    c2.metric("Chofer", e["chofer"] or "—")
    c3.metric("Placas", e["placas"] or "—")
    c4.metric("Destino", e["destino"] or "—")

    nuevo_estado = st.selectbox("Cambiar estado del embarque", options=db.ESTADOS_EMBARQUE,
                                 index=db.ESTADOS_EMBARQUE.index(e["estado"]))
    if st.button("Actualizar estado"):
        db.actualizar_embarque(eid, estado=nuevo_estado)
        # si el embarque pasa a "EN VIAJE" o "ENTREGADO", reflejarlo en sus pallets también
        if nuevo_estado in ("EN VIAJE", "ENTREGADO"):
            for p in db.listar_pallets(embarque_id=eid):
                db.actualizar_pallet(p["id"], estado=nuevo_estado)
        st.success("Estado actualizado.")
        st.rerun()

    st.divider()
    st.subheader("📲 Aviso de embarque por WhatsApp")
    st.caption(
        "Genera el mensaje y el link de WhatsApp ya listos — solo se abre WhatsApp con el "
        "mensaje escrito, tú le das 'Enviar'. No se manda nada solo ni necesita ninguna cuenta especial."
    )
    telefono_actual = e.get("telefono_cliente") or ""
    nuevo_telefono = st.text_input("WhatsApp del cliente/destino", value=telefono_actual, key="tel_cliente_edit")
    if nuevo_telefono != telefono_actual and st.button("Guardar teléfono"):
        db.actualizar_embarque(eid, telefono_cliente=nuevo_telefono)
        st.rerun()

    link, mensaje = db.link_aviso_whatsapp(eid)
    st.text_area("Mensaje que se va a enviar (puedes editarlo antes de mandarlo)", value=mensaje, height=140,
                 key="mensaje_preview", disabled=True)
    if not telefono_actual.strip():
        st.info("Agrega el WhatsApp del cliente arriba para que el link abra directo la conversación "
                 "(si lo dejas vacío, igual puedes abrir el link y elegir el contacto a mano).")
    c1, c2 = st.columns([1, 3])
    with c1:
        st.link_button("Abrir en WhatsApp", link, use_container_width=True)
    with c2:
        if e.get("aviso_enviado"):
            st.success("Ya se marcó como enviado para este embarque.")
        elif st.button("Marcar aviso como enviado"):
            db.actualizar_embarque(eid, aviso_enviado=1)
            st.rerun()

    st.divider()
    st.subheader("Pallets en este embarque")
    pallets_embarque = db.listar_pallets(embarque_id=eid)
    if pallets_embarque:
        st.dataframe(pd.DataFrame(pallets_embarque)[
            ["numero", "productor_nombre", "variedad", "calibre", "cajas", "estado"]],
            use_container_width=True, hide_index=True)

        st.markdown("**Resumen por productor (pallets y cajas que le tocan de este viaje)**")
        resumen = db.resumen_embarque_por_productor(eid)
        st.dataframe(pd.DataFrame([
            dict(Productor=k, Pallets=v["pallets"], Cajas=v["cajas"]) for k, v in resumen.items()
        ]), use_container_width=True, hide_index=True)

        quitar = st.multiselect("Quitar pallets de este embarque",
                                 options=[p["id"] for p in pallets_embarque],
                                 format_func=lambda i: f'#{next(p["numero"] for p in pallets_embarque if p["id"]==i)}')
        if st.button("Quitar seleccionados") and quitar:
            for pid in quitar:
                db.quitar_pallet_de_embarque(pid)
            st.rerun()
    else:
        st.info("Este embarque todavía no tiene pallets asignados.")

    st.divider()
    st.subheader("Agregar pallets disponibles a este embarque")
    disponibles = [p for p in db.listar_pallets() if p["embarque_id"] is None]
    if disponibles:
        seleccion = st.multiselect(
            "Pallets en cuarto frío disponibles",
            options=[p["id"] for p in disponibles],
            format_func=lambda i: next(
                f'#{p["numero"]} — {p["productor_nombre"]} — {p["cajas"]} cajas' for p in disponibles if p["id"] == i))
        if st.button("Agregar al embarque") and seleccion:
            for pid in seleccion:
                db.asignar_pallet_a_embarque(pid, eid)
            st.success(f"{len(seleccion)} pallets agregados al embarque.")
            st.rerun()
    else:
        st.info("No hay pallets disponibles sin asignar en este momento.")
