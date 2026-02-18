import os
from datetime import date, datetime, timedelta

import pandas as pd
import streamlit as st

from db import (
    init_db,
    add_scheduled_records,
    get_month_records,
    update_records_status_and_notes,
    admin_update_scheduled_date,
    admin_delete_record,
    export_month_matrix_xlsx_bytes,
)

st.set_page_config(page_title="Registro de Actividades", layout="wide")
init_db()

def month_bounds(any_day: date):
    first = any_day.replace(day=1)
    # next month
    if first.month == 12:
        next_first = date(first.year + 1, 1, 1)
    else:
        next_first = date(first.year, first.month + 1, 1)
    last = next_first - timedelta(days=1)
    return first, last

def get_admin_code() -> str | None:
    # Prefer Streamlit secrets, fallback to env var
    try:
        return st.secrets.get("ADMIN_CODE")
    except Exception:
        return None

def verify_admin(code: str) -> bool:
    secret = get_admin_code() or os.getenv("ADMIN_CODE")
    return bool(secret) and code == secret

# --- Sidebar / access ---
st.sidebar.title("Acceso")
role = st.sidebar.radio("Rol", ["Especialista", "Administrador"], index=0)
actor_name = st.sidebar.text_input("Tu nombre", value=st.session_state.get("actor_name", ""), placeholder="Ej. Ana Pérez")

is_admin = False
if role == "Administrador":
    code = st.sidebar.text_input("Código especial (admin)", type="password")
    is_admin = verify_admin(code.strip()) if code else False
    st.sidebar.caption("Consejo: guarda el código en `.streamlit/secrets.toml` o variable de entorno `ADMIN_CODE`.")
    if not is_admin and code:
        st.sidebar.error("Código incorrecto.")
else:
    is_admin = False

if actor_name:
    st.session_state["actor_name"] = actor_name.strip()
else:
    st.sidebar.warning("Escribe tu nombre para registrar/actualizar actividades.")

st.sidebar.divider()
today = date.today()
selected_month = st.sidebar.date_input("Mes de trabajo", value=today, help="Se usa para filtrar el tablero y la matriz mensual.")
month_first, month_last = month_bounds(selected_month)

st.title("📋 Registro de Actividades")
st.caption("Registra actividades por día y marca el resultado: ✓ cumplido / ✗ incumplido.")

tab_reg, tab_estado, tab_tablero, tab_export, tab_admin = st.tabs(
    ["➕ Registrar", "✅/✗ Marcar estado", "📊 Tablero", "⬇️ Exportar", "🛡️ Admin"]
)

# --- Registrar ---
with tab_reg:
    st.subheader("Registrar nueva actividad (programación)")
    with st.form("form_registro", clear_on_submit=True):
        col1, col2, col3 = st.columns([2, 3, 2])
        with col1:
            especialista = st.text_input("Especialista", value=actor_name, placeholder="Nombre completo")
        with col2:
            actividad = st.text_input("Actividad (tarea)", placeholder="Ej. Elaborar informe mensual")
        with col3:
            unidad = st.selectbox(
                "Unidad de medida",
                ["Documento", "Informe", "Reunión", "Visita", "Análisis", "Otro"],
                index=1,
            )
            unidad_otro = st.text_input("Si elegiste 'Otro', especifica", placeholder="Ej. Presentación")
        st.markdown("**Programación**")
        mode = st.radio("Tipo", ["Fecha única", "Rango"], horizontal=True)

        if mode == "Fecha única":
            fecha = st.date_input("Fecha programada", value=today)
            fechas = [fecha]
        else:
            c1, c2, c3 = st.columns([1, 1, 2])
            with c1:
                f_ini = st.date_input("Inicio", value=today)
            with c2:
                f_fin = st.date_input("Fin", value=today)
            with c3:
                dias = st.multiselect(
                    "Días de la semana (opcional)",
                    ["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"],
                    default=["Lun", "Mar", "Mié", "Jue", "Vie"],
                    help="Si lo dejas como está, se programará de lunes a viernes.",
                )
                incluir_finde = st.checkbox("Incluir fines de semana", value=False)

            # Build date list
            fechas = []
            if f_ini <= f_fin:
                cur = f_ini
                map_wd = {0: "Lun", 1: "Mar", 2: "Mié", 3: "Jue", 4: "Vie", 5: "Sáb", 6: "Dom"}
                allowed = set(dias) if dias else set(map_wd.values())
                while cur <= f_fin:
                    wd = map_wd[cur.weekday()]
                    if (incluir_finde or wd not in {"Sáb", "Dom"}) and wd in allowed:
                        fechas.append(cur)
                    cur += timedelta(days=1)

        notas = st.text_area("Notas / detalle (opcional)", placeholder="Criterios, entregables, etc.")
        submitted = st.form_submit_button("Registrar")

    if submitted:
        if not (especialista and actividad and (unidad or unidad_otro) and fechas):
            st.error("Faltan datos: especialista, actividad, unidad y al menos 1 fecha.")
        else:
            unidad_final = unidad_otro.strip() if unidad == "Otro" else unidad
            records = [
                {
                    "specialist": especialista.strip(),
                    "activity": actividad.strip(),
                    "unit": unidad_final,
                    "scheduled_date": f.isoformat(),
                    "status": "",
                    "notes": notas.strip(),
                    "created_by": (actor_name or "—").strip(),
                }
                for f in fechas
            ]
            add_scheduled_records(records)
            st.success(f"Listo: se registraron {len(records)} actividad(es).")

# --- Marcar estado ---
with tab_estado:
    st.subheader("Marcar estado (✓/✗) y notas")
    if not actor_name:
        st.info("Escribe tu nombre en la barra lateral para continuar.")
    else:
        c1, c2, c3 = st.columns([2, 2, 2])
        with c1:
            filtro_especialista = st.text_input("Filtrar por especialista", value=actor_name if role == "Especialista" else "")
        with c2:
            rango_ini = st.date_input("Desde", value=month_first)
        with c3:
            rango_fin = st.date_input("Hasta", value=month_last)

        df = get_month_records(rango_ini, rango_fin, specialist=(filtro_especialista.strip() or None))
        if df.empty:
            st.warning("No hay registros en el rango seleccionado.")
        else:
            df = df.sort_values(["scheduled_date", "specialist", "activity"]).reset_index(drop=True)
            original = df.copy()

            st.caption("Edita solo **Estado** y **Notas**. El plazo (fecha programada) solo lo puede cambiar el administrador.")
            edited = st.data_editor(
                df,
                hide_index=True,
                use_container_width=True,
                column_config={
                    "id": st.column_config.NumberColumn("ID", disabled=True),
                    "scheduled_date": st.column_config.DateColumn("Fecha", disabled=not is_admin),
                    "specialist": st.column_config.TextColumn("Especialista", disabled=True),
                    "activity": st.column_config.TextColumn("Actividad", disabled=True),
                    "unit": st.column_config.TextColumn("UM", disabled=True),
                    "status": st.column_config.SelectboxColumn("Estado", options=["", "✓", "✗"], required=False),
                    "notes": st.column_config.TextColumn("Notas", width="large"),
                    "updated_at": st.column_config.TextColumn("Actualizado", disabled=True),
                },
            )

            col_save, col_hint = st.columns([1, 3])
            with col_save:
                if st.button("Guardar cambios"):
                    # Determine changes
                    changes = []
                    for _, row in edited.iterrows():
                        rid = int(row["id"])
                        old = original.loc[original["id"] == rid].iloc[0]
                        # Enforce: non-admin cannot change scheduled_date
                        if not is_admin and row["scheduled_date"] != old["scheduled_date"]:
                            st.error(f"No autorizado: el registro ID {rid} cambió fecha. Se ignorará ese cambio.")
                            row["scheduled_date"] = old["scheduled_date"]
                        if (row["status"] != old["status"]) or (str(row["notes"]) != str(old["notes"])):
                            changes.append(
                                {
                                    "id": rid,
                                    "status": row["status"],
                                    "notes": row["notes"] if pd.notna(row["notes"]) else "",
                                    "actor": actor_name.strip(),
                                }
                            )

                    if not changes:
                        st.info("No detecté cambios para guardar.")
                    else:
                        update_records_status_and_notes(changes)
                        st.success(f"Guardado: {len(changes)} cambio(s).")
                        st.rerun()
            with col_hint:
                st.caption("Tip: deja **Estado** en blanco si no aplica / aún no se ejecuta.")

# --- Tablero ---
with tab_tablero:
    st.subheader("Tablero mensual (carga y cumplimiento)")
    dfm = get_month_records(month_first, month_last, specialist=None)
    if dfm.empty:
        st.info("Aún no hay registros en el mes seleccionado.")
    else:
        dfm["status"] = dfm["status"].fillna("")
        resumen = (
            dfm.assign(
                planned=1,
                completed=(dfm["status"] == "✓").astype(int),
                missed=(dfm["status"] == "✗").astype(int),
                pending=(dfm["status"] == "").astype(int),
            )
            .groupby("specialist", as_index=False)[["planned", "completed", "missed", "pending"]]
            .sum()
        )
        resumen["tasa_cumplimiento"] = resumen.apply(
            lambda r: (r["completed"] / (r["completed"] + r["missed"])) if (r["completed"] + r["missed"]) > 0 else 0,
            axis=1,
        )
        st.dataframe(resumen.sort_values("planned", ascending=False), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1:
            st.caption("Carga (programadas) por especialista")
            st.bar_chart(resumen.set_index("specialist")["planned"])
        with c2:
            st.caption("Cumplidas vs. Incumplidas (suma mensual)")
            chart_df = resumen.set_index("specialist")[["completed", "missed"]]
            st.bar_chart(chart_df)

# --- Exportar ---
with tab_export:
    st.subheader("Exportar matriz mensual a Excel")
    st.caption("Genera una matriz tipo Excel (días del mes como columnas) a partir de lo registrado en la app.")
    exp_especialista = st.text_input("Filtrar por especialista (opcional)", value="")
    if st.button("Generar Excel"):
        bytes_xlsx = export_month_matrix_xlsx_bytes(month_first, month_last, specialist=(exp_especialista.strip() or None))
        filename = f"matriz_{month_first.strftime('%Y_%m')}.xlsx"
        st.download_button(
            label="Descargar Excel",
            data=bytes_xlsx,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# --- Admin ---
with tab_admin:
    if not is_admin:
        st.info("Sección solo para Administrador. Ingresa el código en la barra lateral.")
    else:
        st.subheader("Administración (editar plazos / borrar registros)")
        st.warning("Los cambios quedan registrados en un log de auditoría.")

        c1, c2, c3 = st.columns([1, 2, 2])
        with c1:
            rid = st.number_input("ID del registro", min_value=1, step=1)
        with c2:
            nueva_fecha = st.date_input("Nueva fecha programada", value=today)
        with c3:
            motivo = st.text_input("Motivo del cambio", placeholder="Ej. Reprogramación por solicitud del jefe")

        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Actualizar fecha (plazo)"):
                ok = admin_update_scheduled_date(
                    record_id=int(rid),
                    new_date=nueva_fecha.isoformat(),
                    actor=actor_name.strip() or "ADMIN",
                    reason=motivo.strip() or "—",
                )
                if ok:
                    st.success("Fecha actualizada.")
                else:
                    st.error("No se pudo actualizar (ID no encontrado).")

        with col_b:
            if st.button("Borrar registro"):
                ok = admin_delete_record(int(rid), actor=actor_name.strip() or "ADMIN", reason=motivo.strip() or "—")
                if ok:
                    st.success("Registro borrado.")
                else:
                    st.error("No se pudo borrar (ID no encontrado).")
