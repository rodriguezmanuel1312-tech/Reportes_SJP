"""Cuaderno Digital de Transportes - Versión Pro

Incluye:
- PIN de acceso
- Guardado en tiempo real en Google Sheets + respaldo local
- Eliminación de registros mal digitados
- Campo de Cantidad (m3 / vueltas / litros)
- Rentabilidad por camión (Patente)
- Cobros por rango de fechas y exportación a Excel (.xlsx) y CSV
"""

from __future__ import annotations

import csv
import io
import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

import requests
import streamlit as st


# 👇 1. PEGA AQUÍ ENTRE LAS COMILLAS TU URL DE GOOGLE SHEETS (termina en /exec)
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwEpJD-1GBfaLKQ00RJm6CxDCBoAZJ_t-jM1OWXSoHTGTOHtoluhsCgNDT2VBPVpA89/exec"

# 👇 2. CLAVE DE ACCESO A LA APP (puedes cambiar "1234" por el PIN que quieras)
PIN_ACCESO = "1313"

DATA_FILE = Path("cuaderno_registros.json")
PATENTES = ["FDKH99", "DRXX69", "SX3407"]
INGRESOS = [
    "Reparto General",
    "Movimiento de tierra",
    "Bolones",
    "Estuco",
    "Grava",
    "Arena",
]
EGRESOS = ["Petróleo", "Peaje / TAG", "Neumáticos", "Repuestos", "Mantención Taller", "Otro"]
CLIENTES = ["Obra Buin", "Constructora Paine", "Particular", "Otro"]
PROVEEDORES = ["Servicentro", "Autopista Ruta 5 Sur", "Vulcanización", "Taller", "Otro"]


def usa_google_sheets() -> bool:
    return GOOGLE_SHEET_URL.startswith("https://script.google.com/")


@st.cache_data(ttl=15, show_spinner=False)
def cargar_desde_sheets() -> list[dict[str, Any]] | None:
    try:
        resp = requests.get(GOOGLE_SHEET_URL, timeout=10)
        if resp.status_code == 200:
            data = resp.json()
            return data if isinstance(data, list) else None
    except Exception:
        return None
    return None


def cargar_registros() -> list[dict[str, Any]]:
    if usa_google_sheets():
        remotos = cargar_desde_sheets()
        if remotos is not None:
            return remotos

    if not DATA_FILE.exists():
        return []
    try:
        data = json.loads(DATA_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, list) else []
    except (OSError, json.JSONDecodeError):
        return []


def guardar_registros_local(registros: list[dict[str, Any]]) -> None:
    DATA_FILE.write_text(
        json.dumps(registros, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def pesos(valor: int | float) -> str:
    return f"${int(valor):,}".replace(",", ".")


def limpiar(valor: Any) -> str:
    return str(valor or "").replace(",", ";").replace("\n", " ").strip()


def filtrar_por_fechas(
    registros: list[dict[str, Any]], desde: date, hasta: date
) -> list[dict[str, Any]]:
    d_str = desde.isoformat()
    h_str = hasta.isoformat()
    return [r for r in registros if d_str <= str(r.get("fecha", "")) <= h_str]


def exportar_csv(registros: list[dict[str, Any]]) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(
        [
            "Fecha",
            "N° Reporte",
            "Patente",
            "Tipo",
            "Categoría",
            "Cantidad",
            "Cliente/Proveedor",
            "Detalle",
            "Monto",
        ]
    )
    for r in registros:
        writer.writerow(
            [
                r.get("fecha", ""),
                r.get("numero_reporte", ""),
                r.get("patente", ""),
                r.get("tipo", ""),
                r.get("categoria", ""),
                r.get("cantidad", ""),
                r.get("contraparte", ""),
                limpiar(r.get("detalle")),
                r.get("monto", 0),
            ]
        )
    return output.getvalue().encode("utf-8-sig")


def exportar_xlsx(registros: list[dict[str, Any]]) -> bytes | None:
    try:
        import pandas as pd
    except ImportError:
        return None

    rows = [
        {
            "Fecha": r.get("fecha", ""),
            "N° Reporte": r.get("numero_reporte", ""),
            "Patente": r.get("patente", ""),
            "Tipo": r.get("tipo", ""),
            "Categoría": r.get("categoria", ""),
            "Cantidad": r.get("cantidad", ""),
            "Cliente/Proveedor": r.get("contraparte", ""),
            "Detalle": limpiar(r.get("detalle")),
            "Monto": int(r.get("monto", 0)),
        }
        for r in registros
    ]
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Cobros")
    except (ImportError, ModuleNotFoundError):
        return None
    return buffer.getvalue()


def agregar_registro(form_data: dict[str, Any]) -> None:
    form_data["id"] = datetime.now().strftime("%Y%m%d%H%M%S%f")
    form_data["accion"] = "agregar"

    if usa_google_sheets():
        try:
            requests.post(GOOGLE_SHEET_URL, json=form_data, timeout=10)
            cargar_desde_sheets.clear()
        except Exception as e:
            st.warning(f"Guardado local, pero falló la conexión con Google Sheets: {e}")

    registros = cargar_registros()
    if not any(str(r.get("id")) == form_data["id"] for r in registros):
        registros.insert(0, form_data)
    guardar_registros_local(registros)


def eliminar_registro(id_registro: str) -> None:
    if usa_google_sheets():
        try:
            requests.post(
                GOOGLE_SHEET_URL,
                json={"accion": "eliminar", "id": str(id_registro)},
                timeout=10,
            )
            cargar_desde_sheets.clear()
        except Exception as e:
            st.warning(f"No se pudo eliminar en Google Sheets: {e}")

    registros = [r for r in cargar_registros() if str(r.get("id")) != str(id_registro)]
    guardar_registros_local(registros)


def verificar_pin() -> bool:
    if not PIN_ACCESO:
        return True
    if st.session_state.get("autenticado", False):
        return True

    st.subheader("🔒 Acceso al Cuaderno")
    pin_ingresado = st.text_input(
        "Ingresa el PIN de 4 dígitos", type="password", max_chars=8
    )
    if st.button("Entrar", type="primary"):
        if pin_ingresado == PIN_ACCESO:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("PIN incorrecto. Intenta nuevamente.")
    return False


def mostrar_anotar() -> None:
    st.subheader("Anotar viaje o gasto")
    tipo = st.radio(
        "Tipo de registro", ["Viaje / Ingreso", "Gasto / Egreso"], horizontal=True
    )
    es_ingreso = tipo == "Viaje / Ingreso"
    categorias = INGRESOS if es_ingreso else EGRESOS
    contrapartes = CLIENTES if es_ingreso else PROVEEDORES

    col_a, col_b = st.columns(2)
    with col_a:
        fecha = st.date_input("Fecha", value=date.today())
        patente = st.selectbox("Patente del camión", PATENTES)
        numero_reporte = st.number_input("N° de reporte (opcional)", min_value=0, step=1, value=0)
        categoria = st.selectbox("Material / categoría", categorias)

    with col_b:
        contraparte = st.selectbox(
            "Cliente / destino" if es_ingreso else "Proveedor",
            contrapartes,
        )
        contraparte_manual = ""
        if contraparte == "Otro":
            contraparte_manual = st.text_input(
                "Nombre del cliente, destino o proveedor",
                placeholder="Ej: Constructora Santa María",
            )
        cantidad = st.text_input(
            "Cantidad (opcional: m³, vueltas o litros)",
            placeholder="Ej: 12 m3, 2 vueltas, 150 L",
        )
        monto = st.number_input(
            "Valor cobrado ($)" if es_ingreso else "Monto pagado ($)",
            min_value=0,
            step=5000,
        )

    detalle = st.text_input("Detalle u observación (opcional)")

    if st.button("💾 Guardar en el cuaderno", type="primary", use_container_width=True):
        contraparte_final = (
            contraparte_manual.strip() if contraparte == "Otro" else contraparte
        )
        if not contraparte_final:
            st.error("Por favor escribe el nombre del cliente, destino o proveedor.")
            return
        if monto <= 0:
            st.error("Ingresa un monto mayor que cero.")
            return

        with st.spinner("Guardando registro..."):
            agregar_registro(
                {
                    "fecha": fecha.isoformat(),
                    "patente": patente,
                    "tipo": "ingreso" if es_ingreso else "egreso",
                    "categoria": categoria,
                    "cantidad": cantidad.strip(),
                    "contraparte": contraparte_final,
                    "detalle": detalle.strip(),
                    "monto": int(monto),
                    "numero_reporte": int(numero_reporte) if numero_reporte else "",
                }
            )
        st.success("✅ Registro guardado correctamente.")
        st.rerun()

    registros = cargar_registros()
    if registros:
        st.divider()
        st.subheader("Últimos registros")
        for r in registros[:8]:
            signo = "+" if r.get("tipo") == "ingreso" else "-"
            cant_txt = f" ({r.get('cantidad')})" if r.get("cantidad") else ""
            rep_txt = f" · Rep #{r.get('numero_reporte')}" if r.get("numero_reporte") else ""
            col_texto, col_borrar = st.columns([5, 1])
            with col_texto:
                st.write(
                    f"**[{r.get('patente', '')}] {r.get('categoria', '')}{cant_txt}** · "
                    f"{r.get('contraparte', '')}{rep_txt} · {r.get('fecha', '')} · "
                    f"**{signo}{pesos(r.get('monto', 0))}**"
                )
            with col_borrar:
                reg_id = str(r.get("id", ""))
                if reg_id and st.button("🗑️", key=f"del_{reg_id}", help="Eliminar este registro"):
                    with st.spinner("Eliminando..."):
                        eliminar_registro(reg_id)
                    st.rerun()


def mostrar_ganancias() -> None:
    st.subheader("Ganancias y Rendimiento por Camión")
    registros = cargar_registros()

    hoy = date.today()
    inicio_mes = hoy.replace(day=1)

    col_f1, col_f2, col_f3 = st.columns(3)
    with col_f1:
        desde = st.date_input("Desde", value=inicio_mes, key="gan_desde")
    with col_f2:
        hasta = st.date_input("Hasta", value=hoy, key="gan_hasta")
    with col_f3:
        patente_filtro = st.selectbox("Camión (Patente)", ["Todos"] + PATENTES)

    filtrados = filtrar_por_fechas(registros, desde, hasta)
    if patente_filtro != "Todos":
        filtrados = [r for r in filtrados if r.get("patente") == patente_filtro]

    ingresos = sum(int(r.get("monto", 0)) for r in filtrados if r.get("tipo") == "ingreso")
    egresos = sum(int(r.get("monto", 0)) for r in filtrados if r.get("tipo") == "egreso")

    col1, col2, col3 = st.columns(3)
    col1.metric("Ingresos", pesos(ingresos))
    col2.metric("Gastos", pesos(egresos))
    col3.metric("Ganancia neta", pesos(ingresos - egresos))

    st.divider()
    st.subheader("🚚 Rentabilidad por Camión en el período")
    todos_periodo = filtrar_por_fechas(registros, desde, hasta)
    resumen_camiones = []
    for pat in PATENTES:
        reg_pat = [r for r in todos_periodo if r.get("patente") == pat]
        ing_p = sum(int(r.get("monto", 0)) for r in reg_pat if r.get("tipo") == "ingreso")
        egr_p = sum(int(r.get("monto", 0)) for r in reg_pat if r.get("tipo") == "egreso")
        viajes_p = sum(1 for r in reg_pat if r.get("tipo") == "ingreso")
        resumen_camiones.append(
            {
                "Patente": pat,
                "N° Viajes": viajes_p,
                "Ingresos": pesos(ing_p),
                "Gastos": pesos(egr_p),
                "Ganancia Neta": pesos(ing_p - egr_p),
            }
        )
    st.dataframe(resumen_camiones, use_container_width=True, hide_index=True)

    if filtrados:
        st.subheader("Desglose por categoría")
        categorias: dict[str, int] = {}
        for r in filtrados:
            cat = r.get("categoria", "Sin categoría")
            categorias[cat] = categorias.get(cat, 0) + int(r.get("monto", 0))
        st.bar_chart(categorias)
    else:
        st.info("No hay registros en el rango de fechas seleccionado.")


def mostrar_cobros() -> None:
    st.subheader("Estado de Pago / Cobros a Clientes")
    registros = cargar_registros()
    ingresos = [r for r in registros if r.get("tipo") == "ingreso"]

    clientes_registrados = sorted(
        {r.get("contraparte", "") for r in ingresos if r.get("contraparte")}
    )
    lista_clientes = ["Todos los clientes"] + (clientes_registrados or CLIENTES)

    hoy = date.today()
    inicio_mes = hoy.replace(day=1)

    col1, col2, col3 = st.columns(3)
    with col1:
        cliente = st.selectbox("Cliente a cobrar", lista_clientes)
    with col2:
        desde = st.date_input("Desde", value=inicio_mes, key="cob_desde")
    with col3:
        hasta = st.date_input("Hasta", value=hoy, key="cob_hasta")

    viajes = filtrar_por_fechas(ingresos, desde, hasta)
    if cliente != "Todos los clientes":
        viajes = [r for r in viajes if r.get("contraparte") == cliente]

    total = sum(int(r.get("monto", 0)) for r in viajes)

    if viajes:
        st.dataframe(
            [
                {
                    "Fecha": r.get("fecha", ""),
                    "N° Reporte": r.get("numero_reporte", ""),
                    "Patente": r.get("patente", ""),
                    "Cliente": r.get("contraparte", ""),
                    "Categoría": r.get("categoria", ""),
                    "Cantidad": r.get("cantidad", ""),
                    "Detalle": r.get("detalle", ""),
                    "Monto": pesos(r.get("monto", 0)),
                }
                for r in viajes
            ],
            use_container_width=True,
            hide_index=True,
        )
        st.metric(f"Total a cobrar ({len(viajes)} viajes)", pesos(total))

        nombre_archivo = f"cobro_{cliente.replace(' ', '_')}_{desde}_a_{hasta}"
        xlsx_data = exportar_xlsx(viajes)
        if xlsx_data:
            st.download_button(
                "📥 Descargar Excel (.xlsx)",
                data=xlsx_data,
                file_name=f"{nombre_archivo}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                type="primary",
            )
        csv_data = exportar_csv(viajes)
        st.download_button(
            "📄 Descargar CSV",
            data=csv_data,
            file_name=f"{nombre_archivo}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    else:
        st.info("No hay viajes registrados para este cliente y rango de fechas.")


def main() -> None:
    st.set_page_config(
        page_title="Cuaderno Digital de Transportes", page_icon="🚚", layout="wide"
    )
    st.title("🚚 Cuaderno Digital de Transportes")

    if not verificar_pin():
        return

    if usa_google_sheets():
        st.caption("☁ Conectado en tiempo real con Google Sheets")
    else:
        st.caption("💾 Modo local")

    anotar, ganancias, cobros = st.tabs(["📝 Anotar", "📊 Ganancias", "💰 Cobros"])
    with anotar:
        mostrar_anotar()
    with ganancias:
        mostrar_ganancias()
    with cobros:
        mostrar_cobros()


if __name__ == "__main__":
    main()
