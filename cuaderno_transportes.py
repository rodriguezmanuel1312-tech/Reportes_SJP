"""Cuaderno Digital de Transportes

Versión Streamlit conectada a Google Sheets + respaldo local.

Ejecutar:
    streamlit run cuaderno_transportes.py
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


# 👇 PEGA AQUÍ ENTRE LAS COMILLAS LA URL QUE TE DIO TU GOOGLE SHEETS (termina en /exec)
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwEpJD-1GBfaLKQ00RJm6CxDCBoAZJ_t-jM1OWXSoHTGTOHtoluhsCgNDT2VBPVpA89/exec"

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


def registros_del_mes(registros: list[dict[str, Any]], mes: str) -> list[dict[str, Any]]:
    return [r for r in registros if str(r.get("fecha", "")).startswith(mes)]


def exportar_csv(registros: list[dict[str, Any]]) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(
        ["Fecha", "N° Reporte", "Patente", "Tipo", "Categoría", "Cliente/Proveedor", "Detalle", "Monto"]
    )
    for registro in registros:
        writer.writerow(
            [
                registro.get("fecha", ""),
                registro.get("numero_reporte", ""),
                registro.get("patente", ""),
                registro.get("tipo", ""),
                registro.get("categoria", ""),
                registro.get("contraparte", ""),
                limpiar(registro.get("detalle")),
                registro.get("monto", 0),
            ]
        )
    return output.getvalue().encode("utf-8-sig")


def exportar_xlsx(registros: list[dict[str, Any]]) -> bytes | None:
    """Genera XLSX si pandas/openpyxl están instalados; si no, usa CSV."""
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
            "Cliente/Proveedor": r.get("contraparte", ""),
            "Detalle": limpiar(r.get("detalle")),
            "Monto": r.get("monto", 0),
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

    if usa_google_sheets():
        try:
            requests.post(GOOGLE_SHEET_URL, json=form_data, timeout=10)
            cargar_desde_sheets.clear()
        except Exception as e:
            st.warning(f"Se guardó localmente, pero hubo un problema al conectar con Google Sheets: {e}")

    registros = cargar_registros()
    if not any(r.get("id") == form_data["id"] for r in registros):
        registros.insert(0, form_data)
    guardar_registros_local(registros)


def mostrar_anotar() -> None:
    st.subheader("Anotar viaje o gasto")
    tipo = st.radio("Tipo de registro", ["Viaje / Ingreso", "Gasto / Egreso"], horizontal=True)
    es_ingreso = tipo == "Viaje / Ingreso"
    categorias = INGRESOS if es_ingreso else EGRESOS
    contrapartes = CLIENTES if es_ingreso else PROVEEDORES

    with st.form("registro_form", clear_on_submit=True):
        fecha = st.date_input("Fecha", value=date.today())
        patente = st.selectbox("Patente del camión", PATENTES)
        numero_reporte = st.number_input("N° de reporte", min_value=0, step=1, value=0)
        categoria = st.selectbox("Material / categoría", categorias)
        contraparte = st.selectbox(
            "Cliente / destino" if es_ingreso else "Proveedor",
            contrapartes,
        )
        contraparte_manual = ""
        if contraparte == "Otro":
            contraparte_manual = st.text_input(
                "Detalle del cliente o destino",
                placeholder="Escribe el nombre del cliente o destino",
            )
        detalle = st.text_input("Detalle (opcional)")
        monto = st.number_input(
            "Valor cobrado" if es_ingreso else "Monto pagado",
            min_value=0,
            step=5000,
        )
        guardar = st.form_submit_button("Guardar en el cuaderno", type="primary")

    if guardar:
        contraparte_final = contraparte_manual.strip() if contraparte == "Otro" else contraparte
        if not contraparte_final:
            st.error("Escribe el detalle del cliente o destino.")
            return
        if monto <= 0:
            st.error("Ingresa un monto mayor que cero.")
            return
        with st.spinner("Guardando en Google Sheets..."):
            agregar_registro(
                {
                    "fecha": fecha.isoformat(),
                    "patente": patente,
                    "tipo": "ingreso" if es_ingreso else "egreso",
                    "categoria": categoria,
                    "contraparte": contraparte_final,
                    "detalle": detalle.strip(),
                    "monto": int(monto),
                    "numero_reporte": int(numero_reporte) if numero_reporte else "",
                }
            )
        st.success("Registro guardado en el cuaderno y Google Sheets.")
        st.rerun()

    registros = cargar_registros()
    if registros:
        st.subheader("Últimos registros")
        for registro in registros[:8]:
            signo = "+" if registro.get("tipo") == "ingreso" else "-"
            st.write(
                f"**{registro.get('categoria', '')}** · {registro.get('contraparte', '')} · "
                f"{registro.get('fecha', '')} · {signo}{pesos(registro.get('monto', 0))}"
            )


def mostrar_ganancias() -> None:
    st.subheader("Ganancias")
    registros = cargar_registros()
    mes = st.date_input("Mes", value=date.today(), key="ganancias_mes").strftime("%Y-%m")
    del_mes = registros_del_mes(registros, mes)
    ingresos = sum(r.get("monto", 0) for r in del_mes if r.get("tipo") == "ingreso")
    egresos = sum(r.get("monto", 0) for r in del_mes if r.get("tipo") == "egreso")

    col1, col2, col3 = st.columns(3)
    col1.metric("Ingresos", pesos(ingresos))
    col2.metric("Gastos", pesos(egresos))
    col3.metric("Ganancia neta", pesos(ingresos - egresos))

    if del_mes:
        st.subheader("Desglose por categoría")
        categorias: dict[str, int] = {}
        for registro in del_mes:
            categoria = registro.get("categoria", "Sin categoría")
            categorias[categoria] = categorias.get(categoria, 0) + int(registro.get("monto", 0))
        st.bar_chart(categorias)
    else:
        st.info("No hay registros para el mes seleccionado.")


def mostrar_cobros() -> None:
    st.subheader("Cobros")
    registros = cargar_registros()
    ingresos = [r for r in registros if r.get("tipo") == "ingreso"]
    clientes = sorted({r.get("contraparte", "") for r in ingresos if r.get("contraparte")}) or CLIENTES
    cliente = st.selectbox("Cliente a cobrar", clientes)
    mes = st.date_input("Mes", value=date.today(), key="cobros_mes").strftime("%Y-%m")
    viajes = [r for r in ingresos if r.get("contraparte") == cliente and str(r.get("fecha", "")).startswith(mes)]
    total = sum(int(r.get("monto", 0)) for r in viajes)

    if viajes:
        st.dataframe(
            [
                {
                    "Fecha": r.get("fecha", ""),
                    "N° Reporte": r.get("numero_reporte", ""),
                    "Patente": r.get("patente", ""),
                    "Categoría": r.get("categoria", ""),
                    "Detalle": r.get("detalle", ""),
                    "Monto": r.get("monto", 0),
                }
                for r in viajes
            ],
            use_container_width=True,
            hide_index=True,
        )
        st.metric("Total a cobrar", pesos(total))
        csv_data = exportar_csv(viajes)
        st.download_button(
            "Descargar CSV para Excel / Google Sheets",
            data=csv_data,
            file_name=f"cobro_{cliente.replace(' ', '_')}_{mes}.csv",
            mime="text/csv",
            use_container_width=True,
        )
        xlsx_data = exportar_xlsx(viajes)
        if xlsx_data:
            st.download_button(
                "Descargar Excel (.xlsx)",
                data=xlsx_data,
                file_name=f"cobro_{cliente.replace(' ', '_')}_{mes}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    else:
        st.info("No hay viajes registrados para este cliente y mes.")


def main() -> None:
    st.set_page_config(page_title="Cuaderno Digital de Transportes", page_icon="🚚", layout="wide")
    st.title("Cuaderno Digital de Transportes")
    if usa_google_sheets():
        st.caption("☁️️ Conectado en tiempo real con Google Sheets")
    else:
        st.caption("💾 Modo local (pega tu GOOGLE_SHEET_URL en el código para activar Google Sheets)")
    anotar, ganancias, cobros = st.tabs(["Anotar", "Ganancias", "Cobros"])
    with anotar:
        mostrar_anotar()
    with ganancias:
        mostrar_ganancias()
    with cobros:
        mostrar_cobros()


if __name__ == "__main__":
    main()