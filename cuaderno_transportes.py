"""Cuaderno Digital S.J.P Transporte - Versión Pro (Diseño + Lógica Gerencial)"""

from __future__ import annotations

import csv
import io
import json
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from fpdf import FPDF
from streamlit_option_menu import option_menu

# 👇 1. PEGA AQUÍ TU URL DE GOOGLE SHEETS
GOOGLE_SHEET_URL = "https://script.google.com/macros/s/AKfycbwEpJD-1GBfaLKQ00RJm6CxDCBoAZJ_t-jM1OWXSoHTGTOHtoluhsCgNDT2VBPVpA89/exec"

# 👇 2. CLAVE DE ACCESO
PIN_ACCESO = "1313"

DATA_FILE = Path("cuaderno_registros.json")
PATENTES = ["FDKH99", "DRXX69", "SX3407"]
INGRESOS = ["Reparto General", "Movimiento de tierra", "Bolones", "Estuco", "Grava", "Arena"]
EGRESOS = ["Petróleo", "Peaje / TAG", "Neumáticos", "Repuestos", "Mantención Taller", "Otro"]
CLIENTES = ["Obra Buin", "Constructora Paine", "Particular", "Otro"]
PROVEEDORES = ["Servicentro", "Autopista Ruta 5 Sur", "Vulcanización", "Taller", "Otro"]

# --- FUNCIONES DE BASE DE DATOS Y UTILIDADES ---

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
    DATA_FILE.write_text(json.dumps(registros, ensure_ascii=False, indent=2), encoding="utf-8")

def pesos(valor: int | float) -> str:
    return f"${int(valor):,}".replace(",", ".")

def limpiar(valor: Any) -> str:
    return str(valor or "").replace(",", ";").replace("\n", " ").strip()

def filtrar_por_fechas(registros: list[dict[str, Any]], desde: date, hasta: date) -> list[dict[str, Any]]:
    d_str = desde.isoformat()
    h_str = hasta.isoformat()
    return [r for r in registros if d_str <= str(r.get("fecha", "")) <= h_str]

# --- GENERADORES DE ARCHIVOS (EXCEL Y PDF) ---

def exportar_xlsx(registros: list[dict[str, Any]]) -> bytes | None:
    rows = [{"Fecha": r.get("fecha", ""), "Camión": r.get("patente", ""), "Categoría": r.get("categoria", ""), "Cant": r.get("cantidad", ""), "Monto": int(r.get("monto", 0))} for r in registros]
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            pd.DataFrame(rows).to_excel(writer, index=False, sheet_name="Cobros")
        return buffer.getvalue()
    except Exception:
        return None

def generar_pdf_cobro(viajes: list[dict[str, Any]], cliente: str, total: int, desde: date, hasta: date) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    
    # 📸 Intentar insertar el logo en la esquina superior izquierda
    try:
        # x=10 (margen izquierdo), y=8 (margen superior), w=30 (ancho de la imagen en mm)
        pdf.image("logo.png", x=10, y=8, w=30)
    except Exception:
        pass # Si no encuentra el logo por algún motivo, crea el PDF sin caerse
        
    # Textos de encabezado
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "S.J.P TRANSPORTE", ln=True, align="C")
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, f"Estado de Pago: {cliente}", ln=True, align="C")
    pdf.set_font("Arial", "", 10)
    pdf.cell(0, 10, f"Período: {desde.strftime('%d/%m/%Y')} al {hasta.strftime('%d/%m/%Y')}", ln=True, align="C")
    pdf.ln(10) # Espacio extra antes de empezar la tabla
    
    # Encabezados de tabla
    pdf.set_font("Arial", "B", 10)
    pdf.cell(30, 8, "Fecha", border=1)
    pdf.cell(30, 8, "Patente", border=1)
    pdf.cell(60, 8, "Detalle", border=1)
    pdf.cell(30, 8, "Cantidad", border=1)
    pdf.cell(40, 8, "Monto", border=1, ln=True)
    
    # Filas
    pdf.set_font("Arial", "", 9)
    for r in viajes:
        pdf.cell(30, 8, str(r.get("fecha", "")), border=1)
        pdf.cell(30, 8, str(r.get("patente", "")), border=1)
        pdf.cell(60, 8, str(r.get("categoria", "")), border=1)
        pdf.cell(30, 8, str(r.get("cantidad", "")), border=1)
        pdf.cell(40, 8, pesos(r.get("monto", 0)), border=1, ln=True)
        
    # Total
    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(150, 10, "TOTAL A COBRAR:", align="R")
    pdf.cell(40, 10, pesos(total), border=1, align="C", ln=True)
    
    return pdf.output(dest="S").encode("latin-1")

def agregar_registro(form_data: dict[str, Any]) -> None:
    form_data["id"] = datetime.now().strftime("%Y%m%d%H%M%S%f")
    form_data["accion"] = "agregar"
    if usa_google_sheets():
        try:
            requests.post(GOOGLE_SHEET_URL, json=form_data, timeout=10)
            cargar_desde_sheets.clear()
        except Exception:
            pass
    registros = cargar_registros()
    registros.insert(0, form_data)
    guardar_registros_local(registros)

def eliminar_registro(id_registro: str) -> None:
    if usa_google_sheets():
        try:
            requests.post(GOOGLE_SHEET_URL, json={"accion": "eliminar", "id": str(id_registro)}, timeout=10)
            cargar_desde_sheets.clear()
        except Exception:
            pass
    registros = [r for r in cargar_registros() if str(r.get("id")) != str(id_registro)]
    guardar_registros_local(registros)

# --- VISTAS DE LA APP ---

def mostrar_anotar() -> None:
    tipo = st.radio("Movimiento", ["Ingreso", "Egreso"], horizontal=True, label_visibility="collapsed")
    es_ingreso = tipo == "Ingreso"
    
    with st.container(border=True):
        col_a, col_b = st.columns(2)
        with col_a:
            fecha = st.date_input("Fecha", value=date.today())
            patente = st.selectbox("Patente", PATENTES)
            categoria = st.selectbox("Categoría", INGRESOS if es_ingreso else EGRESOS)
        with col_b:
            contraparte = st.selectbox("Cliente/Proveedor", CLIENTES if es_ingreso else PROVEEDORES)
            contraparte_manual = st.text_input("Nombre (si es 'Otro')") if contraparte == "Otro" else ""
            cantidad = st.text_input("Cantidad (m³, L, Vueltas)")
        
        monto = st.number_input("Monto total ($)", min_value=0, step=5000)
        detalle = st.text_input("Observación")
        
        foto = None
        if not es_ingreso:
            with st.expander("📸 Adjuntar Foto de Boleta / Recibo (Opcional)"):
                foto = st.camera_input("Tomar foto al comprobante")

        if st.button("💾 Guardar Movimiento", type="primary", use_container_width=True):
            contraparte_final = contraparte_manual.strip() if contraparte == "Otro" else contraparte
            if not contraparte_final or monto <= 0:
                st.toast("⚠️ Revisa el cliente y el monto.", icon="⚠️")
                return
            
            with st.spinner("Guardando..."):
                agregar_registro({
                    "fecha": fecha.isoformat(),
                    "patente": patente,
                    "tipo": "ingreso" if es_ingreso else "egreso",
                    "categoria": categoria,
                    "cantidad": cantidad.strip(),
                    "contraparte": contraparte_final,
                    "detalle": detalle.strip(),
                    "monto": int(monto)
                })
            st.toast("Movimiento registrado con éxito", icon="✅")
            if foto:
                st.toast("📷 Comprobante adjuntado al registro local.", icon="📎")
            st.rerun()

    registros = cargar_registros()
    if registros:
        st.write("### Últimos Movimientos")
        for r in registros[:6]:
            es_ing = r.get("tipo") == "ingreso"
            color = "#10B981" if es_ing else "#EF4444"
            with st.container(border=True):
                col1, col2, col3 = st.columns([1, 6, 2])
                with col1: st.write("🟢" if es_ing else "🔴")
                with col2:
                    st.write(f"**{r.get('categoria', '')}** | {r.get('patente', '')}")
                    st.caption(f"{r.get('contraparte', '')} • {r.get('fecha', '')}")
                with col3:
                    st.markdown(f"<h4 style='color: {color}; text-align: right; margin-bottom:0;'>{pesos(r.get('monto', 0))}</h4>", unsafe_allow_html=True)
                    reg_id = str(r.get("id", ""))
                    if st.button("🗑️", key=f"del_{reg_id}", use_container_width=True):
                        eliminar_registro(reg_id)
                        st.rerun()

def mostrar_flota() -> None:
    st.write("### 🛠️ Control de Flota y Mantenciones")
    st.caption("Monitoreo automático basado en los registros históricos de 'Mantención Taller'.")
    registros = cargar_registros()
    
    for pat in PATENTES:
        reg_camion = [r for r in registros if r.get("patente") == pat]
        mantenciones = [r for r in reg_camion if r.get("categoria") == "Mantención Taller"]
        viajes_realizados = len([r for r in reg_camion if r.get("tipo") == "ingreso"])
        
        with st.container(border=True):
            col1, col2 = st.columns([1, 2])
            with col1:
                st.subheader(f"🚛 {pat}")
                st.caption(f"{viajes_realizados} viajes históricos")
            with col2:
                if mantenciones:
                    ultima = mantenciones[0] 
                    dias = (date.today() - date.fromisoformat(ultima["fecha"])).days
                    if dias > 30:
                        st.error(f"⚠️ Atención: Última mantención hace {dias} días.")
                    else:
                        st.success(f"✅ Al día. Última mantención hace {dias} días.")
                    st.write(f"**Último taller:** {ultima.get('contraparte')} ({pesos(ultima.get('monto', 0))})")
                else:
                    st.warning("⚠️ No hay registros de mantención en sistema.")

def mostrar_finanzas() -> None:
    st.write("### 📊 Tablero Financiero")
    registros = cargar_registros()
    
    col_f1, col_f2 = st.columns(2)
    with col_f1: desde = st.date_input("Desde", value=date.today().replace(day=1))
    with col_f2: hasta = st.date_input("Hasta", value=date.today())

    filtrados = filtrar_por_fechas(registros, desde, hasta)
    ingresos = sum(int(r.get("monto", 0)) for r in filtrados if r.get("tipo") == "ingreso")
    egresos = sum(int(r.get("monto", 0)) for r in filtrados if r.get("tipo") == "egreso")

    col1, col2, col3 = st.columns(3)
    col1.metric("Ingresos", pesos(ingresos))
    col2.metric("Gastos", pesos(egresos))
    col3.metric("Utilidad", pesos(ingresos - egresos), delta=pesos(ingresos - egresos))

    st.write("---")
    st.write("#### 💸 Distribución de Gastos")
    gastos = [r for r in filtrados if r.get("tipo") == "egreso"]
    if gastos:
        fig = px.pie(pd.DataFrame(gastos), values='monto', names='categoria', hole=0.6, color_discrete_sequence=px.colors.sequential.RdBu)
        fig.update_layout(margin=dict(t=0, b=0, l=0, r=0), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True)

def mostrar_cobros() -> None:
    st.write("### 💼 Estados de Pago")
    registros = cargar_registros()
    ingresos = [r for r in registros if r.get("tipo") == "ingreso"]
    
    clientes_registrados = sorted({r.get("contraparte", "") for r in ingresos if r.get("contraparte")})
    cliente = st.selectbox("Seleccionar Cliente", ["Todos"] + (clientes_registrados or CLIENTES))
    
    col1, col2 = st.columns(2)
    with col1: desde = st.date_input("Desde", value=date.today().replace(day=1), key="cob_desde")
    with col2: hasta = st.date_input("Hasta", value=date.today(), key="cob_hasta")

    viajes = filtrar_por_fechas(ingresos, desde, hasta)
    if cliente != "Todos":
        viajes = [r for r in viajes if r.get("contraparte") == cliente]

    if viajes:
        total = sum(int(r.get("monto", 0)) for r in viajes)
        st.metric(f"Total a cobrar ({len(viajes)} viajes)", pesos(total))
        
        col_btn1, col_btn2 = st.columns(2)
        with col_btn1:
            pdf_bytes = generar_pdf_cobro(viajes, cliente, total, desde, hasta)
            st.download_button("📄 Generar PDF Formal", data=pdf_bytes, file_name=f"Estado_Pago_{cliente}.pdf", mime="application/pdf", use_container_width=True, type="primary")
        with col_btn2:
            xlsx = exportar_xlsx(viajes)
            if xlsx:
                st.download_button("📥 Descargar Excel", data=xlsx, file_name=f"Detalle_{cliente}.xlsx", use_container_width=True)
                
        st.dataframe([{
            "Fecha": r.get("fecha", ""), "Camión": r.get("patente", ""), "Material": r.get("categoria", ""), 
            "Cant": r.get("cantidad", ""), "Monto": pesos(r.get("monto", 0))
        } for r in viajes], use_container_width=True, hide_index=True)

def main() -> None:
    st.set_page_config(page_title="S.J.P Transporte", page_icon="🚚", layout="centered")

    st.markdown("""
        <style>
        #MainMenu {visibility: hidden;}
        header {visibility: hidden;}
        footer {visibility: hidden;}
        .block-container {padding-top: 1rem; padding-bottom: 5rem;}
        </style>
    """, unsafe_allow_html=True)

    col_logo, col_titulo = st.columns([1, 4])
    with col_logo:
        try: st.image("logo.png", width=80)
        except: st.write("🚚")
    with col_titulo:
        st.title("S.J.P Transporte")

    if PIN_ACCESO and not st.session_state.get("autenticado", False):
        st.caption("Acceso Restringido al Sistema")
        pin = st.text_input("PIN de Seguridad", type="password")
        if st.button("Ingresar", use_container_width=True):
            if pin == PIN_ACCESO:
                st.session_state["autenticado"] = True
                st.rerun()
            else:
                st.toast("PIN Incorrecto", icon="❌")
        return

    seleccion = option_menu(
        menu_title=None,
        options=["Anotar", "Flota", "Finanzas", "Cobros"],
        icons=["pencil-square", "truck", "pie-chart", "file-earmark-pdf"],
        default_index=0,
        orientation="horizontal",
        styles={
            "container": {"padding": "0!important", "background-color": "#1E293B", "border-radius": "10px", "margin-bottom": "20px"},
            "icon": {"color": "#00B4D8", "font-size": "16px"}, 
            "nav-link": {"font-size": "13px", "text-align": "center", "margin":"0px", "--hover-color": "#334155"},
            "nav-link-selected": {"background-color": "#0F172A", "color": "white", "font-weight": "bold"},
        }
    )

    if seleccion == "Anotar": mostrar_anotar()
    elif seleccion == "Flota": mostrar_flota()
    elif seleccion == "Finanzas": mostrar_finanzas()
    elif seleccion == "Cobros": mostrar_cobros()

if __name__ == "__main__":
    main()
