import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import io
import json
import requests
import base64
import os
import plotly.express as px

st.set_page_config(
    page_title="Generador de Certificaciones", 
    layout="wide",
    page_icon="🛠️",
    initial_sidebar_state="expanded"
)

# --- CONFIGURACIÓN GITHUB ---
try:
    GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
except Exception:
    GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

GITHUB_USUARIO = "jaxma77-cell"
GITHUB_REPO = "certificaciones-fibra"
GITHUB_ARCHIVO_BOT = "datos_bot.json"

# --- PLANTILLAS ---
PLANTILLAS = {
    "FIBRAMOL JUNIO Y JULIO": {
        "empresa": "Fibranet",
        "fecha": "Jul-26",
        "conceptos": [
            'Preparación extremo de cable', 'Preparación Sangrado',
            'Instalación Caja de Empalme, CTO', 'Instalación Spliter',
            'Empalme a fusión hasta 64fo', 'Empalme a fusión a partir de 64fo',
            'Medida de potencia de 1 fibra en 2a y 3a ventana'
        ],
        "precios": [12.50, 25.00, 8.00, 3.00, 5.00, 2.50, 2.00]
    },
    "Santomera Mayo 2024": {
        "empresa": "Fibranet Tecnologia y Sistemas SLU",
        "fecha": "May-24",
        "conceptos": [
            'Preparación Extremo de Cable', 'Preparación Sangrado',
            'Instalación Caja de Empalme, CTO', 'Instalación de Spliter',
            'Empalme a Fusión hasta 24fo', 'Medida de potencia'
        ],
        "precios": [18.00, 36.00, 12.00, 3.10, 7.00, 2.50]
    }
}

# --- INICIALIZACIÓN ---
if 'proyectos' not in st.session_state:
    st.session_state.proyectos = {}
    for nombre, datos in PLANTILLAS.items():
        st.session_state.proyectos[nombre] = {
            "empresa": datos["empresa"], "fecha": datos["fecha"],
            "conceptos": pd.DataFrame({"Concepto": datos["conceptos"], "Precio": datos["precios"]}),
            "filas": []
        }
    st.session_state.proyecto_activo = "FIBRAMOL JUNIO Y JULIO"

if 'pie_pagina' not in st.session_state:
    st.session_state.pie_pagina = "F'BERED INGENIERIA EN REDES DE FIBRA"

# --- FUNCIONES ---
def obtener_conceptos_validos(proyecto):
    df = proyecto["conceptos"]
    return df[df['Concepto'].notna() & (df['Concepto'].astype(str).str.strip() != '')].reset_index(drop=True)

def adaptar_filas(filas, conceptos_validos):
    conceptos_list = conceptos_validos['Concepto'].tolist()
    return [{'Dia': f.get('Dia', ''), 'Nombre': f.get('Nombre', ''), **{c: f.get(c, 0) for c in conceptos_list}} for f in filas]

def calcular_totales(filas, conceptos_list, precios_list):
    totales = []
    for fila in filas:
        total = sum(float(fila.get(c, 0) or 0) * precios_list[i] for i, c in enumerate(conceptos_list))
        totales.append(total)
    return totales

def formato_euro(valor):
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

# --- CSS MINIMALISTA ---
st.markdown("""
    <style>
    .stApp { background-color: #ffffff; }
    .stMetric { background-color: #f8f9fa; padding: 20px; border-radius: 8px; }
    .stMetricLabel { color: #6c757d !important; font-size: 13px !important; }
    .stMetricValue { color: #212529 !important; font-size: 28px !important; font-weight: 700 !important; }
    div[data-testid="stSidebar"] { background-color: #f8f9fa; }
    div[data-testid="stSidebar"] * { color: #212529 !important; }
    .stButton>button { background-color: #0d6efd; color: white !important; border-radius: 6px; }
    .stTabs [data-baseweb="tab"] { background-color: #f8f9fa; }
    .stTabs [aria-selected="true"] { background-color: #ffffff !important; }
    </style>
""", unsafe_allow_html=True)

# --- BARRA LATERAL - PROYECTOS VISIBLES ---
with st.sidebar:
    st.markdown("### 📁 PROYECTOS")
    
    # Selector MUY visible
    proyecto_activo = st.selectbox(
        "Selecciona proyecto:",
        list(st.session_state.proyectos.keys()),
        index=list(st.session_state.proyectos.keys()).index(st.session_state.proyecto_activo)
    )
    st.session_state.proyecto_activo = proyecto_activo
    
    st.markdown("---")
    
    # Botones de gestión
    if st.button("➕ Nuevo Proyecto", use_container_width=True):
        nombre = st.text_input("Nombre:", "Nuevo Proyecto")
        if nombre and nombre not in st.session_state.proyectos:
            st.session_state.proyectos[nombre] = {
                "empresa": "", "fecha": "",
                "conceptos": pd.DataFrame({"Concepto": [""], "Precio": [0.0]}),
                "filas": []
            }
            st.session_state.proyecto_activo = nombre
            st.rerun()
    
    if st.button("️ Renombrar", use_container_width=True):
        nuevo = st.text_input("Nuevo nombre:", st.session_state.proyecto_activo)
        if nuevo and nuevo != st.session_state.proyecto_activo and nuevo not in st.session_state.proyectos:
            st.session_state.proyectos[nuevo] = st.session_state.proyectos.pop(st.session_state.proyecto_activo)
            st.session_state.proyecto_activo = nuevo
            st.rerun()
    
    if st.button("🗑️ Eliminar", use_container_width=True):
        if len(st.session_state.proyectos) > 1:
            del st.session_state.proyectos[st.session_state.proyecto_activo]
            st.session_state.proyecto_activo = list(st.session_state.proyectos.keys())[0]
            st.rerun()
    
    st.markdown("---")
    st.markdown("### ️ Configuración")
    st.session_state.pie_pagina = st.text_input("Pie de página:", st.session_state.pie_pagina)

# --- LÓGICA PRINCIPAL ---
p = st.session_state.proyectos[st.session_state.proyecto_activo]
conceptos_validos = obtener_conceptos_validos(p)
conceptos_list = conceptos_validos['Concepto'].tolist()
precios_list = conceptos_validos['Precio'].tolist()

if p["filas"]:
    p["filas"] = adaptar_filas(p["filas"], conceptos_validos)

totales = calcular_totales(p["filas"], conceptos_list, precios_list)
total_general = sum(totales) if totales else 0
p["totales"] = totales
p["total_general"] = total_general

# --- TÍTULO ---
st.title(f"🛠️ {st.session_state.proyecto_activo}")
st.caption(f"{p['empresa']} · {p['fecha']}")

# --- MÉTRICAS ---
col1, col2, col3, col4 = st.columns(4)
col1.metric(" Total", formato_euro(total_general))
col2.metric("📋 Filas", len(p["filas"]))
col3.metric("🏷️ Conceptos", len(conceptos_list))
col4.metric("📊 Media", formato_euro(total_general/len(p["filas"]) if p["filas"] else 0))

st.markdown("---")

# --- PESTAÑAS ---
tab1, tab2, tab3 = st.tabs(["📝 Registro", "📊 Análisis", "⚙️ Configuración"])

with tab1:
    if not conceptos_list:
        st.error("❌ Añade conceptos en Configuración")
    else:
        # Botones
        c1, c2, c3 = st.columns(3)
        if c1.button("➕ Añadir", use_container_width=True):
            p["filas"].append({'Dia': '', 'Nombre': '', **{c: 0 for c in conceptos_list}})
            st.rerun()
        if c2.button("📋 Duplicar", use_container_width=True) and p["filas"]:
            copia = p["filas"][-1].copy()
            copia['Nombre'] = ''
            p["filas"].append(copia)
            st.rerun()
        if c3.button("️ Borrar todo", use_container_width=True):
            p["filas"] = []
            st.rerun()
        
        st.markdown("---")
        
        # Filtros
        col_f1, col_f2 = st.columns(2)
        filtro_nombre = col_f1.text_input("Buscar:", placeholder="ALT-CE07")
        dias = sorted(set(f.get('Dia', '') for f in p["filas"] if f.get('Dia')))
        filtro_dia = col_f2.selectbox("Día:", ["Todos"] + dias)
        
        # Filtrar
        filas_vistas = p["filas"]
        if filtro_nombre:
            filas_vistas = [f for f in filas_vistas if filtro_nombre.upper() in str(f.get('Nombre', '')).upper()]
        if filtro_dia != "Todos":
            filas_vistas = [f for f in filas_vistas if f.get('Dia') == filtro_dia]
        
        st.caption(f"Viendo {len(filas_vistas)} de {len(p['filas'])} filas")
        
        if p["filas"]:
            df = pd.DataFrame(filas_vistas)
            for c in conceptos_list:
                if c not in df.columns:
                    df[c] = 0
            df = df[['Dia', 'Nombre'] + conceptos_list]
            
            col_config = {"Dia": st.column_config.TextColumn("DÍA", width="small"), "Nombre": st.column_config.TextColumn("NOMBRE")}
            for c in conceptos_list:
                col_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1)
            
            with st.form("editar"):
                df_edit = st.data_editor(df, num_rows="dynamic", column_config=col_config, hide_index=True, use_container_width=True)
                if st.form_submit_button(" Guardar", use_container_width=True, type="primary"):
                    nuevas = df_edit.fillna(0).to_dict('records')
                    if filtro_nombre or filtro_dia != "Todos":
                        no_filtradas = [f for f in p["filas"] if f not in filas_vistas]
                        p["filas"] = no_filtradas + nuevas
                    else:
                        p["filas"] = nuevas
                    st.success(f"✅ {len(nuevas)} filas guardadas")
                    st.rerun()
        
        # Resumen por nombre
        if p["filas"]:
            st.markdown("---")
            st.markdown("#### 📍 Por Ubicación")
            resumen = {}
            for i, f in enumerate(p["filas"]):
                nom = f.get('Nombre', '') or '(sin nombre)'
                resumen[nom] = resumen.get(nom, {'filas': 0, 'total': 0})
                resumen[nom]['filas'] += 1
                resumen[nom]['total'] += totales[i]
            
            st.dataframe(pd.DataFrame([
                {'Nombre': n, 'Registros': d['filas'], 'Total': formato_euro(d['total'])}
                for n, d in sorted(resumen.items(), key=lambda x: x[1]['total'], reverse=True)
            ]), use_container_width=True, hide_index=True)

with tab2:
    if not p["filas"]:
        st.info("Sin datos")
    else:
        st.markdown("### 📈 Por Concepto")
        totales_concepto = [sum(float(f.get(c, 0) or 0) * precios_list[i] for f in p["filas"]) for i, c in enumerate(conceptos_list)]
        
        fig = px.bar(x=conceptos_list, y=totales_concepto, labels={'x': 'Concepto', 'y': 'Total (€)'}, color=totales_concepto, color_continuous_scale='Blues')
        fig.update_layout(height=400, showlegend=False, xaxis_tickangle=-30)
        st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.markdown("### 🏷️ Conceptos y Precios")
    p["empresa"] = st.text_input("Empresa:", p["empresa"])
    p["fecha"] = st.text_input("Fecha:", p["fecha"])
    
    edited = st.data_editor(p["conceptos"], num_rows="dynamic", use_container_width=True,
                           column_config={"Concepto": st.column_config.TextColumn(width="large"), "Precio": st.column_config.NumberColumn(format="%.2f €")})
    p["conceptos"] = edited.reset_index(drop=True)
    
    if st.button("🔄 Aplicar cambios", use_container_width=True):
        p["filas"] = adaptar_filas(p["filas"], obtener_conceptos_validos(p))
        st.success("✅ Actualizado")
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📤 Exportar Excel")
    if st.button("🚀 Generar Excel", type="primary", use_container_width=True) and p["filas"]:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = st.session_state.proyecto_activo[:31]
        
        num_cols = 2 + len(conceptos_list) + 1
        last_col = get_column_letter(num_cols)
        
        # Encabezados
        for r in range(1, 4):
            ws.merge_cells(f'A{r}:{last_col}{r}')
        ws['A1'], ws['A1'].font = p["empresa"], Font(bold=True, size=14)
        ws['A2'], ws['A2'].font = st.session_state.proyecto_activo, Font(bold=True, size=12)
        ws['A3'] = p["fecha"]
        
        ws['A4'], ws['B4'] = 'DÍA', 'NOMBRE'
        for i, c in enumerate(conceptos_list):
            ws.cell(row=4, column=i+3, value=c)
        ws.cell(row=4, column=num_cols, value='TOTAL')
        
        # Precios
        for i, precio in enumerate(precios_list):
            ws.cell(row=5, column=i+3, value=precio).number_format = '#,##0.00 "€"'
        
        # Datos
        start_row = 6
        for idx, fila in enumerate(p["filas"]):
            ws.cell(row=start_row, column=1, value=fila.get('Dia', ''))
            ws.cell(row=start_row, column=2, value=fila.get('Nombre', ''))
            for i, c in enumerate(conceptos_list):
                val = int(fila.get(c, 0)) if fila.get(c, 0) > 0 else ""
                ws.cell(row=start_row, column=i+3, value=val)
            ws.cell(row=start_row, column=num_cols, value=p["totales"][idx]).number_format = '#,##0.00 "€"'
            start_row += 1
        
        # Total general
        ws.merge_cells(f'A{start_row}:B{start_row}')
        ws.cell(row=start_row, column=1, value='TOTAL GENERAL').font = Font(bold=True)
        ws.cell(row=start_row, column=num_cols, value=total_general).font = Font(bold=True, color="FF0000")
        ws.cell(row=start_row, column=num_cols).number_format = '#,##0.00 "€"'
        
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button("📥 Descargar", buffer, f"Certificacion_{st.session_state.proyecto_activo.replace(' ', '_')}.xlsx", use_container_width=True)
        st.success("✅ Generado")
