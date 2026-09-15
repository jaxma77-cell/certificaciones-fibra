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

st.set_page_config(page_title="Certificaciones Fibra", layout="wide", page_icon="🛠️")

# --- CONFIGURACIÓN ---
try:
    GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
except:
    GITHUB_TOKEN = ""

GITHUB_USUARIO = "jaxma77-cell"
GITHUB_REPO = "certificaciones-fibra"

# --- DATOS INICIALES ---
if 'proyectos' not in st.session_state:
    st.session_state.proyectos = {
        "FIBRAMOL JUNIO Y JULIO": {
            "empresa": "Fibranet",
            "fecha": "Jul-26",
            "conceptos": pd.DataFrame({
                'Concepto': ['Preparación extremo de cable', 'Preparación Sangrado', 
                           'Instalación Caja de Empalme, CTO', 'Instalación Spliter',
                           'Empalme a fusión hasta 64fo', 'Empalme a fusión a partir de 64fo',
                           'Medida de potencia de 1 fibra en 2a y 3a ventana'],
                'Precio': [12.50, 25.00, 8.00, 3.00, 5.00, 2.50, 2.00]
            }),
            "filas": []
        }
    }
    st.session_state.proyecto_activo = "FIBRAMOL JUNIO Y JULIO"

# --- CSS PARA QUE TODO SE VEA ---
st.markdown("""
<style>
/* Fondo principal BLANCO */
.stApp { background-color: #ffffff !important; }

/* Texto NEGRO en todas partes */
.stApp * { color: #000000 !important; }

/* Métricas con fondo GRIS y texto NEGRO */
.stMetric { 
    background-color: #f0f0f0 !important; 
    border: 2px solid #cccccc !important;
    border-radius: 8px !important;
    padding: 15px !important;
}
.stMetric label { color: #000000 !important; font-weight: bold !important; }
.stMetric div[data-testid="stMetricValue"] { color: #000000 !important; font-weight: bold !important; font-size: 24px !important; }

/* Sidebar gris claro */
section[data-testid="stSidebar"] { background-color: #f5f5f5 !important; }
section[data-testid="stSidebar"] * { color: #000000 !important; }

/* Botones azules con texto blanco */
.stButton button { 
    background-color: #0066cc !important; 
    color: #ffffff !important;
    border: none !important;
    font-weight: bold !important;
}

/* Tabs visibles */
.stTabs [data-baseweb="tab"] { 
    background-color: #e0e0e0 !important;
    color: #000000 !important;
    font-weight: bold !important;
}
.stTabs [aria-selected="true"] { 
    background-color: #ffffff !important;
    border-top: 3px solid #0066cc !important;
}

/* Inputs y selectores */
input, select { 
    background-color: #ffffff !important; 
    color: #000000 !important;
    border: 1px solid #999999 !important;
}

/* Tablas */
div[data-testid="stDataFrame"] { 
    background-color: #ffffff !important;
    border: 1px solid #cccccc !important;
}
</style>
""", unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.header(" PROYECTOS", divider=True)
    
    # Selector de proyecto
    proyecto_seleccionado = st.selectbox(
        "Selecciona proyecto:",
        options=list(st.session_state.proyectos.keys()),
        index=list(st.session_state.proyectos.keys()).index(st.session_state.proyecto_activo)
    )
    st.session_state.proyecto_activo = proyecto_seleccionado
    
    st.divider()
    
    # Botones de gestión
    if st.button("➕ Nuevo Proyecto", use_container_width=True):
        nombre = st.text_input("Nombre del proyecto:", "Nuevo Proyecto")
        if nombre and nombre not in st.session_state.proyectos:
            st.session_state.proyectos[nombre] = {
                "empresa": "",
                "fecha": "",
                "conceptos": pd.DataFrame({"Concepto": [""], "Precio": [0.0]}),
                "filas": []
            }
            st.session_state.proyecto_activo = nombre
            st.rerun()
    
    if st.button("✏️ Renombrar", use_container_width=True):
        nuevo_nombre = st.text_input("Nuevo nombre:", st.session_state.proyecto_activo)
        if nuevo_nombre and nuevo_nombre != st.session_state.proyecto_activo:
            st.session_state.proyectos[nuevo_nombre] = st.session_state.proyectos.pop(st.session_state.proyecto_activo)
            st.session_state.proyecto_activo = nuevo_nombre
            st.rerun()
    
    if st.button("🗑️ Eliminar", use_container_width=True):
        if len(st.session_state.proyectos) > 1:
            del st.session_state.proyectos[st.session_state.proyecto_activo]
            st.session_state.proyecto_activo = list(st.session_state.proyectos.keys())[0]
            st.rerun()
        else:
            st.error("No puedes eliminar el último proyecto")
    
    st.divider()
    st.subheader("⚙️ Configuración")
    pie_pagina = st.text_input("Pie de página:", "F'BERED INGENIERIA EN REDES DE FIBRA")

# --- PROYECTO ACTIVO ---
p = st.session_state.proyectos[st.session_state.proyecto_activo]
conceptos = p["conceptos"][p["conceptos"]["Concepto"].str.strip() != ""]
conceptos_list = conceptos["Concepto"].tolist()
precios_list = conceptos["Precio"].tolist()

# --- TÍTULO PRINCIPAL ---
st.title(f"️ {st.session_state.proyecto_activo}")
st.caption(f"**{p['empresa']}** · {p['fecha']}")

st.divider()

# --- MÉTRICAS ---
total_general = sum(f.get("total", 0) for f in p["filas"]) if p["filas"] else 0
col1, col2, col3, col4 = st.columns(4)
col1.metric(" TOTAL", f"{total_general:,.2f} €")
col2.metric("📋 FILAS", len(p["filas"]))
col3.metric("🏷️ CONCEPTOS", len(conceptos_list))
col4.metric(" MEDIA", f"{total_general/len(p['filas']) if p['filas'] else 0:,.2f} €")

st.divider()

# --- PESTAÑAS ---
tab1, tab2, tab3 = st.tabs(["📝 REGISTRO", " ANÁLISIS", "⚙️ CONFIGURACIÓN"])

with tab1:
    if not conceptos_list:
        st.error(" No hay conceptos definidos. Ve a CONFIGURACIÓN para añadirlos.")
    else:
        # Botones de acción
        c1, c2, c3 = st.columns(3)
        if c1.button("➕ AÑADIR FILA", use_container_width=True):
            nueva_fila = {"Dia": "", "Nombre": ""}
            for c in conceptos_list:
                nueva_fila[c] = 0
            p["filas"].append(nueva_fila)
            st.rerun()
        
        if c2.button("📋 DUPLICAR ÚLTIMA", use_container_width=True, disabled=not p["filas"]):
            if p["filas"]:
                copia = p["filas"][-1].copy()
                copia["Nombre"] = ""
                p["filas"].append(copia)
                st.rerun()
        
        if c3.button("🗑️ BORRAR TODO", use_container_width=True):
            p["filas"] = []
            st.rerun()
        
        st.divider()
        
        # Filtros
        col_f1, col_f2 = st.columns(2)
        filtro_buscar = col_f1.text_input("🔍 Buscar por nombre:", placeholder="Ej: ALT-CE07")
        dias_disponibles = sorted(set(f.get("Dia", "") for f in p["filas"] if f.get("Dia")))
        filtro_dia = col_f2.selectbox("Filtrar por día:", ["TODOS"] + dias_disponibles)
        
        # Aplicar filtros
        filas_filtradas = p["filas"]
        if filtro_buscar:
            filas_filtradas = [f for f in filas_filtradas if filtro_buscar.upper() in str(f.get("Nombre", "")).upper()]
        if filtro_dia != "TODOS":
            filas_filtradas = [f for f in filas_filtradas if f.get("Dia") == filtro_dia]
        
        st.caption(f"Mostrando {len(filas_filtradas)} de {len(p['filas'])} filas")
        
        if p["filas"]:
            # Crear DataFrame editable
            df = pd.DataFrame(filas_filtradas)
            for c in conceptos_list:
                if c not in df.columns:
                    df[c] = 0
            df = df[["Dia", "Nombre"] + conceptos_list]
            
            # Configurar columnas
            column_config = {
                "Dia": st.column_config.TextColumn("DÍA", width="small"),
                "Nombre": st.column_config.TextColumn("NOMBRE", width="medium")
            }
            for c in conceptos_list:
                column_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1)
            
            # Formulario de edición
            with st.form("formulario_edicion"):
                df_editado = st.data_editor(df, num_rows="dynamic", column_config=column_config, 
                                          hide_index=True, use_container_width=True)
                
                if st.form_submit_button("💾 GUARDAR CAMBIOS", use_container_width=True, type="primary"):
                    nuevas_filas = df_editado.fillna(0).to_dict('records')
                    
                    # Calcular totales
                    for fila in nuevas_filas:
                        total = sum(float(fila.get(c, 0)) * precios_list[i] for i, c in enumerate(conceptos_list))
                        fila["total"] = total
                    
                    if filtro_buscar or filtro_dia != "TODOS":
                        no_filtradas = [f for f in p["filas"] if f not in filas_filtradas]
                        p["filas"] = no_filtradas + nuevas_filas
                    else:
                        p["filas"] = nuevas_filas
                    
                    st.success(f"✅ {len(nuevas_filas)} filas guardadas correctamente")
                    st.rerun()
        
        # Resumen por ubicación
        if p["filas"]:
            st.divider()
            st.subheader("📍 RESUMEN POR UBICACIÓN")
            resumen = {}
            for fila in p["filas"]:
                nombre = fila.get("Nombre", "") or "(sin nombre)"
                if nombre not in resumen:
                    resumen[nombre] = {"filas": 0, "total": 0}
                resumen[nombre]["filas"] += 1
                resumen[nombre]["total"] += fila.get("total", 0)
            
            df_resumen = pd.DataFrame([
                {"Ubicación": nom, "Registros": d["filas"], "Total (€)": f"{d['total']:,.2f}"}
                for nom, d in sorted(resumen.items(), key=lambda x: x[1]["total"], reverse=True)
            ])
            st.dataframe(df_resumen, use_container_width=True, hide_index=True)

with tab2:
    if not p["filas"]:
        st.info(" No hay datos para analizar")
    else:
        st.subheader(" TOTAL POR CONCEPTO")
        totales_concepto = []
        for i, c in enumerate(conceptos_list):
            total = sum(float(f.get(c, 0)) * precios_list[i] for f in p["filas"])
            totales_concepto.append(total)
        
        df_grafico = pd.DataFrame({
            "Concepto": conceptos_list,
            "Total (€)": [f"{t:,.2f}" for t in totales_concepto]
        })
        st.bar_chart(df_grafico.set_index("Concepto"), use_container_width=True)
        
        st.divider()
        st.subheader("📋 ESTADÍSTICAS")
        st.dataframe(pd.DataFrame({
            "Concepto": conceptos_list,
            "Precio": [f"{p:,.2f} €" for p in precios_list],
            "Total": [f"{t:,.2f} €" for t in totales_concepto]
        }), use_container_width=True, hide_index=True)

with tab3:
    st.subheader("🏷️ CONCEPTOS Y PRECIOS")
    
    # Datos de empresa
    col1, col2 = st.columns(2)
    p["empresa"] = col1.text_input("Empresa:", p["empresa"])
    p["fecha"] = col2.text_input("Fecha/Mes:", p["fecha"])
    
    st.divider()
    st.write("**Edita los conceptos y precios:**")
    
    # Editor de conceptos
    conceptos_editados = st.data_editor(
        p["conceptos"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Concepto": st.column_config.TextColumn("CONCEPTO", width="large"),
            "Precio": st.column_config.NumberColumn("PRECIO (€)", format="%.2f", min_value=0.0, step=0.5)
        }
    )
    p["conceptos"] = conceptos_editados
    
    if st.button("🔄 APLICAR CAMBIOS A LAS FILAS", use_container_width=True):
        # Adaptar filas existentes a nuevos conceptos
        nuevos_conceptos = conceptos_editados[conceptos_editados["Concepto"].str.strip() != ""]
        nuevos_conceptos_list = nuevos_conceptos["Concepto"].tolist()
        
        nuevas_filas = []
        for fila in p["filas"]:
            nueva_fila = {"Dia": fila.get("Dia", ""), "Nombre": fila.get("Nombre", "")}
            for c in nuevos_conceptos_list:
                nueva_fila[c] = fila.get(c, 0)
            nuevas_filas.append(nueva_fila)
        
        p["filas"] = nuevas_filas
        st.success("✅ Conceptos actualizados correctamente")
        st.rerun()
    
    st.divider()
    st.subheader("📤 EXPORTAR A EXCEL")
    
    if st.button("🚀 GENERAR EXCEL", type="primary", use_container_width=True, disabled=not p["filas"]):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = st.session_state.proyecto_activo[:31]
        
        num_cols = 2 + len(conceptos_list) + 1
        last_col = get_column_letter(num_cols)
        
        # Encabezados
        for r in range(1, 4):
            ws.merge_cells(f'A{r}:{last_col}{r}')
        
        ws['A1'] = p["empresa"]
        ws['A1'].font = Font(bold=True, size=14)
        ws['A1'].alignment = Alignment(horizontal="center")
        
        ws['A2'] = st.session_state.proyecto_activo
        ws['A2'].font = Font(bold=True, size=12)
        ws['A2'].alignment = Alignment(horizontal="center")
        
        ws['A3'] = p["fecha"]
        ws['A3'].alignment = Alignment(horizontal="center")
        
        # Cabeceras de columna
        ws['A4'] = 'DÍA'
        ws['B4'] = 'NOMBRE'
        for i, c in enumerate(conceptos_list):
            ws.cell(row=4, column=i+3, value=c)
        ws.cell(row=4, column=num_cols, value='TOTAL')
        
        # Estilo cabeceras
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        for col in range(1, num_cols + 1):
            cell = ws.cell(row=4, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center")
        
        # Precios
        for i, precio in enumerate(precios_list):
            cell = ws.cell(row=5, column=i+3, value=precio)
            cell.number_format = '#,##0.00 "€"'
            cell.alignment = Alignment(horizontal="center")
        
        # Datos
        start_row = 6
        for fila in p["filas"]:
            ws.cell(row=start_row, column=1, value=fila.get('Dia', ''))
            ws.cell(row=start_row, column=2, value=fila.get('Nombre', ''))
            
            for i, c in enumerate(conceptos_list):
                val = int(fila.get(c, 0)) if fila.get(c, 0) > 0 else ""
                ws.cell(row=start_row, column=i+3, value=val)
            
            total_celda = ws.cell(row=start_row, column=num_cols, value=fila.get("total", 0))
            total_celda.number_format = '#,##0.00 "€"'
            total_celda.font = Font(bold=True)
            
            start_row += 1
        
        # Total general
        ws.merge_cells(f'A{start_row}:B{start_row}')
        ws.cell(row=start_row, column=1, value='TOTAL GENERAL')
        ws.cell(row=start_row, column=1).font = Font(bold=True, size=12)
        ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="right")
        
        total_general_celda = ws.cell(row=start_row, column=num_cols, value=total_general)
        total_general_celda.number_format = '#,##0.00 "€"'
        total_general_celda.font = Font(bold=True, size=12, color="FF0000")
        
        # Pie de página
        start_row += 2
        ws.merge_cells(f'A{start_row}:{last_col}{start_row}')
        ws.cell(row=start_row, column=1, value=pie_pagina)
        ws.cell(row=start_row, column=1).font = Font(italic=True, size=10)
        ws.cell(row=start_row, column=1).alignment = Alignment(horizontal="center")
        
        # Ajustar anchos
        ws.column_dimensions['A'].width = 12
        ws.column_dimensions['B'].width = 20
        for i in range(3, num_cols + 1):
            ws.column_dimensions[get_column_letter(i)].width = 18
        
        # Descargar
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR EXCEL",
            data=buffer,
            file_name=f"Certificacion_{st.session_state.proyecto_activo.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        st.success("✅ Excel generado correctamente")
