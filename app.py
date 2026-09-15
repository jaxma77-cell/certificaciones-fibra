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
import plotly.graph_objects as go

st.set_page_config(
    page_title="Generador de Certificaciones", 
    layout="wide",
    page_icon="🛠️",
    initial_sidebar_state="expanded"
)

# --- CONFIGURACIÓN GITHUB PARA EL BOT ---
try:
    GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
except Exception:
    GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

GITHUB_USUARIO = "jaxma77-cell"
GITHUB_REPO = "certificaciones-fibra"
GITHUB_ARCHIVO_BOT = "datos_bot.json"

# --- PLANTILLAS DE PROYECTOS ---
PLANTILLAS = {
    "FIBRAMOL JUNIO Y JULIO": {
        "empresa": "Fibranet",
        "fecha": "Jul-26",
        "conceptos": [
            'Preparación extremo de cable',
            'Preparación Sangrado',
            'Instalación Caja de Empalme, CTO',
            'Instalación Spliter',
            'Empalme a fusión hasta 64fo',
            'Empalme a fusión a partir de 64fo',
            'Medida de potencia de 1 fibra en 2a y 3a ventana'
        ],
        "precios": [12.50, 25.00, 8.00, 3.00, 5.00, 2.50, 2.00]
    },
    "Santomera Mayo 2024 (Fusionador)": {
        "empresa": "Fibranet Tecnologia y Sistemas SLU",
        "fecha": "May-24",
        "conceptos": [
            'Preparación Extremo de Cable',
            'Preparación Sangrado',
            'Instalación Caja de Empalme, CTO',
            'Instalación de Spliter',
            'Empalme a Fusión hasta 24fo',
            'Medida de potencia de 1 fibra en 2a y 3a ventana'
        ],
        "precios": [18.00, 36.00, 12.00, 3.10, 7.00, 2.50]
    }
}

# --- INICIALIZACIÓN ---
if 'proyectos' not in st.session_state:
    st.session_state.proyectos = {}
    for nombre, datos in PLANTILLAS.items():
        st.session_state.proyectos[nombre] = {
            "empresa": datos["empresa"],
            "fecha": datos["fecha"],
            "conceptos": pd.DataFrame({
                "Concepto": datos["conceptos"],
                "Precio": datos["precios"]
            }),
            "filas": []
        }
    st.session_state.proyecto_activo = "FIBRAMOL JUNIO Y JULIO"

if 'pie_pagina' not in st.session_state:
    st.session_state.pie_pagina = "F'BERED INGENIERIA EN REDES DE FIBRA"

# --- FUNCIONES AUXILIARES ---
def obtener_conceptos_validos(proyecto):
    df = proyecto["conceptos"]
    validos = df[df['Concepto'].notna() & (df['Concepto'].astype(str).str.strip() != '')]
    return validos.reset_index(drop=True)

def adaptar_filas_a_conceptos(filas, conceptos_validos):
    conceptos_list = conceptos_validos['Concepto'].tolist()
    filas_adaptadas = []
    for fila in filas:
        nueva_fila = {'Dia': fila.get('Dia', ''), 'Nombre': fila.get('Nombre', '')}
        for c in conceptos_list:
            nueva_fila[c] = fila.get(c, 0)
        filas_adaptadas.append(nueva_fila)
    return filas_adaptadas

def calcular_totales(filas, conceptos_list, precios_list):
    totales = []
    for fila in filas:
        total = 0.0
        for i, c in enumerate(conceptos_list):
            try:
                qty = float(fila.get(c, 0) or 0)
                total += qty * precios_list[i]
            except:
                pass
        totales.append(total)
    return totales

def formato_euro(valor):
    return f"{valor:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")

# --- ESTILOS CSS MEJORADOS ---
modo_oscuro = False

with st.sidebar:
    modo_oscuro = st.toggle("🌙 Modo Oscuro", value=False, key="dark_mode_toggle")

# Aplicar estilos según el modo
if modo_oscuro:
    st.markdown("""
        <style>
        .stApp {
            background-color: #0e1117;
            color: #fafafa;
        }
        .stMetric {
            background-color: #262730;
            border: 1px solid #3a3a3a;
            padding: 15px;
            border-radius: 10px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        .stMetricLabel {
            color: #fafafa !important;
        }
        .stMetricValue {
            color: #4CAF50 !important;
        }
        div[data-testid="stSidebar"] {
            background-color: #1a1a1a;
        }
        </style>
    """, unsafe_allow_html=True)
else:
    # MODO CLARO MEJORADO - MÉTRICAS BIEN VISIBLES
    st.markdown("""
        <style>
        .stApp {
            background-color: #f0f2f6;
            color: #212529;
        }
        
        /* Métricas MUY visibles en modo claro */
        .stMetric {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border: none;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 8px 16px rgba(102, 126, 234, 0.3);
            transition: all 0.3s ease;
        }
        .stMetric:hover {
            box-shadow: 0 12px 24px rgba(102, 126, 234, 0.4);
            transform: translateY(-3px);
        }
        .stMetricLabel {
            color: #ffffff !important;
            font-weight: 600 !important;
            font-size: 15px !important;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        .stMetricValue {
            color: #ffffff !important;
            font-weight: 800 !important;
            font-size: 32px !important;
            text-shadow: 0 2px 4px rgba(0,0,0,0.2);
        }
        
        /* Barra lateral */
        div[data-testid="stSidebar"] {
            background-color: #ffffff;
            border-right: 2px solid #e9ecef;
        }
        div[data-testid="stSidebar"] h1, 
        div[data-testid="stSidebar"] h2, 
        div[data-testid="stSidebar"] h3 {
            color: #212529 !important;
        }
        
        /* Pestañas */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
        }
        .stTabs [data-baseweb="tab"] {
            background-color: #ffffff;
            border-radius: 8px 8px 0 0;
            padding: 12px 20px;
            border: 2px solid #e9ecef;
            color: #495057;
            font-weight: 600;
        }
        .stTabs [aria-selected="true"] {
            background-color: #0d6efd !important;
            color: #ffffff !important;
            border-color: #0d6efd !important;
        }
        
        /* Botones */
        .stButton>button {
            background-color: #0d6efd;
            color: #ffffff;
            border: none;
            border-radius: 8px;
            padding: 10px 20px;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        .stButton>button:hover {
            background-color: #0b5ed7;
            box-shadow: 0 4px 8px rgba(13, 110, 253, 0.3);
        }
        .stButton>button[kind="primary"] {
            background-color: #198754;
        }
        .stButton>button[kind="primary"]:hover {
            background-color: #157347;
        }
        
        /* Tablas */
        .stDataFrame {
            border: 2px solid #e9ecef;
            border-radius: 8px;
            overflow: hidden;
        }
        
        /* Títulos */
        h1, h2, h3, h4 {
            color: #212529 !important;
            font-weight: 700 !important;
        }
        
        /* Alertas */
        .stAlert {
            border-radius: 8px;
            border: 2px solid;
        }
        
        /* Selectores */
        .stSelectbox > div > div {
            background-color: #ffffff;
            border: 2px solid #e9ecef;
            border-radius: 8px;
        }
        
        /* Inputs de texto */
        .stTextInput > div > div > input {
            background-color: #ffffff;
            border: 2px solid #e9ecef;
            border-radius: 8px;
            color: #212529;
        }
        .stTextInput > div > div > input:focus {
            border-color: #0d6efd;
            box-shadow: 0 0 0 0.2rem rgba(13, 110, 253, 0.25);
        }
        </style>
    """, unsafe_allow_html=True)

# --- BARRA LATERAL ---
with st.sidebar:
    st.markdown("## 🛠️ Certificaciones Fibra")
    st.markdown("---")
    
    st.markdown("### 📁 Proyectos")
    nombres_proyectos = list(st.session_state.proyectos.keys())
    idx_actual = nombres_proyectos.index(st.session_state.proyecto_activo) if st.session_state.proyecto_activo in nombres_proyectos else 0
    proyecto_activo = st.selectbox(
        "Proyecto activo",
        nombres_proyectos,
        index=idx_actual,
        label_visibility="collapsed"
    )
    st.session_state.proyecto_activo = proyecto_activo
    
    with st.expander("➕ Crear nuevo proyecto"):
        nombre_nuevo = st.text_input("Nombre del nuevo proyecto", "Nuevo Proyecto", key="nuevo_proy")
        if st.button("Crear proyecto", use_container_width=True):
            if nombre_nuevo and nombre_nuevo not in st.session_state.proyectos:
                st.session_state.proyectos[nombre_nuevo] = {
                    "empresa": "Mi Empresa", "fecha": "",
                    "conceptos": pd.DataFrame({"Concepto": [""], "Precio": [0.0]}),
                    "filas": []
                }
                st.session_state.proyecto_activo = nombre_nuevo
                st.rerun()
            elif nombre_nuevo in st.session_state.proyectos:
                st.error("Ya existe un proyecto con ese nombre")
    
    with st.expander("✏️ Renombrar proyecto"):
        nuevo_nombre = st.text_input("Nuevo nombre", st.session_state.proyecto_activo, key="input_renombrar")
        if st.button("Aplicar nuevo nombre", use_container_width=True):
            if nuevo_nombre and nuevo_nombre != st.session_state.proyecto_activo:
                if nuevo_nombre not in st.session_state.proyectos:
                    st.session_state.proyectos[nuevo_nombre] = st.session_state.proyectos[st.session_state.proyecto_activo]
                    del st.session_state.proyectos[st.session_state.proyecto_activo]
                    st.session_state.proyecto_activo = nuevo_nombre
                    st.success(f"✅ Renombrado a '{nuevo_nombre}'")
                    st.rerun()
                else:
                    st.error("Ya existe un proyecto con ese nombre")
    
    with st.expander("🗑️ Borrar proyecto"):
        st.warning("⚠️ Esta acción no se puede deshacer.")
        confirmacion = st.text_input(f"Escribe '{st.session_state.proyecto_activo}' para confirmar", key="input_borrar")
        if st.button("BORRAR PROYECTO", use_container_width=True, type="primary"):
            if confirmacion == st.session_state.proyecto_activo:
                if len(st.session_state.proyectos) > 1:
                    del st.session_state.proyectos[st.session_state.proyecto_activo]
                    st.session_state.proyecto_activo = list(st.session_state.proyectos.keys())[0]
                    st.success("✅ Proyecto borrado")
                    st.rerun()
                else:
                    st.error("No puedes borrar el único proyecto")
            else:
                st.error("❌ El nombre no coincide")
    
    st.markdown("---")
    st.markdown("### ⚙️ General")
    st.session_state.pie_pagina = st.text_input("Pie de página", st.session_state.pie_pagina, label_visibility="collapsed")

# --- LÓGICA PRINCIPAL ---
p = st.session_state.proyectos[st.session_state.proyecto_activo]
conceptos_validos = obtener_conceptos_validos(p)
conceptos_list = conceptos_validos['Concepto'].tolist()
precios_list = conceptos_validos['Precio'].tolist()

if p["filas"]:
    p["filas"] = adaptar_filas_a_conceptos(p["filas"], conceptos_validos)

totales_fila = calcular_totales(p["filas"], conceptos_list, precios_list)
total_general = sum(totales_fila)
p["totales"] = totales_fila
p["total_general"] = total_general

# --- ENCABEZADO CON DASHBOARD ---
st.title(f"️ {st.session_state.proyecto_activo}")
st.caption(f"**{p['empresa']}** · {p['fecha']}")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric(" Total General", formato_euro(total_general))
with col2:
    st.metric("📋 Filas", len(p["filas"]))
with col3:
    st.metric("🏷️ Conceptos", len(conceptos_list))
with col4:
    promedio = total_general / len(p["filas"]) if p["filas"] else 0
    st.metric("📊 Media por fila", formato_euro(promedio))

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs([
    "📝 Registro de Trabajos", 
    "📊 Análisis y Gráficos",
    "📥 Importar Datos",
    "⚙️ Configuración"
])

# --- PESTAÑA 1: REGISTRO DE TRABAJOS ---
with tab1:
    if len(conceptos_list) == 0:
        st.error("❌ **No hay conceptos definidos.** Ve a la pestaña de Configuración y añade al menos uno.")
    else:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            if st.button("➕ Añadir fila", use_container_width=True):
                nueva = {'Dia': '', 'Nombre': ''}
                for c in conceptos_list:
                    nueva[c] = 0
                p["filas"].append(nueva)
                st.rerun()
        with col2:
            if st.button("📋 Duplicar última", use_container_width=True) and p["filas"]:
                copia = p["filas"][-1].copy()
                copia['Nombre'] = ''
                p["filas"].append(copia)
                st.rerun()
        with col3:
            if st.button("🗑️ Borrar todo", use_container_width=True):
                p["filas"] = []
                st.rerun()
        
        st.markdown("---")
        
        st.markdown("#### 🔍 Filtros")
        col_f1, col_f2, col_f3 = st.columns([2, 2, 1])
        
        with col_f1:
            filtro_nombre = st.text_input("Buscar por nombre", placeholder="Ej: ALT-CE07", key="filtro_nombre")
        
        with col_f2:
            dias_unicos = sorted(list(set(f.get('Dia', '') for f in p["filas"] if f.get('Dia', ''))))
            dias_opciones = ["Todos los días"] + dias_unicos
            filtro_dia = st.selectbox("Filtrar por día", dias_opciones, key="filtro_dia")
        
        with col_f3:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("🧹 Limpiar filtros", use_container_width=True):
                st.session_state.filtro_nombre = ""
                st.session_state.filtro_dia = "Todos los días"
                st.rerun()
        
        filas_filtradas = p["filas"]
        if filtro_nombre:
            filas_filtradas = [f for f in filas_filtradas if filtro_nombre.upper() in str(f.get('Nombre', '')).upper()]
        if filtro_dia != "Todos los días":
            filas_filtradas = [f for f in filas_filtradas if f.get('Dia', '') == filtro_dia]
        
        st.caption(f"👁️ Mostrando **{len(filas_filtradas)}** de **{len(p['filas'])}** filas")
        
        if not p["filas"]:
            st.info("💡 Pulsa **'➕ Añadir fila'** para empezar a registrar trabajos.")
        else:
            df_edit = pd.DataFrame(filas_filtradas) if filas_filtradas else pd.DataFrame(columns=['Dia', 'Nombre'] + conceptos_list)
            if not df_edit.empty:
                for c in conceptos_list:
                    if c not in df_edit.columns:
                        df_edit[c] = 0
                df_edit = df_edit[['Dia', 'Nombre'] + conceptos_list]

                col_config = {
                    "Dia": st.column_config.TextColumn("DÍA", width="small"),
                    "Nombre": st.column_config.TextColumn("NOMBRE", width="medium"),
                }
                for c in conceptos_list:
                    col_config[c] = st.column_config.NumberColumn(c, min_value=0, step=1, format="%d", width="small")

                with st.form(key=f"form_datos_{st.session_state.proyecto_activo}"):
                    df_editado = st.data_editor(
                        df_edit,
                        num_rows="dynamic",
                        use_container_width=True,
                        column_config=col_config,
                        key=f"editor_datos_{st.session_state.proyecto_activo}",
                        hide_index=True
                    )
                    submit = st.form_submit_button("💾 Guardar cambios", use_container_width=True, type="primary")
                    
                    if submit:
                        filas_nuevas = df_editado.fillna(0).to_dict('records')
                        if filtro_nombre or filtro_dia != "Todos los días":
                            filas_no_filtradas = [f for f in p["filas"] if f not in filas_filtradas]
                            p["filas"] = filas_no_filtradas + filas_nuevas
                        else:
                            p["filas"] = filas_nuevas
                        st.success(f"✅ {len(filas_nuevas)} filas guardadas correctamente en memoria")
                        st.rerun()

        if p["filas"]:
            st.markdown("---")
            st.markdown("#### 📍 Resumen acumulado por Nombre / Ubicación")
            
            resumen_nombre = {}
            for i, fila in enumerate(p["filas"]):
                nombre = fila.get('Nombre', '') or '(sin nombre)'
                if nombre not in resumen_nombre:
                    resumen_nombre[nombre] = {'filas': 0, 'total': 0.0}
                resumen_nombre[nombre]['filas'] += 1
                resumen_nombre[nombre]['total'] += totales_fila[i]
            
            df_resumen_nombre = pd.DataFrame([
                {'Nombre': nombre, 'Registros': datos['filas'], 'Total Acumulado (€)': formato_euro(datos['total'])}
                for nombre, datos in sorted(resumen_nombre.items(), key=lambda x: x[1]['total'], reverse=True)
            ])
            st.dataframe(df_resumen_nombre, use_container_width=True, hide_index=True)

# --- PESTAÑA 2: ANÁLISIS Y GRÁFICOS ---
with tab2:
    if not p["filas"] or len(conceptos_list) == 0:
        st.info("📊 No hay datos para analizar. Añade filas en la pestaña de Registro.")
    else:
        st.markdown("### 📈 Gráficos y estadísticas")
        
        st.markdown("#### 💼 Total acumulado por concepto")
        totales_por_concepto = []
        for i, c in enumerate(conceptos_list):
            total_c = 0.0
            for fila in p["filas"]:
                try:
                    qty = float(fila.get(c, 0) or 0)
                    total_c += qty * precios_list[i]
                except:
                    pass
            totales_por_concepto.append(total_c)
        
        df_conceptos = pd.DataFrame({'Concepto': conceptos_list, 'Total (€)': totales_por_concepto})
        
        fig1 = px.bar(df_conceptos, x='Concepto', y='Total (€)', color='Total (€)', 
                     color_continuous_scale='Blues', text_auto='.2f €')
        fig1.update_layout(height=400, xaxis_tickangle=-30, showlegend=False)
        st.plotly_chart(fig1, use_container_width=True)
        
        st.markdown("---")
        
        st.markdown("#### 📋 Estadísticas detalladas por concepto")
        stats_conceptos = []
        for i, c in enumerate(conceptos_list):
            cantidades = []
            for fila in p["filas"]:
                try:
                    qty = float(fila.get(c, 0) or 0)
                    if qty > 0: cantidades.append(qty)
                except: pass
            
            stats_conceptos.append({
                'Concepto': c,
                'Precio unit.': formato_euro(precios_list[i]),
                'Veces usado': len(cantidades),
                'Unidades tot.': int(sum(cantidades)),
                'Total (€)': formato_euro(totales_por_concepto[i])
            })
        st.dataframe(pd.DataFrame(stats_conceptos), use_container_width=True, hide_index=True)

# --- PESTAÑA 3: IMPORTAR DATOS ---
with tab3:
    tab_telegram, tab_excel = st.tabs(["🤖 Desde Telegram", "📄 Desde Excel"])
    
    with tab_telegram:
        st.info("Lee los datos guardados por el bot de Telegram en GitHub")
        if not GITHUB_TOKEN:
            st.warning("⚠️ No hay token de GitHub configurado en los Secrets.")
        else:
            if st.button("🔄 Leer datos del bot", type="primary", use_container_width=True):
                try:
                    url = f"https://api.github.com/repos/{GITHUB_USUARIO}/{GITHUB_REPO}/contents/{GITHUB_ARCHIVO_BOT}"
                    headers = {"Authorization": f"token {GITHUB_TOKEN}"}
                    r = requests.get(url, headers=headers)
                    if r.status_code == 200:
                        contenido = base64.b64decode(r.json()["content"]).decode("utf-8")
                        datos_bot = json.loads(contenido)
                        st.success("✅ Datos leídos correctamente")
                        if datos_bot:
                            st.markdown("**Proyectos disponibles:**")
                            for pn, dp in datos_bot.items():
                                nf = len(dp.get("filas", []))
                                tp = sum(f.get("total", 0) for f in dp.get("filas", []))
                                st.markdown(f"- **{pn}**: {nf} filas → {formato_euro(tp)}")
                            st.session_state.datos_bot = datos_bot
                            st.rerun()
                    elif r.status_code == 404:
                        st.warning("⚠️ El archivo `datos_bot.json` no existe aún.")
                    else:
                        st.error(f"❌ Error: {r.status_code}")
                except Exception as e:
                    st.error(f"❌ Error: {e}")
            
            if 'datos_bot' in st.session_state and st.session_state.datos_bot:
                st.markdown("---")
                proyecto_seleccionado = st.selectbox("¿Qué proyecto del bot importar?", list(st.session_state.datos_bot.keys()))
                filas_bot = st.session_state.datos_bot[proyecto_seleccionado].get("filas", [])
                
                if filas_bot:
                    st.markdown(f"**{len(filas_bot)} filas disponibles**")
                    modo_importacion = st.radio("¿Cómo importar?", ["Añadir a las filas existentes", "Reemplazar todas las filas"], horizontal=True)
                    if st.button("🚀 IMPORTAR DATOS", type="primary", use_container_width=True):
                        filas_importadas = []
                        for f in filas_bot:
                            fila_nueva = {'Dia': '', 'Nombre': f['Nombre']}
                            cantidades = f.get('cantidades', [])
                            for i, c in enumerate(conceptos_list):
                                fila_nueva[c] = cantidades[i] if i < len(cantidades) else 0
                            filas_importadas.append(fila_nueva)
                        
                        if modo_importacion == "Añadir a las filas existentes":
                            p["filas"].extend(filas_importadas)
                        else:
                            p["filas"] = filas_importadas
                        st.success(f"✅ ¡{len(filas_importadas)} filas importadas!")
                        st.balloons()
                        st.rerun()
    
    with tab_excel:
        st.info("Sube un Excel generado previamente por esta app")
        archivo = st.file_uploader("Selecciona archivo .xlsx", type=["xlsx"])
        if archivo:
            try:
                wb = openpyxl.load_workbook(archivo)
                ws = wb.active
                empresa_imp = ws['A1'].value or ""
                proyecto_imp = ws['A2'].value or "Importado"
                fecha_imp = ws['A3'].value or ""
                conceptos_imp, precios_imp = [], []
                col = 3
                while ws.cell(row=4, column=col).value and str(ws.cell(row=4, column=col).value).strip().upper() != 'TOTAL':
                    conceptos_imp.append(ws.cell(row=4, column=col).value)
                    precios_imp.append(float(ws.cell(row=5, column=col).value or 0))
                    col += 1
                filas_imp = []
                row = 6
                while ws.cell(row=row, column=1).value or ws.cell(row=row, column=2).value:
                    fila = {'Dia': ws.cell(row=row, column=1).value or '', 'Nombre': ws.cell(row=row, column=2).value or ''}
                    for i, c in enumerate(conceptos_imp):
                        val = ws.cell(row=row, column=i+3).value
                        fila[c] = int(float(val)) if val else 0
                    filas_imp.append(fila)
                    row += 1
                st.success(f"✅ Leído: {len(filas_imp)} filas, {len(conceptos_imp)} conceptos")
                if st.button("💾 Cargar en un nuevo proyecto", type="primary"):
                    nombre_nuevo = f"{proyecto_imp} (importado)"
                    st.session_state.proyectos[nombre_nuevo] = {
                        "empresa": empresa_imp, "fecha": str(fecha_imp),
                        "conceptos": pd.DataFrame({"Concepto": conceptos_imp, "Precio": precios_imp}),
                        "filas": filas_imp
                    }
                    st.session_state.proyecto_activo = nombre_nuevo
                    st.rerun()
            except Exception as e:
                st.error(f"Error al leer: {e}")

# --- PESTAÑA 4: CONFIGURACIÓN ---
with tab4:
    tab_conceptos, tab_exportar, tab_precios = st.tabs(["️ Conceptos y Precios", " Exportar Excel", "💾 Guardar/Cargar"])
    
    with tab_conceptos:
        st.markdown(f"### 📋 Proyecto: {st.session_state.proyecto_activo}")
        col1, col2 = st.columns(2)
        with col1: p["empresa"] = st.text_input("Nombre de la empresa", p["empresa"])
        with col2: p["fecha"] = st.text_input("Fecha / Mes", p["fecha"])
        
        st.markdown("---")
        st.markdown("**Conceptos y precios unitarios**")
        edited = st.data_editor(p["conceptos"], num_rows="dynamic", use_container_width=True, 
                               key=f"editor_conceptos_{st.session_state.proyecto_activo}", 
                               column_config={"Concepto": st.column_config.TextColumn("Concepto", width="large"), 
                                             "Precio": st.column_config.NumberColumn("Precio (€)", format="%.2f €", min_value=0.0, step=0.5)})
        p["conceptos"] = edited.reset_index(drop=True)
        
        if st.button("🔄 Aplicar cambios de conceptos a las filas", use_container_width=True, type="secondary"):
            cv = obtener_conceptos_validos(p)
            if len(cv) == 0: st.error("❌ No hay ningún concepto válido")
            else:
                p["filas"] = adaptar_filas_a_conceptos(p["filas"], cv)
                st.success(f"✅ Filas actualizadas con {len(cv)} conceptos")
                st.rerun()
    
    with tab_exportar:
        st.markdown("### 📊 Exportar a Excel con formato oficial")
        if not p["filas"] or len(conceptos_list) == 0:
            st.error("⚠️ No hay datos o no hay conceptos definidos")
        else:
            if st.button("🚀 Generar Excel", type="primary", use_container_width=True):
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.title = st.session_state.proyecto_activo[:31]
                num_cols = 2 + len(conceptos_list) + 1
                last_col = get_column_letter(num_cols)
                bold14, bold12 = Font(bold=True, size=14), Font(bold=True, size=12)
                header_font, header_fill = Font(bold=True, size=11, color="FFFFFF"), PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
                center, right = Alignment(horizontal="center", vertical="center", wrap_text=True), Alignment(horizontal="right", vertical="center")
                thin = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
                currency_fmt = '#,##0.00 "€"'

                for r in range(1, 4): ws.merge_cells(f'A{r}:{last_col}{r}')
                ws['A1'], ws['A1'].font, ws['A1'].alignment = p["empresa"], bold14, center
                ws['A2'], ws['A2'].font, ws['A2'].alignment = st.session_state.proyecto_activo, bold12, center
                ws['A3'], ws['A3'].alignment = p["fecha"], center

                ws['A4'], ws['B4'] = 'DÍA', 'NOMBRE'
                for i, c in enumerate(conceptos_list): ws.cell(row=4, column=i+3, value=c)
                ws.cell(row=4, column=num_cols, value='TOTAL')
                for col in range(1, num_cols + 1):
                    cell = ws.cell(row=4, column=col)
                    cell.font, cell.fill, cell.alignment, cell.border = header_font, header_fill, center, thin

                ws['A5'].border, ws['B5'].border = thin, thin
                for i, precio in enumerate(precios_list):
                    cell = ws.cell(row=5, column=i+3, value=precio)
                    cell.number_format, cell.alignment, cell.border = currency_fmt, center, thin
                ws.cell(row=5, column=num_cols, value='').border = thin

                start_row = 6
                for idx, fila in enumerate(p["filas"]):
                    ws.cell(row=start_row, column=1, value=fila.get('Dia', '')).border = thin
                    ws.cell(row=start_row, column=2, value=fila.get('Nombre', '')).border = thin
                    for i, c in enumerate(conceptos_list):
                        val = fila.get(c, 0)
                        try: val = int(float(val)) if float(val) > 0 else ""
                        except: val = ""
                        cell = ws.cell(row=start_row, column=i+3, value=val)
                        cell.border, cell.alignment = thin, center
                    total_cell = ws.cell(row=start_row, column=num_cols, value=p["totales"][idx])
                    total_cell.number_format, total_cell.border, total_cell.font = currency_fmt, thin, Font(bold=True)
                    start_row += 1

                ws.merge_cells(f'A{start_row}:B{start_row}')
                ws.cell(row=start_row, column=1, value='TOTAL GENERAL').font, ws.cell(row=start_row, column=1).alignment = bold12, right
                gt = ws.cell(row=start_row, column=num_cols, value=p["total_general"])
                gt.number_format, gt.font, gt.border = currency_fmt, Font(bold=True, size=12, color="FF0000"), thin

                start_row += 2
                ws.merge_cells(f'A{start_row}:{last_col}{start_row}')
                ws.cell(row=start_row, column=1, value=st.session_state.pie_pagina).font, ws.cell(row=start_row, column=1).alignment = Font(italic=True, size=10), center

                ws.column_dimensions['A'].width, ws.column_dimensions['B'].width = 14, 20
                for i in range(3, num_cols + 1): ws.column_dimensions[get_column_letter(i)].width = 20

                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)
                st.download_button(label="📥 Descargar Excel", data=buffer, 
                                  file_name=f"Certificacion_{st.session_state.proyecto_activo.replace(' ', '_')}.xlsx", 
                                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", 
                                  use_container_width=True)
                st.success("✅ ¡Excel generado correctamente!")
    
    with tab_precios:
        st.markdown("###  Guardar y Cargar configuración")
        col_save, col_load = st.columns(2)
        with col_save:
            config_data = {"proyecto": st.session_state.proyecto_activo, "empresa": p["empresa"], 
                          "fecha": p["fecha"], "conceptos": conceptos_validos.to_dict(orient="records")}
            st.download_button(label="⬇️ Descargar Configuración", 
                              data=json.dumps(config_data, indent=2, ensure_ascii=False), 
                              file_name=f"precios_{st.session_state.proyecto_activo.replace(' ', '_')}.json", 
                              mime="application/json", use_container_width=True)
        with col_load:
            uploaded_config = st.file_uploader("⬆️ Cargar Configuración (.json)", type=["json"])
            if uploaded_config is not None:
                try:
                    loaded_data = json.load(uploaded_config)
                    p["empresa"], p["fecha"] = loaded_data.get("empresa", p["empresa"]), loaded_data.get("fecha", p["fecha"])
                    p["conceptos"], p["filas"] = pd.DataFrame(loaded_data.get("conceptos", [])), []
                    st.success("✅ ¡Configuración cargada!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error: {e}")
