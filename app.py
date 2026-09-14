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

st.set_page_config(page_title="Generador de Certificaciones", layout="wide")
st.title("🛠️ Generador de Certificaciones de Fibra")

# --- CONFIGURACIÓN GITHUB PARA EL BOT (Segura) ---
try:
    GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN", "")
except Exception:
    GITHUB_TOKEN = os.environ.get("GITHUB_TOKEN", "")

if not GITHUB_TOKEN:
    st.warning("⚠️ No se ha encontrado GITHUB_TOKEN. La importación desde Telegram no funcionará hasta que lo configures en los Secrets.")

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
        "precios": [12.50, 25.00, 8.00, 3.00, 5.00, 2.50, 2.00]  # 7 conceptos, 7 precios
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
        "precios": [18.00, 36.00, 12.00, 3.10, 7.00, 2.50]  # 6 conceptos, 6 precios
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

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("📁 Proyectos")
    nombres_proyectos = list(st.session_state.proyectos.keys())
    
    idx_actual = nombres_proyectos.index(st.session_state.proyecto_activo) if st.session_state.proyecto_activo in nombres_proyectos else 0
    proyecto_activo = st.selectbox(
        "Proyecto activo",
        nombres_proyectos,
        index=idx_actual
    )
    st.session_state.proyecto_activo = proyecto_activo
    
    st.markdown("---")
    st.subheader("🔧 Gestionar proyecto")
    
    with st.expander("➕ Crear nuevo proyecto"):
        nombre_nuevo = st.text_input("Nombre del nuevo proyecto", "Nuevo Proyecto")
        if st.button("Crear proyecto", use_container_width=True):
            if nombre_nuevo and nombre_nuevo not in st.session_state.proyectos:
                st.session_state.proyectos[nombre_nuevo] = {
                    "empresa": "Mi Empresa",
                    "fecha": "",
                    "conceptos": pd.DataFrame({"Concepto": [""], "Precio": [0.0]}),
                    "filas": []
                }
                st.session_state.proyecto_activo = nombre_nuevo
                st.rerun()
            elif nombre_nuevo in st.session_state.proyectos:
                st.error("Ya existe un proyecto con ese nombre")
    
    with st.expander("✏️ Renombrar proyecto"):
        nuevo_nombre = st.text_input(
            "Nuevo nombre", 
            st.session_state.proyecto_activo,
            key="input_renombrar"
        )
        if st.button
