import streamlit as st
import pandas as pd
from streamlit_localstorage import LocalStorage
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
from PIL import Image

# Configuración básica
st.set_page_config(page_title="Evaluación Asesores GO", page_icon=".streamlit/logo.png", layout="wide")
localS = LocalStorage()

# --- LÓGICA DE PERSISTENCIA OFFLINE ---
def guardar_en_telefono():
    # Guarda todo el estado en el chip de memoria del teléfono
    for key, value in st.session_state.items():
        if not key.startswith("up_"): # Excluimos las fotos por tamaño
            localS.setItem(key, str(value))

def cargar_desde_telefono():
    # Recupera al iniciar
    datos = localS.getItemList()
    if datos:
        for k, v in datos.items():
            st.session_state[k] = v

cargar_desde_telefono()

# --- ESTRUCTURA ---
estructura = {
    "Actitud de Ventas": [("Buena presencia y puntualidad", 9), ("Secuencia de visitas de acuerdo al rutero", 10), ("Asistencia a reuniones, auditorias, contacto 1 a 1", 11), ("Organización y Uso de material de trabajo", 12), ("Toma pedidos y cobranzas en el movil y sincroniza en los tiempos oportunos", 13), ("Informa las incideancias de la ruta al supervisor en los tiempos oportunos", 14)],
    "Pasos de la Visita Organizada": [("Preparación de la visita", 16), ("Contacto y acercamiento", 17), ("Chequeo de Inventario", 18), ("Identificación de oportunidades en el PDV", 19), ("Ejecución del producto en stock de inventario", 20), ("Asesoramiento y negociación del pedido", 21), ("Cobranza y cierre de la visita", 22)],
    "Dominio de la Venta": [("Gestiona todo el portafolio de productos", 24), ("Ofrece innovaciones o promociones", 25), ("Revisión y actualización de indicadores de precios en PDV", 26), ("Negociacion de espacios adicionales", 27)],
    "Material P.O.P y Planogramas": [("Presencia de material P.O.P vigente", 29), ("Evalua el cumplimiento del Planograma de Marcas", 30), ("Garantiza que los exhibidores no esten invadidos", 31), ("Evalua limpieza", 32)],
    "PROFUNDIDAD DE LINEAS": [("Ofrece Marcas Parmalat, PAVECA, Polar, Pepsi", 34), ("Ofrece Marcas Isola, Alfonzo Rivas", 35), ("Ofrece Marcas Alinieve, Incosa, Alesca", 36), ("Ofrece Marcas Miceven, Puig, Gaduca, Calven", 37)]
}

# --- DIÁLOGO DE LIMPIEZA ---
@st.dialog("⚠️ Confirmar Limpieza")
def confirmar_limpieza():
    st.write("¿Estás seguro de borrar todos los datos del teléfono?")
    if st.button("Sí, borrar"):
        localS.clear()
        st.session_state.clear()
        st.rerun()

# --- INTERFAZ ---
st.title("📝 Evaluación Asesores GO (Offline)")
col1, col2, col3 = st.columns(3)
st.date_input("Fecha", key="fecha", on_change=guardar_en_telefono)
st.text_input("Vendedor", key="vendedor", on_change=guardar_en_telefono)
st.text_input("Ruta", key="ruta", on_change=guardar_en_telefono)
st.text_input("Evaluador", key="evaluador", on_change=guardar_en_telefono)

for seg, items in estructura.items():
    st.subheader(f"📍 {seg}")
    for item, fila in items:
        with st.expander(item):
            cols = st.columns(5)
            for i in range(5):
                cols[i].selectbox(f"C{i+1}", [0,1,2,3,4], key=f"{item}_{i}", on_change=guardar_en_telefono)

# --- REPORTE FOTOGRÁFICO ---
st.subheader("📸 Reporte Fotográfico")
nombres_fotos = ["Buena presencia y puntualidad", "Ejecución productos stock", "Presencia material POP", "Cumplimiento Planograma", "Exhibidores no invadidos", "Fachada PDB", "Anaquel principal", "Chequeado", "Exhibición adicional"]
for nombre in nombres_fotos:
    st.file_uploader(f"Foto: {nombre}", type=['png', 'jpg'], key=f"up_{nombre}")

# --- BOTONES ---
if st.button("📥 Exportar a Excel"):
    st.info("Generando archivo...")
    # (Aquí tu lógica de exportación que ya teníamos)

if st.button("🧹 Limpiar"):
    confirmar_limpieza()

