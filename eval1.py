import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from io import BytesIO
from PIL import Image

st.set_page_config(
    page_title="Evaluación Asesores GO", 
    page_icon=".streamlit/logo.PNG", 
    layout="wide"
)
st.subheader("Evaluacion de Asesores 2026")

# --- ESTRUCTURA ---
estructura = {
    "Actitud de Ventas": [("Buena presencia y puntualidad", 9), ("Secuencia de visitas de acuerdo al rutero", 10), ("Asistencia a reuniones, auditorias, contacto 1 a 1", 11), ("Organización y Uso de material de trabajo", 12), ("Toma pedidos y cobranzas en el movil y sincroniza en los tiempos oportunos", 13), ("Informa las incideancias de la ruta al supervisor en los tiempos oportunos", 14)],
    "Pasos de la Visita Organizada": [("Preparación de la visita (Planificacion mensual, semanal, diaria)", 16), ("Contacto y acercamiento", 17), ("Chequeo de Inventario", 18), ("Identificación de oportunidades en el PDV", 19), ("Ejecución del producto en stock de inventario (Merchandising)", 20), ("Asesoramiento y negociación del pedido", 21), ("Cobranza y cierre de la visita (Indicacion de Politicas)", 22)],
    "Dominio de la Venta": [("Gestiona todo el portafolio de productos", 24), ("Ofrece innovaciones o promociones", 25), ("Revisión y actualización de indicadores de precios en PDV", 26), ("Negociacion de espacios adicionales", 27)],
    "Material P.O.P y Planogramas": [("Presencia de material P.O.P vigente", 29), ("Evalua el cumplimiento del Planograma de Marcas", 30), ("Garantiza que los exhibidores o espacios negociados no esten invadidos", 31), ("Evalua limpieza de los exhibidores o espacios negociados", 32)],
    "PROFUNDIDAD DE LINEAS": [("Ofrece Marcas Parmalat, PAVECA, Polar, Pepsi", 34), ("Ofrece Marcas Isola, Alfonzo Rivas", 35), ("Ofrece Marcas Alinieve/Giomar, Incosa, Alesca", 36), ("Ofrece Marcas Miceven, Puig, Gaduca, Calven", 37)]
}

# --- INTERFAZ ---
col1, col2, col3 = st.columns(3)
fecha = col1.date_input("Fecha", key="fecha")
vendedor = col2.text_input("Vendedor", key="vendedor")
ruta = col3.text_input("Ruta", key="ruta")
evaluador = st.text_input("Evaluador", key="evaluador")

for seg, items in estructura.items():
    st.subheader(f"📍 {seg}")
    for item, fila in items:
        with st.expander(item):
            cols = st.columns(5)
            for i in range(5):
                cols[i].selectbox(f"C{i+1}", [0, 1, 2, 3, 4], key=f"{item}_{i}")

fortalezas = st.text_area("Fortalezas", key="fortalezas")
debilidades = st.text_area("Debilidades", key="debilidades")
acciones = st.text_area("Acciones y acuerdos hasta el próximo contacto", key="acciones")

# Sección de fotos
st.subheader("📸 Registro Fotografico (Hoja 2)")
nombres_fotos = [
    "Buena presencia y puntualidad", "Ejecución productos Merchandising", "Presencia material POP", 
    "Cumplimiento Planograma", "Exhibidores no invadidos", "Fachada PDV", 
    "Anaquel principal", "Checkout", "Exhibición adicional"
]
mapeo_fotos = {}
for nombre in nombres_fotos:
    mapeo_fotos[nombre] = st.file_uploader(f"Foto: {nombre}", type=['jpg', 'jpeg', 'png'], key=f"up_{nombre}")

# --- LÓGICA DE EXPORTACIÓN ---
def exportar_excel():
    wb = load_workbook('Eval1.xlsx')
    ws_main = wb.active 
    ws_fotos = wb['Registro Fotografico']
    
    # Función inteligente para celdas combinadas
    def escribir_celda(ws, ref, valor):
        for merged_range in ws.merged_cells.ranges:
            if ref in merged_range:
                ws.cell(row=merged_range.min_row, column=merged_range.min_col).value = valor
                return
        ws[ref] = valor

    # Escribir textos en hoja principal
    escribir_celda(ws_main, 'E4', str(fecha))
    escribir_celda(ws_main, 'E6', vendedor)
    escribir_celda(ws_main, 'H6', ruta)
    escribir_celda(ws_main, 'E7', evaluador)
    escribir_celda(ws_main, 'N15', fortalezas)
    escribir_celda(ws_main, 'N22', debilidades)
    escribir_celda(ws_main, 'N30', acciones)
    
    columnas = ['H', 'I', 'J', 'K', 'L']
    for seg, items in estructura.items():
        for item, fila in items:
            for i in range(5):
                escribir_celda(ws_main, f"{columnas[i]}{fila}", st.session_state[f"{item}_{i}"])
    
    # Insertar fotos en hoja Registro Fotografico
    celdas = {
        'Buena presencia y puntualidad': 'D2', 'Ejecución productos Merchandising': 'D3', 'Presencia material POP': 'D4', 
        'Cumplimiento Planograma': 'D5', 'Exhibidores no invadidos': 'D6', 'Fachada PDV': 'D7', 
        'Anaquel principal': 'D8', 'Checkout': 'D9', 'Exhibición adicional': 'D10'
    }
    
    ws_fotos.column_dimensions['D'].width = 30 # Ajuste ancho
    for nombre, celda in celdas.items():
        if mapeo_fotos[nombre]:
            # Ajuste de altura de fila para que la foto quepa bien
            fila_idx = int(celda[1:])
            ws_fotos.row_dimensions[fila_idx].height = 80
            
            img = XLImage(mapeo_fotos[nombre])
            img.width = 150
            img.height = 100
            ws_fotos.add_image(img, celda)
    
    output = BytesIO()
    wb.save(output)
    return output.getvalue()

# --- BOTONES ---
st.download_button(
    label="📥 Exportar a Excel con Fotos",
    data=exportar_excel(),
    file_name=f"Reporte_{vendedor}_{fecha}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

if st.button("🧹 Limpiar y Guardar Limpio"):
    st.session_state.clear()
    st.rerun()
