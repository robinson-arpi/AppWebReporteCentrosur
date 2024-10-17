from datetime import datetime
import streamlit as st

from utils.db.database_manager import get_data_between_days
from utils.report_generation import process_data_for_report

def show_page():
    st.title("Historial de cortes")

    # Título de la aplicación
    st.subheader("Rango de fechas para reporte")
    # Crear dos columnas
    col1, col2 = st.columns(2)

    # Agregar contenido en la primera columna
    with col1:
        # Selección de fecha de inicio
        start_date = st.date_input("Fecha de inicio", datetime.now())
    with col2:
        # Selección de fecha de fin
        end_date = st.date_input("Fecha de fin", datetime.now())

    # Comprobar si el rango de fechas es válido
    if end_date < start_date:
        st.error("La fecha de fin debe ser mayor o igual a la fecha de inicio.", icon="🚨")

    # Crear el botón para generar el reporte
    if st.button("Generar"):
        # Mostrar las fechas seleccionadas
        st.write("Fecha de inicio:", start_date)
        st.write("Fecha de fin:", end_date)

        # Obtener los datos entre las fechas seleccionadas
        df = get_data_between_days(start_date, end_date)
        
        # Procesar los datos para el reporte
        process_data_for_report(df)