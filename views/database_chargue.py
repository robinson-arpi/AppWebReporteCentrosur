from sqlalchemy import create_engine
import streamlit as st
import numpy as np
import utils.data_preprocessing as d_p
from utils.db.database_manager import load_data, check_existing_data, get_data_by_specific_dates
from utils.report_generation import process_data_for_report

def show_page():
    st.title("Cargar datos desde excel")
    uploaded_file = st.file_uploader("Elija un archivo Excel", type="xlsx")
    st.divider()
    if uploaded_file:
        try:
            # Lectura de  df
            df, sheet_names, unique_days = d_p.process_data(uploaded_file)
            with st.expander("Previsualización de datos leídos"):
                st.write("Hojas encontradas:")
                for idx, sheet in enumerate(sheet_names, 1):
                    st.write(f"{idx}. {sheet}")       
                st.write(f"Filas para procesar: {df.shape[0]}")
                st.write(df)

            
            # Botón para cargar a la base de datos
            if st.button("Cargar datos"):
                try:
                    #Verificación de que los datos no hayan sido subidos
                    new_data, existing_data = check_existing_data(df)
                    if not existing_data.empty:
                        with st.expander("Se han encontrado registros duplicados"):
                            st.write(existing_data)
                    #Verificación de que se carguen nuevos datos
                    if new_data.empty:
                        st.warning("No se han encontrado datos nuevos para agregar.")
                    else:
                        with st.expander("Registros que fueron agregados"):
                            st.write(new_data)
                        load_data(new_data) 
                        st.success(f"Se han agregado {new_data.shape[0]} entradas a la base de datos.")

                    df = get_data_by_specific_dates(unique_days.astype(str))
                    process_data_for_report(df)

                except Exception as e:
                    st.error(f"Error al cargar los datos: {e}")
            
        except Exception as e:
            st.write("Error en main: " + str(e))
