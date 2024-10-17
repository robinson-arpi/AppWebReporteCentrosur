import streamlit as st
import utils.data_preprocessing as d_p

def show_page():
    st.title("Actualización de datos")
    uploaded_file = st.file_uploader("Elija un archivo Excel", type="xlsx")
    
    st.divider()
    if uploaded_file:
        try:
            # Lectura de  df
            df, sheet_names, unique_days = d_p.process_data(uploaded_file)
        except Exception as ex:
            pass