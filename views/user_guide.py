import streamlit as st

def show_page():
    st.title("Guía de usuario")
    with st.expander("Reporte"):
        st.write("""
            Por favor, suba un archivo Excel sin ninguna edición, debería contar con al menos las siguientes cabeceras:
            - HORA INICIO
            - HORA FINAL
            - DIA
            - BLOQUE
            - SUBESTACIÓN
            - PRIMARIOS A DESCONECTAR
            - EQUIPO ABRIR
            - EQUIPO TRANSF
            - CARGA EST MW
            - PROVINCIA
            - CANTON
            - ZONA
            - SECTORES
            - Prevalencia del Alimentador CTipo de Cliente)
            - NUMERO CLIENTES

            El programa procesará los datos, separará por días y generará un reporte para descargar.
            """)
    
    with st.expander("Jessica es mensa?"):
        st.write("Sí")

