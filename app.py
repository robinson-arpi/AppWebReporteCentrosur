import streamlit as st
import views.database_chargue as database
import views.report_generation as report
import views.user_guide as guide
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base

# Definir la base de los modelos
Base = declarative_base()
# Crear el motor de la base de datos usando los secretos
engine = create_engine(st.secrets['DB_URL'])

# Función para crear la tabla si no existe
def create_table_if_not_exists():
    Base.metadata.create_all(engine)

# Crear la tabla al iniciar la aplicación
create_table_if_not_exists()

# Info page
st.set_page_config(
    page_title="Automatizaciones",
    layout="centered",
    initial_sidebar_state="expanded",  # Para que la barra lateral esté siempre expandida
)
# Sidebar para instrucciones
logo_url = 'image/logo-centrosur.png'
st.sidebar.image(logo_url)
st.sidebar.header("Menú")

# Inicializar el menú con "Reporte" como opción por defecto
default_menu = 'Cargar datos'
menu = st.sidebar.selectbox(
    'Seleccione una función:',
    ('Cargar datos', 'Reporte histórico'),
    index=0  # Establece el índice inicial en 0 para "Reporte"
)

with st.sidebar.expander("Contenido esperado"):
    st.write("""
            \n-C: hora_inicio
            \n-D: hora_final
            \n-E: dia
            \n-F: bloque
            \n-G: subestacion
            \n-H: primarios_a_desconectar
            \n-I: equipo_abrir
            \n-J: equipo_transf
            \n-K: carga_est_mw
            \n-L: provincia
            \n-M: canton
            \n-N: zona
            \n-O: sectores
            \n-P: prevalencia_del_alimentador
            \n-Q: numero_clientes
            \n-R: clientes_residenciales
            \n-S: aporte_residencial
            \n-T: clientes_comerciales
            \n-U: aporte_comercial
            \n-V: clientes_industriales
            \n-W: aporte_industrial
        """)

# Mostrar la página seleccionada
if menu == 'Cargar datos':
    database.show_page()
elif menu == 'Reporte histórico':
    report.show_page()
