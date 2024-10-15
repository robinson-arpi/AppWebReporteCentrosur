# utils/db/database_manager.py

from sqlalchemy import create_engine, text
from sqlalchemy.ext.declarative import declarative_base
import pandas as pd
import streamlit as st
from sqlalchemy import String, Integer, Float, Date
Base = declarative_base()

# Mapeo de tipos de pandas a tipos de SQLAlchemy
type_map = {
    'int64': Integer,
    'float64': Float,
    'object': String,
    'datetime64[ns]': Date
}

def get_engine():
    """Crea y devuelve un motor de conexión a la base de datos."""
    return create_engine(st.secrets['DB_URL'])

def create_table_if_not_exists(engine):
    """Verifica si la tabla existe y la crea si no."""
    from sqlalchemy import inspect
    inspector = inspect(engine)
    table_name = st.secrets['TABLE_NAME']
    
    # Sentencia SQL para crear la tabla
    create_table_sql = f"""
    CREATE TABLE IF NOT EXISTS `{table_name}` (
        `hora_inicio` TIME NOT NULL,
        `hora_final` TIME NOT NULL,
        `dia` DATE NOT NULL,
        `bloque` VARCHAR(24) NOT NULL,
        `subestacion` INT NOT NULL,
        `primarios_a_desconectar` VARCHAR(255),
        `equipo_abrir` INT,
        `equipo_transf` INT,
        `carga_est_mw` FLOAT NOT NULL,
        `provincia` VARCHAR(255) NOT NULL,
        `canton` VARCHAR(255) NOT NULL,
        `zona` LONGTEXT,
        `sectores` LONGTEXT,
        `prevalencia_del_alimentador` VARCHAR(255),
        `numero_clientes` INT,
        `clientes_residenciales` INT NOT NULL,
        `aporte_residencial` FLOAT NOT NULL,
        `clientes_comerciales` INT NOT NULL,
        `aporte_comercial` FLOAT NOT NULL,
        `clientes_industriales` INT NOT NULL,
        `aporte_industrial` FLOAT NOT NULL,
        PRIMARY KEY (`hora_inicio`, `hora_final`, `dia`, `primarios_a_desconectar`, `provincia`, `canton`)
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    """
    
    if table_name not in inspector.get_table_names():
        try:
            with engine.connect() as connection:
                connection.execute(text(create_table_sql))
            st.success(f"Tabla '{table_name}' creada exitosamente.")
        except Exception as e:
            st.error(f"Error al crear la tabla: {e}")

def load_data(df):
    """Carga un DataFrame en la tabla 'Cortes' en la base de datos."""
    engine = get_engine()
    
    # Crear el diccionario de dtype dinámicamente en base a los tipos de datos del DataFrame
    dtype = {col: type_map.get(str(df[col].dtype), String) for col in df.columns}

    # Crear la tabla si no existe
    create_table_if_not_exists(engine)

    # Cargar los datos en la base de datos
    df.to_sql(st.secrets['TABLE_NAME'], con=engine, if_exists='append', index=False, dtype=dtype)

def check_existing_data(df):
    """Verifica los registros existentes en la base de datos y divide los datos en existentes y nuevos sin duplicados."""
    engine = get_engine()
    create_table_if_not_exists(engine)
    # Inicializamos dos DataFrames vacíos
    existing_data = pd.DataFrame()
    new_data = pd.DataFrame()

    # Iteramos sobre las filas del DataFrame para verificar si cada registro ya existe
    for index, row in df.iterrows():
        query = f"""
        SELECT 1 FROM {st.secrets['TABLE_NAME']}
        WHERE hora_inicio = '{row['hora_inicio']}'
          AND hora_final = '{row['hora_final']}'
          AND dia = '{row['dia']}'
          AND canton = '{row['canton']}'
          AND provincia = '{row['provincia']}'
          AND primarios_a_desconectar = '{row['primarios_a_desconectar']}' 
        LIMIT 1
        """

        # Ejecutamos la consulta para verificar si el registro ya existe
        existing = pd.read_sql(query, con=engine)

        # Si el registro ya existe, lo agregamos a existing_data
        if not existing.empty:
            existing_data = pd.concat([existing_data, df.iloc[[index]]])
        else:
            # Si el registro no existe, lo agregamos a new_data
            new_data = pd.concat([new_data, df.iloc[[index]]])

    # Retornamos los DataFrames de datos nuevos y existentes
    return new_data, existing_data

def get_data_between_days(start_date, end_date):
    """Consulta datos en la tabla 'Cortes' por un rango de fechas y ajusta el formato de las horas."""
    engine = get_engine()
    query = f"""
    SELECT * FROM {st.secrets['TABLE_NAME']} 
    WHERE DIA BETWEEN '{start_date}' AND '{end_date}'
    """
    
    # Ejecutamos la consulta
    df = pd.read_sql(query, con=engine)
    
    # Convertir las columnas 'hora_inicio' y 'hora_final' para que solo muestren las horas, minutos y segundos
    df['hora_inicio'] = df['hora_inicio'].apply(lambda x: str(x).split(' ')[-1])
    df['hora_final'] = df['hora_final'].apply(lambda x: str(x).split(' ')[-1])
    return df

def get_data_by_specific_dates(date_list):
    """Consulta datos en la tabla 'Cortes' para un conjunto específico de fechas."""
    engine = get_engine()
    
    # Convertimos la lista de fechas a un formato compatible para la consulta SQL
    formatted_dates = "', '".join(date_list)
    
    # Creamos la consulta SQL usando la cláusula IN para las fechas específicas
    query = f"""
    SELECT * FROM {st.secrets['TABLE_NAME']} 
    WHERE DIA IN ('{formatted_dates}')
    """
    
    # Ejecutamos la consulta
    df = pd.read_sql(query, con=engine)
    
    # Convertir las columnas 'hora_inicio' y 'hora_final' para que solo muestren horas, minutos y segundos
    df['hora_inicio'] = df['hora_inicio'].apply(lambda x: str(x).split(' ')[-1])
    df['hora_final'] = df['hora_final'].apply(lambda x: str(x).split(' ')[-1])
    
    return df
