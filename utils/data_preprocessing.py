import pandas as pd
import re


# Crear un diccionario para asignar nombres de columnas
cells_to_read = ['C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W']
limits_to_read = "C:W"
column_names = {
    'C': 'hora_inicio',
    'D': 'hora_final',
    'E': 'dia',
    'F': 'bloque',
    'G': 'subestacion',
    'H': 'primarios_a_desconectar',
    'I': 'equipo_abrir',
    'J': 'equipo_transf',
    'K': 'carga_est_mw',
    'L': 'provincia',
    'M': 'canton',
    'N': 'zona',
    'O': 'sectores',
    'P': 'prevalencia_del_alimentador',
    'Q': 'numero_clientes',
    'R': 'clientes_residenciales',
    'S': 'aporte_residencial',
    'T': 'clientes_comerciales',
    'U': 'aporte_comercial',
    'V': 'clientes_industriales',
    'W': 'aporte_industrial',
}

def clean_cell(cell):
    """Función para limpiar caracteres no válidos de una celda."""
    if isinstance(cell, str):
        return cell.replace('\n', ' ').replace('\r', ' ').strip()
    return cell

def read_excel_to_df(input_path):
    """Lee un archivo Excel y lo convierte en un DataFrame, limpiando datos en el proceso."""
    try:
        all_sheets = pd.read_excel(input_path, sheet_name=None, usecols=limits_to_read, skiprows=0)
        df_list = []

        for sheet_name, df in all_sheets.items():
            # Definir cabecera
            df.columns = list(cells_to_read)
            
            # Renombrar las  columnas
            df.columns = [column_names[col] for col in df.columns]

            # Limpiar los datos en cada celda
            for column in df.columns:
                df[column] = df[column].map(clean_cell)

            # # Convertir hora_inicio y hora_final a formato de hora
            # df['hora_inicio'] = pd.to_datetime(df['hora_inicio'], format='%H:%M:%S', errors='coerce').dt.time
            # df['hora_final'] = pd.to_datetime(df['hora_final'], format='%H:%M:%S', errors='coerce').dt.time

            # # Convertir dia a formato de fecha
            # df['dia'] = pd.to_datetime(df['dia'], format='%Y-%m-%d', errors='coerce').dt.date

            # Filtrar filas que tienen más de 5 NaN o están vacías
            df = df[df.isnull().sum(axis=1) <= 6]

            # Aquí aseguramos que se eliminen las filas completamente vacías
            df = df[~df.apply(lambda x: x.astype(str).str.strip().eq('').all(), axis=1)]

            # Agregar el DataFrame limpio a la lista si no está vacío
            if not df.empty:
                df_list.append(df)

        # Concatenar todos los DataFrames de las hojas en uno solo
        concatenated_df = pd.concat(df_list, ignore_index=True)

        # Retornar el DataFrame concatenado y los nombres de las hojas
        sheet_names = list(all_sheets.keys())
        return concatenated_df, sheet_names
    except ValueError as e:
        print(f'Error (faltan columnas en su archivo): {e}')
        return None, None
    
    except Exception as e:
        print(f'Error al leer el archivo: {e}')
        return None, None

def process_data(input_file):
    df, sheet_names = read_excel_to_df(input_file)
    unique_days = df['dia'].unique()  # Obtiene los días únicos como un vector
    return df, sheet_names, unique_days
