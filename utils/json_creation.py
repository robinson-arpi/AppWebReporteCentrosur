import os
from flask import jsonify
import pandas as pd
import simplejson

def convert_timestamps_to_string(df):
    """Convierte las columnas de tipo Timestamp a cadenas."""
    try:
        for column in df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]', 'timedelta']):
            df[column] = df[column].astype(str)
        return df
    except Exception as e:
        raise Exception(f"Error (convert_timestamps_to_string): {e}")

# Función para guardar el DataFrame como un archivo JSON
def get_df_as_json(sheet_dfs):
    try:
        # Inicializamos una lista para los registros
        all_data = []

        for index, (sheet_name, df) in enumerate(sheet_dfs.items(), start=1):
            # Limpiar y convertir tipos
            df = df.where(pd.notnull(df), None)
            df['hora_inicio'] = df['hora_inicio'].astype(str)
            df['hora_final'] = df['hora_final'].astype(str)
            df['dia'] = df['dia'].astype(str)

            # Agregar las columnas num_sheet y nombre_sheet
            df['numero_sheet'] = index  # Número de la hoja
            df['nombre_sheet'] = sheet_name.strip()  # Eliminar espacios en blanco

            # Convertir el DataFrame a un diccionario y agregarlo directamente a all_data
            all_data.extend(df.to_dict(orient='records'))  # Agregar los registros

        # Devolver el JSON
        return all_data  # Devolver directamente la lista de registros
    
    except Exception as e:
        raise Exception(f"Error (get_df_as_json): {e}")
