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
        # Crear un diccionario que contendrá todas las hojas
        all_data = {}

        for sheet_name, df in sheet_dfs.items():
            # Limpiar y convertir tipos
            df = df.where(pd.notnull(df), None)
            df['hora_inicio'] = df['hora_inicio'].astype(str)
            df['hora_final'] = df['hora_final'].astype(str)
            df['dia'] = df['dia'].astype(str)

            # Convertir el DataFrame a un diccionario y luego a JSON
            all_data[sheet_name] = df.to_dict(orient='records')
     
        # Devolver el JSON
        return all_data
    
    except Exception as e:
            raise Exception(f"Error (get_df_as_json): {e}")

        