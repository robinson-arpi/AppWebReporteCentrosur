import pandas as pd
import simplejson

def convert_timestamps_to_string(df):
    """Convierte las columnas de tipo Timestamp a cadenas."""
    for column in df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]', 'timedelta']):
        df[column] = df[column].astype(str)
    return df

# Función para guardar el DataFrame como un archivo JSON
def save_df_as_json(sheet_dfs, filename):
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
        
        # Guardar todo el diccionario en un solo archivo JSON
        with open(filename, 'w', encoding='utf-8') as json_file:
            simplejson.dump(all_data, json_file, ensure_ascii=False, indent=4, ignore_nan=True)
        
        print(f"Datos guardados exitosamente en: {filename}")
    except Exception as e:
        print(f"Error al guardar el archivo JSON: {e}")