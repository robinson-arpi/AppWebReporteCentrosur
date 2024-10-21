import pandas as pd
import simplejson

def convert_timestamps_to_string(df):
    """Convierte las columnas de tipo Timestamp a cadenas."""
    for column in df.select_dtypes(include=['datetime64[ns]', 'datetime64[ns, UTC]', 'timedelta']):
        df[column] = df[column].astype(str)  # O usa df[column].dt.strftime('%Y-%m-%d %H:%M:%S')
    return df


# Función para guardar el DataFrame como un archivo JSON
def save_df_as_json(df, filename):
    try:
        df = df.where(pd.notnull(df), None)
        df['hora_inicio'] = df['hora_inicio'].astype(str)
        df['hora_final'] = df['hora_final'].astype(str)
        df['dia'] = df['dia'].astype(str)

        # Convertir el DataFrame a un diccionario y luego a JSON
        data_to_save = df.to_dict(orient='records')
        
        # Guardar como JSON
        with open(filename, 'w', encoding='utf-8') as json_file:
            simplejson.dump(data_to_save, json_file, ensure_ascii=False, indent=4,ignore_nan=True)
        
        print(f"Datos guardados exitosamente en: {filename}")
    except Exception as e:
        print(f"Error al guardar el archivo JSON: {e}")