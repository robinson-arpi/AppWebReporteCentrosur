import utils.data_preprocessing as d_p
from utils.json_creation import save_df_as_json

# Ruta estática del archivo Excel
file_path = "formato_centrosur.xlsx"
try:
    # Lectura de df desde la ruta estática
    df, sheet_names, unique_days = d_p.process_data(file_path)
    # Guardar el DataFrame procesado como JSON
    save_df_as_json(df, 'formato_salida.json')
except Exception as e:
    print(f"Error al cargar los datos: {e}")        
except Exception as e:
    print("Error: " + str(e))
