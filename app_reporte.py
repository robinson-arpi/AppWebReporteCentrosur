import utils.data_preprocessing as d_p
from utils.json_creation import save_df_as_json
from utils.report_generation import process_data_for_report

# Ruta estática del archivo Excel
file_path = "formato_centrosur.xlsx"
try:
    # Lectura de df desde la ruta estática
    df  = d_p.process_data(file_path)
    # Guardar el DataFrame procesado como JSON
    save_df_as_json(df, 'formato_centrosur.json')
    #Generación de reporte
    process_data_for_report(df)
except Exception as e:
    print(f"Error al cargar los datos: {e}")        
except Exception as e:
    print("Error: " + str(e))
