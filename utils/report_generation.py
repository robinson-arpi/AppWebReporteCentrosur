import pandas as pd
from openpyxl import Workbook
import streamlit as st
import io
from datetime import datetime, timedelta

from utils.workbook_creation import create_worksheet

def combine_hours(group):
    try:
        ordered_hours = group[['hora_inicio', 'hora_final']].sort_values(by='hora_inicio')
        return ' '.join([f"{row['hora_inicio']}-{row['hora_final']}" for index, row in ordered_hours.iterrows()])
    except Exception as e:
        st.write(f"Error al combinar los periodos: {e}")
        return None
    
def process_data_for_report(df):
    # Procesar datos por día
    df_by_day = {day.strftime('%Y-%m-%d'): datos for day, datos in df.groupby(df['dia'])}
    
    # Crear un nuevo libro de trabajo
    wb = Workbook()
    
    for day, df_with_gruped_data in df_by_day.items():
        # Forma en la que se vana  trabajr lso datos en la priemr a agrupación
        df_with_gruped_data = df_with_gruped_data.groupby('primarios_a_desconectar').agg({
            'hora_inicio': lambda x: list(x),
            'hora_final': lambda x: list(x),
            'subestacion': 'first',
            'carga_est_mw': 'first',
            'provincia': 'first',
            'canton': 'first',
            'sectores': 'first',
            'prevalencia_del_alimentador': 'first',
            'numero_clientes': 'first',
            'zona': 'first',
            'clientes_residenciales':'first',
            'clientes_industriales': 'first',
            'clientes_comerciales': 'first',
            'aporte_residencial': 'mean',
            'aporte_industrial': 'mean',
            'aporte_comercial': 'mean'
        }).reset_index()
        df_with_gruped_data['periodo'] = df_with_gruped_data.apply(lambda row: combine_hours(pd.DataFrame({'hora_inicio': row['hora_inicio'], 'hora_final': row['hora_final']})), axis=1)
        #df_with_gruped_data = df_with_gruped_data.sort_values(by='periodo')
        df_with_gruped_data = df_with_gruped_data[['periodo', 'subestacion', 'primarios_a_desconectar','clientes_residenciales','clientes_industriales','clientes_comerciales','aporte_residencial', 'aporte_industrial','aporte_comercial', 'provincia', 'canton', 'sectores', 'carga_est_mw']]

        # Crear una hoja por cada día
        create_worksheet(wb, df_with_gruped_data, day)
        # Eliminar la hoja por defecto llamada "Sheet"
        if "Sheet" in wb.sheetnames:
            std_sheet = wb["Sheet"]
            wb.remove(std_sheet)
        
    # Guardar el archivo en un objeto BytesIO
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)

    # Botón de descarga
    st.write("Reporte generado:")
    st.download_button(
        label="Descargar reporte",
        data=output_file,
        file_name='Formato MEM.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )    
