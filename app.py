import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import streamlit as st
import io

# Variables globales
sheet_names = []



def fila_vacia(fila):
    return fila.isnull().all()

def leer_excel(input_file):
    # Leer todas las hojas de un archivo Excel
    df = pd.read_excel(input_file, sheet_name=None)  # Lee todas las hojas
    filas_validas = []  # Lista para almacenar todas las filas válidas

    # Iterar sobre las hojas
    for sheet_name, data in df.items():
        for index, fila in data.iterrows():
            if fila_vacia(fila):
                break
            filas_validas.append(fila)

    # Crear un DataFrame concatenado con todas las filas válidas
    df_concatenado = pd.DataFrame(filas_validas)
    
    # Obtener los nombres de las hojas
    sheet_names = list(df.keys())

    return df_concatenado, sheet_names

def verificar_sectores(df):
    grupos = df.groupby(['CANTON', 'ZONA', 'NUMERO CLIENTES'])
    correcciones = {}
    filas_con_error = []

    for (canton, zona, num_clientes), grupo in grupos:
        sectores = grupo['SECTORES'].tolist()
        if len(set(sectores)) > 1:
            sector_menor = min(sectores, key=len)
            sector_mayor = max(sectores, key=len)
            nuevo_sector = sector_mayor
            correcciones[(canton, zona, num_clientes)] = nuevo_sector
            filas_con_error.extend(grupo.index.tolist())

    for (canton, zona, num_clientes), nuevo_sector in correcciones.items():
        df.loc[(df['CANTON'] == canton) & (df['ZONA'] == zona) & (df['NUMERO CLIENTES'] == num_clientes), 'SECTORES'] = nuevo_sector

    return df

def combinar_horas(grupo):
    horas_ordenadas = grupo[['HORA_INICIO', 'HORA_FINAL']].sort_values(by='HORA_INICIO')
    return ' '.join([f"{row['HORA_INICIO']}-{row['HORA_FINAL']}" for index, row in horas_ordenadas.iterrows()])

def create_worksheet(wb, df_agrupado, day):
    ws = wb.create_sheet(title=f"Dia {day.replace('/', '-')}")

    # Creación de estilos
    bold_font = Font(bold=True)
    highlight = PatternFill("solid", fgColor="FFFF00")
    header_fill = PatternFill("solid", fgColor="dbf3d3")
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    row = 1
    contador =1
    df_periodos = df_agrupado.groupby('PERIODO')

    for periodo, datos in df_periodos:
        ws[f"A{row}"] = f"PERIODO {contador}"
        ws[f"B{row}"] = "SUBESTACIÓN"
        ws[f"C{row}"] = "PRIMARIOS A DESCONECTAR"
        ws[f"D{row}"] = "CARGA EST MW"
        ws[f"E{row}"] = "PROVINCIA"
        ws[f"F{row}"] = "CANTON"
        ws[f"G{row}"] = "SECTORES"
        ws[f"H{row}"] = "Prevalencia del Alimentador CTipo de Cliente)"
        ws[f"I{row}"] = "NUMERO CLIENTES"
    
        for col in range(1, 10):
            cell = ws.cell(row=row, column=col)
            cell.font = bold_font
            cell.fill = header_fill
            cell.border = thin_border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        ws.row_dimensions[row].height = 30

        if isinstance(datos, pd.Series):
            datos = datos.to_frame().T

        for _, fila in datos.iterrows():
            ws.append(list(fila))
            row += 1
        
        for r in range(row - len(datos), row + 1):
            for c in range(1, len(fila) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border

        contador += 1    
        row += 3

    for r in range(2, row):
        ws.cell(row=r, column=4).number_format = '0.00'

    ancho_columnas = {
        'A': 25,
        'B': 10,
        'C': 25,
        'D': 15,
        'E': 15,
        'F': 15,
        'G': 20,
        'H': 10,
        'I': 15
    }

    for columna, ancho in ancho_columnas.items():
        ws.column_dimensions[columna].width = ancho


def procesar_datos_por_dia(df):
    # Convertir la columna DIA a datetime
    df['DIA'] = pd.to_datetime(df['DIA'], format='%d/%m/%Y', errors='coerce')
    # Agrupar por día con formato seguro para el nombre de la hoja
    df_por_dia = {day.strftime('%Y-%m-%d'): datos for day, datos in df.groupby(df['DIA'])}
    return df_por_dia

# Streamlit Interface
st.title("Reporte de cortes de energía Centrosur")

uploaded_file = st.file_uploader("Elige un archivo Excel", type="xlsx")

if uploaded_file:
    df, sheet_names = leer_excel(uploaded_file)
    st.write("Nombres de las hojas:", sheet_names)
    
    df.columns = df.columns.str.strip().str.replace('\n', ' ', regex=True)

    df_seleccionado = df[['SECTORES', 'SUBESTACIÓN', 'CARGA EST MW', 'HORA_INICIO', 'HORA_FINAL', 'PROVINCIA', 'PRIMARIOS A DESCONECTAR', 'CANTON', 'Prevalencia del Alimentador CTipo de Cliente)', 'NUMERO CLIENTES', 'DIA', 'ZONA']]
    
    df_seleccionado = verificar_sectores(df_seleccionado)

    df_seleccionado.loc[:, 'DIA'] = pd.to_datetime(df_seleccionado['DIA'], format='%d/%m/%Y')
    # Procesar los datos por día
    df_por_dia = procesar_datos_por_dia(df_seleccionado)

    # Crear un nuevo libro de trabajo
    wb = Workbook()

    for day, df_agrupado in df_por_dia.items():
        df_agrupado = df_agrupado.groupby('SECTORES').agg({
            'HORA_INICIO': lambda x: list(x),
            'HORA_FINAL': lambda x: list(x),
            'SUBESTACIÓN': 'first',
            'CARGA EST MW': 'first',
            'PROVINCIA': 'first',
            'CANTON': 'first',
            'PRIMARIOS A DESCONECTAR': 'first',
            'Prevalencia del Alimentador CTipo de Cliente)': 'first',
            'NUMERO CLIENTES': 'sum',
            'ZONA': 'first'
        }).reset_index()

        df_agrupado['PERIODO'] = df_agrupado.apply(lambda row: combinar_horas(pd.DataFrame({'HORA_INICIO': row['HORA_INICIO'], 'HORA_FINAL': row['HORA_FINAL']})), axis=1)
        df_agrupado = df_agrupado.sort_values(by='PERIODO')
        df_agrupado = df_agrupado[['PERIODO', 'SUBESTACIÓN', 'PRIMARIOS A DESCONECTAR', 'CARGA EST MW', 'PROVINCIA', 'CANTON', 'SECTORES', 'Prevalencia del Alimentador CTipo de Cliente)', 'NUMERO CLIENTES']]


        # Crear una hoja por cada día
        create_worksheet(wb, df_agrupado, day)

    # Guardar el archivo en un objeto BytesIO
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)

    st.write("Reporte generado con hojas por día.")
    
    # Botón de descarga
    st.download_button(
        label="Descargar archivo Excel",
        data=output_file,
        file_name='data_horas_agrupadas_por_dia.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
