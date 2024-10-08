import pandas as pd
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
import streamlit as st
import io

# Variables globales
sheet_names = []

# Estilos para el reporte
bold_font_title = Font(size=14)  # Establece el tamaño de la fuente a 16
bold_font = Font(bold=True)
highlight = PatternFill("solid", fgColor="FFFF00")
header_fill = PatternFill("solid", fgColor="dbf3d3")
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

def empty_row(fila):
    return fila.isnull().all()

def read_excel(input_file):
    try:
        # Leer todas las hojas de un archivo Excel
        df = pd.read_excel(input_file, sheet_name=None)  # Lee todas las hojas
        valid_rows = []  # Lista para almacenar todas las filas válidas

        # Iterar sobre las hojas
        for sheet_name, data in df.items():
            for index, fila in data.iterrows():
                if empty_row(fila):
                    break
                valid_rows.append(fila)

        # Crear un DataFrame concatenado con todas las filas válidas
        valid_df = pd.DataFrame(valid_rows)
        
        # Obtenención y limpieza de los nombres de las hojas
        sheet_names = list(df.keys())
        valid_df.columns = valid_df.columns.str.strip().str.replace('\n', ' ', regex=True).str.replace('_',' ')

        return valid_df, sheet_names
    except Exception as e:
        st.write("Revise el archvivo subido, asegúrese que contiene los cmapso solicitados" + e)

        return None
    
# Se presentó un problema respecto a la selección de sectores
# En caso de haber problema  se  reporta en la interfaz
def check_sectors(df):
    # Clave para verificar que los sectores sena correctos ['CANTON', 'ZONA', 'NUMERO CLIENTES']
    groupings = df.groupby(['CANTON', 'ZONA', 'NUMERO CLIENTES'])
    corrections = {}
    rows_with_error = []

    for (canton, zona, num_clientes), grupo in groupings:
        sectores = grupo['SECTORES'].tolist()
        if len(set(sectores)) > 1:
            sector_mayor = max(sectores, key=len)
            nuevo_sector = sector_mayor
            corrections[(canton, zona, num_clientes)] = nuevo_sector
            rows_with_error.extend(grupo.index.tolist())

    for (canton, zona, num_clientes), nuevo_sector in corrections.items():
        df.loc[(df['CANTON'] == canton) & (df['ZONA'] == zona) & (df['NUMERO CLIENTES'] == num_clientes), 'SECTORES'] = nuevo_sector

    return df

def combine_hours(group):
    ordered_hours = group[['HORA INICIO', 'HORA FINAL']].sort_values(by='HORA INICIO')
    return ' '.join([f"{row['HORA INICIO']}-{row['HORA FINAL']}" for index, row in ordered_hours.iterrows()])

def create_worksheet(wb, df_agrupado, day):
    ws = wb.create_sheet(title=f"Dia {day.replace('/', '-')}")

    ws[f"A{2}"] = "FORMATO DE  DESCONEXIONES DIARIAS"
    ws[f"A{3}"] = "EMPRESA:"
    ws[f"B{3}"] = "CENTROSUR"
    ws[f"A{4}"] = "FECHA:"
    ws[f"B{4}"] = day.replace('/', '-')

    ws[f"A{2}"].font = bold_font_title
    ws[f"A{3}"].font = bold_font_title
    ws[f"A{4}"].font = bold_font_title
    ws[f"B{3}"].font = bold_font_title
    ws[f"B{4}"].font = bold_font_title

    ws[f"A{3}"].border = thin_border
    ws[f"A{4}"].border = thin_border
    ws[f"B{3}"].border = thin_border
    ws[f"B{4}"].border = thin_border

    ws[f"B{3}"].fill = header_fill
    ws[f"B{4}"].fill = header_fill

    row = 6
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

        merge_start = row + 1
        for _, fila in datos.iterrows():
            ws.append(list(fila))
            row += 1
        merge_end = row

        for r in range(row - len(datos), row + 1):
            for c in range(1, len(fila) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = thin_border

        contador += 1    
        row += 3
        ws.merge_cells(f'A{merge_start}:A{merge_end}')

        # Centrar el contenido después del merge
        merged_cell = ws[f"A{merge_start}"]
        merged_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)


        # Usar la fórmula de Excel para sumar el rango
        ws[f"C{merge_end + 1}"] = "TOTAL:"

        ws[f"D{merge_end + 1}"] = f"=SUM(D{merge_start}:D{merge_end})"  # Escribir la fórmula en D18


    for r in range(2, row):
        ws.cell(row=r, column=4).number_format = '0.00'


    ancho_columnas = {
        'A': 20,
        'B': 16,
        'C': 25,
        'D': 15,
        'E': 15,
        'F': 15,
        'G': 20,
        'H': 15,
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


# Info page
st.set_page_config(
    page_title="Reporte",
    page_icon='images/icono-centrosur.ico',
    layout="centered",
    initial_sidebar_state="expanded",  # Para que la barra lateral esté siempre expandida
)
# Sidebar para instrucciones
logo_url = 'images/logo-centrosur.png'
st.sidebar.image(logo_url)
st.sidebar.header("Instrucciones")
st.sidebar.write("""
Por favor, suba un archivo Excel sin ninguna edición, debería contar con al menos las siguientes cabeceras:
- HORA INICIO
- HORA FINAL
- DIA
- BLOQUE
- SUBESTACIÓN
- PRIMARIOS A DESCONECTAR
- EQUIPO ABRIR
- EQUIPO TRANSF
- CARGA EST MW
- PROVINCIA
- CANTON
- ZONA
- SECTORES
- Prevalencia del Alimentador CTipo de Cliente)
- NUMERO CLIENTES

El programa procesará los datos, separará por días y generará un reporte para descargar.
""")

# Streamlit Interface
st.title("Reporte de cortes de energía Centrosur")

uploaded_file = st.file_uploader("Elige un archivo Excel", type="xlsx")

st.divider()
if uploaded_file:
    df, sheet_names = read_excel(uploaded_file)
    
    # Crear dos columnas
    col1, col2 = st.columns(2)
    with col1:
        st.write("Hojas encontradas:", sheet_names)

    df_seleccionado = df[['SECTORES', 'SUBESTACIÓN', 'CARGA EST MW', 'HORA INICIO', 'HORA FINAL', 'PROVINCIA', 'PRIMARIOS A DESCONECTAR', 'CANTON', 'Prevalencia del Alimentador CTipo de Cliente)', 'NUMERO CLIENTES', 'DIA', 'ZONA']]
    
    df_seleccionado = check_sectors(df_seleccionado)

    df_seleccionado.loc[:, 'DIA'] = pd.to_datetime(df_seleccionado['DIA'], format='%d/%m/%Y')
    # Procesar los datos por día
    df_por_dia = procesar_datos_por_dia(df_seleccionado)

    # Crear un nuevo libro de trabajo
    wb = Workbook()

    for day, df_agrupado in df_por_dia.items():
        df_agrupado = df_agrupado.groupby('SECTORES').agg({
            'HORA INICIO': lambda x: list(x),
            'HORA FINAL': lambda x: list(x),
            'SUBESTACIÓN': 'first',
            'CARGA EST MW': 'first',
            'PROVINCIA': 'first',
            'CANTON': 'first',
            'PRIMARIOS A DESCONECTAR': 'first',
            'Prevalencia del Alimentador CTipo de Cliente)': 'first',
            'NUMERO CLIENTES': 'sum',
            'ZONA': 'first'
        }).reset_index()

        df_agrupado['PERIODO'] = df_agrupado.apply(lambda row: combine_hours(pd.DataFrame({'HORA INICIO': row['HORA INICIO'], 'HORA FINAL': row['HORA FINAL']})), axis=1)
        df_agrupado = df_agrupado.sort_values(by='PERIODO')
        df_agrupado = df_agrupado[['PERIODO', 'SUBESTACIÓN', 'PRIMARIOS A DESCONECTAR', 'CARGA EST MW', 'PROVINCIA', 'CANTON', 'SECTORES', 'Prevalencia del Alimentador CTipo de Cliente)', 'NUMERO CLIENTES']]


        # Crear una hoja por cada día
        create_worksheet(wb, df_agrupado, day)

    # Eliminar la hoja por defecto llamada "Sheet"
    if "Sheet" in wb.sheetnames:
        std_sheet = wb["Sheet"]
        wb.remove(std_sheet)

    # Guardar el archivo en un objeto BytesIO
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)

    with col2:
        st.write("Reporte generado:")
        
        # Botón de descarga
        st.download_button(
            label="Descargar archivo Excel",
            data=output_file,
            file_name='Formato MEM.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
