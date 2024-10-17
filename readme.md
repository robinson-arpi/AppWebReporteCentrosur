# Automatización de reportes Centrosur

### Desarrollo

* VSCode
* Python 3.12.3

### Ejecución

Crear una base de datos, para acceder a esta, desde `.streamlit/secrets` colocar las credenciales:

`DB_URL = 'url_db'`
`TABLE_NAME = 'nombre_table'`

> También existe un docker-compose en caso de querer ejecutar la base como un contenedor
>
> ```
> docker-compose up -d
> ```

Situese en la carpeta del proyecto

> Instale requirements.txt
>
> ```python
> pip install -r requirements.txt
> ```

> Ejecute la aplicación en streamlit
>
> ```python
> streamlit run app.py
> ```

#### GUI

![1729015234642](image/readme/1729015234642.png)

## Estructura

> El proyecto esta organizado de la siguiente manera (se incluyen tmabién los archivos que no se  están guardando en el repositorio debido al gitignore):
>
> ```
>
> └── 📁.streamlit
>     └── config.toml
>     └── secrets.toml
> └── 📁image
>     └── 📁readme
>     └── logo-centrosur.png
> └── 📁mysql_data
> └── 📁resources
>     └── database.sql
> └── 📁utils
>     └── 📁db
>         └── database_manager.py
>     └── data_preprocessing.py
>     └── report_generation.py
>     └── workbook_creation.py
> └── 📁views
>     └── database_chargue.py
>     └── report_generation.py
>     └── user_guide.py
>     └── .gitignore
> └── app.py
> └── docker-compose.yml
> └── readme.md
> └── requirements.txt
>
> ```

#### .streamlit

- **config.toml**: configuración estética
- **secrets.toml:** variables de entorno para desarrollo

#### image

* **readme:** imágenes para estilizar el readme
* **logo-centrosur.png**: logo para app

#### mysql_data

Carpeta generada por el docker-compose para el volumen de la base de datos.

#### reports

Carpeta para almacenar los reportes generados durante las pruebas.

#### resources

Carpeta con el sql para una visualización más adecuada de la tabal generada.

#### utils

* **db:** gestión de funciones para conexión a base de datos

  * **database_manager:**
* **data_preparation:**
* 
* **report_generation:**

  * **`combine_hours`** :

    * Esta función toma un grupo de datos con horas de inicio y final y las combina en un solo texto en formato `HH:MM:SS-HH:MM:SS`. Asegura que los periodos estén ordenados antes de concatenar.
    * Si ocurre un error, se captura y muestra un mensaje en la interfaz de Streamlit.
  * **`process_data_for_report`** :

    * Esta función procesa los datos agrupándolos por día y luego por `primarios_a_desconectar`.
    * Utiliza la función `combine_hours` para combinar los periodos de horas (`hora_inicio` y `hora_final`) en una nueva columna llamada `periodo`.
    * Después, el DataFrame es reordenado y seleccionado solo con las columnas necesarias para el reporte.
    * Luego, la función `create_worksheet` se encarga de generar las hojas de Excel correspondientes a cada día.
    * El archivo de Excel se guarda en un objeto `BytesIO`, lo que permite descargarlo como archivo desde Streamlit.
    * Se agrega un botón de descarga para que el usuario pueda obtener el reporte generado en formato `.xlsx`.
* **workbook_creation:**

  * **`calculate_hours:`**
    * Calcula el número total de horas entre dos periodos de tiempo en formato "HH:MM", considerando si el periodo pasa de un día a otro.
  * **`create_worksheet: `**
    * Crea una nueva hoja en un archivo de Excel y llena las celdas con información formateada.
    * Añade títulos, bordes, encabezados y detalles de cada bloque de datos.
    * Aplica los estilos definidos, como el formato en negrita y los rellenos de color.
    * Para cada periodo de desconexión, agrega los datos relacionados con subestaciones, primarios a desconectar, número de clientes, demanda promedio, etc.
    * Realiza cálculos de demanda usando fórmulas de Excel y agrupa los resultados por periodo, sumando los totales.

#### views

Contiene las 2 vistas generadas, desde aquí se llama a la carpeta de `utils` que contiene las funciones para realizar todos los procesos.

* database_chargue_view:
* report_generation_view:
* update_view:



## Despliegue

El despleigue se relaizó en los servidores de streamlit. La base de datos se ejecuta en un servidor propio.

En la página [https://share.streamlit.io/](Streamlit Cloud) vinculamos el repositorio y configuramos el repositorio.

![1729195156031](image/readme/1729195156031.png)


En advanced settings se puede configurar la versión de python y es necesario agregar los secretos:

![1729195258910](image/readme/1729195258910.png)

Una vez realizado eso solo damos click a Deploy y con eso finaliza el despliegue.

---



<table align="center">
  <tr>
    <td width="300px">
      <img src="https://user-images.githubusercontent.com/128093285/266722309-901daa36-a94c-4269-a195-88604c4454d0.gif" />
    </td>
    <td width="300px">
      <h3>Robinson Arpi</h3>
      <p>Computer Science Engineer | Full Stack Developer | Data Analyst</p>
      <h3>Contact Me</h3>
      <a href="https://www.linkedin.com/in/robinson-arpi-ayala-b258821b0">
        <img src="https://img.shields.io/badge/linkedin-%230077B5.svg?style=for-the-badge&logo=linkedin&logoColor=white" alt="LinkedIn" />
      </a>
      <a href="https://wa.me/593998320642" target="_blank">
        <img src="https://img.shields.io/badge/WhatsApp-25D366?style=for-the-badge&logo=whatsapp&logoColor=white" alt="WhatsApp" />
      </a>
      <a href="mailto:robinson.arpi@gmail.com">
        <img src="https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white" alt="GMail" />
      </a>
  </tr>
</table>
