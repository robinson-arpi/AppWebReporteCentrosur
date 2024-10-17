# Imagen base de Python
FROM python:3.11-slim

# Establecer el directorio de trabajo
WORKDIR /app

# Copiar archivo de requerimientos a la imagen
COPY requirements.txt .

# Instalación de las dependencias desde requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

# Copiar código fuente al contenedor
COPY . .

# Exposición del puerto en el que corre Streamlit
EXPOSE 8501

# Comando que se ejecutará al iniciar el contenedor
CMD ["streamlit", "run", "app.py"]