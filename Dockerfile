FROM cdrx/pyinstaller-windows:python3

# Configurar el directorio de trabajo
WORKDIR /src

# Copiar los archivos necesarios
COPY main.py .
COPY requirements.txt .
COPY resultados.xlsx .

# Instalar las dependencias
RUN pip install -r requirements.txt

# Crear el ejecutable
CMD pyinstaller --onefile --noconsole --add-data "resultados.xlsx;." main.py 