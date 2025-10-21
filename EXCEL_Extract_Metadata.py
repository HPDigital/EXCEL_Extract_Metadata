"""
EXCEL_Extract_Metadata
"""

#!/usr/bin/env python
# coding: utf-8

# In[ ]:


get_ipython().system('pip install pandas openpyxl')


# In[1]:


import os
import pandas as pd
from openpyxl import load_workbook

def extraer_metadata_de_subcarpetas(ruta_carpeta_madre):
    # Lista para almacenar los datos de cada archivo
    datos = []

    # Iterar a través de todas las subcarpetas y archivos en la carpeta madre
    for root, dirs, files in os.walk(ruta_carpeta_madre):
        for file in files:
            if file.endswith(('.xlsx', '.xls')):
                # Construir la ruta completa del archivo
                file_path = os.path.join(root, file)

                # Intentar leer la propiedad del autor del archivo Excel
                try:
                    wb = load_workbook(file_path, read_only=True, data_only=True)
                    author = wb.properties.author
                except Exception as e:
                    author = "Desconocido"

                # Obtener metadata del archivo
                stat = os.stat(file_path)
                datos.append({
                    "Nombre de la Carpeta": os.path.basename(root),
                    "Nombre del Archivo": file,
                    "Fecha de Creación": stat.st_ctime,
                    "Última Fecha de Modificación": stat.st_mtime,
                    "Tamaño (Bytes)": stat.st_size,
                    "Autor": author
                })

    # Crear un DataFrame con los datos recopilados
    df = pd.DataFrame(datos)

    # Escribir el DataFrame a un archivo Excel
    df.to_excel("metadata_de_subcarpetas.xlsx", index=False)

# Ejecuta la función con la ruta de tu carpeta madre
ruta_carpeta_madre = "C:\\Users\\HP\\Downloads\\Class_4545208_Assignment_45693717"
extraer_metadata_de_subcarpetas(ruta_carpeta_madre)


# In[ ]:






if __name__ == "__main__":
    pass
