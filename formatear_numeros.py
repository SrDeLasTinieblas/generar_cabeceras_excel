import openpyxl
import locale
from decimal import Decimal

# Establece la configuración regional para usar comas como separador de miles
locale.setlocale(locale.LC_ALL, '')

# Leer números desde un archivo de texto (reemplaza 'tu_archivo.txt' con el nombre de tu archivo)
with open('numeros.txt', 'r') as archivo:
    lineas = archivo.readlines()

# Crear un nuevo archivo de Excel
libro = openpyxl.Workbook()
hoja = libro.active

# Itera a través de las líneas del archivo y formatea los números
for i, linea in enumerate(lineas):
    linea = linea.strip().replace(',', '')  # Elimina las comas y luego formatea la línea
    try:
        numero = Decimal(linea)  # Convierte la línea a un número decimal
        numero_formateado = locale.format_string('%.2f', numero, grouping=True)  # Formatea el número
    except Decimal.InvalidOperation as e:
        numero_formateado = f"Error: Línea no válida (Línea {i + 1}) - {str(e)}"  # Imprime el número de línea y el mensaje de error
        print(f"Error en la línea {i + 1}: {str(e)}")  # Imprime el número de línea y el mensaje de error

    # Escribe el número formateado en la celda A1, A2, A3, ...
    hoja.cell(row=i + 1, column=1, value=numero_formateado)

# Guarda el archivo de Excel (reemplaza 'resultado.xlsx' con el nombre que desees)
libro.save('resultado.xlsx')

# Cierra el archivo de Excel
libro.close()
