import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Color, Border

# Abre el archivo Excel
archivo_excel = load_workbook('C:\\Users\\SrDeLasTinieblas\\Downloads\\Plantilla_Asiento.xlsx')

hoja = archivo_excel.active

# Inicializa listas para almacenar la información
celdas_de_texto = []
celdas_numericas = []
celdas_de_estilo = []
celdas_fusionadas = []

codigo_csharp = ''

# Recorre todas las celdas en la hoja y obtiene las celdas fusionadas
for fila in hoja.iter_rows():
    for celda in fila:
        # Verifica si la celda está fusionada
        if celda.coordinate in hoja.merged_cells:
            celdas_fusionadas.append(celda.coordinate)

        # Obtiene el valor de la celda
        valor = celda.value
        if valor:
            # Verifica si el valor de la celda es un número o una cadena de texto
            if isinstance(valor, (int, float)):
                celdas_numericas.append({
                    "fila": celda.row,
                    "columna": celda.column,
                    "valor": valor
                })
            elif isinstance(valor, str):
                celdas_de_texto.append({
                    "fila": celda.row,
                    "columna": celda.column,
                    "valor": valor
                })

        # Obtiene el estilo de la celda
        estilo = celda.font
        alineacion = celda.alignment
        fondo = celda.fill
        borde = celda.border

        estilo_celda = {
            "fila": celda.row,
            "columna": celda.column,
            "fuente": {
                "nombre": estilo.name,
                "tamaño": estilo.size,
                "color": estilo.color.rgb if estilo.color else None
            },
            "alineacion": {
                "horizontal": alineacion.horizontal,
                "vertical": alineacion.vertical
            },
            "fondo": fondo.fgColor.rgb if fondo.fgColor else None,
            "borde": {
                "top": borde.top.style,
                "left": borde.left.style,
                "right": borde.right.style,
                "bottom": borde.bottom.style
            }
        }
        celdas_de_estilo.append(estilo_celda)

# Función para encontrar grupos continuos de números
def encontrar_grupos_continuos(numeros):
    grupos = []
    grupo_actual = [numeros[0]]

    for i in range(1, len(numeros)):
        if numeros[i] == numeros[i - 1] + 1:
            grupo_actual.append(numeros[i])
        else:
            grupos.append(grupo_actual)
            grupo_actual = [numeros[i]]

    grupos.append(grupo_actual)
    return grupos

# Inicializa un conjunto (set) para almacenar las celdas fusionadas únicas
celdas_fusionadas_set = set()

# Inicializa un diccionario para almacenar las celdas fusionadas agrupadas por letra
celdas_fusionadas_agrupadas = {}

# Recorre todas las celdas en la hoja y obtiene las celdas fusionadas
for fila in hoja.iter_rows():
    for celda in fila:
        # Verifica si la celda está fusionada
        if celda.coordinate in hoja.merged_cells:
                
            #print(f'Celda fusionada: {hoja.merged_cells}')
            # Extrae la letra y número de la coordenada
            letra, numero = celda.coordinate[0], int(celda.coordinate[1:])

            # Agrega la celda fusionada al conjunto
            celdas_fusionadas_set.add(hoja.merged_cells)
            
# Recorre las celdas de texto
for celda_texto in celdas_de_texto:
    fila = celda_texto["fila"]
    columna = celda_texto["columna"]
    valor = celda_texto["valor"]

    # Agrega código para establecer el valor de la celda en C#
    codigo_csharp += f'worksheet.Cells[{fila}, {columna}].Value = "{valor}";\n'

# Recorre las celdas de estilo
for estilo_celda in celdas_de_estilo:
    fila = 11# estilo_celda["fila"]
    columna = estilo_celda["columna"]
    estilo = estilo_celda["fuente"]
    alineacion = estilo_celda["alineacion"]
    fondo = estilo_celda["fondo"]
    borde = estilo_celda["borde"]

    # Convierte el tamaño de fuente a un entero (redondea al entero más cercano)
    tamaño_fuente = int(round(float(estilo["tamaño"])))

    # Agrega código para aplicar estilos en C#
    codigo_csharp += f'using (ExcelRange r = worksheet.Cells[{fila}, {columna}, {fila}, {columna}])\n'
    codigo_csharp += '{\n'

    codigo_csharp += f'    r.Style.Font.SetFromFont(new Font("{estilo["nombre"]}", {tamaño_fuente}));\n'

    # Agrega código para establecer el fondo
    if fondo:
        codigo_csharp += '    r.Style.Fill.PatternType = ExcelFillStyle.Solid;\n'
        codigo_csharp += '    r.Style.Fill.BackgroundColor.SetColor(Color.White);\n'

    # Agrega código para establecer la alineación
    codigo_csharp += '    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;\n'
    codigo_csharp += '    r.Style.VerticalAlignment = ExcelVerticalAlignment.Center;\n'
    #codigo_csharp += f'    r.Style.HorizontalAlignment = ExcelHorizontalAlignment.{alineacion["horizontal"]};\n'
    #codigo_csharp += f'    r.Style.VerticalAlignment = ExcelVerticalAlignment.{alineacion["vertical"]};\n'

    # Agrega código para establecer los bordes
    if borde["top"]:
        codigo_csharp += f'    r.Style.Border.Top.Style = ExcelBorderStyle.{borde["top"].title()};\n'
    if borde["left"]:
        codigo_csharp += f'    r.Style.Border.Left.Style = ExcelBorderStyle.{borde["left"].title()};\n'
    if borde["right"]:
        codigo_csharp += f'    r.Style.Border.Right.Style = ExcelBorderStyle.{borde["right"].title()};\n'
    if borde["bottom"]:
        codigo_csharp += f'    r.Style.Border.Bottom.Style = ExcelBorderStyle.{borde["bottom"].title()};\n'



    codigo_csharp += '}\n'
    codigo_csharp += f'worksheet.Column({columna}).AutoFit();\n\n'

# Función para generar el código C# para fusionar celdas en una columna
#(ExcelRange r = worksheet.Cells["A8:A11"])
def generar_codigo_fusion_columna(fila_inicial, fila_final, columna):
    return f'using (ExcelRange r = worksheet.Cells["{fila_inicial}:{fila_final}"])\n' + \
           '{\n' + \
           '    r.Merge = true;\n' + \
           '}\n'
def column_index_from_string(column_string):
    result = 0
    for char in column_string:
        result = result * 26 + ord(char) - ord('A') + 1
    return result


# Luego, puedes recorrer el diccionario de celdas fusionadas agrupadas por letra
# Recorre todas las celdas fusionadas
for celdas_rango in hoja.merged_cells.ranges:
    coord_inicial = celdas_rango.coord.split(':')[0]
    coord_final = celdas_rango.coord.split(':')[1]
    
    # Agrega código para fusionar celdas en C#
    codigo_csharp += generar_codigo_fusion_columna(coord_inicial, coord_final, coord_inicial[0])

# Ruta del archivo de texto donde deseas guardar el código
ruta_archivo = 'codigo_csharp.txt'  # Reemplaza 'ruta/del/archivo' con la ubicación deseada

# Guarda el código C# en un archivo de texto
with open(ruta_archivo, 'w') as archivo_txt:
    archivo_txt.write(codigo_csharp)
    
archivo_excel.close()
