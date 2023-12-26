import win32com.client
import os
import re


print("Iniciando la aplicación de Word y abriendo el documento...")

# Inicializar la aplicación de Word y abrir el documento
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = False

file_path = os.path.abspath("documento_generado.docx")
doc = word_app.Documents.Open(file_path)



# Reemplazar "↵" por "^p" en el documento
print("Reemplazando '↵' por '^p' en el documento...")
find_object = doc.Content.Find
find_object.ClearFormatting()
find_object.Text = '^l'  # '^l' es el código para '↵'
find_object.Replacement.ClearFormatting()
find_object.Replacement.Text = '^p'  # '^p' es el código para '¶'
find_object.Execute(Replace=2)  # 2 = wdReplaceAll


# Aplicar las tablas para el documento word
print("Agregando tablas formato markdown en el documento word")

# Lista para almacenar todas las tablas encontradas
all_tables_data = []

# Patrón regex para encontrar enlaces en Markdown
pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')

# Patrón regex para encontrar imágenes en Markdown rodeadas por tuberías
image1_pattern1 = re.compile(r'\|\!\[([^\]]*)\]\((.*?)\)\|')

table = None  # Añade esta línea para inicializar 'table' como None


def add_missing_delimiters(table_lines):
    new_table_lines = []
    for line in table_lines:
        stripped_line = line.rstrip()  # Elimina los espacios en blanco al final
        if stripped_line and not stripped_line.endswith("|"):
            # Si la línea no está vacía y no termina con '|', añade '|'
            stripped_line += " |"
        new_table_lines.append(stripped_line)
    return new_table_lines

while True:
    start_table = None
    end_table = None
    table_lines = []

    # Búsqueda de tablas en el documento
    for index, para in enumerate(doc.Paragraphs):
        line = para.Range.Text.strip()
        placeholders = []  # Añade esta línea para definir 'placeholders'
        
        # Verifica si la línea contiene una imagen en formato Markdown rodeada por tuberías
        if image1_pattern1.search(line):
            continue  # Si es así, ignora la línea y pasa a la siguiente

        # Ignora las líneas que comienzan con una tubería seguida de cualquier número de espacios y luego '+-' o '-+'
        if re.match(r'^\|\s*\+-', line) or re.match(r'^\|\s*-+\+', line):
            continue
        
        # Manejar inicio y fin de tabla
        if "|" in line:
            if start_table is None:
                start_table = index
            blank_line_count = 0  # Reiniciar contador de líneas vacías
            table_lines.append(line.strip())
        elif start_table is not None:
            # Considerar una línea vacía como posible fin de tabla
            if line == '':
                blank_line_count += 1
            if blank_line_count > 1:
                end_table = index
                break

    # Llama a la función para añadir delimitadores faltantes
    table_lines = add_missing_delimiters(table_lines)

    # Si no se encuentra ninguna tabla, se detiene el bucle
    if not table_lines:
        break

    # Procesamiento de la tabla en markdown
    data = []
    headers_found = False
    header_separator_line_index = None  # Índice de la línea de separación del encabezado
    for line in table_lines:
        if not headers_found and "|" in line:
            # Procesar los encabezados
            headers = [cell.strip() for cell in line.split('|')[1:-1]]
            headers = [re.sub(r':?----+:?', '', header).strip() for header in headers]  # Eliminar patrones de alineación
            if headers[-1].strip() == '':
                headers = headers[:-1]
            data.append(headers)
            headers_found = True
        elif headers_found and (re.match(r'^\|\s*-+\s*\|', line) or re.match(r'^\|\s*(:?----+:?\s*\|)+', line)):
            # Ignora la línea de separación del encabezado, no se añade a 'data'
            continue
        elif "|" in line:
            # Procesar las celdas de datos
            cells = [cell.strip() for cell in line.split('|')[1:-1]]
            if cells[-1].strip() == '':
                cells = cells[:-1]
            cells = [cell.replace(':white_check_mark:', '✅') for cell in cells]
            data.append(cells)

    # Eliminar la fila de separación de encabezados si existe
    if header_separator_line_index is not None:
        del data[header_separator_line_index]

    counter = 0
    for para in doc.Paragraphs:
        counter += 1
        if counter == start_table:
            start_range = para.Range.Start
        if counter == end_table:
            end_range = para.Range.End
            break

    doc.Range(start_range, end_range).Delete()

    table_range = doc.Range(start_range , end_range)

    cm_to_points = 2.15 * 28.3465  # 1 cm es aproximadamente 28.3465 puntos
    
    if len(data) > 2 and len(data[0]) > 0:
        table = doc.Tables.Add(table_range, len(data), len(data[0]))
        # Configura el ancho de cada celda de la tabla
        if table is not None:
            for row in table.Rows:
                 for cell in row.Cells:
                      try:
                          cell.Width = cm_to_points
                      except Exception as e:
                              print(f"Error al ajustar el ancho de la celda: {e}")
    else:
          print("The markdown table in this bot is not recognized.")
          continue  # Continúa con el siguiente ciclo del bucle while si la tabla no se reconoce


    markdown_link_found_in_table = False  # Añade esta línea para inicializar 'markdown_link_found_in_table' en False

    for i, row_data in enumerate(data):
        for j, cell_data in enumerate(row_data):
            # Comprueba si los índices están dentro del rango de la tabla
            if i < table.Rows.Count and j < table.Columns.Count:
                cell = table.Cell(i+1, j+1)
                cell_range = cell.Range             

                # Intenta ajustar el formato del párrafo
                cell_range.ParagraphFormat.ListTemplate = None
                # Ajusta el estilo del párrafo a 'Normal'
                cell_range.Style = doc.Styles('Normal')
                
                cell_range.Text = cell_data.strip()
 
                # Manipulación de hipervínculos
                matches = pattern.findall(cell_data)
                
                if matches:
                    markdown_link_found_in_table = True  # Se encontró un enlace Markdown en una celda de la tabla

                    for text, url in matches:
                        hyperlink_range = cell_range.Duplicate
                      
                        # Limpia el texto de la celda antes de añadir el hipervínculo.
                        cell_range.Text = text.strip()
                        
                        
                        hyperlink_range.Find.Execute(FindText=text)
                        doc.Hyperlinks.Add(Anchor=hyperlink_range, Address=url)
                        
                        
            else:
                print(f"Índice fuera de rango: i={i}, j={j}")

    table.Style = "Acc_Table_1"
    all_tables_data.append(data)


# Inicializa 'sheet_resized' como False
sheet_resized = False

# Recorrer las tablas y aumentar el tamaño de la hoja si hay más de 5 columnas
for table in doc.Tables:
    if table.Columns.Count > 5:
        try:
            # Aumentar el tamaño de la hoja en 2.54 cm (equivalente a 1 pulgada)
            points_in_cm = 5.45 * 28.3465  # 1 cm = 28.3465 puntos
            doc.PageSetup.PageWidth = doc.PageSetup.PageWidth + points_in_cm
            doc.PageSetup.PageHeight = doc.PageSetup.PageHeight + points_in_cm
            print(f"The size of the sheet has been increased. New width: {doc.PageSetup.PageWidth}, Nuevo alto: {doc.PageSetup.PageHeight}")
            sheet_resized = True
            break  # Salir del bucle después de la primera tabla encontrada
        except Exception as e:
            print(f"Se produjo un error al intentar ajustar el tamaño de la hoja: {e}")

if sheet_resized:
    print("The document sheet size has been resized.")
else:
    print("The size of the document sheet has not been altered.")


print("etiquetas ya automaticas de colores html")

# Mapea los nombres de los colores a sus correspondientes valores RGB
colores = {
    'azul': win32api.RGB(0, 0, 255),
    'blue': win32api.RGB(0, 0, 255),
    'yellow': win32api.RGB(255, 255, 0),
    'amarillo': win32api.RGB(255, 255, 0),
    'verde': win32api.RGB(0, 128, 0),
    'green': win32api.RGB(0, 128, 0),
    'marrón': win32api.RGB(165, 42, 42),  # Definir el color marrón usando RGB
    'brown': win32api.RGB(165, 42, 42),  # Definir el color marrón usando RGB
    'rosa': win32api.RGB(255, 105, 180),
    'pink': win32api.RGB(255, 105, 180),
    'rojo': win32api.RGB(255, 0, 0),
    'red': win32api.RGB(255, 0, 0),
    'crimson': win32api.RGB(220, 20, 60),  # Color Crimson
    'teal': win32api.RGB(0, 128, 128),     # Color Teal
    'purple': win32api.RGB(128, 0, 128),   # Color Purple
    'violeta': win32api.RGB(128, 0, 128),   # Color Violeta
    'colour': win32api.RGB(0, 0, 0),        # Color Negro para colour
    'black': win32api.RGB(0, 0, 0)
}


def aplicar_formatos(rango, texto):
    # Procesar etiquetas de color
    color_matches = re.findall(r'\(span style="color:(.*?)".*?\)(.*?)\(/span\)', texto, re.IGNORECASE)
    for match in color_matches:
        color_original = match[0]
        color = color_original.lower() 
        content = match[1]
        # Comprueba si el color está en el diccionario de colores
        if color in colores:
            try:
                rango.Font.Color = colores[color]
            except Exception as e:
                print(f"No se pudo cambiar el color: {e}")
            etiqueta_original = f'(span style="color:{color_original}")'
            texto = texto.replace(etiqueta_original, '', 1)
            texto = texto.replace('(/span)', '', 1)
        else:
            print(f"Color desconocido: {color}")

    # Procesar etiquetas de negrita
    bold_matches = re.findall(r'\(b\)(.*?)\(/b\)', texto, re.IGNORECASE)
    for match in bold_matches:
        rango.Font.Bold = constants.wdToggle
        texto = texto.replace('(b)', '').replace('(/b)', '')

    return texto

# Aplicar formatos en párrafos
for i in range(len(doc.Paragraphs), 0, -1):
    paragraph = doc.Paragraphs.Item(i)
    text = paragraph.Range.Text

    # Ignorar si no hay etiquetas HTML o si está en un campo especial
    if not re.search(r'\(span style="color:(.*?)".*?\)(.*?)\(/span\)|\(b\)(.*?)\(/b\)', text, re.IGNORECASE) or paragraph.Range.Fields.Count > 0:
        continue

    # Aplicar formatos y actualizar el texto del párrafo
    new_text = aplicar_formatos(paragraph.Range, text)
    if i <= len(doc.Paragraphs):
        doc.Paragraphs.Item(i).Range.Text = new_text

# Aplicar formatos en tablas
for table in doc.Tables:
    for row in table.Rows:
        for cell in row.Cells:
            cell_text = cell.Range.Text
            # Aplicar formatos y actualizar el texto de la celda
            new_cell_text = aplicar_formatos(cell.Range, cell_text)
            cell.Range.Text = new_cell_text


print("Paragraph Cleaning...")

# Comprobar si el documento tiene al menos 6 páginas
if doc.ComputeStatistics(win32.constants.wdStatisticPages) >= 6:
    # Obtener el rango de la sexta página
    sixth_page_range = word_app.Selection.GoTo(What=win32.constants.wdGoToPage, Which=win32.constants.wdGoToAbsolute, Count=6)
else:
    print("The document is less than 6 pages.")
    sixth_page_range = None

# Recorrer todos los párrafos del documento en orden inverso
for i in range(doc.Paragraphs.Count, 0, -1):
    paragraph = doc.Paragraphs.Item(i)
    
    # Comprobar si el párrafo está en las primeras 5 páginas
    if sixth_page_range and paragraph.Range.Start < sixth_page_range.Start:
        continue  # Si está en las primeras 5 páginas, ignorarlo
    
    # Comprobar si el párrafo es un salto de párrafo
    if paragraph.Range.Text.strip() == "":
        # Si el párrafo anterior también es un salto de párrafo, eliminarlo
        if i > 1:  # Asegurarse de que no es el primer párrafo
            prev_paragraph = doc.Paragraphs.Item(i - 1)
            if prev_paragraph.Range.Text.strip() == "":
                try:
                    prev_paragraph.Range.Delete()
                except Exception as e:
                    print(f"Could not delete paragraph: {e}")

try:
    # Acceder a la tabla de contenido
    table_of_contents = doc.TablesOfContents(1)
except Exception as e:
    print(f"Se produjo un error al intentar acceder a la tabla de contenido: {e}")

try:
    # Actualizar la tabla de contenido
    table_of_contents.Update()
except Exception as e:
    print(f"Se produjo un error al intentar actualizar la tabla de contenido: {e}")

try:
    # Guardar y cerrar el documento
    doc.Save()
    doc.Close()
except Exception as e:
    print(f"Se produjo un error al intentar guardar y cerrar el documento: {e}")

try:
    word_app.Quit()
except Exception as e:
    print(f"Se produjo un error al intentar cerrar la aplicación de Word: {e}")


print("¡finished!")