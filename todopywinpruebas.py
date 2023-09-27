import win32com.client
import win32com.client as win32
import os
import re
import time
import win32api

print("Iniciando la aplicación de Word y abriendo el documento...")

# Inicializar la aplicación de Word y abrir el documento
word_app = win32com.client.Dispatch("Word.Application")
word_app.Visible = False

file_path = os.path.abspath("documento_generado.docx")
doc = word_app.Documents.Open(file_path)

# Patrón específico que identifica la estructura que nos interesa
specific_pattern = re.compile(r'(\[.*?\]\(/.*?\))\s*(\d+\.\d+\.\s*.*?$)')

# Compilar las expresiones regulares para los patrones de enlace
link_patterns = [
    re.compile(r'\[.*\]\(/.*\)'),  # Enlace tipo1: System-Information-Leak
    re.compile(r'- \[.*\]\(/.*\)'),  # Enlace tipo2: - System-Information-Leak
    re.compile(r'- https://.*'),  # Enlace tipo3: - https://www.enlaceprueba.com
    re.compile(r'!\[.*\]\(https://.*\)')  # Enlace tipo4: !ejemplo otro tipo de enlace
]

for para in doc.Paragraphs:
    match = specific_pattern.search(para.Range.Text)
    if match:
        # Separar el enlace del título
        para.Range.Text = match.group(1) + '\n' + match.group(2)
        para.Range.Text = match.group(1) + '\n' + match.group(2)
        print(f"Insertando salto de párrafo entre '{match.group(1)}' y '{match.group(2)}'...")
    else:
        # Si no coincide con el patrón específico, busca otros patrones de enlace
        for pattern in link_patterns:
            if pattern.search(para.Range.Text):
                end_pos = para.Range.End
                doc.Range(end_pos, end_pos).InsertParagraphAfter()  # Inserta un salto de párrafo
                doc.Range(end_pos, end_pos).InsertParagraphAfter()  # Inserta un salto de párrafo
                doc.Range(end_pos, end_pos).InsertParagraphAfter()  # Inserta un salto de párrafo
                doc.Range(end_pos, end_pos).InsertParagraphAfter()  # Inserta un salto de párrafo
                doc.Range(end_pos, end_pos).InsertParagraphAfter()  # Inserta un salto de párrafo
                print(f"Insertando salto de párrafo después de '{para.Range.Text.strip()}'...")

# Una segunda pasada para eliminar párrafos vacíos que tienen numeración
for para in doc.Paragraphs:
    if para.Range.Text.strip() == '' and para.Range.ListFormat.ListType != 0:  # Párrafo vacío con numeración
        para.Range.Delete()
        print("Eliminando párrafo numerado vacío...")



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

# Analizar cada párrafo buscando el inicio y el final de una tabla en markdown
while True:
    start_table = None
    end_table = None
    table_lines = []

    # Búsqueda de tablas en el documento
    for index, para in enumerate(doc.Paragraphs):
        line = para.Range.Text
        placeholders = []  # Añade esta línea para definir 'placeholders'
        
        if "|" in line:
            if start_table is None:
                start_table = index
            # Vuelve a insertar los enlaces originales
            for placeholder, match in placeholders:
                line = line.replace(placeholder, match[0])
            table_lines.append(line.strip())
        elif start_table is not None:
            end_table = index
            break

    # Si no se encuentra ninguna tabla, se detiene el bucle
    if not table_lines:
        break

    # Procesamiento de la tabla en markdown
    data = []
    headers_found = False
    for line in table_lines:
        if "|--" in line:
            headers_found = True
            continue
        if not headers_found:
            headers = [cell.strip() for cell in line.split('|')[1:]]
            if headers[-1].strip() == '':
                headers = headers[:-1]
            continue
        else:
            cells = [cell.strip() for cell in line.split('|')[1:]]
            if cells[-1].strip() == '':
                cells = cells[:-1]
            print(f"Celdas procesadas en esta línea: {cells}")
            data.append(cells)

    data.insert(0, headers)

    counter = 0
    for para in doc.Paragraphs:
        counter += 1
        if counter == start_table:
            start_range = para.Range.Start
        if counter == end_table:
            end_range = para.Range.End
            break

    doc.Range(start_range, end_range).Delete()

    table_range = doc.Range(start_range, start_range)
    table = doc.Tables.Add(table_range, len(data), len(data[0]))

    markdown_link_found_in_table = False  # Añade esta línea para inicializar 'markdown_link_found_in_table' en False

    for i, row_data in enumerate(data):
        for j, cell_data in enumerate(row_data):
            cell = table.Cell(i+1, j+1)
            cell_range = cell.Range
            
            # Imprimir información útil antes de cambiar el texto de la celda.
            print(f"Texto original de la celda: {cell_range.Text}")
            
            cell_range.Text = cell_data.strip()
            
            # Imprimir información útil después de cambiar el texto de la celda.
            print(f"Texto de la celda después del cambio: {cell_range.Text}")
            
            # Manipulación de hipervínculos
            matches = pattern.findall(cell_data)
            
            if matches:
                markdown_link_found_in_table = True  # Se encontró un enlace Markdown en una celda de la tabla
                print(f"Encontrado enlace Markdown en la celda (Fila: {i+1}, Columna: {j+1}): {matches}")
                for text, url in matches:
                    hyperlink_range = cell_range.Duplicate
                  
                    # Limpia el texto de la celda antes de añadir el hipervínculo.
                    cell_range.Text = text.strip()
                    
                    # Imprimir información útil antes de buscar el texto del anclaje.
                    print(f"Buscando el texto del anclaje: {text}")
                    
                    hyperlink_range.Find.Execute(FindText=text)
                    doc.Hyperlinks.Add(Anchor=hyperlink_range, Address=url)
                    
                    

                    # Imprimir información útil después de agregar el hipervínculo.
                    print(f"Texto de la celda después de agregar el hipervínculo: {cell_range.Text}")

    table.Style = "Acc_Table_1"
    all_tables_data.append(data)

    # Comprobación si existen enlaces markdown en la tabla
    if markdown_link_found_in_table:
        print("Enlaces Markdown encontrados dentro de la tabla.")
    else:
        print("Enlace Markdown no encontrado dentro de la tabla.")


print("Añadiendo imagenes  para tu documento")

# Compilar la expresión regular para buscar imágenes en formato Markdown
image_pattern = re.compile(r'!\[([^\]]*)\]\(([^)]+)\)')

# Buscar y reemplazar imágenes en formato Markdown con imágenes incrustadas
for paragraph in doc.Paragraphs:
    match = image_pattern.search(paragraph.Range.Text)
    if match:
        print("Imagen encontrada en formato Markdown. Procesando...")
        
        # Obtener la descripción y la URL de la imagen
        description = match.group(1)
        image_url = match.group(2)

        print(f"Descripción de la imagen: {description}")
        print(f"URL de la imagen: {image_url}")

        # Borrar el texto original de la imagen en formato Markdown
        print("Borrando el texto original de la imagen en formato Markdown...")
        paragraph.Range.Delete()

        # Crear un nuevo párrafo con estilo "Normal" para insertar la imagen
        print("Creando un nuevo párrafo con estilo 'Normal' para insertar la imagen...")
        image_paragraph = doc.Paragraphs.Add(paragraph.Range)
        image_paragraph.Format.Style = doc.Styles.Item("Normal")

        # Incrustar la imagen en el documento
        print("Incrustando la imagen en el documento...")
        image_range = image_paragraph.Range
        image_range.ParagraphFormat.Alignment = win32.constants.wdAlignParagraphCenter
        image = image_range.InlineShapes.AddPicture(FileName=image_url, LinkToFile=False, SaveWithDocument=True)

        # Ajustar la altura de la imagen a un máximo de 6 cm (6 cm * 28.3465 puntos/cm = 170.078 puntos)
        max_height = 6 * 28.3465
        if image.Height > max_height:
            print("La imagen es demasiado alta, ajustando a 6 cm...")
            image.Height = max_height

        # Agregar un subtítulo a la imagen si se proporcionó una descripción
        if description:
            print("Agregando descripción a la imagen...")
            
            # Crear un nuevo párrafo con estilo "Normal" para insertar la descripción
            desc_paragraph = doc.Paragraphs.Add(image_range)
            desc_paragraph.Range.Text = f"\n{description}"
            desc_paragraph.Format.Alignment = win32.constants.wdAlignParagraphCenter

# Insertar un salto de página antes del título "Imagen prueba"
print("Insertando un salto de página antes del título 'Imagen prueba'...")
for paragraph in doc.Paragraphs:
        if "Imagen prueba" in paragraph.Range.Text:
                paragraph.Range.InsertBefore('\f')  # '\f' es el código para un salto de página


print("creando bloques de código para tu documento...")
# Variables para detección de bloques de código
found_codeblock = False
in_codeblock = False

# Iteración sobre párrafos
for paragraph in doc.Paragraphs:
    text = paragraph.Range.Text

    # Verificar si hemos encontrado el inicio de un bloque de código
    if '```' in text:
        in_codeblock = not in_codeblock
        found_codeblock = True
        
        # Eliminar el párrafo que contiene "```"
        paragraph.Range.Delete()
        continue  # Saltar el delimitador

    # Si estamos dentro de un bloque de código, cambiar la fuente y el color del texto y del fondo
    if in_codeblock:
        paragraph.Range.Font.Name = "Consolas"
        paragraph.Range.Font.Color = win32api.RGB(255, 255, 255)  # Blanco
        paragraph.Range.Shading.BackgroundPatternColor = win32api.RGB(0, 0, 0)  # Negro

    # Si estamos dentro o fuera del bloque, eliminar los caracteres ``` antes y después.
    if '´´´' in text and not in_codeblock:  
        updated_text = text.replace('´´´', '')
        paragraph.Range.Text = updated_text


# Compilamos las expresiones regulares
markdown_pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')
url_pattern = re.compile(r'-\s(http[s]?://\S+)')
markdown_link_regex = re.compile(r'-\s\[(.+?)\]\((http[s]?://\S+)\)')


print("Buscando enlaces de Markdown y agregando un salto de párrafo después si no hay un título después...")
for i in range(1, doc.Paragraphs.Count + 1):
    paragraph = doc.Paragraphs.Item(i)
    match = markdown_pattern.search(paragraph.Range.Text)
    if match:
        # Comprobar si el siguiente párrafo es un título
        next_paragraph = doc.Paragraphs.Item(i + 1) if i + 1 <= doc.Paragraphs.Count else None
        if next_paragraph and next_paragraph.Style.NameLocal.startswith("Heading"):
            print(f"No se agrega un salto de línea después del enlace de Markdown porque hay un título después: {next_paragraph.Range.Text}")
        else:
            # Agrega un salto de párrafo después del enlace de Markdown
            paragraph.Range.InsertAfter('\r')


# Buscando párrafos que contengan un guión, un espacio luego nombre de la página entre [] seguido (URL)...
print("Buscando párrafos que contengan un guión, un espacio luego nombre de la página entre [] seguido (URL)...")
for paragraph in doc.Paragraphs:
    match = re.search(markdown_link_regex, paragraph.Range.Text)
    if match:
        name = match.group(1)
        url = match.group(2)
        hyperlink_range = paragraph.Range
        doc.Hyperlinks.Add(hyperlink_range, url, TextToDisplay=name)
        print(f"Hipervínculo agregado: {url}")

print("Reemplazando enlaces de Markdown con hipervínculos de Word en el documento...")
for paragraph in doc.Paragraphs:
    match = markdown_pattern.search(paragraph.Range.Text)
    if match:
        name = match.group(1)
        hyperlink_url = ""

        # Comprueba si el texto de anclaje del enlace de Markdown contiene una dirección web
        if re.match(r'https?://.+', match.group(2)):
            hyperlink_url = match.group(2)

        # Si el texto de anclaje del enlace de Markdown contiene una dirección web, crea un hipervínculo de Word
        if hyperlink_url:
            print(f"Encontrado enlace de Markdown: {match.group()}")
            print(f"Texto del enlace: {name}")
            print(f"URL: {hyperlink_url}")

            # Borra el texto original
            print("Borrando el texto original del enlace de Markdown...")
            paragraph.Range.Delete()

            # Añade un hipervínculo al párrafo
            print(f"Añadiendo hipervínculo a '{hyperlink_url}' con el texto '{name}'...")
            hyperlink_range = paragraph.Range
            doc.Hyperlinks.Add(hyperlink_range, hyperlink_url, TextToDisplay=name)



# Buscar párrafos que contengan un guión, un espacio y una URL
print("Buscando párrafos que contengan un guión, un espacio y una URL...")
for paragraph in doc.Paragraphs:
    match = url_pattern.search(paragraph.Range.Text)
    if match:
        url = match.group(1)
        hyperlink_range = paragraph.Range
        # Añadir un bucle de comprobación antes de llamar al método doc.Hyperlinks.Add
        doc.Hyperlinks.Add(hyperlink_range, url)
        print(f"Hipervínculo agregado: {url}")


print("creando listas y viñetas  para tu documento")

for par in doc.Paragraphs:
    paragraph = par.Range.Text    
    if paragraph.startswith('- '):
        par.Format.SpaceAfter = 0
        par.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=par.Range.Application.ListGalleries.Item(1).ListTemplates.Item(1),
            ContinuePreviousList=False,
            ApplyTo=win32.constants.wdListApplyToWholeList,
            DefaultListBehavior=win32.constants.wdWord10ListBehavior)
        par.Range.ListFormat.ListLevelNumber = 1  # Aquí se especifica el nivel de lista
    
    elif paragraph.startswith('  - '):
        par.Format.SpaceAfter = 0
        par.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=par.Range.Application.ListGalleries.Item(1).ListTemplates.Item(1),
            ContinuePreviousList=False,
            ApplyTo=win32.constants.wdListApplyToWholeList,
            DefaultListBehavior=win32.constants.wdWord10ListBehavior)
        par.Range.ListFormat.ListLevelNumber = 2



# Definir el texto objetivo y el texto de reemplazo
target_text_1 = "###"
replacement_text_1 = ""
target_text_2 = "##"
replacement_text_2 = ""
target_text_3 = "_"
replacement_text_3 = ""
target_text_4 = "#"
replacement_text_4 = ""
target_text_5 = "*"
replacement_text_5 = ""
target_text_6 = "**"
replacement_text_6 = ""


# Utilizar el objeto Find para buscar en todo el documento
find_object.ClearFormatting()

# Buscar todos los párrafos que contienen "###" y guardar su texto
found_paragraphs_1 = []
find_object.Text = target_text_1
for paragraph in doc.Paragraphs:
    if target_text_1 in paragraph.Range.Text:
        found_paragraphs_1.append(paragraph.Range.Text)

# Realizar el reemplazo de "###" por " "
if found_paragraphs_1:
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_1
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_1} -> {replacement_text_1}")

# Aplicar el formato de negrita a los párrafos donde se ha eliminado "###"
print("Aplicando formato a los párrafos que contenían '###'...")
for found_text in found_paragraphs_1:
    print(f"Párrafo encontrado: {found_text}")
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_1, replacement_text_1) == paragraph.Range.Text:
            paragraph.Range.Bold = True
            paragraph.Range.Font.Size = 11
            print(f"El párrafo ahora tiene formato Bold y tamaño de letra 11. Texto del párrafo: {paragraph.Range.Text}")
            

# Buscar todos los párrafos que contienen "##" y guardar su texto
print("Buscando '##' en el documento...")
found_paragraphs_2 = []
find_object.Text = target_text_2
for paragraph in doc.Paragraphs:
    if target_text_2 in paragraph.Range.Text:
        found_paragraphs_2.append(paragraph.Range.Text)

# Realizar el reemplazo de "##" por " "
if found_paragraphs_2:
    print("Reemplazando '##' por ' '...")
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_2
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_2} -> {replacement_text_2}")

# Aplicar el formato de negrita a los párrafos donde se ha eliminado "##"
print("Párrafos que contienen '##':")
for found_text in found_paragraphs_2:
    print(found_text)
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_2, replacement_text_2) == paragraph.Range.Text:
            paragraph.Range.Bold = True
            paragraph.Range.Font.Size = 16
            print(f"El párrafo ahora tiene formato Bold y tamaño de letra 16. Texto del párrafo: {paragraph.Range.Text}")
           

# Buscar todos los párrafos que contienen "_" y guardar su texto
print("Buscando '_' en el documento...para ponerla en cursiva")
found_paragraphs_3 = []
find_object.Text = target_text_3
for paragraph in doc.Paragraphs:
    if target_text_3 in paragraph.Range.Text:
        found_paragraphs_3.append(paragraph.Range.Text)

# Realizar el reemplazo de "_" por " "
if found_paragraphs_3:
    print("Reemplazando '_' por ' '...")
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_3
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_3} -> {replacement_text_3}")

# Aplicar el formato de cursiva a los párrafos donde se ha eliminado "_"
print("Párrafos que contienen '_':")
for found_text in found_paragraphs_3:
    print(found_text)
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_3, replacement_text_3) == paragraph.Range.Text:
            paragraph.Range.Italic = True
            print(f"El párrafo ahora tiene formato Italic. Texto del párrafo: {paragraph.Range.Text}")
            
print("aplicando  parrafos que contienen # en negrita en tamaño 16")
# Buscar todos los párrafos que contienen "#" y guardar su texto
found_paragraphs_4 = []
find_object.Text = target_text_4
for paragraph in doc.Paragraphs:
    if target_text_4 in paragraph.Range.Text:
        found_paragraphs_4.append(paragraph.Range.Text)

# Realizar el reemplazo de "#" por " "
if found_paragraphs_4:
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_4
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_4} -> {replacement_text_4}")

# Aplicar el formato de negrita y tamaño 16 a los párrafos donde se ha eliminado "#"
print("Aplicando formato a los párrafos que contenían '#'...")
for found_text in found_paragraphs_4:
    print(f"Párrafo encontrado: {found_text}")
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_4, replacement_text_4) == paragraph.Range.Text:
            paragraph.Range.Bold = True
            paragraph.Range.Font.Size = 16
            print(f"El párrafo ahora tiene formato Bold y tamaño de letra 16. Texto del párrafo: {paragraph.Range.Text}")

print("aplicando  parrafos que contienen ** para poner letra en negrita")
# Buscar todos los párrafos que contienen "**" y guardar su texto
found_paragraphs_6 = []
find_object.Text = target_text_6
for paragraph in doc.Paragraphs:
    if target_text_6 in paragraph.Range.Text:
        found_paragraphs_6.append(paragraph.Range.Text)

# Realizar el reemplazo de "**" por " "
if found_paragraphs_6:
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_6
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_6} -> {replacement_text_6}")

# Aplicar el formato de negrita a los párrafos donde se ha eliminado "**"
print("Aplicando formato a los párrafos que contenían '**'...")
for found_text in found_paragraphs_6:
    print(f"Párrafo encontrado: {found_text}")
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_6, replacement_text_6) == paragraph.Range.Text:
            paragraph.Range.Bold = True
            print(f"El párrafo ahora tiene formato Bold. Texto del párrafo: {paragraph.Range.Text}")

print("aplicando  parrafos que contienen * para poner letra cursiva")
# Buscar todos los párrafos que contienen "*" y guardar su texto
found_paragraphs_5 = []
find_object.Text = target_text_5
for paragraph in doc.Paragraphs:
    if target_text_5 in paragraph.Range.Text:
        found_paragraphs_5.append(paragraph.Range.Text)

# Realizar el reemplazo de "*" por " "
if found_paragraphs_5:
    find_object.Replacement.ClearFormatting()
    find_object.Replacement.Text = replacement_text_5
    find_object.Execute(Replace=2)  # 2 = wdReplaceAll
    print(f"Texto reemplazado: {target_text_5} -> {replacement_text_5}")

# Aplicar el formato de cursiva a los párrafos donde se ha eliminado "*"
print("Aplicando formato a los párrafos que contenían '*'...")
for found_text in found_paragraphs_5:
    print(f"Párrafo encontrado: {found_text}")
    for paragraph in doc.Paragraphs:
        if found_text.replace(target_text_5, replacement_text_5) == paragraph.Range.Text:
            paragraph.Range.Italic = True
            print(f"El párrafo ahora tiene formato Italic. Texto del párrafo: {paragraph.Range.Text}")



# Aplicar las tablas para el documento word
print("Agregando tablas formato markdown en el documento word")

# Lista para almacenar todas las tablas encontradas
all_tables_data = []

# Patrón regex para encontrar enlaces en Markdown
pattern = re.compile(r'\[([^\]]+)\]\(([^)]+)\)')

# Analizar cada párrafo buscando el inicio y el final de una tabla en markdown
while True:
    start_table = None
    end_table = None
    table_lines = []

    # Búsqueda de tablas en el documento
    for index, para in enumerate(doc.Paragraphs):
        line = para.Range.Text
        placeholders = []  # Añade esta línea para definir 'placeholders'
        if pattern.search(line):
           markdown_link_found = True  # Se encontró un enlace Markdown
        
        if "|" in line:
            if start_table is None:
                start_table = index
            # Vuelve a insertar los enlaces originales
            for placeholder, match in placeholders:
                line = line.replace(placeholder, match[0])
            table_lines.append(line.strip())
        elif start_table is not None:
            end_table = index
            break

    # Si no se encuentra ninguna tabla, se detiene el bucle
    if not table_lines:
        break

    # Procesamiento de la tabla en markdown
    data = []
    headers_found = False
    for line in table_lines:
        if "|--" in line:
            headers_found = True
            continue
        if not headers_found:
            headers = [cell.strip() for cell in line.split('|')[1:]]
            if headers[-1].strip() == '':
                headers = headers[:-1]
            continue
        else:
            cells = [cell.strip() for cell in line.split('|')[1:]]
            if cells[-1].strip() == '':
                cells = cells[:-1]
            print(f"Celdas procesadas en esta línea: {cells}")
            data.append(cells)

    data.insert(0, headers)

    counter = 0
    for para in doc.Paragraphs:
        counter += 1
        if counter == start_table:
            start_range = para.Range.Start
        if counter == end_table:
            end_range = para.Range.End
            break

    doc.Range(start_range, end_range).Delete()

    table_range = doc.Range(start_range, start_range)
    table = doc.Tables.Add(table_range, len(data), len(data[0]))

    for i, row_data in enumerate(data):
        for j, cell_data in enumerate(row_data):
            cell = table.Cell(i+1, j+1)
            cell_range = cell.Range
            # Imprimir información útil antes de cambiar el texto de la celda.
            print(f"Texto original de la celda: {cell_range.Text}")
            cell_range.Text = cell_data.strip()
            # Imprimir información útil después de cambiar el texto de la celda.
            print(f"Texto de la celda después del cambio: {cell_range.Text}")
            # Manipulación de hipervínculos
            matches = pattern.findall(cell_data)
            if matches:
                print(f"Encontrado enlace Markdown en la celda (Fila: {i+1}, Columna: {j+1}): {matches}")
                for text, url in matches:
                    hyperlink_range = cell_range.Duplicate
                    # Imprimir información útil antes de buscar el texto del anclaje.
                    print(f"Buscando el texto del anclaje: {text}")
                    hyperlink_range.Find.Execute(FindText=text)
                    doc.Hyperlinks.Add(Anchor=hyperlink_range, Address=url)
                    # Imprimir información útil después de agregar el hipervínculo.
                    print(f"Texto de la celda después de agregar el hipervínculo: {cell_range.Text}")

    table.Style = "Acc_Table_1"
    all_tables_data.append(data)


print("poniendo estilo formato bueno correcto si se ve tiple guión")
# Buscar todos los párrafos que contienen "---"
for paragraph in doc.Paragraphs:
    if "---" in paragraph.Range.Text:
        # Borrar el texto original
        paragraph.Range.Text = "\n"
        
        # Agregar una línea horizontal
        border = paragraph.Range.Borders(win32.constants.wdBorderTop)
        border.LineStyle = win32.constants.wdLineStyleSingle
        border.LineWidth = win32.constants.wdLineWidth050pt

        # Establecer el color de la línea a gris claro
        border.Color = win32api.RGB(234, 234, 234)

        

# Itera sobre todos los campos en el documento
for field in doc.Fields:
    # Comprueba si el campo es un hipervínculo
    if field.Type == win32com.client.constants.wdFieldHyperlink:
        # Comprueba si el campo es parte de la tabla de contenido
        if field.Code.Text.startswith(" TOC "):
            # Si es parte de la tabla de contenido, no cambies el formato
            continue
        # Cambia el color del texto a azul
        field.Result.Font.Color = win32com.client.constants.wdColorBlue
        # Cambia el color del subrayado a azul
        field.Result.Font.UnderlineColor = win32com.client.constants.wdColorBlue
        # Aplica un subrayado simple
        field.Result.Font.Underline = win32com.client.constants.wdUnderlineSingle


# Recorrer todos los párrafos del documento en orden inverso
for i in range(doc.Paragraphs.Count, 0, -1):
    paragraph = doc.Paragraphs.Item(i)
    
    # Comprobar si el párrafo contiene una imagen
    if paragraph.Range.InlineShapes.Count > 0:
        # Si el párrafo anterior es un salto de párrafo, eliminarlo
        if i > 1:  # Asegurarse de que no es el primer párrafo
            prev_paragraph = doc.Paragraphs.Item(i - 1)
            if prev_paragraph.Range.Text.strip() == "":
                prev_paragraph.Range.Delete()



# Acceder a la tabla de contenido
table_of_contents = doc.TablesOfContents(1)

# Actualizar la tabla de contenido
table_of_contents.Update()


# Guardar y cerrar el documento
doc.Save()
doc.Close()
word_app.Quit()


print("¡Finalizado!")