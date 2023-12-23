import json
import base64
import win32com.client
import win32com.client as win32
import re
import win32api
import urllib.parse  # Importado para analizar la URL
import requests
import os
from docxtpl import DocxTemplate
from docx import Document
from docx.text.paragraph import Paragraph
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap
import docx
import time
from collections import OrderedDict
from win32com.client import constants


print("Starting the Word application and opening the document...")

try:
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False
    word_app.DisplayAlerts = False  # Desactivar alertas automáticas
    file_path = os.path.abspath("documento_generado.docx")
    doc = word_app.Documents.Open(file_path)
except Exception as e:
    print(f"Error al iniciar Word o abrir el documento: {e}")
    exit(1)  # Sale del script si hay un error

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

# Una segunda pasada para eliminar párrafos vacíos que tienen numeración
for para in doc.Paragraphs:
    if para.Range.Text.strip() == '' and para.Range.ListFormat.ListType != 0:  # Párrafo vacío con numeración
        para.Range.Delete()


find_object = doc.Content.Find
find_object.ClearFormatting()
find_object.Text = '^l'  # '^l' es el código para '↵'
find_object.Replacement.ClearFormatting()
find_object.Replacement.Text = '^p'  # '^p' es el código para '¶'
find_object.Execute(Replace=2)  # 2 = wdReplaceAll


def verificar_etiquetas(doc):
    # Pila para rastrear las etiquetas de apertura
    pila_etiquetas = []

    for par in doc.Paragraphs:
        paragraph = par.Range.Text.strip()

        # Si encontramos una etiqueta de apertura, la añadimos a la pila
        if '(ul)' in paragraph or '(li)' in paragraph:
            pila_etiquetas.append(paragraph)

        # Si encontramos una etiqueta de cierre, comprobamos si su correspondiente etiqueta de apertura está en la parte superior de la pila
        elif '(/ul)' in paragraph or '(/li)' in paragraph:
            if pila_etiquetas and pila_etiquetas[-1] == paragraph.replace('/', ''):
                # Si la etiqueta de apertura correspondiente está en la parte superior de la pila, la eliminamos de la pila
                pila_etiquetas.pop()
            else:
                # Si la etiqueta de apertura correspondiente no está en la parte superior de la pila, tenemos una etiqueta suelta
                return True  # Devolver True si se encontró una etiqueta suelta

    # Al final, cualquier etiqueta de apertura que quede en la pila es una etiqueta suelta
    return len(pila_etiquetas) > 0  # Devolver True si se encontraron etiquetas sueltas, False en caso contrario

def eliminar_etiquetas_sueltas(doc):
    # Variables para rastrear etiquetas sueltas
    etiquetas_sueltas = []

    for par in doc.Paragraphs:
        paragraph = par.Range.Text.strip()

        # Agregar etiquetas sueltas a la lista para su posterior eliminación
        if ('(ul)' in paragraph and '(/ul)' not in paragraph) or \
           ('(li)' in paragraph and '(/li)' not in paragraph):
            etiquetas_sueltas.append(par)

    # Comprobar si se encontraron etiquetas sueltas
    if etiquetas_sueltas:
        # Eliminar párrafos con etiquetas sueltas
        for par in etiquetas_sueltas:
            par.Range.Delete()
        return True  # Devolver True si se encontraron etiquetas sueltas

    return False  # Devolver False si no se encontraron etiquetas sueltas

# Proceso principal
print("Verificando etiquetas HTML...")
if verificar_etiquetas(doc):
    print("Se encontraron etiquetas HTML sueltas. Procediendo a eliminarlas...")
    if eliminar_etiquetas_sueltas(doc):
        print("Se eliminaron etiquetas HTML sueltas. Por favor revise en su wiki el orden de las etiquetas HTML emparejadas.")
else:
    print("No se encontraron etiquetas HTML sueltas.")



print("Bullet and sub ul li HTML robot etiquetas solamente emparejadas")
# Variables para rastrear el estado de las listas, sublistas y bloques de código Markdown
en_lista = False
en_sublista = False
en_bloque_codigo = False

# Primero, verificamos si hay etiquetas sueltas
etiquetas_ul_suelta = '(ul)' in doc.Range().Text and '(/ul)' not in doc.Range().Text
etiquetas_li_suelta = '(li)' in doc.Range().Text and '(/li)' not in doc.Range().Text

for par in doc.Paragraphs:
    paragraph = par.Range.Text

    # Verificar si estamos en un bloque de código Markdown
    if '```' in paragraph:
        en_bloque_codigo = not en_bloque_codigo
        continue

    # Omitir la aplicación de formatos dentro de bloques de código
    if en_bloque_codigo:
        continue

    # Manejar etiquetas sueltas antes de procesar las listas y sublistas
    if etiquetas_ul_suelta and '(ul)' in paragraph:
        par.Range.Delete()
        continue
    if etiquetas_li_suelta and '(li)' in paragraph:
        par.Range.Delete()
        continue

    # Verificar y actualizar el estado de las listas y sublistas
    if '(ul)' in paragraph:
        en_lista = True
        par.Range.Delete()
        continue
    elif '(/ul)' in paragraph:
        en_lista = False
        par.Range.Delete()
        continue
    if '(li)' in paragraph:
        en_sublista = True
        par.Range.Delete()
        continue
    elif '(/li)' in paragraph:
        en_sublista = False
        par.Range.Delete()
        continue

    # Aplicar formato de lista o sublista según corresponda, excepto en bloques de código
    if en_lista or paragraph.startswith('- '):
        par.Format.SpaceAfter = 0
        par.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=par.Range.Application.ListGalleries.Item(1).ListTemplates.Item(1),
            ContinuePreviousList=True,
            ApplyTo=win32.constants.wdListApplyToWholeList,
            DefaultListBehavior=win32.constants.wdWord10ListBehavior)
        par.Range.ListFormat.ListLevelNumber = 1

    if en_sublista or paragraph.startswith('  - '):
        par.Format.SpaceAfter = 0
        par.Range.ListFormat.ApplyListTemplateWithLevel(
            ListTemplate=par.Range.Application.ListGalleries.Item(1).ListTemplates.Item(1),
            ContinuePreviousList=True,
            ApplyTo=win32.constants.wdListApplyToWholeList,
            DefaultListBehavior=win32.constants.wdWord10ListBehavior)
        par.Range.ListFormat.ListLevelNumber = 2


print("Paragraph Cleaning...")

# Comprobar si el documento tiene al menos 6 páginas
if doc.ComputeStatistics(win32.constants.wdStatisticPages) >= 6:
    # Obtener el rango de la sexta página
    sixth_page_range = word_app.Selection.GoTo(What=win32.constants.wdGoToPage, Which=win32.constants.wdGoToAbsolute, Count=6)
else:
    print("El documento tiene menos de 6 páginas.")
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
                    print(f"No se pudo eliminar el párrafo: {e}")

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



print("imprime texto de los párrafos  indica si es una lista o sub-lista")
# Iterar sobre cada párrafo en el documento
for para in doc.Paragraphs:
    # Comprobar si el párrafo es un elemento de lista
    if para.Range.ListFormat.ListType == win32.constants.wdListBullet:
        if para.Range.ListFormat.ListLevelNumber == 1:
            print("Bullet:", para.Range.Text.strip())
        elif para.Range.ListFormat.ListLevelNumber == 2:
            print("Sub-bullet:", para.Range.Text.strip())
    else:
        print("Normal Paragraph:", para.Range.Text.strip())



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

