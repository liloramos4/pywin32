import win32com.client
import win32com.client as win32
import os
import re
import win32api
from win32com.client import constants
from docxtpl import DocxTemplate
from docx import Document
from docx.text.paragraph import Paragraph
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap
import docx
from collections import OrderedDict
import requests
import traceback

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



print("procesa saltos de parrafos")
# Compilar una expresión regular para identificar tus etiquetas HTML específicas y tablas Markdown
html_tag_pattern = re.compile(r'\(b\)\(a color:blue\)development/DevOps tools\(/b\)')
markdown_table_pattern = re.compile(r'\|\s*(.+?)\s*\|')

# Indicador para saber si ya encontramos la etiqueta
found_html_tag = False

# Recorrer todos los párrafos del documento
for paragraph in doc.Paragraphs:
    if found_html_tag:
        break  # Salir del bucle si ya encontramos la etiqueta

    # Buscar tablas Markdown
    if markdown_table_pattern.search(paragraph.Range.Text):
        # Revisar los párrafos anteriores buscando la etiqueta HTML específica
        prev_paragraph = paragraph.Previous()
        while prev_paragraph:
            if html_tag_pattern.search(prev_paragraph.Range.Text):
                # Encontramos la etiqueta HTML deseada cerca de una tabla Markdown
                print("Etiqueta HTML cerca de una tabla Markdown encontrada:", prev_paragraph.Range.Text)
                prev_paragraph.Range.InsertAfter("\n")
                found_html_tag = True  # Indicar que ya encontramos la etiqueta
                break
            prev_paragraph = prev_paragraph.Previous()




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