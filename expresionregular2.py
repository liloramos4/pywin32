import re

def aplicar_primera_regla(content):
    pattern = r'\|\s*(\(span style="color:[^)]+\))\s*([^|]+?)\s*(?=\|)'
    replacement = r'| \1 \2(/span)'
    return re.sub(pattern, replacement, content, flags=re.IGNORECASE)

def aplicar_segunda_regla(content):
    pattern = r'(.*\(span style="color:[^)]+\) s*[^|]+?)\s*( \|\n)'
    replacement = r'\1 (/span)\2'
    return re.sub(pattern, replacement, content)

# Ejemplo de uso
content = """
| ejemplo1 |ejemplo2  |ejemplo 3  |ejemplo4  |ejemplo5  |
|--|--|--|--|--|
|holap.   |(span style="color:rosa") marrón  | (span style="color:marrón") marrón |(span style="color:yellow") sin etiqueta  | (span style="color:green") sin etiqueta |
|(span style="color:violeta")violeta coche  |(span style="color:azul")violeta coche  |(span style="color:verde") OUT of Scope  | (span style="color:red") OUT of Scope |jaja  |
|(span style="color:violeta")violeta coche  |(span style="color:Teal")coche teal  |(span style="color:green") sin etiqueta   |(span style="color:Teal")sin etiqueta| Olé |
| Hola | (span style="color:azul") azul coche | (span style="color:teal")teal coche | Vale | ajjaja |
"""

content = aplicar_primera_regla(content)
content = aplicar_segunda_regla(content)

print(content