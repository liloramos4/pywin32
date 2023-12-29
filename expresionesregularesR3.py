import re

# Tu contenido original con cadenas multilínea
content = """(b)(a color:blue)development/DevOps tools(/b) (/span)| ejemplo1 |ejemplo2  |ejemplo 3  |ejemplo4  |ejemplo5  | 
<hola-quetal>
holaa
"""

# Lista de expresiones regulares y sus correspondientes reemplazos
regex_replacements = [
    (r'\(/b\)\s*\(/span\)\|', '(/b)|'),
    (r'\(/b\)\|', '(/b) \n\n|'),
    (r'(\(b\)\(a color:(?!blue)[^)]*\)[^|]*\(/b\))\s*\|', r'\1\n\n|'),
    (r'(<span style="color:([^>]*)">([^<]*?))\n', r'\1(/span)\n'),
    (r'<([^>]*)>', r'(\1)'),
    (r'<b><span style="color:([^>]*)">([^<]*)</span></b>', r'(b)(span style="color:\1")\2(/span)(/b)', re.IGNORECASE),
    (r'<center>(.*?)</center>', r'\1'),
    (r'<code>```[^\n]*\n', '```\n'),
    (r'<br>(.*?)</br>', r'\1'),
    (r'<br>', '\n'),
    (r'TO_DO: @<([A-Fa-f0-9-]+)>', r'TO_DO: \1'),
    (r'<Lista> @<([^>]+)>', r'Lista @\1'),
    (r'\.(PNG|JPG|JPEG|GIF)', lambda x: x.group().lower()),
    (r'\(/code\)', ''),
    (r'\|\s*(\(span style="color:[^)]+\))\s*([^|]+?)\s*(?=\|)', r'| \1 \2(/span)'),
    (r'(\(span style="color:[^)]+\)[^|]*?)\s*\|', r'\1 (/span)|'),
    (r'(\(/span\))(\s*\(/span\)\s*\|)', r'\1|'),
    (r'\|\(/span\)', '|'),
    (r'\*\*\s*\(/span\)\s*\|', '** |'),
    (r'\(/b\)\s*\(/span\)\s*\|', '(/b) |')
]

# Aplica cada expresión regular
for item in regex_replacements:
    pattern = item[0]
    replacement = item[1]
    flags = item[2] if len(item) > 2 else 0
    content = re.sub(pattern, replacement, content, flags=flags)

# Imprime las expresiones regulares aplicadas
print("Las expresiones regulares aplicadas son:")
for item in regex_replacements:
    pattern = item[0]
    replacement = item[1]
    # Verifica si el reemplazo es una cadena antes de aplicar 'replace'
    if isinstance(replacement, str):
        replacement = replacement.replace('\n', '\\n').replace('\\', '\\\\')
    else:
        # Para funciones lambda, proporciona una descripción legible
        replacement = "<función lambda>"
    print(f"content = re.sub(r'{pattern}', r'{replacement}', content)")

print("\nContenido final:\n", content)
