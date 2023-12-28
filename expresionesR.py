import re

def modify_regex_pattern(pattern, replacement):
    # Modifica la expresión regular y el string de reemplazo,
    # reemplazando \\ con \, pero maneja adecuadamente las funciones lambda
    modified_pattern = pattern.replace('\\\\', '\\')
    if isinstance(replacement, str):
        modified_replacement = replacement.replace('\\\\', '\\')
    else:
        # Para funciones lambda, el reemplazo no se modifica
        modified_replacement = replacement
    return modified_pattern, modified_replacement


# Tu contenido original
content = "(b)(a color:blue)development/DevOps tools(/b) | ejemplo1 |ejemplo2  |ejemplo 3  |ejemplo4  |ejemplo5  |"

# Lista de expresiones regulares y sus correspondientes reemplazos
regex_replacements = [
    (r'(\(b\)\(a color:(?!blue)[^)]*\)[^|]*\(/b\))\s*\|', r'\\1\\n\\n|'),
    (r'\(/b\)\s*\(/span\)\|', '(/b)|'),
    (r'(<span style="color:([^>]*)">([^<]*?))\\n', r'\\1(/span)\\n'),
    (r'<b><span style="color:([^>]*)">([^<]*)</span></b>', r'(b)(span style="color:\\1")\\2(/span)(/b)'),
    (r'<center>(.*?)</center>', r'\\1'),
    (r'<code>```[^\\n]*\\n', '```\\n'),
    (r'<br>(.*?)</br>', r'\\1'),
    (r'<br>', '\\n'),
    (r'TO_DO: @<([A-Fa-f0-9-]+)>', r'TO_DO: \\1'),
    (r'<Lista> @<([^>]+)>', r'Lista @\\1'),
    (r'\\.(PNG|JPG|JPEG|GIF)', lambda x: x.group().lower()),
    (r'<([^>]*)>', r'(\\1)'),
    (r'\(/code\\)', ''),
    (r'\|\s*(\(span style="color:[^)]+\))\s*([^|]+?)\s*(?=\|)', r'| \\1 \\2(/span)'),
    (r'(\(span style="color:[^)]+\)[^|]*?)\s*\|', r'\\1 (/span)|'),
    (r'(\(/span\))(\s*\(/span\)\s*\|)', r'\\1|'),
    (r'\|\(/span\)', '|'),
    (r'\*\*\s*\(/span\)\s*\|', '** |'),
    (r'\(/b\)\s*\(/span\)\s*\|', '(/b) |'),
]

# Aplica cada expresión regular modificada
for pattern, replacement in regex_replacements:
    modified_pattern, modified_replacement = modify_regex_pattern(pattern, replacement)
    content = re.sub(modified_pattern, modified_replacement, content)

print(content)




