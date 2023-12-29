import re

# Tu contenido original con cadenas multilínea
content = """(b)(a color:blue)development/DevOps tools(/b) (/span)| ejemplo1 |ejemplo2  |ejemplo 3  |ejemplo4  |ejemplo5  | 
<hola-quetal>
holaa
"""

# Aplica las expresiones regulares
content = re.sub(r'\(/b\)\s*\(/span\)\|', '(/b)|', content)
content = re.sub(r'\(/b\)\|', '(/b) \n\n|', content)
content = re.sub(r'(\(b\)\(a color:(?!blue)[^)]*\)[^|]*\(/b\))\s*\|', r'\1\n\n|', content)
content = re.sub(r'(<span style="color:([^>]*)">([^<]*?))\n', r'\\1(/span)\n', content)
content = re.sub(r'<([^>]*)>', r'(\1)', content)
content = re.sub(r'<b><span style="color:([^>]*)">([^<]*)</span></b>', r'(b)(span style="color:\1")\2(/span)(/b)', content, flags=re.IGNORECASE)
content = re.sub(r'<center>(.*?)</center>', r'\1', content)
content = re.sub(r'<code>```[^\n]*\n', '```\n', content)
content = re.sub(r'<br>(.*?)</br>', r'\1', content)
content = re.sub(r'<br>', '\n', content)
content = re.sub(r'TO_DO: @<([A-Fa-f0-9-]+)>', r'TO_DO: \1', content)
content = re.sub(r'<Lista> @<([^>]+)>', r'Lista @\1', content)
content = re.sub(r'\\.(PNG|JPG|JPEG|GIF)', lambda x: x.group().lower(), content)
content = re.sub(r'\(/code\)', '', content)
content = re.sub(r'\|\s*(\(span style="color:[^)]+\))\s*([^|]+?)\s*(?=\|)', r'| \1 \2(/span)', content, flags=re.IGNORECASE)
content = re.sub(r'(\(span style="color:[^)]+\)[^|]*?)\s*\|', r'\1 (/span)|', content)
content = re.sub(r'(\(/span\))(\s*\(/span\)\s*\|)', r'\1|', content)
content = re.sub(r'\|\(/span\)', '|', content) 
content = re.sub(r'\*\*\s*\(/span\)\s*\|', '** |', content) 
content = re.sub(r'\(/b\)\s*\(/span\)\s*\|', '(/b) |', content)

print("\nContenido final:\n", content)
