import re

# Tu contenido original con cadenas multilínea
content = """# Ejemplos de viñetas y sub-viñetas
<span style="color:COLOUR"> This text will be in COLOUR </span>

<span style="color:Crimson"> This text will be in COLOUR </span>

<span style="color:Teal"> This text will be in COLOUR </span>

<b><a color:blue>development/DevOps tools</b>



| ejemplo1 |ejemplo2  |ejemplo 3  |ejemplo4  |ejemplo5  |
|--|--|--|--|--|
|<b><span style="color:blue"> NOTE: </span></b>  |<span style="color:rosa"> marrón  | <span style="color:marrón"> marrón |<span style="color:yellow"> sin etiqueta  | <span style="color:green"> sin etiqueta |
|<span style="color:violeta">violeta coche  |<span style="color:azul">violeta coche  |<span style="color:verde"> OUT of Scope  | <span style="color:red"> OUT of Scope |jaja  |
|<span style="color:violeta">violeta coche  |<span style="color:Teal">coche teal  |<span style="color:green"> sin etiqueta   |<span style="color:Teal">sin etiqueta| Olé |
| Hola | <span style="color:azul"> azul coche | <span style="color:teal">teal coche | Vale | ajjaja |



<b><span style="color:red"> IMPORTANT: </span></b> 

<b><span style="color:blue"> NOTE: </span></b> 

<span style="color:red"> es un ejemplo etiqueta sin cerrar. 

Hola

<span style="color:red"> es un ejemplo 2.  

Hola así es el texto  <span style="color:red"> OUT of Scope
"""

# Aplica las expresiones regulares
content = re.sub(r'<b><span style="color:([^>]*)">([^<]*)</span></b>', r'(b)(span style="color:\1")\2(/span)(/b)', content, flags=re.IGNORECASE)
content = re.sub(r'\*\*\s*\(/span\)\s*\|', '** |', content) 
content = re.sub(r'<center>(.*?)</center>', r'\1', content)
content = re.sub(r'<code>```[^\n]*\n', '```\n', content)
content = re.sub(r'<br>(.*?)</br>', r'\1', content)
content = re.sub(r'(<span style="color:([^>]*)">([^<]*?))\n', r'\1(/span)\n', content)
content = re.sub(r'<br>', '\n', content)
content = re.sub(r'TO_DO: @<([A-Fa-f0-9-]+)>', r'TO_DO: \1', content)
content = re.sub(r'<Lista> @<([^>]+)>', r'Lista @\1', content)
content = re.sub(r'\.(PNG|JPG|JPEG|GIF)', lambda x: x.group().lower(), content)
content = re.sub(r'<([^>]*)>', r'(\1)', content)
content = re.sub(r'\|\s*(\(span style="color:[^)]+\))\s*([^|]+?)\s*(?=\|)', r'| \1 \2(/span)', content, flags=re.IGNORECASE)
content = re.sub(r'(\(span style="color:[^)]+\)[^|]*?)\s*\|', r'\1 (/span)|', content)
content = re.sub(r'\|\(/span\)', '|', content)
content = re.sub(r'(\(/span\))(\s*\(/span\)\s*\|)', r'\1|', content) 
content = re.sub(r'\(/b\)\s*\(/span\)\|', '(/b)|', content)
content = re.sub(r'\(/b\)\|', '(/b) \n\n|', content)
content = re.sub(r'\(/b\)\s*\(/span\)\s*\|', '(/b) |', content)  

print("\nContenido final:\n", content)
