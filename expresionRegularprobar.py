import re

content = """
| Name | IP | Purpose | User | Key |
|--|--|--|--|--|
| showcasepreprodvm | 172.10.32.4 | Jumphost to the AKS cluster | agentuser | [secret]
| AzureDevOpsHostedAgent - (span style="color:red") Out of Scope |
| DNSForwarder - (span style="color:red") Out of Scope |
| FortifyOnDemand - (span style="color:red") Out of Scope |
| York - (span style="color:red") Out of Scope |

<br>
holis
(/code)
"""

# Replace <span> tags with plain text
content = re.sub(r'(<span style="color:([^>]*)">([^<]*?))\n', r'\1(/span)\n', content)
# Reemplazar <center>(.*?)</center> con '(.*?)'
content = re.sub(r'<center>(.*?)</center>', r'\1', content)
# reemplazar ciertos bloques códigos wiki
content = re.sub(r'<code>```[^\n]*\n', '```\n', content)
# Reemplazar <br>(.*?)</br> con '(.*?)'
content = re.sub(r'<br>(.*?)</br>', r'\1', content)
# reemplaza etiqueta <br> con un salto de linea
content = re.sub(r'<br>', '\n', content)
# Reemplazar To_do @<lo que sea> con TO_DO @lo que sea
content = re.sub(r'TO_DO: @<([A-Fa-f0-9-]+)>', r'TO_DO: \1', content)
# Reemplazar <Lista> @<lo que sea> con Lista @lo que sea
content = re.sub(r'<Lista> @<([^>]+)>', r'Lista @\1', content)
# quitar mayusculas formatos de imagen
content = re.sub(r'\.(PNG|JPG|JPEG|GIF)', lambda x: x.group().lower(), content)
content = re.sub(r'<([^>]*)>', r'(\1)', content)
# Reemplazar '(/code)' con ''
content = re.sub(r'\(/code\)', '', content)
# En las tablas etiquetas html
content = re.sub(r'(.*\(span style="color:[^)]+\) s*[^|]+?)\s*( \|\n)', r'\1 (/span)\2', content)


print(content)
