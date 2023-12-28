import re

content = "| **(span style=\"color:Black\")Scripted Syntax(/span)** (/span)| **(span style=\"color:Black\")Declarative Syntax(/span)** (/span)|"

# Elimina los casos extra de (/span) antes de un separador |
# Elimina (/span) si está justo antes de |
content = re.sub(r'\(/span\)\s*\|', '|', content)


print(content)

