import re

content = """
| Name | IP | Purpose | User | Key |
|--|--|--|--|--|
| showcasepreprodvm | 172.10.32.4 | Jumphost to the AKS cluster | agentuser | [secret]
| AzureDevOpsHostedAgent - (span style="color:red") Out of Scope |
| DNSForwarder - (span style="color:red") Out of Scope |
| FortifyOnDemand - (span style="color:red") Out of Scope |
| York - (span style="color:red") Out of Scope |
"""

# Expresión regular para buscar la cadena específica y capturar todo lo que precede hasta el último |
content = re.sub(r'(.*\(span style="color:[^)]+\) s*[^|]+?)\s*( \|\n)', r'\1 (/span)\2', content)


print(content)