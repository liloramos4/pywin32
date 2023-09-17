import os
import sys
import subprocess

# Crear un entorno virtual llamado "shell"
subprocess.run([sys.executable, "-m", "venv", "shell"])

# Activar el entorno virtual en sistemas basados en Windows
activate_path = os.path.join("shell", "Scripts", "activate")
if sys.platform == "linux":
    activate_path = os.path.join("shell", "bin", "activate")

# Activar el entorno virtual
subprocess.run([activate_path], shell=True)

# Instalar las dependencias en el entorno virtual
subprocess.run([os.path.join("shell", "Scripts", "pip"), "install", "-r", "requirements.txt"])

# Ejecutar el script main.py dentro del entorno virtual
subprocess.run([os.path.join("shell", "Scripts", "python"), "main.py"])

