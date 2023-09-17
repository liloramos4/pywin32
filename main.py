import runpy

# Preguntar al usuario si quiere descargar todas las páginas y subpáginas de la Wiki
respuesta = input("¿Quieres descargar todas las páginas y subpáginas de la Wiki? (s/n): ")

if respuesta.lower() == 's':
    # Ejecutar el archivo todaslaspaginas.py
    runpy.run_path('todaslaspaginas.py')
else:
    # Ejecutar el archivo paginaconcreta2.py
    runpy.run_path('paginaconcreta2.py')

# ambos scrips ejecutan finalmente este programa antestodospywin32.py
runpy.run_path('antestodospywin32.py')

# ambos scrips ejecutan finalmente este programa todopywin32.py
runpy.run_path('todopywin32.py')
