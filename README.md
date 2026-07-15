# SIMULACROS

Este repositorio contiene scripts para procesar PDFs y generar reportes.

## Entorno de producción

Sigue estos pasos para preparar el entorno de ejecución:

1. Abre PowerShell en la carpeta raíz del proyecto (`c:\Scripts\3. SIMULACROS`).
2. Ejecuta el script para crear el entorno virtual e instalar dependencias:

   ```powershell
   .\crear_venv.ps1
   ```

3. Si el entorno ya está creado o necesitas activarlo en una nueva sesión, ejecuta:

   ```powershell
   .\.venv\Scripts\Activate.ps1
   ```

4. Verifica que las dependencias estén instaladas correctamente:

   ```powershell
   python -m pip list
   ```

5. Ejecuta los scripts del proyecto según sea necesario. Por ejemplo:

   ```powershell
   python .\Reporte\main.py
   ```

### Notas

- El script `crear_venv.ps1` usa el Python que esté disponible en la variable de entorno `PATH`.
- Si no tienes Python instalado, instala Python 3.11+ antes de ejecutar el script.
- El archivo `requirements.txt` contiene las dependencias necesarias para este proyecto.
