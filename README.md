# Sistema de Liquidación - Procesamiento de Cargas

Descripción
-	Aplicación GUI en Tkinter para seleccionar, validar y consolidar archivos Excel con datos de liquidaciones. Permite editar cargas consolidadas y exportar los datasets resultantes a archivos TXT con separador `|`.

Requisitos
-	Sistema operativo: Windows (probado), también puede funcionar en Linux/Mac con las librerías indicadas.
-	Python: Instalar Python 3.11, 3.12 o 3.13 (recomendado usar la misma versión que tengas instalada; la app se probó con Python 3.13).

Instalación de dependencias
-	Crear y activar un entorno virtual (recomendado):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1   # PowerShell
# o en cmd: .\.venv\Scripts\activate
```

-	Instalar las librerías necesarias con pip:

```powershell
pip install pandas pillow tkcalendar pyinstaller
```

Notas sobre dependencias
-	`tkinter` viene incluido con la mayoría de instalaciones oficiales de Python en Windows; en Linux puede requerir paquete adicional (`python3-tk`).
-	`tkcalendar` es opcional en tiempo de ejecución: si no está instalado, la app usa un `Entry` en lugar del DateEntry.
-	`Pillow` (PIL) se usa para escalar la imagen del pie de página; si no está presente, el programa intenta un fallback con `PhotoImage`.

Archivo `requirements.txt` sugerido
```
pandas
pillow
tkcalendar
pyinstaller
```

Estructura del proyecto
-	`procesarInfos.py` : código principal de la aplicación.
-	`img/` : carpeta con imágenes necesarias por la app (p. ej. `datosPAS_pie.png`).

Ejecución durante desarrollo
-	Desde la carpeta del proyecto activa con el entorno virtual:

```powershell
python procesarInfos.py
```

Crear un ejecutable (.exe) con PyInstaller (Windows)
-	Instalar PyInstaller si no se instaló:

```powershell
pip install pyinstaller
```

-	Generar un `.exe` incluyendo la carpeta `img` (nota: la sintaxis `--add-data "SRC;DEST"` es para Windows):

```powershell
pyinstaller --noconfirm --onefile --windowed --add-data "img;img" procesarInfos.py
```

Explicación:
-	`--onefile` genera un único ejecutable en `dist\`.
-	`--windowed` evita abrir una consola (útil para aplicaciones GUI).
-	`--add-data "img;img"` empaqueta la carpeta `img` dentro del ejecutable; después de ejecutar el exe, la app podrá acceder a los archivos en la ruta relativa `img\...`.

Advertencias y comprobaciones
-	Si el exe no encuentra las imágenes, verifica que el `--add-data` se usó correctamente y que la ruta relativa en el código (`os.path.join(os.path.dirname(__file__), 'img', 'datosPAS_pie.png')`) resuelva correctamente.
-	Durante desarrollo prueba primero con `python procesarInfos.py` para confirmar que todas las dependencias están instaladas.

Soporte y ajustes
-	Si quieres que el exe incluya un icono, añade `--icon=mi_icono.ico` al comando de PyInstaller.
-	Si deseas que la app arranque en modo fullscreen con F11, puedo añadir el atajo y el comportamiento.

---

Si quieres, genero también un `requirements.txt` automáticamente y agrego el atajo F11 para alternar fullscreen.
