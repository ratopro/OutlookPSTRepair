# Herramienta de Reparación de Archivos PST de Outlook

Esta herramienta permite analizar archivos PST de Outlook dañados, visualizar su estructura de carpetas en formato de árbol, identificar carpetas correctas, dañadas y eliminadas, y seleccionar qué carpetas recuperar.

## Características

- **Interfaz Gráfica (GUI)**: Aplicación con interfaz gráfica usando Tkinter
- **Análisis de archivos PST**: Utiliza la librería `pypff` para leer y analizar archivos PST
- **Visualización de árbol de carpetas**: Muestra la estructura jerárquica de carpetas
- **Barra de progreso**: Indicador visual durante el análisis y reparación
- **Identificación de estados**: 
  - Verde: Carpetas correctas
  - Naranja: Carpetas dañadas
  - Rojo: Carpetas eliminadas
  - Rosa: Carpetas faltantes (missing)
- **Selección de carpetas**: Permite seleccionar qué carpetas recuperar
- **Reparación**: Crea una copia del archivo PST con las carpetas seleccionadas

## Requisitos

- Python 3.x
- Librería `pypff` (instalada mediante `pip install pypff-python --break-system-packages`)

## Uso

### Interfaz Gráfica (GUI)
```bash
python3 pst_repair_gui.py
```

1. Haga clic en "Examinar..." para seleccionar un archivo PST
2. Haga clic en "Analizar" para cargar y analizar el archivo
3. Seleccione las carpetas que desea recuperar usando los checkboxes
4. Haga clic en "Reparar Seleccionadas" para crear el archivo recuperado

### Línea de comandos
```bash
python3 pst_repair_simple.py "ruta/al/archivo.pst"
```

#### Selección de carpetas (versión CLI)
```bash
# Seleccionar carpetas específicas (por número en el árbol mostrado)
python3 pst_repair_simple.py "ruta/al/archivo.pst" "1,3,5"

# Seleccionar un rango
python3 pst_repair_simple.py "ruta/al/archivo.pst" "2-4"

# Seleccionar todas las carpetas
python3 pst_repair_simple.py "ruta/al/archivo.pst" "all"

# No seleccionar ninguna carpeta
python3 pst_repair_simple.py "ruta/al/archivo.pst" "none"
```

#### Con archivo de salida (simulación de recuperación)
```bash
python3 pst_repair_simple.py "ruta/al/archivo.pst" "1,3,5" "ruta/al/archivo_recuperado.pst"
```

## Ejemplo de salida (CLI)

```
======================================================================
ÁRBOL DE CARPETAS PST
======================================================================
Estados: ✓ Correcto | ⚠ Dañado | ✗ Eliminado
----------------------------------------------------------------------
✓ <Carpeta Raíz> (0 correos)
✓     Buscar raíz (0 correos)
✓     IPM_COMMON_VIEWS (0 correos)
✓     Principio del archivo de datos de Outlook (0 correos)
✗         Elementos eliminados (0 correos)
✓         carpeta (5 correos)
✓         carpeta 3 (25 correos)
✓             carpeta 3_1 (14 correos)
✓         carpeta2 (22 correos)
✓     SPAM Search Folder 2 (0 correos)
----------------------------------------------------------------------
Total: 10 carpetas, 66 correos
======================================================================
```

## Notas importantes

1. Esta herramienta realiza una **copia** del archivo original con las carpetas seleccionadas.

2. La herramienta identifica carpetas eliminadas mediante heurísticas basadas en nombres comunes como "Elementos eliminados", "Deleted Items", "Trash", etc.

3. Las carpetas dañadas se detectan cuando no se pueden leer sus propiedades básicas o contenido.

4. Las carpetas "missing" son aquellas cuyo padre no existe en la estructura del archivo.

## Compilar para Windows

### Requisitos de compilación
```bash
pip install pyinstaller
```

### Compilar ejecutable
```bash
pyinstaller --onefile --name OutlookPSTRepairCi --icon=icono.ico pst_repair_gui.py
```

El parámetro `--name OutlookPSTRepairCi` establece el nombre del ejecutable en Windows.

Opcionalmente puedes agregar un icono:
```bash
pyinstaller --onefile --name OutlookPSTRepairCi --icon=icono.ico pst_repair_gui.py
```

El archivo `.ico` debe ser un icono válido de Windows (puedes crear uno en icoconvert.com).

## Archivos en este directorio

- `pst_repair_gui.py`: Aplicación con interfaz gráfica (Tkinter)
- `pst_repair_simple.py`: Versión de línea de comandos
- `README.md`: Este archivo