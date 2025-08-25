
# 📊 Limpieza de Datos para Reportes Power BI

Este script en Python automatiza la limpieza y transformación de datos provenientes de un archivo Excel (`archivo.xlsx`(cambiar por el nombre de tu archivo)). El objetivo es generar una versión estandarizada y enriquecida de los datos que sirva como base para reportes en Power BI.

## 🧩 Funcionalidad del Script

El script realiza las siguientes tareas:

### Carga de datos desde varias hojas del archivo Excel:
- **Hoja1**: Datos crudos.
- **Hoja2**: Estructura deseada.
- **Codigos**: Equivalencias de ciudad y zona.
- **Equipos**: Relación entre equipos y departamentos.

### Filtrado de columnas:
- Se conservan solo las columnas presentes en `Hoja2`.

### Limpieza de fechas:
- Convierte fechas en formato texto.
- Reemplaza `"Yesterday"` por la fecha correspondiente.

### Normalización de valores vacíos:
- Elimina valores como `''`, `'-'`, `'_'`.

### Enriquecimiento de datos:
- Asigna `Element.REGION` y `Element.PLAZA` usando equivalencias de ciudad.
- Asigna `Element.UNIDAD DE NEGOCIO` según el equipo.

### Relleno de campos vacíos:
- Se completa con `'SIN DATO'` cuando no hay información disponible.

### Exportación:
- Guarda el archivo limpio como `archivo_limpio.xlsx`.(puedes cambiar el nombre al que desees)

## 🚀 Cómo Ejecutarlo

### Requisitos
- Python 3.8 o superior
- Paquetes necesarios:
  - `pandas`
  - `openpyxl`

### Instalación de dependencias

```bash
pip install pandas openpyxl
