import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel
archivo = 'zona20.xlsx'
hoja1 = pd.read_excel(archivo, sheet_name='Hoja1', engine='openpyxl')
hoja2 = pd.read_excel(archivo, sheet_name='Hoja2', engine='openpyxl')
codigos = pd.read_excel(archivo, sheet_name='Codigos', engine='openpyxl')
equipos = pd.read_excel(archivo, sheet_name='Equipos', engine='openpyxl')

# Obtener columnas deseadas
columnas_deseadas = hoja2.columns.tolist()
hoja1_filtrada = hoja1[columnas_deseadas].copy()
hoja1_filtrada.dropna(how='all', inplace=True)

# Procesar 'Creation time'
if 'Creation time' in hoja1_filtrada.columns:
    fecha_ayer = (datetime.now() - timedelta(days=1)).strftime('%d/%m/%Y')
    def limpiar_fecha(valor):
        if isinstance(valor, str) and 'Yesterday' in valor:
            return fecha_ayer
        try:
            return pd.to_datetime(valor, errors='coerce').strftime('%d/%m/%Y')
        except:
            return None
    hoja1_filtrada['Creation time'] = hoja1_filtrada['Creation time'].apply(limpiar_fecha)

# Función para buscar equivalencias de ciudad y zona
def buscar_equivalencia_ciudad(nombre):
    partes = str(nombre).split('_')
    for parte in partes:
        fila = codigos[codigos['CODIGO'] == parte]
        if not fila.empty:
            return {
                'Element.REGION': fila['ZONA'].values[0],
                'Element.PLAZA': fila['CIUDAD'].values[0]
            }
    return {
        'Element.REGION': 'nulo',
        'Element.PLAZA': 'nulo'
    }

# Función para buscar departamento por equipo
def buscar_departamento(nombre):
    for _, fila in equipos.iterrows():
        if fila['EQUIPOS'] in str(nombre):
            return fila['DEPARTAMENTO']
    return 'nulo'

# Normalizar valores vacíos
hoja1_filtrada = hoja1_filtrada.applymap(lambda x: x if pd.notna(x) and str(x).strip() not in ['', '-', '_'] else None)

# Rellenar campos vacíos
for i, fila in hoja1_filtrada.iterrows():
    if any(pd.isna(fila[campo]) for campo in ['Element.REGION', 'Element.PLAZA']):
        valores = buscar_equivalencia_ciudad(fila['Element name'])
        for campo, valor in valores.items():
            if pd.isna(fila[campo]):
                hoja1_filtrada.at[i, campo] = valor
    if pd.isna(fila['Element.UNIDAD DE NEGOCIO']):
        hoja1_filtrada.at[i, 'Element.UNIDAD DE NEGOCIO'] = buscar_departamento(fila['Element name'])

# Rellenar cualquier otro campo vacío con 'SIN DATO'
hoja1_filtrada.fillna('SIN DATO', inplace=True)

# Reordenar y guardar
hoja1_filtrada = hoja1_filtrada[columnas_deseadas]
hoja1_filtrada.to_excel('zona20_limpio_completo.xlsx', index=False)

print("Archivo limpio generado exitosamente como 'zona20_limpio_completo.xlsx'.")
