import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# Lee el archivo CSV
df = pd.read_csv(r'C:\Users\Asus\Downloads\Users25-06.csv')

# Muestra los nombres de las columnas y las primeras filas del DataFrame
print("Nombres de las columnas en el archivo CSV:")
print(df.columns)
print("\nPrimeras filas del DataFrame original:")
print(df.head())

# Lista de columnas deseadas
columnas_deseadas = [
    'active', 'name.familyName', 'name.givenName', 'userName',
    'urn:ietf:params:scim:schemas:extension:enterprise:2.0:User:division',
    'urn:ietf:params:scim:schemas:extension:enterprise:2.0:User:employeeNumber',
    'urn:ietf:params:scim:schemas:extension:sap:2.0:User:loginTime', 'emails[0].value'
]

# Verifica que las columnas deseadas existen en el DataFrame
columnas_existentes = [col for col in columnas_deseadas if col in df.columns]
print("\nColumnas seleccionadas que existen en el DataFrame:")
print(columnas_existentes)

# Selecciona solo las columnas deseadas que existen
df_reducido = df[columnas_existentes]

# Renombra las columnas
nuevos_nombres = {
    'active': 'Estado',
    'name.familyName': 'Apellido',
    'name.givenName': 'Nombre',
    'userName': 'Usuario',
    'urn:ietf:params:scim:schemas:extension:enterprise:2.0:User:division': 'Gerencia',
    'urn:ietf:params:scim:schemas:extension:enterprise:2.0:User:employeeNumber': 'Area',
    'urn:ietf:params:scim:schemas:extension:sap:2.0:User:loginTime': 'Hora de conexion',
    'emails[0].value': 'correo'
}

df_reducido.rename(columns=nuevos_nombres, inplace=True)

# Muestra una muestra de las primeras filas del DataFrame reducido
print("\nPrimeras filas del DataFrame reducido (antes del filtrado):")
print(df_reducido.head())

# Verificar los valores únicos de 'Estado'
print("\nValores únicos de 'Estado' antes del filtrado:")
print(df_reducido['Estado'].unique())

# Filtrar por 'Estado' y cambiar True a "ACTIVO"
df_reducido = df_reducido[df_reducido['Estado'] == True]
df_reducido['Estado'] = 'ACTIVO'

# Verificar los valores únicos de 'correo' que contienen '@mallplaza.com'
print("\nValores únicos de 'correo' que contienen '@mallplaza.com':")
print(df_reducido[df_reducido['correo'].str.contains('@mallplaza.com', na=False)]['correo'].unique())

# Filtrar por correos que contengan "@mallplaza.com"
df_reducido = df_reducido[df_reducido['correo'].str.contains('@mallplaza.com', na=False)]

# Formatear la columna 'Hora de conexion' a formato de fecha y hora
df_reducido['Hora de conexion'] = pd.to_datetime(df_reducido['Hora de conexion'], errors='coerce').dt.strftime('%Y-%m-%d %H:%M:%S')

# Guarda el DataFrame reducido en un nuevo archivo Excel
ruta_excel = r'C:\Users\Asus\Downloads\ReporteConexion2 5-06.xlsx'
df_reducido.to_excel(ruta_excel, index=False, engine='openpyxl')

# Aplicar formato al encabezado
wb = load_workbook(ruta_excel)
ws = wb.active

# Definir el estilo de relleno verde y la fuente blanca
relleno_verde = PatternFill(start_color="5ccb5f", end_color="5ccb5f", fill_type="solid")
fuente_blanca = Font(color="FFFFFF")

# Aplicar el estilo al encabezado
for cell in ws[1]:
    cell.fill = relleno_verde
    cell.font = fuente_blanca

# Guardar el archivo Excel con el formato aplicado
wb.save(ruta_excel)

# Muestra las primeras filas del DataFrame reducido
print("\nPrimeras filas del DataFrame reducido (después del filtrado):")
print(df_reducido.head())
