import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

print("=== INICIO PIPELINE ===")

# =========================
# CONFIGURACIÓN
# =========================

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

BASE_URL = "https://www.bcp.gov.py/documents/20117/213063/Bolet%C3%ADn+Estad%C3%ADstico+de+Sistemas+de+Pago_{mes}_{anio}.xlsx"

SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1FJnqyffjqsEg_3Qt-ww6hqJYMd6eLXgG1S_FHhbBqmk/edit#gid=0"

INDICADORES_VALIDOS = [
    "Compras en POS con Tarjeta de Crédito",
    "Compras en Internet con Tarjeta de Crédito",
    "Compras en POS con Tarjeta de Débito",
    "Compras en Internet con Tarjeta de Débito",
    "Extracciones en ATM",
    "Compras en POS con Tarjetas Prepagas",
    "Compras en Internet con Tarjetas Prepagas",
    "Extracciones en ATM con Tarjetas Prepagas",
    "Cantidad de ATM por Operadora",
    "Cantidad de Comercios Adheridos por Operadora",
    "Cantidad de POS por Operadora",
    "Total QR"
]

# =========================
# FUNCIÓN BUSCAR ARCHIVO
# =========================

def obtener_ultimo_excel():
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}
    hoy = datetime.today()
    anio_actual = hoy.year

    for year in [anio_actual, anio_actual - 1]:
        for mes in reversed(MESES):
            url = BASE_URL.format(mes=mes, anio=year)
            try:
                r = session.get(url, headers=headers, timeout=10)
                if r.status_code == 200 and len(r.content) > 10000:
                    print(f"Archivo encontrado: {mes} {year}")
                    return url
            except:
                continue
    raise Exception("No se encontró ningún archivo válido en el servidor del BCP.")

# =========================
# DESCARGA Y CARGA DE XLS
# =========================

url_descarga = obtener_ultimo_excel()
response = requests.get(url_descarga)

if response.status_code != 200:
    raise Exception("Error al descargar el archivo desde la URL.")

# Definición del objeto xls para que sea accesible en todo el script
excel_file = BytesIO(response.content)
xls = pd.ExcelFile(excel_file)

# =========================
# PROCESAMIENTO DE DATOS
# =========================

data_final = []

for sheet in xls.sheet_names:
    if not sheet.startswith("OMP"):
        continue

    print(f"Procesando hoja: {sheet}")
    # Leemos la hoja sin header para procesar manualmente las filas múltiples
    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    # Localizar la fila "Año Mes" que sirve como ancla
    # Buscamos en la columna B (índice 1)
    try:
        header_idx = df[df.iloc[:, 1].astype(str).str.contains("Año Mes", na=False)].index[0]
    except IndexError:
        print(f"Saltando {sheet}: No se encontró la fila 'Año Mes'.")
        continue

    # --- LÓGICA DE CABECERAS ---
    # Los indicadores están una fila arriba de "Año Mes" y suelen estar en celdas combinadas
    indicadores_row = df.iloc[header_idx - 1].copy().ffill()
   
    # Las operadoras están en la misma fila que "Año Mes"
    operadoras_row = df.iloc[header_idx].copy()

    # Extraer columna de fechas (desde la fila siguiente al ancla)
    fechas_raw = df.iloc[header_idx + 1:, 1]

    # Iterar por cada columna de datos (desde la columna C en adelante)
    for col in range(2, df.shape[1]):
        indicador = str(indicadores_row[col]).strip()
        operadora = str(operadoras_row[col]).strip()

        # Validaciones de indicadores
        if indicador == "nan" or "Indice" in indicador:
            continue
       
        if not any(k in indicador for k in INDICADORES_VALIDOS):
            continue

        # Si la operadora no está definida en esa fila, buscamos el detalle de "Cantidad/Monto" debajo
        if operadora == "nan":
            detalle_extra = str(df.iloc[header_idx + 1, col]).strip()
            operadora = detalle_extra if detalle_extra != "nan" else "Total/General"

        # Extraer valores históricos
        valores = df.iloc[header_idx + 1:, col]

        # Crear DataFrame temporal para la columna actual
        temp = pd.DataFrame({
            "fecha": fechas_raw,
            "indicador": indicador,
            "operadora": operadora,
            "valor": valores,
            "origen_valor": sheet
        })

        # --- LIMPIEZA DE REGISTROS ---
        # Eliminar valores vacíos, nulos o errores de Excel
        temp = temp.dropna(subset=["valor"])
        temp = temp[~temp["valor"].astype(str).str.contains("#REF!|nan|^$", na=False)]

        # Convertir fecha de YYYY/MM a DD/MM/YYYY
        def convertir_fecha(val):
            try:
                p = str(val).split('/')
                return f"01/{p[1]}/{p[0]}"
            except:
                return None

        temp["fecha"] = temp["fecha"].apply(convertir_fecha)
        temp = temp.dropna(subset=["fecha"])

        data_final.append(temp)

# =========================
# CONSOLIDACIÓN FINAL
# =========================

if not data_final:
    print("No se extrajeron datos. Revisa los INDICADORES_VALIDOS.")
    exit()

df_final = pd.concat(data_final, ignore_index=True)

# Limpieza final de operadoras
df_final["operadora"] = df_final["operadora"].replace("nan", "Total/General")

print(f"Total filas a subir: {len(df_final)}")

# =========================
# CARGA A GOOGLE SHEETS
# =========================

try:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    sheet_dest = spreadsheet.sheet1

    # Limpiar y actualizar
    sheet_dest.clear()
    # Convertimos todo a string para evitar errores de JSON con tipos de datos complejos
    upload_data = [df_final.columns.tolist()] + df_final.astype(str).values.tolist()
    sheet_dest.update(upload_data)

    print("✅ Google Sheets actualizado con éxito.")
except Exception as e:
    print(f"❌ Error al subir a Google Sheets: {e}")

# =========================
# BACKUP LOCAL
# =========================

nombre_archivo = f"bcp_consolidado_{datetime.today().strftime('%Y%m%d')}.csv"
df_final.to_csv(nombre_archivo, index=False)
print(f"💾 Backup guardado como: {nombre_archivo}")
print("=== FIN PIPELINE ===")
