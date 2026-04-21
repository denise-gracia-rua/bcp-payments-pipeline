import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

print("=== INICIO PIPELINE ===")

# =========================
# CONFIG
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
# BUSCAR ARCHIVO
# =========================

def obtener_ultimo_excel():
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}

    hoy = datetime.today()
    anio = hoy.year

    for year in [anio, anio - 1]:
        for mes in reversed(MESES):
            url = BASE_URL.format(mes=mes, anio=year)

            try:
                r = session.get(url, headers=headers)
                if r.status_code == 200 and len(r.content) > 10000:
                    print(f"Archivo encontrado: {mes} {year}")
                    return url
            except:
                continue

    raise Exception("No se encontró archivo")

# =========================
# DESCARGA
# =========================

url = obtener_ultimo_excel()
file = requests.get(url)

if file.status_code != 200:
    raise Exception("Error descargando archivo")

excel = BytesIO(file.content)
xls = pd.ExcelFile(excel)

# =========================
# PROCESAMIENTO
# =========================

data_final = []

for sheet in xls.sheet_names:

    if not sheet.startswith("OMP"):
        continue

    print(f"Procesando {sheet}")

    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    # Buscar fila Año Mes
    header_row = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("Año Mes", na=False).any():
            header_row = i
            break

    if header_row is None:
        continue

    # Fechas
    fechas = df.iloc[header_row + 2:, 1]
    fechas = pd.to_datetime(fechas, format="%Y/%m", errors="coerce")

    # 🔥 FILA INDICADOR
    metric_row = df.iloc[header_row - 1]

    # 🔥 FILA OPERADORAS (FIX CLAVE)
    operadoras_row = df.iloc[header_row + 1].copy()
    operadoras_row = operadoras_row.ffill()  # ← ESTE ES EL FIX

    for col in range(2, df.shape[1]):

        indicador = metric_row[col]
        operadora = operadoras_row[col]

        if pd.isna(indicador):
            continue

        indicador = str(indicador).strip()

        if not any(k in indicador for k in INDICADORES_VALIDOS):
            continue

        if pd.isna(operadora):
            continue

        operadora = str(operadora).strip()

        # evitar columnas basura
        if operadora.lower() in ["cantidad", "monto"]:
            continue

        valores = df.iloc[header_row + 2:, col]

        temp = pd.DataFrame({
            "fecha": fechas,
            "indicador": indicador,
            "operadora": operadora,
            "valor": valores,
            "origen_valor": sheet
        })

        temp = temp.dropna(subset=["fecha", "valor"])

        temp["fecha"] = temp["fecha"].dt.strftime("%d/%m/%Y")

        data_final.append(temp)

# =========================
# CONCAT
# =========================

df_final = pd.concat(data_final, ignore_index=True)

# =========================
# LIMPIEZA
# =========================

df_final = df_final.replace([float("inf"), float("-inf")], None)
df_final = df_final.where(pd.notnull(df_final), None)

print("Filas finales:", len(df_final))

# =========================
# GOOGLE SHEETS
# =========================

scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

spreadsheet = client.open_by_url(SPREADSHEET_URL)
sheet = spreadsheet.sheet1

sheet.clear()
sheet.update([df_final.columns.tolist()] + df_final.values.tolist())

print("✅ Google Sheets actualizado")

# =========================
# BACKUP CSV
# =========================

hoy = datetime.today().strftime("%Y%m%d")
df_final.to_csv(f"bcp_datos_{hoy}.csv", index=False)

print("=== FIN PIPELINE ===")
