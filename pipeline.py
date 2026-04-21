import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
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

# =========================
# VALIDAR CREDENTIALS
# =========================

if not os.path.exists("credentials.json"):
    raise Exception("❌ credentials.json NO EXISTE")

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
                response = session.get(url, headers=headers)

                if response.status_code == 200 and len(response.content) > 10000:
                    print(f"Archivo encontrado: {mes} {year}")
                    return url

            except:
                continue

    raise Exception("❌ No se encontró archivo")

# =========================
# DESCARGA
# =========================

FILE_URL = obtener_ultimo_excel()

response = requests.get(FILE_URL)

if response.status_code != 200:
    raise Exception("❌ Error descarga")

excel_file = BytesIO(response.content)

# =========================
# PROCESAMIENTO
# =========================

xls = pd.ExcelFile(excel_file)
omp_sheets = [s for s in xls.sheet_names if s.startswith("OMP")]

data_final = []

for sheet_name in omp_sheets:
    print(f"Procesando {sheet_name}")

    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("Año Mes", case=False, na=False).any():
            header_row = i
            break

    if header_row is None:
        continue

    metric_row = df.iloc[header_row - 1]
    proc_row = df.iloc[header_row + 1]

    # ✅ FIX
    proc_row = proc_row.ffill()

    df_data = df.iloc[header_row + 2:].copy()

    for col in range(len(df.columns)):

        metrica = metric_row[col]
        procesadora = proc_row[col]

        if pd.isna(metrica):
            continue

        if isinstance(procesadora, str) and "Unnamed" in procesadora:
            continue

        temp = df_data.iloc[:, [1, col]].copy()
        temp.columns = ["fecha_raw", "valor"]

        temp = temp.dropna(subset=["fecha_raw"])

        temp["fecha"] = pd.to_datetime(temp["fecha_raw"], format="%Y/%m", errors="coerce")
        temp = temp.dropna(subset=["fecha"])
        temp["fecha"] = temp["fecha"].dt.strftime("%d/%m/%Y")

        temp["metrica"] = str(metrica).strip()
        temp["procesadora"] = str(procesadora).strip()
        temp["origen"] = sheet_name

        temp = temp[["fecha", "metrica", "procesadora", "valor", "origen"]]

        data_final.append(temp)

df_final = pd.concat(data_final, ignore_index=True)

# =========================
# LIMPIEZA FINAL
# =========================

df_final = df_final.replace([float("inf"), float("-inf")], None)
df_final = df_final.where(pd.notnull(df_final), None)

print("Filas:", len(df_final))

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

sheet.update([df_final.columns.values.tolist()] + df_final.values.tolist())

print("✅ DATA OK")

# =========================
# BACKUP
# =========================

hoy_str = datetime.today().strftime("%Y%m%d")
df_final.to_csv(f"bcp_datos_{hoy_str}.csv", index=False)

print("=== FIN PIPELINE ===")
