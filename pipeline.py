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

print("Verificando credentials.json...")

if not os.path.exists("credentials.json"):
    raise Exception("❌ credentials.json NO EXISTE")

with open("credentials.json") as f:
    try:
        creds_json = json.load(f)
        print("✅ credentials.json válido")
        print("Service account:", creds_json.get("client_email"))
    except:
        raise Exception("❌ credentials.json inválido")

# =========================
# BUSCAR ÚLTIMO ARCHIVO
# =========================

def obtener_ultimo_excel():
    session = requests.Session()
    headers = {"User-Agent": "Mozilla/5.0"}

    hoy = datetime.today()
    anio = hoy.year

    print("Buscando archivo BCP...")

    for year in [anio, anio - 1]:
        for mes in reversed(MESES):
            url = BASE_URL.format(mes=mes, anio=year)

            try:
                response = session.get(url, headers=headers)

                if response.status_code == 200 and len(response.content) > 10000:
                    print(f"✅ Archivo encontrado: {mes} {year}")
                    return url

            except Exception as e:
                print("Error request:", e)

    raise Exception("❌ No se encontró archivo")

# =========================
# DESCARGA
# =========================

FILE_URL = obtener_ultimo_excel()

print("Descargando archivo...")
response = requests.get(FILE_URL)

print("Status descarga:", response.status_code)

if response.status_code != 200:
    raise Exception("❌ Error descarga")

excel_file = BytesIO(response.content)

# =========================
# PROCESAMIENTO
# =========================

print("Leyendo Excel...")

xls = pd.ExcelFile(excel_file)
omp_sheets = [s for s in xls.sheet_names if s.startswith("OMP")]

print("Sheets:", omp_sheets)

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
        print("Saltado")
        continue

    metric_row = df.iloc[header_row - 1].ffill()
    proc_row = df.iloc[header_row + 1].ffill()
    df_data = df.iloc[header_row + 2:].copy()

    for col in range(2, df.shape[1]):
        metrica = metric_row[col]
        procesadora = proc_row[col]

        if pd.isna(metrica) and pd.isna(procesadora):
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

if not data_final:
    raise Exception("❌ No se generaron datos")

df_final = pd.concat(data_final, ignore_index=True)

print("Filas:", len(df_final))

# =========================
# GOOGLE SHEETS
# =========================

print("Conectando a Google Sheets...")

try:
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)

    print("✅ Auth OK")

    spreadsheet = client.open_by_url(SPREADSHEET_URL)
    print("✅ Sheet abierto")

    sheet = spreadsheet.sheet1
    print("✅ Worksheet OK")

    sheet.clear()
    print("✅ Sheet limpio")

    sheet.update([df_final.columns.values.tolist()] + df_final.values.tolist())

    print("✅ DATA OK - GOOGLE SHEETS ACTUALIZADO")

except Exception as e:
    print("❌ ERROR GOOGLE SHEETS:")
    print(str(e))
    raise e

# =========================
# BACKUP CSV
# =========================

hoy_str = datetime.today().strftime("%Y%m%d")
OUTPUT_FILE = f"bcp_datos_{hoy_str}.csv"

df_final.to_csv(OUTPUT_FILE, index=False)

print(f"Archivo backup guardado: {OUTPUT_FILE}")

print("=== FIN PIPELINE ===")
