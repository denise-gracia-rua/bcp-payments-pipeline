import requests
import pandas as pd
from io import BytesIO

# =========================
# CONFIG
# =========================

FILE_URL = "https://www.bcp.gov.py/documents/20117/213063/Bolet%C3%ADn+Estad%C3%ADstico+de+Sistemas+de+Pago_Marzo_2026.xlsx"
OUTPUT_FILE = "bcp_limpio.csv"

# =========================
# DESCARGA
# =========================

session = requests.Session()
headers = {
    "User-Agent": "Mozilla/5.0",
    "Referer": "https://www.bcp.gov.py/web/institucional/anexo-estadistico-de-pagos"
}

print("Descargando archivo...")
response = session.get(FILE_URL, headers=headers)

if response.status_code != 200:
    raise Exception("Error al descargar archivo")

excel_file = BytesIO(response.content)

# =========================
# LEER EXCEL
# =========================

xls = pd.ExcelFile(excel_file)
omp_sheets = [s for s in xls.sheet_names if s.startswith("OMP")]

data_final = []

# =========================
# PROCESAMIENTO
# =========================

for sheet_name in omp_sheets:
    print(f"Procesando: {sheet_name}")
    
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    # Buscar fila "Año Mes"
    header_row = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("Año Mes", case=False, na=False).any():
            header_row = i
            break

    if header_row is None:
        continue

    # Headers
    metric_row = df.iloc[header_row - 1].ffill()
    proc_row = df.iloc[header_row + 1].ffill()

    # Data
    df_data = df.iloc[header_row + 2:].copy()

    for col in range(2, df.shape[1]):

        metrica = metric_row[col]
        procesadora = proc_row[col]

        if pd.isna(metrica) and pd.isna(procesadora):
            continue

        temp = df_data.iloc[:, [1, col]].copy()
        temp.columns = ["fecha_raw", "valor"]

        temp = temp.dropna(subset=["fecha_raw"])

        temp["fecha"] = pd.to_datetime(
            temp["fecha_raw"], format="%Y/%m", errors="coerce"
        )

        temp = temp.dropna(subset=["fecha"])
        temp["fecha"] = temp["fecha"].dt.strftime("%d/%m/%Y")

        temp["metrica"] = str(metrica).strip()
        temp["procesadora"] = str(procesadora).strip()
        temp["origen"] = sheet_name

        temp = temp[["fecha", "metrica", "procesadora", "valor", "origen"]]

        data_final.append(temp)

# =========================
# FINAL
# =========================

df_final = pd.concat(data_final, ignore_index=True)

print("Filas generadas:", len(df_final))

# =========================
# EXPORT
# =========================

df_final.to_csv(OUTPUT_FILE, index=False)

print(f"Archivo guardado: {OUTPUT_FILE}")