import requests
import pandas as pd
from io import BytesIO
from datetime import datetime

# =========================
# CONFIG
# =========================

MESES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

BASE_URL = "https://www.bcp.gov.py/documents/20117/213063/Bolet%C3%ADn+Estad%C3%ADstico+de+Sistemas+de+Pago_{mes}_{anio}.xlsx"

# =========================
# FUNCION: DETECTAR ÚLTIMO ARCHIVO
# =========================

def obtener_ultimo_excel():
    session = requests.Session()

    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    hoy = datetime.today()
    anio = hoy.year

    print("Buscando último archivo válido...")

    # probar meses hacia atrás
    for year in [anio, anio - 1]:
        for mes in reversed(MESES):

            url = BASE_URL.format(mes=mes, anio=year)

            try:
                response = session.get(url, headers=headers)

                if response.status_code == 200 and len(response.content) > 10000:
                    print(f"Archivo encontrado: {mes} {year}")
                    print("URL:", url)
                    return url

            except:
                continue

    raise Exception("No se encontró ningún archivo válido")


# =========================
# CONFIG DINAMICA
# =========================

FILE_URL = obtener_ultimo_excel()

hoy_str = datetime.today().strftime("%Y%m%d")
OUTPUT_FILE = f"bcp_datos_{hoy_str}.csv"

# =========================
# DESCARGA
# =========================

print("Descargando archivo...")
response = requests.get(FILE_URL)

if response.status_code != 200:
    raise Exception("Error al descargar archivo")

excel_file = BytesIO(response.content)

# =========================
# LEER EXCEL
# =========================

xls = pd.ExcelFile(excel_file)
omp_sheets = [s for s in xls.sheet_names if s.startswith("OMP")]

print("Hojas encontradas:", omp_sheets)

data_final = []

# =========================
# PROCESAMIENTO
# =========================

for sheet_name in omp_sheets:
    print(f"Procesando: {sheet_name}")
    
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    header_row = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("Año Mes", case=False, na=False).any():
            header_row = i
            break

    if header_row is None:
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

if not data_final:
    raise Exception("No se generaron datos")

df_final = pd.concat(data_final, ignore_index=True)

print("Filas generadas:", len(df_final))

# =========================
# EXPORT
# =========================

df_final.to_csv(OUTPUT_FILE, index=False)

print(f"Archivo guardado: {OUTPUT_FILE}")
