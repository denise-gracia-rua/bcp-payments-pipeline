import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
from bs4 import BeautifulSoup

# =========================
# FUNCION: DETECTAR ÚLTIMO EXCEL
# =========================

def obtener_ultimo_excel():
    url = "https://www.bcp.gov.py/web/institucional/anexo-estadistico-de-pagos"
    
    headers = {
        "User-Agent": "Mozilla/5.0"
    }

    print("Buscando último archivo en BCP...")
    response = requests.get(url, headers=headers)
    
    if response.status_code != 200:
        raise Exception("No se pudo acceder a la página del BCP")

    soup = BeautifulSoup(response.text, "html.parser")

    links = soup.find_all("a", href=True)

    excel_links = []

    for link in links:
        href = link["href"]
        if ".xlsx" in href and "Bolet" in href:
            excel_links.append(href)

    if not excel_links:
        raise Exception("No se encontró archivo Excel")

    ultimo = excel_links[0]

    if not ultimo.startswith("http"):
        ultimo = "https://www.bcp.gov.py" + ultimo

    print("Excel detectado:", ultimo)

    return ultimo


# =========================
# CONFIG
# =========================

FILE_URL = obtener_ultimo_excel()

hoy = datetime.today().strftime("%Y%m%d")
OUTPUT_FILE = f"bcp_datos_{hoy}.csv"

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

    # Headers (manejo de celdas combinadas)
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

if not data_final:
    raise Exception("No se generaron datos")

df_final = pd.concat(data_final, ignore_index=True)

print("Filas generadas:", len(df_final))

# =========================
# EXPORT
# =========================

df_final.to_csv(OUTPUT_FILE, index=False)

print(f"Archivo guardado: {OUTPUT_FILE}")
