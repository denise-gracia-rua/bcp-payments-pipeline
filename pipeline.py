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
    
    print("Status página:", response.status_code)

    if response.status_code != 200:
        raise Exception("No se pudo acceder a la página del BCP")

    soup = BeautifulSoup(response.text, "html.parser")

    links = soup.find_all("a", href=True)

    excel_links = []

    for link in links:
        href = link["href"]
        if ".xlsx" in href:
            if not href.startswith("http"):
                href = "https://www.bcp.gov.py" + href
            excel_links.append(href)

    print(f"Cantidad de excels encontrados: {len(excel_links)}")

    for l in excel_links[:5]:
        print("Ejemplo link:", l)

    if not excel_links:
        raise Exception("No se encontraron archivos Excel")

    # tomar el más reciente (heurística)
    excel_links = sorted(excel_links, key=len, reverse=True)

    ultimo = excel_links[0]

    print("Excel elegido:", ultimo)

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

print("Status descarga:", response.status_code)

if response.status_code != 200:
    raise Exception("Error al descargar archivo")

if len(response.content) < 10000:
    raise Exception("Archivo descargado sospechosamente pequeño")

excel_file = BytesIO(response.content)

# =========================
# LEER EXCEL
# =========================

try:
    xls = pd.ExcelFile(excel_file)
except Exception as e:
    raise Exception(f"Error leyendo Excel: {e}")

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
        print(f"⚠️ No se encontró 'Año Mes' en {sheet_name}")
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
    raise Exception("No se generaron datos (todas las hojas fallaron)")

df_final = pd.concat(data_final, ignore_index=True)

print("Filas generadas:", len(df_final))

# =========================
# EXPORT
# =========================

df_final.to_csv(OUTPUT_FILE, index=False)

print(f"Archivo guardado: {OUTPUT_FILE}")
