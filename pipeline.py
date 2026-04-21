import requests
import pandas as pd
from io import BytesIO
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os

print("=== INICIO PIPELINE ===")

# ... (Configuración de MESES, BASE_URL y SPREADSHEET_URL igual a tu código)

# =========================
# PROCESAMIENTO CORREGIDO
# =========================

data_final = []

for sheet in xls.sheet_names:
    if not sheet.startswith("OMP"):
        continue

    print(f"Procesando {sheet}")
    # Leemos la hoja completa sin cabeceras para manejar la estructura manualmente
    df = pd.read_excel(xls, sheet_name=sheet, header=None)

    # 1. Localizar la fila "Año Mes" que es nuestra ancla
    header_idx = None
    for i in range(len(df)):
        if df.iloc[i].astype(str).str.contains("Año Mes", na=False).any():
            header_idx = i
            break

    if header_idx is None:
        print(f"No se encontró ancla en {sheet}, saltando...")
        continue

    # 2. Extraer Filas de Metadatos (ajustado a la estructura real del BCP)
    # Fila de Indicadores (ej: Cantidad de ATM...) suele estar 1 fila ARRIBA de Año Mes
    indicadores_row = df.iloc[header_idx - 1].copy()
    indicadores_row.iloc[2:] = indicadores_row.iloc[2:].ffill() # Completar celdas combinadas

    # Fila de Operadoras (ej: BANCARD, UPAY...) es la fila donde está "Año Mes"
    operadoras_row = df.iloc[header_idx].copy()

    # 3. Extraer Fechas y Datos
    # Las fechas están en la columna index 1, desde 1 fila abajo del header
    fechas_raw = df.iloc[header_idx + 1:, 1]
   
    # 4. Iterar por columnas de datos (desde la columna 2 en adelante)
    for col in range(2, df.shape[1]):
        indicador = str(indicadores_row[col]).strip()
        operadora = str(operadoras_row[col]).strip()

        # Limpieza de nombres
        if indicador == "nan" or "Indice" in indicador:
            continue
       
        # Si la operadora es "nan", suele ser porque el indicador no tiene desglose
        # En ese caso usamos "Total/General" o lo que indique la fila de abajo
        if operadora == "nan":
            # Intentamos ver si hay un detalle de "Cantidad/Monto" en la fila siguiente
            detalle_debajo = str(df.iloc[header_idx + 1, col]).strip()
            operadora = detalle_debajo if detalle_debajo != "nan" else "Total/General"

        # Extraer valores de la columna actual
        valores = df.iloc[header_idx + 1:, col]

        # Crear DataFrame temporal para esta columna
        temp = pd.DataFrame({
            "fecha": fechas_raw,
            "indicador": indicador,
            "operadora": operadora,
            "valor": valores,
            "origen_valor": sheet
        })

        # Limpieza de datos dentro de la columna
        temp = temp.dropna(subset=["valor"])
        temp = temp[temp["valor"].astype(str).str.strip() != ""]
        temp = temp[temp["valor"].astype(str).str.strip() != "#REF!"]
       
        # Formatear fecha dd/mm/aaaa
        # El BCP usa YYYY/MM, lo convertimos:
        def transformar_fecha(x):
            try:
                partes = str(x).split('/')
                return f"01/{partes[1]}/{partes[0]}"
            except: return None

        temp["fecha"] = temp["fecha"].apply(transformar_fecha)
        temp = temp.dropna(subset=["fecha"])

        data_final.append(temp)

# =========================
# CONSOLIDACIÓN Y ENVÍO
# =========================

if data_final:
    df_final = pd.concat(data_final, ignore_index=True)
   
    # Limpieza final de strings
    df_final["operadora"] = df_final["operadora"].replace("nan", "Total/General")
   
    # ... (Resto de tu código de Google Sheets y Backup igual)
    print(f"Éxito: {len(df_final)} filas procesadas.")
else:
    print("No se recolectaron datos.")
