import pandas as pd
import re
from datetime import datetime


# -------------------------
# 🔧 UTILIDADES
# -------------------------
def limpiar(valor):
    if pd.isna(valor):
        return ""
    return str(valor).replace("\n", " ").strip()


def to_float(valor):
    if pd.isna(valor) or valor == "":
        return 0.0
    
    # Si ya es un número (int o float), lo devolvemos tal cual
    if isinstance(valor, (int, float)):
        return float(valor)
    
    # Si es un string, aplicamos la limpieza de formato
    try:
        # Quitamos puntos de miles y cambiamos coma decimal por punto
        # Solo si detectamos que parece un formato con coma decimal (ej: 9.000,00)
        if "," in str(valor) and "." in str(valor):
            valor = str(valor).replace(".", "").replace(",", ".")
        elif "," in str(valor):
            valor = str(valor).replace(",", ".")
            
        return float(valor)
    except:
        return 0.0


def extraer_bl_hijo(bl):
    match = re.search(r'\(H\)([^)]+)', bl)
    return f"(H){match.group(1)}" if match else bl


def formatear_fecha(fecha):
    try:
        d, m, y = fecha.split(".")
        return f"{d}-{m}-{y}"
    except:
        return fecha


def fecha_actual():
    return datetime.now().strftime("%d-%m-%Y %H:%M")


# -------------------------
# 🔍 DETECTAR HOJA + HEADER
# -------------------------
def detectar_hoja_y_header(ruta):
    xls = pd.ExcelFile(ruta)

    columnas_clave = ["BL NUMBER", "SPECIFICATION", "CONSIGNEE"]

    for hoja in xls.sheet_names:
        for i in range(10):
            try:
                df = pd.read_excel(xls, sheet_name=hoja, header=i)

                columnas = [
                    col.strip().upper()
                    for col in df.columns
                    if isinstance(col, str)
                ]

                if all(col in columnas for col in columnas_clave):
                    return df

            except:
                continue

    raise ValueError("No se encontró hoja válida")


# -------------------------
# 🚀 PRINCIPAL
# -------------------------
def leer_excel_multiple(ruta):
    df = detectar_hoja_y_header(ruta)

    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()

    # detectar columna BL
    posibles_bl = ["BL NUMBER", "B/NUMBER"]
    col_bl = next((c for c in posibles_bl if c in df.columns), None)

    if not col_bl:
        raise ValueError("No se encontró columna BL")

    df[col_bl] = df[col_bl].astype(str).str.strip()

    resultados = []

    # -------------------------
    # 🔥 AGRUPAR POR BL
    # -------------------------
    for bl, grupo in df.groupby(col_bl):
        if not bl or bl.lower() == "nan":
            continue

        fila = grupo.iloc[0]

        # -------------------------
        # ITEMS (solo specification)
        # -------------------------
        items = []

        for _, row in grupo.iterrows():
            spec = limpiar(row.get("SPECIFICATION"))

            if not spec:
                continue

            items.append({
                "specification": spec
            })

        if not items:
            continue

        # -------------------------
        # VEHÍCULO (UNO POR BL)
        # -------------------------
        make = limpiar(fila.get("MAKE"))
        tipo = limpiar(fila.get("TYPE"))
        vin = limpiar(fila.get("VIN"))
        transit = limpiar(fila.get("TRANSIT"))

        # -------------------------
        # PESO TOTAL (desde columna, no por item)
        # -------------------------
        peso_total = to_float(fila.get("WEIGHT"))
        freight = to_float(fila.get("FREIGHT FOR BL"))

        # -------------------------
        # DATOS FINALES
        # -------------------------
        datos = {
            "bl_completo": limpiar(bl),
            "bl": extraer_bl_hijo(bl),
            "nit": str(limpiar(fila.get("NIT"))).replace(".0", "") if "NIT" in df.columns else "",
            "pol": limpiar(fila.get("POL")),
            "pod": limpiar(fila.get("POD")),
            "consignee": limpiar(fila.get("CONSIGNEE")),
            "fecha_bl": formatear_fecha(limpiar(fila.get("DATE OF BL"))),
            "fecha_pres": fecha_actual(),
            "nave": limpiar(fila.get("NAVE")) if "NAVE" in df.columns else "",
            "transit": limpiar(fila.get("TRANSIT")),

            # 🔥 CAMBIOS IMPORTANTES
            "peso_total": peso_total,
            "bultos": len(items),

            "freight": freight,
            "items": items,

            # 🚛 VEHÍCULO FUERA DE ITEMS
            "make": make,
            "type": tipo,
            "vin": vin
        }

        resultados.append(datos)

    return resultados