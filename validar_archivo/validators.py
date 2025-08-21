import pandas as pd
import tkinter as tk
from tkinter import filedialog

# ======================
# CONFIGURACIÓN
# ======================
REQUIRED_COLUMNS = [
    "PRODUCT_ID",
    "BRAND_NAME",
    "GENERIC_NAME",
    "CATEGORY_ID",
    "FAMILY",
    "LABORATORY_ID",
    "PRICE_US",
    "PRICE_NATIONAL",
    "REF",
    "RX",
    "STATUS",
]

OPTIONAL_COLUMNS = [
    "DESCRIPTION",
    "DOSAGE",
    "PRESENTATION",
    "QUANTITY_IN_A_BOX",
    "INVENTORY",
    "INDICATIONS",
    "DETAILS",
    "INGREDIENTS",
    "IMAGE_NAME",
]

# ======================
# FUNCIONES
# ======================
def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    """Limpia y normaliza los nombres de columnas a MAYÚSCULAS con _"""
    df.columns = (
        df.columns.str.strip()
        .str.upper()
        .str.replace(" ", "_")
        .str.replace("-", "_")
    )
    return df

def validar_columnas(df: pd.DataFrame):
    df = normalizar_columnas(df)
    errores = []
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            errores.append(f"Falta columna obligatoria: {col}")
    return errores

def validar_tipos(df: pd.DataFrame):
    df = normalizar_columnas(df)
    errores = []
    for i, row in df.iterrows():
        fila_vacia = all(
            pd.isna(row.get(col)) or row.get(col) == "" for col in REQUIRED_COLUMNS + OPTIONAL_COLUMNS
        )
        if fila_vacia:
            continue

        # Validar obligatorias vacías
        for col in REQUIRED_COLUMNS:
            valor = row.get(col)
            if pd.isna(valor) or valor == "":
                errores.append(f"Fila {i+1}: Columna obligatoria '{col}' vacía")

        # Validar campos obligatorios de texto
        for col in [
            "PRODUCT_ID",
            "BRAND_NAME",
            "GENERIC_NAME",
            "CATEGORY_ID",
            "FAMILY",
            "LABORATORY_ID",
            "REF",
            "RX",
            "STATUS",
        ]:
            valor = row.get(col)
            if valor is not None and not pd.isna(valor) and valor != "":
                try:
                    str(valor)
                except Exception:
                    errores.append(f"Fila {i+1}: '{col}' debe ser texto")

        # Validar precios (numéricos)
        for col in ["PRICE_US", "PRICE_NATIONAL"]:
            valor = row.get(col)
            if valor is not None and not pd.isna(valor) and valor != "":
                try:
                    float(valor)
                except Exception:
                    errores.append(f"Fila {i+1}: '{col}' debe ser numérico")

        # Validar opcionales (si tienen valor, deben ser texto)
        for col in OPTIONAL_COLUMNS:
            valor = row.get(col)
            if valor is not None and not pd.isna(valor) and valor != "":
                try:
                    str(valor)
                except Exception:
                    errores.append(f"Fila {i+1}: '{col}' debe ser texto")
    return errores

# ======================
# PROGRAMA PRINCIPAL
# ======================
def main():
    # Ventana para seleccionar archivo
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="Selecciona el archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")]
    )

    if not file_path:
        print("❌ No se seleccionó ningún archivo.")
        return

    # Leer hoja Productos
    try:
        df_productos = pd.read_excel(file_path, sheet_name="Productos")
    except Exception as e:
        print(f"❌ Error al leer la hoja 'Productos': {e}")
        return

    # Validaciones
    errores_columnas = validar_columnas(df_productos)
    errores_tipos = validar_tipos(df_productos)

    # Mostrar resultados
    if errores_columnas or errores_tipos:
        print("\n⚠️ Se encontraron los siguientes errores:")
        for e in errores_columnas + errores_tipos:
            print("-", e)
        print("\nPuedes corregir el archivo o continuar bajo tu responsabilidad.")
    else:
        print("✅ El archivo pasó todas las validaciones.")

    # Aquí puedes poner lo que sigue en tu flujo,
    # aunque haya errores igual va a continuar.
    print("\n👉 Continuando con el proceso...")

if __name__ == "__main__":
    main()