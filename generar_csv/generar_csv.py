import pandas as pd
import os
import re
import unicodedata
from tkinter import Tk, filedialog

# Ocultar ventana principal de Tkinter
Tk().withdraw()

# Abrir diálogo para seleccionar archivo Excel
archivo_excel = filedialog.askopenfilename(
    title="Selecciona el archivo Excel",
    filetypes=[("Archivos Excel", "*.xlsx *.xls")]
)

# Validar si se seleccionó archivo
if not archivo_excel:
    print("❌ No seleccionaste ningún archivo.")
    exit()

try:
    # Leer la hoja "Productos"
    df_productos = pd.read_excel(archivo_excel, sheet_name="Productos")
    print("✅ Hoja 'Productos' encontrada y leída correctamente.")
except Exception as e:
    print("❌ Error al leer la hoja 'Productos':")
    print(e)
    exit()

# Columnas requeridas por Shopify
columnas_shopify = [
    "Handle","Title","Body (HTML)","Vendor","Product Category","Type","Tags","Published",
    "Option1 Name","Option1 Value","Option1 Linked To",
    "Option2 Name","Option2 Value","Option2 Linked To",
    "Option3 Name","Option3 Value","Option3 Linked To",
    "Variant SKU","Variant Grams","Variant Inventory Tracker","Variant Inventory Qty",
    "Variant Inventory Policy","Variant Fulfillment Service","Variant Price",
    "Variant Compare At Price","Variant Requires Shipping","Variant Taxable","Variant Barcode",
    "Image Src","Image Position","Image Alt Text","Gift Card","SEO Title","SEO Description",
    "Google Shopping / Google Product Category","Google Shopping / Gender","Google Shopping / Age Group",
    "Google Shopping / MPN","Google Shopping / Condition","Google Shopping / Custom Product",
    "Google Shopping / Custom Label 0","Google Shopping / Custom Label 1","Google Shopping / Custom Label 2",
    "Google Shopping / Custom Label 3","Google Shopping / Custom Label 4",
    "ATC (product.metafields.ficha_tecnica.atc)",
    "Cantidad PUM (product.metafields.ficha_tecnica.cantidad_pum)",
    "Condiciones Almacenamiento (product.metafields.ficha_tecnica.condiciones_almacenamiento)",
    "Condiciones de dispensacion (product.metafields.ficha_tecnica.condiciones_de_dispensacion)",
    "Consecutivo CUM (product.metafields.ficha_tecnica.consecutivo_cum)",
    "Contenido (product.metafields.ficha_tecnica.contenido)",
    "Cronico (product.metafields.ficha_tecnica.cronico)",
    "CUM (product.metafields.ficha_tecnica.cum)",
    "Descripción (product.metafields.ficha_tecnica.descripcion)",
    "Estado Registro Sanitario (product.metafields.ficha_tecnica.estado_registro_sanitario)",
    "Fecha de vencimiento Registro Sanitario (product.metafields.ficha_tecnica.fecha_de_vencimiento_registro_sanitario)",
    "Forma Farmacéutica (product.metafields.ficha_tecnica.forma_farmaceutica)",
    "Laboratorio (product.metafields.ficha_tecnica.laboratorio)",
    "Marca (product.metafields.ficha_tecnica.marca)",
    "Otros Principios Activos (product.metafields.ficha_tecnica.otros_principios_activos)",
    "Presentacion Comercial Cantidad (product.metafields.ficha_tecnica.presentacion_comercial_cantidad)",
    "Presentacion Comercial Embalaje (product.metafields.ficha_tecnica.presentacion_comercial_embalaje)",
    "Presentacion Comercial Interna (product.metafields.ficha_tecnica.presentacion_comercial_interna)",
    "Presentación PUM (product.metafields.ficha_tecnica.presentacion_pum)",
    "Principio Activo 1 (product.metafields.ficha_tecnica.principio_activo_1)",
    "Principio Activo 2 (product.metafields.ficha_tecnica.principio_activo_2)",
    "Registro Sanitario (product.metafields.ficha_tecnica.registro_sanitario)",
    "RX (product.metafields.ficha_tecnica.rx)",
    "SKU (product.metafields.ficha_tecnica.sku)",
    "Vencimiento Registro Sanitario (product.metafields.ficha_tecnica.vencimiento_registro_sanitario)",
    "Via Administracion (product.metafields.ficha_tecnica.via_administracion)",
    "Google: Custom Product (product.metafields.mm-google-shopping.custom_product)",
    "Grupo de edad (product.metafields.shopify.age-group)",
    "Tipo de frasco (product.metafields.shopify.bottle-type)",
    "Ingredientes detallados (product.metafields.shopify.detailed-ingredients)",
    "Preferencias alimentarias (product.metafields.shopify.dietary-preferences)",
    "Sabor (product.metafields.shopify.flavor)",
    "Tipo de solución (product.metafields.shopify.solution-type)",
    "Productos complementarios (product.metafields.shopify--discovery--product_recommendation.complementary_products)",
    "Productos relacionados (product.metafields.shopify--discovery--product_recommendation.related_products)",
    "Configuración de productos relacionados (product.metafields.shopify--discovery--product_recommendation.related_products_display)",
    "Buscar impulsos de productos (product.metafields.shopify--discovery--product_search_boost.queries)",
    "Variant Image","Variant Weight Unit","Variant Tax Code","Cost per item","Status"
]

# Creamos DataFrame vacío con columnas
df_csv = pd.DataFrame(columns=columnas_shopify)

# Función para limpiar y formatear Handle
def crear_handle(brand, product_id):
    if pd.isna(brand) or pd.isna(product_id):
        return ""
    handle = str(brand).lower()
    handle = handle.replace('ñ', 'n')
    handle = unicodedata.normalize('NFKD', handle).encode('ASCII', 'ignore').decode('utf-8')
    handle = re.sub(r'[^\w]+', '-', handle)
    handle = re.sub(r'-+', '-', handle)
    handle = handle.strip('-')
    handle = f"{handle}-{product_id}"
    return handle

# Llenamos la columna Handle en df_csv
df_csv['Handle'] = df_productos.apply(
    lambda row: crear_handle(row.get('BRAND NAME'), row.get('PRODUCT ID')), axis=1
)

# Llenamos la columna Title con BRAND NAME directamente
df_csv['Title'] = df_productos['BRAND NAME']

# --- COLUMNA BODY (HTML) ---
def crear_body(row):
    partes = []
    if not pd.isna(row.get('GENERIC NAME')):
        partes.append(f"<p>{row['GENERIC NAME']}</p>")
    if not pd.isna(row.get('PRESENTATION')):
        partes.append(f"<p>{row['PRESENTATION']}</p>")
    if not pd.isna(row.get('LABORATORY ID')):
        partes.append(f"<p>{row['LABORATORY ID']}</p>")
    if not pd.isna(row.get('REF')):
        partes.append(f"<p>{row['REF']}</p>")

    # buscar RX aunque se llame distinto en el Excel
    posibles_rx = [col for col in row.index if "rx" in col.lower()]
    if posibles_rx:
        col_rx = posibles_rx[0]
        valor_rx = row[col_rx]
        if not pd.isna(valor_rx):
            valor_rx_str = str(valor_rx).strip().upper()
            if valor_rx_str in ["TRUE", "VERDADERO"]:
                partes.append("<p>Necesita receta médica</p>")
    return "".join(partes)

df_csv['Body (HTML)'] = df_productos.apply(crear_body, axis=1)

# --- COLUMNAS VENDOR Y TYPE (PAIS) ---
# Lista de farmacias conocidas y sus países
farmacias = {
    "ALSA QUERETARO": "MEXICO",
    "JI COHEN": "GUATEMALA",
    "FARMACIA BRASIL": "SALVADOR",
    "FARMACIA BATRES": "GUATEMALA",
    "FARMACIA FARMAGO": "VENEZUELA"
}

# Mostrar opciones al usuario
print("Farmacias disponibles:")
for f in farmacias:
    print(f"- {f}")

nombre_farmacia = input("Ingresa el nombre de la farmacia (tal cual o nueva si no está en la lista): ").strip()

# Si la farmacia no está en la lista, pedimos el país
if nombre_farmacia not in farmacias:
    pais_farmacia = input(f"Ingrese el país de la farmacia '{nombre_farmacia}': ").strip()
else:
    pais_farmacia = farmacias[nombre_farmacia]

# Llenar las columnas Vendor y Type
df_csv['Vendor'] = nombre_farmacia
df_csv['Type'] = pais_farmacia

# --- COLUMNA TAGS ---
try:
    # Leer las hojas Categoria y Familia
    df_categoria = pd.read_excel(archivo_excel, sheet_name="Categoria")
    df_familia = pd.read_excel(archivo_excel, sheet_name="Familia")
except Exception as e:
    print("❌ Error al leer las hojas 'Categoria' o 'Familia':", e)
    exit()

# Convertimos a diccionarios {id: name} para fácil búsqueda
dict_categoria = pd.Series(df_categoria['name'].values, index=df_categoria['id']).to_dict()
dict_familia   = pd.Series(df_familia['name'].values, index=df_familia['id']).to_dict()

# --- COLUMNA TAGS ---
try:
    # Leer las hojas Categoria y Familia
    df_categoria = pd.read_excel(archivo_excel, sheet_name="Categoria")
    df_familia = pd.read_excel(archivo_excel, sheet_name="Familia")
except Exception as e:
    print("❌ Error al leer las hojas 'Categoria' o 'Familia':", e)
    exit()

# Convertimos a diccionarios {id: name} para fácil búsqueda
dict_categoria = pd.Series(df_categoria['name'].values, index=df_categoria['id']).to_dict()
dict_familia   = pd.Series(df_familia['name'].values, index=df_familia['id']).to_dict()

def generar_tags(row, pais):
    tags = []

    # --- Familia primero ---
    fam_id = row.get('FAMILY')
    if pd.notna(fam_id) and fam_id in dict_familia:
        fam_name = dict_familia[fam_id]
        fam_parts = [part.strip() for part in fam_name.split(',')]
        # 🔹 país al final
        fam_with_country = ','.join([f"{part} {pais.capitalize()}" for part in fam_parts])
        tags.append(fam_with_country)

    # --- Categoría después ---
    cat_id = row.get('CATEGORY ID')
    if pd.notna(cat_id) and cat_id in dict_categoria:
        cat_name = dict_categoria[cat_id]
        cat_parts = [part.strip() for part in cat_name.split(',')]
        # 🔹 país al final
        cat_with_country = ','.join([f"{part} {pais.capitalize()}" for part in cat_parts])
        tags.append(cat_with_country)

    # --- País y Farmacia al final ---
    tags.append(pais.capitalize())
    tags.append(f"Farmacia {pais.capitalize()}")

    return ', '.join(tags)

# Llenamos la columna Tags en df_csv
df_csv['Tags'] = df_productos.apply(lambda row: generar_tags(row, pais_farmacia), axis=1)

# --- COLUMNA PUBLISHED ---
df_csv['Published'] = "TRUE"

# ==============================
# Columnas Option1 Name y Option1 Value
# ==============================
df_csv["Option1 Name"] = "Title"
df_csv["Option1 Value"] = "Default Title"

# ==============================
# Columnas OptionX vacías
# ==============================
df_csv["Option1 Linked To"] = ""
df_csv["Option2 Name"] = ""
df_csv["Option2 Value"] = ""
df_csv["Option2 Linked To"] = ""
df_csv["Option3 Name"] = ""
df_csv["Option3 Value"] = ""
df_csv["Option3 Linked To"] = ""

# ==============================
# Columna Variant SKU desde REF
# ==============================
if "REF" in df_productos.columns:
    df_csv["Variant SKU"] = df_productos["REF"]
else:
    df_csv["Variant SKU"] = ""
    print("⚠️ Advertencia: No se encontró la columna 'REF' en la hoja 'Productos'. Se dejará Variant SKU vacío.")

# ==============================
# Columnas Variant Grams y Variant Inventory Tracker
# ==============================
df_csv['Variant Grams'] = 0
df_csv['Variant Inventory Tracker'] = "shopify"

# ==============================
# Llenar Variant Inventory Qty desde la columna INVENTORY del Excel
# ==============================
if 'INVENTORY' in df_productos.columns:
    df_csv['Variant Inventory Qty'] = df_productos['INVENTORY'].fillna(0)
else:
    # Si no existe la columna INVENTORY, dejar en 0
    df_csv['Variant Inventory Qty'] = 0

# ==============================
# Llenar Variant Inventory Policy y Variant Fulfillment Service
# ==============================
df_csv['Variant Inventory Policy'] = "continue"   # Valor fijo para toda la columna
df_csv['Variant Fulfillment Service'] = "manual"  # Valor fijo para toda la columna

# ==============================
# Llenar Variant Price desde PRICE US
# ==============================
df_csv['Variant Price'] = df_productos['PRICE US']

# ==============================
# Columnas Variant Compare At Price, Variant Requires Shipping, Variant Taxable, Variant Barcode
# ==============================
df_csv['Variant Compare At Price'] = ""
df_csv['Variant Requires Shipping'] = "TRUE"
df_csv['Variant Taxable'] = "TRUE"
df_csv['Variant Barcode'] = ""

# ==============================
# Columnas Image Src, Image Position, Image Alt Text, Gift Card
# ==============================
df_csv['Image Src'] = ""
df_csv['Image Position'] = ""
df_csv['Image Alt Text'] = ""
df_csv['Gift Card'] = "FALSE"

# ==============================
# Columnas vacías individuales
# ==============================
df_csv["SEO Title"] = ""
df_csv["SEO Description"] = ""
df_csv["Google Shopping / Google Product Category"] = ""
df_csv["Google Shopping / Gender"] = ""
df_csv["Google Shopping / Age Group"] = ""
df_csv["Google Shopping / MPN"] = ""
df_csv["Google Shopping / Condition"] = ""
df_csv["Google Shopping / Custom Product"] = ""
df_csv["Google Shopping / Custom Label 0"] = ""
df_csv["Google Shopping / Custom Label 1"] = ""
df_csv["Google Shopping / Custom Label 2"] = ""
df_csv["Google Shopping / Custom Label 3"] = ""
df_csv["Google Shopping / Custom Label 4"] = ""
df_csv["ATC (product.metafields.ficha_tecnica.atc)"] = ""
df_csv["Cantidad PUM (product.metafields.ficha_tecnica.cantidad_pum)"] = ""
df_csv["Condiciones Almacenamiento (product.metafields.ficha_tecnica.condiciones_almacenamiento)"] = ""
df_csv["Condiciones de dispensacion (product.metafields.ficha_tecnica.condiciones_de_dispensacion)"] = ""
df_csv["Consecutivo CUM (product.metafields.ficha_tecnica.consecutivo_cum)"] = ""
df_csv["Contenido (product.metafields.ficha_tecnica.contenido)"] = ""
df_csv["Cronico (product.metafields.ficha_tecnica.cronico)"] = ""
df_csv["CUM (product.metafields.ficha_tecnica.cum)"] = ""
df_csv["Descripción (product.metafields.ficha_tecnica.descripcion)"] = ""
df_csv["Estado Registro Sanitario (product.metafields.ficha_tecnica.estado_registro_sanitario)"] = ""
df_csv["Fecha de vencimiento Registro Sanitario (product.metafields.ficha_tecnica.fecha_de_vencimiento_registro_sanitario)"] = ""
df_csv["Forma Farmacéutica (product.metafields.ficha_tecnica.forma_farmaceutica)"] = ""
df_csv["Laboratorio (product.metafields.ficha_tecnica.laboratorio)"] = ""
df_csv["Marca (product.metafields.ficha_tecnica.marca)"] = ""
df_csv["Otros Principios Activos (product.metafields.ficha_tecnica.otros_principios_activos)"] = ""
df_csv["Presentacion Comercial Cantidad (product.metafields.ficha_tecnica.presentacion_comercial_cantidad)"] = ""
df_csv["Presentacion Comercial Embalaje (product.metafields.ficha_tecnica.presentacion_comercial_embalaje)"] = ""
df_csv["Presentacion Comercial Interna (product.metafields.ficha_tecnica.presentacion_comercial_interna)"] = ""
df_csv["Presentación PUM (product.metafields.ficha_tecnica.presentacion_pum)"] = ""
df_csv["Principio Activo 1 (product.metafields.ficha_tecnica.principio_activo_1)"] = ""
df_csv["Principio Activo 2 (product.metafields.ficha_tecnica.principio_activo_2)"] = ""
df_csv["Registro Sanitario (product.metafields.ficha_tecnica.registro_sanitario)"] = ""
df_csv["RX (product.metafields.ficha_tecnica.rx)"] = ""
df_csv["SKU (product.metafields.ficha_tecnica.sku)"] = ""
df_csv["Vencimiento Registro Sanitario (product.metafields.ficha_tecnica.vencimiento_registro_sanitario)"] = ""
df_csv["Via Administracion (product.metafields.ficha_tecnica.via_administracion)"] = ""
df_csv["Google: Custom Product (product.metafields.mm-google-shopping.custom_product)"] = ""
df_csv["Grupo de edad (product.metafields.shopify.age-group)"] = ""
df_csv["Tipo de frasco (product.metafields.shopify.bottle-type)"] = ""
df_csv["Ingredientes detallados (product.metafields.shopify.detailed-ingredients)"] = ""
df_csv["Preferencias alimentarias (product.metafields.shopify.dietary-preferences)"] = ""
df_csv["Sabor (product.metafields.shopify.flavor)"] = ""
df_csv["Tipo de solución (product.metafields.shopify.solution-type)"] = ""
df_csv["Productos complementarios (product.metafields.shopify--discovery--product_recommendation.complementary_products)"] = ""
df_csv["Productos relacionados (product.metafields.shopify--discovery--product_recommendation.related_products)"] = ""
df_csv["Configuración de productos relacionados (product.metafields.shopify--discovery--product_recommendation.related_products_display)"] = ""
df_csv["Buscar impulsos de productos (product.metafields.shopify--discovery--product_search_boost.queries)"] = ""
df_csv["Variant Image"] = ""

# ==============================
# Columnas finales con valores fijos o vacíos
# ==============================
df_csv["Variant Weight Unit"] = "kg"
df_csv["Variant Tax Code"] = ""
df_csv["Cost per item"] = ""

# ==============================
# Columna Status
# ==============================
df_csv["Status"] = df_productos["STATUS"]


# Nombre de salida
nombre_salida = os.path.splitext(os.path.basename(archivo_excel))[0] + "_shopify.csv"

# Guardamos CSV
df_csv.to_csv(nombre_salida, index=False, encoding="utf-8-sig")

print(f"✅ Archivo '{nombre_salida}' creado con Handle, Title, Body, Vendor y Type correctamente llenos.")