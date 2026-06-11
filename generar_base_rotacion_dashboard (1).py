# -*- coding: utf-8 -*-
"""
Genera Excel base para rotación de inventario y alertas.

Versión corregida:
- Ventas finales = Odoo + Autoazur.
- Odoo toma cotizaciones/ventas solo si Tipo de venta = Full o Drop.
- Autoazur y Odoo se complementan:
    - Si la misma referencia/pedido aparece en ambas fuentes, se conserva Odoo.
    - Si aparece solo en Autoazur, sí cuenta como venta.
- Inventario Odoo toma solo CUATI/Existencias, ID 125.
- Stock marketplaces: Walmart, Liverpool 99MIN, Mercado Libre Full, Amazon FBA.
- Diccionario origen4:
    - Lee aliases de todas las hojas y columnas.
    - Acepta SKUs con números, con guiones, IQ y también SKUs de texto como outspeakblue.
- Match en capas:
    1) sku_original / sku_default_code / sku_desde_columna / sku_desde_titulo
    2) variante sin ceros a la izquierda: 05024173182 = 5024173182
    3) referencia Autoazur -> referencia en Odoo -> SKU Odoo -> IQ
"""

from pathlib import Path
import os
import re
import unicodedata
import xmlrpc.client
import warnings

import pandas as pd
import numpy as np


warnings.filterwarnings("ignore", category=UserWarning)


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

#DESKTOP = Path.home() / "Desktop"
#CARPETA_SALIDA = DESKTOP / "rotacion_inventario_base_dashboard"
DESKTOP = Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")
CARPETA_SALIDA =  Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

ARCHIVO_SALIDA = CARPETA_SALIDA / "rotacion_inventario_base_dashboard_odoo_autoazur.xlsx"

# ============================================================
# ARCHIVO AUTOAZUR ACTUALIZADO
# ============================================================
# Aquí puedes poner el nombre exacto de tu archivo de ventas Autoazur actualizado.
# El código lo buscará primero en Desktop y luego en Downloads.
#
# Ejemplo actual:
#   VentasAZ10JUN.xlsx
#
# Si lo dejas vacío, volverá a buscar por patrones:
#   "listado detallado" + "pedidos"
ARCHIVO_AUTOAZUR_ACTUALIZADO = "VentasAZ11
JUN.xlsx"

FECHA_INICIO = pd.Timestamp("2025-11-01")
FECHA_FIN = pd.Timestamp.today().normalize()


# ============================================================
# CONFIGURACIÓN ODOO - AJUSTA ESTA SECCIÓN
# ============================================================

ENABLE_ODOO = True

ODOO_URL = os.getenv("ODOO_URL", "https://comercializadora-iqtech-productos-innovadores.odoo.com")
ODOO_DB = os.getenv("ODOO_DB", "comercializadora-iqtech-productos-innovadores-sh-ma-27238691")
ODOO_USER = os.getenv("ODOO_USER", "hocampou@gmail.com")
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "d4340db6fc78dcba7c75bda7ba5fe2f5e6d57347")

# En tu Odoo ya se detectó este campo:
ODOO_CAMPO_TIPO_VENTA = os.getenv("ODOO_CAMPO_TIPO_VENTA", "x_studio_tipo_de_venta")

# Solo se toman ventas/cotizaciones con estos tipos.
ODOO_TIPOS_VENTA_VALIDOS = ["full", "drop"]

# Inventario Odoo. Ya detectamos CUATI/Existencias = ID 125.
ODOO_LOCATION_IDS = [125]

# Si algún día quieres buscar ubicación por nombre, deja ODOO_LOCATION_IDS = [].
ODOO_LOCATION_KEYWORDS = ["CUATI", "Existencias"]

# available = quantity - reserved_quantity
# quantity = cantidad física total
ODOO_STOCK_MODE = "available"


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def normalizar_texto(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    x = unicodedata.normalize("NFKD", x)
    x = "".join(c for c in x if not unicodedata.combining(c))
    x = x.lower()
    x = re.sub(r"\s+", " ", x)
    return x


def limpiar_sku(x):
    if pd.isna(x):
        return ""
    x = str(x).strip()
    if x.lower() in ["nan", "none", ""]:
        return ""
    if re.fullmatch(r"\d+\.0", x):
        x = x[:-2]
    return x.strip()


def sku_key(x):
    """
    Llave SKU general:
    - mayúsculas
    - sin espacios
    - conserva guiones
    """
    return limpiar_sku(x).upper().replace(" ", "")


def sku_key_sin_ceros(x):
    """
    Quita ceros a la izquierda solo si el SKU es numérico puro.

    Ejemplo:
    05024173182 -> 5024173182
    """
    k = sku_key(x)

    if re.fullmatch(r"\d+", k):
        return k.lstrip("0") or "0"

    return k


def generar_sku_keys_match(x):
    """
    Variantes para match por SKU.
    Ejemplo:
    05024173182 -> ["05024173182", "5024173182"]
    outspeakblue -> ["OUTSPEAKBLUE"]
    """
    base = sku_key(x)
    sin_ceros = sku_key_sin_ceros(x)

    keys = []

    if base:
        keys.append(base)

    if sin_ceros and sin_ceros not in keys:
        keys.append(sin_ceros)

    return keys


def referencia_key(x):
    """
    Normaliza referencias de pedidos/canales para cruzar Autoazur contra Odoo.

    Quita todo lo que no sea letra o número.
    Ejemplo:
    LIV - 2950099473 -> LIV2950099473
    2950099473 -> 2950099473
    """
    if pd.isna(x):
        return ""
    x = str(x).strip().upper()
    if x.lower() in ["nan", "none", ""]:
        return ""
    x = re.sub(r"[^A-Z0-9]", "", x)
    return x


def referencia_variantes(x):
    """
    Genera variantes de referencia para encontrar coincidencias entre Autoazur y Odoo.
    Incluye:
    - texto limpio completo
    - solo números
    - números sin ceros a la izquierda
    """
    base = referencia_key(x)
    keys = []

    if base:
        keys.append(base)

    nums = re.sub(r"\D", "", base)

    if nums:
        keys.append(nums)
        keys.append(nums.lstrip("0") or "0")

    out = []
    for k in keys:
        if k and k not in out:
            out.append(k)

    return out


def to_number(s):
    if isinstance(s, pd.Series):
        return pd.to_numeric(
            s.astype(str)
            .str.replace(",", "", regex=False)
            .str.replace("$", "", regex=False)
            .str.replace("MXN", "", regex=False)
            .str.strip(),
            errors="coerce"
        ).fillna(0)

    try:
        return float(str(s).replace(",", "").replace("$", "").strip())
    except Exception:
        return 0


def encontrar_columna(df, posibles_nombres):
    mapa = {normalizar_texto(c): c for c in df.columns}

    for nombre in posibles_nombres:
        n = normalizar_texto(nombre)
        if n in mapa:
            return mapa[n]

    for col_norm, col_real in mapa.items():
        for nombre in posibles_nombres:
            n = normalizar_texto(nombre)
            if n and n in col_norm:
                return col_real

    raise KeyError(
        f"No encontré columna {posibles_nombres}. "
        f"Columnas disponibles: {list(df.columns)}"
    )


def encontrar_columna_opcional(df, posibles_nombres):
    try:
        return encontrar_columna(df, posibles_nombres)
    except Exception:
        return None


def encontrar_archivo(carpeta, patrones, requerido=True):
    archivos = [p for p in carpeta.iterdir() if p.is_file()]

    for p in archivos:
        nombre = normalizar_texto(p.name)
        if all(normalizar_texto(patron) in nombre for patron in patrones):
            return p

    if requerido:
        raise FileNotFoundError(
            f"No encontré archivo con patrones {patrones} en {carpeta}"
        )

    return None


def encontrar_archivo_autoazur_actualizado():
    """
    Busca primero el archivo Autoazur actualizado por nombre exacto.
    Prioridad:
      1) Desktop/VentasAZ10JUN.xlsx
      2) Downloads/VentasAZ10JUN.xlsx
      3) Desktop/rotacion_inventario_base_dashboard/VentasAZ10JUN.xlsx
      4) Búsqueda por patrones en Desktop: listado detallado + pedidos
      5) Búsqueda por patrones en Downloads: listado detallado + pedidos
    """
    candidatos_carpetas = [
        DESKTOP,
        Path.home() / "Downloads",
        CARPETA_SALIDA,
    ]

    if ARCHIVO_AUTOAZUR_ACTUALIZADO:
        for carpeta in candidatos_carpetas:
            ruta = carpeta / ARCHIVO_AUTOAZUR_ACTUALIZADO
            if ruta.exists():
                return ruta

    for carpeta in candidatos_carpetas:
        try:
            ruta = encontrar_archivo(
                carpeta,
                ["listado detallado", "pedidos"],
                requerido=False
            )
            if ruta:
                return ruta
        except Exception:
            pass

    return None


def preparar_para_excel(df):
    df = df.copy()

    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            try:
                df[col] = df[col].dt.tz_localize(None)
            except Exception:
                pass

    return df


def m2o_id(value):
    if isinstance(value, (list, tuple)) and value:
        return value[0]
    return None


def m2o_name(value):
    if isinstance(value, (list, tuple)) and len(value) > 1:
        return value[1]
    return ""


# ============================================================
# CLIENTE ODOO
# ============================================================

class OdooClient:
    def __init__(self, url, db, user, api_key):
        self.url = url.rstrip("/")
        self.db = db
        self.user = user
        self.api_key = api_key
        self.uid = None
        self.models = None

    def connect(self):
        if (
            "TU-ODOO" in self.url
            or self.db.startswith("TU_")
            or self.user.startswith("TU_")
            or self.api_key.startswith("TU_")
        ):
            raise ValueError(
                "Faltan credenciales de Odoo. Ajusta ODOO_URL, ODOO_DB, "
                "ODOO_USER y ODOO_API_KEY."
            )

        common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self.uid = common.authenticate(self.db, self.user, self.api_key, {})

        if not self.uid:
            raise ConnectionError(
                "No se pudo autenticar en Odoo. Revisa URL, DB, usuario y API key."
            )

        self.models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        print(f"Conectado a Odoo. UID: {self.uid}")

    def execute(self, model, method, *args, **kwargs):
        return self.models.execute_kw(
            self.db,
            self.uid,
            self.api_key,
            model,
            method,
            args,
            kwargs
        )

    def search_read_all(self, model, domain, fields, batch=1000, order=None):
        out = []
        offset = 0

        while True:
            kwargs = {
                "fields": fields,
                "limit": batch,
                "offset": offset,
            }

            if order:
                kwargs["order"] = order

            rows = self.execute(model, "search_read", domain, **kwargs)

            if not rows:
                break

            out.extend(rows)

            if len(rows) < batch:
                break

            offset += batch

        return out


# ============================================================
# ODOO: CAMPO TIPO DE VENTA
# ============================================================

def descubrir_campo_tipo_venta(odoo):
    if ODOO_CAMPO_TIPO_VENTA:
        print(f"Usando campo Tipo de venta definido: {ODOO_CAMPO_TIPO_VENTA}")
        return ODOO_CAMPO_TIPO_VENTA

    fields = odoo.search_read_all(
        "ir.model.fields",
        [("model", "=", "sale.order")],
        ["name", "field_description", "ttype", "relation"],
        batch=3000,
        order="name asc"
    )

    candidatos = []

    for f in fields:
        name_norm = normalizar_texto(f.get("name", ""))
        desc_norm = normalizar_texto(f.get("field_description", ""))
        texto = f"{name_norm} {desc_norm}"

        if (
            "tipo de venta" in texto
            or "tipo venta" in texto
            or "sale type" in texto
            or "sales type" in texto
            or "x_studio_tipo" in texto
            or "x_tipo" in texto
        ):
            candidatos.append(f)

    if not candidatos:
        ruta = CARPETA_SALIDA / "odoo_campos_sale_order_para_buscar_tipo_venta.xlsx"
        pd.DataFrame(fields).to_excel(ruta, index=False)

        raise ValueError(
            "No pude detectar automáticamente el campo técnico de 'Tipo de venta'.\n"
            f"Exporté todos los campos de sale.order aquí:\n{ruta}\n"
            "Busca 'Tipo de venta' y copia el valor de la columna name."
        )

    print("Candidatos detectados para Tipo de venta en sale.order:")

    for c in candidatos[:20]:
        print(
            f"- {c.get('name')} | "
            f"{c.get('field_description')} | "
            f"{c.get('ttype')}"
        )

    elegido = candidatos[0]

    print(
        f"Usando campo Tipo de venta: "
        f"{elegido.get('name')} ({elegido.get('field_description')})"
    )

    return elegido.get("name")


# ============================================================
# ODOO: EXTRAER VENTAS
# ============================================================

def extraer_ventas_odoo(odoo):
    print("Descargando ventas desde Odoo...")

    campo_tipo_venta = descubrir_campo_tipo_venta(odoo)

    dt_ini = FECHA_INICIO.strftime("%Y-%m-%d 00:00:00")
    dt_fin = (FECHA_FIN + pd.Timedelta(days=1)).strftime("%Y-%m-%d 00:00:00")

    # No se filtra por estado: permite draft/cotización y sale/venta.
    # Luego se filtra por Tipo de venta = Full o Drop.
    domain = [
        ("order_id.date_order", ">=", dt_ini),
        ("order_id.date_order", "<", dt_fin),
        ("display_type", "=", False),
        ("product_id", "!=", False),
    ]

    fields_line = [
        "id",
        "order_id",
        "product_id",
        "name",
        "product_uom_qty",
        "qty_delivered",
        "price_total",
        "price_unit",
    ]

    lines = odoo.search_read_all(
        "sale.order.line",
        domain,
        fields_line,
        batch=2000,
        order="id asc"
    )

    if not lines:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    product_ids = sorted({
        m2o_id(x.get("product_id"))
        for x in lines
        if m2o_id(x.get("product_id"))
    })

    order_ids = sorted({
        m2o_id(x.get("order_id"))
        for x in lines
        if m2o_id(x.get("order_id"))
    })

    product_fields = [
        "id",
        "display_name",
        "default_code",
        "barcode",
        "categ_id",
    ]

    products = odoo.execute(
        "product.product",
        "read",
        product_ids,
        product_fields
    ) if product_ids else []

    product_map = {p["id"]: p for p in products}

    order_fields = [
        "id",
        "name",
        "date_order",
        "state",
        "client_order_ref",
        "origin",
        "partner_id",
        "team_id",
        "warehouse_id",
        campo_tipo_venta,
    ]

    orders = odoo.execute(
        "sale.order",
        "read",
        order_ids,
        order_fields
    ) if order_ids else []

    order_map = {o["id"]: o for o in orders}

    rows = []
    rows_sin_sku = []

    for l in lines:
        order_id = m2o_id(l.get("order_id"))
        product_id = m2o_id(l.get("product_id"))

        o = order_map.get(order_id, {})
        p = product_map.get(product_id, {})

        tipo_raw = o.get(campo_tipo_venta, "")

        if isinstance(tipo_raw, (list, tuple)):
            tipo_venta = m2o_name(tipo_raw)
        else:
            tipo_venta = str(tipo_raw or "")

        tipo_venta_norm = normalizar_texto(tipo_venta)

        if tipo_venta_norm not in ODOO_TIPOS_VENTA_VALIDOS:
            continue

        default_code = limpiar_sku(p.get("default_code", ""))
        barcode = limpiar_sku(p.get("barcode", ""))

        sku = default_code or barcode

        row = {
            "fuente": "ODOO_API",
            "fecha": pd.to_datetime(o.get("date_order"), errors="coerce"),
            "pedido": o.get("name", ""),
            "estado": o.get("state", ""),
            "estado_odoo": o.get("state", ""),
            "tipo_venta": tipo_venta,
            "canal": tipo_venta,
            "cliente": m2o_name(o.get("partner_id")),
            "equipo_ventas": m2o_name(o.get("team_id")),
            "almacen": m2o_name(o.get("warehouse_id")),
            "referencia_cliente": o.get("client_order_ref", ""),
            "referencia": o.get("client_order_ref", ""),
            "item_id": "",
            "origen": o.get("origin", ""),
            "line_id": l.get("id"),
            "product_id": product_id,
            "producto": p.get("display_name", l.get("name", "")),
            "sku_original": sku,
            "sku_desde_columna": default_code,
            "sku_desde_titulo": "",
            "sku_default_code": default_code,
            "barcode": barcode,
            "categoria": m2o_name(p.get("categ_id")),
            "cantidad": float(l.get("product_uom_qty") or 0),
            "cantidad_entregada": float(l.get("qty_delivered") or 0),
            "precio_unitario": float(l.get("price_unit") or 0),
            "venta_total": float(l.get("price_total") or 0),
        }

        if not sku:
            rows_sin_sku.append(row)
        else:
            rows.append(row)

    ventas = pd.DataFrame(rows)
    ventas_sin_sku = pd.DataFrame(rows_sin_sku)
    productos_df = pd.DataFrame(products)

    print(f"Ventas Odoo descargadas con SKU: {len(ventas)}")
    print(f"Ventas Odoo sin SKU: {len(ventas_sin_sku)}")

    return ventas, productos_df, ventas_sin_sku


# ============================================================
# ODOO: INVENTARIO
# ============================================================

def buscar_ubicaciones_odoo(odoo):
    if ODOO_LOCATION_IDS:
        ubicaciones = odoo.execute(
            "stock.location",
            "read",
            ODOO_LOCATION_IDS,
            ["id", "name", "complete_name", "usage"]
        )

        print("Ubicaciones Odoo usadas para inventario:")

        for u in ubicaciones:
            print(f"- ID {u.get('id')} | {u.get('complete_name')}")

        return pd.DataFrame(ubicaciones), ODOO_LOCATION_IDS

    domain = [("usage", "=", "internal")]

    ubicaciones = odoo.search_read_all(
        "stock.location",
        domain,
        ["id", "name", "complete_name", "usage"],
        batch=2000,
        order="complete_name asc"
    )

    ubicaciones_df = pd.DataFrame(ubicaciones)

    if ubicaciones_df.empty:
        return ubicaciones_df, []

    mask = pd.Series(False, index=ubicaciones_df.index)

    for kw in ODOO_LOCATION_KEYWORDS:
        kw_norm = normalizar_texto(kw)

        mask = mask | ubicaciones_df["complete_name"].apply(
            lambda x: kw_norm in normalizar_texto(x)
        )

    filtradas = ubicaciones_df[mask].copy()

    if filtradas.empty:
        ruta = CARPETA_SALIDA / "odoo_ubicaciones_internas_para_buscar_cuautitlan.xlsx"
        ubicaciones_df.to_excel(ruta, index=False)

        raise ValueError(
            "No encontré ubicaciones de Odoo con los keywords configurados.\n"
            f"Exporté las ubicaciones internas aquí:\n{ruta}"
        )

    location_ids = filtradas["id"].astype(int).tolist()

    print("Ubicaciones Odoo detectadas para inventario:")

    for _, r in filtradas.iterrows():
        print(f"- ID {r['id']} | {r['complete_name']}")

    return filtradas, location_ids


def extraer_inventario_odoo(odoo):
    print("Descargando inventario desde Odoo...")

    ubicaciones_df, location_ids = buscar_ubicaciones_odoo(odoo)

    if not location_ids:
        return pd.DataFrame(), ubicaciones_df

    domain = [
        ("location_id", "in", location_ids),
        ("product_id", "!=", False),
    ]

    quant_fields = [
        "id",
        "product_id",
        "location_id",
        "quantity",
        "reserved_quantity",
    ]

    quants = odoo.search_read_all(
        "stock.quant",
        domain,
        quant_fields,
        batch=3000,
        order="id asc"
    )

    if not quants:
        return pd.DataFrame(), ubicaciones_df

    product_ids = sorted({
        m2o_id(q.get("product_id"))
        for q in quants
        if m2o_id(q.get("product_id"))
    })

    products = odoo.execute(
        "product.product",
        "read",
        product_ids,
        ["id", "display_name", "default_code", "barcode", "categ_id"]
    ) if product_ids else []

    product_map = {p["id"]: p for p in products}

    rows = []

    for q in quants:
        product_id = m2o_id(q.get("product_id"))
        p = product_map.get(product_id, {})

        qty = float(q.get("quantity") or 0)
        reserved = float(q.get("reserved_quantity") or 0)

        if ODOO_STOCK_MODE == "available":
            stock = qty - reserved
        else:
            stock = qty

        default_code = limpiar_sku(p.get("default_code", ""))
        barcode = limpiar_sku(p.get("barcode", ""))

        sku = default_code or barcode

        if not sku:
            continue

        rows.append({
            "canal_stock": "ODOO_CUAUTITLAN",
            "sku_original": sku,
            "producto_stock": p.get("display_name", ""),
            "stock": stock,
            "fuente_archivo": "ODOO_API_STOCK_QUANT",
            "odoo_product_id": product_id,
            "odoo_location_id": m2o_id(q.get("location_id")),
            "odoo_location": m2o_name(q.get("location_id")),
            "quantity_odoo": qty,
            "reserved_quantity_odoo": reserved,
            "categoria_odoo": m2o_name(p.get("categ_id")),
        })

    stock_odoo = pd.DataFrame(rows)

    print(f"Renglones inventario Odoo: {len(stock_odoo)}")

    return stock_odoo, ubicaciones_df


# ============================================================
# DICCIONARIO ORIGEN4 ROBUSTO
# ============================================================

def parece_sku_madre(x):
    x = limpiar_sku(x).upper()
    return bool(re.fullmatch(r"IQ\d+", x))


def parece_alias_valido(x):
    """
    Permite aliases numéricos, alfanuméricos, con guión y también texto corto
    como outspeakblue, XiaomiPocket, BoseSoundLinkBlanco.
    """
    x = limpiar_sku(x)

    if not x:
        return False

    x_norm = normalizar_texto(x)
    x_up = x.upper().strip()

    descartes_exactos = {
        "CON SKU",
        "SIN SKU",
        "SIN SKU EN",
        "SKU MADRE",
        "AMAZON",
        "MERCADO LIBRE",
        "LIVERPOOL",
        "WALMART",
        "COPPEL",
        "ELEKTRA",
        "CANAL",
        "PRODUCTO",
        "NOMBRE",
        "MARCA",
        "FALSE",
        "TRUE",
        "FALSO",
        "VERDADERO",
        "SI",
        "NO",
        "N/A",
    }

    if x_up in descartes_exactos:
        return False

    if x_norm in ["nan", "none", "falso", "false", "verdadero", "true", "sin sku", "con sku"]:
        return False

    if "sin sku" in x_norm or "con sku" in x_norm:
        return False

    # Evita descripciones largas completas, pero permite SKUs texto cortos.
    if len(x) > 60:
        return False

    # Evita valores de una sola letra.
    if len(x) <= 1:
        return False

    # Descarta frases largas con muchos espacios.
    if x.count(" ") >= 3:
        return False

    # Si es IQ, siempre.
    if re.fullmatch(r"IQ\d+", x_up):
        return True

    # Si tiene números, suele ser SKU.
    if re.search(r"\d", x):
        return True

    # Si tiene guión o guion bajo.
    if "-" in x or "_" in x:
        return True

    # NUEVO: SKUs de texto tipo outspeakblue, XiaomiPocket, BoseSoundLinkBlanco
    # y aliases con diagonal como Wave/VibeBeamNegro o Wave/VibeBeamBlanco.
    # Permite texto corto sin espacios con letras, números, guion, guion bajo y diagonal.
    if re.fullmatch(r"[A-Za-z][A-Za-z0-9_\-/]{2,55}", x):
        return True

    return False


def detectar_columna_madre(df):
    for c in df.columns:
        cn = normalizar_texto(c)
        if cn in ["sku_madre", "sku madre", "referencia madre", "referencia_madre"] or "sku madre" in cn:
            return c

    scores = {}

    for c in df.columns:
        scores[c] = df[c].apply(parece_sku_madre).sum()

    if not scores:
        return None

    mejor_col = max(scores, key=scores.get)

    if scores[mejor_col] > 0:
        return mejor_col

    return None


def cargar_diccionario_origen4(ruta_diccionario):
    xls = pd.ExcelFile(ruta_diccionario)
    registros = []

    for hoja in xls.sheet_names:
        df = pd.read_excel(ruta_diccionario, sheet_name=hoja, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        col_madre = detectar_columna_madre(df)

        if col_madre is None:
            continue

        for _, row in df.iterrows():
            sku_madre = limpiar_sku(row.get(col_madre, ""))

            if not parece_sku_madre(sku_madre):
                posibles_iq = [
                    limpiar_sku(v)
                    for v in row.values
                    if parece_sku_madre(v)
                ]

                if posibles_iq:
                    sku_madre = posibles_iq[0]
                else:
                    continue

            producto_madre = ""

            # Intenta detectar nombre de producto, pero no es crítico.
            for c in df.columns:
                cn = normalizar_texto(c)
                if "producto" in cn or "nombre" in cn or "descripcion" in cn:
                    val_prod = str(row.get(c, "") or "").strip()
                    if val_prod and val_prod.lower() not in ["nan", "none"]:
                        producto_madre = val_prod
                        break

            # Agrega el propio IQ como alias.
            registros.append({
                "sku_key": sku_key(sku_madre),
                "alias_diccionario": sku_madre,
                "sku_madre": sku_madre,
                "producto_madre": producto_madre,
                "hoja_diccionario": hoja,
                "columna_alias": col_madre,
            })

            # Agrega todos los aliases válidos de la fila.
            for c in df.columns:
                val = limpiar_sku(row.get(c, ""))

                if not parece_alias_valido(val):
                    continue

                registros.append({
                    "sku_key": sku_key(val),
                    "alias_diccionario": val,
                    "sku_madre": sku_madre,
                    "producto_madre": producto_madre,
                    "hoja_diccionario": hoja,
                    "columna_alias": c,
                })

    dic_match = pd.DataFrame(registros)

    if dic_match.empty:
        raise ValueError("No pude construir diccionario de aliases desde origen4.")

    dic_match = dic_match[
        (dic_match["sku_key"] != "")
        & (dic_match["sku_madre"] != "")
    ].copy()

    # Agregar variante sin ceros a la izquierda para SKUs numéricos.
    extra_rows = []

    for _, r in dic_match.iterrows():
        key_original = r["sku_key"]
        key_sin_ceros = sku_key_sin_ceros(key_original)

        if key_sin_ceros and key_sin_ceros != key_original:
            nuevo = r.copy()
            nuevo["sku_key"] = key_sin_ceros
            nuevo["alias_diccionario"] = str(r["alias_diccionario"]) + " | variante_sin_ceros"
            extra_rows.append(nuevo)

    if extra_rows:
        dic_match = pd.concat(
            [dic_match, pd.DataFrame(extra_rows)],
            ignore_index=True
        )

    dic_duplicados = dic_match[
        dic_match.duplicated("sku_key", keep=False)
    ].sort_values(["sku_key", "sku_madre"])

    dic_match = dic_match.drop_duplicates("sku_key", keep="first").copy()

    return dic_match, dic_duplicados


# ============================================================
# AUTOAZUR
# ============================================================

def extraer_sku_desde_texto(texto):
    if pd.isna(texto):
        return ""

    texto = str(texto)

    # Patrón SKU: XXXXX
    m = re.search(r"SKU[:\s]+([A-Za-z0-9\-_]+)", texto, flags=re.IGNORECASE)
    if m:
        return limpiar_sku(m.group(1))

    # Patrón IQ.
    m = re.search(r"\b(IQ\d+)\b", texto, flags=re.IGNORECASE)
    if m:
        return limpiar_sku(m.group(1))

    # Patrón con guiones tipo 1167642485-FBL.
    m = re.search(r"\b([A-Za-z0-9]+-[A-Za-z0-9\-_]+)\b", texto)
    if m:
        return limpiar_sku(m.group(1))

    # Código numérico largo.
    m = re.search(r"\b(\d{7,})\b", texto)
    if m:
        return limpiar_sku(m.group(1))

    return ""


def preparar_ventas_autoazur(ruta_autoazur):
    if not ruta_autoazur:
        return pd.DataFrame()

    print("Preparando ventas Autoazur...")

    df = pd.read_excel(ruta_autoazur, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    col_fecha = encontrar_columna_opcional(df, ["fecha", "Fecha", "Fecha de creación", "Fecha pedido"])
    col_folio = encontrar_columna_opcional(df, ["folio", "Folio", "pedido", "Pedido"])
    col_status = encontrar_columna_opcional(df, ["status", "estatus", "Estado"])
    col_canal = encontrar_columna_opcional(df, ["canal", "Canal", "Marketplace"])
    col_tienda = encontrar_columna_opcional(df, ["tienda", "Tienda"])
    col_ref = encontrar_columna_opcional(df, ["referencia", "Referencia", "Referencia cliente", "referencia_cliente"])
    col_item = encontrar_columna_opcional(df, ["item_id", "Item ID", "ItemID", "item"])
    col_producto = encontrar_columna_opcional(df, ["producto_autoazur", "producto", "Producto", "Título", "Titulo", "Descripción", "Descripcion", "Nombre"])
    col_sku_original = encontrar_columna_opcional(df, ["sku_original", "SKU Original", "sku", "SKU"])
    col_sku_columna = encontrar_columna_opcional(df, ["sku_desde_columna", "SKU desde columna"])
    col_sku_titulo = encontrar_columna_opcional(df, ["sku_desde_titulo", "SKU desde título", "SKU desde titulo"])
    col_cantidad = encontrar_columna_opcional(df, ["cantidad", "Cantidad", "Qty", "Unidades"])
    col_total = encontrar_columna_opcional(df, ["venta_total", "Venta total", "Total", "Venta", "Importe", "Precio total"])

    sku_vals = []

    for _, row in df.iterrows():
        candidatos = []

        for c in [col_sku_original, col_sku_columna, col_sku_titulo, col_item, col_ref]:
            if c:
                candidatos.append(limpiar_sku(row.get(c, "")))

        if col_producto:
            candidatos.append(extraer_sku_desde_texto(row.get(col_producto, "")))

        texto_fila = " ".join([str(v) for v in row.values if pd.notna(v)])
        candidatos.append(extraer_sku_desde_texto(texto_fila))

        sku_final = ""

        for cand in candidatos:
            if cand:
                sku_final = cand
                break

        sku_vals.append(sku_final)

    out = pd.DataFrame({
        "fuente": "AUTOAZUR",
        "fecha": pd.to_datetime(df[col_fecha], errors="coerce") if col_fecha else pd.NaT,
        "pedido": df[col_folio] if col_folio else "",
        "estado": df[col_status] if col_status else "",
        "estado_odoo": "",
        "tipo_venta": df[col_canal] if col_canal else "Autoazur",
        "canal": df[col_canal] if col_canal else "Autoazur",
        "tienda": df[col_tienda] if col_tienda else "",
        "cliente": "",
        "equipo_ventas": "",
        "almacen": "",
        "referencia_cliente": df[col_ref] if col_ref else "",
        "referencia": df[col_ref] if col_ref else "",
        "item_id": df[col_item] if col_item else "",
        "origen": "AUTOAZUR",
        "line_id": "",
        "product_id": "",
        "producto": df[col_producto] if col_producto else "",
        "sku_original": sku_vals,
        "sku_desde_columna": df[col_sku_columna] if col_sku_columna else "",
        "sku_desde_titulo": df[col_sku_titulo] if col_sku_titulo else "",
        "sku_default_code": "",
        "barcode": "",
        "categoria": "",
        "cantidad": to_number(df[col_cantidad]) if col_cantidad else 1,
        "cantidad_entregada": 0,
        "precio_unitario": 0,
        "venta_total": to_number(df[col_total]) if col_total else 0,
    })

    out["sku_original"] = out["sku_original"].apply(limpiar_sku)

    out = out[
        (out["sku_original"] != "")
        | (out["referencia"].astype(str).str.strip() != "")
        | (out["item_id"].astype(str).str.strip() != "")
    ].copy()

    print(f"Ventas Autoazur preparadas: {len(out)}")

    return out


# ============================================================
# DEDUPLICAR ODOO + AUTOAZUR
# ============================================================

def construir_llave_venta(row):
    """
    Crea una llave para evitar duplicados entre Odoo y Autoazur.
    Prioridad: referencia, referencia_cliente, item_id, pedido.
    """
    candidatos = [
        row.get("referencia", ""),
        row.get("referencia_cliente", ""),
        row.get("item_id", ""),
        row.get("pedido", ""),
    ]

    for c in candidatos:
        for k in referencia_variantes(c):
            if k:
                return k

    return ""


def deduplicar_ventas_odoo_autoazur(ventas_odoo, ventas_autoazur):
    """
    Odoo y Autoazur se complementan:
    - Si la referencia aparece en Odoo y Autoazur, se conserva Odoo.
    - Si Autoazur no aparece en Odoo, se conserva.
    """
    ventas_odoo = ventas_odoo.copy()
    ventas_autoazur = ventas_autoazur.copy()

    if not ventas_odoo.empty:
        ventas_odoo["llave_venta"] = ventas_odoo.apply(construir_llave_venta, axis=1)
    else:
        ventas_odoo["llave_venta"] = []

    if not ventas_autoazur.empty:
        ventas_autoazur["llave_venta"] = ventas_autoazur.apply(construir_llave_venta, axis=1)
    else:
        ventas_autoazur["llave_venta"] = []

    llaves_odoo = set(
        ventas_odoo.loc[
            ventas_odoo["llave_venta"].astype(str).str.strip() != "",
            "llave_venta"
        ].astype(str)
    ) if not ventas_odoo.empty else set()

    if ventas_autoazur.empty:
        autoazur_filtrado = ventas_autoazur
    else:
        autoazur_filtrado = ventas_autoazur[
            ~ventas_autoazur["llave_venta"].astype(str).isin(llaves_odoo)
            | (ventas_autoazur["llave_venta"].astype(str).str.strip() == "")
        ].copy()

        autoazur_filtrado["motivo_deduplicacion"] = "conservada_autoazur_no_en_odoo"

    autoazur_duplicado = pd.DataFrame()

    if not ventas_autoazur.empty:
        autoazur_duplicado = ventas_autoazur[
            ventas_autoazur["llave_venta"].astype(str).isin(llaves_odoo)
            & (ventas_autoazur["llave_venta"].astype(str).str.strip() != "")
        ].copy()
        autoazur_duplicado["motivo_deduplicacion"] = "omitida_por_existir_en_odoo"

    if not ventas_odoo.empty:
        ventas_odoo["motivo_deduplicacion"] = "conservada_odoo"

    ventas_conjunto = pd.concat(
        [
            ventas_odoo,
            autoazur_filtrado,
        ],
        ignore_index=True
    )

    return ventas_conjunto, autoazur_duplicado


# ============================================================
# MATCH EN CAPAS
# ============================================================

def construir_mapa_referencia_odoo(ventas_odoo):
    """
    Mapa referencia -> SKU Odoo.
    Usa referencia_cliente, referencia, origen y pedido.
    """
    if ventas_odoo.empty:
        return {}

    mapa = {}

    columnas_ref = [
        "referencia_cliente",
        "referencia",
        "origen",
        "pedido",
    ]

    for _, row in ventas_odoo.iterrows():
        sku = limpiar_sku(row.get("sku_original", ""))

        if not sku:
            continue

        for col in columnas_ref:
            ref = limpiar_sku(row.get(col, ""))

            if not ref:
                continue

            for k in referencia_variantes(ref):
                if k and k not in mapa:
                    mapa[k] = sku

    return mapa


def obtener_candidatos_sku_row(row):
    """
    Candidatos de SKU por fila.
    Incluye sku_original, sku_default_code, sku_desde_columna, sku_desde_titulo y barcode.
    """
    cols = [
        "sku_original",
        "sku_default_code",
        "sku_desde_columna",
        "sku_desde_titulo",
        "barcode",
    ]

    candidatos = []

    for col in cols:
        val = limpiar_sku(row.get(col, ""))
        if val and val.lower() not in ["false", "falso", "true", "verdadero"]:
            candidatos.append(val)

    out = []
    for c in candidatos:
        if c not in out:
            out.append(c)

    return out


def buscar_match_diccionario_por_candidatos(candidatos_sku, dic_map):
    """
    Busca match por múltiples candidatos.
    Regresa: match, metodo, sku_usado
    """
    for sku in candidatos_sku:
        for k in generar_sku_keys_match(sku):
            if k in dic_map:
                metodo = "sku_original" if k == sku_key(sku) else "sku_sin_ceros_izquierda"
                return dic_map[k], metodo, sku

    return None, "", ""


def cruzar_ventas_con_diccionario(ventas_conjunto, dic_match, ventas_odoo):
    """
    Match ventas:
    1. SKU directo con múltiples columnas.
    2. SKU sin ceros a la izquierda.
    3. Si es Autoazur y no encontró:
       referencia Autoazur -> ventas Odoo -> sku Odoo -> diccionario.
    """
    if ventas_conjunto.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")
    mapa_ref_odoo = construir_mapa_referencia_odoo(ventas_odoo)

    rows = []

    for _, row in ventas_conjunto.iterrows():
        row = row.copy()

        fuente = str(row.get("fuente", "")).upper()
        candidatos_sku = obtener_candidatos_sku_row(row)

        match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos(
            candidatos_sku,
            dic_map
        )

        sku_odoo_desde_referencia = ""
        referencia_usada_para_odoo = ""

        if match is None and fuente == "AUTOAZUR":
            refs = [
                row.get("referencia", ""),
                row.get("referencia_cliente", ""),
                row.get("item_id", ""),
                row.get("pedido", ""),
            ]

            for ref in refs:
                for ref_k in referencia_variantes(ref):
                    if ref_k in mapa_ref_odoo:
                        sku_odoo_desde_referencia = mapa_ref_odoo[ref_k]
                        referencia_usada_para_odoo = str(ref)
                        break

                if sku_odoo_desde_referencia:
                    break

            if sku_odoo_desde_referencia:
                match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos(
                    [sku_odoo_desde_referencia],
                    dic_map
                )

                if match is not None:
                    metodo_match = "referencia_autoazur_vs_odoo"

        sku_original = limpiar_sku(row.get("sku_original", ""))

        row["sku_key"] = sku_key(sku_original)
        row["sku_key_sin_ceros"] = sku_key_sin_ceros(sku_original)
        row["sku_odoo_desde_referencia"] = sku_odoo_desde_referencia
        row["referencia_usada_para_odoo"] = referencia_usada_para_odoo
        row["sku_usado_para_match"] = sku_usado_para_match
        row["metodo_match"] = metodo_match

        if match:
            row["alias_diccionario"] = match.get("alias_diccionario", "")
            row["sku_madre"] = match.get("sku_madre", "")
            row["producto_madre"] = match.get("producto_madre", "")
            row["hoja_diccionario"] = match.get("hoja_diccionario", "")
            row["columna_alias"] = match.get("columna_alias", "")
            row["tiene_referencia_madre"] = "SI"
        else:
            row["alias_diccionario"] = ""
            row["sku_madre"] = ""
            row["producto_madre"] = ""
            row["hoja_diccionario"] = ""
            row["columna_alias"] = ""
            row["tiene_referencia_madre"] = "NO"

        rows.append(row)

    ventas_con_match = pd.DataFrame(rows)

    ventas_sin_match = ventas_con_match[
        ventas_con_match["tiene_referencia_madre"] == "NO"
    ].copy()

    ventas_sku_madre = (
        ventas_con_match[ventas_con_match["tiene_referencia_madre"] == "SI"]
        .groupby(["sku_madre", "producto_madre"], as_index=False)
        .agg(
            unidades_vendidas=("cantidad", "sum"),
            venta_total=("venta_total", "sum"),
            pedidos=("pedido", "nunique"),
            lineas=("sku_usado_para_match", "count"),
            fuentes_venta=("fuente", lambda x: " | ".join(sorted(set(map(str, x))))),
            metodos_match=("metodo_match", lambda x: " | ".join(sorted(set(map(str, x))))),
            skus_usados_para_match=("sku_usado_para_match", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
            skus_originales=("sku_original", lambda x: " | ".join(sorted(set(map(str, x)))[:80])),
            skus_odoo_desde_referencia=("sku_odoo_desde_referencia", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
        )
        .sort_values("unidades_vendidas", ascending=False)
    )

    return ventas_con_match, ventas_sin_match, ventas_sku_madre


def cruzar_stock_con_diccionario(stock_por_sku, dic_match):
    if stock_por_sku.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")

    rows = []

    for _, row in stock_por_sku.iterrows():
        row = row.copy()

        sku_original = limpiar_sku(row.get("sku_original", ""))
        match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos(
            [sku_original],
            dic_map
        )

        row["sku_key"] = sku_key(sku_original)
        row["sku_key_sin_ceros"] = sku_key_sin_ceros(sku_original)
        row["sku_usado_para_match"] = sku_usado_para_match
        row["metodo_match"] = metodo_match

        if match:
            row["alias_diccionario"] = match.get("alias_diccionario", "")
            row["sku_madre"] = match.get("sku_madre", "")
            row["producto_madre"] = match.get("producto_madre", "")
            row["hoja_diccionario"] = match.get("hoja_diccionario", "")
            row["columna_alias"] = match.get("columna_alias", "")
            row["tiene_referencia_madre"] = "SI"
        else:
            row["alias_diccionario"] = ""
            row["sku_madre"] = ""
            row["producto_madre"] = ""
            row["hoja_diccionario"] = ""
            row["columna_alias"] = ""
            row["tiene_referencia_madre"] = "NO"

        rows.append(row)

    stock_con_match = pd.DataFrame(rows)

    stock_sin_match = stock_con_match[
        stock_con_match["tiene_referencia_madre"] == "NO"
    ].copy()

    stock_sku_madre = (
        stock_con_match[stock_con_match["tiene_referencia_madre"] == "SI"]
        .groupby(["sku_madre", "producto_madre"], as_index=False)
        .agg(
            stock_walmart_wfs=("WALMART_WFS", "sum"),
            stock_liverpool_99min=("LIVERPOOL_FULL_99MIN", "sum"),
            stock_meli_full=("MERCADO_LIBRE_FULL", "sum"),
            stock_amazon_fba=("AMAZON_FBA", "sum"),
            stock_odoo_cuautitlan=("ODOO_CUAUTITLAN", "sum"),
            stock_total=("stock_total", "sum"),
            metodos_match=("metodo_match", lambda x: " | ".join(sorted(set(map(str, x))))),
            skus_usados_para_match=("sku_usado_para_match", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
            skus_originales=("sku_original", lambda x: " | ".join(sorted(set(map(str, x)))[:80])),
        )
        .sort_values("stock_total", ascending=False)
    )

    return stock_con_match, stock_sin_match, stock_sku_madre


# ============================================================
# ARCHIVOS LOCALES
# ============================================================

print("Buscando archivos en Desktop...")

archivo_diccionario = encontrar_archivo(
    DESKTOP,
    ["diccionario", "origen4"]
)

archivo_walmart = encontrar_archivo(
    DESKTOP,
    ["items_list"]
)

archivo_liverpool_1 = encontrar_archivo(
    DESKTOP,
    ["reporte de ofertas actuales", "(3)"]
)

archivo_liverpool_2 = encontrar_archivo(
    DESKTOP,
    ["reporte de ofertas actuales", "(4)"]
)

archivo_meli = encontrar_archivo(
    DESKTOP,
    ["stock_general_full"]
)

archivo_amazon = encontrar_archivo(
    DESKTOP,
    ["inventario en bodega"]
)

archivo_autoazur = encontrar_archivo_autoazur_actualizado()

print("Archivos detectados:")
print(f"- Diccionario: {archivo_diccionario.name}")
print(f"- Walmart: {archivo_walmart.name}")
print(f"- Liverpool 1: {archivo_liverpool_1.name}")
print(f"- Liverpool 2: {archivo_liverpool_2.name}")
print(f"- Mercado Libre: {archivo_meli.name}")
print(f"- Amazon: {archivo_amazon.name}")

if archivo_autoazur:
    print(f"- Autoazur: {archivo_autoazur.name}")
else:
    print("- Autoazur: NO ENCONTRADO")


# ============================================================
# DICCIONARIO
# ============================================================

dic_match, dic_duplicados = cargar_diccionario_origen4(archivo_diccionario)

print(f"Aliases válidos cargados del diccionario: {len(dic_match)}")
print(f"Aliases duplicados detectados: {len(dic_duplicados)}")


# ============================================================
# ODOO
# ============================================================

ventas_odoo = pd.DataFrame()
productos_odoo = pd.DataFrame()
ventas_sin_sku_odoo = pd.DataFrame()
stock_odoo = pd.DataFrame()
odoo_ubicaciones = pd.DataFrame()

if ENABLE_ODOO:
    odoo = OdooClient(
        ODOO_URL,
        ODOO_DB,
        ODOO_USER,
        ODOO_API_KEY
    )

    odoo.connect()

    ventas_odoo, productos_odoo, ventas_sin_sku_odoo = extraer_ventas_odoo(odoo)
    stock_odoo, odoo_ubicaciones = extraer_inventario_odoo(odoo)

else:
    raise ValueError("ENABLE_ODOO está en False. Para esta versión necesitamos Odoo.")


# ============================================================
# VENTAS = ODOO + AUTOAZUR, SIN DUPLICAR REFERENCIAS
# ============================================================

ventas_autoazur = preparar_ventas_autoazur(archivo_autoazur)

ventas_conjunto, autoazur_duplicados_omitidos = deduplicar_ventas_odoo_autoazur(
    ventas_odoo,
    ventas_autoazur
)

ventas_con_match, ventas_sin_match, ventas_sku_madre = cruzar_ventas_con_diccionario(
    ventas_conjunto,
    dic_match,
    ventas_odoo
)

print(f"Ventas Odoo: {len(ventas_odoo)}")
print(f"Ventas Autoazur originales: {len(ventas_autoazur)}")
print(f"Ventas Autoazur omitidas por duplicado con Odoo: {len(autoazur_duplicados_omitidos)}")
print(f"Ventas conjunto final Odoo + Autoazur: {len(ventas_conjunto)}")
print(f"Ventas vinculadas a SKU madre: {len(ventas_con_match[ventas_con_match['tiene_referencia_madre'] == 'SI']) if len(ventas_con_match) else 0}")
print(f"Ventas sin referencia madre: {len(ventas_sin_match)}")


# ============================================================
# STOCK WALMART
# ============================================================

# El archivo items_list trae una fila de título arriba.
# Los encabezados reales están en la fila 2 de Excel.
walmart = pd.read_excel(archivo_walmart, dtype=str, header=1)
walmart.columns = [str(c).strip() for c in walmart.columns]

stock_walmart = pd.DataFrame({
    "canal_stock": "WALMART_WFS",
    "sku_original": walmart[encontrar_columna(walmart, ["SKU"])].apply(limpiar_sku),
    "producto_stock": walmart[encontrar_columna(walmart, ["Nombre del artículo", "Nombre del articulo", "Product Name", "Nombre"])],
    "stock": to_number(walmart[encontrar_columna(walmart, ["WFS MX"])]),
    "fuente_archivo": archivo_walmart.name,
})

stock_walmart = stock_walmart[
    stock_walmart["sku_original"] != ""
].copy()


# ============================================================
# STOCK LIVERPOOL
# ============================================================

def leer_liverpool(ruta):
    df = pd.read_excel(ruta, sheet_name=0, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]

    col_sku_prod = encontrar_columna(df, ["SKU de producto"])
    col_cantidad = encontrar_columna(df, ["Cantidad"])
    col_expedido = encontrar_columna(df, ["Expedido por"])

    col_producto = encontrar_columna_opcional(df, ["Producto", "Nombre", "Título", "Titulo"])

    if col_producto is None:
        col_producto = col_sku_prod

    df["expedido_limpio"] = df[col_expedido].apply(
        lambda x: normalizar_texto(x).replace(" ", "")
    )

    df = df[df["expedido_limpio"] == "99min"].copy()

    return pd.DataFrame({
        "canal_stock": "LIVERPOOL_FULL_99MIN",
        "sku_original": df[col_sku_prod].apply(limpiar_sku),
        "producto_stock": df[col_producto],
        "stock": to_number(df[col_cantidad]),
        "fuente_archivo": ruta.name,
    })


stock_liverpool = pd.concat(
    [
        leer_liverpool(archivo_liverpool_1),
        leer_liverpool(archivo_liverpool_2),
    ],
    ignore_index=True
)

stock_liverpool = stock_liverpool[
    stock_liverpool["sku_original"] != ""
].copy()


# ============================================================
# STOCK MERCADO LIBRE FULL
# ============================================================

meli_raw = pd.read_excel(
    archivo_meli,
    sheet_name="Resumen",
    header=None,
    dtype=str
)

# Según archivo compartido:
# SKU = col 3, Producto = col 6, Unidades Full = col 21, datos desde fila 13 aprox.
meli_data = meli_raw.iloc[12:].copy()

stock_meli = pd.DataFrame({
    "canal_stock": "MERCADO_LIBRE_FULL",
    "sku_original": meli_data.iloc[:, 3].apply(limpiar_sku),
    "producto_stock": meli_data.iloc[:, 6],
    "stock": to_number(meli_data.iloc[:, 21]),
    "fuente_archivo": archivo_meli.name,
})

stock_meli = stock_meli[
    stock_meli["sku_original"] != ""
].copy()


# ============================================================
# STOCK AMAZON FBA
# ============================================================

amazon = pd.read_excel(archivo_amazon, dtype=str)
amazon.columns = [str(c).strip() for c in amazon.columns]

stock_amazon = pd.DataFrame({
    "canal_stock": "AMAZON_FBA",
    "sku_original": amazon[encontrar_columna(amazon, ["Sku", "SKU"])].apply(limpiar_sku),
    "producto_stock": amazon[encontrar_columna(amazon, ["Titulo del producto", "Título del producto", "Producto", "Nombre"])],
    "stock": to_number(amazon[encontrar_columna(amazon, ["Balance final"])]),
    "fuente_archivo": archivo_amazon.name,
})

stock_amazon = stock_amazon[
    stock_amazon["sku_original"] != ""
].copy()


# ============================================================
# CONSOLIDAR STOCK
# ============================================================

stock_detalle = pd.concat(
    [
        stock_walmart,
        stock_liverpool,
        stock_meli,
        stock_amazon,
        stock_odoo,
    ],
    ignore_index=True
)

stock_detalle["sku_key"] = stock_detalle["sku_original"].apply(sku_key)
stock_detalle["sku_key_sin_ceros"] = stock_detalle["sku_original"].apply(sku_key_sin_ceros)

stock_detalle = stock_detalle[
    stock_detalle["sku_key"] != ""
].copy()

stock_por_sku_canal = (
    stock_detalle
    .groupby(["canal_stock", "sku_key", "sku_original"], as_index=False)
    .agg(
        stock=("stock", "sum"),
        producto_stock=("producto_stock", lambda x: " | ".join(sorted(set(map(str, x.dropna())))[:3])),
        fuentes=("fuente_archivo", lambda x: " | ".join(sorted(set(map(str, x))))),
    )
)

stock_por_sku = (
    stock_por_sku_canal
    .pivot_table(
        index=["sku_key", "sku_original"],
        columns="canal_stock",
        values="stock",
        aggfunc="sum",
        fill_value=0
    )
    .reset_index()
)

stock_cols = [
    "WALMART_WFS",
    "LIVERPOOL_FULL_99MIN",
    "MERCADO_LIBRE_FULL",
    "AMAZON_FBA",
    "ODOO_CUAUTITLAN",
]

for col in stock_cols:
    if col not in stock_por_sku.columns:
        stock_por_sku[col] = 0

stock_por_sku["stock_total"] = stock_por_sku[stock_cols].sum(axis=1)

stock_con_match, stock_sin_match, stock_sku_madre = cruzar_stock_con_diccionario(
    stock_por_sku,
    dic_match
)

print(f"Stock SKUs únicos: {len(stock_por_sku)}")
print(f"Stock vinculado a SKU madre: {len(stock_con_match[stock_con_match['tiene_referencia_madre'] == 'SI']) if len(stock_con_match) else 0}")
print(f"Stock sin referencia madre: {len(stock_sin_match)}")


# ============================================================
# ROTACIÓN BASE
# ============================================================

if ventas_sku_madre.empty:
    ventas_resumen = pd.DataFrame(
        columns=[
            "sku_madre", "unidades_vendidas", "venta_total", "pedidos",
            "fuentes_venta", "metodos_match", "skus_usados_para_match",
            "skus_originales", "skus_odoo_desde_referencia"
        ]
    )
else:
    columnas_ventas_resumen = [
        "sku_madre", "unidades_vendidas", "venta_total", "pedidos",
        "fuentes_venta", "metodos_match", "skus_usados_para_match",
        "skus_originales", "skus_odoo_desde_referencia"
    ]
    columnas_ventas_resumen = [c for c in columnas_ventas_resumen if c in ventas_sku_madre.columns]
    ventas_resumen = ventas_sku_madre[columnas_ventas_resumen]

rotacion_base = stock_sku_madre.merge(
    ventas_resumen,
    on="sku_madre",
    how="outer"
)

for col in [
    "stock_walmart_wfs",
    "stock_liverpool_99min",
    "stock_meli_full",
    "stock_amazon_fba",
    "stock_odoo_cuautitlan",
    "stock_total",
    "unidades_vendidas",
    "venta_total",
    "pedidos",
]:
    if col in rotacion_base.columns:
        rotacion_base[col] = pd.to_numeric(
            rotacion_base[col],
            errors="coerce"
        ).fillna(0)

if "producto_madre" not in rotacion_base.columns:
    rotacion_base["producto_madre"] = ""

dias_periodo = max(
    (FECHA_FIN - FECHA_INICIO).days + 1,
    1
)

rotacion_base["dias_periodo"] = dias_periodo
rotacion_base["venta_diaria_promedio"] = (
    rotacion_base["unidades_vendidas"] / dias_periodo
)

# Preliminar, hasta integrar arribos:
# stock inicial = stock actual + ventas
rotacion_base["stock_inicial_preliminar"] = (
    rotacion_base["stock_total"] + rotacion_base["unidades_vendidas"]
)

rotacion_base["dias_inventario"] = np.where(
    rotacion_base["venta_diaria_promedio"] > 0,
    rotacion_base["stock_total"] / rotacion_base["venta_diaria_promedio"],
    np.nan
)


def clasificar_alerta(row):
    stock = row.get("stock_total", 0)
    ventas_u = row.get("unidades_vendidas", 0)
    dias_inv = row.get("dias_inventario", np.nan)

    if stock == 0 and ventas_u > 0:
        return "SIN STOCK CON VENTAS"

    if stock == 0 and ventas_u == 0:
        return "SIN STOCK Y SIN VENTAS"

    if ventas_u == 0 and stock > 0:
        return "STOCK SIN VENTAS"

    if pd.notna(dias_inv) and dias_inv <= 15:
        return "BAJO STOCK"

    if pd.notna(dias_inv) and dias_inv >= 120:
        return "SOBRESTOCK"

    return "OK"


rotacion_base["alerta_preliminar"] = rotacion_base.apply(
    clasificar_alerta,
    axis=1
)


# ============================================================
# VALIDACIONES ÚTILES
# ============================================================

validacion_1167642485 = ventas_con_match[
    ventas_con_match.astype(str).apply(
        lambda row: row.str.contains("1167642485", case=False, na=False).any(),
        axis=1
    )
].copy() if not ventas_con_match.empty else pd.DataFrame()

validacion_outspeakblue = ventas_con_match[
    ventas_con_match.astype(str).apply(
        lambda row: row.str.contains("outspeakblue", case=False, na=False).any(),
        axis=1
    )
].copy() if not ventas_con_match.empty else pd.DataFrame()

validacion_ceros = pd.DataFrame({
    "ejemplo": ["05024173182", "5024173182"],
    "sku_key": [sku_key("05024173182"), sku_key("5024173182")],
    "sku_key_sin_ceros": [sku_key_sin_ceros("05024173182"), sku_key_sin_ceros("5024173182")],
})

# Base auxiliar para dashboard:
# ventas por IQ y por SKU sincronizado/alias que hizo match.
if not ventas_con_match.empty:
    ventas_match_si = ventas_con_match[ventas_con_match["tiene_referencia_madre"] == "SI"].copy()

    ventas_match_si["sku_sincronizado"] = ventas_match_si["sku_usado_para_match"].replace("", np.nan)
    ventas_match_si["sku_sincronizado"] = ventas_match_si["sku_sincronizado"].fillna(ventas_match_si["sku_original"])

    ventas_iq_sincronizados = (
        ventas_match_si
        .groupby(["sku_madre", "producto_madre", "sku_sincronizado", "fuente", "canal"], as_index=False)
        .agg(
            unidades_vendidas=("cantidad", "sum"),
            venta_total=("venta_total", "sum"),
            pedidos=("pedido", "nunique"),
            lineas=("sku_original", "count"),
            metodos_match=("metodo_match", lambda x: " | ".join(sorted(set(map(str, x))))),
            aliases_diccionario=("alias_diccionario", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
            skus_originales=("sku_original", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
        )
        .sort_values(["sku_madre", "unidades_vendidas"], ascending=[True, False])
    )
else:
    ventas_iq_sincronizados = pd.DataFrame()


# ============================================================
# RESUMEN
# ============================================================

resumen_control = pd.DataFrame([
    ["fecha_inicio", FECHA_INICIO.strftime("%Y-%m-%d")],
    ["fecha_fin", FECHA_FIN.strftime("%Y-%m-%d")],
    ["dias_periodo", dias_periodo],
    ["ventas_fuente_principal", "ODOO_API + AUTOAZUR"],
    ["ventas_odoo_renglones", len(ventas_odoo)],
    ["ventas_autoazur_renglones_originales", len(ventas_autoazur)],
    ["ventas_autoazur_duplicadas_omitidas", len(autoazur_duplicados_omitidos)],
    ["ventas_conjunto_final_renglones", len(ventas_conjunto)],
    [
        "ventas_conjunto_unidades",
        float(ventas_con_match["cantidad"].sum()) if len(ventas_con_match) else 0
    ],
    [
        "ventas_conjunto_total",
        float(ventas_con_match["venta_total"].sum()) if len(ventas_con_match) else 0
    ],
    ["ventas_odoo_sin_sku", len(ventas_sin_sku_odoo)],
    ["ventas_sin_referencia_madre", len(ventas_sin_match)],
    ["ventas_iq_sincronizados_renglones", len(ventas_iq_sincronizados)],
    ["stock_detalle_renglones", len(stock_detalle)],
    [
        "stock_skus_unicos",
        stock_por_sku["sku_key"].nunique() if len(stock_por_sku) else 0
    ],
    [
        "stock_total_unidades",
        float(stock_por_sku["stock_total"].sum()) if len(stock_por_sku) else 0
    ],
    ["stock_sin_referencia_madre", len(stock_sin_match)],
    ["ubicaciones_odoo_detectadas", len(odoo_ubicaciones)],
    ["aliases_validos_diccionario", len(dic_match)],
    ["aliases_duplicados_diccionario", len(dic_duplicados)],
], columns=["metrica", "valor"])


# ============================================================
# EXPORTACIÓN A EXCEL
# ============================================================

salidas = {
    "resumen_control": resumen_control,
    "rotacion_base": rotacion_base,
    "ventas_conjunto_detalle": ventas_con_match,
    "ventas_sku_madre": ventas_sku_madre,
    "ventas_iq_sincronizados": ventas_iq_sincronizados,
    "ventas_sin_referencia": ventas_sin_match,
    "ventas_odoo_detalle": ventas_odoo,
    "ventas_autoazur_detalle": ventas_autoazur,
    "autoazur_duplicados_omitidos": autoazur_duplicados_omitidos,
    "ventas_odoo_sin_sku": ventas_sin_sku_odoo,
    "stock_detalle": stock_detalle,
    "stock_por_sku": stock_con_match,
    "stock_sku_madre": stock_sku_madre,
    "stock_sin_referencia": stock_sin_match,
    "stock_odoo_detalle": stock_odoo,
    "odoo_ubicaciones": odoo_ubicaciones,
    "productos_odoo": productos_odoo,
    "diccionario_usado": dic_match,
    "diccionario_duplicados": dic_duplicados,
    "validacion_1167642485": validacion_1167642485,
    "validacion_outspeakblue": validacion_outspeakblue,
    "validacion_ceros": validacion_ceros,
}

with pd.ExcelWriter(
    ARCHIVO_SALIDA,
    engine="openpyxl",
    datetime_format="yyyy-mm-dd hh:mm",
    date_format="yyyy-mm-dd"
) as writer:

    for nombre_hoja, df in salidas.items():
        if df is None:
            df = pd.DataFrame()

        df = preparar_para_excel(df)
        nombre_final = nombre_hoja[:31]

        df.to_excel(
            writer,
            sheet_name=nombre_final,
            index=False
        )

        ws = writer.book[nombre_final]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Estilo encabezado.
        for cell in ws[1]:
            try:
                cell.font = cell.font.copy(
                    bold=True,
                    color="FFFFFF"
                )
                cell.fill = cell.fill.copy(
                    fill_type="solid",
                    fgColor="1F4E78"
                )
            except Exception:
                pass

        # Ancho columnas.
        for col_cells in ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter

            for cell in col_cells[:1000]:
                if cell.value is not None:
                    max_len = max(
                        max_len,
                        len(str(cell.value))
                    )

            ws.column_dimensions[col_letter].width = min(
                max(max_len + 2, 12),
                45
            )


print("\nPROCESO TERMINADO")
print(f"Excel generado en: {ARCHIVO_SALIDA}")
print("\nHojas creadas:")

for hoja in salidas:
    print(f"- {hoja[:31]}")
