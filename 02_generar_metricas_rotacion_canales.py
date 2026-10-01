# -*- coding: utf-8 -*-
"""
generar_base_dashboard_rotacion_v2.py

Crea la base limpia para el dashboard a partir de:
Desktop/rotacion_inventario_base_dashboard/rotacion_inventario_base_dashboard_odoo_autoazur.xlsx

Salida:
Desktop/rotacion_inventario_base_dashboard/base_dashboard_rotacion.xlsx

Cambios principales:
- Lead time default = 20 días.
- Sugerencia de compra = max(promedio_ventas_diarias_positivas * 45 - stock_total, 0).
- Ritmo de producto = ventas / días con stock estimado; al llegar a stock 0 se congela hasta el siguiente arribo.
- Alertas:
    Inventario acabado - Comprar ya
    Compra con urgencia
    Inventario suficiente
    Inventario sin ventas
    Sin stock y sin ventas
- Top 80% muestra ranking.
- Agrega stock no vinculado.
- Agrega chips de SKU por marketplace usando origen4/diccionario_usado como fuente principal.
- Ventas no vinculadas se construyen desde:
    1) ventas no vinculadas del Excel operativo
    2) logs Autoazur si existe archivo con "logs autoazur"
    3) Odoo API si configuras credenciales
"""

from pathlib import Path
import os
import re
import json
import math
import unicodedata
import xmlrpc.client
import warnings

import pandas as pd
import numpy as np

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=False)
except ImportError:
    pass

warnings.filterwarnings("ignore", category=UserWarning)


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA = Path(os.getenv(
    "DASHBOARD_ROTACION_DIR",
    str(Path(__file__).resolve().parent)
))
DESKTOP = CARPETA

ARCHIVO_ENTRADA = CARPETA / "rotacion_inventario_base_dashboard_odoo_autoazur.xlsx"
ARCHIVO_SALIDA = CARPETA / "base_dashboard_rotacion.xlsx"

# ============================================================
# CORTE MAESTRO DE STOCK
# ============================================================
# IMPORTANTE:
# El script 01 YA construye el inventario actual correctamente:
#
# Full actual = corte Full + traslados Hecho posteriores - ventas Full posteriores
# Odoo actual = inventario disponible consultado en vivo por API.
#
# Por eso este script 02 NO debe volver a aplicar un segundo descuento global
# sobre el stock. Hacerlo duplicaría ventas/arribos y distorsionaría el inventario.
USAR_STOCK_CONGELADO = False

# Se conserva la fecha para auditoría y reconstrucción auxiliar del ritmo.
# 00_calcular_stock_inicial.py actualiza FECHA_CORTE_FULL en .env.
FECHA_CORTE_STOCK = pd.Timestamp(
    os.getenv("FECHA_CORTE_FULL", "2026-08-20 08:30:00").strip()
)
ARCHIVO_STOCK_CONGELADO = CARPETA / os.getenv(
    "FULL_STOCK_FILENAME",
    "stock_full_corte_actual.xlsx"
).strip()

# Antigüedad por canal (versión 2026-09).
# "capa"     -> la alerta usa la capa MÁS ANTIGUA del canal: si hay piezas con
#               más de 90 días, el SKU entra en alerta aunque el promedio sea menor.
# "promedio" -> criterio anterior: promedio ponderado de días del canal.
CRITERIO_ALERTA_ANTIGUEDAD = os.getenv("CRITERIO_ALERTA_ANTIGUEDAD", "capa").strip().lower()

LEAD_TIME_DEFAULT = 30
COBERTURA_OBJETIVO_DIAS = 45
DIAS_ANALISIS_3M = 90

# Ritmo de venta ajustado por disponibilidad de inventario.
# Los días con stock total estimado = 0 NO avanzan el denominador del ritmo.
# Si el producto se agota, el ritmo queda "congelado" hasta que vuelva a haber stock.
DIAS_RITMO_STOCK = int(os.getenv("DIAS_RITMO_STOCK", "90"))

# Política de abastecimiento desde Odoo hacia almacenes Full.
# Puede ajustarse desde .env sin modificar el código.
OBJETIVO_FULL_DIAS = int(os.getenv("OBJETIVO_FULL_DIAS", "30"))
DIAS_TRANSITO_ODOO_FULL = int(os.getenv("DIAS_TRANSITO_ODOO_FULL", "5"))
MAX_PIEZAS_PRODUCTO_NUEVO_FULL = int(os.getenv("MAX_PIEZAS_PRODUCTO_NUEVO_FULL", "10"))
CANALES_CON_FULL = {
    c.strip()
    for c in os.getenv(
        "CANALES_CON_FULL",
        "Amazon,Mercado Libre,Walmart,Liverpool",
    ).split(",")
    if c.strip()
}

# Política para redistribución interna entre canales.
# La recomendación es analítica: nunca crea ni valida movimientos en Odoo.
DIAS_TRANSITO_INTERNO = int(os.getenv("DIAS_TRANSITO_INTERNO", "5"))
OBJETIVO_REDISTRIBUCION_DIAS = int(os.getenv("OBJETIVO_REDISTRIBUCION_DIAS", "30"))
RESERVA_MINIMA_SIN_VENTAS = int(os.getenv("RESERVA_MINIMA_SIN_VENTAS", "10"))
PESO_DEMANDA_10D = float(os.getenv("PESO_DEMANDA_10D", "0.50"))
PESO_DEMANDA_30D = float(os.getenv("PESO_DEMANDA_30D", "0.30"))
PESO_DEMANDA_90D = float(os.getenv("PESO_DEMANDA_90D", "0.20"))

# Política para REPARTICIÓN de stock central (CUATI/Existencias + CUATI/B2B)
# hacia los canales comerciales. El 20% base evita que un canal sin historia
# quede permanentemente sin oportunidad; el 80% restante sigue el desempeño.
REPARTICION_CANALES = [
    "Amazon", "Mercado Libre", "Walmart", "Liverpool",
    "Coppel", "Elektra", "TikTok",
]
REPARTICION_PESO_BASE_IGUAL = min(max(float(os.getenv("REPARTICION_PESO_BASE_IGUAL", "0.20")), 0.0), 1.0)
REPARTICION_MIN_VENTAS_SKU = max(int(os.getenv("REPARTICION_MIN_VENTAS_SKU", "5")), 1)
REPARTICION_MIN_DIAS_VENTA_SKU = max(int(os.getenv("REPARTICION_MIN_DIAS_VENTA_SKU", "3")), 1)
REPARTICION_MIN_VENTAS_CAT_MARCA = max(int(os.getenv("REPARTICION_MIN_VENTAS_CAT_MARCA", "10")), 1)
REPARTICION_MIN_SKUS_CAT_MARCA = max(int(os.getenv("REPARTICION_MIN_SKUS_CAT_MARCA", "2")), 1)

# Opcional: Odoo para logs/no vinculadas.
ENABLE_ODOO_LOGS = True
ODOO_URL = os.getenv("ODOO_URL", "").strip()
ODOO_DB = os.getenv("ODOO_DB", "").strip()
ODOO_USER = os.getenv("ODOO_USER", "").strip()
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "").strip()
ODOO_CAMPO_TIPO_VENTA = os.getenv("ODOO_CAMPO_TIPO_VENTA", "x_studio_tipo_de_venta")
ENABLE_ODOO_IMAGENES_PRODUCTOS = True
ODOO_CAMPO_IMAGEN_PRODUCTO = "image_512"
ODOO_TIPOS_VENTA_VALIDOS = ["full", "drop"]

# ============================================================
# ARRIBOS ODOO COMPRAS
# ============================================================
# Sustituye el archivo manual de arribos.
# El código toma arribos desde purchase.order.line usando qty_received.
ENABLE_ODOO_ARRIBOS_COMPRAS = True
ODOO_ESTADOS_COMPRA_VALIDOS = ["purchase", "done"]

# Para auditar, extraemos TODOS los arribos de Odoo Compras con qty_received > 0.
# Para el cálculo de stock, solo se SUMAN los arribos posteriores al corte.
EXTRAER_TODOS_LOS_ARRIBOS_ODOO = True

# Si EXTRAER_TODOS_LOS_ARRIBOS_ODOO = False, usa este rango.
FECHA_INICIO_ARRIBOS_ODOO = pd.Timestamp("2025-11-01")

# Para el cálculo de stock, solo se suman arribos posteriores al corte:
# FECHA_CORTE_STOCK = 2026-06-02 15:00:00
# La fecha se reporta como fecha corta en el Excel.
# Si conoces el modelo técnico del log, configúralo aquí o por variable de entorno.
# Ejemplos posibles: "autoazur.log", "x_autoazur_log", "queue.job", etc.
ODOO_LOG_MODEL = os.getenv("ODOO_LOG_MODEL", "").strip()
ODOO_LOG_DATE_FIELD = os.getenv("ODOO_LOG_DATE_FIELD", "").strip()


# ============================================================
# FUNCIONES BASE
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
    if x.lower() in ["nan", "none", "false", "falso", "true", "verdadero"]:
        return ""
    if re.fullmatch(r"\d+\.0", x):
        x = x[:-2]
    return x.strip()


def sku_key(x):
    return limpiar_sku(x).upper().replace(" ", "")


def sku_key_sin_ceros(x):
    k = sku_key(x)
    if re.fullmatch(r"\d+", k):
        return k.lstrip("0") or "0"
    return k


def generar_sku_keys_match(x):
    base = sku_key(x)
    sin_ceros = sku_key_sin_ceros(x)
    keys = []
    if base:
        keys.append(base)
    if sin_ceros and sin_ceros not in keys:
        keys.append(sin_ceros)
    return keys


def referencia_key(x):
    if pd.isna(x):
        return ""
    x = str(x).strip().upper()
    if x.lower() in ["nan", "none", ""]:
        return ""
    return re.sub(r"[^A-Z0-9]", "", x)


def referencia_variantes(x):
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


def to_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)


def first_non_empty(series):
    for v in series:
        sv = str(v).strip()
        if sv and sv.lower() not in ["nan", "none"]:
            return sv
    return ""


def first_image_base64(series):
    """
    Devuelve la primera imagen base64 válida para mostrar en el dashboard.
    """
    for v in series:
        sv = str(v).strip()
        if sv and sv.lower() not in ["nan", "none", "false"]:
            return sv
    return ""


def encontrar_archivo_entrada():
    if ARCHIVO_ENTRADA.exists():
        return ARCHIVO_ENTRADA

    candidatos = sorted(
        CARPETA.glob("*odoo_autoazur*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    if candidatos:
        return candidatos[0]

    candidatos = sorted(
        CARPETA.glob("rotacion_inventario*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    if candidatos:
        return candidatos[0]

    raise FileNotFoundError(
        f"No encontré archivo de entrada en {CARPETA}. "
        "Primero corre el script operativo."
    )


def leer_hoja(xls, nombre, required=False):
    if nombre in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=nombre)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    if required:
        raise ValueError(f"No encontré la hoja requerida: {nombre}")
    return pd.DataFrame()


def encontrar_columna(df, posibles):
    mapa = {normalizar_texto(c): c for c in df.columns}
    for p in posibles:
        pn = normalizar_texto(p)
        if pn in mapa:
            return mapa[pn]
    for cn, real in mapa.items():
        for p in posibles:
            pn = normalizar_texto(p)
            if pn and pn in cn:
                return real
    return None


def es_credencial_odoo_valida():
    return not (
        "TU-ODOO" in ODOO_URL
        or ODOO_DB.startswith("TU_")
        or ODOO_USER.startswith("TU_")
        or not str(ODOO_API_KEY or "").strip()
    )


# ============================================================
# MARKETPLACE / CHIPS
# ============================================================

def clasificar_marketplace(alias="", columna_alias="", hoja_diccionario=""):
    texto = f"{alias} {columna_alias} {hoja_diccionario}".lower()

    reglas = [
        ("Amazon", ["amazon", "amz", "b0", "fba"]),
        ("Mercado Libre", ["mercado libre", "meli", "mlm", "full meli", "mercadolibre"]),
        ("Walmart", ["walmart", "walm", "wal-", "wfs", "jz-"]),
        ("Liverpool", ["liverpool", "liv", "fbl"]),
        ("Coppel", ["coppel", "coppe"]),
        ("Elektra", ["elektra", "elekt"]),
        ("TikTok", ["tiktok", "tik tok", "tiktk"]),
        ("Odoo/Interno", ["odoo", "interno", "default_code"]),
        ("IQ", ["sku madre", "iq"]),
    ]

    alias_up = str(alias).upper().strip()
    if re.fullmatch(r"IQ\d+", alias_up):
        return "IQ"

    for market, pats in reglas:
        if any(p in texto for p in pats):
            return market

    return "Otro"



def es_referencia_celda_excel(x):
    s = str(x or "").strip().upper()
    if not s:
        return False
    return bool(re.fullmatch(r"[A-Z]{1,3}\d{1,7}", s))


def es_upc_ean_puro(x):
    s = str(x or "").strip()
    return bool(re.fullmatch(r"\d{8,14}", s))


def sku_visible_para_dashboard(sku, marketplace="", columna_alias="", hoja_diccionario=""):
    """
    Deja visibles solo SKUs útiles por canal.
    Oculta referencias tipo A9/AA9/G9, UPC/EAN puros y textos auxiliares.
    """
    s = str(sku or "").strip()
    if not s:
        return False

    sl = s.lower()
    su = s.upper()

    if sl in ["nan", "none", "false", "falso", "sin sku", "n/a", "na", "null"]:
        return False

    if es_referencia_celda_excel(su):
        return False

    if es_upc_ean_puro(su):
        return False

    auxiliares = [
        "SINETIQUETA", "SIN ETIQUETA", "COMBINADO", "COMBINADOS",
        "ETIQUETA", "NO USAR", "PRUEBA", "TEST"
    ]
    if any(a in su for a in auxiliares):
        return False

    if not su.startswith("IQ") and len(su) <= 3:
        return False

    return True


def limpiar_sku_visible(sku):
    s = str(sku or "").strip()
    if s.lower() in ["nan", "none", "false", "falso"]:
        return ""
    return s



def normalizar_canal_dashboard(canal):
    """
    Normaliza el canal solo para color/agrupación visual.
    El nombre exacto del canal se conserva en canal_original.
    """
    c = str(canal or "").upper()
    if "AMAZON" in c:
        return "Amazon"
    if "MERCADO" in c or "ML" in c:
        return "Mercado Libre"
    if "WALMART" in c:
        return "Walmart"
    if "LIVERPOOL" in c:
        return "Liverpool"
    if "COPPEL" in c:
        return "Coppel"
    if "ELEKTRA" in c:
        return "Elektra"
    if "TIK" in c:
        return "TikTok"
    if "ODOO" in c or "INTERNO" in c:
        return "Odoo/Interno"
    if "SKU MADRE" in c or c.strip() == "IQ":
        return "IQ"
    if "CANAL SIN MAPEAR" in c:
        return "Otro"
    return str(canal or "Otro").strip() or "Otro"


def cargar_diccionario_usado(diccionario_df):
    """
    Lee el diccionario usando la lógica correcta:
    - Canal = columna 'canal'
    - SKU = columna 'sku_alias'
    - Si sku_alias está vacío, NO se muestra.
    - No se toma celda_origen/columna_origen como SKU visual.
    """
    if diccionario_df.empty:
        return pd.DataFrame(columns=[
            "sku_madre", "sku_sincronizado", "marketplace", "canal_original",
            "columna_alias", "hoja_diccionario", "sku_visible"
        ])

    df = diccionario_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # Soporta tanto el nuevo diccionario como el Excel operativo anterior.
    col_sku = encontrar_columna(df, ["sku_alias", "alias_diccionario", "sku_sincronizado"])
    col_canal = encontrar_columna(df, ["canal", "marketplace"])
    col_producto = encontrar_columna(df, ["producto", "producto_madre"])
    col_status = encontrar_columna(df, ["estatus_alias", "status_alias"])
    col_columna = encontrar_columna(df, ["columna_origen", "columna_alias"])
    col_hoja = encontrar_columna(df, ["hoja_diccionario", "hoja_origen"])

    if "sku_madre" not in df.columns:
        df["sku_madre"] = ""

    df["sku_sincronizado"] = df[col_sku].apply(limpiar_sku) if col_sku else ""
    df["canal_original"] = df[col_canal].astype(str).str.strip() if col_canal else ""
    df["producto_diccionario"] = df[col_producto].astype(str).str.strip() if col_producto else ""
    df["estatus_alias"] = df[col_status].astype(str).str.strip() if col_status else ""
    df["columna_alias"] = df[col_columna].astype(str).str.strip() if col_columna else ""
    df["hoja_diccionario"] = df[col_hoja].astype(str).str.strip() if col_hoja else ""

    # Visualmente solo queremos la columna sku_alias si trae valor real.
    df = df[df["sku_madre"].notna() & df["sku_sincronizado"].astype(str).str.strip().ne("")].copy()
    df = df[
        ~df["sku_sincronizado"].astype(str).str.lower().isin(["false", "falso", "nan", "none", ""])
    ].copy()

    # Si viene estatus_alias, usarlo solo para excluir claramente "SIN SKU",
    # pero no borrar filas válidas cuando el texto venga vacío o distinto.
    if "estatus_alias" in df.columns and df["estatus_alias"].astype(str).str.strip().ne("").any():
        st = df["estatus_alias"].astype(str).str.upper()
        df = df[~st.str.contains("SIN SKU", na=False)].copy()

    df["marketplace"] = df["canal_original"].apply(normalizar_canal_dashboard)
    df["sku_visible"] = True
    df["alias_diccionario"] = df["sku_sincronizado"]
    df["sku_key"] = df["sku_sincronizado"].apply(sku_key)

    return df[[
        "sku_madre", "sku_sincronizado", "alias_diccionario", "sku_key",
        "marketplace", "canal_original", "producto_diccionario",
        "columna_alias", "hoja_diccionario", "sku_visible"
    ]].drop_duplicates()



def crear_chips_por_iq(dic_aliases):
    if dic_aliases.empty:
        return pd.DataFrame(columns=["sku_madre", "sku_chips_json", "skus_sincronizados", "num_skus_sincronizados"])

    rows = []
    dic_aliases = dic_aliases.copy()
    if "sku_visible" in dic_aliases.columns:
        dic_aliases = dic_aliases[dic_aliases["sku_visible"] == True].copy()

    for sku_madre, g in dic_aliases.groupby("sku_madre"):
        g = g.copy()

        # Mantener IQ primero y luego mercados conocidos.
        orden = {
            "IQ": 0, "Amazon": 1, "Mercado Libre": 2, "Walmart": 3, "Liverpool": 4,
            "Coppel": 5, "Elektra": 6, "TikTok": 7, "Odoo/Interno": 8, "Otro": 9
        }
        g["orden"] = g["marketplace"].map(orden).fillna(99)
        g = g.sort_values(["orden", "sku_sincronizado"])

        chips = []
        vistos = set()
        for _, r in g.iterrows():
            sku = str(r["sku_sincronizado"]).strip()
            mkt = str(r["marketplace"]).strip()
            key = (sku, mkt)
            if not sku or key in vistos:
                continue
            vistos.add(key)
            chips.append({"sku": sku, "marketplace": mkt})

        skus_txt = " | ".join([c["sku"] for c in chips[:120]])
        rows.append({
            "sku_madre": sku_madre,
            "sku_chips_json": json.dumps(chips[:120], ensure_ascii=False),
            "skus_sincronizados": skus_txt,
            "num_skus_sincronizados": len(chips),
        })

    return pd.DataFrame(rows)



def crear_skus_por_canal(dic_aliases):
    """
    Crea una vista amigable de SKUs por canal usando:
    - canal_original para el título exacto del canal
    - sku_sincronizado desde sku_alias

    No usa celda_origen ni columna_origen como SKU.
    """
    cols = [
        "sku_madre",
        "sku_por_canal_json",
        "skus_iq",
        "skus_odoo_interno",
        "skus_amazon",
        "skus_mercado_libre",
        "skus_walmart",
        "skus_liverpool",
        "skus_coppel",
        "skus_elektra",
        "skus_tiktok",
        "skus_otro",
        "num_skus_iq",
        "num_skus_odoo_interno",
        "num_skus_amazon",
        "num_skus_mercado_libre",
        "num_skus_walmart",
        "num_skus_liverpool",
        "num_skus_coppel",
        "num_skus_elektra",
        "num_skus_tiktok",
        "num_skus_otro",
    ]

    if dic_aliases.empty:
        return pd.DataFrame(columns=cols), pd.DataFrame(columns=[
            "sku_madre", "canal", "marketplace", "sku_sincronizado",
            "columna_alias", "hoja_diccionario"
        ])

    df = dic_aliases.copy()
    for c in ["sku_madre", "sku_sincronizado", "marketplace", "canal_original", "columna_alias", "hoja_diccionario"]:
        if c not in df.columns:
            df[c] = ""

    if "sku_visible" in df.columns:
        df = df[df["sku_visible"] == True].copy()

    df["sku_madre"] = df["sku_madre"].astype(str).str.strip().str.upper()
    df["sku_sincronizado"] = df["sku_sincronizado"].astype(str).str.strip()
    df["marketplace"] = df["marketplace"].astype(str).str.strip().replace("", "Otro")
    df["canal"] = df["canal_original"].astype(str).str.strip()
    df.loc[df["canal"].eq(""), "canal"] = df["marketplace"]

    df = df[
        df["sku_madre"].ne("")
        & df["sku_sincronizado"].ne("")
        & ~df["sku_sincronizado"].str.lower().isin(["nan", "none", "false", "falso"])
    ].copy()

    orden_mkt = {
        "IQ": 0,
        "Odoo/Interno": 1,
        "Amazon": 2,
        "Mercado Libre": 3,
        "Walmart": 4,
        "Liverpool": 5,
        "Coppel": 6,
        "Elektra": 7,
        "TikTok": 8,
        "Otro": 9,
    }
    df["_orden"] = df["marketplace"].map(orden_mkt).fillna(99)
    df = df.sort_values(["sku_madre", "_orden", "canal", "sku_sincronizado"])

    detalle = df.drop_duplicates(
        ["sku_madre", "canal", "sku_sincronizado"],
        keep="first"
    )[["sku_madre", "canal", "marketplace", "sku_sincronizado", "columna_alias", "hoja_diccionario"]].copy()

    mapa_col = {
        "IQ": "skus_iq",
        "Odoo/Interno": "skus_odoo_interno",
        "Amazon": "skus_amazon",
        "Mercado Libre": "skus_mercado_libre",
        "Walmart": "skus_walmart",
        "Liverpool": "skus_liverpool",
        "Coppel": "skus_coppel",
        "Elektra": "skus_elektra",
        "TikTok": "skus_tiktok",
        "Otro": "skus_otro",
    }

    rows = []
    for sku_madre, g in detalle.groupby("sku_madre"):
        row = {"sku_madre": sku_madre}
        json_groups = []

        # Columnas resumidas por marketplace normalizado.
        for mkt, col in mapa_col.items():
            skus_mkt = (
                g.loc[g["marketplace"].eq(mkt), "sku_sincronizado"]
                .dropna().astype(str).str.strip().drop_duplicates().tolist()
            )
            row[col] = " | ".join([s for s in skus_mkt if s])
            row["num_" + col.replace("skus_", "skus_")] = len([s for s in skus_mkt if s])

        # JSON por canal exacto.
        for canal, gc in g.groupby("canal"):
            skus = (
                gc["sku_sincronizado"]
                .dropna().astype(str).str.strip().drop_duplicates().tolist()
            )
            skus = [s for s in skus if s]
            if not skus:
                continue

            marketplace = first_non_empty(gc["marketplace"])
            json_groups.append({
                "canal": canal,
                "marketplace": marketplace,
                "skus": skus,
                "count": len(skus),
            })

        row["sku_por_canal_json"] = json.dumps(json_groups, ensure_ascii=False)
        rows.append(row)

    out = pd.DataFrame(rows)
    for c in cols:
        if c not in out.columns:
            out[c] = 0 if c.startswith("num_") else ""

    return out[cols].copy(), detalle


# ============================================================
# MATCH DICCIONARIO
# ============================================================

def construir_dic_map(diccionario_df):
    if diccionario_df.empty:
        return {}

    dic = diccionario_df.copy()
    if "sku_key" not in dic.columns and "alias_diccionario" in dic.columns:
        dic["sku_key"] = dic["alias_diccionario"].apply(sku_key)

    dic = dic[dic["sku_key"].notna() & dic["sku_key"].astype(str).str.strip().ne("")].copy()

    # variante sin ceros
    extras = []
    for _, r in dic.iterrows():
        k = str(r["sku_key"])
        k2 = sku_key_sin_ceros(k)
        if k2 and k2 != k:
            nr = r.copy()
            nr["sku_key"] = k2
            extras.append(nr)
    if extras:
        dic = pd.concat([dic, pd.DataFrame(extras)], ignore_index=True)

    return dic.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")


def buscar_iq_por_sku(sku, dic_map):
    for k in generar_sku_keys_match(sku):
        if k in dic_map:
            return dic_map[k].get("sku_madre", "")
    return ""


def candidatos_sku_de_row(row):
    cols = [
        "sku_original", "sku_default_code", "sku_desde_columna", "sku_desde_titulo",
        "barcode", "sku_log", "sku_autoazur", "sku_odoo"
    ]
    out = []
    for c in cols:
        v = limpiar_sku(row.get(c, ""))
        if v and v not in out:
            out.append(v)
    return out


def asignar_iq_a_logs(df, dic_map):
    if df.empty:
        return df

    df = df.copy()
    sku_madre = []
    sku_usado = []
    motivo = []

    for _, r in df.iterrows():
        encontrado = ""
        usado = ""

        for sku in candidatos_sku_de_row(r):
            iq = buscar_iq_por_sku(sku, dic_map)
            if iq:
                encontrado = iq
                usado = sku
                break

        sku_madre.append(encontrado)
        sku_usado.append(usado)

        if encontrado:
            motivo.append("Vinculado")
        else:
            motivo.append("No encontró IQ / falta alias en origen4")

    df["sku_madre"] = sku_madre
    df["sku_usado_para_match"] = sku_usado
    df["motivo_no_vinculado"] = motivo
    df["accion_sugerida"] = np.where(
        df["sku_madre"].astype(str).str.strip().ne(""),
        "Ya vinculado",
        "Agregar alias a origen4 o corregir SKU sincronizado"
    )
    return df


# ============================================================
# ODOO API PARA LOGS
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
        faltantes = []
        if not str(self.url or "").strip():
            faltantes.append("ODOO_URL")
        if not str(self.db or "").strip():
            faltantes.append("ODOO_DB")
        if not str(self.user or "").strip():
            faltantes.append("ODOO_USER")
        if not str(self.api_key or "").strip() or str(self.api_key).startswith("TU_"):
            faltantes.append("ODOO_API_KEY")
        if faltantes:
            raise ValueError(
                "Faltan credenciales en .env: " + ", ".join(faltantes) + ". "
                "Ejecuta ABRIR_DASHBOARD.command para configurarlas mediante una ventana segura."
            )

        common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self.uid = common.authenticate(self.db, self.user, self.api_key, {})
        if not self.uid:
            raise ConnectionError("No se pudo autenticar en Odoo.")
        self.models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        print(f"Conectado a Odoo. UID: {self.uid}")

    def execute(self, model, method, *args, **kwargs):
        return self.models.execute_kw(self.db, self.uid, self.api_key, model, method, args, kwargs)

    def search_read_all(self, model, domain, fields, batch=2000, order=None):
        out, offset = [], 0
        while True:
            kw = {"fields": fields, "limit": batch, "offset": offset}
            if order:
                kw["order"] = order
            rows = self.execute(model, "search_read", domain, **kw)
            if not rows:
                break
            out.extend(rows)
            if len(rows) < batch:
                break
            offset += batch
        return out


def m2o_id(v):
    if isinstance(v, (list, tuple)) and v:
        return v[0]
    return None


def m2o_name(v):
    if isinstance(v, (list, tuple)) and len(v) > 1:
        return v[1]
    return ""


def descubrir_modelos_log_odoo(odoo):
    """
    Busca modelos candidatos para logs reales.
    No toma sale.order.line porque eso son ventas, no logs.
    Exporta candidatos a Excel para que puedas decirme cuál usar.
    """
    try:
        modelos = odoo.search_read_all(
            "ir.model",
            [
                "|", "|", "|",
                ("model", "ilike", "log"),
                ("model", "ilike", "autoazur"),
                ("model", "ilike", "queue"),
                ("name", "ilike", "log"),
            ],
            ["name", "model"],
            batch=1000,
            order="model asc"
        )

        rows = []
        for m in modelos:
            model_name = m.get("model", "")
            if model_name in ["sale.order", "sale.order.line"]:
                continue

            try:
                fields = odoo.search_read_all(
                    "ir.model.fields",
                    [("model", "=", model_name)],
                    ["name", "field_description", "ttype"],
                    batch=2000,
                    order="name asc"
                )
            except Exception:
                fields = []

            field_text = " | ".join([
                f"{f.get('name')} ({f.get('field_description')})"
                for f in fields
            ])

            score = 0
            txt = normalizar_texto(model_name + " " + m.get("name", "") + " " + field_text)
            for kw in ["sku", "pedido", "order", "folio", "referencia", "autoazur", "log", "error"]:
                if kw in txt:
                    score += 1

            rows.append({
                "model": model_name,
                "name": m.get("name", ""),
                "score": score,
                "fields_preview": field_text[:1000],
            })

        cand = pd.DataFrame(rows).sort_values("score", ascending=False)
        ruta = CARPETA / "odoo_modelos_logs_candidatos.xlsx"
        cand.to_excel(ruta, index=False)
        print(f"Exporté modelos candidatos de logs Odoo en: {ruta}")
        return cand

    except Exception as e:
        print(f"No pude descubrir modelos de logs Odoo: {e}")
        return pd.DataFrame()


def detectar_campo_fecha_log(odoo, model):
    if ODOO_LOG_DATE_FIELD:
        return ODOO_LOG_DATE_FIELD

    fields = odoo.search_read_all(
        "ir.model.fields",
        [("model", "=", model)],
        ["name", "field_description", "ttype"],
        batch=2000
    )

    posibles = ["create_date", "write_date", "date", "fecha", "datetime", "timestamp"]
    names = {f["name"]: f for f in fields}

    for p in posibles:
        if p in names:
            return p

    # primer datetime/date
    for f in fields:
        if f.get("ttype") in ["datetime", "date"]:
            return f.get("name")

    return "create_date"


def detectar_campos_log(odoo, model):
    fields = odoo.search_read_all(
        "ir.model.fields",
        [("model", "=", model)],
        ["name", "field_description", "ttype"],
        batch=3000,
        order="name asc"
    )

    def pick(keywords):
        best = None
        for f in fields:
            txt = normalizar_texto(f.get("name", "") + " " + f.get("field_description", ""))
            if any(k in txt for k in keywords):
                best = f.get("name")
                break
        return best

    fecha = detectar_campo_fecha_log(odoo, model)

    return {
        "fecha": fecha,
        "sku": pick(["sku", "seller sku", "default code", "codigo", "código"]),
        "pedido": pick(["pedido", "order", "folio", "orden"]),
        "referencia": pick(["referencia", "reference", "external", "marketplace"]),
        "producto": pick(["producto", "product", "name", "descripcion", "descripción"]),
        "canal": pick(["canal", "channel", "marketplace", "site"]),
        "cantidad": pick(["cantidad", "quantity", "qty", "unidades"]),
        "venta_total": pick(["total", "amount", "importe", "monto", "price"]),
        "mensaje": pick(["message", "mensaje", "log", "error", "description", "body"]),
    }



def extraer_imagenes_odoo_por_iq(iq_list):
    """
    Extrae imágenes desde Odoo usando default_code (IQ).

    No filtra por image_512 en el dominio porque image_512
    no es un campo almacenado y Odoo genera:
        Cannot convert product.product.image_512 to SQL

    El filtrado se hace después de descargar los registros.
    """

    cols = [
        "sku_madre",
        "imagen_odoo_base64_api",
        "producto_imagen_odoo"
    ]

    if not ENABLE_ODOO_IMAGENES_PRODUCTOS:
        return pd.DataFrame(columns=cols)

    if not es_credencial_odoo_valida():
        print("Imágenes Odoo omitidas.")
        return pd.DataFrame(columns=cols)

    iq_list = sorted(set(
        str(x).strip().upper()
        for x in iq_list
        if str(x).strip().upper().startswith("IQ")
    ))

    if len(iq_list) == 0:
        return pd.DataFrame(columns=cols)

    try:

        odoo = OdooClient(
            ODOO_URL,
            ODOO_DB,
            ODOO_USER,
            ODOO_API_KEY
        )
        odoo.connect()

        rows_all = []

        chunk_size = 300

        for i in range(0, len(iq_list), chunk_size):

            chunk = iq_list[i:i+chunk_size]

            domain = [
                ("default_code", "in", chunk)
            ]

            fields = [
                "default_code",
                "display_name",
                ODOO_CAMPO_IMAGEN_PRODUCTO
            ]

            rows = odoo.search_read_all(
                "product.product",
                domain,
                fields,
                batch=1000,
                order="id asc"
            )

            rows_all.extend(rows)

        if len(rows_all) == 0:
            print("No encontré productos Odoo para imágenes.")
            return pd.DataFrame(columns=cols)

        tmp = pd.DataFrame(rows_all)

        tmp["sku_madre"] = (
            tmp["default_code"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        tmp["imagen_odoo_base64_api"] = (
            tmp[ODOO_CAMPO_IMAGEN_PRODUCTO]
            .fillna("")
            .astype(str)
        )

        tmp["producto_imagen_odoo"] = (
            tmp["display_name"]
            .fillna("")
            .astype(str)
        )

        tmp = tmp[
            tmp["sku_madre"].str.match(r"^IQ\d+$", na=False)
        ]

        tmp = tmp[
            tmp["imagen_odoo_base64_api"]
            .astype(str)
            .str.strip()
            .ne("")
        ]

        tmp = tmp[
            tmp["imagen_odoo_base64_api"]
            .astype(str)
            .str.lower()
            .ne("false")
        ]

        if tmp.empty:
            print("No encontré imágenes válidas.")
            return pd.DataFrame(columns=cols)

        imagenes = (
            tmp
            .groupby("sku_madre", as_index=False)
            .agg(
                imagen_odoo_base64_api=(
                    "imagen_odoo_base64_api",
                    first_image_base64
                ),
                producto_imagen_odoo=(
                    "producto_imagen_odoo",
                    first_non_empty
                )
            )
        )

        print(f"Imágenes Odoo recuperadas: {len(imagenes)}")

        return imagenes

    except Exception as e:

        print(f"No pude extraer imágenes desde Odoo: {e}")

        return pd.DataFrame(columns=cols)


def extraer_logs_odoo(fecha_inicio, fecha_fin, dic_map):
    """
    Extrae SOLO logs reales de Odoo. No usa sale.order.line.
    Si ODOO_LOG_MODEL no está configurado, exporta un archivo de modelos candidatos
    y deja la pestaña sin ventas no vinculadas de Odoo hasta que elijas el modelo.
    """
    if not ENABLE_ODOO_LOGS or not es_credencial_odoo_valida():
        print("Odoo logs omitido: credenciales no configuradas.")
        return pd.DataFrame()

    try:
        odoo = OdooClient(ODOO_URL, ODOO_DB, ODOO_USER, ODOO_API_KEY)
        odoo.connect()

        if not ODOO_LOG_MODEL:
            descubrir_modelos_log_odoo(odoo)
            print("ODOO_LOG_MODEL no está configurado. No tomaré ventas Odoo como logs.")
            return pd.DataFrame()

        campos = detectar_campos_log(odoo, ODOO_LOG_MODEL)
        fecha_field = campos["fecha"]

        dt_ini = fecha_inicio.strftime("%Y-%m-%d 00:00:00")
        dt_fin = (fecha_fin + pd.Timedelta(days=1)).strftime("%Y-%m-%d 00:00:00")

        domain = []
        if fecha_field:
            domain = [
                (fecha_field, ">=", dt_ini),
                (fecha_field, "<", dt_fin),
            ]

        fields_to_read = sorted(set([v for v in campos.values() if v]))
        if "id" not in fields_to_read:
            fields_to_read.insert(0, "id")

        rows_raw = odoo.search_read_all(
            ODOO_LOG_MODEL,
            domain,
            fields_to_read,
            batch=2000,
            order=f"{fecha_field} asc" if fecha_field else None
        )

        if not rows_raw:
            return pd.DataFrame()

        out_rows = []

        for r in rows_raw:
            def val(c):
                if not c:
                    return ""
                v = r.get(c, "")
                if isinstance(v, (list, tuple)) and len(v) > 1:
                    return v[1]
                return v

            sku = limpiar_sku(val(campos.get("sku")))
            mensaje = str(val(campos.get("mensaje")) or "")

            # Si no hay campo SKU explícito, intenta extraerlo del mensaje.
            if not sku and mensaje:
                m = re.search(r"(?:SKU|sku)[:\s]+([A-Za-z0-9_\\-/\\.]+)", mensaje)
                if m:
                    sku = limpiar_sku(m.group(1))

            row = {
                "fecha": pd.to_datetime(val(campos.get("fecha")), errors="coerce"),
                "fuente_log": "ODOO_LOG",
                "modelo_log": ODOO_LOG_MODEL,
                "fuente": "ODOO_LOG",
                "canal": val(campos.get("canal")),
                "pedido": val(campos.get("pedido")),
                "referencia": val(campos.get("referencia")),
                "producto": val(campos.get("producto")) or mensaje[:200],
                "sku_log": sku,
                "sku_odoo": sku,
                "sku_autoazur": "",
                "sku_original": sku,
                "cantidad": to_num(pd.Series([val(campos.get("cantidad"))])).iloc[0] if campos.get("cantidad") else 0,
                "venta_total": to_num(pd.Series([val(campos.get("venta_total"))])).iloc[0] if campos.get("venta_total") else 0,
                "mensaje_log": mensaje[:500],
            }
            out_rows.append(row)

        df = pd.DataFrame(out_rows)
        df = asignar_iq_a_logs(df, dic_map)
        df = df[df["sku_madre"].astype(str).str.strip().eq("")].copy()
        return df

    except Exception as e:
        print(f"No pude extraer logs reales de Odoo: {e}")
        return pd.DataFrame()


# ============================================================
# LOGS AUTOAZUR
# ============================================================

def encontrar_archivo_logs_autoazur():
    patrones = ["log", "autoazur"]
    candidatos = []
    for carpeta in [DESKTOP, CARPETA]:
        if carpeta.exists():
            for p in carpeta.glob("*.xlsx"):
                n = normalizar_texto(p.name)
                if all(x in n for x in patrones):
                    candidatos.append(p)
    candidatos = sorted(set(candidatos), key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0] if candidatos else None


def extraer_logs_autoazur(dic_map):
    ruta = encontrar_archivo_logs_autoazur()
    if not ruta:
        print("Logs Autoazur omitidos: no encontré archivo con 'logs autoazur'.")
        return pd.DataFrame()

    print(f"Leyendo logs Autoazur: {ruta.name}")
    try:
        df = pd.read_excel(ruta, dtype=str)
        df.columns = [str(c).strip() for c in df.columns]

        col_fecha = encontrar_columna(df, ["fecha", "Fecha", "create_date", "Fecha creación"])
        col_canal = encontrar_columna(df, ["canal", "Canal", "marketplace", "Marketplace"])
        col_pedido = encontrar_columna(df, ["pedido", "Pedido", "folio", "Folio", "orden", "Orden"])
        col_ref = encontrar_columna(df, ["referencia", "Referencia", "reference", "order_reference"])
        col_prod = encontrar_columna(df, ["producto", "Producto", "name", "display_name", "titulo", "Título"])
        col_sku = encontrar_columna(df, ["sku", "SKU", "sku_log", "default_code", "seller_sku"])
        col_cant = encontrar_columna(df, ["cantidad", "Cantidad", "qty", "quantity"])
        col_total = encontrar_columna(df, ["venta_total", "total", "Total", "price_total", "monto"])

        out = pd.DataFrame({
            "fecha": pd.to_datetime(df[col_fecha], errors="coerce") if col_fecha else pd.NaT,
            "fuente_log": "AUTOAZUR_LOG",
            "fuente": "AUTOAZUR",
            "canal": df[col_canal] if col_canal else "",
            "pedido": df[col_pedido] if col_pedido else "",
            "referencia": df[col_ref] if col_ref else "",
            "origen": "AUTOAZUR_LOG",
            "producto": df[col_prod] if col_prod else "",
            "sku_log": df[col_sku].apply(limpiar_sku) if col_sku else "",
            "sku_odoo": "",
            "sku_autoazur": df[col_sku].apply(limpiar_sku) if col_sku else "",
            "sku_original": df[col_sku].apply(limpiar_sku) if col_sku else "",
            "sku_default_code": "",
            "barcode": "",
            "cantidad": to_num(df[col_cant]) if col_cant else 0,
            "venta_total": to_num(df[col_total]) if col_total else 0,
        })

        out = asignar_iq_a_logs(out, dic_map)
        out = out[out["sku_madre"].astype(str).str.strip().eq("")].copy()
        return out

    except Exception as e:
        print(f"No pude leer logs Autoazur: {e}")
        return pd.DataFrame()



# ============================================================
# DICCIONARIO ORIGEN4 DIRECTO
# ============================================================

def encontrar_diccionario_origen4():
    """
    Busca el archivo original del diccionario, porque ahí sí existe:
    canal + sku_alias.

    El Excel operativo puede traer columnas auxiliares del match, pero para
    mostrar SKUs por canal necesitamos el diccionario real.
    """
    carpetas = [DESKTOP, CARPETA, Path.cwd()]
    patrones = [
        "diccionario_skus_madre_con_origen4*.xlsx",
        "*origen4*.xlsx",
        "*diccionario*sku*.xlsx",
    ]

    candidatos = []
    for carpeta in carpetas:
        if not carpeta.exists():
            continue
        for patron in patrones:
            candidatos.extend(list(carpeta.glob(patron)))

    candidatos = [
        p for p in candidatos
        if not p.name.startswith("~$")
        and p.suffix.lower() in [".xlsx", ".xlsm", ".xls"]
    ]

    if not candidatos:
        return None

    candidatos = sorted(candidatos, key=lambda p: p.stat().st_mtime, reverse=True)
    return candidatos[0]


def leer_diccionario_origen4_directo():
    """
    Lee el diccionario original y lo devuelve en formato compatible:
    sku_madre, producto, canal, sku_alias, alias_diccionario, sku_key...
    """
    archivo_dic = encontrar_diccionario_origen4()

    if archivo_dic is None:
        print("No encontré diccionario origen4 directo; usaré diccionario_usado del Excel operativo.")
        return pd.DataFrame()

    try:
        xls_dic = pd.ExcelFile(archivo_dic)
        frames = []

        for sh in xls_dic.sheet_names:
            try:
                df = pd.read_excel(xls_dic, sheet_name=sh)
                df.columns = [str(c).strip() for c in df.columns]

                col_iq = encontrar_columna(df, ["sku_madre", "SKU madre", "referencia madre", "Ref Interna"])
                col_sku = encontrar_columna(df, ["sku_alias", "SKU alias", "alias_diccionario", "sku sincronizado"])
                col_canal = encontrar_columna(df, ["canal", "marketplace"])
                col_prod = encontrar_columna(df, ["producto", "producto_madre"])
                col_status = encontrar_columna(df, ["estatus_alias", "status_alias"])
                col_celda = encontrar_columna(df, ["celda_origen"])
                col_col = encontrar_columna(df, ["columna_origen", "columna_alias"])
                col_fila = encontrar_columna(df, ["fila_origen"])
                col_codigo = encontrar_columna(df, ["codigo_canal"])

                if not col_iq or not col_sku or not col_canal:
                    continue

                tmp = pd.DataFrame()
                tmp["sku_madre"] = df[col_iq].astype(str).str.strip().str.upper()
                tmp["producto"] = df[col_prod].astype(str).str.strip() if col_prod else ""
                tmp["canal"] = df[col_canal].astype(str).str.strip()
                tmp["sku_alias"] = df[col_sku].apply(limpiar_sku)
                tmp["alias_diccionario"] = tmp["sku_alias"]
                tmp["estatus_alias"] = df[col_status].astype(str).str.strip() if col_status else ""
                tmp["celda_origen"] = df[col_celda].astype(str).str.strip() if col_celda else ""
                tmp["columna_origen"] = df[col_col].astype(str).str.strip() if col_col else ""
                tmp["fila_origen"] = df[col_fila].astype(str).str.strip() if col_fila else ""
                tmp["codigo_canal"] = df[col_codigo].astype(str).str.strip() if col_codigo else ""
                tmp["hoja_diccionario"] = sh
                tmp["sku_key"] = tmp["sku_alias"].apply(sku_key)

                tmp = tmp[
                    tmp["sku_madre"].str.match(r"^IQ\d+$", na=False)
                    & tmp["sku_alias"].astype(str).str.strip().ne("")
                    & ~tmp["sku_alias"].astype(str).str.lower().isin(["nan", "none", "false", "falso"])
                ].copy()

                if not tmp.empty:
                    frames.append(tmp)

            except Exception:
                continue

        if not frames:
            print(f"Encontré {archivo_dic.name}, pero no identifiqué columnas canal/sku_alias.")
            return pd.DataFrame()

        out = pd.concat(frames, ignore_index=True).drop_duplicates()
        print(f"Diccionario origen4 directo usado: {archivo_dic.name}")
        print(f"Aliases con canal/sku_alias cargados: {len(out)}")
        return out

    except Exception as e:
        print(f"No pude leer diccionario origen4 directo: {e}")
        return pd.DataFrame()



# ============================================================
# CARGA
# ============================================================

CARPETA.mkdir(parents=True, exist_ok=True)
archivo_entrada = encontrar_archivo_entrada()
print(f"Archivo de entrada: {archivo_entrada}")

xls = pd.ExcelFile(archivo_entrada)

ventas = leer_hoja(xls, "ventas_conjunto_detalle", required=True)
ventas_sin_ref_excel = leer_hoja(xls, "ventas_sin_referencia")
stock_iq = leer_hoja(xls, "stock_sku_madre")
stock_por_sku = leer_hoja(xls, "stock_por_sku")
stock_sin_ref = leer_hoja(xls, "stock_sin_referencia")
rotacion = leer_hoja(xls, "rotacion_base")
diccionario_usado = leer_hoja(xls, "diccionario_usado")
inventario_canal_operativo = leer_hoja(xls, "inventario_canal_sku_madre")
ventas_canal_operativo = leer_hoja(xls, "ventas_canal_sku_madre")
traslados_full_operativo = leer_hoja(xls, "traslados_full_detalle")
traslados_full_excluidos = leer_hoja(xls, "traslados_full_excluidos")
# Hojas nuevas del 01 (versión 2026-09). Si el 01 es anterior, llegan vacías
# y el dashboard sigue funcionando con la lógica previa.
antiguedad_capas = leer_hoja(xls, "antiguedad_capas")
movimientos_odoo_auditoria = leer_hoja(xls, "movimientos_odoo_auditoria")
movimientos_odoo_resumen = leer_hoja(xls, "movimientos_odoo_resumen")
costos_sku_madre = leer_hoja(xls, "costos_sku_madre")
publicaciones_autoazur = leer_hoja(xls, "publicaciones_autoazur")
publicaciones_sku_canal = leer_hoja(xls, "publicaciones_sku_canal")
publicaciones_sin_sku_madre = leer_hoja(xls, "publicaciones_sin_sku_madre")
publicaciones_log = leer_hoja(xls, "publicaciones_log")
publicaciones_sin_sku_resumen = leer_hoja(xls, "publicaciones_sin_sku_resumen")

diccionario_origen4_directo = leer_diccionario_origen4_directo()
if not diccionario_origen4_directo.empty:
    diccionario_para_skus = diccionario_origen4_directo.copy()
else:
    diccionario_para_skus = diccionario_usado.copy()

dic_aliases = cargar_diccionario_usado(diccionario_para_skus)
sku_aliases_excluidos_visual = dic_aliases[dic_aliases.get("sku_visible", False) == False].copy() if "sku_visible" in dic_aliases.columns else pd.DataFrame()
sku_chips = crear_chips_por_iq(dic_aliases)
sku_por_canal, sku_por_canal_detalle = crear_skus_por_canal(dic_aliases)

# Para logs/no vinculadas conviene usar el diccionario más completo disponible.
dic_map = construir_dic_map(diccionario_para_skus)

# ============================================================
# IMÁGENES ODOO DESDE EXCEL OPERATIVO
# ============================================================

def extraer_imagenes_desde_hojas_operativas(hojas):
    """
    Recupera imagen_odoo_base64 desde cualquier hoja del Excel operativo
    que tenga sku_madre y/o referencia IQ.

    Esto sirve como respaldo cuando la imagen ya venía desde la primera etapa.
    """
    rows = []

    for nombre_hoja, df in hojas.items():
        if df is None or df.empty:
            continue

        col_img = encontrar_columna(df, ["imagen_odoo_base64", "image_512", "imagen odoo", "imagen"])
        if not col_img:
            continue

        col_iq = encontrar_columna(df, ["sku_madre", "referencia_madre", "Ref Interna", "referencia_interna", "default_code", "sku_odoo"])
        if not col_iq:
            continue

        tmp = df[[col_iq, col_img]].copy()
        tmp.columns = ["sku_madre", "imagen_odoo_base64_excel"]
        tmp["sku_madre"] = tmp["sku_madre"].astype(str).str.strip().str.upper()
        tmp["imagen_odoo_base64_excel"] = tmp["imagen_odoo_base64_excel"].fillna("").astype(str).str.strip()
        tmp = tmp[
            tmp["sku_madre"].str.match(r"^IQ\d+$", na=False)
            & tmp["imagen_odoo_base64_excel"].ne("")
            & tmp["imagen_odoo_base64_excel"].str.lower().ne("false")
        ].copy()

        if not tmp.empty:
            tmp["fuente_imagen_excel"] = nombre_hoja
            rows.append(tmp)

    if not rows:
        return pd.DataFrame(columns=["sku_madre", "imagen_odoo_base64_excel", "fuente_imagen_excel"])

    out = pd.concat(rows, ignore_index=True)
    out = (
        out.groupby("sku_madre", as_index=False)
        .agg(
            imagen_odoo_base64_excel=("imagen_odoo_base64_excel", first_image_base64),
            fuente_imagen_excel=("fuente_imagen_excel", first_non_empty),
        )
    )
    print(f"Imágenes recuperadas desde Excel operativo: {len(out)} IQ")
    return out


imagenes_excel = extraer_imagenes_desde_hojas_operativas({
    "ventas_conjunto_detalle": ventas,
    "stock_sku_madre": stock_iq,
    "stock_por_sku": stock_por_sku,
    "rotacion_base": rotacion,
})


# ============================================================
# VENTAS ÚLTIMOS 3 MESES
# ============================================================

if "fecha" in ventas.columns:
    ventas["fecha"] = pd.to_datetime(ventas["fecha"], errors="coerce")
else:
    ventas["fecha"] = pd.NaT

if ventas["fecha"].notna().any():
    fecha_fin = ventas["fecha"].max().normalize()
else:
    fecha_fin = pd.Timestamp.today().normalize()

fecha_inicio_3m = fecha_fin - pd.Timedelta(days=DIAS_ANALISIS_3M - 1)

ventas["cantidad"] = to_num(ventas.get("cantidad", 0))
ventas["venta_total"] = to_num(ventas.get("venta_total", 0))

ventas_link = ventas[
    (ventas.get("tiene_referencia_madre", "") == "SI")
    & ventas["sku_madre"].notna()
    & ventas["sku_madre"].astype(str).str.strip().ne("")
].copy()

ventas_3m = ventas_link[
    (ventas_link["fecha"].notna())
    & (ventas_link["fecha"] >= fecha_inicio_3m)
    & (ventas_link["fecha"] <= fecha_fin + pd.Timedelta(days=1))
].copy()

# Ventas por IQ en 3M.
ventas_iq_3m = (
    ventas_3m.groupby("sku_madre", as_index=False)
    .agg(
        ventas_3m_unidades=("cantidad", "sum"),
        ventas_3m_monto=("venta_total", "sum"),
        pedidos_3m=("pedido", "nunique"),
        fuentes_venta=("fuente", lambda x: " | ".join(sorted(set(map(str, x))))),
        canales_venta=("canal", lambda x: " | ".join(sorted(set(map(str, x))))),
    )
)

# Promedio diario solo con días donde venta > 0.
if not ventas_3m.empty:
    tmp_daily = ventas_3m.copy()
    tmp_daily["dia"] = tmp_daily["fecha"].dt.date
    daily = tmp_daily.groupby(["sku_madre", "dia"], as_index=False).agg(unidades_dia=("cantidad", "sum"))
    daily_pos = daily[daily["unidades_dia"] > 0].copy()
    avg_pos = (
        daily_pos.groupby("sku_madre", as_index=False)
        .agg(
            dias_con_venta_3m=("dia", "nunique"),
            venta_diaria_promedio_positiva_3m=("unidades_dia", "mean"),
        )
    )
else:
    avg_pos = pd.DataFrame(columns=["sku_madre", "dias_con_venta_3m", "venta_diaria_promedio_positiva_3m"])

ventas_iq_3m = ventas_iq_3m.merge(avg_pos, on="sku_madre", how="left")
ventas_iq_3m["dias_con_venta_3m"] = to_num(ventas_iq_3m.get("dias_con_venta_3m", 0))
ventas_iq_3m["venta_diaria_promedio_positiva_3m"] = to_num(ventas_iq_3m.get("venta_diaria_promedio_positiva_3m", 0))

# Producto madre e imagen Odoo desde ventas.
if "producto_madre" in ventas_link.columns:
    if "imagen_odoo_base64" in ventas_link.columns:
        prod_ventas = (
            ventas_link.groupby("sku_madre", as_index=False)
            .agg(
                producto_madre=("producto_madre", first_non_empty),
                imagen_odoo_base64_ventas=("imagen_odoo_base64", first_image_base64),
            )
        )
    else:
        prod_ventas = (
            ventas_link.groupby("sku_madre", as_index=False)
            .agg(producto_madre=("producto_madre", first_non_empty))
        )
        prod_ventas["imagen_odoo_base64_ventas"] = ""
else:
    prod_ventas = pd.DataFrame(columns=["sku_madre", "producto_madre", "imagen_odoo_base64_ventas"])




# ============================================================
# ARRIBOS DESDE ODOO COMPRAS
# ============================================================

def detectar_iq_en_texto(texto):
    texto = str(texto or "")
    m = re.search(r"\b(IQ\d+)\b", texto, flags=re.IGNORECASE)
    if m:
        return m.group(1).upper()
    return ""


def fecha_corta(x):
    dt = pd.to_datetime(x, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%Y-%m-%d")


def extraer_arribos_odoo_compras(fecha_inicio=None, fecha_fin=None):
    """
    Extrae TODOS los arribos directamente de Odoo Compras.

    Fuente:
    - purchase.order.line
    - qty_received = cantidad recibida acumulada en la línea de compra

    Importante:
    - Esta función trae todos los renglones recibidos.
    - NO se limita al último arribo.
    - Después, otra función filtra cuáles son posteriores al corte para sumarlos al stock.
    """
    cols = [
        "fecha", "fecha_dt", "orden_compra", "proveedor", "referencia_interna",
        "producto", "cantidad_arribada", "cantidad_ordenada", "sku_odoo",
        "barcode", "estado_compra", "origen", "precio_unitario", "subtotal",
        "purchase_line_id", "product_id", "post_corte_stock"
    ]

    if not ENABLE_ODOO_ARRIBOS_COMPRAS:
        print("Arribos Odoo omitidos: ENABLE_ODOO_ARRIBOS_COMPRAS=False.")
        return pd.DataFrame(columns=cols)

    if not es_credencial_odoo_valida():
        print("Arribos Odoo omitidos: credenciales no configuradas.")
        return pd.DataFrame(columns=cols)

    try:
        odoo = OdooClient(ODOO_URL, ODOO_DB, ODOO_USER, ODOO_API_KEY)
        odoo.connect()

        # Base del dominio: TODOS los renglones de compras ya recibidos.
        domain = [
            ("order_id.state", "in", ODOO_ESTADOS_COMPRA_VALIDOS),
            ("product_id", "!=", False),
            ("qty_received", ">", 0),
        ]

        # Si quieres limitar por rango, apaga EXTRAER_TODOS_LOS_ARRIBOS_ODOO.
        if not EXTRAER_TODOS_LOS_ARRIBOS_ODOO:
            fecha_inicio = fecha_inicio or FECHA_INICIO_ARRIBOS_ODOO
            fecha_fin = fecha_fin or pd.Timestamp.today().normalize()
            dt_ini = pd.to_datetime(fecha_inicio).strftime("%Y-%m-%d 00:00:00")
            dt_fin = (pd.to_datetime(fecha_fin).normalize() + pd.Timedelta(days=1)).strftime("%Y-%m-%d 00:00:00")
            domain.extend([
                ("order_id.date_order", ">=", dt_ini),
                ("order_id.date_order", "<", dt_fin),
            ])

        fields_line = [
            "id",
            "order_id",
            "product_id",
            "name",
            "product_qty",
            "qty_received",
            "price_unit",
            "price_subtotal",
            "date_planned",
        ]

        lineas = odoo.search_read_all(
            "purchase.order.line",
            domain,
            fields_line,
            batch=2000,
            order="id asc"
        )

        if not lineas:
            print("No encontré arribos Odoo Compras con qty_received > 0.")
            return pd.DataFrame(columns=cols)

        order_ids = sorted({
            m2o_id(l.get("order_id"))
            for l in lineas
            if m2o_id(l.get("order_id"))
        })

        product_ids = sorted({
            m2o_id(l.get("product_id"))
            for l in lineas
            if m2o_id(l.get("product_id"))
        })

        order_fields = [
            "id",
            "name",
            "date_order",
            "date_approve",
            "effective_date",
            "state",
            "partner_id",
            "origin",
            "amount_total",
        ]

        try:
            orders = odoo.execute(
                "purchase.order",
                "read",
                order_ids,
                order_fields
            ) if order_ids else []
        except Exception:
            orders = odoo.execute(
                "purchase.order",
                "read",
                order_ids,
                ["id", "name", "date_order", "state", "partner_id", "origin", "amount_total"]
            ) if order_ids else []

        products = odoo.execute(
            "product.product",
            "read",
            product_ids,
            ["id", "display_name", "default_code", "barcode", "categ_id"]
        ) if product_ids else []

        order_map = {o["id"]: o for o in orders}
        product_map = {p["id"]: p for p in products}

        rows = []

        for l in lineas:
            order_id = m2o_id(l.get("order_id"))
            product_id = m2o_id(l.get("product_id"))

            o = order_map.get(order_id, {})
            p = product_map.get(product_id, {})

            producto_nombre = p.get("display_name") or l.get("name", "")
            default_code = limpiar_sku(p.get("default_code", ""))
            barcode = limpiar_sku(p.get("barcode", ""))

            referencia_interna = ""
            if re.fullmatch(r"IQ\d+", default_code.upper()):
                referencia_interna = default_code.upper()
            else:
                referencia_interna = (
                    detectar_iq_en_texto(producto_nombre)
                    or detectar_iq_en_texto(l.get("name", ""))
                    or default_code
                )

            # Odoo purchase.order.line no siempre tiene fecha real por recepción individual.
            # Para este primer cruce se usa la mejor fecha disponible:
            # effective_date > date_approve > date_order > date_planned
            fecha_base = (
                o.get("effective_date")
                or o.get("date_approve")
                or o.get("date_order")
                or l.get("date_planned")
            )
            fecha_dt = pd.to_datetime(fecha_base, errors="coerce")

            rows.append({
                "fecha": fecha_corta(fecha_dt),
                "fecha_dt": fecha_dt,
                "orden_compra": o.get("name", ""),
                "proveedor": m2o_name(o.get("partner_id")),
                "referencia_interna": str(referencia_interna).strip(),
                "producto": producto_nombre,
                "cantidad_arribada": float(l.get("qty_received") or 0),
                "cantidad_ordenada": float(l.get("product_qty") or 0),
                "sku_odoo": default_code,
                "barcode": barcode,
                "estado_compra": o.get("state", ""),
                "origen": o.get("origin", ""),
                "precio_unitario": float(l.get("price_unit") or 0),
                "subtotal": float(l.get("price_subtotal") or 0),
                "purchase_line_id": l.get("id"),
                "product_id": product_id,
                "post_corte_stock": "SI" if pd.notna(fecha_dt) and fecha_dt > FECHA_CORTE_STOCK else "NO",
            })

        df = pd.DataFrame(rows)

        if df.empty:
            return pd.DataFrame(columns=cols)

        df["referencia_interna"] = df["referencia_interna"].astype(str).str.strip()
        df = df[df["referencia_interna"].ne("")].copy()
        df["cantidad_arribada"] = to_num(df["cantidad_arribada"])
        df["cantidad_ordenada"] = to_num(df["cantidad_ordenada"])

        df = df.sort_values(["fecha_dt", "orden_compra", "referencia_interna"]).reset_index(drop=True)

        print(f"Arribos Odoo Compras extraídos TODOS: {len(df)}")
        print(f"Cantidad arribada total Odoo: {df['cantidad_arribada'].sum():,.0f}")
        print(f"Arribos post corte para stock: {(df['post_corte_stock'] == 'SI').sum()}")

        return df

    except Exception as e:
        print(f"No pude extraer arribos desde Odoo Compras: {e}")
        return pd.DataFrame(columns=cols)


def calcular_arribos_posteriores_corte_odoo():
    """
    Extrae TODOS los arribos de Odoo, pero para el cálculo de inventario
    solo suma los marcados como posteriores al corte.
    """
    arribos_todos = extraer_arribos_odoo_compras(FECHA_INICIO_ARRIBOS_ODOO, pd.Timestamp.today().normalize())

    cols_res = [
        "sku_madre", "arribos_post_corte_unidades", "arribos_post_corte_monto",
        "ordenes_compra_arribos", "proveedores_arribos", "primera_fecha_arribo",
        "ultima_fecha_arribo"
    ]

    if arribos_todos.empty:
        return pd.DataFrame(columns=cols_res), arribos_todos

    arribos = arribos_todos.copy()
    arribos["sku_madre"] = arribos["referencia_interna"].astype(str).str.strip().str.upper()
    arribos = arribos[arribos["sku_madre"].str.match(r"^IQ\d+$", na=False)].copy()

    # Solo estos se suman al stock congelado.
    arribos_post = arribos[arribos["post_corte_stock"].astype(str).str.upper().eq("SI")].copy()

    if arribos_post.empty:
        return pd.DataFrame(columns=cols_res), arribos_todos

    resumen = (
        arribos_post.groupby("sku_madre", as_index=False)
        .agg(
            arribos_post_corte_unidades=("cantidad_arribada", "sum"),
            arribos_post_corte_monto=("subtotal", "sum"),
            ordenes_compra_arribos=("orden_compra", lambda x: " | ".join(sorted(set(map(str, x))))),
            proveedores_arribos=("proveedor", lambda x: " | ".join(sorted(set(map(str, x))))),
            primera_fecha_arribo=("fecha", "min"),
            ultima_fecha_arribo=("fecha", "max"),
        )
    )

    return resumen, arribos_todos



# ============================================================
# RITMO DE VENTA AJUSTADO POR DÍAS CON STOCK
# ============================================================

def construir_ritmo_stock(
    stock_actual_iq,
    ventas_detalle,
    arribos_detalle,
    dic_map,
    fecha_fin=None,
    dias=90,
):
    """
    Calcula el ritmo de venta sin castigar los días sin inventario.

    Regla:
        ritmo_stock = ventas del periodo / días estimados con stock disponible

    Si el stock llega a 0:
    - los días posteriores con stock = 0 NO se agregan al denominador;
    - por lo tanto el ritmo queda congelado;
    - cuando vuelve a entrar inventario, el denominador vuelve a avanzar.

    La disponibilidad diaria se reconstruye hacia atrás desde el stock total
    actual usando:
        stock_inicio_dia ~= stock_fin_dia + ventas_dia - arribos_dia

    Los movimientos internos Odoo -> canal / Full no modifican el stock total
    de la empresa, por eso no afectan esta reconstrucción agregada por IQ.
    """
    cols_res = [
        "sku_madre",
        "ventas_periodo_ritmo",
        "dias_con_stock_90d",
        "dias_sin_stock_90d",
        "venta_diaria_stock_90d",
        "ritmo_congelado_stockout",
        "fecha_ultimo_dia_con_stock",
        "dias_desde_ultimo_stock",
        "metodo_ritmo_stock",
        "fuente_arribos_ritmo",
    ]
    cols_daily = [
        "fecha", "sku_madre", "stock_inicio_estimado", "arribos_dia",
        "ventas_dia", "stock_fin_estimado", "dia_con_stock"
    ]

    fecha_fin = pd.to_datetime(fecha_fin if fecha_fin is not None else pd.Timestamp.today()).normalize()
    dias = max(int(dias or 90), 1)
    fecha_inicio = fecha_fin - pd.Timedelta(days=dias - 1)
    fechas = pd.date_range(fecha_inicio, fecha_fin, freq="D")

    # ----------------------------
    # Stock actual por IQ
    # ----------------------------
    stock = stock_actual_iq.copy() if isinstance(stock_actual_iq, pd.DataFrame) else pd.DataFrame()
    if stock.empty or "sku_madre" not in stock.columns:
        stock_map = {}
    else:
        stock["sku_madre"] = stock["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        stock["stock_total"] = to_num(stock.get("stock_total", 0)).clip(lower=0)
        stock_map = stock.groupby("sku_madre")["stock_total"].sum().to_dict()

    # ----------------------------
    # Ventas diarias por IQ
    # ----------------------------
    ven = ventas_detalle.copy() if isinstance(ventas_detalle, pd.DataFrame) else pd.DataFrame()
    if ven.empty:
        ventas_daily = pd.DataFrame(columns=["sku_madre", "fecha", "ventas_dia"])
    else:
        ven["fecha"] = pd.to_datetime(ven.get("fecha", pd.NaT), errors="coerce").dt.normalize()
        ven["sku_madre"] = ven.get("sku_madre", "").fillna("").astype(str).str.strip().str.upper()
        ven["cantidad"] = to_num(ven.get("cantidad", 0)).clip(lower=0)
        ven = ven[
            ven["sku_madre"].ne("")
            & ven["fecha"].notna()
            & ven["fecha"].between(fecha_inicio, fecha_fin)
        ].copy()
        ventas_daily = (
            ven.groupby(["sku_madre", "fecha"], as_index=False)["cantidad"]
            .sum()
            .rename(columns={"cantidad": "ventas_dia"})
        )

    # ----------------------------
    # Arribos diarios por IQ
    # ----------------------------
    arr = arribos_detalle.copy() if isinstance(arribos_detalle, pd.DataFrame) else pd.DataFrame()
    fuente_arribos = "ODOO_COMPRAS" if not arr.empty else "SIN_ARRIBOS_DETECTADOS"
    if arr.empty:
        arribos_daily = pd.DataFrame(columns=["sku_madre", "fecha", "arribos_dia"])
    else:
        fecha_col = "fecha_dt" if "fecha_dt" in arr.columns else "fecha"
        arr["fecha"] = pd.to_datetime(arr.get(fecha_col, pd.NaT), errors="coerce").dt.normalize()
        arr["cantidad_arribada"] = to_num(arr.get("cantidad_arribada", 0)).clip(lower=0)

        def resolver_iq_arribo(row):
            candidatos = []
            for c in ["referencia_interna", "sku_odoo", "barcode"]:
                v = limpiar_sku(row.get(c, ""))
                if v and v not in candidatos:
                    candidatos.append(v)
            for v in candidatos:
                up = str(v).strip().upper()
                if re.fullmatch(r"IQ\d+", up):
                    return up
                iq = buscar_iq_por_sku(v, dic_map)
                if iq:
                    return str(iq).strip().upper()
            return ""

        arr["sku_madre"] = arr.apply(resolver_iq_arribo, axis=1)
        arr = arr[
            arr["sku_madre"].ne("")
            & arr["fecha"].notna()
            & arr["fecha"].between(fecha_inicio, fecha_fin)
        ].copy()
        arribos_daily = (
            arr.groupby(["sku_madre", "fecha"], as_index=False)["cantidad_arribada"]
            .sum()
            .rename(columns={"cantidad_arribada": "arribos_dia"})
        )

    venta_lookup = {
        (str(r["sku_madre"]), pd.Timestamp(r["fecha"])): float(r["ventas_dia"] or 0)
        for _, r in ventas_daily.iterrows()
    }
    arribo_lookup = {
        (str(r["sku_madre"]), pd.Timestamp(r["fecha"])): float(r["arribos_dia"] or 0)
        for _, r in arribos_daily.iterrows()
    }

    universo = set(stock_map)
    universo |= set(ventas_daily["sku_madre"].astype(str)) if not ventas_daily.empty else set()
    universo |= set(arribos_daily["sku_madre"].astype(str)) if not arribos_daily.empty else set()

    resumen_rows = []
    diario_rows = []

    for sku in sorted(s for s in universo if str(s).strip()):
        stock_fin = max(float(stock_map.get(sku, 0) or 0), 0.0)
        registros_sku = []

        # Reconstruimos hacia atrás desde el stock actual.
        for fecha in reversed(fechas):
            fecha = pd.Timestamp(fecha)
            venta_dia = max(float(venta_lookup.get((sku, fecha), 0) or 0), 0.0)
            arribo_dia = max(float(arribo_lookup.get((sku, fecha), 0) or 0), 0.0)

            stock_inicio_raw = stock_fin + venta_dia - arribo_dia
            stock_inicio = max(stock_inicio_raw, 0.0)

            # Hubo disponibilidad si arrancó con piezas, entraron piezas o hubo venta.
            # "venta_dia > 0" también protege contra pequeñas inconsistencias de fecha.
            con_stock = int(stock_inicio > 1e-9 or arribo_dia > 1e-9 or venta_dia > 1e-9)

            registros_sku.append({
                "fecha": fecha,
                "sku_madre": sku,
                "stock_inicio_estimado": stock_inicio,
                "arribos_dia": arribo_dia,
                "ventas_dia": venta_dia,
                "stock_fin_estimado": stock_fin,
                "dia_con_stock": con_stock,
            })

            stock_fin = stock_inicio

        registros_sku.reverse()
        diario_rows.extend(registros_sku)

        dias_con_stock = int(sum(x["dia_con_stock"] for x in registros_sku))
        ventas_periodo = float(sum(x["ventas_dia"] for x in registros_sku))
        ritmo = ventas_periodo / dias_con_stock if dias_con_stock > 0 else 0.0
        dias_sin_stock = max(len(fechas) - dias_con_stock, 0)

        fechas_stock = [x["fecha"] for x in registros_sku if x["dia_con_stock"] == 1]
        ultimo_stock = max(fechas_stock) if fechas_stock else pd.NaT
        stock_actual = max(float(stock_map.get(sku, 0) or 0), 0.0)

        congelado = "SI" if stock_actual <= 1e-9 and ritmo > 0 else "NO"
        dias_desde = (
            int((fecha_fin - ultimo_stock).days)
            if pd.notna(ultimo_stock) else np.nan
        )

        resumen_rows.append({
            "sku_madre": sku,
            "ventas_periodo_ritmo": ventas_periodo,
            "dias_con_stock_90d": dias_con_stock,
            "dias_sin_stock_90d": dias_sin_stock,
            "venta_diaria_stock_90d": ritmo,
            "ritmo_congelado_stockout": congelado,
            "fecha_ultimo_dia_con_stock": ultimo_stock,
            "dias_desde_ultimo_stock": dias_desde,
            "metodo_ritmo_stock": "VENTAS / DIAS_CON_STOCK_ESTIMADO",
            "fuente_arribos_ritmo": fuente_arribos,
        })

    return (
        pd.DataFrame(resumen_rows, columns=cols_res),
        pd.DataFrame(diario_rows, columns=cols_daily),
    )



# ============================================================
# STOCK CONGELADO / DESCUENTO DE VENTAS POSTERIORES
# ============================================================

def normalizar_stock_base_para_congelar(stock_iq_original):
    """
    Toma el stock por IQ del Excel operativo y lo convierte en la foto oficial
    de inventario congelado.
    """
    df = stock_iq_original.copy()

    if df.empty:
        return pd.DataFrame(columns=[
            "sku_madre", "producto_madre_stock",
            "stock_walmart_wfs_congelado",
            "stock_liverpool_99min_congelado",
            "stock_meli_full_congelado",
            "stock_amazon_fba_congelado",
            "stock_odoo_cuautitlan_congelado",
            "stock_total_congelado",
            "fecha_corte_stock",
        ])

    if "sku_madre" not in df.columns:
        df["sku_madre"] = ""

    for col in [
        "stock_walmart_wfs",
        "stock_liverpool_99min",
        "stock_meli_full",
        "stock_amazon_fba",
        "stock_odoo_cuautitlan",
        "stock_total",
    ]:
        if col not in df.columns:
            df[col] = 0
        df[col] = to_num(df[col])

    if "producto_madre" not in df.columns:
        df["producto_madre"] = ""

    base = (
        df.groupby("sku_madre", as_index=False)
        .agg(
            producto_madre_stock=("producto_madre", first_non_empty),
            stock_walmart_wfs_congelado=("stock_walmart_wfs", "sum"),
            stock_liverpool_99min_congelado=("stock_liverpool_99min", "sum"),
            stock_meli_full_congelado=("stock_meli_full", "sum"),
            stock_amazon_fba_congelado=("stock_amazon_fba", "sum"),
            stock_odoo_cuautitlan_congelado=("stock_odoo_cuautitlan", "sum"),
            stock_total_congelado=("stock_total", "sum"),
        )
    )

    base = base[base["sku_madre"].astype(str).str.strip().ne("")].copy()
    base["fecha_corte_stock"] = FECHA_CORTE_STOCK
    base["nota"] = "Stock congelado oficial antes de descontar ventas posteriores al corte"

    return base


def cargar_o_crear_stock_congelado(stock_iq_original):
    """
    Si existe ARCHIVO_STOCK_CONGELADO, lo usa como inventario base fijo.
    Si no existe, lo crea con el stock que venga en el Excel operativo actual.
    """
    if not USAR_STOCK_CONGELADO:
        return None

    if ARCHIVO_STOCK_CONGELADO.exists():
        base = pd.read_excel(ARCHIVO_STOCK_CONGELADO)
        base.columns = [str(c).strip() for c in base.columns]
        print(f"Stock congelado cargado: {ARCHIVO_STOCK_CONGELADO}")
    else:
        base = normalizar_stock_base_para_congelar(stock_iq_original)
        with pd.ExcelWriter(ARCHIVO_STOCK_CONGELADO, engine="openpyxl") as writer:
            base.to_excel(writer, sheet_name="stock_congelado_iq", index=False)
        print(f"Stock congelado creado: {ARCHIVO_STOCK_CONGELADO}")

    for c in [
        "stock_walmart_wfs_congelado",
        "stock_liverpool_99min_congelado",
        "stock_meli_full_congelado",
        "stock_amazon_fba_congelado",
        "stock_odoo_cuautitlan_congelado",
        "stock_total_congelado",
    ]:
        if c not in base.columns:
            base[c] = 0
        base[c] = to_num(base[c])

    if "sku_madre" not in base.columns:
        base["sku_madre"] = ""

    return base


def calcular_ventas_posteriores_corte(ventas_df):
    """
    Calcula unidades vendidas por IQ después del corte.
    Solo descuenta ventas ya vinculadas a sku_madre.
    """
    if ventas_df.empty:
        return pd.DataFrame(columns=[
            "sku_madre", "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte"
        ])

    df = ventas_df.copy()

    if "fecha" not in df.columns:
        return pd.DataFrame(columns=[
            "sku_madre", "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte"
        ])

    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["cantidad"] = to_num(df.get("cantidad", 0))
    df["venta_total"] = to_num(df.get("venta_total", 0))

    if "tiene_referencia_madre" in df.columns:
        df = df[df["tiene_referencia_madre"].astype(str).str.upper().eq("SI")].copy()

    df = df[
        df["sku_madre"].notna()
        & df["sku_madre"].astype(str).str.strip().ne("")
        & df["fecha"].notna()
        & (df["fecha"] > FECHA_CORTE_STOCK)
    ].copy()

    if df.empty:
        return pd.DataFrame(columns=[
            "sku_madre", "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte"
        ])

    return (
        df.groupby("sku_madre", as_index=False)
        .agg(
            ventas_post_corte_unidades=("cantidad", "sum"),
            ventas_post_corte_monto=("venta_total", "sum"),
            pedidos_post_corte=("pedido", "nunique"),
        )
    )


def aplicar_descuento_stock_congelado(stock_iq_original, ventas_df):
    """
    Regresa stock_iq ajustado:
    stock actual calculado =
        stock congelado
        + arribos Odoo Compras posteriores al corte
        - ventas posteriores al corte.

    Para conservar el desglose por canal, el stock total calculado se distribuye
    proporcionalmente sobre los canales del stock congelado. El ajuste por
    redondeo se manda a Odoo Cuautitlán.
    """
    if not USAR_STOCK_CONGELADO:
        return stock_iq_original.copy(), pd.DataFrame(), pd.DataFrame()

    stock_congelado = cargar_o_crear_stock_congelado(stock_iq_original)
    ventas_post = calcular_ventas_posteriores_corte(ventas_df)
    arribos_post, arribos_odoo_detalle = calcular_arribos_posteriores_corte_odoo()

    base = stock_congelado.merge(ventas_post, on="sku_madre", how="left")
    base = base.merge(arribos_post, on="sku_madre", how="left")

    for c in [
        "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte",
        "arribos_post_corte_unidades", "arribos_post_corte_monto",
    ]:
        if c not in base.columns:
            base[c] = 0
        base[c] = to_num(base[c])

    for c in [
        "ordenes_compra_arribos", "proveedores_arribos",
        "primera_fecha_arribo", "ultima_fecha_arribo"
    ]:
        if c not in base.columns:
            base[c] = ""

    base["stock_total_antes_movimientos"] = base["stock_total_congelado"]
    base["stock_total_antes_descuento"] = base["stock_total_congelado"] + base["arribos_post_corte_unidades"]

    base["stock_total"] = np.maximum(
        base["stock_total_congelado"]
        + base["arribos_post_corte_unidades"]
        - base["ventas_post_corte_unidades"],
        0
    )

    # Factor proporcional para ajustar stock de canales sin inventar asignación exacta por canal.
    base["base_para_distribucion"] = base["stock_total_congelado"] + base["arribos_post_corte_unidades"]

    base["factor_stock_restante"] = np.where(
        base["base_para_distribucion"] > 0,
        base["stock_total"] / base["base_para_distribucion"],
        0
    )

    # Arribos se agregan provisionalmente a Odoo Cuautitlán antes de distribuir.
    base["stock_walmart_wfs_base_mov"] = base.get("stock_walmart_wfs_congelado", 0)
    base["stock_liverpool_99min_base_mov"] = base.get("stock_liverpool_99min_congelado", 0)
    base["stock_meli_full_base_mov"] = base.get("stock_meli_full_congelado", 0)
    base["stock_amazon_fba_base_mov"] = base.get("stock_amazon_fba_congelado", 0)
    base["stock_odoo_cuautitlan_base_mov"] = (
        base.get("stock_odoo_cuautitlan_congelado", 0)
        + base["arribos_post_corte_unidades"]
    )

    canales = [
        ("stock_walmart_wfs_base_mov", "stock_walmart_wfs"),
        ("stock_liverpool_99min_base_mov", "stock_liverpool_99min"),
        ("stock_meli_full_base_mov", "stock_meli_full"),
        ("stock_amazon_fba_base_mov", "stock_amazon_fba"),
        ("stock_odoo_cuautitlan_base_mov", "stock_odoo_cuautitlan"),
    ]

    for base_col, actual in canales:
        if base_col not in base.columns:
            base[base_col] = 0
        base[actual] = np.maximum(
            np.floor(to_num(base[base_col]) * base["factor_stock_restante"]),
            0
        )

    # Ajuste por redondeo: suma de canales debe coincidir con stock_total.
    canal_actual_cols = [actual for _, actual in canales]
    base["_suma_canales"] = base[canal_actual_cols].sum(axis=1)
    base["_delta_redondeo"] = base["stock_total"] - base["_suma_canales"]
    base["stock_odoo_cuautitlan"] = np.maximum(
        base["stock_odoo_cuautitlan"] + base["_delta_redondeo"],
        0
    )
    base = base.drop(columns=["_suma_canales", "_delta_redondeo"], errors="ignore")

    base["producto_madre"] = base.get("producto_madre_stock", "")
    base["stock_calculado_desde_congelado"] = "SI"
    base["fecha_corte_stock"] = FECHA_CORTE_STOCK
    base["formula_stock"] = "NO_APLICA_EN_02; stock actual proviene del script 01"

    cols = [
        "sku_madre", "producto_madre",
        "stock_walmart_wfs", "stock_liverpool_99min", "stock_meli_full",
        "stock_amazon_fba", "stock_odoo_cuautitlan", "stock_total",
        "stock_total_congelado", "arribos_post_corte_unidades", "arribos_post_corte_monto",
        "ordenes_compra_arribos", "proveedores_arribos", "primera_fecha_arribo", "ultima_fecha_arribo",
        "ventas_post_corte_unidades", "ventas_post_corte_monto",
        "pedidos_post_corte", "stock_total_antes_movimientos", "stock_total_antes_descuento",
        "stock_calculado_desde_congelado", "fecha_corte_stock", "formula_stock"
    ]
    for c in cols:
        if c not in base.columns:
            base[c] = ""

    auditoria = base.copy()

    return base[cols].copy(), auditoria, arribos_odoo_detalle



# ============================================================
# STOCK
# ============================================================

if stock_iq.empty and not rotacion.empty:
    stock_iq = rotacion.copy()

# El ajuste congelado queda desactivado: el stock actual ya viene resuelto desde 01.
stock_iq, auditoria_stock_congelado, arribos_odoo_compras_detalle = aplicar_descuento_stock_congelado(stock_iq, ventas)

# Para el ritmo ajustado por stock necesitamos conocer entradas de inventario.
# Si el cálculo de stock congelado no las cargó (caso habitual con
# USAR_STOCK_CONGELADO=False), se consultan aquí solo para reconstruir
# disponibilidad histórica; NO modifican el stock actual.
if not isinstance(arribos_odoo_compras_detalle, pd.DataFrame) or arribos_odoo_compras_detalle.empty:
    arribos_odoo_compras_detalle = extraer_arribos_odoo_compras(
        pd.Timestamp.today().normalize() - pd.Timedelta(days=DIAS_RITMO_STOCK + 30),
        pd.Timestamp.today().normalize(),
    )

ritmo_stock_90d, ritmo_stock_diario = construir_ritmo_stock(
    stock_actual_iq=stock_iq,
    ventas_detalle=ventas_link,
    arribos_detalle=arribos_odoo_compras_detalle,
    dic_map=dic_map,
    fecha_fin=pd.Timestamp.today().normalize(),
    dias=DIAS_RITMO_STOCK,
)


def proveedor_mas_reciente_desde_arribos(df):
    """
    Regresa el proveedor del arribo/compra más reciente.

    Acepta:
    - DataFrame completo con columnas fecha/proveedor/orden_compra.
    - Series de proveedor, cuando se usa dentro de groupby.agg.
    """
    if df is None:
        return ""

    # Caso Series: ocurre cuando pandas llama la función dentro de groupby.agg.
    if isinstance(df, pd.Series):
        prov = df.dropna().astype(str).str.strip()
        prov = prov[~prov.str.lower().isin(["", "nan", "none", "false"])]
        if prov.empty:
            return ""
        return prov.iloc[-1]

    if df.empty:
        return ""

    tmp = df.copy()
    if "fecha" not in tmp.columns:
        if "proveedor" in tmp.columns:
            return proveedor_mas_reciente_desde_arribos(tmp["proveedor"])
        return ""

    tmp["_fecha_sort"] = pd.to_datetime(tmp["fecha"], errors="coerce")
    tmp["_orden_sort"] = tmp.get("orden_compra", "").astype(str) if "orden_compra" in tmp.columns else ""
    tmp = tmp.sort_values(["_fecha_sort", "_orden_sort"], ascending=[True, True])

    if "proveedor" not in tmp.columns:
        return ""

    prov = tmp["proveedor"].dropna().astype(str).str.strip()
    prov = prov[~prov.str.lower().isin(["", "nan", "none", "false"])]
    if prov.empty:
        return ""
    return prov.iloc[-1]


# ============================================================
# PROVEEDORES POR PRODUCTO DESDE ARRIBOS ODOO
# ============================================================
# Se usa TODOS los arribos de Odoo, no solo post-corte.
# Esto permite mostrar el proveedor en Detalle por producto aunque el arribo sea anterior al corte.
if "arribos_odoo_compras_detalle" in globals() and isinstance(arribos_odoo_compras_detalle, pd.DataFrame) and not arribos_odoo_compras_detalle.empty:
    prov_tmp = arribos_odoo_compras_detalle.copy()

    if "referencia_interna" in prov_tmp.columns:
        prov_tmp["sku_madre"] = prov_tmp["referencia_interna"].astype(str).str.strip().str.upper()
    elif "sku_madre" not in prov_tmp.columns:
        prov_tmp["sku_madre"] = ""

    for c in ["proveedor", "orden_compra", "fecha", "cantidad_arribada"]:
        if c not in prov_tmp.columns:
            prov_tmp[c] = ""

    prov_tmp["cantidad_arribada"] = to_num(prov_tmp["cantidad_arribada"])
    prov_tmp = prov_tmp[prov_tmp["sku_madre"].str.match(r"^IQ\d+$", na=False)].copy()

    if not prov_tmp.empty:
        proveedores_producto = (
            prov_tmp.groupby("sku_madre", as_index=False)
            .apply(lambda g: pd.Series({
                "proveedor_principal": proveedor_mas_reciente_desde_arribos(g),
                "proveedores_odoo_compras": " | ".join(sorted(set([
                    str(v).strip() for v in g["proveedor"]
                    if str(v).strip() and str(v).lower() not in ["nan", "none", "false"]
                ]))),
                "ordenes_compra_odoo": " | ".join(sorted(set([
                    str(v).strip() for v in g["orden_compra"]
                    if str(v).strip() and str(v).lower() not in ["nan", "none", "false"]
                ]))[:30]),
                "primera_fecha_compra_odoo": pd.to_datetime(g["fecha"], errors="coerce").min().strftime("%Y-%m-%d") if pd.to_datetime(g["fecha"], errors="coerce").notna().any() else "",
                "ultima_fecha_compra_odoo": pd.to_datetime(g["fecha"], errors="coerce").max().strftime("%Y-%m-%d") if pd.to_datetime(g["fecha"], errors="coerce").notna().any() else "",
                "cantidad_arribada_historica": g["cantidad_arribada"].sum(),
            }))
            .reset_index()
        )
    else:
        proveedores_producto = pd.DataFrame(columns=[
            "sku_madre", "proveedor_principal", "proveedores_odoo_compras",
            "ordenes_compra_odoo", "primera_fecha_compra_odoo",
            "ultima_fecha_compra_odoo", "cantidad_arribada_historica"
        ])
else:
    proveedores_producto = pd.DataFrame(columns=[
        "sku_madre", "proveedor_principal", "proveedores_odoo_compras",
        "ordenes_compra_odoo", "primera_fecha_compra_odoo",
        "ultima_fecha_compra_odoo", "cantidad_arribada_historica"
    ])


# ============================================================
# SELL-THROUGH ÚLTIMOS 7 DÍAS VS ÚLTIMO LOTE
# ============================================================
# Métrica interna:
# ventas_ultimos_7d / cantidad_ultimo_lote
# Si es >= 30%, la recomendación será "Recompra inmediata".
# Esto evalúa ventas recientes, no ventas acumuladas desde el arribo.
if "arribos_odoo_compras_detalle" in globals() and isinstance(arribos_odoo_compras_detalle, pd.DataFrame) and not arribos_odoo_compras_detalle.empty:
    lotes_tmp = arribos_odoo_compras_detalle.copy()

    if "referencia_interna" in lotes_tmp.columns:
        lotes_tmp["sku_madre"] = lotes_tmp["referencia_interna"].astype(str).str.strip().str.upper()
    elif "sku_madre" not in lotes_tmp.columns:
        lotes_tmp["sku_madre"] = ""

    for c in ["fecha", "cantidad_arribada", "proveedor", "orden_compra"]:
        if c not in lotes_tmp.columns:
            lotes_tmp[c] = ""

    lotes_tmp["fecha_arribo_dt"] = pd.to_datetime(lotes_tmp["fecha"], errors="coerce")
    lotes_tmp["cantidad_arribada"] = to_num(lotes_tmp["cantidad_arribada"])
    lotes_tmp = lotes_tmp[
        lotes_tmp["sku_madre"].str.match(r"^IQ\d+$", na=False)
        & lotes_tmp["fecha_arribo_dt"].notna()
        & (lotes_tmp["cantidad_arribada"] > 0)
    ].copy()

    if not lotes_tmp.empty:
        ult_fecha = (
            lotes_tmp.groupby("sku_madre", as_index=False)
            .agg(fecha_ultimo_lote=("fecha_arribo_dt", "max"))
        )
        lotes_ult = lotes_tmp.merge(ult_fecha, on="sku_madre", how="inner")
        lotes_ult = lotes_ult[lotes_ult["fecha_arribo_dt"].eq(lotes_ult["fecha_ultimo_lote"])].copy()

        ultimo_lote = (
            lotes_ult.groupby("sku_madre", as_index=False)
            .agg(
                fecha_ultimo_lote=("fecha_ultimo_lote", "max"),
                cantidad_ultimo_lote=("cantidad_arribada", "sum"),
                proveedor_ultimo_lote=("proveedor", lambda x: proveedor_mas_reciente_desde_arribos(x)),
                ordenes_ultimo_lote=("orden_compra", lambda x: " | ".join(sorted(set([str(v).strip() for v in x if str(v).strip() and str(v).lower() not in ["nan", "none", "false"]])))),
            )
        )
    else:
        ultimo_lote = pd.DataFrame(columns=[
            "sku_madre", "fecha_ultimo_lote", "cantidad_ultimo_lote",
            "proveedor_ultimo_lote", "ordenes_ultimo_lote"
        ])
else:
    ultimo_lote = pd.DataFrame(columns=[
        "sku_madre", "fecha_ultimo_lote", "cantidad_ultimo_lote",
        "proveedor_ultimo_lote", "ordenes_ultimo_lote"
    ])

if not ultimo_lote.empty and "ventas_link" in globals() and isinstance(ventas_link, pd.DataFrame) and not ventas_link.empty:
    ventas_tmp_st = ventas_link.copy()
    ventas_tmp_st["fecha"] = pd.to_datetime(ventas_tmp_st.get("fecha", pd.NaT), errors="coerce")
    ventas_tmp_st["cantidad"] = to_num(ventas_tmp_st.get("cantidad", 0))
    ventas_tmp_st["sku_madre"] = ventas_tmp_st["sku_madre"].astype(str).str.strip().str.upper()

    ventas_lote = ventas_tmp_st.merge(
        ultimo_lote[["sku_madre", "fecha_ultimo_lote", "cantidad_ultimo_lote"]],
        on="sku_madre",
        how="inner"
    )

    # Tomar únicamente ventas recientes de los últimos 7 días.
    # Se usa fecha_fin del análisis si existe; si no, la fecha máxima de ventas.
    fecha_ref_7d = pd.to_datetime(fecha_fin, errors="coerce") if "fecha_fin" in globals() else pd.NaT
    if pd.isna(fecha_ref_7d):
        fecha_ref_7d = ventas_tmp_st["fecha"].max()
    fecha_inicio_7d = pd.to_datetime(fecha_ref_7d).normalize() - pd.Timedelta(days=7)

    ventas_lote = ventas_lote[
        ventas_lote["fecha"].notna()
        & (ventas_lote["fecha"] > fecha_inicio_7d)
        & (ventas_lote["fecha"] <= pd.to_datetime(fecha_ref_7d))
    ].copy()

    ventas_7d_lote = (
        ventas_lote.groupby("sku_madre", as_index=False)
        .agg(ventas_ultimos_7d=("cantidad", "sum"))
    )
else:
    ventas_7d_lote = pd.DataFrame(columns=["sku_madre", "ventas_ultimos_7d"])

sell_through_lote = ultimo_lote.merge(ventas_7d_lote, on="sku_madre", how="left")
if sell_through_lote.empty:
    sell_through_lote = pd.DataFrame(columns=[
        "sku_madre", "fecha_ultimo_lote", "cantidad_ultimo_lote",
        "ventas_ultimos_7d", "sell_through_ultimos_7d_vs_lote",
        "recompra_inmediata_flag"
    ])
else:
    sell_through_lote["ventas_ultimos_7d"] = to_num(sell_through_lote.get("ventas_ultimos_7d", 0))
    sell_through_lote["cantidad_ultimo_lote"] = to_num(sell_through_lote.get("cantidad_ultimo_lote", 0))
    sell_through_lote["sell_through_ultimos_7d_vs_lote"] = np.where(
        sell_through_lote["cantidad_ultimo_lote"] > 0,
        sell_through_lote["ventas_ultimos_7d"] / sell_through_lote["cantidad_ultimo_lote"],
        0
    )
    sell_through_lote["recompra_inmediata_flag"] = np.where(
        sell_through_lote["sell_through_ultimos_7d_vs_lote"] >= 0.30,
        "SI",
        "NO"
    )

for col in [
    "stock_walmart_wfs",
    "stock_liverpool_99min",
    "stock_meli_full",
    "stock_amazon_fba",
    "stock_odoo_cuautitlan",
    "stock_total",
]:
    if col not in stock_iq.columns:
        stock_iq[col] = 0
    stock_iq[col] = to_num(stock_iq[col])

if "sku_madre" not in stock_iq.columns:
    stock_iq["sku_madre"] = ""

if "producto_madre" not in stock_iq.columns:
    stock_iq["producto_madre"] = ""

if "imagen_odoo_base64" not in stock_iq.columns:
    stock_iq["imagen_odoo_base64"] = ""

stock_base = (
    stock_iq.groupby("sku_madre", as_index=False)
    .agg(
        producto_madre_stock=("producto_madre", first_non_empty),
        imagen_odoo_base64_stock=("imagen_odoo_base64", first_image_base64),
        stock_walmart_wfs=("stock_walmart_wfs", "sum"),
        stock_liverpool_99min=("stock_liverpool_99min", "sum"),
        stock_meli_full=("stock_meli_full", "sum"),
        stock_amazon_fba=("stock_amazon_fba", "sum"),
        stock_odoo_cuautitlan=("stock_odoo_cuautitlan", "sum"),
        stock_total=("stock_total", "sum"),
    )
)

# Stock no vinculado.
stock_no_vinculado = stock_sin_ref.copy()
if not stock_no_vinculado.empty:
    for c in ["stock_total", "WALMART_WFS", "LIVERPOOL_FULL_99MIN", "MERCADO_LIBRE_FULL", "AMAZON_FBA", "ODOO_CUAUTITLAN"]:
        if c in stock_no_vinculado.columns:
            stock_no_vinculado[c] = to_num(stock_no_vinculado[c])

    # Solo conservar SKUs con stock total real mayor a 0.
    if "stock_total" in stock_no_vinculado.columns:
        stock_no_vinculado = stock_no_vinculado[
            stock_no_vinculado["stock_total"] > 0
        ].copy()

    stock_no_vinculado["motivo_no_vinculado"] = "Stock sin referencia madre IQ"
    stock_no_vinculado["accion_sugerida"] = "Agregar alias a origen4 o corregir SKU de inventario"

stock_no_vinculado_skus = stock_no_vinculado["sku_original"].nunique() if "sku_original" in stock_no_vinculado.columns and not stock_no_vinculado.empty else 0
stock_no_vinculado_unidades = stock_no_vinculado["stock_total"].sum() if "stock_total" in stock_no_vinculado.columns and not stock_no_vinculado.empty else 0


# ============================================================
# BASE DASHBOARD
# ============================================================

all_iq = pd.DataFrame({
    "sku_madre": sorted(
        set(stock_base["sku_madre"].dropna().astype(str))
        | set(ventas_iq_3m["sku_madre"].dropna().astype(str))
    )
})
all_iq = all_iq[all_iq["sku_madre"].str.strip().ne("")].copy()

dashboard = all_iq.merge(stock_base, on="sku_madre", how="left")

# Adjuntar auditoría de stock congelado al producto.
if USAR_STOCK_CONGELADO and "auditoria_stock_congelado" in globals() and not auditoria_stock_congelado.empty:
    cols_aud = [
        "sku_madre", "stock_total_congelado",
        "arribos_post_corte_unidades", "arribos_post_corte_monto",
        "ordenes_compra_arribos", "proveedores_arribos",
        "primera_fecha_arribo", "ultima_fecha_arribo",
        "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte",
        "stock_total_antes_movimientos", "stock_total_antes_descuento",
        "stock_calculado_desde_congelado", "fecha_corte_stock", "formula_stock"
    ]
    cols_aud = [c for c in cols_aud if c in auditoria_stock_congelado.columns]
    dashboard = dashboard.merge(
        auditoria_stock_congelado[cols_aud].drop_duplicates("sku_madre"),
        on="sku_madre",
        how="left"
    )

dashboard = dashboard.merge(ventas_iq_3m, on="sku_madre", how="left")
if isinstance(ritmo_stock_90d, pd.DataFrame) and not ritmo_stock_90d.empty:
    dashboard = dashboard.merge(ritmo_stock_90d, on="sku_madre", how="left")
dashboard = dashboard.merge(prod_ventas, on="sku_madre", how="left")
dashboard = dashboard.merge(sku_chips, on="sku_madre", how="left")
dashboard = dashboard.merge(sku_por_canal, on="sku_madre", how="left")
dashboard = dashboard.merge(proveedores_producto, on="sku_madre", how="left")
dashboard = dashboard.merge(sell_through_lote, on="sku_madre", how="left")

dashboard["producto_madre"] = dashboard["producto_madre"].fillna(dashboard.get("producto_madre_stock", ""))
dashboard["producto_madre"] = dashboard["producto_madre"].fillna("")

# Imagen Odoo para detalle de producto.
# Prioridad:
# 1) Imagen que ya venga del Excel operativo.
# 2) Imagen recuperada directamente desde API de Odoo por IQ.
# 3) Imagen desde ventas/stock si ya existía.
for c in ["imagen_odoo_base64_ventas", "imagen_odoo_base64_stock"]:
    if c not in dashboard.columns:
        dashboard[c] = ""

if "imagenes_excel" in globals() and isinstance(imagenes_excel, pd.DataFrame) and not imagenes_excel.empty:
    dashboard = dashboard.merge(imagenes_excel, on="sku_madre", how="left")
else:
    dashboard["imagen_odoo_base64_excel"] = ""
    dashboard["fuente_imagen_excel"] = ""

imagenes_api = extraer_imagenes_odoo_por_iq(dashboard["sku_madre"].dropna().astype(str).tolist())
if not imagenes_api.empty:
    dashboard = dashboard.merge(imagenes_api, on="sku_madre", how="left")
else:
    dashboard["imagen_odoo_base64_api"] = ""
    dashboard["producto_imagen_odoo"] = ""

for c in [
    "imagen_odoo_base64_excel",
    "imagen_odoo_base64_api",
    "imagen_odoo_base64_ventas",
    "imagen_odoo_base64_stock",
]:
    if c not in dashboard.columns:
        dashboard[c] = ""

dashboard["imagen_odoo_base64"] = dashboard[
    [
        "imagen_odoo_base64_excel",
        "imagen_odoo_base64_api",
        "imagen_odoo_base64_ventas",
        "imagen_odoo_base64_stock",
    ]
].apply(first_image_base64, axis=1)

dashboard["tiene_imagen_odoo"] = np.where(
    dashboard["imagen_odoo_base64"].astype(str).str.strip().ne(""),
    "SI",
    "NO"
)

for col in [
    "ventas_3m_unidades",
    "ventas_3m_monto",
    "pedidos_3m",
    "dias_con_venta_3m",
    "venta_diaria_promedio_positiva_3m",
    "ventas_periodo_ritmo",
    "dias_con_stock_90d",
    "dias_sin_stock_90d",
    "venta_diaria_stock_90d",
    "dias_desde_ultimo_stock",
    "stock_walmart_wfs",
    "stock_liverpool_99min",
    "stock_meli_full",
    "stock_amazon_fba",
    "stock_odoo_cuautitlan",
    "stock_total",
    "stock_total_congelado",
    "arribos_post_corte_unidades",
    "arribos_post_corte_monto",
    "ventas_post_corte_unidades",
    "ventas_post_corte_monto",
    "pedidos_post_corte",
    "stock_total_antes_movimientos",
    "stock_total_antes_descuento",
    "num_skus_sincronizados",
    "cantidad_arribada_historica",
    "dias_desde_ultimo_arribo",
    "cantidad_ultimo_lote",
    "ventas_ultimos_7d",
    "sell_through_ultimos_7d_vs_lote",
]:
    if col not in dashboard.columns:
        dashboard[col] = 0
    dashboard[col] = to_num(dashboard[col])

# Ritmo oficial de producto: ventas / días con stock estimado.
# Si no se pudo reconstruir disponibilidad, conserva como respaldo el cálculo
# previo de días con venta positiva para no dejar el dashboard sin dato.
dashboard["venta_diaria_promedio_3m"] = np.where(
    dashboard["dias_con_stock_90d"] > 0,
    dashboard["venta_diaria_stock_90d"],
    dashboard["venta_diaria_promedio_positiva_3m"],
)

dashboard["dias_inventario"] = np.where(
    dashboard["venta_diaria_promedio_3m"] > 0,
    dashboard["stock_total"] / dashboard["venta_diaria_promedio_3m"],
    np.where(dashboard["stock_total"] > 0, np.inf, 0)
)

dashboard["lead_time_dias"] = LEAD_TIME_DEFAULT
dashboard["cobertura_objetivo_dias"] = COBERTURA_OBJETIVO_DIAS
dashboard["objetivo_stock_45_dias"] = dashboard["venta_diaria_promedio_3m"] * COBERTURA_OBJETIVO_DIAS
dashboard["sugerencia_compra"] = np.maximum(
    np.ceil(dashboard["objetivo_stock_45_dias"] - dashboard["stock_total"]),
    0
)

dashboard["dias_para_comprar"] = np.where(
    np.isfinite(dashboard["dias_inventario"]),
    dashboard["dias_inventario"] - dashboard["lead_time_dias"],
    np.inf
)



# Defaults de métricas internas de rotación rápida.
for c in ["cantidad_ultimo_lote", "ventas_ultimos_7d", "sell_through_ultimos_7d_vs_lote"]:
    if c not in dashboard.columns:
        dashboard[c] = 0
    dashboard[c] = to_num(dashboard[c])

if "recompra_inmediata_flag" not in dashboard.columns:
    dashboard["recompra_inmediata_flag"] = "NO"
dashboard["recompra_inmediata_flag"] = dashboard["recompra_inmediata_flag"].fillna("NO")

def clasificar_rotacion_30d(row):
    dias = row.get("dias_desde_ultimo_arribo", np.nan)
    ventas = float(row.get("ventas_3m_unidades", 0) or 0)
    sell = float(row.get("sell_through_ultimos_7d_vs_lote", 0) or 0)
    stock = float(row.get("stock_total", 0) or 0)

    if pd.isna(dias):
        return "Sin arribo"
    dias = float(dias)

    # Regla comercial fuerte, pero solo si la cobertura NO es suficiente.
    # Si tiene inventario suficiente, se respeta el semáforo/recomendación normal.
    dias_inv = row.get("dias_inventario", np.nan)
    lead = float(row.get("lead_time_dias", LEAD_TIME_DEFAULT) or LEAD_TIME_DEFAULT)
    dias_para_comprar = np.inf
    if not pd.isna(dias_inv) and dias_inv != np.inf:
        dias_para_comprar = float(dias_inv) - lead

    if sell >= 0.30 and dias_para_comprar <= 5:
        return "Recompra inmediata"

    # Liquidación fuerte se mantiene hasta 90 días.
    if dias >= 90 and stock > 0:
        return "Liquidación / pérdida"

    # Semáforo rápido de 30 días.
    if dias <= 14:
        if ventas > 0 or sell > 0:
            return "Verde — Producto sano"
        return "Verde — Ventana inicial"

    if dias <= 21:
        return "Amarillo — Producto en observación"

    if dias <= 30:
        return "Rojo — Sigue sin ventas"

    # Después de 30 días, pero antes de 90, se mantiene presión comercial.
    return "Rojo — Rotación lenta"


dashboard["estado_rotacion_30d"] = dashboard.apply(clasificar_rotacion_30d, axis=1)


def clasificar_riesgo_sobrestock(dias_inv):
    if pd.isna(dias_inv) or dias_inv == np.inf:
        return "Sin venta"
    if dias_inv > 90:
        return "Riesgo +90 días"
    if dias_inv > 60:
        return "Riesgo +60 días"
    if dias_inv > 30:
        return "Riesgo +30 días"
    return "Sano ≤30 días"


def clasificar_alerta_compra(row):
    stock = float(row.get("stock_total", 0) or 0)
    ventas_3m = float(row.get("ventas_3m_unidades", 0) or 0)
    dias_inv = row.get("dias_inventario", np.nan)
    lead = float(row.get("lead_time_dias", LEAD_TIME_DEFAULT) or LEAD_TIME_DEFAULT)

    if ventas_3m <= 0 and stock > 0:
        return "Inventario sin ventas"
    if ventas_3m <= 0 and stock <= 0:
        return "Sin stock y sin ventas"
    if stock <= 0 and ventas_3m > 0:
        return "Inventario acabado - Comprar ya"

    if pd.isna(dias_inv) or dias_inv == np.inf:
        return "Inventario sin ventas"

    dias_para_comprar = dias_inv - lead

    if dias_para_comprar <= 0:
        return "Inventario acabado - Comprar ya"
    if dias_para_comprar <= 5:
        return "Compra con urgencia"
    return "Inventario suficiente"


def prioridad(alerta, riesgo, estado_rotacion=""):
    alerta = str(alerta)
    riesgo = str(riesgo)
    estado_rotacion = str(estado_rotacion)

    if "Recompra inmediata" in estado_rotacion and "suficiente" not in alerta.lower():
        return "Muy alta"
    if "Comprar ya" in alerta:
        return "Alta"
    if "Liquidación" in estado_rotacion or "+90" in riesgo:
        return "Alta"
    if "Sigue sin ventas" in estado_rotacion or "Rotación lenta" in estado_rotacion:
        return "Media"
    if "Producto en observación" in estado_rotacion or "urgencia" in alerta:
        return "Media"
    return "Baja"


def recomendacion(row):
    alerta = str(row.get("alerta_compra", ""))
    riesgo = str(row.get("riesgo_sobrestock", ""))
    estado = str(row.get("estado_rotacion_30d", ""))
    sug = float(row.get("sugerencia_compra", 0) or 0)
    sell = float(row.get("sell_through_ultimos_7d_vs_lote", 0) or 0)
    ventas_7d = float(row.get("ventas_ultimos_7d", 0) or 0)
    lote = float(row.get("cantidad_ultimo_lote", 0) or 0)
    dias_arribo = row.get("dias_desde_ultimo_arribo", np.nan)
    stock = float(row.get("stock_total", 0) or 0)

    # 1) Recompra inmediata por sell-through reciente.
    # Solo se activa si NO hay inventario suficiente.
    # Si el stock alcanza, se mantiene la recomendación normal.
    if sell >= 0.30 and "suficiente" not in alerta.lower() and sug > 0:
        return (
            f"Recompra inmediata / vendió {sell*100:.0f}% "
            f"del último lote en los últimos 7 días ({ventas_7d:,.0f}/{lote:,.0f} pzs)"
        )

    # 2) Compra por cobertura insuficiente.
    if "Comprar ya" in alerta:
        return f"Comprar ya / sugerido {sug:,.0f} unidades"
    if "urgencia" in alerta:
        return f"Compra con urgencia / sugerido {sug:,.0f} unidades"

    # 3) Liquidación dura hasta 90 días.
    if not pd.isna(dias_arribo) and float(dias_arribo) >= 90 and stock > 0:
        return "Liquidar / mover agresivo / aceptar pérdida si aplica"

    # 4) Semáforo rápido de 30 días para acelerar venta.
    if "Producto sano" in estado or "Ventana inicial" in estado:
        return "Producto sano / mantener, optimizar ligeramente y vigilar crecimiento"

    if "Producto en observación" in estado:
        return "Producto en observación / ajustar precio, ads, listing y competencia; ROI objetivo 5%"

    if "Sigue sin ventas" in estado:
        return "Sigue sin ventas / mover canal y bajar ROI objetivo a 3%"

    if "Rotación lenta" in estado:
        return "Rotación lenta / mover canal, ajustar agresivo y revisar conversión"

    # 5) Respaldo de sobrestock tradicional.
    if "+90" in riesgo:
        return "Revisar exceso de inventario / posible liquidación"
    if "+60" in riesgo:
        return "Revisar velocidad de venta y compras abiertas"
    if "+30" in riesgo:
        return "Monitorear inventario"
    if "sin ventas" in alerta.lower():
        return "Validar publicación, mapeo o producto lento"
    return "Inventario sano"



dashboard["riesgo_sobrestock"] = dashboard["dias_inventario"].apply(clasificar_riesgo_sobrestock)
dashboard["alerta_compra"] = dashboard.apply(clasificar_alerta_compra, axis=1)
dashboard["prioridad"] = dashboard.apply(lambda r: prioridad(r["alerta_compra"], r["riesgo_sobrestock"], r.get("estado_rotacion_30d", "")), axis=1)
dashboard["recomendacion"] = dashboard.apply(recomendacion, axis=1)
dashboard["recompra_inmediata_efectiva"] = np.where(
    dashboard["recomendacion"].astype(str).str.contains("Recompra inmediata", case=False, na=False),
    "SI",
    "NO"
)

# Pareto y ranking por MONTO vendido 3M.
# Importante:
# - El Top 80% ya no se calcula por piezas.
# - Ahora se calcula por ventas_3m_monto ($).
dashboard = dashboard.sort_values("ventas_3m_monto", ascending=False).reset_index(drop=True)
dashboard["ranking_ventas_3m"] = np.arange(1, len(dashboard) + 1)
dashboard["ranking_monto_3m"] = dashboard["ranking_ventas_3m"]

total_ventas_3m = dashboard["ventas_3m_monto"].sum()
if total_ventas_3m > 0:
    dashboard["participacion_ventas_3m"] = dashboard["ventas_3m_monto"] / total_ventas_3m
    dashboard["participacion_acumulada_3m"] = dashboard["participacion_ventas_3m"].cumsum()
else:
    dashboard["participacion_ventas_3m"] = 0
    dashboard["participacion_acumulada_3m"] = 0

# Incluye los productos hasta 80% y también el producto que cruza el umbral.
dashboard["top_80_flag"] = "NO"
if len(dashboard) > 0 and total_ventas_3m > 0:
    mask_top = dashboard["participacion_acumulada_3m"] <= 0.80

    if mask_top.any():
        pos_ultimo = int(np.where(mask_top.values)[0][-1])
        dashboard.loc[:pos_ultimo, "top_80_flag"] = "SI"

        if pos_ultimo + 1 < len(dashboard):
            dashboard.loc[pos_ultimo + 1, "top_80_flag"] = "SI"
    else:
        dashboard.loc[0, "top_80_flag"] = "SI"

dashboard["ranking_top80"] = np.where(
    dashboard["top_80_flag"] == "SI",
    "#" + dashboard["ranking_monto_3m"].astype(str),
    "No"
)

dashboard["dias_inventario_num"] = dashboard["dias_inventario"].replace(np.inf, np.nan)
dashboard["dias_inventario_texto"] = dashboard["dias_inventario"].apply(
    lambda x: "Sin venta" if x == np.inf or pd.isna(x) else round(float(x), 1)
)

# ============================================================
# ANTIGÜEDAD DE INVENTARIO / DÍAS DESDE ÚLTIMO ARRIBO
# ============================================================
# Esta métrica NO indica cuánto va a durar el stock.
# Indica cuántos días han pasado desde el último arribo/compra registrada en Odoo.
if "ultima_fecha_compra_odoo" in dashboard.columns:
    dashboard["ultima_fecha_compra_odoo_dt"] = pd.to_datetime(
        dashboard["ultima_fecha_compra_odoo"],
        errors="coerce"
    )
else:
    dashboard["ultima_fecha_compra_odoo_dt"] = pd.NaT

fecha_referencia_antiguedad = pd.Timestamp.today().normalize()
dashboard["dias_desde_ultimo_arribo"] = (
    fecha_referencia_antiguedad - dashboard["ultima_fecha_compra_odoo_dt"]
).dt.days

dashboard["dias_desde_ultimo_arribo"] = dashboard["dias_desde_ultimo_arribo"].where(
    dashboard["ultima_fecha_compra_odoo_dt"].notna(),
    np.nan
)

dashboard["dias_desde_ultimo_arribo_texto"] = dashboard["dias_desde_ultimo_arribo"].apply(
    lambda x: "Sin arribo" if pd.isna(x) else int(max(float(x), 0))
)

def clasificar_antiguedad_inventario(dias):
    if pd.isna(dias):
        return "Sin arribo"
    dias = float(dias)
    if dias > 90:
        return "Alerta +90 días"
    if dias > 60:
        return "Alerta +60 días"
    if dias > 30:
        return "Alerta +30 días"
    return "Reciente ≤30 días"

dashboard["alerta_antiguedad_inventario"] = dashboard["dias_desde_ultimo_arribo"].apply(
    clasificar_antiguedad_inventario
)

dashboard["sku_chips_json"] = dashboard["sku_chips_json"].fillna("[]")
dashboard["skus_sincronizados"] = dashboard["skus_sincronizados"].fillna("")

# ============================================================
# COSTO UNITARIO ODOO (SOLO UBICACIONES CUATI)
# ============================================================
# inventario_canal_operativo trae, por sku_madre + canal, el costo total y
# las unidades físicas acumuladas de todas las ubicaciones CUATI (ver
# construir_inventario_canal_sku_madre en el script 01). Es el mismo costo
# para todos los canales de un mismo sku_madre, así que basta con tomar un
# valor por sku_madre (max evita duplicar si algún canal quedó en NaN/0).
if not inventario_canal_operativo.empty and "costo_unitario_cuati" in inventario_canal_operativo.columns:
    costo_cuati_sku = (
        inventario_canal_operativo
        .assign(
            costo_unitario_cuati=to_num(inventario_canal_operativo.get("costo_unitario_cuati", 0)),
            costo_total_cuati=to_num(inventario_canal_operativo.get("costo_total_cuati", 0)),
            unidades_fisicas_cuati=to_num(inventario_canal_operativo.get("unidades_fisicas_cuati", 0)),
        )
        .groupby("sku_madre", as_index=False)
        .agg(
            costo_unitario_odoo_cuati=("costo_unitario_cuati", "max"),
            costo_total_odoo_cuati=("costo_total_cuati", "max"),
            unidades_fisicas_odoo_cuati=("unidades_fisicas_cuati", "max"),
        )
    )
else:
    costo_cuati_sku = pd.DataFrame(columns=[
        "sku_madre", "costo_unitario_odoo_cuati", "costo_total_odoo_cuati", "unidades_fisicas_odoo_cuati"
    ])

dashboard = dashboard.merge(costo_cuati_sku, on="sku_madre", how="left")
for c in ["costo_unitario_odoo_cuati", "costo_total_odoo_cuati", "unidades_fisicas_odoo_cuati"]:
    dashboard[c] = to_num(dashboard.get(c, 0)).fillna(0)

# ============================================================
# CATEGORÍA Y MARCA POR IQ (para el motor de repartición)
# ============================================================
def _primero_texto_no_vacio(serie):
    for x in serie:
        s = str(x or "").strip()
        if s and s.lower() not in {"nan", "none", "false"}:
            return s
    return ""


def construir_metadata_producto_reparticion(inventario_canal, ventas_periodo):
    frames = []
    if inventario_canal is not None and not inventario_canal.empty and "sku_madre" in inventario_canal.columns:
        tmp = inventario_canal.copy()
        tmp["categoria"] = tmp.get("categoria", tmp.get("categoria_odoo", ""))
        tmp["marca"] = tmp.get("marca", tmp.get("marca_odoo", ""))
        frames.append(tmp[["sku_madre", "categoria", "marca"]])
    if ventas_periodo is not None and not ventas_periodo.empty and "sku_madre" in ventas_periodo.columns:
        tmp = ventas_periodo.copy()
        if "categoria" not in tmp.columns:
            tmp["categoria"] = ""
        if "marca" not in tmp.columns:
            tmp["marca"] = ""
        frames.append(tmp[["sku_madre", "categoria", "marca"]])
    if not frames:
        return pd.DataFrame(columns=["sku_madre", "categoria", "marca"])
    meta = pd.concat(frames, ignore_index=True)
    meta["sku_madre"] = meta["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    meta = meta[meta["sku_madre"].ne("")].copy()
    for c in ["categoria", "marca"]:
        meta[c] = meta[c].fillna("").astype(str).str.strip()
    return meta.groupby("sku_madre", as_index=False).agg(
        categoria=("categoria", _primero_texto_no_vacio),
        marca=("marca", _primero_texto_no_vacio),
    )


metadata_reparticion = construir_metadata_producto_reparticion(inventario_canal_operativo, ventas_3m)
if not metadata_reparticion.empty:
    dashboard = dashboard.merge(metadata_reparticion, on="sku_madre", how="left", suffixes=("", "_meta"))
    for _c in ["categoria", "marca"]:
        _meta = f"{_c}_meta"
        if _c not in dashboard.columns:
            dashboard[_c] = ""
        if _meta in dashboard.columns:
            actual = dashboard[_c].fillna("").astype(str).str.strip()
            respaldo = dashboard[_meta].fillna("").astype(str).str.strip()
            dashboard[_c] = actual.where(actual.ne(""), respaldo)
            dashboard = dashboard.drop(columns=[_meta])
for _c in ["categoria", "marca"]:
    if _c not in dashboard.columns:
        dashboard[_c] = ""
    dashboard[_c] = dashboard[_c].fillna("").astype(str).str.strip()

# Columnas ordenadas.
cols_dashboard = [
    "sku_madre", "producto_madre", "categoria", "marca", "proveedor_principal", "proveedores_odoo_compras",
    "ordenes_compra_odoo", "primera_fecha_compra_odoo", "ultima_fecha_compra_odoo",
    "cantidad_arribada_historica",
    "dias_desde_ultimo_arribo", "dias_desde_ultimo_arribo_texto", "alerta_antiguedad_inventario",
    "estado_rotacion_30d", "cantidad_ultimo_lote", "ventas_ultimos_7d",
    "sell_through_ultimos_7d_vs_lote", "recompra_inmediata_flag", "recompra_inmediata_efectiva",
    "imagen_odoo_base64", "tiene_imagen_odoo", "fuente_imagen_excel", "producto_imagen_odoo", "ranking_ventas_3m", "ranking_monto_3m", "ranking_top80",
    "skus_sincronizados", "sku_chips_json", "sku_por_canal_json", "num_skus_sincronizados",
    "skus_iq", "skus_odoo_interno", "skus_amazon", "skus_mercado_libre",
    "skus_walmart", "skus_liverpool", "skus_coppel", "skus_elektra",
    "skus_tiktok", "skus_otro",
    "num_skus_iq", "num_skus_odoo_interno", "num_skus_amazon", "num_skus_mercado_libre",
    "num_skus_walmart", "num_skus_liverpool", "num_skus_coppel", "num_skus_elektra",
    "num_skus_tiktok", "num_skus_otro",
    "ventas_3m_unidades", "ventas_3m_monto", "pedidos_3m", "dias_con_venta_3m",
    "venta_diaria_promedio_positiva_3m",
    "dias_con_stock_90d", "dias_sin_stock_90d", "venta_diaria_stock_90d",
    "ritmo_congelado_stockout", "fecha_ultimo_dia_con_stock", "dias_desde_ultimo_stock",
    "metodo_ritmo_stock", "fuente_arribos_ritmo",
    "venta_diaria_promedio_3m",
    "stock_total", "stock_total_congelado",
    "arribos_post_corte_unidades", "arribos_post_corte_monto",
    "ordenes_compra_arribos", "proveedores_arribos",
    "primera_fecha_arribo", "ultima_fecha_arribo",
    "ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte",
    "stock_total_antes_movimientos", "stock_total_antes_descuento",
    "stock_calculado_desde_congelado", "fecha_corte_stock", "formula_stock",
    "stock_odoo_cuautitlan", "stock_amazon_fba", "stock_meli_full",
    "stock_walmart_wfs", "stock_liverpool_99min",
    "costo_unitario_odoo_cuati", "costo_total_odoo_cuati", "unidades_fisicas_odoo_cuati",
    "dias_inventario_num", "dias_inventario_texto",
    "lead_time_dias", "cobertura_objetivo_dias", "objetivo_stock_45_dias",
    "sugerencia_compra", "dias_para_comprar",
    "alerta_compra", "riesgo_sobrestock", "prioridad", "recomendacion",
    "top_80_flag", "participacion_ventas_3m", "participacion_acumulada_3m",
    "fuentes_venta", "canales_venta",
]
for c in cols_dashboard:
    if c not in dashboard.columns:
        dashboard[c] = ""

dashboard_productos = dashboard[cols_dashboard].copy()

# ============================================================
# ROTACIÓN INDIVIDUAL POR CANAL
# ============================================================

def construir_rotacion_por_canal(inventario_canal, ventas_detalle, fecha_fin, dias=90, catalogo_productos=None):
    columnas = [
        "sku_madre", "producto_madre", "canal", "inventario_odoo",
        "inventario_transito", "inventario_full", "inventario_total",
        "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades",
        "ventas_full_90d", "ventas_drop_90d", "ventas_pendiente_odoo_90d",
        "dias_con_venta_10d", "dias_con_venta_30d", "dias_con_venta_90d", "semanas_con_venta_90d",
        "venta_diaria_10d", "venta_diaria_30d", "venta_diaria_90d",
        "demanda_diaria_ponderada", "aceleracion_10_vs_90",
        "venta_diaria_calendario", "venta_diaria_full", "demanda_tipo",
        "cobertura_full_dias", "cobertura_total_dias",
        "canal_tiene_full", "objetivo_full_dias", "dias_transito_full",
        "consumo_estimado_transito", "full_proyectado_arribo",
        "objetivo_full_piezas", "cobertura_proyectada_arribo_dias",
        "transferencia_sugerida", "check_manual_full", "criterio_transferencia_full",
        "accion_preliminar", "justificacion", "validacion_total_componentes",
        # Antigüedad de inventario por canal (rastreada por Lote en Odoo) y
        # ritmo de venta del canal, para la sección "Antigüedad de inventario".
        "dias_inventario_lote", "unidades_con_lote", "unidades_sin_lote_odoo",
        "cobertura_lote_pct", "fecha_recepcion_lote_min", "fecha_recepcion_lote_max",
        "dias_inventario_canal", "fuente_dias_inventario", "alerta_antiguedad_canal",
        "ritmo_venta_canal",
        # Versión 2026-09: capas FIFO, criterio de alerta y valuación.
        "dias_inventario_max", "unidades_mas_90d", "unidades_mas_60d", "unidades_mas_30d",
        "detalle_fuente_dias", "fecha_ultimo_envio_full", "fecha_envio_full_vigente_mas_antiguo",
        "dias_desde_compra_sku", "dias_alerta_antiguedad", "criterio_alerta_antiguedad",
        "costo_unitario_inventario", "fuente_costo_inventario",
        "valor_odoo", "valor_transito", "valor_full", "valor_inventario_canal", "valor_mas_90d",
    ]
    if inventario_canal is None or inventario_canal.empty:
        return pd.DataFrame(columns=columnas)

    inv = inventario_canal.copy()
    for c in ["inventario_odoo", "inventario_transito", "inventario_full", "inventario_total"]:
        inv[c] = to_num(inv.get(c, 0)).clip(lower=0)

    # Mapa sku_madre -> producto_madre construido desde todas las filas que sí
    # traen el nombre resuelto por el diccionario origen4. Se usa más abajo
    # para rellenar los huecos que deja el merge outer con ventas, ya que
    # ventas_detalle no trae producto_madre y por lo tanto cualquier SKU-canal
    # que exista solo del lado de ventas quedaría sin nombre de producto.
    mapa_producto_madre = (
        inv[inv["producto_madre"].astype(str).str.strip().ne("")]
        .drop_duplicates("sku_madre")
        .set_index("sku_madre")["producto_madre"]
        .to_dict()
        if "producto_madre" in inv.columns else {}
    )

    # Respaldo adicional: el catálogo maestro (dashboard_productos) ya trae
    # producto_madre resuelto para todo el universo de SKU, incluyendo casos
    # que el diccionario origen4 no logró resolver a nivel de inventario por
    # canal. Se usa para completar cualquier hueco que el mapa anterior no
    # haya cubierto, sin sobreescribir nombres que sí vinieron de inv.
    if catalogo_productos is not None and not catalogo_productos.empty \
            and {"sku_madre", "producto_madre"}.issubset(catalogo_productos.columns):
        mapa_catalogo = (
            catalogo_productos[catalogo_productos["producto_madre"].astype(str).str.strip().ne("")]
            .drop_duplicates("sku_madre")
            .set_index("sku_madre")["producto_madre"]
            .to_dict()
        )
        for sku, nombre in mapa_catalogo.items():
            mapa_producto_madre.setdefault(sku, nombre)

    ven = ventas_detalle.copy() if ventas_detalle is not None else pd.DataFrame()
    if ven.empty:
        resumen_ventas = pd.DataFrame(columns=["sku_madre", "canal"])
    else:
        ven["fecha"] = pd.to_datetime(ven.get("fecha", pd.NaT), errors="coerce")
        ven["cantidad"] = to_num(ven.get("cantidad", 0))
        if "canal_venta" in ven.columns:
            canal_serie = ven["canal_venta"]
        elif "canal" in ven.columns:
            canal_serie = ven["canal"]
        else:
            canal_serie = pd.Series("Sin identificar", index=ven.index)
        ven["canal_venta"] = canal_serie.apply(normalizar_canal_dashboard)
        if "modalidad_venta" in ven.columns:
            modalidad_serie = ven["modalidad_venta"]
        else:
            modalidad_serie = pd.Series("SIN_IDENTIFICAR", index=ven.index)
        ven["modalidad_venta"] = modalidad_serie.astype(str).str.upper()
        inicio = pd.to_datetime(fecha_fin).normalize() - pd.Timedelta(days=dias - 1)
        ven = ven[
            ven.get("tiene_referencia_madre", "").astype(str).str.upper().eq("SI")
            & ven["fecha"].notna()
            & (ven["fecha"] >= inicio)
            & (ven["fecha"] <= pd.to_datetime(fecha_fin) + pd.Timedelta(days=1))
        ].copy()
        ven["canal"] = ven["canal_venta"]
        ven["dia"] = ven["fecha"].dt.date
        ven["semana"] = ven["fecha"].dt.to_period("W").astype(str)
        fin_ventas = pd.to_datetime(fecha_fin).normalize() + pd.Timedelta(days=1)
        inicio_10d = pd.to_datetime(fecha_fin).normalize() - pd.Timedelta(days=9)
        inicio_30d = pd.to_datetime(fecha_fin).normalize() - pd.Timedelta(days=29)
        ven["u_10d"] = np.where(ven["fecha"] >= inicio_10d, ven["cantidad"], 0)
        ven["u_30d"] = np.where(ven["fecha"] >= inicio_30d, ven["cantidad"], 0)
        ven["dia_10d"] = np.where(ven["fecha"] >= inicio_10d, ven["dia"].astype(str), "")
        ven["dia_30d"] = np.where(ven["fecha"] >= inicio_30d, ven["dia"].astype(str), "")
        ven["u_full"] = np.where(ven["modalidad_venta"].eq("FULL"), ven["cantidad"], 0)
        ven["u_drop"] = np.where(ven["modalidad_venta"].eq("DROP"), ven["cantidad"], 0)
        ven["u_pend"] = np.where(ven["modalidad_venta"].eq("PENDIENTE_ODOO"), ven["cantidad"], 0)
        resumen_ventas = ven.groupby(["sku_madre", "canal"], as_index=False).agg(
            ventas_10d_unidades=("u_10d", "sum"),
            ventas_30d_unidades=("u_30d", "sum"),
            ventas_90d_unidades=("cantidad", "sum"),
            ventas_full_90d=("u_full", "sum"),
            ventas_drop_90d=("u_drop", "sum"),
            ventas_pendiente_odoo_90d=("u_pend", "sum"),
            dias_con_venta_10d=("dia_10d", lambda s: len({x for x in s if str(x).strip()})),
            dias_con_venta_30d=("dia_30d", lambda s: len({x for x in s if str(x).strip()})),
            dias_con_venta_90d=("dia", "nunique"),
            semanas_con_venta_90d=("semana", "nunique"),
        )

    out = inv.merge(resumen_ventas, on=["sku_madre", "canal"], how="outer")

    # El merge outer puede crear filas sku_madre/canal que solo existían del
    # lado de ventas (por ejemplo, se vendió en un canal sin tener aún fila
    # de inventario ahí), o que sí venían de inventario pero sin nombre
    # resuelto por el diccionario origen4. En ambos casos producto_madre
    # queda vacío/NaN. Se rellena aquí usando mapa_producto_madre, que
    # combina lo resuelto en inv y, como respaldo, el catálogo maestro
    # (dashboard_productos) para cubrir el mayor número posible de SKU.
    if "producto_madre" not in out.columns:
        out["producto_madre"] = ""
    out["producto_madre"] = out["producto_madre"].fillna("").astype(str).str.strip()
    faltantes = out["producto_madre"].eq("")
    if faltantes.any() and mapa_producto_madre:
        out.loc[faltantes, "producto_madre"] = out.loc[faltantes, "sku_madre"].map(mapa_producto_madre).fillna("")

    for c in [
        "inventario_odoo", "inventario_transito", "inventario_full", "inventario_total",
        "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades",
        "ventas_full_90d", "ventas_drop_90d", "ventas_pendiente_odoo_90d",
        "dias_con_venta_10d", "dias_con_venta_30d", "dias_con_venta_90d", "semanas_con_venta_90d"
    ]:
        out[c] = to_num(out.get(c, 0)).clip(lower=0)

    # Antigüedad por lote: NO se rellena con 0, porque NaN aquí significa
    # "sin dato de lote todavía", no "cero días en stock".
    for c in ["dias_inventario_lote", "unidades_con_lote", "unidades_sin_lote_odoo", "cobertura_lote_pct"]:
        if c not in out.columns:
            out[c] = np.nan
        out[c] = pd.to_numeric(out[c], errors="coerce")
    for c in ["fecha_recepcion_lote_min", "fecha_recepcion_lote_max"]:
        if c not in out.columns:
            out[c] = pd.NaT
        out[c] = pd.to_datetime(out[c], errors="coerce")
    if "fuente_dias_inventario" not in out.columns:
        out["fuente_dias_inventario"] = ""
    out["fuente_dias_inventario"] = out["fuente_dias_inventario"].fillna("").astype(str)

    out["venta_diaria_10d"] = out["ventas_10d_unidades"] / 10.0
    out["venta_diaria_30d"] = out["ventas_30d_unidades"] / 30.0
    out["venta_diaria_90d"] = out["ventas_90d_unidades"] / float(dias)
    out["demanda_diaria_ponderada"] = (
        out["venta_diaria_10d"] * PESO_DEMANDA_10D
        + out["venta_diaria_30d"] * PESO_DEMANDA_30D
        + out["venta_diaria_90d"] * PESO_DEMANDA_90D
    )
    out["aceleracion_10_vs_90"] = np.where(
        out["venta_diaria_90d"] > 0,
        out["venta_diaria_10d"] / out["venta_diaria_90d"],
        np.where(out["venta_diaria_10d"] > 0, np.inf, 0),
    )
    out["venta_diaria_calendario"] = out["venta_diaria_90d"]
    out["venta_diaria_full"] = out["ventas_full_90d"] / float(dias)

    # "Ritmo de ventas" para la sección de antigüedad de inventario:
    # promedio de ventas diarias del canal en los últimos `dias` (90 por
    # defecto), calendario completo (no solo días con venta).
    out["ritmo_venta_canal"] = out["venta_diaria_calendario"]
    out["demanda_tipo"] = np.where(
        (out["ventas_90d_unidades"] <= 12) | (out["dias_con_venta_90d"] <= 8),
        "INTERMITENTE", "CONTINUA"
    )
    out["cobertura_full_dias"] = np.where(
        out["venta_diaria_full"] > 0,
        out["inventario_full"] / out["venta_diaria_full"],
        np.nan
    )
    out["cobertura_total_dias"] = np.where(
        out["venta_diaria_calendario"] > 0,
        out["inventario_total"] / out["venta_diaria_calendario"],
        np.nan
    )

    # ========================================================
    # SUGERENCIA ODOO -> FULL: 30 DÍAS DISPONIBLES AL ARRIBO
    # ========================================================
    # El traslado tarda aproximadamente DIAS_TRANSITO_ODOO_FULL días. Por eso
    # primero se proyecta cuánto stock Full quedará después de consumir durante
    # ese plazo y luego se completa hasta 30 días de cobertura al momento del
    # arribo. El tránsito ya registrado se considera disponible para esa fecha.
    # Para productos sin ventas históricas, el objetivo máximo es 10 piezas y
    # la recomendación queda marcada para validación manual.
    out["canal_tiene_full"] = out["canal"].isin(CANALES_CON_FULL)
    out["objetivo_full_dias"] = np.where(
        out["canal_tiene_full"], OBJETIVO_FULL_DIAS, 0
    )
    out["dias_transito_full"] = np.where(
        out["canal_tiene_full"], DIAS_TRANSITO_ODOO_FULL, 0
    )

    objetivo_historico = np.ceil(
        out["venta_diaria_calendario"] * float(OBJETIVO_FULL_DIAS)
    )
    objetivo_nuevo = np.full(len(out), MAX_PIEZAS_PRODUCTO_NUEVO_FULL, dtype=float)

    out["consumo_estimado_transito"] = np.where(
        out["canal_tiene_full"] & (out["ventas_90d_unidades"] > 0),
        np.ceil(out["venta_diaria_calendario"] * float(DIAS_TRANSITO_ODOO_FULL)),
        0,
    )
    out["full_proyectado_arribo"] = np.where(
        out["canal_tiene_full"],
        (out["inventario_full"] - out["consumo_estimado_transito"]).clip(lower=0),
        out["inventario_full"],
    )

    out["objetivo_full_piezas"] = np.where(
        ~out["canal_tiene_full"],
        0,
        np.where(
            out["ventas_90d_unidades"] > 0,
            objetivo_historico,
            objetivo_nuevo,
        ),
    )

    faltante_full_al_arribo = (
        out["objetivo_full_piezas"]
        - out["full_proyectado_arribo"]
        - out["inventario_transito"]
    ).clip(lower=0)

    out["transferencia_sugerida"] = np.minimum(
        np.ceil(faltante_full_al_arribo),
        out["inventario_odoo"],
    ).clip(lower=0)

    out["cobertura_proyectada_arribo_dias"] = np.where(
        out["venta_diaria_calendario"] > 0,
        (
            out["full_proyectado_arribo"]
            + out["inventario_transito"]
            + out["transferencia_sugerida"]
        ) / out["venta_diaria_calendario"],
        np.nan,
    )

    out["check_manual_full"] = np.where(
        out["canal_tiene_full"]
        & (out["ventas_90d_unidades"] <= 0)
        & (out["inventario_odoo"] > 0),
        "SI - PRODUCTO NUEVO",
        "NO",
    )

    def criterio_full(r):
        if not bool(r["canal_tiene_full"]):
            return "Canal sin operación Full configurada; no se sugiere traslado."
        if r["inventario_odoo"] <= 0:
            return "Sin stock disponible en Odoo para enviar."
        if r["ventas_90d_unidades"] <= 0:
            return (
                f"Sin ventas históricas del canal: objetivo máximo "
                f"{MAX_PIEZAS_PRODUCTO_NUEVO_FULL} piezas al arribo. "
                "No se estima consumo durante el tránsito y el CHECK MANUAL es obligatorio."
            )
        if r["transferencia_sugerida"] > 0:
            return (
                f"Se proyectan {DIAS_TRANSITO_ODOO_FULL} días de consumo antes del arribo "
                f"y después se completa a {OBJETIVO_FULL_DIAS} días de cobertura. "
                f"Ventas históricas: {r['ventas_90d_unidades']:.0f} unidades en {dias} días."
            )
        return (
            f"El Full proyectado al arribo + tránsito ya cubre el objetivo de "
            f"{OBJETIVO_FULL_DIAS} días, o no hay piezas adicionales disponibles en Odoo."
        )

    out["criterio_transferencia_full"] = out.apply(criterio_full, axis=1)

    def accion(r):
        if r["check_manual_full"].startswith("SI") and r["transferencia_sugerida"] > 0:
            return "CHECK MANUAL - PRODUCTO NUEVO A FULL"
        if r["transferencia_sugerida"] > 0:
            return f"TRANSFERIR A FULL - {OBJETIVO_FULL_DIAS} DÍAS AL ARRIBO"
        if r["demanda_tipo"] == "INTERMITENTE":
            if r["inventario_total"] <= 0 and r["ventas_90d_unidades"] > 0:
                return "REVISAR STOCKOUT - DEMANDA INTERMITENTE"
            return "REVISIÓN MANUAL - DEMANDA INTERMITENTE"
        if r["venta_diaria_calendario"] > 0 and r["cobertura_total_dias"] < 30:
            return "EVALUAR COMPRA A PROVEEDOR"
        return "SIN ACCIÓN"

    out["accion_preliminar"] = out.apply(accion, axis=1)
    out["justificacion"] = out.apply(
        lambda r: (
            f"Full={r['inventario_full']:.0f}; tránsito={r['inventario_transito']:.0f}; "
            f"Odoo={r['inventario_odoo']:.0f}; total={r['inventario_total']:.0f}; "
            f"ventas 90d={r['ventas_90d_unidades']:.0f}; "
            f"consumo tránsito {DIAS_TRANSITO_ODOO_FULL}d={r['consumo_estimado_transito']:.0f}; "
            f"Full proyectado al arribo={r['full_proyectado_arribo']:.0f}; "
            f"objetivo al arribo={r['objetivo_full_piezas']:.0f}; "
            f"sugerido enviar={r['transferencia_sugerida']:.0f}; "
            f"cobertura proyectada al arribo={r['cobertura_proyectada_arribo_dias']:.1f} días."
        ), axis=1
    )
    out["validacion_total_componentes"] = out["inventario_total"] - (
        out["inventario_odoo"] + out["inventario_transito"] + out["inventario_full"]
    )

    # ========================================================
    # DÍAS DE INVENTARIO POR CANAL (antigüedad real, no cobertura)
    # ========================================================
    # 1) Si hay lote (fecha real de recepción en Compras, rastreada por
    #    Odoo -> Inventario -> Reporte -> Trazabilidad), se usa ese dato,
    #    ponderado por las piezas vigentes de cada lote en este canal.
    # 2) Si no hay lote (producto sin trazabilidad activada en Odoo, o
    #    stock que ya salió de Odoo hacia Full/tránsito y por eso perdió
    #    el rastro de lote en esta base), se usa como respaldo la fecha
    #    del último arribo de compra del SKU completo (mismo valor para
    #    todos los canales de ese SKU), marcando la fuente como
    #    aproximada para que el equipo sepa qué tan confiable es el dato.
    mapa_dias_fallback = {}
    if catalogo_productos is not None and not catalogo_productos.empty \
            and {"sku_madre", "dias_desde_ultimo_arribo"}.issubset(catalogo_productos.columns):
        mapa_dias_fallback = (
            catalogo_productos.dropna(subset=["dias_desde_ultimo_arribo"])
            .drop_duplicates("sku_madre")
            .set_index("sku_madre")["dias_desde_ultimo_arribo"]
            .to_dict()
        )

    # Versión 2026-09: si el 01 ya trae la antigüedad FIFO por canal
    # (lote + envíos a Full + respaldo de compras), esa es la fuente de verdad.
    for c in ["dias_inventario_canal", "dias_inventario_max", "unidades_mas_90d",
              "unidades_mas_60d", "unidades_mas_30d", "dias_desde_compra_sku"]:
        if c not in out.columns:
            out[c] = np.nan
        out[c] = pd.to_numeric(out[c], errors="coerce")
    for c in ["fecha_ultimo_envio_full", "fecha_envio_full_vigente_mas_antiguo"]:
        if c not in out.columns:
            out[c] = pd.NaT
        out[c] = pd.to_datetime(out[c], errors="coerce")
    if "detalle_fuente_dias" not in out.columns:
        out["detalle_fuente_dias"] = ""
    out["detalle_fuente_dias"] = out["detalle_fuente_dias"].fillna("").astype(str)

    usa_fifo = out["dias_inventario_max"].notna().any()
    if not usa_fifo:
        # Base del 01 anterior: se conserva la lógica previa (lote CUATI o
        # último arribo del SKU).
        out["dias_inventario_canal"] = out["dias_inventario_lote"]
        out["dias_inventario_max"] = out["dias_inventario_lote"]
    sin_dato = out["dias_inventario_canal"].isna()
    if sin_dato.any() and mapa_dias_fallback:
        out.loc[sin_dato, "dias_inventario_canal"] = out.loc[sin_dato, "sku_madre"].map(mapa_dias_fallback)
        out.loc[sin_dato, "dias_inventario_max"] = out.loc[sin_dato, "dias_inventario_canal"]

    out.loc[sin_dato & out["dias_inventario_canal"].notna(), "fuente_dias_inventario"] = "APROX_ULTIMO_ARRIBO"
    out.loc[sin_dato & out["dias_inventario_canal"].isna(), "fuente_dias_inventario"] = "SIN_DATO"
    out.loc[~sin_dato & out["fuente_dias_inventario"].isin(["", "nan"]), "fuente_dias_inventario"] = "LOTE"
    if not usa_fifo:
        # Sin capas no se sabe cuántas piezas pasan el umbral: se aproxima
        # con todo el stock del canal cuando el promedio ya lo rebasa.
        out["unidades_mas_90d"] = np.where(out["dias_inventario_canal"] > 90, out["inventario_total"], 0.0)

    criterio = "capa" if (CRITERIO_ALERTA_ANTIGUEDAD == "capa" and usa_fifo) else "promedio"
    out["criterio_alerta_antiguedad"] = (
        "Capa más antigua del canal" if criterio == "capa" else "Promedio ponderado del canal"
    )
    out["dias_alerta_antiguedad"] = out["dias_inventario_max"] if criterio == "capa" else out["dias_inventario_canal"]

    def clasificar_antiguedad_canal(dias):
        if pd.isna(dias):
            return "Sin dato"
        dias = float(dias)
        if dias > 90:
            return "Alerta +90 días"
        if dias > 60:
            return "Alerta +60 días"
        if dias > 30:
            return "Alerta +30 días"
        return "Reciente ≤30 días"

    out["alerta_antiguedad_canal"] = out["dias_alerta_antiguedad"].apply(clasificar_antiguedad_canal)
    out.loc[out["inventario_total"] <= 0, "alerta_antiguedad_canal"] = "Sin stock"
    for c in ["unidades_mas_90d", "unidades_mas_60d", "unidades_mas_30d"]:
        out[c] = out[c].fillna(0).clip(lower=0)
    out.loc[out["inventario_total"] <= 0, ["unidades_mas_90d", "unidades_mas_60d", "unidades_mas_30d"]] = 0

    # ========================================================
    # COSTO Y VALOR (INVERSIÓN) POR SKU-CANAL
    # ========================================================
    # 01 calcula un costo por SKU madre (valor de quants CUATI, avg_cost,
    # standard_price o última recepción). Si el 01 es anterior, se usa el
    # costo CUATI. Las filas que solo existen del lado de ventas heredan el
    # costo del SKU para que ninguna fila con stock quede valuada en cero.
    if "costo_unitario_inventario" not in out.columns:
        out["costo_unitario_inventario"] = out.get("costo_unitario_cuati", np.nan)
    out["costo_unitario_inventario"] = pd.to_numeric(out["costo_unitario_inventario"], errors="coerce")
    if "fuente_costo_inventario" not in out.columns:
        out["fuente_costo_inventario"] = np.where(
            out["costo_unitario_inventario"].fillna(0) > 0, "ODOO_VALOR_QUANT", "SIN_COSTO"
        )
    mapa_costo = (
        out[out["costo_unitario_inventario"].fillna(0) > 0]
        .groupby("sku_madre")["costo_unitario_inventario"].max().to_dict()
    )
    mapa_fuente_costo = (
        out[out["costo_unitario_inventario"].fillna(0) > 0]
        .drop_duplicates("sku_madre").set_index("sku_madre")["fuente_costo_inventario"].to_dict()
    )
    sin_costo = out["costo_unitario_inventario"].fillna(0) <= 0
    out.loc[sin_costo, "costo_unitario_inventario"] = out.loc[sin_costo, "sku_madre"].map(mapa_costo)
    out.loc[sin_costo, "fuente_costo_inventario"] = out.loc[sin_costo, "sku_madre"].map(mapa_fuente_costo)
    out["costo_unitario_inventario"] = out["costo_unitario_inventario"].fillna(0)
    out["fuente_costo_inventario"] = out["fuente_costo_inventario"].fillna("SIN_COSTO").replace("", "SIN_COSTO")
    costo = out["costo_unitario_inventario"]
    out["valor_odoo"] = out["inventario_odoo"] * costo
    out["valor_transito"] = out["inventario_transito"] * costo
    out["valor_full"] = out["inventario_full"] * costo
    out["valor_inventario_canal"] = out["inventario_total"] * costo
    out["valor_mas_90d"] = out["unidades_mas_90d"] * costo

    return out[columnas].sort_values(["canal", "ventas_90d_unidades"], ascending=[True, False])


dashboard_rotacion_canal = construir_rotacion_por_canal(
    inventario_canal_operativo,
    ventas,
    fecha_fin,
    DIAS_ANALISIS_3M,
    catalogo_productos=dashboard_productos,
)

# Costo y valor total por SKU madre en el catálogo de productos (el 04 los
# usa como costo prioritario del modelo B2B).
if not dashboard_rotacion_canal.empty:
    _valor_sku = dashboard_rotacion_canal.groupby("sku_madre", as_index=False).agg(
        costo_unitario_inventario=("costo_unitario_inventario", "max"),
        valor_inventario_total=("valor_inventario_canal", "sum"),
        valor_inventario_mas_90d=("valor_mas_90d", "sum"),
    )
    _fuente_sku = (
        dashboard_rotacion_canal[dashboard_rotacion_canal["costo_unitario_inventario"] > 0]
        .drop_duplicates("sku_madre")[["sku_madre", "fuente_costo_inventario"]]
    )
    dashboard_productos = (
        dashboard_productos
        .drop(columns=[c for c in ["costo_unitario_inventario", "fuente_costo_inventario",
                                   "valor_inventario_total", "valor_inventario_mas_90d"]
                       if c in dashboard_productos.columns])
        .merge(_valor_sku, on="sku_madre", how="left")
        .merge(_fuente_sku, on="sku_madre", how="left")
    )
    for _c in ["costo_unitario_inventario", "valor_inventario_total", "valor_inventario_mas_90d"]:
        dashboard_productos[_c] = to_num(dashboard_productos[_c]).fillna(0)
    dashboard_productos["fuente_costo_inventario"] = dashboard_productos["fuente_costo_inventario"].fillna("SIN_COSTO")

# ============================================================
# SUGERENCIAS DE REDISTRIBUCIÓN INTERNA ENTRE CANALES
# ============================================================

def construir_sugerencias_redistribucion(rotacion_canal):
    columnas = [
        "sku_madre", "producto_madre", "canal_origen", "canal_destino",
        "cantidad_sugerida", "tipo_stock_origen", "stock_transferible_total",
        "stock_odoo_utilizado", "stock_full_utilizado",
        "stock_origen_actual", "stock_destino_actual", "stock_odoo_origen", "stock_full_origen",
        "ventas_10d_origen", "ventas_30d_origen", "ventas_90d_origen",
        "ventas_10d_destino", "ventas_30d_destino", "ventas_90d_destino",
        "demanda_ponderada_origen", "demanda_ponderada_destino",
        "cobertura_origen_antes", "cobertura_origen_despues",
        "cobertura_destino_antes", "cobertura_destino_despues",
        "stock_origen_proyectado_arribo", "stock_destino_proyectado_arribo",
        "objetivo_origen_piezas", "objetivo_destino_piezas",
        "dias_transito", "objetivo_dias", "confianza", "veredicto",
        "check_manual", "motivo", "beneficio_estimado",
    ]
    if rotacion_canal is None or rotacion_canal.empty:
        return pd.DataFrame(columns=columnas)

    base = rotacion_canal.copy()
    for c in [
        "inventario_odoo", "inventario_transito", "inventario_full", "inventario_total",
        "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades",
        "dias_con_venta_90d", "demanda_diaria_ponderada",
    ]:
        base[c] = to_num(base.get(c, 0)).clip(lower=0)

    # General representa la bolsa central y ya se atiende en Odoo → Full.
    # Aquí solo comparamos canales comerciales entre sí.
    base = base[~base["canal"].astype(str).str.strip().eq("General")].copy()

    sugerencias = []
    for sku_madre, grupo in base.groupby("sku_madre", dropna=False):
        sku_madre = str(sku_madre or "").strip()
        if not sku_madre or len(grupo) < 2:
            continue

        registros = []
        for _, r in grupo.iterrows():
            demanda = float(r.get("demanda_diaria_ponderada", 0) or 0)
            stock_total = float(r.get("inventario_total", 0) or 0)
            stock_odoo = float(r.get("inventario_odoo", 0) or 0)
            stock_full = float(r.get("inventario_full", 0) or 0)
            proyectado = max(stock_total - demanda * DIAS_TRANSITO_INTERNO, 0)
            objetivo = demanda * OBJETIVO_REDISTRIBUCION_DIAS if demanda > 0 else float(RESERVA_MINIMA_SIN_VENTAS)
            cobertura = proyectado / demanda if demanda > 0 else np.inf
            exceso = max(proyectado - objetivo, 0)
            capacidad_odoo = max(min(stock_odoo, exceso), 0)
            capacidad_full = max(min(stock_full, exceso - capacidad_odoo), 0)
            capacidad_total = capacidad_odoo + capacidad_full
            necesidad = max(objetivo - proyectado, 0) if demanda > 0 else 0
            registros.append({
                "row": r, "canal": str(r.get("canal", "")).strip(),
                "demanda": demanda, "stock_total": stock_total,
                "stock_odoo": stock_odoo, "stock_full": stock_full,
                "proyectado": proyectado, "objetivo": objetivo,
                "cobertura": cobertura, "exceso": exceso,
                "capacidad_odoo": capacidad_odoo, "capacidad_full": capacidad_full,
                "capacidad_total": capacidad_total, "necesidad": necesidad,
            })

        destinos = sorted(
            [x for x in registros if x["necesidad"] >= 1 and x["demanda"] > 0],
            key=lambda x: (x["cobertura"], -x["demanda"]),
        )
        cap_odoo = {x["canal"]: x["capacidad_odoo"] for x in registros}
        cap_full = {x["canal"]: x["capacidad_full"] for x in registros}

        for destino in destinos:
            necesidad_restante = float(np.ceil(destino["necesidad"]))
            fuentes = sorted(
                [x for x in registros if x["canal"] != destino["canal"] and (cap_odoo.get(x["canal"], 0) + cap_full.get(x["canal"], 0)) >= 1],
                key=lambda x: (-(cap_odoo.get(x["canal"], 0) + cap_full.get(x["canal"], 0)), x["demanda"]),
            )
            for origen in fuentes:
                disponible_odoo = cap_odoo.get(origen["canal"], 0)
                disponible_full = cap_full.get(origen["canal"], 0)
                disponible_total = disponible_odoo + disponible_full
                cantidad = int(max(min(np.floor(disponible_total), np.ceil(necesidad_restante)), 0))
                if cantidad <= 0:
                    continue
                odoo_usado = int(min(cantidad, np.floor(disponible_odoo)))
                full_usado = int(cantidad - odoo_usado)
                tipo_stock = "ODOO ASIGNADO" if full_usado <= 0 else ("MIXTO ODOO + FULL" if odoo_usado > 0 else "FULL - RETIRO / RELABELING")

                demanda_o, demanda_d = origen["demanda"], destino["demanda"]
                o_antes, d_antes = origen["cobertura"], destino["cobertura"]
                o_despues = max(origen["proyectado"] - cantidad, 0) / demanda_o if demanda_o > 0 else np.inf
                d_despues = (destino["proyectado"] + cantidad) / demanda_d
                ro, rd = origen["row"], destino["row"]
                dias_venta_o = float(ro.get("dias_con_venta_90d", 0) or 0)
                dias_venta_d = float(rd.get("dias_con_venta_90d", 0) or 0)
                check_manual = full_usado > 0 or demanda_o <= 0 or dias_venta_d < 3 or dias_venta_o < 2
                aceleracion_d = float(rd.get("aceleracion_10_vs_90", 0) or 0)
                if full_usado > 0:
                    confianza = "MEDIA" if dias_venta_d >= 3 else "BAJA"
                    veredicto = "APROBAR CON REVISIÓN OPERATIVA"
                elif not check_manual and dias_venta_d >= 8 and 0.45 <= aceleracion_d <= 3.0:
                    confianza, veredicto = "ALTA", "APROBAR"
                elif dias_venta_d >= 3:
                    confianza, veredicto = "MEDIA", "APROBAR CON REVISIÓN"
                else:
                    confianza, veredicto, check_manual = "BAJA", "REVISIÓN MANUAL", True

                operativa = (
                    " Requiere retiro/relabeling desde Full antes de confirmar." if full_usado > 0 else ""
                )
                motivo = (
                    f"{destino['canal']} proyecta {d_antes:.1f} días; {origen['canal']} puede ceder "
                    f"{cantidad} piezas y conservar {o_despues:.1f} días.{operativa}"
                )
                beneficio = (
                    f"La cobertura de {destino['canal']} subiría de {d_antes:.1f} a {d_despues:.1f} días; "
                    f"{origen['canal']} conservaría {o_despues:.1f} días."
                )
                sugerencias.append({
                    "sku_madre": sku_madre,
                    "producto_madre": first_non_empty(pd.Series([rd.get("producto_madre", ""), ro.get("producto_madre", "")])),
                    "canal_origen": origen["canal"], "canal_destino": destino["canal"],
                    "cantidad_sugerida": cantidad, "tipo_stock_origen": tipo_stock,
                    "stock_transferible_total": int(np.floor(disponible_total)),
                    "stock_odoo_utilizado": odoo_usado, "stock_full_utilizado": full_usado,
                    "stock_origen_actual": origen["stock_total"], "stock_destino_actual": destino["stock_total"],
                    "stock_odoo_origen": origen["stock_odoo"], "stock_full_origen": origen["stock_full"],
                    "ventas_10d_origen": float(ro.get("ventas_10d_unidades", 0) or 0),
                    "ventas_30d_origen": float(ro.get("ventas_30d_unidades", 0) or 0),
                    "ventas_90d_origen": float(ro.get("ventas_90d_unidades", 0) or 0),
                    "ventas_10d_destino": float(rd.get("ventas_10d_unidades", 0) or 0),
                    "ventas_30d_destino": float(rd.get("ventas_30d_unidades", 0) or 0),
                    "ventas_90d_destino": float(rd.get("ventas_90d_unidades", 0) or 0),
                    "demanda_ponderada_origen": demanda_o, "demanda_ponderada_destino": demanda_d,
                    "cobertura_origen_antes": o_antes if np.isfinite(o_antes) else 9999,
                    "cobertura_origen_despues": o_despues if np.isfinite(o_despues) else 9999,
                    "cobertura_destino_antes": d_antes, "cobertura_destino_despues": d_despues,
                    "stock_origen_proyectado_arribo": origen["proyectado"],
                    "stock_destino_proyectado_arribo": destino["proyectado"],
                    "objetivo_origen_piezas": origen["objetivo"], "objetivo_destino_piezas": destino["objetivo"],
                    "dias_transito": DIAS_TRANSITO_INTERNO, "objetivo_dias": OBJETIVO_REDISTRIBUCION_DIAS,
                    "confianza": confianza, "veredicto": veredicto,
                    "check_manual": "SI" if check_manual else "NO",
                    "motivo": motivo, "beneficio_estimado": beneficio,
                })
                cap_odoo[origen["canal"]] = max(disponible_odoo - odoo_usado, 0)
                cap_full[origen["canal"]] = max(disponible_full - full_usado, 0)
                necesidad_restante = max(necesidad_restante - cantidad, 0)
                if necesidad_restante < 1:
                    break

    out = pd.DataFrame(sugerencias, columns=columnas)
    if out.empty:
        return out
    orden_conf = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}
    out["_orden_conf"] = out["confianza"].map(orden_conf).fillna(9)
    out = out.sort_values(["_orden_conf", "cantidad_sugerida", "cobertura_destino_antes"], ascending=[True, False, True]).drop(columns=["_orden_conf"])
    return out.reset_index(drop=True)


redistribuciones_sugeridas = construir_sugerencias_redistribucion(dashboard_rotacion_canal)

# ============================================================
# REPARTICIÓN DE STOCK CENTRAL HACIA CANALES
# ============================================================
def _asignar_enteros_por_peso(total, pesos):
    """Reparte un total entero respetando pesos y garantizando suma exacta."""
    total = int(max(total, 0))
    arr = np.array([max(float(x or 0), 0.0) for x in pesos], dtype=float)
    if total <= 0:
        return [0] * len(arr)
    if len(arr) == 0:
        return []
    if arr.sum() <= 0:
        arr = np.ones(len(arr), dtype=float)
    raw = total * arr / arr.sum()
    base = np.floor(raw).astype(int)
    faltan = total - int(base.sum())
    if faltan > 0:
        frac = raw - base
        orden = np.argsort(-frac, kind="stable")
        for idx in orden[:faltan]:
            base[idx] += 1
    return base.tolist()


def _perfil_dict(hist, keys):
    """Devuelve señal por canal y estadísticas de muestra para un perfil."""
    if hist.empty:
        return {}, {}
    tmp = hist.copy()
    for k in keys:
        tmp[k] = tmp[k].fillna("").astype(str).str.strip()
    tmp = tmp[np.logical_and.reduce([tmp[k].ne("") for k in keys])].copy()
    if tmp.empty:
        return {}, {}
    canal = tmp.groupby(keys + ["canal"], as_index=False).agg(
        senal=("demanda_diaria_ponderada", "sum"),
        ventas_90d=("ventas_90d_unidades", "sum"),
    )
    stats = tmp.groupby(keys, as_index=False).agg(
        ventas_90d=("ventas_90d_unidades", "sum"),
        skus=("sku_madre", "nunique"),
        dias_venta=("dias_con_venta_90d", "sum"),
    )
    perfiles, muestras = {}, {}
    for _, r in canal.iterrows():
        key = tuple(r[k] for k in keys)
        perfiles.setdefault(key, {})[r["canal"]] = float(r["senal"] or 0)
    for _, r in stats.iterrows():
        key = tuple(r[k] for k in keys)
        muestras[key] = {
            "ventas_90d": float(r["ventas_90d"] or 0),
            "skus": int(r["skus"] or 0),
            "dias_venta": float(r["dias_venta"] or 0),
        }
    return perfiles, muestras


def construir_reparticion_stock_central(rotacion_canal, catalogo, meta_productos):
    cols_resumen = [
        "sku_madre", "producto_madre", "categoria", "marca",
        "stock_existencias", "stock_b2b", "stock_repartir",
        "metodo_reparticion", "confianza", "ventas_sku_90d", "dias_venta_sku_90d",
        "muestra_ventas_90d", "muestra_skus", "peso_base_igual",
        "ultima_fecha_arribo", "cantidad_ultimo_lote", "fundamento",
    ]
    cols_detalle = [
        "sku_madre", "producto_madre", "categoria", "marca", "canal",
        "stock_actual_canal", "ventas_10d_sku", "ventas_30d_sku", "ventas_90d_sku",
        "demanda_diaria_sku", "senal_referencia", "participacion_referencia",
        "participacion_objetivo", "stock_objetivo_teorico", "cantidad_sugerida",
        "stock_despues", "cobertura_despues_dias", "desde_existencias", "desde_b2b",
        "metodo_reparticion", "confianza", "fundamento",
    ]
    if rotacion_canal is None or rotacion_canal.empty:
        return pd.DataFrame(columns=cols_resumen), pd.DataFrame(columns=cols_detalle), pd.DataFrame()

    base = rotacion_canal.copy()
    base["sku_madre"] = base["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    base["canal"] = base["canal"].fillna("").astype(str).str.strip()
    for c in [
        "inventario_odoo", "inventario_total", "ventas_10d_unidades", "ventas_30d_unidades",
        "ventas_90d_unidades", "dias_con_venta_90d", "demanda_diaria_ponderada",
    ]:
        if c not in base.columns:
            base[c] = 0
        base[c] = to_num(base[c]).fillna(0).clip(lower=0)

    meta = meta_productos.copy() if meta_productos is not None else pd.DataFrame()
    if meta.empty:
        meta = pd.DataFrame(columns=["sku_madre", "categoria", "marca"])
    for c in ["sku_madre", "categoria", "marca"]:
        if c not in meta.columns:
            meta[c] = ""
    meta["sku_madre"] = meta["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    meta = meta.drop_duplicates("sku_madre")
    meta_map = meta.set_index("sku_madre")[["categoria", "marca"]].to_dict("index") if not meta.empty else {}

    # Historial comercial con metadatos para construir afinidad por categoría/marca.
    hist = base[base["canal"].isin(REPARTICION_CANALES)].copy()
    hist["categoria"] = hist["sku_madre"].map(lambda x: meta_map.get(x, {}).get("categoria", ""))
    hist["marca"] = hist["sku_madre"].map(lambda x: meta_map.get(x, {}).get("marca", ""))

    sku_perf, sku_stats = _perfil_dict(hist, ["sku_madre"])
    cb_perf, cb_stats = _perfil_dict(hist, ["categoria", "marca"])
    cat_perf, cat_stats = _perfil_dict(hist, ["categoria"])
    brand_perf, brand_stats = _perfil_dict(hist, ["marca"])

    cat_profile = pd.DataFrame()
    if not hist.empty:
        cat_profile = hist[hist["categoria"].astype(str).str.strip().ne("")].groupby(
            ["categoria", "canal"], as_index=False
        ).agg(
            demanda_diaria_ponderada=("demanda_diaria_ponderada", "sum"),
            ventas_90d_unidades=("ventas_90d_unidades", "sum"),
            skus=("sku_madre", "nunique"),
        )

    catalog = catalogo.copy() if catalogo is not None else pd.DataFrame()
    if not catalog.empty:
        catalog["sku_madre"] = catalog["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        catalog = catalog.drop_duplicates("sku_madre").set_index("sku_madre")

    resumen_rows, detalle_rows = [], []
    source_skus = sorted(set(base.loc[base["canal"].isin(["General", "B2B"]), "sku_madre"]) - {""})
    for sku in source_skus:
        g = base[base["sku_madre"].eq(sku)].copy()
        exist = float(g.loc[g["canal"].eq("General"), "inventario_odoo"].sum())
        b2b = float(g.loc[g["canal"].eq("B2B"), "inventario_odoo"].sum())
        exist_i, b2b_i = int(np.floor(max(exist, 0))), int(np.floor(max(b2b, 0)))
        pool = exist_i + b2b_i
        if pool <= 0:
            continue

        meta_sku = meta_map.get(sku, {})
        categoria = str(meta_sku.get("categoria", "") or "").strip()
        marca = str(meta_sku.get("marca", "") or "").strip()
        product = ""
        if not catalog.empty and sku in catalog.index:
            product = str(catalog.loc[sku].get("producto_madre", "") or "")
        if not product:
            product = _primero_texto_no_vacio(g.get("producto_madre", pd.Series(dtype=str)))

        direct_key = (sku,)
        direct_stat = sku_stats.get(direct_key, {"ventas_90d": 0, "skus": 1, "dias_venta": 0})
        direct_perf = sku_perf.get(direct_key, {})
        ventas_sku = float(direct_stat.get("ventas_90d", 0) or 0)
        dias_sku = float(direct_stat.get("dias_venta", 0) or 0)

        metodo, perf, muestra = "Parejo", {}, {"ventas_90d": 0, "skus": 0, "dias_venta": 0}
        if sum(direct_perf.values()) > 0 and (ventas_sku >= REPARTICION_MIN_VENTAS_SKU or dias_sku >= REPARTICION_MIN_DIAS_VENTA_SKU):
            metodo, perf, muestra = "Histórico SKU", direct_perf, direct_stat
        else:
            cb_key = (categoria, marca)
            cb_m = cb_stats.get(cb_key, {}) if categoria and marca else {}
            if (
                categoria and marca
                and sum(cb_perf.get(cb_key, {}).values()) > 0
                and float(cb_m.get("ventas_90d", 0) or 0) >= REPARTICION_MIN_VENTAS_CAT_MARCA
                and int(cb_m.get("skus", 0) or 0) >= REPARTICION_MIN_SKUS_CAT_MARCA
            ):
                metodo, perf, muestra = "Categoría + marca", cb_perf[cb_key], cb_m
            elif categoria and sum(cat_perf.get((categoria,), {}).values()) > 0:
                metodo, perf, muestra = "Categoría", cat_perf[(categoria,)], cat_stats.get((categoria,), {})
            elif marca and sum(brand_perf.get((marca,), {}).values()) > 0:
                metodo, perf, muestra = "Marca", brand_perf[(marca,)], brand_stats.get((marca,), {})

        # Confianza basada en cantidad de evidencia, no en una opinión subjetiva.
        mv = float(muestra.get("ventas_90d", 0) or 0)
        ms = int(muestra.get("skus", 0) or 0)
        md = float(muestra.get("dias_venta", 0) or 0)
        if metodo == "Histórico SKU":
            confianza = "ALTA" if ventas_sku >= 20 and dias_sku >= 8 else "MEDIA"
        elif metodo in {"Categoría + marca", "Categoría", "Marca"}:
            confianza = "ALTA" if mv >= 100 and ms >= 5 else ("MEDIA" if mv >= 30 and ms >= 3 else "BAJA")
        else:
            confianza = "BAJA"

        equal_share = np.ones(len(REPARTICION_CANALES), dtype=float) / len(REPARTICION_CANALES)
        ref_signal = np.array([max(float(perf.get(c, 0) or 0), 0) for c in REPARTICION_CANALES], dtype=float)
        if metodo == "Parejo" or ref_signal.sum() <= 0:
            target_share = equal_share.copy()
            ref_share = equal_share.copy()
        else:
            ref_share = ref_signal / ref_signal.sum()
            target_share = REPARTICION_PESO_BASE_IGUAL * equal_share + (1 - REPARTICION_PESO_BASE_IGUAL) * ref_share

        cur = []
        sku_rows = {}
        for c in REPARTICION_CANALES:
            rr = g[g["canal"].eq(c)]
            sku_rows[c] = rr.iloc[0] if not rr.empty else None
            cur.append(float(rr["inventario_total"].sum()) if not rr.empty else 0.0)
        current_total = float(sum(cur))

        if metodo == "Parejo":
            alloc = _asignar_enteros_por_peso(pool, np.ones(len(REPARTICION_CANALES)))
            desired = np.array(cur, dtype=float) + np.array(alloc, dtype=float)
        else:
            desired = target_share * (current_total + pool)
            deficits = np.maximum(desired - np.array(cur, dtype=float), 0)
            alloc = _asignar_enteros_por_peso(pool, deficits if deficits.sum() > 0 else target_share)

        # Divide las fuentes proporcionalmente entre destinos para no sesgar un
        # canal por el orden de la tabla. La suma sigue siendo exacta.
        desde_exist = _asignar_enteros_por_peso(exist_i, alloc) if exist_i else [0] * len(alloc)
        desde_b2b = [max(int(q) - int(e), 0) for q, e in zip(alloc, desde_exist)]

        if metodo == "Histórico SKU":
            fundamento = f"Ritmo del IQ: {ventas_sku:.0f} ventas en 90d y {dias_sku:.0f} días con venta; {REPARTICION_PESO_BASE_IGUAL:.0%} base parejo + desempeño reciente."
        elif metodo == "Categoría + marca":
            fundamento = f"Sin historia suficiente del IQ; se usa afinidad de {categoria} / {marca} ({mv:.0f} ventas, {ms} IQ)."
        elif metodo == "Categoría":
            fundamento = f"Sin historia suficiente del IQ; se usa desempeño de la categoría {categoria} ({mv:.0f} ventas, {ms} IQ)."
        elif metodo == "Marca":
            fundamento = f"Sin historia suficiente del IQ/categoría; se usa desempeño de la marca {marca} ({mv:.0f} ventas, {ms} IQ)."
        else:
            fundamento = "Sin histórico útil de IQ, categoría o marca: el stock nuevo se reparte parejo entre canales."

        ultima_fecha = ""
        cantidad_lote = 0
        if not catalog.empty and sku in catalog.index:
            cr = catalog.loc[sku]
            ultima_fecha = cr.get("ultima_fecha_arribo", cr.get("ultima_fecha_compra_odoo", ""))
            cantidad_lote = float(cr.get("cantidad_ultimo_lote", 0) or 0)

        resumen_rows.append({
            "sku_madre": sku, "producto_madre": product, "categoria": categoria, "marca": marca,
            "stock_existencias": exist_i, "stock_b2b": b2b_i, "stock_repartir": pool,
            "metodo_reparticion": metodo, "confianza": confianza,
            "ventas_sku_90d": ventas_sku, "dias_venta_sku_90d": dias_sku,
            "muestra_ventas_90d": mv, "muestra_skus": ms,
            "peso_base_igual": REPARTICION_PESO_BASE_IGUAL,
            "ultima_fecha_arribo": ultima_fecha, "cantidad_ultimo_lote": cantidad_lote,
            "fundamento": fundamento,
        })

        for i, c in enumerate(REPARTICION_CANALES):
            rr = sku_rows[c]
            v10 = float(rr.get("ventas_10d_unidades", 0) or 0) if rr is not None else 0.0
            v30 = float(rr.get("ventas_30d_unidades", 0) or 0) if rr is not None else 0.0
            v90 = float(rr.get("ventas_90d_unidades", 0) or 0) if rr is not None else 0.0
            dsku = float(rr.get("demanda_diaria_ponderada", 0) or 0) if rr is not None else 0.0
            after = float(cur[i]) + int(alloc[i])
            cobertura = after / dsku if dsku > 0 else np.nan
            detalle_rows.append({
                "sku_madre": sku, "producto_madre": product, "categoria": categoria, "marca": marca,
                "canal": c, "stock_actual_canal": cur[i],
                "ventas_10d_sku": v10, "ventas_30d_sku": v30, "ventas_90d_sku": v90,
                "demanda_diaria_sku": dsku, "senal_referencia": float(ref_signal[i]),
                "participacion_referencia": float(ref_share[i]),
                "participacion_objetivo": float(target_share[i]),
                "stock_objetivo_teorico": float(desired[i]),
                "cantidad_sugerida": int(alloc[i]), "stock_despues": after,
                "cobertura_despues_dias": cobertura,
                "desde_existencias": int(desde_exist[i]), "desde_b2b": int(desde_b2b[i]),
                "metodo_reparticion": metodo, "confianza": confianza, "fundamento": fundamento,
            })

    resumen = pd.DataFrame(resumen_rows, columns=cols_resumen)
    detalle = pd.DataFrame(detalle_rows, columns=cols_detalle)
    if not resumen.empty:
        resumen = resumen.sort_values(["stock_repartir", "confianza"], ascending=[False, True]).reset_index(drop=True)
    if not detalle.empty:
        detalle = detalle.sort_values(["sku_madre", "cantidad_sugerida", "canal"], ascending=[True, False, True]).reset_index(drop=True)
    return resumen, detalle, cat_profile


reparticion_resumen, reparticion_detalle, reparticion_perfiles_categoria = construir_reparticion_stock_central(
    dashboard_rotacion_canal, dashboard_productos, metadata_reparticion
)

# El total se valida contra la suma de todos los canales y sus tres bolsas.
if not dashboard_rotacion_canal.empty:
    control_total_canal = dashboard_rotacion_canal.groupby("sku_madre", as_index=False).agg(
        stock_suma_canales=("inventario_total", "sum"),
        odoo_suma_canales=("inventario_odoo", "sum"),
        transito_suma_canales=("inventario_transito", "sum"),
        full_suma_canales=("inventario_full", "sum"),
        ventas_suma_canales_90d=("ventas_90d_unidades", "sum"),
    )
    control_total_canal = control_total_canal.merge(
        dashboard_productos[["sku_madre", "stock_total", "ventas_3m_unidades"]],
        on="sku_madre", how="outer"
    )
    for c in ["stock_suma_canales", "stock_total", "ventas_suma_canales_90d", "ventas_3m_unidades"]:
        control_total_canal[c] = to_num(control_total_canal.get(c, 0))
    control_total_canal["diferencia_stock"] = control_total_canal["stock_suma_canales"] - control_total_canal["stock_total"]
    control_total_canal["diferencia_ventas"] = control_total_canal["ventas_suma_canales_90d"] - control_total_canal["ventas_3m_unidades"]
else:
    control_total_canal = pd.DataFrame()

# ============================================================
# NO VINCULADAS Y TABLAS SECUNDARIAS
# ============================================================

logs_odoo_no = extraer_logs_odoo(fecha_inicio_3m, fecha_fin, dic_map)

# IMPORTANTE:
# Ventas no vinculadas ahora toma SOLO logs reales de Odoo.
# No usa ventas_sin_referencia del Excel operativo ni sale.order.line.
ventas_no_vinculadas = logs_odoo_no.copy()

# quitar duplicados simples por fuente/pedido/sku/cantidad
if not ventas_no_vinculadas.empty:
    for c in ["fuente_log", "pedido", "referencia", "sku_log", "sku_original", "producto", "cantidad"]:
        if c not in ventas_no_vinculadas.columns:
            ventas_no_vinculadas[c] = ""
    ventas_no_vinculadas["_dedup"] = (
        ventas_no_vinculadas["fuente_log"].astype(str) + "|" +
        ventas_no_vinculadas["pedido"].astype(str) + "|" +
        ventas_no_vinculadas["referencia"].astype(str) + "|" +
        ventas_no_vinculadas["sku_log"].astype(str) + "|" +
        ventas_no_vinculadas["sku_original"].astype(str) + "|" +
        ventas_no_vinculadas["cantidad"].astype(str)
    )
    ventas_no_vinculadas = ventas_no_vinculadas.drop_duplicates("_dedup", keep="first").drop(columns=["_dedup"])

dashboard_top80 = dashboard_productos[dashboard_productos["top_80_flag"] == "SI"].copy()

dashboard_alertas = dashboard_productos[
    dashboard_productos["alerta_compra"].astype(str).str.contains("Comprar|urgencia", case=False, na=False)
    | dashboard_productos["riesgo_sobrestock"].astype(str).str.contains(r"\+60|\+90", regex=True, na=False)
].copy()

dashboard_stock_canal = (
    inventario_canal_operativo.copy()
    if not inventario_canal_operativo.empty
    else dashboard_productos[[
        "sku_madre", "producto_madre", "stock_odoo_cuautitlan",
        "stock_amazon_fba", "stock_meli_full", "stock_walmart_wfs",
        "stock_liverpool_99min", "stock_total"
    ]].copy()
)

dashboard_ventas_canal = ventas_3m.copy()
if not dashboard_ventas_canal.empty:
    if "canal_venta" in dashboard_ventas_canal.columns:
        canal_serie = dashboard_ventas_canal["canal_venta"]
    elif "canal" in dashboard_ventas_canal.columns:
        canal_serie = dashboard_ventas_canal["canal"]
    else:
        canal_serie = pd.Series("Sin identificar", index=dashboard_ventas_canal.index)
    dashboard_ventas_canal["canal_venta"] = canal_serie.apply(normalizar_canal_dashboard)
    if "modalidad_venta" in dashboard_ventas_canal.columns:
        modalidad_serie = dashboard_ventas_canal["modalidad_venta"]
    else:
        modalidad_serie = pd.Series("SIN_IDENTIFICAR", index=dashboard_ventas_canal.index)
    dashboard_ventas_canal["modalidad_venta"] = modalidad_serie.astype(str).str.upper()
    dashboard_ventas_canal = dashboard_ventas_canal.groupby(
        ["sku_madre", "canal_venta", "modalidad_venta", "fuente"], as_index=False
    ).agg(
        unidades=("cantidad", "sum"),
        venta_total=("venta_total", "sum"),
        pedidos=("pedido", "nunique")
    )


# ============================================================
# VENTAS DIARIAS PARA CONSULTA Y DESCARGA EN EL DASHBOARD
# ============================================================
# Se conserva un detalle diario de los últimos 90 días por:
# IQ + canal + SKU original + modalidad + fuente.
# Esto permite consultar y exportar ventas sin cargar cada línea transaccional
# dentro del HTML, manteniendo un archivo razonable y auditable.
def construir_ventas_detalle_dashboard(ventas_periodo):
    columnas = [
        "fecha", "sku_madre", "producto_madre", "canal", "sku_original",
        "modalidad_venta", "fuente", "unidades", "venta_total", "pedidos"
    ]
    if ventas_periodo is None or ventas_periodo.empty:
        return pd.DataFrame(columns=columnas)

    df = ventas_periodo.copy()
    df["fecha"] = pd.to_datetime(df.get("fecha", pd.NaT), errors="coerce").dt.normalize()
    df["cantidad"] = to_num(df.get("cantidad", 0))
    df["venta_total"] = to_num(df.get("venta_total", 0))

    def serie_texto(col):
        if col not in df.columns:
            return pd.Series("", index=df.index, dtype="object")
        out = df[col].fillna("").astype(str).str.strip()
        return out.mask(out.str.lower().isin(["nan", "none", "false", "falso"]), "")

    # Elegir canal sin confundir la modalidad Full/Drop con el marketplace.
    canal_cols = []
    for col in ["canal_venta", "equipo_ventas", "canal", "tienda"]:
        serie = serie_texto(col)
        serie = serie.mask(
            serie.str.lower().isin(["full", "drop", "pendiente_odoo", "sin identificar"]),
            "",
        )
        canal_cols.append(serie)
    canal_raw = pd.concat(canal_cols, axis=1).replace("", np.nan).bfill(axis=1).iloc[:, 0].fillna("")
    df["canal_dashboard"] = canal_raw.apply(normalizar_canal_dashboard)

    # Elegir solo identificadores de producto; nunca referencias de pedido.
    sku_cols = []
    for col in [
        "sku_original", "sku_default_code", "sku_usado_para_match",
        "alias_diccionario", "sku_desde_columna", "sku_desde_titulo"
    ]:
        sku_cols.append(serie_texto(col).apply(limpiar_sku))
    sku_raw = pd.concat(sku_cols, axis=1).replace("", np.nan).bfill(axis=1).iloc[:, 0]
    df["sku_original_dashboard"] = sku_raw.fillna(serie_texto("sku_madre"))

    modalidad_cols = [serie_texto("modalidad_venta"), serie_texto("tipo_venta")]
    modalidad_raw = pd.concat(modalidad_cols, axis=1).replace("", np.nan).bfill(axis=1).iloc[:, 0].fillna("")
    modalidad_fallback = serie_texto("canal").str.upper().where(
        serie_texto("canal").str.upper().isin(["FULL", "DROP"]),
        "SIN_IDENTIFICAR",
    )
    df["modalidad_dashboard"] = modalidad_raw.str.upper().where(modalidad_raw.ne(""), modalidad_fallback)

    df["producto_madre"] = serie_texto("producto_madre")
    if "producto" in df.columns:
        faltante_producto = df["producto_madre"].eq("")
        df.loc[faltante_producto, "producto_madre"] = serie_texto("producto").loc[faltante_producto]
    df["fuente"] = serie_texto("fuente")
    df["pedido"] = serie_texto("pedido")
    df["sku_madre"] = serie_texto("sku_madre").str.upper()

    df = df[
        df["fecha"].notna()
        & df["sku_madre"].ne("")
        & df["canal_dashboard"].ne("Sin identificar")
    ].copy()

    if df.empty:
        return pd.DataFrame(columns=columnas)

    detalle = df.groupby(
        [
            "fecha", "sku_madre", "producto_madre", "canal_dashboard",
            "sku_original_dashboard", "modalidad_dashboard", "fuente"
        ],
        as_index=False,
        dropna=False,
    ).agg(
        unidades=("cantidad", "sum"),
        venta_total=("venta_total", "sum"),
        pedidos=("pedido", lambda x: x[x.astype(str).str.strip().ne("")].nunique()),
    )

    detalle = detalle.rename(columns={
        "canal_dashboard": "canal",
        "sku_original_dashboard": "sku_original",
        "modalidad_dashboard": "modalidad_venta",
    })
    return detalle[columnas].sort_values(
        ["fecha", "sku_madre", "canal", "sku_original"],
        ascending=[False, True, True, True]
    )

ventas_detalle_dashboard = construir_ventas_detalle_dashboard(ventas_3m)

parametros_lead_time = dashboard_productos[["sku_madre", "producto_madre", "lead_time_dias", "cobertura_objetivo_dias"]].copy()
parametros_lead_time["comentario"] = "Editar lead_time_dias si aplica por producto/proveedor"

resumen = pd.DataFrame([
    ["archivo_entrada", str(archivo_entrada)],
    ["fecha_inicio_3m", fecha_inicio_3m.strftime("%Y-%m-%d")],
    ["fecha_fin_3m", fecha_fin.strftime("%Y-%m-%d")],
    ["dias_analisis_3m", DIAS_ANALISIS_3M],
    ["lead_time_default", LEAD_TIME_DEFAULT],
    ["cobertura_objetivo_dias", COBERTURA_OBJETIVO_DIAS],
    ["usar_stock_congelado", USAR_STOCK_CONGELADO],
    ["fecha_corte_stock", FECHA_CORTE_STOCK.strftime("%Y-%m-%d %H:%M:%S")],
    ["archivo_stock_congelado", str(ARCHIVO_STOCK_CONGELADO)],
    ["arribos_odoo_compras_activo", ENABLE_ODOO_ARRIBOS_COMPRAS],
    ["arribos_post_corte_unidades", float(dashboard_productos["arribos_post_corte_unidades"].sum()) if "arribos_post_corte_unidades" in dashboard_productos.columns else 0],
    ["arribos_post_corte_monto", float(dashboard_productos["arribos_post_corte_monto"].sum()) if "arribos_post_corte_monto" in dashboard_productos.columns else 0],
    ["renglones_arribos_odoo_compras", len(arribos_odoo_compras_detalle) if "arribos_odoo_compras_detalle" in globals() and isinstance(arribos_odoo_compras_detalle, pd.DataFrame) else 0],
    ["iq_con_imagen_odoo", int(dashboard_productos["tiene_imagen_odoo"].astype(str).str.upper().eq("SI").sum()) if "tiene_imagen_odoo" in dashboard_productos.columns else 0],
    ["iq_en_dashboard", len(dashboard_productos)],
    ["iq_top80", int((dashboard_productos["top_80_flag"] == "SI").sum())],
    ["ventas_3m_unidades", float(dashboard_productos["ventas_3m_unidades"].sum())],
    ["ventas_3m_monto", float(dashboard_productos["ventas_3m_monto"].sum())],
    ["iq_ritmo_congelado_stockout", int(dashboard_productos.get("ritmo_congelado_stockout", pd.Series(dtype=str)).astype(str).str.upper().eq("SI").sum())],
    ["dias_ritmo_stock", DIAS_RITMO_STOCK],
    ["stock_total", float(dashboard_productos["stock_total"].sum())],
    ["inventario_canal_renglones", len(dashboard_rotacion_canal)],
    ["traslados_full_renglones", len(traslados_full_operativo)],
    ["diferencia_stock_total_canales", float(control_total_canal["diferencia_stock"].sum()) if not control_total_canal.empty else 0],
    ["comprar_ya", int(dashboard_productos["alerta_compra"].astype(str).str.contains("Comprar ya", case=False, na=False).sum())],
    ["compra_con_urgencia", int(dashboard_productos["alerta_compra"].astype(str).str.contains("urgencia", case=False, na=False).sum())],
    ["riesgo_90", int(dashboard_productos["riesgo_sobrestock"].astype(str).str.contains(r"\+90", regex=True, na=False).sum())],
    ["ventas_no_vinculadas", len(ventas_no_vinculadas)],
    ["stock_no_vinculado_skus", int(stock_no_vinculado_skus)],
    ["stock_no_vinculado_unidades", float(stock_no_vinculado_unidades)],
], columns=["metrica", "valor"])


# ============================================================
# EXPORTAR
# ============================================================

with pd.ExcelWriter(ARCHIVO_SALIDA, engine="openpyxl") as writer:
    resumen.to_excel(writer, sheet_name="resumen", index=False)
    dashboard_productos.to_excel(writer, sheet_name="dashboard_productos", index=False)
    dashboard_top80.to_excel(writer, sheet_name="dashboard_top80", index=False)
    dashboard_alertas.to_excel(writer, sheet_name="dashboard_alertas", index=False)
    dashboard_stock_canal.to_excel(writer, sheet_name="dashboard_stock_canal", index=False)
    dashboard_ventas_canal.to_excel(writer, sheet_name="dashboard_ventas_canal", index=False)
    ventas_detalle_dashboard.to_excel(writer, sheet_name="ventas_detalle_dashboard", index=False)
    ritmo_stock_90d.to_excel(writer, sheet_name="ritmo_stock_90d", index=False)
    ritmo_stock_diario.to_excel(writer, sheet_name="ritmo_stock_diario", index=False)
    dashboard_rotacion_canal.to_excel(writer, sheet_name="rotacion_por_canal", index=False)
    redistribuciones_sugeridas.to_excel(writer, sheet_name="redistribucion_interna", index=False)
    reparticion_resumen.to_excel(writer, sheet_name="reparticion_resumen", index=False)
    reparticion_detalle.to_excel(writer, sheet_name="reparticion_detalle", index=False)
    if isinstance(reparticion_perfiles_categoria, pd.DataFrame) and not reparticion_perfiles_categoria.empty:
        reparticion_perfiles_categoria.to_excel(writer, sheet_name="reparticion_perfiles_cat", index=False)
    control_total_canal.to_excel(writer, sheet_name="control_total_canal", index=False)
    if not inventario_canal_operativo.empty:
        inventario_canal_operativo.to_excel(writer, sheet_name="inventario_canal_base", index=False)
    if not traslados_full_operativo.empty:
        traslados_full_operativo.to_excel(writer, sheet_name="traslados_full", index=False)
    if not traslados_full_excluidos.empty:
        traslados_full_excluidos.to_excel(writer, sheet_name="traslados_excluidos", index=False)
    ventas_no_vinculadas.to_excel(writer, sheet_name="ventas_no_vinculadas", index=False)
    stock_no_vinculado.to_excel(writer, sheet_name="stock_no_vinculado", index=False)
    # Hojas nuevas (versión 2026-09) que consume el 03.
    for _nombre, _df in [
        ("antiguedad_capas", antiguedad_capas),
        ("movimientos_odoo_auditoria", movimientos_odoo_auditoria),
        ("movimientos_odoo_resumen", movimientos_odoo_resumen),
        ("costos_sku_madre", costos_sku_madre),
        ("publicaciones_autoazur", publicaciones_autoazur),
        ("publicaciones_sku_canal", publicaciones_sku_canal),
        ("publicaciones_sin_sku_madre", publicaciones_sin_sku_madre),
        ("publicaciones_log", publicaciones_log),
        ("publicaciones_sin_sku_resumen", publicaciones_sin_sku_resumen),
    ]:
        if isinstance(_df, pd.DataFrame) and not _df.empty:
            _df.to_excel(writer, sheet_name=_nombre, index=False)
    if "auditoria_stock_congelado" in globals() and isinstance(auditoria_stock_congelado, pd.DataFrame):
        auditoria_stock_congelado.to_excel(writer, sheet_name="auditoria_stock_congelado", index=False)
    if "arribos_odoo_compras_detalle" in globals() and isinstance(arribos_odoo_compras_detalle, pd.DataFrame):
        arribos_odoo_compras_detalle.to_excel(writer, sheet_name="arribos_odoo_compras", index=False)
    parametros_lead_time.to_excel(writer, sheet_name="parametros_lead_time", index=False)
    dic_aliases.to_excel(writer, sheet_name="sku_marketplace_aliases", index=False)
    if "diccionario_origen4_directo" in globals() and isinstance(diccionario_origen4_directo, pd.DataFrame) and not diccionario_origen4_directo.empty:
        diccionario_origen4_directo.to_excel(writer, sheet_name="diccionario_origen4_directo", index=False)
    if "sku_por_canal_detalle" in globals() and isinstance(sku_por_canal_detalle, pd.DataFrame):
        sku_por_canal_detalle.to_excel(writer, sheet_name="sku_por_canal_detalle", index=False)
    if "sku_por_canal" in globals() and isinstance(sku_por_canal, pd.DataFrame):
        sku_por_canal.to_excel(writer, sheet_name="sku_por_canal_resumen", index=False)
    if "sku_aliases_excluidos_visual" in globals() and isinstance(sku_aliases_excluidos_visual, pd.DataFrame):
        sku_aliases_excluidos_visual.to_excel(writer, sheet_name="sku_aliases_excluidos", index=False)

    wb = writer.book
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True, color="FFFFFF")
            cell.fill = cell.fill.copy(fill_type="solid", fgColor="111827")
        for col_cells in ws.columns:
            max_len = 0
            col_letter = col_cells[0].column_letter
            for cell in col_cells[:1000]:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 48)

print("\nBASE DASHBOARD V20 GENERADA")
print(f"Archivo: {ARCHIVO_SALIDA}")
print(f"IQ en dashboard: {len(dashboard_productos)}")
print(f"Top 80%: {(dashboard_productos['top_80_flag'] == 'SI').sum()}")
print(f"Ventas no vinculadas: {len(ventas_no_vinculadas)}")
print(f"Stock no vinculado SKUs: {stock_no_vinculado_skus}")
print(f"Stock congelado activo: {USAR_STOCK_CONGELADO}")
print(f"Fecha corte stock: {FECHA_CORTE_STOCK}")
print(f"Archivo stock congelado: {ARCHIVO_STOCK_CONGELADO}")
print(f"Arribos Odoo compras activo: {ENABLE_ODOO_ARRIBOS_COMPRAS}")
if "arribos_odoo_compras_detalle" in globals() and isinstance(arribos_odoo_compras_detalle, pd.DataFrame):
    print(f"Renglones arribos Odoo compras: {len(arribos_odoo_compras_detalle)}")
    if not arribos_odoo_compras_detalle.empty and "cantidad_arribada" in arribos_odoo_compras_detalle.columns:
        print(f"Cantidad arribada Odoo compras: {arribos_odoo_compras_detalle['cantidad_arribada'].sum():,.0f}")

print(f"SKU por canal IQ: {len(sku_por_canal) if 'sku_por_canal' in globals() else 0}")
