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
- Promedio diario se calcula solo con días donde hubo ventas > 0.
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

warnings.filterwarnings("ignore", category=UserWarning)


# ============================================================
# CONFIGURACIÓN
# ============================================================

#DESKTOP = Path.home() / "Desktop"
#CARPETA = DESKTOP / "rotacion_inventario_base_dashboard"
DESKTOP = Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")
CARPETA = Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")


ARCHIVO_ENTRADA = CARPETA / "rotacion_inventario_base_dashboard_odoo_autoazur.xlsx"
ARCHIVO_SALIDA = CARPETA / "base_dashboard_rotacion.xlsx"

# ============================================================
# STOCK CONGELADO
# ============================================================
# El inventario base oficial queda congelado al:
# 02 de junio de 2026 a las 15:00 hrs.
#
# A partir de ese corte, el dashboard NO debe usar el stock descargado
# como stock actual final, sino:
#
# stock_actual_calculado = stock_congelado_02062026_1500 - ventas_posteriores_al_corte
#
# La primera vez que corras este script, si no existe el archivo
# stock_congelado_02062026_1500.xlsx, el código lo crea usando el stock
# que venga en el Excel operativo de entrada.
#
# IMPORTANTE:
# Esa primera corrida debe hacerse usando una base operativa que tenga el
# stock correcto del 02/06/2026 a las 15:00 hrs.
USAR_STOCK_CONGELADO = True
FECHA_CORTE_STOCK = pd.Timestamp("2026-06-02 15:00:00")
ARCHIVO_STOCK_CONGELADO = CARPETA / "stock_congelado_02062026_1500.xlsx"

LEAD_TIME_DEFAULT = 30
COBERTURA_OBJETIVO_DIAS = 45
DIAS_ANALISIS_3M = 90

# Opcional: Odoo para logs/no vinculadas.
ENABLE_ODOO_LOGS = True
ODOO_URL = os.getenv("ODOO_URL", "https://comercializadora-iqtech-productos-innovadores.odoo.com")
ODOO_DB = os.getenv("ODOO_DB", "comercializadora-iqtech-productos-innovadores-sh-ma-27238691")
ODOO_USER = os.getenv("ODOO_USER", "hocampou@gmail.com")
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "d4340db6fc78dcba7c75bda7ba5fe2f5e6d57347")
ODOO_CAMPO_TIPO_VENTA = os.getenv("ODOO_CAMPO_TIPO_VENTA", "x_studio_tipo_de_venta")
ODOO_TIPOS_VENTA_VALIDOS = ["full", "drop"]

# Para ventas no vinculadas ahora usaremos SOLO logs reales de Odoo.
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
        or ODOO_API_KEY.startswith("TU_")
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


def cargar_diccionario_usado(diccionario_df):
    if diccionario_df.empty:
        return pd.DataFrame(columns=[
            "sku_madre", "sku_sincronizado", "marketplace", "columna_alias", "hoja_diccionario"
        ])

    df = diccionario_df.copy()

    for c in ["sku_madre", "alias_diccionario", "columna_alias", "hoja_diccionario"]:
        if c not in df.columns:
            df[c] = ""

    df["sku_sincronizado"] = df["alias_diccionario"].apply(limpiar_sku)
    df = df[df["sku_madre"].notna() & df["sku_sincronizado"].ne("")].copy()

    df["marketplace"] = df.apply(
        lambda r: clasificar_marketplace(
            r.get("sku_sincronizado", ""),
            r.get("columna_alias", ""),
            r.get("hoja_diccionario", "")
        ),
        axis=1
    )

    # Quitar aliases evidentemente vacíos o falsos.
    df = df[
        df["sku_sincronizado"].astype(str).str.lower().isin(["false", "falso", "nan", "none"]) == False
    ].copy()

    return df[["sku_madre", "sku_sincronizado", "marketplace", "columna_alias", "hoja_diccionario"]].drop_duplicates()


def crear_chips_por_iq(dic_aliases):
    if dic_aliases.empty:
        return pd.DataFrame(columns=["sku_madre", "sku_chips_json", "skus_sincronizados", "num_skus_sincronizados"])

    rows = []
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
        common = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/common")
        self.uid = common.authenticate(self.db, self.user, self.api_key, {})
        if not self.uid:
            raise ConnectionError("No se pudo autenticar en Odoo.")
        self.models = xmlrpc.client.ServerProxy(f"{self.url}/xmlrpc/2/object")
        print(f"Conectado a Odoo logs. UID: {self.uid}")

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

dic_aliases = cargar_diccionario_usado(diccionario_usado)
sku_chips = crear_chips_por_iq(dic_aliases)
dic_map = construir_dic_map(diccionario_usado)

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

# Producto madre desde ventas.
if "producto_madre" in ventas_link.columns:
    prod_ventas = (
        ventas_link[["sku_madre", "producto_madre"]]
        .dropna()
        .drop_duplicates("sku_madre")
    )
else:
    prod_ventas = pd.DataFrame(columns=["sku_madre", "producto_madre"])



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
    stock actual calculado = stock congelado - ventas posteriores al corte.

    El descuento se hace primero sobre stock_total.
    Para conservar el desglose por canal, se descuenta proporcionalmente de cada canal
    según la composición del stock congelado.
    """
    if not USAR_STOCK_CONGELADO:
        return stock_iq_original.copy(), pd.DataFrame()

    stock_congelado = cargar_o_crear_stock_congelado(stock_iq_original)
    ventas_post = calcular_ventas_posteriores_corte(ventas_df)

    base = stock_congelado.merge(ventas_post, on="sku_madre", how="left")

    for c in ["ventas_post_corte_unidades", "ventas_post_corte_monto", "pedidos_post_corte"]:
        if c not in base.columns:
            base[c] = 0
        base[c] = to_num(base[c])

    base["stock_total_antes_descuento"] = base["stock_total_congelado"]
    base["stock_total"] = np.maximum(
        base["stock_total_congelado"] - base["ventas_post_corte_unidades"],
        0
    )

    # Factor proporcional para bajar stock de canales sin inventar asignación por canal.
    base["factor_stock_restante"] = np.where(
        base["stock_total_congelado"] > 0,
        base["stock_total"] / base["stock_total_congelado"],
        0
    )

    canales = [
        ("stock_walmart_wfs_congelado", "stock_walmart_wfs"),
        ("stock_liverpool_99min_congelado", "stock_liverpool_99min"),
        ("stock_meli_full_congelado", "stock_meli_full"),
        ("stock_amazon_fba_congelado", "stock_amazon_fba"),
        ("stock_odoo_cuautitlan_congelado", "stock_odoo_cuautitlan"),
    ]

    for congelado, actual in canales:
        if congelado not in base.columns:
            base[congelado] = 0
        base[actual] = np.maximum(
            np.floor(base[congelado] * base["factor_stock_restante"]),
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
    base["formula_stock"] = "stock_congelado_02062026_1500 - ventas_posteriores_al_corte"

    # Deja columnas compatibles con el resto del script.
    cols = [
        "sku_madre", "producto_madre",
        "stock_walmart_wfs", "stock_liverpool_99min", "stock_meli_full",
        "stock_amazon_fba", "stock_odoo_cuautitlan", "stock_total",
        "stock_total_congelado", "ventas_post_corte_unidades", "ventas_post_corte_monto",
        "pedidos_post_corte", "stock_total_antes_descuento",
        "stock_calculado_desde_congelado", "fecha_corte_stock", "formula_stock"
    ]
    for c in cols:
        if c not in base.columns:
            base[c] = ""

    auditoria = base.copy()

    return base[cols].copy(), auditoria


# ============================================================
# STOCK
# ============================================================

if stock_iq.empty and not rotacion.empty:
    stock_iq = rotacion.copy()

# Aplicar stock congelado:
# stock actual calculado = stock congelado al 02/06/2026 15:00 - ventas posteriores al corte.
stock_iq, auditoria_stock_congelado = aplicar_descuento_stock_congelado(stock_iq, ventas)

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

stock_base = (
    stock_iq.groupby("sku_madre", as_index=False)
    .agg(
        producto_madre_stock=("producto_madre", first_non_empty),
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
        "sku_madre", "stock_total_congelado", "ventas_post_corte_unidades",
        "ventas_post_corte_monto", "pedidos_post_corte",
        "stock_total_antes_descuento", "stock_calculado_desde_congelado",
        "fecha_corte_stock", "formula_stock"
    ]
    cols_aud = [c for c in cols_aud if c in auditoria_stock_congelado.columns]
    dashboard = dashboard.merge(
        auditoria_stock_congelado[cols_aud].drop_duplicates("sku_madre"),
        on="sku_madre",
        how="left"
    )

dashboard = dashboard.merge(ventas_iq_3m, on="sku_madre", how="left")
dashboard = dashboard.merge(prod_ventas, on="sku_madre", how="left")
dashboard = dashboard.merge(sku_chips, on="sku_madre", how="left")

dashboard["producto_madre"] = dashboard["producto_madre"].fillna(dashboard.get("producto_madre_stock", ""))
dashboard["producto_madre"] = dashboard["producto_madre"].fillna("")

for col in [
    "ventas_3m_unidades",
    "ventas_3m_monto",
    "pedidos_3m",
    "dias_con_venta_3m",
    "venta_diaria_promedio_positiva_3m",
    "stock_walmart_wfs",
    "stock_liverpool_99min",
    "stock_meli_full",
    "stock_amazon_fba",
    "stock_odoo_cuautitlan",
    "stock_total",
    "stock_total_congelado",
    "ventas_post_corte_unidades",
    "ventas_post_corte_monto",
    "pedidos_post_corte",
    "stock_total_antes_descuento",
    "num_skus_sincronizados",
]:
    if col not in dashboard.columns:
        dashboard[col] = 0
    dashboard[col] = to_num(dashboard[col])

dashboard["venta_diaria_promedio_3m"] = dashboard["venta_diaria_promedio_positiva_3m"]

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


def prioridad(alerta, riesgo):
    alerta = str(alerta)
    riesgo = str(riesgo)
    if "Comprar ya" in alerta:
        return "Alta"
    if "urgencia" in alerta:
        return "Media"
    if "+90" in riesgo:
        return "Revisar sobrestock"
    if "+60" in riesgo:
        return "Media sobrestock"
    return "Baja"


def recomendacion(row):
    alerta = str(row.get("alerta_compra", ""))
    riesgo = str(row.get("riesgo_sobrestock", ""))
    sug = float(row.get("sugerencia_compra", 0) or 0)

    if "Comprar ya" in alerta:
        return f"Comprar ya / sugerido {sug:,.0f} unidades"
    if "urgencia" in alerta:
        return f"Compra con urgencia / sugerido {sug:,.0f} unidades"
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
dashboard["prioridad"] = dashboard.apply(lambda r: prioridad(r["alerta_compra"], r["riesgo_sobrestock"]), axis=1)
dashboard["recomendacion"] = dashboard.apply(recomendacion, axis=1)

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

dashboard["sku_chips_json"] = dashboard["sku_chips_json"].fillna("[]")
dashboard["skus_sincronizados"] = dashboard["skus_sincronizados"].fillna("")

# Columnas ordenadas.
cols_dashboard = [
    "sku_madre", "producto_madre", "ranking_ventas_3m", "ranking_monto_3m", "ranking_top80",
    "skus_sincronizados", "sku_chips_json", "num_skus_sincronizados",
    "ventas_3m_unidades", "ventas_3m_monto", "pedidos_3m", "dias_con_venta_3m",
    "venta_diaria_promedio_3m",
    "stock_total", "stock_total_congelado", "ventas_post_corte_unidades",
    "ventas_post_corte_monto", "pedidos_post_corte", "stock_total_antes_descuento",
    "stock_calculado_desde_congelado", "fecha_corte_stock", "formula_stock",
    "stock_odoo_cuautitlan", "stock_amazon_fba", "stock_meli_full",
    "stock_walmart_wfs", "stock_liverpool_99min",
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

dashboard_stock_canal = dashboard_productos[
    [
        "sku_madre", "producto_madre",
        "stock_odoo_cuautitlan", "stock_amazon_fba", "stock_meli_full",
        "stock_walmart_wfs", "stock_liverpool_99min", "stock_total",
    ]
].copy()

dashboard_ventas_canal = ventas_3m.copy()
if not dashboard_ventas_canal.empty:
    dashboard_ventas_canal = (
        dashboard_ventas_canal.groupby(["sku_madre", "fuente", "canal"], as_index=False)
        .agg(unidades=("cantidad", "sum"), venta_total=("venta_total", "sum"), pedidos=("pedido", "nunique"))
    )

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
    ["iq_en_dashboard", len(dashboard_productos)],
    ["iq_top80", int((dashboard_productos["top_80_flag"] == "SI").sum())],
    ["ventas_3m_unidades", float(dashboard_productos["ventas_3m_unidades"].sum())],
    ["ventas_3m_monto", float(dashboard_productos["ventas_3m_monto"].sum())],
    ["stock_total", float(dashboard_productos["stock_total"].sum())],
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
    ventas_no_vinculadas.to_excel(writer, sheet_name="ventas_no_vinculadas", index=False)
    stock_no_vinculado.to_excel(writer, sheet_name="stock_no_vinculado", index=False)
    if "auditoria_stock_congelado" in globals() and isinstance(auditoria_stock_congelado, pd.DataFrame):
        auditoria_stock_congelado.to_excel(writer, sheet_name="auditoria_stock_congelado", index=False)
    parametros_lead_time.to_excel(writer, sheet_name="parametros_lead_time", index=False)
    dic_aliases.to_excel(writer, sheet_name="sku_marketplace_aliases", index=False)

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

print("\nBASE DASHBOARD V2 GENERADA")
print(f"Archivo: {ARCHIVO_SALIDA}")
print(f"IQ en dashboard: {len(dashboard_productos)}")
print(f"Top 80%: {(dashboard_productos['top_80_flag'] == 'SI').sum()}")
print(f"Ventas no vinculadas: {len(ventas_no_vinculadas)}")
print(f"Stock no vinculado SKUs: {stock_no_vinculado_skus}")
print(f"Stock congelado activo: {USAR_STOCK_CONGELADO}")
print(f"Fecha corte stock: {FECHA_CORTE_STOCK}")
print(f"Archivo stock congelado: {ARCHIVO_STOCK_CONGELADO}")

