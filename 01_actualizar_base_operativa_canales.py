# -*- coding: utf-8 -*-
"""
Dashboard de rotación IQ Tech — base operativa consolidada.

Lógica principal:
- Ventas: Odoo + AutoAzur, sin duplicar por referencia/pedido.
- Los SKU de producto se mantienen separados de referencias de venta, folios e Item ID.
- Diccionario origen4: relaciona aliases de producto con un SKU madre IQ.
- Stock Odoo: ubicaciones CUATI autorizadas y sus sububicaciones; disponible = quantity - reserved_quantity.
- Full base: corte maestro indicado en .env (actual: 20/08/2026 08:30 CDMX).
- Tránsito: todos los traslados que siguen en Listo/assigned hacia destinos Full.
- Full actual: Full congelado + traslados Hecho posteriores al corte - ventas Full posteriores.
- Cancelados: no suman.
- Exporta hojas separadas para SKU sin madre, productos Odoo sin código y referencias AutoAzur no resueltas.

Versión 2026-09:
- Días de inventario por canal = capas FIFO: lote (compra/ajuste, nunca
  devolución) para Odoo, fecha de envío para Full, respaldo FIFO de compras
  para lo demás. Hoja 'antiguedad_capas' para auditar cada SKU.
- Costo unitario y valor (inversión) por SKU madre y canal, también para SKU
  que hoy solo tienen piezas en Full.
- Registro acumulativo de movimientos Odoo (registro_movimientos_odoo.csv)
  con banderas de revisión para detectar traslados que maquillan antigüedad.
"""

from pathlib import Path
import os
import re
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

VERSION_CODIGO = "2026-09-25_ANTIGUEDAD_FIFO_VALUACION_AUDITORIA"


# ============================================================
# CONFIGURACIÓN GENERAL
# ============================================================

# Cambia únicamente esta ruta si mueves la carpeta del proyecto.
# CORRECCIÓN: la versión anterior leía os.getenv("API Odoo"), una variable que
# no existe; por eso la carpeta siempre caía en la del script aunque .env
# definiera DASHBOARD_ROTACION_DIR como el resto de los scripts (00, 02, 03).
CARPETA_SALIDA = Path(os.getenv(
    "DASHBOARD_ROTACION_DIR",
    str(Path(__file__).resolve().parent)
))
DESKTOP = CARPETA_SALIDA
CARPETA_SALIDA.mkdir(parents=True, exist_ok=True)

ARCHIVO_SALIDA = CARPETA_SALIDA / "rotacion_inventario_base_dashboard_odoo_autoazur.xlsx"

# Nombre opcional del archivo AutoAzur. Si queda vacío, se toma el archivo
# VentasAZ*.xlsx más reciente de la carpeta, Desktop o Downloads.
ARCHIVO_AUTOAZUR_ACTUALIZADO = os.getenv("AUTOAZUR_FILENAME", "").strip()

FECHA_INICIO = pd.Timestamp("2025-11-01")
FECHA_FIN = pd.Timestamp.today().normalize()

# CORTE MAESTRO FULL.
# 00_calcular_stock_inicial.py mantiene estas tres variables en .env.
# Hora local = Ciudad de México; Odoo registra las fechas técnicas en UTC.
FECHA_CORTE_FULL = pd.Timestamp(
    os.getenv("FECHA_CORTE_FULL", "2026-08-20 08:30:00").strip()
)
FECHA_CORTE_FULL_ODOO_UTC = pd.Timestamp(
    os.getenv("FECHA_CORTE_FULL_ODOO_UTC", "2026-08-20 14:30:00").strip()
)
FECHA_INICIO_REGLA_TRASLADOS = FECHA_CORTE_FULL_ODOO_UTC
ARCHIVO_STOCK_FULL_CONGELADO = os.getenv(
    "FULL_STOCK_FILENAME",
    "stock_full_corte_actual.xlsx"
).strip()

# ============================================================
# CONFIGURACIÓN ODOO
# ============================================================

ENABLE_ODOO = True
ODOO_URL = os.getenv("ODOO_URL", "").strip()
ODOO_DB = os.getenv("ODOO_DB", "").strip()
ODOO_USER = os.getenv("ODOO_USER", "").strip()
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "").strip()
ODOO_CAMPO_TIPO_VENTA = os.getenv("ODOO_CAMPO_TIPO_VENTA", "x_studio_tipo_de_venta")
ODOO_TIPOS_VENTA_VALIDOS = ["full", "drop"]

# available = quantity - reserved_quantity. Es la opción correcta para no
# duplicar las piezas reservadas en un traslado con estado Listo.
ODOO_STOCK_MODE = "available"

# Ubicaciones que sí forman parte del inventario comercial por canal.
ODOO_UBICACIONES_VALIDAS = {
    "cuati/existencias": ("General", "ODOO_GENERAL"),
    "cuati/amazon": ("Amazon", "ODOO_AMAZON"),
    "cuati/mercadolibre": ("Mercado Libre", "ODOO_MERCADO_LIBRE"),
    "cuati/walmart": ("Walmart", "ODOO_WALMART"),
    "cuati/liverpool": ("Liverpool", "ODOO_LIVERPOOL"),
    "cuati/coppel": ("Coppel", "ODOO_COPPEL"),
    "cuati/elektra": ("Elektra", "ODOO_ELEKTRA"),
    "cuati/tiktok": ("TikTok", "ODOO_TIKTOK"),
    # B2B existe en Odoo con stock real y antes quedaba fuera de la base.
    "cuati/b2b": ("B2B", "ODOO_B2B"),
}

# Raíz de almacén que se vigila para detectar canales nuevos creados en Odoo
# sin avisar. Todo lo que cuelgue de aquí y no esté arriba se reporta.
ODOO_RAIZ_VIGILADA = "cuati"

# ------------------------------------------------------------------
# BANDERAS DE COBERTURA (añadidas en la revisión de cobertura total)
# ------------------------------------------------------------------
# Estados de picking que se consideran mercancía en camino a Full.
# 'assigned' (Listo) es el histórico. 'confirmed'/'waiting' son traslados
# creados pero sin reserva: existen, pero antes eran invisibles.
ESTADOS_TRANSITO_FULL = ["assigned", "confirmed", "waiting"]

# Si es True, los traslados sin reserva suman al inventario en tránsito.
# Déjalo en False para conservar exactamente la métrica histórica; la
# información se exporta igual en la hoja de detalle para que la revises.
INCLUIR_TRANSITO_SIN_RESERVA = False

# Prefijo para productos de Odoo sin código interno ni código de barras.
# Antes se descartaban de la base; ahora entran con una llave sintética
# para que su stock y su costo no desaparezcan del inventario.
SKU_FALLBACK_PREFIJO = "ODOO-PID-"

# Si es True, un SKU sin alias en el diccionario origen4 se convierte en su
# propio SKU madre en vez de quedar excluido de las métricas. Sigue
# apareciendo en las hojas de excepción para que lo cures a mano.
AUTOGENERAR_SKU_MADRE = True

# ------------------------------------------------------------------
# ANTIGÜEDAD, VALUACIÓN Y AUDITORÍA DE MOVIMIENTOS (versión 2026-09)
# ------------------------------------------------------------------
# Días hacia atrás para reconstruir las capas FIFO de Full: se leen los
# envíos Odoo -> Full ya "Hecho" y se asume que lo que hoy queda en Full son
# las piezas de los envíos más recientes (Full vende primero lo más viejo).
ANTIGUEDAD_LOOKBACK_ENVIOS_FULL_DIAS = int(os.getenv("ANTIGUEDAD_LOOKBACK_ENVIOS_FULL_DIAS", "540"))

# Días hacia atrás de recepciones de compra (proveedor -> bodega). Se usan
# como respaldo FIFO cuando una pieza no tiene lote o su lote solo entró por
# devolución, y para la "antigüedad desde compra" a nivel SKU.
ANTIGUEDAD_LOOKBACK_COMPRAS_DIAS = int(os.getenv("ANTIGUEDAD_LOOKBACK_COMPRAS_DIAS", "900"))

# Umbral de antigüedad que se vigila en las métricas de los KAMs.
ANTIGUEDAD_UMBRAL_ALERTA_DIAS = int(os.getenv("ANTIGUEDAD_UMBRAL_ALERTA_DIAS", "90"))

# Orígenes de un lote que SÍ cuentan como ingreso real a bodega. Una
# devolución de cliente (customer -> interno) o un tránsito ya NO reinician
# la antigüedad: ese era el error que hacía que, por ejemplo, IQ1187 en
# Mercado Libre mostrara la fecha de una devolución y no la del envío.
# El número es la prioridad: primero compra/producción, luego ajuste.
ORIGENES_RECEPCION_LOTE = {
    "supplier": ("COMPRA", 1),
    "production": ("PRODUCCION", 1),
    "inventory": ("AJUSTE_INVENTARIO", 2),
}

# Auditoría de movimientos Odoo (registro anti-maquillaje de antigüedad).
AUDITORIA_MOVIMIENTOS_DIAS = int(os.getenv("AUDITORIA_MOVIMIENTOS_DIAS", "120"))
AUDITORIA_UMBRAL_EDAD_SOSPECHA = int(os.getenv("AUDITORIA_UMBRAL_EDAD_SOSPECHA", "75"))
AUDITORIA_VENTANA_IDA_VUELTA_DIAS = int(os.getenv("AUDITORIA_VENTANA_IDA_VUELTA_DIAS", "30"))
AUDITORIA_DIAS_EN_DASHBOARD = int(os.getenv("AUDITORIA_DIAS_EN_DASHBOARD", "180"))
ARCHIVO_REGISTRO_MOVIMIENTOS_NOMBRE = os.getenv(
    "ARCHIVO_REGISTRO_MOVIMIENTOS", "registro_movimientos_odoo.csv"
).strip()

# Palabras para pre-filtrar en el servidor los contactos que pueden ser Full.
# Después se aplica regla_canal_full(), que es la que decide.
PALABRAS_CONTACTO_FULL = ("full", "fba", "wfs", "amz", "amazon", "walmart", "99min")

# Sububicaciones operativas de CUATI que no son un canal (se ignoran en la
# auditoría porque su flujo es logístico, no comercial).
UBICACIONES_OPERATIVAS = {
    "entrada", "salida", "input", "output", "stock", "control de calidad",
    "quality control", "zona de empaquetado", "packing zone",
    "preproduccion", "posproduccion", "no apto",
}

# Canales comerciales cuyo KAM tiene métrica de antigüedad. Mover inventario
# desde uno de estos hacia General/B2B/no comercial se revisa en la auditoría.
CANALES_CON_KAM = {
    "Amazon", "Mercado Libre", "Walmart", "Liverpool", "Coppel", "Elektra", "TikTok",
}

# Contactos que identifican una salida hacia Full.
CONTACTOS_FULL = {
    "mercado libre full": "Mercado Libre",
    "amazon fba": "Amazon",
    "amazon": "Amazon",
    "amazon fba, amazon": "Amazon",
    "walmart wfs": "Walmart",
    "liverpool full 99min": "Liverpool",
    "distribuidora liverpool, s.a. de c.v., liverpool full": "Liverpool",
    "coppel full": "Coppel",
}


# Reglas adicionales de contactos Full capturados por los KAMs.
# IMPORTANTE: estas reglas son deliberadamente estrictas con los separadores.
# Amazon acepta solo contactos que EMPIECEN exactamente con "AMZ | FBA19"
# (ignorando mayúsculas/minúsculas, pero NO corrigiendo espacios internos).
# Ejemplo válido: "AMZ | FBA19KXPXQ3D | SONY BLACK"
# Ejemplos NO válidos: "AMZ|FBA19...", "AMZ  | FBA19...", "AMZ |  FBA19..."
# Walmart acepta "Walmart | ..." salvo contactos que indiquen Venta diaria(s).
AMAZON_FULL_PREFIX_ESTRICTO = "amz | fba19"
WALMART_FULL_PREFIX_ESTRICTO = "walmart | "
WALMART_CONTACTOS_EXCLUIR = ("venta diaria", "ventas diarias")

# Frases que descalifican un contacto como Full en CUALQUIER canal.
# Antes esta protección solo existía para Walmart, lo que dejaba a Amazon
# expuesto a contar ventas diarias como reabasto a FBA.
CONTACTOS_EXCLUIR_GLOBAL = (
    "venta diaria", "ventas diarias", "devolucion", "devoluciones",
    "retiro", "retiros", "muestra", "garantia",
)

# Prefijos tolerantes: aceptan que el contacto traiga información extra
# después (folio de envío, producto, plaza). Se evalúan sobre el texto
# NORMALIZADO, así que sí perdonan acentos y espacios dobles.
#
# Mercado Libre no tenía ninguna regla de prefijo: solo entraba con la
# coincidencia exacta "mercado libre full", de modo que cualquier captura
# tipo "Mercado Libre Full | ENV123" se perdía silenciosamente. Este es el
# hueco principal que cierra esta versión.
PREFIJOS_FULL_NORMALIZADOS = (
    ("mercado libre full", "Mercado Libre"),
    ("mercadolibre full", "Mercado Libre"),
    ("meli full", "Mercado Libre"),
    ("ml full", "Mercado Libre"),
    ("amazon fba", "Amazon"),
    ("liverpool full", "Liverpool"),
    ("coppel full", "Coppel"),
    ("elektra full", "Elektra"),
)


def regla_canal_full(contacto):
    """Devuelve (canal, regla_aplicada, motivo_rechazo) para un contacto Odoo.

    Orden de evaluación:
      1. Exclusión global (venta diaria, devolución, retiro...).
      2. Diccionario de contactos exactos (histórico, no se toca).
      3. Prefijos estrictos en texto crudo: Amazon ``AMZ | FBA19`` y
         Walmart ``Walmart | ``. Se conservan tal cual para no cambiar la
         metodología: si el KAM usó otro separador, sigue quedando fuera.
      4. Prefijos tolerantes sobre texto normalizado (Mercado Libre y demás).

    Devolver la regla permite auditar después por qué entró cada traslado.
    """
    contacto_raw = str(contacto or "").strip()
    if not contacto_raw:
        return "", "", "CONTACTO_VACIO"

    contacto_lower = contacto_raw.lower()
    contacto_norm = normalizar_texto(contacto_raw)

    # 1) Exclusiones que aplican a todos los canales.
    for frase in CONTACTOS_EXCLUIR_GLOBAL:
        if frase in contacto_norm:
            return "", "", f"EXCLUIDO_POR_FRASE:{frase}"

    # 2) Contactos exactos históricos.
    if contacto_norm in CONTACTOS_FULL:
        return CONTACTOS_FULL[contacto_norm], "EXACTO", ""

    # 3) Prefijos estrictos históricos (sobre texto crudo, a propósito).
    if contacto_lower.startswith(AMAZON_FULL_PREFIX_ESTRICTO):
        return "Amazon", "PREFIJO_ESTRICTO_AMAZON", ""

    if contacto_lower.startswith(WALMART_FULL_PREFIX_ESTRICTO):
        if any(f in contacto_lower for f in WALMART_CONTACTOS_EXCLUIR):
            return "", "", "WALMART_VENTA_DIARIA"
        return "Walmart", "PREFIJO_ESTRICTO_WALMART", ""

    # 4) Prefijos tolerantes (cierra el hueco de Mercado Libre Full).
    for prefijo, canal in PREFIJOS_FULL_NORMALIZADOS:
        if contacto_norm.startswith(prefijo):
            return canal, f"PREFIJO_NORMALIZADO:{prefijo}", ""

    return "", "", "CONTACTO_NO_FULL"


def canal_full_por_contacto(contacto):
    """Compatibilidad: devuelve solo el canal Full, o cadena vacía."""
    return regla_canal_full(contacto)[0]

CANALES_COMERCIALES = [
    "Amazon", "Mercado Libre", "Walmart", "Liverpool",
    "Coppel", "Elektra", "TikTok"
]


def clasificar_ubicacion_odoo(ubicacion):
    """Clasifica una ubicación raíz CUATI o cualquiera de sus sububicaciones."""
    norm = normalizar_texto(ubicacion)
    for raiz in sorted(ODOO_UBICACIONES_VALIDAS, key=len, reverse=True):
        if norm == raiz or norm.startswith(raiz + "/"):
            canal, canal_stock = ODOO_UBICACIONES_VALIDAS[raiz]
            return raiz, canal, canal_stock
    return "", "", ""


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


def obtener_api_key_odoo():
    """Lee la API key únicamente desde ODOO_API_KEY; nunca solicita datos en Terminal."""
    key = str(ODOO_API_KEY or "").strip()
    if not key:
        raise ValueError(
            "Falta ODOO_API_KEY en el archivo .env. "
            "Ejecuta ABRIR_DASHBOARD.command para configurarla mediante una ventana segura."
        )
    return key


def normalizar_canal_venta(valor):
    texto = normalizar_texto(valor)
    if "amazon" in texto:
        return "Amazon"
    if "mercado libre" in texto or "mercadolibre" in texto or texto in {"meli", "ml"}:
        return "Mercado Libre"
    if "walmart" in texto:
        return "Walmart"
    if "liverpool" in texto:
        return "Liverpool"
    if "coppel" in texto:
        return "Coppel"
    if "elektra" in texto:
        return "Elektra"
    if "tiktok" in texto or "tik tok" in texto:
        return "TikTok"
    return "Sin identificar"


def normalizar_modalidad_venta(valor, fuente=""):
    texto = normalizar_texto(valor)
    if texto == "full" or " full" in f" {texto}":
        return "FULL"
    if texto == "drop" or " drop" in f" {texto}":
        return "DROP"
    if normalizar_texto(fuente) == "autoazur":
        return "PENDIENTE_ODOO"
    return "SIN_IDENTIFICAR"


def estado_picking_dashboard(estado):
    mapa = {
        "assigned": "Listo",
        "confirmed": "En espera",
        "waiting": "En espera de otra operación",
        "done": "Hecho",
        "cancel": "Cancelado",
    }
    return mapa.get(str(estado or "").strip().lower(), str(estado or ""))


def campos_disponibles_odoo(odoo, modelo):
    """Obtiene los campos reales del modelo para tolerar diferencias entre versiones de Odoo."""
    try:
        data = odoo.execute(modelo, "fields_get", [], attributes=["type", "string", "relation"])
        return set(data.keys())
    except Exception:
        return set()


def seleccionar_campos(disponibles, candidatos):
    return [c for c in candidatos if not disponibles or c in disponibles]


# ============================================================
# METADATOS DE PRODUCTO PARA REPARTICIÓN
# ============================================================
# La categoría ya existe en Odoo mediante categ_id. La marca puede ser un
# campo estándar de un módulo adicional o un campo Studio/custom. Se detecta
# de forma tolerante para que el pipeline siga funcionando aunque no exista.
CANDIDATOS_CAMPO_MARCA_ODOO = [
    "product_brand_id", "brand_id", "x_studio_marca", "x_marca",
    "product_brand", "marca", "brand",
]


def detectar_campo_marca_odoo(disponibles):
    disponibles = set(disponibles or [])
    for campo in CANDIDATOS_CAMPO_MARCA_ODOO:
        if campo in disponibles:
            return campo
    return ""


def valor_texto_odoo(value):
    """Convierte many2one/selection/texto de Odoo a una etiqueta legible."""
    if isinstance(value, (list, tuple)):
        return m2o_name(value)
    if value in (None, False):
        return ""
    return str(value).strip()


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
    """Busca el archivo AutoAzur indicado o el VentasAZ*.xlsx más reciente."""
    carpetas = [CARPETA_SALIDA, Path.home() / "Desktop", Path.home() / "Downloads"]

    if ARCHIVO_AUTOAZUR_ACTUALIZADO:
        for carpeta in carpetas:
            ruta = carpeta / ARCHIVO_AUTOAZUR_ACTUALIZADO
            if ruta.exists():
                return ruta

    candidatos = []
    for carpeta in carpetas:
        if not carpeta.exists():
            continue
        candidatos.extend(carpeta.glob("VentasAZ*.xlsx"))
        candidatos.extend(carpeta.glob("*listado*detallado*pedidos*.xlsx"))

    candidatos = [p for p in candidatos if p.is_file() and not p.name.startswith("~$")]
    if not candidatos:
        return None

    return max(candidatos, key=lambda p: p.stat().st_mtime)


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

    campos_producto_disponibles = campos_disponibles_odoo(odoo, "product.product")
    campo_marca_odoo = detectar_campo_marca_odoo(campos_producto_disponibles)
    product_fields = seleccionar_campos(
        campos_producto_disponibles,
        ["id", "display_name", "default_code", "barcode", "categ_id"]
        + ([campo_marca_odoo] if campo_marca_odoo else [])
    )

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
            "canal_venta": normalizar_canal_venta(m2o_name(o.get("team_id"))),
            "modalidad_venta": normalizar_modalidad_venta(tipo_venta, "ODOO_API"),
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
            "marca": valor_texto_odoo(p.get(campo_marca_odoo)) if campo_marca_odoo else "",
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
    """Localiza las ocho raíces CUATI autorizadas y todas sus sububicaciones."""
    ubicaciones = odoo.search_read_all(
        "stock.location",
        [("usage", "=", "internal")],
        ["id", "name", "complete_name", "usage"],
        batch=3000,
        order="complete_name asc"
    )
    ubicaciones_df = pd.DataFrame(ubicaciones)
    if ubicaciones_df.empty:
        return ubicaciones_df, []

    ubicaciones_df["ubicacion_norm"] = ubicaciones_df["complete_name"].apply(normalizar_texto)
    clasificaciones = ubicaciones_df["complete_name"].apply(clasificar_ubicacion_odoo)
    ubicaciones_df["raiz_autorizada"] = clasificaciones.apply(lambda x: x[0])
    ubicaciones_df["canal_asignado"] = clasificaciones.apply(lambda x: x[1])
    ubicaciones_df["canal_stock"] = clasificaciones.apply(lambda x: x[2])

    filtradas = ubicaciones_df[ubicaciones_df["raiz_autorizada"].ne("")].copy()
    raices_encontradas = set(filtradas["raiz_autorizada"])
    faltantes = sorted(set(ODOO_UBICACIONES_VALIDAS) - raices_encontradas)
    if faltantes:
        print("ADVERTENCIA: no se encontraron estas raíces de ubicación en Odoo:")
        for nombre in faltantes:
            print(f"- {nombre}")

    # Canales nuevos: ubicaciones que cuelgan del almacén vigilado pero que
    # nadie dio de alta en ODOO_UBICACIONES_VALIDAS. Antes su stock se perdía
    # en silencio; ahora se avisa para que se agreguen a la configuración.
    prefijo = ODOO_RAIZ_VIGILADA + "/"
    huerfanas = ubicaciones_df[
        ubicaciones_df["raiz_autorizada"].eq("")
        & ubicaciones_df["ubicacion_norm"].str.startswith(prefijo)
    ].copy()
    if not huerfanas.empty:
        print("ADVERTENCIA: ubicaciones de CUATI sin canal asignado "
              "(su stock NO entra a las métricas):")
        for _, r in huerfanas.iterrows():
            print(f"- ID {r['id']} | {r['complete_name']}")
        print("  Agrégalas a ODOO_UBICACIONES_VALIDAS si deben contar.")

    print("Ubicaciones Odoo usadas para inventario:")
    for _, r in filtradas.iterrows():
        print(f"- ID {r['id']} | {r['complete_name']} | {r['canal_asignado']}")

    return filtradas, filtradas["id"].astype(int).tolist()


def extraer_fecha_recepcion_lotes(odoo, lote_ids):
    """Fecha real de ingreso a bodega de cada lote de Odoo.

    Regla corregida (2026-09):
    - Solo cuentan como ingreso los movimientos "Hecho" que entran a una
      ubicación interna desde un proveedor (Compras) o producción; en
      segundo término, desde un ajuste de inventario (carga inicial).
    - Una devolución de cliente (customer -> interno) o un tránsito NO son
      ingreso: antes sí contaban, y como se tomaba la fecha más antigua de
      cualquier entrada, un lote creado al recibir una devolución de Full
      aparecía con la fecha de la devolución (caso IQ1187 Mercado Libre).
    - Si el lote nunca tuvo una entrada de compra/producción/ajuste, se deja
      fecha_recepcion_lote vacía y se reporta la fecha de su primer
      movimiento para que el cálculo de antigüedad use el respaldo FIFO.

    Devuelve: lote_id_odoo, fecha_recepcion_lote, origen_recepcion_lote,
    fuente_recepcion_lote, fecha_primer_movimiento_lote.
    """
    columnas = [
        "lote_id_odoo", "fecha_recepcion_lote", "origen_recepcion_lote",
        "fuente_recepcion_lote", "fecha_primer_movimiento_lote",
    ]
    lote_ids = sorted({int(x) for x in (lote_ids or []) if x and not pd.isna(x)})
    if not lote_ids:
        return pd.DataFrame(columns=columnas)

    ml_avail = campos_disponibles_odoo(odoo, "stock.move.line")
    if ml_avail and "lot_id" not in ml_avail:
        print("ADVERTENCIA: stock.move.line no expone lot_id; no se puede calcular antigüedad por lote.")
        return pd.DataFrame(columns=columnas)

    campo_fecha = "date" if (not ml_avail or "date" in ml_avail) else (
        "date_done" if "date_done" in ml_avail else None
    )
    if not campo_fecha:
        print("ADVERTENCIA: stock.move.line no expone fecha; no se puede calcular antigüedad por lote.")
        return pd.DataFrame(columns=columnas)
    campos_ml = ["id", "lot_id", "location_id", "location_dest_id", "picking_id", campo_fecha]

    movimientos = []
    # Se consulta por bloques para no mandar dominios gigantes a Odoo.
    for i in range(0, len(lote_ids), 2000):
        bloque = lote_ids[i:i + 2000]
        movimientos.extend(odoo.search_read_all(
            "stock.move.line",
            [("lot_id", "in", bloque), ("state", "=", "done")],
            campos_ml,
            batch=5000,
            order="id asc",
        ))
    if not movimientos:
        return pd.DataFrame(columns=columnas)

    location_ids = sorted({
        loc_id
        for m in movimientos
        for loc_id in (m2o_id(m.get("location_id")), m2o_id(m.get("location_dest_id")))
        if loc_id
    })
    usage_map = {}
    for i in range(0, len(location_ids), 1000):
        for u in odoo.execute(
            "stock.location", "read", location_ids[i:i + 1000], ["id", "usage"],
            context={"active_test": False},
        ):
            usage_map[u["id"]] = u.get("usage", "")

    picking_ids = sorted({m2o_id(m.get("picking_id")) for m in movimientos if m2o_id(m.get("picking_id"))})
    picking_map = {}
    for i in range(0, len(picking_ids), 1000):
        for p in odoo.execute("stock.picking", "read", picking_ids[i:i + 1000], ["id", "name", "origin"]):
            picking_map[p["id"]] = p

    primer_mov = {}
    candidatos = []
    for m in movimientos:
        lote_id = m2o_id(m.get("lot_id"))
        if not lote_id:
            continue
        fecha_dt = pd.to_datetime(m.get(campo_fecha), errors="coerce")
        if pd.isna(fecha_dt):
            continue
        if lote_id not in primer_mov or fecha_dt < primer_mov[lote_id]:
            primer_mov[lote_id] = fecha_dt

        origen_usage = usage_map.get(m2o_id(m.get("location_id")), "")
        destino_usage = usage_map.get(m2o_id(m.get("location_dest_id")), "")
        if destino_usage != "internal" or origen_usage not in ORIGENES_RECEPCION_LOTE:
            continue
        fuente, prioridad = ORIGENES_RECEPCION_LOTE[origen_usage]
        picking = picking_map.get(m2o_id(m.get("picking_id")), {})
        candidatos.append({
            "lote_id_odoo": lote_id,
            "fecha_recepcion_lote": fecha_dt,
            "origen_recepcion_lote": picking.get("origin") or picking.get("name") or "",
            "fuente_recepcion_lote": fuente,
            "_prioridad": prioridad,
        })

    base = pd.DataFrame({
        "lote_id_odoo": list(primer_mov.keys()),
        "fecha_primer_movimiento_lote": list(primer_mov.values()),
    })
    if candidatos:
        cand = pd.DataFrame(candidatos).sort_values(["lote_id_odoo", "_prioridad", "fecha_recepcion_lote"])
        mejor = cand.drop_duplicates("lote_id_odoo", keep="first").drop(columns=["_prioridad"])
        base = base.merge(mejor, on="lote_id_odoo", how="left")
    for c in ["fecha_recepcion_lote"]:
        if c not in base.columns:
            base[c] = pd.NaT
    for c in ["origen_recepcion_lote", "fuente_recepcion_lote"]:
        if c not in base.columns:
            base[c] = ""
        base[c] = base[c].fillna("")
    base.loc[base["fecha_recepcion_lote"].isna(), "fuente_recepcion_lote"] = "SIN_ENTRADA_DE_COMPRA"

    resueltos = int(base["fecha_recepcion_lote"].notna().sum())
    print(
        f"Fechas de recepción por lote: {resueltos} con compra/ajuste, "
        f"{len(base) - resueltos} solo con devolución/otro, de {len(lote_ids)} lotes consultados."
    )
    return base[columnas]


def extraer_inventario_odoo(odoo):
    print("Descargando inventario Odoo por canal...")
    ubicaciones_df, location_ids = buscar_ubicaciones_odoo(odoo)

    columnas_sin_sku = [
        "odoo_quant_id", "odoo_product_id", "producto_stock", "odoo_location",
        "canal_asignado", "quantity_odoo", "reserved_quantity_odoo",
        "stock_original", "stock", "motivo"
    ]
    if not location_ids:
        return pd.DataFrame(), ubicaciones_df, pd.DataFrame(columns=columnas_sin_sku)

    # Se intenta traer "value" (costo total en ubicación, tal como se ve en
    # Odoo -> Inventario -> Reporte -> Ubicaciones). Si la valuación de
    # inventario no está activa en esta instancia, el campo puede no existir
    # o venir vacío; en ese caso se usa como respaldo el costo estándar del
    # producto (standard_price) multiplicado por la cantidad del quant.
    campos_quant = ["id", "product_id", "location_id", "quantity", "reserved_quantity"]
    campos_quant_disponibles = campos_disponibles_odoo(odoo, "stock.quant")
    tiene_value_quant = "value" in campos_quant_disponibles
    if tiene_value_quant:
        campos_quant.append("value")
    tiene_lote_quant = "lot_id" in campos_quant_disponibles
    if tiene_lote_quant:
        campos_quant.append("lot_id")

    quants = odoo.search_read_all(
        "stock.quant",
        [("location_id", "in", location_ids), ("product_id", "!=", False)],
        campos_quant,
        batch=5000,
        order="id asc"
    )
    if not quants:
        return pd.DataFrame(), ubicaciones_df, pd.DataFrame(columns=columnas_sin_sku)

    product_ids = sorted({m2o_id(q.get("product_id")) for q in quants if m2o_id(q.get("product_id"))})
    campos_producto_base = ["id", "display_name", "default_code", "barcode", "categ_id"]
    campos_producto_disponibles = campos_disponibles_odoo(odoo, "product.product")
    campo_marca_odoo = detectar_campo_marca_odoo(campos_producto_disponibles)
    candidatos_producto = campos_producto_base + ["standard_price"]
    if campo_marca_odoo:
        candidatos_producto.append(campo_marca_odoo)
    campos_producto = seleccionar_campos(
        campos_producto_disponibles, candidatos_producto
    ) if campos_producto_disponibles else campos_producto_base
    products = odoo.execute(
        "product.product", "read", product_ids, campos_producto
    ) if product_ids else []
    product_map = {p["id"]: p for p in products}
    loc_map = {int(r["id"]): r for _, r in ubicaciones_df.iterrows()}

    rows, rows_sin_sku = [], []
    for q in quants:
        product_id = m2o_id(q.get("product_id"))
        location_id = m2o_id(q.get("location_id"))
        p = product_map.get(product_id, {})
        loc = loc_map.get(location_id, {})

        qty = float(q.get("quantity") or 0)
        reserved = float(q.get("reserved_quantity") or 0)
        stock_original = qty - reserved if ODOO_STOCK_MODE == "available" else qty
        stock_metricas = max(stock_original, 0)

        default_code = limpiar_sku(p.get("default_code", ""))
        barcode = limpiar_sku(p.get("barcode", ""))
        sku = default_code or barcode

        # Antes, un producto sin código interno ni código de barras salía de
        # la base por completo: su stock y su costo simplemente desaparecían
        # del inventario. Ahora se le asigna una llave sintética estable
        # (ODOO-PID-<id>) para que el valor cuadre, y se sigue reportando en
        # la hoja de excepciones para que alguien le capture el SKU real.
        if sku:
            origen_sku = "ODOO_DEFAULT_CODE" if default_code else "ODOO_BARCODE"
            sku_es_generado = "NO"
        else:
            sku = f"{SKU_FALLBACK_PREFIJO}{product_id}"
            origen_sku = "GENERADO_POR_ID"
            sku_es_generado = "SI"

        # Costo total del quant. Fuente principal: "value" del stock.quant
        # (igual a la columna "Costo total en ubicación" de Odoo). Respaldo:
        # standard_price del producto por la cantidad física (quantity),
        # que se usa cuando la valuación automática no está disponible.
        costo_total_quant_odoo = float(q.get("value") or 0) if tiene_value_quant else 0.0
        costo_unitario_producto = float(p.get("standard_price") or 0)
        fuente_costo_quant = ""
        if costo_total_quant_odoo:
            fuente_costo_quant = "ODOO_QUANT_VALUE"
        elif costo_unitario_producto:
            costo_total_quant_odoo = costo_unitario_producto * qty
            fuente_costo_quant = "PRODUCT_STANDARD_PRICE"

        base = {
            "canal_stock": loc.get("canal_stock", ""),
            "canal_asignado": loc.get("canal_asignado", ""),
            "tipo_inventario": "ODOO",
            "sku_original": sku,
            "producto_stock": p.get("display_name", ""),
            "stock_original": stock_original,
            "stock": stock_metricas,
            "alerta_calidad_stock": "NEGATIVO_AJUSTADO_A_CERO" if stock_original < 0 else "",
            "fuente_archivo": "ODOO_API_STOCK_QUANT",
            "fecha_corte": pd.Timestamp.now(),
            "odoo_quant_id": q.get("id"),
            "odoo_product_id": product_id,
            "odoo_location_id": location_id,
            "odoo_location": m2o_name(q.get("location_id")),
            "quantity_odoo": qty,
            "reserved_quantity_odoo": reserved,
            "categoria_odoo": m2o_name(p.get("categ_id")),
            "marca_odoo": valor_texto_odoo(p.get(campo_marca_odoo)) if campo_marca_odoo else "",
            "costo_total_ubicacion_odoo": costo_total_quant_odoo,
            "costo_unitario_producto_odoo": costo_unitario_producto,
            "fuente_costo_quant": fuente_costo_quant,
            "lote_id_odoo": m2o_id(q.get("lot_id")) if tiene_lote_quant else None,
            "lote_nombre": m2o_name(q.get("lot_id")) if tiene_lote_quant else "",
        }

        base["origen_sku"] = origen_sku
        base["sku_generado"] = sku_es_generado

        # La fila SIEMPRE entra a la base; si el SKU fue generado también se
        # copia a la hoja de excepciones para su corrección manual.
        rows.append(base)
        if sku_es_generado == "SI":
            excepcion = dict(base)
            excepcion["motivo"] = (
                "Producto sin código interno ni código de barras. "
                "Se incluyó en métricas con SKU generado; captúrale el código en Odoo."
            )
            rows_sin_sku.append(excepcion)

    stock_odoo = pd.DataFrame(rows)
    stock_odoo_sin_sku = pd.DataFrame(rows_sin_sku)
    print(f"Renglones inventario Odoo con SKU: {len(stock_odoo)}")
    print(f"Renglones inventario Odoo sin SKU: {len(stock_odoo_sin_sku)}")
    if not tiene_value_quant:
        print(
            "ADVERTENCIA: stock.quant no expone el campo 'value' en esta instancia; "
            "el costo unitario CUATI se calculará con standard_price del producto."
        )

    # ------------------------------------------------------------------
    # ANTIGÜEDAD POR LOTE: fecha real de recepción en Compras (Odoo).
    # Permite calcular "Días de inventario" a partir de cuándo entró
    # físicamente cada lote, en vez de cuántos días dura el stock al
    # ritmo de venta actual (eso ya existe como cobertura/ritmo aparte).
    # ------------------------------------------------------------------
    if not stock_odoo.empty and tiene_lote_quant:
        lote_ids = sorted({
            int(v) for v in stock_odoo["lote_id_odoo"].dropna().tolist() if v
        })
        fechas_lote = extraer_fecha_recepcion_lotes(odoo, lote_ids)
        if not fechas_lote.empty:
            stock_odoo = stock_odoo.merge(fechas_lote, on="lote_id_odoo", how="left")
        for c in ["fecha_recepcion_lote", "fecha_primer_movimiento_lote"]:
            if c not in stock_odoo.columns:
                stock_odoo[c] = pd.NaT
        for c in ["origen_recepcion_lote", "fuente_recepcion_lote"]:
            if c not in stock_odoo.columns:
                stock_odoo[c] = ""
            stock_odoo[c] = stock_odoo[c].fillna("")
        hoy = pd.Timestamp.now().normalize()
        stock_odoo["dias_en_stock_lote"] = np.where(
            stock_odoo["fecha_recepcion_lote"].notna(),
            (hoy - stock_odoo["fecha_recepcion_lote"]).dt.days,
            np.nan
        )
    else:
        if not stock_odoo.empty:
            print("ADVERTENCIA: stock.quant no expone lot_id en esta instancia; 'Días de inventario' usará el respaldo aproximado por SKU.")
        stock_odoo["fecha_recepcion_lote"] = pd.NaT
        stock_odoo["origen_recepcion_lote"] = ""
        stock_odoo["fuente_recepcion_lote"] = ""
        stock_odoo["fecha_primer_movimiento_lote"] = pd.NaT
        stock_odoo["dias_en_stock_lote"] = np.nan

    return stock_odoo, ubicaciones_df, stock_odoo_sin_sku


def extraer_auditoria_stock_todas_ubicaciones(odoo):
    """Descarga todos los quants de ubicaciones internas solo para auditoría; no altera métricas."""
    # active_test=False incluye ubicaciones ARCHIVADAS. Odoo las oculta de la
    # interfaz pero sus quants siguen vivos: ahí apareció el stock atrapado en
    # el tránsito inter-almacén que no salía en ningún reporte.
    ubicaciones = odoo.execute(
        "stock.location", "search_read",
        [("usage", "in", ["internal", "transit"])],
        fields=["id", "name", "complete_name", "usage", "active"],
        order="complete_name asc",
        context={"active_test": False},
    )
    ubicaciones_df = pd.DataFrame(ubicaciones)
    if ubicaciones_df.empty:
        return pd.DataFrame(), pd.DataFrame(), ubicaciones_df

    archivadas = ubicaciones_df[~ubicaciones_df["active"].astype(bool)]
    if not archivadas.empty:
        print(f"AVISO: {len(archivadas)} ubicaciones archivadas incluidas en la auditoría.")

    location_ids = ubicaciones_df["id"].astype(int).tolist()
    quants = odoo.search_read_all(
        "stock.quant",
        [("location_id", "in", location_ids), ("product_id", "!=", False)],
        ["id", "product_id", "location_id", "quantity", "reserved_quantity"],
        batch=8000, order="id asc"
    )
    if not quants:
        return pd.DataFrame(), pd.DataFrame(), ubicaciones_df

    product_ids = sorted({m2o_id(q.get("product_id")) for q in quants if m2o_id(q.get("product_id"))})
    products = odoo.execute(
        "product.product", "read", product_ids,
        ["id", "display_name", "default_code", "barcode", "categ_id"]
    ) if product_ids else []
    product_map = {p["id"]: p for p in products}
    loc_map = {int(r["id"]): r for r in ubicaciones}

    rows = []
    for q in quants:
        pid = m2o_id(q.get("product_id"))
        lid = m2o_id(q.get("location_id"))
        p = product_map.get(pid, {})
        loc = loc_map.get(lid, {})
        qty = float(q.get("quantity") or 0)
        reserved = float(q.get("reserved_quantity") or 0)
        available = qty - reserved
        sku = limpiar_sku(p.get("default_code", "")) or limpiar_sku(p.get("barcode", ""))
        raiz, canal, canal_stock = clasificar_ubicacion_odoo(loc.get("complete_name", ""))
        rows.append({
            "odoo_quant_id": q.get("id"), "odoo_product_id": pid,
            "sku_original": sku, "producto": p.get("display_name", ""),
            "categoria": m2o_name(p.get("categ_id")),
            "odoo_location_id": lid, "odoo_location": loc.get("complete_name", ""),
            "tipo_ubicacion": loc.get("usage", ""),
            "ubicacion_archivada": "SI" if not loc.get("active", True) else "NO",
            "quantity_odoo": qty, "reserved_quantity_odoo": reserved,
            "stock_disponible": max(available, 0), "stock_disponible_original": available,
            "raiz_autorizada": raiz, "canal_asignado": canal, "canal_stock": canal_stock,
            "incluido_metricas": "SI" if raiz else "NO",
        })
    detalle = pd.DataFrame(rows)
    resumen = detalle.groupby(
        ["odoo_location", "tipo_ubicacion", "ubicacion_archivada",
         "incluido_metricas", "canal_asignado"],
        as_index=False, dropna=False
    ).agg(
        skus=("sku_original", "nunique"),
        cantidad_fisica=("quantity_odoo", "sum"),
        cantidad_reservada=("reserved_quantity_odoo", "sum"),
        stock_disponible=("stock_disponible", "sum"),
    ).sort_values("stock_disponible", ascending=False)
    return detalle, resumen, ubicaciones_df


def extraer_traslados_full_odoo(odoo, ubicaciones_df):
    """Extrae todos los Listo actuales y los Hecho/Cancelado posteriores al corte Full."""
    columnas = [
        "picking_id", "referencia", "contacto", "canal", "regla_full", "ubicacion_origen",
        "ubicacion_destino", "estado_odoo", "estado", "fecha_programada",
        "fecha_efectiva", "fecha_creacion", "ultima_modificacion", "documento_origen",
        "move_id", "product_id", "sku_original", "producto",
        "cantidad_solicitada", "cantidad_realizada", "cantidad_metricas",
        "clasificacion_inventario"
    ]
    if ubicaciones_df.empty:
        return pd.DataFrame(columns=columnas), pd.DataFrame()

    location_ids = ubicaciones_df["id"].astype(int).tolist()
    picking_fields_available = campos_disponibles_odoo(odoo, "stock.picking")
    picking_fields = seleccionar_campos(picking_fields_available, [
        "id", "name", "partner_id", "location_id", "location_dest_id", "state",
        "scheduled_date", "date_done", "create_date", "write_date", "origin"
    ])

    dt_inicio = FECHA_INICIO_REGLA_TRASLADOS.strftime("%Y-%m-%d %H:%M:%S")
    # Traslados vigentes. Antes solo se leían los 'assigned' (Listo), así que
    # un envío a Full creado pero sin reserva ('confirmed'/'waiting') era
    # invisible. Ahora se traen los tres y se distinguen por clasificación.
    pickings_listo = odoo.search_read_all(
        "stock.picking",
        [
            ("state", "in", ESTADOS_TRANSITO_FULL),
            ("location_id", "in", location_ids),
            ("partner_id", "!=", False),
        ],
        picking_fields, batch=3000, order="id asc"
    )
    # Historial reciente: solo Hecho o Cancelado posteriores al corte.
    pickings_hist = odoo.search_read_all(
        "stock.picking",
        [
            "|", ("date_done", ">=", dt_inicio), ("write_date", ">=", dt_inicio),
            ("state", "in", ["done", "cancel"]),
            ("location_id", "in", location_ids),
            ("partner_id", "!=", False),
        ],
        picking_fields, batch=3000, order="id asc"
    )
    pickings_map = {int(p["id"]): p for p in (pickings_listo + pickings_hist)}
    pickings = list(pickings_map.values())

    aceptados, excluidos = [], []
    for p in pickings:
        contacto = m2o_name(p.get("partner_id"))
        contacto_norm = normalizar_texto(contacto)
        origen = m2o_name(p.get("location_id"))
        origen_norm = normalizar_texto(origen)
        estado = str(p.get("state") or "").lower()

        fechas = [
            pd.to_datetime(p.get("date_done"), errors="coerce"),
            pd.to_datetime(p.get("scheduled_date"), errors="coerce"),
            pd.to_datetime(p.get("write_date"), errors="coerce"),
            pd.to_datetime(p.get("create_date"), errors="coerce"),
        ]
        fechas_validas = [f for f in fechas if pd.notna(f)]
        fecha_evento = max(fechas_validas) if fechas_validas else pd.NaT

        canal_full, regla_full, motivo_regla = regla_canal_full(contacto)

        motivo = ""
        if not canal_full:
            motivo = motivo_regla or "CONTACTO_NO_FULL"
        elif not clasificar_ubicacion_odoo(origen)[0]:
            motivo = "ORIGEN_NO_VALIDO"
        elif estado not in ESTADOS_TRANSITO_FULL and (
            pd.isna(fecha_evento) or fecha_evento < FECHA_INICIO_REGLA_TRASLADOS
        ):
            motivo = "ANTERIOR_AL_CORTE_FULL"

        p2 = dict(p)
        p2.update({
            "contacto": contacto,
            "ubicacion_origen": origen,
            "fecha_evento": fecha_evento,
            "motivo_exclusion": motivo,
            "regla_full": regla_full,
        })
        if motivo:
            excluidos.append(p2)
        else:
            p2["canal"] = canal_full
            aceptados.append(p2)

    if not aceptados:
        return pd.DataFrame(columns=columnas), pd.DataFrame(excluidos)

    picking_ids = [int(p["id"]) for p in aceptados]
    picking_map = {int(p["id"]): p for p in aceptados}

    move_avail = campos_disponibles_odoo(odoo, "stock.move")
    move_fields = seleccionar_campos(move_avail, [
        "id", "picking_id", "product_id", "product_uom_qty", "quantity",
        "quantity_done", "state"
    ])
    moves = odoo.search_read_all(
        "stock.move",
        [("picking_id", "in", picking_ids), ("product_id", "!=", False)],
        move_fields,
        batch=5000,
        order="id asc"
    )

    product_ids = sorted({m2o_id(m.get("product_id")) for m in moves if m2o_id(m.get("product_id"))})
    products = odoo.execute(
        "product.product", "read", product_ids,
        ["id", "display_name", "default_code", "barcode"]
    ) if product_ids else []
    product_map = {p["id"]: p for p in products}

    # Cantidad realmente hecha por move; compatible con Odoo antiguo y reciente.
    move_line_done = {}
    ml_avail = campos_disponibles_odoo(odoo, "stock.move.line")
    qty_ml_field = "quantity" if "quantity" in ml_avail else "qty_done" if "qty_done" in ml_avail else None
    if qty_ml_field:
        ml_fields = seleccionar_campos(ml_avail, ["id", "move_id", "picking_id", "product_id", qty_ml_field])
        move_lines = odoo.search_read_all(
            "stock.move.line",
            [("picking_id", "in", picking_ids), ("product_id", "!=", False)],
            ml_fields,
            batch=8000,
            order="id asc"
        )
        for ml in move_lines:
            mid = m2o_id(ml.get("move_id"))
            if mid:
                move_line_done[mid] = move_line_done.get(mid, 0.0) + float(ml.get(qty_ml_field) or 0)

    rows = []
    for m in moves:
        picking_id = m2o_id(m.get("picking_id"))
        p = picking_map.get(picking_id, {})
        estado_odoo = str(p.get("state") or "").lower()
        product_id = m2o_id(m.get("product_id"))
        prod = product_map.get(product_id, {})
        sku = limpiar_sku(prod.get("default_code", "")) or limpiar_sku(prod.get("barcode", ""))

        solicitada = float(m.get("product_uom_qty") or 0)
        realizada = move_line_done.get(m.get("id"), None)
        if realizada is None:
            if "quantity_done" in m:
                realizada = float(m.get("quantity_done") or 0)
            else:
                realizada = float(m.get("quantity") or 0)

        if estado_odoo == "assigned":
            cantidad_metricas = max(solicitada, 0)
            clasificacion = "TRANSITO_FULL"
        elif estado_odoo in ESTADOS_TRANSITO_FULL:
            # confirmed / waiting: el traslado existe pero no tiene reserva.
            # Se etiqueta aparte para no alterar la métrica histórica salvo
            # que INCLUIR_TRANSITO_SIN_RESERVA lo autorice.
            cantidad_metricas = max(solicitada, 0)
            clasificacion = "TRANSITO_FULL_SIN_RESERVA"
        elif estado_odoo == "done":
            cantidad_metricas = max(realizada, 0)
            clasificacion = "RECIBIDO_FULL"
        else:
            cantidad_metricas = 0
            clasificacion = "CANCELADO"

        rows.append({
            "picking_id": picking_id,
            "referencia": p.get("name", ""),
            "contacto": m2o_name(p.get("partner_id")),
            "canal": p.get("canal", ""),
            "regla_full": p.get("regla_full", ""),
            "ubicacion_origen": m2o_name(p.get("location_id")),
            "ubicacion_destino": m2o_name(p.get("location_dest_id")),
            "estado_odoo": estado_odoo,
            "estado": estado_picking_dashboard(estado_odoo),
            "fecha_programada": pd.to_datetime(p.get("scheduled_date"), errors="coerce"),
            "fecha_efectiva": pd.to_datetime(p.get("date_done"), errors="coerce"),
            "fecha_creacion": pd.to_datetime(p.get("create_date"), errors="coerce"),
            "ultima_modificacion": pd.to_datetime(p.get("write_date"), errors="coerce"),
            "documento_origen": p.get("origin", ""),
            "move_id": m.get("id"),
            "product_id": product_id,
            "sku_original": sku,
            "producto": prod.get("display_name", ""),
            "cantidad_solicitada": solicitada,
            "cantidad_realizada": realizada,
            "cantidad_metricas": cantidad_metricas,
            "clasificacion_inventario": clasificacion,
        })

    detalle = pd.DataFrame(rows, columns=columnas)
    return detalle, pd.DataFrame(excluidos)


def resumir_contactos_no_full(odoo, excluidos_df):
    """Agrupa los traslados descartados por contacto y cuantifica el hueco.

    Sirve para responder con números la pregunta '¿cuánto se nos está
    quedando fuera por cómo capturan el contacto?'. Trae las unidades de
    cada picking excluido para que el resumen no sea solo un conteo.
    """
    columnas = ["contacto", "motivo_exclusion", "pickings", "unidades",
                "canal_sugerido", "ejemplo_referencia", "ultima_fecha"]
    if excluidos_df is None or excluidos_df.empty:
        return pd.DataFrame(columns=columnas)

    df = excluidos_df.copy()
    if "id" not in df.columns:
        return pd.DataFrame(columns=columnas)

    picking_ids = [int(v) for v in df["id"].dropna().tolist()]
    unidades = {}
    if picking_ids:
        try:
            moves = odoo.search_read_all(
                "stock.move",
                [("picking_id", "in", picking_ids), ("product_id", "!=", False)],
                ["id", "picking_id", "product_uom_qty"],
                batch=5000, order="id asc"
            )
            for m in moves:
                pid = m2o_id(m.get("picking_id"))
                if pid:
                    unidades[pid] = unidades.get(pid, 0.0) + float(m.get("product_uom_qty") or 0)
        except Exception as e:
            print(f"AVISO: no se pudieron contar unidades de traslados excluidos: {e}")

    df["unidades_picking"] = df["id"].apply(lambda v: unidades.get(int(v), 0.0) if pd.notna(v) else 0.0)
    df["contacto"] = df.get("contacto", "").fillna("")
    df["motivo_exclusion"] = df.get("motivo_exclusion", "").fillna("")
    # Qué canal PARECE ser, aunque la regla no lo haya aceptado. Es la pista
    # para decidir si hay que corregir la captura o ampliar la regla.
    df["canal_sugerido"] = df["contacto"].apply(normalizar_canal_venta)

    resumen = df.groupby(["contacto", "motivo_exclusion", "canal_sugerido"], as_index=False).agg(
        pickings=("id", "nunique"),
        unidades=("unidades_picking", "sum"),
        ejemplo_referencia=("name", lambda x: next((str(v) for v in x if str(v).strip()), "")),
        ultima_fecha=("fecha_evento", "max"),
    )
    # Lo más sospechoso primero: contactos que parecen de un canal comercial
    # pero fueron rechazados.
    resumen["prioridad_revision"] = np.where(
        resumen["canal_sugerido"].isin(CANALES_COMERCIALES), "ALTA", "BAJA"
    )
    return resumen.sort_values(
        ["prioridad_revision", "unidades"], ascending=[True, False]
    )


# ============================================================
# ODOO: RECEPCIONES, ENVÍOS FULL HISTÓRICOS, COSTOS Y AUDITORÍA
# ============================================================

def _cantidad_hecha_move(m):
    """Cantidad realizada de un stock.move, tolerante a versiones de Odoo."""
    if "quantity_done" in m and m.get("quantity_done") is not None:
        return float(m.get("quantity_done") or 0)
    if "quantity" in m and m.get("quantity") is not None:
        return float(m.get("quantity") or 0)
    return float(m.get("product_uom_qty") or 0)


def _cantidad_move_line(ml):
    if "quantity" in ml and ml.get("quantity") is not None:
        return float(ml.get("quantity") or 0)
    return float(ml.get("qty_done") or 0)


def _leer_productos(odoo, product_ids, campos_extra=None):
    """Lee product.product en bloques y devuelve {id: registro}."""
    product_ids = sorted({int(x) for x in product_ids if x})
    if not product_ids:
        return {}
    disponibles = campos_disponibles_odoo(odoo, "product.product")
    campos = seleccionar_campos(
        disponibles,
        ["id", "display_name", "default_code", "barcode"] + list(campos_extra or []),
    )
    out = {}
    for i in range(0, len(product_ids), 500):
        for p in odoo.execute(
            "product.product", "read", product_ids[i:i + 500], campos,
            context={"active_test": False},
        ):
            out[p["id"]] = p
    return out


def extraer_recepciones_compra(odoo, dias=None):
    """Recepciones reales de proveedor (stock.move Hecho: supplier -> interno).

    Es la fuente para el respaldo FIFO de antigüedad y para el último precio
    de compra. A diferencia de purchase.order.line, trae la fecha real de
    cada recepción parcial.
    """
    columnas = [
        "move_id", "product_id", "sku_original", "producto", "fecha_recepcion",
        "cantidad_recibida", "precio_unitario_recepcion", "referencia",
    ]
    dias = ANTIGUEDAD_LOOKBACK_COMPRAS_DIAS if dias is None else dias
    desde = (pd.Timestamp.now() - pd.Timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")
    disponibles = campos_disponibles_odoo(odoo, "stock.move")
    campos = seleccionar_campos(disponibles, [
        "id", "product_id", "date", "product_uom_qty", "quantity", "quantity_done",
        "price_unit", "reference", "origin", "picking_id",
    ])
    try:
        moves = odoo.search_read_all(
            "stock.move",
            [
                ("state", "=", "done"),
                ("date", ">=", desde),
                ("location_id.usage", "=", "supplier"),
                ("location_dest_id.usage", "=", "internal"),
            ],
            campos, batch=5000, order="date asc",
        )
    except Exception as e:
        print(f"ADVERTENCIA: no se pudieron leer recepciones de compra: {e}")
        return pd.DataFrame(columns=columnas)
    if not moves:
        return pd.DataFrame(columns=columnas)

    productos = _leer_productos(odoo, [m2o_id(m.get("product_id")) for m in moves])
    rows = []
    for m in moves:
        pid = m2o_id(m.get("product_id"))
        p = productos.get(pid, {})
        qty = _cantidad_hecha_move(m)
        if qty <= 0:
            continue
        sku = limpiar_sku(p.get("default_code", "")) or limpiar_sku(p.get("barcode", ""))
        rows.append({
            "move_id": m.get("id"),
            "product_id": pid,
            "sku_original": sku or f"{SKU_FALLBACK_PREFIJO}{pid}",
            "producto": p.get("display_name", ""),
            "fecha_recepcion": pd.to_datetime(m.get("date"), errors="coerce"),
            "cantidad_recibida": qty,
            "precio_unitario_recepcion": float(m.get("price_unit") or 0),
            "referencia": m.get("reference") or m2o_name(m.get("picking_id")) or m.get("origin") or "",
        })
    df = pd.DataFrame(rows, columns=columnas)
    print(f"Recepciones de compra (últimos {dias} días): {len(df)} movimientos, "
          f"{df['cantidad_recibida'].sum():,.0f} unidades.")
    return df


def _dominio_or(terminos):
    """Arma un OR en notación polaca de Odoo: ['|','|',a,b,c]."""
    terminos = list(terminos)
    if not terminos:
        return []
    return ["|"] * (len(terminos) - 1) + terminos


def extraer_envios_full_historicos(odoo, ubicaciones_df, dias=None):
    """Envíos Odoo -> Full ya Hecho en los últimos ``dias`` días, por línea.

    Solo se usan para la antigüedad FIFO de Full y para la auditoría; NO
    alteran el inventario (ese se sigue calculando contra el corte maestro).
    Devuelve una fila por stock.move.line para conservar el lote enviado.
    """
    columnas = [
        "picking_id", "referencia", "contacto", "canal", "regla_full",
        "ubicacion_origen", "canal_origen", "fecha_envio", "move_line_id",
        "product_id", "sku_original", "producto", "lote_id_odoo", "lote_nombre",
        "cantidad_enviada", "usuario", "responsable",
    ]
    if ubicaciones_df is None or ubicaciones_df.empty:
        return pd.DataFrame(columns=columnas)
    dias = ANTIGUEDAD_LOOKBACK_ENVIOS_FULL_DIAS if dias is None else dias
    desde = (pd.Timestamp.now() - pd.Timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")
    location_ids = ubicaciones_df["id"].astype(int).tolist()

    disponibles_p = campos_disponibles_odoo(odoo, "stock.picking")
    campos_p = seleccionar_campos(disponibles_p, [
        "id", "name", "partner_id", "location_id", "state", "date_done", "user_id", "write_uid",
    ])
    dominio = [
        ("state", "=", "done"),
        ("date_done", ">=", desde),
        ("location_id", "in", location_ids),
        ("partner_id", "!=", False),
    ] + _dominio_or([("partner_id.name", "ilike", w) for w in PALABRAS_CONTACTO_FULL])
    try:
        pickings = odoo.search_read_all("stock.picking", dominio, campos_p, batch=3000, order="date_done asc")
    except Exception as e:
        print(f"ADVERTENCIA: no se pudieron leer envíos Full históricos: {e}")
        return pd.DataFrame(columns=columnas)

    aceptados = {}
    for p in pickings:
        canal, regla, _ = regla_canal_full(m2o_name(p.get("partner_id")))
        origen = m2o_name(p.get("location_id"))
        _, canal_origen, _ = clasificar_ubicacion_odoo(origen)
        if canal and canal_origen:
            p2 = dict(p)
            p2.update({"canal": canal, "regla_full": regla, "ubicacion_origen": origen, "canal_origen": canal_origen})
            aceptados[int(p["id"])] = p2
    if not aceptados:
        return pd.DataFrame(columns=columnas)

    disponibles_ml = campos_disponibles_odoo(odoo, "stock.move.line")
    campo_qty = "quantity" if (not disponibles_ml or "quantity" in disponibles_ml) else "qty_done"
    campos_ml = seleccionar_campos(disponibles_ml, [
        "id", "picking_id", "product_id", "lot_id", campo_qty, "state",
    ])
    lineas = []
    ids = sorted(aceptados)
    for i in range(0, len(ids), 2000):
        lineas.extend(odoo.search_read_all(
            "stock.move.line",
            [("picking_id", "in", ids[i:i + 2000]), ("product_id", "!=", False)],
            campos_ml, batch=8000, order="id asc",
        ))
    productos = _leer_productos(odoo, [m2o_id(l.get("product_id")) for l in lineas])

    rows = []
    for l in lineas:
        p = aceptados.get(m2o_id(l.get("picking_id")), {})
        pid = m2o_id(l.get("product_id"))
        prod = productos.get(pid, {})
        qty = _cantidad_move_line(l)
        if qty <= 0:
            continue
        sku = limpiar_sku(prod.get("default_code", "")) or limpiar_sku(prod.get("barcode", ""))
        rows.append({
            "picking_id": p.get("id"),
            "referencia": p.get("name", ""),
            "contacto": m2o_name(p.get("partner_id")),
            "canal": p.get("canal", ""),
            "regla_full": p.get("regla_full", ""),
            "ubicacion_origen": p.get("ubicacion_origen", ""),
            "canal_origen": p.get("canal_origen", ""),
            "fecha_envio": pd.to_datetime(p.get("date_done"), errors="coerce"),
            "move_line_id": l.get("id"),
            "product_id": pid,
            "sku_original": sku or f"{SKU_FALLBACK_PREFIJO}{pid}",
            "producto": prod.get("display_name", ""),
            "lote_id_odoo": m2o_id(l.get("lot_id")),
            "lote_nombre": m2o_name(l.get("lot_id")),
            "cantidad_enviada": qty,
            "usuario": m2o_name(p.get("write_uid")),
            "responsable": m2o_name(p.get("user_id")),
        })
    df = pd.DataFrame(rows, columns=columnas)
    print(f"Envíos Full históricos (últimos {dias} días): {df['picking_id'].nunique() if not df.empty else 0} "
          f"traslados, {df['cantidad_enviada'].sum() if not df.empty else 0:,.0f} unidades.")
    return df


def extraer_costos_productos(odoo, product_ids, skus_madre):
    """Costo unitario por producto Odoo: avg_cost (si existe) y standard_price.

    Incluye los productos cuyo código interno es un SKU madre, para poder
    valuar el inventario Full de SKU que hoy no tienen piezas en CUATI
    (antes su costo quedaba en 0 y su inversión no aparecía).
    """
    columnas = ["product_id", "sku_original", "producto", "avg_cost", "standard_price"]
    disponibles = campos_disponibles_odoo(odoo, "product.product")
    extra = [c for c in ["avg_cost", "standard_price"] if not disponibles or c in disponibles]
    ids = {int(x) for x in product_ids if x}
    skus = sorted({str(s).strip() for s in skus_madre if str(s).strip()})
    for i in range(0, len(skus), 500):
        try:
            encontrados = odoo.execute(
                "product.product", "search",
                [("default_code", "in", skus[i:i + 500])],
                context={"active_test": False},
            )
            ids.update(int(x) for x in encontrados)
        except Exception as e:
            print(f"ADVERTENCIA: no se pudieron buscar productos por SKU madre: {e}")
            break
    productos = _leer_productos(odoo, ids, extra)
    rows = []
    for pid, p in productos.items():
        sku = limpiar_sku(p.get("default_code", "")) or limpiar_sku(p.get("barcode", ""))
        rows.append({
            "product_id": pid,
            "sku_original": sku or f"{SKU_FALLBACK_PREFIJO}{pid}",
            "producto": p.get("display_name", ""),
            "avg_cost": float(p.get("avg_cost") or 0),
            "standard_price": float(p.get("standard_price") or 0),
        })
    return pd.DataFrame(rows, columns=columnas)


def _clasificar_extremo_ubicacion(nombre, usage):
    """Devuelve (tipo_extremo, canal) para un lado de un movimiento."""
    usage = str(usage or "").lower()
    if usage == "internal":
        raiz, canal, _ = clasificar_ubicacion_odoo(nombre)
        if raiz:
            return "CANAL", canal
        norm = normalizar_texto(nombre)
        ultimo = norm.split("/")[-1].strip()
        if norm.startswith(ODOO_RAIZ_VIGILADA + "/") and ultimo in UBICACIONES_OPERATIVAS:
            return "OPERATIVA", ""
        return "INTERNA_NO_COMERCIAL", ""
    if usage == "inventory":
        return "AJUSTE", ""
    if usage == "customer":
        return "CLIENTE", ""
    if usage == "transit":
        return "TRANSITO", ""
    if usage == "production":
        return "PRODUCCION", ""
    return (usage.upper() or "OTRA"), ""


def _tipo_movimiento_auditoria(t_origen, c_origen, t_destino, c_destino):
    if t_origen == "CANAL" and t_destino == "CANAL":
        return "MISMO_CANAL" if c_origen == c_destino else "REASIGNACION_ENTRE_CANALES"
    if t_destino == "AJUSTE":
        return "SALIDA_POR_AJUSTE"
    if t_origen == "AJUSTE":
        return "ENTRADA_POR_AJUSTE"
    if t_origen == "CLIENTE":
        return "DEVOLUCION_CLIENTE"
    if "OPERATIVA" in (t_origen, t_destino):
        return "FLUJO_OPERATIVO"
    if t_origen == "CANAL" and t_destino in ("INTERNA_NO_COMERCIAL", "TRANSITO"):
        return "SALIDA_A_UBICACION_NO_COMERCIAL"
    if t_destino == "CANAL" and t_origen in ("INTERNA_NO_COMERCIAL", "TRANSITO"):
        return "ENTRADA_DESDE_UBICACION_NO_COMERCIAL"
    return "OTRO"


TIPOS_AUDITORIA_OMITIDOS = {"MISMO_CANAL", "FLUJO_OPERATIVO"}


def extraer_movimientos_auditoria(odoo, ubicaciones_todas_df, dias=None):
    """Movimientos Hecho que tocan CUATI y NO son compra ni venta.

    Captura reasignaciones entre canales, ajustes de inventario, salidas a
    ubicaciones no comerciales y devoluciones. Es la materia prima del
    registro de movimientos que se usa para detectar traslados hechos para
    "reiniciar" la antigüedad del inventario de un canal.
    """
    columnas = [
        "move_line_id", "fecha_movimiento", "referencia", "documento_origen",
        "contacto", "tipo_operacion", "product_id", "sku_original", "producto",
        "lote_id_odoo", "lote_nombre", "cantidad", "ubicacion_origen",
        "ubicacion_destino", "tipo_origen", "tipo_destino", "canal_origen",
        "canal_destino", "tipo_movimiento", "usuario", "creado_por", "responsable",
    ]
    if ubicaciones_todas_df is None or ubicaciones_todas_df.empty:
        return pd.DataFrame(columns=columnas)
    dias = AUDITORIA_MOVIMIENTOS_DIAS if dias is None else dias
    desde = (pd.Timestamp.now() - pd.Timedelta(days=dias)).strftime("%Y-%m-%d %H:%M:%S")

    ub = ubicaciones_todas_df.copy()
    ub["norm"] = ub["complete_name"].apply(normalizar_texto)
    cuati_ids = ub.loc[
        ub["norm"].eq(ODOO_RAIZ_VIGILADA) | ub["norm"].str.startswith(ODOO_RAIZ_VIGILADA + "/"),
        "id",
    ].astype(int).tolist()
    if not cuati_ids:
        return pd.DataFrame(columns=columnas)

    disponibles = campos_disponibles_odoo(odoo, "stock.move.line")
    campo_qty = "quantity" if (not disponibles or "quantity" in disponibles) else "qty_done"
    campos = seleccionar_campos(disponibles, [
        "id", "date", "reference", "picking_id", "product_id", "lot_id",
        "location_id", "location_dest_id", campo_qty, "create_uid", "write_uid",
    ])
    dominio = [
        ("state", "=", "done"),
        ("date", ">=", desde),
        "|", ("location_id", "in", cuati_ids), ("location_dest_id", "in", cuati_ids),
        ("location_id.usage", "!=", "supplier"),
        ("location_dest_id.usage", "!=", "customer"),
    ]
    try:
        lineas = odoo.search_read_all("stock.move.line", dominio, campos, batch=5000, order="date asc")
    except Exception as e:
        print(f"ADVERTENCIA: no se pudieron leer movimientos para auditoría: {e}")
        return pd.DataFrame(columns=columnas)
    if not lineas:
        return pd.DataFrame(columns=columnas)

    loc_ids = sorted({
        x for l in lineas
        for x in (m2o_id(l.get("location_id")), m2o_id(l.get("location_dest_id"))) if x
    })
    loc_info = {}
    for i in range(0, len(loc_ids), 1000):
        for u in odoo.execute(
            "stock.location", "read", loc_ids[i:i + 1000], ["id", "complete_name", "usage"],
            context={"active_test": False},
        ):
            loc_info[u["id"]] = u

    picking_ids = sorted({m2o_id(l.get("picking_id")) for l in lineas if m2o_id(l.get("picking_id"))})
    disponibles_p = campos_disponibles_odoo(odoo, "stock.picking")
    campos_p = seleccionar_campos(disponibles_p, ["id", "name", "origin", "partner_id", "user_id", "picking_type_id"])
    pickings = {}
    for i in range(0, len(picking_ids), 1000):
        for p in odoo.execute("stock.picking", "read", picking_ids[i:i + 1000], campos_p):
            pickings[p["id"]] = p

    productos = _leer_productos(odoo, [m2o_id(l.get("product_id")) for l in lineas])

    rows = []
    for l in lineas:
        o = loc_info.get(m2o_id(l.get("location_id")), {})
        d = loc_info.get(m2o_id(l.get("location_dest_id")), {})
        t_o, c_o = _clasificar_extremo_ubicacion(o.get("complete_name", ""), o.get("usage", ""))
        t_d, c_d = _clasificar_extremo_ubicacion(d.get("complete_name", ""), d.get("usage", ""))
        tipo = _tipo_movimiento_auditoria(t_o, c_o, t_d, c_d)
        if tipo in TIPOS_AUDITORIA_OMITIDOS:
            continue
        qty = _cantidad_move_line(l)
        if qty <= 0:
            continue
        pid = m2o_id(l.get("product_id"))
        prod = productos.get(pid, {})
        sku = limpiar_sku(prod.get("default_code", "")) or limpiar_sku(prod.get("barcode", ""))
        pk = pickings.get(m2o_id(l.get("picking_id")), {})
        rows.append({
            "move_line_id": l.get("id"),
            "fecha_movimiento": pd.to_datetime(l.get("date"), errors="coerce"),
            "referencia": l.get("reference") or pk.get("name", ""),
            "documento_origen": pk.get("origin", "") or "",
            "contacto": m2o_name(pk.get("partner_id")),
            "tipo_operacion": m2o_name(pk.get("picking_type_id")),
            "product_id": pid,
            "sku_original": sku or f"{SKU_FALLBACK_PREFIJO}{pid}",
            "producto": prod.get("display_name", ""),
            "lote_id_odoo": m2o_id(l.get("lot_id")),
            "lote_nombre": m2o_name(l.get("lot_id")),
            "cantidad": qty,
            "ubicacion_origen": o.get("complete_name", ""),
            "ubicacion_destino": d.get("complete_name", ""),
            "tipo_origen": t_o,
            "tipo_destino": t_d,
            "canal_origen": c_o,
            "canal_destino": c_d,
            "tipo_movimiento": tipo,
            "usuario": m2o_name(l.get("write_uid")) or m2o_name(l.get("create_uid")),
            "creado_por": m2o_name(l.get("create_uid")),
            "responsable": m2o_name(pk.get("user_id")),
        })
    df = pd.DataFrame(rows, columns=columnas)
    print(f"Movimientos para auditoría (últimos {dias} días): {len(df)} líneas.")
    return df


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

    # No se toman números largos aislados: pueden ser referencias de venta.
    return ""


def preparar_ventas_autoazur(ruta_autoazur):
    if not ruta_autoazur:
        return pd.DataFrame()

    print("Preparando ventas Autoazur con SKU y referencias separados...")
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

    sku_vals, origen_sku = [], []
    for _, row in df.iterrows():
        candidatos = []
        for nombre, c in [
            ("sku_original", col_sku_original),
            ("sku_desde_columna", col_sku_columna),
            ("sku_desde_titulo", col_sku_titulo),
        ]:
            if c:
                valor = limpiar_sku(row.get(c, ""))
                if valor:
                    candidatos.append((valor, nombre))
        if col_producto:
            valor_titulo = extraer_sku_desde_texto(row.get(col_producto, ""))
            if valor_titulo:
                candidatos.append((valor_titulo, "producto_titulo"))

        if candidatos:
            sku_vals.append(candidatos[0][0])
            origen_sku.append(candidatos[0][1])
        else:
            sku_vals.append("")
            origen_sku.append("")

    out = pd.DataFrame({
        "fuente": "AUTOAZUR",
        "fecha": pd.to_datetime(df[col_fecha], errors="coerce") if col_fecha else pd.NaT,
        "pedido": df[col_folio] if col_folio else "",
        "estado": df[col_status] if col_status else "",
        "estado_odoo": "",
        "tipo_venta": df[col_canal] if col_canal else "Autoazur",
        "canal": df[col_canal] if col_canal else "Autoazur",
        "canal_venta": (df[col_canal].apply(normalizar_canal_venta) if col_canal else "Sin identificar"),
        "modalidad_venta": "PENDIENTE_ODOO",
        "tienda": df[col_tienda] if col_tienda else "",
        "cliente": "", "equipo_ventas": "", "almacen": "",
        "referencia_cliente": df[col_ref] if col_ref else "",
        "referencia": df[col_ref] if col_ref else "",
        "item_id": df[col_item] if col_item else "",
        "origen": "AUTOAZUR", "line_id": "", "product_id": "",
        "producto": df[col_producto] if col_producto else "",
        "sku_original": sku_vals,
        "origen_sku_producto": origen_sku,
        "sku_desde_columna": df[col_sku_columna] if col_sku_columna else "",
        "sku_desde_titulo": df[col_sku_titulo] if col_sku_titulo else "",
        "sku_default_code": "", "barcode": "", "categoria": "", "marca": "",
        "cantidad": to_number(df[col_cantidad]) if col_cantidad else 1,
        "cantidad_entregada": 0, "precio_unitario": 0,
        "venta_total": to_number(df[col_total]) if col_total else 0,
    })
    out["sku_original"] = out["sku_original"].apply(limpiar_sku)
    out = out[
        out["sku_original"].ne("")
        | out["referencia"].astype(str).str.strip().ne("")
        | out["item_id"].astype(str).str.strip().ne("")
        | out["pedido"].astype(str).str.strip().ne("")
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
    """Mapa referencia -> conjunto de SKU Odoo; evita adivinar cuando una orden tiene varios SKU."""
    if ventas_odoo.empty:
        return {}
    mapa = {}
    columnas_ref = ["referencia_cliente", "referencia", "origen", "pedido"]
    for _, row in ventas_odoo.iterrows():
        sku = limpiar_sku(row.get("sku_original", ""))
        if not sku:
            continue
        for col in columnas_ref:
            ref = limpiar_sku(row.get(col, ""))
            if not ref:
                continue
            for k in referencia_variantes(ref):
                if k:
                    mapa.setdefault(k, set()).add(sku)
    return {k: sorted(v) for k, v in mapa.items()}


def obtener_candidatos_sku_row(row):
    """Solo identificadores de producto; nunca referencia, pedido ni Item ID."""
    cols = ["sku_original", "sku_default_code", "sku_desde_columna", "sku_desde_titulo", "barcode"]
    candidatos = []
    for col in cols:
        val = limpiar_sku(row.get(col, ""))
        if val and val.lower() not in ["false", "falso", "true", "verdadero"] and val not in candidatos:
            candidatos.append(val)
    return candidatos


def buscar_match_diccionario_por_candidatos(candidatos_sku, dic_map):
    for sku in candidatos_sku:
        for k in generar_sku_keys_match(sku):
            if k in dic_map:
                metodo = "sku_original" if k == sku_key(sku) else "sku_sin_ceros_izquierda"
                return dic_map[k], metodo, sku
    return None, "", ""


def cruzar_ventas_con_diccionario(ventas_conjunto, dic_match, ventas_odoo):
    """Vincula por SKU de producto; las referencias solo sirven para consultar Odoo."""
    if ventas_conjunto.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")
    mapa_ref_odoo = construir_mapa_referencia_odoo(ventas_odoo)
    rows = []

    for _, original in ventas_conjunto.iterrows():
        row = original.copy()
        fuente = str(row.get("fuente", "")).upper()
        candidatos_sku = obtener_candidatos_sku_row(row)
        match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos(candidatos_sku, dic_map)

        sku_odoo_desde_referencia = ""
        referencia_usada_para_odoo = ""
        estado_referencia_odoo = "NO_APLICA"
        skus_odoo_candidatos = []

        if match is None and fuente == "AUTOAZUR":
            refs = [row.get("referencia", ""), row.get("referencia_cliente", ""), row.get("pedido", "")]
            encontro_clave = False
            for ref in refs:
                for ref_k in referencia_variantes(ref):
                    if ref_k in mapa_ref_odoo:
                        encontro_clave = True
                        skus_odoo_candidatos = mapa_ref_odoo[ref_k]
                        referencia_usada_para_odoo = str(ref)
                        break
                if encontro_clave:
                    break

            if len(skus_odoo_candidatos) == 1:
                estado_referencia_odoo = "UNICA"
                sku_odoo_desde_referencia = skus_odoo_candidatos[0]
                match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos(
                    [sku_odoo_desde_referencia], dic_map
                )
                if match is not None:
                    metodo_match = "referencia_autoazur_vs_odoo"
            elif len(skus_odoo_candidatos) > 1:
                estado_referencia_odoo = "AMBIGUA"
            else:
                tiene_ref = any(str(v or "").strip() for v in refs)
                estado_referencia_odoo = "NO_ENCONTRADA" if tiene_ref else "SIN_REFERENCIA"

        sku_original = limpiar_sku(row.get("sku_original", ""))
        sku_producto_pendiente = candidatos_sku[0] if candidatos_sku else sku_odoo_desde_referencia
        row["sku_key"] = sku_key(sku_original)
        row["sku_key_sin_ceros"] = sku_key_sin_ceros(sku_original)
        row["sku_odoo_desde_referencia"] = sku_odoo_desde_referencia
        row["referencia_usada_para_odoo"] = referencia_usada_para_odoo
        row["estado_referencia_odoo"] = estado_referencia_odoo
        row["skus_odoo_candidatos"] = " | ".join(skus_odoo_candidatos)
        row["sku_producto_pendiente"] = limpiar_sku(sku_producto_pendiente)
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
            for c in ["alias_diccionario", "sku_madre", "producto_madre", "hoja_diccionario", "columna_alias"]:
                row[c] = ""
            row["tiene_referencia_madre"] = "NO"
        rows.append(row)

    ventas_con_match = pd.DataFrame(rows)
    ventas_sin_match = ventas_con_match[ventas_con_match["tiene_referencia_madre"].eq("NO")].copy()
    vinculadas = ventas_con_match[ventas_con_match["tiene_referencia_madre"].eq("SI")].copy()
    if vinculadas.empty:
        ventas_sku_madre = pd.DataFrame()
    else:
        ventas_sku_madre = vinculadas.groupby(["sku_madre", "producto_madre"], as_index=False).agg(
            unidades_vendidas=("cantidad", "sum"), venta_total=("venta_total", "sum"),
            pedidos=("pedido", "nunique"), lineas=("sku_usado_para_match", "count"),
            fuentes_venta=("fuente", lambda x: " | ".join(sorted(set(map(str, x))))),
            metodos_match=("metodo_match", lambda x: " | ".join(sorted(set(map(str, x))))),
            skus_usados_para_match=("sku_usado_para_match", lambda x: " | ".join(sorted(set(str(v) for v in x if str(v).strip()))[:80])),
            skus_originales=("sku_original", lambda x: " | ".join(sorted(set(map(str, x)))[:80])),
            skus_odoo_desde_referencia=("sku_odoo_desde_referencia", lambda x: " | ".join(sorted(set(str(v) for v in x if str(v).strip()))[:80])),
        ).sort_values("unidades_vendidas", ascending=False)
    return ventas_con_match, ventas_sin_match, ventas_sku_madre


def cruzar_stock_con_diccionario(stock_por_sku, dic_match):
    if stock_por_sku.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")
    rows = []
    for _, row in stock_por_sku.iterrows():
        row = row.copy()
        sku_original = limpiar_sku(row.get("sku_original", ""))
        match, metodo_match, sku_usado_para_match = buscar_match_diccionario_por_candidatos([sku_original], dic_map)
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
            row["origen_sku_madre"] = "DICCIONARIO"
        elif AUTOGENERAR_SKU_MADRE and sku_original:
            # Sin alias en origen4 el producto quedaba fuera de todas las
            # métricas. Se vuelve su propio SKU madre para que su stock y su
            # costo aparezcan, y se marca para curación manual.
            row["alias_diccionario"] = ""
            row["sku_madre"] = sku_original
            row["producto_madre"] = str(row.get("producto_stock", "") or sku_original)
            row["hoja_diccionario"] = ""
            row["columna_alias"] = ""
            row["tiene_referencia_madre"] = "SI"
            row["origen_sku_madre"] = "AUTOGENERADA"
        else:
            row["alias_diccionario"] = ""
            row["sku_madre"] = ""
            row["producto_madre"] = ""
            row["hoja_diccionario"] = ""
            row["columna_alias"] = ""
            row["tiene_referencia_madre"] = "NO"
            row["origen_sku_madre"] = "SIN_RESOLVER"
        rows.append(row)

    stock_con_match = pd.DataFrame(rows)
    # La hoja de excepciones ahora lista lo que necesita curación (madre
    # autogenerada o sin resolver), no solo lo que quedó fuera.
    stock_sin_match = stock_con_match[
        stock_con_match["origen_sku_madre"].isin(["AUTOGENERADA", "SIN_RESOLVER"])
    ].copy()

    vinculados = stock_con_match[stock_con_match["tiene_referencia_madre"] == "SI"].copy()
    if vinculados.empty:
        return stock_con_match, stock_sin_match, pd.DataFrame()

    # Solo columnas numéricas de canal. 'producto_stock' y las banderas de
    # trazabilidad son texto y sumarlas concatenaría cadenas.
    no_agregables = {
        "sku_key", "sku_original", "stock_total", "producto_stock",
        "sku_key_sin_ceros", "sku_usado_para_match", "metodo_match",
        "alias_diccionario", "sku_madre", "producto_madre",
        "hoja_diccionario", "columna_alias", "tiene_referencia_madre",
        "origen_sku_madre",
    }
    canales_stock = [
        c for c in stock_por_sku.columns
        if c not in no_agregables
        and pd.api.types.is_numeric_dtype(stock_por_sku[c])
    ]
    agg = {c: (c, "sum") for c in canales_stock if c in vinculados.columns}
    agg.update({
        "stock_total": ("stock_total", "sum"),
        "metodos_match": ("metodo_match", lambda x: " | ".join(sorted(set(map(str, x))))),
        "skus_usados_para_match": ("sku_usado_para_match", lambda x: " | ".join(sorted(set([str(v) for v in x if str(v).strip()]))[:80])),
        "skus_originales": ("sku_original", lambda x: " | ".join(sorted(set(map(str, x)))[:80])),
    })
    stock_sku_madre = vinculados.groupby(["sku_madre", "producto_madre"], as_index=False).agg(**agg)

    legacy = {
        "stock_walmart_wfs": "WALMART_WFS",
        "stock_liverpool_99min": "LIVERPOOL_FULL_99MIN",
        "stock_meli_full": "MERCADO_LIBRE_FULL",
        "stock_amazon_fba": "AMAZON_FBA",
    }
    for destino, origen in legacy.items():
        # Si la columna no existe, .get() devuelve un entero y .fillna() truena.
        # Se construye una Serie del largo correcto para que siempre funcione.
        if origen in stock_sku_madre.columns:
            serie = pd.to_numeric(stock_sku_madre[origen], errors="coerce").fillna(0)
        else:
            serie = pd.Series(0.0, index=stock_sku_madre.index)
        stock_sku_madre[destino] = serie

    odoo_cols = [c for c in stock_sku_madre.columns if c.startswith("ODOO_")]
    stock_sku_madre["stock_odoo_cuautitlan"] = stock_sku_madre[odoo_cols].sum(axis=1) if odoo_cols else 0
    return stock_con_match, stock_sin_match, stock_sku_madre.sort_values("stock_total", ascending=False)


def vincular_detalle_sku(df, dic_match, sku_col="sku_original", autogenerar=None):
    """Agrega sku_madre a un detalle sin perder sus columnas originales.

    Si ``autogenerar`` está activo y el SKU no tiene alias en el diccionario,
    el propio SKU pasa a ser su madre. Así un traslado o una venta de un
    producto todavía no catalogado deja de desaparecer de las métricas.
    """
    if df is None or df.empty:
        return pd.DataFrame()
    if autogenerar is None:
        autogenerar = AUTOGENERAR_SKU_MADRE
    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")
    columnas_nombre = ["producto", "producto_stock", "producto_madre", "descripcion"]
    rows = []
    for _, original in df.iterrows():
        row = original.copy()
        sku = limpiar_sku(row.get(sku_col, ""))
        match, metodo, usado = buscar_match_diccionario_por_candidatos([sku], dic_map)
        row["sku_usado_para_match"] = usado
        row["metodo_match"] = metodo
        if match:
            row["sku_madre"] = match.get("sku_madre", "")
            row["producto_madre"] = match.get("producto_madre", "")
            row["tiene_referencia_madre"] = "SI"
            row["origen_sku_madre"] = "DICCIONARIO"
        elif autogenerar and sku:
            nombre = ""
            for c in columnas_nombre:
                valor = str(row.get(c, "") or "").strip()
                if valor:
                    nombre = valor
                    break
            row["sku_madre"] = sku
            row["producto_madre"] = nombre or sku
            row["tiene_referencia_madre"] = "SI"
            row["origen_sku_madre"] = "AUTOGENERADA"
        else:
            row["sku_madre"] = ""
            row["producto_madre"] = ""
            row["tiene_referencia_madre"] = "NO"
            row["origen_sku_madre"] = "SIN_RESOLVER"
        rows.append(row)
    return pd.DataFrame(rows)


def construir_inventario_canal_sku_madre(stock_detalle_vinculado, traslados_vinculados, ventas_vinculadas):
    """Construye primero cada canal y después el total, sin promediar coberturas."""
    keys = ["sku_madre", "producto_madre", "canal"]

    def vacio(nombre):
        return pd.DataFrame(columns=keys + [nombre])

    stock_ok = stock_detalle_vinculado[
        stock_detalle_vinculado.get("tiene_referencia_madre", "").astype(str).eq("SI")
    ].copy() if not stock_detalle_vinculado.empty else pd.DataFrame()

    if not stock_ok.empty:
        stock_ok["canal"] = stock_ok.get("canal_asignado", "General").replace("", "General")
        odoo_rows = stock_ok[stock_ok.get("tipo_inventario", "").eq("ODOO")].copy()
        odoo = odoo_rows.groupby(keys, as_index=False).agg(
            inventario_odoo=("stock", "sum")
        )
        # Costo CUATI: se acumula el costo total (value de Odoo, con respaldo
        # standard_price) y la cantidad física de todas las ubicaciones CUATI
        # ligadas a cada SKU madre, para obtener después un costo unitario
        # promedio ponderado por unidad física (no solo por canal).
        if "costo_total_ubicacion_odoo" in odoo_rows.columns:
            costo_odoo = odoo_rows.groupby("sku_madre", as_index=False).agg(
                costo_total_cuati=("costo_total_ubicacion_odoo", "sum"),
                unidades_fisicas_cuati=("quantity_odoo", "sum"),
            )
        else:
            costo_odoo = pd.DataFrame(columns=["sku_madre", "costo_total_cuati", "unidades_fisicas_cuati"])
        full_corte = stock_ok[stock_ok.get("tipo_inventario", "").eq("FULL")].groupby(keys, as_index=False).agg(
            inventario_full_corte=("stock", "sum")
        )
    else:
        odoo, full_corte = vacio("inventario_odoo"), vacio("inventario_full_corte")
        costo_odoo = pd.DataFrame(columns=["sku_madre", "costo_total_cuati", "unidades_fisicas_cuati"])

    tr_ok = traslados_vinculados[
        traslados_vinculados.get("tiene_referencia_madre", "").astype(str).eq("SI")
    ].copy() if not traslados_vinculados.empty else pd.DataFrame()
    if not tr_ok.empty:
        clases_transito = ["TRANSITO_FULL"]
        if INCLUIR_TRANSITO_SIN_RESERVA:
            clases_transito.append("TRANSITO_FULL_SIN_RESERVA")
        transito = tr_ok[
            tr_ok["clasificacion_inventario"].isin(clases_transito)
        ].groupby(keys, as_index=False).agg(
            inventario_transito=("cantidad_metricas", "sum")
        )
        # Siempre visible por separado, se incluya o no en la métrica.
        transito_sin_reserva = tr_ok[
            tr_ok["clasificacion_inventario"].eq("TRANSITO_FULL_SIN_RESERVA")
        ].groupby(keys, as_index=False).agg(
            inventario_transito_sin_reserva=("cantidad_metricas", "sum")
        )
        hechos_tmp = tr_ok[
            tr_ok["clasificacion_inventario"].eq("RECIBIDO_FULL")
            & (pd.to_datetime(tr_ok["fecha_efectiva"], errors="coerce") > FECHA_CORTE_FULL_ODOO_UTC)
        ].copy()
        hechos = hechos_tmp.groupby(keys, as_index=False).agg(
            traslados_hecho_post_corte=("cantidad_metricas", "sum")
        ) if not hechos_tmp.empty else vacio("traslados_hecho_post_corte")
    else:
        transito, hechos = vacio("inventario_transito"), vacio("traslados_hecho_post_corte")
        transito_sin_reserva = vacio("inventario_transito_sin_reserva")

    ven = ventas_vinculadas.copy() if ventas_vinculadas is not None else pd.DataFrame()
    if not ven.empty:
        ven["fecha"] = pd.to_datetime(ven.get("fecha"), errors="coerce")
        vf = ven[
            ven.get("tiene_referencia_madre", "").astype(str).eq("SI")
            & ven.get("modalidad_venta", "").astype(str).eq("FULL")
            & ven.get("canal_venta", "").astype(str).isin(CANALES_COMERCIALES)
            & (ven["fecha"] > FECHA_CORTE_FULL_ODOO_UTC)
        ].copy()
        if not vf.empty:
            vf["canal"] = vf["canal_venta"]
            ventas_full = vf.groupby(keys, as_index=False).agg(
                ventas_full_post_corte=("cantidad", "sum")
            )
        else:
            ventas_full = vacio("ventas_full_post_corte")
    else:
        ventas_full = vacio("ventas_full_post_corte")

    frames = [odoo, full_corte, transito, transito_sin_reserva, hechos, ventas_full]
    base = None
    for frame in frames:
        base = frame.copy() if base is None else base.merge(frame, on=keys, how="outer")
    if base is None or base.empty:
        return pd.DataFrame()

    for c in ["inventario_odoo", "inventario_full_corte", "inventario_transito",
              "inventario_transito_sin_reserva", "traslados_hecho_post_corte",
              "ventas_full_post_corte"]:
        base[c] = pd.to_numeric(base.get(c, 0), errors="coerce").fillna(0).clip(lower=0)

    base["inventario_full"] = (
        base["inventario_full_corte"]
        + base["traslados_hecho_post_corte"]
        - base["ventas_full_post_corte"]
    ).clip(lower=0)
    base["inventario_total"] = base["inventario_odoo"] + base["inventario_transito"] + base["inventario_full"]
    base["validacion_componentes"] = base["inventario_total"] - (
        base["inventario_odoo"] + base["inventario_transito"] + base["inventario_full"]
    )
    base["fecha_corte_full"] = FECHA_CORTE_FULL
    base["fecha_actualizacion_odoo"] = pd.Timestamp.now()

    # Metadatos de producto para el motor de repartición. Se resuelven a nivel
    # SKU madre desde el stock vivo de Odoo y después se replican a todos los
    # canales del mismo IQ. Si no existe marca en esta instancia, queda vacío.
    if not stock_ok.empty:
        meta = stock_ok.copy()
        if "categoria_odoo" not in meta.columns:
            meta["categoria_odoo"] = ""
        if "marca_odoo" not in meta.columns:
            meta["marca_odoo"] = ""
        def _primero_no_vacio(s):
            vals = [str(x).strip() for x in s if str(x).strip() and str(x).lower() not in {"nan", "none"}]
            return vals[0] if vals else ""
        meta_sku = meta.groupby("sku_madre", as_index=False).agg(
            categoria=("categoria_odoo", _primero_no_vacio),
            marca=("marca_odoo", _primero_no_vacio),
        )
        base = base.merge(meta_sku, on="sku_madre", how="left")
    for _c in ["categoria", "marca"]:
        if _c not in base.columns:
            base[_c] = ""
        base[_c] = base[_c].fillna("").astype(str).str.strip()

    # Costo unitario CUATI por SKU madre: promedio ponderado de todas las
    # ubicaciones CUATI (costo total / unidades físicas). Es el mismo valor
    # para todos los canales de un SKU madre porque el costo no distingue
    # por canal de venta, solo por ubicación de inventario en Odoo.
    if not costo_odoo.empty:
        costo_odoo = costo_odoo.copy()
        costo_odoo["costo_unitario_cuati"] = np.where(
            costo_odoo["unidades_fisicas_cuati"] > 0,
            costo_odoo["costo_total_cuati"] / costo_odoo["unidades_fisicas_cuati"],
            0.0
        )
        base = base.merge(
            costo_odoo[["sku_madre", "costo_total_cuati", "unidades_fisicas_cuati", "costo_unitario_cuati"]],
            on="sku_madre", how="left"
        )
    for c in ["costo_total_cuati", "unidades_fisicas_cuati", "costo_unitario_cuati"]:
        base[c] = pd.to_numeric(base.get(c, 0), errors="coerce").fillna(0)
    return base.sort_values(["canal", "sku_madre"])


def construir_dias_inventario_lote_canal(stock_detalle_vinculado):
    """Días de inventario por canal, ponderados por las piezas de cada lote.

    Solo cubre el inventario que hoy vive en ubicaciones Odoo/CUATI (columna
    tipo_inventario == "ODOO"), porque es ahí donde se puede rastrear el
    lote (Odoo -> Inventario -> Reporte -> Trazabilidad) y su fecha de
    recepción original en Compras. El inventario que ya está físicamente en
    Full/tránsito no conserva el lote en esta base, así que ese excedente
    queda marcado en 'unidades_sin_lote_odoo' y el script de métricas le
    aplica un respaldo aproximado (fecha del último arribo de compra del
    SKU completo).
    """
    columnas = [
        "sku_madre", "producto_madre", "canal",
        "dias_inventario_lote", "unidades_con_lote", "unidades_sin_lote_odoo",
        "cobertura_lote_pct", "fecha_recepcion_lote_min", "fecha_recepcion_lote_max",
        "fuente_dias_inventario",
    ]
    if stock_detalle_vinculado is None or stock_detalle_vinculado.empty:
        return pd.DataFrame(columns=columnas)

    df = stock_detalle_vinculado[
        stock_detalle_vinculado.get("tiene_referencia_madre", "").astype(str).eq("SI")
        & stock_detalle_vinculado.get("tipo_inventario", "").astype(str).eq("ODOO")
    ].copy()
    if df.empty:
        return pd.DataFrame(columns=columnas)

    df["canal"] = df.get("canal_asignado", "General").replace("", "General")
    df["stock"] = pd.to_numeric(df.get("stock", 0), errors="coerce").fillna(0).clip(lower=0)
    df["dias_en_stock_lote"] = pd.to_numeric(df.get("dias_en_stock_lote"), errors="coerce")
    df["fecha_recepcion_lote"] = pd.to_datetime(df.get("fecha_recepcion_lote"), errors="coerce")

    keys = ["sku_madre", "producto_madre", "canal"]

    total_odoo = df.groupby(keys, as_index=False).agg(unidades_odoo_total=("stock", "sum"))

    con_lote = df[df["dias_en_stock_lote"].notna() & (df["stock"] > 0)].copy()
    if not con_lote.empty:
        con_lote["peso_dias"] = con_lote["stock"] * con_lote["dias_en_stock_lote"]
        agg_con_lote = con_lote.groupby(keys, as_index=False).agg(
            unidades_con_lote=("stock", "sum"),
            suma_peso_dias=("peso_dias", "sum"),
            fecha_recepcion_lote_min=("fecha_recepcion_lote", "min"),
            fecha_recepcion_lote_max=("fecha_recepcion_lote", "max"),
        )
        agg_con_lote["dias_inventario_lote"] = np.where(
            agg_con_lote["unidades_con_lote"] > 0,
            agg_con_lote["suma_peso_dias"] / agg_con_lote["unidades_con_lote"],
            np.nan,
        )
    else:
        agg_con_lote = pd.DataFrame(columns=keys + [
            "unidades_con_lote", "dias_inventario_lote",
            "fecha_recepcion_lote_min", "fecha_recepcion_lote_max"
        ])

    out = total_odoo.merge(
        agg_con_lote[[
            "sku_madre", "producto_madre", "canal", "unidades_con_lote",
            "dias_inventario_lote", "fecha_recepcion_lote_min", "fecha_recepcion_lote_max"
        ]],
        on=keys, how="left"
    )
    out["unidades_con_lote"] = out["unidades_con_lote"].fillna(0)
    out["unidades_sin_lote_odoo"] = (out["unidades_odoo_total"] - out["unidades_con_lote"]).clip(lower=0)
    out["cobertura_lote_pct"] = np.where(
        out["unidades_odoo_total"] > 0,
        out["unidades_con_lote"] / out["unidades_odoo_total"],
        0.0
    )
    out["fuente_dias_inventario"] = np.where(
        out["dias_inventario_lote"].notna() & (out["cobertura_lote_pct"] >= 0.5),
        "LOTE",
        np.where(out["dias_inventario_lote"].notna(), "LOTE_PARCIAL", "SIN_LOTE")
    )
    return out[columnas]


# ============================================================
# ANTIGÜEDAD FIFO POR CANAL (versión 2026-09)
# ============================================================
#
# Antes: los días de inventario de un canal salían SOLO de los lotes que
# hoy están en ubicaciones Odoo/CUATI. Si un canal tenía aunque fuera una
# pieza en CUATI (por ejemplo, devoluciones que regresan de Full a
# CUATI/MercadoLibre), el resultado ignoraba por completo las piezas que
# están en Full. Además, la devolución contaba como "recepción" del lote.
#
# Ahora cada pieza del canal recibe una fecha de referencia según dónde está:
#   - Odoo/CUATI : fecha de ingreso del lote (compra > ajuste; nunca una
#                  devolución). Sin lote -> respaldo FIFO de compras del SKU.
#   - Full       : fecha del envío Odoo -> Full, por capas FIFO (lo que queda
#                  en Full son las piezas de los envíos más recientes).
#   - Tránsito   : respaldo FIFO de compras del SKU.
# Con esas capas se calcula el promedio ponderado, la capa más antigua y las
# piezas que ya superan el umbral (90 días por defecto).

FUENTES_ANTIGUEDAD_LEGIBLES = {
    "LOTE_COMPRA": "LOTE",
    "LOTE_PRODUCCION": "LOTE",
    "LOTE_AJUSTE_INVENTARIO": "LOTE_AJUSTE",
    "ENVIO_FULL": "ENVIO_FULL",
    "APROX_FIFO_COMPRAS": "APROX_FIFO_COMPRAS",
    "PRIMER_MOVIMIENTO_LOTE": "APROX_PRIMER_MOVIMIENTO",
    "SIN_DATO": "SIN_DATO",
}


def _fifo_capas(stock, eventos):
    """Asigna ``stock`` a eventos ordenados del más reciente al más antiguo.

    eventos: lista de (fecha, cantidad, referencia).
    Devuelve (capas[(cantidad, fecha, referencia)], faltante_sin_evento).
    """
    restante = float(max(stock or 0, 0))
    capas = []
    for fecha, qty, ref in eventos:
        if restante <= 1e-9:
            break
        qty = float(qty or 0)
        if qty <= 0 or pd.isna(fecha):
            continue
        toma = min(restante, qty)
        capas.append((toma, fecha, ref))
        restante -= toma
    return capas, max(restante, 0.0)


def _eventos_por_llave(df, llaves, col_fecha, col_qty, col_ref):
    """{llave: [(fecha, qty, ref), ...]} ordenado del más reciente al más antiguo."""
    out = {}
    if df is None or df.empty:
        return out
    tmp = df.copy()
    tmp[col_fecha] = pd.to_datetime(tmp[col_fecha], errors="coerce")
    tmp[col_qty] = pd.to_numeric(tmp[col_qty], errors="coerce").fillna(0)
    tmp = tmp[tmp[col_fecha].notna() & (tmp[col_qty] > 0)]
    tmp = tmp.sort_values(col_fecha, ascending=False)
    for r in tmp[llaves + [col_fecha, col_qty, col_ref]].itertuples(index=False):
        key = tuple(r[:len(llaves)]) if len(llaves) > 1 else r[0]
        out.setdefault(key, []).append((r[-3], r[-2], r[-1]))
    return out


def construir_fifo_compras_sku(inventario_canal, recepciones_vinc):
    """Antigüedad de TODO el inventario de cada SKU madre según compras (FIFO).

    Se usa como respaldo y como referencia "desde compra". No depende de
    lotes ni de dónde esté la pieza, así que ningún traslado interno, ajuste
    o devolución la reinicia.
    """
    columnas = [
        "sku_madre", "fifo_compras_fecha_prom", "fifo_compras_dias_prom",
        "fifo_compras_dias_max", "fifo_compras_estado", "fecha_ultima_recepcion_compra",
    ]
    if inventario_canal is None or inventario_canal.empty:
        return pd.DataFrame(columns=columnas)
    hoy = pd.Timestamp.now().normalize()
    stock = inventario_canal.groupby("sku_madre")["inventario_total"].sum()

    rec = pd.DataFrame()
    if recepciones_vinc is not None and not recepciones_vinc.empty:
        rec = recepciones_vinc[
            recepciones_vinc.get("tiene_referencia_madre", "").astype(str).eq("SI")
        ].copy()
    eventos = _eventos_por_llave(rec, ["sku_madre"], "fecha_recepcion", "cantidad_recibida", "referencia") \
        if not rec.empty else {}

    rows = []
    for sku, st in stock.items():
        evs = eventos.get(sku, [])
        fila = {
            "sku_madre": sku,
            "fifo_compras_fecha_prom": pd.NaT,
            "fifo_compras_dias_prom": np.nan,
            "fifo_compras_dias_max": np.nan,
            "fifo_compras_estado": "SIN_COMPRAS_EN_VENTANA" if not evs else "COMPLETO",
            "fecha_ultima_recepcion_compra": evs[0][0] if evs else pd.NaT,
        }
        if st > 0 and evs:
            capas, faltante = _fifo_capas(st, evs)
            if faltante > 1e-9:
                # Piezas más viejas que la ventana: al menos tan viejas como
                # la recepción más antigua encontrada.
                capas.append((faltante, evs[-1][0], "ANTERIOR_A_VENTANA"))
                fila["fifo_compras_estado"] = "INCOMPLETO"
            unidades = sum(c[0] for c in capas)
            dias = [max((hoy - pd.Timestamp(c[1]).normalize()).days, 0) for c in capas]
            prom = sum(c[0] * d for c, d in zip(capas, dias)) / unidades if unidades else np.nan
            fila["fifo_compras_dias_prom"] = prom
            fila["fifo_compras_dias_max"] = max(dias) if dias else np.nan
            fila["fifo_compras_fecha_prom"] = (hoy - pd.Timedelta(days=float(prom))).normalize() if pd.notna(prom) else pd.NaT
        rows.append(fila)
    return pd.DataFrame(rows, columns=columnas)


def construir_antiguedad_canal(inventario_canal, stock_detalle_vinculado, envios_full_vinc,
                               recepciones_vinc, umbral=None):
    """Devuelve (resumen por sku_madre+canal, capas de antigüedad, fifo por SKU)."""
    umbral = ANTIGUEDAD_UMBRAL_ALERTA_DIAS if umbral is None else umbral
    hoy = pd.Timestamp.now().normalize()
    cols_capas = [
        "sku_madre", "producto_madre", "canal", "componente", "unidades",
        "fecha_referencia", "dias", "fuente", "referencia",
    ]
    if inventario_canal is None or inventario_canal.empty:
        return pd.DataFrame(), pd.DataFrame(columns=cols_capas), pd.DataFrame()

    fifo_sku = construir_fifo_compras_sku(inventario_canal, recepciones_vinc)
    fifo_map = fifo_sku.set_index("sku_madre").to_dict("index") if not fifo_sku.empty else {}

    def respaldo(sku):
        f = fifo_map.get(sku, {}).get("fifo_compras_fecha_prom", pd.NaT)
        return pd.Timestamp(f) if pd.notna(f) else pd.NaT

    nombres = inventario_canal.drop_duplicates("sku_madre").set_index("sku_madre")["producto_madre"].to_dict()
    capas = []

    # 1) Piezas en Odoo/CUATI (por quant, con su lote).
    det = pd.DataFrame()
    if stock_detalle_vinculado is not None and not stock_detalle_vinculado.empty:
        det = stock_detalle_vinculado[
            stock_detalle_vinculado.get("tiene_referencia_madre", "").astype(str).eq("SI")
            & stock_detalle_vinculado.get("tipo_inventario", "").astype(str).eq("ODOO")
        ].copy()
    if not det.empty:
        det["stock"] = pd.to_numeric(det.get("stock", 0), errors="coerce").fillna(0)
        det = det[det["stock"] > 0].copy()
        det["canal"] = det.get("canal_asignado", "General").fillna("").replace("", "General")
        for c in ["fecha_recepcion_lote", "fecha_primer_movimiento_lote"]:
            det[c] = pd.to_datetime(det[c], errors="coerce") if c in det.columns else pd.NaT
        for c in ["fuente_recepcion_lote", "lote_nombre", "odoo_location"]:
            det[c] = det[c].fillna("").astype(str) if c in det.columns else ""
        for r in det.to_dict("records"):
            sku = r["sku_madre"]
            fecha = r["fecha_recepcion_lote"]
            if pd.notna(fecha):
                fuente = f"LOTE_{r['fuente_recepcion_lote'] or 'COMPRA'}"
            else:
                fr, fpm = respaldo(sku), r["fecha_primer_movimiento_lote"]
                opciones = [x for x in (fr, fpm) if pd.notna(x)]
                if opciones:
                    fecha = min(opciones)
                    fuente = "APROX_FIFO_COMPRAS" if (pd.notna(fr) and fecha == fr) else "PRIMER_MOVIMIENTO_LOTE"
                else:
                    fuente = "SIN_DATO"
            capas.append({
                "sku_madre": sku, "producto_madre": r.get("producto_madre", ""),
                "canal": r["canal"], "componente": "ODOO", "unidades": float(r["stock"]),
                "fecha_referencia": fecha, "fuente": fuente,
                "referencia": r["lote_nombre"] or r["odoo_location"],
            })

    # 2) Piezas en Full: FIFO contra envíos Odoo -> Full ya recibidos.
    env = pd.DataFrame()
    if envios_full_vinc is not None and not envios_full_vinc.empty:
        env = envios_full_vinc[
            envios_full_vinc.get("tiene_referencia_madre", "").astype(str).eq("SI")
        ].copy()
    eventos_full = _eventos_por_llave(env, ["sku_madre", "canal"], "fecha_envio", "cantidad_enviada", "referencia") \
        if not env.empty else {}

    inv = inventario_canal.copy()
    for c in ["inventario_full", "inventario_transito"]:
        inv[c] = pd.to_numeric(inv.get(c, 0), errors="coerce").fillna(0)
    for r in inv[inv["inventario_full"] > 0].to_dict("records"):
        sku, canal = r["sku_madre"], r["canal"]
        capas_f, faltante = _fifo_capas(r["inventario_full"], eventos_full.get((sku, canal), []))
        for qty, fecha, ref in capas_f:
            capas.append({
                "sku_madre": sku, "producto_madre": r.get("producto_madre", ""), "canal": canal,
                "componente": "FULL", "unidades": qty, "fecha_referencia": fecha,
                "fuente": "ENVIO_FULL", "referencia": ref,
            })
        if faltante > 1e-9:
            f = respaldo(sku)
            capas.append({
                "sku_madre": sku, "producto_madre": r.get("producto_madre", ""), "canal": canal,
                "componente": "FULL", "unidades": faltante, "fecha_referencia": f,
                "fuente": "APROX_FIFO_COMPRAS" if pd.notna(f) else "SIN_DATO",
                "referencia": f"Full sin envío registrado en {ANTIGUEDAD_LOOKBACK_ENVIOS_FULL_DIAS} días",
            })

    # 3) Tránsito (piezas reservadas en traslados Listo).
    for r in inv[inv["inventario_transito"] > 0].to_dict("records"):
        f = respaldo(r["sku_madre"])
        capas.append({
            "sku_madre": r["sku_madre"], "producto_madre": r.get("producto_madre", ""),
            "canal": r["canal"], "componente": "TRANSITO", "unidades": r["inventario_transito"],
            "fecha_referencia": f, "fuente": "APROX_FIFO_COMPRAS" if pd.notna(f) else "SIN_DATO",
            "referencia": "Traslado a Full en tránsito",
        })

    capas_df = pd.DataFrame(capas, columns=[c for c in cols_capas if c != "dias"])
    if capas_df.empty:
        return pd.DataFrame(), pd.DataFrame(columns=cols_capas), fifo_sku
    capas_df["producto_madre"] = capas_df["producto_madre"].fillna("")
    faltan_nombre = capas_df["producto_madre"].astype(str).str.strip().eq("")
    capas_df.loc[faltan_nombre, "producto_madre"] = capas_df.loc[faltan_nombre, "sku_madre"].map(nombres).fillna("")
    capas_df["fecha_referencia"] = pd.to_datetime(capas_df["fecha_referencia"], errors="coerce")
    capas_df["dias"] = (hoy - capas_df["fecha_referencia"].dt.normalize()).dt.days.clip(lower=0)
    capas_df = capas_df[cols_capas]

    # Resumen por SKU madre + canal.
    tmp = capas_df.copy()
    tmp["con_fecha"] = tmp["dias"].notna()
    tmp["u_fecha"] = np.where(tmp["con_fecha"], tmp["unidades"], 0.0)
    tmp["u_x_d"] = np.where(tmp["con_fecha"], tmp["unidades"] * tmp["dias"].fillna(0), 0.0)
    tmp["u_mas_umbral"] = np.where(tmp["dias"].fillna(-1) > umbral, tmp["unidades"], 0.0)
    tmp["u_mas_60"] = np.where(tmp["dias"].fillna(-1) > 60, tmp["unidades"], 0.0)
    tmp["u_mas_30"] = np.where(tmp["dias"].fillna(-1) > 30, tmp["unidades"], 0.0)
    tmp["fuente_legible"] = tmp["fuente"].map(FUENTES_ANTIGUEDAD_LEGIBLES).fillna(tmp["fuente"])
    tmp["fecha_envio_vigente"] = tmp["fecha_referencia"].where(tmp["fuente"].eq("ENVIO_FULL"))

    keys = ["sku_madre", "canal"]
    res = tmp.groupby(keys, as_index=False).agg(
        unidades_antiguedad=("unidades", "sum"),
        unidades_con_fecha=("u_fecha", "sum"),
        suma_u_x_d=("u_x_d", "sum"),
        dias_inventario_max=("dias", "max"),
        unidades_mas_90d=("u_mas_umbral", "sum"),
        unidades_mas_60d=("u_mas_60", "sum"),
        unidades_mas_30d=("u_mas_30", "sum"),
        fecha_envio_full_vigente_mas_antiguo=("fecha_envio_vigente", "min"),
    )
    res["dias_inventario_canal"] = np.where(
        res["unidades_con_fecha"] > 0, res["suma_u_x_d"] / res["unidades_con_fecha"].replace(0, np.nan), np.nan
    )
    res["unidades_sin_dato_antiguedad"] = (res["unidades_antiguedad"] - res["unidades_con_fecha"]).clip(lower=0)

    fuentes = tmp.groupby(keys + ["fuente_legible"], as_index=False)["unidades"].sum()
    fuentes["pct"] = fuentes["unidades"] / fuentes.groupby(keys)["unidades"].transform("sum").replace(0, np.nan)
    fuentes = fuentes.sort_values(keys + ["unidades"], ascending=[True, True, False])
    principal = fuentes.drop_duplicates(keys, keep="first")[keys + ["fuente_legible"]].rename(
        columns={"fuente_legible": "fuente_dias_inventario"}
    )
    detalle = fuentes.groupby(keys).apply(
        lambda g: " · ".join(f"{a} {b:.0%}" for a, b in zip(g["fuente_legible"], g["pct"].fillna(0)))
    ).reset_index(name="detalle_fuente_dias")
    res = res.merge(principal, on=keys, how="left").merge(detalle, on=keys, how="left")

    if not env.empty:
        ult = env.groupby(keys, as_index=False).agg(fecha_ultimo_envio_full=("fecha_envio", "max"))
        res = res.merge(ult, on=keys, how="left")
    else:
        res["fecha_ultimo_envio_full"] = pd.NaT

    if not fifo_sku.empty:
        res = res.merge(
            fifo_sku[["sku_madre", "fifo_compras_dias_prom", "fifo_compras_dias_max", "fifo_compras_estado"]].rename(
                columns={
                    "fifo_compras_dias_prom": "dias_desde_compra_sku",
                    "fifo_compras_dias_max": "dias_desde_compra_sku_max",
                    "fifo_compras_estado": "estado_fifo_compras_sku",
                }
            ),
            on="sku_madre", how="left",
        )
    res = res.drop(columns=["suma_u_x_d"])
    res["umbral_antiguedad_dias"] = umbral
    return res, capas_df, fifo_sku


# Columnas de antigüedad que produce construir_antiguedad_canal y que se
# pegan a inventario_canal_sku_madre (y viajan a 02 y 03).
COLUMNAS_ANTIGUEDAD_CANAL = [
    "unidades_antiguedad", "unidades_con_fecha", "dias_inventario_canal", "dias_inventario_max",
    "unidades_mas_90d", "unidades_mas_60d", "unidades_mas_30d", "unidades_sin_dato_antiguedad",
    "fuente_dias_inventario", "detalle_fuente_dias", "fecha_ultimo_envio_full",
    "fecha_envio_full_vigente_mas_antiguo", "dias_desde_compra_sku", "dias_desde_compra_sku_max",
    "estado_fifo_compras_sku", "umbral_antiguedad_dias",
]


# ============================================================
# COSTO UNITARIO Y VALOR DEL INVENTARIO POR SKU MADRE
# ============================================================

def construir_costos_sku_madre(inventario_canal, costos_productos_vinc, recepciones_vinc):
    """Un costo unitario por SKU madre, con su fuente, en este orden:

    1. ODOO_VALOR_QUANT     : valor contable de los quants CUATI / piezas.
    2. ODOO_AVG_COST        : costo promedio del producto (si el campo existe).
    3. ODOO_STANDARD_PRICE  : costo del producto en Odoo.
    4. ULTIMA_RECEPCION     : precio unitario de la última recepción de compra.

    Así los SKU que hoy solo tienen piezas en Full también quedan valuados.
    """
    columnas = ["sku_madre", "costo_unitario_inventario", "fuente_costo_inventario"]
    if inventario_canal is None or inventario_canal.empty:
        return pd.DataFrame(columns=columnas)

    quant = inventario_canal.groupby("sku_madre", as_index=False).agg(
        costo_unitario_cuati=("costo_unitario_cuati", "max")
    ) if "costo_unitario_cuati" in inventario_canal.columns else pd.DataFrame(columns=["sku_madre", "costo_unitario_cuati"])
    quant_map = dict(zip(quant["sku_madre"], pd.to_numeric(quant["costo_unitario_cuati"], errors="coerce").fillna(0)))

    prod_map = {}
    if costos_productos_vinc is not None and not costos_productos_vinc.empty:
        cp = costos_productos_vinc[
            costos_productos_vinc.get("tiene_referencia_madre", "").astype(str).eq("SI")
        ].copy()
        for c in ["avg_cost", "standard_price"]:
            cp[c] = pd.to_numeric(cp.get(c, 0), errors="coerce").fillna(0)
        cp["exacto"] = cp["sku_original"].astype(str).str.upper().eq(cp["sku_madre"].astype(str).str.upper())
        cp = cp.sort_values(["sku_madre", "exacto"], ascending=[True, False])
        for sku, g in cp.groupby("sku_madre"):
            prod_map[sku] = g

    rec_map = {}
    if recepciones_vinc is not None and not recepciones_vinc.empty:
        rc = recepciones_vinc[
            recepciones_vinc.get("tiene_referencia_madre", "").astype(str).eq("SI")
        ].copy()
        rc["precio_unitario_recepcion"] = pd.to_numeric(rc.get("precio_unitario_recepcion", 0), errors="coerce").fillna(0)
        rc["fecha_recepcion"] = pd.to_datetime(rc.get("fecha_recepcion"), errors="coerce")
        rc = rc[rc["precio_unitario_recepcion"] > 0].sort_values("fecha_recepcion")
        rec_map = rc.groupby("sku_madre")["precio_unitario_recepcion"].last().to_dict()

    rows = []
    for sku in sorted(inventario_canal["sku_madre"].dropna().unique()):
        costo, fuente = float(quant_map.get(sku, 0) or 0), "ODOO_VALOR_QUANT"
        if costo <= 0 and sku in prod_map:
            g = prod_map[sku]
            exactos = g[g["exacto"]]
            base = exactos if not exactos.empty else g
            avg = base.loc[base["avg_cost"] > 0, "avg_cost"]
            std = base.loc[base["standard_price"] > 0, "standard_price"]
            if not avg.empty:
                costo, fuente = float(avg.median()), "ODOO_AVG_COST"
            elif not std.empty:
                costo, fuente = float(std.median()), "ODOO_STANDARD_PRICE"
        if costo <= 0 and sku in rec_map:
            costo, fuente = float(rec_map[sku]), "ULTIMA_RECEPCION_COMPRA"
        if costo <= 0:
            costo, fuente = 0.0, "SIN_COSTO"
        rows.append({"sku_madre": sku, "costo_unitario_inventario": costo, "fuente_costo_inventario": fuente})
    return pd.DataFrame(rows, columns=columnas)


def agregar_valor_inventario(inventario_canal, costos_sku):
    """Pega costo y calcula el valor (inversión) de cada bolsa del canal."""
    out = inventario_canal.copy()
    if costos_sku is not None and not costos_sku.empty:
        out = out.drop(columns=[c for c in ["costo_unitario_inventario", "fuente_costo_inventario"] if c in out.columns])
        out = out.merge(costos_sku, on="sku_madre", how="left")
    out["costo_unitario_inventario"] = pd.to_numeric(out.get("costo_unitario_inventario", 0), errors="coerce").fillna(0)
    out["fuente_costo_inventario"] = out.get("fuente_costo_inventario", "SIN_COSTO")
    out["fuente_costo_inventario"] = out["fuente_costo_inventario"].fillna("SIN_COSTO")
    costo = out["costo_unitario_inventario"]
    for comp in ["odoo", "transito", "full"]:
        out[f"valor_{comp}"] = pd.to_numeric(out.get(f"inventario_{comp}", 0), errors="coerce").fillna(0) * costo
    out["valor_inventario_canal"] = pd.to_numeric(out.get("inventario_total", 0), errors="coerce").fillna(0) * costo
    out["valor_mas_90d"] = pd.to_numeric(out.get("unidades_mas_90d", 0), errors="coerce").fillna(0) * costo
    return out


# ============================================================
# REGISTRO DE MOVIMIENTOS ODOO (AUDITORÍA DE ANTIGÜEDAD)
# ============================================================

TIPOS_SALIDA_AUDITORIA = {
    "REASIGNACION_ENTRE_CANALES", "SALIDA_POR_AJUSTE",
    "SALIDA_A_UBICACION_NO_COMERCIAL", "ENVIO_A_FULL",
}
TIPOS_ENTRADA_AUDITORIA = {
    "ENTRADA_POR_AJUSTE", "ENTRADA_DESDE_UBICACION_NO_COMERCIAL", "DEVOLUCION_CLIENTE",
}
NIVEL_ORDEN = {"BAJA": 0, "MEDIA": 1, "ALTA": 2}


def _envios_como_movimientos(envios_full_vinc, desde):
    """Convierte los envíos Full de la ventana al formato de auditoría."""
    if envios_full_vinc is None or envios_full_vinc.empty:
        return pd.DataFrame()
    e = envios_full_vinc.copy()
    e["fecha_envio"] = pd.to_datetime(e["fecha_envio"], errors="coerce")
    e = e[e["fecha_envio"] >= desde]
    if e.empty:
        return pd.DataFrame()
    return pd.DataFrame({
        "move_line_id": e["move_line_id"],
        "fecha_movimiento": e["fecha_envio"],
        "referencia": e["referencia"],
        "documento_origen": "",
        "contacto": e["contacto"],
        "tipo_operacion": "Envío a Full",
        "product_id": e["product_id"],
        "sku_original": e["sku_original"],
        "producto": e["producto"],
        "lote_id_odoo": e["lote_id_odoo"],
        "lote_nombre": e["lote_nombre"],
        "cantidad": e["cantidad_enviada"],
        "ubicacion_origen": e["ubicacion_origen"],
        "ubicacion_destino": "Full " + e["canal"].astype(str),
        "tipo_origen": "CANAL",
        "tipo_destino": "FULL",
        "canal_origen": e["canal_origen"],
        "canal_destino": "Full " + e["canal"].astype(str),
        "tipo_movimiento": "ENVIO_A_FULL",
        "usuario": e["usuario"],
        "creado_por": "",
        "responsable": e["responsable"],
        "sku_madre": e.get("sku_madre", ""),
        "producto_madre": e.get("producto_madre", ""),
        "tiene_referencia_madre": e.get("tiene_referencia_madre", ""),
    })


def _marcar_idas_y_vueltas(reg, ventana):
    """Marca pares A->B / B->A y ajuste salida/entrada del mismo producto.

    Cada salida se empareja con UNA sola regresión (la primera dentro de la
    ventana que no se haya usado) y, si ambos movimientos traen lote, debe
    ser el mismo lote. Antes se emparejaba cada salida con todas las
    regresiones del producto, lo que en productos con traslados frecuentes
    marcaba casi todo como revisión ALTA.
    """
    motivos = {i: [] for i in reg.index}
    pares = [
        ("REASIGNACION_ENTRE_CANALES", "REASIGNACION_ENTRE_CANALES", True),
        ("SALIDA_POR_AJUSTE", "ENTRADA_POR_AJUSTE", False),
        ("SALIDA_A_UBICACION_NO_COMERCIAL", "ENTRADA_DESDE_UBICACION_NO_COMERCIAL", False),
    ]
    base = reg[reg["fecha_movimiento"].notna()].copy()
    base["lote_par"] = pd.to_numeric(base.get("lote_id_odoo"), errors="coerce") if "lote_id_odoo" in base.columns else np.nan
    for _, g in base.groupby("product_id"):
        if len(g) < 2:
            continue
        g = g.sort_values("fecha_movimiento")
        filas = list(g[["fecha_movimiento", "tipo_movimiento", "canal_origen", "canal_destino", "lote_par"]].itertuples())
        usados = set()
        for a_i, a in enumerate(filas):
            for tipo_ida, tipo_vuelta, cruzado in pares:
                if a.tipo_movimiento != tipo_ida:
                    continue
                if cruzado and not ({a.canal_origen, a.canal_destino} & CANALES_CON_KAM):
                    continue
                for b in filas[a_i + 1:]:
                    dias = (b.fecha_movimiento - a.fecha_movimiento).days
                    if dias > ventana:
                        break
                    if b.Index in usados or b.tipo_movimiento != tipo_vuelta:
                        continue
                    if cruzado and not (b.canal_origen == a.canal_destino and b.canal_destino == a.canal_origen):
                        continue
                    if not cruzado and (a.canal_origen or b.canal_destino) and a.canal_origen != b.canal_destino:
                        continue
                    if pd.notna(a.lote_par) and pd.notna(b.lote_par) and a.lote_par != b.lote_par:
                        continue
                    texto = (f"Ida y vuelta {a.canal_origen or 'bodega'} → "
                             f"{a.canal_destino or 'fuera'} → {b.canal_destino or 'bodega'} en {dias} días")
                    motivos[a.Index].append(texto)
                    motivos[b.Index].append(texto)
                    usados.add(b.Index)
                    break
    return motivos


def evaluar_movimientos_auditoria(reg):
    """Asigna nivel de revisión (ALTA/MEDIA/BAJA) y motivo a cada movimiento."""
    if reg.empty:
        return reg
    reg = reg.copy()
    reg["fecha_movimiento"] = pd.to_datetime(reg["fecha_movimiento"], errors="coerce")
    ida_vuelta = _marcar_idas_y_vueltas(reg, AUDITORIA_VENTANA_IDA_VUELTA_DIAS)
    umbral = AUDITORIA_UMBRAL_EDAD_SOSPECHA
    niveles, textos = [], []
    for idx, r in reg.iterrows():
        tipo = r.get("tipo_movimiento", "")
        edad = pd.to_numeric(r.get("edad_lote_al_movimiento"), errors="coerce")
        nivel, motivos = "BAJA", []
        # La edad solo cuenta cuando las piezas SALEN de un canal con KAM:
        # mover inventario viejo de General hacia un canal no le baja la
        # antigüedad a nadie (se la sube al canal que lo recibe).
        if (tipo in TIPOS_SALIDA_AUDITORIA and r.get("canal_origen") in CANALES_CON_KAM
                and pd.notna(edad) and edad >= umbral):
            motivos.append(f"Piezas con {int(edad)} días de antigüedad al salir de {r.get('canal_origen')} (umbral {umbral})")
            nivel = "MEDIA" if tipo == "ENVIO_A_FULL" else "ALTA"
        if ida_vuelta.get(idx):
            motivos.extend(sorted(set(ida_vuelta[idx])))
            nivel = "ALTA"
        if tipo == "REASIGNACION_ENTRE_CANALES" and r.get("canal_origen") in CANALES_CON_KAM \
                and r.get("canal_destino") not in CANALES_CON_KAM:
            motivos.append(f"Sale de {r.get('canal_origen')} hacia {r.get('canal_destino') or 'otra bolsa'}")
            nivel = max(nivel, "MEDIA", key=NIVEL_ORDEN.get)
        if tipo in {"SALIDA_POR_AJUSTE", "SALIDA_A_UBICACION_NO_COMERCIAL"} and r.get("canal_origen") in CANALES_CON_KAM:
            motivos.append("Saca inventario de un canal con KAM sin venderlo")
            nivel = max(nivel, "MEDIA", key=NIVEL_ORDEN.get)
        if tipo == "DEVOLUCION_CLIENTE" and not motivos:
            motivos.append("Informativo: devolución que regresa a bodega")
        if not motivos:
            motivos.append("Sin señales de riesgo")
        niveles.append(nivel)
        textos.append(" | ".join(motivos))
    reg["nivel_revision"] = niveles
    reg["motivo_revision"] = textos
    # Canal cuyo KAM se ve afectado: en salidas el que pierde piezas, en
    # entradas el que las recibe. En una reasignación se atribuye al canal
    # con KAM (si solo uno de los dos lo tiene, es a quien le cambia la métrica).
    entrada = reg["tipo_movimiento"].isin(list(TIPOS_ENTRADA_AUDITORIA))
    reasig = reg["tipo_movimiento"].eq("REASIGNACION_ENTRE_CANALES")
    origen_kam = reg["canal_origen"].isin(list(CANALES_CON_KAM))
    destino_kam = reg["canal_destino"].isin(list(CANALES_CON_KAM))
    reg["canal_kam_afectado"] = np.where(
        entrada | (reasig & ~origen_kam & destino_kam), reg["canal_destino"], reg["canal_origen"]
    )
    return reg


def procesar_auditoria_movimientos(mov_vinc, envios_full_vinc, fechas_lote, costos_sku, ruta_registro):
    """Une movimientos + envíos Full, calcula edad y valor, y actualiza el registro.

    El registro (CSV) es ACUMULATIVO: cada corrida agrega los movimientos
    nuevos por move_line_id y conserva los anteriores aunque ya salgan de la
    ventana de consulta. Así queda evidencia aunque después se reviertan.
    """
    ahora = pd.Timestamp.now()
    desde = ahora - pd.Timedelta(days=AUDITORIA_MOVIMIENTOS_DIAS)
    partes = [p for p in [mov_vinc, _envios_como_movimientos(envios_full_vinc, desde)]
              if p is not None and not p.empty]
    nuevos = pd.concat(partes, ignore_index=True, sort=False) if partes else pd.DataFrame()

    if not nuevos.empty:
        nuevos["fecha_movimiento"] = pd.to_datetime(nuevos["fecha_movimiento"], errors="coerce")
        if fechas_lote is not None and not fechas_lote.empty:
            fl = fechas_lote[["lote_id_odoo", "fecha_recepcion_lote", "fuente_recepcion_lote"]].copy()
            fl["lote_id_odoo"] = pd.to_numeric(fl["lote_id_odoo"], errors="coerce")
            nuevos["lote_id_odoo"] = pd.to_numeric(nuevos["lote_id_odoo"], errors="coerce")
            nuevos = nuevos.merge(fl.drop_duplicates("lote_id_odoo"), on="lote_id_odoo", how="left")
        else:
            nuevos["fecha_recepcion_lote"] = pd.NaT
            nuevos["fuente_recepcion_lote"] = ""
        nuevos["fecha_recepcion_lote"] = pd.to_datetime(nuevos["fecha_recepcion_lote"], errors="coerce")
        nuevos["edad_lote_al_movimiento"] = (
            nuevos["fecha_movimiento"] - nuevos["fecha_recepcion_lote"]
        ).dt.days
        if costos_sku is not None and not costos_sku.empty:
            nuevos = nuevos.merge(costos_sku, on="sku_madre", how="left")
        nuevos["costo_unitario_inventario"] = pd.to_numeric(
            nuevos.get("costo_unitario_inventario", 0), errors="coerce").fillna(0)
        nuevos["cantidad"] = pd.to_numeric(nuevos["cantidad"], errors="coerce").fillna(0)
        nuevos["valor_movido"] = nuevos["cantidad"] * nuevos["costo_unitario_inventario"]
        nuevos["primera_deteccion"] = ahora.strftime("%Y-%m-%d %H:%M:%S")

    historico = pd.DataFrame()
    if ruta_registro.exists():
        try:
            historico = pd.read_csv(ruta_registro, encoding="utf-8-sig", low_memory=False)
        except Exception as e:
            print(f"ADVERTENCIA: no pude leer el registro anterior {ruta_registro.name}: {e}")

    if not historico.empty and not nuevos.empty:
        historico["move_line_id"] = pd.to_numeric(historico["move_line_id"], errors="coerce")
        nuevos["move_line_id"] = pd.to_numeric(nuevos["move_line_id"], errors="coerce")
        # Lo nuevo actualiza datos (costo, sku madre); se conserva la primera detección.
        primera = historico.set_index("move_line_id")["primera_deteccion"].to_dict() \
            if "primera_deteccion" in historico.columns else {}
        nuevos["primera_deteccion"] = nuevos["move_line_id"].map(primera).fillna(nuevos["primera_deteccion"])
        registro = pd.concat(
            [historico[~historico["move_line_id"].isin(nuevos["move_line_id"])], nuevos],
            ignore_index=True, sort=False,
        )
    else:
        registro = nuevos if not nuevos.empty else historico

    if registro.empty:
        return pd.DataFrame(), pd.DataFrame()

    registro = registro.drop(columns=[c for c in ["nivel_revision", "motivo_revision", "canal_kam_afectado"]
                                      if c in registro.columns])
    for c in ["canal_origen", "canal_destino", "tipo_movimiento", "usuario"]:
        registro[c] = registro.get(c, "").fillna("").astype(str)
    registro = evaluar_movimientos_auditoria(registro)
    registro["ultima_actualizacion_registro"] = ahora.strftime("%Y-%m-%d %H:%M:%S")
    _alta = registro[registro["nivel_revision"].eq("ALTA")]
    if not _alta.empty:
        _mot = _alta["motivo_revision"].astype(str)
        _cat = np.select(
            [_mot.str.contains("Ida y vuelta") & _mot.str.contains("antigüedad"),
             _mot.str.contains("Ida y vuelta"), _mot.str.contains("antigüedad")],
            ["edad + ida y vuelta", "solo ida y vuelta", "solo edad al salir"], default="otro",
        )
        _desglose = pd.Series(_cat).value_counts()
        print(f"Auditoría · revisión ALTA {len(_alta)} de {len(registro)}: "
              + ", ".join(f"{k} {v}" for k, v in _desglose.items())
              + " · por tipo: " + ", ".join(f"{k} {v}" for k, v in _alta["tipo_movimiento"].value_counts().head(4).items()))
    registro = registro.sort_values("fecha_movimiento", ascending=False)

    try:
        preparar_para_excel(registro).to_csv(ruta_registro, index=False, encoding="utf-8-sig")
        print(f"Registro de movimientos actualizado: {ruta_registro.name} ({len(registro)} movimientos acumulados)")
    except Exception as e:
        print(f"ADVERTENCIA: no pude guardar el registro de movimientos: {e}")

    limite = ahora - pd.Timedelta(days=AUDITORIA_DIAS_EN_DASHBOARD)
    visible = registro[pd.to_datetime(registro["fecha_movimiento"], errors="coerce") >= limite].copy()

    resumen = pd.DataFrame()
    if not visible.empty:
        visible["es_alta"] = visible["nivel_revision"].eq("ALTA").astype(int)
        visible["es_media"] = visible["nivel_revision"].eq("MEDIA").astype(int)
        resumen = visible.groupby(
            ["canal_kam_afectado", "usuario", "tipo_movimiento"], as_index=False, dropna=False
        ).agg(
            movimientos=("move_line_id", "nunique"),
            piezas=("cantidad", "sum"),
            valor_movido=("valor_movido", "sum"),
            revision_alta=("es_alta", "sum"),
            revision_media=("es_media", "sum"),
            primer_movimiento=("fecha_movimiento", "min"),
            ultimo_movimiento=("fecha_movimiento", "max"),
        ).sort_values(["revision_alta", "valor_movido"], ascending=[False, False])
        visible = visible.drop(columns=["es_alta", "es_media"])
    return visible, resumen


# ============================================================
# AUTOAZUR: PUBLICACIONES POR CANAL (versión 2026-09)
# ============================================================
#
# Cruza las publicaciones de AutoAzur (GET /item/listings) contra el stock
# asignado a cada canal para responder: "¿qué SKU tienen piezas en el canal
# pero no están publicados, o su publicación está pausada / en revisión /
# con error?".
#
# La documentación de AutoAzur no publica la estructura de la respuesta de
# /item/listings, así que la lectura es tolerante a varios nombres de campo
# (Sku/SellerSku, Status/ListingStatus, etc.). En cada corrida se guarda una
# muestra cruda en autoazur_publicaciones_muestra.json para verificar el
# mapeo; si algún campo sale vacío, basta con agregar su nombre real a la
# lista de candidatos correspondiente en AA_CAMPOS.

def _env_bool(nombre, default="1"):
    return os.getenv(nombre, default).strip().lower() in {"1", "true", "si", "sí", "yes"}


ENABLE_AUTOAZUR_PUBLICACIONES = _env_bool("ENABLE_AUTOAZUR_PUBLICACIONES", "1")
AUTOAZUR_BASE_URL = os.getenv("AUTOAZUR_BASE_URL", "https://api.autoazur.com").strip().rstrip("/")
# El GUID NO se escribe en el código: va en .env como AUTOAZUR_UNIQUE_GUID.
# Se busca en el .env junto al script y también en el .env de la carpeta de
# datos (DASHBOARD_ROTACION_DIR), que es donde el 00 escribe sus variables.
_ENVS_REVISADOS = []
try:
    from dotenv import load_dotenv as _load_dotenv
    for _env in [Path(__file__).resolve().parent / ".env", CARPETA_SALIDA / ".env"]:
        _env = _env.resolve()
        if _env not in _ENVS_REVISADOS:
            _ENVS_REVISADOS.append(_env)
            if _env.exists():
                _load_dotenv(_env, override=False)
except ImportError:
    pass
# Se aceptan nombres alternativos por si el .env ya tenía el GUID con otro nombre.
AUTOAZUR_VARIABLES_GUID = ["AUTOAZUR_UNIQUE_GUID", "AUTOAZUR_GUID", "UNIQUE_GUID"]
AUTOAZUR_UNIQUE_GUID = next(
    (os.getenv(v, "").strip().strip('"').strip("'") for v in AUTOAZUR_VARIABLES_GUID
     if os.getenv(v, "").strip().strip('"').strip("'")),
    "",
)
# Nombres de canal tal como los acepta AutoAzur (los mismos del reporte de ventas).
AUTOAZUR_CANALES_PUBLICACIONES = [
    c.strip() for c in os.getenv(
        "AUTOAZUR_CANALES_PUBLICACIONES",
        "MERCADO LIBRE,AMAZON,WALMART,LIVERPOOL,COPPEL,ELEKTRA,TIK TOK",
    ).split(",") if c.strip()
]
AUTOAZUR_PER_PAGE = min(int(os.getenv("AUTOAZUR_PER_PAGE", "500")), 500)  # máximo de la API
AUTOAZUR_MAX_PAGINAS = int(os.getenv("AUTOAZUR_MAX_PAGINAS", "400"))
AUTOAZUR_TIMEOUT = int(os.getenv("AUTOAZUR_TIMEOUT", "180"))
AUTOAZUR_ARCHIVO_MUESTRA = "autoazur_publicaciones_muestra.json"

# Candidatos de nombre por campo (se buscan sin distinguir mayúsculas; el
# punto indica un objeto anidado, p. ej. "Account.Name").
AA_CAMPOS = {
    "sku_publicacion": ["Sku", "SKU", "SellerSku", "SellerSKU", "SkuSeller", "SellerCustomField", "ProductSku"],
    "item_id": ["ItemID", "ItemId", "ListingID", "ListingId", "PublicationID", "PublicationId", "ASIN", "ID", "Id"],
    "titulo": ["Title", "Titulo", "Name", "ProductName"],
    "upc": ["UPC", "EAN", "Barcode", "GTIN"],
    "variation_id": ["VariationID", "VariationId"],
    "estado_original": ["Status", "Estatus", "Estado", "ListingStatus", "PublicationStatus", "ItemStatus",
                        "StatusName", "StatusListing", "ListingState", "PublicationState", "MarketplaceStatus",
                        "ChannelStatus", "State", "Active", "IsActive", "Enabled", "IsEnabled", "Activo",
                        "Published", "IsPublished", "Paused", "IsPaused"],
    "subestado": ["SubStatus", "Sub_Status", "Substatus", "StatusDetail", "StatusDescription", "StatusReason", "Reason"],
    "cuenta": ["AccountName", "Account.Name", "Nickname", "NickName", "UserName", "Account"],
    "user_id": ["UserID", "UserId", "AccountID", "AccountId", "SellerID", "SellerId"],
    "precio": ["Price", "Precio", "SalePrice"],
    "stock_publicado": ["Stock", "AvailableQuantity", "Available_Quantity", "Quantity", "Inventory", "Existencias"],
    "tipo_publicacion": ["ListingType", "ListingTypeID", "Type"],
    "logistica": ["LogisticType", "Logistic", "Fulfillment", "ShippingMode", "ShippingType"],
    "permalink": ["Permalink", "Url", "URL", "Link"],
    "canal_api": ["Channel", "ChannelName", "Channel.Name"],
    "fecha_actualizacion": ["LastUpdated", "LastUpdate", "UpdateDate", "ModifiedDate", "UpdatedAt"],
}
AA_LLAVES_LISTA = ["Items", "Listings", "Publications", "Publicaciones", "Products", "Data", "Results", "Message", "Response"]

# Estados normalizados y su prioridad al resumir varias publicaciones de un
# mismo SKU en un canal (se muestra el MEJOR estado: basta una activa).
ESTADOS_PUBLICACION_PRIORIDAD = {
    "Activa": 1,
    "En revisión": 2,
    "Pausada · sin stock": 3,
    "Pausada": 4,
    "Con error / rechazada": 5,
    "Inactiva": 6,
    "Cerrada / eliminada": 7,
    "Otro": 8,
    "Sin estado": 9,
}


# Nombres genéricos que NO se buscan dentro de objetos anidados (p. ej. el
# "Id" de la cuenta no debe confundirse con el ItemID de la publicación).
AA_GENERICOS_NO_ANIDADOS = {"id", "name", "type", "account", "channel", "status", "state"}
# Envolturas típicas de un registro ({"Publication": {...}}).
AA_ENVOLTURAS = ["Item", "Listing", "Publication", "Publicacion", "Product", "Producto", "Data"]


def aa_campos_efectivos():
    """Candidatos por campo + los que se agreguen en .env.

    Ejemplo en .env:  AUTOAZUR_CAMPO_SKU_PUBLICACION=SkuVendedor
    (se pueden poner varios separados por coma). Así se corrige el mapeo
    sin tocar el código.
    """
    out = {}
    for campo, cands in AA_CAMPOS.items():
        extra = os.getenv(f"AUTOAZUR_CAMPO_{campo.upper()}", "")
        out[campo] = [e.strip() for e in extra.split(",") if e.strip()] + list(cands)
    return out


def _aa_normalizar_valor(val):
    if isinstance(val, dict):
        val, _ = _aa_buscar(val, ["Name", "Nombre", "Description", "ID", "Id"])
    if isinstance(val, list) and val and all(not isinstance(x, (dict, list)) for x in val):
        val = ", ".join(str(x) for x in val)
    if isinstance(val, list):
        return None
    return val


def _aa_buscar(registro, candidatos, profundidad=2):
    """Busca un campo por varios nombres (sin distinguir mayúsculas).

    Devuelve (valor, nombre_de_la_llave_usada). Primero en el nivel actual
    (incluidas rutas con punto como "Account.Name"); si no aparece, dentro
    de objetos anidados, salvo para nombres genéricos como "Id".
    """
    if not isinstance(registro, dict):
        return None, ""
    mapa = {str(k).lower(): (k, v) for k, v in registro.items()}
    for cand in candidatos:
        partes = cand.lower().split(".")
        par = mapa.get(partes[0])
        if par is None:
            continue
        clave, val = par
        for parte in partes[1:]:
            if isinstance(val, dict):
                sub = {str(k).lower(): v for k, v in val.items()}.get(parte)
                val, clave = sub, f"{clave}.{parte}"
            else:
                val = None
        val = _aa_normalizar_valor(val)
        if val not in (None, "", [], {}):
            return val, str(clave)
    if profundidad > 0:
        anidables = [c for c in candidatos if c.lower() not in AA_GENERICOS_NO_ANIDADOS and "." not in c]
        for k, v in registro.items():
            if isinstance(v, dict) and anidables:
                val, clave = _aa_buscar(v, anidables, profundidad - 1)
                if val not in (None, "", [], {}):
                    return val, f"{k}.{clave}"
    return None, ""


def _aa_valor(registro, candidatos):
    return _aa_buscar(registro, candidatos)[0]


def _aa_extraer_lista(data, profundidad=0):
    """Devuelve la lista de publicaciones dentro de la respuesta, esté donde esté."""
    if isinstance(data, list):
        return [x for x in data if isinstance(x, dict)]
    if not isinstance(data, dict) or profundidad > 3:
        return []
    mapa = {str(k).lower(): v for k, v in data.items()}
    for llave in AA_LLAVES_LISTA:
        val = mapa.get(llave.lower())
        if isinstance(val, list) and (not val or isinstance(val[0], dict)):
            return [x for x in val if isinstance(x, dict)]
        if isinstance(val, dict):
            sub = _aa_extraer_lista(val, profundidad + 1)
            if sub:
                return sub
    # La lista más larga de objetos que haya en la respuesta.
    mejor = []
    for val in data.values():
        cand = _aa_extraer_lista(val, profundidad + 1) if isinstance(val, dict) else (
            [x for x in val if isinstance(x, dict)] if isinstance(val, list) else [])
        if len(cand) > len(mejor):
            mejor = cand
    return mejor


def _aa_totales(data, profundidad=0):
    """(total_publicaciones, total_paginas) si la respuesta los informa.

    Antes se tomaba el primer campo con "total" en el nombre, que podía ser
    TotalPages; por eso los canales salían "INCOMPLETO".
    """
    total_items, total_paginas = None, None
    if not isinstance(data, dict) or profundidad > 3:
        return None, None
    for k, v in data.items():
        kl = str(k).lower()
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            if "page" in kl or "pagina" in kl:
                if ("total" in kl or kl in ("pages", "pagecount", "paginas")) and total_paginas is None:
                    total_paginas = int(v)
            elif ("total" in kl or kl in ("count", "records", "recordcount")) and total_items is None:
                total_items = int(v)
        elif isinstance(v, dict):
            ti, tp = _aa_totales(v, profundidad + 1)
            total_items = total_items if total_items is not None else ti
            total_paginas = total_paginas if total_paginas is not None else tp
    return total_items, total_paginas


AA_ITEM_ID_ESPECIFICOS = [c for c in AA_CAMPOS["item_id"] if c.lower() not in AA_GENERICOS_NO_ANIDADOS]


def _aa_parece_publicacion(d, campos):
    """Un objeto es publicación si trae SKU o un ItemID específico (no un "Id" genérico)."""
    if not isinstance(d, dict):
        return False
    if _aa_buscar(d, campos["sku_publicacion"], profundidad=0)[0] not in (None, ""):
        return True
    especificos = [c for c in campos["item_id"] if c.lower() not in AA_GENERICOS_NO_ANIDADOS]
    return _aa_buscar(d, especificos, profundidad=0)[0] not in (None, "")


def _aa_registros(data, campos, contexto=None, en_lista=False, profundidad=0):
    """Aplana la respuesta de /item/listings en una lista de publicaciones.

    AutoAzur puede agrupar las publicaciones dentro de contenedores (por
    cuenta, por canal, etc.). Antes el contenedor se tomaba como UNA
    publicación y la paginación se detenía en la primera página. Ahora se
    baja hasta cada publicación y los datos simples de los contenedores que
    vienen en listas (p. ej. la cuenta) se heredan como contexto. El sobre
    de la respuesta (Status/Message del nivel superior) NO se hereda.
    """
    contexto = contexto or {}
    if profundidad > 8:
        return []
    if isinstance(data, list):
        out = []
        for x in data:
            out.extend(_aa_registros(x, campos, contexto, True, profundidad + 1))
        return out
    if not isinstance(data, dict):
        return []
    if _aa_parece_publicacion(data, campos):
        rec = dict(contexto)
        rec.update(data)
        return [rec]
    ctx = dict(contexto)
    if en_lista:
        for k, v in data.items():
            if not isinstance(v, (dict, list)) and k not in ctx:
                ctx[k] = v
    out = []
    for v in data.values():
        if isinstance(v, (dict, list)):
            out.extend(_aa_registros(v, campos, ctx, False, profundidad + 1))
    return out


def _aa_desenvolver(rec, campos):
    """{"Publication": {...}} -> {...}; conserva los campos del nivel superior."""
    for _ in range(3):
        if not isinstance(rec, dict):
            return {}
        tiene_id = any(_aa_buscar(rec, campos[c], profundidad=0)[0] not in (None, "")
                       for c in ("sku_publicacion", "item_id", "estado_original"))
        if tiene_id:
            return rec
        mapa = {str(k).lower(): k for k in rec}
        envoltura = next((mapa[e.lower()] for e in AA_ENVOLTURAS
                          if e.lower() in mapa and isinstance(rec[mapa[e.lower()]], dict)), None)
        if envoltura is None and len(rec) == 1 and isinstance(next(iter(rec.values())), dict):
            envoltura = next(iter(rec))
        if envoltura is None:
            return rec
        interior = dict(rec[envoltura])
        for k, v in rec.items():
            if k != envoltura and k not in interior:
                interior[k] = v
        rec = interior
    return rec


def _aa_expandir_variaciones(rec, campos):
    """Una publicación con variaciones (cada una con su SKU) -> una fila por variación."""
    cand_sku = {c.lower() for c in campos["sku_publicacion"]}
    for k, v in rec.items():
        if isinstance(v, list) and v and all(isinstance(x, dict) for x in v):
            if any(any(str(ck).lower() in cand_sku for ck in x) for x in v):
                padre = {pk: pv for pk, pv in rec.items() if pk != k}
                filas = []
                for x in v:
                    fila = dict(padre)
                    fila.update({ck: cv for ck, cv in x.items() if cv not in (None, "", [])})
                    filas.append(fila)
                return filas
    return [rec]


def normalizar_estado_publicacion(estado, subestado=""):
    """Traduce el estatus del marketplace a un estado común en español.

    El orden importa: "inactive" contiene "active", y una publicación
    "paused / out_of_stock" se reporta aparte porque suele significar que
    AutoAzur mandó stock 0 aunque haya piezas asignadas al canal.
    """
    t = normalizar_texto(f"{estado or ''} {subestado or ''}").replace("_", " ")
    if not t.strip():
        return "Sin estado"
    if any(k in t for k in ("review", "revision", "moderat", "pending", "pendiente",
                            "en proceso", "processing", "validac", "in progress")):
        return "En revisión"
    if any(k in t for k in ("pause", "pausad")):
        if any(k in t for k in ("out of stock", "sin stock", "sin existencia", "no stock")):
            return "Pausada · sin stock"
        return "Pausada"
    if any(k in t for k in ("error", "rechaz", "reject", "suppress", "incomplet",
                            "invalid", "bloque", "block", "denied")):
        return "Con error / rechazada"
    if any(k in t for k in ("closed", "cerrad", "finaliz", "ended", "delet", "elimin", "baja")):
        return "Cerrada / eliminada"
    if any(k in t for k in ("inactiv", "unpublish", "no publicad", "not published",
                            "not live", "disabled", "deshabilit")):
        return "Inactiva"
    if any(k in t for k in ("active", "activ", "publicad", "published", "live",
                            "buyable", "vigente", "enabled", "habilitad")):
        return "Activa"
    return "Otro"


class ClienteAutoAzur:
    """Cliente mínimo de AutoAzur con reintentos y renovación de token.

    OJO: AutoAzur revoca el token anterior cada vez que se genera uno nuevo.
    Si el reporte de ventas corre al mismo tiempo que este script, uno de
    los dos perderá su token; conviene ejecutarlos en secuencia.
    """

    def __init__(self, guid, base_url=AUTOAZUR_BASE_URL, session=None):
        if session is None:
            import requests
            session = requests.Session()
        self.s = session
        self.guid = guid
        self.base = base_url
        self.token = None
        import threading
        self._lock = threading.Lock()

    def renovar_token(self):
        r = self.s.post(f"{self.base}/token/redeem", json={"UniqueGUID": self.guid},
                        headers={"Content-Type": "application/json"}, timeout=AUTOAZUR_TIMEOUT)
        r.raise_for_status()
        data = r.json()
        token = _aa_valor(data.get("Message", {}) if isinstance(data.get("Message"), dict) else data,
                          ["Token", "AccessToken"])
        if not token:
            raise RuntimeError("AutoAzur no devolvió token en /token/redeem")
        self.token = token

    def get(self, ruta, params):
        with self._lock:
            if not self.token:
                self.renovar_token()
        ultimo_error = None
        renovado = False
        for intento in range(4):
            try:
                token_usado = self.token
                r = self.s.get(f"{self.base}{ruta}", headers={"Authorization": token_usado},
                               params=params, timeout=AUTOAZUR_TIMEOUT)
                if r.status_code == 401 and not renovado:
                    # Con varias consultas en paralelo solo un hilo renueva;
                    # los demás reutilizan el token nuevo (renovar de más
                    # revocaría el token de los otros hilos).
                    with self._lock:
                        if self.token == token_usado:
                            self.renovar_token()
                    renovado = True
                    continue
                if r.status_code in (429, 500, 502, 503, 504):
                    ultimo_error = f"HTTP {r.status_code}"
                    import time
                    time.sleep(3 * (intento + 1))
                    continue
                if r.status_code == 204:
                    return {}
                r.raise_for_status()
                return r.json()
            except Exception as e:  # red, JSON inválido, etc.
                ultimo_error = str(e)
                import time
                time.sleep(3 * (intento + 1))
        raise RuntimeError(f"{ruta} {params}: {ultimo_error}")


def _aa_huella(rec):
    """Identificador de un registro crudo: detecta páginas repetidas SIN
    confundir publicaciones distintas (antes, si no se detectaban SKU e
    ItemID, todas se veían iguales y solo se conservaba la primera)."""
    import hashlib
    import json
    return hashlib.md5(json.dumps(rec, sort_keys=True, default=str).encode("utf-8")).hexdigest()


def extraer_publicaciones_autoazur(cliente=None, ruta_muestra=None):
    """Descarga todas las publicaciones por canal. Devuelve (detalle, log)."""
    columnas = ["canal_autoazur", "canal"] + list(AA_CAMPOS.keys()) + ["estado_publicacion"]
    log = []
    if cliente is None:
        if not ENABLE_AUTOAZUR_PUBLICACIONES:
            print("AutoAzur publicaciones: desactivado (ENABLE_AUTOAZUR_PUBLICACIONES=0).")
            return pd.DataFrame(columns=columnas), pd.DataFrame(log)
        if not AUTOAZUR_UNIQUE_GUID:
            revisados = " | ".join(
                f"{e} ({'existe' if e.exists() else 'NO existe'})" for e in _ENVS_REVISADOS
            )
            detalle = (f"No se encontró {AUTOAZUR_VARIABLES_GUID[0]} en ningún .env. "
                       f"Archivos revisados: {revisados}")
            print("\n" + "!" * 72)
            print("AutoAzur publicaciones: NO se consultó. " + detalle)
            print(f"Agrega esta línea a uno de esos .env:  {AUTOAZUR_VARIABLES_GUID[0]}=<tu GUID>")
            print("!" * 72 + "\n")
            return pd.DataFrame(columns=columnas), pd.DataFrame([{
                "canal_autoazur": "TODOS", "canal": "TODOS", "estado_consulta": "SIN_CREDENCIAL",
                "publicaciones": 0, "paginas": 0, "detalle": detalle,
            }])
        print(f"AutoAzur publicaciones: GUID detectado (termina en ...{AUTOAZUR_UNIQUE_GUID[-4:]}).")
        cliente = ClienteAutoAzur(AUTOAZUR_UNIQUE_GUID)

    campos = aa_campos_efectivos()
    filas, muestra = [], {}
    claves_ya_mostradas = set()
    for canal_aa in AUTOAZUR_CANALES_PUBLICACIONES:
        canal = normalizar_canal_venta(canal_aa)
        pagina, total, total_paginas, error = 1, None, None, ""
        vistos, filas_canal, recibidos, repetidos = set(), [], 0, 0
        primer_registro = None
        mapeo = {}
        try:
            while pagina <= AUTOAZUR_MAX_PAGINAS:
                data = cliente.get("/item/listings", {
                    "ChannelID": canal_aa, "PerPage": AUTOAZUR_PER_PAGE, "Page": pagina,
                })
                items = _aa_registros(data, campos)
                if not items:
                    # Respaldo: la estructura anterior (lista directa).
                    items = [_aa_desenvolver(x, campos) for x in _aa_extraer_lista(data)]
                if pagina == 1:
                    total, total_paginas = _aa_totales(data)
                    muestra[canal_aa] = {
                        "llaves_respuesta": list(data.keys()) if isinstance(data, dict) else "lista",
                        "total_publicaciones_informado": total,
                        "total_paginas_informado": total_paginas,
                        "registros_en_pagina_1": len(items),
                        "claves_primer_registro": list(items[0].keys()) if items else [],
                        "primeros_registros": items[:3],
                    }
                if not items:
                    break
                nuevos = 0
                for it in items:
                    recibidos += 1
                    huella = _aa_huella(it)
                    if huella in vistos:
                        repetidos += 1
                        continue
                    vistos.add(huella)
                    nuevos += 1
                    if primer_registro is None:
                        primer_registro = it
                    for rec in _aa_expandir_variaciones(it, campos):
                        fila = {}
                        for campo, cands in campos.items():
                            val, clave = _aa_buscar(rec, cands)
                            if campo == "estado_original" and isinstance(val, bool):
                                # Campo sí/no: "Paused: true" es pausada;
                                # "Active/Enabled/Published: false" es inactiva.
                                if "paus" in clave.lower():
                                    val = "paused" if val else "active"
                                else:
                                    val = "active" if val else "inactive"
                            fila[campo] = val
                            if clave and campo not in mapeo:
                                mapeo[campo] = clave
                        fila["canal_autoazur"] = canal_aa
                        fila["canal"] = canal
                        fila["estado_publicacion"] = normalizar_estado_publicacion(
                            fila["estado_original"], fila["subestado"]
                        )
                        filas_canal.append(fila)
                # Fin: la API informó el número de páginas, la página vino
                # incompleta o la API repitió la página (nada nuevo).
                if nuevos == 0:
                    break
                if total_paginas:
                    if pagina >= total_paginas:
                        break
                elif total:
                    if len(vistos) >= total:
                        break
                elif len(items) < AUTOAZUR_PER_PAGE:
                    break
                pagina += 1
        except Exception as e:
            error = str(e)[:300]
            print(f"ADVERTENCIA AutoAzur {canal_aa}: {error}")

        n_pub = len(vistos)
        claves_pub = []
        for f_raw in (primer_registro,) if primer_registro else ():
            claves_pub = [f"{k}" + (" (objeto)" if isinstance(v, dict) else " (lista)" if isinstance(v, list) else "")
                          for k, v in f_raw.items()]
        con_id = sum(1 for f in filas_canal if f.get("sku_publicacion") or f.get("item_id"))
        con_estado = sum(1 for f in filas_canal if f.get("estado_original"))
        n_filas = len(filas_canal)
        if error:
            estado = "ERROR"
        elif n_pub == 0:
            estado = "SIN_PUBLICACIONES"
        elif n_filas and con_id / n_filas < 0.5:
            # Sin SKU/ItemID no se puede ligar al IQ: evaluar el canal
            # marcaría todo como "sin publicación".
            estado = "CAMPOS_NO_DETECTADOS"
        elif n_filas and con_estado / n_filas < 0.5:
            # Se puede detectar "sin publicación", pero no pausadas/revisión.
            estado = "SIN_CAMPO_ESTADO"
        elif total and n_pub < total:
            estado = "INCOMPLETO"
        else:
            estado = "OK"
        muestra.setdefault(canal_aa, {})["mapeo_detectado"] = mapeo
        log.append({
            "canal_autoazur": canal_aa, "canal": canal, "estado_consulta": estado,
            "publicaciones": n_pub, "filas_con_variaciones": n_filas, "total_informado": total,
            "paginas": pagina, "total_paginas_informado": total_paginas,
            "registros_repetidos": repetidos, "pct_con_sku_o_itemid": round(con_id / n_filas, 3) if n_filas else 0,
            "pct_con_estado": round(con_estado / n_filas, 3) if n_filas else 0,
            "campos_detectados": ", ".join(f"{k}={v}" for k, v in mapeo.items()),
            "detalle": error, "fecha_consulta": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
        })
        total_txt = f" de {total:,} informadas" if total else ""
        print(f"AutoAzur {canal_aa}: {n_pub:,} publicaciones{total_txt} ({estado}) · "
              f"SKU/ItemID {con_id}/{n_filas} · estado {con_estado}/{n_filas} · "
              f"campos: sku={mapeo.get('sku_publicacion', '—')}, item={mapeo.get('item_id', '—')}, "
              f"estado={mapeo.get('estado_original', '—')}, stock={mapeo.get('stock_publicado', '—')}")
        valores_estado = pd.Series([f.get("estado_original") for f in filas_canal if f.get("estado_original") not in (None, "")],
                                   dtype=object).astype(str).value_counts().head(8)
        muestra.setdefault(canal_aa, {})["claves_publicacion"] = claves_pub
        muestra[canal_aa]["valores_estado"] = valores_estado.to_dict()
        if not valores_estado.empty:
            print("  estados recibidos: " + ", ".join(f"{k} ({v})" for k, v in valores_estado.items()))
        if estado in ("CAMPOS_NO_DETECTADOS", "SIN_CAMPO_ESTADO") and tuple(claves_pub) not in claves_ya_mostradas:
            claves_ya_mostradas.add(tuple(claves_pub))
            if estado == "CAMPOS_NO_DETECTADOS":
                print(f"  -> Campos que trae cada publicación: {claves_pub}")
                print("  -> Si alguno es el SKU, agrégalo al .env: AUTOAZUR_CAMPO_SKU_PUBLICACION=<nombre>")
            else:
                print("  -> /item/listings no trae el estado de la publicación; se intentará con el "
                      "detalle de vinculaciones (/prod/relatedpublications) para los SKU con stock.")
        filas.extend(filas_canal)

    if ruta_muestra is not None and muestra:
        try:
            import json
            ruta_muestra.write_text(json.dumps(muestra, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
            print(f"Muestra de AutoAzur guardada en {ruta_muestra.name}")
        except Exception as e:
            print(f"ADVERTENCIA: no pude guardar la muestra de AutoAzur: {e}")

    df = pd.DataFrame(filas, columns=columnas)
    if not df.empty:
        for c in ["precio", "stock_publicado"]:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        for c in ["sku_publicacion", "item_id", "cuenta", "user_id", "estado_original", "subestado"]:
            df[c] = df[c].apply(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())
    return df, pd.DataFrame(log)


# ------------------------------------------------------------------
# Estado de publicación vía "detalle de vinculaciones" (/prod/relatedpublications)
# ------------------------------------------------------------------
# /item/listings NO trae el estado (activa, pausada, en revisión). El detalle
# de vinculaciones de AutoAzur se consulta por SKU; primero se sondea con
# pocos SKU y, solo si trae un estado reconocible, se consulta para los SKU
# que tienen stock en algún canal (los únicos donde el estado importa).
AUTOAZUR_ESTADO_DETALLE = _env_bool("AUTOAZUR_ESTADO_DETALLE", "1")
AUTOAZUR_MAX_CONSULTAS_ESTADO = int(os.getenv("AUTOAZUR_MAX_CONSULTAS_ESTADO", "1500"))
AUTOAZUR_HILOS = max(1, min(int(os.getenv("AUTOAZUR_HILOS", "4")), 8))


def _aa_registros_detalle(data, campos):
    regs = _aa_registros(data, campos)
    return regs if regs else [x for x in _aa_extraer_lista(data)]


def _aa_estado_de_registro(reg, campos):
    val, clave = _aa_buscar(reg, campos["estado_original"])
    if isinstance(val, bool):
        val = ("paused" if val else "active") if "paus" in clave.lower() else ("active" if val else "inactive")
    sub, _ = _aa_buscar(reg, campos["subestado"])
    return val, sub, clave


def sondear_estado_detalle(cliente, skus, campos, n=3):
    """Prueba /prod/relatedpublications con unos cuantos SKU.

    Devuelve (trae_estado, muestra_para_json).
    """
    muestra = {"endpoint": "/prod/relatedpublications", "pruebas": []}
    trae = False
    for sku in skus[:n]:
        try:
            data = cliente.get("/prod/relatedpublications", {"Sku": sku})
        except Exception as e:
            muestra["pruebas"].append({"sku": sku, "error": str(e)[:200]})
            continue
        regs = _aa_registros_detalle(data, campos)
        estados = [_aa_estado_de_registro(r, campos) for r in regs]
        con_estado = [e for e in estados if e[0] not in (None, "")]
        muestra["pruebas"].append({
            "sku": sku,
            "llaves_respuesta": list(data.keys()) if isinstance(data, dict) else "lista",
            "registros": len(regs),
            "claves_registro": list(regs[0].keys()) if regs else [],
            "campo_estado": con_estado[0][2] if con_estado else "",
            "respuesta_cruda": data if len(str(data)) < 6000 else str(data)[:6000],
        })
        if con_estado:
            trae = True
            break
    return trae, muestra


def consultar_estado_detalle(cliente, skus, campos):
    """{(item_id, canal) y (sku, canal): (estado, subestado)} desde el detalle."""
    from concurrent.futures import ThreadPoolExecutor
    resultados, errores = {}, 0

    def uno(sku):
        try:
            return sku, cliente.get("/prod/relatedpublications", {"Sku": sku}), ""
        except Exception as e:
            return sku, None, str(e)[:120]

    with ThreadPoolExecutor(max_workers=AUTOAZUR_HILOS) as ex:
        for sku, data, err in ex.map(uno, skus):
            if err or data is None:
                errores += 1
                continue
            for r in _aa_registros_detalle(data, campos):
                estado, sub, _ = _aa_estado_de_registro(r, campos)
                if estado in (None, ""):
                    continue
                canal_r = _aa_valor(r, campos["canal_api"])
                canal_r = normalizar_canal_venta(str(canal_r)) if canal_r not in (None, "") else ""
                item = limpiar_sku(_aa_valor(r, campos["item_id"]) or "")
                sku_r = limpiar_sku(_aa_valor(r, campos["sku_publicacion"]) or sku)
                if item:
                    resultados[("ITEM", item.upper(), canal_r)] = (estado, sub)
                    resultados[("ITEM", item.upper(), "")] = (estado, sub)
                if canal_r:
                    resultados.setdefault(("SKU", sku_r.upper(), canal_r), (estado, sub))
    return resultados, errores


def enriquecer_estado_publicaciones(pub_vinc, inventario_canal, cliente=None, ruta_muestra=None):
    """Completa estado_original/estado_publicacion con el detalle de vinculaciones.

    Solo consulta SKU publicados de IQ que tienen piezas en ese canal.
    Devuelve (pub_vinc, resumen_texto).
    """
    if pub_vinc is None or pub_vinc.empty or not AUTOAZUR_ESTADO_DETALLE:
        return pub_vinc, "desactivado"
    sin_estado = pub_vinc["estado_original"].astype(str).str.strip().eq("")
    if not sin_estado.any():
        return pub_vinc, "no necesario"
    if cliente is None:
        if not AUTOAZUR_UNIQUE_GUID:
            return pub_vinc, "sin credencial"
        cliente = ClienteAutoAzur(AUTOAZUR_UNIQUE_GUID)
    campos = aa_campos_efectivos()

    con_stock = set()
    if inventario_canal is not None and not inventario_canal.empty:
        inv = inventario_canal[pd.to_numeric(inventario_canal["inventario_total"], errors="coerce").fillna(0) > 0]
        con_stock = set(zip(inv["sku_madre"].astype(str).str.upper(), inv["canal"]))
    cand = pub_vinc[sin_estado & pub_vinc["sku_publicacion"].astype(str).str.strip().ne("")].copy()
    cand = cand[[(str(s).upper(), c) in con_stock for s, c in zip(cand["sku_madre"], cand["canal"])]]
    skus = list(dict.fromkeys(cand["sku_publicacion"].astype(str).str.strip()))
    if not skus:
        return pub_vinc, "sin candidatos con stock"

    trae, muestra = sondear_estado_detalle(cliente, skus, campos)
    if ruta_muestra is not None:
        try:
            import json
            previo = json.loads(ruta_muestra.read_text(encoding="utf-8")) if ruta_muestra.exists() else {}
            previo["_detalle_vinculaciones"] = muestra
            ruta_muestra.write_text(json.dumps(previo, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        except Exception:
            pass
    if not trae:
        claves = next((p.get("claves_registro") for p in muestra["pruebas"] if p.get("claves_registro")), [])
        print("AutoAzur detalle de vinculaciones: tampoco se reconoció un campo de estado. "
              f"Campos que trae: {claves or 'sin registros'}. Revisa '_detalle_vinculaciones' en "
              f"{AUTOAZUR_ARCHIVO_MUESTRA}; si alguno es el estado, agrégalo como AUTOAZUR_CAMPO_ESTADO_ORIGINAL.")
        return pub_vinc, "detalle sin estado"

    if len(skus) > AUTOAZUR_MAX_CONSULTAS_ESTADO:
        print(f"AutoAzur detalle: {len(skus)} SKU candidatos; se consultan los primeros "
              f"{AUTOAZUR_MAX_CONSULTAS_ESTADO} (AUTOAZUR_MAX_CONSULTAS_ESTADO).")
        skus = skus[:AUTOAZUR_MAX_CONSULTAS_ESTADO]
    print(f"AutoAzur detalle de vinculaciones: consultando estado de {len(skus)} SKU con stock "
          f"({AUTOAZUR_HILOS} en paralelo)...")
    mapa, errores = consultar_estado_detalle(cliente, skus, campos)

    out = pub_vinc.copy()
    completados = 0
    for idx in out.index[sin_estado]:
        item = limpiar_sku(out.at[idx, "item_id"]).upper()
        canal = out.at[idx, "canal"]
        sku = limpiar_sku(out.at[idx, "sku_publicacion"]).upper()
        hit = (mapa.get(("ITEM", item, canal)) or mapa.get(("ITEM", item, ""))) if item else None
        hit = hit or mapa.get(("SKU", sku, canal))
        if hit:
            out.at[idx, "estado_original"] = str(hit[0])
            out.at[idx, "subestado"] = str(hit[1] or "")
            out.at[idx, "estado_publicacion"] = normalizar_estado_publicacion(hit[0], hit[1])
            out.at[idx, "fuente_estado"] = "detalle_vinculaciones"
            completados += 1
    texto = f"{completados} publicaciones con estado desde el detalle ({errores} SKU con error)"
    print(f"AutoAzur detalle de vinculaciones: {texto}.")
    return out, texto


def extraer_mapa_barcode_odoo(odoo):
    """{código de barras normalizado: código interno} de productos Odoo."""
    try:
        prods = odoo.search_read_all(
            "product.product", [("barcode", "!=", False), ("default_code", "!=", False)],
            ["barcode", "default_code"], batch=5000, order="id asc",
        )
    except Exception as e:
        print(f"ADVERTENCIA: no se pudo leer códigos de barras de Odoo: {e}")
        return {}
    return {sku_key(p["barcode"]): limpiar_sku(p["default_code"]) for p in prods
            if p.get("barcode") and p.get("default_code")}


def resumir_publicaciones_sin_sku(pub_vinc):
    """SKU de AutoAzur sin IQ, agrupados: lista de trabajo para el diccionario."""
    if pub_vinc is None or pub_vinc.empty:
        return pd.DataFrame()
    p = pub_vinc[pub_vinc["metodo_match_publicacion"].eq("SIN_MATCH")].copy()
    if p.empty:
        return pd.DataFrame()
    p["stock_publicado"] = pd.to_numeric(p["stock_publicado"], errors="coerce").fillna(0)
    res = p.groupby(["sku_publicacion"], as_index=False, dropna=False).agg(
        canales=("canal", lambda s: " | ".join(sorted(set(map(str, s))))),
        n_canales=("canal", "nunique"),
        publicaciones=("item_id", "size"),
        stock_publicado=("stock_publicado", "sum"),
        upc=("upc", lambda s: next((str(x) for x in s if str(x).strip() not in ("", "None", "nan")), "")),
        titulo=("titulo", "first"),
        item_ids=("item_id", lambda s: " | ".join(list(dict.fromkeys(map(str, s)))[:5])),
    )
    # Primero lo que más impacta: con stock publicado y en más canales.
    return res.sort_values(["stock_publicado", "n_canales", "publicaciones"], ascending=False)


def vincular_publicaciones(publicaciones, dic_match, skus_madre_conocidos, mapa_barcode=None):
    """Asigna SKU madre a cada publicación.

    Intenta con el SKU de la publicación y luego con el ItemID (el
    diccionario incluye ASIN/MLM/Item ID en varias columnas). NO se
    autogenera SKU madre: una publicación sin match se reporta aparte en
    lugar de inventar un producto.
    """
    if publicaciones is None or publicaciones.empty:
        return pd.DataFrame()
    dic_map = dic_match.drop_duplicates("sku_key", keep="first").set_index("sku_key").to_dict("index")
    conocidos = {str(s).strip().upper() for s in skus_madre_conocidos if str(s).strip()}
    rows = []
    for r in publicaciones.to_dict("records"):
        sku, item = limpiar_sku(r.get("sku_publicacion", "")), limpiar_sku(r.get("item_id", ""))
        upc = limpiar_sku(r.get("upc", ""))
        match, metodo, usado = buscar_match_diccionario_por_candidatos([c for c in (sku, item, upc) if c], dic_map)
        if match and usado == upc and upc not in (sku, item):
            metodo = f"upc_{metodo}"
        iq = ""
        if not match and mapa_barcode:
            # UPC o SKU de AutoAzur = código de barras de un producto Odoo:
            # se usa el código interno de ese producto (normalmente el IQ).
            for codigo in (upc, sku):
                interno = mapa_barcode.get(sku_key(codigo)) if codigo else None
                if not interno:
                    continue
                match, metodo, usado = buscar_match_diccionario_por_candidatos([interno], dic_map)
                if match:
                    metodo = f"barcode_odoo_{metodo}"
                    break
                if interno.upper() in conocidos:
                    iq = interno.upper()
                    break
        if not match and not iq:
            # Respaldo: SKU de AutoAzur que contiene el IQ con prefijos o
            # sufijos (p. ej. "IQ1187-FULL", "KIT_IQ2000"). Solo si hay UN
            # único IQ en el texto y ese IQ existe.
            encontrados = sorted(set(re.findall(r"IQ\d+", sku.upper())))
            if len(encontrados) == 1 and encontrados[0] in conocidos:
                iq = encontrados[0]
            else:
                # Primer bloque antes de "-", "_" o espacio contra el diccionario.
                primero = re.split(r"[-_\s/]+", sku)[0] if sku else ""
                if primero and primero != sku:
                    match, metodo, usado = buscar_match_diccionario_por_candidatos([primero], dic_map)
                    metodo = f"prefijo_{metodo}" if match else metodo
        if iq and not match:
            r["sku_madre"], r["producto_madre"] = iq, ""
            r["metodo_match_publicacion"] = "barcode_odoo_directo" if mapa_barcode and iq not in sku.upper() \
                else "patron_IQ_en_sku"
        elif match:
            r["sku_madre"] = match.get("sku_madre", "")
            r["producto_madre"] = match.get("producto_madre", "")
            r["metodo_match_publicacion"] = f"{metodo}:{'item_id' if usado == item and usado != sku else 'sku'}"
        elif sku.upper() in conocidos:
            r["sku_madre"], r["producto_madre"] = sku.upper(), ""
            r["metodo_match_publicacion"] = "sku_madre_directo"
        else:
            r["sku_madre"], r["producto_madre"] = "", ""
            r["metodo_match_publicacion"] = "SIN_MATCH"
        rows.append(r)
    return pd.DataFrame(rows)


def resumir_publicaciones_sku_canal(pub_vinc):
    """Un renglón por SKU madre + canal con el mejor estado y el detalle."""
    columnas = [
        "sku_madre", "canal", "estado_publicacion", "n_publicaciones", "n_activas",
        "n_pausadas", "n_en_revision", "n_con_error", "n_inactivas_cerradas",
        "estados_detalle", "item_ids", "cuentas", "skus_publicados",
        "stock_publicado_activo", "stock_publicado_detectado", "subestados",
    ]
    if pub_vinc is None or pub_vinc.empty:
        return pd.DataFrame(columns=columnas)
    p = pub_vinc[pub_vinc["sku_madre"].astype(str).str.strip().ne("")].copy()
    if p.empty:
        return pd.DataFrame(columns=columnas)
    p["prioridad"] = p["estado_publicacion"].map(ESTADOS_PUBLICACION_PRIORIDAD).fillna(8)
    p["es_activa"] = p["estado_publicacion"].eq("Activa")
    # Stock que AutoAzur publica: de las activas, o de todas si el estado no
    # se conoce (así se detecta "publicada con stock 0" aun sin estado).
    p["vigente"] = p["estado_publicacion"].isin(["Activa", "Sin estado", "Otro"])
    p["stock_activo"] = np.where(p["vigente"], pd.to_numeric(p["stock_publicado"], errors="coerce"), np.nan)

    def unir(serie, maximo=6):
        vals = []
        for v in serie:
            v = str(v or "").strip()
            if v and v not in vals:
                vals.append(v)
        extra = f" (+{len(vals) - maximo})" if len(vals) > maximo else ""
        return " | ".join(vals[:maximo]) + extra

    res = p.groupby(["sku_madre", "canal"], as_index=False).agg(
        prioridad=("prioridad", "min"),
        n_publicaciones=("estado_publicacion", "size"),
        n_activas=("es_activa", "sum"),
        n_pausadas=("estado_publicacion", lambda s: int(s.str.startswith("Pausada").sum())),
        n_en_revision=("estado_publicacion", lambda s: int((s == "En revisión").sum())),
        n_con_error=("estado_publicacion", lambda s: int((s == "Con error / rechazada").sum())),
        n_inactivas_cerradas=("estado_publicacion", lambda s: int(s.isin(["Inactiva", "Cerrada / eliminada"]).sum())),
        estados_detalle=("estado_publicacion", lambda s: " · ".join(f"{k} {v}" for k, v in s.value_counts().items())),
        item_ids=("item_id", unir),
        cuentas=("cuenta", unir),
        skus_publicados=("sku_publicacion", unir),
        stock_publicado_activo=("stock_activo", "sum"),
        stock_publicado_detectado=("stock_activo", lambda s: bool(s.notna().any())),
        subestados=("subestado", unir),
    )
    inverso = {v: k for k, v in ESTADOS_PUBLICACION_PRIORIDAD.items()}
    res["estado_publicacion"] = res["prioridad"].map(inverso)
    return res[columnas]


def construir_stock_sku_madre_desde_canales(inventario_canal):
    if inventario_canal is None or inventario_canal.empty:
        return pd.DataFrame()

    inv = inventario_canal.copy()
    total = inv.groupby(["sku_madre", "producto_madre"], as_index=False).agg(
        stock_odoo_total=("inventario_odoo", "sum"),
        stock_transito_total=("inventario_transito", "sum"),
        stock_full_total=("inventario_full", "sum"),
        stock_total=("inventario_total", "sum"),
        inventario_full_corte_total=("inventario_full_corte", "sum"),
        traslados_hecho_post_corte_total=("traslados_hecho_post_corte", "sum"),
        ventas_full_post_corte_total=("ventas_full_post_corte", "sum"),
    )

    slugs = {
        "Amazon": "amazon", "Mercado Libre": "mercado_libre", "Walmart": "walmart",
        "Liverpool": "liverpool", "Coppel": "coppel", "Elektra": "elektra",
        "TikTok": "tiktok", "General": "general"
    }
    for canal, slug in slugs.items():
        tmp = inv[inv["canal"].eq(canal)].groupby("sku_madre", as_index=False).agg(
            **{
                f"stock_{slug}_odoo": ("inventario_odoo", "sum"),
                f"stock_{slug}_transito": ("inventario_transito", "sum"),
                f"stock_{slug}_full": ("inventario_full", "sum"),
                f"stock_{slug}_total": ("inventario_total", "sum"),
            }
        )
        total = total.merge(tmp, on="sku_madre", how="left")

    numeric_cols = [c for c in total.columns if c.startswith("stock_") or c.endswith("_total")]
    for c in numeric_cols:
        total[c] = pd.to_numeric(total[c], errors="coerce").fillna(0)

    # Compatibilidad con el segundo y tercer código existentes.
    total["stock_amazon_fba"] = total.get("stock_amazon_full", 0)
    total["stock_meli_full"] = total.get("stock_mercado_libre_full", 0)
    total["stock_walmart_wfs"] = total.get("stock_walmart_full", 0)
    total["stock_liverpool_99min"] = total.get("stock_liverpool_full", 0)
    total["stock_odoo_cuautitlan"] = total["stock_odoo_total"]
    total["fecha_corte_full"] = FECHA_CORTE_FULL
    return total.sort_values("stock_total", ascending=False)


# ============================================================
# ARCHIVOS LOCALES
# ============================================================

print("Buscando archivos del proyecto...")

archivo_diccionario = encontrar_archivo(
    DESKTOP,
    ["diccionario", "origen4"]
)

ruta_full_configurada = DESKTOP / ARCHIVO_STOCK_FULL_CONGELADO
if ruta_full_configurada.exists():
    archivo_full_congelado = ruta_full_configurada
else:
    archivo_full_congelado = encontrar_archivo(
        DESKTOP,
        ["stock_full_congelado", "asin", "corregido"]
    )

archivo_autoazur = encontrar_archivo_autoazur_actualizado()

print("Archivos detectados:")
print(f"- Diccionario: {archivo_diccionario.name}")
print(f"- Corte Full congelado: {archivo_full_congelado.name}")

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
stock_odoo_sin_sku = pd.DataFrame()
odoo_ubicaciones = pd.DataFrame()
odoo_ubicaciones_todas = pd.DataFrame()
stock_odoo_todas_ubicaciones = pd.DataFrame()
resumen_stock_odoo_ubicacion = pd.DataFrame()
traslados_full_detalle = pd.DataFrame()
traslados_full_excluidos = pd.DataFrame()
contactos_no_full_resumen = pd.DataFrame()
recepciones_compra = pd.DataFrame()
envios_full_historicos = pd.DataFrame()
movimientos_auditoria_raw = pd.DataFrame()
fechas_lote_todas = pd.DataFrame()
costos_productos = pd.DataFrame()

if ENABLE_ODOO:
    odoo = OdooClient(
        ODOO_URL,
        ODOO_DB,
        ODOO_USER,
        obtener_api_key_odoo()
    )

    odoo.connect()

    ventas_odoo, productos_odoo, ventas_sin_sku_odoo = extraer_ventas_odoo(odoo)
    stock_odoo, odoo_ubicaciones, stock_odoo_sin_sku = extraer_inventario_odoo(odoo)
    stock_odoo_todas_ubicaciones, resumen_stock_odoo_ubicacion, odoo_ubicaciones_todas = extraer_auditoria_stock_todas_ubicaciones(odoo)
    traslados_full_detalle, traslados_full_excluidos = extraer_traslados_full_odoo(odoo, odoo_ubicaciones)
    contactos_no_full_resumen = resumir_contactos_no_full(odoo, traslados_full_excluidos)
    if not contactos_no_full_resumen.empty:
        alta = contactos_no_full_resumen[contactos_no_full_resumen["prioridad_revision"].eq("ALTA")]
        if not alta.empty:
            print("\nREVISAR: contactos que parecen de canal comercial pero NO entraron a Full:")
            for _, r in alta.head(15).iterrows():
                print(f"- {r['contacto'][:48]:<50} {r['canal_sugerido']:<15} "
                      f"{r['unidades']:>8,.0f} u  ({r['motivo_exclusion']})")
            print(f"  Total en riesgo: {alta['unidades'].sum():,.0f} unidades "
                  f"en {int(alta['pickings'].sum())} traslados.\n")

    # --- Antigüedad FIFO, valuación y auditoría (versión 2026-09) ---------
    # Son consultas ADICIONALES: si alguna falla, el inventario y las
    # ventas siguen calculándose igual y solo se pierde esa capa de análisis.
    try:
        recepciones_compra = extraer_recepciones_compra(odoo)
    except Exception as e:
        print(f"ADVERTENCIA: recepciones de compra no disponibles: {e}")
    try:
        envios_full_historicos = extraer_envios_full_historicos(odoo, odoo_ubicaciones)
    except Exception as e:
        print(f"ADVERTENCIA: envíos Full históricos no disponibles: {e}")
    try:
        movimientos_auditoria_raw = extraer_movimientos_auditoria(odoo, odoo_ubicaciones_todas)
    except Exception as e:
        print(f"ADVERTENCIA: movimientos para auditoría no disponibles: {e}")

    columnas_lote = ["lote_id_odoo", "fecha_recepcion_lote", "fuente_recepcion_lote"]
    lotes_quants = pd.DataFrame(columns=columnas_lote)
    if not stock_odoo.empty and "lote_id_odoo" in stock_odoo.columns:
        lotes_quants = stock_odoo[[c for c in columnas_lote if c in stock_odoo.columns]].dropna(
            subset=["lote_id_odoo"]
        ).drop_duplicates("lote_id_odoo")
    lotes_ya = set(pd.to_numeric(lotes_quants["lote_id_odoo"], errors="coerce").dropna().astype(int))
    lotes_extra = set()
    for _df in (envios_full_historicos, movimientos_auditoria_raw):
        if not _df.empty and "lote_id_odoo" in _df.columns:
            lotes_extra |= set(pd.to_numeric(_df["lote_id_odoo"], errors="coerce").dropna().astype(int))
    lotes_extra -= lotes_ya
    fechas_lote_extra = pd.DataFrame(columns=columnas_lote)
    if lotes_extra:
        try:
            fechas_lote_extra = extraer_fecha_recepcion_lotes(odoo, lotes_extra)
        except Exception as e:
            print(f"ADVERTENCIA: no se pudieron fechar los lotes de envíos/movimientos: {e}")
    fechas_lote_todas = pd.concat(
        [lotes_quants, fechas_lote_extra[[c for c in columnas_lote if c in fechas_lote_extra.columns]]],
        ignore_index=True, sort=False,
    ).drop_duplicates("lote_id_odoo")

    try:
        _pids = set()
        for _df, _col in ((stock_odoo, "odoo_product_id"), (recepciones_compra, "product_id"),
                          (envios_full_historicos, "product_id")):
            if not _df.empty and _col in _df.columns:
                _pids |= set(pd.to_numeric(_df[_col], errors="coerce").dropna().astype(int))
        _skus_madre = dic_match["sku_madre"].dropna().astype(str).unique().tolist() \
            if "sku_madre" in dic_match.columns else []
        costos_productos = extraer_costos_productos(odoo, _pids, _skus_madre)
    except Exception as e:
        print(f"ADVERTENCIA: costos de producto no disponibles: {e}")

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

# Separación explícita de pendientes.
ventas_sku_sin_madre = ventas_sin_match[
    ventas_sin_match.get("sku_producto_pendiente", "").astype(str).str.strip().ne("")
].copy() if not ventas_sin_match.empty else pd.DataFrame()
referencias_no_encontradas = ventas_sin_match[
    ventas_sin_match.get("fuente", "").astype(str).str.upper().eq("AUTOAZUR")
    & ventas_sin_match.get("sku_producto_pendiente", "").astype(str).str.strip().eq("")
    & ventas_sin_match.get("estado_referencia_odoo", "").astype(str).eq("NO_ENCONTRADA")
].copy() if not ventas_sin_match.empty else pd.DataFrame()
referencias_ambiguas = ventas_sin_match[
    ventas_sin_match.get("fuente", "").astype(str).str.upper().eq("AUTOAZUR")
    & ventas_sin_match.get("estado_referencia_odoo", "").astype(str).eq("AMBIGUA")
].copy() if not ventas_sin_match.empty else pd.DataFrame()
autoazur_resuelto_ref = ventas_con_match[
    ventas_con_match.get("metodo_match", "").astype(str).eq("referencia_autoazur_vs_odoo")
].copy() if not ventas_con_match.empty else pd.DataFrame()
item_ids_sin_sku_madre = ventas_sin_match[
    ventas_sin_match.get("item_id", "").astype(str).str.strip().ne("")
].copy() if not ventas_sin_match.empty else pd.DataFrame()

print(f"Ventas Odoo: {len(ventas_odoo)}")
print(f"Ventas Autoazur originales: {len(ventas_autoazur)}")
print(f"Ventas Autoazur omitidas por duplicado con Odoo: {len(autoazur_duplicados_omitidos)}")
print(f"Ventas conjunto final Odoo + Autoazur: {len(ventas_conjunto)}")
print(f"Ventas vinculadas a SKU madre: {len(ventas_con_match[ventas_con_match['tiene_referencia_madre'] == 'SI']) if len(ventas_con_match) else 0}")
print(f"Ventas sin referencia madre: {len(ventas_sin_match)}")


# ============================================================
# STOCK FULL DEL CORTE MAESTRO
# ============================================================

full_congelado = pd.read_excel(
    archivo_full_congelado,
    sheet_name="stock_full_congelado",
    dtype=str
)
full_congelado.columns = [str(c).strip() for c in full_congelado.columns]

columnas_full_requeridas = [
    "canal_stock", "sku_original", "producto_stock", "stock",
    "canal_asignado", "tipo_inventario", "fecha_corte", "fuente_archivo"
]
faltantes_full = [c for c in columnas_full_requeridas if c not in full_congelado.columns]
if faltantes_full:
    raise ValueError(
        f"El archivo de corte Full no contiene estas columnas: {faltantes_full}"
    )

# Evita mezclar accidentalmente un Excel de corte con otra fecha.
_fechas_corte_archivo = pd.to_datetime(
    full_congelado["fecha_corte"], errors="coerce"
).dropna().drop_duplicates()
if len(_fechas_corte_archivo) != 1:
    raise ValueError(
        "El archivo Full debe contener exactamente una fecha_corte válida. "
        f"Encontré: {[str(x) for x in _fechas_corte_archivo.tolist()[:10]]}"
    )
_fecha_archivo = pd.Timestamp(_fechas_corte_archivo.iloc[0])
if _fecha_archivo != FECHA_CORTE_FULL:
    raise ValueError(
        "La fecha del Excel Full no coincide con FECHA_CORTE_FULL. "
        f"Excel={_fecha_archivo}; configuración={FECHA_CORTE_FULL}. "
        "Ejecuta primero 00_calcular_stock_inicial.py."
    )

stock_full = full_congelado.copy()
stock_full["sku_original"] = stock_full["sku_original"].apply(limpiar_sku)
stock_full["stock"] = to_number(stock_full["stock"]).clip(lower=0)
stock_full["stock_original"] = stock_full["stock"]
stock_full["tipo_inventario"] = "FULL"
stock_full["fecha_corte"] = pd.to_datetime(
    stock_full["fecha_corte"], errors="coerce"
).fillna(FECHA_CORTE_FULL)

stock_full = stock_full[
    (stock_full["sku_original"] != "")
    & (stock_full["stock"] > 0)
].copy()

print(
    f"Stock Full congelado cargado: {len(stock_full)} renglones, "
    f"{stock_full['stock'].sum():,.0f} unidades"
)

stock_detalle = pd.concat(
    [
        stock_full,
        stock_odoo,
    ],
    ignore_index=True,
    sort=False
)


if "stock_original" not in stock_detalle.columns:
    stock_detalle["stock_original"] = pd.to_numeric(stock_detalle.get("stock", 0), errors="coerce").fillna(0)
else:
    stock_detalle["stock_original"] = pd.to_numeric(stock_detalle["stock_original"], errors="coerce").fillna(
        pd.to_numeric(stock_detalle.get("stock", 0), errors="coerce").fillna(0)
    )
stock_detalle["stock"] = pd.to_numeric(stock_detalle.get("stock", 0), errors="coerce").fillna(0).clip(lower=0)
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

# El pivot se queda solo con las cantidades, así que el nombre del producto
# se vuelve a pegar aquí. Sin esto, un SKU madre autogenerado quedaría sin
# descripción y sería ilegible en el dashboard.
nombres_producto = (
    stock_por_sku_canal
    .groupby(["sku_key", "sku_original"], as_index=False)
    .agg(producto_stock=("producto_stock",
                         lambda x: next((str(v) for v in x if str(v).strip()), "")))
)
stock_por_sku = stock_por_sku.merge(
    nombres_producto, on=["sku_key", "sku_original"], how="left"
)
stock_por_sku["producto_stock"] = stock_por_sku["producto_stock"].fillna("")

# Columnas de stock congelado Full (archivo de corte).
stock_cols_full = [
    "WALMART_WFS", "LIVERPOOL_FULL_99MIN", "MERCADO_LIBRE_FULL", "AMAZON_FBA",
]

# Columnas de stock Odoo: se derivan de ODOO_UBICACIONES_VALIDAS en vez de
# estar escritas a mano. Antes, agregar un canal a la configuración no lo
# sumaba al stock_total porque esta lista se quedaba atrás.
stock_cols_odoo = sorted({v[1] for v in ODOO_UBICACIONES_VALIDAS.values()})

# Cualquier canal que haya aparecido en los datos y no esté en la config
# tampoco se pierde: se agrega aquí para que entre al total.
stock_cols_detectadas = [
    c for c in stock_por_sku.columns
    if str(c).startswith("ODOO_") and c not in stock_cols_odoo
]
if stock_cols_detectadas:
    print("AVISO: canales de stock presentes en datos pero no en la configuración:")
    for c in stock_cols_detectadas:
        print(f"- {c}")

stock_cols = stock_cols_full + stock_cols_odoo + stock_cols_detectadas

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
# INVENTARIO INDIVIDUAL POR CANAL: ODOO + TRÁNSITO + FULL
# ============================================================

stock_detalle_vinculado = vincular_detalle_sku(stock_detalle, dic_match)
traslados_full_vinculados = vincular_detalle_sku(traslados_full_detalle, dic_match)

inventario_canal_sku_madre = construir_inventario_canal_sku_madre(
    stock_detalle_vinculado,
    traslados_full_vinculados,
    ventas_con_match,
)

# Antigüedad de inventario por canal, rastreada por Lote (Compras -> ficha
# de ingreso; Inventario -> Reporte -> Trazabilidad para saber cuántas
# piezas de cada lote quedaron asignadas a cada canal).
dias_inventario_lote_canal = construir_dias_inventario_lote_canal(stock_detalle_vinculado)
if not inventario_canal_sku_madre.empty:
    if not dias_inventario_lote_canal.empty:
        inventario_canal_sku_madre = inventario_canal_sku_madre.merge(
            dias_inventario_lote_canal,
            on=["sku_madre", "producto_madre", "canal"],
            how="left"
        )
    for col in ["dias_inventario_lote", "unidades_con_lote", "unidades_sin_lote_odoo", "cobertura_lote_pct"]:
        if col not in inventario_canal_sku_madre.columns:
            inventario_canal_sku_madre[col] = np.nan
    for col in ["fecha_recepcion_lote_min", "fecha_recepcion_lote_max"]:
        if col not in inventario_canal_sku_madre.columns:
            inventario_canal_sku_madre[col] = pd.NaT
    if "fuente_dias_inventario" not in inventario_canal_sku_madre.columns:
        inventario_canal_sku_madre["fuente_dias_inventario"] = "SIN_LOTE"
    inventario_canal_sku_madre["fuente_dias_inventario"] = inventario_canal_sku_madre["fuente_dias_inventario"].fillna("SIN_LOTE")
    # La fuente anterior (solo lotes CUATI) se conserva con otro nombre para
    # auditoría; la nueva antigüedad FIFO por canal toma su lugar.
    inventario_canal_sku_madre = inventario_canal_sku_madre.rename(
        columns={"fuente_dias_inventario": "fuente_dias_inventario_lote"}
    )

# ------------------------------------------------------------------
# ANTIGÜEDAD FIFO POR CANAL + COSTO + VALOR (versión 2026-09)
# ------------------------------------------------------------------
recepciones_compra_vinc = vincular_detalle_sku(recepciones_compra, dic_match)
envios_full_vinc = vincular_detalle_sku(envios_full_historicos, dic_match)
antiguedad_canal = pd.DataFrame()
antiguedad_capas = pd.DataFrame()
fifo_compras_sku = pd.DataFrame()
if not inventario_canal_sku_madre.empty:
    antiguedad_canal, antiguedad_capas, fifo_compras_sku = construir_antiguedad_canal(
        inventario_canal_sku_madre, stock_detalle_vinculado, envios_full_vinc, recepciones_compra_vinc,
    )
    if not antiguedad_canal.empty:
        inventario_canal_sku_madre = inventario_canal_sku_madre.drop(
            columns=[c for c in COLUMNAS_ANTIGUEDAD_CANAL if c in inventario_canal_sku_madre.columns]
        ).merge(
            antiguedad_canal[["sku_madre", "canal"] + [c for c in COLUMNAS_ANTIGUEDAD_CANAL if c in antiguedad_canal.columns]],
            on=["sku_madre", "canal"], how="left",
        )
    for c in COLUMNAS_ANTIGUEDAD_CANAL:
        if c not in inventario_canal_sku_madre.columns:
            inventario_canal_sku_madre[c] = np.nan
    inventario_canal_sku_madre["fuente_dias_inventario"] = (
        inventario_canal_sku_madre["fuente_dias_inventario"].fillna("SIN_DATO")
    )

costos_productos_vinc = vincular_detalle_sku(costos_productos, dic_match)
costos_sku_madre = construir_costos_sku_madre(inventario_canal_sku_madre, costos_productos_vinc, recepciones_compra_vinc)
if not inventario_canal_sku_madre.empty:
    inventario_canal_sku_madre = agregar_valor_inventario(inventario_canal_sku_madre, costos_sku_madre)
if not antiguedad_capas.empty:
    antiguedad_capas = antiguedad_capas.merge(costos_sku_madre, on="sku_madre", how="left")
    antiguedad_capas["costo_unitario_inventario"] = pd.to_numeric(
        antiguedad_capas["costo_unitario_inventario"], errors="coerce").fillna(0)
    antiguedad_capas["valor"] = antiguedad_capas["unidades"] * antiguedad_capas["costo_unitario_inventario"]

movimientos_auditoria_vinc = vincular_detalle_sku(movimientos_auditoria_raw, dic_match)
movimientos_odoo_auditoria, movimientos_odoo_resumen = procesar_auditoria_movimientos(
    movimientos_auditoria_vinc, envios_full_vinc, fechas_lote_todas, costos_sku_madre,
    CARPETA_SALIDA / ARCHIVO_REGISTRO_MOVIMIENTOS_NOMBRE,
)
# ------------------------------------------------------------------
# PUBLICACIONES AUTOAZUR POR CANAL (stock asignado vs publicado)
# ------------------------------------------------------------------
publicaciones_autoazur = pd.DataFrame()
publicaciones_log = pd.DataFrame()
publicaciones_sku_canal = pd.DataFrame()
publicaciones_sin_sku_madre = pd.DataFrame()
publicaciones_sin_sku_resumen = pd.DataFrame()
try:
    _pub_raw, publicaciones_log = extraer_publicaciones_autoazur(
        ruta_muestra=CARPETA_SALIDA / AUTOAZUR_ARCHIVO_MUESTRA
    )
    _skus_conocidos = set(inventario_canal_sku_madre["sku_madre"].astype(str)) \
        if not inventario_canal_sku_madre.empty else set()
    _mapa_barcode = {}
    if not _pub_raw.empty:
        try:
            _mapa_barcode = extraer_mapa_barcode_odoo(odoo)
        except NameError:
            pass
    publicaciones_autoazur = vincular_publicaciones(_pub_raw, dic_match, _skus_conocidos, _mapa_barcode)
    publicaciones_autoazur, _txt_estado = enriquecer_estado_publicaciones(
        publicaciones_autoazur, inventario_canal_sku_madre,
        ruta_muestra=CARPETA_SALIDA / AUTOAZUR_ARCHIVO_MUESTRA,
    )
    if not publicaciones_log.empty and "canal" in publicaciones_log.columns and not publicaciones_autoazur.empty:
        _con_det = publicaciones_autoazur.get("fuente_estado", pd.Series("", index=publicaciones_autoazur.index))
        _con_det = publicaciones_autoazur[_con_det.astype(str).eq("detalle_vinculaciones")].groupby("canal").size()
        publicaciones_log["estado_desde_detalle"] = publicaciones_log["canal"].map(_con_det).fillna(0).astype(int)
    publicaciones_sku_canal = resumir_publicaciones_sku_canal(publicaciones_autoazur)
    if not publicaciones_autoazur.empty:
        publicaciones_sin_sku_madre = publicaciones_autoazur[
            publicaciones_autoazur["metodo_match_publicacion"].eq("SIN_MATCH")
        ].copy()
        print(f"Publicaciones AutoAzur: {len(publicaciones_autoazur)} "
              f"({len(publicaciones_sin_sku_madre)} sin SKU madre).")
        _metodos = publicaciones_autoazur["metodo_match_publicacion"].astype(str).str.split(":").str[0].value_counts()
        print("  Cómo se ligaron al IQ: " + ", ".join(f"{k} {v}" for k, v in _metodos.items()))
        publicaciones_sin_sku_resumen = resumir_publicaciones_sin_sku(publicaciones_autoazur)
        if not publicaciones_sin_sku_resumen.empty:
            _con_stock = int((publicaciones_sin_sku_resumen["stock_publicado"] > 0).sum())
            print(f"  SKU de AutoAzur sin IQ: {len(publicaciones_sin_sku_resumen)} distintos "
                  f"({_con_stock} con stock publicado). Lista en la hoja 'publicaciones_sin_sku_resumen'.")
except Exception as e:
    print(f"ADVERTENCIA: no se pudo construir la sección de publicaciones AutoAzur: {e}")

if not movimientos_odoo_auditoria.empty:
    _alta = movimientos_odoo_auditoria[movimientos_odoo_auditoria["nivel_revision"].eq("ALTA")]
    print(f"Auditoría: {len(movimientos_odoo_auditoria)} movimientos visibles, "
          f"{len(_alta)} con revisión ALTA.")

stock_sku_madre_nuevo = construir_stock_sku_madre_desde_canales(inventario_canal_sku_madre)
if not stock_sku_madre_nuevo.empty:
    stock_sku_madre = stock_sku_madre_nuevo

ventas_canal_sku_madre = pd.DataFrame()
if not ventas_con_match.empty:
    ventas_tmp_canal = ventas_con_match[
        ventas_con_match.get("tiene_referencia_madre", "").astype(str).eq("SI")
    ].copy()
    if not ventas_tmp_canal.empty:
        ventas_tmp_canal["canal_venta"] = ventas_tmp_canal.get("canal_venta", "Sin identificar").replace("", "Sin identificar")
        ventas_tmp_canal["modalidad_venta"] = ventas_tmp_canal.get("modalidad_venta", "SIN_IDENTIFICAR").replace("", "SIN_IDENTIFICAR")
        ventas_canal_sku_madre = ventas_tmp_canal.groupby(
            ["sku_madre", "producto_madre", "canal_venta", "modalidad_venta"], as_index=False
        ).agg(
            unidades_vendidas=("cantidad", "sum"),
            venta_total=("venta_total", "sum"),
            pedidos=("pedido", "nunique"),
            fecha_primera_venta=("fecha", "min"),
            fecha_ultima_venta=("fecha", "max"),
            fuentes=("fuente", lambda x: " | ".join(sorted(set(map(str, x))))),
        )

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
# ALIASES SUGERIDOS PARA ORIGEN4
# ============================================================
ventas_alias = pd.DataFrame()
if not ventas_sku_sin_madre.empty:
    tmp = ventas_sku_sin_madre.copy()
    tmp["sku_alias"] = tmp.get("sku_producto_pendiente", "").apply(limpiar_sku)
    tmp = tmp[tmp["sku_alias"].ne("")]
    ventas_alias = tmp.groupby("sku_alias", as_index=False).agg(
        producto=("producto", lambda x: " | ".join(sorted(set(str(v) for v in x if str(v).strip()))[:3])),
        unidades_vendidas=("cantidad", "sum"), monto_vendido=("venta_total", "sum"),
        fuentes=("fuente", lambda x: " | ".join(sorted(set(map(str, x))))),
        pedidos=("pedido", "nunique"),
    )

stock_alias = pd.DataFrame()
if not stock_sin_match.empty:
    tmp = stock_sin_match.copy()
    tmp["sku_alias"] = tmp.get("sku_original", "").apply(limpiar_sku)
    tmp = tmp[tmp["sku_alias"].ne("")]
    stock_alias = tmp.groupby("sku_alias", as_index=False).agg(stock=("stock_total", "sum"))

if ventas_alias.empty and stock_alias.empty:
    aliases_sugeridos_origen4 = pd.DataFrame(columns=[
        "sku_madre_por_completar", "sku_alias", "producto", "unidades_vendidas",
        "monto_vendido", "stock", "fuentes", "pedidos", "prioridad", "accion"
    ])
else:
    aliases_sugeridos_origen4 = ventas_alias.merge(stock_alias, on="sku_alias", how="outer")
    for c in ["unidades_vendidas", "monto_vendido", "stock", "pedidos"]:
        aliases_sugeridos_origen4[c] = pd.to_numeric(aliases_sugeridos_origen4.get(c, 0), errors="coerce").fillna(0)
    for c in ["producto", "fuentes"]:
        aliases_sugeridos_origen4[c] = aliases_sugeridos_origen4.get(c, "").fillna("")
    aliases_sugeridos_origen4.insert(0, "sku_madre_por_completar", "")
    aliases_sugeridos_origen4["prioridad"] = np.where(
        aliases_sugeridos_origen4["monto_vendido"] > 0, "1_VENTAS",
        np.where(aliases_sugeridos_origen4["stock"] > 0, "2_STOCK", "3_REVISAR")
    )
    aliases_sugeridos_origen4["accion"] = "Completar sku_madre y copiar sku_madre + sku_alias a aliases_odoo_manual en origen4"
    aliases_sugeridos_origen4 = aliases_sugeridos_origen4.sort_values(
        ["prioridad", "monto_vendido", "unidades_vendidas", "stock"],
        ascending=[True, False, False, False]
    )


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
    ["ventas_odoo_sin_codigo", len(ventas_sin_sku_odoo)],
    ["ventas_sku_sin_madre", len(ventas_sku_sin_madre)],
    ["referencias_no_encontradas", len(referencias_no_encontradas)],
    ["referencias_ambiguas", len(referencias_ambiguas)],
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
    ["stock_sku_sin_madre", len(stock_sin_match)],
    ["stock_odoo_sin_codigo", len(stock_odoo_sin_sku)],
    ["traslados_full_detalle", len(traslados_full_detalle)],
    ["traslados_full_listos_unidades", float(traslados_full_detalle.loc[traslados_full_detalle.get("estado", "").eq("Listo"), "cantidad_metricas"].sum()) if not traslados_full_detalle.empty else 0],
    ["inventario_canal_renglones", len(inventario_canal_sku_madre)],
    [
        "sku_madre_con_costo_cuati",
        int((inventario_canal_sku_madre.get("costo_unitario_cuati", 0) > 0).sum())
        if not inventario_canal_sku_madre.empty else 0
    ],
    ["fecha_corte_full", FECHA_CORTE_FULL.strftime("%Y-%m-%d %H:%M:%S")],
    ["valor_inventario_total", float(inventario_canal_sku_madre.get("valor_inventario_canal", pd.Series(dtype=float)).sum())
        if not inventario_canal_sku_madre.empty else 0],
    ["valor_inventario_mas_90d", float(inventario_canal_sku_madre.get("valor_mas_90d", pd.Series(dtype=float)).sum())
        if not inventario_canal_sku_madre.empty else 0],
    ["sku_madre_sin_costo", int(costos_sku_madre["fuente_costo_inventario"].eq("SIN_COSTO").sum())
        if not costos_sku_madre.empty else 0],
    ["envios_full_historicos_lineas", len(envios_full_historicos)],
    ["recepciones_compra_lineas", len(recepciones_compra)],
    ["movimientos_auditoria_visibles", len(movimientos_odoo_auditoria)],
    ["publicaciones_autoazur", len(publicaciones_autoazur)],
    ["publicaciones_sin_sku_madre", len(publicaciones_sin_sku_madre)],
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
    # Compatibilidad: ventas_sin_referencia ahora contiene solo SKU de producto sin madre.
    "ventas_sin_referencia": ventas_sku_sin_madre,
    "ventas_sku_sin_madre": ventas_sku_sin_madre,
    "referencias_no_encontradas": referencias_no_encontradas,
    "referencias_ambiguas": referencias_ambiguas,
    "autoazur_resuelto_ref": autoazur_resuelto_ref,
    "item_ids_sin_sku_madre": item_ids_sin_sku_madre,
    "aliases_sugeridos_origen4": aliases_sugeridos_origen4,
    "ventas_odoo_detalle": ventas_odoo,
    "ventas_autoazur_detalle": ventas_autoazur,
    "autoazur_duplicados_omitidos": autoazur_duplicados_omitidos,
    "ventas_odoo_sin_sku": ventas_sin_sku_odoo,
    "ventas_odoo_sin_codigo": ventas_sin_sku_odoo,
    "stock_detalle": stock_detalle,
    "stock_por_sku": stock_con_match,
    "stock_sku_madre": stock_sku_madre,
    "stock_sin_referencia": stock_sin_match,
    "stock_sku_sin_madre": stock_sin_match,
    "stock_odoo_detalle": stock_odoo,
    "stock_odoo_sin_sku": stock_odoo_sin_sku,
    "stock_odoo_sin_codigo": stock_odoo_sin_sku,
    "stock_odoo_todas_ubicaciones": stock_odoo_todas_ubicaciones,
    "resumen_stock_odoo_ubicacion": resumen_stock_odoo_ubicacion,
    "odoo_ubicaciones_todas": odoo_ubicaciones_todas,
    "stock_detalle_vinculado": stock_detalle_vinculado,
    "inventario_canal_sku_madre": inventario_canal_sku_madre,
    "dias_inventario_lote_canal": dias_inventario_lote_canal,
    "antiguedad_capas": antiguedad_capas,
    "fifo_compras_sku": fifo_compras_sku,
    "costos_sku_madre": costos_sku_madre,
    "recepciones_compra": recepciones_compra_vinc,
    "envios_full_historicos": envios_full_vinc,
    "movimientos_odoo_auditoria": movimientos_odoo_auditoria,
    "movimientos_odoo_resumen": movimientos_odoo_resumen,
    "publicaciones_autoazur": publicaciones_autoazur,
    "publicaciones_sku_canal": publicaciones_sku_canal,
    "publicaciones_sin_sku_madre": publicaciones_sin_sku_madre,
    "publicaciones_log": publicaciones_log,
    "publicaciones_sin_sku_resumen": publicaciones_sin_sku_resumen,
    "ventas_canal_sku_madre": ventas_canal_sku_madre,
    "traslados_full_detalle": traslados_full_vinculados,
    "traslados_full_excluidos": traslados_full_excluidos,
    "contactos_no_full_resumen": contactos_no_full_resumen,
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
print(f"Versión: {VERSION_CODIGO}")
print(f"Excel generado en: {ARCHIVO_SALIDA}")
print("\nHojas creadas:")

for hoja in salidas:
    print(f"- {hoja[:31]}")
