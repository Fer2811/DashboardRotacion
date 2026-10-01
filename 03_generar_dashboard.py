# -*- coding: utf-8 -*-
"""
Dashboard profesional de rotación de inventario por canal.

Entrada:
    base_dashboard_rotacion.xlsx

Salida:
    index.html

Características:
- Barra lateral colapsable.
- Resumen ejecutivo.
- Vista individual por KAM/canal.
- Transferencias sugeridas.
- Compras sugeridas.
- Lotes y antigüedad.
- Alertas.
- Calidad de datos.
- Metodología.
- Tablas con búsqueda y exportación CSV.
- Compatible con la base nueva por canal y con bases anteriores (fallback).

Ejecución:
    python "03_generar_dashboard_html_canales.py"
"""

from __future__ import annotations

from pathlib import Path
from datetime import datetime
import base64
import html
import json
import math
import os
import re
import subprocess
import xmlrpc.client
from typing import Any

import numpy as np
import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=False)
except ImportError:
    pass


# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA = Path(
    os.getenv(
        "DASHBOARD_ROTACION_DIR",
        str(Path(__file__).resolve().parent),
    )
)
ARCHIVO_BASE = CARPETA / "base_dashboard_rotacion.xlsx"
ARCHIVO_HTML = CARPETA / "index.html"
ARCHIVO_ACCIONES_ADMIN = CARPETA / "dashboard_actions.json"

# ============================================================
# MÓDULO B2B · EXCEDENTE + COSTO DE OPORTUNIDAD VPN
# ============================================================
# Fuente financiera semanal. El costo de oportunidad Marketplace se mantiene
# como referencia económica, pero B2B aplica un piso comercial de 10% por pieza.
ARCHIVO_ROI_B2B = CARPETA / os.getenv("ARCHIVO_ROI_B2B", "base_utilidad_roi.xlsx")
B2B_HORIZONTE_DIAS = int(os.getenv("B2B_HORIZONTE_DIAS", "30"))
B2B_TASA_DESCUENTO_DIARIA = float(os.getenv("B2B_TASA_DESCUENTO_DIARIA", "0.000355"))
B2B_ROI_VENTANA = os.getenv("B2B_ROI_VENTANA", "90").strip().lower()
# Regla comercial solicitada: nunca recomendar B2B por debajo de 10% ROI.
B2B_ROI_MINIMO = max(float(os.getenv("B2B_ROI_MINIMO", "0.10")), 0.10)
# Para canales con payout incompleto se usan 30 días como supuesto temporal.
B2B_PAYOUT_FALLBACK_DIAS = 30.0

# Stock B2B: por defecto se refresca directamente desde Odoo al generar el HTML.
# Las credenciales permanecen en .env y NUNCA se incrustan en el navegador.
B2B_ODOO_LIVE_STOCK = os.getenv("B2B_ODOO_LIVE_STOCK", "1").strip().lower() in {"1", "true", "si", "sí", "yes"}
ODOO_URL = os.getenv("ODOO_URL", "").strip().rstrip("/")
ODOO_DB = os.getenv("ODOO_DB", "").strip()
ODOO_USER = os.getenv("ODOO_USER", "").strip()
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "").strip()

# Stock físicamente disponible = quantity - reserved_quantity, igual que el 01.
B2B_ODOO_STOCK_MODE = os.getenv("B2B_ODOO_STOCK_MODE", "available").strip().lower()
B2B_ODOO_ROOTS = [
    "cuati/existencias",
    "cuati/amazon",
    "cuati/mercadolibre",
    "cuati/walmart",
    "cuati/liverpool",
    "cuati/coppel",
    "cuati/elektra",
    "cuati/tiktok",
    # CUATI/B2B tiene stock real y el 01 ya la integra; antes el módulo B2B
    # la ignoraba justo al calcular lo que B2B puede vender.
    "cuati/b2b",
]


def _env_float_b2b(nombre: str, default: float | None = None) -> float | None:
    raw = os.getenv(nombre)
    if raw is None or str(raw).strip() == "":
        return default
    try:
        return float(str(raw).strip())
    except Exception:
        return default


B2B_PAYOUT_DIAS = {
    "Amazon": _env_float_b2b("B2B_PAYOUT_AMAZON_DIAS", None),
    "Mercado Libre": _env_float_b2b("B2B_PAYOUT_MERCADO_LIBRE_DIAS", 5.0),
    "Walmart": _env_float_b2b("B2B_PAYOUT_WALMART_DIAS", 24.0),
    "Liverpool": _env_float_b2b("B2B_PAYOUT_LIVERPOOL_DIAS", None),
    "Coppel": _env_float_b2b("B2B_PAYOUT_COPPEL_DIAS", 39.0),
    "Elektra": _env_float_b2b("B2B_PAYOUT_ELEKTRA_DIAS", None),
    "TikTok": _env_float_b2b("B2B_PAYOUT_TIKTOK_DIAS", 14.0),
}

FECHA_GENERACION = datetime.now().strftime("%d/%m/%Y %H:%M")

# Parámetros preliminares visibles en Metodología.
DIAS_TRANSITO_ODOO_FULL = int(os.getenv("DIAS_TRANSITO_ODOO_FULL", "5"))
# Se conserva el nombre anterior para compatibilidad visual y de riesgo,
# pero la fuente de verdad es DIAS_TRANSITO_ODOO_FULL.
LEAD_TIME_FULL_DIAS = DIAS_TRANSITO_ODOO_FULL
SEGURIDAD_FULL_DIAS = int(os.getenv("SEGURIDAD_FULL_DIAS", "5"))
LEAD_TIME_PROVEEDOR_DIAS = int(os.getenv("LEAD_TIME_PROVEEDOR_DIAS", "30"))
COBERTURA_COMPRA_DIAS = int(os.getenv("COBERTURA_COMPRA_DIAS", "45"))
OBJETIVO_FULL_DIAS = int(os.getenv("OBJETIVO_FULL_DIAS", "30"))
MAX_PIEZAS_PRODUCTO_NUEVO_FULL = int(os.getenv("MAX_PIEZAS_PRODUCTO_NUEVO_FULL", "10"))
CANALES_CON_FULL = {
    c.strip()
    for c in os.getenv(
        "CANALES_CON_FULL",
        "Amazon,Mercado Libre,Walmart,Liverpool",
    ).split(",")
    if c.strip()
}

# Nombres editables. No implican control de acceso; son etiquetas visuales.
KAM_POR_CANAL = {
    "Amazon": os.getenv("KAM_AMAZON", "KAM Amazon"),
    "Mercado Libre": os.getenv("KAM_MERCADO_LIBRE", "KAM Mercado Libre"),
    "Walmart": os.getenv("KAM_WALMART", "KAM Walmart"),
    "Liverpool": os.getenv("KAM_LIVERPOOL", "KAM Liverpool"),
    "Coppel": os.getenv("KAM_COPPEL", "KAM Coppel"),
    "Elektra": os.getenv("KAM_ELEKTRA", "KAM Elektra"),
    "TikTok": os.getenv("KAM_TIKTOK", "KAM TikTok"),
    "General": os.getenv("KAM_GENERAL", "Inventario general"),
}

# Umbral y criterio de antigüedad (deben coincidir con 01/02).
ANTIGUEDAD_UMBRAL_ALERTA_DIAS = int(os.getenv("ANTIGUEDAD_UMBRAL_ALERTA_DIAS", "90"))
AUDITORIA_UMBRAL_EDAD_SOSPECHA = int(os.getenv("AUDITORIA_UMBRAL_EDAD_SOSPECHA", "75"))
AUDITORIA_DIAS_EN_DASHBOARD = int(os.getenv("AUDITORIA_DIAS_EN_DASHBOARD", "180"))

# Semáforo de cobertura (días) para la sección "Cobertura" de cada canal.
COBERTURA_QUIEBRE_DIAS = int(os.getenv("COBERTURA_QUIEBRE_DIAS", "7"))
COBERTURA_BAJA_DIAS = int(os.getenv("COBERTURA_BAJA_DIAS", "15"))
COBERTURA_ALTA_DIAS = int(os.getenv("COBERTURA_ALTA_DIAS", "60"))
COBERTURA_EXCESO_DIAS = int(os.getenv("COBERTURA_EXCESO_DIAS", "90"))

CANALES_ORDEN = [
    "Amazon",
    "Mercado Libre",
    "Walmart",
    "Liverpool",
    "Coppel",
    "Elektra",
    "TikTok",
]


# ============================================================
# UTILIDADES
# ============================================================


def safe_float(value: Any, default: float = 0.0) -> float:
    try:
        if pd.isna(value):
            return default
        value = float(value)
        if math.isfinite(value):
            return value
    except Exception:
        pass
    return default


def to_num(series: pd.Series | Any) -> pd.Series:
    if isinstance(series, pd.Series):
        return pd.to_numeric(series, errors="coerce").fillna(0)
    return pd.Series(dtype=float)


def normalizar_canal(value: Any) -> str:
    text = str(value or "").strip().lower()
    if "amazon" in text or "fba" in text:
        return "Amazon"
    if "mercado" in text or "meli" in text or text == "ml":
        return "Mercado Libre"
    if "walmart" in text or "wfs" in text:
        return "Walmart"
    if "liverpool" in text or "99min" in text:
        return "Liverpool"
    if "coppel" in text:
        return "Coppel"
    if "elektra" in text:
        return "Elektra"
    if "tiktok" in text or "tik tok" in text:
        return "TikTok"
    if "general" in text or "existencias" in text or "odoo" in text:
        return "General"
    return str(value or "Sin identificar").strip() or "Sin identificar"


# Columnas de 'stock_no_vinculado' que NO representan una ubicación/canal de
# stock (son metadatos del match contra el diccionario), para no confundirlas
# con columnas tipo AMAZON_FBA, WALMART_WFS, ODOO_COPPEL, etc.
COLUMNAS_STOCK_NO_VINCULADO_METADATA = {
    "sku_original", "sku_key", "sku_key_sin_ceros", "sku_usado_para_match",
    "metodo_match", "alias_diccionario", "sku_madre", "producto_madre",
    "hoja_diccionario", "columna_alias", "tiene_referencia_madre",
    "stock_total", "motivo_no_vinculado", "accion_sugerida",
}


def anotar_canales_stock_no_vinculado(df: pd.DataFrame) -> pd.DataFrame:
    """'stock_no_vinculado' trae el stock en formato ancho (una columna por
    ubicación: AMAZON_FBA, WALMART_WFS, ODOO_COPPEL, ODOO_ELEKTRA, etc.), no
    una columna 'canal'. Para poder mostrar estos SKU en la página de cada
    canal, se reutiliza normalizar_canal() sobre el NOMBRE de cada columna
    de stock: si esa columna normaliza a un canal real (CANALES_ORDEN) y su
    valor es mayor a 0 en ese renglón, se considera que ese SKU afecta a
    ese canal. Un mismo renglón puede afectar a más de un canal (p.ej. si
    hay stock sin vincular tanto en AMAZON_FBA como en ODOO_AMAZON)."""
    out = df.copy()
    if out.empty:
        out["canales_afectados_texto"] = pd.Series(dtype="object")
        return out

    mapa_columna_canal: dict[str, str] = {}
    for col in out.columns:
        if col in COLUMNAS_STOCK_NO_VINCULADO_METADATA:
            continue
        canal = normalizar_canal(col)
        if canal in CANALES_ORDEN:
            mapa_columna_canal[col] = canal

    if not mapa_columna_canal:
        out["canales_afectados_texto"] = "Sin canal identificado"
        return out

    indicadores = pd.DataFrame(index=out.index)
    for col, canal in mapa_columna_canal.items():
        tiene_stock = pd.to_numeric(out[col], errors="coerce").fillna(0) > 0
        # Si dos columnas mapean al mismo canal (p.ej. AMAZON_FBA y
        # ODOO_AMAZON), basta con que una tenga stock > 0 (OR).
        indicadores[canal] = indicadores[canal] | tiene_stock if canal in indicadores.columns else tiene_stock

    def _texto(row: pd.Series) -> str:
        canales = [c for c in CANALES_ORDEN if c in indicadores.columns and row.get(c, False)]
        return " | ".join(canales) if canales else "Sin canal identificado"

    out["canales_afectados_texto"] = indicadores.apply(_texto, axis=1)
    return out


def leer_hoja(xls: pd.ExcelFile, nombre: str) -> pd.DataFrame:
    if nombre not in xls.sheet_names:
        return pd.DataFrame()
    try:
        df = pd.read_excel(xls, sheet_name=nombre)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as exc:
        print(f"ADVERTENCIA: no se pudo leer la hoja {nombre}: {exc}")
        return pd.DataFrame()


def js_records(df: pd.DataFrame, max_rows: int | None = None) -> list[dict[str, Any]]:
    if df is None or df.empty:
        return []
    out = df.copy()
    if max_rows is not None:
        out = out.head(max_rows)
    out = out.replace([np.inf, -np.inf], np.nan)
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%Y-%m-%d %H:%M:%S")
        elif out[col].dtype == "object":
            out[col] = out[col].map(
                lambda x: x.strftime("%Y-%m-%d %H:%M:%S")
                if hasattr(x, "strftime") and not isinstance(x, str)
                else x
            )
    for col in out.columns:
        if str(out[col].dtype) == "category":
            out[col] = out[col].astype(object)
    return out.fillna("").to_dict(orient="records")


def first_existing(df: pd.DataFrame, names: list[str], default: Any = "") -> pd.Series:
    for name in names:
        if name in df.columns:
            return df[name]
    return pd.Series(default, index=df.index)


def find_summary_value(resumen: pd.DataFrame, key_terms: list[str], default: float = 0) -> float:
    if resumen.empty or not {"metrica", "valor"}.issubset(resumen.columns):
        return default
    metricas = resumen["metrica"].astype(str).str.lower()
    mask = pd.Series(True, index=resumen.index)
    for term in key_terms:
        mask &= metricas.str.contains(term.lower(), regex=False, na=False)
    if not mask.any():
        return default
    return safe_float(resumen.loc[mask, "valor"].iloc[0], default)



# Columnas de utilidad/ROI recientes (30 y 90 días) que trae la hoja
# 'por_producto' de base_utilidad_roi.xlsx (generada por build_utilidad_roi.py).
# Se usan tanto en 'Candidatos B2B' como en 'Compra consolidada por SKU'.
COLUMNAS_ROI_FINANCIERO = [
    "utilidad_30d",
    "roi_30d",
    "n_ventas_30d",
    "utilidad_90d",
    "roi_90d",
    "n_ventas_90d",
]


def cargar_roi_b2b() -> pd.DataFrame:
    """Carga la base semanal de ROI. Para B2B solo se usa el ROI por producto.
    El ritmo de venta proviene del dashboard y se congela durante stockout.
    También expone utilidad/ROI de 30 y 90 días para las tablas de
    Candidatos B2B y Compra consolidada por SKU."""
    if not ARCHIVO_ROI_B2B.exists():
        print(f"ADVERTENCIA B2B: no encontré {ARCHIVO_ROI_B2B.name}")
        return pd.DataFrame()
    try:
        df = pd.read_excel(ARCHIVO_ROI_B2B, sheet_name="por_producto")
        df.columns = [str(c).strip() for c in df.columns]
        if "sku_madre" in df.columns:
            df["sku_madre"] = df["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        columnas_numericas = ["roi_ponderado", "ritmo_ventas_unidades_dia"] + COLUMNAS_ROI_FINANCIERO
        for col in columnas_numericas:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df
    except Exception as exc:
        print(f"ADVERTENCIA B2B: no pude leer {ARCHIVO_ROI_B2B}: {exc}")
        return pd.DataFrame()


def adjuntar_financiero_roi(df: pd.DataFrame, roi_base: pd.DataFrame) -> pd.DataFrame:
    """Añade utilidad_30d/roi_30d/utilidad_90d/roi_90d (y sus conteos de
    ventas) a cualquier DataFrame que tenga una columna 'sku_madre', tomando
    los valores ya calculados en base_utilidad_roi.xlsx (hoja 'por_producto').

    No filtra ni reordena renglones. Los SKU sin match en esa base quedan
    en 0 (mismo criterio que el resto de columnas financieras del módulo
    B2B, p.ej. roi_marketplace / costo_unitario_odoo)."""
    out = df.copy()
    if out.empty:
        for col in COLUMNAS_ROI_FINANCIERO:
            out[col] = pd.Series(dtype="float64")
        return out

    out["sku_madre"] = out["sku_madre"].fillna("").astype(str).str.strip().str.upper()

    if roi_base is None or roi_base.empty or "sku_madre" not in roi_base.columns:
        for col in COLUMNAS_ROI_FINANCIERO:
            out[col] = 0.0
        return out

    cols_disponibles = [c for c in COLUMNAS_ROI_FINANCIERO if c in roi_base.columns]
    financiero = roi_base[["sku_madre"] + cols_disponibles].drop_duplicates("sku_madre")
    out = out.merge(financiero, on="sku_madre", how="left")
    for col in COLUMNAS_ROI_FINANCIERO:
        if col not in out.columns:
            out[col] = 0.0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    return out


# ============================================================
# UTILIDAD / ROI POR CANAL Y SERIE DIARIA DE VENTAS
# ============================================================
# 'por_producto' resuelve el ROI consolidado del producto (todos los
# canales juntos). La pestaña de cada canal necesita otra cosa: lo que ESE
# canal generó. Eso vive en la hoja 'por_producto_canal' de
# base_utilidad_roi.xlsx, que build_utilidad_roi.py ahora entrega con las
# ventanas de 30 y 90 días ya calculadas por canal.

VENTANAS_ROI_CANAL = [30, 90]

COLUMNAS_ROI_CANAL = [
    f"{campo}_canal_{dias}d"
    for dias in VENTANAS_ROI_CANAL
    for campo in ("utilidad", "roi", "n_ventas")
]

# Días de historia que se cargan para la gráfica de ventas por día de cada
# canal. El selector de rango de fechas del dashboard no puede ir más atrás
# que esto.
GRAFICA_VENTAS_DIAS = int(os.getenv("GRAFICA_VENTAS_DIAS", "180"))
GRAFICA_VENTAS_MAX_FILAS = int(os.getenv("GRAFICA_VENTAS_MAX_FILAS", "90000"))


def _leer_hoja_roi(nombre_hoja: str) -> pd.DataFrame:
    """Lee una hoja de base_utilidad_roi.xlsx sin romper si no existe.

    Un archivo generado con una versión anterior de build_utilidad_roi.py
    no traerá las hojas nuevas; en ese caso el dashboard se genera igual,
    solo que sin las columnas financieras por canal.
    """
    if not ARCHIVO_ROI_B2B.exists():
        return pd.DataFrame()
    try:
        xls_roi = pd.ExcelFile(ARCHIVO_ROI_B2B)
        if nombre_hoja not in xls_roi.sheet_names:
            print(
                f"ADVERTENCIA: {ARCHIVO_ROI_B2B.name} no tiene la hoja "
                f"'{nombre_hoja}'. Vuelve a correr build_utilidad_roi.py "
                f"para habilitar utilidad/ROI por canal."
            )
            return pd.DataFrame()
        df = pd.read_excel(ARCHIVO_ROI_B2B, sheet_name=nombre_hoja)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as exc:
        print(f"ADVERTENCIA: no pude leer '{nombre_hoja}' de {ARCHIVO_ROI_B2B.name}: {exc}")
        return pd.DataFrame()


def cargar_roi_por_canal() -> pd.DataFrame:
    """Utilidad y ROI de 30 y 90 días por (sku_madre, canal del dashboard).

    Un mismo canal del dashboard puede recibir varias cuentas del
    diccionario (p.ej. dos tiendas de Amazon). Se consolidan sumando la
    utilidad y recalculando el ROI como utilidad/base, no promediando los
    ROI de cada cuenta: el promedio de razones no es la razón del total.
    Si el archivo viene de una versión anterior sin columna base_roi, se
    cae al promedio ponderado por utilidad, que es el criterio que ya usa
    el resto del módulo financiero.
    """
    df = _leer_hoja_roi("por_producto_canal")
    if df.empty or "sku_madre" not in df.columns or "canal" not in df.columns:
        return pd.DataFrame()

    # La hoja existe desde siempre, pero las columnas por ventana son nuevas.
    # Sin ellas todo quedaría en 0/vacío sin explicación visible.
    ventanas_presentes = [
        dias for dias in VENTANAS_ROI_CANAL
        if f"utilidad_{dias}d" in df.columns or f"roi_{dias}d" in df.columns
    ]
    if not ventanas_presentes:
        print(
            f"ADVERTENCIA: 'por_producto_canal' de {ARCHIVO_ROI_B2B.name} no trae "
            f"utilidad/ROI de 30 ni 90 días por canal. Las columnas nuevas de las "
            f"tablas quedarán vacías; vuelve a correr build_utilidad_roi.py."
        )
        return pd.DataFrame()

    out = df.copy()
    out["sku_madre"] = out["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    out["canal_dashboard"] = out["canal"].apply(normalizar_canal)
    out = out[out["sku_madre"].ne("")]
    if out.empty:
        return pd.DataFrame()

    numericas = [
        c for c in out.columns
        if c.startswith(("utilidad_", "roi_", "n_ventas_", "unidades_", "base_roi_"))
    ]
    for col in numericas:
        out[col] = pd.to_numeric(out[col], errors="coerce")

    filas: list[dict[str, Any]] = []
    for (sku, canal), grupo in out.groupby(["sku_madre", "canal_dashboard"]):
        fila: dict[str, Any] = {"sku_madre": sku, "canal": canal}
        for dias in VENTANAS_ROI_CANAL:
            col_u, col_r = f"utilidad_{dias}d", f"roi_{dias}d"
            col_n = f"n_ventas_{dias}d"
            col_base, col_ubase = f"base_roi_{dias}d", f"utilidad_con_base_{dias}d"

            utilidad = safe_float(grupo[col_u].sum(), 0) if col_u in grupo else 0.0
            n_ventas = safe_float(grupo[col_n].sum(), 0) if col_n in grupo else 0.0

            roi: Any = None
            if col_base in grupo and col_ubase in grupo:
                # Base <= 0 significa reportes incoherentes en el origen
                # (ver _base_roi en build_utilidad_roi.py): mejor caer al
                # promedio ponderado que publicar un ROI absurdo.
                base = safe_float(grupo[col_base].sum(), 0)
                if base > 0:
                    roi = safe_float(grupo[col_ubase].sum(), 0) / base
            if roi is None and col_r in grupo:
                pesos = grupo[col_u].clip(lower=0) if col_u in grupo else None
                rois = grupo[col_r]
                if pesos is not None and safe_float(pesos.sum(), 0) > 0:
                    roi = float(np.average(rois.fillna(0), weights=pesos))
                elif rois.notna().any():
                    roi = float(rois.mean())

            fila[f"utilidad_canal_{dias}d"] = utilidad
            fila[f"roi_canal_{dias}d"] = roi if roi is not None else np.nan
            fila[f"n_ventas_canal_{dias}d"] = n_ventas
        filas.append(fila)

    return pd.DataFrame(filas)


def adjuntar_roi_por_canal(rotacion: pd.DataFrame, roi_canal: pd.DataFrame) -> pd.DataFrame:
    """Pega utilidad/ROI del canal a cada renglón (sku_madre, canal).

    Un SKU sin ventas registradas en ese canal dentro de la ventana queda
    con utilidad 0 (no hubo nada que sumar) y ROI vacío (no hay de dónde
    calcularlo). Se distingue a propósito: un ROI de 0% y un ROI sin dato
    no significan lo mismo, y las tablas del dashboard los muestran
    distinto ("0.00%" vs "—").
    """
    out = rotacion.copy()
    if out.empty:
        return out

    if roi_canal is None or roi_canal.empty:
        for col in COLUMNAS_ROI_CANAL:
            out[col] = 0.0 if col.startswith(("utilidad", "n_ventas")) else np.nan
        return out

    out["_sku_key"] = out["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    out["_canal_key"] = out["canal"].apply(normalizar_canal) if "canal" in out.columns else ""

    izq = roi_canal.rename(columns={"sku_madre": "_sku_key", "canal": "_canal_key"})
    out = out.merge(izq, on=["_sku_key", "_canal_key"], how="left")

    for dias in VENTANAS_ROI_CANAL:
        col_u, col_n = f"utilidad_canal_{dias}d", f"n_ventas_canal_{dias}d"
        col_r = f"roi_canal_{dias}d"
        for col in (col_u, col_n):
            if col not in out.columns:
                out[col] = 0.0
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
        if col_r not in out.columns:
            out[col_r] = np.nan
        out[col_r] = pd.to_numeric(out[col_r], errors="coerce")

    return out.drop(columns=["_sku_key", "_canal_key"], errors="ignore")


def construir_serie_ventas_canal(ventas_detalle: pd.DataFrame) -> pd.DataFrame:
    """Serie diaria por (fecha, canal, sku_madre) para la gráfica del canal.

    Combina dos fuentes a propósito:
      - unidades y monto vienen de 'ventas_detalle_dashboard' (script 02),
        que es la fuente operativa de ventas del dashboard;
      - utilidad y ROI vienen de 'ventas_diarias_canal' (build_utilidad_roi),
        que es la única fuente que tiene el dato financiero por venta.

    Se cruzan por (fecha, canal, sku_madre) con merge externo: un día puede
    tener utilidad reportada sin venta operativa conciliada o al revés, y
    esconder cualquiera de los dos casos daría una gráfica más limpia pero
    menos cierta. Los huecos quedan en 0 y el ROI sin base se deja vacío
    para que la línea muestre el corte en vez de fingir un 0%.
    """
    columnas = [
        "fecha", "canal", "sku_madre", "producto_madre",
        "unidades", "monto", "utilidad", "base_roi", "utilidad_con_base",
    ]

    operativa = pd.DataFrame()
    if ventas_detalle is not None and not ventas_detalle.empty:
        op = ventas_detalle.copy()
        op["fecha"] = pd.to_datetime(op.get("fecha"), errors="coerce").dt.normalize()
        op["canal"] = op["canal"].apply(normalizar_canal) if "canal" in op.columns else ""
        op["sku_madre"] = op.get("sku_madre", "").fillna("").astype(str).str.strip().str.upper()
        op["producto_madre"] = op.get("producto_madre", "").fillna("").astype(str)
        op["unidades"] = to_num(op.get("unidades", 0))
        op["monto"] = to_num(op.get("venta_total", 0))
        op = op[op["fecha"].notna() & op["sku_madre"].ne("")]
        if not op.empty:
            operativa = op.groupby(["fecha", "canal", "sku_madre"], as_index=False).agg(
                producto_madre=("producto_madre", lambda s: next((x for x in s if str(x).strip()), "")),
                unidades=("unidades", "sum"),
                monto=("monto", "sum"),
            )

    financiera = pd.DataFrame()
    diarias = _leer_hoja_roi("ventas_diarias_canal")
    if not diarias.empty and {"fecha", "sku_madre"}.issubset(diarias.columns):
        fin = diarias.copy()
        fin["fecha"] = pd.to_datetime(fin["fecha"], errors="coerce").dt.normalize()
        fin["canal"] = fin["canal"].apply(normalizar_canal) if "canal" in fin.columns else ""
        fin["sku_madre"] = fin["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        fin["producto_roi"] = fin.get("producto", "").fillna("").astype(str)
        for col in ("utilidad", "base_roi", "utilidad_con_base"):
            fin[col] = to_num(fin.get(col, 0))
        fin = fin[fin["fecha"].notna() & fin["sku_madre"].ne("")]
        if not fin.empty:
            financiera = fin.groupby(["fecha", "canal", "sku_madre"], as_index=False).agg(
                producto_roi=("producto_roi", lambda s: next((x for x in s if str(x).strip()), "")),
                utilidad=("utilidad", "sum"),
                base_roi=("base_roi", "sum"),
                utilidad_con_base=("utilidad_con_base", "sum"),
            )

    if operativa.empty and financiera.empty:
        return pd.DataFrame(columns=columnas)
    if operativa.empty:
        serie = financiera.rename(columns={"producto_roi": "producto_madre"})
        serie["unidades"] = 0.0
        serie["monto"] = 0.0
    elif financiera.empty:
        serie = operativa.copy()
        serie["utilidad"] = 0.0
        serie["base_roi"] = 0.0
        serie["utilidad_con_base"] = 0.0
    else:
        serie = operativa.merge(financiera, on=["fecha", "canal", "sku_madre"], how="outer")
        serie["producto_madre"] = (
            serie["producto_madre"].fillna("").astype(str).replace("", np.nan)
            .fillna(serie["producto_roi"].fillna("").astype(str))
            .fillna("")
        )
        serie = serie.drop(columns=["producto_roi"], errors="ignore")

    for col in ("unidades", "monto", "utilidad", "base_roi", "utilidad_con_base"):
        serie[col] = to_num(serie.get(col, 0))
    serie["producto_madre"] = serie.get("producto_madre", "").fillna("").astype(str)

    # Recorte de historia: la gráfica se ancla al último día CON datos, no a
    # la fecha del sistema, porque los reportes de canal casi nunca llegan
    # hasta hoy y un eje que termina en un hueco se lee como caída de ventas.
    serie = serie[serie["fecha"].notna()]
    if serie.empty:
        return pd.DataFrame(columns=columnas)
    ultimo_dia = serie["fecha"].max()
    serie = serie[serie["fecha"] > ultimo_dia - pd.Timedelta(days=GRAFICA_VENTAS_DIAS)]

    serie = serie.sort_values(["fecha", "canal", "sku_madre"], ascending=[False, True, True])
    if len(serie) > GRAFICA_VENTAS_MAX_FILAS:
        print(
            f"AVISO: la serie diaria por canal tiene {len(serie):,} renglones; "
            f"se recortan a los {GRAFICA_VENTAS_MAX_FILAS:,} más recientes. "
            f"Ajusta GRAFICA_VENTAS_DIAS o GRAFICA_VENTAS_MAX_FILAS si necesitas más."
        )
        serie = serie.head(GRAFICA_VENTAS_MAX_FILAS)

    serie["fecha"] = serie["fecha"].dt.strftime("%Y-%m-%d")
    return serie[columnas].reset_index(drop=True)


def _first_positive(row: pd.Series, names: list[str]) -> float:
    for name in names:
        if name in row.index:
            value = safe_float(row.get(name, 0), 0)
            if value > 0:
                return value
    return 0.0


def _b2b_sku_key(value: Any) -> str:
    text = str(value or "").strip().upper()
    if text.lower() in {"", "nan", "none"}:
        return ""
    return re.sub(r"\s+", "", text)


def _b2b_norm_location(value: Any) -> str:
    return re.sub(r"\s+", "", str(value or "").strip().lower())


def _b2b_payout(canal: str) -> tuple[float, str, bool]:
    value = B2B_PAYOUT_DIAS.get(canal)
    if value is None:
        return 30.0, "Supuesto 30d · dato incompleto", True
    return max(float(value), 0.0), "Config canal", False


def _b2b_roi_equivalente_hoy(roi_marketplace: float, dias_cobro: float) -> float:
    """ROI equivalente si el capital se recuperara HOY.

    ROI_hoy = (1 + ROI_MP) / (1 + r)^dias_cobro - 1
    """
    factor = (1.0 + max(B2B_TASA_DESCUENTO_DIARIA, 0.0)) ** max(float(dias_cobro), 0.0)
    if factor <= 0:
        return float(roi_marketplace)
    return (1.0 + float(roi_marketplace)) / factor - 1.0


def _b2b_roi_canal(row: pd.Series) -> tuple[float | None, str]:
    """Elige ROI Marketplace por canal, priorizando la ventana configurada."""
    if B2B_ROI_VENTANA in {"30", "30d"}:
        candidatos = [
            ("roi_canal_30d", "ROI canal 30d", "n_ventas_canal_30d"),
            ("roi_canal_90d", "ROI canal 90d respaldo", "n_ventas_canal_90d"),
        ]
    else:
        candidatos = [
            ("roi_canal_90d", "ROI canal 90d", "n_ventas_canal_90d"),
            ("roi_canal_30d", "ROI canal 30d respaldo", "n_ventas_canal_30d"),
        ]

    for roi_col, fuente, n_col in candidatos:
        if roi_col not in row.index or pd.isna(row.get(roi_col)):
            continue
        if n_col in row.index and safe_float(row.get(n_col), 0) <= 0:
            continue
        return float(row.get(roi_col)), fuente
    return None, "Sin ROI canal"


def _b2b_alias_map(
    sku_detalle: pd.DataFrame,
    sku_aliases_b2b: pd.DataFrame,
    skus_madre: set[str],
) -> tuple[dict[str, str], set[str]]:
    """Mapea default_code de Odoo -> SKU madre sin aceptar aliases ambiguos."""
    alias_sets: dict[str, set[str]] = {}

    def add(alias: Any, madre: Any) -> None:
        a = _b2b_sku_key(alias)
        m = _b2b_sku_key(madre)
        if not a or not m:
            return
        alias_sets.setdefault(a, set()).add(m)

    for madre in skus_madre:
        add(madre, madre)

    for frame in [sku_detalle, sku_aliases_b2b]:
        if frame is None or frame.empty or "sku_madre" not in frame.columns:
            continue
        candidatos_alias = [
            c for c in ["sku_sincronizado", "sku_original", "sku_alias", "alias_diccionario", "sku_visible"]
            if c in frame.columns
        ]
        for _, rr in frame.iterrows():
            madre = rr.get("sku_madre", "")
            add(madre, madre)
            for col in candidatos_alias:
                add(rr.get(col, ""), madre)

    ambiguos = {a for a, madres in alias_sets.items() if len(madres) > 1}
    mapa = {a: next(iter(madres)) for a, madres in alias_sets.items() if len(madres) == 1}
    return mapa, ambiguos


def cargar_stock_b2b_odoo_vivo(
    sku_detalle: pd.DataFrame,
    sku_aliases_b2b: pd.DataFrame,
    resumen_general_productos: pd.DataFrame,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Consulta stock.quant al generar el dashboard.

    Solo se usa para la pestaña B2B. Si Odoo no está disponible, el módulo cae
    a la base operativa ya generada por 01/02. Nunca se exponen credenciales al HTML.
    """
    timestamp = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    meta: dict[str, Any] = {
        "enabled": bool(B2B_ODOO_LIVE_STOCK),
        "status": "DESACTIVADO",
        "source": "base_dashboard_rotacion.xlsx",
        "timestamp": timestamp,
        "mapped_skus": 0,
        "available_units": 0.0,
        "unmapped_units": 0.0,
        "error": "",
    }

    if not B2B_ODOO_LIVE_STOCK:
        return pd.DataFrame(), meta

    if not all([ODOO_URL, ODOO_DB, ODOO_USER, ODOO_API_KEY]):
        meta.update({
            "status": "FALLBACK BASE · FALTAN CREDENCIALES",
            "error": "Faltan ODOO_URL / ODOO_DB / ODOO_USER / ODOO_API_KEY en .env",
        })
        return pd.DataFrame(), meta

    try:
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common", allow_none=True)
        uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
        if not uid:
            raise RuntimeError("Odoo rechazó la autenticación")
        models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object", allow_none=True)

        locations = models.execute_kw(
            ODOO_DB, uid, ODOO_API_KEY,
            "stock.location", "search_read",
            [[("usage", "=", "internal")]],
            {"fields": ["id", "name", "complete_name"], "limit": 0},
        )

        roots = [_b2b_norm_location(x) for x in B2B_ODOO_ROOTS]
        location_ids = []
        for loc in locations:
            name = _b2b_norm_location(loc.get("complete_name") or loc.get("name"))
            if any(name == root or name.startswith(root + "/") for root in roots):
                location_ids.append(int(loc["id"]))

        if not location_ids:
            raise RuntimeError("No se encontraron ubicaciones CUATI válidas en Odoo")

        quant_ids = models.execute_kw(
            ODOO_DB, uid, ODOO_API_KEY,
            "stock.quant", "search",
            [[("location_id", "in", location_ids), ("quantity", "!=", 0)]],
        )

        quants: list[dict[str, Any]] = []
        for i in range(0, len(quant_ids), 1000):
            quants.extend(models.execute_kw(
                ODOO_DB, uid, ODOO_API_KEY,
                "stock.quant", "read",
                [quant_ids[i:i + 1000]],
                {"fields": ["product_id", "location_id", "quantity", "reserved_quantity"]},
            ))

        product_ids = sorted({
            int(q["product_id"][0])
            for q in quants
            if q.get("product_id")
        })

        products_odoo: list[dict[str, Any]] = []
        for i in range(0, len(product_ids), 500):
            ids = product_ids[i:i + 500]
            try:
                batch = models.execute_kw(
                    ODOO_DB, uid, ODOO_API_KEY,
                    "product.product", "read", [ids],
                    {"fields": ["default_code", "name", "standard_price"]},
                )
            except Exception:
                batch = models.execute_kw(
                    ODOO_DB, uid, ODOO_API_KEY,
                    "product.product", "read", [ids],
                    {"fields": ["default_code", "name"]},
                )
            products_odoo.extend(batch)

        prod_by_id = {int(p["id"]): p for p in products_odoo}
        skus_madre = set()
        if resumen_general_productos is not None and not resumen_general_productos.empty and "sku_madre" in resumen_general_productos.columns:
            skus_madre = {
                _b2b_sku_key(x)
                for x in resumen_general_productos["sku_madre"]
                if _b2b_sku_key(x)
            }
        mapa_alias, ambiguos = _b2b_alias_map(sku_detalle, sku_aliases_b2b, skus_madre)

        acumulado: dict[str, dict[str, float]] = {}
        unmapped_units = 0.0

        for q in quants:
            pid = int(q["product_id"][0]) if q.get("product_id") else 0
            p = prod_by_id.get(pid, {})
            default_code = _b2b_sku_key(p.get("default_code", ""))
            madre = mapa_alias.get(default_code)

            qty = safe_float(q.get("quantity", 0), 0)
            reserved = safe_float(q.get("reserved_quantity", 0), 0)
            available = qty - reserved if B2B_ODOO_STOCK_MODE == "available" else qty

            if not madre or default_code in ambiguos:
                unmapped_units += max(available, 0.0)
                continue

            slot = acumulado.setdefault(madre, {
                "stock_odoo_live": 0.0,
                "costo_x_unidad": 0.0,
                "costo_peso": 0.0,
                "productos_odoo_mapeados": 0.0,
            })
            slot["stock_odoo_live"] += available
            cost = safe_float(p.get("standard_price", 0), 0)
            if cost > 0 and available > 0:
                slot["costo_x_unidad"] += cost * available
                slot["costo_peso"] += available
            slot["productos_odoo_mapeados"] += 1

        rows = []
        for madre, vals in acumulado.items():
            stock = max(vals["stock_odoo_live"], 0.0)
            cost = vals["costo_x_unidad"] / vals["costo_peso"] if vals["costo_peso"] > 0 else 0.0
            rows.append({
                "sku_madre": madre,
                "stock_odoo_live": stock,
                "costo_unitario_odoo_live": cost,
                "productos_odoo_mapeados": int(vals["productos_odoo_mapeados"]),
                "fuente_stock_b2b": "Odoo API · stock.quant available",
            })

        df = pd.DataFrame(rows)
        meta.update({
            "status": "CONECTADO",
            "source": "Odoo API · stock.quant",
            "mapped_skus": int(len(df)),
            "available_units": float(df["stock_odoo_live"].sum()) if not df.empty else 0.0,
            "unmapped_units": float(unmapped_units),
        })
        return df, meta

    except Exception as exc:
        meta.update({
            "status": "FALLBACK BASE · ERROR ODOO",
            "error": str(exc),
        })
        print(f"ADVERTENCIA B2B ODOO: {exc}. Se usará el stock de base_dashboard_rotacion.xlsx")
        return pd.DataFrame(), meta


def construir_portafolio_b2b(
    rotacion: pd.DataFrame,
    productos: pd.DataFrame,
    roi_base: pd.DataFrame,
    roi_por_canal: pd.DataFrame,
    stock_odoo_live: pd.DataFrame,
    b2b_live_meta: dict[str, Any],
) -> pd.DataFrame:
    """Excedente B2B + costo de oportunidad Marketplace llevado a VPN.

    Marketplace se protege primero por canal. B2B solo usa el excedente de
    bodega. El VPN Marketplace conserva la referencia económica, pero cada pieza
    B2B queda sujeta a un ROI comercial mínimo de 10%; el precio mínimo se
    calcula sobre el costo Odoo con ese ROI o con un costo de oportunidad mayor.
    """
    if rotacion is None or rotacion.empty:
        return pd.DataFrame()

    work = rotacion.copy()
    work["sku_madre"] = work["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    work["canal"] = work["canal"].apply(normalizar_canal)

    # ROI canal ya consolidado por la rutina financiera existente del dashboard.
    roi_map: dict[tuple[str, str], dict[str, Any]] = {}
    if roi_por_canal is not None and not roi_por_canal.empty:
        rp = roi_por_canal.copy()
        rp["sku_madre"] = rp["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        rp["canal"] = rp["canal"].apply(normalizar_canal)
        for _, rr in rp.iterrows():
            roi_map[(rr["sku_madre"], rr["canal"])] = rr.to_dict()

    product_map: dict[str, dict[str, Any]] = {}
    if productos is not None and not productos.empty and "sku_madre" in productos.columns:
        pp = productos.copy()
        pp["sku_madre"] = pp["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        product_map = pp.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    roi_product_map: dict[str, dict[str, Any]] = {}
    if roi_base is not None and not roi_base.empty and "sku_madre" in roi_base.columns:
        rb = roi_base.copy()
        rb["sku_madre"] = rb["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        roi_product_map = rb.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    live_map: dict[str, dict[str, Any]] = {}
    if stock_odoo_live is not None and not stock_odoo_live.empty:
        live = stock_odoo_live.copy()
        live["sku_madre"] = live["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        live_map = live.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    rows: list[dict[str, Any]] = []

    for sku, g in work.groupby("sku_madre", dropna=False):
        sku = str(sku or "").strip().upper()
        if not sku:
            continue

        producto = next((str(x) for x in g.get("producto_madre", pd.Series(dtype=str)) if str(x).strip()), "")
        rb = roi_product_map.get(sku, {})
        if not producto:
            producto = str(rb.get("producto", "") or "")

        p = product_map.get(sku, {})
        live = live_map.get(sku, {})

        stock_base = max(safe_float(g["inventario_odoo"].sum(), 0), 0)
        if live and "stock_odoo_live" in live:
            stock_odoo = max(safe_float(live.get("stock_odoo_live", 0), 0), 0)
            fuente_stock = str(live.get("fuente_stock_b2b", "Odoo API"))
        else:
            stock_odoo = stock_base
            fuente_stock = "Dashboard 01/02 · inventario_odoo"

        stock_full = max(safe_float(g["inventario_full"].sum(), 0), 0)
        stock_transito = max(safe_float(g["inventario_transito"].sum(), 0), 0)
        stock_total = stock_odoo + stock_full + stock_transito

        # Costo: primero el costo Odoo ya consolidado por 01/02; API live como respaldo.
        costo = _first_positive(pd.Series(p), [
            "costo_unitario_odoo_cuati",
            "costo_unitario_odoo",
            "costo_promedio_odoo",
            "costo_odoo_unitario",
        ])
        fuente_costo = "Odoo CUATI · base 01/02" if costo > 0 else "Pendiente"
        if costo <= 0 and live:
            costo = safe_float(live.get("costo_unitario_odoo_live", 0), 0)
            if costo > 0:
                fuente_costo = "Odoo API · standard_price respaldo"

        canales_detalle: list[dict[str, Any]] = []
        warnings: list[str] = []
        reserva_requerida_total = 0

        for canal in CANALES_ORDEN:
            gc = g[g["canal"] == canal]
            if gc.empty:
                continue
            rr = gc.iloc[0]

            forecast = max(safe_float(rr.get("demanda_diaria_ponderada", 0), 0), 0)
            if forecast <= 0:
                forecast = max(safe_float(rr.get("ventas_90d_unidades", 0), 0) / 90.0, 0)

            full = max(safe_float(rr.get("inventario_full", 0), 0), 0)
            transito = max(safe_float(rr.get("inventario_transito", 0), 0), 0)
            necesidad = math.ceil(forecast * B2B_HORIZONTE_DIAS - 1e-12) if forecast > 0 else 0
            fuera_bodega = math.floor(full + transito + 1e-9)
            reserva = max(int(necesidad - fuera_bodega), 0)
            reserva_requerida_total += reserva

            rfin = pd.Series(roi_map.get((sku, canal), {}))
            roi_mp, fuente_roi = _b2b_roi_canal(rfin) if not rfin.empty else (None, "Sin ROI canal")
            payout, fuente_payout, payout_fallback = _b2b_payout(canal)

            if forecast > 0 and roi_mp is None:
                warnings.append(f"{canal}: demanda sin ROI canal")
            if forecast > 0 and payout_fallback:
                warnings.append(f"{canal}: payout no configurado; se usó {B2B_PAYOUT_FALLBACK_DIAS:g}d")

            posicion_fin_reserva = int(fuera_bodega + reserva)
            dia_primera = None
            roi_primera_hoy = None
            if forecast > 0 and roi_mp is not None:
                dia_primera = max(math.ceil((posicion_fin_reserva + 1) / forecast - 1e-12), 1)
                roi_primera_hoy = _b2b_roi_equivalente_hoy(roi_mp, dia_primera + payout)

            canales_detalle.append({
                "canal": canal,
                "forecast_dia": forecast,
                "necesidad_lead_time": int(necesidad),
                "full": full,
                "transito": transito,
                "reserva_bodega": int(reserva),
                "roi_marketplace": roi_mp,
                "roi_fuente": fuente_roi,
                "payout_dias": payout,
                "payout_fuente": fuente_payout,
                "payout_es_fallback": payout_fallback,
                "posicion_fin_reserva": posicion_fin_reserva,
                "dia_primera_venta_excedente": dia_primera,
                "roi_equivalente_hoy_primera_excedente": roi_primera_hoy,
            })

        reserva_cubierta = min(reserva_requerida_total, math.floor(stock_odoo + 1e-9))
        excedente = max(math.floor(stock_odoo - reserva_requerida_total + 1e-9), 0)

        # Oportunidades futuras que el excedente podría capturar si se conserva para Marketplace.
        candidatos: list[dict[str, Any]] = []
        if excedente > 0:
            for c in canales_detalle:
                forecast = safe_float(c.get("forecast_dia", 0), 0)
                roi_mp = c.get("roi_marketplace")
                if forecast <= 0 or roi_mp is None:
                    continue
                payout = safe_float(c.get("payout_dias", 0), 0)
                base_pos = int(c.get("posicion_fin_reserva", 0))
                for j in range(1, excedente + 1):
                    posicion = base_pos + j
                    dia_venta = max(math.ceil(posicion / forecast - 1e-12), 1)
                    dia_cobro = dia_venta + payout
                    roi_hoy = _b2b_roi_equivalente_hoy(float(roi_mp), dia_cobro)
                    candidatos.append({
                        "canal": c["canal"],
                        "roi_marketplace": float(roi_mp),
                        "roi_equivalente_hoy": roi_hoy,
                        "dia_venta": int(dia_venta),
                        "dia_cobro": float(dia_cobro),
                    })

        # Si se conserva el excedente, se asigna a las mejores oportunidades Marketplace.
        pool = sorted(candidatos, key=lambda x: x["roi_equivalente_hoy"], reverse=True)[:excedente]
        # B2B desplaza primero las oportunidades menos valiosas dentro de ese pool.
        desplazadas = sorted(pool, key=lambda x: x["roi_equivalente_hoy"])

        # Curva comercial B2B. Cada pieza tiene un piso de ROI de 10%; si el
        # costo de oportunidad VPN es mayor, se respeta el mayor de ambos.
        # Si falta una referencia Marketplace para una pieza del excedente, su
        # costo de oportunidad económico se toma como 0 y el piso de 10% sigue
        # permitiendo cotizarla sin inventar un ROI Marketplace.
        curva: list[dict[str, Any]] = []
        suma_roi_economico = 0.0
        suma_roi_minimo = 0.0
        suma_roi_mp = 0.0
        n_roi_mp = 0
        for q in range(1, excedente + 1):
            opp = desplazadas[q - 1] if q - 1 < len(desplazadas) else None
            marginal_economico = float(opp["roi_equivalente_hoy"]) if opp else 0.0
            marginal_minimo = max(B2B_ROI_MINIMO, marginal_economico)
            suma_roi_economico += marginal_economico
            suma_roi_minimo += marginal_minimo
            if opp is not None:
                suma_roi_mp += float(opp["roi_marketplace"])
                n_roi_mp += 1
            roi_economico_prom = suma_roi_economico / q
            roi_min_prom = suma_roi_minimo / q
            precio_min = costo * (1 + roi_min_prom) if costo > 0 else None
            curva.append({
                "q": q,
                # Compatibilidad: el campo histórico ahora representa el mínimo comercial.
                "roi_b2b_equivalente_hoy": roi_min_prom,
                "roi_b2b_minimo": roi_min_prom,
                "roi_b2b_economico_vpn": roi_economico_prom,
                "roi_oportunidad_marginal": marginal_economico,
                "roi_b2b_marginal_minimo": marginal_minimo,
                "roi_marketplace_promedio": (suma_roi_mp / n_roi_mp) if n_roi_mp else None,
                "precio_equivalente_hoy_unitario": precio_min,
                "precio_minimo_b2b_unitario": precio_min,
                "canal_oportunidad": opp["canal"] if opp else "Sin referencia Marketplace",
                "roi_marketplace_referencia": opp["roi_marketplace"] if opp else None,
                "dia_venta_marketplace": opp["dia_venta"] if opp else None,
                "dia_cobro_marketplace": opp["dia_cobro"] if opp else None,
                "piso_roi_b2b": B2B_ROI_MINIMO,
            })

        valoradas = min(len(desplazadas), excedente)
        roi_total = curva[-1]["roi_b2b_minimo"] if curva else None
        roi_economico_total = curva[-1]["roi_b2b_economico_vpn"] if curva else None
        precio_total = curva[-1]["precio_minimo_b2b_unitario"] if curva else None
        roi_mp_prom = curva[-1]["roi_marketplace_promedio"] if curva else None
        canales_opp = " | ".join(sorted({x["canal_oportunidad"] for x in curva if x["canal_oportunidad"] != "Sin referencia Marketplace"})) if curva else ""

        for c in canales_detalle:
            vals = [x["roi_equivalente_hoy"] for x in pool if x["canal"] == c["canal"]]
            c["piezas_excedente_asignadas_mp"] = len(vals)
            c["roi_equivalente_hoy_prom_excedente"] = float(np.mean(vals)) if vals else None

        if stock_odoo <= 0:
            estado = "SIN STOCK ODOO"
        elif excedente <= 0:
            estado = "SIN EXCEDENTE B2B"
        elif costo <= 0:
            estado = "ROI MÍN. 10% · PENDIENTE COSTO ODOO"
        elif valoradas < excedente:
            estado = "LISTO · PISO 10% · VPN PARCIAL"
        elif warnings:
            estado = "LISTO · PISO 10% · REVISAR SUPUESTOS"
        else:
            estado = "LISTO · PISO 10%"

        rows.append({
            "sku_madre": sku,
            "producto_madre": producto,
            "stock_total": stock_total,
            "stock_odoo": stock_odoo,
            "stock_full": stock_full,
            "stock_transito": stock_transito,
            "fuente_stock": fuente_stock,
            "reserva_marketplace_30d": int(reserva_cubierta),
            "reserva_marketplace_requerida": int(reserva_requerida_total),
            "excedente_economico": int(excedente),
            "b2b_disponible": int(excedente),
            "excedente_valorado_vpn": int(valoradas),
            "costo_unitario_odoo": costo,
            "fuente_costo": fuente_costo,
            "roi_marketplace": roi_mp_prom,
            "roi_marketplace_referencia": roi_mp_prom,
            "roi_b2b_equivalente_hoy_total": roi_total,
            "roi_b2b_minimo_total": roi_total,
            "roi_b2b_economico_vpn_total": roi_economico_total,
            "roi_minimo_b2b": B2B_ROI_MINIMO,
            "precio_equivalente_hoy_total": precio_total,
            "precio_minimo_b2b_total": precio_total,
            "canales_oportunidad": canales_opp,
            "estado_b2b": estado,
            "curva_b2b": curva,
            "canales_b2b": canales_detalle,
            "warnings_b2b": sorted(set(warnings)),
            "fuente_roi": f"{ARCHIVO_ROI_B2B.name} · ROI canal {B2B_ROI_VENTANA}d",
            "odoo_live_status": str(b2b_live_meta.get("status", "")),
            "odoo_live_timestamp": str(b2b_live_meta.get("timestamp", "")),
        })

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    out = adjuntar_financiero_roi(out, roi_base)
    return out.sort_values(["b2b_disponible", "stock_odoo"], ascending=[False, False])


def enrich_rotacion(rotacion: pd.DataFrame) -> pd.DataFrame:
    if rotacion.empty:
        return rotacion
    out = rotacion.copy()
    numeric = [
        "inventario_odoo",
        "inventario_transito",
        "inventario_full",
        "inventario_total",
        "ventas_10d_unidades",
        "ventas_30d_unidades",
        "ventas_90d_unidades",
        "demanda_diaria_ponderada",
        "venta_diaria_10d",
        "venta_diaria_30d",
        "venta_diaria_90d",
        "ventas_full_90d",
        "ventas_drop_90d",
        "ventas_pendiente_odoo_90d",
        "venta_diaria_calendario",
        "venta_diaria_full",
        "cobertura_full_dias",
        "cobertura_total_dias",
        "costo_unitario_inventario",
        "valor_odoo",
        "valor_transito",
        "valor_full",
        "valor_inventario_canal",
        "valor_mas_90d",
        "unidades_mas_90d",
    ]
    for col in numeric:
        if col not in out.columns:
            out[col] = 0
        out[col] = pd.to_numeric(out[col], errors="coerce")

    if "canal" not in out.columns:
        out["canal"] = "Sin identificar"
    out["canal"] = out["canal"].apply(normalizar_canal)
    if "producto_madre" not in out.columns:
        out["producto_madre"] = ""
    if "sku_madre" not in out.columns:
        out["sku_madre"] = ""
    if "demanda_tipo" not in out.columns:
        out["demanda_tipo"] = np.where(
            out["ventas_90d_unidades"].fillna(0) <= 12,
            "INTERMITENTE",
            "CONTINUA",
        )
    if "accion_preliminar" not in out.columns:
        out["accion_preliminar"] = "SIN ACCIÓN"
    if "justificacion" not in out.columns:
        out["justificacion"] = ""

    # Cantidades preliminares para apoyar la decisión; no ejecutan movimientos.
    # Política Odoo -> Full: proyectar el consumo durante 5 días de traslado y
    # dejar 30 días de cobertura disponibles al momento del arribo.
    out["canal_tiene_full"] = out["canal"].isin(CANALES_CON_FULL)
    out["objetivo_full_dias"] = np.where(
        out["canal_tiene_full"], OBJETIVO_FULL_DIAS, 0
    )
    out["dias_transito_full"] = np.where(
        out["canal_tiene_full"], DIAS_TRANSITO_ODOO_FULL, 0
    )
    objetivo_historico = np.ceil(
        out["venta_diaria_calendario"].fillna(0) * OBJETIVO_FULL_DIAS
    )
    out["consumo_estimado_transito"] = np.where(
        out["canal_tiene_full"] & (out["ventas_90d_unidades"].fillna(0) > 0),
        np.ceil(out["venta_diaria_calendario"].fillna(0) * DIAS_TRANSITO_ODOO_FULL),
        0,
    )
    out["full_proyectado_arribo"] = np.where(
        out["canal_tiene_full"],
        (out["inventario_full"].fillna(0) - out["consumo_estimado_transito"]).clip(lower=0),
        out["inventario_full"].fillna(0),
    )
    out["objetivo_full_piezas"] = np.where(
        ~out["canal_tiene_full"],
        0,
        np.where(
            out["ventas_90d_unidades"].fillna(0) > 0,
            objetivo_historico,
            MAX_PIEZAS_PRODUCTO_NUEVO_FULL,
        ),
    )
    faltante_full = (
        out["objetivo_full_piezas"]
        - out["full_proyectado_arribo"]
        - out["inventario_transito"].fillna(0)
    ).clip(lower=0)
    out["transferencia_sugerida"] = np.minimum(
        np.ceil(faltante_full), out["inventario_odoo"].fillna(0)
    ).clip(lower=0)
    out["cobertura_proyectada_arribo_dias"] = np.where(
        out["venta_diaria_calendario"].fillna(0) > 0,
        (
            out["full_proyectado_arribo"]
            + out["inventario_transito"].fillna(0)
            + out["transferencia_sugerida"]
        ) / out["venta_diaria_calendario"].fillna(0),
        np.nan,
    )
    out["check_manual_full"] = np.where(
        out["canal_tiene_full"]
        & (out["ventas_90d_unidades"].fillna(0) <= 0)
        & (out["inventario_odoo"].fillna(0) > 0),
        "SI - PRODUCTO NUEVO",
        "NO",
    )

    def criterio_full(row: pd.Series) -> str:
        if not bool(row.get("canal_tiene_full", False)):
            return "Canal sin operación Full configurada; no se sugiere traslado."
        if safe_float(row.get("inventario_odoo", 0)) <= 0:
            return "Sin stock disponible en Odoo para enviar."
        if safe_float(row.get("ventas_90d_unidades", 0)) <= 0:
            return (
                f"Sin ventas históricas: máximo {MAX_PIEZAS_PRODUCTO_NUEVO_FULL} "
                "piezas al arribo. No se estima consumo y el CHECK MANUAL es obligatorio."
            )
        if safe_float(row.get("transferencia_sugerida", 0)) > 0:
            return (
                f"Se descuentan {DIAS_TRANSITO_ODOO_FULL} días de consumo del Full actual "
                f"y se completa a {OBJETIVO_FULL_DIAS} días disponibles al arribo."
            )
        return f"El Full proyectado al arribo + tránsito cubre el objetivo de {OBJETIVO_FULL_DIAS} días."

    out["criterio_transferencia_full"] = out.apply(criterio_full, axis=1)

    manual_mask = out["check_manual_full"].astype(str).str.startswith("SI") & (out["transferencia_sugerida"] > 0)
    transfer_mask = (~manual_mask) & (out["transferencia_sugerida"] > 0)
    out.loc[manual_mask, "accion_preliminar"] = "CHECK MANUAL - PRODUCTO NUEVO A FULL"
    out.loc[transfer_mask, "accion_preliminar"] = f"TRANSFERIR A FULL - {OBJETIVO_FULL_DIAS} DÍAS AL ARRIBO"

    objetivo_compra = out["venta_diaria_calendario"].fillna(0) * COBERTURA_COMPRA_DIAS
    out["compra_sugerida"] = np.ceil(
        (objetivo_compra - out["inventario_total"].fillna(0)).clip(lower=0)
    )

    def nivel_riesgo(row: pd.Series) -> str:
        accion = str(row.get("accion_preliminar", "")).upper()
        demanda = str(row.get("demanda_tipo", "")).upper()
        cobertura_full = row.get("cobertura_full_dias")
        cobertura_total = row.get("cobertura_total_dias")
        if "COMPRA" in accion or (
            pd.notna(cobertura_total) and safe_float(cobertura_total) < 15
        ):
            return "CRÍTICO"
        if "TRANSFERIR" in accion or (
            pd.notna(cobertura_full) and safe_float(cobertura_full) < LEAD_TIME_FULL_DIAS
        ):
            return "ALTO"
        if demanda == "INTERMITENTE":
            return "INTERMITENTE"
        if pd.notna(cobertura_total) and safe_float(cobertura_total) > 90:
            return "SOBRESTOCK"
        return "SANO"

    out["nivel_riesgo"] = out.apply(nivel_riesgo, axis=1)
    out["kam"] = out["canal"].map(KAM_POR_CANAL).fillna("Por asignar")
    out = agregar_cobertura_kam(out)
    return out


def agregar_cobertura_kam(out: pd.DataFrame) -> pd.DataFrame:
    """Semáforo de cobertura y acción en lenguaje simple para el KAM.

    estado_cobertura / prioridad_cobertura (1 = atender primero):
      1 Sin stock con ventas · 2 Quiebre < 7 días · 3 Full se agota antes de
      que llegue un envío · 4 Cobertura baja 7-15 días · 5 Sana 15-60 ·
      6 Alta 60-90 · 7 Exceso > 90 · 8 Stock sin ventas · 9 Sin stock ni ventas.
    """
    out = out.copy()
    inv = out["inventario_total"].fillna(0)
    ventas = out["ventas_90d_unidades"].fillna(0)
    cob = pd.to_numeric(out["cobertura_total_dias"], errors="coerce")
    cob_full = pd.to_numeric(out["cobertura_full_dias"], errors="coerce")
    full_canal = out["canal"].isin(CANALES_CON_FULL)

    condiciones = [
        (inv <= 0) & (ventas > 0),
        (inv > 0) & (ventas > 0) & (cob < COBERTURA_QUIEBRE_DIAS),
        full_canal & (ventas > 0) & (out["ventas_full_90d"].fillna(0) > 0)
        & (cob_full < DIAS_TRANSITO_ODOO_FULL) & (out["inventario_odoo"].fillna(0) + out["inventario_transito"].fillna(0) > 0),
        (inv > 0) & (ventas > 0) & (cob < COBERTURA_BAJA_DIAS),
        (inv > 0) & (ventas > 0) & (cob <= COBERTURA_ALTA_DIAS),
        (inv > 0) & (ventas > 0) & (cob <= COBERTURA_EXCESO_DIAS),
        (inv > 0) & (ventas > 0),
        (inv > 0) & (ventas <= 0),
    ]
    estados = [
        "Sin stock · con ventas",
        f"Quiebre < {COBERTURA_QUIEBRE_DIAS} días",
        "Full se agota antes del envío",
        f"Cobertura baja {COBERTURA_QUIEBRE_DIAS}–{COBERTURA_BAJA_DIAS} días",
        f"Cobertura sana {COBERTURA_BAJA_DIAS}–{COBERTURA_ALTA_DIAS} días",
        f"Cobertura alta {COBERTURA_ALTA_DIAS}–{COBERTURA_EXCESO_DIAS} días",
        f"Exceso > {COBERTURA_EXCESO_DIAS} días",
        "Stock sin ventas 90d",
    ]
    out["estado_cobertura"] = np.select(condiciones, estados, default="Sin stock ni ventas")
    out["prioridad_cobertura"] = np.select(condiciones, list(range(1, 9)), default=9)
    out["dias_para_quiebre"] = np.where(ventas > 0, cob, np.nan)
    hoy = pd.Timestamp.now().normalize()
    out["fecha_estimada_quiebre"] = [
        (hoy + pd.Timedelta(days=float(d))).strftime("%d/%m/%Y") if pd.notna(d) and d < 365 else ""
        for d in out["dias_para_quiebre"]
    ]

    def accion(r: pd.Series) -> str:
        env = safe_float(r.get("transferencia_sugerida"), 0)
        compra = safe_float(r.get("compra_sugerida"), 0)
        p = int(r.get("prioridad_cobertura", 9))
        if str(r.get("check_manual_full", "")).upper().startswith("SI") and env > 0:
            return f"Validar envío inicial de {env:.0f} pzas a Full (producto nuevo)"
        if p in (1, 2, 3, 4) and env > 0:
            return f"Enviar {env:.0f} pzas a Full hoy"
        if p in (1, 2, 4) and compra > 0:
            return f"Pedir compra: faltan {compra:.0f} pzas para {COBERTURA_COMPRA_DIAS} días"
        if p == 3:
            return "Full se agota: priorizar el traslado en curso"
        if env > 0:
            return f"Enviar {env:.0f} pzas a Full"
        if p == 7:
            return "No enviar ni comprar: hay exceso; evaluar promoción o reasignar"
        if p == 8:
            return "Sin ventas en 90 días: revisar publicación, precio o liquidar"
        if p == 1:
            return "Sin stock: revisar reabasto"
        return "Sin acción"

    out["accion_kam"] = out.apply(accion, axis=1)
    return out


def fallback_rotacion(
    stock_canal: pd.DataFrame,
    ventas_canal: pd.DataFrame,
    productos: pd.DataFrame,
) -> pd.DataFrame:
    """Construye una vista básica si la base aún no tiene rotacion_por_canal."""
    if stock_canal.empty:
        return pd.DataFrame()

    stock = stock_canal.copy()
    rows: list[dict[str, Any]] = []
    map_cols = {
        "stock_amazon_fba": "Amazon",
        "stock_meli_full": "Mercado Libre",
        "stock_walmart_wfs": "Walmart",
        "stock_liverpool_99min": "Liverpool",
        "stock_odoo_cuautitlan": "General",
    }
    product_map = {}
    if not productos.empty and "sku_madre" in productos.columns:
        product_map = productos.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    ventas_lookup: dict[tuple[str, str], float] = {}
    if not ventas_canal.empty:
        vc = ventas_canal.copy()
        if "canal" in vc.columns:
            vc["canal_norm"] = vc["canal"].apply(normalizar_canal)
        else:
            vc["canal_norm"] = "Sin identificar"
        if "unidades" not in vc.columns:
            vc["unidades"] = 0
        vc["unidades"] = pd.to_numeric(vc["unidades"], errors="coerce").fillna(0)
        if "sku_madre" in vc.columns:
            grouped = vc.groupby(["sku_madre", "canal_norm"], as_index=False)["unidades"].sum()
            ventas_lookup = {
                (str(r["sku_madre"]), str(r["canal_norm"])): safe_float(r["unidades"])
                for _, r in grouped.iterrows()
            }

    for _, r in stock.iterrows():
        sku = str(r.get("sku_madre", ""))
        prod = str(r.get("producto_madre", "") or product_map.get(sku, {}).get("producto_madre", ""))
        for col, canal in map_cols.items():
            if col not in stock.columns:
                continue
            qty = max(safe_float(r.get(col, 0)), 0)
            ventas = ventas_lookup.get((sku, canal), 0)
            rows.append(
                {
                    "sku_madre": sku,
                    "producto_madre": prod,
                    "canal": canal,
                    "inventario_odoo": qty if canal == "General" else 0,
                    "inventario_transito": 0,
                    "inventario_full": qty if canal != "General" else 0,
                    "inventario_total": qty,
                    "ventas_90d_unidades": ventas,
                    "ventas_full_90d": ventas if canal != "General" else 0,
                    "ventas_drop_90d": 0,
                    "ventas_pendiente_odoo_90d": 0,
                    "venta_diaria_calendario": ventas / 90,
                    "venta_diaria_full": ventas / 90 if canal != "General" else 0,
                    "cobertura_full_dias": qty / (ventas / 90) if ventas > 0 and canal != "General" else np.nan,
                    "cobertura_total_dias": qty / (ventas / 90) if ventas > 0 else np.nan,
                    "demanda_tipo": "INTERMITENTE" if ventas <= 12 else "CONTINUA",
                    "accion_preliminar": "SIN ACCIÓN",
                    "justificacion": "Vista de compatibilidad: la base todavía no contiene rotación por canal.",
                }
            )
    return pd.DataFrame(rows)


# ============================================================
# CARGA Y PREPARACIÓN DE DATOS
# ============================================================

if not ARCHIVO_BASE.exists():
    raise FileNotFoundError(
        f"No encontré {ARCHIVO_BASE}. Primero ejecuta 02_generar_metricas_rotacion_canales.py"
    )

xls = pd.ExcelFile(ARCHIVO_BASE)

resumen = leer_hoja(xls, "resumen")
productos = leer_hoja(xls, "dashboard_productos")
alertas_productos = leer_hoja(xls, "dashboard_alertas")
stock_canal = leer_hoja(xls, "dashboard_stock_canal")
ventas_canal = leer_hoja(xls, "dashboard_ventas_canal")
rotacion = leer_hoja(xls, "rotacion_por_canal")
control_total = leer_hoja(xls, "control_total_canal")
inventario_base = leer_hoja(xls, "inventario_canal_base")
traslados = leer_hoja(xls, "traslados_full")
traslados_excluidos = leer_hoja(xls, "traslados_excluidos")
ventas_no = leer_hoja(xls, "ventas_no_vinculadas")
stock_no = leer_hoja(xls, "stock_no_vinculado")

# Para poder mostrar "SKU sin vincular" dentro de cada página de canal
# (no solo en Calidad de datos global).
if not ventas_no.empty:
    ventas_no["canal_normalizado"] = (
        ventas_no["canal"].apply(normalizar_canal)
        if "canal" in ventas_no.columns
        else "Sin identificar"
    )
stock_no = anotar_canales_stock_no_vinculado(stock_no)
arribos = leer_hoja(xls, "arribos_odoo_compras")
redistribuciones = leer_hoja(xls, "redistribucion_interna")
reparticion_resumen = leer_hoja(xls, "reparticion_resumen")
reparticion_detalle = leer_hoja(xls, "reparticion_detalle")
reparticion_perfiles_categoria = leer_hoja(xls, "reparticion_perfiles_cat")
sku_detalle = leer_hoja(xls, "sku_por_canal_detalle")
sku_aliases_b2b = leer_hoja(xls, "sku_marketplace_aliases")
ventas_detalle = leer_hoja(xls, "ventas_detalle_dashboard")
antiguedad_capas = leer_hoja(xls, "antiguedad_capas")
movimientos_auditoria = leer_hoja(xls, "movimientos_odoo_auditoria")
movimientos_auditoria_resumen = leer_hoja(xls, "movimientos_odoo_resumen")
publicaciones_sku_canal = leer_hoja(xls, "publicaciones_sku_canal")
publicaciones_log = leer_hoja(xls, "publicaciones_log")
publicaciones_sin_sku = leer_hoja(xls, "publicaciones_sin_sku_madre")
publicaciones_sin_sku_resumen = leer_hoja(xls, "publicaciones_sin_sku_resumen")

# Compatibilidad de trazabilidad para traslados excluidos.
# En el Excel operativo, los pickings excluidos pueden conservar el nombre
# técnico de Odoo `origin`; el dashboard lo expone como `documento_origen`.
if not traslados_excluidos.empty:
    traslados_excluidos = traslados_excluidos.copy()
    if "documento_origen" not in traslados_excluidos.columns:
        traslados_excluidos["documento_origen"] = ""
    if "origin" in traslados_excluidos.columns:
        actual = traslados_excluidos["documento_origen"].fillna("").astype(str).str.strip()
        desde_odoo = traslados_excluidos["origin"].fillna("").astype(str).str.strip()
        traslados_excluidos.loc[actual.eq(""), "documento_origen"] = desde_odoo[actual.eq("")]

if rotacion.empty:
    rotacion = fallback_rotacion(stock_canal, ventas_canal, productos)
rotacion = enrich_rotacion(rotacion)

# Utilidad y ROI de 30 y 90 días del canal, pegados a cada renglón
# (sku_madre, canal) para que las tablas de la pestaña de canal puedan
# mostrarlos y filtrar por ellos.
roi_por_canal = cargar_roi_por_canal()
rotacion = adjuntar_roi_por_canal(rotacion, roi_por_canal)
if not roi_por_canal.empty:
    print(
        f"Utilidad/ROI por canal cargados: {len(roi_por_canal)} pares "
        f"(sku_madre, canal) desde {ARCHIVO_ROI_B2B.name}"
    )

# Movimientos Odoo en revisión por SKU-canal (últimos 90 días). Se pegan a
# la rotación para que la tabla de antigüedad muestre si un SKU "joven"
# tuvo traslados sospechosos que pudieron reiniciar su antigüedad.
if not movimientos_auditoria.empty:
    movimientos_auditoria = movimientos_auditoria.copy()
    for _c in ["sku_madre", "canal_origen", "canal_destino", "canal_kam_afectado", "nivel_revision",
               "tipo_movimiento", "usuario", "motivo_revision"]:
        if _c not in movimientos_auditoria.columns:
            movimientos_auditoria[_c] = ""
        movimientos_auditoria[_c] = movimientos_auditoria[_c].fillna("").astype(str)
    movimientos_auditoria["fecha_movimiento"] = pd.to_datetime(movimientos_auditoria.get("fecha_movimiento"), errors="coerce")
    for _c in ["cantidad", "valor_movido", "edad_lote_al_movimiento"]:
        movimientos_auditoria[_c] = pd.to_numeric(movimientos_auditoria.get(_c), errors="coerce")
    movimientos_auditoria = movimientos_auditoria.sort_values("fecha_movimiento", ascending=False)
    _recientes = movimientos_auditoria[
        movimientos_auditoria["fecha_movimiento"] >= pd.Timestamp.now() - pd.Timedelta(days=90)
    ]
    _pares = []
    for _campo in ["canal_origen", "canal_destino"]:
        _t = _recientes[_recientes["nivel_revision"].isin(["ALTA", "MEDIA"])][["sku_madre", _campo, "nivel_revision", "move_line_id"]]
        _t = _t.rename(columns={_campo: "canal"})
        _pares.append(_t)
    _pares = pd.concat(_pares, ignore_index=True).drop_duplicates(["sku_madre", "canal", "move_line_id"])
    if not _pares.empty and not rotacion.empty:
        _cnt = _pares.groupby(["sku_madre", "canal"], as_index=False).agg(
            movs_revision_90d=("move_line_id", "nunique"),
            movs_alta_90d=("nivel_revision", lambda x: int((x == "ALTA").sum())),
        )
        rotacion = rotacion.merge(_cnt, on=["sku_madre", "canal"], how="left")
for _c in ["movs_revision_90d", "movs_alta_90d"]:
    if not rotacion.empty:
        rotacion[_c] = pd.to_numeric(rotacion[_c], errors="coerce").fillna(0) if _c in rotacion.columns else 0

# ------------------------------------------------------------------
# PUBLICACIONES AUTOAZUR vs STOCK ASIGNADO POR CANAL
# ------------------------------------------------------------------
# Un canal solo se evalúa si AutoAzur respondió bien y trajo publicaciones;
# si la consulta falló no se afirma que un SKU "no está publicado".
# El 01 deja un renglón general "TODOS" cuando no pudo consultar nada
# (p. ej. falta el GUID). Ese caso se muestra con su causa real en lugar de
# decir, canal por canal, que el canal no está configurado.
autoazur_error_general = ""
if not publicaciones_log.empty:
    if "canal" not in publicaciones_log.columns:
        publicaciones_log["canal"] = publicaciones_log.get("canal_autoazur", "")
    publicaciones_log["canal"] = publicaciones_log["canal"].fillna(publicaciones_log.get("canal_autoazur", "")).astype(str)
    _global = publicaciones_log[publicaciones_log["canal"].str.upper().eq("TODOS")
                                | publicaciones_log.get("canal_autoazur", pd.Series("", index=publicaciones_log.index)).astype(str).str.upper().eq("TODOS")]
    if not _global.empty:
        _g = _global.iloc[0]
        autoazur_error_general = f"{_g.get('estado_consulta', '')}: {_g.get('detalle', '')}".strip(": ")
    publicaciones_log = publicaciones_log.drop(index=_global.index)
autoazur_disponible = not publicaciones_log.empty
canales_autoazur_ok: set[str] = set()
canales_autoazur_sin_estado: set[str] = set()
if autoazur_disponible and {"canal", "estado_consulta"}.issubset(publicaciones_log.columns):
    # Solo canales con descarga confiable. Una descarga INCOMPLETA se acepta
    # si trae al menos el 95% de lo que informa AutoAzur; con menos, los SKU
    # faltantes aparecerían como "sin publicación" sin serlo.
    _pubs = pd.to_numeric(publicaciones_log.get("publicaciones", 0), errors="coerce").fillna(0)
    _tot = pd.to_numeric(publicaciones_log.get("total_informado", np.nan), errors="coerce")
    _cobertura = np.where(_tot > 0, _pubs / _tot, 1.0)
    publicaciones_log["cobertura_descarga"] = _cobertura
    _estado = publicaciones_log["estado_consulta"].astype(str)
    # SIN_CAMPO_ESTADO: el SKU sí se detecta, así que el canal se evalúa
    # para "Sin publicación"; lo que no se puede saber es si está pausada.
    _ok = publicaciones_log[
        (_pubs > 0)
        & (_estado.isin(["OK", "SIN_CAMPO_ESTADO"])
           | (_estado.eq("INCOMPLETO") & (_cobertura >= 0.95)))
    ]
    canales_autoazur_sin_estado = {
        normalizar_canal(c) for c in publicaciones_log.loc[_estado.eq("SIN_CAMPO_ESTADO"), "canal"].astype(str)
    }
    canales_autoazur_ok = {normalizar_canal(c) for c in _ok["canal"].astype(str)}

ALERTAS_PUBLICACION_ACCION = {
    "Sin publicación": "Publicar el SKU en {canal}: tiene piezas asignadas y ninguna publicación",
    "Pausada · sin stock": "Reactivar: la publicación está pausada por falta de stock aunque hay piezas; revisar el stock que manda AutoAzur",
    "Pausada": "Reactivar la publicación pausada o confirmar por qué se pausó",
    "En revisión": "Dar seguimiento a la revisión del marketplace",
    "Con error / rechazada": "Corregir el error del marketplace y volver a publicar",
    "Inactiva": "Reactivar la publicación o crear una nueva",
    "Cerrada / eliminada": "Crear una publicación nueva: la anterior se cerró",
    "Activa con stock 0": "La publicación está activa pero con stock 0: sincronizar stock en AutoAzur",
    "Publicada con stock 0": "AutoAzur publica 0 piezas aunque hay stock en bodega: revisar la sincronización de stock (el marketplace suele pausarla)",
    "Otro": "Revisar el estado de la publicación en AutoAzur",
    "Sin estado": "Revisar el estado de la publicación en AutoAzur",
}
listings_sin_stock = pd.DataFrame()
if not rotacion.empty:
    for _c in ["estado_publicacion", "estados_detalle", "item_ids", "cuentas", "skus_publicados", "subestados"]:
        if _c in rotacion.columns:
            rotacion = rotacion.drop(columns=[_c])
    if not publicaciones_sku_canal.empty:
        _pub = publicaciones_sku_canal.copy()
        _pub["sku_madre"] = _pub["sku_madre"].astype(str).str.strip().str.upper()
        _pub["canal"] = _pub["canal"].apply(normalizar_canal)
        _pub = _pub.drop_duplicates(["sku_madre", "canal"])
        rotacion["_sku_up"] = rotacion["sku_madre"].astype(str).str.strip().str.upper()
        rotacion = rotacion.merge(
            _pub.rename(columns={"sku_madre": "_sku_up"}), on=["_sku_up", "canal"], how="left"
        ).drop(columns=["_sku_up"])
        # Publicaciones activas en canales donde hoy no hay piezas asignadas
        # (riesgo de venta sin inventario para surtir).
        _con_stock = set(zip(
            rotacion.loc[rotacion["inventario_total"].fillna(0) > 0, "sku_madre"].astype(str).str.upper(),
            rotacion.loc[rotacion["inventario_total"].fillna(0) > 0, "canal"],
        ))
        _nombres = rotacion.drop_duplicates("sku_madre").set_index(rotacion.drop_duplicates("sku_madre")["sku_madre"].astype(str).str.upper())["producto_madre"].to_dict()
        listings_sin_stock = _pub[
            (pd.to_numeric(_pub["n_activas"], errors="coerce").fillna(0) > 0)
            & ~_pub.apply(lambda r: (r["sku_madre"], r["canal"]) in _con_stock, axis=1)
        ].copy()
        listings_sin_stock["producto_madre"] = listings_sin_stock["sku_madre"].map(_nombres).fillna("")
    for _c, _d in [("estado_publicacion", ""), ("estados_detalle", ""), ("item_ids", ""), ("cuentas", ""),
                   ("skus_publicados", ""), ("subestados", "")]:
        if _c not in rotacion.columns:
            rotacion[_c] = _d
        rotacion[_c] = rotacion[_c].fillna(_d).astype(str).replace("nan", "")
    rotacion["estados_detalle"] = rotacion["estados_detalle"].str.replace("Sin estado", "Estado no informado", regex=False)
    for _c in ["n_publicaciones", "n_activas", "stock_publicado_activo"]:
        rotacion[_c] = pd.to_numeric(rotacion[_c], errors="coerce").fillna(0) if _c in rotacion.columns else 0
    _detectado = rotacion["stock_publicado_detectado"].astype(str).str.lower().isin(["true", "1", "si"]) \
        if "stock_publicado_detectado" in rotacion.columns else pd.Series(False, index=rotacion.index)

    def _alerta_pub(r: pd.Series, detectado: bool) -> str:
        if not autoazur_disponible or safe_float(r.get("inventario_total"), 0) <= 0:
            return ""
        if r["canal"] not in canales_autoazur_ok:
            return "Sin datos AutoAzur"
        estado = str(r.get("estado_publicacion") or "")
        if not estado:
            return "Sin publicación"
        if r["canal"] in canales_autoazur_sin_estado and estado in ("Sin estado", "Otro"):
            # Sin estado: la señal disponible es el stock que publica AutoAzur.
            if (detectado and safe_float(r.get("stock_publicado_activo"), 0) <= 0
                    and safe_float(r.get("inventario_odoo"), 0) > 0
                    and safe_float(r.get("inventario_full"), 0) <= 0):
                return "Publicada con stock 0"
            return "OK"
        if estado == "Activa":
            if (detectado and safe_float(r.get("stock_publicado_activo"), 0) <= 0
                    and safe_float(r.get("inventario_odoo"), 0) > 0
                    and safe_float(r.get("inventario_full"), 0) <= 0):
                return "Activa con stock 0"
            return "OK"
        return estado

    rotacion["alerta_publicacion"] = [
        _alerta_pub(r, bool(d)) for (_, r), d in zip(rotacion.iterrows(), _detectado)
    ]
    rotacion["accion_publicacion"] = [
        ALERTAS_PUBLICACION_ACCION.get(a, "").format(canal=c) if a not in ("", "OK", "Sin datos AutoAzur") else ""
        for a, c in zip(rotacion["alerta_publicacion"], rotacion["canal"])
    ]
    if autoazur_disponible:
        _pend = rotacion["alerta_publicacion"].isin([k for k in ALERTAS_PUBLICACION_ACCION])
        print(f"Publicaciones: {int(_pend.sum())} SKU-canal con stock sin publicación activa.")

# Mapa de SKUs individuales (por canal/marketplace) asociados a cada SKU madre,
# para poder buscar por SKU individual además de SKU madre o producto.
sku_individual_map: dict[str, list[str]] = {}
if not sku_detalle.empty and {"sku_madre", "sku_sincronizado"}.issubset(sku_detalle.columns):
    sd = sku_detalle.copy()
    sd["sku_madre"] = sd["sku_madre"].astype(str).str.strip()
    sd["sku_sincronizado"] = sd["sku_sincronizado"].astype(str).str.strip()
    sd = sd[(sd["sku_madre"] != "") & (sd["sku_sincronizado"] != "")]
    sku_individual_map = (
        sd.groupby("sku_madre")["sku_sincronizado"]
        .apply(lambda s: sorted(set(s)))
        .to_dict()
    )

if not rotacion.empty:
    rotacion["skus_individuales"] = rotacion["sku_madre"].astype(str).map(
        lambda sku: " ".join(sku_individual_map.get(sku, []))
    )

# Normalización del detalle diario de ventas que se mostrará y exportará.
columnas_ventas_detalle = [
    "fecha", "sku_madre", "producto_madre", "canal", "sku_original",
    "modalidad_venta", "fuente", "unidades", "venta_total", "pedidos"
]
if ventas_detalle.empty:
    ventas_detalle = pd.DataFrame(columns=columnas_ventas_detalle)
else:
    ventas_detalle = ventas_detalle.copy()
    for col in columnas_ventas_detalle:
        if col not in ventas_detalle.columns:
            ventas_detalle[col] = 0 if col in ["unidades", "venta_total", "pedidos"] else ""
    ventas_detalle["fecha"] = pd.to_datetime(ventas_detalle["fecha"], errors="coerce")
    ventas_detalle["sku_madre"] = ventas_detalle["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    ventas_detalle["canal"] = ventas_detalle["canal"].apply(normalizar_canal)
    ventas_detalle["sku_original"] = ventas_detalle["sku_original"].fillna("").astype(str).str.strip()
    for col in ["unidades", "venta_total", "pedidos"]:
        ventas_detalle[col] = pd.to_numeric(ventas_detalle[col], errors="coerce").fillna(0)
    ventas_detalle = ventas_detalle[columnas_ventas_detalle].copy()

# Resumen general por IQ para la página principal: ventas y stock, sin usar
# cobertura como indicador principal.
if rotacion.empty:
    resumen_general_productos = pd.DataFrame(columns=[
        "sku_madre", "producto_madre", "skus_individuales", "canales",
        "inventario_odoo", "inventario_transito", "inventario_full", "inventario_total",
        "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades", "venta_90d_monto"
    ])
else:
    rg = rotacion.copy()
    for col in [
        "inventario_odoo", "inventario_transito", "inventario_full", "inventario_total",
        "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades"
    ]:
        if col not in rg.columns:
            rg[col] = 0
        rg[col] = pd.to_numeric(rg[col], errors="coerce").fillna(0)
    rg["producto_madre"] = rg.get("producto_madre", "").fillna("").astype(str)
    rg["skus_individuales"] = rg.get("skus_individuales", "").fillna("").astype(str)
    rg["canal"] = rg.get("canal", "").fillna("").astype(str)

    def unir_unicos(series):
        valores = []
        for valor in series:
            for parte in str(valor or "").replace("|", " ").split():
                parte = parte.strip()
                if parte and parte not in valores:
                    valores.append(parte)
        return " | ".join(valores)

    resumen_general_productos = rg.groupby("sku_madre", as_index=False).agg(
        producto_madre=("producto_madre", lambda s: next((x for x in s if str(x).strip()), "")),
        skus_individuales=("skus_individuales", unir_unicos),
        canales=("canal", lambda s: " | ".join(sorted({str(x).strip() for x in s if str(x).strip() and str(x).strip() != "General"}))),
        inventario_odoo=("inventario_odoo", "sum"),
        inventario_transito=("inventario_transito", "sum"),
        inventario_full=("inventario_full", "sum"),
        inventario_total=("inventario_total", "sum"),
        ventas_10d_unidades=("ventas_10d_unidades", "sum"),
        ventas_30d_unidades=("ventas_30d_unidades", "sum"),
        ventas_90d_unidades=("ventas_90d_unidades", "sum"),
    )
    if not ventas_detalle.empty:
        monto_sku = ventas_detalle.groupby("sku_madre", as_index=False)["venta_total"].sum().rename(columns={"venta_total": "venta_90d_monto"})
        resumen_general_productos = resumen_general_productos.merge(monto_sku, on="sku_madre", how="left")
    else:
        resumen_general_productos["venta_90d_monto"] = 0
    resumen_general_productos["venta_90d_monto"] = pd.to_numeric(
        resumen_general_productos["venta_90d_monto"], errors="coerce"
    ).fillna(0)
    resumen_general_productos = resumen_general_productos.sort_values(
        ["ventas_90d_unidades", "inventario_total"], ascending=[False, False]
    )

# Serie diaria de ventas por canal (unidades, monto, utilidad y ROI).
serie_ventas_canal = construir_serie_ventas_canal(ventas_detalle)
if serie_ventas_canal.empty:
    rango_ventas_canal = {"min": "", "max": ""}
else:
    rango_ventas_canal = {
        "min": str(serie_ventas_canal["fecha"].min()),
        "max": str(serie_ventas_canal["fecha"].max()),
    }
    print(
        f"Serie diaria por canal: {len(serie_ventas_canal):,} renglones "
        f"({rango_ventas_canal['min']} → {rango_ventas_canal['max']})"
    )
serie_ventas_tiene_utilidad = bool(
    not serie_ventas_canal.empty
    and safe_float(serie_ventas_canal["base_roi"].abs().sum(), 0) > 0
)
if not serie_ventas_tiene_utilidad:
    print(
        "AVISO: la serie diaria no trae utilidad/ROI. La gráfica del canal "
        "funcionará con unidades y monto; para habilitar utilidad y ROI, "
        "corre build_utilidad_roi.py con la hoja 'ventas_diarias_canal'."
    )

# Portafolio B2B financiero · excedente + costo de oportunidad VPN.
roi_b2b_base = cargar_roi_b2b()
stock_b2b_live, b2b_live_meta = cargar_stock_b2b_odoo_vivo(sku_detalle, sku_aliases_b2b, resumen_general_productos)
print(
    f"B2B stock: {b2b_live_meta.get('status')} · "
    f"fuente={b2b_live_meta.get('source')} · "
    f"SKU mapeados={b2b_live_meta.get('mapped_skus', 0):,}"
)
b2b_portafolio = construir_portafolio_b2b(
    rotacion,
    productos,
    roi_b2b_base,
    roi_por_canal,
    stock_b2b_live,
    b2b_live_meta,
)

# Normalización mínima de producto principal.
if not productos.empty:
    for c in [
        "stock_total",
        "ventas_3m_unidades",
        "ventas_3m_monto",
        "dias_inventario_num",
        "dias_desde_ultimo_arribo",
        "sugerencia_compra",
    ]:
        if c in productos.columns:
            productos[c] = pd.to_numeric(productos[c], errors="coerce").fillna(0)

# Resumen por canal.
if not rotacion.empty:
    canal_resumen = rotacion.groupby("canal", as_index=False).agg(
        inventario_odoo=("inventario_odoo", "sum"),
        inventario_transito=("inventario_transito", "sum"),
        inventario_full=("inventario_full", "sum"),
        inventario_total=("inventario_total", "sum"),
        ventas_90d_unidades=("ventas_90d_unidades", "sum"),
        ventas_full_90d=("ventas_full_90d", "sum"),
        ventas_drop_90d=("ventas_drop_90d", "sum"),
        transferencias_sugeridas=("transferencia_sugerida", "sum"),
        compras_sugeridas=("compra_sugerida", "sum"),
        skus=("sku_madre", "nunique"),
        valor_inventario=("valor_inventario_canal", "sum"),
        valor_mas_90d=("valor_mas_90d", "sum"),
        unidades_mas_90d=("unidades_mas_90d", "sum"),
    )
    _skus90 = rotacion[rotacion["unidades_mas_90d"].fillna(0) > 0].groupby("canal")["sku_madre"].nunique()
    canal_resumen["skus_mas_90d"] = canal_resumen["canal"].map(_skus90).fillna(0)
    canal_resumen["venta_diaria"] = canal_resumen["ventas_90d_unidades"] / 90
    canal_resumen["cobertura_total_dias"] = np.where(
        canal_resumen["venta_diaria"] > 0,
        canal_resumen["inventario_total"] / canal_resumen["venta_diaria"],
        np.nan,
    )
    canal_resumen["venta_diaria_full"] = canal_resumen["ventas_full_90d"] / 90
    canal_resumen["cobertura_full_dias"] = np.where(
        canal_resumen["venta_diaria_full"] > 0,
        canal_resumen["inventario_full"] / canal_resumen["venta_diaria_full"],
        np.nan,
    )
    canal_resumen["kam"] = canal_resumen["canal"].map(KAM_POR_CANAL).fillna("Por asignar")
    canal_resumen["orden"] = canal_resumen["canal"].map(
        {c: i for i, c in enumerate(CANALES_ORDEN + ["General"])}
    ).fillna(99)
    canal_resumen = canal_resumen.sort_values("orden").drop(columns=["orden"])
else:
    canal_resumen = pd.DataFrame()

transferencias = rotacion[
    rotacion["accion_preliminar"].astype(str).str.contains("TRANSFERIR", case=False, na=False)
    | (rotacion["transferencia_sugerida"] > 0)
].copy()
transferencias = transferencias.sort_values(
    ["transferencia_sugerida", "ventas_90d_unidades"], ascending=[False, False]
)

compras = rotacion[
    rotacion["accion_preliminar"].astype(str).str.contains("COMPRA", case=False, na=False)
    | (rotacion["compra_sugerida"] > 0)
].copy()
compras = compras.sort_values(["compra_sugerida", "ventas_90d_unidades"], ascending=[False, False])

# Compra consolidada por SKU (visión empresa, no por canal).
# Evita sumar déficits por canal cuando el mismo SKU tiene excedente en otra bolsa.
# Fórmula preserva la política actual de 45 días:
#   compra SKU = max(demanda diaria total SKU * cobertura objetivo - inventario total empresa SKU, 0)
def _primer_texto_no_vacio(series):
    for value in series:
        s = str(value or "").strip()
        if s and s.lower() not in {"nan", "none", "false"}:
            return s
    return ""

if not rotacion.empty:
    _rot_compra = rotacion.copy()
    for _c in [
        "inventario_odoo", "inventario_transito", "inventario_full", "inventario_total",
        "ventas_90d_unidades", "venta_diaria_calendario",
    ]:
        if _c not in _rot_compra.columns:
            _rot_compra[_c] = 0
        _rot_compra[_c] = pd.to_numeric(_rot_compra[_c], errors="coerce").fillna(0)

    _rot_compra["canal_compra_activo"] = np.where(
        (_rot_compra["ventas_90d_unidades"] > 0) | (_rot_compra["inventario_total"] > 0),
        _rot_compra["canal"].astype(str),
        "",
    )
    compras_consolidadas = (
        _rot_compra.groupby("sku_madre", as_index=False)
        .agg(
            producto_madre=("producto_madre", _primer_texto_no_vacio),
            canales=("canal_compra_activo", lambda s: " | ".join(sorted({str(x).strip() for x in s if str(x).strip() and str(x).strip() != "General"}))),
            inventario_odoo=("inventario_odoo", "sum"),
            inventario_transito=("inventario_transito", "sum"),
            inventario_full=("inventario_full", "sum"),
            inventario_total=("inventario_total", "sum"),
            ventas_90d_unidades=("ventas_90d_unidades", "sum"),
            venta_diaria_total=("venta_diaria_calendario", "sum"),
        )
    )
    # Utilidad/ROI reales de 30 y 90 días (base_utilidad_roi.xlsx), para
    # priorizar qué SKU conviene comprar primero, no solo por rotación.
    compras_consolidadas = adjuntar_financiero_roi(compras_consolidadas, roi_b2b_base)
    compras_consolidadas["objetivo_stock_piezas"] = np.ceil(
        compras_consolidadas["venta_diaria_total"] * COBERTURA_COMPRA_DIAS
    )
    compras_consolidadas["compra_sugerida"] = np.ceil(
        (compras_consolidadas["objetivo_stock_piezas"] - compras_consolidadas["inventario_total"]).clip(lower=0)
    )
    compras_consolidadas["cobertura_total_dias"] = np.where(
        compras_consolidadas["venta_diaria_total"] > 0,
        compras_consolidadas["inventario_total"] / compras_consolidadas["venta_diaria_total"],
        np.nan,
    )
    compras_consolidadas["nivel_riesgo"] = np.select(
        [
            compras_consolidadas["compra_sugerida"] <= 0,
            compras_consolidadas["cobertura_total_dias"].fillna(np.inf) < 15,
            compras_consolidadas["cobertura_total_dias"].fillna(np.inf) < 30,
        ],
        ["SANO", "CRÍTICO", "ALTO"],
        default="REVISAR",
    )
    compras_consolidadas["justificacion"] = compras_consolidadas.apply(
        lambda r: (
            f"Inventario empresa={r['inventario_total']:.0f}; ventas 90d={r['ventas_90d_unidades']:.0f}; "
            f"ritmo={r['venta_diaria_total']:.2f} u/día; objetivo {COBERTURA_COMPRA_DIAS}d="
            f"{r['objetivo_stock_piezas']:.0f}; compra consolidada={r['compra_sugerida']:.0f}."
        ),
        axis=1,
    )
    compras_consolidadas = compras_consolidadas[
        compras_consolidadas["compra_sugerida"] > 0
    ].sort_values(["compra_sugerida", "ventas_90d_unidades"], ascending=[False, False])
else:
    compras_consolidadas = pd.DataFrame()

# Buckets de antigüedad.
# Versión 2026-09: si el 01 exportó las capas FIFO, los buckets se arman con
# las piezas reales de cada capa (y su valor). Si no, se conserva el cálculo
# anterior por "días desde último arribo".
age_buckets: list[dict[str, Any]] = []
age_source = "capas"
if not antiguedad_capas.empty and {"dias", "unidades"}.issubset(antiguedad_capas.columns):
    age = antiguedad_capas.copy()
    age["dias"] = pd.to_numeric(age["dias"], errors="coerce")
    age["unidades"] = pd.to_numeric(age["unidades"], errors="coerce").fillna(0)
    age["valor"] = pd.to_numeric(age.get("valor", 0), errors="coerce").fillna(0)
    age = age[age["unidades"] > 0].copy()
    labels = ["0–30", "31–60", "61–90", "91–180", "+180", "Sin fecha"]
    age["bucket"] = pd.cut(age["dias"], bins=[-np.inf, 30, 60, 90, 180, np.inf], labels=labels[:-1])
    age["bucket"] = age["bucket"].astype(object).where(age["dias"].notna(), "Sin fecha")
    grouped = age.groupby("bucket", dropna=False).agg(
        unidades=("unidades", "sum"), valor=("valor", "sum"), skus=("sku_madre", "nunique")
    ).reindex(labels).fillna(0).reset_index().rename(columns={"index": "bucket"})
    age_buckets = js_records(grouped)
elif not productos.empty and "dias_desde_ultimo_arribo" in productos.columns:
    age_source = "ultimo_arribo"
    age = productos.copy()
    age["dias_desde_ultimo_arribo"] = pd.to_numeric(age["dias_desde_ultimo_arribo"], errors="coerce")
    if "stock_total" not in age.columns:
        age["stock_total"] = 0
    age["stock_total"] = pd.to_numeric(age["stock_total"], errors="coerce").fillna(0)
    age = age[age["stock_total"] > 0].copy()
    labels = ["0–30", "31–60", "61–90", "91–180", "+180", "Sin fecha"]
    bins = [-np.inf, 30, 60, 90, 180, np.inf]
    age["bucket"] = pd.cut(age["dias_desde_ultimo_arribo"], bins=bins, labels=labels[:-1])
    age["bucket"] = age["bucket"].astype(object).where(age["dias_desde_ultimo_arribo"].notna(), "Sin fecha")
    grouped = age.groupby("bucket", dropna=False, observed=False).agg(
        unidades=("stock_total", "sum"), skus=("sku_madre", "nunique")
    ).reset_index()
    grouped["bucket"] = pd.Categorical(grouped["bucket"], labels, ordered=True)
    grouped = grouped.sort_values("bucket")
    age_buckets = js_records(grouped)

# Calidad de datos.
quality_summary = {
    "stock_no_vinculado_skus": int(stock_no["sku_original"].nunique())
    if not stock_no.empty and "sku_original" in stock_no.columns
    else int(len(stock_no)),
    "stock_no_vinculado_unidades": safe_float(stock_no["stock_total"].sum())
    if not stock_no.empty and "stock_total" in stock_no.columns
    else 0,
    "ventas_no_vinculadas": int(len(ventas_no)),
    "traslados_excluidos": int(len(traslados_excluidos)),
    "diferencia_stock_total": safe_float(control_total["diferencia_stock"].sum())
    if not control_total.empty and "diferencia_stock" in control_total.columns
    else 0,
    "diferencia_ventas_total": safe_float(control_total["diferencia_ventas"].sum())
    if not control_total.empty and "diferencia_ventas" in control_total.columns
    else 0,
}

# KPIs globales.
stock_total = safe_float(rotacion["inventario_total"].sum()) if not rotacion.empty else 0
stock_odoo_total = safe_float(rotacion["inventario_odoo"].sum()) if not rotacion.empty else 0
stock_transito_total = safe_float(rotacion["inventario_transito"].sum()) if not rotacion.empty else 0
stock_full_total = safe_float(rotacion["inventario_full"].sum()) if not rotacion.empty else 0
ventas_10_total = safe_float(rotacion["ventas_10d_unidades"].sum()) if not rotacion.empty and "ventas_10d_unidades" in rotacion.columns else 0
ventas_30_total = safe_float(rotacion["ventas_30d_unidades"].sum()) if not rotacion.empty and "ventas_30d_unidades" in rotacion.columns else 0
ventas_90_total = safe_float(rotacion["ventas_90d_unidades"].sum()) if not rotacion.empty else 0
venta_diaria_total = ventas_90_total / 90 if ventas_90_total else 0
cobertura_total = stock_total / venta_diaria_total if venta_diaria_total else None
skus_criticos = int(rotacion["nivel_riesgo"].isin(["CRÍTICO", "ALTO"]).sum()) if not rotacion.empty else 0
skus_sobrestock = int((rotacion["nivel_riesgo"] == "SOBRESTOCK").sum()) if not rotacion.empty else 0

payload = {
    "meta": {
        "generated_at": FECHA_GENERACION,
        "source_file": ARCHIVO_BASE.name,
        "channels": CANALES_ORDEN,
        "kam_map": KAM_POR_CANAL,
        "sales_daily_range": rango_ventas_canal,
        "sales_daily_has_profit": serie_ventas_tiene_utilidad,
        "b2b_odoo_live": b2b_live_meta,
        "params": {
            "lead_time_full": LEAD_TIME_FULL_DIAS,
            "safety_full": SEGURIDAD_FULL_DIAS,
            "lead_time_supplier": LEAD_TIME_PROVEEDOR_DIAS,
            "purchase_coverage": COBERTURA_COMPRA_DIAS,
            "full_target_days": OBJETIVO_FULL_DIAS,
            "full_transit_days": DIAS_TRANSITO_ODOO_FULL,
            "new_product_full_max": MAX_PIEZAS_PRODUCTO_NUEVO_FULL,
            "full_channels": sorted(CANALES_CON_FULL),
            "internal_transfer_days": 5,
            "redistribution_target_days": 30,
            "demand_weights": {"d10": 0.50, "d30": 0.30, "d90": 0.20},
            "b2b_horizon_days": B2B_HORIZONTE_DIAS,
            "b2b_daily_discount_rate": B2B_TASA_DESCUENTO_DIARIA,
            "b2b_roi_window": B2B_ROI_VENTANA,
            "b2b_roi_minimum": B2B_ROI_MINIMO,
            "b2b_payout_fallback_days": B2B_PAYOUT_FALLBACK_DIAS,
            "b2b_odoo_live_stock": B2B_ODOO_LIVE_STOCK,
            "age_threshold": ANTIGUEDAD_UMBRAL_ALERTA_DIAS,
            "autoazur_available": autoazur_disponible,
            "autoazur_global_error": autoazur_error_general,
            "autoazur_channels_ok": sorted(canales_autoazur_ok),
            "audit_age_threshold": AUDITORIA_UMBRAL_EDAD_SOSPECHA,
            "audit_days": AUDITORIA_DIAS_EN_DASHBOARD,
            "coverage_break_days": COBERTURA_QUIEBRE_DIAS,
            "coverage_low_days": COBERTURA_BAJA_DIAS,
            "coverage_high_days": COBERTURA_ALTA_DIAS,
            "coverage_excess_days": COBERTURA_EXCESO_DIAS,
        },
    },
    "kpis": {
        "stock_total": stock_total,
        "stock_odoo": stock_odoo_total,
        "stock_transito": stock_transito_total,
        "stock_full": stock_full_total,
        "ventas_10": ventas_10_total,
        "ventas_30": ventas_30_total,
        "ventas_90": ventas_90_total,
        "cobertura_total": cobertura_total,
        "skus_criticos": skus_criticos,
        "skus_sobrestock": skus_sobrestock,
        "transferencias_unidades": safe_float(transferencias["transferencia_sugerida"].sum())
        if not transferencias.empty
        else 0,
        "compras_unidades": safe_float(compras_consolidadas["compra_sugerida"].sum()) if not compras_consolidadas.empty else 0,
        "valor_inventario": safe_float(rotacion["valor_inventario_canal"].sum()) if not rotacion.empty else 0,
        "valor_mas_90d": safe_float(rotacion["valor_mas_90d"].sum()) if not rotacion.empty else 0,
    },
    "channel_sales_daily": js_records(serie_ventas_canal, GRAFICA_VENTAS_MAX_FILAS),
    "channel_summary": js_records(canal_resumen),
    "sku_map": sku_individual_map,
    "product_summary": js_records(resumen_general_productos, 10000),
    "sales_detail": js_records(ventas_detalle, 30000),
    "rotation": js_records(rotacion),
    "transfers": js_records(transferencias),
    "redistributions": js_records(redistribuciones, 5000),
    "distribution_summary": js_records(reparticion_resumen, 10000),
    "distribution_detail": js_records(reparticion_detalle, 70000),
    "distribution_category_profiles": js_records(reparticion_perfiles_categoria, 10000),
    "b2b_portfolio": js_records(b2b_portafolio, 10000),
    "purchases": js_records(compras),
    "purchases_consolidated": js_records(compras_consolidadas),
    "products": js_records(productos),
    "product_alerts": js_records(alertas_productos),
    "arrivals": js_records(arribos, 6000),
    "age_buckets": age_buckets,
    "quality": quality_summary,
    "stock_unlinked": js_records(stock_no, 5000),
    "sales_unlinked": js_records(ventas_no, 5000),
    "excluded_transfers": js_records(traslados_excluidos, 5000),
    "control_totals": js_records(control_total, 5000),
    "age_source": age_source,
    "audit_moves": js_records(movimientos_auditoria, 20000),
    "audit_summary": js_records(movimientos_auditoria_resumen, 5000),
    "autoazur_log": js_records(publicaciones_log, 100),
    "listings_no_stock": js_records(listings_sin_stock, 10000),
    "listings_unlinked": js_records(
        publicaciones_sin_sku_resumen if not publicaciones_sin_sku_resumen.empty else (
            publicaciones_sin_sku[[c for c in ["canal", "sku_publicacion", "item_id", "titulo", "estado_publicacion",
                                               "cuenta", "stock_publicado"] if c in publicaciones_sin_sku.columns]]
            if not publicaciones_sin_sku.empty else publicaciones_sin_sku),
        5000,
    ),
    "listings_unlinked_by_channel": (
        {normalizar_canal(str(k)): int(v) for k, v in publicaciones_sin_sku["canal"].value_counts().items()}
        if not publicaciones_sin_sku.empty and "canal" in publicaciones_sin_sku.columns else {}
    ),
}

DATA_JSON = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

# Archivo auxiliar consumido por el backend HTTP/Odoo.
# El navegador NO decide por sí solo qué cantidad ejecutar: el servidor vuelve
# a validar esta fotografía y luego consulta el stock vivo de Odoo.
distribution_actions = []
for r in payload.get("distribution_detail", []) or []:
    sku = str(r.get("sku_madre", "") or "").strip().upper()
    dest = str(r.get("canal", "") or "").strip()
    for source_key, field_name, source_channel in [
        ("existencias", "desde_existencias", "General"),
        ("b2b", "desde_b2b", "B2B"),
    ]:
        qty = int(max(safe_float(r.get(field_name, 0), 0), 0))
        if sku and dest and qty > 0:
            distribution_actions.append({
                "sku_madre": sku,
                "producto_madre": str(r.get("producto_madre", "") or ""),
                "canal_origen": source_channel,
                "canal_destino": dest,
                "fuente": source_key,
                "cantidad_sugerida": qty,
                "metodo_reparticion": str(r.get("metodo_reparticion", "") or ""),
                "confianza": str(r.get("confianza", "") or ""),
            })

acciones_admin_payload = {
    "generated_at": FECHA_GENERACION,
    "full_channels": sorted(CANALES_CON_FULL),
    "redistributions": [
        {
            "sku_madre": str(r.get("sku_madre", "") or "").strip().upper(),
            "producto_madre": str(r.get("producto_madre", "") or ""),
            "canal_origen": str(r.get("canal_origen", "") or ""),
            "canal_destino": str(r.get("canal_destino", "") or ""),
            "tipo_stock_origen": str(r.get("tipo_stock_origen", "") or ""),
            "cantidad_sugerida": safe_float(r.get("cantidad_sugerida", 0), 0),
            "check_manual": str(r.get("check_manual", "") or ""),
            "confianza": str(r.get("confianza", "") or ""),
            "veredicto": str(r.get("veredicto", "") or ""),
        }
        for r in payload.get("redistributions", []) or []
        if safe_float(r.get("cantidad_sugerida", 0), 0) > 0
    ],
    "distribution": distribution_actions,
    "full_deliveries": [
        {
            "sku_madre": str(r.get("sku_madre", "") or "").strip().upper(),
            "producto_madre": str(r.get("producto_madre", "") or ""),
            "canal": str(r.get("canal", "") or ""),
            "cantidad_sugerida": safe_float(r.get("transferencia_sugerida", 0), 0),
            "check_manual_full": str(r.get("check_manual_full", "") or ""),
        }
        for r in payload.get("transfers", []) or []
        if safe_float(r.get("transferencia_sugerida", 0), 0) > 0
    ],
}
ARCHIVO_ACCIONES_ADMIN.write_text(
    json.dumps(acciones_admin_payload, ensure_ascii=False, indent=2),
    encoding="utf-8",
)


# ============================================================
# PLANTILLA HTML
# ============================================================

HTML_TEMPLATE = r'''<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="theme-color" content="#0b1220" />
  <title>Control Tower · Rotación de Inventario</title>
  <script src="https://cdn.plot.ly/plotly-2.35.2.min.js"></script>
  <style>
    :root {
      --nav:#0b1220;
      --nav-2:#111c31;
      --page:#f4f7fb;
      --card:#ffffff;
      --text:#172033;
      --muted:#697386;
      --line:#e4e9f1;
      --blue:#2563eb;
      --teal:#0f9f9a;
      --purple:#7559d9;
      --amber:#f59e0b;
      --red:#e5484d;
      --green:#1f9d66;
      --slate:#64748b;
      --shadow:0 12px 34px rgba(15,23,42,.08);
      --radius:18px;
      --sidebar:286px;
      --sidebar-collapsed:88px;
    }
    *{box-sizing:border-box}
    html,body{margin:0;min-height:100%;font-family:Inter,ui-sans-serif,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;background:var(--page);color:var(--text)}
    button,input,select{font:inherit}
    body.dark{
      --page:#0b1120;--card:#111827;--text:#edf2f7;--muted:#9aa6b6;--line:#253147;--shadow:0 14px 36px rgba(0,0,0,.3)
    }
    .app{display:flex;min-height:100vh}
    .sidebar{position:fixed;inset:0 auto 0 0;width:var(--sidebar);background:linear-gradient(180deg,var(--nav),#0c1729 62%,#101c30);color:white;z-index:50;display:flex;flex-direction:column;transition:width .25s ease;box-shadow:10px 0 30px rgba(4,10,20,.14)}
    .sidebar.collapsed{width:var(--sidebar-collapsed)}
    .brand{display:flex;align-items:center;gap:13px;padding:22px 18px;border-bottom:1px solid rgba(255,255,255,.08);min-height:82px}
    .brand-mark{width:44px;height:44px;border-radius:14px;background:linear-gradient(145deg,#2e6df6,#12a6a0);display:grid;place-items:center;font-weight:900;box-shadow:0 9px 24px rgba(37,99,235,.32)}
    .brand-copy{min-width:0;transition:opacity .2s}.brand-title{font-size:15px;font-weight:850;letter-spacing:.01em}.brand-sub{font-size:11px;color:#9fb0ca;margin-top:3px}
    .sidebar.collapsed .brand-copy,.sidebar.collapsed .nav-label,.sidebar.collapsed .nav-group-title,.sidebar.collapsed .sidebar-footer-copy,.sidebar.collapsed .nav-chevron{display:none}
    .collapse-btn{margin-left:auto;border:0;background:rgba(255,255,255,.08);color:#d8e1ee;width:34px;height:34px;border-radius:11px;cursor:pointer}
    .sidebar.collapsed .collapse-btn{position:absolute;right:-15px;top:23px;background:#182944;border:1px solid #2b4266}
    .nav{overflow:auto;padding:14px 12px 20px;flex:1}
    .nav-group{margin:7px 0 15px}.nav-group-title{font-size:10px;letter-spacing:.13em;text-transform:uppercase;color:#7f91ac;padding:10px 12px 7px;font-weight:800}
    .nav-item{width:100%;border:0;background:transparent;color:#b9c5d8;display:flex;align-items:center;gap:13px;padding:12px 13px;border-radius:12px;cursor:pointer;text-align:left;margin:2px 0;transition:.18s;position:relative}
    .nav-item:hover{background:rgba(255,255,255,.07);color:white}.nav-item.active{background:linear-gradient(90deg,rgba(37,99,235,.25),rgba(15,159,154,.13));color:white;box-shadow:inset 3px 0 0 #53a5ff}
    .nav-icon{width:24px;height:24px;display:grid;place-items:center;font-size:16px;flex:none}.nav-label{font-size:13px;font-weight:700;white-space:nowrap}.nav-badge{margin-left:auto;background:rgba(255,255,255,.1);color:#d9e4f3;border-radius:999px;padding:3px 7px;font-size:10px;font-weight:800}.sidebar.collapsed .nav-badge{position:absolute;top:5px;right:5px;padding:2px 5px}
    .sidebar-footer{padding:15px;border-top:1px solid rgba(255,255,255,.08);display:flex;gap:11px;align-items:center}.status-dot{width:9px;height:9px;border-radius:50%;background:#36d399;box-shadow:0 0 0 5px rgba(54,211,153,.12)}.sidebar-footer-copy{font-size:11px;color:#9fb0ca;line-height:1.4}
    .main{margin-left:var(--sidebar);width:calc(100% - var(--sidebar));min-height:100vh;transition:margin-left .25s,width .25s}.sidebar.collapsed~.main{margin-left:var(--sidebar-collapsed);width:calc(100% - var(--sidebar-collapsed))}
    .topbar{position:sticky;top:0;z-index:35;background:rgba(244,247,251,.88);backdrop-filter:blur(16px);border-bottom:1px solid rgba(217,224,234,.85);display:flex;align-items:center;justify-content:space-between;gap:16px;padding:14px 28px}.dark .topbar{background:rgba(11,17,32,.88);border-color:#263147}
    .topbar-left{display:flex;align-items:center;gap:14px;min-width:0}.mobile-menu{display:none;border:0;background:var(--card);color:var(--text);border-radius:10px;width:40px;height:40px;box-shadow:var(--shadow)}
    .page-eyebrow{font-size:11px;text-transform:uppercase;letter-spacing:.13em;color:var(--blue);font-weight:850}.page-title{font-size:21px;font-weight:900;letter-spacing:-.025em;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
    .topbar-right{display:flex;align-items:center;gap:10px}.search-global{position:relative}.search-global input{width:260px;border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:13px;padding:11px 13px 11px 38px;outline:none}.search-global:before{content:'⌕';position:absolute;left:13px;top:8px;font-size:20px;color:var(--muted)}
    .icon-btn{width:42px;height:42px;border:1px solid var(--line);border-radius:13px;background:var(--card);color:var(--text);cursor:pointer;box-shadow:0 4px 14px rgba(15,23,42,.04)}
    .content{padding:26px 28px 46px;max-width:1920px;margin:0 auto}.page{display:none}.page.active{display:block;animation:fade .22s ease}@keyframes fade{from{opacity:.2;transform:translateY(5px)}to{opacity:1;transform:none}}
    .semaphore-strip{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:16px}
    .semaphore-card{border-radius:20px;padding:20px 22px;display:flex;align-items:center;gap:16px;cursor:pointer;border:1px solid var(--line);background:var(--card);box-shadow:var(--shadow);transition:.15s;position:relative;overflow:hidden}
    .semaphore-card:hover{transform:translateY(-2px);box-shadow:0 18px 40px rgba(15,23,42,.12)}
    .semaphore-dot{width:16px;height:16px;border-radius:50%;flex:none;box-shadow:0 0 0 6px currentColor22}
    .semaphore-num{font-size:34px;font-weight:900;letter-spacing:-.04em;line-height:1}
    .semaphore-label{font-size:12px;color:var(--muted);font-weight:750;margin-top:4px}
    .semaphore-card.red{border-color:color-mix(in srgb,var(--red) 35%,var(--line))}.semaphore-card.red .semaphore-dot{background:var(--red);color:var(--red)}.semaphore-card.red .semaphore-num{color:var(--red)}
    .semaphore-card.amber{border-color:color-mix(in srgb,var(--amber) 35%,var(--line))}.semaphore-card.amber .semaphore-dot{background:var(--amber);color:var(--amber)}.semaphore-card.amber .semaphore-num{color:#b86f00}
    .semaphore-card.green{border-color:color-mix(in srgb,var(--green) 35%,var(--line))}.semaphore-card.green .semaphore-dot{background:var(--green);color:var(--green)}.semaphore-card.green .semaphore-num{color:var(--green)}
    .action-banner{border-radius:18px;background:linear-gradient(95deg,#0d6f76,#17335a);color:#fff;padding:16px 22px;display:flex;align-items:center;gap:14px;margin-bottom:16px;font-size:13px;line-height:1.5;box-shadow:var(--shadow)}
    .action-banner b{font-weight:900}.action-banner-icon{width:36px;height:36px;border-radius:11px;background:rgba(255,255,255,.14);display:grid;place-items:center;flex:none;font-size:17px}
    @media(max-width:1050px){.semaphore-strip{grid-template-columns:1fr}}
    .hero{border-radius:24px;background:linear-gradient(115deg,#101e35 0%,#17335a 52%,#0d6f76 120%);color:white;padding:28px 30px;box-shadow:0 20px 45px rgba(16,30,53,.2);display:flex;align-items:flex-end;justify-content:space-between;gap:20px;margin-bottom:22px;overflow:hidden;position:relative}.hero:after{content:'';position:absolute;width:380px;height:380px;border-radius:50%;right:-110px;top:-225px;background:radial-gradient(circle,rgba(91,169,255,.33),transparent 64%)}
    .hero-copy{position:relative;z-index:2}.hero-kicker{font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:#8fd6ff;font-weight:850}.hero h1{margin:8px 0 8px;font-size:31px;line-height:1.06;letter-spacing:-.04em}.hero p{margin:0;color:#c9d7e8;max-width:780px;font-size:14px;line-height:1.55}.hero-meta{position:relative;z-index:2;display:flex;gap:10px;flex-wrap:wrap;justify-content:flex-end}.hero-chip{padding:9px 12px;border-radius:999px;background:rgba(255,255,255,.11);border:1px solid rgba(255,255,255,.12);font-size:11px;font-weight:750;color:#e7f2ff;white-space:nowrap}
    .kpi-grid{display:grid;grid-template-columns:repeat(5,minmax(150px,1fr));gap:15px;margin-bottom:22px}.kpi{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);padding:18px 19px;box-shadow:var(--shadow);position:relative;overflow:hidden;min-height:132px}.kpi-accent{position:absolute;inset:0 auto 0 0;width:4px;background:var(--accent,var(--blue))}.kpi-top{display:flex;justify-content:space-between;gap:10px}.kpi-label{font-size:12px;color:var(--muted);font-weight:750}.kpi-icon{width:32px;height:32px;border-radius:10px;display:grid;place-items:center;background:color-mix(in srgb,var(--accent,var(--blue)) 12%,transparent);color:var(--accent,var(--blue));font-size:16px}.kpi-value{font-size:29px;font-weight:900;letter-spacing:-.04em;margin-top:12px}.kpi-sub{font-size:11px;color:var(--muted);margin-top:7px;line-height:1.4}.kpi-progress{height:5px;border-radius:999px;background:var(--line);overflow:hidden;margin-top:10px}.kpi-progress span{display:block;height:100%;background:var(--accent,var(--blue));border-radius:999px}
    .kpi.clickable{cursor:pointer;transition:transform .18s ease,box-shadow .18s ease,border-color .18s ease}.kpi.clickable:hover{transform:translateY(-3px);box-shadow:0 18px 38px rgba(15,23,42,.14);border-color:color-mix(in srgb,var(--accent,var(--purple)) 38%,var(--line))}.kpi.clickable:focus{outline:3px solid color-mix(in srgb,var(--accent,var(--purple)) 25%,transparent);outline-offset:2px}.kpi-link{margin-top:10px;font-size:11px;font-weight:850;color:var(--accent,var(--purple));display:flex;align-items:center;gap:5px}
    .odoo-detail-card{display:none;margin:16px 0 22px;border:1px solid color-mix(in srgb,var(--purple) 30%,var(--line));background:linear-gradient(145deg,color-mix(in srgb,var(--purple) 6%,var(--card)),var(--card));scroll-margin-top:90px}.odoo-detail-card.open{display:block;animation:panelIn .24s ease}.odoo-detail-card.flash{box-shadow:0 0 0 4px rgba(117,89,217,.16),var(--shadow)}@keyframes panelIn{from{opacity:0;transform:translateY(-8px)}to{opacity:1;transform:translateY(0)}}
    .odoo-rule{display:flex;gap:12px;align-items:flex-start;padding:13px 14px;border:1px solid color-mix(in srgb,var(--purple) 20%,var(--line));background:color-mix(in srgb,var(--purple) 5%,var(--card));border-radius:13px;margin:12px 0 15px;color:var(--muted);font-size:12px;line-height:1.55}.odoo-rule strong{color:var(--text)}.odoo-rule-icon{width:30px;height:30px;border-radius:10px;background:rgba(117,89,217,.12);color:var(--purple);display:grid;place-items:center;font-weight:900;flex:0 0 auto}
    .odoo-panel-tools{display:flex;gap:9px;flex-wrap:wrap;align-items:center}.odoo-panel-tools .filter-input{min-width:260px}.odoo-close{border:1px solid var(--line);background:var(--card);color:var(--muted);width:34px;height:34px;border-radius:10px;cursor:pointer;font-weight:900}.odoo-close:hover{color:var(--text);border-color:var(--purple)}
    .executive-products-card{margin-bottom:18px}.sales-detail-card{display:none;margin-top:18px;scroll-margin-top:92px;border-color:color-mix(in srgb,var(--blue) 28%,var(--line));background:linear-gradient(145deg,color-mix(in srgb,var(--blue) 4%,var(--card)),var(--card))}.sales-detail-card.open{display:block;animation:panelIn .24s ease}.sales-detail-grid{align-items:start;margin-bottom:0}.mini-section-title{font-size:12px;font-weight:850;color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin:2px 0 10px}.sales-open-btn{border:0;background:linear-gradient(135deg,var(--blue),#4182f7);color:#fff;border-radius:9px;padding:7px 10px;font-size:11px;font-weight:800;cursor:pointer;white-space:nowrap}.sales-open-btn:hover{filter:brightness(1.06);transform:translateY(-1px)}
    .kpi-help{position:relative;display:inline-flex;width:16px;height:16px;border-radius:50%;background:var(--line);color:var(--muted);font-size:10px;font-weight:800;align-items:center;justify-content:center;cursor:help;flex:none}
    .kpi-help .kpi-tooltip{display:none;position:absolute;bottom:22px;left:50%;transform:translateX(-50%);width:200px;background:#101828;color:#fff;font-size:11px;font-weight:500;line-height:1.5;padding:10px 11px;border-radius:10px;box-shadow:0 12px 30px rgba(0,0,0,.3);z-index:20;text-align:left}
    .kpi-help:hover .kpi-tooltip{display:block}
    .kpi-top{display:flex;justify-content:space-between;gap:8px;align-items:flex-start}
    .grid-2{display:grid;grid-template-columns:1fr 1fr;gap:18px;margin-bottom:18px}.grid-3{display:grid;grid-template-columns:1.25fr .9fr .9fr;gap:18px;margin-bottom:18px}.grid-73{display:grid;grid-template-columns:1.7fr 1fr;gap:18px;margin-bottom:18px}.card{background:var(--card);border:1px solid var(--line);border-radius:var(--radius);box-shadow:var(--shadow);padding:20px 21px;min-width:0}.card-head{display:flex;align-items:flex-start;justify-content:space-between;gap:14px;margin-bottom:14px}.card-title{font-size:16px;font-weight:880;letter-spacing:-.02em}.card-sub{font-size:11px;color:var(--muted);margin-top:4px;line-height:1.45}.card-action{border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:10px;padding:7px 10px;font-size:11px;font-weight:750;cursor:pointer}.chart{height:350px}.chart.small{height:290px}.chart.tall{height:420px}
    .view-toggle{display:inline-flex;border:1px solid var(--line);border-radius:11px;overflow:hidden}.view-toggle button{border:0;background:var(--card);color:var(--muted);padding:9px 13px;font-size:12px;font-weight:750;cursor:pointer}.view-toggle button.active{background:var(--blue);color:#fff}
    .section-heading{display:flex;align-items:flex-end;justify-content:space-between;gap:14px;margin:2px 0 16px}.section-heading h2{margin:0;font-size:25px;letter-spacing:-.035em}.section-heading p{margin:5px 0 0;color:var(--muted);font-size:13px}.section-tools{display:flex;gap:9px;flex-wrap:wrap}.date-range{display:none;align-items:center;gap:6px}.date-range.open{display:inline-flex}.date-sep{font-size:11px;color:var(--muted);font-weight:750}.date-range input[type=date]{font-size:11px;padding:8px 9px}.select,.filter-input{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:9px 11px;font-size:12px;outline:none}.filter-input{min-width:230px}
    .status-strip{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}.status-card{padding:15px 16px;background:var(--card);border:1px solid var(--line);border-radius:15px;display:flex;align-items:center;gap:12px}.status-icon{width:38px;height:38px;border-radius:12px;display:grid;place-items:center;font-weight:900}.status-info b{display:block;font-size:19px}.status-info span{font-size:11px;color:var(--muted)}
    .insights{display:grid;gap:10px}.insight{border:1px solid var(--line);border-left:4px solid var(--tone,var(--blue));border-radius:13px;padding:13px 14px;background:color-mix(in srgb,var(--tone,var(--blue)) 4%,var(--card));display:flex;gap:11px;align-items:flex-start}.insight-icon{width:28px;height:28px;border-radius:9px;background:color-mix(in srgb,var(--tone,var(--blue)) 14%,transparent);color:var(--tone,var(--blue));display:grid;place-items:center;flex:none}.insight-title{font-size:12px;font-weight:850}.insight-text{font-size:11px;color:var(--muted);line-height:1.5;margin-top:3px}
    .table-inline-tools{display:flex;align-items:center;justify-content:space-between;gap:10px;margin:0 0 10px;flex-wrap:wrap}.table-inline-tools-left,.table-inline-tools-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.table-inline-search{width:min(330px,100%);border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:9px 11px;font-size:12px;outline:none}.table-inline-search:focus,.table-inline-select:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.10)}.table-inline-select{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:9px 10px;font-size:11px;outline:none;max-width:220px}.table-sort-dir{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:9px 11px;font-size:11px;font-weight:800;cursor:pointer;white-space:nowrap}.table-sort-dir:hover{border-color:var(--blue);color:var(--blue)}.table-inline-count{font-size:11px;color:var(--muted);font-weight:750;white-space:nowrap}.table-sort-note{font-size:10px;color:var(--muted);font-weight:700}.table-filter-chip{font-size:10px;color:var(--blue);background:rgba(37,99,235,.08);border-radius:999px;padding:5px 8px;font-weight:800}.table-wrap{border:1px solid var(--line);border-radius:14px;overflow:auto;max-height:620px}.data-table{width:100%;border-collapse:separate;border-spacing:0;min-width:920px;font-size:11px}.data-table th{position:sticky;top:0;z-index:2;background:#f8fafc;color:#5d687a;text-transform:uppercase;letter-spacing:.055em;font-size:9px;text-align:left;padding:11px 12px;border-bottom:1px solid var(--line);white-space:nowrap}.dark .data-table th{background:#172033;color:#aab6c7}.data-table td{padding:11px 12px;border-bottom:1px solid var(--line);vertical-align:middle}.data-table tr:hover td{background:rgba(37,99,235,.035)}.data-table td.num{text-align:right;font-variant-numeric:tabular-nums}.data-table td.neg{color:var(--red);font-weight:750}.data-table td.muted-cell{color:var(--muted)}.table-inline-select.active-filter{border-color:var(--blue);color:var(--blue);font-weight:800}.sku{font-weight:850;color:var(--blue)}.product{max-width:260px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.badge{display:inline-flex;align-items:center;gap:5px;border-radius:999px;padding:5px 8px;font-size:9px;font-weight:850;letter-spacing:.025em;white-space:nowrap}.badge.green{background:rgba(31,157,102,.12);color:var(--green)}.badge.red{background:rgba(229,72,77,.12);color:var(--red)}.badge.amber{background:rgba(245,158,11,.13);color:#b86f00}.badge.blue{background:rgba(37,99,235,.12);color:var(--blue)}.badge.purple{background:rgba(117,89,217,.13);color:var(--purple)}.badge.gray{background:rgba(100,116,139,.12);color:var(--slate)}
    .empty{padding:42px;text-align:center;color:var(--muted);font-size:13px}.empty-icon{font-size:32px;margin-bottom:8px}.pager{display:flex;align-items:center;justify-content:space-between;gap:12px;padding-top:12px;color:var(--muted);font-size:11px}.pager-controls{display:flex;gap:6px}.pager button{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:9px;padding:6px 9px;cursor:pointer}.pager button:disabled{opacity:.4;cursor:not-allowed}
    .channel-hero{display:flex;align-items:center;gap:15px}.channel-logo{width:54px;height:54px;border-radius:16px;background:#fff;display:grid;place-items:center;font-size:21px;font-weight:900;color:white;box-shadow:0 10px 24px color-mix(in srgb,var(--channel,#2563eb) 25%,transparent);border:1px solid var(--line);padding:8px}.channel-logo svg{width:100%;height:100%}.channel-name{font-size:26px;font-weight:900;letter-spacing:-.04em}.channel-kam{font-size:12px;color:var(--muted);margin-top:4px}.channel-state{margin-left:auto;text-align:right}.channel-state b{font-size:13px}.channel-state span{display:block;font-size:11px;color:var(--muted);margin-top:4px}
    .method-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:15px}.method-card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px}.method-number{width:34px;height:34px;border-radius:11px;background:rgba(37,99,235,.11);color:var(--blue);display:grid;place-items:center;font-weight:900;margin-bottom:12px}.method-card h3{font-size:14px;margin:0 0 7px}.method-card p{font-size:11px;line-height:1.6;color:var(--muted);margin:0}.formula{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#0f172a;color:#dbeafe;padding:13px;border-radius:12px;font-size:11px;overflow:auto;margin-top:10px}.dark .formula{background:#060b14}
    .quality-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:18px}.quality-kpi{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:17px}.quality-kpi strong{display:block;font-size:25px}.quality-kpi span{font-size:11px;color:var(--muted)}
    .decision-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:18px}.decision-card{border:1px solid var(--line);border-radius:16px;padding:16px;background:var(--card);box-shadow:0 8px 24px rgba(15,23,42,.05)}.decision-route{display:flex;align-items:center;gap:9px;font-weight:850;font-size:13px}.decision-route .arrow{color:var(--amber);font-size:18px}.decision-qty{font-size:28px;font-weight:900;letter-spacing:-.04em;margin:12px 0 3px}.decision-meta{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin:12px 0}.decision-metric{background:var(--page);border-radius:11px;padding:9px}.decision-metric b{display:block;font-size:13px}.decision-metric span{font-size:9px;color:var(--muted)}.decision-actions{display:flex;gap:7px;flex-wrap:wrap;margin-top:12px}.primary-btn,.secondary-btn{border:0;border-radius:10px;padding:9px 12px;font-size:11px;font-weight:800;cursor:pointer}.primary-btn{background:var(--blue);color:#fff}.secondary-btn{background:var(--page);color:var(--text);border:1px solid var(--line)}.form-grid{display:grid;grid-template-columns:1.2fr 1fr 1fr .7fr auto;gap:10px;align-items:end}.form-field label{display:block;font-size:10px;font-weight:800;color:var(--muted);margin:0 0 5px;text-transform:uppercase;letter-spacing:.05em}.form-field input,.form-field select,.kam-question{width:100%;border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:10px 11px;font-size:12px;outline:none}.decision-result{border:1px solid var(--line);border-left:5px solid var(--tone,var(--blue));border-radius:15px;padding:17px;background:color-mix(in srgb,var(--tone,var(--blue)) 5%,var(--card));margin-top:14px}.decision-result h3{margin:0 0 6px;font-size:17px}.decision-result p{margin:5px 0;color:var(--muted);font-size:11px;line-height:1.55}.scenario-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:13px}.scenario-box{border:1px solid var(--line);border-radius:12px;padding:12px;background:var(--card)}.scenario-box b{font-size:12px}.scenario-box div{font-size:10px;color:var(--muted);margin-top:5px;line-height:1.5}.question-row{display:grid;grid-template-columns:1fr auto;gap:9px}.history-empty{padding:24px;text-align:center;color:var(--muted);font-size:11px}
    .b2b-strip{display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:18px}.b2b-sim-grid{display:grid;grid-template-columns:1.15fr .75fr .75fr auto;gap:10px;align-items:end}.b2b-result{margin-top:14px;border:1px solid var(--line);border-left:5px solid var(--tone,var(--blue));border-radius:15px;padding:17px;background:color-mix(in srgb,var(--tone,var(--blue)) 5%,var(--card))}.b2b-result-main{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:12px}.b2b-result-metric{border:1px solid var(--line);border-radius:12px;padding:11px;background:var(--card)}.b2b-result-metric b{display:block;font-size:18px}.b2b-result-metric span{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:.05em}.b2b-explanation{display:none;margin-top:14px;border:1px solid color-mix(in srgb,var(--purple) 26%,var(--line));border-radius:15px;background:color-mix(in srgb,var(--purple) 4%,var(--card));padding:17px}.b2b-explanation.open{display:block;animation:panelIn .2s ease}.b2b-steps{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin-top:12px}.b2b-step{border:1px solid var(--line);border-radius:12px;padding:12px;background:var(--card)}.b2b-step b{font-size:12px}.b2b-step p{font-size:10px;color:var(--muted);line-height:1.55;margin:6px 0 0}.b2b-volume-bar{height:9px;background:var(--line);border-radius:999px;overflow:hidden;margin-top:10px}.b2b-volume-bar span{display:block;height:100%;background:linear-gradient(90deg,var(--blue),var(--teal));border-radius:999px}.b2b-policy{padding:13px 14px;border-radius:13px;border:1px solid color-mix(in srgb,var(--green) 25%,var(--line));background:color-mix(in srgb,var(--green) 5%,var(--card));font-size:11px;color:var(--muted);line-height:1.55;margin-bottom:14px}.b2b-policy b{color:var(--text)}@media(max-width:1050px){.b2b-strip{grid-template-columns:repeat(2,1fr)}.b2b-result-main,.b2b-steps{grid-template-columns:1fr 1fr}.b2b-sim-grid{grid-template-columns:1fr 1fr}}@media(max-width:700px){.b2b-strip,.b2b-result-main,.b2b-steps,.b2b-sim-grid{grid-template-columns:1fr}}
    .b2b-detail-shell{margin-top:18px;border:1px solid var(--line);border-radius:18px;background:var(--card);padding:20px;box-shadow:var(--shadow)}
    .b2b-detail-toolbar{display:flex;justify-content:space-between;align-items:end;gap:14px;margin-bottom:16px}.b2b-detail-toolbar .form-field{min-width:min(420px,100%)}
    .b2b-detail-head{display:flex;justify-content:space-between;gap:16px;align-items:flex-start;margin-bottom:12px}.b2b-detail-title h3{margin:0;font-size:24px;line-height:1.15;letter-spacing:-.02em}.b2b-detail-title p{margin:4px 0 0;color:var(--muted);font-size:12px}.b2b-live-badge{display:inline-flex;align-items:center;gap:6px;padding:7px 10px;border-radius:999px;border:1px solid var(--line);font-size:10px;font-weight:800;letter-spacing:.05em;text-transform:uppercase;background:var(--card)}.b2b-live-badge.live{color:var(--green);background:color-mix(in srgb,var(--green) 7%,var(--card));border-color:color-mix(in srgb,var(--green) 28%,var(--line))}.b2b-live-badge.base{color:var(--amber);background:color-mix(in srgb,var(--amber) 7%,var(--card));border-color:color-mix(in srgb,var(--amber) 28%,var(--line))}
    .b2b-kpi-grid{display:grid;grid-template-columns:repeat(5,minmax(0,1fr));gap:12px;margin:14px 0}.b2b-kpi-card{border:1px solid var(--line);border-radius:14px;padding:14px;background:color-mix(in srgb,var(--blue) 1.8%,var(--card));min-height:94px}.b2b-kpi-card b{display:block;font-size:23px;line-height:1.05;margin-bottom:8px}.b2b-kpi-card span{font-size:10px;color:var(--muted);text-transform:uppercase;letter-spacing:.045em;line-height:1.35}.b2b-kpi-card.negative b{color:var(--red)}.b2b-kpi-card.positive b{color:var(--green)}
    .b2b-info-box{margin:12px 0 16px;padding:13px 15px;border-radius:13px;border:1px solid color-mix(in srgb,var(--blue) 32%,var(--line));background:color-mix(in srgb,var(--blue) 6%,var(--card));font-size:11px;line-height:1.5;color:color-mix(in srgb,var(--blue) 85%,var(--text))}
    .b2b-detail-grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.b2b-detail-card{border:1px solid var(--line);border-radius:15px;padding:16px;background:var(--card)}.b2b-detail-card h4{margin:0;font-size:17px}.b2b-detail-card .detail-sub{font-size:11px;color:var(--muted);margin:4px 0 14px;line-height:1.45}
    .b2b-qty-line{display:grid;grid-template-columns:1fr 118px;gap:12px;align-items:center;margin:10px 0 14px}.b2b-qty-line input[type=range]{width:100%;accent-color:var(--blue)}.b2b-qty-line input[type=number]{width:100%;padding:10px 11px;border:1px solid var(--line);border-radius:10px;background:var(--card);color:var(--text);font-weight:700}
    .b2b-sim-output{border-left:5px solid var(--blue);background:color-mix(in srgb,var(--blue) 3.5%,var(--card));border-radius:14px;padding:14px 14px 13px}.b2b-sim-metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px}.b2b-sim-metric b{display:block;font-size:22px;line-height:1.05;margin-bottom:5px}.b2b-sim-metric span{font-size:9px;color:var(--muted);line-height:1.35;text-transform:uppercase;letter-spacing:.04em}.b2b-marginal-note{margin-top:12px;font-size:11px;color:var(--muted);line-height:1.5}.b2b-offer-row{display:grid;grid-template-columns:1fr auto;gap:8px;align-items:end;margin-top:12px}.b2b-offer-row .form-field{margin:0}.b2b-offer-verdict{margin-top:8px;font-size:10px;color:var(--muted)}
    .b2b-curve-wrap{height:280px;border:1px solid var(--line);border-radius:12px;background:var(--card);overflow:hidden}.b2b-curve-wrap svg{width:100%;height:100%;display:block}.b2b-curve-line{fill:none;stroke:var(--blue);stroke-width:2.5;vector-effect:non-scaling-stroke}.b2b-curve-point{fill:var(--red)}.b2b-curve-axis{stroke:var(--line);stroke-width:1;vector-effect:non-scaling-stroke}.b2b-curve-zero{stroke:color-mix(in srgb,var(--muted) 60%,transparent);stroke-width:1;stroke-dasharray:5 4;vector-effect:non-scaling-stroke}.b2b-curve-label{font-size:10px;fill:var(--muted)}.b2b-curve-foot{font-size:10px;color:var(--muted);margin-top:7px;line-height:1.4}
    .b2b-channel-block{margin-top:18px}.b2b-channel-block h4{margin:0 0 8px;font-size:16px}.b2b-channel-table{max-height:390px;overflow:auto;border:1px solid var(--line);border-radius:13px}.b2b-channel-table table{min-width:1120px}.b2b-channel-table td,.b2b-channel-table th{white-space:nowrap}.b2b-channel-table .warn{color:var(--amber);font-weight:700}
    @media(max-width:1150px){.b2b-kpi-grid{grid-template-columns:repeat(3,1fr)}.b2b-detail-grid{grid-template-columns:1fr}}@media(max-width:760px){.b2b-detail-toolbar,.b2b-detail-head{flex-direction:column;align-items:stretch}.b2b-kpi-grid{grid-template-columns:repeat(2,1fr)}.b2b-sim-metrics{grid-template-columns:repeat(2,1fr)}.b2b-detail-shell{padding:13px}}@media(max-width:480px){.b2b-kpi-grid,.b2b-sim-metrics{grid-template-columns:1fr}.b2b-qty-line,.b2b-offer-row{grid-template-columns:1fr}}

    .combo{position:relative}.combo input{width:100%;border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:11px;padding:10px 34px 10px 11px;font-size:12px;outline:none}.combo input:focus{border-color:var(--blue)}.combo-clear{position:absolute;right:8px;top:50%;transform:translateY(-50%);border:0;background:none;color:var(--muted);cursor:pointer;font-size:14px;width:22px;height:22px;border-radius:7px;display:none;align-items:center;justify-content:center}.combo-clear.show{display:flex}.combo-clear:hover{background:var(--page);color:var(--text)}.combo-panel{position:absolute;top:calc(100% + 6px);left:0;min-width:100%;width:max-content;max-width:min(560px,92vw);background:var(--card);border:1px solid var(--line);border-radius:13px;box-shadow:0 18px 40px rgba(15,23,42,.14);max-height:340px;overflow:auto;z-index:40;display:none}.combo-panel.open{display:block;animation:panelIn .15s ease}.combo-option{display:flex;align-items:center;gap:14px;padding:11px 14px;cursor:pointer;font-size:13px;border-bottom:1px solid var(--line)}.combo-option:last-child{border-bottom:0}.combo-option:hover,.combo-option.active{background:rgba(37,99,235,.08)}.combo-option .co-sku{font-weight:850;color:var(--blue);white-space:nowrap;flex:none}.combo-option .co-name{color:var(--text);white-space:normal;line-height:1.4;flex:1}.combo-option mark{background:rgba(245,158,11,.35);color:inherit;border-radius:3px;padding:0 1px}.combo-empty{padding:14px 12px;text-align:center;color:var(--muted);font-size:11px}.combo-count{padding:7px 14px;font-size:9px;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);border-bottom:1px solid var(--line);background:var(--page)}
    .toast{position:fixed;right:22px;bottom:22px;background:#101828;color:white;border-radius:13px;padding:12px 15px;font-size:12px;box-shadow:0 15px 35px rgba(0,0,0,.25);opacity:0;pointer-events:none;transform:translateY(8px);transition:.2s;z-index:100}.toast.show{opacity:1;transform:none}
    .overlay{display:none}
    .admin-status-btn{border:1px solid var(--border);background:var(--card);color:var(--muted);height:38px;border-radius:10px;padding:0 12px;font-weight:800;cursor:pointer;display:inline-flex;align-items:center;gap:7px}
    .admin-status-btn.on{color:var(--green);border-color:rgba(31,157,102,.35);background:rgba(31,157,102,.08)}
    .admin-transfer-btn{border:0;background:#111827;color:#fff;border-radius:8px;padding:7px 10px;font-weight:800;font-size:12px;cursor:pointer;white-space:nowrap}
    .admin-transfer-btn:hover{transform:translateY(-1px);box-shadow:0 5px 14px rgba(15,23,42,.18)}
    .admin-transfer-btn:disabled{opacity:.45;cursor:not-allowed;transform:none;box-shadow:none}
    .admin-modal-backdrop{position:fixed;inset:0;background:rgba(3,8,17,.58);z-index:200;display:none;align-items:center;justify-content:center;padding:18px}
    .admin-modal-backdrop.open{display:flex}
    .admin-modal{width:min(520px,100%);background:var(--card);border:1px solid var(--border);border-radius:18px;box-shadow:0 30px 80px rgba(0,0,0,.32);padding:22px}
    .admin-modal h3{margin:0 0 6px;font-size:20px}
    .admin-modal p{margin:0;color:var(--muted);line-height:1.55}
    .admin-modal-grid{display:grid;grid-template-columns:1fr 1fr;gap:10px;margin:18px 0}
    .admin-fact{border:1px solid var(--border);border-radius:12px;padding:11px;background:var(--soft)}
    .admin-fact span{display:block;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.06em;font-weight:800}
    .admin-fact b{display:block;margin-top:4px;font-size:15px}
    .admin-password{width:100%;height:42px;border:1px solid var(--border);background:var(--bg);color:var(--text);border-radius:10px;padding:0 12px;margin-top:14px}
    .admin-actions{display:flex;gap:10px;justify-content:flex-end;margin-top:18px;flex-wrap:wrap}
    .admin-actions button{height:40px;border-radius:10px;padding:0 14px;font-weight:800;cursor:pointer}
    .admin-secondary{border:1px solid var(--border);background:var(--card);color:var(--text)}
    .admin-primary{border:0;background:#111827;color:#fff}
    .admin-danger-note{margin-top:12px;padding:10px 12px;border-radius:10px;background:rgba(229,72,77,.08);color:#b42318;font-size:12px;font-weight:700}
    .admin-message{min-height:20px;margin-top:10px;font-size:12px;font-weight:700}
    .admin-message.error{color:var(--red)}
    .admin-message.ok{color:var(--green)}
    @media(max-width:1350px){.decision-grid{grid-template-columns:1fr 1fr}.form-grid{grid-template-columns:1fr 1fr}.kpi-grid{grid-template-columns:repeat(3,1fr)}.grid-3{grid-template-columns:1fr 1fr}.grid-3>.card:first-child{grid-column:1/-1}.quality-grid{grid-template-columns:repeat(2,1fr)}}
    @media(max-width:1050px){.sidebar{transform:translateX(-100%);width:var(--sidebar)!important}.sidebar.mobile-open{transform:none}.main,.sidebar.collapsed~.main{margin-left:0;width:100%}.mobile-menu{display:block}.overlay{display:block;position:fixed;inset:0;background:rgba(3,8,17,.45);z-index:45;opacity:0;pointer-events:none;transition:.2s}.overlay.show{opacity:1;pointer-events:auto}.topbar{padding:12px 17px}.content{padding:20px 17px}.search-global{display:none}.grid-2,.grid-73{grid-template-columns:1fr}.method-grid{grid-template-columns:1fr 1fr}}
    @media(max-width:700px){.decision-grid,.form-grid,.scenario-grid{grid-template-columns:1fr}.kpi-grid{grid-template-columns:1fr 1fr}.status-strip{grid-template-columns:1fr 1fr}.grid-3{grid-template-columns:1fr}.grid-3>.card:first-child{grid-column:auto}.method-grid,.quality-grid{grid-template-columns:1fr}.hero{padding:23px;align-items:flex-start;flex-direction:column}.hero h1{font-size:25px}.hero-meta{justify-content:flex-start}.topbar-right .icon-btn:first-child{display:none}.page-title{font-size:17px}.channel-state{display:none}}
    @media(max-width:460px){.kpi-grid{grid-template-columns:1fr}.status-strip{grid-template-columns:1fr}}
  
    /* ---- Versión 2026-09: canal orientado a decisiones del KAM ---- */
    .today-card{margin-bottom:12px}
    .action-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(175px,1fr));gap:12px}
    .action-tile{all:unset;box-sizing:border-box;cursor:pointer;display:flex;flex-direction:column;gap:6px;padding:15px 16px;border-radius:14px;border:1px solid var(--line);border-top:4px solid var(--tone,var(--blue));background:color-mix(in srgb,var(--tone,var(--blue)) 5%,var(--card));min-height:118px}
    .action-tile:hover{border-color:color-mix(in srgb,var(--tone,var(--blue)) 45%,var(--line))}
    .action-tile:focus-visible{outline:3px solid color-mix(in srgb,var(--tone,var(--blue)) 35%,transparent);outline-offset:2px}
    .action-tile .a-title{font-size:12px;font-weight:850;color:var(--text)}
    .action-tile .a-value{font-size:24px;font-weight:900;letter-spacing:-.03em;color:var(--tone,var(--blue))}
    .action-tile .a-text{font-size:11px;color:var(--muted);line-height:1.45}
    .action-tile.quiet{--tone:var(--green)}
    .channel-jump{position:sticky;top:var(--topbar-h,64px);z-index:5;display:flex;gap:6px;flex-wrap:wrap;padding:8px;margin:0 0 18px;border:1px solid var(--line);border-radius:13px;background:color-mix(in srgb,var(--card) 92%,transparent);backdrop-filter:blur(6px)}
    .channel-jump button{border:0;background:transparent;color:var(--muted);font-size:12px;font-weight:750;padding:7px 11px;border-radius:9px;cursor:pointer}
    .channel-jump button:hover,.channel-jump button:focus-visible{background:color-mix(in srgb,var(--blue) 10%,transparent);color:var(--blue);outline:none}
    #secAge,#secCoverage,#secListings,#channelSalesCard,#secMoves,#secDetail,#secToday,#channelCoverageCard{scroll-margin-top:140px}
    .block-heading{margin:30px 0 14px;padding-top:18px;border-top:1px solid var(--line)}
    .block-heading h3{margin:0;font-size:21px;letter-spacing:-.03em}
    .block-heading p{margin:5px 0 0;color:var(--muted);font-size:13px;max-width:820px}
    .status-strip.wide{grid-template-columns:repeat(auto-fit,minmax(170px,1fr))}
    .status-card.clickable{cursor:pointer;position:relative}
    .status-card.clickable:hover{border-color:color-mix(in srgb,var(--blue) 40%,var(--line))}
    .status-card.clickable:focus-visible{outline:3px solid rgba(37,99,235,.25);outline-offset:2px}
    .status-card.active{border-color:var(--blue);box-shadow:0 0 0 3px rgba(37,99,235,.12)}
    .status-card .status-hint{display:block;font-size:10px;font-weight:800;color:var(--blue);margin-top:3px}
    .status-dl{margin-left:auto;align-self:flex-start;border:1px solid var(--line);background:var(--card);color:var(--muted);border-radius:8px;padding:5px 8px;font-size:10px;font-weight:800;cursor:pointer;white-space:nowrap}
    .status-dl:hover{color:var(--blue);border-color:var(--blue)}
    .coverage-guide{display:grid;gap:7px}
    .coverage-guide .cg-row{display:grid;grid-template-columns:150px 1fr;gap:10px;align-items:center;font-size:11px;color:var(--muted);line-height:1.4}
    .filter-banner{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:9px 12px;margin:0 0 12px;border-radius:11px;background:rgba(37,99,235,.07);color:var(--text);font-size:12px}
    .filter-banner button{border:1px solid var(--line);background:var(--card);color:var(--text);border-radius:8px;padding:5px 9px;font-size:11px;font-weight:750;cursor:pointer}
    
    @media(max-width:760px){.action-grid{grid-template-columns:1fr 1fr}.channel-jump{position:static}.coverage-guide .cg-row{grid-template-columns:1fr}}
    @media(prefers-reduced-motion:reduce){*{scroll-behavior:auto!important;animation:none!important;transition:none!important}}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar" id="sidebar">
    <div class="brand">
      <div class="brand-mark">IQ</div>
      <div class="brand-copy"><div class="brand-title">Inventory Control Tower</div><div class="brand-sub">Rotación · Canales · Full</div></div>
      <button class="collapse-btn" id="collapseBtn" title="Contraer menú">‹</button>
    </div>
    <nav class="nav">
      <div class="nav-group">
        <div class="nav-group-title">Control general</div>
        <button class="nav-item active" data-page="executive"><span class="nav-icon">⌂</span><span class="nav-label">Resumen ejecutivo</span></button>
      </div>
      <div class="nav-group">
        <div class="nav-group-title">Canales y KAMs</div>
        <div id="channelNav"></div>
      </div>
      <div class="nav-group">
        <div class="nav-group-title">Decisiones</div>
        <button class="nav-item" data-page="suggestions"><span class="nav-icon">✦</span><span class="nav-label">Sugerencias</span><span class="nav-badge" id="suggestionBadge">0</span></button>
        <button class="nav-item" data-page="distribution"><span class="nav-icon">⇉</span><span class="nav-label">Repartición</span><span class="nav-badge" id="distributionBadge">0</span></button>
        <button class="nav-item" data-page="b2b"><span class="nav-icon">$</span><span class="nav-label">B2B financiero</span><span class="nav-badge" id="b2bBadge">0</span></button>
        <button class="nav-item" data-page="transfers"><span class="nav-icon">⇄</span><span class="nav-label">Odoo → Full</span><span class="nav-badge" id="transferBadge">0</span></button>
        <button class="nav-item" data-page="purchases"><span class="nav-icon">◫</span><span class="nav-label">Compras sugeridas</span><span class="nav-badge" id="purchaseBadge">0</span></button>
      </div>
      <div class="nav-group">
        <div class="nav-group-title">Análisis</div>
        <button class="nav-item" data-page="lots"><span class="nav-icon">◫</span><span class="nav-label">Lotes y antigüedad</span></button>
        <button class="nav-item" data-page="audit"><span class="nav-icon">⇆</span><span class="nav-label">Movimientos Odoo</span><span class="nav-badge" id="auditBadge">0</span></button>
        <button class="nav-item" data-page="alerts"><span class="nav-icon">!</span><span class="nav-label">Alertas</span><span class="nav-badge" id="alertBadge">0</span></button>
        <button class="nav-item" data-page="quality"><span class="nav-icon">✓</span><span class="nav-label">Calidad de datos</span></button>
        <button class="nav-item" data-page="methodology"><span class="nav-icon">i</span><span class="nav-label">Metodología</span></button>
      </div>
    </nav>
    <div class="sidebar-footer"><div class="status-dot"></div><div class="sidebar-footer-copy">Base actualizada<br><span id="footerDate"></span></div></div>
  </aside>
  <div class="overlay" id="overlay"></div>

  <main class="main">
    <header class="topbar">
      <div class="topbar-left">
        <button class="mobile-menu" id="mobileMenu">☰</button>
        <div><div class="page-eyebrow" id="pageEyebrow">Control Tower</div><div class="page-title" id="pageTitle">Resumen ejecutivo</div></div>
      </div>
      <div class="topbar-right">
        <div class="search-global"><input id="globalSearch" placeholder="Filtrar esta pestaña por IQ, SKU o producto…" /></div>
        <button class="admin-status-btn" id="adminStatusBtn" type="button" title="Usuarios y permisos se habilitarán en la fase multiusuario" style="display:none">🔒 Admin</button>
        <button class="icon-btn" id="themeBtn" title="Cambiar tema">◐</button>
        <button class="icon-btn" id="refreshBtn" title="Volver a cargar">↻</button>
      </div>
    </header>

    <div class="content">
      <section class="page active" id="page-executive">
        <div class="semaphore-strip" id="semaphoreStrip"></div>
        <div class="action-banner" id="actionBanner"></div>
        <div class="hero">
          <div class="hero-copy"><div class="hero-kicker">Inventario omnicanal</div><h1>Stock y ventas en una sola vista.</h1><p>Consulta cada IQ, sus SKU originales, el inventario disponible y las ventas recientes. Desde aquí puedes abrir y descargar el desglose diario de ventas por canal y SKU.</p></div>
          <div class="hero-meta"><span class="hero-chip" id="heroDate"></span><span class="hero-chip">Ventas: <b>10 / 30 / 90 días</b></span><span class="hero-chip">Odoo → Full: <b>__FULL_TRANSIT_DAYS__ días</b></span><span class="hero-chip" style="cursor:pointer" onclick="goPage('methodology')">Metodología →</span></div>
        </div>
        <div class="kpi-grid" id="executiveKpis"></div>
        <div class="grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Stock por canal</div><div class="card-sub">Odoo, tránsito y Full en cada canal.</div></div></div><div id="chartInventoryChannel" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Ventas por canal</div><div class="card-sub">Comparación de unidades vendidas en 10, 30 y 90 días.</div></div></div><div id="chartSalesChannel" class="chart"></div></div>
        </div>
        <div class="grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Productos con más ventas</div><div class="card-sub">Top de IQ por unidades vendidas en los últimos 90 días.</div></div></div><div id="chartTopProducts" class="chart"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Mix de inventario</div><div class="card-sub">Dónde se encuentran físicamente las piezas.</div></div></div><div id="chartInventoryMix" class="chart"></div></div>
        </div>
        <div class="card executive-products-card">
          <div class="card-head">
            <div><div class="card-title">Resumen general por producto</div><div class="card-sub">Busca por IQ, SKU original o producto. La tabla muestra stock y ventas; el botón abre el desglose descargable.</div></div>
            <div class="section-tools"><input class="filter-input" id="executiveProductSearch" placeholder="Buscar IQ, SKU original o producto"/><button class="card-action" onclick="exportCurrentTable('executiveProducts')">Exportar resumen</button></div>
          </div>
          <div id="executiveProductTable"></div>
        </div>
        <div class="card sales-detail-card" id="salesDetailPanel">
          <div class="card-head">
            <div><div class="card-title" id="salesDetailTitle">Detalle de ventas</div><div class="card-sub" id="salesDetailSubtitle">Selecciona un producto para consultar su desglose.</div></div>
            <div class="section-tools"><select class="select" id="salesDetailPeriod"><option value="90">Últimos 90 días</option><option value="30">Últimos 30 días</option><option value="10">Últimos 10 días</option></select><button class="card-action" onclick="exportSkuSales()">Descargar ventas del SKU</button><button class="odoo-close" onclick="closeSalesDetail()">×</button></div>
          </div>
          <div class="status-strip" id="salesDetailStatus"></div>
          <div class="grid-2 sales-detail-grid"><div><div class="mini-section-title">Resumen por canal y SKU original</div><div id="salesBreakdownTable"></div></div><div><div class="mini-section-title">Ventas diarias</div><div id="salesDailyTable"></div></div></div>
        </div>
      </section>

      <section class="page" id="page-channel">
        <div class="section-heading">
          <div class="channel-hero"><div class="channel-logo" id="channelLogo">A</div><div><div class="channel-name" id="channelName">Amazon</div><div class="channel-kam" id="channelKam">KAM Amazon</div></div></div>
          <div class="channel-state"><b id="channelState">Estado del canal</b><span id="channelUpdated"></span></div>
        </div>
        <div class="kpi-grid" id="channelKpis"></div>

        <div class="card today-card" id="secToday">
          <div class="card-head"><div><div class="card-title">Qué hacer hoy</div><div class="card-sub">Las decisiones del canal ordenadas por urgencia. Cada recuadro abre el detalle correspondiente.</div></div></div>
          <div class="action-grid" id="channelActions"></div>
        </div>
        <nav class="channel-jump" aria-label="Secciones del canal">
          <button type="button" data-target="secToday">Qué hacer hoy</button>
          <button type="button" data-target="secAge">Antigüedad</button>
          <button type="button" data-target="secCoverage">Cobertura</button>
          <button type="button" data-target="secListings">Publicaciones</button>
          <button type="button" data-target="channelSalesCard">Ventas</button>
          <button type="button" data-target="secMoves">Movimientos Odoo</button>
          <button type="button" data-target="secDetail">Detalle</button>
        </nav>

        <div class="card" id="secAge"><div class="card-head"><div><div class="card-title">Antigüedad de inventario por canal (+90 días)</div><div class="card-sub" id="channelStaleSub">Cuánto tiempo llevan las piezas en este canal. En Odoo se cuenta desde la compra del lote; en Full, desde el envío a Full (las piezas que quedan son las de los envíos más recientes). Una devolución ya no reinicia el conteo. La tabla inicia por la inversión más alta.</div></div><div class="section-tools"><input class="filter-input" id="channelStaleSearch" placeholder="Buscar SKU o producto"/><select class="select" id="channelStaleStockFilter"><option value="">Toda antigüedad</option><option value="30">Más de 30 días</option><option value="60">Más de 60 días</option><option value="90">Más de 90 días</option></select><button class="card-action" onclick="exportCurrentTable('channelStale')">Exportar CSV</button></div></div><div class="status-strip wide" id="channelStaleStatus"></div><div id="channelStaleTable"></div></div>

        <div class="block-heading" id="secCoverage"><div><h3>Cobertura</h3><p>Cuántos días alcanza el inventario del canal al ritmo de venta actual, y qué hacer para no quedarse sin stock ni acumular de más.</p></div></div>
        <div class="status-strip wide" id="channelCoverageStatus"></div>
        <div class="grid-73">
          <div class="card"><div class="card-head"><div><div class="card-title">Mapa de cobertura del canal</div><div class="card-sub">Cada punto es un SKU: a la derecha vende más, arriba tiene más días de inventario. La línea roja marca 30 días y la azul 90.</div></div></div><div id="chartChannelRisk" class="chart tall"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Sugerencias del KAM</div><div class="card-sub">Los SKU que conviene atender primero y qué hacer con cada uno.</div></div></div><div class="insights" id="channelInsights"></div></div>
        </div>
        <div class="grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Dónde está el inventario</div><div class="card-sub">Piezas del canal en bodega (Odoo), en camino a Full y ya en Full.</div></div></div><div id="chartChannelInventory" class="chart small"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Cómo leer la cobertura</div><div class="card-sub">Días de cobertura = piezas disponibles ÷ venta diaria de los últimos 90 días.</div></div></div><div class="coverage-guide" id="coverageGuide"></div></div>
        </div>
        <div class="card odoo-detail-card" id="channelOdooPanel" aria-live="polite">
          <div class="card-head">
            <div><div class="card-title">Stock Odoo y sugerencia de envío a Full</div><div class="card-sub" id="channelOdooSubtitle">Detalle por producto del canal seleccionado.</div></div>
            <div class="odoo-panel-tools">
              <input class="filter-input" id="channelOdooSearch" placeholder="Buscar IQ, SKU individual o producto"/>
              <select class="select" id="channelOdooFilter"><option value="">Todos</option><option value="send">Con envío sugerido</option><option value="manual">Check manual</option><option value="covered">Sin envío requerido</option></select>
              <button class="card-action" onclick="exportCurrentTable('odooPanel')">Exportar CSV</button>
              <button class="odoo-close" onclick="closeOdooPanel()" aria-label="Cerrar detalle">×</button>
            </div>
          </div>
          <div class="status-strip" id="channelOdooStatus"></div>
          <div class="odoo-rule" id="channelOdooRule"></div>
          <div id="channelOdooTable"></div>
        </div>
        <div class="card" id="channelCoverageCard"><div class="card-head"><div><div class="card-title">Cobertura por SKU</div><div class="card-sub">Ordenada por urgencia: primero lo que ya se quedó sin stock o está por agotarse. La acción está escrita para ejecutarse tal cual.</div></div><div class="section-tools"><input class="filter-input" id="channelCoverageSearch" placeholder="Buscar SKU o producto"/><select class="select" id="channelCoverageFilter"><option value="">Todos los estados</option><option value="urgent">Urgente: sin stock o por agotarse</option><option value="low">Cobertura baja</option><option value="ok">Cobertura sana</option><option value="excess">Cobertura alta o exceso</option><option value="nosales">Stock sin ventas</option></select><button class="card-action" onclick="exportCurrentTable('channelCoverage')">Exportar CSV</button></div></div><div id="channelCoverageTable"></div></div>

        <div class="card" id="secListings"><div class="card-head"><div><div class="card-title">Stock asignado sin publicación activa</div><div class="card-sub" id="channelListingsSub">SKU con piezas en este canal (bodega o Full) que en AutoAzur no tienen publicación, o la tienen pausada, en revisión, con error o cerrada. Inicia por la inversión más alta.</div></div><div class="section-tools"><input class="filter-input" id="channelListingsSearch" placeholder="Buscar SKU, ItemID o producto"/><select class="select" id="channelListingsFilter"><option value="">Todos los pendientes</option><option value="nopub">Sin publicación</option><option value="paused">Pausada</option><option value="review">En revisión</option><option value="error">Con error / rechazada</option><option value="closed">Inactiva o cerrada</option><option value="zero">Publicada con stock 0</option></select><button class="card-action" onclick="exportCurrentTable('channelListings')">Exportar CSV</button></div></div><div id="channelListingsNotice"></div><div class="status-strip wide" id="channelListingsStatus"></div><div id="channelListingsTable"></div></div>


        <div class="card" id="channelSalesCard"><div class="card-head"><div><div class="card-title">Ventas del canal por día</div><div class="card-sub" id="channelSalesSub">Serie diaria del canal. El eje X son los días; el eje Y cambia según la métrica que elijas. Puedes acotar a un SKU o mover el periodo.</div></div><div class="section-tools"><select class="select" id="channelSalesMetric" title="Métrica del eje Y"><option value="unidades">Unidades vendidas</option><option value="monto">Monto vendido</option><option value="utilidad">Utilidad</option><option value="roi">ROI</option></select><div class="combo" id="channelSalesSkuCombo" style="min-width:250px"><input id="channelSalesSku" autocomplete="off" placeholder="Venta general del canal"/><button type="button" class="combo-clear" id="channelSalesSkuClear">&times;</button><div class="combo-panel" id="channelSalesSkuPanel"></div></div><select class="select" id="channelSalesRange" title="Periodo"><option value="90">Últimos 90 días</option><option value="60">Últimos 60 días</option><option value="30">Últimos 30 días</option><option value="10">Últimos 10 días</option><option value="all">Todo el periodo disponible</option><option value="custom">Rango de fechas…</option></select><span class="date-range" id="channelSalesDates"><input type="date" class="select" id="channelSalesFrom"/><span class="date-sep">a</span><input type="date" class="select" id="channelSalesTo"/></span><button class="card-action" onclick="exportChannelSales()">Exportar CSV</button></div></div><div class="status-strip" id="channelSalesStatus"></div><div id="chartChannelSales" class="chart tall"></div><div class="card-sub" id="channelSalesNote"></div></div>
        <div class="card"><div class="card-head"><div><div class="card-title">Top productos por venta</div><div class="card-sub">Unidades vendidas en los últimos 90 días.</div></div></div><div id="chartChannelTop" class="chart"></div></div>

        <div class="card" id="secMoves"><div class="card-head"><div><div class="card-title">Movimientos en Odoo de este canal</div><div class="card-sub">Traslados, ajustes y envíos que sacaron o metieron piezas al canal sin ser compra ni venta. Se marcan en revisión cuando mueven piezas viejas o regresan al poco tiempo, porque pueden maquillar la antigüedad.</div></div><div class="section-tools"><input class="filter-input" id="channelMovesSearch" placeholder="Buscar SKU, usuario o referencia"/><select class="select" id="channelMovesFilter"><option value="review">En revisión (alta y media)</option><option value="ALTA">Solo revisión alta</option><option value="">Todos los movimientos</option></select><button class="card-action" onclick="exportCurrentTable('channelMoves')">Exportar CSV</button></div></div><div class="status-strip" id="channelMovesStatus"></div><div id="channelMovesTable"></div></div>

        <div class="card" id="secDetail"><div class="card-head"><div><div class="card-title">Detalle de productos</div><div class="card-sub">Inventario, ventas, coberturas y recomendación por SKU.</div></div><div class="section-tools"><div class="view-toggle" id="channelViewToggle"><button data-view="simple" class="active">Vista simple</button><button data-view="detail">Vista detallada</button></div><input class="filter-input" id="channelTableSearch" placeholder="Buscar SKU o producto"/><select class="select" id="channelRiskFilter"><option value="">Todos los riesgos</option><option>CRÍTICO</option><option>ALTO</option><option>INTERMITENTE</option><option>SOBRESTOCK</option><option>SANO</option></select><button class="card-action" onclick="exportCurrentTable('channel')">Exportar CSV</button></div></div><div id="channelTable"></div></div>
        <div class="card" id="channelDictionaryCard"><div class="card-head"><div><div class="card-title">Calidad del diccionario</div><div class="card-sub" id="channelDictionarySub">SKU de este canal que la sincronización no pudo vincular a un SKU madre (IQ).</div></div><div class="section-tools"><button class="card-action" onclick="exportCurrentTable('channelUnlinkedStock')">Exportar stock CSV</button><button class="card-action" onclick="exportCurrentTable('channelUnlinkedSales')">Exportar ventas CSV</button></div></div><div id="channelDictionaryBody"></div></div>
      </section>

      <section class="page" id="page-suggestions">
        <div class="section-heading"><div><h2>Centro de decisiones</h2><p>Redistribuciones viables, comparación de escenarios y consultas de los KAMs. Toda acción requiere autorización humana.</p></div><div class="section-tools"><input class="filter-input" id="suggestionSearch" placeholder="Buscar IQ, producto, origen o destino"/><select class="select" id="suggestionConfidence"><option value="">Toda confianza</option><option>ALTA</option><option>MEDIA</option><option>BAJA</option></select><button class="card-action" onclick="exportCurrentTable('redistributions')">Exportar CSV</button></div></div>
        <div class="status-strip" id="suggestionStatus"></div>
        <div class="card"><div class="card-head"><div><div class="card-title">Redistribuciones prioritarias</div><div class="card-sub">El sistema identifica excedentes entre canales. Usa primero Odoo asignado; si requiere retirar piezas de Full, lo marca para revisión operativa.</div></div></div><div class="decision-grid" id="suggestionCards"></div><div id="redistributionTable"></div></div>
        <div class="grid-2" style="margin-top:18px">
          <div class="card"><div class="card-head"><div><div class="card-title">Evaluar propuesta KAM</div><div class="card-sub">Compara la cantidad propuesta contra mantener la distribución actual y contra la cantidad optimizada.</div></div></div><div class="form-grid" style="grid-template-columns:1.15fr 1fr 1fr .7fr auto"><div class="form-field"><label>SKU madre o individual</label><input id="kamSku" list="skuDecisionList" placeholder="IQ123 o SKU canal"/><datalist id="skuDecisionList"></datalist></div><div class="form-field"><label>Origen</label><select id="kamOrigin"></select></div><div class="form-field"><label>Destino</label><select id="kamDestination"></select></div><div class="form-field"><label>Piezas</label><input id="kamQty" type="number" min="0" step="1" placeholder="30"/></div><button class="primary-btn" onclick="evaluateKamForm()">Evaluar</button></div><div id="kamEvaluationResult"><div class="history-empty">Selecciona un SKU, origen, destino y cantidad para comparar escenarios.</div></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Pregúntale al dashboard</div><div class="card-sub">Asistente determinista: interpreta preguntas de traslado usando únicamente los datos cargados.</div></div></div><div class="question-row"><input class="kam-question" id="kamQuestion" placeholder="¿Conviene mover 30 piezas de Coppel a Liverpool del IQ123?"/><button class="primary-btn" onclick="answerKamQuestion()">Analizar</button></div><div id="kamQuestionResult"><div class="history-empty">Ejemplo: “¿Cuánto conviene mover de Walmart a Liverpool del IQ825?”</div></div></div>
        </div>
        <div class="card"><div class="card-head"><div><div class="card-title">Solicitudes guardadas en este navegador</div><div class="card-sub">Registro local para revisión. No ejecuta movimientos ni sustituye la autorización en Odoo.</div></div><button class="card-action" onclick="clearDecisionHistory()">Limpiar historial</button></div><div id="decisionHistory"></div></div>
      </section>

      <section class="page" id="page-distribution">
        <div class="section-heading">
          <div><h2>Repartición de stock</h2><p>Distribuye el stock disponible de CUATI/Existencias y CUATI/B2B entre canales usando primero el ritmo del IQ; si no existe suficiente historia, baja a categoría + marca, categoría, marca y finalmente reparto parejo.</p></div>
          <div class="section-tools"><input class="filter-input" id="distributionSearch" placeholder="Buscar IQ, producto, categoría o marca"/><select class="select" id="distributionMethod"><option value="">Todos los métodos</option><option>Histórico SKU</option><option>Categoría + marca</option><option>Categoría</option><option>Marca</option><option>Parejo</option></select><button class="card-action" onclick="exportCurrentTable('distribution')">Exportar CSV</button></div>
        </div>
        <div class="status-strip wide" id="distributionStatus"></div>
        <div class="card">
          <div class="card-head"><div><div class="card-title">Stock pendiente de repartir</div><div class="card-sub">La orden de compra y el último arribo son referencia; la cantidad repartible sale del stock disponible vivo de CUATI/Existencias y CUATI/B2B. La sugerencia ya descuenta el inventario que cada canal tiene actualmente.</div></div></div>
          <div id="distributionTable"></div>
        </div>
        <div class="card sales-detail-card" id="distributionDetailPanel">
          <div class="card-head">
            <div><div class="card-title" id="distributionDetailTitle">Detalle de repartición</div><div class="card-sub" id="distributionDetailSubtitle">Selecciona un IQ para ver la propuesta por canal.</div></div>
            <div class="section-tools"><button class="primary-btn" onclick="createAllDistributionDrafts()">Crear borradores válidos</button><button class="card-action" onclick="exportDistributionDetail()">Exportar detalle CSV</button><button class="odoo-close" onclick="closeDistributionDetail()">×</button></div>
          </div>
          <div class="status-strip wide" id="distributionDetailStatus"></div>
          <div class="grid-2">
            <div><div class="mini-section-title">Stock actual vs. stock después</div><div id="distributionChart" class="chart"></div></div>
            <div><div class="mini-section-title">Cómo se decidió</div><div class="insights" id="distributionMethodology"></div></div>
          </div>
          <div style="margin-top:18px"><div class="mini-section-title">Propuesta por canal</div><div id="distributionDetailTable"></div></div>
        </div>
      </section>

      <section class="page" id="page-b2b">
        <div class="section-heading"><div><h2>Portafolio B2B · mínimo 10% + VPN</h2><p>Marketplace aparta primero su cobertura. El excedente disponible se compara contra el costo de oportunidad VPN, pero ninguna recomendación B2B baja de 10% ROI.</p></div><div class="section-tools"><button class="card-action" onclick="exportCurrentTable('b2b')">Exportar CSV</button></div></div>
        <div class="b2b-strip" id="b2bStatus"></div>
        <div class="b2b-policy"><b>Regla financiera:</b> ROI mínimo comercial B2B = 10% por pieza. Reserva por canal = Forecast × 30 días − Full − tránsito. Para cada pieza se calcula el costo de oportunidad Marketplace a valor presente y se usa <b>max(10%, ROI oportunidad VPN)</b>. Si el payout del canal está incompleto, se asumen <b>30 días</b> hasta contar con el dato real.</div>

        <div class="card">
          <div class="card-head"><div><div class="card-title">Candidatos B2B</div><div class="card-sub">Busca el SKU y usa “Simular” para abrir abajo el detalle financiero completo: stock, reserva, excedente, VPN, curva por cantidad y desglose por canal.</div></div></div>
          <div class="combo" id="b2bSearchCombo" style="margin-bottom:12px"><input id="b2bSearch" autocomplete="off" placeholder="Buscar IQ o producto"/><button type="button" class="combo-clear" id="b2bSearchClear">×</button><div class="combo-panel" id="b2bSearchPanel"></div></div>
          <div id="b2bTable"></div>
        </div>

        <div class="b2b-detail-shell" id="b2bDetailShell">
          <div class="b2b-detail-toolbar">
            <div class="form-field"><label>Producto seleccionado</label><div class="combo" id="b2bSkuCombo"><input id="b2bSku" autocomplete="off" placeholder="Buscar IQ o producto"/><button type="button" class="combo-clear" id="b2bSkuClear">×</button><div class="combo-panel" id="b2bSkuPanel"></div></div></div>
            <button class="secondary-btn" type="button" id="b2bExplainBtnTop">Ver metodología del SKU</button>
          </div>
          <div id="b2bDetailHeader"><div class="history-empty">Selecciona un candidato B2B.</div></div>
          <div id="b2bKpiGrid" class="b2b-kpi-grid"></div>
          <div class="b2b-info-box" id="b2bInfoBox">La referencia VPN sigue mostrando el costo económico de dejar de vender en Marketplace, pero la cotización B2B aplica un piso comercial de 10%. El precio mínimo por pieza se calcula con el mayor entre 10% y el costo de oportunidad correspondiente.</div>

          <div class="b2b-detail-grid">
            <div class="b2b-detail-card">
              <h4>Simulador por cantidad B2B</h4>
              <div class="detail-sub">El máximo es el excedente físico disponible. A mayor cantidad, se incorporan oportunidades Marketplace progresivamente más valiosas; donde falte VPN, permanece el piso comercial de 10%.</div>
              <div class="b2b-qty-line"><input id="b2bQtyRange" type="range" min="1" max="1" value="1" oninput="syncB2BQtyFromRange(this.value)"/><input id="b2bQty" type="number" min="1" step="1" value="1"/></div>
              <div id="b2bSimulatorResult"><div class="history-empty">Selecciona un candidato para calcular el ROI equivalente hoy.</div></div>
              <div class="b2b-offer-row"><div class="form-field"><label>ROI oferta B2B % · opcional</label><input id="b2bProposedRoi" type="number" step="0.01" placeholder="Comparar contra costo de oportunidad"/></div><button class="primary-btn" id="b2bEvaluateBtn" type="button">Evaluar oferta</button></div>
              <div id="b2bOfferVerdict" class="b2b-offer-verdict"></div>
            </div>
            <div class="b2b-detail-card">
              <h4>ROI mínimo B2B vs cantidad</h4>
              <div class="detail-sub">La curva usa el mayor entre el piso comercial de 10% y el VPN de las ventas Marketplace futuras desplazadas.</div>
              <div class="b2b-curve-wrap" id="b2bCurveChart"></div>
              <div class="b2b-curve-foot">La línea punteada marca el piso comercial de 10%. El punto rojo corresponde a la cantidad seleccionada.</div>
            </div>
          </div>

          <div class="b2b-channel-block">
            <h4>Marketplace · reserva y oportunidad futura</h4>
            <div id="b2bChannelDetail" class="b2b-channel-table"><div class="history-empty">Selecciona un SKU para ver el detalle por canal.</div></div>
          </div>
          <div id="b2bExplanation" class="b2b-explanation"></div>
        </div>
      </section>

      <section class="page" id="page-transfers">
        <div class="section-heading"><div><h2>Transferencias sugeridas</h2><p>Inventario total suficiente, pero mal ubicado para cubrir el lead time hacia Full.</p></div><div class="section-tools"><input class="filter-input" id="transferSearch" placeholder="Buscar SKU o canal"/><button class="card-action" onclick="exportCurrentTable('transfers')">Exportar CSV</button></div></div>
        <div class="status-strip" id="transferStatus"></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Unidades sugeridas por canal</div><div class="card-sub">Cantidad preliminar limitada por el inventario disponible en Odoo. La tabla inicia de mayor a menor por piezas a transferir.</div></div></div><div id="chartTransfers" class="chart"></div></div><div class="card"><div class="card-head"><div><div class="card-title">Justificación de la decisión</div><div class="card-sub">Por qué el sistema recomienda transferir antes de comprar.</div></div></div><div class="insights" id="transferInsights"></div></div></div>
        <div class="card"><div id="transferTable"></div></div>
      </section>

      <section class="page" id="page-purchases">
        <div class="section-heading"><div><h2>Compras sugeridas</h2><p>La decisión de compra se consolida por IQ para evitar comprar de más cuando el mismo producto tiene inventario en otra bolsa o canal.</p></div><div class="section-tools"><input class="filter-input" id="purchaseSearch" placeholder="Buscar IQ, producto o canal"/><button class="card-action" onclick="exportCurrentTable('purchasesConsolidated')">Exportar consolidado CSV</button></div></div>
        <div class="status-strip" id="purchaseStatus"></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Compra consolidada por SKU</div><div class="card-sub">Top de piezas a ordenar considerando todo el inventario de la empresa para cada IQ.</div></div></div><div id="chartPurchases" class="chart"></div></div><div class="card"><div class="card-head"><div><div class="card-title">Criterios de compra</div><div class="card-sub">La recomendación revisa primero Odoo, tránsito y Full en conjunto.</div></div></div><div class="insights" id="purchaseInsights"></div></div></div>
        <div class="card" style="margin-top:18px"><div class="card-head"><div><div class="card-title">Compra consolidada por SKU</div><div class="card-sub">Una sola recomendación por IQ. Los excedentes de un canal compensan faltantes de otro antes de pedir al proveedor. Incluye utilidad y ROI reales de 30 y 90 días (base_utilidad_roi.xlsx) para priorizar qué comprar primero.</div></div></div><div id="purchaseConsolidatedTable"></div></div>
        <div class="card" style="margin-top:18px"><div class="card-head"><div><div class="card-title">Detalle por canal</div><div class="card-sub">Vista diagnóstica para entender dónde se origina la necesidad. Se ordena por piezas requeridas de mayor a menor; puedes cambiar el orden y los filtros en la propia tabla.</div></div><div class="section-tools"><button class="card-action" onclick="exportCurrentTable('purchases')">Exportar detalle CSV</button></div></div><div id="purchaseTable"></div></div>
      </section>

      <section class="page" id="page-lots">
        <div class="section-heading"><div><h2>Lotes y antigüedad</h2><p>Edad del inventario y trazabilidad de arribos registrados en Odoo.</p></div></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Inventario por antigüedad</div><div class="card-sub" id="chartAgeSub">Unidades con stock actual, agrupadas por antigüedad.</div></div></div><div id="chartAge" class="chart"></div></div><div class="card"><div class="card-head"><div><div class="card-title">SKU-canal con inventario +90 días</div><div class="card-sub">Ordenados por inversión afectada: prioridad para promoción, reasignación o liquidación.</div></div><button class="card-action" onclick="exportCurrentTable('oldProducts')">Exportar CSV</button></div><div id="oldProductsTable"></div></div></div>
        <div class="card"><div class="card-head"><div><div class="card-title">Historial de arribos</div><div class="card-sub">Órdenes de compra, cantidades y monto registrado.</div></div><div class="section-tools"><input class="filter-input" id="arrivalSearch" placeholder="Buscar orden, proveedor o SKU"/><button class="card-action" onclick="exportCurrentTable('arrivals')">Exportar CSV</button></div></div><div id="arrivalsTable"></div></div>
      </section>

      <section class="page" id="page-audit">
        <div class="section-heading"><div><h2>Movimientos Odoo</h2><p>Registro acumulativo de traslados entre canales, ajustes, salidas a ubicaciones no comerciales y envíos a Full. Sirve para detectar movimientos hechos para bajar la antigüedad de un canal.</p></div><div class="section-tools"><input class="filter-input" id="auditSearch" placeholder="Buscar SKU, usuario o referencia"/><select class="select" id="auditChannel"><option value="">Todos los canales</option></select><select class="select" id="auditLevel"><option value="review">En revisión (alta y media)</option><option value="ALTA">Solo revisión alta</option><option value="">Todos</option></select><button class="card-action" onclick="exportCurrentTable('auditMoves')">Exportar CSV</button></div></div>
        <div class="status-strip wide" id="auditStatus"></div>
        <div class="grid-2">
          <div class="card"><div class="card-head"><div><div class="card-title">Movimientos por canal afectado</div><div class="card-sub">Piezas movidas en los filtros actuales, separadas por nivel de revisión.</div></div></div><div id="chartAudit" class="chart small"></div></div>
          <div class="card"><div class="card-head"><div><div class="card-title">Por usuario</div><div class="card-sub">Quién validó los movimientos en los filtros actuales.</div></div><button class="card-action" onclick="exportCurrentTable('auditUsers')">Exportar CSV</button></div><div id="auditUsersTable"></div></div>
        </div>
        <div class="card"><div class="card-head"><div><div class="card-title">Detalle de movimientos</div><div class="card-sub" id="auditRulesNote"></div></div></div><div id="auditMovesTable"></div></div>
      </section>

      <section class="page" id="page-alerts">
        <div class="section-heading"><div><h2>Centro de alertas</h2><p>Abastecimiento, sobrestock, demanda intermitente y excepciones.</p></div><div class="section-tools"><div class="view-toggle" id="alertsViewToggle"><button data-view="simple" class="active">Vista simple</button><button data-view="detail">Vista detallada</button></div><select class="select" id="alertRiskFilter"><option value="">Todos los riesgos</option><option>CRÍTICO</option><option>ALTO</option><option>INTERMITENTE</option><option>SOBRESTOCK</option></select><input class="filter-input" id="alertSearch" placeholder="Buscar SKU, producto o canal"/></div></div>
        <div class="status-strip" id="alertStatus"></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Alertas por canal</div><div class="card-sub">Cantidad de SKUs críticos, altos e intermitentes.</div></div></div><div id="chartAlertsChannel" class="chart"></div></div><div class="card"><div class="card-head"><div><div class="card-title">Distribución de riesgo</div><div class="card-sub">Clasificación global de los SKU-canal.</div></div></div><div id="chartRiskDistribution" class="chart"></div></div></div>
        <div class="card"><div id="alertsTable"></div></div>
      </section>

      <section class="page" id="page-quality">
        <div class="section-heading"><div><h2>Calidad de datos</h2><p>Transparencia sobre homologación, movimientos excluidos y consistencia del consolidado.</p></div></div>
        <div class="quality-grid" id="qualityKpis"></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Control detalle vs. total</div><div class="card-sub">La diferencia esperada es cero.</div></div></div><div id="controlTable"></div></div><div class="card"><div class="card-head"><div><div class="card-title">Traslados excluidos</div><div class="card-sub">Operaciones que no cumplieron contacto, origen, estado o fecha.</div></div></div><div id="excludedTransfersTable"></div></div></div>
        <div class="grid-2"><div class="card"><div class="card-head"><div><div class="card-title">Stock sin homologar</div><div class="card-sub">Referencias que todavía no se relacionan con un SKU madre.</div></div></div><div id="stockUnlinkedTable"></div></div><div class="card"><div class="card-head"><div><div class="card-title">Ventas no vinculadas</div><div class="card-sub">Registros pendientes de mapeo o validación.</div></div></div><div id="salesUnlinkedTable"></div></div></div>
        <div class="card"><div class="card-head"><div><div class="card-title">SKU de AutoAzur sin IQ</div><div class="card-sub">SKU de publicaciones que no se pudieron ligar a un IQ (ni por diccionario, ni por código de barras de Odoo, ni por patrón IQ). Agrégalos al diccionario como alias del IQ correcto; primero los que tienen stock publicado.</div></div><button class="card-action" onclick="exportCurrentTable('listingsUnlinked')">Exportar CSV</button></div><div id="listingsUnlinkedTable"></div></div>
      </section>

      <section class="page" id="page-methodology">
        <div class="section-heading"><div><h2>Metodología y reglas</h2><p>Definiciones utilizadas para construir la rotación individual y el consolidado.</p></div></div>
        <div class="method-grid">
          <div class="method-card"><div class="method-number">1</div><h3>Tres inventarios por canal</h3><p>Odoo asignado al canal, inventario en tránsito hacia Full e inventario recibido en Full. Las bolsas son mutuamente excluyentes.</p><div class="formula">Inventario total = Odoo + Tránsito + Full</div></div>
          <div class="method-card"><div class="method-number">2</div><h3>Traslados a Full</h3><p>Estado Listo se clasifica como tránsito; Hecho como recibido; Cancelado vale cero. El stock de Odoo se refresca desde la API.</p><div class="formula">Listo → Tránsito · Hecho → Full · Cancelado → 0</div></div>
          <div class="method-card"><div class="method-number">3</div><h3>Ventas por origen</h3><p>Odoo identifica Full o Drop y el equipo de ventas identifica el canal. AutoAzur complementa pedidos no encontrados en Odoo.</p><div class="formula">Venta total canal = Full + Drop + Pendiente Odoo</div></div>
          <div class="method-card"><div class="method-number">4</div><h3>Cobertura inmediata</h3><p>Mide cuántos días puede sostenerse el canal únicamente con su inventario Full frente a ventas Full.</p><div class="formula">Cobertura Full = Inventario Full / Venta diaria Full</div></div>
          <div class="method-card"><div class="method-number">5</div><h3>Cobertura total</h3><p>Evalúa si el conjunto de inventarios puede sostener la demanda completa del canal antes de considerar una compra.</p><div class="formula">Cobertura total = Inventario total / Venta diaria total</div></div>
          <div class="method-card"><div class="method-number">6</div><h3>Transferencia sugerida</h3><p>Se proyecta el consumo durante __FULL_TRANSIT_DAYS__ días de traslado y se completa el inventario para conservar __FULL_TARGET_DAYS__ días de cobertura cuando las piezas lleguen.</p><div class="formula">min(Odoo, demanda × __FULL_TARGET_DAYS__ − max(Full − demanda × __FULL_TRANSIT_DAYS__, 0) − Tránsito)</div></div>
          <div class="method-card"><div class="method-number">7</div><h3>Compra sugerida</h3><p>Solo se propone después de revisar las tres bolsas. El objetivo actual es 45 días y el proveedor tiene 30 días de lead time.</p><div class="formula">max(demanda diaria × 45 días − inventario total, 0)</div></div>
          <div class="method-card"><div class="method-number">8</div><h3>Demanda intermitente</h3><p>Los productos con pocas ventas o pocos días de movimiento no se juzgan únicamente con promedios diarios. Requieren revisión manual.</p><div class="formula">≤12 unidades o ≤8 días con venta en 90 días</div></div>
          <div class="method-card"><div class="method-number">9</div><h3>Consolidación</h3><p>El total se construye después del detalle individual. El control compara la suma por canal contra la métrica consolidada.</p><div class="formula">Σ detalle por canal = total consolidado</div></div><div class="method-card"><div class="method-number">10</div><h3>Utilidad y ROI por canal</h3><p>Las columnas <b>Utilidad 30d/90d</b> y <b>ROI 30d/90d</b> de las tablas de cada canal miden solo las ventas de ESE canal, no el consolidado del producto. Vienen de la hoja <b>por_producto_canal</b> de base_utilidad_roi.xlsx. Las ventanas se cuentan contra la venta más reciente observada en los reportes, no contra la fecha de hoy, porque los canales rara vez entregan su información al día. Un guion (—) significa que no hubo ventas con fecha reconocible en esa ventana: no es un ROI de 0%.</p><div class="formula">ROI del periodo = Σ utilidad / Σ base &nbsp; (no es el promedio de los ROI diarios)</div></div><div class="method-card"><div class="method-number">11</div><h3>Gráfica de ventas por día</h3><p>Unidades y monto salen de la base operativa del script 02; utilidad y ROI, de los reportes de utilidad por canal. El eje X recorre día por día, incluidos los días sin venta, para que un hueco se vea como hueco y no como una barra ausente. En la métrica de ROI los días sin venta se dibujan como corte de la línea, no como 0%. El rango de fechas está acotado al periodo realmente disponible en la base.</p><div class="formula">ROI del rango = Σ utilidad / Σ base &nbsp;·&nbsp; Media móvil = promedio de los 7 días previos</div></div>
          <div class="method-card"><div class="method-number">12</div><h3>Redistribución entre canales</h3><p>Compara demanda reciente e histórica, proyecta cinco días de consumo y protege 30 días en el origen. Solo utiliza como transferible directo el stock Odoo asignado.</p><div class="formula">Demanda = 50% ritmo 10d + 30% ritmo 30d + 20% ritmo 90d</div></div>
          <div class="method-card"><div class="method-number">13</div><h3>Portafolio B2B</h3><p>Marketplace conserva primero 30 días. El excedente se valúa con el ROI semanal y el costo unitario Odoo. Las utilidades futuras se descuentan al 0.0355% diario.</p><div class="formula">ROI B2B equivalente hoy(q) = VPN de los ingresos Marketplace desplazados / costo de q piezas − 1</div></div>
          <div class="method-card"><div class="method-number">14</div><h3>Antigüedad por capas</h3><p>Cada pieza del canal tiene una fecha de referencia. En Odoo, la compra del lote (una devolución o un traslado interno no la reinician). En Full, el envío a Full: lo que queda en Full son las piezas de los envíos más recientes. Sin lote o sin envío registrado, se usa la antigüedad de las compras del SKU. La alerta se activa si alguna pieza supera el umbral.</p><div class="formula">Piezas +90 = Σ piezas de capas con días &gt; 90 · Días promedio = Σ(piezas × días) / Σ piezas</div></div>
          <div class="method-card"><div class="method-number">15</div><h3>Inversión</h3><p>Costo unitario por IQ: valor contable de Odoo en CUATI; si no hay piezas en CUATI, costo promedio o estándar del producto; si falta, precio de la última recepción de compra. Las tablas de antigüedad inician por la inversión más alta.</p><div class="formula">Inversión = piezas × costo unitario · Inversión +90 = piezas +90 × costo</div></div>
          <div class="method-card"><div class="method-number">16</div><h3>Semáforo de cobertura</h3><p>Ordena los SKU por urgencia para el KAM: sin stock con ventas, quiebre en menos de 7 días, Full que se agota antes de que llegue un envío, cobertura baja, sana, alta, exceso y stock sin ventas. Los umbrales se ajustan en .env.</p><div class="formula">Días de cobertura = inventario ÷ venta diaria de 90 días</div></div>
          <div class="method-card"><div class="method-number">17</div><h3>Registro de movimientos Odoo</h3><p>Cada corrida del 01 agrega a registro_movimientos_odoo.csv los traslados entre canales, ajustes, salidas a ubicaciones no comerciales, devoluciones y envíos a Full, con usuario, piezas, valor y edad del lote al moverse. El registro no se borra aunque el movimiento se revierta después.</p><div class="formula">Revisión alta = edad ≥ umbral al mover · ida y vuelta · ajuste de salida y reingreso</div></div>
          <div class="method-card"><div class="method-number">18</div><h3>Stock sin publicación activa</h3><p>El 01 descarga de AutoAzur (/item/listings) todas las publicaciones de cada canal y las liga al IQ por su SKU o ItemID con el diccionario. Un SKU con piezas en el canal se reporta si no tiene publicación, o si ninguna está activa (pausada, en revisión, con error, inactiva o cerrada). Si tiene varias, se toma la mejor. Si la consulta de un canal falla, ese canal no se evalúa.</p><div class="formula">Pendiente = stock en canal &gt; 0 y ninguna publicación Activa</div></div>
          <div class="method-card"><div class="method-number">19</div><h3>Repartición de stock central</h3><p>La cantidad sale del stock disponible en CUATI/Existencias y CUATI/B2B. La señal usa 50% del ritmo de 10 días, 30% del de 30 y 20% del de 90. Si el IQ no tiene muestra suficiente, baja a categoría + marca, categoría, marca y finalmente reparto parejo. Con datos, 20% del peso queda como base igual para mantener exploración y 80% sigue desempeño; el stock que ya tiene cada canal reduce lo que recibe.</p><div class="formula">Participación objetivo = 20% × reparto igual + 80% × participación por demanda</div></div>
        </div>
      </section>
    </div>
  </main>
</div>

<div class="admin-modal-backdrop" id="adminModalBackdrop" aria-hidden="true">
  <div class="admin-modal" role="dialog" aria-modal="true" aria-labelledby="adminModalTitle">
    <div style="display:flex;justify-content:space-between;gap:14px;align-items:flex-start">
      <div>
        <div class="page-eyebrow">Acción protegida</div>
        <h3 id="adminModalTitle">Administrador Odoo</h3>
        <p id="adminModalSubtitle">Inicia sesión para ejecutar sugerencias desde el dashboard.</p>
      </div>
      <button class="odoo-close" type="button" onclick="closeAdminModal()">×</button>
    </div>

    <div id="adminLoginStep">
      <input class="admin-password" id="adminPassword" type="password" autocomplete="current-password" placeholder="Contraseña de administrador"/>
      <div class="admin-message" id="adminLoginMessage"></div>
      <div class="admin-actions">
        <button class="admin-secondary" type="button" onclick="closeAdminModal()">Cancelar</button>
        <button class="admin-primary" id="adminLoginBtn" type="button" onclick="submitAdminLogin()">Entrar</button>
      </div>
    </div>

    <div id="adminTransferStep" style="display:none">
      <div class="admin-modal-grid">
        <div class="admin-fact"><span>SKU</span><b id="adminTransferSku">—</b></div>
        <div class="admin-fact"><span>Canal</span><b id="adminTransferChannel">—</b></div>
        <div class="admin-fact"><span>Cantidad</span><b id="adminTransferQty">—</b></div>
        <div class="admin-fact"><span>Resultado</span><b>Marcar HECHO</b></div>
      </div>
      <p>El servidor volverá a validar la sugerencia, el stock disponible, la ubicación, el lote y la plantilla de transferencia antes de escribir en Odoo.</p>
      <div class="admin-danger-note">Esta acción modifica el inventario real de Odoo. Úsala únicamente cuando el movimiento físico sí corresponda a la transferencia.</div>
      <div class="admin-message" id="adminTransferMessage"></div>
      <div class="admin-actions">
        <button class="admin-secondary" type="button" onclick="adminLogout()">Cerrar sesión</button>
        <button class="admin-secondary" type="button" onclick="closeAdminModal()">Cancelar</button>
        <button class="admin-primary" id="adminExecuteBtn" type="button" onclick="executeAdminTransfer()">Mover y marcar HECHO</button>
      </div>
    </div>
  </div>
</div>

<div class="toast" id="toast"></div>

<script>
const DATA = __DATA_JSON__;
const COLORS = {blue:'#2563eb',teal:'#0f9f9a',purple:'#7559d9',amber:'#f59e0b',red:'#e5484d',green:'#1f9d66',slate:'#64748b',odoo:'#7559d9',transit:'#f59e0b',full:'#0f9f9a'};
const CHANNEL_COLORS = {'Amazon':'#ff9900','Mercado Libre':'#3483fa','Walmart':'#0071ce','Liverpool':'#d51b5e','Coppel':'#f5c400','Elektra':'#d9252a','TikTok':'#111827','General':'#7559d9'};
const state = {page:'executive',channel:DATA.meta.channels[0]||'Amazon',tables:{},views:{channel:'simple',alerts:'simple'},lastEvaluation:null,selectedSalesSku:'',selectedB2BSku:'',selectedDistributionSku:'',pageQueries:{},tableQueries:{},tableSorts:{},tableFilters:{},tableNumFilters:{},channelSales:{metric:'unidades',sku:'',range:'90',from:'',to:''}};
const adminState = {
  token: sessionStorage.getItem('iqtech_admin_token') || '',
  pending: null,
  completed: new Set(),
};
const fmt = new Intl.NumberFormat('es-MX',{maximumFractionDigits:0});
const fmt1 = new Intl.NumberFormat('es-MX',{maximumFractionDigits:1});
const money = new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0});
const money2 = new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',minimumFractionDigits:2,maximumFractionDigits:2});
const pct2 = v => `${(n(v)*100).toFixed(2)}%`;
const clamp=(v,min,max)=>Math.max(min,Math.min(max,v));
const n=v=>{const x=Number(v);return Number.isFinite(x)?x:0};
const text=v=>(v===null||v===undefined||v==='')?'—':String(v);
const escapeHtml=v=>String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
const cap=s=>String(s||'').toLowerCase().replace(/(^|\s)\S/g,c=>c.toUpperCase());
function coverage(v){return Number.isFinite(Number(v))?`${fmt1.format(Number(v))} días`:'Sin venta';}
function daysAge(v){return Number.isFinite(Number(v))?`${fmt1.format(Number(v))} días`:'Sin dato';}
function badge(value){const s=String(value||'').toUpperCase();let cls='gray';if(s.includes('CRÍTICO')||s.includes('COMPRA')||s==='BAJA')cls='red';else if(s.includes('MANUAL')||s.includes('PENDIENTE'))cls='purple';else if(s.includes('ALTO')||s.includes('TRANSFERIR')||s==='MEDIA'||s.includes('REVISIÓN'))cls='amber';else if(s.includes('SANO')||s.includes('SIN ACCIÓN')||s==='NO'||s==='ALTA'||s==='APROBAR')cls='green';else if(s.includes('INTERMITENTE'))cls='purple';else if(s.includes('SOBRESTOCK'))cls='blue';else if(s.includes('+90'))cls='red';else if(s.includes('+60')||s.includes('+30'))cls='amber';else if(s.includes('RECIENTE'))cls='green';else if(s.includes('LOTE_PARCIAL')||s.includes('APROX')||s.includes('LOTE_AJUSTE'))cls='amber';else if(s==='LOTE'||s==='ENVIO_FULL')cls='green';return `<span class="badge ${cls}">${escapeHtml(value||'Sin clasificar')}</span>`;}
function compactAction(v){const s=String(v||'').toUpperCase();if(s.includes('TRANSFERIR'))return 'Transferir';if(s.includes('COMPRA'))return 'Comprar';if(s.includes('INTERMITENTE'))return 'Revisión manual';return 'Sin acción';}
function showToast(msg){const el=document.getElementById('toast');el.textContent=msg;el.classList.add('show');setTimeout(()=>el.classList.remove('show'),2200);}


// ============================================================
// ODOO · BORRADORES REVISADOS
// ============================================================
const odooDraftState={busy:new Set(),completed:new Set(JSON.parse(sessionStorage.getItem('iqtech_odoo_drafts_vnext')||'[]'))};
function persistOdooDrafts(){sessionStorage.setItem('iqtech_odoo_drafts_vnext',JSON.stringify([...odooDraftState.completed]));}
function redistReqId(r){return `${DATA.meta.generated_at}|REDIST|${r.sku_madre}|${r.canal_origen}|${r.canal_destino}|${Math.floor(n(r.cantidad_sugerida))}`;}
function redistKey(r){return `R|${redistReqId(r)}`;}
function redistExecutable(r){const t=String(r?.tipo_stock_origen||'').toUpperCase(),m=String(r?.check_manual||'').toUpperCase();return t.startsWith('ODOO')&&!t.includes('FULL')&&!m.startsWith('SI')&&n(r?.cantidad_sugerida)>0;}
async function apiJson(path,payload){const res=await fetch(path,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});let body={};try{body=await res.json();}catch(_){body={mensaje:`HTTP ${res.status}`};}if(!res.ok)throw new Error(body.mensaje||body.message||`HTTP ${res.status}`);return body;}
function selectedSuggestion(){const i=Number(state.pendingSuggestionIndex);return Number.isInteger(i)&&i>=0?(DATA.redistributions||[])[i]:null;}
async function createEvaluatedRedistributionDraft(){
  const r=selectedSuggestion();if(!r||!redistExecutable(r)){showToast('Esta propuesta requiere revisión manual');return;}
  const key=redistKey(r);if(odooDraftState.completed.has(key)){showToast('Este borrador ya fue creado en esta sesión');return;}if(odooDraftState.busy.has(key))return;
  if(!confirm(`Crear BORRADOR en Odoo?\n\n${r.sku_madre}\n${r.canal_origen} → ${r.canal_destino}\n${fmt.format(r.cantidad_sugerida)} piezas\n\nEl servidor volverá a validar producto, lotes y stock vivo.`))return;
  odooDraftState.busy.add(key);renderEvaluation(state.lastEvaluation);
  try{
    const body=await apiJson('/api/odoo/redistribution/draft',{generated_at:DATA.meta.generated_at,sku_madre:r.sku_madre,canal_origen:r.canal_origen,canal_destino:r.canal_destino,cantidad_sugerida:r.cantidad_sugerida,request_id:redistReqId(r)});
    odooDraftState.completed.add(key);persistOdooDrafts();
    const lots=body?.resolucion_producto?.lotes_disponibles?.length||0;
    alert(`Borrador creado en Odoo: ${body?.picking?.name||'—'}\n${r.sku_madre} · ${r.canal_origen} → ${r.canal_destino} · ${fmt.format(r.cantidad_sugerida)} pzs\nLotes detectados en origen: ${lots}`);
  }catch(err){alert(`No se pudo crear el borrador en Odoo:\n\n${err.message}`);}finally{odooDraftState.busy.delete(key);renderEvaluation(state.lastEvaluation);}
}
function distributionReqId(sku,origin,dest,qty){return `${DATA.meta.generated_at}|DIST|${String(sku).toUpperCase()}|${origin}|${dest}|${Math.floor(n(qty))}`;}
function distributionKey(sku,origin,dest,qty){return `D|${distributionReqId(sku,origin,dest,qty)}`;}
async function createDistributionSourceDraft(sku,dest,origin,qty){
  qty=Math.floor(n(qty));if(qty<=0)return null;const key=distributionKey(sku,origin,dest,qty);if(odooDraftState.completed.has(key))return {duplicate:true,local:true};
  const body=await apiJson('/api/odoo/distribution/draft',{generated_at:DATA.meta.generated_at,sku_madre:sku,canal_origen:origin,canal_destino:dest,cantidad_sugerida:qty,request_id:distributionReqId(sku,origin,dest,qty)});
  odooDraftState.completed.add(key);persistOdooDrafts();return body;
}
async function createDistributionRowDrafts(sku,canal){
  const row=distributionDetailRows(sku).find(r=>String(r.canal||'')===String(canal||''));if(!row)return;
  const parts=[];if(n(row.desde_existencias)>0)parts.push(`CUATI/Existencias: ${fmt.format(row.desde_existencias)}`);if(n(row.desde_b2b)>0)parts.push(`CUATI/B2B: ${fmt.format(row.desde_b2b)}`);
  if(!parts.length){showToast('Este canal no tiene piezas sugeridas');return;}
  if(!confirm(`Crear borrador(es) Odoo para ${sku} → ${canal}?\n\n${parts.join('\n')}\n\nSe volverá a validar el stock real antes de crear.`))return;
  try{
    const created=[];
    if(n(row.desde_existencias)>0){const x=await createDistributionSourceDraft(sku,canal,'General',row.desde_existencias);if(x?.picking?.name)created.push(x.picking.name);}
    if(n(row.desde_b2b)>0){const x=await createDistributionSourceDraft(sku,canal,'B2B',row.desde_b2b);if(x?.picking?.name)created.push(x.picking.name);}
    alert(`Borrador(es) listos para ${sku} → ${canal}.\n${created.length?created.join('\n'):'Ya existían en esta sesión.'}`);renderDistributionDetail();
  }catch(err){alert(`No se pudo crear la repartición en Odoo:\n\n${err.message}`);}
}
async function createAllDistributionDrafts(){
  const sku=state.selectedDistributionSku;if(!sku)return;const rows=distributionDetailRows(sku).filter(r=>n(r.cantidad_sugerida)>0);if(!rows.length)return;
  const total=rows.reduce((a,r)=>a+n(r.cantidad_sugerida),0);if(!confirm(`Crear todos los borradores válidos de ${sku}?\n\n${rows.length} destinos · ${fmt.format(total)} piezas.\nCada fuente se validará por separado contra Odoo.`))return;
  const ok=[],fail=[];
  for(const r of rows){try{if(n(r.desde_existencias)>0){const x=await createDistributionSourceDraft(sku,r.canal,'General',r.desde_existencias);if(x?.picking?.name)ok.push(x.picking.name);}if(n(r.desde_b2b)>0){const x=await createDistributionSourceDraft(sku,r.canal,'B2B',r.desde_b2b);if(x?.picking?.name)ok.push(x.picking.name);}}catch(err){fail.push(`${r.canal}: ${err.message}`);}}
  alert(`Proceso terminado.\nBorradores nuevos: ${ok.length}${fail.length?`\n\nRevisar:\n${fail.join('\n')}`:''}`);renderDistributionDetail();
}

function adminRequestId(p){
  return `${DATA.meta.generated_at}|${p.sku_madre}|${p.canal}|${Math.floor(n(p.cantidad))}`;
}
function updateAdminStatus(){
  const btn=document.getElementById('adminStatusBtn');
  if(!btn)return;
  if(adminState.token){btn.classList.add('on');btn.textContent='✓ Admin';}
  else{btn.classList.remove('on');btn.textContent='🔒 Admin';}
}
function openAdminModal(){
  const modal=document.getElementById('adminModalBackdrop');
  if(!modal)return;
  modal.classList.add('open');
  modal.setAttribute('aria-hidden','false');
}
function closeAdminModal(){
  const modal=document.getElementById('adminModalBackdrop');
  if(!modal)return;
  modal.classList.remove('open');
  modal.setAttribute('aria-hidden','true');
  const msg=document.getElementById('adminLoginMessage');if(msg){msg.textContent='';msg.className='admin-message';}
  const tmsg=document.getElementById('adminTransferMessage');if(tmsg){tmsg.textContent='';tmsg.className='admin-message';}
}
function showAdminLogin(){
  document.getElementById('adminLoginStep').style.display='';
  document.getElementById('adminTransferStep').style.display='none';
  document.getElementById('adminModalSubtitle').textContent='Inicia sesión para ejecutar sugerencias desde el dashboard.';
  openAdminModal();
  setTimeout(()=>document.getElementById('adminPassword')?.focus(),50);
}
function showAdminHome(){
  adminState.pending=null;
  document.getElementById('adminLoginStep').style.display='none';
  document.getElementById('adminTransferStep').style.display='';
  document.getElementById('adminModalSubtitle').textContent='Sesión administrativa activa.';
  document.getElementById('adminTransferSku').textContent='—';
  document.getElementById('adminTransferChannel').textContent='—';
  document.getElementById('adminTransferQty').textContent='—';
  document.getElementById('adminExecuteBtn').style.display='none';
  openAdminModal();
}
function showAdminTransfer(){
  const p=adminState.pending;
  if(!p){showAdminHome();return;}
  document.getElementById('adminLoginStep').style.display='none';
  document.getElementById('adminTransferStep').style.display='';
  document.getElementById('adminExecuteBtn').style.display='';
  document.getElementById('adminModalSubtitle').textContent='Confirma la sugerencia antes de modificar Odoo.';
  document.getElementById('adminTransferSku').textContent=p.sku_madre;
  document.getElementById('adminTransferChannel').textContent=p.canal;
  document.getElementById('adminTransferQty').textContent=`${fmt.format(p.cantidad)} pzs`;
  openAdminModal();
}
async function adminFetch(path,options={}){
  const headers={'Content-Type':'application/json',...(options.headers||{})};
  if(adminState.token)headers.Authorization=`Bearer ${adminState.token}`;
  const res=await fetch(path,{...options,headers});
  let body={};
  try{body=await res.json();}catch(_){body={message:`HTTP ${res.status}`};}
  if(res.status===401){
    adminState.token='';
    sessionStorage.removeItem('iqtech_admin_token');
    updateAdminStatus();
  }
  return {res,body};
}
async function submitAdminLogin(){
  const password=document.getElementById('adminPassword')?.value||'';
  const msg=document.getElementById('adminLoginMessage');
  const btn=document.getElementById('adminLoginBtn');
  if(!password){msg.textContent='Escribe la contraseña.';msg.className='admin-message error';return;}
  btn.disabled=true;msg.textContent='Validando…';msg.className='admin-message';
  try{
    const {res,body}=await adminFetch('/api/admin/login',{method:'POST',body:JSON.stringify({password})});
    if(!res.ok)throw new Error(body.message||'No se pudo iniciar sesión.');
    adminState.token=body.token||'';
    sessionStorage.setItem('iqtech_admin_token',adminState.token);
    document.getElementById('adminPassword').value='';
    updateAdminStatus();
    msg.textContent='';
    if(adminState.pending)showAdminTransfer();else showAdminHome();
  }catch(err){
    msg.textContent=err.message||'Contraseña incorrecta.';
    msg.className='admin-message error';
  }finally{btn.disabled=false;}
}
async function adminLogout(){
  try{if(adminState.token)await adminFetch('/api/admin/logout',{method:'POST',body:'{}'});}catch(_){}
  adminState.token='';
  adminState.pending=null;
  sessionStorage.removeItem('iqtech_admin_token');
  updateAdminStatus();
  closeAdminModal();
  showToast('Sesión de administrador cerrada');
}
function openAdminTransfer(sku,canal,cantidad){
  const q=Math.floor(n(cantidad));
  if(q<=0)return;
  const key=`${sku}|${canal}|${q}`;
  if(adminState.completed.has(key)){showToast('Esta sugerencia ya se ejecutó en esta sesión');return;}
  adminState.pending={sku_madre:String(sku||'').toUpperCase(),canal:String(canal||''),cantidad:q};
  if(adminState.token)showAdminTransfer();else showAdminLogin();
}
async function executeAdminTransfer(){
  const p=adminState.pending;
  if(!p)return;
  const btn=document.getElementById('adminExecuteBtn');
  const msg=document.getElementById('adminTransferMessage');
  btn.disabled=true;msg.textContent='Validando sugerencia y ejecutando en Odoo…';msg.className='admin-message';
  try{
    const payload={
      ...p,
      generated_at:DATA.meta.generated_at,
      request_id:adminRequestId(p),
      modo:'done',
    };
    const {res,body}=await adminFetch('/api/transfer/full',{method:'POST',body:JSON.stringify(payload)});
    if(!res.ok)throw new Error(body.message||'No se pudo completar la transferencia.');
    if(body.state!=='done')throw new Error(body.message||`Odoo dejó la transferencia en ${body.state||'estado desconocido'}.`);
    adminState.completed.add(`${p.sku_madre}|${p.canal}|${p.cantidad}`);
    msg.textContent=`Hecho: ${body.picking_name||'transferencia'} · ${fmt.format(p.cantidad)} pzs.`;
    msg.className='admin-message ok';
    showToast(`Odoo: ${body.picking_name||'transferencia'} marcada como HECHO`);
    setTimeout(()=>{closeAdminModal();renderPage(state.page);},1400);
  }catch(err){
    msg.textContent=err.message||'Error al ejecutar la transferencia.';
    msg.className='admin-message error';
    if(!adminState.token)setTimeout(showAdminLogin,600);
  }finally{btn.disabled=false;}
}
function plot(id,traces,layout={}){const dark=document.body.classList.contains('dark');const base={margin:{l:50,r:20,t:20,b:48},paper_bgcolor:'rgba(0,0,0,0)',plot_bgcolor:'rgba(0,0,0,0)',font:{family:'Inter,Segoe UI,sans-serif',size:11,color:dark?'#cbd5e1':'#5f6b7c'},xaxis:{gridcolor:dark?'#243147':'#edf0f5',zeroline:false},yaxis:{gridcolor:dark?'#243147':'#edf0f5',zeroline:false},legend:{orientation:'h',y:-.18},hoverlabel:{bgcolor:dark?'#172033':'#fff',font:{color:dark?'#fff':'#172033'}},modebar:{bgcolor:'rgba(0,0,0,0)'}};Plotly.react(id,traces,{...base,...layout},{displaylogo:false,responsive:true,modeBarButtonsToRemove:['lasso2d','select2d']});}
function channelRows(channel){return DATA.rotation.filter(r=>r.canal===channel);}
function channelSummary(channel){return DATA.channel_summary.find(r=>r.canal===channel)||{};}
function channelUnlinkedStock(channel){return (DATA.stock_unlinked||[]).filter(r=>String(r.canales_afectados_texto||'').split('|').map(s=>s.trim()).includes(channel));}
function channelUnlinkedSales(channel){return (DATA.sales_unlinked||[]).filter(r=>String(r.canal_normalizado||'')===channel);}
function renderChannelDictionary(){
  const channel=state.channel,box=document.getElementById('channelDictionaryBody');
  if(!box)return;
  const stockRows=channelUnlinkedStock(channel),salesRows=channelUnlinkedSales(channel),total=stockRows.length+salesRows.length;
  document.getElementById('channelDictionarySub').textContent=`SKU de ${channel} que la sincronización no pudo vincular a un SKU madre (IQ).`;
  if(total===0){
    box.innerHTML=insight('¡Felicidades, diccionario al día!',`Todos los SKU de ${channel} están vinculados correctamente al SKU madre (IQ). No hay pendientes de stock ni de ventas por resolver en este canal.`,COLORS.green,'✓');
    return;
  }
  box.innerHTML=`<div class="status-strip">${statusCard(fmt.format(stockRows.length),'SKU de stock sin vincular',COLORS.amber,'◌')}${statusCard(fmt.format(salesRows.length),'Ventas sin vincular',COLORS.red,'!')}</div><div style="margin-top:14px"><div class="card-sub" style="margin-bottom:6px">Stock sin vincular</div><div id="channelUnlinkedStockTable"></div></div><div style="margin-top:18px"><div class="card-sub" style="margin-bottom:6px">Ventas sin vincular</div><div id="channelUnlinkedSalesTable"></div></div>`;
  renderTable('channelUnlinkedStockTable',stockRows,[
    {key:'sku_original',label:'SKU',format:'sku'},
    {key:'producto_madre',label:'Producto',format:'product'},
    {key:'stock_total',label:'Stock',format:'num'},
    {key:'canales_afectados_texto',label:'Canales con este stock',format:'badge'},
    {key:'motivo_no_vinculado',label:'Motivo',format:'badge'},
    {key:'accion_sugerida',label:'Acción',format:'product'}
  ],{pageSize:10,id:'channelUnlinkedStock'});
  renderTable('channelUnlinkedSalesTable',salesRows,[
    {key:'fecha',label:'Fecha'},
    {key:'pedido',label:'Pedido'},
    {key:'sku_original',label:'SKU',format:'sku'},
    {key:'producto',label:'Producto',format:'product'},
    {key:'cantidad',label:'Cantidad',format:'num'},
    {key:'motivo_no_vinculado',label:'Motivo',format:'badge'}
  ],{pageSize:10,id:'channelUnlinkedSales'});
}
function riskTone(r){return ({'CRÍTICO':COLORS.red,'ALTO':COLORS.amber,'INTERMITENTE':COLORS.purple,'SOBRESTOCK':COLORS.blue,'SANO':COLORS.green})[r]||COLORS.slate;}
function channelIcon(c){return ({'Amazon':'A','Mercado Libre':'ML','Walmart':'W','Liverpool':'L','Coppel':'C','Elektra':'E','TikTok':'T'})[c]||c.slice(0,2).toUpperCase();}
const CHANNEL_MARKS={
'Amazon':'<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M12 15.8c-3.6 0-6.5-1-8.9-2.8-.2-.1 0-.4.2-.3 2.6 1.5 5.7 2.4 8.9 2.4 2.2 0 4.6-.5 6.8-1.4.3-.1.6.2.3.4-2 1.5-4.5 2.3-7.3 2.3z" fill="#fff"/><path d="M12.8 12.6c-.3.3-.6.2-.9.1-.5-.4-.6-.6-.9-1-.9 1-1.6 1.3-2.7 1.3-1.4 0-2.5-.9-2.5-2.6 0-1.4.7-2.3 1.8-2.8 1-.4 2.3-.5 3.3-.6v-.2c0-.4 0-.9-.2-1.3-.2-.3-.6-.5-1-.5-.6 0-1.2.3-1.3 1-.1.2-.2.4-.4.4l-2.2-.2c-.2 0-.4-.2-.3-.5.5-2.4 2.6-3.2 4.5-3.2 1 0 2.3.3 3 1 1 .9.9 2.1.9 3.5v3.1c0 .9.4 1.3.7 1.8.1.2.1.4 0 .5-.4.3-1.1.9-1.5 1.2h0z" fill="#fff"/><circle cx="12" cy="12" r="11.5" stroke="#37475A" stroke-width="1"/></svg>',
'Mercado Libre':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#FFE600"/><path d="M6 13.5c0-3 2-5.5 5-5.5s5 2.5 5 5.5c0 .6-.1 1.1-.3 1.6.7-.1 1.4-.5 1.9-1 .2-.2.5 0 .4.3-.6 1.4-2.1 2.4-3.8 2.4-.6 0-1.1-.1-1.6-.3-.5.5-1.2.8-2 .8-2.5 0-4.6-1.8-4.6-3.8z" fill="#2D3277"/></svg>',
'Walmart':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#0071CE"/><g fill="#FFC220"><circle cx="12" cy="5.5" r="1.15"/><circle cx="12" cy="18.5" r="1.15"/><circle cx="5.5" cy="12" r="1.15"/><circle cx="18.5" cy="12" r="1.15"/><circle cx="7.6" cy="7.6" r="1.15"/><circle cx="16.4" cy="16.4" r="1.15"/><circle cx="7.6" cy="16.4" r="1.15"/><circle cx="16.4" cy="7.6" r="1.15"/></g><circle cx="12" cy="12" r="2.4" fill="#FFC220"/></svg>',
'Liverpool':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#E4032E"/><text x="12" y="16" font-family="Arial,sans-serif" font-size="13" font-weight="900" fill="#fff" text-anchor="middle">L</text></svg>',
'Coppel':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#E10600"/><text x="12" y="16" font-family="Arial,sans-serif" font-size="12" font-weight="900" fill="#FFE600" text-anchor="middle">CP</text></svg>',
'Elektra':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#FFDD00"/><text x="12" y="16" font-family="Arial,sans-serif" font-size="12" font-weight="900" fill="#D9252A" text-anchor="middle">EK</text></svg>',
'TikTok':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#010101"/><path d="M14.7 5.5c.4 1.4 1.5 2.5 2.9 2.8v2.1c-1.1 0-2.1-.3-3-.9v4.4c0 2.3-1.8 4.1-4.1 4.1S6.4 16.2 6.4 13.9c0-2.2 1.7-4 3.9-4.1v2.2c-1 .1-1.7.9-1.7 1.9 0 1.1.9 1.9 1.9 1.9 1.1 0 2-.9 2-2V5.5h2.2z" fill="#25F4EE"/><path d="M14.2 5.5c.4 1.4 1.5 2.5 2.9 2.8v2.1c-1.1 0-2.1-.3-3-.9v4.4c0 2.3-1.8 4.1-4.1 4.1S5.9 16.2 5.9 13.9c0-2.2 1.7-4 3.9-4.1v2.2c-1 .1-1.7.9-1.7 1.9 0 1.1.9 1.9 1.9 1.9 1.1 0 2-.9 2-2V5.5h2.2z" fill="#FE2C55" opacity=".75"/></svg>',
'General':'<svg viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg"><circle cx="12" cy="12" r="11.5" fill="#7559D9"/><text x="12" y="16" font-family="Arial,sans-serif" font-size="12" font-weight="900" fill="#fff" text-anchor="middle">IQ</text></svg>',
};
function channelMark(c,size=22){return CHANNEL_MARKS[c]?`<span style="display:inline-flex;width:${size}px;height:${size}px;flex:none">${CHANNEL_MARKS[c]}</span>`:`<span class="nav-icon">${channelIcon(c)}</span>`;}

function kpiCard(label,value,sub,accent,icon,progress=0,help='',onClick=''){const helpHtml=help?`<span class="kpi-help">?<span class="kpi-tooltip">${escapeHtml(help)}</span></span>`:'';const clickable=onClick?' clickable':'';const attrs=onClick?` role="button" tabindex="0" onclick="${onClick}" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${onClick}}"`:'';const link=onClick?`<div class="kpi-link">Ver detalle y sugerencias →</div>`:'';return `<div class="kpi${clickable}" style="--accent:${accent}"${attrs}><span class="kpi-accent"></span><div class="kpi-top"><div class="kpi-label" style="display:flex;align-items:center;gap:6px">${label}${helpHtml}</div><div class="kpi-icon">${icon}</div></div><div class="kpi-value">${value}</div><div class="kpi-sub">${sub}</div><div class="kpi-progress"><span style="width:${clamp(progress,0,100)}%"></span></div>${link}</div>`;}
function insight(title,body,tone=COLORS.blue,icon='→'){return `<div class="insight" style="--tone:${tone}"><div class="insight-icon">${icon}</div><div><div class="insight-title">${escapeHtml(title)}</div><div class="insight-text">${escapeHtml(body)}</div></div></div>`;}
function statusCard(value,label,tone,icon,opts={}){const click=opts.onClick?` clickable${opts.active?' active':''}`:'';const attrs=opts.onClick?` role="button" tabindex="0" onclick="${opts.onClick}" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();${opts.onClick}}"`:'';const hint=opts.hint?`<span class="status-hint">${escapeHtml(opts.hint)}</span>`:'';const dl=opts.download?`<button type="button" class="status-dl" title="Descargar CSV" onclick="event.stopPropagation();${opts.download}">CSV ↓</button>`:'';return `<div class="status-card${click}"${attrs}><div class="status-icon" style="background:${tone}18;color:${tone}">${icon}</div><div class="status-info"><b>${value}</b><span>${label}</span>${hint}</div>${dl}</div>`;}

function initNav(){const box=document.getElementById('channelNav');box.innerHTML=DATA.meta.channels.map(c=>`<button class="nav-item" data-page="channel" data-channel="${escapeHtml(c)}">${channelMark(c,22)}<span class="nav-label">${escapeHtml(c)}</span><span class="nav-badge">${fmt.format(channelRows(c).length)}</span></button>`).join('');document.querySelectorAll('.nav-item').forEach(btn=>btn.addEventListener('click',()=>{if(btn.dataset.channel)state.channel=btn.dataset.channel;goPage(btn.dataset.page);}));document.getElementById('suggestionBadge').textContent=fmt.format((DATA.redistributions||[]).length);document.getElementById('distributionBadge').textContent=fmt.format((DATA.distribution_summary||[]).length);document.getElementById('b2bBadge').textContent=fmt.format((DATA.b2b_portfolio||[]).filter(r=>n(r.b2b_disponible)>0).length);document.getElementById('transferBadge').textContent=fmt.format(DATA.transfers.length);document.getElementById('purchaseBadge').textContent=fmt.format((DATA.purchases_consolidated||[]).length);document.getElementById('auditBadge').textContent=fmt.format((DATA.audit_moves||[]).filter(m=>String(m.nivel_revision).toUpperCase()==='ALTA').length);document.getElementById('alertBadge').textContent=fmt.format(DATA.rotation.filter(r=>['CRÍTICO','ALTO'].includes(r.nivel_riesgo)).length);populateDecisionControls();populateB2BControls();}
function goPage(page){state.page=page;document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));document.getElementById(`page-${page}`).classList.add('active');document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.page===page && (!b.dataset.channel||b.dataset.channel===state.channel)));const titles={executive:['Control Tower','Resumen general'],channel:['Canales y KAMs',state.channel],suggestions:['Decisiones','Sugerencias y consulta KAM'],distribution:['Decisiones','Repartición de stock'],b2b:['Decisiones','B2B financiero'],transfers:['Decisiones','Odoo → Full'],purchases:['Decisiones','Compras sugeridas'],lots:['Análisis','Lotes y antigüedad'],audit:['Análisis','Movimientos Odoo'],alerts:['Análisis','Centro de alertas'],quality:['Gobierno de datos','Calidad de datos'],methodology:['Documentación','Metodología']};document.getElementById('pageEyebrow').textContent=titles[page][0];document.getElementById('pageTitle').textContent=titles[page][1];const global=document.getElementById('globalSearch');if(global)global.value=state.pageQueries[page]||'';renderPage(page);closeMobile();window.scrollTo({top:0,behavior:'smooth'});}
function renderPage(page){if(page==='executive')renderExecutive();if(page==='channel')renderChannel();if(page==='suggestions')renderSuggestions();if(page==='distribution')renderDistribution();if(page==='b2b')renderB2B();if(page==='transfers')renderTransfers();if(page==='purchases')renderPurchases();if(page==='lots')renderLots();if(page==='audit')renderAudit();if(page==='alerts')renderAlerts();if(page==='quality'){renderQuality();renderListingsQuality();}}

function renderSemaphore(){const rows=DATA.rotation;const critical=rows.filter(r=>r.nivel_riesgo==='CRÍTICO').length;const warning=rows.filter(r=>['ALTO','INTERMITENTE'].includes(r.nivel_riesgo)).length;const healthy=rows.filter(r=>['SANO','SOBRESTOCK'].includes(r.nivel_riesgo)).length;document.getElementById('semaphoreStrip').innerHTML=[
{cls:'red',num:critical,label:'SKU-canal críticos · requieren compra ya',filter:'CRÍTICO'},
{cls:'amber',num:warning,label:'En riesgo alto o demanda intermitente · revisar esta semana',filter:'ALTO'},
{cls:'green',num:healthy,label:'Sanos o con sobrestock · sin acción urgente',filter:''},
].map(c=>`<div class="semaphore-card ${c.cls}" onclick="goToAlerts('${c.filter}')"><span class="semaphore-dot"></span><div><div class="semaphore-num">${fmt.format(c.num)}</div><div class="semaphore-label">${c.label}</div></div></div>`).join('');
const critRows=rows.filter(r=>r.nivel_riesgo==='CRÍTICO');const topTransfer=DATA.transfers[0],topPurchase=(DATA.purchases_consolidated||[])[0];
const transferTotal=safeSum(DATA.transfers,'transferencia_sugerida'),purchaseTotal=safeSum(DATA.purchases_consolidated||[],'compra_sugerida');
let bannerText;
if(critRows.length||transferTotal>0||purchaseTotal>0){
const parts=[];
if(transferTotal>0)parts.push(`transferir <b>${fmt.format(transferTotal)}</b> piezas entre canales`);
if(purchaseTotal>0)parts.push(`comprar <b>${fmt.format(purchaseTotal)}</b> piezas${topPurchase?` (prioridad: ${escapeHtml(topPurchase.sku_madre)})`:''}`);
bannerText=`Esta semana conviene ${parts.join(' y ')}.`;
}else{
bannerText='No se detectaron acciones urgentes con la base actual. Operación estable en todos los canales.';
}
document.getElementById('actionBanner').innerHTML=`<div class="action-banner-icon">→</div><div><b>Próxima acción sugerida:</b> ${bannerText}</div>`;
}
function safeSum(rows,key){return (rows||[]).reduce((a,r)=>a+n(r[key]),0);}
function goToAlerts(risk){goPage('alerts');setTimeout(()=>{const sel=document.getElementById('alertRiskFilter');if(sel){sel.value=risk;renderAlerts();}},50);}
function productSummaryRows(){return DATA.product_summary||[];}
function renderExecutive(){
  renderSemaphore();
  const k=DATA.kpis,total=Math.max(n(k.stock_total),1),summary=productSummaryRows();
  const active=summary.filter(r=>n(r.ventas_90d_unidades)>0).length;
  document.getElementById('executiveKpis').innerHTML=[
    kpiCard('Inventario total',fmt.format(k.stock_total),'Odoo + tránsito + Full',COLORS.blue,'▦',100,'Todas las piezas consolidadas por IQ.'),
    kpiCard('Inventario Odoo',fmt.format(k.stock_odoo),`${fmt1.format(n(k.stock_odoo)/total*100)}% del inventario`,COLORS.purple,'O',n(k.stock_odoo)/total*100,'Piezas disponibles en bodega central.'),
    kpiCard('En tránsito',fmt.format(k.stock_transito),'Traslados en estado Listo',COLORS.amber,'⇄',n(k.stock_transito)/total*100,'Piezas que están viajando hacia un canal.'),
    kpiCard('Inventario Full',fmt.format(k.stock_full),'Disponible en canales Full',COLORS.teal,'F',n(k.stock_full)/total*100,'Piezas listas para venta inmediata en Full.'),
    kpiCard('Ventas 10 días',fmt.format(k.ventas_10||0),'Ritmo más reciente',COLORS.green,'↗',100,'Unidades vendidas en los últimos 10 días.'),
    kpiCard('Ventas 30 días',fmt.format(k.ventas_30||0),'Ritmo mensual',COLORS.blue,'↗',100,'Unidades vendidas en los últimos 30 días.'),
    kpiCard('Ventas 90 días',fmt.format(k.ventas_90||0),'Histórico consolidado',COLORS.blue,'↗',100,'Unidades vendidas en los últimos 90 días.'),
    kpiCard('SKU con ventas',fmt.format(active),'Al menos una venta en 90 días',COLORS.slate,'#',Math.min(active/Math.max(summary.length,1)*100,100),'Cantidad de IQ con movimiento de venta reciente.'),
    kpiCard('Inversión en inventario',money.format(k.valor_inventario||0),'A costo, todos los canales',COLORS.blue,'$',100,'Valor a costo de todas las piezas: Odoo + tránsito + Full.'),
    kpiCard(`Inventario +${AGE_LIMIT()} días`,money.format(k.valor_mas_90d||0),`${fmt1.format(n(k.valor_mas_90d)/Math.max(n(k.valor_inventario),1)*100)}% de la inversión`,COLORS.red,'◷',n(k.valor_mas_90d)/Math.max(n(k.valor_inventario),1)*100,'Inversión en piezas con más días que el umbral de antigüedad. Abre Lotes y antigüedad.',"goPage('lots')")
  ].join('');
  const cs=DATA.channel_summary.filter(r=>DATA.meta.channels.includes(r.canal)),channels=cs.map(r=>r.canal);
  plot('chartInventoryChannel',[{type:'bar',name:'Odoo',x:channels,y:cs.map(r=>n(r.inventario_odoo)),marker:{color:COLORS.odoo}},{type:'bar',name:'Tránsito',x:channels,y:cs.map(r=>n(r.inventario_transito)),marker:{color:COLORS.transit}},{type:'bar',name:'Full',x:channels,y:cs.map(r=>n(r.inventario_full)),marker:{color:COLORS.full}}],{barmode:'stack',yaxis:{title:'Unidades',rangemode:'tozero'},xaxis:{tickangle:-18}});
  const salesByChannel=channels.map(c=>{const rows=DATA.rotation.filter(r=>r.canal===c);return {canal:c,d10:safeSum(rows,'ventas_10d_unidades'),d30:safeSum(rows,'ventas_30d_unidades'),d90:safeSum(rows,'ventas_90d_unidades')}});
  plot('chartSalesChannel',[{type:'bar',name:'10 días',x:channels,y:salesByChannel.map(r=>r.d10),marker:{color:COLORS.green}},{type:'bar',name:'30 días',x:channels,y:salesByChannel.map(r=>r.d30),marker:{color:COLORS.blue}},{type:'bar',name:'90 días',x:channels,y:salesByChannel.map(r=>r.d90),marker:{color:COLORS.purple}}],{barmode:'group',yaxis:{title:'Unidades vendidas',rangemode:'tozero'},xaxis:{tickangle:-18}});
  const top=[...summary].sort((a,b)=>n(b.ventas_90d_unidades)-n(a.ventas_90d_unidades)).slice(0,15).reverse();
  plot('chartTopProducts',[{type:'bar',orientation:'h',y:top.map(r=>r.sku_madre),x:top.map(r=>n(r.ventas_90d_unidades)),text:top.map(r=>fmt.format(r.ventas_90d_unidades)),textposition:'outside',marker:{color:COLORS.blue},hovertext:top.map(r=>r.producto_madre||''),hovertemplate:'%{y}<br>%{hovertext}<br>Ventas 90d: %{x:,.0f}<extra></extra>'}],{margin:{l:72,r:38,t:10,b:40},xaxis:{title:'Unidades',rangemode:'tozero'}});
  plot('chartInventoryMix',[{type:'pie',hole:.67,labels:['Odoo','Tránsito','Full'],values:[k.stock_odoo,k.stock_transito,k.stock_full],marker:{colors:[COLORS.odoo,COLORS.transit,COLORS.full]},textinfo:'percent',hovertemplate:'%{label}: %{value:,.0f}<extra></extra>'}],{showlegend:true,margin:{l:10,r:10,t:10,b:35}});
  renderExecutiveProductTable();
  if(state.selectedSalesSku)renderSalesDetail();
}
function renderExecutiveProductTable(){
  const q=(document.getElementById('executiveProductSearch')?.value||state.pageQueries.executive||'').toLowerCase().trim();
  const rows=productSummaryRows().filter(r=>!q||`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''} ${r.canales||''}`.toLowerCase().includes(q));
  state.tables.executiveProducts=rows;
  renderTable('executiveProductTable',rows,[{key:'sku_madre',label:'IQ',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'skus_individuales',label:'SKU(s) originales',format:'product'},{key:'canales',label:'Canales',format:'product'},{key:'inventario_odoo',label:'Odoo',format:'num'},{key:'inventario_transito',label:'Tránsito',format:'num'},{key:'inventario_full',label:'Full',format:'num'},{key:'inventario_total',label:'Stock total',format:'numStrong'},{key:'ventas_10d_unidades',label:'Ventas 10d',format:'num'},{key:'ventas_30d_unidades',label:'Ventas 30d',format:'num'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'numStrong'},{key:'venta_90d_monto',label:'Monto 90d',format:'money'},{key:'sku_madre',label:'Desglose',format:'salesButton'}],{pageSize:25,id:'executiveProducts',searchInputId:'executiveProductSearch'});
}
function openSalesDetail(sku){state.selectedSalesSku=String(sku||'').toUpperCase();const panel=document.getElementById('salesDetailPanel');panel.classList.add('open');renderSalesDetail();panel.scrollIntoView({behavior:'smooth',block:'start'});}
function closeSalesDetail(){state.selectedSalesSku='';document.getElementById('salesDetailPanel')?.classList.remove('open');}
function selectedSalesRows(){const sku=state.selectedSalesSku,days=n(document.getElementById('salesDetailPeriod')?.value||90);let rows=(DATA.sales_detail||[]).filter(r=>String(r.sku_madre||'').toUpperCase()===sku);if(!rows.length)return [];const dates=rows.map(r=>new Date(r.fecha)).filter(d=>!Number.isNaN(d.getTime()));if(!dates.length)return rows;const maxDate=new Date(Math.max(...dates.map(d=>d.getTime())));const minDate=new Date(maxDate);minDate.setDate(minDate.getDate()-(days-1));return rows.filter(r=>{const d=new Date(r.fecha);return !Number.isNaN(d.getTime())&&d>=minDate&&d<=maxDate;});}
function renderSalesDetail(){
  const sku=state.selectedSalesSku;if(!sku)return;const rows=selectedSalesRows(),product=productSummaryRows().find(r=>String(r.sku_madre).toUpperCase()===sku)||{};
  document.getElementById('salesDetailTitle').textContent=`Ventas de ${sku}`;
  document.getElementById('salesDetailSubtitle').textContent=`${product.producto_madre||'Producto'} · desglose por canal y SKU original.`;
  const units=safeSum(rows,'unidades'),amount=safeSum(rows,'venta_total'),orders=safeSum(rows,'pedidos'),channels=new Set(rows.map(r=>r.canal)).size;
  document.getElementById('salesDetailStatus').innerHTML=[statusCard(fmt.format(units),'Unidades vendidas',COLORS.blue,'↗'),statusCard(money.format(amount),'Monto vendido',COLORS.green,'$'),statusCard(fmt.format(orders),'Pedidos',COLORS.purple,'#'),statusCard(fmt.format(channels),'Canales',COLORS.teal,'▦')].join('');
  const map={};rows.forEach(r=>{const key=`${r.canal}|||${r.sku_original||sku}`;if(!map[key])map[key]={canal:r.canal,sku_original:r.sku_original||sku,unidades:0,venta_total:0,pedidos:0,dias_con_venta:new Set()};map[key].unidades+=n(r.unidades);map[key].venta_total+=n(r.venta_total);map[key].pedidos+=n(r.pedidos);map[key].dias_con_venta.add(String(r.fecha).slice(0,10));});
  const breakdown=Object.values(map).map(r=>({...r,dias_con_venta:r.dias_con_venta.size})).sort((a,b)=>b.unidades-a.unidades);
  state.tables.salesBreakdown=breakdown;state.tables.salesDaily=rows;
  renderTable('salesBreakdownTable',breakdown,[{key:'canal',label:'Canal'},{key:'sku_original',label:'SKU original',format:'sku'},{key:'unidades',label:'Unidades',format:'numStrong'},{key:'venta_total',label:'Monto',format:'money'},{key:'pedidos',label:'Pedidos',format:'num'},{key:'dias_con_venta',label:'Días con venta',format:'num'}],{pageSize:15,id:'salesBreakdown'});
  renderTable('salesDailyTable',rows,[{key:'fecha',label:'Fecha'},{key:'canal',label:'Canal'},{key:'sku_original',label:'SKU original',format:'sku'},{key:'modalidad_venta',label:'Tipo',format:'badge'},{key:'unidades',label:'Unidades',format:'numStrong'},{key:'venta_total',label:'Monto',format:'money'},{key:'pedidos',label:'Pedidos',format:'num'},{key:'fuente',label:'Fuente'}],{pageSize:20,id:'salesDaily'});
}
function exportSkuSales(){if(!state.selectedSalesSku){showToast('Selecciona un SKU');return;}const rows=selectedSalesRows();downloadCsv(rows,`ventas_${state.selectedSalesSku}_${document.getElementById('salesDetailPeriod')?.value||90}d`);}
function channelSupportsFull(channel){return (DATA.meta.params.full_channels||[]).includes(channel);}
function showOdooPanel(){const panel=document.getElementById('channelOdooPanel');if(!panel)return;panel.classList.add('open','flash');renderOdooPanel();scrollToId('channelOdooPanel');setTimeout(()=>panel.classList.remove('flash'),1100);}
function closeOdooPanel(){const panel=document.getElementById('channelOdooPanel');if(panel)panel.classList.remove('open','flash');}
function renderOdooPanel(){
  const panel=document.getElementById('channelOdooPanel');if(!panel)return;
  const channel=state.channel,supports=channelSupportsFull(channel),q=(document.getElementById('channelOdooSearch')?.value||'').toLowerCase().trim(),filter=document.getElementById('channelOdooFilter')?.value||'';
  let rows=channelRows(channel).filter(r=>n(r.inventario_odoo)>0);
  rows=rows.filter(r=>!q||`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''}`.toLowerCase().includes(q));
  if(filter==='send')rows=rows.filter(r=>n(r.transferencia_sugerida)>0);
  if(filter==='manual')rows=rows.filter(r=>String(r.check_manual_full||'').toUpperCase().startsWith('SI'));
  if(filter==='covered')rows=rows.filter(r=>n(r.transferencia_sugerida)<=0);
  rows=[...rows].sort((a,b)=>{const ma=String(a.check_manual_full||'').startsWith('SI')?1:0,mb=String(b.check_manual_full||'').startsWith('SI')?1:0;return n(b.transferencia_sugerida)-n(a.transferencia_sugerida)||mb-ma||n(b.ventas_90d_unidades)-n(a.ventas_90d_unidades)});
  const odoo=safeSum(rows,'inventario_odoo'),send=safeSum(rows,'transferencia_sugerida'),manual=rows.filter(r=>String(r.check_manual_full||'').toUpperCase().startsWith('SI')).length,withSend=rows.filter(r=>n(r.transferencia_sugerida)>0).length;
  document.getElementById('channelOdooSubtitle').textContent=`${channel}: stock disponible en Odoo y recomendación para Full por producto.`;
  document.getElementById('channelOdooStatus').innerHTML=[statusCard(fmt.format(rows.length),'SKU con stock Odoo',COLORS.purple,'O'),statusCard(fmt.format(odoo),'Piezas Odoo mostradas',COLORS.purple,'▦'),statusCard(fmt.format(send),'Piezas sugeridas a Full',COLORS.amber,'⇄'),statusCard(fmt.format(manual),'Checks manuales',COLORS.purple,'i')].join('');
  document.getElementById('channelOdooRule').innerHTML=supports?`<div class="odoo-rule-icon">${DATA.meta.params.full_target_days||30}</div><div><strong>Regla de sugerencia.</strong> Primero se proyecta cuánto Full quedará después de <strong>${DATA.meta.params.full_transit_days||5} días</strong> de venta mientras viaja el traslado. Después se completa para que, al llegar, existan <strong>${DATA.meta.params.full_target_days||30} días</strong> de cobertura, contando también el tránsito actual. La cantidad nunca supera el stock disponible en Odoo. Productos sin historial: máximo <strong>${DATA.meta.params.new_product_full_max||10} piezas</strong> y <strong>CHECK MANUAL</strong>.</div>`:`<div class="odoo-rule-icon">—</div><div><strong>${channel} no tiene operación Full configurada.</strong> Se muestra el stock Odoo para consulta, pero la sugerencia de traslado permanece en cero.</div>`;
  state.tables.odooPanel=rows;
  renderTable('channelOdooTable',rows,[{key:'sku_madre',label:'IQ',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'skus_individuales',label:'SKU(s) canal',format:'product'},{key:'inventario_odoo',label:'Odoo',format:'numStrong'},{key:'inventario_full',label:'Full hoy',format:'num'},{key:'consumo_estimado_transito',label:'Consumo tránsito',format:'num'},{key:'full_proyectado_arribo',label:'Full al arribo',format:'num'},{key:'inventario_transito',label:'Tránsito actual',format:'num'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},{key:'objetivo_full_piezas',label:'Objetivo al arribo',format:'num'},{key:'transferencia_sugerida',label:'Enviar a Full',format:'numStrong'},{key:'cobertura_proyectada_arribo_dias',label:'Cob. al arribo',format:'coverage'},{key:'check_manual_full',label:'Validación',format:'badge'},{key:'criterio_transferencia_full',label:'Criterio',format:'product'},{key:'sku_madre',label:'Admin',format:'adminTransferButton'}],{pageSize:20,id:'odooPanel',searchInputId:'channelOdooSearch'});
}
function openChannelTransfers(){const channel=state.channel;goPage('transfers');setTimeout(()=>{const el=document.getElementById('transferSearch');if(el)el.value=channel;state.pageQueries.transfers=channel;renderTransfers();},30);}
function openChannelPurchases(){const channel=state.channel;goPage('purchases');setTimeout(()=>{const el=document.getElementById('purchaseSearch');if(el)el.value=channel;state.pageQueries.purchases=channel;renderPurchases();},30);}
function openChannelPriorities(){goPage('alerts');setTimeout(()=>{const el=document.getElementById('alertSearch');if(el)el.value=state.channel;state.pageQueries.alerts=state.channel;renderAlerts();},30);}
function renderChannel(){const channel=state.channel,rows=channelRows(channel),s=channelSummary(channel),color=CHANNEL_COLORS[channel]||COLORS.blue;document.getElementById('channelLogo').innerHTML=CHANNEL_MARKS[channel]||channelIcon(channel);document.getElementById('channelLogo').style.setProperty('--channel',color);document.getElementById('channelName').textContent=channel;document.getElementById('channelKam').textContent=`Responsable: ${DATA.meta.kam_map[channel]||'Por asignar'}`;const critical=rows.filter(r=>['CRÍTICO','ALTO'].includes(r.nivel_riesgo)).length;document.getElementById('channelState').innerHTML=critical?`<span class="badge red" style="cursor:pointer" onclick="openChannelPriorities()" title="Ver en el Centro de alertas">${critical} prioridades</span>`:`<span class="badge green">Operación estable</span>`;document.getElementById('channelUpdated').textContent=`Actualizado ${DATA.meta.generated_at}`;const total=Math.max(n(s.inventario_total),1);document.getElementById('channelKpis').innerHTML=[kpiCard('Inventario Odoo',fmt.format(s.inventario_odoo||0),'Click para ver envíos a Full',COLORS.purple,'O',n(s.inventario_odoo)/total*100,'Piezas en nuestra bodega central. Abre el detalle por producto y la sugerencia de envío para un mes de cobertura Full.','showOdooPanel()'),kpiCard('En tránsito',fmt.format(s.inventario_transito||0),'Traslados en estado Listo',COLORS.amber,'⇄',n(s.inventario_transito)/total*100,'Piezas que ya salieron hacia este canal pero todavía no llegan al almacén Full.'),kpiCard('Inventario Full',fmt.format(s.inventario_full||0),'Disponible en Full',COLORS.teal,'F',n(s.inventario_full)/total*100,'Piezas ya recibidas en el almacén del canal, listas para venderse hoy mismo.'),kpiCard('Ventas 90 días',fmt.format(s.ventas_90d_unidades||0),'Unidades del canal',color,'↗',100,'Unidades vendidas en este canal durante los últimos 90 días.'),kpiCard('Cobertura Full',coverage(s.cobertura_full_dias),'Contra ventas Full',COLORS.green,'◷',Math.min(n(s.cobertura_full_dias)/90*100,100),'Días que dura el inventario ya disponible en Full, al ritmo de venta actual del canal.'),kpiCard('Cobertura total',coverage(s.cobertura_total_dias),'Odoo + tránsito + Full',COLORS.blue,'▦',Math.min(n(s.cobertura_total_dias)/90*100,100),'Días que dura el inventario del canal sumando Odoo, tránsito y Full.'),kpiCard('Inversión en canal',money.format(s.valor_inventario||0),'Costo de las piezas del canal',COLORS.blue,'$',100,'Valor a costo de todas las piezas del canal (Odoo + tránsito + Full). Costo Odoo; si falta, costo del producto o última compra.'),kpiCard(`Inventario +${AGE_LIMIT()} días`,money.format(s.valor_mas_90d||0),`${fmt.format(s.skus_mas_90d||0)} SKU con piezas viejas`,COLORS.red,'◷',Math.min(n(s.valor_mas_90d)/Math.max(n(s.valor_inventario),1)*100,100),'Inversión en piezas que llevan más del umbral en el canal. Abre la tabla filtrada de antigüedad.',"setStaleFilter('90')"),kpiCard('Transferir',fmt.format(s.transferencias_sugeridas||0),'Piezas sugeridas Odoo → Full',COLORS.amber,'⇄',Math.min(n(s.transferencias_sugeridas)/total*100,100),'Sí: este número es la suma de transferencia_sugerida del canal. Abre el listado exacto de piezas por IQ para revisar o ejecutar.','openChannelTransfers()'),kpiCard('Comprar',fmt.format(s.compras_sugeridas||0),'Necesidad diagnóstica del canal',COLORS.red,'◫',Math.min(n(s.compras_sugeridas)/Math.max(n(s.ventas_90d_unidades),1)*100,100),'Abre Compras sugeridas. La orden final se consulta en la tabla consolidada por SKU, no sumando canales.','openChannelPurchases()')].join('');plot('chartChannelInventory',[{type:'pie',hole:.62,labels:['Odoo','Tránsito','Full'],values:[s.inventario_odoo||0,s.inventario_transito||0,s.inventario_full||0],marker:{colors:[COLORS.purple,COLORS.amber,COLORS.teal]},textinfo:'label+percent',hovertemplate:'%{label}: %{value:,.0f}<extra></extra>'}],{margin:{l:10,r:10,t:10,b:30}});const top=[...rows].sort((a,b)=>n(b.ventas_90d_unidades)-n(a.ventas_90d_unidades)).slice(0,12).reverse();plot('chartChannelTop',[{type:'bar',orientation:'h',y:top.map(r=>r.sku_madre),x:top.map(r=>n(r.ventas_90d_unidades)),text:top.map(r=>fmt.format(r.ventas_90d_unidades)),textposition:'outside',marker:{color},hovertext:top.map(r=>r.producto_madre||''),hovertemplate:'%{y}<br>%{hovertext}<br>Ventas: %{x:,.0f}<extra></extra>'}],{margin:{l:70,r:40,t:10,b:40},xaxis:{title:'Unidades',rangemode:'tozero'}});const points=rows.filter(r=>n(r.ventas_90d_unidades)>0&&Number.isFinite(Number(r.cobertura_total_dias)));plot('chartChannelRisk',[{type:'scatter',mode:'markers',x:points.map(r=>n(r.ventas_90d_unidades)),y:points.map(r=>n(r.cobertura_total_dias)),text:points.map(r=>`${r.sku_madre}<br>${r.producto_madre||''}<br>${r.estado_cobertura||''}`),hovertemplate:'%{text}<br>Ventas: %{x:,.0f}<br>Cobertura: %{y:.1f} días<extra></extra>',marker:{size:points.map(r=>clamp(Math.sqrt(n(r.inventario_total))+6,8,30)),color:points.map(r=>({red:COLORS.red,amber:COLORS.amber,green:COLORS.green,blue:COLORS.blue,purple:COLORS.purple})[coverageTone(r.prioridad_cobertura)]||COLORS.slate),opacity:.78,line:{color:'#fff',width:1}}}],{xaxis:{title:'Ventas 90 días',rangemode:'tozero'},yaxis:{title:'Cobertura total (días)',range:[0,Math.min(180,Math.max(...points.map(r=>n(r.cobertura_total_dias)),30))]},shapes:[{type:'line',x0:0,x1:1,xref:'paper',y0:30,y1:30,line:{color:COLORS.red,dash:'dot'}},{type:'line',x0:0,x1:1,xref:'paper',y0:90,y1:90,line:{color:COLORS.blue,dash:'dot'}}]});const priority=[...rows].filter(r=>n(r.prioridad_cobertura)<=4||n(r.transferencia_sugerida)>0).sort((a,b)=>n(a.prioridad_cobertura)-n(b.prioridad_cobertura)||n(b.ventas_90d_unidades)-n(a.ventas_90d_unidades)).slice(0,7);document.getElementById('channelInsights').innerHTML=priority.map(r=>insight(`${r.sku_madre} · ${r.accion_kam||compactAction(r.accion_preliminar)}`,`${r.producto_madre||'Producto'}. ${r.estado_cobertura||''}: Full ${fmt.format(r.inventario_full)}, tránsito ${fmt.format(r.inventario_transito)}, Odoo ${fmt.format(r.inventario_odoo)}; alcanza ${coverage(r.cobertura_total_dias)}.`,COLORS[{red:'red',amber:'amber',green:'green',blue:'blue',purple:'purple'}[coverageTone(r.prioridad_cobertura)]]||COLORS.slate,n(r.prioridad_cobertura)<=3?'!':'→')).join('')||insight('Sin acciones urgentes','Ningún SKU del canal está sin stock, por agotarse o con envío pendiente.',COLORS.green,'✓');renderChannelActions();renderChannelStale();renderChannelCoverage();renderChannelListings();renderChannelSales();renderChannelMoves();renderChannelTable();renderChannelDictionary();if(document.getElementById('channelOdooPanel')?.classList.contains('open'))renderOdooPanel();}
/* ============================================================
   GRÁFICA DE VENTAS DIARIAS DEL CANAL
   ============================================================
   Fuente: DATA.channel_sales_daily, un renglón por (fecha, canal, sku_madre)
   con unidades y monto de la base operativa (script 02) y utilidad/base_roi
   de base_utilidad_roi.xlsx.

   El ROI NO se promedia entre días: se recalcula como utilidad/base sobre el
   periodo completo. Promediar razones da un número que no corresponde al
   total y que además cambiaría de forma incoherente al mover el rango. */
const CHANNEL_SALES_METRICS={
  unidades:{label:'Unidades vendidas',axis:'Unidades',hover:'%{x|%d %b %Y}<br>Unidades: %{y:,.0f}<extra></extra>',fmt:v=>fmt.format(v)},
  monto:{label:'Monto vendido',axis:'Monto vendido (MXN)',hover:'%{x|%d %b %Y}<br>Monto: $%{y:,.0f}<extra></extra>',fmt:v=>money.format(v)},
  utilidad:{label:'Utilidad',axis:'Utilidad (MXN)',hover:'%{x|%d %b %Y}<br>Utilidad: $%{y:,.0f}<extra></extra>',fmt:v=>money.format(v)},
  roi:{label:'ROI',axis:'ROI del día',hover:'%{x|%d %b %Y}<br>ROI: %{y:.2%}<extra></extra>',fmt:v=>pct2(v)},
};
/* Fechas en UTC a propósito: las claves vienen como 'YYYY-MM-DD' y usar la
   zona local haría que el mismo día se corriera según el huso del navegador. */
function isoShift(iso,days){const d=new Date(`${iso}T00:00:00Z`);d.setUTCDate(d.getUTCDate()+days);return d.toISOString().slice(0,10);}
function isoRange(fromIso,toIso){const out=[];let cur=fromIso,guard=0;while(cur<=toIso&&guard++<1500){out.push(cur);cur=isoShift(cur,1);}return out;}
function channelSalesRows(){const ch=state.channel;return (DATA.channel_sales_daily||[]).filter(r=>r.canal===ch);}
function channelSalesBounds(rows){const fechas=rows.map(r=>String(r.fecha||'')).filter(Boolean);if(!fechas.length)return null;return {min:fechas.reduce((a,b)=>a<b?a:b),max:fechas.reduce((a,b)=>a>b?a:b)};}
function channelSalesSkuOptions(){
  const mapa=new Map();
  channelSalesRows().forEach(r=>{const sku=String(r.sku_madre||'').toUpperCase();if(!sku)return;const cur=mapa.get(sku)||{sku,label:'',units:0};cur.units+=n(r.unidades);if(!cur.label&&r.producto_madre)cur.label=String(r.producto_madre);mapa.set(sku,cur);});
  return [...mapa.values()].sort((a,b)=>b.units-a.units).map(o=>({sku:o.sku,label:o.label||'Producto'}));
}
let channelSalesCombo=null;
function ensureChannelSalesCombo(){
  if(channelSalesCombo)return;
  channelSalesCombo=initCombo({inputId:'channelSalesSku',panelId:'channelSalesSkuPanel',clearId:'channelSalesSkuClear',getOptions:channelSalesSkuOptions,
    onSelect:(opt)=>{state.channelSales.sku=opt?opt.sku:'';renderChannelSales();},
    onInput:(val)=>{if(!String(val||'').trim()){state.channelSales.sku='';renderChannelSales();}}});
}
function channelSalesWindow(bounds){
  const cfg=state.channelSales;
  if(cfg.range==='all')return {from:bounds.min,to:bounds.max};
  if(cfg.range==='custom'){
    let from=cfg.from||isoShift(bounds.max,-89),to=cfg.to||bounds.max;
    if(from>to){const tmp=from;from=to;to=tmp;}
    return {from:from<bounds.min?bounds.min:from,to:to>bounds.max?bounds.max:to};
  }
  const dias=Math.max(1,n(cfg.range)||90);
  const from=isoShift(bounds.max,-(dias-1));
  return {from:from<bounds.min?bounds.min:from,to:bounds.max};
}
function movingAverage(values,ventana){
  return values.map((_,i)=>{
    if(i<ventana-1)return null;
    let suma=0,cuenta=0;
    for(let j=i-ventana+1;j<=i;j++){const v=values[j];if(v===null||v===undefined||!Number.isFinite(v))continue;suma+=v;cuenta++;}
    return cuenta?suma/cuenta:null;
  });
}
function channelSalesSeries(){
  const rows=channelSalesRows(),bounds=channelSalesBounds(rows);
  if(!bounds)return null;
  const cfg=state.channelSales,{from,to}=channelSalesWindow(bounds);
  const sku=String(cfg.sku||'').toUpperCase();
  const sel=rows.filter(r=>{const f=String(r.fecha||'');return f>=from&&f<=to&&(!sku||String(r.sku_madre||'').toUpperCase()===sku);});
  const porDia=new Map();
  const total={unidades:0,monto:0,utilidad:0,base:0,utilBase:0};
  sel.forEach(r=>{
    const k=String(r.fecha),cur=porDia.get(k)||{unidades:0,monto:0,utilidad:0,base:0,utilBase:0};
    cur.unidades+=n(r.unidades);cur.monto+=n(r.monto);cur.utilidad+=n(r.utilidad);cur.base+=n(r.base_roi);cur.utilBase+=n(r.utilidad_con_base);
    porDia.set(k,cur);
    total.unidades+=n(r.unidades);total.monto+=n(r.monto);total.utilidad+=n(r.utilidad);total.base+=n(r.base_roi);total.utilBase+=n(r.utilidad_con_base);
  });
  const dias=isoRange(from,to);
  const values=dias.map(d=>{
    const v=porDia.get(d);
    if(cfg.metric==='roi')return (v&&v.base>0)?v.utilBase/v.base:null;
    if(!v)return 0;
    if(cfg.metric==='monto')return v.monto;
    if(cfg.metric==='utilidad')return v.utilidad;
    return v.unidades;
  });
  return {bounds,from,to,dias,values,total,rows:sel,porDia,diasConVenta:porDia.size,
          roiPeriodo:total.base>0?total.utilBase/total.base:null};
}
function renderChannelSales(){
  const card=document.getElementById('channelSalesCard');if(!card)return;
  ensureChannelSalesCombo();
  const cfg=state.channelSales;
  /* El SKU elegido es de un canal concreto. Al cambiar de pestaña se limpia
     en vez de mostrar una gráfica vacía sin explicación. */
  if(cfg.sku&&!channelSalesRows().some(r=>String(r.sku_madre||'').toUpperCase()===cfg.sku)){cfg.sku='';channelSalesCombo&&channelSalesCombo.setValue('');}
  const metric=CHANNEL_SALES_METRICS[cfg.metric]||CHANNEL_SALES_METRICS.unidades;
  const metricSel=document.getElementById('channelSalesMetric');if(metricSel)metricSel.value=cfg.metric;
  const rangeSel=document.getElementById('channelSalesRange');if(rangeSel)rangeSel.value=cfg.range;
  document.getElementById('channelSalesDates')?.classList.toggle('open',cfg.range==='custom');
  const statusBox=document.getElementById('channelSalesStatus'),note=document.getElementById('channelSalesNote');

  const serie=channelSalesSeries();
  if(!serie){
    if(statusBox)statusBox.innerHTML='';
    if(note)note.textContent='';
    document.getElementById('chartChannelSales').innerHTML='<div class="empty"><div class="empty-icon">◌</div>Este canal no tiene ventas con fecha en la base actual.</div>';
    state.tables.channelSales=[];
    return;
  }

  /* Los inputs de fecha se acotan al periodo realmente disponible: pedir un
     rango fuera de la base devolvería una gráfica vacía sin explicar por qué. */
  const inputFrom=document.getElementById('channelSalesFrom'),inputTo=document.getElementById('channelSalesTo');
  [inputFrom,inputTo].forEach(el=>{if(el){el.min=serie.bounds.min;el.max=serie.bounds.max;}});
  if(inputFrom)inputFrom.value=serie.from;
  if(inputTo)inputTo.value=serie.to;

  const color=CHANNEL_COLORS[state.channel]||COLORS.blue;
  const traces=[];
  if(cfg.metric==='roi'){
    traces.push({type:'scatter',mode:'lines+markers',name:'ROI del día',x:serie.dias,y:serie.values,connectgaps:false,
      line:{color,width:2},marker:{size:5,color},hovertemplate:metric.hover});
  }else{
    traces.push({type:'bar',name:metric.label,x:serie.dias,y:serie.values,
      marker:{color:serie.values.map(v=>n(v)<0?COLORS.red:color)},hovertemplate:metric.hover});
    if(serie.dias.length>=14){
      traces.push({type:'scatter',mode:'lines',name:'Media móvil 7 días',x:serie.dias,y:movingAverage(serie.values,7),
        line:{color:COLORS.amber,width:2,dash:'dot'},hovertemplate:'%{x|%d %b %Y}<br>Media 7d: %{y:,.1f}<extra></extra>'});
    }
  }
  const layout={xaxis:{title:'Día',type:'date'},yaxis:{title:metric.axis},margin:{l:64,r:20,t:14,b:52},
                legend:{orientation:'h',y:-.22},bargap:.25};
  if(cfg.metric==='roi'){layout.yaxis.tickformat='.0%';layout.shapes=[{type:'line',xref:'paper',x0:0,x1:1,y0:0,y1:0,line:{color:COLORS.slate,dash:'dot',width:1}}];}
  else if(cfg.metric==='utilidad'){layout.shapes=[{type:'line',xref:'paper',x0:0,x1:1,y0:0,y1:0,line:{color:COLORS.slate,dash:'dot',width:1}}];}
  else layout.yaxis.rangemode='tozero';
  plot('chartChannelSales',traces,layout);

  const dias=serie.dias.length,promedio=dias?serie.total.unidades/dias:0;
  if(statusBox)statusBox.innerHTML=[
    statusCard(fmt.format(serie.total.unidades),`Unidades · ${fmt.format(dias)} días`,color,'▦'),
    statusCard(money.format(serie.total.monto),'Monto vendido',COLORS.blue,'$'),
    statusCard(money.format(serie.total.utilidad),'Utilidad del periodo',serie.total.utilidad<0?COLORS.red:COLORS.green,'↗'),
    statusCard(serie.roiPeriodo===null?'Sin dato':pct2(serie.roiPeriodo),'ROI del periodo',COLORS.purple,'%'),
  ].join('');

  const skuTxt=cfg.sku?`SKU ${cfg.sku}`:'venta general del canal';
  const avisoUtilidad=(!DATA.meta.sales_daily_has_profit&&(cfg.metric==='utilidad'||cfg.metric==='roi'))
    ? ' La base actual no trae utilidad por día: vuelve a correr build_utilidad_roi.py para poblar esta métrica.'
    : '';
  if(note)note.textContent=`${metric.label} · ${skuTxt} · ${serie.from} a ${serie.to} (${fmt.format(serie.diasConVenta)} días con movimiento de ${fmt.format(dias)}). Periodo disponible en la base: ${serie.bounds.min} a ${serie.bounds.max}. Promedio ${fmt1.format(promedio)} unidades/día.${avisoUtilidad}`;

  const vacio={unidades:0,monto:0,utilidad:0,base:0,utilBase:0};
  state.tables.channelSales=serie.dias.map((d,i)=>{
    const agg=serie.porDia.get(d)||vacio;
    return {fecha:d,canal:state.channel,sku:cfg.sku||'TODOS',unidades:agg.unidades,monto:agg.monto,utilidad:agg.utilidad,roi:agg.base>0?agg.utilBase/agg.base:'',metrica_graficada:serie.values[i]===null?'':serie.values[i]};
  });
}
function exportChannelSales(){
  const rows=state.tables.channelSales||[];
  if(!rows.length){showToast('No hay serie que exportar');return;}
  const sku=state.channelSales.sku||'general';
  downloadCsv(rows,`ventas_${state.channel}_${sku}_${state.channelSales.metric}`);
}
function renderChannelTable(){const q=(document.getElementById('channelTableSearch')?.value||'').toLowerCase(),risk=document.getElementById('channelRiskFilter')?.value||'';const rows=channelRows(state.channel).filter(r=>(!risk||r.nivel_riesgo===risk)&&(!q||`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''}`.toLowerCase().includes(q)));state.tables.channel=rows;renderTable('channelTable',rows,rotationColumns(state.views.channel),{pageSize:20,id:'channel',searchInputId:'channelTableSearch'});}
/* Utilidad y ROI que ESTE canal generó con el SKU (hoja 'por_producto_canal'
   de base_utilidad_roi.xlsx). No es el ROI consolidado del producto: un mismo
   IQ puede ser rentable en un canal y no en otro, y en la pestaña del canal
   solo importa lo que pasó ahí. Se declara `numFilter` para que renderTable
   genere el filtro por rango automáticamente. */
function financeChannelColumns(){return[
  {key:'utilidad_canal_30d',label:'Utilidad 30d',format:'moneySigned',numFilter:{label:'Utilidad 30d',options:UTILIDAD_FILTER_OPTIONS}},
  {key:'roi_canal_30d',label:'ROI 30d',format:'roiPct',numFilter:{label:'ROI 30d',options:ROI_FILTER_OPTIONS}},
  {key:'utilidad_canal_90d',label:'Utilidad 90d',format:'moneySigned',numFilter:{label:'Utilidad 90d',options:UTILIDAD_FILTER_OPTIONS}},
  {key:'roi_canal_90d',label:'ROI 90d',format:'roiPct',numFilter:{label:'ROI 90d',options:ROI_FILTER_OPTIONS}},
];}
/* ============================================================
   VERSIÓN 2026-09 · ANTIGÜEDAD POR INVERSIÓN, COBERTURA Y AUDITORÍA
   ============================================================ */
const AGE_LIMIT=()=>n(DATA.meta.params.age_threshold)||90;
const MOVE_TYPE_LABELS={REASIGNACION_ENTRE_CANALES:'Reasignación entre canales',SALIDA_POR_AJUSTE:'Ajuste de salida',ENTRADA_POR_AJUSTE:'Ajuste de entrada',SALIDA_A_UBICACION_NO_COMERCIAL:'Salida a ubicación no comercial',ENTRADA_DESDE_UBICACION_NO_COMERCIAL:'Entrada desde ubicación no comercial',DEVOLUCION_CLIENTE:'Devolución de cliente',ENVIO_A_FULL:'Envío a Full',OTRO:'Otro'};
function moveTypeLabel(v){return MOVE_TYPE_LABELS[String(v||'')]||String(v||'—');}
function levelTone(v){return ({ALTA:'red',MEDIA:'amber',BAJA:'gray'})[String(v||'').toUpperCase()]||'gray';}
function coverageTone(p){p=n(p);if(p>=1&&p<=3)return 'red';if(p===4)return 'amber';if(p===5)return 'green';if(p===6)return 'blue';if(p===7)return 'purple';return 'gray';}
/* Días que definen la alerta: capa más antigua (o promedio, según criterio del 02). */
function ageDays(r){const a=numFilterValue(r,'dias_alerta_antiguedad');if(Number.isFinite(a))return a;const b=numFilterValue(r,'dias_inventario_canal');return Number.isFinite(b)?b:NaN;}
function isOverAge(r,limit){const d=ageDays(r);return Number.isFinite(d)&&d>limit;}
function is90(r){return n(r.unidades_mas_90d)>0||isOverAge(r,AGE_LIMIT());}
/* Desplaza solo la ventana (scrollIntoView también movía contenedores con
   overflow oculto y descuadraba el diseño) y deja libre la barra fija. */
function scrollToId(id){const el=document.getElementById(id);if(!el)return;const top=document.querySelector('.topbar')?.offsetHeight||64,jump=document.querySelector('.channel-jump');const extra=jump&&getComputedStyle(jump).position==='sticky'?jump.offsetHeight+12:12;const reduce=window.matchMedia('(prefers-reduced-motion: reduce)').matches;window.scrollTo({top:Math.max(el.getBoundingClientRect().top+window.scrollY-top-extra,0),behavior:reduce?'auto':'smooth'});}
function channelStockRows(channel){return channelRows(channel).filter(r=>n(r.inventario_total)>0);}

/* ---------- Antigüedad ---------- */
function staleColumns(){return[
  {key:'sku_madre',label:'SKU',format:'sku'},
  {key:'producto_madre',label:'Producto',format:'product'},
  {key:'inventario_total',label:'Stock en canal',format:'num'},
  {key:'valor_inventario_canal',label:'Inversión en canal',format:'money'},
  {key:'dias_inventario_canal',label:'Días promedio',format:'daysAge'},
  {key:'dias_inventario_max',label:'Pieza más antigua',format:'daysAge'},
  {key:'unidades_mas_90d',label:'Piezas +90 días',format:'num'},
  {key:'valor_mas_90d',label:'Inversión +90 días',format:'money'},
  {key:'ritmo_venta_canal',label:'Ritmo (u/día)',format:'num1'},
  {key:'alerta_antiguedad_canal',label:'Alerta',format:'badge'},
  {key:'fuente_dias_inventario',label:'Origen del dato',format:'badge'},
  {key:'movs_revision_90d',label:'Movs. en revisión',format:'num'},
  {key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},
  ...financeChannelColumns(),
  {key:'accion_kam',label:'Acción sugerida',format:'product'}];}
function staleRowsForFilter(filter){const q=(document.getElementById('channelStaleSearch')?.value||'').toLowerCase();let rows=channelStockRows(state.channel);if(filter==='90')rows=rows.filter(is90);else if(filter==='60')rows=rows.filter(r=>isOverAge(r,60));else if(filter==='30')rows=rows.filter(r=>isOverAge(r,30));if(q)rows=rows.filter(r=>`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''}`.toLowerCase().includes(q));return rows;}
function setStaleFilter(value,scroll=true){const sel=document.getElementById('channelStaleStockFilter');if(sel)sel.value=value;state.tables.channelStale_page=1;state.tableSorts.channelStale={key:value==='90'?'valor_mas_90d':'valor_inventario_canal',dir:'desc'};renderChannelStale();if(scroll)scrollToId('secAge');}
function downloadStale90(){const rows=channelStockRows(state.channel).filter(is90).sort((a,b)=>n(b.valor_mas_90d)-n(a.valor_mas_90d));downloadCsv(rows.map(r=>({sku_madre:r.sku_madre,producto:r.producto_madre,canal:r.canal,stock_canal:r.inventario_total,piezas_mas_90d:r.unidades_mas_90d,inversion_canal:r.valor_inventario_canal,inversion_mas_90d:r.valor_mas_90d,dias_promedio:r.dias_inventario_canal,pieza_mas_antigua_dias:r.dias_inventario_max,origen_dato:r.fuente_dias_inventario,detalle_origen:r.detalle_fuente_dias,ventas_90d:r.ventas_90d_unidades,movs_en_revision_90d:r.movs_revision_90d,accion:r.accion_kam})),`antiguedad_90d_${state.channel.replace(/\s+/g,'_')}`);}
function renderChannelStale(){
  const filter=document.getElementById('channelStaleStockFilter')?.value||'';
  const rows=staleRowsForFilter(filter);
  state.tables.channelStale=rows;
  const all=channelStockRows(state.channel),rows90=all.filter(is90);
  const units=all.reduce((a,r)=>a+n(r.inventario_total),0),valor=safeSum(all,'valor_inventario_canal'),valor90=safeSum(all,'valor_mas_90d'),piezas90=safeSum(all,'unidades_mas_90d');
  const conDato=all.filter(r=>Number.isFinite(numFilterValue(r,'dias_inventario_canal'))&&n(r.inventario_total)>0);
  const uDato=conDato.reduce((a,r)=>a+n(r.inventario_total),0),prom=uDato?conDato.reduce((a,r)=>a+n(r.dias_inventario_canal)*n(r.inventario_total),0)/uDato:NaN;
  const on90=filter==='90';
  document.getElementById('channelStaleStatus').innerHTML=[
    statusCard(fmt.format(all.length),'SKU con stock en el canal',COLORS.slate,'◌',{onClick:"setStaleFilter('')",active:filter==='',hint:'Ver todos'}),
    statusCard(fmt.format(rows90.length),`En alerta +${AGE_LIMIT()} días · ${fmt.format(piezas90)} piezas`,COLORS.red,'!',{onClick:on90?"setStaleFilter('')":"setStaleFilter('90')",active:on90,hint:on90?'Quitar filtro':'Ver en la tabla',download:rows90.length?'downloadStale90()':''}),
    statusCard(money.format(valor90),`Inversión con +${AGE_LIMIT()} días`,COLORS.red,'$',{onClick:"setStaleFilter('90')",active:on90,hint:'Ver SKUs'}),
    statusCard(money.format(valor),`Inversión en el canal · ${fmt.format(units)} piezas`,COLORS.blue,'▦'),
    statusCard(Number.isFinite(prom)?`${fmt1.format(prom)} días`:'Sin dato','Antigüedad promedio por pieza',COLORS.amber,'◷')
  ].join('');
  renderTable('channelStaleTable',rows,staleColumns(),{pageSize:20,id:'channelStale',searchInputId:'channelStaleSearch',defaultSortKey:'valor_inventario_canal',defaultSortDir:'desc'});
  const box=document.getElementById('channelStaleTable');
  if(box&&on90)box.insertAdjacentHTML('afterbegin',`<div class="filter-banner"><span>Mostrando los ${fmt.format(rows.length)} SKU con piezas de más de ${AGE_LIMIT()} días, de mayor a menor inversión afectada.</span><span><button type="button" onclick="downloadStale90()">Descargar CSV</button> <button type="button" onclick="setStaleFilter('',false)">Quitar filtro</button></span></div>`);
}

/* ---------- Cobertura ---------- */
const COVERAGE_GROUPS={urgent:[1,2,3],low:[4],ok:[5],excess:[6,7],nosales:[8]};
function setCoverageFilter(value,scroll=true){const sel=document.getElementById('channelCoverageFilter');if(sel)sel.value=value;state.tables.channelCoverage_page=1;renderChannelCoverage();if(scroll)scrollToId(value?'channelCoverageCard':'secCoverage');}
function coverageColumns(){return[
  {key:'prioridad_cobertura',label:'Urgencia',format:'num'},
  {key:'sku_madre',label:'SKU',format:'sku'},
  {key:'producto_madre',label:'Producto',format:'product'},
  {key:'estado_cobertura',label:'Estado',format:'coverageBadge'},
  {key:'venta_diaria_calendario',label:'Venta diaria',format:'num1'},
  {key:'inventario_full',label:'Full',format:'num'},
  {key:'inventario_transito',label:'Tránsito',format:'num'},
  {key:'inventario_odoo',label:'Odoo',format:'num'},
  {key:'cobertura_full_dias',label:'Días Full',format:'coverage'},
  {key:'cobertura_total_dias',label:'Días total',format:'coverage'},
  {key:'fecha_estimada_quiebre',label:'Se agota aprox.',format:'text'},
  {key:'transferencia_sugerida',label:'Enviar a Full',format:'numStrong'},
  {key:'compra_sugerida',label:'Faltante compra',format:'num'},
  {key:'accion_kam',label:'Qué hacer',format:'product'},
  {key:'sku_madre',label:'Admin',format:'adminTransferButton'}];}
function renderChannelCoverage(){
  const all=channelRows(state.channel).filter(r=>n(r.prioridad_cobertura)<9);
  const filter=document.getElementById('channelCoverageFilter')?.value||'',q=(document.getElementById('channelCoverageSearch')?.value||'').toLowerCase();
  const count=g=>all.filter(r=>COVERAGE_GROUPS[g].includes(n(r.prioridad_cobertura)));
  const P=DATA.meta.params;
  const card=(g,value,label,tone,icon)=>statusCard(fmt.format(value),label,tone,icon,{onClick:`setCoverageFilter('${filter===g?'':g}')`,active:filter===g,hint:filter===g?'Quitar filtro':'Ver SKUs'});
  document.getElementById('channelCoverageStatus').innerHTML=[
    card('urgent',count('urgent').length,`Sin stock o se agotan en < ${P.coverage_break_days} días`,COLORS.red,'!'),
    card('low',count('low').length,`Cobertura baja (${P.coverage_break_days}–${P.coverage_low_days} días)`,COLORS.amber,'◔'),
    card('ok',count('ok').length,`Cobertura sana (${P.coverage_low_days}–${P.coverage_high_days} días)`,COLORS.green,'✓'),
    card('excess',count('excess').length,`Cobertura alta o exceso (> ${P.coverage_high_days} días)`,COLORS.purple,'▲'),
    card('nosales',count('nosales').length,'Con stock y sin ventas en 90 días',COLORS.slate,'○')
  ].join('');
  let rows=filter?count(filter):all.slice();
  if(q)rows=rows.filter(r=>`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''}`.toLowerCase().includes(q));
  rows.sort((a,b)=>n(a.prioridad_cobertura)-n(b.prioridad_cobertura)||(Number.isFinite(numFilterValue(a,'cobertura_total_dias'))?n(a.cobertura_total_dias):1e9)-(Number.isFinite(numFilterValue(b,'cobertura_total_dias'))?n(b.cobertura_total_dias):1e9)||n(b.ventas_90d_unidades)-n(a.ventas_90d_unidades));
  state.tables.channelCoverage=rows;
  renderTable('channelCoverageTable',rows,coverageColumns(),{pageSize:20,id:'channelCoverage',searchInputId:'channelCoverageSearch',defaultSortKey:'prioridad_cobertura',defaultSortDir:'asc'});
  const guide=document.getElementById('coverageGuide');
  if(guide)guide.innerHTML=[
    ['red',`Sin stock / < ${P.coverage_break_days} días`,'Se pierde venta esta semana. Enviar a Full o comprar hoy.'],
    ['red','Full se agota antes del envío',`Full no alcanza para los ${P.full_transit_days} días que tarda un traslado.`],
    ['amber',`${P.coverage_break_days}–${P.coverage_low_days} días`,'Programar envío o compra en los próximos días.'],
    ['green',`${P.coverage_low_days}–${P.coverage_high_days} días`,'Nivel sano. No requiere acción.'],
    ['purple',`> ${P.coverage_high_days} días`,'Hay de más. No enviar ni comprar; evaluar promoción o reasignar.'],
    ['gray','Sin ventas 90 días','Revisar publicación, precio o liquidar.']
  ].map(([t,a,b])=>`<div class="cg-row"><span class="badge ${t}">${escapeHtml(a)}</span><span>${escapeHtml(b)}</span></div>`).join('');
}

/* ---------- Qué hacer hoy ---------- */
function actionTile(title,value,text,tone,onClick){return `<button type="button" class="action-tile" style="--tone:${tone}" onclick="${onClick}"><span class="a-title">${escapeHtml(title)}</span><span class="a-value">${value}</span><span class="a-text">${escapeHtml(text)}</span></button>`;}
function renderChannelActions(){
  const ch=state.channel,rows=channelRows(ch),stock=rows.filter(r=>n(r.inventario_total)>0);
  const send=rows.filter(r=>n(r.transferencia_sugerida)>0),urgent=rows.filter(r=>[1,2,3].includes(n(r.prioridad_cobertura)));
  const aged=stock.filter(is90),excess=rows.filter(r=>[6,7].includes(n(r.prioridad_cobertura)));
  const moves=channelMoveRows(ch).filter(m=>String(m.nivel_revision).toUpperCase()==='ALTA');
  const supports=channelSupportsFull(ch);
  document.getElementById('channelActions').innerHTML=[
    actionTile('Enviar a Full',`${fmt.format(safeSum(send,'transferencia_sugerida'))} pzas`,supports?`${fmt.format(send.length)} SKU necesitan envío para cubrir ${DATA.meta.params.full_target_days} días al llegar.`:'Este canal no tiene operación Full configurada.',COLORS.amber,'showOdooPanel()'),
    actionTile('Stock por agotar',`${fmt.format(urgent.length)} SKU`,urgent.length?`Sin stock o por agotarse. En conjunto venden ${fmt1.format(safeSum(urgent,'venta_diaria_calendario'))} pzas por día.`:'Ningún SKU con venta está por agotarse.',urgent.length?COLORS.red:COLORS.green,"setCoverageFilter('urgent')"),
    actionTile(`Inventario +${AGE_LIMIT()} días`,money.format(safeSum(aged,'valor_mas_90d')),`${fmt.format(aged.length)} SKU con ${fmt.format(safeSum(aged,'unidades_mas_90d'))} piezas viejas. Prioriza por inversión.`,aged.length?COLORS.red:COLORS.green,"setStaleFilter('90')"),
    actionTile('Exceso de cobertura',money.format(safeSum(excess,'valor_inventario_canal')),`${fmt.format(excess.length)} SKU con más de ${DATA.meta.params.coverage_high_days} días. No enviar más.`,COLORS.purple,"setCoverageFilter('excess')"),
    (DATA.meta.params.autoazur_available&&(DATA.meta.params.autoazur_channels_ok||[]).includes(ch)?(()=>{const pend=channelListingRows(ch);return actionTile('Stock sin publicar',money.format(safeSum(pend,'valor_inventario_canal')),pend.length?`${fmt.format(pend.length)} SKU con piezas en el canal y sin publicación activa en AutoAzur.`:'Todo el stock asignado tiene publicación activa.',pend.length?COLORS.red:COLORS.green,"setListingsFilter('')");})():actionTile('Stock sin publicar','—','Sin datos de AutoAzur para este canal.',COLORS.slate,"scrollToId('secListings')")),
    actionTile('Movimientos en revisión',fmt.format(moves.length),moves.length?'Traslados o ajustes con piezas viejas o de ida y vuelta.':'Sin movimientos de revisión alta.',moves.length?COLORS.red:COLORS.green,"scrollToId('secMoves')")
  ].join('');
}

/* ---------- Movimientos Odoo ---------- */
function channelMoveRows(ch){return (DATA.audit_moves||[]).filter(m=>m.canal_origen===ch||m.canal_destino===ch||m.canal_kam_afectado===ch||m.canal_destino===`Full ${ch}`);}
function moveColumns(withChannel=false){return[
  {key:'fecha_movimiento',label:'Fecha',format:'dateShort'},
  {key:'nivel_revision',label:'Revisión',format:'levelBadge'},
  {key:'tipo_movimiento',label:'Tipo',format:'moveType'},
  ...(withChannel?[{key:'canal_kam_afectado',label:'Canal afectado'}]:[]),
  {key:'sku_madre',label:'SKU',format:'sku'},
  {key:'producto_madre',label:'Producto',format:'product'},
  {key:'cantidad',label:'Piezas',format:'num'},
  {key:'valor_movido',label:'Valor',format:'money'},
  {key:'edad_lote_al_movimiento',label:'Edad al mover',format:'daysAge'},
  {key:'ubicacion_origen',label:'De',format:'product'},
  {key:'ubicacion_destino',label:'A',format:'product'},
  {key:'usuario',label:'Usuario'},
  {key:'referencia',label:'Referencia'},
  {key:'motivo_revision',label:'Motivo',format:'product'}];}
function filterMoves(rows,level,q){if(level==='review')rows=rows.filter(m=>['ALTA','MEDIA'].includes(String(m.nivel_revision).toUpperCase()));else if(level)rows=rows.filter(m=>String(m.nivel_revision).toUpperCase()===level);q=(q||'').toLowerCase().trim();if(q)rows=rows.filter(m=>`${m.sku_madre} ${m.producto_madre} ${m.usuario} ${m.referencia} ${m.sku_original}`.toLowerCase().includes(q));return rows;}
function moveStatus(rows){return[
  statusCard(fmt.format(rows.length),'Movimientos',COLORS.slate,'⇆'),
  statusCard(fmt.format(safeSum(rows,'cantidad')),'Piezas movidas',COLORS.blue,'▦'),
  statusCard(money.format(safeSum(rows,'valor_movido')),'Valor movido',COLORS.purple,'$'),
  statusCard(fmt.format(rows.filter(m=>String(m.nivel_revision).toUpperCase()==='ALTA').length),'Con revisión alta',COLORS.red,'!')].join('');}
function renderChannelMoves(){
  const all=channelMoveRows(state.channel),level=document.getElementById('channelMovesFilter')?.value??'review';
  const rows=filterMoves(all,level,document.getElementById('channelMovesSearch')?.value);
  state.tables.channelMoves=rows;
  document.getElementById('channelMovesStatus').innerHTML=moveStatus(rows);
  if(!(DATA.audit_moves||[]).length){document.getElementById('channelMovesTable').innerHTML=insight('Sin registro de movimientos','La base no trae la hoja movimientos_odoo_auditoria. Corre el 01 actualizado para empezar el registro.',COLORS.slate,'i');return;}
  renderTable('channelMovesTable',rows,moveColumns(),{pageSize:15,id:'channelMoves',searchInputId:'channelMovesSearch',defaultSortKey:'fecha_movimiento',defaultSortDir:'desc'});
}
function populateAuditControls(){const sel=document.getElementById('auditChannel');if(!sel||sel.dataset.ready)return;const chans=[...new Set((DATA.audit_moves||[]).map(m=>m.canal_kam_afectado).filter(Boolean))].sort();sel.innerHTML='<option value="">Todos los canales</option>'+chans.map(c=>`<option>${escapeHtml(c)}</option>`).join('');sel.dataset.ready='1';}
function renderAudit(){
  populateAuditControls();
  const ch=document.getElementById('auditChannel')?.value||'',level=document.getElementById('auditLevel')?.value??'review';
  let rows=DATA.audit_moves||[];if(ch)rows=rows.filter(m=>m.canal_kam_afectado===ch);
  rows=filterMoves(rows,level,document.getElementById('auditSearch')?.value);
  state.tables.auditMoves=rows;
  const users=new Set(rows.map(m=>m.usuario).filter(Boolean));
  document.getElementById('auditStatus').innerHTML=moveStatus(rows)+statusCard(fmt.format(users.size),'Usuarios involucrados',COLORS.amber,'◉');
  const P=DATA.meta.params;
  document.getElementById('auditRulesNote').textContent=`Revisión alta: piezas con ${P.audit_age_threshold} días o más al moverse (umbral editable), idas y vueltas entre canales o ajustes de salida y reingreso del mismo producto. Revisión media: salidas de un canal con KAM hacia General, B2B o ajustes, y envíos a Full de piezas viejas. Se muestran los últimos ${P.audit_days} días; el histórico completo queda en registro_movimientos_odoo.csv.`;
  const canales=[...new Set(rows.map(m=>m.canal_kam_afectado||'Sin canal'))];
  const serie=lvl=>canales.map(c=>rows.filter(m=>(m.canal_kam_afectado||'Sin canal')===c&&String(m.nivel_revision).toUpperCase()===lvl).reduce((a,m)=>a+n(m.cantidad),0));
  plot('chartAudit',[{type:'bar',name:'Revisión alta',x:canales,y:serie('ALTA'),marker:{color:COLORS.red}},{type:'bar',name:'Revisión media',x:canales,y:serie('MEDIA'),marker:{color:COLORS.amber}},{type:'bar',name:'Sin señales',x:canales,y:serie('BAJA'),marker:{color:COLORS.slate}}],{barmode:'stack',yaxis:{title:'Piezas',rangemode:'tozero'}});
  const byUser={};rows.forEach(m=>{const u=m.usuario||'Sin usuario';const x=byUser[u]||(byUser[u]={usuario:u,movimientos:0,piezas:0,valor:0,alta:0,canales:new Set()});x.movimientos++;x.piezas+=n(m.cantidad);x.valor+=n(m.valor_movido);if(String(m.nivel_revision).toUpperCase()==='ALTA')x.alta++;if(m.canal_kam_afectado)x.canales.add(m.canal_kam_afectado);});
  const userRows=Object.values(byUser).map(x=>({...x,canales:[...x.canales].join(' | ')}));
  state.tables.auditUsers=userRows;
  renderTable('auditUsersTable',userRows,[{key:'usuario',label:'Usuario'},{key:'alta',label:'Revisión alta',format:'num'},{key:'movimientos',label:'Movimientos',format:'num'},{key:'valor',label:'Valor',format:'money'},{key:'piezas',label:'Piezas',format:'num'},{key:'canales',label:'Canales'}],{pageSize:8,id:'auditUsers',defaultSortKey:'alta',defaultSortDir:'desc'});
  renderTable('auditMovesTable',rows,moveColumns(true),{pageSize:25,id:'auditMoves',searchInputId:'auditSearch',defaultSortKey:'fecha_movimiento',defaultSortDir:'desc'});
}
function syncTopbarHeight(){const t=document.querySelector('.topbar');if(t)document.documentElement.style.setProperty('--topbar-h',`${t.offsetHeight}px`);}
/* ---------- Publicaciones AutoAzur ---------- */
const LISTING_GROUPS={nopub:['Sin publicación'],paused:['Pausada','Pausada · sin stock'],review:['En revisión'],error:['Con error / rechazada'],closed:['Inactiva','Cerrada / eliminada','Otro','Sin estado'],zero:['Activa con stock 0','Publicada con stock 0']};
const LISTING_PENDING=Object.values(LISTING_GROUPS).flat();
function listingTone(v){v=String(v||'');if(v==='Sin publicación'||v.startsWith('Con error'))return 'red';if(v.startsWith('Pausada'))return 'amber';if(v==='En revisión')return 'blue';if(v==='Activa con stock 0'||v==='Publicada con stock 0')return 'purple';if(v==='Activa'||v==='OK')return 'green';return 'gray';}
function channelListingRows(ch){return channelRows(ch).filter(r=>n(r.inventario_total)>0&&LISTING_PENDING.includes(String(r.alerta_publicacion||'')));}
function setListingsFilter(value,scroll=true){const sel=document.getElementById('channelListingsFilter');if(sel)sel.value=value;state.tables.channelListings_page=1;renderChannelListings();if(scroll)scrollToId('secListings');}
function downloadListingsNoStock(){const ch=state.channel,rows=(DATA.listings_no_stock||[]).filter(r=>r.canal===ch);downloadCsv(rows.map(r=>({sku_madre:r.sku_madre,producto:r.producto_madre,canal:r.canal,publicaciones_activas:r.n_activas,item_ids:r.item_ids,cuentas:r.cuentas,stock_publicado:r.stock_publicado_activo,estados:r.estados_detalle})),`publicadas_sin_stock_${ch.replace(/\s+/g,'_')}`);}
function listingsColumns(){return[
  {key:'sku_madre',label:'SKU',format:'sku'},
  {key:'producto_madre',label:'Producto',format:'product'},
  {key:'alerta_publicacion',label:'Estado publicación',format:'listingBadge'},
  {key:'inventario_total',label:'Stock en canal',format:'numStrong'},
  {key:'inventario_odoo',label:'Odoo',format:'num'},
  {key:'inventario_full',label:'Full',format:'num'},
  {key:'valor_inventario_canal',label:'Inversión',format:'money'},
  {key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},
  {key:'n_publicaciones',label:'Publicaciones',format:'num'},
  {key:'estados_detalle',label:'Detalle AutoAzur',format:'product'},
  {key:'item_ids',label:'ItemID',format:'product'},
  {key:'cuentas',label:'Cuenta',format:'product'},
  {key:'accion_publicacion',label:'Qué hacer',format:'product'}];}
function renderChannelListings(){
  const ch=state.channel,notice=document.getElementById('channelListingsNotice'),P=DATA.meta.params;
  if(!P.autoazur_available){notice.innerHTML=P.autoazur_global_error?insight('AutoAzur no se consultó en la última corrida del 01',`${P.autoazur_global_error}. Agrega AUTOAZUR_UNIQUE_GUID=<tu GUID> a ese .env y vuelve a correr 01 → 02 → 03.`,COLORS.amber,'!'):insight('Sin datos de AutoAzur','Agrega AUTOAZUR_UNIQUE_GUID al .env y corre el 01 para consultar las publicaciones de cada canal.',COLORS.slate,'i');document.getElementById('channelListingsStatus').innerHTML='';document.getElementById('channelListingsTable').innerHTML='';state.tables.channelListings=[];return;}
  const log=(DATA.autoazur_log||[]).find(l=>l.canal===ch);
  if(!(P.autoazur_channels_ok||[]).includes(ch)){const motivos={ERROR:()=>`La consulta falló (${log.detalle||'sin detalle'}).`,SIN_PUBLICACIONES:()=>'AutoAzur no devolvió publicaciones para este canal.',CAMPOS_NO_DETECTADOS:()=>`AutoAzur respondió ${fmt.format(n(log.publicaciones))} publicaciones, pero no se reconoció el campo de SKU ni el de ItemID, así que no se pueden ligar a los IQ. Revisa autoazur_publicaciones_muestra.json y agrega el nombre del campo al .env (AUTOAZUR_CAMPO_SKU_PUBLICACION).`,SIN_CAMPO_ESTADO:()=>`Se descargaron ${fmt.format(n(log.publicaciones))} publicaciones, pero no se reconoció el campo de estado. Agrega su nombre al .env (AUTOAZUR_CAMPO_ESTADO_ORIGINAL).`,INCOMPLETO:()=>`Solo se descargaron ${fmt.format(n(log.publicaciones))} de ${fmt.format(n(log.total_informado))} publicaciones que informa AutoAzur (${fmt1.format(n(log.cobertura_descarga)*100)}%).`};const motivo=log?(motivos[log.estado_consulta]?motivos[log.estado_consulta]():`Estado de la consulta: ${log.estado_consulta}.`):'El canal no está en AUTOAZUR_CANALES_PUBLICACIONES.';notice.innerHTML=insight(`No se pudo evaluar ${ch}`,`${motivo} Para no marcar SKU como "sin publicar" por error, esta sección queda vacía hasta que la consulta sea confiable.`,COLORS.amber,'!');document.getElementById('channelListingsStatus').innerHTML='';document.getElementById('channelListingsTable').innerHTML='';state.tables.channelListings=[];return;}
  const sinLigar=n((DATA.listings_unlinked_by_channel||{})[ch]);
  const avisoLiga=sinLigar?`<div class="filter-banner"><span><b>${fmt.format(sinLigar)} publicaciones</b> de este canal tienen un SKU que no está en el diccionario. Si un IQ se publicó con uno de esos SKU, aquí aparecerá como "Sin publicación" aunque sí esté publicado. La lista para completar el diccionario está en Calidad de datos.</span></div>`:'';
  notice.innerHTML=avisoLiga+(log&&log.estado_consulta==='SIN_CAMPO_ESTADO'?(n(log.estado_desde_detalle)>0?`<div class="filter-banner"><span>El estado (activa, pausada, en revisión) se obtuvo del detalle de vinculaciones de AutoAzur para ${fmt.format(n(log.estado_desde_detalle))} publicaciones con stock. Las demás se evalúan por el stock publicado.</span></div>`:`<div class="filter-banner"><span>AutoAzur no informa el estado de las publicaciones: se detectan SKU <b>sin publicación</b> y <b>publicados con stock 0</b> (el marketplace suele pausarlos). Pausadas por otro motivo, en revisión o con error no se pueden distinguir.</span></div>`):log&&log.estado_consulta==='INCOMPLETO'?`<div class="filter-banner"><span>AutoAzur informó ${fmt.format(n(log.total_informado))} publicaciones y se descargaron ${fmt.format(n(log.publicaciones))}. Algunos SKU podrían aparecer como "sin publicación" por error.</span></div>`:'');
  const all=channelListingRows(ch),filter=document.getElementById('channelListingsFilter')?.value||'',q=(document.getElementById('channelListingsSearch')?.value||'').toLowerCase();
  const grp=g=>all.filter(r=>LISTING_GROUPS[g].includes(String(r.alerta_publicacion)));
  const card=(g,label,tone,icon)=>{const rows=grp(g);return statusCard(fmt.format(rows.length),`${label} · ${money.format(safeSum(rows,'valor_inventario_canal'))}`,tone,icon,{onClick:`setListingsFilter('${filter===g?'':g}')`,active:filter===g,hint:filter===g?'Quitar filtro':'Ver SKUs'});};
  const noStock=(DATA.listings_no_stock||[]).filter(r=>r.canal===ch);
  document.getElementById('channelListingsStatus').innerHTML=[
    card('nopub','Sin publicación',COLORS.red,'∅'),card('paused','Pausadas',COLORS.amber,'Ⅱ'),card('review','En revisión',COLORS.blue,'◷'),
    card('error','Con error o rechazadas',COLORS.red,'!'),card('closed','Inactivas o cerradas',COLORS.slate,'×'),card('zero','Publicadas con stock 0',COLORS.purple,'0'),
    statusCard(fmt.format(noStock.length),'Activas sin stock asignado al canal',COLORS.slate,'○',{download:noStock.length?'downloadListingsNoStock()':''})
  ].join('');
  let rows=filter?grp(filter):all.slice();
  if(q)rows=rows.filter(r=>`${r.sku_madre} ${r.producto_madre} ${r.item_ids||''} ${r.skus_publicados||''} ${r.skus_individuales||''}`.toLowerCase().includes(q));
  state.tables.channelListings=rows;
  renderTable('channelListingsTable',rows,listingsColumns(),{pageSize:20,id:'channelListings',searchInputId:'channelListingsSearch',defaultSortKey:'valor_inventario_canal',defaultSortDir:'desc'});
}
function renderListingsQuality(){const box=document.getElementById('listingsUnlinkedTable');if(!box)return;const rows=DATA.listings_unlinked||[];state.tables.listingsUnlinked=rows;if(!DATA.meta.params.autoazur_available){box.innerHTML='<div class="empty">Sin datos de AutoAzur en esta base.</div>';return;}const resumen=rows.length&&'canales' in rows[0];renderTable('listingsUnlinkedTable',rows,resumen?[{key:'sku_publicacion',label:'SKU en AutoAzur',format:'sku'},{key:'titulo',label:'Título',format:'product'},{key:'upc',label:'UPC'},{key:'canales',label:'Canales'},{key:'publicaciones',label:'Publicaciones',format:'num'},{key:'stock_publicado',label:'Stock publicado',format:'num'},{key:'item_ids',label:'ItemID',format:'product'}]:[{key:'canal',label:'Canal'},{key:'sku_publicacion',label:'SKU publicación',format:'sku'},{key:'item_id',label:'ItemID'},{key:'titulo',label:'Título',format:'product'},{key:'cuenta',label:'Cuenta'},{key:'stock_publicado',label:'Stock publicado',format:'num'}],{pageSize:12,id:'listingsUnlinked',defaultSortKey:resumen?'stock_publicado':undefined});}

function initChannelJump(){syncTopbarHeight();window.addEventListener('resize',syncTopbarHeight);document.querySelectorAll('.channel-jump button').forEach(b=>b.addEventListener('click',()=>scrollToId(b.dataset.target)));}

function rotationColumns(mode='detail'){const simple=[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'inventario_total',label:'Stock total',format:'num'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},{key:'cobertura_total_dias',label:'Cobertura',format:'coverage'},...financeChannelColumns(),{key:'nivel_riesgo',label:'Riesgo',format:'badge'},{key:'accion_preliminar',label:'Acción',format:'badge'}];const detail=[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'inventario_odoo',label:'Odoo',format:'num'},{key:'inventario_transito',label:'Tránsito',format:'num'},{key:'inventario_full',label:'Full',format:'num'},{key:'inventario_total',label:'Total',format:'num'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},{key:'ventas_full_90d',label:'Full 90d',format:'num'},{key:'ventas_drop_90d',label:'Drop 90d',format:'num'},{key:'cobertura_full_dias',label:'Cob. Full',format:'coverage'},{key:'cobertura_total_dias',label:'Cob. Total',format:'coverage'},...financeChannelColumns(),{key:'demanda_tipo',label:'Demanda',format:'badge'},{key:'nivel_riesgo',label:'Riesgo',format:'badge'},{key:'accion_preliminar',label:'Acción',format:'badge'}];return mode==='simple'?simple:detail;}
function alertsColumns(mode='detail'){const simple=[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'canal',label:'Canal'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},{key:'cobertura_total_dias',label:'Días de stock',format:'coverage'},{key:'nivel_riesgo',label:'Riesgo',format:'badge'},{key:'accion_preliminar',label:'Acción',format:'badge'}];const detail=[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'canal',label:'Canal'},{key:'inventario_odoo',label:'Odoo',format:'num'},{key:'inventario_transito',label:'Tránsito',format:'num'},{key:'inventario_full',label:'Full',format:'num'},{key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},{key:'cobertura_full_dias',label:'Cob. Full',format:'coverage'},{key:'cobertura_total_dias',label:'Cob. Total',format:'coverage'},{key:'nivel_riesgo',label:'Riesgo',format:'badge'},{key:'accion_preliminar',label:'Acción',format:'badge'}];return mode==='simple'?simple:detail;}
function setupViewToggle(toggleId,onChange){const box=document.getElementById(toggleId);if(!box)return;box.querySelectorAll('button').forEach(btn=>btn.addEventListener('click',()=>{box.querySelectorAll('button').forEach(b=>b.classList.remove('active'));btn.classList.add('active');onChange(btn.dataset.view);}));}


function resolveSku(value){const raw=String(value||'').trim(),up=raw.toUpperCase();if(!raw)return '';const direct=DATA.rotation.find(r=>String(r.sku_madre||'').toUpperCase()===up);if(direct)return direct.sku_madre;for(const [iq,aliases] of Object.entries(DATA.sku_map||{})){if((aliases||[]).some(a=>String(a).toUpperCase()===up))return iq;}const partial=DATA.rotation.find(r=>`${r.sku_madre} ${r.producto_madre} ${r.skus_individuales||''}`.toUpperCase().includes(up));return partial?partial.sku_madre:'';}
function decisionRow(sku,channel){return DATA.rotation.find(r=>String(r.sku_madre)===String(sku)&&r.canal===channel)||null;}
function weightedDemand(r){const w=Number(r?.demanda_diaria_ponderada);return Number.isFinite(w)&&w>0?w:n(r?.ventas_90d_unidades)/90;}
function projectedAtArrival(r){return Math.max(n(r?.inventario_total)-weightedDemand(r)*n(DATA.meta.params.internal_transfer_days||5),0);}
function channelScenario(r){const d=weightedDemand(r),projected=projectedAtArrival(r),target=d>0?d*n(DATA.meta.params.redistribution_target_days||30):10,reserve=target,coverage=d>0?projected/d:9999,excess=Math.max(projected-reserve,0),odooCap=Math.max(Math.min(n(r?.inventario_odoo),excess),0),fullCap=Math.max(Math.min(n(r?.inventario_full),Math.max(excess-odooCap,0)),0),totalCap=odooCap+fullCap,need=d>0?Math.max(target-projected,0):0;return {d,projected,target,reserve,coverage,excess,odooCap,fullCap,totalCap,need};}
function evaluateInternalTransfer(skuInput,origin,destination,qtyInput){const sku=resolveSku(skuInput),qty=Math.max(Math.floor(n(qtyInput)),0);if(!sku)return {ok:false,title:'SKU no encontrado',tone:COLORS.red,body:'No se encontró el IQ ni un SKU individual relacionado en el diccionario cargado.'};if(!origin||!destination||origin===destination)return {ok:false,title:'Selecciona dos canales distintos',tone:COLORS.red,body:'El origen y el destino deben ser diferentes.'};const ro=decisionRow(sku,origin),rd=decisionRow(sku,destination);if(!ro||!rd)return {ok:false,title:'Producto no disponible en ambos canales',tone:COLORS.red,body:'El IQ debe existir en el detalle de origen y destino para poder comparar.'};const so=channelScenario(ro),sd=channelScenario(rd),ideal=Math.max(Math.floor(Math.min(so.totalCap,Math.ceil(sd.need))),0),proposed=qty||ideal,odooUse=Math.min(proposed,Math.floor(so.odooCap)),fullUse=Math.max(proposed-odooUse,0);const oAfter=Math.max(so.projected-proposed,0),dAfter=sd.projected+proposed,oCov=so.d>0?oAfter/so.d:9999,dCov=sd.d>0?dAfter/sd.d:9999;let verdict,tone,body;if(ideal<=0){verdict='MANTENER DISTRIBUCIÓN ACTUAL';tone=COLORS.green;body=sd.need<=0?`${destination} ya alcanza el objetivo proyectado; mover piezas no mejora la decisión.`:`${origin} no tiene excedente transferible sin comprometer su reserva.`;}else if(proposed>ideal){verdict='APROBAR CON MENOR CANTIDAD';tone=COLORS.amber;body=`La propuesta de ${fmt.format(proposed)} piezas supera la cantidad optimizada de ${fmt.format(ideal)}. Conviene reducirla para proteger el origen.`;}else if(proposed<ideal){verdict='APROBAR; PUEDE AUMENTARSE';tone=COLORS.blue;body=`La propuesta es viable. La cantidad optimizada es ${fmt.format(ideal)} piezas y ofrece mayor protección al destino.`;}else{verdict='APROBAR PROPUESTA';tone=COLORS.green;body=`La cantidad coincide con la recomendación optimizada y mejora el destino sin bajar el origen de su reserva proyectada.`;}const limitedHistory=(n(ro.dias_con_venta_90d)<2||n(rd.dias_con_venta_90d)<3),operationalReview=fullUse>0,manual=limitedHistory||operationalReview;if(operationalReview){verdict+=' · REVISIÓN OPERATIVA';tone=COLORS.amber;body+=` ${fmt.format(fullUse)} piezas provienen de Full y requieren retiro, relabeling y validación de costos/plazos.`;}if(limitedHistory){verdict+=' · CHECK MANUAL';tone=COLORS.purple;body+=' El historial es limitado o intermitente.';}return {ok:true,sku,origin,destination,qty:proposed,ideal,verdict,tone,body,ro,rd,so,sd,oAfter,dAfter,oCov,dCov,manual,odooUse,fullUse,operationalReview};}
function renderEvaluation(ev,targetId='kamEvaluationResult'){const box=document.getElementById(targetId);if(!box)return;if(!ev.ok){box.innerHTML=`<div class="decision-result" style="--tone:${ev.tone}"><h3>${escapeHtml(ev.title)}</h3><p>${escapeHtml(ev.body)}</p></div>`;return;}state.lastEvaluation=ev;const r=selectedSuggestion(),canOdoo=targetId==='kamEvaluationResult'&&r&&redistExecutable(r),key=r?redistKey(r):'',done=key&&odooDraftState.completed.has(key),busy=key&&odooDraftState.busy.has(key);const odooBtn=canOdoo?`<button class="admin-transfer-btn" ${done||busy?'disabled':''} onclick="createEvaluatedRedistributionDraft()">${done?'✓ Borrador creado':busy?'Creando…':'Crear borrador en Odoo'}</button>`:'';box.innerHTML=`<div class="decision-result" style="--tone:${ev.tone}"><h3>${escapeHtml(ev.verdict)}</h3><p>${escapeHtml(ev.body)}</p><p><b>${escapeHtml(ev.sku)}</b> · ${escapeHtml(ev.origin)} → ${escapeHtml(ev.destination)} · propuesta ${fmt.format(ev.qty)} · óptimo ${fmt.format(ev.ideal)}</p><div class="scenario-grid"><div class="scenario-box"><b>${escapeHtml(ev.origin)}</b><div>Stock actual: ${fmt.format(ev.ro.inventario_total)}<br>Cobertura al arribo sin mover: ${coverage(ev.so.coverage)}<br>Cobertura después: ${coverage(ev.oCov)}<br>Transferible Odoo: ${fmt.format(ev.so.odooCap)}<br>Transferible desde Full: ${fmt.format(ev.so.fullCap)}<br>Uso propuesto: ${fmt.format(ev.odooUse)} Odoo + ${fmt.format(ev.fullUse)} Full</div></div><div class="scenario-box"><b>${escapeHtml(ev.destination)}</b><div>Stock actual: ${fmt.format(ev.rd.inventario_total)}<br>Cobertura al arribo sin mover: ${coverage(ev.sd.coverage)}<br>Cobertura después: ${coverage(ev.dCov)}<br>Necesidad para 30 días: ${fmt.format(ev.sd.need)}</div></div></div><div class="decision-actions"><button class="primary-btn" onclick="saveCurrentDecision()">Guardar solicitud</button><button class="secondary-btn" onclick="copyCurrentDecision()">Copiar resumen</button>${odooBtn}</div>${canOdoo?'<div class="card-sub" style="margin-top:8px">Odoo se vuelve a consultar al ejecutar: producto correcto, lotes y stock disponible.</div>':''}</div>`;}
function evaluateKamForm(){state.pendingSuggestionIndex=null;const ev=evaluateInternalTransfer(document.getElementById('kamSku').value,document.getElementById('kamOrigin').value,document.getElementById('kamDestination').value,document.getElementById('kamQty').value);renderEvaluation(ev);}
function populateDecisionControls(){const channels=DATA.meta.channels||[],options=channels.map(c=>`<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join('');for(const id of ['kamOrigin','kamDestination']){const el=document.getElementById(id);if(el)el.innerHTML=`<option value="">Seleccionar</option>${options}`;}const list=document.getElementById('skuDecisionList');if(list){const uniq=[...new Map(DATA.rotation.map(r=>[r.sku_madre,r])).values()];list.innerHTML=uniq.map(r=>`<option value="${escapeHtml(r.sku_madre)}">${escapeHtml(r.producto_madre||'')}</option>`).join('');}}
function prefillSuggestion(index){const r=(DATA.redistributions||[])[index];if(!r)return;state.pendingSuggestionIndex=index;document.getElementById('kamSku').value=r.sku_madre;document.getElementById('kamOrigin').value=r.canal_origen;document.getElementById('kamDestination').value=r.canal_destino;document.getElementById('kamQty').value=r.cantidad_sugerida;const ev=evaluateInternalTransfer(r.sku_madre,r.canal_origen,r.canal_destino,r.cantidad_sugerida);renderEvaluation(ev);document.getElementById('kamEvaluationResult').scrollIntoView({behavior:'smooth',block:'center'});}
function renderSuggestions(){const q=(document.getElementById('suggestionSearch')?.value||'').toLowerCase().trim(),conf=document.getElementById('suggestionConfidence')?.value||'';let rows=(DATA.redistributions||[]).map((r,i)=>({...r,_index:i})).filter(r=>(!conf||r.confianza===conf)&&(!q||`${r.sku_madre} ${r.producto_madre} ${r.canal_origen} ${r.canal_destino}`.toLowerCase().includes(q)));state.tables.redistributions=rows;const high=rows.filter(r=>r.confianza==='ALTA').length,units=safeSum(rows,'cantidad_sugerida'),routes=new Set(rows.map(r=>`${r.canal_origen}>${r.canal_destino}`)).size,manual=rows.filter(r=>String(r.check_manual).toUpperCase()==='SI').length;document.getElementById('suggestionStatus').innerHTML=[statusCard(fmt.format(rows.length),'Oportunidades viables',COLORS.blue,'✦'),statusCard(fmt.format(units),'Piezas sugeridas',COLORS.amber,'⇄'),statusCard(fmt.format(high),'Confianza alta',COLORS.green,'✓'),statusCard(fmt.format(manual),'Check manual',COLORS.purple,'?')].join('');const cards=rows.slice(0,6);document.getElementById('suggestionCards').innerHTML=cards.length?cards.map(r=>`<div class="decision-card"><div class="decision-route"><span>${escapeHtml(r.canal_origen)}</span><span class="arrow">→</span><span>${escapeHtml(r.canal_destino)}</span></div><div class="decision-qty">${fmt.format(r.cantidad_sugerida)} <span style="font-size:11px;color:var(--muted)">pzs</span></div><div class="card-sub"><b>${escapeHtml(r.sku_madre)}</b> · ${escapeHtml(r.producto_madre||'Producto')}</div><div class="decision-meta"><div class="decision-metric"><b>${coverage(r.cobertura_origen_despues)}</b><span>Origen después</span></div><div class="decision-metric"><b>${coverage(r.cobertura_destino_despues)}</b><span>Destino después</span></div></div><div>${badge(r.confianza)} ${badge(r.veredicto)}</div><div class="decision-actions"><button class="primary-btn" onclick="prefillSuggestion(${r._index})">Evaluar movimiento</button></div></div>`).join(''):`<div class="history-empty" style="grid-column:1/-1">No hay sugerencias para los filtros actuales.</div>`;renderTable('redistributionTable',rows,[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'canal_origen',label:'Origen'},{key:'canal_destino',label:'Destino'},{key:'tipo_stock_origen',label:'Origen físico',format:'badge'},{key:'cantidad_sugerida',label:'Mover',format:'numStrong'},{key:'ventas_10d_origen',label:'Vta 10d origen',format:'num'},{key:'ventas_10d_destino',label:'Vta 10d destino',format:'num'},{key:'cobertura_origen_despues',label:'Origen después',format:'coverage'},{key:'cobertura_destino_despues',label:'Destino después',format:'coverage'},{key:'confianza',label:'Confianza',format:'badge'},{key:'veredicto',label:'Resultado',format:'badge'}],{pageSize:20,id:'redistributions',searchInputId:'suggestionSearch'});renderDecisionHistory();}
function answerKamQuestion(){const raw=document.getElementById('kamQuestion').value.trim();if(!raw){renderEvaluation({ok:false,title:'Escribe una pregunta',tone:COLORS.amber,body:'Incluye el SKU y los canales de origen y destino.'},'kamQuestionResult');return;}const upper=raw.toUpperCase(),skuMatch=upper.match(/IQ\s*\d+/),sku=skuMatch?skuMatch[0].replace(/\s+/g,''):resolveSku(raw),channels=(DATA.meta.channels||[]).filter(c=>upper.includes(c.toUpperCase())).sort((a,b)=>upper.indexOf(a.toUpperCase())-upper.indexOf(b.toUpperCase())),cleanForQty=raw.replace(/IQ\s*\d+/ig,''),qtyMatch=cleanForQty.match(/(\d+)\s*(?:PZS?|PIEZAS?)?/i),qty=qtyMatch?n(qtyMatch[1]):0;if(!sku||channels.length<2){renderEvaluation({ok:false,title:'No pude interpretar la consulta',tone:COLORS.amber,body:'Escribe algo como: “¿Conviene mover 30 piezas de Coppel a Liverpool del IQ123?”'},'kamQuestionResult');return;}const ev=evaluateInternalTransfer(sku,channels[0],channels[1],qty);renderEvaluation(ev,'kamQuestionResult');}
function decisionHistoryData(){try{return JSON.parse(localStorage.getItem('inventoryDecisionHistory')||'[]')}catch{return []}}
function saveCurrentDecision(){const ev=state.lastEvaluation;if(!ev)return;const hist=decisionHistoryData();hist.unshift({fecha:new Date().toLocaleString('es-MX'),sku:ev.sku,origen:ev.origin,destino:ev.destination,propuesta:ev.qty,recomendada:ev.ideal,resultado:ev.verdict,estado:'PENDIENTE DE AUTORIZACIÓN'});localStorage.setItem('inventoryDecisionHistory',JSON.stringify(hist.slice(0,100)));renderDecisionHistory();showToast('Solicitud guardada localmente');}
function copyCurrentDecision(){const ev=state.lastEvaluation;if(!ev)return;const txt=`${ev.verdict}\n${ev.sku}: ${ev.origin} → ${ev.destination}\nPropuesta: ${ev.qty} piezas\nRecomendación: ${ev.ideal} piezas\nOrigen después: ${fmt1.format(ev.oCov)} días\nDestino después: ${fmt1.format(ev.dCov)} días`;navigator.clipboard?.writeText(txt);showToast('Resumen copiado');}
function renderDecisionHistory(){const el=document.getElementById('decisionHistory');if(!el)return;const rows=decisionHistoryData();if(!rows.length){el.innerHTML='<div class="history-empty">Todavía no hay solicitudes guardadas en este navegador.</div>';return;}renderTable('decisionHistory',rows,[{key:'fecha',label:'Fecha'},{key:'sku',label:'SKU',format:'sku'},{key:'origen',label:'Origen'},{key:'destino',label:'Destino'},{key:'propuesta',label:'Propuesta',format:'num'},{key:'recomendada',label:'Recomendada',format:'numStrong'},{key:'resultado',label:'Resultado',format:'badge'},{key:'estado',label:'Estado',format:'badge'}],{pageSize:15,id:'decisionHistoryRows'});}
function clearDecisionHistory(){localStorage.removeItem('inventoryDecisionHistory');renderDecisionHistory();showToast('Historial local eliminado');}

function b2bRows(){return (DATA.b2b_portfolio||[]).filter(r=>n(r.b2b_disponible)>0);}
function selectedB2BRow(){const sku=state.selectedB2BSku||'';return b2bRows().find(r=>String(r.sku_madre||'').toUpperCase()===String(sku).toUpperCase())||null;}
function highlightMatch(text,q){const s=String(text??'');if(!q)return escapeHtml(s);const idx=s.toLowerCase().indexOf(q.toLowerCase());if(idx===-1)return escapeHtml(s);return `${escapeHtml(s.slice(0,idx))}<mark>${escapeHtml(s.slice(idx,idx+q.length))}</mark>${escapeHtml(s.slice(idx+q.length))}`;}
/* Combobox de autocomplete genérico: input de texto + panel desplegable que se reduce
   conforme se escribe. options: array de {sku, label}. onSelect(opt) al elegir o al escribir
   coincidencia exacta. onInput(rawQuery) se dispara en cada tecleo (para filtros en vivo). */
function initCombo({inputId,panelId,clearId,getOptions,onSelect,onInput,placeholderAll}){
  const input=document.getElementById(inputId),panel=document.getElementById(panelId),clearBtn=document.getElementById(clearId);
  if(!input||!panel)return null;
  let activeIndex=-1,currentOpts=[];
  function renderPanel(query){
    const all=getOptions();
    const q=(query||'').trim().toLowerCase();
    currentOpts=!q?all:all.filter(o=>o.sku.toLowerCase().includes(q)||o.label.toLowerCase().includes(q));
    activeIndex=-1;
    if(clearBtn)clearBtn.classList.toggle('show',!!input.value);
    if(!currentOpts.length){panel.innerHTML=`<div class="combo-empty">Sin coincidencias${q?` para “${escapeHtml(query)}”`:''}.</div>`;return;}
    const countLabel=q?`${currentOpts.length} de ${all.length} productos`:`${all.length} productos disponibles`;
    const rows=currentOpts.slice(0,60).map((o,i)=>`<div class="combo-option" data-idx="${i}" data-sku="${escapeHtml(o.sku)}"><span class="co-sku">${highlightMatch(o.sku,query)}</span><span class="co-name">${highlightMatch(o.label,query)}</span></div>`).join('');
    panel.innerHTML=`<div class="combo-count">${countLabel}</div>${rows}`;
  }
  function open(query){renderPanel(query!==undefined?query:input.value);panel.classList.add('open');}
  function close(){panel.classList.remove('open');activeIndex=-1;}
  function setActive(idx){const opts=[...panel.querySelectorAll('.combo-option')];opts.forEach(o=>o.classList.remove('active'));if(opts[idx]){opts[idx].classList.add('active');opts[idx].scrollIntoView({block:'nearest'});}activeIndex=idx;}
  function selectOpt(opt){if(!opt)return;input.value=`${opt.sku} · ${opt.label}`;if(clearBtn)clearBtn.classList.add('show');close();onSelect&&onSelect(opt);}
  input.addEventListener('input',()=>{open(input.value);onInput&&onInput(input.value);});
  input.addEventListener('focus',()=>open(input.value));
  input.addEventListener('keydown',e=>{
    if(!panel.classList.contains('open')&&e.key==='ArrowDown'){open(input.value);return;}
    if(!panel.classList.contains('open')&&e.key==='Enter'){open(input.value);}
    const opts=[...panel.querySelectorAll('.combo-option')];
    if(e.key==='ArrowDown'){e.preventDefault();setActive(Math.min(activeIndex+1,opts.length-1));}
    else if(e.key==='ArrowUp'){e.preventDefault();setActive(Math.max(activeIndex-1,0));}
    else if(e.key==='Enter'){e.preventDefault();const pick=(activeIndex>=0&&currentOpts[activeIndex])?currentOpts[activeIndex]:currentOpts[0];if(pick)selectOpt(pick);else close();}
    else if(e.key==='Escape'){close();}
  });
  panel.addEventListener('mousedown',e=>{const item=e.target.closest('.combo-option');if(!item)return;e.preventDefault();const opt=currentOpts.find(o=>o.sku===item.dataset.sku);selectOpt(opt);});
  if(clearBtn)clearBtn.addEventListener('click',()=>{input.value='';clearBtn.classList.remove('show');onInput&&onInput('');onSelect&&onSelect(null);close();input.focus();});
  document.addEventListener('click',e=>{if(!input.contains(e.target)&&!panel.contains(e.target)&&e.target!==clearBtn)close();});
  return {refresh:()=>{if(clearBtn)clearBtn.classList.toggle('show',!!input.value);},setValue:(text)=>{input.value=text;if(clearBtn)clearBtn.classList.toggle('show',!!text);},close};
}
function b2bCurvePoint(row,qty){
  const curve=Array.isArray(row?.curva_b2b)?row.curva_b2b:[];
  if(!curve.length)return null;
  const q=Math.max(1,Math.min(Math.floor(n(qty)),curve.length));
  return curve[q-1]||null;
}
function b2bEquivalentRoi(row,qty){const p=b2bCurvePoint(row,qty);return p?n(p.roi_b2b_equivalente_hoy):0;}
function b2bSkuOptions(){return b2bRows().map(r=>({sku:r.sku_madre,label:r.producto_madre||'Producto'}));}
let b2bSkuCombo=null,b2bSearchCombo=null;
function b2bDefaultQty(row){const qmax=Math.max(0,(row?.curva_b2b||[]).length);return qmax?Math.max(1,Math.min(qmax,Math.round(qmax/2))):1;}
function ensureB2BCombos(){
  if(!b2bSkuCombo)b2bSkuCombo=initCombo({inputId:'b2bSku',panelId:'b2bSkuPanel',clearId:'b2bSkuClear',getOptions:b2bSkuOptions,onSelect:(opt)=>{if(!opt)return;state.selectedB2BSku=opt.sku;const row=selectedB2BRow();if(row){const q=b2bDefaultQty(row);const qInput=document.getElementById('b2bQty');if(qInput)qInput.value=q;const p=document.getElementById('b2bProposedRoi');if(p)p.value='';}renderB2BDetail(false);}});
  if(!b2bSearchCombo)b2bSearchCombo=initCombo({inputId:'b2bSearch',panelId:'b2bSearchPanel',clearId:'b2bSearchClear',getOptions:b2bSkuOptions,onInput:(val)=>{state.pageQueries.b2b=val;renderB2B();},onSelect:(opt)=>{state.pageQueries.b2b=opt?`${opt.sku} · ${opt.label}`:'';renderB2B();}});
}
function populateB2BControls(){ensureB2BCombos();const rows=b2bRows();const input=document.getElementById('b2bSku');if(!input)return;if(rows.length){if(!state.selectedB2BSku||!rows.some(r=>r.sku_madre===state.selectedB2BSku))state.selectedB2BSku=rows[0].sku_madre;const row=selectedB2BRow();if(row){b2bSkuCombo.setValue(`${row.sku_madre} · ${row.producto_madre||'Producto'}`);const q=b2bDefaultQty(row);const qty=document.getElementById('b2bQty');if(qty&&!qty.value)qty.value=q;}}else{b2bSkuCombo.setValue('');}}

function distributionSummaryRows(){return DATA.distribution_summary||[];}
function distributionDetailRows(sku=''){const key=String(sku||'').toUpperCase();return (DATA.distribution_detail||[]).filter(r=>!key||String(r.sku_madre||'').toUpperCase()===key);}
function renderDistribution(){
  const q=(document.getElementById('distributionSearch')?.value||state.pageQueries.distribution||'').toLowerCase().trim();
  const method=document.getElementById('distributionMethod')?.value||'';
  let rows=distributionSummaryRows().filter(r=>(!q||`${r.sku_madre} ${r.producto_madre} ${r.categoria} ${r.marca}`.toLowerCase().includes(q))&&(!method||r.metodo_reparticion===method));
  const all=distributionSummaryRows(),pieces=safeSum(all,'stock_repartir'),exist=safeSum(all,'stock_existencias'),b2b=safeSum(all,'stock_b2b');
  const direct=all.filter(r=>r.metodo_reparticion==='Histórico SKU').length,category=all.filter(r=>String(r.metodo_reparticion).includes('Categoría')).length,equal=all.filter(r=>r.metodo_reparticion==='Parejo').length;
  document.getElementById('distributionStatus').innerHTML=[
    statusCard(fmt.format(pieces),'Piezas disponibles para repartir',COLORS.blue,'⇉'),
    statusCard(fmt.format(all.length),'IQ con stock central',COLORS.purple,'#'),
    statusCard(fmt.format(exist),'CUATI / Existencias',COLORS.green,'O'),
    statusCard(fmt.format(b2b),'CUATI / B2B',COLORS.amber,'B'),
    statusCard(`${fmt.format(direct)} / ${fmt.format(category)} / ${fmt.format(equal)}`,'Histórico / categoría / parejo',COLORS.teal,'≋')
  ].join('');
  state.tables.distribution=rows;
  renderTable('distributionTable',rows,[
    {key:'sku_madre',label:'IQ',format:'sku'},
    {key:'producto_madre',label:'Producto',format:'product'},
    {key:'categoria',label:'Categoría'},
    {key:'marca',label:'Marca'},
    {key:'stock_existencias',label:'Existencias',format:'num'},
    {key:'stock_b2b',label:'B2B',format:'num'},
    {key:'stock_repartir',label:'A repartir',format:'numStrong'},
    {key:'metodo_reparticion',label:'Método',format:'badge'},
    {key:'confianza',label:'Confianza',format:'badge'},
    {key:'ventas_sku_90d',label:'Ventas IQ 90d',format:'num'},
    {key:'muestra_skus',label:'IQ muestra',format:'num'},
    {key:'sku_madre',label:'Detalle',format:'distributionButton'}
  ],{pageSize:25,id:'distribution',searchInputId:'distributionSearch',defaultSortKey:'stock_repartir'});
  if(state.selectedDistributionSku)renderDistributionDetail();
}
function openDistributionDetail(sku){state.selectedDistributionSku=String(sku||'').toUpperCase();const p=document.getElementById('distributionDetailPanel');p?.classList.add('open');renderDistributionDetail();p?.scrollIntoView({behavior:'smooth',block:'start'});}
function closeDistributionDetail(){state.selectedDistributionSku='';document.getElementById('distributionDetailPanel')?.classList.remove('open');}
function exportDistributionDetail(){downloadCsv(distributionDetailRows(state.selectedDistributionSku),`reparticion_${state.selectedDistributionSku||'detalle'}`);}
function renderDistributionDetail(){
  const sku=state.selectedDistributionSku;if(!sku)return;
  const summary=distributionSummaryRows().find(r=>String(r.sku_madre||'').toUpperCase()===sku);const rows=distributionDetailRows(sku);if(!summary||!rows.length)return;
  document.getElementById('distributionDetailTitle').textContent=`Repartición de ${sku}`;
  document.getElementById('distributionDetailSubtitle').textContent=`${summary.producto_madre||''}${summary.categoria?` · ${summary.categoria}`:''}${summary.marca?` · ${summary.marca}`:''}`;
  document.getElementById('distributionDetailStatus').innerHTML=[
    statusCard(fmt.format(summary.stock_repartir),'Piezas a repartir',COLORS.blue,'⇉'),
    statusCard(fmt.format(summary.stock_existencias),'Desde Existencias',COLORS.green,'O'),
    statusCard(fmt.format(summary.stock_b2b),'Desde B2B',COLORS.amber,'B'),
    statusCard(summary.metodo_reparticion||'—','Método',COLORS.purple,'ƒ'),
    statusCard(summary.confianza||'—','Confianza de evidencia',summary.confianza==='ALTA'?COLORS.green:(summary.confianza==='MEDIA'?COLORS.amber:COLORS.red),'✓')
  ].join('');
  const channels=rows.map(r=>r.canal);
  plot('distributionChart',[
    {type:'bar',name:'Stock actual',x:channels,y:rows.map(r=>n(r.stock_actual_canal)),marker:{color:COLORS.slate}},
    {type:'bar',name:'Stock después',x:channels,y:rows.map(r=>n(r.stock_despues)),marker:{color:COLORS.blue}}
  ],{barmode:'group',yaxis:{title:'Piezas',rangemode:'tozero'},xaxis:{tickangle:-18},margin:{l:50,r:20,t:15,b:70}});
  const basePct=Math.round(n(summary.peso_base_igual)*100);
  document.getElementById('distributionMethodology').innerHTML=[
    insight('Fuente de cantidad',`Se toman ${fmt.format(summary.stock_repartir)} piezas disponibles de CUATI/Existencias + CUATI/B2B. Los arribos solo se muestran como referencia.`,COLORS.blue,'1'),
    insight('Jerarquía de datos',summary.fundamento||'Histórico SKU → categoría + marca → categoría → marca → parejo.',COLORS.purple,'2'),
    insight('Exploración controlada',summary.metodo_reparticion==='Parejo'?'Sin datos útiles: se reparte el stock nuevo en partes iguales.':`${basePct}% del peso se reserva como base igual y el resto sigue el desempeño; además se considera el stock que ya tiene cada canal.`,COLORS.teal,'3'),
    insight('Resultado',`La propuesta suma exactamente ${fmt.format(summary.stock_repartir)} piezas y no descuenta stock de los canales actuales; solo asigna la bolsa central disponible.`,COLORS.green,'4')
  ].join('');
  state.tables.distributionDetail=rows;
  renderTable('distributionDetailTable',rows,[
    {key:'canal',label:'Canal'},
    {key:'stock_actual_canal',label:'Stock actual',format:'num'},
    {key:'ventas_10d_sku',label:'Ventas 10d',format:'num'},
    {key:'ventas_30d_sku',label:'Ventas 30d',format:'num'},
    {key:'ventas_90d_sku',label:'Ventas 90d',format:'num'},
    {key:'participacion_objetivo',label:'Peso objetivo',format:'percent'},
    {key:'cantidad_sugerida',label:'Enviar',format:'numStrong'},
    {key:'desde_existencias',label:'Desde Exist.',format:'num'},
    {key:'desde_b2b',label:'Desde B2B',format:'num'},
    {key:'stock_despues',label:'Stock después',format:'num'},
    {key:'cobertura_despues_dias',label:'Cobertura post',format:'coverage'},
    {key:'sku_madre',label:'Odoo',format:'distributionDraftButton'}
  ],{pageSize:10,id:'distributionDetail',defaultSortKey:'cantidad_sugerida'});
}

function renderB2B(){
  ensureB2BCombos();
  const rawQ=(document.getElementById('b2bSearch')?.value||state.pageQueries.b2b||'');
  const q=rawQ.split('·')[0].trim().toLowerCase();
  let rows=(DATA.b2b_portfolio||[]).filter(r=>!q||`${r.sku_madre} ${r.producto_madre}`.toLowerCase().includes(q));
  state.tables.b2b=rows;
  const eligible=rows.filter(r=>n(r.b2b_disponible)>0),units=safeSum(eligible,'b2b_disponible'),valued=safeSum(eligible,'excedente_valorado_vpn');
  const ready=eligible.filter(r=>(r.curva_b2b||[]).length>0).length,pending=eligible.length-ready;
  const capital=eligible.reduce((a,r)=>a+n(r.b2b_disponible)*n(r.costo_unitario_odoo),0);
  const live=DATA.meta.b2b_odoo_live||{};
  document.getElementById('b2bStatus').innerHTML=[
    statusCard(fmt.format(eligible.length),'SKU con excedente',COLORS.blue,'$'),
    statusCard(fmt.format(units),'Piezas excedentes',COLORS.teal,'▦'),
    statusCard(fmt.format(valued),'Piezas valoradas VPN',COLORS.green,'✓'),
    statusCard(money.format(capital),'Costo del excedente',COLORS.purple,'$'),
    statusCard(escapeHtml(live.status||'BASE'),'Stock B2B · '+escapeHtml(live.source||'dashboard'),live.status==='CONECTADO'?COLORS.green:COLORS.amber,'O')
  ].join('');
  renderTable('b2bTable',rows,[
    {key:'sku_madre',label:'IQ',format:'sku'},
    {key:'producto_madre',label:'Producto',format:'product'},
    {key:'stock_odoo',label:'Stock bodega',format:'numStrong'},
    {key:'reserva_marketplace_30d',label:'Reserva MKT',format:'num'},
    {key:'b2b_disponible',label:'Excedente B2B',format:'numStrong'},
    {key:'excedente_valorado_vpn',label:'Valorado VPN',format:'num'},
    {key:'roi_b2b_minimo_total',label:'ROI mín. B2B · todo excedente',format:'percent'},
    {key:'precio_minimo_b2b_total',label:'Precio mín. / pza',format:'money2'},
    {key:'costo_unitario_odoo',label:'Costo Odoo',format:'money2'},
    {key:'fuente_stock',label:'Fuente stock',format:'badge'},
    {key:'estado_b2b',label:'Estado',format:'badge'},
    {key:'sku_madre',label:'Detalle',format:'b2bButton'}
  ],{pageSize:18,id:'b2b',searchInputId:'b2bSearch'});
  populateB2BControls();
  renderB2BDetail(false);
}
function openB2BSimulator(sku){ensureB2BCombos();state.selectedB2BSku=String(sku||'').toUpperCase();const row=selectedB2BRow();if(row){b2bSkuCombo.setValue(`${row.sku_madre} · ${row.producto_madre||'Producto'}`);const q=b2bDefaultQty(row);const qty=document.getElementById('b2bQty');if(qty)qty.value=q;const proposed=document.getElementById('b2bProposedRoi');if(proposed)proposed.value='';}renderB2BDetail(false);document.getElementById('b2bDetailShell')?.scrollIntoView({behavior:'smooth',block:'start'});}
function syncB2BQtyFromRange(value){const qty=document.getElementById('b2bQty');if(qty)qty.value=value;renderB2BSimulator(false);}
function renderB2BDetail(fromUser=false){
  const row=selectedB2BRow();
  const head=document.getElementById('b2bDetailHeader'),kpis=document.getElementById('b2bKpiGrid'),channels=document.getElementById('b2bChannelDetail');
  if(!row){if(head)head.innerHTML='<div class="history-empty">No hay candidatos B2B disponibles.</div>';if(kpis)kpis.innerHTML='';if(channels)channels.innerHTML='<div class="history-empty">Sin detalle.</div>';return;}
  state.selectedB2BSku=row.sku_madre;
  const live=String(row.odoo_live_status||DATA.meta.b2b_odoo_live?.status||'BASE').toUpperCase();
  const liveClass=live==='CONECTADO'?'live':'base';
  if(head)head.innerHTML=`<div class="b2b-detail-head"><div class="b2b-detail-title"><h3>${escapeHtml(row.sku_madre)} · ${escapeHtml(row.producto_madre||'Producto')}</h3><p>${escapeHtml(row.estado_b2b||'')} · costo: ${n(row.costo_unitario_odoo)>0?money2.format(n(row.costo_unitario_odoo)):'—'} (${escapeHtml(row.fuente_costo||'sin costo')}) · stock: ${escapeHtml(row.fuente_stock||'')}</p></div><span class="b2b-live-badge ${liveClass}">${escapeHtml(live==='CONECTADO'?'Odoo API live':live)}</span></div>`;
  const roiTotal=row.roi_b2b_minimo_total??row.roi_b2b_equivalente_hoy_total;
  const roiClass=n(roiTotal)<0?'negative':'positive';
  if(kpis)kpis.innerHTML=`
    <div class="b2b-kpi-card"><b>${fmt.format(n(row.stock_odoo))}</b><span>Stock bodega</span></div>
    <div class="b2b-kpi-card"><b>${fmt.format(n(row.reserva_marketplace_30d))}</b><span>Reserva Marketplace</span></div>
    <div class="b2b-kpi-card"><b>${fmt.format(n(row.b2b_disponible))}</b><span>Excedente B2B</span></div>
    <div class="b2b-kpi-card"><b>${fmt.format(n(row.excedente_valorado_vpn))}</b><span>Excedente valorado VPN</span></div>
    <div class="b2b-kpi-card ${roiClass}"><b>${roiTotal===null||roiTotal===''?'—':pct2(n(roiTotal))}</b><span>ROI mínimo B2B · todo excedente</span></div>
    <div class="b2b-kpi-card"><b>${n(row.costo_unitario_odoo)>0&&row.precio_minimo_b2b_total!==null?money2.format(n(row.precio_minimo_b2b_total)):'—'}</b><span>Precio mínimo / pza</span></div>`;
  renderB2BChannelDetail(row);
  renderB2BSimulator(fromUser);
}
function renderB2BChannelDetail(row){
  const el=document.getElementById('b2bChannelDetail');if(!el)return;
  const rows=Array.isArray(row.canales_b2b)?row.canales_b2b:[];
  if(!rows.length){el.innerHTML='<div class="history-empty">Este SKU no tiene detalle Marketplace calculable.</div>';return;}
  el.innerHTML=`<table><thead><tr><th>Canal</th><th class="num">Forecast/día</th><th class="num">Necesidad LT</th><th class="num">Full</th><th class="num">Tránsito</th><th class="num">Reserva</th><th class="num">ROI MP</th><th class="num">Payout</th><th class="num">1a venta excedente</th><th class="num">ROI VPN 1a pza</th><th class="num">Pzas excedente asignadas</th><th class="num">ROI VPN prom.</th></tr></thead><tbody>${rows.map(c=>`<tr><td>${escapeHtml(c.canal)}</td><td class="num">${fmt1.format(n(c.forecast_dia))}</td><td class="num">${fmt.format(n(c.necesidad_lead_time))}</td><td class="num">${fmt.format(n(c.full))}</td><td class="num">${fmt.format(n(c.transito))}</td><td class="num"><b>${fmt.format(n(c.reserva_bodega))}</b></td><td class="num">${c.roi_marketplace===null||c.roi_marketplace===''?'—':pct2(n(c.roi_marketplace))}</td><td class="num ${c.payout_es_fallback?'warn':''}">${fmt1.format(n(c.payout_dias))}d${c.payout_es_fallback?' *':''}</td><td class="num">${c.dia_primera_venta_excedente??'—'}</td><td class="num">${c.roi_equivalente_hoy_primera_excedente===null||c.roi_equivalente_hoy_primera_excedente===''?'—':pct2(n(c.roi_equivalente_hoy_primera_excedente))}</td><td class="num">${fmt.format(n(c.piezas_excedente_asignadas_mp))}</td><td class="num">${c.roi_equivalente_hoy_prom_excedente===null||c.roi_equivalente_hoy_prom_excedente===''?'—':pct2(n(c.roi_equivalente_hoy_prom_excedente))}</td></tr>`).join('')}</tbody></table><div class="card-sub" style="margin-top:8px">* Payout de 30 días = supuesto temporal para canales con dato incompleto.</div>`;
}
function renderB2BSimulator(fromUser=true){
  const row=selectedB2BRow(),box=document.getElementById('b2bSimulatorResult');if(!box)return;
  const chart=document.getElementById('b2bCurveChart'),offerBox=document.getElementById('b2bOfferVerdict');
  if(!row){box.innerHTML='<div class="history-empty">No hay candidatos B2B disponibles.</div>';if(chart)chart.innerHTML='<div class="history-empty">Sin curva.</div>';return;}
  const curve=Array.isArray(row.curva_b2b)?row.curva_b2b:[],qmax=curve.length;
  const range=document.getElementById('b2bQtyRange'),qtyInput=document.getElementById('b2bQty');
  if(range){range.max=Math.max(qmax,1);range.disabled=qmax<=0;}
  if(qtyInput){qtyInput.max=Math.max(qmax,1);qtyInput.disabled=qmax<=0;}
  if(qmax<=0){box.innerHTML=`<div class="b2b-sim-output"><div class="card-sub"><b>${escapeHtml(row.estado_b2b||'Sin curva VPN calculable')}</b> · Excedente físico: ${fmt.format(row.b2b_disponible)} piezas, valoradas: ${fmt.format(row.excedente_valorado_vpn)}.</div></div>`;if(chart)chart.innerHTML='<div class="history-empty">No hay suficientes datos de ROI / forecast para construir la curva.</div>';if(offerBox)offerBox.textContent='';return;}
  let qty=Math.floor(n(qtyInput?.value||b2bDefaultQty(row)));qty=Math.max(1,Math.min(qty,qmax));
  if(qtyInput)qtyInput.value=qty;if(range)range.value=qty;
  const point=b2bCurvePoint(row,qty),reference=n(point?.roi_b2b_minimo??point?.roi_b2b_equivalente_hoy),economic=n(point?.roi_b2b_economico_vpn),marginal=n(point?.roi_oportunidad_marginal),marginalMin=n(point?.roi_b2b_marginal_minimo??Math.max(reference,marginal)),cost=n(row.costo_unitario_odoo),price=point?.precio_minimo_b2b_unitario!==null&&point?.precio_minimo_b2b_unitario!==undefined?n(point.precio_minimo_b2b_unitario):(cost>0?cost*(1+reference):0);
  box.innerHTML=`<div class="b2b-sim-output"><div class="b2b-sim-metrics"><div class="b2b-sim-metric"><b>${fmt.format(qty)}</b><span>Piezas B2B del excedente</span></div><div class="b2b-sim-metric"><b>${pct2(reference)}</b><span>ROI mínimo B2B</span></div><div class="b2b-sim-metric"><b>${pct2(economic)}</b><span>Referencia económica VPN</span></div><div class="b2b-sim-metric"><b>${cost>0?money2.format(price):'—'}</b><span>Precio mínimo / pza</span></div></div><div class="b2b-marginal-note">Pieza marginal: costo de oportunidad ${pct2(marginal)} → mínimo comercial ${pct2(marginalMin)}. Referencia: <b>${escapeHtml(point?.canal_oportunidad||'—')}</b>${point?.roi_marketplace_referencia===null||point?.roi_marketplace_referencia===undefined?'':`, ROI Marketplace ${pct2(n(point.roi_marketplace_referencia))}`}${point?.dia_venta_marketplace?`, venta estimada día ${point.dia_venta_marketplace}`:''}${point?.dia_cobro_marketplace?`, cobro día ${point.dia_cobro_marketplace}`:''}.</div></div>`;
  const proposedRaw=document.getElementById('b2bProposedRoi')?.value||'',hasProposed=String(proposedRaw).trim()!=='';
  if(offerBox){if(!hasProposed){offerBox.innerHTML='Ingresa un ROI de oferta si quieres compararlo contra el mínimo comercial calculado.';}else{const proposed=n(proposedRaw)/100,covers=proposed+1e-9>=reference;offerBox.innerHTML=`${badge(covers?'CUMPLE MÍNIMO B2B':'POR DEBAJO DEL MÍNIMO B2B')} · Oferta ${pct2(proposed)} vs mínimo ${pct2(reference)}.`;}}
  drawB2BCurve(row,qty);
  renderB2BExplanation(row,qty,point,price);
}
function drawB2BCurve(row,qtySel){
  const box=document.getElementById('b2bCurveChart');if(!box)return;const curve=Array.isArray(row.curva_b2b)?row.curva_b2b:[];
  if(!curve.length){box.innerHTML='<div class="history-empty">Sin curva VPN calculable.</div>';return;}
  const W=720,H=280,pL=52,pR=22,pT=22,pB=38,floor=n(DATA.meta.params.b2b_roi_minimum||0.10);const vals=curve.map(x=>n(x.roi_b2b_minimo??x.roi_b2b_equivalente_hoy));let minY=Math.min(0,...vals),maxY=Math.max(floor,...vals);if(Math.abs(maxY-minY)<0.01){maxY+=0.01;minY-=0.01;}const pad=(maxY-minY)*0.08;maxY+=pad;minY-=pad;const span=maxY-minY;
  const xPos=q=>pL+(W-pL-pR)*(q-1)/Math.max(curve.length-1,1);const yPos=v=>pT+(H-pT-pB)*(maxY-v)/span;const pts=curve.map(x=>`${xPos(x.q)},${yPos(n(x.roi_b2b_minimo??x.roi_b2b_equivalente_hoy))}`).join(' ');const floorY=yPos(floor);const sel=curve[Math.max(0,Math.min(qtySel-1,curve.length-1))],selRoi=n(sel.roi_b2b_minimo??sel.roi_b2b_equivalente_hoy),cx=xPos(sel.q),cy=yPos(selRoi);
  box.innerHTML=`<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none"><line class="b2b-curve-axis" x1="${pL}" y1="${H-pB}" x2="${W-pR}" y2="${H-pB}"/><line class="b2b-curve-axis" x1="${pL}" y1="${pT}" x2="${pL}" y2="${H-pB}"/><line class="b2b-curve-zero" x1="${pL}" y1="${floorY}" x2="${W-pR}" y2="${floorY}"/><polyline class="b2b-curve-line" points="${pts}"/><circle class="b2b-curve-point" cx="${cx}" cy="${cy}" r="4"/><text class="b2b-curve-label" x="${pL}" y="${H-12}">1</text><text class="b2b-curve-label" x="${W-pR-24}" y="${H-12}">${curve.length}</text><text class="b2b-curve-label" x="5" y="${pT+5}">${(maxY*100).toFixed(1)}%</text><text class="b2b-curve-label" x="5" y="${H-pB+4}">${(minY*100).toFixed(1)}%</text><text class="b2b-curve-label" x="${Math.min(cx+8,W-100)}" y="${Math.max(cy-8,16)}">${(selRoi*100).toFixed(2)}%</text></svg>`;
}
function renderB2BExplanation(row,qty,point,price){
  const el=document.getElementById('b2bExplanation');if(!el)return;
  const rate=n(DATA.meta.params.b2b_daily_discount_rate||0.000355),h=n(DATA.meta.params.b2b_horizon_days||30),floor=n(DATA.meta.params.b2b_roi_minimum||0.10),cost=n(row.costo_unitario_odoo),ref=point?n(point.roi_b2b_minimo??point.roi_b2b_equivalente_hoy):floor,econ=point?n(point.roi_b2b_economico_vpn):0,marginal=point?n(point.roi_oportunidad_marginal):0,channel=point?.canal_oportunidad||'—',roiMp=point?.roi_marketplace_referencia===null||point?.roi_marketplace_referencia===undefined?null:n(point.roi_marketplace_referencia),daySale=point?.dia_venta_marketplace??'—',dayCash=point?.dia_cobro_marketplace??'—';
  el.innerHTML=`<div class="card-head" style="margin-bottom:6px"><div><div class="card-title">¿De dónde sale el mínimo B2B?</div><div class="card-sub">Trazabilidad financiera del pedido de ${fmt.format(qty)} piezas.</div></div><button class="card-action" id="b2bCloseExplain" type="button">Cerrar</button></div><div class="b2b-steps"><div class="b2b-step"><b>1. Marketplace se protege primero</b><p>Reserva por canal = Forecast × ${h} días − Full − tránsito. Reserva total de bodega: ${fmt.format(row.reserva_marketplace_30d)} piezas.</p></div><div class="b2b-step"><b>2. Stock B2B desde Odoo</b><p>Stock bodega: ${fmt.format(row.stock_odoo)}. Excedente B2B: ${fmt.format(row.b2b_disponible)}. Fuente: ${escapeHtml(row.fuente_stock||'')}.</p></div><div class="b2b-step"><b>3. Alternativa Marketplace</b><p>${roiMp===null?`La pieza marginal no tiene referencia Marketplace completa; se conserva el piso comercial.`:`La pieza marginal sustituye ${escapeHtml(channel)} con ROI ${pct2(roiMp)}, venta estimada día ${daySale} y cobro día ${dayCash}.`}</p></div><div class="b2b-step"><b>4. Valor presente</b><p>ROI VPN = (1 + ROI Marketplace) / (1 + tasa diaria)^días de cobro − 1. Tasa diaria: ${(rate*100).toFixed(4)}%. Cuando el payout es incompleto se usan 30 días.</p></div><div class="b2b-step"><b>5. Piso comercial</b><p>Por pieza: ROI mínimo B2B = max(${pct2(floor)}, ROI oportunidad VPN). Para ${fmt.format(qty)} piezas, referencia económica promedio ${pct2(econ)} y mínimo comercial promedio <b>${pct2(ref)}</b>. Costo oportunidad marginal: ${pct2(marginal)}.</p></div><div class="b2b-step"><b>6. Precio mínimo</b><p>${cost>0?`Costo Odoo ${money2.format(cost)} × (1 + ${pct2(ref)}) = <b>${money2.format(price)}</b> por pieza.`:'El ROI mínimo está calculado, pero falta costo Odoo para convertirlo a precio.'}</p></div></div><div class="formula" style="margin-top:12px">ROI mínimo pieza = max(10%, ROI oportunidad VPN) · ROI mínimo pedido = promedio(ROI mínimo pieza) · Precio mínimo = Costo × (1 + ROI mínimo pedido)</div>`;
  document.getElementById('b2bCloseExplain')?.addEventListener('click',()=>el.classList.remove('open'));
}
function toggleB2BExplanation(){document.getElementById('b2bExplanation')?.classList.toggle('open');}

function renderTransfers(){const rows=filterText(DATA.transfers,document.getElementById('transferSearch')?.value,['sku_madre','producto_madre','canal']);state.tables.transfers=rows;const channels=new Set(rows.map(r=>r.canal)).size,total=rows.reduce((a,r)=>a+n(r.transferencia_sugerida),0),odoo=rows.reduce((a,r)=>a+n(r.inventario_odoo),0),sales=rows.reduce((a,r)=>a+n(r.ventas_90d_unidades),0);document.getElementById('transferStatus').innerHTML=[statusCard(fmt.format(rows.length),'SKU-canal sugeridos',COLORS.amber,'⇄'),statusCard(fmt.format(total),'Piezas a transferir',COLORS.amber,'▦'),statusCard(fmt.format(channels),'Canales involucrados',COLORS.blue,'#'),statusCard(fmt.format(sales),'Ventas 90 días asociadas',COLORS.teal,'↗')].join('');const agg=groupSum(rows,'canal','transferencia_sugerida');plot('chartTransfers',[{type:'bar',x:agg.map(r=>r.key),y:agg.map(r=>r.value),marker:{color:agg.map(r=>CHANNEL_COLORS[r.key]||COLORS.amber)},text:agg.map(r=>fmt.format(r.value)),textposition:'outside'}],{yaxis:{title:'Unidades',rangemode:'tozero'}});document.getElementById('transferInsights').innerHTML=[insight('Primero se revisa Odoo',`Las piezas sugeridas nunca superan el inventario Odoo disponible (${fmt.format(odoo)} unidades en los registros mostrados).`,COLORS.purple,'O'),insight('Cobertura al arribo',`Se proyectan ${DATA.meta.params.full_transit_days||5} días de consumo y se busca conservar ${DATA.meta.params.full_target_days||30} días al llegar. Productos nuevos: máximo ${DATA.meta.params.new_product_full_max||10} piezas y check manual.`,COLORS.amber,'◷'),insight('No implica compra',`Una transferencia corrige la ubicación del inventario. El total de la empresa ya contiene las piezas.`,COLORS.green,'✓')].join('');renderTable('transferTable',rows,[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'canal',label:'Canal'},{key:'inventario_full',label:'Full',format:'num'},{key:'inventario_transito',label:'Tránsito',format:'num'},{key:'inventario_odoo',label:'Odoo',format:'num'},{key:'ventas_full_90d',label:'Ventas Full 90d',format:'num'},{key:'cobertura_full_dias',label:'Cob. Full',format:'coverage'},{key:'objetivo_full_piezas',label:'Objetivo Full',format:'num'},{key:'transferencia_sugerida',label:'Transferir',format:'numStrong'},{key:'check_manual_full',label:'Validación',format:'badge'},{key:'nivel_riesgo',label:'Riesgo',format:'badge'},{key:'justificacion',label:'Justificación',format:'product'},{key:'sku_madre',label:'Admin',format:'adminTransferButton'}],{pageSize:25,id:'transfers',searchInputId:'transferSearch',defaultSortKey:'transferencia_sugerida',defaultSortDir:'desc'});}
function renderPurchases(){
  const query=document.getElementById('purchaseSearch')?.value||'';
  const consolidated=filterText(DATA.purchases_consolidated||[],query,['sku_madre','producto_madre','canales']);
  const detail=filterText(DATA.purchases||[],query,['sku_madre','producto_madre','canal']);
  state.tables.purchasesConsolidated=consolidated;
  state.tables.purchases=detail;

  const total=consolidated.reduce((a,r)=>a+n(r.compra_sugerida),0);
  const critical=consolidated.filter(r=>r.nivel_riesgo==='CRÍTICO').length;
  const inventory=consolidated.reduce((a,r)=>a+n(r.inventario_total),0);
  const sales=consolidated.reduce((a,r)=>a+n(r.ventas_90d_unidades),0);
  document.getElementById('purchaseStatus').innerHTML=[
    statusCard(fmt.format(consolidated.length),'SKU a comprar',COLORS.red,'◫'),
    statusCard(fmt.format(total),'Piezas consolidadas',COLORS.red,'▦'),
    statusCard(fmt.format(critical),'Casos críticos',COLORS.red,'!'),
    statusCard(fmt.format(sales),'Ventas 90 días',COLORS.blue,'↗')
  ].join('');

  const top=[...consolidated].sort((a,b)=>n(b.compra_sugerida)-n(a.compra_sugerida)).slice(0,15).reverse();
  plot('chartPurchases',[{
    type:'bar',orientation:'h',
    y:top.map(r=>r.sku_madre),
    x:top.map(r=>n(r.compra_sugerida)),
    text:top.map(r=>fmt.format(r.compra_sugerida)),
    textposition:'outside',marker:{color:COLORS.red},
    hovertext:top.map(r=>r.producto_madre||''),
    hovertemplate:'%{y}<br>%{hovertext}<br>Comprar: %{x:,.0f}<extra></extra>'
  }],{margin:{l:72,r:38,t:10,b:40},xaxis:{title:'Piezas a comprar',rangemode:'tozero'}});

  document.getElementById('purchaseInsights').innerHTML=[
    insight('Una orden por SKU',`La cifra principal ya no suma déficits independientes por canal. Se agregan ventas e inventario de todas las bolsas del IQ antes de calcular la compra.`,COLORS.blue,'▦'),
    insight('Objetivo configurable',`Compra consolidada = máx(ritmo diario total × ${DATA.meta.params.purchase_coverage} días − inventario total empresa, 0). Lead time proveedor: ${DATA.meta.params.lead_time_supplier} días.`,COLORS.red,'◷'),
    insight('Detalle por canal',`La segunda tabla se conserva para diagnóstico: ayuda a identificar dónde falta inventario, pero no debe sumarse para generar la OC.`,COLORS.amber,'i')
  ].join('');

  renderTable('purchaseConsolidatedTable',consolidated,[
    {key:'sku_madre',label:'SKU',format:'sku'},
    {key:'producto_madre',label:'Producto',format:'product'},
    {key:'canales',label:'Canales',format:'product'},
    {key:'inventario_odoo',label:'Odoo',format:'num'},
    {key:'inventario_transito',label:'Tránsito',format:'num'},
    {key:'inventario_full',label:'Full',format:'num'},
    {key:'inventario_total',label:'Inventario empresa',format:'numStrong'},
    {key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},
    {key:'venta_diaria_total',label:'Ritmo total u/día',format:'num1'},
    {key:'cobertura_total_dias',label:'Cobertura',format:'coverage'},
    {key:'objetivo_stock_piezas',label:'Objetivo',format:'num'},
    {key:'compra_sugerida',label:'Comprar',format:'numStrong'},
    {key:'utilidad_30d',label:'Utilidad 30d',format:'money2'},
    {key:'roi_30d',label:'ROI 30d',format:'percent'},
    {key:'utilidad_90d',label:'Utilidad 90d',format:'money2'},
    {key:'roi_90d',label:'ROI 90d',format:'percent'},
    {key:'nivel_riesgo',label:'Riesgo',format:'badge'},
    {key:'justificacion',label:'Justificación',format:'product'}
  ],{pageSize:25,id:'purchasesConsolidated',searchInputId:'purchaseSearch',defaultSortKey:'compra_sugerida',defaultSortDir:'desc'});

  renderTable('purchaseTable',detail,[
    {key:'sku_madre',label:'SKU',format:'sku'},
    {key:'producto_madre',label:'Producto',format:'product'},
    {key:'canal',label:'Canal'},
    {key:'inventario_total',label:'Inventario total',format:'num'},
    {key:'ventas_90d_unidades',label:'Ventas 90d',format:'num'},
    {key:'cobertura_total_dias',label:'Cob. Total',format:'coverage'},
    {key:'compra_sugerida',label:'Necesidad canal',format:'numStrong'},
    {key:'demanda_tipo',label:'Demanda',format:'badge'},
    {key:'nivel_riesgo',label:'Riesgo',format:'badge'},
    {key:'justificacion',label:'Justificación',format:'product'}
  ],{pageSize:25,id:'purchases',searchInputId:'purchaseSearch',defaultSortKey:'compra_sugerida',defaultSortDir:'desc'});
}

function renderLots(){const buckets=DATA.age_buckets||[],fromLayers=DATA.age_source==='capas';plot('chartAge',[{type:'bar',x:buckets.map(r=>r.bucket),y:buckets.map(r=>n(r.unidades)),marker:{color:[COLORS.green,'#84cc16',COLORS.amber,'#f97316',COLORS.red,COLORS.slate]},text:buckets.map(r=>fromLayers?money.format(n(r.valor)):fmt.format(r.unidades)),textposition:'outside',customdata:buckets.map(r=>[r.skus,money.format(n(r.valor))]),hovertemplate:fromLayers?'%{x} días<br>Piezas: %{y:,.0f}<br>Inversión: %{customdata[1]}<br>SKUs: %{customdata[0]}<extra></extra>':'%{x} días<br>Unidades: %{y:,.0f}<br>SKUs: %{customdata[0]}<extra></extra>'}],{yaxis:{title:'Piezas',rangemode:'tozero'}});const sub=document.getElementById('chartAgeSub');if(sub)sub.textContent=fromLayers?'Piezas con stock actual por antigüedad real de cada capa (lote en Odoo, envío en Full). La etiqueta muestra la inversión.':'Unidades con stock actual, agrupadas por días desde el último arribo (base anterior).';const old=(DATA.rotation||[]).filter(r=>n(r.inventario_total)>0&&is90(r)).sort((a,b)=>n(b.valor_mas_90d)-n(a.valor_mas_90d)||n(b.dias_inventario_max)-n(a.dias_inventario_max));state.tables.oldProducts=old;renderTable('oldProductsTable',old,[{key:'sku_madre',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'canal',label:'Canal'},{key:'unidades_mas_90d',label:'Piezas +90',format:'num'},{key:'valor_mas_90d',label:'Inversión +90',format:'money'},{key:'dias_inventario_max',label:'Pieza más antigua',format:'daysAge'},{key:'movs_revision_90d',label:'Movs. revisión',format:'num'}],{pageSize:12,id:'oldProducts',defaultSortKey:'valor_mas_90d'});const rows=filterText(DATA.arrivals,document.getElementById('arrivalSearch')?.value,['orden_compra','proveedor','sku_odoo','producto']);state.tables.arrivals=rows;renderTable('arrivalsTable',rows,[{key:'fecha',label:'Fecha'},{key:'orden_compra',label:'Orden'},{key:'proveedor',label:'Proveedor'},{key:'sku_odoo',label:'SKU',format:'sku'},{key:'producto',label:'Producto',format:'product'},{key:'cantidad_arribada',label:'Arribado',format:'num'},{key:'cantidad_ordenada',label:'Ordenado',format:'num'},{key:'precio_unitario',label:'Precio',format:'money'},{key:'subtotal',label:'Subtotal',format:'money'},{key:'estado_compra',label:'Estado',format:'badge'}],{pageSize:25,id:'arrivals',searchInputId:'arrivalSearch'});}
function renderAlerts(){const q=(document.getElementById('alertSearch')?.value||'').toLowerCase(),risk=document.getElementById('alertRiskFilter')?.value||'';const rows=DATA.rotation.filter(r=>(r.nivel_riesgo!=='SANO')&&(!risk||r.nivel_riesgo===risk)&&(!q||`${r.sku_madre} ${r.producto_madre} ${r.canal}`.toLowerCase().includes(q)));state.tables.alerts=rows;const critical=rows.filter(r=>r.nivel_riesgo==='CRÍTICO').length,high=rows.filter(r=>r.nivel_riesgo==='ALTO').length,inter=rows.filter(r=>r.nivel_riesgo==='INTERMITENTE').length,over=rows.filter(r=>r.nivel_riesgo==='SOBRESTOCK').length;document.getElementById('alertStatus').innerHTML=[statusCard(fmt.format(critical),'Críticos',COLORS.red,'!'),statusCard(fmt.format(high),'Riesgo alto',COLORS.amber,'↑'),statusCard(fmt.format(inter),'Demanda intermitente',COLORS.purple,'≈'),statusCard(fmt.format(over),'Sobrestock',COLORS.blue,'▦')].join('');const channels=DATA.meta.channels;const levels=['CRÍTICO','ALTO','INTERMITENTE','SOBRESTOCK'];plot('chartAlertsChannel',levels.map(level=>({type:'bar',name:cap(level),x:channels,y:channels.map(c=>rows.filter(r=>r.canal===c&&r.nivel_riesgo===level).length),marker:{color:riskTone(level)}})),{barmode:'stack',yaxis:{title:'SKU-canal',rangemode:'tozero'},xaxis:{tickangle:-18}});const counts={};rows.forEach(r=>counts[r.nivel_riesgo]=(counts[r.nivel_riesgo]||0)+1);plot('chartRiskDistribution',[{type:'pie',hole:.58,labels:Object.keys(counts),values:Object.values(counts),marker:{colors:Object.keys(counts).map(riskTone)},textinfo:'label+percent'}],{margin:{l:10,r:10,t:10,b:40}});renderTable('alertsTable',rows,alertsColumns(state.views.alerts),{pageSize:25,id:'alerts',searchInputId:'alertSearch'});}
function renderQuality(){const q=DATA.quality;document.getElementById('qualityKpis').innerHTML=[`<div class="quality-kpi"><strong>${fmt.format(q.stock_no_vinculado_skus)}</strong><span>SKU de stock sin homologar</span></div>`,`<div class="quality-kpi"><strong>${fmt.format(q.stock_no_vinculado_unidades)}</strong><span>Unidades sin homologar</span></div>`,`<div class="quality-kpi"><strong>${fmt.format(q.ventas_no_vinculadas)}</strong><span>Ventas no vinculadas</span></div>`,`<div class="quality-kpi"><strong>${fmt.format(q.traslados_excluidos)}</strong><span>Traslados excluidos</span></div>`,`<div class="quality-kpi"><strong>${fmt1.format(q.diferencia_stock_total)}</strong><span>Diferencia total de stock</span></div>`,`<div class="quality-kpi"><strong>${fmt1.format(q.diferencia_ventas_total)}</strong><span>Diferencia total de ventas</span></div>`].join('');renderTable('controlTable',DATA.control_totals,[{key:'sku_madre',label:'SKU',format:'sku'},{key:'stock_suma_canales',label:'Σ Canales',format:'num'},{key:'stock_total',label:'Total',format:'num'},{key:'diferencia_stock',label:'Dif. Stock',format:'difference'},{key:'ventas_suma_canales_90d',label:'Σ Ventas',format:'num'},{key:'ventas_3m_unidades',label:'Ventas total',format:'num'},{key:'diferencia_ventas',label:'Dif. Ventas',format:'difference'}],{pageSize:18,id:'controlQuality'});renderTable('excludedTransfersTable',DATA.excluded_transfers,[{key:'name',label:'Referencia'},{key:'documento_origen',label:'Documento origen'},{key:'contacto',label:'Contacto'},{key:'ubicacion_origen',label:'Origen'},{key:'state',label:'Estado',format:'badge'},{key:'fecha_evento',label:'Fecha'},{key:'motivo_exclusion',label:'Motivo',format:'badge'}],{pageSize:18,id:'excludedTransfers'});renderTable('stockUnlinkedTable',DATA.stock_unlinked,[{key:'sku_original',label:'SKU',format:'sku'},{key:'producto_madre',label:'Producto',format:'product'},{key:'stock_total',label:'Stock',format:'num'},{key:'motivo_no_vinculado',label:'Motivo',format:'badge'},{key:'accion_sugerida',label:'Acción',format:'product'}],{pageSize:18,id:'stockUnlinked'});renderTable('salesUnlinkedTable',DATA.sales_unlinked,[{key:'fecha',label:'Fecha'},{key:'pedido',label:'Pedido'},{key:'canal',label:'Canal'},{key:'sku_original',label:'SKU',format:'sku'},{key:'producto',label:'Producto',format:'product'},{key:'cantidad',label:'Cantidad',format:'num'},{key:'motivo_no_vinculado',label:'Motivo',format:'badge'}],{pageSize:18,id:'salesUnlinked'});}

function filterText(rows,query,keys){const q=(query||'').toLowerCase().trim();if(!q)return rows;return rows.filter(r=>keys.some(k=>String(r[k]||'').toLowerCase().includes(q)));}
function groupSum(rows,key,value){const map={};rows.forEach(r=>map[r[key]]=(map[r[key]]||0)+n(r[value]));return Object.entries(map).map(([key,value])=>({key,value})).sort((a,b)=>b.value-a.value);}
function cell(value,format,row={}){if(format==='num')return `<td class="num">${fmt.format(n(value))}</td>`;if(format==='num1')return `<td class="num">${fmt1.format(n(value))}</td>`;if(format==='numStrong')return `<td class="num"><b>${fmt.format(n(value))}</b></td>`;if(format==='money')return `<td class="num">${money.format(n(value))}</td>`;if(format==='money2')return `<td class="num">${money2.format(n(value))}</td>`;if(format==='percent')return `<td class="num">${pct2(n(value))}</td>`;if(format==='roiPct'){const x=numFilterValue({v:value},'v');return Number.isFinite(x)?`<td class="num${x<0?' neg':''}">${pct2(x)}</td>`:'<td class="num muted-cell">—</td>';}if(format==='moneySigned'){const x=numFilterValue({v:value},'v');return Number.isFinite(x)?`<td class="num${x<0?' neg':''}">${money.format(x)}</td>`:'<td class="num muted-cell">—</td>';}if(format==='coverage')return `<td class="num">${coverage(value)}</td>`;if(format==='daysAge')return `<td class="num">${daysAge(value)}</td>`;if(format==='sku')return `<td><span class="sku">${escapeHtml(value)}</span></td>`;if(format==='product')return `<td><div class="product" title="${escapeHtml(value)}">${escapeHtml(value)}</div></td>`;if(format==='badge')return `<td>${badge(value)}</td>`;if(format==='salesButton')return `<td><button class="sales-open-btn" onclick="openSalesDetail('${escapeHtml(value)}')">Ver ventas</button></td>`;if(format==='b2bButton')return `<td><button class="sales-open-btn" onclick="openB2BSimulator('${escapeHtml(value)}')">Simular</button></td>`;if(format==='distributionDraftButton'){const q=n(row?.cantidad_sugerida);if(q<=0)return '<td>—</td>';return `<td><button class="admin-transfer-btn" onclick="createDistributionRowDrafts('${escapeHtml(row.sku_madre)}','${escapeHtml(row.canal)}')">Crear borrador(s)</button></td>`;}if(format==='distributionButton')return `<td><button class="sales-open-btn" onclick="openDistributionDetail('${escapeHtml(value)}')">Ver reparto</button></td>`;if(format==='adminTransferButton'){const qty=Math.floor(n(row?.transferencia_sugerida)),manual=String(row?.check_manual_full||'').toUpperCase().startsWith('SI'),key=`${row?.sku_madre}|${row?.canal}|${qty}`,done=adminState.completed.has(key);if(qty<=0)return '<td>—</td>';if(manual)return '<td><span class="badge purple">CHECK MANUAL</span></td>';return `<td><button class="admin-transfer-btn" ${done?'disabled':''} onclick="openAdminTransfer('${escapeHtml(row?.sku_madre||'')}','${escapeHtml(row?.canal||'')}',${qty})">${done?'✓ Ejecutado':'🔒 Ejecutar'}</button></td>`;}if(format==='coverageBadge')return `<td><span class="badge ${coverageTone(row?.prioridad_cobertura)}">${escapeHtml(value||'—')}</span></td>`;if(format==='listingBadge')return `<td><span class="badge ${listingTone(value)}">${escapeHtml(value||'—')}</span></td>`;if(format==='levelBadge')return `<td><span class="badge ${levelTone(value)}">${escapeHtml(value||'—')}</span></td>`;if(format==='moveType')return `<td>${escapeHtml(moveTypeLabel(value))}</td>`;if(format==='dateShort')return `<td style="white-space:nowrap">${escapeHtml(String(value||'').slice(0,10)||'—')}</td>`;if(format==='text')return `<td>${escapeHtml(value||'—')}</td>`;if(format==='difference'){const v=n(value),cls=Math.abs(v)<.001?'green':'red';return `<td class="num"><span class="badge ${cls}">${fmt1.format(v)}</span></td>`;}return `<td>${escapeHtml(text(value))}</td>`;}
function tableRowMatches(row,columns,query){const q=String(query||'').toLowerCase().trim();if(!q)return true;return columns.some(c=>String(row?.[c.key]??'').toLowerCase().includes(q));}
function setInlineTableSearch(containerId,id,value){state.tableQueries[id]=value;state.tables[`${id}_page`]=1;const cfg=state.tables[`${id}_cfg`];if(cfg)renderTable(containerId,cfg.rows,cfg.columns,cfg.opts);}
function tableNumericFormat(format){return ['num','num1','numStrong','money','money2','percent','roiPct','moneySigned','coverage','daysAge','difference'].includes(String(format||''));}
function inferDefaultSortKey(columns,opts={}){
  if(opts.defaultSortKey&&columns.some(c=>c.key===opts.defaultSortKey))return opts.defaultSortKey;
  const priorityKeys=['transferencia_sugerida','compra_sugerida','cantidad_sugerida','piezas_sugeridas','cantidad','inventario_total','ventas_90d_unidades','ventas_3m_unidades','dias_inventario_canal','dias_inventario_num'];
  for(const key of priorityKeys){if(columns.some(c=>c.key===key))return key;}
  const actionLabel=columns.find(c=>/transferir|comprar|mover|cantidad|piezas|necesidad|stock|ventas/i.test(String(c.label||''))&&tableNumericFormat(c.format));
  if(actionLabel)return actionLabel.key;
  const numeric=columns.find(c=>tableNumericFormat(c.format));
  return numeric?numeric.key:(columns[0]?.key||'');
}
/* Filtros numéricos por rango.
   tableAutoFilterDefs solo sabe filtrar por valores repetidos (Riesgo, Canal,
   Acción...). Para columnas continuas como ROI o Utilidad ese enfoque no sirve:
   habría un valor distinto por renglón. Estas funciones permiten que cualquier
   columna declare `numFilter` con umbrales y el filtro aparece solo. */
const ROI_FILTER_OPTIONS=[
  {value:'gte:0.4',label:'≥ 40%'},
  {value:'gte:0.3',label:'≥ 30%'},
  {value:'gte:0.2',label:'≥ 20%'},
  {value:'gte:0.1',label:'≥ 10%'},
  {value:'gt:0',label:'positivo'},
  {value:'lte:0',label:'cero o negativo'},
  {value:'nodata',label:'sin dato'},
];
const UTILIDAD_FILTER_OPTIONS=[
  {value:'gte:100000',label:'≥ $100,000'},
  {value:'gte:50000',label:'≥ $50,000'},
  {value:'gte:10000',label:'≥ $10,000'},
  {value:'gte:1000',label:'≥ $1,000'},
  {value:'gt:0',label:'positiva'},
  {value:'lt:0',label:'negativa'},
  {value:'lte:0',label:'cero o negativa'},
];
/* Un valor vacío NO es cero: el generador exporta los NaN como cadena vacía,
   así que un ROI sin ventas en la ventana debe leerse como "sin dato" y no
   como 0%. Number('') devuelve 0, por eso no se puede usar n() aquí. */
function numFilterValue(row,key){const raw=row?.[key];if(raw===''||raw===null||raw===undefined)return NaN;const x=Number(raw);return Number.isFinite(x)?x:NaN;}
function tableNumericFilterDefs(columns){return (columns||[]).filter(c=>c&&c.numFilter&&Array.isArray(c.numFilter.options)&&c.numFilter.options.length).slice(0,6);}
function passesNumFilter(row,key,expr){
  if(!expr)return true;
  const parts=String(expr).split(':'),op=parts[0],target=Number(parts[1]),v=numFilterValue(row,key);
  if(op==='nodata')return Number.isNaN(v);
  if(Number.isNaN(v))return false;
  if(op==='gte')return v>=target;
  if(op==='gt')return v>target;
  if(op==='lte')return v<=target;
  if(op==='lt')return v<target;
  return true;
}
function setTableNumFilter(containerId,id,key,value){state.tableNumFilters[id]=state.tableNumFilters[id]||{};state.tableNumFilters[id][key]=value;state.tables[`${id}_page`]=1;const cfg=state.tables[`${id}_cfg`];if(cfg)renderTable(containerId,cfg.rows,cfg.columns,cfg.opts);}
function tableUniqueValues(rows,key){return [...new Set((rows||[]).map(r=>String(r?.[key]??'').trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es',{numeric:true,sensitivity:'base'}));}
function tableAutoFilterDefs(rows,columns){
  const defs=[];
  const keys=new Set(columns.map(c=>c.key));
  const add=(key,label)=>{if(keys.has(key)){const values=tableUniqueValues(rows,key);if(values.length>1&&values.length<=30)defs.push({key,label,values});}};
  add('canal','Canal');add('canal_origen','Origen');add('canal_destino','Destino');add('nivel_riesgo','Riesgo');add('demanda_tipo','Demanda');add('check_manual_full','Validación');add('confianza','Confianza');add('metodo_reparticion','Método');add('categoria','Categoría');add('veredicto','Resultado');add('accion_preliminar','Acción');
  return defs.slice(0,4);
}
function setTableSort(containerId,id,key){const current=state.tableSorts[id]||{};state.tableSorts[id]={key,dir:current.dir||'desc'};state.tables[`${id}_page`]=1;const cfg=state.tables[`${id}_cfg`];if(cfg)renderTable(containerId,cfg.rows,cfg.columns,cfg.opts);}
function toggleTableSortDir(containerId,id){const current=state.tableSorts[id]||{};state.tableSorts[id]={key:current.key||'',dir:current.dir==='asc'?'desc':'asc'};state.tables[`${id}_page`]=1;const cfg=state.tables[`${id}_cfg`];if(cfg)renderTable(containerId,cfg.rows,cfg.columns,cfg.opts);}
function setTableFilter(containerId,id,key,value){state.tableFilters[id]=state.tableFilters[id]||{};state.tableFilters[id][key]=value;state.tables[`${id}_page`]=1;const cfg=state.tables[`${id}_cfg`];if(cfg)renderTable(containerId,cfg.rows,cfg.columns,cfg.opts);}
function compareTableValues(a,b,col,dir){
  const av=a?.[col.key],bv=b?.[col.key];let cmp=0;
  if(tableNumericFormat(col.format)){const an=Number(av),bn=Number(bv),aok=Number.isFinite(an),bok=Number.isFinite(bn);if(aok&&bok)cmp=an-bn;else if(aok)cmp=1;else if(bok)cmp=-1;}
  else cmp=String(av??'').localeCompare(String(bv??''),'es',{numeric:true,sensitivity:'base'});
  return dir==='asc'?cmp:-cmp;
}
function renderTable(containerId,rows,columns,opts={}){
  const container=document.getElementById(containerId);if(!container)return;
  const id=opts.id||containerId,pageSize=opts.pageSize||20,sourceRows=rows||[];
  state.tables[`${id}_cfg`]={rows:sourceRows,columns,opts};

  const hasExternalSearch=Boolean(opts.searchInputId);
  const localQuery=hasExternalSearch?'':(state.tableQueries[id]||'');
  let filteredRows=localQuery?sourceRows.filter(r=>tableRowMatches(r,columns,localQuery)):sourceRows.slice();

  const filterDefs=tableAutoFilterDefs(sourceRows,columns);
  state.tableFilters[id]=state.tableFilters[id]||{};
  filterDefs.forEach(f=>{const wanted=state.tableFilters[id][f.key]||'';if(wanted)filteredRows=filteredRows.filter(r=>String(r?.[f.key]??'')===wanted);});

  const numFilterDefs=tableNumericFilterDefs(columns);
  state.tableNumFilters[id]=state.tableNumFilters[id]||{};
  numFilterDefs.forEach(c=>{const expr=state.tableNumFilters[id][c.key]||'';if(expr)filteredRows=filteredRows.filter(r=>passesNumFilter(r,c.key,expr));});

  const defaultSortKey=inferDefaultSortKey(columns,opts);
  if(!state.tableSorts[id]||!columns.some(c=>c.key===state.tableSorts[id].key))state.tableSorts[id]={key:defaultSortKey,dir:opts.defaultSortDir||'desc'};
  const sortState=state.tableSorts[id],sortCol=columns.find(c=>c.key===sortState.key)||columns[0];
  if(sortCol)filteredRows=[...filteredRows].sort((a,b)=>compareTableValues(a,b,sortCol,sortState.dir));

  state.tables[id]=filteredRows;
  const current=state.tables[`${id}_page`]||1,totalPages=Math.max(1,Math.ceil(filteredRows.length/pageSize)),page=Math.min(current,totalPages);
  state.tables[`${id}_page`]=page;
  const slice=filteredRows.slice((page-1)*pageSize,page*pageSize);

  const searchHtml=hasExternalSearch?'':`<input class="table-inline-search" value="${escapeHtml(localQuery)}" placeholder="Buscar en esta tabla…" oninput="setInlineTableSearch('${containerId}','${id}',this.value)"/>`;
  const sortableCols=columns.filter(c=>c.format!=='adminTransferButton'&&c.format!=='salesButton'&&c.format!=='b2bButton'&&c.format!=='distributionDraftButton'&&c.format!=='distributionButton');
  const sortOptions=sortableCols.map(c=>`<option value="${escapeHtml(c.key)}" ${c.key===sortState.key?'selected':''}>${escapeHtml(c.label)}</option>`).join('');
  const filtersHtml=filterDefs.map(f=>{const currentValue=state.tableFilters[id][f.key]||'';return `<select class="table-inline-select" onchange="setTableFilter('${containerId}','${id}','${escapeHtml(f.key)}',this.value)"><option value="">${escapeHtml(f.label)}: todos</option>${f.values.map(v=>`<option value="${escapeHtml(v)}" ${v===currentValue?'selected':''}>${escapeHtml(v)}</option>`).join('')}</select>`;}).join('');
  const numFiltersHtml=numFilterDefs.map(c=>{const label=c.numFilter.label||c.label,currentValue=state.tableNumFilters[id][c.key]||'';return `<select class="table-inline-select${currentValue?' active-filter':''}" onchange="setTableNumFilter('${containerId}','${id}','${escapeHtml(c.key)}',this.value)"><option value="">${escapeHtml(label)}: todos</option>${c.numFilter.options.map(o=>`<option value="${escapeHtml(o.value)}" ${o.value===currentValue?'selected':''}>${escapeHtml(label)} ${escapeHtml(o.label)}</option>`).join('')}</select>`;}).join('');
  const orderLabel=sortCol?.label||'columna';
  const toolsHtml=`<div class="table-inline-tools"><div class="table-inline-tools-left">${searchHtml}${filtersHtml}${numFiltersHtml}<select class="table-inline-select" title="Ordenar tabla" onchange="setTableSort('${containerId}','${id}',this.value)"><option value="" disabled>Ordenar por…</option>${sortOptions}</select><button class="table-sort-dir" type="button" onclick="toggleTableSortDir('${containerId}','${id}')">${sortState.dir==='desc'?'↓ Mayor → menor':'↑ Menor → mayor'}</button></div><div class="table-inline-tools-right"><span class="table-sort-note">Orden: ${escapeHtml(orderLabel)}</span><span class="table-inline-count">${fmt.format(filteredRows.length)} registros</span></div></div>`;

  if(!filteredRows.length){container.innerHTML=`${toolsHtml}<div class="empty"><div class="empty-icon">◌</div>Sin registros para los filtros actuales.</div>`;state.tables[`${id}_render`]=()=>renderTable(containerId,sourceRows,columns,opts);return;}
  container.innerHTML=`${toolsHtml}<div class="table-wrap"><table class="data-table"><thead><tr>${columns.map(c=>`<th>${escapeHtml(c.label)}</th>`).join('')}</tr></thead><tbody>${slice.map(r=>`<tr>${columns.map(c=>cell(r[c.key],c.format,r)).join('')}</tr>`).join('')}</tbody></table></div><div class="pager"><span>Mostrando ${(page-1)*pageSize+1}–${Math.min(page*pageSize,filteredRows.length)} de ${fmt.format(filteredRows.length)}</span><div class="pager-controls"><button ${page<=1?'disabled':''} onclick="changePage('${containerId}','${id}',${page-1})">Anterior</button><button ${page>=totalPages?'disabled':''} onclick="changePage('${containerId}','${id}',${page+1})">Siguiente</button></div></div>`;
  state.tables[`${id}_render`]=()=>renderTable(containerId,sourceRows,columns,opts);
}
function changePage(containerId,id,page){state.tables[`${id}_page`]=page;state.tables[`${id}_render`]();}
function downloadCsv(rows,name='exportacion'){if(!rows||!rows.length){showToast('No hay registros para exportar');return;}const keys=[...new Set(rows.flatMap(r=>Object.keys(r)))],csv=[keys.join(','),...rows.map(r=>keys.map(k=>`"${String(r[k]??'').replace(/"/g,'""')}"`).join(','))].join('\n');const blob=new Blob(['\ufeff'+csv],{type:'text/csv;charset=utf-8'}),url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=`${name}_${new Date().toISOString().slice(0,10)}.csv`;a.click();URL.revokeObjectURL(url);showToast('CSV exportado');}
function exportCurrentTable(id){downloadCsv(state.tables[id]||[],id);}

function closeMobile(){document.getElementById('sidebar').classList.remove('mobile-open');document.getElementById('overlay').classList.remove('show');}
document.getElementById('collapseBtn').addEventListener('click',()=>{const sb=document.getElementById('sidebar');sb.classList.toggle('collapsed');document.getElementById('collapseBtn').textContent=sb.classList.contains('collapsed')?'›':'‹';setTimeout(()=>window.dispatchEvent(new Event('resize')),280);});
document.getElementById('mobileMenu').addEventListener('click',()=>{document.getElementById('sidebar').classList.add('mobile-open');document.getElementById('overlay').classList.add('show');});document.getElementById('overlay').addEventListener('click',closeMobile);
document.getElementById('adminStatusBtn')?.addEventListener('click',()=>{adminState.pending=null;if(adminState.token)showAdminHome();else showAdminLogin();});
document.getElementById('adminPassword')?.addEventListener('keydown',e=>{if(e.key==='Enter')submitAdminLogin();});
document.getElementById('adminModalBackdrop')?.addEventListener('click',e=>{if(e.target===e.currentTarget)closeAdminModal();});
updateAdminStatus();
document.getElementById('themeBtn').addEventListener('click',()=>{document.body.classList.toggle('dark');localStorage.setItem('inventoryTheme',document.body.classList.contains('dark')?'dark':'light');renderPage(state.page);});
document.getElementById('refreshBtn').addEventListener('click',()=>location.reload());
document.getElementById('executiveProductSearch').addEventListener('input',()=>{state.pageQueries.executive=document.getElementById('executiveProductSearch').value;document.getElementById('globalSearch').value=state.pageQueries.executive;renderExecutiveProductTable();});document.getElementById('salesDetailPeriod').addEventListener('change',renderSalesDetail);document.getElementById('suggestionSearch').addEventListener('input',renderSuggestions);document.getElementById('distributionSearch').addEventListener('input',renderDistribution);document.getElementById('distributionMethod').addEventListener('change',renderDistribution);document.getElementById('b2bEvaluateBtn').addEventListener('click',()=>renderB2BSimulator(true));document.getElementById('b2bQty').addEventListener('input',()=>renderB2BSimulator(false));document.getElementById('b2bProposedRoi').addEventListener('input',()=>renderB2BSimulator(false));document.getElementById('b2bExplainBtnTop')?.addEventListener('click',toggleB2BExplanation);document.getElementById('suggestionConfidence').addEventListener('change',renderSuggestions);document.getElementById('distributionSearch')?.addEventListener('input',()=>{state.pageQueries.distribution=document.getElementById('distributionSearch').value;renderDistribution();});document.getElementById('distributionMethod')?.addEventListener('change',renderDistribution);document.getElementById('channelTableSearch').addEventListener('input',renderChannelTable);document.getElementById('channelStaleSearch').addEventListener('input',renderChannelStale);document.getElementById('channelStaleStockFilter').addEventListener('change',e=>setStaleFilter(e.target.value,false));document.getElementById('channelOdooSearch').addEventListener('input',renderOdooPanel);document.getElementById('channelOdooFilter').addEventListener('change',renderOdooPanel);document.getElementById('channelRiskFilter').addEventListener('change',renderChannelTable);document.getElementById('channelSalesMetric').addEventListener('change',e=>{state.channelSales.metric=e.target.value;renderChannelSales();});document.getElementById('channelSalesRange').addEventListener('change',e=>{state.channelSales.range=e.target.value;renderChannelSales();});document.getElementById('channelSalesFrom').addEventListener('change',e=>{state.channelSales.range='custom';state.channelSales.from=e.target.value;renderChannelSales();});document.getElementById('channelSalesTo').addEventListener('change',e=>{state.channelSales.range='custom';state.channelSales.to=e.target.value;renderChannelSales();});document.getElementById('transferSearch').addEventListener('input',renderTransfers);document.getElementById('purchaseSearch').addEventListener('input',renderPurchases);document.getElementById('arrivalSearch').addEventListener('input',renderLots);document.getElementById('alertSearch').addEventListener('input',renderAlerts);document.getElementById('alertRiskFilter').addEventListener('change',renderAlerts);
setupViewToggle('channelViewToggle',v=>{state.views.channel=v;renderChannelTable();});
document.getElementById('channelCoverageSearch').addEventListener('input',renderChannelCoverage);document.getElementById('channelListingsSearch').addEventListener('input',renderChannelListings);document.getElementById('channelListingsFilter').addEventListener('change',()=>setListingsFilter(document.getElementById('channelListingsFilter').value,false));document.getElementById('channelCoverageFilter').addEventListener('change',renderChannelCoverage);document.getElementById('channelMovesSearch').addEventListener('input',renderChannelMoves);document.getElementById('channelMovesFilter').addEventListener('change',renderChannelMoves);document.getElementById('auditSearch').addEventListener('input',renderAudit);document.getElementById('auditChannel').addEventListener('change',renderAudit);document.getElementById('auditLevel').addEventListener('change',renderAudit);initChannelJump();
setupViewToggle('alertsViewToggle',v=>{state.views.alerts=v;renderAlerts();});
function applyGlobalSearch(){const q=document.getElementById('globalSearch')?.value||'';state.pageQueries[state.page]=q;const set=(id)=>{const el=document.getElementById(id);if(el)el.value=q;};if(state.page==='executive'){set('executiveProductSearch');renderExecutiveProductTable();}else if(state.page==='channel'){set('channelTableSearch');set('channelStaleSearch');set('channelOdooSearch');set('channelCoverageSearch');set('channelMovesSearch');set('channelListingsSearch');renderChannelTable();renderChannelStale();renderChannelCoverage();renderChannelListings();renderChannelMoves();if(document.getElementById('channelOdooPanel')?.classList.contains('open'))renderOdooPanel();}else if(state.page==='suggestions'){set('suggestionSearch');renderSuggestions();}else if(state.page==='distribution'){set('distributionSearch');renderDistribution();}else if(state.page==='b2b'){set('b2bSearch');b2bSearchCombo?.refresh();renderB2B();}else if(state.page==='transfers'){set('transferSearch');renderTransfers();}else if(state.page==='purchases'){set('purchaseSearch');renderPurchases();}else if(state.page==='lots'){set('arrivalSearch');renderLots();}else if(state.page==='alerts'){set('alertSearch');renderAlerts();}else if(state.page==='quality'){renderQuality();}else if(state.page==='audit'){set('auditSearch');renderAudit();}}
document.getElementById('globalSearch').addEventListener('input',applyGlobalSearch);
if(localStorage.getItem('inventoryTheme')==='dark')document.body.classList.add('dark');
document.getElementById('heroDate').textContent=`Actualizado ${DATA.meta.generated_at}`;document.getElementById('footerDate').textContent=DATA.meta.generated_at;initNav();renderExecutive();
</script>
</body>
</html>'''

html_doc = (
    HTML_TEMPLATE
    .replace("__DATA_JSON__", DATA_JSON)
    .replace("__FULL_TRANSIT_DAYS__", str(DIAS_TRANSITO_ODOO_FULL))
    .replace("__FULL_TARGET_DAYS__", str(OBJETIVO_FULL_DIAS))
)
ARCHIVO_HTML.write_text(html_doc, encoding="utf-8")

print("\nDASHBOARD PROFESIONAL GENERADO")
print(f"Entrada: {ARCHIVO_BASE}")
print(f"Salida:  {ARCHIVO_HTML}")
print(f"B2B Odoo live: {b2b_live_meta.get('status')} · {b2b_live_meta.get('source')}")
print(f"Acciones admin: {ARCHIVO_ACCIONES_ADMIN}")
print(f"Canales: {', '.join(CANALES_ORDEN)}")
print(f"Registros por canal: {len(rotacion):,}")

# Publicación opcional. Desactivada por defecto.
PUBLICAR_GITHUB = os.getenv("PUBLICAR_GITHUB", "1").strip().lower() in {
    "1",
    "true",
    "si",
    "sí",
    "yes",
}

if PUBLICAR_GITHUB:
    try:
        status = subprocess.run(
            ["git", "-C", str(CARPETA), "status", "--porcelain"],
            capture_output=True,
            text=True,
            check=False,
        )
        if status.stdout.strip():
            subprocess.run(["git", "-C", str(CARPETA), "add", "index.html"], check=True)
            subprocess.run(
                [
                    "git",
                    "-C",
                    str(CARPETA),
                    "commit",
                    "-m",
                    f"Dashboard profesional {datetime.now():%Y-%m-%d %H:%M:%S}",
                ],
                check=False,
            )
            push = subprocess.run(
                ["git", "-C", str(CARPETA), "push", "origin", "main"],
                capture_output=True,
                text=True,
                check=False,
            )
            if push.returncode == 0:
                print("Publicado en GitHub.")
            else:
                print("No se pudo publicar en GitHub (git push falló):")
                if push.stdout.strip():
                    print(push.stdout.strip())
                if push.stderr.strip():
                    print(push.stderr.strip())
        else:
            print("GitHub: no hay cambios que publicar.")
    except Exception as exc:
        print(f"No se pudo publicar en GitHub: {exc}")
else:
    print("GitHub no se publicó. Usa PUBLICAR_GITHUB=1 para habilitarlo.")
