# -*- coding: utf-8 -*-
"""
04_calcular_b2b_excedente_oportunidad.py
=========================================
Modelo B2B por EXCEDENTE + COSTO DE OPORTUNIDAD.

Idea central
------------
1) Marketplace define primero cuánto inventario de bodega necesita para cubrir
   el lead time de reposición.
2) El inventario por encima de esa reserva es "excedente seguro" para B2B.
3) Si B2B pide más que el excedente, empieza a desplazar piezas protegidas de
   Marketplace. Esas piezas se valúan por su ROI Marketplace llevado a valor
   presente según día estimado de venta + payout.
4) El ROI mínimo B2B de un pedido q es el promedio de los ROI mínimos pieza a
   pieza: ROI base comercial para excedente y max(ROI base, costo de oportunidad)
   para piezas protegidas.

Entradas esperadas en la misma carpeta (o DASHBOARD_ROTACION_DIR):
- base_dashboard_rotacion.xlsx
- base_utilidad_roi.xlsx

Producción requiere la hoja rotacion_por_canal. Para generar un PREVIEW con una
base legacy que no la tenga se puede usar B2B_ALLOW_PREVIEW_FALLBACK=1.

Salidas:
- b2b_modelo_excedente.json
- b2b_resumen_excedente.csv

Variables editables (.env opcional):
DASHBOARD_ROTACION_DIR=...
ARCHIVO_DASHBOARD_B2B=base_dashboard_rotacion.xlsx
ARCHIVO_ROI_B2B=base_utilidad_roi.xlsx
B2B_LEAD_TIME_DIAS=30
B2B_ROI_BASE_EXCEDENTE=0.10
B2B_TASA_DESCUENTO_DIARIA=0.000355
B2B_ROI_VENTANA=90
B2B_PAYOUT_FALLBACK_DIAS=30
B2B_ALLOW_PREVIEW_FALLBACK=0

Payouts por canal (días):
B2B_PAYOUT_AMAZON_DIAS
B2B_PAYOUT_MERCADO_LIBRE_DIAS
B2B_PAYOUT_WALMART_DIAS
B2B_PAYOUT_LIVERPOOL_DIAS
B2B_PAYOUT_COPPEL_DIAS
B2B_PAYOUT_ELEKTRA_DIAS
B2B_PAYOUT_TIKTOK_DIAS
"""

from __future__ import annotations

from pathlib import Path
import json
import math
import os
from typing import Any

import numpy as np
import pandas as pd

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=False)
except Exception:
    pass


# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_DIR = Path(os.getenv("DASHBOARD_ROTACION_DIR", str(Path(__file__).resolve().parent))).resolve()
ARCHIVO_DASHBOARD = BASE_DIR / os.getenv("ARCHIVO_DASHBOARD_B2B", "base_dashboard_rotacion.xlsx")
ARCHIVO_ROI = BASE_DIR / os.getenv("ARCHIVO_ROI_B2B", "base_utilidad_roi.xlsx")
SALIDA_JSON = BASE_DIR / "b2b_modelo_excedente.json"
SALIDA_CSV = BASE_DIR / "b2b_resumen_excedente.csv"

LEAD_TIME_DIAS = max(float(os.getenv("B2B_LEAD_TIME_DIAS", "30")), 0.0)
ROI_BASE_B2B = max(float(os.getenv("B2B_ROI_BASE_EXCEDENTE", "0.10")), 0.10)
TASA_DIARIA = max(float(os.getenv("B2B_TASA_DESCUENTO_DIARIA", "0.000355")), 0.0)
ROI_VENTANA = str(os.getenv("B2B_ROI_VENTANA", "90")).strip().lower()
PAYOUT_FALLBACK = 30.0  # dato incompleto: supuesto temporal fijo de 30 días
ALLOW_PREVIEW = os.getenv("B2B_ALLOW_PREVIEW_FALLBACK", "0").strip().lower() in {"1", "true", "si", "sí", "yes"}

CANALES = ["Amazon", "Mercado Libre", "Walmart", "Liverpool", "Coppel", "Elektra", "TikTok"]


def env_float(name: str, default: float | None) -> float | None:
    raw = os.getenv(name)
    if raw is None or str(raw).strip() == "":
        return default
    return float(str(raw).strip())


PAYOUTS = {
    "Amazon": env_float("B2B_PAYOUT_AMAZON_DIAS", None),
    "Mercado Libre": env_float("B2B_PAYOUT_MERCADO_LIBRE_DIAS", 5.0),
    "Walmart": env_float("B2B_PAYOUT_WALMART_DIAS", 24.0),
    "Liverpool": env_float("B2B_PAYOUT_LIVERPOOL_DIAS", None),
    "Coppel": env_float("B2B_PAYOUT_COPPEL_DIAS", 39.0),
    "Elektra": env_float("B2B_PAYOUT_ELEKTRA_DIAS", None),
    "TikTok": env_float("B2B_PAYOUT_TIKTOK_DIAS", 14.0),
}


# ============================================================
# AUXILIARES
# ============================================================

def fnum(x: Any, default: float = 0.0) -> float:
    try:
        if pd.isna(x):
            return default
        v = float(x)
        return v if math.isfinite(v) else default
    except Exception:
        return default


def normalizar_canal(value: Any) -> str:
    t = str(value or "").strip().lower()
    if "amazon" in t or "fba" in t:
        return "Amazon"
    if "mercado" in t or "meli" in t or t == "ml":
        return "Mercado Libre"
    if "walmart" in t or "wfs" in t:
        return "Walmart"
    if "liverpool" in t or "99min" in t:
        return "Liverpool"
    if "coppel" in t:
        return "Coppel"
    if "elektra" in t:
        return "Elektra"
    if "tiktok" in t or "tik tok" in t:
        return "TikTok"
    if "general" in t or "existencias" in t or "odoo" in t:
        return "General"
    return str(value or "Sin identificar").strip() or "Sin identificar"


def elegir_roi(rr: pd.Series) -> tuple[float, str]:
    """Elige ROI del canal. Default 90d para estabilidad; configurable."""
    candidates: list[tuple[str, str, str | None]]
    if ROI_VENTANA in {"30", "30d"}:
        candidates = [("roi_30d", "ROI 30d", "n_ventas_30d"), ("roi_90d", "ROI 90d respaldo", "n_ventas_90d"), ("roi_canal", "ROI canal histórico", None)]
    elif ROI_VENTANA in {"canal", "historico", "histórico"}:
        candidates = [("roi_canal", "ROI canal histórico", None), ("roi_90d", "ROI 90d respaldo", "n_ventas_90d"), ("roi_30d", "ROI 30d respaldo", "n_ventas_30d")]
    else:
        candidates = [("roi_90d", "ROI 90d", "n_ventas_90d"), ("roi_30d", "ROI 30d respaldo", "n_ventas_30d"), ("roi_canal", "ROI canal histórico", None)]

    for roi_col, label, ncol in candidates:
        if roi_col not in rr.index or pd.isna(rr.get(roi_col)):
            continue
        if ncol and ncol in rr.index and fnum(rr.get(ncol), 0) <= 0:
            continue
        roi = fnum(rr.get(roi_col), 0)
        # ROI negativo sí es información económica real; no lo convertimos a positivo.
        return roi, label
    return 0.0, "Sin ROI"


def payout_canal(canal: str) -> tuple[float, str]:
    v = PAYOUTS.get(canal)
    if v is None:
        return PAYOUT_FALLBACK, f"Supuesto fallback {PAYOUT_FALLBACK:g}d"
    return max(float(v), 0.0), "Config canal"


def costo_implicito(rr: pd.Series) -> tuple[float, str]:
    """Proxy de costo unitario a partir de la misma base usada para ROI."""
    for base_col, units_col, label in [
        ("base_roi_90d", "unidades_90d", "base ROI 90d / unidades"),
        ("base_roi_30d", "unidades_30d", "base ROI 30d / unidades"),
    ]:
        base = fnum(rr.get(base_col), 0)
        units = fnum(rr.get(units_col), 0)
        if base > 0 and units > 0:
            return base / units, label
    return 0.0, "Sin costo implícito"


# ============================================================
# CARGA ROI
# ============================================================

def cargar_roi() -> tuple[pd.DataFrame, pd.DataFrame]:
    if not ARCHIVO_ROI.exists():
        raise FileNotFoundError(f"No encontré {ARCHIVO_ROI}")
    xls = pd.ExcelFile(ARCHIVO_ROI)
    if "por_producto_canal" not in xls.sheet_names:
        raise ValueError("base_utilidad_roi.xlsx debe contener 'por_producto_canal'.")
    pc = pd.read_excel(ARCHIVO_ROI, sheet_name="por_producto_canal")
    pc.columns = [str(c).strip() for c in pc.columns]
    pc["sku_madre"] = pc["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    pc["canal_norm"] = pc["canal"].apply(normalizar_canal)

    if "por_producto" in xls.sheet_names:
        pp = pd.read_excel(ARCHIVO_ROI, sheet_name="por_producto")
        pp.columns = [str(c).strip() for c in pp.columns]
        pp["sku_madre"] = pp["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    else:
        pp = pd.DataFrame()
    return pc, pp


# ============================================================
# CARGA OPERATIVA
# ============================================================

def cargar_operacion(pc: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, str, list[str]]:
    """Devuelve detalle SKU-canal y catálogo producto.

    Producción: rotacion_por_canal.
    Preview legacy: reconstruye lo indispensable desde dashboard_stock_canal y ROI 90d.
    """
    if not ARCHIVO_DASHBOARD.exists():
        raise FileNotFoundError(f"No encontré {ARCHIVO_DASHBOARD}")

    xls = pd.ExcelFile(ARCHIVO_DASHBOARD)
    warnings: list[str] = []

    if "rotacion_por_canal" in xls.sheet_names:
        rot = pd.read_excel(ARCHIVO_DASHBOARD, sheet_name="rotacion_por_canal")
        rot.columns = [str(c).strip() for c in rot.columns]
        rot["sku_madre"] = rot["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        rot["canal"] = rot["canal"].apply(normalizar_canal)
        # Versión 2026-09: el 02 también trae costo, valor y antigüedad por
        # canal. Incluye la fila "B2B" (CUATI/B2B), que suma a stock de bodega.
        for c in ["inventario_odoo", "inventario_transito", "inventario_full", "inventario_total", "ventas_10d_unidades", "ventas_30d_unidades", "ventas_90d_unidades", "demanda_diaria_ponderada",
                  "costo_unitario_inventario", "unidades_mas_90d", "valor_mas_90d"]:
            if c not in rot.columns:
                rot[c] = 0.0
            rot[c] = pd.to_numeric(rot[c], errors="coerce").fillna(0.0).clip(lower=0)
        if "producto_madre" not in rot.columns:
            rot["producto_madre"] = ""
        modo = "PRODUCCION · rotacion_por_canal"
    else:
        if not ALLOW_PREVIEW:
            raise ValueError(
                "base_dashboard_rotacion.xlsx no contiene 'rotacion_por_canal'. "
                "Corre el 02 actual. Para un preview legacy únicamente, usa B2B_ALLOW_PREVIEW_FALLBACK=1."
            )
        if "dashboard_stock_canal" not in xls.sheet_names:
            raise ValueError("La base legacy tampoco contiene dashboard_stock_canal.")
        stock = pd.read_excel(ARCHIVO_DASHBOARD, sheet_name="dashboard_stock_canal")
        stock.columns = [str(c).strip() for c in stock.columns]
        stock["sku_madre"] = stock["sku_madre"].fillna("").astype(str).str.strip().str.upper()
        full_cols = {
            "Amazon": "stock_amazon_fba",
            "Mercado Libre": "stock_meli_full",
            "Walmart": "stock_walmart_wfs",
            "Liverpool": "stock_liverpool_99min",
        }
        rows = []
        roi_lookup = pc.groupby(["sku_madre", "canal_norm"], dropna=False).first().reset_index()
        for _, s in stock.iterrows():
            sku = s["sku_madre"]
            prod = str(s.get("producto_madre", "") or "")
            odoo_total = max(fnum(s.get("stock_odoo_cuautitlan"), 0), 0)
            # General conserva todo Odoo en preview; no lo duplicamos por canal.
            rows.append({"sku_madre": sku, "producto_madre": prod, "canal": "General", "inventario_odoo": odoo_total, "inventario_transito": 0.0, "inventario_full": 0.0, "inventario_total": odoo_total, "ventas_90d_unidades": 0.0, "demanda_diaria_ponderada": 0.0})
            sku_roi = roi_lookup[roi_lookup["sku_madre"] == sku]
            for canal in CANALES:
                rr = sku_roi[sku_roi["canal_norm"] == canal]
                unidades90 = fnum(rr.iloc[0].get("unidades_90d"), 0) if not rr.empty else 0.0
                forecast = unidades90 / 90.0
                full = max(fnum(s.get(full_cols.get(canal, ""), 0), 0), 0) if canal in full_cols else 0.0
                if forecast > 0 or full > 0:
                    rows.append({"sku_madre": sku, "producto_madre": prod, "canal": canal, "inventario_odoo": 0.0, "inventario_transito": 0.0, "inventario_full": full, "inventario_total": full, "ventas_90d_unidades": unidades90, "demanda_diaria_ponderada": forecast})
        rot = pd.DataFrame(rows)
        modo = "PREVIEW LEGACY · sin rotacion_por_canal"
        warnings.append("PREVIEW: stock Odoo no está separado por canal; se usa como bolsa General.")
        warnings.append("PREVIEW: tránsito por canal no disponible y se toma como 0.")
        warnings.append("PREVIEW: forecast se aproxima con unidades_90d / 90 desde base_utilidad_roi.")

    productos = pd.read_excel(ARCHIVO_DASHBOARD, sheet_name="dashboard_productos") if "dashboard_productos" in xls.sheet_names else pd.DataFrame()
    if not productos.empty:
        productos.columns = [str(c).strip() for c in productos.columns]
        productos["sku_madre"] = productos["sku_madre"].fillna("").astype(str).str.strip().str.upper()
    return rot, productos, modo, warnings


# ============================================================
# MODELO POR SKU
# ============================================================

def construir_modelo(rot: pd.DataFrame, pc: pd.DataFrame, pp: pd.DataFrame, productos: pd.DataFrame, modo: str, warnings_global: list[str]) -> dict[str, Any]:
    # ROI map SKU-canal
    roi_map: dict[tuple[str, str], dict[str, Any]] = {}
    for _, rr in pc.iterrows():
        sku = str(rr.get("sku_madre", "")).strip().upper()
        canal = normalizar_canal(rr.get("canal_norm", rr.get("canal", "")))
        if not sku or canal not in CANALES:
            continue
        roi, roi_fuente = elegir_roi(rr)
        cost, cost_fuente = costo_implicito(rr)
        # Si hay duplicados tras normalizar, prioriza mayor número de ventas 90d.
        key = (sku, canal)
        score = fnum(rr.get("n_ventas_90d"), 0)
        if key not in roi_map or score > roi_map[key]["score"]:
            roi_map[key] = {
                "roi": roi,
                "roi_fuente": roi_fuente,
                "costo": cost,
                "costo_fuente": cost_fuente,
                "score": score,
            }

    pp_map = {}
    if not pp.empty:
        pp_map = pp.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    prod_map = {}
    if not productos.empty:
        prod_map = productos.drop_duplicates("sku_madre").set_index("sku_madre").to_dict("index")

    result_skus: list[dict[str, Any]] = []

    for sku, g in rot.groupby("sku_madre", dropna=False):
        sku = str(sku).strip().upper()
        if not sku:
            continue
        producto = next((str(x) for x in g.get("producto_madre", pd.Series(dtype=str)) if str(x).strip()), "")
        if not producto:
            producto = str(pp_map.get(sku, {}).get("producto", "") or "")

        stock_bodega = max(int(math.floor(g["inventario_odoo"].sum() + 1e-9)), 0)
        stock_full = max(int(math.floor(g["inventario_full"].sum() + 1e-9)), 0)
        stock_transito = max(int(math.floor(g["inventario_transito"].sum() + 1e-9)), 0)
        stock_empresa = max(int(math.floor(g["inventario_total"].sum() + 1e-9)), stock_bodega + stock_full)

        # Costos proxy por canales con ROI. Ponderamos por ventas recientes disponibles.
        cost_vals = []
        for canal in CANALES:
            info = roi_map.get((sku, canal))
            if info and info["costo"] > 0:
                cost_vals.append(info["costo"])
        costo_proxy = float(np.median(cost_vals)) if cost_vals else 0.0
        costo_fuente = "Mediana costo implícito por canal (base ROI)" if costo_proxy > 0 else "No disponible"

        # Si el dashboard actual incorpora costo Odoo, tiene prioridad.
        p = prod_map.get(sku, {})
        # costo_unitario_inventario (02 versión 2026-09) cubre también SKU que
        # solo tienen piezas en Full; tiene prioridad sobre el costo CUATI.
        for c in ["costo_unitario_inventario", "costo_unitario_odoo_cuati", "costo_unitario_odoo", "costo_promedio_odoo", "costo_odoo_unitario"]:
            v = fnum(p.get(c), 0)
            if v > 0:
                costo_proxy = v
                costo_fuente = f"{c} (dashboard)"
                break
        else:
            v = float(g["costo_unitario_inventario"].max()) if "costo_unitario_inventario" in g.columns else 0.0
            if v > 0:
                costo_proxy = v
                costo_fuente = "costo_unitario_inventario (rotacion_por_canal)"
        piezas_mas_90d = float(g["unidades_mas_90d"].sum()) if "unidades_mas_90d" in g.columns else 0.0
        valor_mas_90d = float(g["valor_mas_90d"].sum()) if "valor_mas_90d" in g.columns else 0.0

        channel_rows: list[dict[str, Any]] = []
        all_opportunities: list[dict[str, Any]] = []
        local_warnings: list[str] = []

        for canal in CANALES:
            gc = g[g["canal"] == canal]
            if gc.empty:
                continue
            rr = gc.iloc[0]
            forecast = max(fnum(rr.get("demanda_diaria_ponderada"), 0), 0)
            if forecast <= 0:
                forecast = max(fnum(rr.get("ventas_90d_unidades"), 0) / 90.0, 0)
            full = max(fnum(rr.get("inventario_full"), 0), 0)
            transito = max(fnum(rr.get("inventario_transito"), 0), 0)
            necesidad = int(math.ceil(forecast * LEAD_TIME_DIAS - 1e-12)) if forecast > 0 else 0
            fuera_bodega = int(math.floor(full + transito + 1e-9))
            reserva_bodega = max(necesidad - fuera_bodega, 0)

            fin = roi_map.get((sku, canal), {})
            roi = fnum(fin.get("roi"), 0)
            roi_source = fin.get("roi_fuente", "Sin ROI")
            payout, payout_source = payout_canal(canal)
            if (sku, canal) not in roi_map and forecast > 0:
                # Respaldo a ROI producto solo para no perder el canal; se marca.
                roi_prod = fnum(pp_map.get(sku, {}).get("roi_90d"), 0)
                if roi_prod == 0:
                    roi_prod = fnum(pp_map.get(sku, {}).get("roi_ponderado"), 0)
                roi = roi_prod
                roi_source = "ROI producto respaldo"
                local_warnings.append(f"{canal}: sin ROI canal; se usó ROI producto.")

            opp_values = []
            if forecast > 0 and reserva_bodega > 0:
                for k in range(1, reserva_bodega + 1):
                    posicion = fuera_bodega + k
                    dia_venta = max(int(math.ceil(posicion / forecast - 1e-12)), 1)
                    dia_cobro = dia_venta + payout
                    roi_vp = roi / ((1 + TASA_DIARIA) ** dia_cobro) if (1 + TASA_DIARIA) > 0 else roi
                    opp = {
                        "canal": canal,
                        "roi_marketplace": roi,
                        "roi_vp": roi_vp,
                        "dia_venta": dia_venta,
                        "dia_cobro": dia_cobro,
                        "posicion_canal": k,
                    }
                    all_opportunities.append(opp)
                    opp_values.append(roi_vp)

            channel_rows.append({
                "canal": canal,
                "forecast_dia": forecast,
                "necesidad_lead_time": necesidad,
                "full": full,
                "transito": transito,
                "reserva_bodega": reserva_bodega,
                "roi_marketplace": roi,
                "roi_fuente": roi_source,
                "payout_dias": payout,
                "payout_fuente": payout_source,
                "roi_vp_prom_reserva": float(np.mean(opp_values)) if opp_values else 0.0,
            })

        # La bolsa de bodega protege primero las oportunidades de mayor valor.
        protected_capacity = min(len(all_opportunities), stock_bodega)
        covered = sorted(all_opportunities, key=lambda x: x["roi_vp"], reverse=True)[:protected_capacity]
        # Cuando B2B rebasa el excedente, desplaza primero la oportunidad cubierta de menor valor.
        displaced_order = sorted(covered, key=lambda x: x["roi_vp"])
        excedente_seguro = max(stock_bodega - protected_capacity, 0)

        curve = []
        running = 0.0
        for q in range(1, stock_bodega + 1):
            if q <= excedente_seguro:
                marginal = ROI_BASE_B2B
                displaced = None
                tipo = "EXCEDENTE"
            else:
                idx = q - excedente_seguro - 1
                opp = displaced_order[idx] if idx < len(displaced_order) else None
                opp_roi = fnum(opp.get("roi_vp"), 0) if opp else 0.0
                marginal = max(ROI_BASE_B2B, opp_roi)
                displaced = opp
                tipo = "PROTEGIDA"
            running += marginal
            avg = running / q
            curve.append({
                "q": q,
                "roi_min_promedio": avg,
                "roi_marginal": marginal,
                "precio_min_unitario": costo_proxy * (1 + avg) if costo_proxy > 0 else 0.0,
                "tipo_pieza_marginal": tipo,
                "canal_desplazado": displaced.get("canal") if displaced else "",
                "dia_venta_desplazada": displaced.get("dia_venta") if displaced else None,
                "dia_cobro_desplazada": displaced.get("dia_cobro") if displaced else None,
                "roi_oportunidad_vp": displaced.get("roi_vp") if displaced else 0.0,
            })

        roi_all = curve[-1]["roi_min_promedio"] if curve else 0.0
        price_all = curve[-1]["precio_min_unitario"] if curve else 0.0
        status = "LISTO" if stock_bodega > 0 else "SIN STOCK BODEGA"
        if stock_bodega > 0 and costo_proxy <= 0:
            status = "ROI LISTO · PRECIO SIN COSTO"

        result_skus.append({
            "sku_madre": sku,
            "producto_madre": producto,
            "stock_bodega": stock_bodega,
            "stock_full": stock_full,
            "stock_transito": stock_transito,
            "stock_empresa": stock_empresa,
            "reserva_marketplace_bodega": protected_capacity,
            "reserva_requerida_total": len(all_opportunities),
            "excedente_b2b_seguro": excedente_seguro,
            "roi_base_b2b": ROI_BASE_B2B,
            "roi_min_todo_bodega": roi_all,
            "precio_min_todo_bodega": price_all,
            "costo_unitario": costo_proxy,
            "fuente_costo": costo_fuente,
            "valor_excedente_seguro": excedente_seguro * costo_proxy,
            "piezas_mas_90d": piezas_mas_90d,
            "valor_mas_90d": valor_mas_90d,
            "estado": status,
            "canales": channel_rows,
            "curva": curve,
            "warnings": sorted(set(local_warnings)),
        })

    result_skus.sort(key=lambda x: (x["excedente_b2b_seguro"], x["stock_bodega"]), reverse=True)

    summary = {
        "modo": modo,
        "lead_time_dias": LEAD_TIME_DIAS,
        "roi_base_b2b": ROI_BASE_B2B,
        "tasa_descuento_diaria": TASA_DIARIA,
        "roi_ventana": ROI_VENTANA,
        "payout_fallback_dias": PAYOUT_FALLBACK,
        "warnings": warnings_global,
        "n_skus": len(result_skus),
        "stock_bodega_total": int(sum(x["stock_bodega"] for x in result_skus)),
        "excedente_seguro_total": int(sum(x["excedente_b2b_seguro"] for x in result_skus)),
        "reserva_marketplace_total": int(sum(x["reserva_marketplace_bodega"] for x in result_skus)),
        "valor_excedente_total": float(sum(x["valor_excedente_seguro"] for x in result_skus)),
        "valor_mas_90d_total": float(sum(x["valor_mas_90d"] for x in result_skus)),
    }
    return {"meta": summary, "skus": result_skus}


def exportar(modelo: dict[str, Any]) -> None:
    SALIDA_JSON.write_text(json.dumps(modelo, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    rows = []
    for x in modelo["skus"]:
        rows.append({
            "sku_madre": x["sku_madre"],
            "producto_madre": x["producto_madre"],
            "stock_bodega": x["stock_bodega"],
            "stock_full": x["stock_full"],
            "stock_transito": x["stock_transito"],
            "stock_empresa": x["stock_empresa"],
            "reserva_marketplace_bodega": x["reserva_marketplace_bodega"],
            "excedente_b2b_seguro": x["excedente_b2b_seguro"],
            "roi_base_b2b": x["roi_base_b2b"],
            "roi_min_todo_bodega": x["roi_min_todo_bodega"],
            "costo_unitario": x["costo_unitario"],
            "valor_excedente_seguro": x["valor_excedente_seguro"],
            "piezas_mas_90d": x["piezas_mas_90d"],
            "valor_mas_90d": x["valor_mas_90d"],
            "precio_min_todo_bodega": x["precio_min_todo_bodega"],
            "estado": x["estado"],
        })
    pd.DataFrame(rows).to_csv(SALIDA_CSV, index=False, encoding="utf-8-sig")


def main() -> int:
    print("=" * 78)
    print("B2B · EXCEDENTE + COSTO DE OPORTUNIDAD")
    print("=" * 78)
    print(f"Carpeta: {BASE_DIR}")
    print(f"Dashboard: {ARCHIVO_DASHBOARD.name}")
    print(f"ROI: {ARCHIVO_ROI.name}")
    pc, pp = cargar_roi()
    rot, productos, modo, warnings = cargar_operacion(pc)
    modelo = construir_modelo(rot, pc, pp, productos, modo, warnings)
    exportar(modelo)
    m = modelo["meta"]
    print(f"Modo: {m['modo']}")
    print(f"SKU: {m['n_skus']:,}")
    print(f"Stock bodega: {m['stock_bodega_total']:,}")
    print(f"Reserva Marketplace: {m['reserva_marketplace_total']:,}")
    print(f"Excedente seguro B2B: {m['excedente_seguro_total']:,}")
    print(f"ROI base B2B excedente: {m['roi_base_b2b']:.2%}")
    print(f"JSON: {SALIDA_JSON}")
    print(f"CSV: {SALIDA_CSV}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
