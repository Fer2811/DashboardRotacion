# -*- coding: utf-8 -*-
"""
generar_dashboard_html_rotacion_v3.py

Genera un dashboard HTML estático a partir de:
Desktop/rotacion_inventario_base_dashboard/base_dashboard_rotacion.xlsx

Salida:
Desktop/rotacion_inventario_base_dashboard/dashboard_rotacion.html

Uso:
python3 generar_dashboard_html_rotacion_v3.py
"""

from pathlib import Path
import json
import html
import pandas as pd
import numpy as np
import subprocess
from datetime import datetime

FECHA_GENERACION = datetime.now().strftime("%d/%m/%Y %H:%M")


# ============================================================
# CONFIGURACIÓN
# ============================================================

#DESKTOP = Path.home() / "Desktop"
#CARPETA = DESKTOP / "rotacion_inventario_base_dashboard"
CARPETA = Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")

ARCHIVO_BASE = CARPETA / "base_dashboard_rotacion.xlsx"
ARCHIVO_HTML = CARPETA / "index.html"


# ============================================================
# FUNCIONES
# ============================================================

def fmt_num(x, dec=0):
    try:
        if pd.isna(x):
            return "-"
        return f"{float(x):,.{dec}f}"
    except Exception:
        return str(x)


def fmt_money(x):
    try:
        if pd.isna(x):
            return "$0"
        return "$" + f"{float(x):,.0f}"
    except Exception:
        return str(x)


def safe_text(x):
    if pd.isna(x):
        return ""
    return html.escape(str(x))


def df_to_js_records(df):
    df = df.copy()
    df = df.replace([np.inf, -np.inf], np.nan)

    # Convertir fechas/Timestamps a texto para que JSON pueda serializar.
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")

    # Convertir cualquier valor suelto tipo Timestamp/date que haya quedado en columnas object.
    # Pandas 3 removió DataFrame.applymap; usamos map por columna.
    for col in df.columns:
        df[col] = df[col].map(
            lambda x: x.strftime("%Y-%m-%d %H:%M:%S")
            if hasattr(x, "strftime") and not isinstance(x, str)
            else x
        )

    df = df.fillna("")
    return json.dumps(df.to_dict(orient="records"), ensure_ascii=False)


def color_alerta_py(alerta):
    alerta = str(alerta).lower()
    if "comprar ya" in alerta:
        return "#ef4444"
    if "urgencia" in alerta:
        return "#f97316"
    if "suficiente" in alerta:
        return "#22c55e"
    if "sin ventas" in alerta:
        return "#64748b"
    return "#64748b"


def color_riesgo_py(riesgo):
    riesgo = str(riesgo)
    if "+90" in riesgo:
        return "#ef4444"
    if "+60" in riesgo:
        return "#f97316"
    if "+30" in riesgo:
        return "#eab308"
    if "Sin venta" in riesgo:
        return "#64748b"
    return "#22c55e"


def clase_recomendacion_py(reco):
    reco = str(reco).lower()
    if "comprar ya" in reco or "+90" in reco:
        return "rec-red"
    if "urgencia" in reco or "+60" in reco:
        return "rec-orange"
    if "sano" in reco or "suficiente" in reco:
        return "rec-green"
    return "rec-gray"


def badge_py(text, color):
    return (
        f'<span class="badge" '
        f'style="background:{color}20;color:{color};border:1px solid {color}55;">'
        f'{safe_text(text)}</span>'
    )


def df_to_table_rows(df, cols, max_rows=500):
    if df.empty:
        return "<tr><td colspan='20'>Sin datos</td></tr>"

    rows = []
    for _, r in df.head(max_rows).iterrows():
        tds = []
        for c in cols:
            val = r.get(c, "")
            if c == "alerta_compra":
                val_html = badge_py(val, color_alerta_py(val))
            elif c == "riesgo_sobrestock":
                val_html = badge_py(val, color_riesgo_py(val))
            elif c == "recomendacion":
                val_html = f'<span class="{clase_recomendacion_py(val)}">{safe_text(val)}</span>'
            elif c in ["ventas_3m_monto", "venta_total"]:
                val_html = fmt_money(val)
            elif c in ["ventas_3m_unidades", "stock_total", "stock_total_congelado",
                       "arribos_post_corte_unidades", "ventas_post_corte_unidades",
                       "sugerencia_compra", "cantidad_arribada", "cantidad_ordenada"]:
                val_html = fmt_num(val, 0)
            elif c in ["ventas_3m_monto", "arribos_post_corte_monto", "ventas_post_corte_monto", "subtotal", "precio_unitario"]:
                val_html = fmt_money(val)
            elif c in ["dias_para_comprar", "dias_inventario_num", "dias_desde_ultimo_arribo"]:
                val_html = fmt_num(val, 1)
            else:
                val_html = safe_text(val)
            tds.append(f"<td>{val_html}</td>")
        rows.append("<tr>" + "".join(tds) + "</tr>")
    return "\n".join(rows)


# ============================================================
# CARGA
# ============================================================

if not ARCHIVO_BASE.exists():
    raise FileNotFoundError(
        f"No encontré {ARCHIVO_BASE}. Primero corre generar_base_dashboard_rotacion_v2.py"
    )

productos = pd.read_excel(ARCHIVO_BASE, sheet_name="dashboard_productos")
ventas_no = pd.read_excel(ARCHIVO_BASE, sheet_name="ventas_no_vinculadas")
stock_no = pd.read_excel(ARCHIVO_BASE, sheet_name="stock_no_vinculado")

# Hoja nueva generada por:
# generar_base_dashboard_rotacion_v8_todos_arribos_odoo.py
try:
    arribos_odoo = pd.read_excel(ARCHIVO_BASE, sheet_name="arribos_odoo_compras")
except Exception:
    arribos_odoo = pd.DataFrame()

for df in [productos, ventas_no, stock_no, arribos_odoo]:
    df.columns = [str(c).strip() for c in df.columns]

# En la pestaña de stock no vinculado solo queremos SKUs con stock_total > 0.
if "stock_total" in stock_no.columns:
    stock_no["stock_total"] = pd.to_numeric(stock_no["stock_total"], errors="coerce").fillna(0)
    stock_no = stock_no[stock_no["stock_total"] > 0].copy()

for col in [
    "ventas_3m_unidades", "ventas_3m_monto", "stock_total",
    "dias_inventario_num", "lead_time_dias", "dias_para_comprar",
    "stock_odoo_cuautitlan", "stock_amazon_fba", "stock_meli_full",
    "stock_walmart_wfs", "stock_liverpool_99min", "sugerencia_compra",
    "objetivo_stock_45_dias", "venta_diaria_promedio_3m",
    "cantidad_arribada_historica",
    "dias_desde_ultimo_arribo",
    "participacion_ventas_3m", "participacion_acumulada_3m",
    "stock_total_congelado", "arribos_post_corte_unidades", "arribos_post_corte_monto",
    "ventas_post_corte_unidades", "ventas_post_corte_monto",
    "stock_total_antes_movimientos", "stock_total_antes_descuento"
]:
    if col in productos.columns:
        productos[col] = pd.to_numeric(productos[col], errors="coerce").fillna(0)

for col in ["cantidad_arribada", "cantidad_ordenada", "precio_unitario", "subtotal"]:
    if col in arribos_odoo.columns:
        arribos_odoo[col] = pd.to_numeric(arribos_odoo[col], errors="coerce").fillna(0)

# ============================================================
# KPIS
# ============================================================

total_iq = len(productos)
stock_total = productos["stock_total"].sum() if "stock_total" in productos.columns else 0
ventas_u = productos["ventas_3m_unidades"].sum() if "ventas_3m_unidades" in productos.columns else 0
ventas_monto = productos["ventas_3m_monto"].sum() if "ventas_3m_monto" in productos.columns else 0
top80_count = int((productos["top_80_flag"].astype(str) == "SI").sum()) if "top_80_flag" in productos.columns else 0
comprar_ya = int(productos["alerta_compra"].astype(str).str.contains("Comprar ya", case=False, na=False).sum()) if "alerta_compra" in productos.columns else 0
urgencia = int(productos["alerta_compra"].astype(str).str.contains("urgencia", case=False, na=False).sum()) if "alerta_compra" in productos.columns else 0
riesgo_90 = int(productos["riesgo_sobrestock"].astype(str).str.contains(r"\+90", regex=True, na=False).sum()) if "riesgo_sobrestock" in productos.columns else 0
no_vinc = len(ventas_no)
stock_no_skus = stock_no["sku_original"].nunique() if "sku_original" in stock_no.columns and not stock_no.empty else len(stock_no)
stock_no_units = stock_no["stock_total"].sum() if "stock_total" in stock_no.columns and not stock_no.empty else 0

arribos_total_rows = len(arribos_odoo)
arribos_total_units = arribos_odoo["cantidad_arribada"].sum() if "cantidad_arribada" in arribos_odoo.columns and not arribos_odoo.empty else 0
if "post_corte_stock" in arribos_odoo.columns:
    arribos_post_rows = int(arribos_odoo["post_corte_stock"].astype(str).str.upper().eq("SI").sum())
    arribos_post_units = arribos_odoo.loc[
        arribos_odoo["post_corte_stock"].astype(str).str.upper().eq("SI"),
        "cantidad_arribada"
    ].sum() if "cantidad_arribada" in arribos_odoo.columns else 0
else:
    arribos_post_rows = 0
    arribos_post_units = productos["arribos_post_corte_unidades"].sum() if "arribos_post_corte_unidades" in productos.columns else 0

iq_con_imagen = int(productos["tiene_imagen_odoo"].astype(str).str.upper().eq("SI").sum()) if "tiene_imagen_odoo" in productos.columns else int(productos["imagen_odoo_base64"].astype(str).str.strip().ne("").sum()) if "imagen_odoo_base64" in productos.columns else 0

# ============================================================
# TABLAS
# ============================================================

tabla_cols = [
    "sku_madre", "producto_madre", "ranking_top80",
    "ventas_3m_unidades", "ventas_3m_monto", "stock_total",
    "sugerencia_compra", "dias_inventario_texto", "dias_desde_ultimo_arribo_texto",
    "lead_time_dias", "dias_para_comprar",
    "alerta_compra", "riesgo_sobrestock", "alerta_antiguedad_inventario",
    "prioridad", "recomendacion",
]
tabla_cols = [c for c in tabla_cols if c in productos.columns]
tabla_rows = df_to_table_rows(productos.sort_values("ventas_3m_monto", ascending=False), tabla_cols, max_rows=1000)
tabla_head = "".join(f"<th>{safe_text(c)}</th>" for c in tabla_cols)

no_cols_pref = [
    "fecha", "fuente_log", "fuente", "canal", "pedido", "referencia",
    "producto", "sku_log", "sku_odoo", "sku_autoazur", "sku_original",
    "cantidad", "venta_total", "motivo_no_vinculado", "accion_sugerida"
]
no_cols = [c for c in no_cols_pref if c in ventas_no.columns]
if not no_cols:
    no_cols = ventas_no.columns.tolist()[:14]
no_rows = df_to_table_rows(ventas_no, no_cols, max_rows=1500)
no_head = "".join(f"<th>{safe_text(c)}</th>" for c in no_cols)

stock_no_cols_pref = [
    "sku_original", "producto_madre", "stock_total", "WALMART_WFS",
    "LIVERPOOL_FULL_99MIN", "MERCADO_LIBRE_FULL", "AMAZON_FBA",
    "ODOO_CUAUTITLAN", "motivo_no_vinculado", "accion_sugerida"
]
stock_no_cols = [c for c in stock_no_cols_pref if c in stock_no.columns]
if not stock_no_cols:
    stock_no_cols = stock_no.columns.tolist()[:12]
stock_no_rows = df_to_table_rows(stock_no, stock_no_cols, max_rows=1000)
stock_no_head = "".join(f"<th>{safe_text(c)}</th>" for c in stock_no_cols)

arribos_cols_pref = [
    "fecha", "post_corte_stock", "orden_compra", "proveedor", "referencia_interna",
    "producto", "cantidad_arribada", "cantidad_ordenada", "sku_odoo",
    "barcode", "estado_compra", "origen", "precio_unitario", "subtotal"
]
arribos_cols = [c for c in arribos_cols_pref if c in arribos_odoo.columns]
if not arribos_cols:
    arribos_cols = arribos_odoo.columns.tolist()[:14]
arribos_rows = df_to_table_rows(arribos_odoo, arribos_cols, max_rows=3000)
arribos_head = "".join(f"<th>{safe_text(c)}</th>" for c in arribos_cols)

productos_js = df_to_js_records(productos)
ventas_no_js = df_to_js_records(ventas_no)
stock_no_js = df_to_js_records(stock_no)
arribos_js = df_to_js_records(arribos_odoo)

# ============================================================
# HTML
# ============================================================

html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Dashboard Rotación Inventario</title>
<script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
:root {{
  --bg:#f4f6fb;
  --card:#ffffff;
  --text:#111827;
  --muted:#667085;
  --nav:#111827;
  --line:#e5e7eb;
  --blue:#2563eb;
  --shadow:0 10px 28px rgba(15,23,42,.08);
  --radius:18px;
  --red:#ef4444;
  --orange:#f97316;
  --green:#22c55e;
  --gray:#64748b;
  --yellow:#eab308;
}}
* {{ box-sizing:border-box; }}
body {{
  margin:0;
  background:var(--bg);
  color:var(--text);
  font-family:Inter,Arial,sans-serif;
}}
.container {{
  max-width:1920px;
  margin:0 auto;
  padding:26px 28px 42px;
}}
.tabs {{
  display:flex;
  gap:10px;
  margin-bottom:24px;
  flex-wrap:wrap;
}}
.tab-btn {{
  border:0;
  background:#e5e7eb;
  color:#374151;
  border-radius:999px;
  padding:14px 23px;
  font-weight:850;
  font-size:15px;
  cursor:pointer;
}}
.tab-btn.active {{
  background:#111827;
  color:white;
  box-shadow:inset 0 0 0 2px #2563eb;
}}
.tab-content {{ display:none; }}
.tab-content.active {{ display:block; }}
.kpi-grid {{
  display:grid;
  grid-template-columns:repeat(11,minmax(145px,1fr));
  gap:18px;
  margin-bottom:28px;
}}
.metric-card {{
  background:var(--card);
  border-radius:var(--radius);
  padding:22px 24px;
  box-shadow:var(--shadow);
  border:1px solid rgba(148,163,184,.18);
  min-height:128px;
}}
.metric-title {{
  color:var(--muted);
  font-size:14px;
  font-weight:750;
  margin-bottom:10px;
}}
.metric-value {{
  font-size:31px;
  line-height:1;
  font-weight:900;
  letter-spacing:-.04em;
}}
.metric-sub {{
  color:var(--muted);
  font-size:13px;
  margin-top:10px;
  line-height:1.35;
}}
.grid-2 {{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:24px;
  margin-bottom:24px;
}}
.grid-alert-stock {{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:24px;
  margin-bottom:24px;
}}
.alert-split {{
  display:grid;
  grid-template-columns:1fr 1fr;
  gap:18px;
}}
.card {{
  background:var(--card);
  border-radius:var(--radius);
  padding:24px 26px;
  box-shadow:var(--shadow);
  border:1px solid rgba(148,163,184,.18);
  overflow:hidden;
}}
.card h2 {{
  margin:0 0 18px;
  font-size:23px;
  letter-spacing:-.02em;
}}
.chart-toolbar {{
  display:flex;
  gap:10px;
  align-items:center;
  justify-content:flex-end;
  margin-top:-8px;
  margin-bottom:8px;
  flex-wrap:wrap;
}}
.select {{
  border:1px solid #d1d5db;
  background:white;
  border-radius:12px;
  padding:9px 12px;
  font-weight:700;
  color:#374151;
}}
.chart-scroll {{
  max-height:560px;
  overflow-y:auto;
  overflow-x:hidden;
  padding-right:8px;
}}
.search {{
  width:100%;
  border:1px solid var(--line);
  background:white;
  border-radius:14px;
  padding:14px 16px;
  font-size:15px;
  outline:none;
}}
.search:focus {{
  border-color:var(--blue);
  box-shadow:0 0 0 4px rgba(37,99,235,.12);
}}
.table-wrap {{
  overflow:auto;
  max-height:620px;
  border:1px solid var(--line);
  border-radius:14px;
}}
table {{
  width:100%;
  border-collapse:collapse;
  background:white;
  font-size:13px;
}}
th {{
  background:#111827;
  color:white;
  text-align:left;
  padding:11px 12px;
  position:sticky;
  top:0;
  z-index:2;
  white-space:nowrap;
}}
td {{
  padding:10px 12px;
  border-bottom:1px solid #edf0f4;
  vertical-align:top;
}}
tr:hover td {{ background:#f8fafc; }}
.badge {{
  display:inline-block;
  border-radius:999px;
  padding:6px 10px;
  font-size:12px;
  font-weight:850;
  white-space:nowrap;
}}


.sku-channel-grid {{
  display:grid;
  grid-template-columns:repeat(2,minmax(220px,1fr));
  gap:12px;
}}
.sku-channel-card {{
  border:1px solid #e5e7eb;
  border-radius:16px;
  padding:14px;
  background:#f8fafc;
}}
.sku-channel-head {{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  margin-bottom:10px;
  font-size:13px;
  font-weight:900;
}}
.sku-channel-count {{
  font-size:12px;
  font-weight:900;
  padding:4px 8px;
  border-radius:999px;
  background:white;
  color:#475569;
  border:1px solid #e5e7eb;
}}
.sku-list {{
  display:flex;
  flex-wrap:wrap;
  gap:7px;
}}
.sku-pill {{
  display:inline-flex;
  align-items:center;
  border-radius:999px;
  padding:7px 10px;
  font-size:12px;
  font-weight:800;
  border:1px solid rgba(148,163,184,.45);
  background:white;
  color:#111827;
  max-width:100%;
  overflow-wrap:anywhere;
}}
.sku-empty {{
  font-size:13px;
  color:#94a3b8;
  font-weight:750;
}}
@media (max-width:900px) {{
  .sku-channel-grid {{
    grid-template-columns:1fr;
  }}
}}

.product-hero {{
  display:grid;
  grid-template-columns:180px 1fr auto;
  gap:20px;
  align-items:center;
}}
.product-image-box {{
  width:180px;
  height:180px;
  border-radius:22px;
  border:1px solid #e5e7eb;
  background:#f8fafc;
  display:flex;
  align-items:center;
  justify-content:center;
  overflow:hidden;
  box-shadow:inset 0 0 0 1px rgba(148,163,184,.12);
}}
.product-image-box img {{
  width:100%;
  height:100%;
  object-fit:contain;
  background:white;
}}
.product-image-placeholder {{
  color:#94a3b8;
  font-size:13px;
  font-weight:800;
  text-align:center;
  padding:16px;
}}
@media (max-width:900px) {{
  
.sku-channel-grid {{
  display:grid;
  grid-template-columns:repeat(2,minmax(220px,1fr));
  gap:12px;
}}
.sku-channel-card {{
  border:1px solid #e5e7eb;
  border-radius:16px;
  padding:14px;
  background:#f8fafc;
}}
.sku-channel-head {{
  display:flex;
  align-items:center;
  justify-content:space-between;
  gap:10px;
  margin-bottom:10px;
  font-size:13px;
  font-weight:900;
}}
.sku-channel-count {{
  font-size:12px;
  font-weight:900;
  padding:4px 8px;
  border-radius:999px;
  background:white;
  color:#475569;
  border:1px solid #e5e7eb;
}}
.sku-list {{
  display:flex;
  flex-wrap:wrap;
  gap:7px;
}}
.sku-pill {{
  display:inline-flex;
  align-items:center;
  border-radius:999px;
  padding:7px 10px;
  font-size:12px;
  font-weight:800;
  border:1px solid rgba(148,163,184,.45);
  background:white;
  color:#111827;
  max-width:100%;
  overflow-wrap:anywhere;
}}
.sku-empty {{
  font-size:13px;
  color:#94a3b8;
  font-weight:750;
}}
@media (max-width:900px) {{
  .sku-channel-grid {{
    grid-template-columns:1fr;
  }}
}}

.product-hero {{
    grid-template-columns:1fr;
  }}
  .product-image-box {{
    width:100%;
    max-width:220px;
    height:220px;
  }}
}}

.product-header {{
  display:flex;
  align-items:flex-start;
  justify-content:space-between;
  gap:16px;
  margin-bottom:20px;
}}
.product-title h1 {{
  margin:0;
  font-size:30px;
  letter-spacing:-.04em;
}}
.sku-chips {{
  display:flex;
  gap:7px;
  flex-wrap:wrap;
  margin-top:12px;
  max-height:72px;
  overflow:hidden;
}}
.chip {{
  display:inline-flex;
  align-items:center;
  gap:6px;
  border-radius:999px;
  padding:6px 9px;
  font-size:12px;
  font-weight:800;
  border:1px solid transparent;
  line-height:1;
}}
.chip small {{
  opacity:.8;
  font-weight:900;
}}
.detail-grid {{
  display:grid;
  grid-template-columns:repeat(6,1fr);
  gap:16px;
  margin-bottom:24px;
}}
.detail-layout {{
  display:grid;
  grid-template-columns:.95fr 1.05fr;
  gap:24px;
}}
.info-list {{ display:grid; gap:10px; }}
.info-row {{
  display:grid;
  grid-template-columns:220px 1fr;
  gap:12px;
  padding:10px 0;
  border-bottom:1px solid #eef2f7;
}}
.info-label {{
  color:var(--muted);
  font-weight:800;
}}
.info-label .hint {{
  display:inline-block;
  margin-left:6px;
  width:18px;
  height:18px;
  line-height:18px;
  text-align:center;
  border-radius:50%;
  background:#eef2ff;
  color:#3730a3;
  font-size:12px;
  cursor:help;
}}
.info-value {{
  font-weight:750;
  word-break:break-word;
}}
.rec-red {{ color:#991b1b; background:#fee2e2; border:1px solid #fecaca; padding:8px 11px; border-radius:12px; display:inline-block; }}
.rec-orange {{ color:#9a3412; background:#ffedd5; border:1px solid #fed7aa; padding:8px 11px; border-radius:12px; display:inline-block; }}
.rec-green {{ color:#166534; background:#dcfce7; border:1px solid #bbf7d0; padding:8px 11px; border-radius:12px; display:inline-block; }}
.rec-gray {{ color:#374151; background:#f3f4f6; border:1px solid #e5e7eb; padding:8px 11px; border-radius:12px; display:inline-block; }}
.empty {{
  padding:28px;
  color:var(--muted);
  text-align:center;
}}
@media (max-width:1500px) {{
  .kpi-grid {{ grid-template-columns:repeat(3,1fr); }}
  .grid-2,.grid-alert-stock {{ grid-template-columns:1fr; }}
  .detail-grid {{ grid-template-columns:repeat(3,1fr); }}
  .detail-layout {{ grid-template-columns:1fr; }}
}}
@media (max-width:900px) {{
  .container {{ padding:18px; }}
  .kpi-grid {{ grid-template-columns:1fr; }}
  .detail-grid {{ grid-template-columns:1fr; }}
  .alert-split {{ grid-template-columns:1fr; }}
}}

/* Vista compacta de SKUs por canal */
.sku-channel-grid {{
  display:flex;
  flex-direction:column;
  gap:10px;
}}
.sku-channel-card {{
  border:1px solid #e5e7eb;
  border-radius:16px;
  padding:12px 14px;
  background:#ffffff;
}}
.sku-channel-head {{
  display:flex;
  align-items:center;
  gap:10px;
  margin-bottom:8px;
  font-size:13px;
  font-weight:900;
}}
.sku-channel-count {{
  font-size:11px;
  font-weight:900;
  padding:3px 8px;
  border-radius:999px;
  background:#f8fafc;
  color:#475569;
  border:1px solid #e5e7eb;
}}
.sku-list {{
  display:flex;
  flex-wrap:wrap;
  gap:7px;
}}
.sku-pill {{
  display:inline-flex;
  align-items:center;
  border-radius:999px;
  padding:6px 10px;
  font-size:12px;
  font-weight:800;
  border:1px solid rgba(148,163,184,.45);
  background:white;
  color:#111827;
  max-width:100%;
  overflow-wrap:anywhere;
}}
.sku-empty {{
  font-size:13px;
  color:#94a3b8;
  font-weight:750;
}}


.sku-bar {{
  margin:16px 0 22px;
  border:1px solid #e5e7eb;
  background:#ffffff;
  border-radius:18px;
  padding:14px 16px;
  box-shadow:0 8px 20px rgba(15,23,42,.05);
}}
.sku-bar-title {{
  font-size:13px;
  font-weight:900;
  color:#475569;
  margin-bottom:10px;
}}
.sku-bar-scroll {{
  display:flex;
  gap:10px;
  overflow-x:auto;
  padding-bottom:4px;
}}
.sku-bar-group {{
  display:flex;
  align-items:center;
  gap:8px;
  white-space:nowrap;
  border-radius:999px;
  padding:8px 10px;
  border:1px solid #e5e7eb;
  background:#f8fafc;
  flex:0 0 auto;
}}
.sku-bar-market {{
  font-size:11px;
  font-weight:900;
  padding:4px 8px;
  border-radius:999px;
}}
.sku-bar-sku {{
  font-size:12px;
  font-weight:850;
  color:#111827;
}}


.chart-description {{
  color:#667085;
  font-size:14px;
  line-height:1.45;
  margin:-4px 0 14px;
}}

.dashboard-header{{
  display:flex;
  justify-content:space-between;
  align-items:center;
  margin-bottom:20px;
  padding:18px 24px;
  background:white;
  border-radius:18px;
  box-shadow:var(--shadow);
  border:1px solid rgba(148,163,184,.18);
}}
.dashboard-title{{
  font-size:32px;
  font-weight:900;
  margin:0;
}}
.dashboard-date{{
  color:#667085;
  font-size:15px;
  font-weight:700;
}}

</style>
</head>
<body>
<div class="container">

  <div class="dashboard-header">
      <div>
          <h1 class="dashboard-title">Dashboard Rotación Inventario</h1>
      </div>
      <div class="dashboard-date">
          Actualizado: {FECHA_GENERACION}
      </div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="showTab('main', this)">Dashboard principal</button>
    <button class="tab-btn" onclick="showTab('detail', this)">Detalle por producto</button>
    <button class="tab-btn" onclick="showTab('unlinked', this)">Ventas no vinculadas</button>
    <button class="tab-btn" onclick="showTab('stockUnlinked', this)">Stock no vinculado</button>
    <button class="tab-btn" onclick="showTab('arribosOdoo', this)">Arribos Odoo</button>
  </div>

  <section id="main" class="tab-content active">
    <div class="kpi-grid">
      <div class="metric-card"><div class="metric-title">IQ en dashboard</div><div class="metric-value">{fmt_num(total_iq)}</div><div class="metric-sub">Con stock o ventas 3M</div></div>
      <div class="metric-card"><div class="metric-title">Ventas últimos 3 meses</div><div class="metric-value">{fmt_num(ventas_u)}</div><div class="metric-sub">{fmt_money(ventas_monto)}</div></div>
      <div class="metric-card"><div class="metric-title">Stock total</div><div class="metric-value">{fmt_num(stock_total)}</div><div class="metric-sub">Congelado + arribos - ventas</div></div>
      <div class="metric-card"><div class="metric-title">Arribos Odoo</div><div class="metric-value">{fmt_num(arribos_total_rows)}</div><div class="metric-sub">{fmt_num(arribos_total_units)} unidades totales · {fmt_num(arribos_post_units)} post-corte</div></div>
      <div class="metric-card"><div class="metric-title">Imágenes Odoo</div><div class="metric-value">{fmt_num(iq_con_imagen)}</div><div class="metric-sub">IQ con imagen en detalle</div></div>
      <div class="metric-card"><div class="metric-title">Top 80% ventas $</div><div class="metric-value">{fmt_num(top80_count)}</div><div class="metric-sub">IQ dentro del Pareto por monto</div></div>
      <div class="metric-card"><div class="metric-title">Comprar ya</div><div class="metric-value">{fmt_num(comprar_ya)}</div><div class="metric-sub">Inventario acabado o debajo del lead time</div></div>
      <div class="metric-card"><div class="metric-title">Compra con urgencia</div><div class="metric-value">{fmt_num(urgencia)}</div><div class="metric-sub">Inventario próximo a acabar</div></div>
      <div class="metric-card"><div class="metric-title">Riesgo +90 días</div><div class="metric-value">{fmt_num(riesgo_90)}</div><div class="metric-sub">Sobrestock alto</div></div>
      <div class="metric-card"><div class="metric-title">Ventas no vinculadas</div><div class="metric-value">{fmt_num(no_vinc)}</div><div class="metric-sub">Logs / ventas pendientes de mapeo</div></div>
      <div class="metric-card"><div class="metric-title">Stock no vinculado</div><div class="metric-value">{fmt_num(stock_no_skus)}</div><div class="metric-sub">{fmt_num(stock_no_units)} unidades</div></div>
    </div>

    <div class="grid-2">
      <div class="card">
        <h2>Top IQ por monto vendido últimos 3 meses</h2>
        <div class="chart-toolbar">
          <label>Mostrar</label>
          <select class="select" id="topLimit" onchange="renderTopChart()">
            <option value="15">Top 15</option>
            <option value="30">Top 30</option>
            <option value="50">Top 50</option>
            <option value="80">Top 80</option>
            <option value="all">Todos</option>
          </select>
        </div>
        <div class="chart-scroll"><div id="topChart"></div></div>
      </div>

      <div class="card">
        <h2>Días de inventario por IQ</h2>
        <p class="chart-description">
          Cobertura estimada: cuántos días duraría el stock actual con el ritmo de venta promedio.
        </p>
        <div class="chart-toolbar">
          <label>Riesgo</label>
          <select class="select" id="riskFilter" onchange="renderDiasChart()">
            <option value="all">Todos</option>
            <option value="+30">+30 días</option>
            <option value="+60">+60 días</option>
            <option value="+90">+90 días</option>
            <option value="sinventa">Sin venta</option>
          </select>
          <label>Mostrar</label>
          <select class="select" id="diasLimit" onchange="renderDiasChart()">
            <option value="15">Top 15</option>
            <option value="30">Top 30</option>
            <option value="50">Top 50</option>
            <option value="80">Top 80</option>
            <option value="all">Todos</option>
          </select>
        </div>
        <div class="chart-scroll"><div id="diasChart"></div></div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <h2>Antigüedad de inventario por IQ</h2>
        <p class="chart-description">
          Días desde el último arribo/compra registrada en Odoo; solo considera productos con stock actual mayor a cero.
        </p>
        <div class="chart-toolbar">
          <label>Alerta</label>
          <select class="select" id="ageFilter" onchange="renderAgeChart()">
            <option value="all">Todos</option>
            <option value="+30">+30 días</option>
            <option value="+60">+60 días</option>
            <option value="+90">+90 días</option>
            <option value="sinarribo">Sin arribo</option>
          </select>
          <label>Mostrar</label>
          <select class="select" id="ageLimit" onchange="renderAgeChart()">
            <option value="15">Top 15</option>
            <option value="30">Top 30</option>
            <option value="50">Top 50</option>
            <option value="80">Top 80</option>
            <option value="all">Todos</option>
          </select>
        </div>
        <div class="chart-scroll"><div id="ageChart"></div></div>
      </div>

      <div class="card">
        <h2>Distribución de alertas</h2>
        <div id="alertGeneralChart"></div>
      </div>
    </div>

    <div class="grid-alert-stock">
      <div class="card">
        <h2>Stock por canal</h2>
        <div id="stockChart"></div>
      </div>
    </div>

    <div class="card">
      <h2>Tabla operativa</h2>
      <input class="search" id="tableSearch" placeholder="Buscar en tabla por IQ, producto o alerta..." onkeyup="filterTable('tableSearch','mainTable')">
      <br><br>
      <div class="table-wrap">
        <table id="mainTable">
          <thead><tr>{tabla_head}</tr></thead>
          <tbody>{tabla_rows}</tbody>
        </table>
      </div>
    </div>
  </section>

  <section id="detail" class="tab-content">
    <div class="card">
      <h2>Buscar producto</h2>
      <input class="search" id="productSearch" placeholder="Buscar por IQ, producto o SKU sincronizado..." onkeyup="renderProductResults()">
      <br><br>
      <div id="productResults"></div>
    </div>
    <br>
    <div id="productDetail"></div>
  </section>

  <section id="unlinked" class="tab-content">
    <div class="card">
      <h2>Ventas no vinculadas</h2>
      <p style="color:#667085;margin-top:-8px;">Logs Odoo / Autoazur y ventas que no encontraron referencia madre IQ.</p>
      <input class="search" id="unlinkedSearch" placeholder="Buscar en ventas no vinculadas..." onkeyup="filterTable('unlinkedSearch','unlinkedTable')">
      <br><br>
      <div class="table-wrap">
        <table id="unlinkedTable">
          <thead><tr>{no_head}</tr></thead>
          <tbody>{no_rows}</tbody>
        </table>
      </div>
    </div>
  </section>

  <section id="stockUnlinked" class="tab-content">
    <div class="card">
      <h2>Stock no vinculado</h2>
      <p style="color:#667085;margin-top:-8px;">SKUs con inventario que no encontraron referencia madre IQ.</p>
      <input class="search" id="stockNoSearch" placeholder="Buscar en stock no vinculado..." onkeyup="filterTable('stockNoSearch','stockNoTable')">
      <br><br>
      <div class="table-wrap">
        <table id="stockNoTable">
          <thead><tr>{stock_no_head}</tr></thead>
          <tbody>{stock_no_rows}</tbody>
        </table>
      </div>
    </div>
  </section>

  <section id="arribosOdoo" class="tab-content">
    <div class="card">
      <h2>Arribos Odoo Compras</h2>
      <p style="color:#667085;margin-top:-8px;">
        Todos los renglones recibidos desde Odoo Compras. La columna <b>post_corte_stock</b>
        indica cuáles se suman al stock congelado.
      </p>
      <div class="detail-grid">
        <div class="mini-metric"><div class="mini-title">Renglones totales</div><div class="mini-value">{fmt_num(arribos_total_rows)}</div></div>
        <div class="mini-metric"><div class="mini-title">Unidades totales recibidas</div><div class="mini-value">{fmt_num(arribos_total_units)}</div></div>
        <div class="mini-metric"><div class="mini-title">Unidades post-corte</div><div class="mini-value">{fmt_num(arribos_post_units)}</div></div>
      </div>
      <br>
      <input class="search" id="arribosSearch" placeholder="Buscar por IQ, producto, proveedor u orden..." onkeyup="filterTable('arribosSearch','arribosTable')">
      <br><br>
      <div class="table-wrap">
        <table id="arribosTable">
          <thead><tr>{arribos_head}</tr></thead>
          <tbody>{arribos_rows}</tbody>
        </table>
      </div>
    </div>
  </section>

</div>

<script>
const PRODUCTS = {productos_js};
const VENTAS_NO = {ventas_no_js};
const STOCK_NO = {stock_no_js};
const ARRIBOS_ODOO = {arribos_js};

const COLORS = {{
  red: "#ef4444",
  orange: "#f97316",
  green: "#22c55e",
  gray: "#64748b",
  yellow: "#eab308",
  blue: "#2563eb",
  navy: "#111827"
}};

const MARKET_COLORS = {{
  "Amazon": ["#fff7ed", "#c2410c"],
  "Mercado Libre": ["#fef9c3", "#854d0e"],
  "Walmart": ["#dbeafe", "#1d4ed8"],
  "Liverpool": ["#fce7f3", "#be185d"],
  "Coppel": ["#dcfce7", "#166534"],
  "Elektra": ["#ffe4e6", "#be123c"],
  "TikTok": ["#e5e7eb", "#111827"],
  "Odoo/Interno": ["#e0f2fe", "#0369a1"],
  "IQ": ["#e0e7ff", "#3730a3"],
  "Otro": ["#f3f4f6", "#374151"]
}};

function showTab(id, btn) {{
  document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  btn.classList.add('active');
  if (id === 'detail') renderProductResults();
}}

function filterTable(inputId, tableId) {{
  const input = document.getElementById(inputId);
  const filter = input.value.toLowerCase();
  const table = document.getElementById(tableId);
  const rows = table.getElementsByTagName("tr");
  for (let i = 1; i < rows.length; i++) {{
    const txt = rows[i].textContent || rows[i].innerText;
    rows[i].style.display = txt.toLowerCase().indexOf(filter) > -1 ? "" : "none";
  }}
}}

function fmtNum(x, dec=0) {{
  const n = Number(x);
  if (!isFinite(n)) return "-";
  return n.toLocaleString("en-US", {{minimumFractionDigits: dec, maximumFractionDigits: dec}});
}}

function fmtMoney(x) {{
  const n = Number(x);
  if (!isFinite(n)) return "$0";
  return "$" + n.toLocaleString("en-US", {{maximumFractionDigits: 0}});
}}

function esc(x) {{
  if (x === null || x === undefined) return "";
  return String(x)
    .replace(/&/g,"&amp;")
    .replace(/</g,"&lt;")
    .replace(/>/g,"&gt;")
    .replace(/"/g,"&quot;");
}}

function alertaColor(alerta) {{
  const a = String(alerta || "").toLowerCase();
  if (a.includes("comprar ya")) return COLORS.red;
  if (a.includes("urgencia")) return COLORS.orange;
  if (a.includes("suficiente")) return COLORS.green;
  if (a.includes("sin ventas")) return COLORS.gray;
  return COLORS.gray;
}}

function riesgoColor(riesgo) {{
  const r = String(riesgo || "");
  if (r.includes("+90")) return COLORS.red;
  if (r.includes("+60")) return COLORS.orange;
  if (r.includes("+30")) return COLORS.yellow;
  if (r.includes("Sin venta")) return COLORS.gray;
  return COLORS.green;
}}

function recClass(reco) {{
  const r = String(reco || "").toLowerCase();
  if (r.includes("comprar ya") || r.includes("+90")) return "rec-red";
  if (r.includes("urgencia") || r.includes("+60")) return "rec-orange";
  if (r.includes("sano") || r.includes("suficiente")) return "rec-green";
  return "rec-gray";
}}

function badgeHtml(text, color) {{
  return `<span class="badge" style="background:${{color}}20;color:${{color}};border:1px solid ${{color}}55;">${{esc(text)}}</span>`;
}}

function metricHtml(title, value, sub) {{
  return `
    <div class="metric-card">
      <div class="metric-title">${{esc(title)}}</div>
      <div class="metric-value">${{value}}</div>
      <div class="metric-sub">${{esc(sub || "")}}</div>
    </div>`;
}}

function chipHtml(chip) {{
  const market = chip.marketplace || "Otro";
  const pair = MARKET_COLORS[market] || MARKET_COLORS["Otro"];
  return `<span class="chip" style="background:${{pair[0]}};color:${{pair[1]}};border-color:${{pair[1]}}33;"><small>${{esc(market)}}</small>${{esc(chip.sku)}}</span>`;
}}

function chipsHtml(jsonText, limit=12) {{
  let chips = [];
  try {{ chips = JSON.parse(jsonText || "[]"); }} catch(e) {{ chips = []; }}
  const shown = chips.slice(0, limit).map(chipHtml).join("");
  const more = chips.length > limit ? `<span class="chip" style="background:#f3f4f6;color:#374151;">+ ${{chips.length - limit}} más</span>` : "";
  return `<div class="sku-chips">${{shown}}${{more}}</div>`;
}}

function splitSkuText(txt) {{
  return String(txt || "")
    .split("|")
    .map(s => s.trim())
    .filter(s => s && !["nan","none","false"].includes(s.toLowerCase()));
}}

function skuChannelHtml(jsonText, p) {{
  let groups = [];
  try {{ groups = JSON.parse(jsonText || "[]"); }} catch(e) {{ groups = []; }}

  if (!groups.length) {{
    const fallback = [
      ["IQ", splitSkuText(p.skus_iq)],
      ["Odoo/Interno", splitSkuText(p.skus_odoo_interno)],
      ["Amazon", splitSkuText(p.skus_amazon)],
      ["Mercado Libre", splitSkuText(p.skus_mercado_libre)],
      ["Walmart", splitSkuText(p.skus_walmart)],
      ["Liverpool", splitSkuText(p.skus_liverpool)],
      ["Coppel", splitSkuText(p.skus_coppel)],
      ["Elektra", splitSkuText(p.skus_elektra)],
      ["TikTok", splitSkuText(p.skus_tiktok)],
      ["Otro", splitSkuText(p.skus_otro)]
    ];
    groups = fallback.filter(g => g[1].length).map(g => ({{marketplace:g[0], skus:g[1], count:g[1].length}}));
  }}

  if (!groups.length) {{
    return `<div class="sku-empty">Sin SKUs sincronizados por canal.</div>`;
  }}

  return `<div class="sku-channel-grid">` + groups.map(g => {{
    const market = g.marketplace || "Otro";
    const canal = g.canal || market;
    const pair = MARKET_COLORS[market] || MARKET_COLORS["Otro"];
    const skus = Array.isArray(g.skus) ? g.skus : [];
    const pills = skus.map(s => `<span class="sku-pill" style="border-color:${{pair[1]}}55;background:${{pair[0]}};color:${{pair[1]}};">${{esc(s)}}</span>`).join("");
    return `
      <div class="sku-channel-card">
        <div class="sku-channel-head">
          <span style="color:${{pair[1]}};">${{esc(canal)}}</span>
          <span class="sku-channel-count">${{skus.length}} SKU</span>
        </div>
        <div class="sku-list">${{pills || `<span class="sku-empty">Sin SKU</span>`}}</div>
      </div>`;
  }}).join("") + `</div>`;
}}


function skuBarHtml(jsonText, p) {{
  let groups = [];
  try {{ groups = JSON.parse(jsonText || "[]"); }} catch(e) {{ groups = []; }}

  if (!groups.length) {{
    return "";
  }}

  const items = [];
  groups.forEach(g => {{
    const market = g.marketplace || "Otro";
    const canal = g.canal || market;
    const pair = MARKET_COLORS[market] || MARKET_COLORS["Otro"];
    const skus = Array.isArray(g.skus) ? g.skus : [];

    skus.forEach(s => {{
      items.push(`
        <div class="sku-bar-group">
          <span class="sku-bar-market" style="background:${{pair[0]}};color:${{pair[1]}};">${{esc(canal)}}</span>
          <span class="sku-bar-sku">${{esc(s)}}</span>
        </div>
      `);
    }});
  }});

  if (!items.length) return "";

  return `
    <div class="sku-bar">
      <div class="sku-bar-title">SKUs vinculados por canal</div>
      <div class="sku-bar-scroll">${{items.join("")}}</div>
    </div>
  `;
}}


function getLimited(arr, limitValue) {{
  if (limitValue === "all") return arr;
  return arr.slice(0, Number(limitValue));
}}

function renderTopChart() {{
  const lim = document.getElementById("topLimit").value;
  let data = PRODUCTS
    .filter(p => Number(p.ventas_3m_monto || 0) > 0)
    .sort((a,b) => Number(b.ventas_3m_monto || 0) - Number(a.ventas_3m_monto || 0));
  data = getLimited(data, lim).reverse();

  const h = Math.max(520, data.length * 30);
  document.getElementById("topChart").style.height = h + "px";

  Plotly.newPlot("topChart", [{{
    type: "bar",
    orientation: "h",
    x: data.map(p => Number(p.ventas_3m_monto || 0)),
    y: data.map(p => p.sku_madre),
    text: data.map(p => fmtMoney(p.ventas_3m_monto)),
    textposition: "outside",
    marker: {{color: COLORS.blue}},
    customdata: data.map(p => [p.producto_madre, p.ventas_3m_unidades, p.stock_total, p.dias_inventario_texto, p.alerta_compra]),
    hovertemplate: "<b>%{{y}}</b><br>%{{customdata[0]}}<br>Monto vendido: $%{{x:,.0f}}<br>Unidades: %{{customdata[1]}}<br>Stock: %{{customdata[2]}}<br>Días inv: %{{customdata[3]}}<br>%{{customdata[4]}}<extra></extra>"
  }}], {{
    height: h,
    margin: {{l:80,r:80,t:10,b:40}},
    xaxis: {{title: "Monto vendido 3M"}},
    yaxis: {{title: "IQ"}},
    plot_bgcolor: "white",
    paper_bgcolor: "white",
    font: {{family:"Inter, Arial", color:"#111827"}}
  }}, {{displayModeBar:false, responsive:true}});
}}

function filterRisk(data, risk) {{
  if (risk === "all") return data;
  if (risk === "+30") return data.filter(p => String(p.riesgo_sobrestock).includes("+30"));
  if (risk === "+60") return data.filter(p => String(p.riesgo_sobrestock).includes("+60"));
  if (risk === "+90") return data.filter(p => String(p.riesgo_sobrestock).includes("+90"));
  if (risk === "sinventa") return data.filter(p => String(p.riesgo_sobrestock).includes("Sin venta"));
  return data;
}}

function renderDiasChart() {{
  const lim = document.getElementById("diasLimit").value;
  const risk = document.getElementById("riskFilter").value;
  let data = PRODUCTS
    .filter(p => isFinite(Number(p.dias_inventario_num)))
    .sort((a,b) => Number(b.dias_inventario_num || 0) - Number(a.dias_inventario_num || 0));
  data = filterRisk(data, risk);
  data = getLimited(data, lim).reverse();

  const h = Math.max(520, data.length * 30);
  document.getElementById("diasChart").style.height = h + "px";

  Plotly.newPlot("diasChart", [{{
    type: "bar",
    orientation: "h",
    x: data.map(p => Number(p.dias_inventario_num || 0)),
    y: data.map(p => p.sku_madre),
    text: data.map(p => fmtNum(p.dias_inventario_num, 1)),
    textposition: "outside",
    marker: {{color: data.map(p => riesgoColor(p.riesgo_sobrestock))}},
    customdata: data.map(p => [p.producto_madre, p.ventas_3m_unidades, p.stock_total, p.riesgo_sobrestock]),
    hovertemplate: "<b>%{{y}}</b><br>%{{customdata[0]}}<br>Días inventario: %{{x:.1f}}<br>Ventas 3M: %{{customdata[1]}}<br>Stock: %{{customdata[2]}}<br>%{{customdata[3]}}<extra></extra>"
  }}], {{
    height: h,
    margin: {{l:80,r:35,t:10,b:40}},
    xaxis: {{title: "Días de inventario"}},
    yaxis: {{title: "IQ"}},
    plot_bgcolor: "white",
    paper_bgcolor: "white",
    font: {{family:"Inter, Arial", color:"#111827"}}
  }}, {{displayModeBar:false, responsive:true}});
}}


function edadColor(alerta) {{
  const a = String(alerta || "");
  if (a.includes("+90")) return COLORS.red;
  if (a.includes("+60")) return COLORS.orange;
  if (a.includes("+30")) return COLORS.yellow;
  if (a.includes("Sin arribo")) return COLORS.gray;
  return COLORS.green;
}}

function filterAge(data, risk) {{
  if (risk === "all") return data;
  if (risk === "+30") return data.filter(p => String(p.alerta_antiguedad_inventario).includes("+30"));
  if (risk === "+60") return data.filter(p => String(p.alerta_antiguedad_inventario).includes("+60"));
  if (risk === "+90") return data.filter(p => String(p.alerta_antiguedad_inventario).includes("+90"));
  if (risk === "sinarribo") return data.filter(p => String(p.alerta_antiguedad_inventario).includes("Sin arribo"));
  return data;
}}

function renderAgeChart() {{
  const lim = document.getElementById("ageLimit").value;
  const risk = document.getElementById("ageFilter").value;

  let data = PRODUCTS
    .filter(p =>
      isFinite(Number(p.dias_desde_ultimo_arribo)) &&
      Number(p.stock_total || 0) > 0
    )
    .sort((a,b) => Number(b.dias_desde_ultimo_arribo || 0) - Number(a.dias_desde_ultimo_arribo || 0));

  data = filterAge(data, risk);
  data = getLimited(data, lim).reverse();

  const h = Math.max(520, data.length * 30);
  document.getElementById("ageChart").style.height = h + "px";

  Plotly.newPlot("ageChart", [{{
    type: "bar",
    orientation: "h",
    x: data.map(p => Number(p.dias_desde_ultimo_arribo || 0)),
    y: data.map(p => p.sku_madre),
    text: data.map(p => fmtNum(p.dias_desde_ultimo_arribo, 0)),
    textposition: "outside",
    marker: {{color: data.map(p => edadColor(p.alerta_antiguedad_inventario))}},
    customdata: data.map(p => [
      p.producto_madre,
      p.ultima_fecha_compra_odoo,
      p.proveedor_principal,
      p.alerta_antiguedad_inventario,
      p.stock_total
    ]),
    hovertemplate: "<b>%{{y}}</b><br>%{{customdata[0]}}<br>Días desde último arribo: %{{x:.0f}}<br>Último arribo: %{{customdata[1]}}<br>Proveedor: %{{customdata[2]}}<br>Stock actual: %{{customdata[4]:,.0f}} unidades<br>%{{customdata[3]}}<extra></extra>"
  }}], {{
    height: h,
    margin: {{l:80,r:35,t:10,b:40}},
    xaxis: {{title: "Días desde último arribo"}},
    yaxis: {{title: "IQ"}},
    plot_bgcolor: "white",
    paper_bgcolor: "white",
    font: {{family:"Inter, Arial", color:"#111827"}}
  }}, {{displayModeBar:false, responsive:true}});
}}


function countBy(arr, fn) {{
  const out = {{}};
  arr.forEach(x => {{
    const k = fn(x);
    if (!k) return;
    out[k] = (out[k] || 0) + 1;
  }});
  return out;
}}

function renderAlertCharts() {{
  const generalCounts = countBy(PRODUCTS, p => {{
    const a = String(p.alerta_compra || "");
    if (a === "Sin stock y sin ventas") return null;
    return a;
  }});
  const labels1 = Object.keys(generalCounts);
  const values1 = labels1.map(k => generalCounts[k]);

  Plotly.newPlot("alertGeneralChart", [{{
    type:"pie",
    labels: labels1,
    values: values1,
    hole:.55,
    marker: {{colors: labels1.map(alertaColor)}},
    textinfo:"percent"
  }}], {{
    height:420,
    margin:{{l:10,r:10,t:10,b:10}},
    font:{{family:"Inter, Arial", color:"#111827"}},
    showlegend:true
  }}, {{displayModeBar:false, responsive:true}});
}}

function renderStockChart() {{
  const sums = [
    ["Odoo Cuautitlán", sum(PRODUCTS, "stock_odoo_cuautitlan")],
    ["Amazon FBA", sum(PRODUCTS, "stock_amazon_fba")],
    ["Mercado Libre Full", sum(PRODUCTS, "stock_meli_full")],
    ["Walmart WFS", sum(PRODUCTS, "stock_walmart_wfs")],
    ["Liverpool 99MIN", sum(PRODUCTS, "stock_liverpool_99min")]
  ];

  Plotly.newPlot("stockChart", [{{
    type:"bar",
    x:sums.map(x => x[0]),
    y:sums.map(x => x[1]),
    text:sums.map(x => fmtNum(x[1])),
    textposition:"outside",
    marker:{{color:COLORS.navy}}
  }}], {{
    height:420,
    margin:{{l:55,r:20,t:10,b:55}},
    xaxis:{{title:"Canal"}},
    yaxis:{{title:"Stock"}},
    plot_bgcolor:"white",
    paper_bgcolor:"white",
    font:{{family:"Inter, Arial", color:"#111827"}}
  }}, {{displayModeBar:false, responsive:true}});
}}

function sum(arr, key) {{
  return arr.reduce((acc,p) => acc + Number(p[key] || 0), 0);
}}

function renderProductResults() {{
  const input = document.getElementById("productSearch");
  if (!input) return;
  const q = input.value.toLowerCase().trim();
  let results = PRODUCTS.filter(p => {{
    const text = `${{p.sku_madre}} ${{p.producto_madre}} ${{p.skus_sincronizados}}`.toLowerCase();
    return q === "" || text.includes(q);
  }});

  results = results
    .sort((a,b) => Number(b.ventas_3m_monto || 0) - Number(a.ventas_3m_monto || 0))
    .slice(0, 40);

  const box = document.getElementById("productResults");
  if (results.length === 0) {{
    box.innerHTML = `<div class="empty">No encontré productos con esa búsqueda.</div>`;
    document.getElementById("productDetail").innerHTML = "";
    return;
  }}

  box.innerHTML = `
    <select class="search" id="productSelect" onchange="renderProductDetail(this.value)">
      ${{results.map((p, idx) => `<option value="${{idx}}">${{esc(p.sku_madre)}} | ${{esc(p.producto_madre)}} | Monto 3M: ${{fmtMoney(p.ventas_3m_monto)}} | Unid: ${{fmtNum(p.ventas_3m_unidades)}}</option>`).join("")}}
    </select>
  `;
  window.currentResults = results;
  renderProductDetail(0);
}}

function productImageHtml(img64) {{
  const img = String(img64 || "").trim();
  if (!img || img.toLowerCase() === "nan" || img.toLowerCase() === "false") {{
    return `<div class="product-image-box"><div class="product-image-placeholder">Sin imagen Odoo</div></div>`;
  }}
  return `<div class="product-image-box"><img src="data:image/png;base64,${{img}}" alt="Imagen producto Odoo"></div>`;
}}

function renderProductDetail(index) {{
  const p = window.currentResults[Number(index)];
  if (!p) return;

  const color = alertaColor(p.alerta_compra);
  const recoClass = recClass(p.recomendacion);

  const stockRows = [
    ["Odoo Cuautitlán", p.stock_odoo_cuautitlan],
    ["Amazon FBA", p.stock_amazon_fba],
    ["Mercado Libre Full", p.stock_meli_full],
    ["Walmart WFS", p.stock_walmart_wfs],
    ["Liverpool 99MIN", p.stock_liverpool_99min]
  ].map(r => `<tr><td>${{esc(r[0])}}</td><td>${{fmtNum(r[1])}}</td></tr>`).join("");

  document.getElementById("productDetail").innerHTML = `
    <div class="card">
      <div class="product-hero">
        ${{productImageHtml(p.imagen_odoo_base64)}}
        <div class="product-title">
          <h1>${{esc(p.sku_madre)}} — ${{esc(p.producto_madre)}}</h1>
          <div class="metric-sub">SKUs vinculados por canal en la barra inferior.</div>
        </div>
        <div>${{badgeHtml(p.alerta_compra, color)}}</div>
      </div>

      ${{skuBarHtml(p.sku_por_canal_json, p)}}

      <div class="detail-grid">
        ${{metricHtml("Ventas 3M", fmtNum(p.ventas_3m_unidades), fmtMoney(p.ventas_3m_monto))}}
        ${{metricHtml("Promedio diario", fmtNum(p.venta_diaria_promedio_3m, 2), "solo días con ventas > 0")}}
        ${{metricHtml("Stock total", fmtNum(p.stock_total), "congelado + arribos - ventas")}}
        ${{metricHtml("Días inventario", esc(p.dias_inventario_texto), "stock / promedio diario")}}
        ${{metricHtml("Días desde último arribo", esc(p.dias_desde_ultimo_arribo_texto || "-"), "antigüedad del inventario")}}
        ${{metricHtml("Lead time", fmtNum(p.lead_time_dias), "días")}}
        ${{metricHtml("Sugerencia compra", fmtNum(p.sugerencia_compra), "cobertura 45 días")}}
      </div>

      <div class="detail-layout">
        <div class="card" style="box-shadow:none;border:1px solid #e5e7eb;">
          <h2>Stock por canal</h2>
          <div class="table-wrap" style="max-height:none;">
            <table>
              <thead><tr><th>Canal</th><th>Stock</th></tr></thead>
              <tbody>${{stockRows}}</tbody>
            </table>
          </div>
        </div>

        <div class="card" style="box-shadow:none;border:1px solid #e5e7eb;">
          <h2>Datos importantes</h2>
          <div class="info-list">
            <div class="info-row"><div class="info-label">Ranking Top 80%</div><div class="info-value">${{esc(p.ranking_top80)}}</div></div>
            <div class="info-row"><div class="info-label">Proveedor último arribo</div><div class="info-value">${{esc(p.proveedor_principal || "-")}}</div></div>
            <div class="info-row"><div class="info-label">Proveedores Odoo</div><div class="info-value">${{esc(p.proveedores_odoo_compras || "-")}}</div></div>
            <div class="info-row"><div class="info-label">Última compra/arribo</div><div class="info-value">${{esc(p.ultima_fecha_compra_odoo || "-")}}</div></div>
            <div class="info-row"><div class="info-label">Participación ventas $ <span class="hint" title="Porcentaje que representa este IQ sobre el monto vendido total de los últimos 3 meses.">?</span></div><div class="info-value">${{fmtNum(Number(p.participacion_ventas_3m || 0) * 100, 2)}}%</div></div>
            <div class="info-row"><div class="info-label">Participación acumulada $ <span class="hint" title="Suma acumulada de participación en monto vendido al ordenar los IQ de mayor a menor venta. Sirve para construir el Pareto 80/20.">?</span></div><div class="info-value">${{fmtNum(Number(p.participacion_acumulada_3m || 0) * 100, 2)}}%</div></div>
            <div class="info-row"><div class="info-label">Prioridad</div><div class="info-value">${{esc(p.prioridad)}}</div></div>
            <div class="info-row"><div class="info-label">Fecha corte stock</div><div class="info-value">${{esc(p.fecha_corte_stock || "")}}</div></div>
            <div class="info-row"><div class="info-label">Riesgo días</div><div class="info-value">${{badgeHtml(p.riesgo_sobrestock, riesgoColor(p.riesgo_sobrestock))}}</div></div>
            <div class="info-row"><div class="info-label">Días para comprar</div><div class="info-value">${{fmtNum(p.dias_para_comprar, 1)}}</div></div>
            <div class="info-row"><div class="info-label">Objetivo stock 45 días</div><div class="info-value">${{fmtNum(p.objetivo_stock_45_dias, 0)}}</div></div>
            <div class="info-row"><div class="info-label">Recomendación</div><div class="info-value"><span class="${{recoClass}}">${{esc(p.recomendacion)}}</span></div></div>
          </div>
        </div>
      </div>
    </div>
  `;
}}

document.addEventListener("DOMContentLoaded", () => {{
  renderTopChart();
  renderDiasChart();
  renderAgeChart();
  renderAlertCharts();
  renderStockChart();
  renderProductResults();
}});
</script>
</body>
</html>
"""

ARCHIVO_HTML.write_text(html_doc, encoding="utf-8")

print("\nDASHBOARD HTML V19 GENERADO")
print(f"Archivo: {ARCHIVO_HTML}")


# ============================================
# PUBLICAR AUTOMÁTICAMENTE EN GITHUB
# ============================================

REPO = Path(r"C:\Users\luisf\IQ Tech\DashboardRotacion")

try:

    estado = subprocess.run(
        ["git", "-C", str(REPO), "status", "--porcelain"],
        capture_output=True,
        text=True
    )

    if estado.stdout.strip():

        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        subprocess.run(
            ["git", "-C", str(REPO), "add", "."],
            check=True
        )

        resultado = subprocess.run(
            [
                "git",
                "-C",
                str(REPO),
                "commit",
                "-m",
                f"Actualización automática {fecha}"
            ],
            capture_output=True,
            text=True
        )

        print(resultado.stdout)
        print(resultado.stderr)

        subprocess.run(
            [
                "git",
                "-C",
                str(REPO),
                "push",
                "origin",
                "main"
            ],
            check=True
        )

        print("\nTodos los archivos fueron publicados correctamente en GitHub.")

    else:
        print("\nNo hubo cambios en el repositorio.")

except Exception as e:
    print("\nError al publicar en GitHub:")
    print(e)

