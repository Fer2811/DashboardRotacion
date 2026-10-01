# -*- coding: utf-8 -*-
"""Genera dashboard_B2B_excedente.html a partir de b2b_modelo_excedente.json."""
from __future__ import annotations
from pathlib import Path
import json
import os

try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).resolve().parent / ".env", override=False)
except Exception:
    pass

BASE_DIR = Path(os.getenv("DASHBOARD_ROTACION_DIR", str(Path(__file__).resolve().parent))).resolve()
INPUT = BASE_DIR / "b2b_modelo_excedente.json"
OUTPUT = BASE_DIR / "dashboard_B2B_excedente.html"

if not INPUT.exists():
    raise FileNotFoundError(f"No encontré {INPUT.name}. Corre primero 04_calcular_b2b_excedente_oportunidad.py")

DATA = json.loads(INPUT.read_text(encoding="utf-8"))
data_json = json.dumps(DATA, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")

html = r'''<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>B2B · Excedente + costo de oportunidad</title>
<style>
:root{--bg:#f3f6fb;--panel:#fff;--nav:#10243e;--blue:#2563eb;--green:#059669;--amber:#d97706;--red:#dc2626;--slate:#64748b;--ink:#142033;--line:#dbe4ef;--soft:#eef4fb;--purple:#7c3aed}
*{box-sizing:border-box}body{margin:0;background:var(--bg);font-family:Inter,ui-sans-serif,system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;color:var(--ink)}
header{background:linear-gradient(120deg,#10243e,#17375e);color:#fff;padding:22px 28px;position:sticky;top:0;z-index:10;box-shadow:0 8px 28px rgba(15,35,61,.18)}
header h1{margin:0;font-size:22px}header p{margin:5px 0 0;color:#cbd7e8;font-size:13px}.top{display:flex;justify-content:space-between;gap:20px;align-items:center}.mode{padding:7px 10px;border-radius:999px;background:#ffffff1c;border:1px solid #ffffff33;font-size:12px;font-weight:800}
main{max-width:1500px;margin:0 auto;padding:24px}.banner{padding:13px 15px;border-radius:12px;margin-bottom:18px;background:#fff7ed;border:1px solid #fed7aa;color:#9a3412;display:none}.banner.show{display:block}
.grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px}.kpi,.card{background:var(--panel);border:1px solid var(--line);border-radius:15px;box-shadow:0 4px 18px rgba(15,35,61,.04)}.kpi{padding:15px}.kpi b{font-size:24px;display:block}.kpi span{font-size:12px;color:var(--slate)}
.layout{display:grid;grid-template-columns:1.05fr 1.95fr;gap:16px;margin-top:16px}.card{padding:16px}.title{font-size:15px;font-weight:800;margin-bottom:4px}.sub{font-size:12px;color:var(--slate);margin-bottom:12px}.search{width:100%;padding:11px 12px;border:1px solid var(--line);border-radius:10px;font:inherit}.sku-list{max-height:620px;overflow:auto;margin-top:10px}.sku-item{padding:10px 9px;border-bottom:1px solid #edf2f7;cursor:pointer}.sku-item:hover,.sku-item.active{background:var(--soft)}.sku-code{font-weight:800}.sku-name{font-size:12px;color:var(--slate);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}.chips{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}.chip{padding:3px 7px;border-radius:999px;font-size:11px;background:#eef2ff;color:#3730a3}.chip.green{background:#ecfdf5;color:#047857}.chip.amber{background:#fff7ed;color:#b45309}
.detail-head{display:flex;justify-content:space-between;gap:12px;align-items:flex-start}.detail-head h2{margin:0;font-size:20px}.badge{font-size:11px;font-weight:800;padding:5px 8px;border-radius:999px;background:#ecfdf5;color:#047857}.badge.preview{background:#fff7ed;color:#b45309}
.metrics{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;margin:14px 0}.metric{padding:11px;border:1px solid var(--line);border-radius:11px;background:#fbfdff}.metric b{font-size:17px;display:block}.metric span{font-size:11px;color:var(--slate)}
.sim{display:grid;grid-template-columns:1fr 1fr;gap:14px}.sim-box{border:1px solid var(--line);border-radius:12px;padding:14px}.qty-row{display:grid;grid-template-columns:1fr 120px;gap:10px;align-items:center}input[type=range]{width:100%}.qty-input{padding:9px;border:1px solid var(--line);border-radius:8px;width:100%}.result{margin-top:12px;padding:14px;border-radius:12px;background:#f8fafc;border-left:4px solid var(--blue)}.result-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:8px}.result-grid b{font-size:20px;display:block}.result-grid span{font-size:11px;color:var(--slate)}
.chart{height:240px;border:1px solid var(--line);border-radius:10px;background:white;position:relative;overflow:hidden}.chart svg{width:100%;height:100%}.axis-label{font-size:10px;fill:#64748b}.curve{fill:none;stroke:#2563eb;stroke-width:2}.base-line{stroke:#94a3b8;stroke-dasharray:5 4}.point{fill:#dc2626}.chart-note{font-size:11px;color:var(--slate);margin-top:6px}
table{width:100%;border-collapse:collapse;font-size:12px}th{background:#f1f5f9;text-align:left;padding:8px;border-bottom:1px solid var(--line);position:sticky;top:0}td{padding:8px;border-bottom:1px solid #edf2f7}.num{text-align:right;font-variant-numeric:tabular-nums}.table-wrap{max-height:320px;overflow:auto;border:1px solid var(--line);border-radius:10px}.method{margin-top:16px}.formula{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#0f172a;color:#dbeafe;padding:12px;border-radius:10px;font-size:12px;overflow:auto}.warning-list{font-size:12px;color:#92400e;margin-top:10px}.empty{padding:40px;text-align:center;color:var(--slate)}
@media(max-width:1050px){.layout{grid-template-columns:1fr}.grid{grid-template-columns:repeat(2,1fr)}.metrics{grid-template-columns:repeat(2,1fr)}.sim{grid-template-columns:1fr}}@media(max-width:620px){main{padding:12px}.grid{grid-template-columns:1fr}.result-grid{grid-template-columns:repeat(2,1fr)}header{padding:16px}.top{align-items:flex-start;flex-direction:column}}
</style>
</head>
<body>
<header><div class="top"><div><h1>B2B · Excedente + costo de oportunidad</h1><p>Marketplace define la reserva; B2B puede usar el excedente y pagar el costo de oportunidad si invade inventario protegido.</p></div><div class="mode" id="mode"></div></div></header>
<main>
<div class="banner" id="banner"></div>
<div class="grid" id="globalKpis"></div>
<div class="layout">
  <div class="card"><div class="title">Portafolio</div><div class="sub">Busca por IQ o nombre. Ordenado por excedente seguro B2B.</div><input class="search" id="search" placeholder="Buscar SKU o producto..."/><div class="sku-list" id="skuList"></div></div>
  <div class="card" id="detail"><div class="empty">Selecciona un SKU.</div></div>
</div>
<div class="card method">
  <div class="title">Metodología</div><div class="sub">El ROI mínimo comercial B2B es 10% por pieza. Si el costo de oportunidad Marketplace es mayor, se usa el valor mayor. Para payout incompleto se asumen 30 días.</div>
  <div class="formula">Reserva_c = max(ceil(Forecast_c × LeadTime) − Full_c − Tránsito_c, 0)<br>ExcedenteSeguro = StockBodega − piezas protegidas cubiertas<br>ROI oportunidad_i = ROI Marketplace_c / (1 + r)^(DíaVenta_i + Payout_c)<br>ROI mínimo B2B(q) = promedio[ max(10%, ROI oportunidad_i) por pieza ] · Precio mínimo = Costo × (1 + ROI mínimo B2B)</div>
</div>
</main>
<script>
const DATA=__DATA__;
const fmt=new Intl.NumberFormat('es-MX',{maximumFractionDigits:0});
const fmt1=new Intl.NumberFormat('es-MX',{maximumFractionDigits:1});
const pct=v=>`${(Number(v||0)*100).toFixed(2)}%`;
const money=v=>new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:2}).format(Number(v||0));
const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
let selected=null;
function init(){
 document.getElementById('mode').textContent=DATA.meta.modo;
 if((DATA.meta.warnings||[]).length){const b=document.getElementById('banner');b.classList.add('show');b.innerHTML='<b>Advertencias de datos:</b> '+DATA.meta.warnings.map(esc).join(' · ')}
 document.getElementById('globalKpis').innerHTML=`<div class="kpi"><b>${fmt.format(DATA.meta.n_skus)}</b><span>SKU modelados</span></div><div class="kpi"><b>${fmt.format(DATA.meta.stock_bodega_total)}</b><span>Stock bodega</span></div><div class="kpi"><b>${fmt.format(DATA.meta.reserva_marketplace_total)}</b><span>Reserva Marketplace cubierta</span></div><div class="kpi"><b>${fmt.format(DATA.meta.excedente_seguro_total)}</b><span>Excedente seguro B2B</span></div><div class="kpi"><b>${money(DATA.meta.valor_excedente_total)}</b><span>Valor a costo del excedente</span></div><div class="kpi"><b>${money(DATA.meta.valor_mas_90d_total)}</b><span>Inventario con más de 90 días (todos los canales)</span></div>`;
 renderList();
 document.getElementById('search').addEventListener('input',renderList);
 if(DATA.skus.length)selectSku(DATA.skus[0].sku_madre);
}
function renderList(){const q=(document.getElementById('search').value||'').toLowerCase();const rows=DATA.skus.filter(r=>(r.sku_madre+' '+r.producto_madre).toLowerCase().includes(q));document.getElementById('skuList').innerHTML=rows.map(r=>`<div class="sku-item ${selected===r.sku_madre?'active':''}" onclick="selectSku('${esc(r.sku_madre)}')"><div class="sku-code">${esc(r.sku_madre)}</div><div class="sku-name">${esc(r.producto_madre)}</div><div class="chips"><span class="chip green">Excedente ${fmt.format(r.excedente_b2b_seguro)}</span><span class="chip">Bodega ${fmt.format(r.stock_bodega)}</span><span class="chip amber">ROI todo ${pct(r.roi_min_todo_bodega)}</span>${Number(r.piezas_mas_90d||0)>0?`<span class="chip amber">+90 días: ${fmt.format(r.piezas_mas_90d)} pzas</span>`:''}</div></div>`).join('')||'<div class="empty">Sin coincidencias.</div>'}
function selectSku(sku){selected=sku;renderList();renderDetail();}
function curveAt(row,q){if(!row.curva?.length||q<=0)return null;return row.curva[Math.min(q,row.curva.length)-1]}
function renderDetail(){const row=DATA.skus.find(x=>x.sku_madre===selected),el=document.getElementById('detail');if(!row){el.innerHTML='<div class="empty">Selecciona un SKU.</div>';return}const qmax=row.stock_bodega;const q0=qmax?Math.min(Math.max(row.excedente_b2b_seguro,1),qmax):0;el.innerHTML=`<div class="detail-head"><div><h2>${esc(row.sku_madre)} · ${esc(row.producto_madre||'Producto')}</h2><div class="sub">${esc(row.estado)} · costo: ${row.costo_unitario>0?money(row.costo_unitario):'no disponible'} (${esc(row.fuente_costo)})</div></div><span class="badge ${DATA.meta.modo.startsWith('PREVIEW')?'preview':''}">${DATA.meta.modo.startsWith('PREVIEW')?'PREVIEW':'PRODUCCIÓN'}</span></div>
<div class="metrics"><div class="metric"><b>${fmt.format(row.stock_bodega)}</b><span>Stock bodega</span></div><div class="metric"><b>${fmt.format(row.reserva_marketplace_bodega)}</b><span>Reserva MP cubierta</span></div><div class="metric"><b>${fmt.format(row.excedente_b2b_seguro)}</b><span>Excedente seguro</span></div><div class="metric"><b>${pct(row.roi_base_b2b)}</b><span>Piso comercial B2B</span></div><div class="metric"><b>${pct(row.roi_min_todo_bodega)}</b><span>ROI mínimo todo bodega</span></div><div class="metric"><b>${row.costo_unitario>0?money(row.valor_excedente_seguro||0):'—'}</b><span>Valor del excedente</span></div><div class="metric"><b>${fmt.format(row.piezas_mas_90d||0)}</b><span>Piezas +90 días (${row.costo_unitario>0?money(row.valor_mas_90d||0):'sin costo'})</span></div></div>
<div class="sim"><div class="sim-box"><div class="title">Simulador por cantidad</div><div class="sub">Hasta el excedente seguro no se desplaza cobertura Marketplace. Después, el sistema toma primero las oportunidades protegidas de menor costo.</div>${qmax?`<div class="qty-row"><input id="qtyRange" type="range" min="1" max="${qmax}" value="${q0}"/><input id="qtyInput" class="qty-input" type="number" min="1" max="${qmax}" value="${q0}"/></div><div id="simResult"></div>`:'<div class="empty">Sin stock de bodega.</div>'}</div><div class="sim-box"><div class="title">Curva ROI mínimo B2B vs volumen</div><div class="chart" id="curveChart"></div><div class="chart-note">La pendiente aparece cuando B2B empieza a consumir inventario protegido de Marketplace.</div></div></div>
<div style="margin-top:14px"><div class="title">Reserva por Marketplace</div><div class="table-wrap"><table><thead><tr><th>Canal</th><th class="num">Forecast/día</th><th class="num">Necesidad LT</th><th class="num">Full</th><th class="num">Tránsito</th><th class="num">Reserva bodega</th><th class="num">ROI MP</th><th class="num">Payout</th><th class="num">ROI VP prom.</th></tr></thead><tbody>${row.canales.map(c=>`<tr><td>${esc(c.canal)}</td><td class="num">${fmt1.format(c.forecast_dia)}</td><td class="num">${fmt.format(c.necesidad_lead_time)}</td><td class="num">${fmt.format(c.full)}</td><td class="num">${fmt.format(c.transito)}</td><td class="num"><b>${fmt.format(c.reserva_bodega)}</b></td><td class="num">${pct(c.roi_marketplace)}</td><td class="num">${fmt1.format(c.payout_dias)}d</td><td class="num">${pct(c.roi_vp_prom_reserva)}</td></tr>`).join('')}</tbody></table></div>${row.warnings?.length?`<div class="warning-list"><b>Revisar:</b> ${row.warnings.map(esc).join(' · ')}</div>`:''}</div>`;
 if(qmax){const r=document.getElementById('qtyRange'),i=document.getElementById('qtyInput');const update=(v)=>{let q=Math.max(1,Math.min(qmax,Math.floor(Number(v)||1)));r.value=q;i.value=q;renderSim(row,q);drawCurve(row,q)};r.addEventListener('input',e=>update(e.target.value));i.addEventListener('change',e=>update(e.target.value));update(q0)}else drawCurve(row,0)
}
function renderSim(row,q){const x=curveAt(row,q),safe=Math.min(q,row.excedente_b2b_seguro),protected=Math.max(q-row.excedente_b2b_seguro,0),label=x?.tipo_pieza_marginal==='PROTEGIDA'?`La pieza marginal desplaza ${esc(x.canal_desplazado||'Marketplace')}`:'La pieza marginal sigue dentro del excedente seguro';document.getElementById('simResult').innerHTML=`<div class="result"><div class="result-grid"><div><b>${fmt.format(q)}</b><span>Piezas B2B</span></div><div><b>${pct(x?.roi_min_promedio)}</b><span>ROI mínimo promedio</span></div><div><b>${pct(x?.roi_marginal)}</b><span>ROI pieza marginal</span></div><div><b>${row.costo_unitario>0?money(x?.precio_min_unitario):'—'}</b><span>Precio mínimo / pza</span></div></div><div class="sub" style="margin:10px 0 0">${fmt.format(safe)} excedentes + ${fmt.format(protected)} protegidas. ${label}${x?.dia_venta_desplazada?` · venta MP estimada día ${fmt.format(x.dia_venta_desplazada)}, cobro día ${fmt.format(x.dia_cobro_desplazada)}`:''}.</div></div>`}
function drawCurve(row,qSel){const box=document.getElementById('curveChart');if(!box)return;if(!row.curva?.length){box.innerHTML='<div class="empty">Sin curva.</div>';return}const W=620,H=230,p=36,maxQ=row.curva.length,maxY=Math.max(...row.curva.map(x=>Number(x.roi_min_promedio||0)),Number(row.roi_base_b2b||0),.01)*1.12;const xy=x=>[p+(W-2*p)*(x.q-1)/Math.max(maxQ-1,1),H-p-(H-2*p)*Number(x.roi_min_promedio||0)/maxY];const pts=row.curva.map(x=>xy(x).join(',')).join(' ');const yBase=H-p-(H-2*p)*Number(row.roi_base_b2b||0)/maxY;let point='';if(qSel){const [cx,cy]=xy(row.curva[qSel-1]);point=`<circle class="point" cx="${cx}" cy="${cy}" r="4"/>`}box.innerHTML=`<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none"><line x1="${p}" y1="${H-p}" x2="${W-p}" y2="${H-p}" stroke="#cbd5e1"/><line x1="${p}" y1="${p}" x2="${p}" y2="${H-p}" stroke="#cbd5e1"/><line class="base-line" x1="${p}" y1="${yBase}" x2="${W-p}" y2="${yBase}"/><polyline class="curve" points="${pts}"/>${point}<text class="axis-label" x="${p}" y="${H-10}">1</text><text class="axis-label" x="${W-p-18}" y="${H-10}">${maxQ}</text><text class="axis-label" x="4" y="${p+4}">${(maxY*100).toFixed(1)}%</text><text class="axis-label" x="5" y="${H-p+4}">0%</text></svg>`}
init();
</script>
</body></html>'''.replace('__DATA__', data_json)

OUTPUT.write_text(html, encoding="utf-8")
print(f"HTML generado: {OUTPUT}")
