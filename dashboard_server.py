# -*- coding: utf-8 -*-
"""Backend HTTP/Odoo para Dashboard IQ Tech.

Fase actual:
- Redistribución canal -> canal: transferencia INTERNA en BORRADOR.
- Repartición: transferencia INTERNA desde CUATI/Existencias o CUATI/B2B en BORRADOR.
- SKU duplicados: resolver usando stock REAL por product_id en la ubicación origen.
- Lotes: se consultan y se reportan; no se confunden con product.product.
- Prueba opcional Odoo -> Full como ENTREGA/outgoing en BORRADOR.

No confirma, no reserva, no valida y no marca Hecho.
"""
from __future__ import annotations
from pathlib import Path
from datetime import datetime
import hashlib
import json
import os
import re
import unicodedata
import xmlrpc.client

from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_file

ROOT = Path(__file__).resolve().parent
load_dotenv(ROOT / ".env", override=False)

DASHBOARD_DIR = Path(os.getenv("DASHBOARD_ROTACION_DIR", str(ROOT))).expanduser().resolve()
INDEX_HTML = DASHBOARD_DIR / "index.html"
ACTIONS_FILE = DASHBOARD_DIR / "dashboard_actions.json"
AUDIT_LOG = DASHBOARD_DIR / "dashboard_odoo_drafts_log.jsonl"

ODOO_URL = os.getenv("ODOO_URL", "").strip().rstrip("/")
ODOO_DB = os.getenv("ODOO_DB", "").strip()
ODOO_USER = os.getenv("ODOO_USER", "").strip()
ODOO_API_KEY = os.getenv("ODOO_API_KEY", "").strip()

HOST = os.getenv("DASHBOARD_HOST", "127.0.0.1").strip() or "127.0.0.1"
PORT = int(os.getenv("PORT", os.getenv("DASHBOARD_PORT", "5001")))
DEBUG = os.getenv("DASHBOARD_DEBUG", "0").strip().lower() in {"1", "true", "yes", "si", "sí"}

ORIGIN_PREFIX = "IQTECH_DASH_DRAFT"

ENABLE_FULL_DELIVERY_TEST = os.getenv("ODOO_ENABLE_FULL_DELIVERY_TEST", "0").strip().lower() in {
    "1", "true", "yes", "si", "sí"
}
FULL_TEST_CHANNEL = os.getenv("ODOO_FULL_TEST_CHANNEL", "").strip()
FULL_TEST_PARTNER_ID = int(os.getenv("ODOO_FULL_TEST_PARTNER_ID", "0") or 0)

CHANNEL_LOCATION_MAP = {
    "general": "CUATI/Existencias",
    "amazon": "CUATI/Amazon",
    "mercado libre": "CUATI/MercadoLibre",
    "mercadolibre": "CUATI/MercadoLibre",
    "walmart": "CUATI/Walmart",
    "liverpool": "CUATI/Liverpool",
    "coppel": "CUATI/Coppel",
    "elektra": "CUATI/Elektra",
    "tiktok": "CUATI/TikTok",
    "b2b": "CUATI/B2B",
}

app = Flask(__name__)


def ascii_text(value) -> str:
    s = unicodedata.normalize("NFKD", str(value or "").strip().lower())
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", s)


def norm_loc(value) -> str:
    return ascii_text(value).replace("\\", "/")


def channel_to_location(channel: str) -> str:
    key = ascii_text(channel)
    if key not in CHANNEL_LOCATION_MAP:
        raise ValueError(f"El canal {channel!r} no tiene ubicación Odoo configurada.")
    return CHANNEL_LOCATION_MAP[key]


def execute(models, uid, model, method, args=None, kwargs=None):
    return models.execute_kw(
        ODOO_DB, uid, ODOO_API_KEY, model, method, args or [], kwargs or {}
    )


def m2o_id(value):
    return int(value[0]) if isinstance(value, (list, tuple)) and value else None


def conectar_odoo():
    missing = [
        k for k, v in {
            "ODOO_URL": ODOO_URL,
            "ODOO_DB": ODOO_DB,
            "ODOO_USER": ODOO_USER,
            "ODOO_API_KEY": ODOO_API_KEY,
        }.items() if not v
    ]
    if missing:
        raise RuntimeError("Faltan variables en .env: " + ", ".join(missing))

    common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(ODOO_DB, ODOO_USER, ODOO_API_KEY, {})
    if not uid:
        raise RuntimeError("Odoo rechazó la autenticación.")

    models = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/object", allow_none=True)
    return uid, models


def cargar_acciones() -> dict:
    if not ACTIONS_FILE.exists():
        raise RuntimeError("Falta dashboard_actions.json. Ejecuta primero 03_generar_dashboard.py.")
    data = json.loads(ACTIONS_FILE.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise RuntimeError("dashboard_actions.json inválido.")
    return data


def ubicaciones_cuati(models, uid):
    rows = execute(
        models, uid, "stock.location", "search_read",
        [[("usage", "=", "internal")]],
        {"fields": ["id", "name", "complete_name"], "limit": 0, "order": "complete_name asc"},
    )
    out = []
    for row in rows:
        complete = str(row.get("complete_name") or row.get("name") or "")
        n = norm_loc(complete)
        if n == "cuati" or n.startswith("cuati/"):
            out.append({"id": int(row["id"]), "name": row.get("name") or "", "complete_name": complete})
    return out


def buscar_loc(models, uid, complete_name: str):
    wanted = norm_loc(complete_name)
    exact = [r for r in ubicaciones_cuati(models, uid) if norm_loc(r["complete_name"]) == wanted]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise RuntimeError(f"La ubicación {complete_name} está duplicada en Odoo.")
    raise ValueError(f"No encontré la ubicación interna {complete_name} en Odoo.")


def stock_disponible(models, uid, product_id: int, location_id: int):
    quant_fields = execute(
        models, uid, "stock.quant", "fields_get", [], {"attributes": ["type"]}
    )
    fields = ["location_id", "quantity", "reserved_quantity"]
    if "lot_id" in quant_fields:
        fields.append("lot_id")

    rows = execute(
        models, uid, "stock.quant", "search_read",
        [[("product_id", "=", int(product_id)), ("location_id", "child_of", int(location_id))]],
        {"fields": fields, "limit": 0},
    )

    quantity = sum(float(r.get("quantity") or 0) for r in rows)
    reserved = sum(float(r.get("reserved_quantity") or 0) for r in rows)

    lots = []
    for row in rows:
        available = float(row.get("quantity") or 0) - float(row.get("reserved_quantity") or 0)
        if available <= 1e-9:
            continue
        lot = row.get("lot_id")
        lots.append({
            "lot_id": m2o_id(lot),
            "lote": lot[1] if isinstance(lot, (list, tuple)) and len(lot) > 1 else "Sin lote",
            "location_id": m2o_id(row.get("location_id")),
            "cantidad": float(row.get("quantity") or 0),
            "reservada": float(row.get("reserved_quantity") or 0),
            "disponible": available,
        })

    return {
        "quantity": quantity,
        "reserved_quantity": reserved,
        "available": quantity - reserved,
        "lotes": lots,
    }


def product_rows(models, uid, sku: str):
    return execute(
        models, uid, "product.product", "search_read",
        [[("default_code", "=", sku)]],
        {
            "fields": ["id", "default_code", "display_name", "uom_id", "product_tmpl_id"],
            "limit": 50,
        },
    )


def name_score(a: str, b: str) -> float:
    aa = set(re.findall(r"[a-z0-9]+", ascii_text(a)))
    bb = set(re.findall(r"[a-z0-9]+", ascii_text(b)))
    return len(aa & bb) / max(len(aa | bb), 1)


def resolver_producto_por_origen(models, uid, sku: str, location_id: int, qty: int, expected_name: str = ""):
    """Resuelve default_code duplicado usando stock en el ORIGEN.

    Los lotes NO cuentan como product.product duplicado. Se consultan dentro de
    stock_disponible() una vez elegido el product_id.
    """
    rows = product_rows(models, uid, sku)
    if not rows:
        raise ValueError(f"No encontré {sku} como product.product en Odoo.")

    evaluated = []
    for product in rows:
        stock = stock_disponible(models, uid, product["id"], location_id)
        evaluated.append({
            "product": product,
            "stock": stock,
            "name_score": name_score(product.get("display_name"), expected_name) if expected_name else 0.0,
        })

    enough = [x for x in evaluated if x["stock"]["available"] + 1e-9 >= qty]
    if len(enough) == 1:
        return enough[0]

    if len(enough) > 1 and expected_name:
        ranked = sorted(
            enough,
            key=lambda x: (-x["name_score"], -x["stock"]["available"], int(x["product"]["id"])),
        )
        # Desempata por nombre solo cuando la diferencia es suficientemente clara.
        if ranked[0]["name_score"] > 0 and (
            len(ranked) == 1 or ranked[0]["name_score"] > ranked[1]["name_score"] + 0.15
        ):
            return ranked[0]

    if len(rows) == 1:
        return evaluated[0]

    positives = [x for x in evaluated if x["stock"]["available"] > 1e-9]
    if len(positives) == 1:
        return positives[0]

    diag = "; ".join(
        f"product_id {x['product']['id']} · {x['product'].get('display_name')} · disponible {x['stock']['available']:.0f}"
        for x in evaluated
    )

    if not enough:
        raise ValueError(
            f"{sku} existe en varios product.product, pero ninguno tiene por sí solo {qty} piezas "
            f"disponibles en el origen. {diag}"
        )

    raise ValueError(
        f"{sku} existe en varios product.product con stock suficiente en el mismo origen; "
        f"no escogeré uno arbitrariamente. {diag}"
    )


def action_id(generated_at: str, kind: str, sku: str, origin: str, dest: str, qty: int) -> str:
    return f"{generated_at}|{kind}|{sku}|{origin}|{dest}|{int(qty)}"


def validate_common(payload: dict, kind: str, list_key: str):
    actions = cargar_acciones()
    generated_at = str(payload.get("generated_at") or "").strip()
    sku = str(payload.get("sku_madre") or "").strip().upper()
    origin = str(payload.get("canal_origen") or "").strip()
    dest = str(payload.get("canal_destino") or "").strip()
    try:
        qty = int(float(payload.get("cantidad_sugerida") or 0))
    except Exception:
        qty = 0
    request_id = str(payload.get("request_id") or "").strip()

    if generated_at != str(actions.get("generated_at") or ""):
        raise ValueError("Dashboard desactualizado. Regenera/recarga antes de ejecutar.")
    if not sku or not origin or not dest or qty <= 0:
        raise ValueError("Solicitud incompleta.")

    expected = action_id(generated_at, kind, sku, origin, dest, qty)
    if request_id != expected:
        raise ValueError("request_id inválido.")

    found = None
    for row in actions.get(list_key, []) or []:
        if (
            str(row.get("sku_madre") or "").strip().upper() == sku
            and ascii_text(row.get("canal_origen")) == ascii_text(origin)
            and ascii_text(row.get("canal_destino")) == ascii_text(dest)
            and int(float(row.get("cantidad_sugerida") or 0)) == qty
        ):
            found = row
            break

    if not found:
        raise ValueError("La acción ya no existe en el snapshot autorizado del dashboard.")

    return {
        "generated_at": generated_at,
        "sku": sku,
        "origin_channel": origin,
        "dest_channel": dest,
        "qty": qty,
        "request_id": expected,
        "action": found,
    }


def validar_redistribucion(payload: dict):
    action = validate_common(payload, "REDIST", "redistributions")
    tipo = ascii_text(action["action"].get("tipo_stock_origen"))
    manual = ascii_text(action["action"].get("check_manual"))
    if not tipo.startswith("odoo") or "full" in tipo:
        raise ValueError("La sugerencia usa Full o no es Odoo puro; requiere revisión operativa.")
    if manual.startswith("si"):
        raise ValueError("La sugerencia está marcada CHECK MANUAL.")
    return action


def validar_reparticion(payload: dict):
    return validate_common(payload, "DIST", "distribution")


def origin_token(request_id: str) -> str:
    return f"{ORIGIN_PREFIX}_{hashlib.sha256(request_id.encode('utf-8')).hexdigest()[:16]}"


def existing_by_origin(models, uid, origin: str):
    rows = execute(
        models, uid, "stock.picking", "search_read",
        [[("origin", "=", origin)]],
        {
            "fields": ["id", "name", "state", "origin", "picking_type_id", "location_id", "location_dest_id"],
            "limit": 5,
        },
    )
    return rows[0] if rows else None


def picking_type_internal(models, uid, source_id: int, dest_id: int):
    rows = execute(
        models, uid, "stock.picking.type", "search_read",
        [[("code", "=", "internal")]],
        {
            "fields": ["id", "name", "sequence_code", "warehouse_id", "default_location_src_id", "default_location_dest_id"],
            "limit": 0,
        },
    )
    if not rows:
        raise RuntimeError("No encontré tipo de operación interna.")

    def score(row):
        return (
            (4 if m2o_id(row.get("default_location_src_id")) == source_id else 0)
            + (4 if m2o_id(row.get("default_location_dest_id")) == dest_id else 0)
            + (1 if not m2o_id(row.get("default_location_src_id")) else 0)
            + (1 if not m2o_id(row.get("default_location_dest_id")) else 0)
        )

    return sorted(rows, key=lambda r: (-score(r), int(r["id"])))[0]


def picking_type_outgoing(models, uid, source_id: int):
    rows = execute(
        models, uid, "stock.picking.type", "search_read",
        [[("code", "=", "outgoing")]],
        {
            "fields": ["id", "name", "warehouse_id", "default_location_src_id", "default_location_dest_id"],
            "limit": 0,
        },
    )
    if not rows:
        raise RuntimeError("No encontré tipo de operación Entregas/outgoing.")
    return sorted(
        rows,
        key=lambda r: (0 if m2o_id(r.get("default_location_src_id")) == source_id else 1, int(r["id"])),
    )[0]


def construir_move_vals(models, uid, product: dict, qty: int, source_id: int, dest_id: int, picking_id: int, sku: str):
    fields = execute(
        models, uid, "stock.move", "fields_get", [], {"attributes": ["type", "required", "readonly"]}
    )
    vals = {
        "product_id": int(product["id"]),
        "location_id": int(source_id),
        "location_dest_id": int(dest_id),
        "picking_id": int(picking_id),
    }
    if "product_uom_qty" in fields:
        vals["product_uom_qty"] = float(qty)
    elif "quantity" in fields:
        vals["quantity"] = float(qty)
    else:
        raise RuntimeError("stock.move no expone un campo de cantidad compatible.")

    uom_id = m2o_id(product.get("uom_id"))
    if uom_id and "product_uom" in fields:
        vals["product_uom"] = uom_id

    if "description_picking" in fields:
        vals["description_picking"] = f"Dashboard IQ Tech · {sku}"
    elif "name" in fields:
        vals["name"] = f"Dashboard IQ Tech · {sku}"

    return {k: v for k, v in vals.items() if k in fields}


def append_audit(payload: dict):
    row = {"logged_at": datetime.now().isoformat(timespec="seconds"), **payload}
    with AUDIT_LOG.open("a", encoding="utf-8") as f:
        f.write(json.dumps(row, ensure_ascii=False) + "\n")


def create_internal_draft(action: dict, kind: str):
    uid, models = conectar_odoo()
    source = buscar_loc(models, uid, channel_to_location(action["origin_channel"]))
    dest = buscar_loc(models, uid, channel_to_location(action["dest_channel"]))

    expected_name = action["action"].get("producto_madre", "")
    resolved = resolver_producto_por_origen(
        models, uid, action["sku"], source["id"], action["qty"], expected_name
    )
    product = resolved["product"]
    stock = resolved["stock"]

    if stock["available"] + 1e-9 < action["qty"]:
        raise ValueError(
            f"Stock insuficiente en {source['complete_name']}. Disponible {stock['available']}, solicitado {action['qty']}."
        )

    origin = origin_token(action["request_id"])
    existing = existing_by_origin(models, uid, origin)
    if existing:
        return {
            "duplicate": True,
            "picking": existing,
            "producto": product,
            "stock": stock,
            "resolucion_producto": {
                "product_id": product["id"],
                "candidatos_mismo_sku": len(product_rows(models, uid, action["sku"])),
                "lotes_disponibles": stock.get("lotes", []),
            },
        }

    picking_type = picking_type_internal(models, uid, source["id"], dest["id"])
    picking_id = execute(
        models, uid, "stock.picking", "create",
        [{
            "picking_type_id": int(picking_type["id"]),
            "location_id": int(source["id"]),
            "location_dest_id": int(dest["id"]),
            "origin": origin,
        }],
    )

    try:
        move_id = execute(
            models, uid, "stock.move", "create",
            [construir_move_vals(models, uid, product, action["qty"], source["id"], dest["id"], picking_id, action["sku"])],
        )
    except Exception:
        try:
            execute(models, uid, "stock.picking", "unlink", [[int(picking_id)]])
        except Exception:
            pass
        raise

    picking = execute(
        models, uid, "stock.picking", "read", [[int(picking_id)]],
        {"fields": ["id", "name", "state", "origin", "picking_type_id", "location_id", "location_dest_id"]},
    )[0]

    if picking.get("state") != "draft":
        raise RuntimeError(f"Protección: Odoo creó estado {picking.get('state')!r}, no draft.")

    append_audit({
        "tipo": kind,
        "request_id": action["request_id"],
        "sku": action["sku"],
        "canal_origen": action["origin_channel"],
        "canal_destino": action["dest_channel"],
        "cantidad": action["qty"],
        "product_id": product["id"],
        "picking_id": picking["id"],
        "picking_name": picking.get("name"),
        "move_id": move_id,
        "state": picking.get("state"),
    })

    return {
        "duplicate": False,
        "picking": picking,
        "move_id": int(move_id),
        "producto": product,
        "stock": stock,
        "resolucion_producto": {
            "product_id": product["id"],
            "candidatos_mismo_sku": len(product_rows(models, uid, action["sku"])),
            "lotes_disponibles": stock.get("lotes", []),
        },
    }


@app.get("/")
def home():
    if not INDEX_HTML.exists():
        return "<h2>Falta index.html</h2><p>Ejecuta 03_generar_dashboard.py.</p>", 404
    return send_file(INDEX_HTML)


@app.get("/api/health")
def health():
    return jsonify(
        ok=True,
        dashboard_exists=INDEX_HTML.exists(),
        actions_exists=ACTIONS_FILE.exists(),
        odoo_configured=all([ODOO_URL, ODOO_DB, ODOO_USER, ODOO_API_KEY]),
        host=HOST,
        delivery_test_enabled=ENABLE_FULL_DELIVERY_TEST,
    )


@app.get("/api/odoo/test")
def test_odoo():
    try:
        uid, _ = conectar_odoo()
        common = xmlrpc.client.ServerProxy(f"{ODOO_URL}/xmlrpc/2/common", allow_none=True)
        return jsonify(
            ok=True,
            uid=uid,
            server_version=common.version().get("server_version"),
            mensaje="Conexión con Odoo correcta.",
        )
    except Exception as exc:
        return jsonify(ok=False, mensaje=str(exc)), 400


@app.post("/api/odoo/redistribution/draft")
def redistribution_draft():
    try:
        action = validar_redistribucion(request.get_json(silent=True) or {})
        result = create_internal_draft(action, "redistribucion")
        return jsonify(
            ok=True,
            modo="WRITE_DRAFT",
            mensaje="Transferencia interna creada EN BORRADOR. No se confirmó, reservó ni validó.",
            cantidad=action["qty"],
            **result,
        )
    except Exception as exc:
        return jsonify(ok=False, modo="WRITE_DRAFT", mensaje=str(exc)), 400


@app.post("/api/odoo/distribution/draft")
def distribution_draft():
    try:
        action = validar_reparticion(request.get_json(silent=True) or {})
        result = create_internal_draft(action, "reparticion")
        return jsonify(
            ok=True,
            modo="WRITE_DRAFT",
            mensaje="Repartición creada EN BORRADOR. No se confirmó, reservó ni validó.",
            cantidad=action["qty"],
            **result,
        )
    except Exception as exc:
        return jsonify(ok=False, modo="WRITE_DRAFT", mensaje=str(exc)), 400


# Prueba opcional: Odoo -> Full como ENTREGA (code='outgoing'), NO traslado interno.
@app.post("/api/odoo/full-delivery/draft")
def full_delivery_draft():
    try:
        if not ENABLE_FULL_DELIVERY_TEST:
            raise ValueError(
                "Prueba de Entregas desactivada. Usa ODOO_ENABLE_FULL_DELIVERY_TEST=1 solo para una prueba controlada."
            )
        if not FULL_TEST_CHANNEL or not FULL_TEST_PARTNER_ID:
            raise ValueError("Configura ODOO_FULL_TEST_CHANNEL y ODOO_FULL_TEST_PARTNER_ID.")

        payload = request.get_json(silent=True) or {}
        actions = cargar_acciones()
        generated_at = str(payload.get("generated_at") or "")
        sku = str(payload.get("sku_madre") or "").upper()
        channel = str(payload.get("canal") or "")
        qty = int(float(payload.get("cantidad_sugerida") or 0))

        if generated_at != str(actions.get("generated_at") or ""):
            raise ValueError("Dashboard desactualizado.")
        if ascii_text(channel) != ascii_text(FULL_TEST_CHANNEL):
            raise ValueError(f"La prueba está limitada a {FULL_TEST_CHANNEL}.")

        found = next(
            (
                r for r in actions.get("full_deliveries", [])
                if str(r.get("sku_madre") or "").upper() == sku
                and ascii_text(r.get("canal")) == ascii_text(channel)
                and int(float(r.get("cantidad_sugerida") or 0)) == qty
            ),
            None,
        )
        if not found:
            raise ValueError("La entrega no coincide con una sugerencia vigente.")

        uid, models = conectar_odoo()
        source = buscar_loc(models, uid, channel_to_location(channel))
        resolved = resolver_producto_por_origen(
            models, uid, sku, source["id"], qty, found.get("producto_madre", "")
        )
        product = resolved["product"]
        stock = resolved["stock"]
        if stock["available"] + 1e-9 < qty:
            raise ValueError(f"Stock insuficiente: {stock['available']}.")

        picking_type = picking_type_outgoing(models, uid, source["id"])
        dest_id = m2o_id(picking_type.get("default_location_dest_id"))
        if not dest_id:
            raise RuntimeError("El tipo Entregas no tiene ubicación destino cliente predeterminada.")

        request_id = f"{generated_at}|DELIVERY|{sku}|{channel}|{qty}|{FULL_TEST_PARTNER_ID}"
        origin = origin_token(request_id)
        existing = existing_by_origin(models, uid, origin)
        if existing:
            return jsonify(ok=True, duplicate=True, picking=existing, mensaje="La entrega de prueba ya existía.")

        picking_id = execute(
            models, uid, "stock.picking", "create",
            [{
                "picking_type_id": int(picking_type["id"]),
                "partner_id": FULL_TEST_PARTNER_ID,
                "location_id": int(source["id"]),
                "location_dest_id": int(dest_id),
                "origin": origin,
            }],
        )
        try:
            move_id = execute(
                models, uid, "stock.move", "create",
                [construir_move_vals(models, uid, product, qty, source["id"], dest_id, picking_id, sku)],
            )
        except Exception:
            try:
                execute(models, uid, "stock.picking", "unlink", [[int(picking_id)]])
            except Exception:
                pass
            raise

        picking = execute(
            models, uid, "stock.picking", "read", [[int(picking_id)]],
            {"fields": ["id", "name", "state", "origin", "picking_type_id", "partner_id", "location_id", "location_dest_id"]},
        )[0]
        if picking.get("state") != "draft":
            raise RuntimeError("La entrega no quedó en draft.")

        append_audit({
            "tipo": "entrega_full_prueba",
            "sku": sku,
            "canal": channel,
            "cantidad": qty,
            "picking_id": picking["id"],
            "picking_name": picking.get("name"),
            "move_id": move_id,
            "state": "draft",
        })

        return jsonify(
            ok=True,
            duplicate=False,
            picking=picking,
            move_id=int(move_id),
            resolucion_producto={
                "product_id": product["id"],
                "lotes_disponibles": stock.get("lotes", []),
            },
            mensaje="ENTREGA creada en BORRADOR. No se confirmó, reservó ni validó.",
        )
    except Exception as exc:
        return jsonify(ok=False, mensaje=str(exc)), 400


if __name__ == "__main__":
    print("\nIQ TECH · Dashboard HTTP + Odoo (BORRADOR)")
    print(f"Servidor: http://{HOST}:{PORT}")
    print("Redistribución y Repartición activas. Entrega Full solo si se habilita explícitamente.\n")
    app.run(host=HOST, port=PORT, debug=DEBUG)
