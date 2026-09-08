from __future__ import annotations

import csv
from datetime import datetime, timedelta, timezone
from pathlib import Path

from dash import Dash, Input, Output, State, ctx, dcc, html, dash_table, no_update
from dash.exceptions import PreventUpdate

from excel_export import write_simulation_excel
from drive_data import (
    MADRYN_FILE_ID, PRICE_FILE_ID, TRELEW_FILE_ID,
    load_catalog_from_drive, load_freshness_from_drive,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
MARGEN_SUGERIDO = 1.30
AR_TZ = timezone(timedelta(hours=-3), name="ART")
MONTHS_ES = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre",
]


def to_float(value, default=0.0):
    try:
        if value in (None, ""):
            return default
        return float(value)
    except (TypeError, ValueError):
        return default


def load_catalog_fallback():
    catalog = {}
    with (DATA_DIR / "catalogo.csv").open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            sku = str(row["sku"]).strip()
            catalog[sku] = {
                "sku": sku,
                "descripcion": row["descripcion"],
                "unidades_bulto": to_float(row["unidades_bulto"]),
                "bultos_pallet": to_float(row["bultos_pallet"]),
                "precio_base_bulto": to_float(row["precio_base_bulto"]),
                "p_final_bulto": to_float(row["p_final_bulto"]),
                "final_actual_bulto": to_float(row["final_actual_bulto"]),
                "final_actual_unidad": to_float(row["final_actual_unidad"]),
                "evidencia": row.get("evidencia", ""),
            }
    return catalog


def load_expiry_fallback():
    rows = []
    with (DATA_DIR / "vencimientos.csv").open(encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            rows.append(row)
    return rows


def reload_sources():
    """Load each operational source from Google Drive, with bundled CSV fallback."""
    fallback_catalog = load_catalog_fallback()
    fallback_expiry = load_expiry_fallback()
    status = {}

    try:
        catalog = load_catalog_from_drive(PRICE_FILE_ID)
        status["Precios"] = {"ok": True, "detail": f"Drive · {len(catalog)} SKU"}
    except Exception as exc:
        catalog = fallback_catalog
        status["Precios"] = {"ok": False, "detail": f"respaldo local · {type(exc).__name__}"}

    expiry = []
    for branch, file_id in (("Trelew", TRELEW_FILE_ID), ("Madryn", MADRYN_FILE_ID)):
        try:
            rows = load_freshness_from_drive(file_id, branch, catalog)
            expiry.extend(rows)
            status[f"Frescura {branch}"] = {"ok": True, "detail": f"Drive · {len(rows)} lotes"}
        except Exception as exc:
            rows = [r for r in fallback_expiry if r.get("sucursal") == branch]
            expiry.extend(rows)
            status[f"Frescura {branch}"] = {"ok": False, "detail": f"respaldo local · {type(exc).__name__}"}

    status["loaded_at"] = datetime.now(AR_TZ).strftime("%d/%m/%Y %H:%M")
    return catalog, expiry, status


CATALOG, EXPIRY, DATA_STATUS = reload_sources()


def now_ar():
    return datetime.now(AR_TZ)


def money(value, decimals=2):
    if value is None:
        return "—"
    text = f"{value:,.{decimals}f}"
    return "$ " + text.replace(",", "X").replace(".", ",").replace("X", ".")


def pct(value, decimals=2):
    if value is None:
        return "—"
    return f"{value:.{decimals}f}%".replace(".", ",")


def number(value, decimals=2):
    if value is None:
        return "—"
    text = f"{value:,.{decimals}f}"
    return text.replace(",", "X").replace(".", ",").replace("X", ".")


def pretty_date(iso_date):
    if not iso_date:
        return "—"
    try:
        return datetime.strptime(iso_date, "%Y-%m-%d").strftime("%d/%m/%Y")
    except ValueError:
        return iso_date


def calculate_from_discount(product, discount_pct, bultos=1):
    discount = discount_pct / 100.0
    p_final_bulto = product["p_final_bulto"]
    precio_base = product["precio_base_bulto"]
    units = product["unidades_bulto"]

    bonif_bulto = precio_base * discount

    # Regla validada contra ERP:
    # - Sin bonificación, el final comercial incluye el recargo observado del 3%
    #   sobre Precio Base (ej. SKU 32642: 28.493,062 + 642,919 = 29.135,981).
    # - Con bonificación > 0%, el ERP calcula el neto sobre P. Final.
    if abs(discount_pct) < 1e-12:
        final_bulto = product.get("final_actual_bulto", p_final_bulto + precio_base * 0.03)
    else:
        final_bulto = p_final_bulto * (1 - discount)

    final_unit = final_bulto / units if units else 0
    suggested = final_unit * MARGEN_SUGERIDO

    return {
        "discount_pct": discount_pct,
        "bonif_bulto": bonif_bulto,
        "final_bulto": final_bulto,
        "final_unit": final_unit,
        "suggested": suggested,
        "gasto_total": bonif_bulto * bultos,
    }


def calculate_from_suggested(product, suggested, bultos=1):
    p_final_bulto = product["p_final_bulto"]
    precio_base = product["precio_base_bulto"]
    units = product["unidades_bulto"]

    final_unit = suggested / MARGEN_SUGERIDO
    final_bulto = final_unit * units
    discount_pct = (1 - final_bulto / p_final_bulto) * 100 if p_final_bulto else 0
    bonif_bulto = precio_base * (discount_pct / 100.0)

    return {
        "discount_pct": discount_pct,
        "bonif_bulto": bonif_bulto,
        "final_bulto": final_bulto,
        "final_unit": final_unit,
        "suggested": suggested,
        "gasto_total": bonif_bulto * bultos,
    }


def lots_for_sku(sku, branch="Todas"):
    current = now_ar()
    lots = []
    for row in EXPIRY:
        if row["sku"].strip() != sku:
            continue
        if branch != "Todas" and row["sucursal"] != branch:
            continue
        iso = row.get("vencimiento", "")
        try:
            dt = datetime.strptime(iso, "%Y-%m-%d")
        except ValueError:
            dt = None
        stock = to_float(row.get("stock_lote_bultos"))
        lots.append({
            "sucursal": row.get("sucursal", ""),
            "vencimiento": pretty_date(iso),
            "vencimiento_iso": iso,
            "estado": row.get("estado", ""),
            "stock": stock,
            "orden": to_float(row.get("orden_lote")),
            "dias": to_float(row.get("dias_bloqueo")),
            "mes_corriente": bool(
                dt and stock > 0 and dt.year == current.year and dt.month == current.month
            ),
        })
    lots.sort(key=lambda r: r["vencimiento_iso"] or "9999-12-31")
    return lots


def build_simulation_entry(sku_raw, mode, value, bultos, branch):
    sku = str(sku_raw or "").strip()
    product = CATALOG.get(sku)
    if not sku:
        return None, "Ingresá un SKU antes de sumarlo."
    if not product:
        return None, f"El SKU {sku} no está en el catálogo."
    if product["p_final_bulto"] <= 0 or product["precio_base_bulto"] <= 0 or product["unidades_bulto"] <= 0:
        return None, f"El SKU {sku} no tiene precios/unidades válidos para simular."
    if value is None:
        return None, "Completá el descuento o el sugerido antes de sumar el SKU."

    bultos_num = max(to_float(bultos, 1), 0)
    if mode == "suggested":
        result = calculate_from_suggested(product, max(float(value), 0), bultos_num)
    else:
        result = calculate_from_discount(product, min(max(float(value), 0), 100), bultos_num)

    if result["discount_pct"] < 0 or result["discount_pct"] > 100:
        return None, "El cálculo genera un descuento fuera del rango 0%–100%. Revisá el valor ingresado."

    lots = lots_for_sku(sku, branch)
    stock_total = sum(lot["stock"] for lot in lots)
    current_month_lots = [lot for lot in lots if lot["mes_corriente"]]

    valid_dates = [lot["vencimiento_iso"] for lot in lots if lot["vencimiento_iso"]]
    next_date = min(valid_dates) if valid_dates else ""
    stock_next = sum(lot["stock"] for lot in lots if lot["vencimiento_iso"] == next_date) if next_date else 0

    if lots:
        expiry_summary = " | ".join(
            f"{lot['sucursal']} {lot['vencimiento']}: {number(lot['stock'], 2)} btos"
            for lot in lots
        )
    else:
        expiry_summary = "Sin frescura informada"

    current = now_ar()
    alert = "⚠ VENCE ESTE MES" if current_month_lots else "OK"
    current_summary = " | ".join(
        f"{lot['sucursal']} {lot['vencimiento']}: {number(lot['stock'], 2)} btos"
        for lot in current_month_lots
    )

    return {
        "key": f"{sku}|{branch}",
        "sku": sku,
        "descripcion": product["descripcion"],
        "sucursal": branch,
        "bultos": bultos_num,
        "unidades_bulto": product["unidades_bulto"],
        "precio_base_bulto": product["precio_base_bulto"],
        "p_final_bulto": product["p_final_bulto"],
        "discount_pct": result["discount_pct"],
        "bonif_bulto": result["bonif_bulto"],
        "final_bulto": result["final_bulto"],
        "final_unit": result["final_unit"],
        "suggested": result["suggested"],
        "gasto_total": result["gasto_total"],
        "stock_total": stock_total,
        "proximo_vencimiento_iso": next_date,
        "stock_proximo": stock_next,
        "vencimientos_resumen": expiry_summary,
        "alerta_mes": alert,
        "vencimientos_mes": current_summary,
        "mes_alerta": f"{MONTHS_ES[current.month - 1]} {current.year}",
        "lots": lots,
    }, None


def card(title, value_id, subtitle=None, accent=False):
    children = [html.Div(title, className="metric-label"), html.Div(id=value_id, className="metric-value")]
    if subtitle:
        children.append(html.Div(subtitle, className="metric-subtitle"))
    return html.Div(children, className="metric-card accent" if accent else "metric-card")


def data_status_children():
    items = []
    for name in ("Precios", "Frescura Trelew", "Frescura Madryn"):
        info = DATA_STATUS.get(name, {})
        ok = bool(info.get("ok"))
        items.append(
            html.Div(
                [
                    html.Span("●", className="source-dot ok" if ok else "source-dot warn"),
                    html.Div([html.Strong(name), html.Span(info.get("detail", "sin datos"))]),
                ],
                className="source-item",
            )
        )
    return items


app = Dash(__name__, title="Simulador Multi-SKU · ERP")
server = app.server

app.layout = html.Div(
    className="app-shell",
    children=[
        dcc.Store(id="simulation-store", storage_type="session"),
        dcc.Store(id="data-version", data=0),
        dcc.Download(id="excel-download"),
        html.Div(
            className="header",
            children=[
                html.Div(
                    [
                        html.Div([html.Span("SIMULADOR COMERCIAL"), html.Span("VERSIÓN MULTI-SKU", className="version-badge")], className="eyebrow header-eyebrow"),
                        html.H1("Precio sugerido y descuento por SKU"),
                        html.P(
                            "Simulá un SKU, sumalo al listado y exportá todos los escenarios juntos con stock y vencimientos."
                        ),
                    ]
                ),
                html.Div("Factor sugerido: × 1,30", className="factor-chip"),
            ],
        ),
        html.Div(
            className="panel source-panel",
            children=[
                html.Div(
                    [
                        html.Div("FUENTES DE DATOS", className="eyebrow"),
                        html.Div(id="data-source-status", className="source-items", children=data_status_children()),
                        html.Div(
                            f"Última carga: {DATA_STATUS.get('loaded_at', '—')} · Google Drive",
                            id="data-loaded-at",
                            className="source-loaded-at",
                        ),
                    ],
                    className="source-copy",
                ),
                html.Div(
                    [
                        html.Button("↻ Actualizar datos", id="refresh-data-btn", n_clicks=0, className="secondary-btn refresh-btn"),
                        html.Div(id="data-refresh-feedback", className="refresh-feedback"),
                    ],
                    className="source-actions",
                ),
            ],
        ),
        html.Div(
            className="panel input-panel",
            children=[
                html.Div(
                    className="field",
                    children=[
                        html.Label("SKU"),
                        dcc.Input(id="sku-input", type="text", value="2218", debounce=True, placeholder="Ej.: 2218", className="text-input"),
                    ],
                ),
                html.Div(
                    className="field mode-field",
                    children=[
                        html.Label("Quiero calcular desde"),
                        dcc.RadioItems(
                            id="mode-input",
                            options=[
                                {"label": "% de descuento", "value": "discount"},
                                {"label": "Precio sugerido", "value": "suggested"},
                            ],
                            value="discount",
                            inline=True,
                            className="radio-group",
                        ),
                    ],
                ),
                html.Div(
                    className="field",
                    children=[
                        html.Label(id="value-label", children="Descuento (%)"),
                        dcc.Input(id="value-input", type="number", value=20, min=0, step=0.01, className="text-input"),
                    ],
                ),
                html.Div(
                    className="field small-field",
                    children=[
                        html.Label("Bultos a bonificar"),
                        dcc.Input(id="bultos-input", type="number", value=1, min=0, step=0.01, className="text-input"),
                    ],
                ),
            ],
        ),
        html.Div(id="status-message", className="status-message"),
        html.Div(
            id="product-panel",
            className="panel product-panel",
            children=[
                html.Div([html.Div(id="product-sku", className="product-sku"), html.H2(id="product-description")]),
                html.Div(
                    className="product-meta",
                    children=[
                        html.Div([html.Span("Unidades/bulto"), html.Strong(id="product-units")]),
                        html.Div([html.Span("Precio base/bulto"), html.Strong(id="current-base")]),
                        html.Div([html.Span("Final sin descuento/bulto"), html.Strong(id="current-bulto")]),
                        html.Div([html.Span("Sugerido s/desc. ×1,30"), html.Strong(id="current-suggested")]),
                    ],
                ),
            ],
        ),
        html.Div(
            className="metrics-grid",
            children=[
                card("Descuento", "result-discount", accent=True),
                card("Bonif. ERP / bulto", "result-cedido"),
                card("Final ERP / bulto", "result-bulto"),
                card("Final por unidad", "result-unit"),
                card("Precio sugerido", "result-suggested", accent=True),
                card("Gasto total bonificación", "result-total"),
            ],
        ),
        html.Div(
            className="calc-action-row",
            children=[
                html.Div(id="calc-note", className="calc-note"),
                html.Button("＋ Sumar SKU", id="add-sku-btn", n_clicks=0, className="primary-btn"),
            ],
        ),
        html.Div(id="list-feedback", className="list-feedback"),
        html.Div(
            className="panel expiry-panel",
            children=[
                html.Div(
                    className="section-title-row",
                    children=[
                        html.Div([html.Div("FRESCURA", className="eyebrow"), html.H3("Lotes y vencimientos del SKU actual")]),
                        dcc.Dropdown(
                            id="branch-filter",
                            options=[
                                {"label": "Todas las sucursales", "value": "Todas"},
                                {"label": "Trelew", "value": "Trelew"},
                                {"label": "Madryn", "value": "Madryn"},
                            ],
                            value="Todas",
                            clearable=False,
                            className="branch-dropdown",
                        ),
                    ],
                ),
                dash_table.DataTable(
                    id="expiry-table",
                    columns=[
                        {"name": "Sucursal", "id": "sucursal"},
                        {"name": "Vencimiento", "id": "vencimiento"},
                        {"name": "Estado", "id": "estado"},
                        {"name": "Stock lote (bultos)", "id": "stock"},
                        {"name": "Orden lote", "id": "orden"},
                        {"name": "Días p/ bloqueo", "id": "dias"},
                    ],
                    data=[],
                    page_size=8,
                    sort_action="native",
                    style_as_list_view=True,
                    style_header={"fontWeight": "700", "backgroundColor": "#111827", "color": "#cbd5e1", "border": "none"},
                    style_cell={
                        "backgroundColor": "#0b1220", "color": "#e5e7eb", "border": "none", "padding": "12px 10px",
                        "fontFamily": "Inter, Arial, sans-serif", "fontSize": "13px", "textAlign": "left",
                    },
                    style_data_conditional=[
                        {"if": {"filter_query": '{estado} = "ACCIONAR"', "column_id": "estado"}, "color": "#fb7185", "fontWeight": "700"},
                        {"if": {"filter_query": '{estado} = "OK"', "column_id": "estado"}, "color": "#4ade80", "fontWeight": "700"},
                        {"if": {"filter_query": '{estado} = "ELIMINAR"', "column_id": "estado"}, "color": "#94a3b8", "fontWeight": "700"},
                    ],
                ),
                html.Div(id="no-expiry", className="empty-note"),
            ],
        ),
        html.Div(
            className="panel simulation-panel",
            children=[
                html.Div(
                    className="simulation-header",
                    children=[
                        html.Div([
                            html.Div("SIMULACIÓN MULTI-SKU", className="eyebrow"),
                            html.H3("Listado para exportar"),
                            html.Div(id="simulation-count", className="simulation-count"),
                        ]),
                        html.Div(
                            className="simulation-actions",
                            children=[
                                html.Button("Vaciar listado", id="clear-list-btn", n_clicks=0, className="secondary-btn"),
                                html.Button("Exportar Excel", id="export-excel-btn", n_clicks=0, className="export-btn", disabled=True),
                            ],
                        ),
                    ],
                ),
                html.Div(id="month-alert", className="month-alert clear"),
                dash_table.DataTable(
                    id="simulation-table",
                    columns=[
                        {"name": "SKU", "id": "sku"},
                        {"name": "Descripción", "id": "descripcion"},
                        {"name": "Sucursal", "id": "sucursal"},
                        {"name": "Bultos", "id": "bultos"},
                        {"name": "Dto. %", "id": "discount"},
                        {"name": "Sugerido", "id": "suggested"},
                        {"name": "Final/u", "id": "final_unit"},
                        {"name": "Final/bulto", "id": "final_bulto"},
                        {"name": "Bonif./bulto", "id": "bonif_bulto"},
                        {"name": "Gasto total", "id": "gasto_total"},
                        {"name": "Stock total", "id": "stock_total"},
                        {"name": "Stock / vencimientos", "id": "vencimientos"},
                        {"name": "Alerta", "id": "alerta"},
                    ],
                    data=[],
                    page_size=12,
                    sort_action="native",
                    filter_action="native",
                    fixed_rows={"headers": True},
                    style_table={"overflowX": "auto", "minWidth": "100%"},
                    style_header={
                        "fontWeight": "800", "backgroundColor": "#111827", "color": "#cbd5e1", "border": "1px solid #22314a",
                        "whiteSpace": "normal", "height": "auto",
                    },
                    style_cell={
                        "backgroundColor": "#0b1220", "color": "#e5e7eb", "border": "1px solid #17243a",
                        "padding": "10px", "fontFamily": "Inter, Arial, sans-serif", "fontSize": "12px", "textAlign": "left",
                        "minWidth": "105px", "maxWidth": "300px", "whiteSpace": "normal", "height": "auto",
                    },
                    style_cell_conditional=[
                        {"if": {"column_id": "descripcion"}, "minWidth": "250px", "width": "250px"},
                        {"if": {"column_id": "vencimientos"}, "minWidth": "360px", "width": "360px", "maxWidth": "520px"},
                        {"if": {"column_id": "alerta"}, "minWidth": "145px"},
                    ],
                    style_data_conditional=[
                        {"if": {"filter_query": '{alerta} contains "VENCE ESTE MES"', "column_id": "alerta"}, "backgroundColor": "#3a2510", "color": "#fbbf24", "fontWeight": "800"},
                        {"if": {"filter_query": '{alerta} = "OK"', "column_id": "alerta"}, "color": "#4ade80", "fontWeight": "700"},
                    ],
                ),
                html.Div("El listado queda guardado durante la sesión del navegador. Volver a sumar el mismo SKU/sucursal actualiza ese escenario.", className="empty-note"),
            ],
        ),
        html.Div(
            className="footer-note",
            children=["Cálculo alineado al ERP: BONIF sobre Precio Base y FINAL sobre P. Final. Factor de sugerido: × 1,30."],
        ),
    ],
)


@app.callback(
    Output("data-version", "data"),
    Output("data-source-status", "children"),
    Output("data-loaded-at", "children"),
    Output("data-refresh-feedback", "children"),
    Input("refresh-data-btn", "n_clicks"),
    State("data-version", "data"),
    prevent_initial_call=True,
)
def refresh_data(n_clicks, version):
    global CATALOG, EXPIRY, DATA_STATUS
    if not n_clicks:
        raise PreventUpdate
    CATALOG, EXPIRY, DATA_STATUS = reload_sources()
    all_drive = all(DATA_STATUS.get(name, {}).get("ok") for name in ("Precios", "Frescura Trelew", "Frescura Madryn"))
    feedback = "Datos recargados desde Google Drive." if all_drive else "Actualización completada con al menos una fuente de respaldo local."
    return (version or 0) + 1, data_status_children(), f"Última carga: {DATA_STATUS.get('loaded_at', '—')} · Google Drive", feedback


@app.callback(
    Output("value-label", "children"),
    Output("value-input", "min"),
    Output("value-input", "max"),
    Output("value-input", "value"),
    Input("mode-input", "value"),
    State("value-input", "value"),
)
def switch_mode(mode, current_value):
    if mode == "suggested":
        return "Precio sugerido objetivo ($)", 0, None, 2000 if current_value is None or current_value <= 100 else current_value
    return "Descuento (%)", 0, 100, 20 if current_value is None or current_value > 100 else current_value


@app.callback(
    Output("status-message", "children"),
    Output("status-message", "className"),
    Output("product-sku", "children"),
    Output("product-description", "children"),
    Output("product-units", "children"),
    Output("current-base", "children"),
    Output("current-bulto", "children"),
    Output("current-suggested", "children"),
    Output("result-discount", "children"),
    Output("result-unit", "children"),
    Output("result-suggested", "children"),
    Output("result-bulto", "children"),
    Output("result-cedido", "children"),
    Output("result-total", "children"),
    Output("calc-note", "children"),
    Output("expiry-table", "data"),
    Output("no-expiry", "children"),
    Input("sku-input", "value"),
    Input("mode-input", "value"),
    Input("value-input", "value"),
    Input("bultos-input", "value"),
    Input("branch-filter", "value"),
    Input("data-version", "data"),
)
def calculate(sku_raw, mode, value, bultos, branch, _data_version):
    sku = str(sku_raw or "").strip()
    product = CATALOG.get(sku)
    blank = "—"

    if not sku:
        return (
            "Ingresá un SKU para comenzar.", "status-message neutral", "SKU —", "Sin producto seleccionado",
            blank, blank, blank, blank, blank, blank, blank, blank, blank, blank, "", [], ""
        )

    if not product:
        return (
            f"El SKU {sku} no está en el catálogo cargado.", "status-message error", f"SKU {sku}", "SKU no encontrado",
            blank, blank, blank, blank, blank, blank, blank, blank, blank, blank, "Revisá el código o actualizá la base.", [], ""
        )

    if product["p_final_bulto"] <= 0 or product["precio_base_bulto"] <= 0 or product["unidades_bulto"] <= 0:
        return (
            "El SKU existe, pero no tiene un precio válido para simular.", "status-message warning", f"SKU {sku}", product["descripcion"],
            number(product["unidades_bulto"], 0), blank, blank, blank, blank, blank, blank, blank, blank, blank,
            "No se puede calcular hasta tener Precio Base y P. Final mayores a cero.", [], ""
        )

    bultos_num = max(to_float(bultos, 1), 0)
    current_final = product.get("final_actual_bulto", product["p_final_bulto"] + product["precio_base_bulto"] * 0.03)
    current_unit = current_final / product["unidades_bulto"]
    current_suggested = current_unit * MARGEN_SUGERIDO

    if value is None:
        result = None
    elif mode == "suggested":
        result = calculate_from_suggested(product, max(float(value), 0), bultos_num)
    else:
        result = calculate_from_discount(product, min(max(float(value), 0), 100), bultos_num)

    lots_raw = lots_for_sku(sku, branch)
    lots = [{
        "sucursal": lot["sucursal"],
        "vencimiento": lot["vencimiento"],
        "estado": lot["estado"],
        "stock": number(lot["stock"], 2),
        "orden": number(lot["orden"], 0),
        "dias": number(lot["dias"], 0),
    } for lot in lots_raw]
    no_expiry = "No hay lotes de frescura informados para este SKU en la sucursal seleccionada." if not lots else ""

    if result is None:
        return (
            "Producto encontrado. Completá el valor para calcular.", "status-message neutral", f"SKU {sku}", product["descripcion"],
            number(product["unidades_bulto"], 0), money(product["precio_base_bulto"], 3), money(current_final, 3), money(current_suggested, 2),
            blank, blank, blank, blank, blank, blank, "", lots, no_expiry
        )

    if result["discount_pct"] < 0:
        status = "El sugerido objetivo está por encima del sugerido actual: requeriría un aumento, no un descuento."
        status_class = "status-message warning"
        note = f"Equivale a un aumento de {pct(abs(result['discount_pct']))} sobre el P. Final del ERP."
    elif result["discount_pct"] > 100:
        status = "El objetivo requiere más de 100% de descuento y no es válido."
        status_class = "status-message error"
        note = "Revisá el precio sugerido ingresado."
    else:
        status = "Cálculo listo."
        status_class = "status-message success"
        if mode == "suggested":
            note = f"Para llegar a un sugerido de {money(result['suggested'])}, el final neto por unidad debe quedar en {money(result['final_unit'])}."
        else:
            if abs(result["discount_pct"]) < 1e-12:
                note = f"Sin bonificación: final comercial {money(result['final_bulto'], 3)} y sugerido {money(result['suggested'])}."
            else:
                note = f"Aplicando {pct(result['discount_pct'])}, el sugerido resultante queda en {money(result['suggested'])}."

    return (
        status, status_class, f"SKU {sku}", product["descripcion"],
        number(product["unidades_bulto"], 0), money(product["precio_base_bulto"], 3), money(current_final, 3), money(current_suggested, 2),
        pct(result["discount_pct"]), money(result["final_unit"]), money(result["suggested"]), money(result["final_bulto"], 3),
        money(result["bonif_bulto"], 3), money(result["gasto_total"], 2), note, lots, no_expiry
    )


@app.callback(
    Output("simulation-store", "data"),
    Output("list-feedback", "children"),
    Output("list-feedback", "className"),
    Input("add-sku-btn", "n_clicks"),
    Input("clear-list-btn", "n_clicks"),
    Input("refresh-data-btn", "n_clicks"),
    State("simulation-store", "data"),
    State("sku-input", "value"),
    State("mode-input", "value"),
    State("value-input", "value"),
    State("bultos-input", "value"),
    State("branch-filter", "value"),
    prevent_initial_call=True,
)
def manage_list(add_clicks, clear_clicks, refresh_clicks, current_items, sku, mode, value, bultos, branch):
    trigger = ctx.triggered_id
    items = list(current_items or [])

    if trigger == "clear-list-btn":
        return [], "Listado vaciado.", "list-feedback neutral"

    if trigger == "refresh-data-btn":
        return [], "Datos actualizados. El listado se vació para no mezclar precios de distintas cargas.", "list-feedback neutral"

    if trigger != "add-sku-btn":
        raise PreventUpdate

    entry, error = build_simulation_entry(sku, mode, value, bultos, branch)
    if error:
        return no_update, error, "list-feedback error"

    replaced = False
    for i, existing in enumerate(items):
        if existing.get("key") == entry["key"]:
            items[i] = entry
            replaced = True
            break
    if not replaced:
        items.append(entry)

    message = f"SKU {entry['sku']} actualizado en el listado." if replaced else f"SKU {entry['sku']} sumado al listado."
    return items, message, "list-feedback success"


@app.callback(
    Output("simulation-table", "data"),
    Output("simulation-count", "children"),
    Output("month-alert", "children"),
    Output("month-alert", "className"),
    Output("export-excel-btn", "disabled"),
    Input("simulation-store", "data"),
)
def render_list(items):
    items = items or []
    rows = []
    for item in items:
        rows.append({
            "sku": item["sku"],
            "descripcion": item["descripcion"],
            "sucursal": item["sucursal"],
            "bultos": number(item["bultos"], 2),
            "discount": pct(item["discount_pct"]),
            "suggested": money(item["suggested"], 2),
            "final_unit": money(item["final_unit"], 2),
            "final_bulto": money(item["final_bulto"], 3),
            "bonif_bulto": money(item["bonif_bulto"], 3),
            "gasto_total": money(item["gasto_total"], 2),
            "stock_total": number(item["stock_total"], 2),
            "vencimientos": item["vencimientos_resumen"],
            "alerta": item["alerta_mes"],
        })

    count_text = f"{len(items)} SKU en el listado" if len(items) != 1 else "1 SKU en el listado"
    alerted = [item for item in items if str(item.get("alerta_mes", "")).startswith("⚠")]
    current = now_ar()
    month_label = f"{MONTHS_ES[current.month - 1]} {current.year}"

    if alerted:
        details = "; ".join(f"{item['sku']} — {item.get('vencimientos_mes', '')}" for item in alerted)
        alert_text = f"⚠ {len(alerted)} SKU con vencimientos en {month_label}: {details}"
        alert_class = "month-alert warning"
    else:
        alert_text = f"Sin SKU del listado con vencimientos en {month_label}."
        alert_class = "month-alert clear"

    return rows, count_text, alert_text, alert_class, not bool(items)


@app.callback(
    Output("excel-download", "data"),
    Input("export-excel-btn", "n_clicks"),
    State("simulation-store", "data"),
    prevent_initial_call=True,
)
def export_excel(n_clicks, items):
    if not n_clicks or not items:
        raise PreventUpdate

    timestamp = now_ar().strftime("%Y%m%d_%H%M")

    def writer(buffer):
        write_simulation_excel(buffer, items)

    return dcc.send_bytes(writer, f"simulacion_skus_{timestamp}.xlsx")


if __name__ == "__main__":
    app.run(debug=False, host="127.0.0.1", port=8051)
