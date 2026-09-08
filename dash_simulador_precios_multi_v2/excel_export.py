from __future__ import annotations

from datetime import datetime
from typing import BinaryIO, Iterable

import xlsxwriter


MARGEN_SUGERIDO = 1.30


def _parse_iso_date(value):
    if not value:
        return None
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except (TypeError, ValueError):
        return None


def write_simulation_excel(buffer: BinaryIO, items: Iterable[dict]) -> None:
    """Escribe un Excel de simulación con resumen por SKU y detalle de vencimientos."""
    items = list(items or [])
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    workbook.set_properties({
        "title": "Simulación comercial por SKU",
        "subject": "Descuentos, sugeridos, stock y vencimientos",
        "company": "Distribuidora del Valle",
        "comments": "Generado desde el simulador Dash.",
    })

    # Paleta alineada con el dashboard.
    navy = "#0D1728"
    navy2 = "#111D30"
    cyan = "#22D3EE"
    line = "#22314A"
    text = "#F8FAFC"
    muted = "#CBD5E1"
    amber = "#FBBF24"
    red = "#FB7185"
    green = "#4ADE80"
    white = "#FFFFFF"

    title_fmt = workbook.add_format({
        "bold": True, "font_size": 18, "font_color": white, "bg_color": navy,
        "align": "left", "valign": "vcenter",
    })
    subtitle_fmt = workbook.add_format({
        "font_size": 10, "font_color": muted, "bg_color": navy,
        "align": "left", "valign": "vcenter",
    })
    header_fmt = workbook.add_format({
        "bold": True, "font_color": white, "bg_color": "#0E7490",
        "border": 1, "border_color": line, "align": "center", "valign": "vcenter",
        "text_wrap": True,
    })
    text_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "valign": "top"})
    int_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "0", "valign": "top"})
    num2_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "#,##0.00", "valign": "top"})
    num3_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "#,##0.000", "valign": "top"})
    money2_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "$ #,##0.00", "valign": "top"})
    money3_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "$ #,##0.000", "valign": "top"})
    pct_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "0.00%", "valign": "top"})
    date_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "num_format": "dd/mm/yyyy", "valign": "top"})
    wrap_fmt = workbook.add_format({"border": 1, "border_color": "#D9E2EC", "text_wrap": True, "valign": "top"})
    alert_fmt = workbook.add_format({
        "border": 1, "border_color": "#D9E2EC", "bg_color": "#FEF3C7", "font_color": "#92400E",
        "bold": True, "valign": "top"
    })
    ok_fmt = workbook.add_format({
        "border": 1, "border_color": "#D9E2EC", "bg_color": "#DCFCE7", "font_color": "#166534",
        "bold": True, "valign": "top"
    })
    month_detail_fmt = workbook.add_format({
        "border": 1, "border_color": "#D9E2EC", "bg_color": "#FEE2E2", "font_color": "#991B1B",
        "bold": True, "valign": "top"
    })

    # --- Hoja 1: Simulación ---
    ws = workbook.add_worksheet("Simulacion")
    ws.hide_gridlines(2)
    ws.freeze_panes(4, 0)
    ws.set_tab_color(cyan)
    ws.merge_range("A1:T1", "SIMULACIÓN COMERCIAL POR SKU", title_fmt)
    ws.merge_range(
        "A2:T2",
        "Resumen de descuentos, precio sugerido, stock y vencimientos. Las columnas calculadas conservan fórmulas editables.",
        subtitle_fmt,
    )
    ws.set_row(0, 28)
    ws.set_row(1, 22)

    headers = [
        "SKU", "Descripción", "Sucursal", "Bultos a bonificar", "Unidades/bulto",
        "Precio base/bulto", "P. Final ERP/bulto", "Descuento %", "Bonif. ERP/bulto",
        "Final ERP/bulto", "Percepción 3%/bulto", "Final cliente/bulto", "Final cliente/unidad",
        "Precio sugerido", "Gasto total bonificación", "Stock total (bultos)",
        "Próximo vencimiento", "Stock próx. vencimiento", "Stock / vencimientos", "Alerta mes corriente",
    ]
    header_row = 3  # zero-based => Excel row 4
    for c, value in enumerate(headers):
        ws.write(header_row, c, value, header_fmt)
    ws.set_row(header_row, 36)

    for idx, item in enumerate(items, start=4):
        # idx es zero-based de xlsxwriter; fila Excel = idx + 1.
        excel_row = idx + 1
        ws.write(idx, 0, str(item.get("sku", "")), text_fmt)
        ws.write(idx, 1, item.get("descripcion", ""), text_fmt)
        ws.write(idx, 2, item.get("sucursal", "Todas"), text_fmt)
        ws.write_number(idx, 3, float(item.get("bultos", 0) or 0), num2_fmt)
        ws.write_number(idx, 4, float(item.get("unidades_bulto", 0) or 0), int_fmt)
        ws.write_number(idx, 5, float(item.get("precio_base_bulto", 0) or 0), money3_fmt)
        ws.write_number(idx, 6, float(item.get("p_final_bulto", 0) or 0), money3_fmt)
        discount_decimal = float(item.get("discount_pct", 0) or 0) / 100.0
        ws.write_number(idx, 7, discount_decimal, pct_fmt)

        # Fórmulas: ERP por un lado y precio cliente con percepción 3% por otro.
        ws.write_formula(idx, 8, f"=F{excel_row}*H{excel_row}", money3_fmt, float(item.get("bonif_bulto", 0) or 0))
        ws.write_formula(idx, 9, f"=G{excel_row}*(1-H{excel_row})", money3_fmt, float(item.get("final_erp_bulto", 0) or 0))
        ws.write_formula(idx, 10, f"=F{excel_row}*3%*(1-H{excel_row})", money3_fmt, float(item.get("percepcion_bulto", 0) or 0))
        ws.write_formula(idx, 11, f"=J{excel_row}+K{excel_row}", money3_fmt, float(item.get("final_bulto", 0) or 0))
        ws.write_formula(idx, 12, f"=IFERROR(L{excel_row}/E{excel_row},0)", money2_fmt, float(item.get("final_unit", 0) or 0))
        ws.write_formula(idx, 13, f"=M{excel_row}*{MARGEN_SUGERIDO}", money2_fmt, float(item.get("suggested", 0) or 0))
        ws.write_formula(idx, 14, f"=I{excel_row}*D{excel_row}", money2_fmt, float(item.get("gasto_total", 0) or 0))
        ws.write_number(idx, 15, float(item.get("stock_total", 0) or 0), num2_fmt)

        next_date = _parse_iso_date(item.get("proximo_vencimiento_iso"))
        if next_date:
            ws.write_datetime(idx, 16, next_date, date_fmt)
        else:
            ws.write_blank(idx, 16, None, date_fmt)
        ws.write_number(idx, 17, float(item.get("stock_proximo", 0) or 0), num2_fmt)
        ws.write(idx, 18, item.get("vencimientos_resumen", ""), wrap_fmt)
        alert = item.get("alerta_mes", "OK")
        ws.write(idx, 19, alert, alert_fmt if alert.startswith("⚠") else ok_fmt)

    data_end_row = 4 + len(items)  # fila Excel final (header=4)
    if items:
        ws.autofilter(3, 0, data_end_row - 1, len(headers) - 1)
        ws.conditional_format(4, 19, data_end_row - 1, 19, {
            "type": "text", "criteria": "containing", "value": "VENCE ESTE MES",
            "format": alert_fmt,
        })

    widths = [11, 38, 14, 18, 15, 20, 20, 14, 20, 20, 20, 20, 20, 18, 23, 20, 19, 22, 56, 22]
    for c, width in enumerate(widths):
        ws.set_column(c, c, width)

    # --- Hoja 2: Vencimientos ---
    lot_ws = workbook.add_worksheet("Vencimientos")
    lot_ws.hide_gridlines(2)
    lot_ws.freeze_panes(4, 0)
    lot_ws.set_tab_color(amber)
    lot_ws.merge_range("A1:J1", "DETALLE DE STOCK Y VENCIMIENTOS", title_fmt)
    lot_ws.merge_range(
        "A2:J2",
        "Una fila por lote de cada SKU agregado. La alerta identifica lotes con vencimiento dentro del mes corriente al momento de la simulación.",
        subtitle_fmt,
    )
    lot_headers = [
        "SKU", "Descripción", "Sucursal simulada", "Sucursal lote", "Vencimiento",
        "Estado", "Stock lote (bultos)", "Orden lote", "Días p/ bloqueo", "Mes corriente",
    ]
    for c, value in enumerate(lot_headers):
        lot_ws.write(3, c, value, header_fmt)
    lot_ws.set_row(3, 34)

    row_idx = 4
    for item in items:
        for lot in item.get("lots", []) or []:
            lot_ws.write(row_idx, 0, str(item.get("sku", "")), text_fmt)
            lot_ws.write(row_idx, 1, item.get("descripcion", ""), text_fmt)
            lot_ws.write(row_idx, 2, item.get("sucursal", "Todas"), text_fmt)
            lot_ws.write(row_idx, 3, lot.get("sucursal", ""), text_fmt)
            dt = _parse_iso_date(lot.get("vencimiento_iso"))
            if dt:
                lot_ws.write_datetime(row_idx, 4, dt, date_fmt)
            else:
                lot_ws.write(row_idx, 4, lot.get("vencimiento", ""), text_fmt)
            lot_ws.write(row_idx, 5, lot.get("estado", ""), text_fmt)
            lot_ws.write_number(row_idx, 6, float(lot.get("stock", 0) or 0), num2_fmt)
            lot_ws.write_number(row_idx, 7, float(lot.get("orden", 0) or 0), int_fmt)
            lot_ws.write_number(row_idx, 8, float(lot.get("dias", 0) or 0), int_fmt)
            is_current = bool(lot.get("mes_corriente"))
            lot_ws.write(row_idx, 9, "⚠ SÍ" if is_current else "NO", month_detail_fmt if is_current else text_fmt)
            row_idx += 1

    if row_idx > 4:
        lot_ws.autofilter(3, 0, row_idx - 1, len(lot_headers) - 1)
    lot_widths = [11, 38, 18, 18, 16, 14, 21, 13, 17, 16]
    for c, width in enumerate(lot_widths):
        lot_ws.set_column(c, c, width)

    # Mensaje útil cuando ningún SKU tiene vencimientos.
    if not items:
        ws.write("A5", "No hay SKU agregados a la simulación.", text_fmt)
    elif row_idx == 4:
        lot_ws.write("A5", "Los SKU agregados no tienen lotes de frescura informados.", text_fmt)

    workbook.close()
