from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
import os
import re
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

PRICE_FILE_ID = os.getenv("DRIVE_PRICE_FILE_ID", "11bArEuSP1JLoHnQRgJSnt6kRY7xeCuv5")
TRELEW_FILE_ID = os.getenv("DRIVE_TRELEW_FILE_ID", "1w5ydKY-p8pkOphqeIRGXiTdrfHnVYFnM")
MADRYN_FILE_ID = os.getenv("DRIVE_MADRYN_FILE_ID", "18rSCdJrYcQ_8nyWHULXd12oFO8HqkC7p")

AR_TZ = timezone(timedelta(hours=-3), name="ART")

_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"main": _MAIN_NS, "rel": _REL_NS, "pkg": _PKG_NS}


def _to_float(value, default=0.0):
    try:
        if value in (None, ""):
            return default
        return float(str(value).replace(",", "."))
    except (TypeError, ValueError):
        return default


def _normalize_sku(value):
    if value in (None, ""):
        return ""
    text = str(value).strip()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def download_public_drive_file(file_id: str, timeout: int = 30) -> bytes:
    """Download a small public Google Drive file by ID."""
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (compatible; SimuladorPrecios/1.0)",
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
        },
    )
    with urllib.request.urlopen(request, timeout=timeout) as response:
        payload = response.read()
    if not payload.startswith(b"PK"):
        preview = payload[:180].decode("utf-8", errors="ignore")
        raise RuntimeError(f"Google Drive no devolvió un XLSX válido. Respuesta: {preview!r}")
    return payload


def _shared_strings(zf: zipfile.ZipFile):
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    out = []
    for si in root.findall("main:si", NS):
        out.append("".join(node.text or "" for node in si.iter(f"{{{_MAIN_NS}}}t")))
    return out


def _first_sheet_path(zf: zipfile.ZipFile) -> str:
    workbook = ET.fromstring(zf.read("xl/workbook.xml"))
    rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
    sheet = workbook.find("main:sheets", NS)[0]
    rid = sheet.attrib[f"{{{_REL_NS}}}id"]
    target = rel_map[rid].lstrip("/")
    if not target.startswith("xl/"):
        target = "xl/" + target
    return target


def _cell_value(cell, shared):
    cell_type = cell.attrib.get("t")
    value = cell.find("main:v", NS)
    inline = cell.find("main:is", NS)
    if cell_type == "s" and value is not None:
        idx = int(value.text)
        return shared[idx] if 0 <= idx < len(shared) else ""
    if cell_type == "inlineStr" and inline is not None:
        return "".join(node.text or "" for node in inline.iter(f"{{{_MAIN_NS}}}t"))
    if cell_type == "b" and value is not None:
        return value.text == "1"
    if value is not None:
        return value.text
    return ""


def read_first_sheet_rows(xlsx_bytes: bytes):
    """Return rows as dicts keyed by Excel column letters (A, B, ... AB)."""
    with zipfile.ZipFile(BytesIO(xlsx_bytes)) as zf:
        shared = _shared_strings(zf)
        sheet_path = _first_sheet_path(zf)
        root = ET.fromstring(zf.read(sheet_path))
        sheet_data = root.find("main:sheetData", NS)
        rows = []
        if sheet_data is None:
            return rows
        for row in sheet_data.findall("main:row", NS):
            values = {}
            for cell in row.findall("main:c", NS):
                ref = cell.attrib.get("r", "")
                match = re.match(r"([A-Z]+)", ref)
                if not match:
                    continue
                values[match.group(1)] = _cell_value(cell, shared)
            rows.append(values)
        return rows


def parse_price_catalog(xlsx_bytes: bytes):
    rows = read_first_sheet_rows(xlsx_bytes)
    catalog = {}
    for row in rows[1:]:
        sku = _normalize_sku(row.get("A"))
        if not sku:
            continue
        description = str(row.get("B") or "").strip()
        units = _to_float(row.get("C"))
        bultos_pallet = _to_float(row.get("D"))
        lista = _to_float(row.get("E"))
        precio_base = _to_float(row.get("F"))
        p_final = _to_float(row.get("G"))
        catalog[sku] = {
            "sku": sku,
            "descripcion": description,
            "unidades_bulto": units,
            "bultos_pallet": bultos_pallet,
            "lista": lista,
            "precio_base_bulto": precio_base,
            "p_final_bulto": p_final,
            # Final comercial sin bonificación: el ERP/exportación suma un 3% del Precio Base
            # sobre el P. Final cuando no existe descuento.
            "final_actual_bulto": p_final + (precio_base * 0.03),
            "final_actual_unidad": ((p_final + (precio_base * 0.03)) / units) if units else 0,
            "evidencia": "Google Drive · plantillaPreciosColumna",
        }
    if not catalog:
        raise RuntimeError("La lista de precios de Drive no contiene SKUs legibles.")
    return catalog


def _excel_date_to_iso(value):
    if value in (None, ""):
        return ""
    text = str(value).strip()
    try:
        serial = float(text)
        if serial > 0:
            date_value = datetime(1899, 12, 30) + timedelta(days=serial)
            return date_value.strftime("%Y-%m-%d")
    except ValueError:
        pass
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    return ""


def parse_freshness(xlsx_bytes: bytes, branch: str, catalog=None):
    rows = read_first_sheet_rows(xlsx_bytes)
    output = []
    groups = [
        ("K", "L", "M", "P", 1),
        ("Q", "R", "S", "V", 2),
        ("W", "X", "Y", "AB", 3),
    ]
    for row in rows[2:]:
        sku = _normalize_sku(row.get("A"))
        if not sku:
            continue
        description = str(row.get("B") or "").strip()
        units = 0
        if catalog and sku in catalog:
            units = catalog[sku].get("unidades_bulto", 0)
        for stock_col, date_col, days_col, status_col, order in groups:
            stock = _to_float(row.get(stock_col))
            expiry = _excel_date_to_iso(row.get(date_col))
            if stock <= 0 or not expiry:
                continue
            output.append({
                "sucursal": branch,
                "sku": sku,
                "descripcion": description,
                "unidades_bulto": units,
                "vencimiento": expiry,
                "estado": str(row.get(status_col) or "").strip(),
                "stock_lote_bultos": stock,
                "orden_lote": order,
                "dias_bloqueo": _to_float(row.get(days_col)),
            })
    return output


def load_catalog_from_drive(file_id: str | None = None):
    return parse_price_catalog(download_public_drive_file(file_id or PRICE_FILE_ID))


def load_freshness_from_drive(file_id: str, branch: str, catalog=None):
    return parse_freshness(download_public_drive_file(file_id), branch, catalog)


def load_all_from_drive():
    catalog = load_catalog_from_drive(PRICE_FILE_ID)
    trelew = load_freshness_from_drive(TRELEW_FILE_ID, "Trelew", catalog)
    madryn = load_freshness_from_drive(MADRYN_FILE_ID, "Madryn", catalog)
    return catalog, trelew + madryn, {
        "loaded_at": datetime.now(AR_TZ).strftime("%d/%m/%Y %H:%M"),
        "prices_count": len(catalog),
        "trelew_count": len(trelew),
        "madryn_count": len(madryn),
    }
