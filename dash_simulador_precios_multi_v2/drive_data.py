from __future__ import annotations

from datetime import datetime, timedelta, timezone
from io import BytesIO
import html as html_lib
import json
import os
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile

# Carpeta operativa. El Dash intenta descubrir automáticamente los archivos
# cada vez que inicia o se pulsa "Actualizar datos".
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID", "1cukgXLUaPsEDK_yD7tSwgaBFZAbiDUot")

# IDs de respaldo. SOLO se usan si falla la detección automática.
PRICE_FILE_ID = os.getenv("DRIVE_PRICE_FILE_ID", "1gqrFQ_b6e4ZdopQHANdr2tgVqefm8Puv")
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


def _http_bytes(url: str, timeout: int = 30, headers=None) -> bytes:
    request_headers = {
        "User-Agent": "Mozilla/5.0 (compatible; SimuladorPrecios/2.0)",
        "Accept": "*/*",
    }
    if headers:
        request_headers.update(headers)
    request = urllib.request.Request(url, headers=request_headers)
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return response.read()


def download_public_drive_file(file_id: str, timeout: int = 30) -> bytes:
    """Descarga un XLSX público de Google Drive por ID."""
    url = f"https://drive.google.com/uc?export=download&id={file_id}"
    payload = _http_bytes(
        url,
        timeout=timeout,
        headers={
            "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*",
        },
    )
    if not payload.startswith(b"PK"):
        preview = payload[:180].decode("utf-8", errors="ignore")
        raise RuntimeError(f"Google Drive no devolvió un XLSX válido. Respuesta: {preview!r}")
    return payload


# ---------------------------------------------------------------------------
# Descubrimiento automático de archivos
# ---------------------------------------------------------------------------

def _list_folder_with_api_key(folder_id: str):
    api_key = os.getenv("GOOGLE_DRIVE_API_KEY", "").strip()
    if not api_key:
        return []
    query = f"'{folder_id}' in parents and trashed = false"
    params = urllib.parse.urlencode({
        "q": query,
        "fields": "files(id,name,mimeType,modifiedTime)",
        "orderBy": "modifiedTime desc",
        "pageSize": 100,
        "key": api_key,
    })
    payload = _http_bytes(f"https://www.googleapis.com/drive/v3/files?{params}")
    data = json.loads(payload.decode("utf-8"))
    return data.get("files", [])


def _service_account_info():
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON no contiene JSON válido") from exc


def _list_folder_with_service_account(folder_id: str):
    info = _service_account_info()
    if not info:
        return []
    try:
        from google.oauth2 import service_account
        from google.auth.transport.requests import AuthorizedSession
    except ImportError as exc:
        raise RuntimeError("Faltan google-auth/requests para usar la cuenta de servicio") from exc

    credentials = service_account.Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive.readonly"],
    )
    session = AuthorizedSession(credentials)
    response = session.get(
        "https://www.googleapis.com/drive/v3/files",
        params={
            "q": f"'{folder_id}' in parents and trashed = false",
            "fields": "files(id,name,mimeType,modifiedTime)",
            "orderBy": "modifiedTime desc",
            "pageSize": 100,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json().get("files", [])


def _decode_js_text(value: str) -> str:
    value = html_lib.unescape(value)
    try:
        return bytes(value, "utf-8").decode("unicode_escape")
    except UnicodeDecodeError:
        return value


def _list_public_folder_html(folder_id: str):
    """
    Último recurso sin credenciales: inspecciona el HTML de la carpeta pública.
    Google cambia esta estructura ocasionalmente, por eso API key / service account
    tienen prioridad cuando están configuradas.
    """
    urls = [
        f"https://drive.google.com/drive/folders/{folder_id}?usp=sharing",
        f"https://drive.google.com/drive/mobile/folders/{folder_id}",
        f"https://drive.google.com/embeddedfolderview?id={folder_id}#list",
    ]
    found = {}
    wanted = ("plantillaprecioscolumna", "plantillafrescura trelew", "plantillafrescura madryn")

    for url in urls:
        try:
            text = _http_bytes(url, timeout=20).decode("utf-8", errors="ignore")
        except Exception:
            continue

        # Pares ID/nombre en ambos órdenes. Filtramos estrictamente por nuestros nombres.
        patterns = [
            r'["\']([A-Za-z0-9_-]{20,})["\']\s*,\s*["\']([^"\']+\.xlsx)["\']',
            r'["\']([^"\']+\.xlsx)["\']\s*,\s*["\']([A-Za-z0-9_-]{20,})["\']',
        ]
        for p_index, pattern in enumerate(patterns):
            for match in re.finditer(pattern, text, flags=re.IGNORECASE):
                if p_index == 0:
                    file_id, name = match.group(1), match.group(2)
                else:
                    name, file_id = match.group(1), match.group(2)
                name = _decode_js_text(name).replace("\\/", "/")
                lower = name.lower()
                if any(token in lower for token in wanted):
                    found[file_id] = {
                        "id": file_id,
                        "name": name,
                        "modifiedTime": "",
                        "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    }
        if found:
            break
    return list(found.values())


def _price_sort_key(item):
    name = str(item.get("name") or "")
    # La exportación diaria comienza con YYYYMMDDHHMMSS. Esto permite ordenar
    # aun cuando el método público no exponga modifiedTime.
    match = re.match(r"(20\d{12})", name)
    timestamp = match.group(1) if match else ""
    return (timestamp, str(item.get("modifiedTime") or ""), name.lower())


def _pick_latest(files, name_contains: str, exact: bool = False):
    needle = name_contains.lower()
    candidates = []
    for item in files:
        name = str(item.get("name") or "")
        lower = name.lower()
        if (exact and lower == needle) or (not exact and needle in lower):
            candidates.append(item)
    if not candidates:
        return None
    if "plantillaprecioscolumna" in needle:
        return max(candidates, key=_price_sort_key)
    return max(candidates, key=lambda x: (str(x.get("modifiedTime") or ""), str(x.get("name") or "")))


def resolve_drive_sources(folder_id: str | None = None):
    """
    Resuelve los 3 archivos operativos actuales.

    Prioridad de descubrimiento:
      1. Cuenta de servicio (GOOGLE_SERVICE_ACCOUNT_JSON)
      2. API key (GOOGLE_DRIVE_API_KEY)
      3. Carpeta pública (sin credenciales)
      4. IDs de respaldo

    La lista de precios se elige SIEMPRE por el archivo más nuevo cuyo nombre
    contiene 'plantillaPreciosColumna'. Por eso un nuevo archivo diario puede
    tener un ID diferente sin modificar Render ni GitHub.
    """
    folder_id = folder_id or DRIVE_FOLDER_ID
    files = []
    method = ""
    errors = []

    attempts = [
        ("service-account", _list_folder_with_service_account),
        ("api-key", _list_folder_with_api_key),
        ("carpeta-pública", _list_public_folder_html),
    ]
    for label, loader in attempts:
        try:
            candidate_files = loader(folder_id)
            if candidate_files:
                files = candidate_files
                method = label
                break
        except Exception as exc:
            errors.append(f"{label}: {type(exc).__name__}")

    price = _pick_latest(files, "plantillaPreciosColumna") if files else None
    trelew = _pick_latest(files, "plantillafrescura trelew") if files else None
    madryn = _pick_latest(files, "plantillafrescura madryn") if files else None

    used_fallback = False
    if not price:
        used_fallback = True
        price = {"id": PRICE_FILE_ID, "name": "lista precios · ID respaldo", "modifiedTime": ""}
    if not trelew:
        used_fallback = True
        trelew = {"id": TRELEW_FILE_ID, "name": "plantillafrescura trelew.xlsx · ID respaldo", "modifiedTime": ""}
    if not madryn:
        used_fallback = True
        madryn = {"id": MADRYN_FILE_ID, "name": "plantillafrescura madryn.xlsx · ID respaldo", "modifiedTime": ""}

    if not method:
        method = "IDs-respaldo"
    elif used_fallback:
        method += "+respaldo"

    return {
        "price": price,
        "trelew": trelew,
        "madryn": madryn,
        "method": method,
        "errors": errors,
        "folder_id": folder_id,
    }


# ---------------------------------------------------------------------------
# Lectura XLSX
# ---------------------------------------------------------------------------

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
        final_actual = p_final + (precio_base * 0.03)
        catalog[sku] = {
            "sku": sku,
            "descripcion": description,
            "unidades_bulto": units,
            "bultos_pallet": bultos_pallet,
            "lista": lista,
            "precio_base_bulto": precio_base,
            "p_final_bulto": p_final,
            "final_actual_bulto": final_actual,
            "final_actual_unidad": (final_actual / units) if units else 0,
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


def load_all_from_drive_auto():
    sources = resolve_drive_sources()
    catalog = load_catalog_from_drive(sources["price"]["id"])
    trelew = load_freshness_from_drive(sources["trelew"]["id"], "Trelew", catalog)
    madryn = load_freshness_from_drive(sources["madryn"]["id"], "Madryn", catalog)
    return catalog, trelew + madryn, sources
