import json
import os
import shutil
from http import HTTPStatus
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile


HOST = "127.0.0.1"
PORT = 8001
WORKSPACE = Path(__file__).resolve().parent
WORKBOOK_NAME = "customer_orders_template_updated.xlsx"
WORKBOOK_PATH = WORKSPACE / WORKBOOK_NAME
SHEET_PATH = "xl/worksheets/sheet1.xml"
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", NS)

HEADERS = [
    "Name",
    "Email",
    "Class",
    "School",
    "Phone Number",
    "Candy Item",
    "Quantity",
    "Notes",
]


def column_name(number):
    name = ""
    current = number
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        name = chr(65 + remainder) + name
    return name


def sheet_tag(name):
    return f"{{{NS}}}{name}"


def build_template(path):
    widths = [22, 28, 18, 28, 20, 22, 14, 28]
    row_cells = []
    for index, header in enumerate(HEADERS, start=1):
        ref = f"{column_name(index)}1"
        row_cells.append(
            f'<c r="{ref}" t="inlineStr"><is><t>{escape_xml(header)}</t></is></c>'
        )

    cols = "".join(
        f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths, start=1)
    )

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS}">'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f"<cols>{cols}</cols>"
        f'<sheetData><row r="1">{"".join(row_cells)}</row></sheetData>'
        "</worksheet>"
    )

    files = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>"
        ),
        "docProps/core.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            "<dc:title>A VAN CANDIES Orders</dc:title>"
            "<dc:creator>A VAN CANDIES</dc:creator>"
            "<cp:lastModifiedBy>A VAN CANDIES</cp:lastModifiedBy>"
            '<dcterms:created xsi:type="dcterms:W3CDTF">2026-04-25T00:00:00Z</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-25T00:00:00Z</dcterms:modified>'
            "</cp:coreProperties>"
        ),
        "docProps/app.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Microsoft Excel</Application>"
            "</Properties>"
        ),
        "xl/workbook.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Orders" sheetId="1" r:id="rId1"/></sheets>'
            "</workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>"
        ),
        "xl/styles.xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<styleSheet xmlns="{NS}">'
            '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
            '<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            "</styleSheet>"
        ),
        SHEET_PATH: sheet_xml,
    }

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as workbook:
        for filename, content in files.items():
            workbook.writestr(filename, content)


def escape_xml(value):
    return (
        value.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def append_order_to_workbook(order):
    if not WORKBOOK_PATH.exists():
        build_template(WORKBOOK_PATH)

    with ZipFile(WORKBOOK_PATH, "r") as workbook:
        files = {name: workbook.read(name) for name in workbook.namelist()}

    root = ET.fromstring(files[SHEET_PATH])
    sheet_data = root.find(sheet_tag("sheetData"))
    if sheet_data is None:
        raise ValueError("Worksheet is missing sheet data.")

    existing_rows = sheet_data.findall(sheet_tag("row"))
    next_row_number = len(existing_rows) + 1
    row_element = ET.SubElement(sheet_data, sheet_tag("row"), {"r": str(next_row_number)})

    values = [
        order["name"],
        order["email"],
        order["class"],
        order["school"],
        order["phone"],
        order["item"],
        order["quantity"],
        order["notes"],
    ]

    for index, value in enumerate(values, start=1):
        reference = f"{column_name(index)}{next_row_number}"
        cell = ET.SubElement(row_element, sheet_tag("c"), {"r": reference, "t": "inlineStr"})
        is_element = ET.SubElement(cell, sheet_tag("is"))
        text_element = ET.SubElement(is_element, sheet_tag("t"))
        text_element.text = str(value)

    files[SHEET_PATH] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with NamedTemporaryFile(delete=False, suffix=".xlsx", dir=WORKSPACE) as temp_file:
        temp_path = Path(temp_file.name)

    try:
        with ZipFile(temp_path, "w", compression=ZIP_DEFLATED) as workbook:
            for filename, content in files.items():
                workbook.writestr(filename, content)
        os.replace(temp_path, WORKBOOK_PATH)
    finally:
        if temp_path.exists():
            temp_path.unlink(missing_ok=True)


class OrderHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(WORKSPACE), **kwargs)

    def do_GET(self):
        requested = self.path.split("?", 1)[0]
        blocked_files = {
            f"/{WORKBOOK_NAME}",
            "/customer_orders_template.xlsx",
            "/customer_orders_template_updated.xlsx",
            "/customer_orders_template_with_notes.xlsx",
        }

        if requested in blocked_files or requested.endswith(".xlsx"):
            self.send_error(HTTPStatus.FORBIDDEN, "This file is private.")
            return

        super().do_GET()

    def end_headers(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(HTTPStatus.NO_CONTENT)
        self.end_headers()

    def do_POST(self):
        if self.path != "/submit-order":
            self.send_error(HTTPStatus.NOT_FOUND, "Not found.")
            return

        content_length = int(self.headers.get("Content-Length", "0"))
        raw_body = self.rfile.read(content_length)

        try:
            payload = json.loads(raw_body.decode("utf-8"))
        except json.JSONDecodeError:
            self.respond_json(HTTPStatus.BAD_REQUEST, {"error": "Invalid order data."})
            return

        required_fields = ["name", "email", "class", "school", "phone", "item", "quantity"]
        missing = [field for field in required_fields if not str(payload.get(field, "")).strip()]
        if missing:
            self.respond_json(HTTPStatus.BAD_REQUEST, {"error": "Please complete all required fields."})
            return

        order = {
            "name": str(payload.get("name", "")).strip(),
            "email": str(payload.get("email", "")).strip(),
            "class": str(payload.get("class", "")).strip(),
            "school": str(payload.get("school", "")).strip(),
            "phone": str(payload.get("phone", "")).strip(),
            "item": str(payload.get("item", "")).strip(),
            "quantity": str(payload.get("quantity", "")).strip(),
            "notes": str(payload.get("notes", "")).strip(),
        }

        if WORKBOOK_PATH.exists():
            try:
                with open(WORKBOOK_PATH, "rb"):
                    pass
            except PermissionError:
                self.respond_json(
                    HTTPStatus.CONFLICT,
                    {"error": "Close the Excel file first, then submit the order again."},
                )
                return

        try:
            append_order_to_workbook(order)
        except PermissionError:
            self.respond_json(
                HTTPStatus.CONFLICT,
                {"error": "Close the Excel file first, then submit the order again."},
            )
            return
        except Exception as error:
            self.respond_json(HTTPStatus.INTERNAL_SERVER_ERROR, {"error": f"Could not save order: {error}"})
            return

        self.respond_json(HTTPStatus.OK, {"ok": True, "file": WORKBOOK_NAME})

    def respond_json(self, status, payload):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


if __name__ == "__main__":
    os.chdir(WORKSPACE)
    server = ThreadingHTTPServer((HOST, PORT), OrderHandler)
    print(f"A VAN CANDIES order server running at http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
