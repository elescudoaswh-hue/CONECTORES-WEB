"""Reporte mínimo reproducible, sin dependencias de Excel u openpyxl."""
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile, ZIP_DEFLATED

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def column_name(number):
    text = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        text = chr(65 + remainder) + text
    return text


def make_report(path: Path, headers=None, extra_accessories=None):
    headers = headers or [
        "W3KT-14401-AAA A01 F0015930201",
        "W3KT-14401-ABA A01 F0015930301",
        "W3KT-14401-ERA A01",
    ]
    yes = ["X"] * len(headers)
    a = ["X"] + ["-"] * (len(headers) - 1)
    b = ["-", "X"] + ["-"] * (len(headers) - 2)
    last = ["-"] * (len(headers) - 1) + ["X"]
    circuit_headers = ["Circuit Name", "Wire Gauge", "Wire Color", "From", "From Cavity", "To", "To Cavity"]
    sheets = {
        "Connectors": (
            ["Connector", "Component", "Type"] + headers,
            [["X321A", "7289656830", "Connector"] + yes,
             ["X322A", "7289656830", "Connector"] + yes],
        ),
        "Circuits": (
            circuit_headers + headers,
            [["LEFT", "0.35", "RD/BU", "X321A", "1", "UNLISTED", "1"] + yes,
             ["RIGHT", "0.5", "BK", "X322A", "1", "UNLISTED", "2"] + yes,
             ["Z910AF", "0.5", "BK", "X322A", "6", "UNLISTED", "3"] + a,
             ["Z910AM", "0.50", "BK", "X322A", "6", "UNLISTED", "4"] + b],
        ),
        "Circuits Detail": (
            ["Circuit Name", "Wire Gauge", "Wire Color", "From Connector", "From Cavity", "To Connector", "To Cavity"] + headers,
            [["UNSAFE-SHARED", "1", "WH", "7289656830", "1", "", ""] + yes],
        ),
        "Connectors Accessories": (
            ["Connector", "Number of Component", "Description", "Cavity", "Quantity"] + headers,
            [["X322A", "7158331550", "CAVITY PLUG", "2", "1"] + last,
             ["X321A", "7152574330", "CLIP", "", "1"] + yes],
        ),
    }
    if extra_accessories:
        sheets["Connectors Accessories"][1].extend(extra_accessories)
    workbook = ET.Element(f"{{{NS}}}workbook")
    listing = ET.SubElement(workbook, f"{{{NS}}}sheets")
    relationships = ET.Element(f"{{{PACKAGE_REL}}}Relationships")
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for index, (name, (titles, values)) in enumerate(sheets.items(), 1):
            ET.SubElement(listing, f"{{{NS}}}sheet", {
                "name": name, "sheetId": str(index), f"{{{REL}}}id": f"rId{index}",
            })
            ET.SubElement(relationships, f"{{{PACKAGE_REL}}}Relationship", {
                "Id": f"rId{index}", "Type": f"{REL}/worksheet", "Target": f"worksheets/sheet{index}.xml",
            })
            sheet = ET.Element(f"{{{NS}}}worksheet")
            ET.SubElement(sheet, f"{{{NS}}}dimension", {"ref": "A1"})
            data = ET.SubElement(sheet, f"{{{NS}}}sheetData")
            for row_number, row_values in [(1, ["REPORTE DE PRUEBA"]), (5, titles)] + list(enumerate(values, 6)):
                row = ET.SubElement(data, f"{{{NS}}}row", {"r": str(row_number)})
                for column, value in enumerate(row_values, 1):
                    cell = ET.SubElement(row, f"{{{NS}}}c", {"r": f"{column_name(column)}{row_number}", "t": "inlineStr"})
                    string = ET.SubElement(cell, f"{{{NS}}}is")
                    ET.SubElement(string, f"{{{NS}}}t").text = str(value)
            archive.writestr(f"xl/worksheets/sheet{index}.xml", ET.tostring(sheet))
        archive.writestr("xl/workbook.xml", ET.tostring(workbook))
        archive.writestr("xl/_rels/workbook.xml.rels", ET.tostring(relationships))
    return path
