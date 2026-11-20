"""Compara la lectura del programa con las matrices originales del XLSX.

Uso: python validar_reporte.py reporte.xlsx familias_guardadas.json
No abre AutoCAD ni modifica el reporte o la base de familias.
"""
from __future__ import annotations
import argparse
from collections import defaultdict
import json
from pathlib import Path
import posixpath
import re
import xml.etree.ElementTree as ET
from zipfile import ZipFile
from analizador.core.engine import (
    InitialReportReader, load_family_database, matching_families,
    assignments_from_family, calculate_requirements, adapt_group_expression,
    parse_group_expression, cavity_analysis,
)
from analizador.services.generation import build_generation_instances
from analizador.services.selection import connector_selection_key

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def original_tables(path):
    """Lectura independiente del lector del programa y de la dimension de Excel."""
    tables = {}
    with ZipFile(path) as archive:
        strings = []
        if "xl/sharedStrings.xml" in archive.namelist():
            strings = ["".join(node.itertext()) for node in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
        relationships = {
            node.attrib["Id"]: node.attrib["Target"]
            for node in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        }
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        for sheet in workbook.findall("s:sheets/s:sheet", NS):
            target = relationships[sheet.attrib[f"{{{REL}}}id"]]
            filename = target.lstrip("/") if target.startswith("/") else posixpath.normpath("xl/" + target)
            rows = []
            headers = {}
            for node in ET.fromstring(archive.read(filename)).findall("s:sheetData/s:row", NS):
                values = {}
                for cell in node.findall("s:c", NS):
                    column = re.sub(r"\d", "", cell.attrib["r"])
                    kind = cell.attrib.get("t", "")
                    text = cell.findtext("s:v", default="", namespaces=NS)
                    if kind == "s":
                        text = strings[int(text)] if text else ""
                    elif kind == "inlineStr":
                        text = "".join(cell.find("s:is", NS).itertext())
                    values[column] = text.strip()
                number = int(node.attrib["r"])
                if number == 5:
                    headers = values
                elif number > 5:
                    rows.append(values)
            tables[sheet.attrib["name"]] = (headers, rows)
    return tables


def source_configurations(headers):
    columns = {}
    for column, title in headers.items():
        tokens = title.upper().split()
        if not tokens:
            continue
        first = tokens[0]
        if len(tokens) > 1 and "-" in first and re.fullmatch(r"[A-Z][0-9_]+", tokens[1]):
            columns[column] = first.rsplit("-", 1)[1]
        else:
            numeric = re.match(r"^(\d{10}|\d{8})(?=[A-Z]|$)", first)
            if numeric:
                columns[column] = numeric.group(1)
    return columns


def title_columns(headers):
    return {title.casefold(): col for col, title in headers.items()}


def source_scope(row, columns):
    return {identifier for col, identifier in columns.items() if row.get(col, "").upper() == "X"}


def component(value):
    text = value.strip().upper()
    if text.startswith("TEMP") and text[4:].isdigit():
        return text[4:]
    if all(c.isdigit() or c.isspace() or c == "-" for c in text):
        return "".join(c for c in text if c.isdigit())
    return text


def cavity(value):
    return str(int(float(value))) if re.fullmatch(r"\d+(?:\.0+)?", value) else value.upper()


def validate(path, database_path):
    report = InitialReportReader(path).read()
    tables = original_tables(path)
    headers, rows = tables["Connectors"]
    columns = source_configurations(headers)
    expected_codes = list(columns.values())
    assert report.part_numbers == expected_codes, "Diferencia en configuraciones del encabezado"
    lookup = title_columns(headers)
    expected_connectors = {}
    for row in rows:
        name = row.get(lookup["connector"], "")
        number = component(row.get(lookup["component"], ""))
        kind = row.get(lookup["type"], "")
        if name and number:
            expected_connectors[(name, number, kind)] = source_scope(row, columns)
    actual_connectors = {(c.name, c.component, c.kind): c.applicable_parts for c in report.connectors}
    assert actual_connectors == expected_connectors, "Diferencia en aplicaciones de conectores"

    by_name = {name.casefold(): (name, number, scope) for (name, number, _), scope in expected_connectors.items()}
    headers, rows = tables["Circuits"]
    lookup = title_columns(headers)
    columns = source_configurations(headers)
    assert set(columns.values()) == set(expected_codes), "Circuits usa otras configuraciones"
    expected_circuits = defaultdict(set)
    for row in rows:
        name = row.get(lookup["circuit name"], "")
        if not name:
            continue
        scope = source_scope(row, columns)
        gauge = row.get(lookup["wire gauge"], "")
        color = row.get(lookup["wire color"], "").upper()
        for endpoint in ("from", "to"):
            owner = by_name.get(row.get(lookup[endpoint], "").casefold())
            cell = row.get(lookup[f"{endpoint} cavity"], "")
            if not owner or not cell:
                continue
            connector_name, number, connector_scope = owner
            applies = scope & connector_scope
            if applies:
                expected_circuits[(number, connector_name, cavity(cell), name, gauge, color)].update(applies)
    actual_circuits = {
        (c.component, c.connector_name, c.cavity, c.circuit, c.gauge, c.color): c.applicable_parts
        for c in report.circuits if c.connector_name
    }
    assert actual_circuits == dict(expected_circuits), "Diferencia en circuito / conector / cavidad / aplicaciones"

    headers, rows = tables["Connectors Accessories"]
    lookup = title_columns(headers)
    columns = source_configurations(headers)
    if columns:
        assert set(columns.values()) == set(expected_codes), "Accesorios usa otras configuraciones"
    expected_blocked = defaultdict(set)
    for row in rows:
        description = row.get(lookup["description"], "")
        part_number = row.get(lookup["number of component"], "")
        blocked_values = {
            re.sub(r"[\s_-]+", " ", value.strip().upper())
            for value in (description, part_number)
        }
        if "BLOCKED CAVITY" not in blocked_values:
            continue
        connector = row.get(lookup["connector"], "")
        cells = row.get(lookup["cavity"], "")
        if connector:
            for cell in re.split(r"[,;\s]+", cells):
                if cell:
                    expected_blocked[(connector, cavity(cell))].update(source_scope(row, columns))
    actual_blocked = {
        (p.connector_name, p.cavity): p.applicable_parts
        for p in report.plugs if p.part_number == "BLOCKED CAVITY"
    }
    assert actual_blocked == dict(expected_blocked), "Diferencia en bloqueos / cavidades / aplicaciones"
    for item in [*report.plugs, *report.accessories]:
        if item.part_number == "BLOCKED CAVITY":
            continue  # Ya se comparó por conector y cavidad, aunque no tenga N/P.
        scopes = set()
        for row in rows:
            if (row.get(lookup["connector"], "") == item.connector_name
                    and component(row.get(lookup["number of component"], "")) == item.part_number):
                scopes.update(source_scope(row, columns))
        # Algunos reportes agregan accesorios desde Subassemblies. Estos no
        # se comparan contra Connectors Accessories cuando no existe esa fila.
        if scopes:
            assert scopes == item.applicable_parts, f"Diferencia en accesorio de {item.connector_name}"

    database = load_family_database(database_path)
    families = matching_families(report, database)
    summary = {
        "configuraciones": len(expected_codes),
        "conectores_leidos": len(report.connectors),
        "registros_circuito_por_extremo": len(report.circuits),
        "registros_plug": sum(p.part_number != "BLOCKED CAVITY" for p in report.plugs),
        "registros_bloqueo": len(actual_blocked),
        "registros_accesorio": len(report.accessories),
        "familias_con_cobertura_completa": families,
    }
    if len(families) == 1:
        assignments, conflicts = assignments_from_family(database, families[0])
        assignments = {code: assignments[code] for code in expected_codes}
        assert not set(conflicts) & set(expected_codes)
        groups = parse_group_expression(adapt_group_expression("ROSA+VERDE;CAFE;AZUL", assignments.values()), assignments)
        requirements = calculate_requirements(report, assignments, groups)
        instances = build_generation_instances(report, requirements, {}, assignments)
        for group, payload in instances.items():
            for item in payload:
                connector = item["connector"]
                data, _ = cavity_analysis(report, connector, requirements[group].part_numbers)
                source_names = {
                    c.circuit for c in report.circuits
                    if c.connector_name == connector.name and c.applicable_parts & connector.applicable_parts
                }
                for cell in data.values():
                    options = cell.get("VARIANTES") or [cell]
                    for option in options:
                        assert set(option.get("CIRCUITOS_UNIDOS", [])) <= source_names, "Circuito ajeno al conector"
        first_group = next(iter(requirements))
        first_connector = requirements[first_group].connectors[0]
        selected = {connector_selection_key(first_group, first_connector)}
        chosen = build_generation_instances(report, requirements, {}, assignments, selected)
        assert sum(len(items) for items in chosen.values()) == 1, "La seleccion no limita la generacion"
        summary.update({
            "configuraciones_con_color": len(assignments),
            "colores": sorted(set(assignments.values())),
            "posiciones_a_generar": sum(len(items) for items in instances.values()),
            "dibujos_unicos": len({c.component for r in requirements.values() for c in r.connectors}),
        })
    summary["resultado"] = "OK: matrices, circuitos y seleccion verificados"
    return summary


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("reporte", type=Path)
    parser.add_argument("base_json", type=Path)
    args = parser.parse_args()
    print(json.dumps(validate(args.reporte, args.base_json), ensure_ascii=False, indent=2))
