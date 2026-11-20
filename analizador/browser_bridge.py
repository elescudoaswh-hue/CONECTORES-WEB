"""Conserva el reporte y las familias dentro del Web Worker del navegador."""

import json
from pathlib import Path
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

from analizador.analysis import analyze_report, apply_colors, apply_family
from analizador.core.engine import (
    COLOR_BUTTONS, InitialReportReader, default_assignments,
    load_family_database, matching_families, normalize_color,
)

_report = None
_families = {}


def dispatch(action, payload, path):
    global _report, _families
    data = json.loads(payload)
    if action == "report":
        try:
            report = InitialReportReader(Path(path)).read()
        except (BadZipFile, KeyError, ParseError) as error:
            raise ValueError("El archivo no es un Initial Report .xlsx válido.") from error
        report.source_path = Path(Path(data["name"]).name)
        _report = report
        result = {
            "parts": _report.part_numbers,
            "headers": _report.report_headers,
            "assignments": default_assignments(_report),
            "connectorCount": len(_report.connectors),
            "matching": matching_families(_report, _families),
            "palette": COLOR_BUTTONS,
        }
    elif action == "families":
        _families = load_family_database(Path(path))
        result = {
            "families": sorted(_families, key=str.casefold),
            "matching": matching_families(_report, _families) if _report else [],
        }
    elif action == "normalize_color":
        result = normalize_color(data["color"])
    else:
        if _report is None:
            raise ValueError("Primero carga el Initial Report.")
        if action == "apply_family":
            result = apply_family(_report, _families, data["family"], data["groups"])
        elif action == "apply_colors":
            result = apply_colors(_report, data["assignments"], data["groups"])
        elif action == "analyze":
            result = analyze_report(
                _report, data["assignments"], data["groups"],
                data["drawings"], data.get("exportFamily"),
            )
        else:
            raise ValueError("Operación de análisis desconocida.")
    return json.dumps(result, ensure_ascii=False)
