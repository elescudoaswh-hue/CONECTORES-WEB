"""
Analizador de conectores por Initial Report y generador de DXF.

Versión reforzada para Windows + AutoCAD 2026.

Funciones principales:
  * Lee Initial Reports XLSX aunque tengan una dimensión de hoja incorrecta.
  * Permite asignar números de parte a colores.
  * Permite unir colores, por ejemplo: ROSA+VERDE;CAFE;AZUL.
  * Excluye únicamente los registros Type=Splice.
  * Cuenta posiciones lógicas y dibujos físicos únicos.
  * Busca los DWG por número de parte, ignorando guiones, espacios y sufijos (1).
  * Inserta los DWG encontrados en un solo DXF mediante AutoCAD.
  * Completa los atributos CIRCUITO, CALIBRE, COLOR, OT y TIPO de cada cavidad.
  * Rellena el cuadro inferior con el color sólido del cable y su franja secundaria.
  * Muestra aplicaciones SOLO/N/A y el nombre de cada conector faltante.
  * Reintenta automáticamente las operaciones COM cuando AutoCAD está ocupado.
  * Reserva una posición visible para cada DWG faltante o que no pudo insertarse.
  * Lee CAVITY PLUG desde Connectors Accessories y lo coloca por cavidad.
  * Usa colores RGB reales y puntos de aplicación iguales a la leyenda.
  * Divide cada número de plug en dos líneas dentro de la cavidad.
  * Dibuja como cavidades extra únicamente alternativas con circuito, calibre y color.
  * Une circuitos de la misma cavidad cuando comparten calibre y color.
  * Separa por nombre los circuitos de conectores que comparten el mismo CON.
  * Recuerda las rutas del reporte, carpeta de conectores y DXF de salida.
  * Carga una base JSON de familias y asigna automáticamente sus colores.
  * Presenta una interfaz oscura Yazaki con paneles y controles organizados.
  * Prioriza siempre el circuito con cable cuando una cavidad también lleva plug.
  * Evita falsos rótulos de faltante y valida el DXF después de SaveAs.
  * Usa Arial Regular en los textos generados y atributos actualizados.
  * Agrupa las aplicaciones de cuatro en cuatro, colorea cada número de parte
    según su familia y usa un fondo amarillo intenso.
  * Agrega debajo del conector los accesorios aplicables: PLUG, PIA, CLIP,
    ASSEMBLY, COVER, STRAP, TUBE, BUTIL y TERMINAL.
  * Usa CAVR como cavidad física cuando difiere de CAV y distribuye circuitos
    distintos entre cuadros que representan la misma cavidad lógica.
  * No repite la última opción cuando existen más cuadros CAVR que circuitos.
  * Representa BLOCKED CAVITY como cuadro vacío con una línea diagonal.
  * Reintenta en pasadas separadas las etiquetas y cavidades gráficas fallidas.
  * Valida los textos creados antes de repetirlos cuando AutoCAD falla en Update.
  * Verifica cada plug dentro de su cavidad y detiene el guardado si falta alguno.
  * Coloca las aplicaciones exteriores a la derecha del límite real del conector.
  * Acomoda cada fila según la huella real de conectores, etiquetas y variantes.
  * Coloca siempre el nombre y los componentes inmediatamente debajo del dibujo.
  * Lee los circuitos COAX de ANTENNA CABLE desde la hoja Circuits.
  * Genera una base JSON reutilizable y un reporte de faltantes/duplicados.

Requisito para generar el DXF:
    pip install pywin32

El análisis, conteo y JSON funcionan sin AutoCAD. La generación del DXF requiere
Windows, AutoCAD instalado y pywin32.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import posixpath
import re
import shutil
import sys
import tempfile
import threading
import time
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, Iterable, List, Optional, Sequence, Set, Tuple
import xml.etree.ElementTree as ET



APP_NAME = "Analizador de Conectores DXF"
APP_VERSION = "0.7.4"


def settings_file_path() -> Path:
    """Configuración persistente por usuario, fuera de la carpeta del programa."""
    base = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
    folder = Path(base) if base else Path.home() / ".config"
    return folder / "Analizador_Conectores_DXF" / "rutas.json"


def load_saved_paths() -> Dict[str, str]:
    try:
        data = json.loads(settings_file_path().read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            return {}
        return {
            key: str(data.get(key, ""))
            for key in ("report", "connectors", "output", "database", "family")
        }
    except (OSError, ValueError, TypeError):
        return {}


def save_paths(
    report: str,
    connectors: str,
    output: str,
    database: str = "",
    family: str = "",
) -> None:
    target = settings_file_path()
    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(
            json.dumps(
                {
                    "report": str(report).strip(),
                    "connectors": str(connectors).strip(),
                    "output": str(output).strip(),
                    "database": str(database).strip(),
                    "family": str(family).strip(),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
    except OSError:
        # La generación no debe detenerse si Windows bloquea la configuración.
        pass

MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
DOC_REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

DEFAULT_COLOR_PARTS: Dict[str, Set[str]] = {
    "ROSA": {
        "68786194", "68786195", "68786196", "68786197", "68786198",
        "68786199", "68786200", "68786201", "68786202", "68786203",
        "68786204", "68786205", "68786206", "68786207", "68786208",
        "68786209", "68786210", "68786211", "68786212",
    },
    "VERDE": {"68786213", "68786214", "68786215", "68786216"},
    "CAFE": {"68786185", "68786186", "68786187", "68786188", "68786189"},
    "AZUL": {
        "68786217", "68786218", "68786219", "68786220", "68786221",
        "68786222", "68786223", "68786224", "68786225", "68786226",
        "68786227", "68786228", "68786229", "68786230",
    },
}

COLOR_BUTTONS = {
    "ROSA": ("#f4a6c8", "#111111"),
    "VERDE": ("#39a852", "#ffffff"),
    "CAFE": ("#8b5a2b", "#ffffff"),
    "AZUL": ("#3979d3", "#ffffff"),
    "CELESTE": ("#18cbd1", "#111111"),
    "MAGENTA": ("#e500d7", "#ffffff"),
}

WIRE_ACI = {
    "BK": 250, "WH": 255, "RD": 1, "GN": 3, "DG": 94,
    "BU": 5, "DB": 174, "YE": 2, "OG": 30, "VT": 6,
    "PK": 231, "BG": 33, "BN": 32, "GY": 8, "GR": 8,
    "LB": 4, "TN": 32, "SI": 9,
}

WIRE_RGB: Dict[str, Tuple[int, int, int]] = {
    "BK": (0, 0, 0),
    "WH": (255, 255, 255),
    "RD": (255, 0, 0),
    "GN": (0, 175, 0),
    "DG": (0, 85, 0),
    "BU": (0, 0, 255),
    "DB": (0, 0, 150),
    "YE": (255, 255, 0),
    "OG": (255, 95, 0),
    "VT": (175, 0, 255),
    "PK": (255, 0, 190),
    "BG": (225, 195, 165),
    "BN": (145, 75, 25),
    "GY": (125, 125, 125),
    "GR": (125, 125, 125),
    "LB": (0, 190, 255),
    "TN": (195, 155, 105),
    "SI": (195, 195, 195),
}

# Abreviaturas en español mostradas dentro de las cavidades. Los códigos
# originales del Initial Report se conservan para dibujar los colores reales.
# Se siguen las abreviaturas suministradas por Ingeniería Yazaki.
WIRE_SPANISH_ABBREVIATIONS: Dict[str, str] = {
    "BG": "CC",  # Café claro
    "DB": "AZ",  # Azul
    "DG": "V",   # Verde oscuro
    "GN": "VC",  # Verde claro
    "GY": "GR",  # Gris
    "GR": "GR",  # Gris (alias)
    "BU": "AZ",  # Azul claro, según la base suministrada
    "LB": "AZ",  # Azul claro (alias)
    "OG": "NA",  # Naranja
    "BR": "CA",  # Café
    "BN": "CA",  # Café (código usado por los Initial Report actuales)
    "TN": "CC",  # Café claro / tan
    "BK": "N",   # Negro
    "PK": "RS",  # Rosa
    "VT": "VI",  # Violeta
    "RD": "R",   # Rojo
    "R": "R",    # Rojo (alias presente en la base)
    "WH": "BL",  # Blanco
    "YE": "AM",  # Amarillo
    "SI": "PL",  # Plata
}

# La secuencia conserva los siete colores históricos y añade tonos distintos
# para conectores que tengan más condiciones visibles. El mapa se construye
# solamente con las condiciones que realmente aparecen en el conector.
APPLICATION_NOTE_ACI = [3, 6, 32, 5, 1, 4, 30, 94, 174, 231, 8, 2]

# Colores de los números de parte en las etiquetas SOLO/N/A.
PART_FAMILY_ACI = {
    "ROSA": 6,
    "VERDE": 3,
    "CAFE": 32,
    "AZUL": 5,
    "CELESTE": 4,
    "MAGENTA": 6,
}

BRIGHT_YELLOW_RGB = (255, 255, 0)
APPLICATIONS_PER_LINE = 4

ACCESSORY_DESCRIPTION_LABELS = (
    ("TERMINAL POSITION ASSURANCE", "TERMINAL"),
    ("TIE-STRAP", "STRAP"),
    ("STRAP", "STRAP"),
    ("CLIP", "CLIP"),
    ("COVER", "COVER"),
    ("TUBE", "TUBE"),
    ("BUTYL", "BUTIL"),
    ("BUTIL", "BUTIL"),
)

ACCESSORY_LABEL_ORDER = {
    "PLUG": 0,
    "PIA": 1,
    "CLIP": 2,
    "ASSEMBLY": 3,
    "COVER": 4,
    "STRAP": 5,
    "TUBE": 6,
    "BUTIL": 7,
    "TERMINAL": 8,
}


def normalize_part_number(value: object) -> str:
    """Conserva completo un componente, incluyendo letras y ceros iniciales."""
    text = str(value if value is not None else "").strip().upper()
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    temporary = re.fullmatch(r"TEMP[\s_-]*(\d+)", text)
    if temporary:
        return temporary.group(1)
    if re.fullmatch(r"[\d\s-]+", text):
        return re.sub(r"\D", "", text)
    return text


def configuration_from_header(value: object) -> str:
    """Identificador de una columna de configuración, separado del N/P F."""
    text = re.sub(r"\s+", " ", str(value or "").strip().upper())
    match = re.fullmatch(
        r"[A-Z0-9_-]+?[-_]([A-Z][A-Z0-9]{0,7}) "
        r"[A-Z](?:\d[A-Z0-9_]*|_\d[A-Z0-9_]*)(?: F\d{8,}(?: .*)?)?",
        text,
    )
    if match:
        return match.group(1)
    match = re.fullmatch(r"(\d{10}|\d{8})(?:[A-Z][A-Z0-9_-]*)?(?: .*)?", text)
    return match.group(1) if match else ""


def normalize_configuration(value: object) -> str:
    """Normaliza claves JSON y encabezados sin convertir AAA en el código F."""
    text = normalize_part_number(value)
    return configuration_from_header(text) or text


def configuration_sort_key(value: object):
    text = str(value)
    return (0, int(text)) if text.isdigit() else (1, text)


def normalize_color(value: str) -> str:
    text = value.strip().upper().replace("É", "E")
    text = re.sub(r"\s+", "_", text)
    return text


def column_number(cell_reference: str) -> int:
    match = re.match(r"([A-Z]+)", cell_reference.upper())
    if not match:
        return 0
    number = 0
    for character in match.group(1):
        number = number * 26 + ord(character) - 64
    return number


def short_configuration(part_number: str) -> str:
    identifier = normalize_configuration(part_number)
    return identifier[-3:] if identifier.isdigit() and len(identifier) >= 3 else identifier


def application_part_colors(
    part_numbers: Iterable[str], assignments: Dict[str, str]
) -> Dict[str, str]:
    """Relaciona el sufijo visible de cada número de parte con su familia."""
    return {
        short_configuration(part): normalize_color(assignments.get(part, ""))
        for part in part_numbers
    }


def format_colored_application_mtext(
    note_type: str,
    values: Sequence[str],
    part_colors: Dict[str, str],
    per_line: int = APPLICATIONS_PER_LINE,
) -> Tuple[str, int, int]:
    """Forma MTEXT de cuatro aplicaciones por línea y color por familia.

    Devuelve contenido MTEXT, cantidad de líneas y longitud visible máxima.
    """
    clean_values = [str(value) for value in values if str(value)]
    if not clean_values:
        visible = f"{note_type}:TODOS" if note_type else "APLICA:TODOS"
        return f"{{\\C250;{visible}}}", 1, len(visible)

    per_line = max(1, int(per_line))
    rows = [
        clean_values[index:index + per_line]
        for index in range(0, len(clean_values), per_line)
    ]
    first_color = PART_FAMILY_ACI.get(
        normalize_color(part_colors.get(clean_values[0], "")), 250
    )
    prefix = f"{note_type}:" if note_type else "APLICA:"
    formatted_rows: List[str] = []
    visible_lengths: List[int] = []
    for row_index, row in enumerate(rows):
        formatted = []
        for value_index, value in enumerate(row):
            color = PART_FAMILY_ACI.get(
                normalize_color(part_colors.get(value, "")), 250
            )
            comma = "," if value_index + 1 < len(row) else ""
            formatted.append(f"{{\\C{color};{value}{comma}}}")
        row_prefix = f"{{\\C{first_color};{prefix}}}" if row_index == 0 else ""
        formatted_rows.append(row_prefix + "".join(formatted))
        visible_lengths.append(len(prefix if row_index == 0 else "") + len(",".join(row)))
    return "\\P".join(formatted_rows), len(rows), max(visible_lengths)


def normalize_cavity(value: object) -> str:
    """Normaliza 01, 1.0 y 1 como la misma cavidad sin alterar nombres alfanuméricos."""
    text = str(value or "").strip()
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        return str(int(float(text)))
    return text.upper()


def select_cavity_option_for_block(
    cavity_data: Dict[str, Dict[str, object]],
    cav_value: object,
    cavr_value: object,
    cavr_occurrences: Dict[str, int],
) -> Tuple[str, Optional[Dict[str, object]], bool, int]:
    """Selecciona la opción de un bloque, usando CAVR como respaldo.

    Cuando CAVR identifica una cavidad distinta y válida del reporte, se usa
    como número físico. Si CAV y CAVR son iguales, se conserva CAV. Cuando
    varios cuadros apuntan a una misma cavidad lógica, cada opción se usa una
    sola vez y los cuadros sobrantes quedan vacíos.
    """
    cavity = normalize_cavity(cav_value)
    direct = cavity_data.get(cavity)
    reverse_cavity = normalize_cavity(cavr_value)
    fallback = cavity_data.get(reverse_cavity)
    if fallback is not None and (direct is None or reverse_cavity != cavity):
        logical_cavity = reverse_cavity
        selected_data = fallback
        matched_by_cavr = True
    elif direct is not None:
        logical_cavity = cavity
        selected_data = direct
        matched_by_cavr = False
    elif fallback is not None:
        logical_cavity = reverse_cavity
        selected_data = fallback
        matched_by_cavr = True
    else:
        return "", None, False, 0

    occurrence = cavr_occurrences.get(logical_cavity, 0)
    cavr_occurrences[logical_cavity] = occurrence + 1
    options = selected_data.get("VARIANTES", [])
    if isinstance(options, list) and options:
        if occurrence >= len(options):
            return logical_cavity, None, matched_by_cavr, occurrence
        selected = options[occurrence]
        if isinstance(selected, dict):
            return logical_cavity, selected, matched_by_cavr, occurrence
        return logical_cavity, None, matched_by_cavr, occurrence

    # Una cavidad sin variantes puede ocupar un único cuadro físico.
    if occurrence == 0:
        return logical_cavity, selected_data, matched_by_cavr, 0
    return logical_cavity, None, matched_by_cavr, occurrence


def split_plug_number(value: object) -> Tuple[str, str]:
    """Divide un plug en dos líneas equilibradas para una cavidad pequeña."""
    text = str(value or "").strip().upper()
    if not text:
        return "", ""
    if text.isdigit() and len(text) > 1:
        midpoint = len(text) // 2
        return text[:midpoint], text[midpoint:]
    words = text.split()
    if len(words) > 1:
        midpoint = max(1, len(words) // 2)
        return " ".join(words[:midpoint]), " ".join(words[midpoint:])
    midpoint = max(1, len(text) // 2)
    return text[:midpoint], text[midpoint:]


def is_blocked_cavity(value: object) -> bool:
    """Reconoce el componente especial que representa una cavidad bloqueada."""
    text = re.sub(r"[\s_-]+", " ", str(value or "").strip().upper())
    return text == "BLOCKED CAVITY"


def blocked_cavity_diagonal(
    segments: Sequence[Tuple[Tuple[float, float], Tuple[float, float]]],
) -> Tuple[Tuple[float, float], Tuple[float, float]]:
    """Une la esquina inferior izquierda con la esquina del recuadro del número.

    Solo considera bordes horizontales y verticales. Las diagonales o textos
    existentes no cambian la medida del cuadro. Si la definición no permite
    reconocerlo, se utiliza el cuadro estándar de 10 x 10 de las plantillas.
    """
    tolerance = 0.001
    borders = [
        (first, second) for first, second in segments
        if (abs(first[0] - second[0]) <= tolerance
            or abs(first[1] - second[1]) <= tolerance)
        and math.dist(first, second) > tolerance
    ]
    if not borders:
        return (0.0, 0.0), (7.0, 7.5)
    points = [point for segment in borders for point in segment]
    left, right = min(p[0] for p in points), max(p[0] for p in points)
    bottom, top = min(p[1] for p in points), max(p[1] for p in points)
    if right - left <= tolerance or top - bottom <= tolerance:
        return (0.0, 0.0), (7.0, 7.5)
    for first, second in borders:
        # Borde inferior del recuadro superior derecho del número.
        if abs(first[1] - second[1]) > tolerance:
            continue
        x_start, x_end = sorted((first[0], second[0]))
        y = first[1]
        if not (left + (right - left) / 2 < x_start < right - tolerance
                and bottom + (top - bottom) / 2 < y < top - tolerance
                and abs(x_end - right) <= tolerance):
            continue
        for vertical_first, vertical_second in borders:
            if abs(vertical_first[0] - vertical_second[0]) > tolerance:
                continue
            y_start, y_end = sorted((vertical_first[1], vertical_second[1]))
            if (abs(vertical_first[0] - x_start) <= tolerance
                    and abs(y_start - y) <= tolerance
                    and abs(y_end - top) <= tolerance):
                return (left, bottom), (x_start, y)
    return (left, bottom), (left + 0.7 * (right - left), bottom + 0.75 * (top - bottom))


def normalize_accessory_component(value: object) -> str:
    """Conserva números largos y componentes alfanuméricos de accesorios."""
    text = str(value or "").strip().upper()
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".", 1)[0]
    return text


def accessory_label_from_description(description: object) -> str:
    text = str(description or "").strip().upper()
    for pattern, label in ACCESSORY_DESCRIPTION_LABELS:
        if pattern in text:
            return label
    return ""


@dataclass
class ConnectorRecord:
    name: str
    component: str
    kind: str
    applicable_parts: Set[str] = field(default_factory=set)


@dataclass
class CircuitRecord:
    component: str
    cavity: str
    circuit: str
    gauge: str
    color: str
    applicable_parts: Set[str] = field(default_factory=set)
    connector_name: str = ""


@dataclass
class PlugRecord:
    connector_name: str
    cavity: str
    part_number: str
    applicable_parts: Set[str] = field(default_factory=set)


@dataclass
class AccessoryRecord:
    connector_name: str
    part_number: str
    label: str
    description: str
    quantity: str
    applicable_parts: Set[str] = field(default_factory=set)


@dataclass
class ReportData:
    source_path: Path
    report_label: str
    part_numbers: List[str]
    report_headers: Dict[str, str]
    connectors: List[ConnectorRecord]
    circuits: List[CircuitRecord]
    plugs: List[PlugRecord]
    accessories: List[AccessoryRecord]


@dataclass
class GroupRequirement:
    name: str
    colors: List[str]
    part_numbers: List[str]
    connectors: List[ConnectorRecord]

    @property
    def unique_components(self) -> List[str]:
        return list(dict.fromkeys(item.component for item in self.connectors))


class InitialReportReader:
    """Lector OOXML que ignora el dimension ref=A1 defectuoso del reporte."""

    def __init__(self, path: Path):
        self.path = Path(path)
        self.shared_strings: List[str] = []
        self.sheet_paths: Dict[str, str] = {}

    def read(self) -> ReportData:
        with zipfile.ZipFile(self.path) as archive:
            self._load_shared_strings(archive)
            self._load_sheet_paths(archive)
            connector_headers, connector_rows, label = self._read_table(
                archive, "Connectors", header_row=5
            )
            circuit_headers, circuit_rows, _ = self._read_table(
                archive, "Circuits Detail", header_row=5
            )
            named_circuit_headers, named_circuit_rows, _ = self._read_table(
                archive, "Circuits", header_row=5
            )
            plug_headers, plug_rows, _ = self._read_table(
                archive, "Connectors Accessories", header_row=5
            )
            if "Subassemblies" in self.sheet_paths:
                subassembly_headers, subassembly_rows, _ = self._read_table(
                    archive, "Subassemblies", header_row=5
                )
            else:
                subassembly_headers, subassembly_rows = {}, []

        connectors, part_numbers, report_headers = self._parse_connectors(
            connector_headers, connector_rows
        )
        if not part_numbers:
            raise ValueError(
                "El reporte no contiene columnas de configuración reconocibles. "
                "Se admiten números de 8/10 dígitos y encabezados como "
                "W3KT-14401-AAA A01 F0015930201."
            )
        detail_circuits = self._parse_circuits(circuit_headers, circuit_rows)
        named_circuits = self._parse_named_connector_circuits(
            named_circuit_headers,
            named_circuit_rows,
            connectors,
        )
        # La hoja Circuits conserva el nombre real de cada extremo y funciona
        # tanto con conectores normales como con PIGTAIL/ANTENNA CABLE. Se usa
        # como fuente principal para impedir cruces entre conectores y admitir
        # reportes donde Circuits Detail trae nombres en vez de componentes.
        named_components = {item.component for item in named_circuits}
        circuits = list(named_circuits)
        circuits.extend(
            item
            for item in detail_circuits
            if item.component not in named_components
        )
        plugs = self._parse_plugs(plug_headers, plug_rows)
        accessories = self._parse_accessories(plug_headers, plug_rows)
        if subassembly_headers:
            accessories.extend(
                self._parse_antenna_subassemblies(
                    subassembly_headers,
                    subassembly_rows,
                    connectors,
                )
            )
        return ReportData(
            source_path=self.path,
            report_label=label,
            part_numbers=part_numbers,
            report_headers=report_headers,
            connectors=connectors,
            circuits=circuits,
            plugs=plugs,
            accessories=accessories,
        )

    def _load_shared_strings(self, archive: zipfile.ZipFile) -> None:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        for item in root.findall(MAIN_NS + "si"):
            value = "".join(node.text or "" for node in item.iter(MAIN_NS + "t"))
            self.shared_strings.append(value)

    def _load_sheet_paths(self, archive: zipfile.ZipFile) -> None:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {
            rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
            for rel in relationships.findall(PKG_REL_NS + "Relationship")
        }
        for sheet in workbook.findall(".//" + MAIN_NS + "sheet"):
            name = sheet.attrib.get("name", "")
            relationship_id = sheet.attrib.get(DOC_REL_NS + "id", "")
            target = targets.get(relationship_id, "")
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            self.sheet_paths[name] = path

    def _cell_value(self, cell: ET.Element) -> str:
        cell_type = cell.attrib.get("t", "")
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.iter(MAIN_NS + "t"))
        value_node = cell.find(MAIN_NS + "v")
        value = "" if value_node is None else (value_node.text or "")
        if cell_type == "s" and value:
            try:
                return self.shared_strings[int(value)]
            except (ValueError, IndexError):
                return value
        return value

    def _iter_rows(self, archive: zipfile.ZipFile, sheet_name: str):
        path = self.sheet_paths.get(sheet_name)
        if not path or path not in archive.namelist():
            raise ValueError(f"No se encontró la hoja requerida: {sheet_name}")
        with archive.open(path) as handle:
            for _, element in ET.iterparse(handle, events=("end",)):
                if element.tag != MAIN_NS + "row":
                    continue
                row_number = int(element.attrib.get("r", "0") or 0)
                row: Dict[int, str] = {}
                for cell in element.findall(MAIN_NS + "c"):
                    index = column_number(cell.attrib.get("r", ""))
                    if index:
                        row[index] = self._cell_value(cell)
                yield row_number, row
                element.clear()

    def _read_table(
        self, archive: zipfile.ZipFile, sheet_name: str, header_row: int
    ) -> Tuple[Dict[int, str], List[Dict[int, str]], str]:
        headers: Dict[int, str] = {}
        rows: List[Dict[int, str]] = []
        report_label = ""
        for row_number, row in self._iter_rows(archive, sheet_name):
            if row_number == 1:
                report_label = row.get(1, "")
            elif row_number == header_row:
                headers = {index: value.strip() for index, value in row.items()}
            elif row_number > header_row and any(str(value).strip() for value in row.values()):
                rows.append(row)
        if not headers:
            raise ValueError(f"No se encontró la fila de encabezados en {sheet_name}")
        return headers, rows, report_label

    @staticmethod
    def _header_lookup(headers: Dict[int, str]) -> Dict[str, int]:
        return {value.strip().casefold(): index for index, value in headers.items()}

    @staticmethod
    def _configuration_columns(headers: Dict[int, str]) -> List[Tuple[int, str]]:
        columns: List[Tuple[int, str]] = []
        seen: Dict[str, int] = {}
        for index in sorted(headers):
            identifier = configuration_from_header(headers[index])
            if not identifier:
                continue
            if identifier in seen:
                raise ValueError(
                    f"La configuración {identifier} está repetida en las columnas "
                    f"{seen[identifier]} y {index}. Revisa los encabezados del reporte."
                )
            seen[identifier] = index
            columns.append((index, identifier))
        return columns

    def _parse_connectors(
        self, headers: Dict[int, str], rows: List[Dict[int, str]]
    ) -> Tuple[List[ConnectorRecord], List[str], Dict[str, str]]:
        lookup = self._header_lookup(headers)
        required = ["connector", "component", "type"]
        for name in required:
            if name not in lookup:
                raise ValueError(f"Falta la columna {name!r} en Connectors")

        config_columns = [
            (index, identifier, headers[index])
            for index, identifier in self._configuration_columns(headers)
        ]

        connectors: List[ConnectorRecord] = []
        for row in rows:
            name = str(row.get(lookup["connector"], "")).strip()
            component = normalize_part_number(row.get(lookup["component"], ""))
            kind = str(row.get(lookup["type"], "")).strip()
            if not name or not component:
                continue
            applicable = {
                part
                for index, part, _ in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            }
            connectors.append(ConnectorRecord(name, component, kind, applicable))

        part_numbers = [part for _, part, _ in config_columns]
        report_headers = {part: original for _, part, original in config_columns}
        return connectors, part_numbers, report_headers

    def _parse_circuits(
        self, headers: Dict[int, str], rows: List[Dict[int, str]]
    ) -> List[CircuitRecord]:
        lookup = self._header_lookup(headers)
        names = {
            "circuit": "circuit name",
            "gauge": "wire gauge",
            "color": "wire color",
            "from_component": "from connector",
            "from_cavity": "from cavity",
            "to_component": "to connector",
            "to_cavity": "to cavity",
        }
        for title in names.values():
            if title not in lookup:
                raise ValueError(f"Falta la columna {title!r} en Circuits Detail")

        config_columns = self._configuration_columns(headers)

        grouped: Dict[Tuple[str, str, str, str, str], Set[str]] = defaultdict(set)
        for row in rows:
            circuit = str(row.get(lookup[names["circuit"]], "")).strip()
            gauge = str(row.get(lookup[names["gauge"]], "")).strip()
            color = str(row.get(lookup[names["color"]], "")).strip().upper()
            applicable = {
                part
                for index, part in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            }
            endpoints = [
                (
                    normalize_part_number(row.get(lookup[names["from_component"]], "")),
                    str(row.get(lookup[names["from_cavity"]], "")).strip(),
                ),
                (
                    normalize_part_number(row.get(lookup[names["to_component"]], "")),
                    str(row.get(lookup[names["to_cavity"]], "")).strip(),
                ),
            ]
            for component, cavity in endpoints:
                if not component or not cavity:
                    continue
                key = (component, cavity, circuit, gauge, color)
                grouped[key].update(applicable)

        return [
            CircuitRecord(component, cavity, circuit, gauge, color, parts)
            for (component, cavity, circuit, gauge, color), parts in grouped.items()
        ]

    def _parse_named_connector_circuits(
        self,
        headers: Dict[int, str],
        rows: List[Dict[int, str]],
        connectors: Sequence[ConnectorRecord],
    ) -> List[CircuitRecord]:
        """Lee extremos por nombre para separar conectores con componente común."""
        lookup = self._header_lookup(headers)
        names = {
            "circuit": "circuit name",
            "gauge": "wire gauge",
            "color": "wire color",
            "from_connector": "from",
            "from_cavity": "from cavity",
            "to_connector": "to",
            "to_cavity": "to cavity",
        }
        for title in names.values():
            if title not in lookup:
                raise ValueError(f"Falta la columna {title!r} en Circuits")

        connector_by_name = {
            item.name.casefold(): item
            for item in connectors
        }
        if not connector_by_name:
            return []

        config_columns = self._configuration_columns(headers)

        grouped: Dict[
            Tuple[str, str, str, str, str, str], Set[str]
        ] = defaultdict(set)
        for row in rows:
            circuit = str(row.get(lookup[names["circuit"]], "")).strip()
            gauge = str(row.get(lookup[names["gauge"]], "")).strip()
            color = str(row.get(lookup[names["color"]], "")).strip().upper()
            if not circuit:
                continue
            applicable = {
                part
                for index, part in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            }
            endpoints = (
                (
                    str(row.get(lookup[names["from_connector"]], "")).strip(),
                    str(row.get(lookup[names["from_cavity"]], "")).strip(),
                ),
                (
                    str(row.get(lookup[names["to_connector"]], "")).strip(),
                    str(row.get(lookup[names["to_cavity"]], "")).strip(),
                ),
            )
            for connector_name, cavity in endpoints:
                connector = connector_by_name.get(connector_name.casefold())
                if connector is None or not cavity:
                    continue
                endpoint_parts = applicable & connector.applicable_parts
                if not endpoint_parts:
                    continue
                key = (
                    connector.component,
                    connector.name,
                    normalize_cavity(cavity),
                    circuit,
                    gauge,
                    color,
                )
                grouped[key].update(endpoint_parts)

        return [
            CircuitRecord(
                component,
                cavity,
                circuit,
                gauge,
                color,
                parts,
                connector_name=connector_name,
            )
            for (
                component,
                connector_name,
                cavity,
                circuit,
                gauge,
                color,
            ), parts in grouped.items()
        ]

    def _parse_plugs(
        self, headers: Dict[int, str], rows: List[Dict[int, str]]
    ) -> List[PlugRecord]:
        """Lee plugs y bloqueos explícitos, conservando su matriz de aplicación."""
        lookup = self._header_lookup(headers)
        required = ["connector", "number of component", "description", "cavity"]
        for title in required:
            if title not in lookup:
                raise ValueError(
                    f"Falta la columna {title!r} en Connectors Accessories"
                )

        config_columns = self._configuration_columns(headers)

        grouped: Dict[Tuple[str, str, str], Set[str]] = defaultdict(set)
        for row in rows:
            description = str(row.get(lookup["description"], "")).strip().upper()
            raw_part = str(row.get(lookup["number of component"], "")).strip()
            blocked = is_blocked_cavity(raw_part) or is_blocked_cavity(description)
            if not blocked and "CAVITY PLUG" not in description:
                continue
            connector_name = str(row.get(lookup["connector"], "")).strip()
            normalized_part = normalize_part_number(raw_part)
            plug_part = (
                "BLOCKED CAVITY" if blocked
                else normalized_part if normalized_part else raw_part.upper()
            )
            cavity_text = str(row.get(lookup["cavity"], "")).strip()
            applicable = {
                part
                for index, part in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            }
            if not connector_name or not plug_part or not cavity_text:
                continue
            cavities = [
                normalize_cavity(value)
                for value in re.split(r"[,;\s]+", cavity_text)
                if value.strip()
            ]
            for cavity in cavities:
                grouped[(connector_name, cavity, plug_part)].update(applicable)

        return [
            PlugRecord(connector_name, cavity, plug_part, parts)
            for (connector_name, cavity, plug_part), parts in grouped.items()
        ]

    def _parse_accessories(
        self, headers: Dict[int, str], rows: List[Dict[int, str]]
    ) -> List[AccessoryRecord]:
        """Lee los accesorios solicitados distintos de CAVITY PLUG."""
        lookup = self._header_lookup(headers)
        required = ["connector", "number of component", "description"]
        for title in required:
            if title not in lookup:
                raise ValueError(
                    f"Falta la columna {title!r} en Connectors Accessories"
                )

        config_columns = self._configuration_columns(headers)

        quantity_column = lookup.get("quantity")
        grouped: Dict[
            Tuple[str, str, str, str, str], Set[str]
        ] = defaultdict(set)
        for row in rows:
            description = str(row.get(lookup["description"], "")).strip().upper()
            if ("CAVITY PLUG" in description or is_blocked_cavity(description)
                    or is_blocked_cavity(row.get(lookup["number of component"], ""))):
                continue
            label = accessory_label_from_description(description)
            if not label:
                continue
            connector_name = str(row.get(lookup["connector"], "")).strip()
            part_number = normalize_accessory_component(
                row.get(lookup["number of component"], "")
            )
            quantity = (
                str(row.get(quantity_column, "")).strip()
                if quantity_column is not None
                else ""
            )
            applicable = {
                part
                for index, part in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            }
            if not connector_name or not part_number:
                continue
            grouped[
                (connector_name, part_number, label, description, quantity)
            ].update(applicable)

        return [
            AccessoryRecord(
                connector_name,
                part_number,
                label,
                description,
                quantity,
                parts,
            )
            for (
                connector_name,
                part_number,
                label,
                description,
                quantity,
            ), parts in grouped.items()
        ]

    def _parse_antenna_subassemblies(
        self,
        headers: Dict[int, str],
        rows: List[Dict[int, str]],
        connectors: Sequence[ConnectorRecord],
    ) -> List[AccessoryRecord]:
        """Lee PIA y ASSEMBLY de los conectores tipo ANTENNA CABLE.

        Solo se aceptan filas cuyo Object Name coincide exactamente con el
        nombre del conector. Las filas anidadas, por ejemplo
        ``Y530A1 | OTH54701``, se omiten porque su componente ya se registra
        como CLIP desde Connectors Accessories.
        """
        lookup = self._header_lookup(headers)
        required = ["assy object part number", "object name", "part number"]
        for title in required:
            if title not in lookup:
                raise ValueError(f"Falta la columna {title!r} en Subassemblies")

        antenna_by_name = {
            item.name.casefold(): item
            for item in connectors
            if "ANTENNA CABLE" in item.kind.upper()
        }
        if not antenna_by_name:
            return []

        config_columns = self._configuration_columns(headers)

        grouped: Dict[
            Tuple[str, str, str, str, str], Set[str]
        ] = defaultdict(set)
        for row in rows:
            object_name = str(row.get(lookup["object name"], "")).strip()
            connector = antenna_by_name.get(object_name.casefold())
            if connector is None:
                continue

            applicable = {
                part
                for index, part in config_columns
                if str(row.get(index, "")).strip().upper() == "X"
            } & connector.applicable_parts
            if not applicable:
                continue

            pia = normalize_accessory_component(
                row.get(lookup["part number"], "")
            )
            assembly = normalize_accessory_component(
                row.get(lookup["assy object part number"], "")
            )
            if pia:
                grouped[
                    (connector.name, pia, "PIA", "ANTENNA PIA", "")
                ].update(applicable)
            if assembly:
                grouped[
                    (
                        connector.name,
                        assembly,
                        "ASSEMBLY",
                        "ANTENNA ASSEMBLY",
                        "",
                    )
                ].update(applicable)

        return [
            AccessoryRecord(
                connector_name,
                part_number,
                label,
                description,
                quantity,
                parts,
            )
            for (
                connector_name,
                part_number,
                label,
                description,
                quantity,
            ), parts in grouped.items()
        ]


def default_assignments(report: ReportData) -> Dict[str, str]:
    result = {part: "" for part in report.part_numbers}
    for color, numbers in DEFAULT_COLOR_PARTS.items():
        for part in report.part_numbers:
            if part in numbers:
                result[part] = color
    return result


def load_family_database(path: Path) -> Dict[str, Dict[str, List[str]]]:
    """Carga una base familia -> color -> números de parte."""
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8-sig"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"No se pudo leer la base JSON: {error}") from error
    if not isinstance(raw, dict):
        raise ValueError("La base JSON debe contener un objeto de familias")

    families: Dict[str, Dict[str, List[str]]] = {}
    for raw_family, raw_colors in raw.items():
        family = str(raw_family).strip()
        if not family or family.startswith("_"):
            continue
        if not isinstance(raw_colors, dict):
            raise ValueError(f"La familia {family!r} no contiene una tabla de colores")
        colors: Dict[str, List[str]] = {}
        for raw_color, raw_parts in raw_colors.items():
            color = normalize_color(str(raw_color))
            if not color or not isinstance(raw_parts, list):
                raise ValueError(
                    f"El color {raw_color!r} de {family!r} debe contener una lista"
                )
            parts = list(
                dict.fromkeys(
                    normalize_configuration(value)
                    for value in raw_parts
                    if normalize_configuration(value)
                )
            )
            if parts:
                colors[color] = parts
        if colors:
            families[family] = colors
    if not families:
        raise ValueError("La base JSON no contiene familias utilizables")
    return families


def assignments_from_family(
    database: Dict[str, Dict[str, List[str]]],
    family: str,
) -> Tuple[Dict[str, str], Dict[str, List[str]]]:
    """Devuelve asignaciones inequívocas y conflictos de una familia."""
    if family not in database:
        raise ValueError(f"La familia {family!r} no existe en la base seleccionada")
    candidates: Dict[str, Set[str]] = defaultdict(set)
    for color, parts in database[family].items():
        for part in parts:
            candidates[normalize_configuration(part)].add(normalize_color(color))
    assignments = {
        part: next(iter(colors))
        for part, colors in candidates.items()
        if len(colors) == 1
    }
    conflicts = {
        part: sorted(colors)
        for part, colors in candidates.items()
        if len(colors) > 1
    }
    return assignments, conflicts


def matching_families(
    report: ReportData, database: Dict[str, Dict[str, List[str]]]
) -> List[str]:
    """Familias que cubren todo el reporte sin conflictos de color."""
    required = set(report.part_numbers)
    if not required:
        return []
    matches = []
    for family in database:
        assignments, conflicts = assignments_from_family(database, family)
        if required.issubset(assignments) and not (required & set(conflicts)):
            matches.append(family)
    return sorted(matches, key=str.casefold)


def adapt_group_expression(expression: str, colors: Iterable[str]) -> str:
    """Conserva grupos válidos y agrega aparte los colores nuevos."""
    available = {normalize_color(color) for color in colors if str(color).strip()}
    groups: List[str] = []
    used: Set[str] = set()
    for raw_group in str(expression).split(";"):
        members = [
            normalize_color(value)
            for value in re.split(r"\s*[+\-]\s*", raw_group.strip())
            if value.strip()
        ]
        kept = [color for color in members if color in available and color not in used]
        if kept:
            groups.append("+".join(kept))
            used.update(kept)
    preferred = list(COLOR_BUTTONS)
    remaining = sorted(
        available - used,
        key=lambda color: (
            preferred.index(color) if color in preferred else len(preferred),
            color,
        ),
    )
    groups.extend(remaining)
    return ";".join(groups)


def parse_group_expression(
    expression: str, assignments: Dict[str, str]
) -> Dict[str, List[str]]:
    available = {normalize_color(value) for value in assignments.values() if value.strip()}
    groups: Dict[str, List[str]] = {}
    for raw_group in expression.split(";"):
        raw_group = raw_group.strip()
        if not raw_group:
            continue
        colors = [
            normalize_color(item)
            for item in re.split(r"\s*[+\-]\s*", raw_group)
            if item.strip()
        ]
        colors = list(dict.fromkeys(colors))
        if not colors:
            continue
        name = "_".join(colors)
        groups[name] = colors

    used = {color for colors in groups.values() for color in colors}
    for color in sorted(available - used):
        groups[color] = [color]
    unknown = used - available
    if unknown:
        raise ValueError("Colores sin números asignados: " + ", ".join(sorted(unknown)))
    if not groups:
        raise ValueError("No se definieron grupos de colores")
    return groups


def calculate_requirements(
    report: ReportData,
    assignments: Dict[str, str],
    groups: Dict[str, List[str]],
) -> Dict[str, GroupRequirement]:
    result: Dict[str, GroupRequirement] = {}
    for group_name, colors in groups.items():
        color_set = {normalize_color(color) for color in colors}
        group_parts = [
            part
            for part in report.part_numbers
            if normalize_color(assignments.get(part, "")) in color_set
        ]
        part_set = set(group_parts)
        connectors = [
            item
            for item in report.connectors
            if item.kind.strip().casefold() != "splice"
            and bool(item.applicable_parts & part_set)
        ]
        result[group_name] = GroupRequirement(
            name=group_name,
            colors=colors,
            part_numbers=group_parts,
            connectors=connectors,
        )
    return result


def application_condition(
    scope: Set[str], universe: Set[str]
) -> Tuple[str, List[str]]:
    if not scope or not universe or scope == universe:
        return "", []
    missing = universe - scope
    if len(scope) <= len(missing):
        note_type = "SOLO"
        values = scope
    else:
        note_type = "N/A"
        values = missing
    suffixes = sorted(
        (short_configuration(part) for part in values),
        key=configuration_sort_key,
    )
    return note_type, suffixes


def connector_plug_parts(
    report: ReportData,
    connector: ConnectorRecord,
    group_parts: Iterable[str],
) -> List[str]:
    relevant = set(group_parts) & connector.applicable_parts
    return sorted(
        {
            item.part_number
            for item in report.plugs
            if item.connector_name == connector.name
            and bool(item.applicable_parts & relevant)
            and not is_blocked_cavity(item.part_number)
        },
        key=lambda value: (0, int(value)) if value.isdigit() else (1, value),
    )


def connector_accessory_records(
    report: ReportData,
    connector: ConnectorRecord,
    group_parts: Iterable[str],
) -> List[AccessoryRecord]:
    """Devuelve solo los accesorios que aplican al conector en este grupo."""
    relevant = set(group_parts) & connector.applicable_parts
    unique: Dict[Tuple[str, str], AccessoryRecord] = {}
    for item in report.accessories:
        if item.connector_name != connector.name:
            continue
        if not (item.applicable_parts & relevant):
            continue
        unique.setdefault((item.label, item.part_number), item)
    return sorted(
        unique.values(),
        key=lambda item: (
            ACCESSORY_LABEL_ORDER.get(item.label, 99),
            0 if item.part_number.isdigit() else 1,
            int(item.part_number) if item.part_number.isdigit() else item.part_number,
        ),
    )


def connector_component_labels(
    report: ReportData,
    connector: ConnectorRecord,
    group_parts: Iterable[str],
) -> List[str]:
    """Líneas mostradas debajo de CON:, primero plugs y luego accesorios."""
    labels = [
        f"PLUG:{part}"
        for part in connector_plug_parts(report, connector, group_parts)
    ]
    labels.extend(
        f"{item.label}:{item.part_number}"
        for item in connector_accessory_records(
            report, connector, group_parts
        )
    )
    return list(dict.fromkeys(labels))


def normalized_gauge_signature(value: object) -> str:
    """Iguala calibres equivalentes como 0.5 y 0.50 sin alterar su texto visible."""
    text = str(value or "").strip()
    try:
        number = float(text.replace(",", "."))
        return f"{number:.8f}".rstrip("0").rstrip(".")
    except (TypeError, ValueError):
        return text.upper()


def normalized_wire_color_signature(value: object) -> str:
    return "/".join(
        part.strip().upper()
        for part in str(value or "").split("/")
        if part.strip()
    )


def spanish_wire_color(value: object) -> str:
    """Traduce un código de cable conservando el orden base/franja."""
    text = str(value or "").strip().upper()
    # Compatibilidad con la entrada RDWH de la base suministrada.
    if text == "RDWH":
        text = "RD/WH"
    codes = [
        part.strip().upper()
        for part in text.split("/")
        if part.strip()
    ]
    return "/".join(
        WIRE_SPANISH_ABBREVIATIONS.get(code, code)
        for code in codes
    )


def compact_circuit_names(names: Iterable[object]) -> str:
    """Compacta Z910AF + Z910AM como Z910AF/AM."""
    unique = list(
        dict.fromkeys(str(name or "").strip().upper() for name in names if str(name or "").strip())
    )
    if not unique:
        return ""
    if len(unique) == 1:
        return unique[0]

    parsed = [re.fullmatch(r"(.*?\d)([A-Z]+)", name) for name in unique]
    if all(match is not None for match in parsed):
        bases = [match.group(1) for match in parsed if match is not None]
        suffixes = [match.group(2) for match in parsed if match is not None]
        if len(set(bases)) == 1 and len(suffixes) == len(unique):
            return unique[0] + "/" + "/".join(suffixes[1:])
    return "/".join(unique)


def circuit_belongs_to_connector(
    report: ReportData,
    circuit: CircuitRecord,
    connector: ConnectorRecord,
) -> bool:
    """Impide compartir circuitos entre conectores con el mismo componente."""
    if circuit.component != connector.component:
        return False
    if circuit.connector_name:
        return circuit.connector_name.casefold() == connector.name.casefold()

    # ``Circuits Detail`` no incluye el nombre del conector. Sus registros son
    # seguros solamente cuando el componente pertenece a un único conector.
    # Si dos nombres comparten el mismo CON, se exige el registro nominal de
    # la hoja ``Circuits`` y se prefiere dejar una cavidad pendiente antes que
    # insertar en ella un circuito perteneciente a otro conector.
    owners = {
        item.name.casefold()
        for item in report.connectors
        if item.component == connector.component
    }
    return len(owners) <= 1


def cavity_analysis(
    report: ReportData,
    connector: ConnectorRecord,
    group_parts: Iterable[str],
) -> Tuple[Dict[str, Dict[str, object]], List[str]]:
    """Combina cables y plugs; cualquier circuito con cable tiene prioridad."""
    relevant = set(group_parts) & connector.applicable_parts
    connector_note_type, connector_suffixes = application_condition(
        relevant, set(group_parts)
    )
    grouped: Dict[str, Dict[Tuple[str, str, str, str], Set[str]]] = defaultdict(
        lambda: defaultdict(set)
    )

    for item in report.circuits:
        if not circuit_belongs_to_connector(report, item, connector):
            continue
        applies = item.applicable_parts & relevant
        if applies:
            key = ("WIRE", item.circuit, item.gauge, item.color)
            grouped[normalize_cavity(item.cavity)][key].update(applies)

    for item in report.plugs:
        if item.connector_name != connector.name:
            continue
        applies = item.applicable_parts & relevant
        if applies:
            kind = "BLOCKED" if is_blocked_cavity(item.part_number) else "PLUG"
            key = (kind, item.part_number, "", "")
            grouped[normalize_cavity(item.cavity)][key].update(applies)

    result: Dict[str, Dict[str, object]] = {}
    variations: List[str] = []
    for cavity, option_map in grouped.items():
        merged_entries: List[Dict[str, object]] = []
        wire_entries: Dict[Tuple[str, str], Dict[str, object]] = {}
        for (kind, name, gauge, color), applies in option_map.items():
            if kind != "WIRE" or not str(gauge).strip() or not str(color).strip():
                merged_entries.append(
                    {
                        "kind": kind,
                        "names": [name],
                        "gauge": gauge,
                        "color": color,
                        "applies": set(applies),
                    }
                )
                continue
            signature = (
                normalized_gauge_signature(gauge),
                normalized_wire_color_signature(color),
            )
            entry = wire_entries.get(signature)
            if entry is None:
                entry = {
                    "kind": kind,
                    "names": [],
                    "gauge": gauge,
                    "color": color,
                    "applies": set(),
                }
                wire_entries[signature] = entry
                merged_entries.append(entry)
            names = entry["names"]
            if isinstance(names, list) and name not in names:
                names.append(name)
            entry_applies = entry["applies"]
            if isinstance(entry_applies, set):
                entry_applies.update(applies)

        option_rows: List[Dict[str, object]] = []
        for entry in merged_entries:
            kind = str(entry["kind"])
            names = entry["names"] if isinstance(entry["names"], list) else []
            name = (
                compact_circuit_names(names)
                if kind == "WIRE"
                else str(names[0]) if names else ""
            )
            gauge = str(entry["gauge"])
            color = str(entry["color"])
            applies = entry["applies"] if isinstance(entry["applies"], set) else set()
            note_type, suffixes = application_condition(applies, relevant)
            # Cuando el circuito aplica a todas las versiones en las que existe
            # el conector, hereda la condición global del conector. Sin esto la
            # leyenda exterior mostraba el color, pero la cavidad quedaba sin punto.
            if kind == "WIRE" and not suffixes and connector_suffixes:
                note_type = connector_note_type
                suffixes = list(connector_suffixes)
            option_rows.append(
                {
                    "KIND": kind,
                    "CIRCUITO": name if kind == "WIRE" else "",
                    "CALIBRE": gauge if kind == "WIRE" else "",
                    "COLOR": color if kind == "WIRE" else "",
                    "PLUG": name if kind == "PLUG" else "",
                    "BLOCKED": kind == "BLOCKED",
                    "TIPO": note_type,
                    "OT": ",".join(suffixes),
                    "APLICA": suffixes,
                    "APLICA_COUNT": len(applies),
                    "CIRCUITOS_UNIDOS": list(names) if kind == "WIRE" else [],
                }
            )

        # Regla de prioridad: cualquier cable queda dentro del conector. El
        # plug pasa a varianza; solo queda principal cuando no existe cable.
        option_rows.sort(
            key=lambda option: {"WIRE": 0, "BLOCKED": 1, "PLUG": 2}.get(
                str(option["KIND"]), 9
            )
        )
        # Entre varios circuitos se conserva el primero encontrado en el RI.
        primary = dict(option_rows[0])
        if primary["KIND"] == "PLUG":
            primary["CIRCUITO"] = str(primary["PLUG"])
            primary["CALIBRE"] = ""
            primary["COLOR"] = ""
        primary["VARIANTES"] = option_rows if len(option_rows) > 1 else []
        result[cavity] = primary

        if len(option_rows) > 1:
            details = []
            for option in option_rows:
                if option["KIND"] == "BLOCKED":
                    label = "BLOCKED CAVITY"
                elif option["KIND"] == "PLUG":
                    label = f"PLUG:{option['PLUG']}"
                else:
                    label = (
                        f"{option['CIRCUITO']} / {option['CALIBRE']} / "
                        f"{option['COLOR']}"
                    )
                condition = (
                    f" {option['TIPO']}:{option['OT']}" if option["TIPO"] else ""
                )
                details.append(label + condition)
            variations.append(
                f"{connector.name} cavidad {cavity}: " + " | ".join(details)
            )
    return result, variations


def application_notes(
    report: ReportData,
    connector: ConnectorRecord,
    group_parts: Iterable[str],
) -> List[Tuple[str, List[str]]]:
    """Resume las configuraciones aplicables como SOLO/N/A, sin duplicados."""
    group_set = set(group_parts)
    relevant = connector.applicable_parts & group_set
    notes: Set[Tuple[str, Tuple[str, ...]]] = set()

    def add_scope(scope: Set[str], universe: Set[str]) -> None:
        note_type, values = application_condition(scope, universe)
        suffixes = tuple(values)
        if suffixes:
            notes.add((note_type, suffixes))

    # Se conserva la lista completa para analizar plugs y accesorios. La paleta
    # visual se asigna después, únicamente sobre las condiciones de circuitos
    # que realmente se muestran.
    add_scope(relevant, group_set)
    for item in report.circuits:
        if not circuit_belongs_to_connector(report, item, connector):
            continue
        applies = item.applicable_parts & relevant
        if applies:
            add_scope(applies, relevant)
    for item in report.plugs:
        if item.connector_name != connector.name:
            continue
        applies = item.applicable_parts & relevant
        if applies:
            add_scope(applies, relevant)
    for item in report.accessories:
        if item.connector_name != connector.name:
            continue
        applies = item.applicable_parts & relevant
        if applies:
            add_scope(applies, relevant)

    order = {"SOLO": 0, "N/A": 1}
    sorted_notes = sorted(
        notes,
        key=lambda item: (
            order.get(item[0], 9),
            tuple(
                f"{int(value):09d}" if value.isdigit() else value
                for value in item[1]
            ),
        ),
    )
    return [(note_type, list(values)) for note_type, values in sorted_notes]


def format_application_notes(notes: Sequence[Tuple[str, Sequence[str]]]) -> str:
    return " | ".join(
        f"{note_type}:{','.join(values)}" for note_type, values in notes
    )


def application_marker_colors(
    notes: Sequence[Tuple[str, Sequence[str]]],
) -> Dict[Tuple[str, Tuple[str, ...]], int]:
    """Asigna a cada condición el mismo ACI usado en la leyenda exterior."""
    return {
        (str(note_type).upper(), tuple(str(value) for value in values)):
        APPLICATION_NOTE_ACI[index % len(APPLICATION_NOTE_ACI)]
        for index, (note_type, values) in enumerate(notes)
    }


def visible_circuit_notes(
    notes: Sequence[Tuple[str, Sequence[str]]],
    cavity_data: Dict[str, Dict[str, object]],
) -> List[Tuple[str, List[str]]]:
    """Filtra la leyenda a condiciones usadas por circuitos WIRE visibles."""
    used: Set[Tuple[str, Tuple[str, ...]]] = set()
    for primary in cavity_data.values():
        sources: List[Dict[str, object]] = [primary]
        variations = primary.get("VARIANTES", [])
        if isinstance(variations, list):
            sources.extend(item for item in variations if isinstance(item, dict))
        for source in sources:
            if str(source.get("KIND", "")).upper() != "WIRE":
                continue
            note_type = str(source.get("TIPO", "")).upper()
            values = tuple(
                value for value in str(source.get("OT", "")).split(",") if value
            )
            if note_type and values:
                used.add((note_type, values))
    return [
        (str(note_type).upper(), [str(value) for value in values])
        for note_type, values in notes
        if (str(note_type).upper(), tuple(str(value) for value in values)) in used
    ]


def complete_application_notes(
    notes: Sequence[Tuple[str, Sequence[str]]],
    cavity_data: Dict[str, Dict[str, object]],
) -> List[Tuple[str, List[str]]]:
    """Añade condiciones nuevas producidas al unir aplicaciones de circuitos."""
    completed = [
        (str(note_type).upper(), [str(value) for value in values])
        for note_type, values in notes
    ]
    existing = {
        (note_type, tuple(values)) for note_type, values in completed
    }
    for primary in cavity_data.values():
        sources: List[Dict[str, object]] = [primary]
        variations = primary.get("VARIANTES", [])
        if isinstance(variations, list):
            sources.extend(item for item in variations if isinstance(item, dict))
        for source in sources:
            if str(source.get("KIND", "")).upper() != "WIRE":
                continue
            key = (
                str(source.get("TIPO", "")).upper(),
                tuple(
                    value
                    for value in str(source.get("OT", "")).split(",")
                    if value
                ),
            )
            if key[0] and key[1] and key not in existing:
                completed.append((key[0], list(key[1])))
                existing.add(key)
    return completed


def cavity_variation_items(
    cavity_data: Dict[str, Dict[str, object]],
    embedded_option_counts: Optional[Dict[str, int]] = None,
) -> List[Dict[str, object]]:
    """Devuelve alternativas que no fueron colocadas en cuadros repetidos."""
    embedded_option_counts = embedded_option_counts or {}
    rows: List[Dict[str, object]] = []
    for cavity in sorted(
        cavity_data,
        key=lambda value: (0, int(value)) if value.isdigit() else (1, value),
    ):
        options = cavity_data[cavity].get("VARIANTES", [])
        if not isinstance(options, list):
            continue
        # La primera opción es principal; CAVR puede alojar opciones adicionales.
        start_index = 1 + max(0, int(embedded_option_counts.get(cavity, 0)))
        for option in options[start_index:]:
            if not isinstance(option, dict):
                continue
            # Las alternativas PLUG ya aparecen en la información del conector
            # y no necesitan cuadros inferiores. Solo se dibujan cavidades extra
            # que contienen la información completa de un circuito.
            if str(option.get("KIND", "")).upper() != "WIRE":
                continue
            if not all(
                str(option.get(field, "")).strip()
                for field in ("CIRCUITO", "CALIBRE", "COLOR")
            ):
                continue
            row = dict(option)
            row["CAV"] = cavity
            rows.append(row)
    return rows


def top_application_notes(
    visible_notes: Sequence[Tuple[str, Sequence[str]]],
    cavity_data: Dict[str, Dict[str, object]],
    embedded_option_counts: Optional[Dict[str, int]] = None,
) -> List[Tuple[str, List[str]]]:
    """Evita repetir arriba una condición ya mostrada con cavidades inferiores."""
    lower_keys: Set[Tuple[str, Tuple[str, ...]]] = set()
    for option in cavity_variation_items(
        cavity_data, embedded_option_counts=embedded_option_counts
    ):
        note_type = str(option.get("TIPO", "")).upper()
        values = tuple(
            value for value in str(option.get("OT", "")).split(",") if value
        )
        if note_type and values:
            lower_keys.add((note_type, values))

    return [
        (str(note_type).upper(), [str(value) for value in values])
        for note_type, values in visible_notes
        if (str(note_type).upper(), tuple(str(value) for value in values))
        not in lower_keys
    ]


def normalized_file_number(path: Path) -> str:
    stem = re.sub(r"\s*\(\d+\)\s*$", "", path.stem)
    return re.sub(r"[^A-Z0-9]", "", normalize_part_number(stem))


def match_connector_files(
    files: Iterable[Tuple[Path, float]], required_components: Iterable[str]
) -> Tuple[Dict[str, List[Path]], List[Path]]:
    required = set(required_components)
    required_by_key: Dict[str, List[str]] = defaultdict(list)
    for component in required:
        required_by_key[re.sub(r"[^A-Z0-9]", "", normalize_part_number(component))].append(component)
    candidates: Dict[str, List[Tuple[Path, float]]] = defaultdict(list)
    extras: List[Path] = []
    for path, modified in files:
        if path.suffix.lower() not in {".dwg", ".dxf"}:
            continue
        number = normalized_file_number(path)
        if number in required_by_key:
            for component in required_by_key[number]:
                candidates[component].append((path, modified))
        else:
            extras.append(path)
    matches = {}
    for component, drawings in candidates.items():
        # Se prioriza el DWG y, entre archivos del mismo formato, el más nuevo.
        drawings.sort(key=lambda item: (item[0].suffix.lower() != ".dwg", -item[1]))
        matches[component] = [path for path, _ in drawings]
    return matches, extras


def scan_connector_folder(
    folder: Path, required_components: Iterable[str]
) -> Tuple[Dict[str, List[Path]], List[Path]]:
    files = (
        (path, path.stat().st_mtime)
        for path in Path(folder).rglob("*")
        if path.suffix.lower() in {".dwg", ".dxf"} and path.is_file()
    )
    return match_connector_files(files, required_components)


def build_database_json(
    family: str,
    report: ReportData,
    assignments: Dict[str, str],
    groups: Dict[str, List[str]],
    requirements: Dict[str, GroupRequirement],
) -> Dict[str, object]:
    colors: Dict[str, List[str]] = defaultdict(list)
    for part in report.part_numbers:
        color = normalize_color(assignments.get(part, ""))
        if color:
            colors[color.lower()].append(part)
    return {
        family: dict(colors),
        "_plantillas": {
            family: {
                name: [color.lower() for color in group_colors]
                for name, group_colors in groups.items()
            }
        },
        "_conteo_conectores": {
            family: {
                name: {
                    "cantidad_posiciones": len(requirement.connectors),
                    "componentes_unicos": len(requirement.unique_components),
                    "numeros_de_parte": requirement.part_numbers,
                    "conectores": [
                        {
                            "nombre": item.name,
                            "numero_de_parte": item.component,
                            "tipo": item.kind,
                            "plugs": connector_plug_parts(
                                report, item, requirement.part_numbers
                            ),
                            "accesorios": [
                                {
                                    "tipo": accessory.label,
                                    "numero_de_parte": accessory.part_number,
                                    "cantidad": accessory.quantity,
                                }
                                for accessory in connector_accessory_records(
                                    report, item, requirement.part_numbers
                                )
                            ],
                            "aplicaciones": [
                                {
                                    "condicion": note_type,
                                    "numeros": values,
                                }
                                for note_type, values in application_notes(
                                    report, item, requirement.part_numbers
                                )
                            ],
                        }
                        for item in requirement.connectors
                    ],
                }
                for name, requirement in requirements.items()
            }
        },
        "_origen": {
            "archivo": report.source_path.name,
            "reporte": report.report_label,
            "generado": datetime.now().isoformat(timespec="seconds"),
            "programa": f"{APP_NAME} {APP_VERSION}",
        },
    }


class AutoCADAssembler:
    """Ensamblador COM. Se ejecuta únicamente en Windows con AutoCAD."""

    def __init__(self, progress: Optional[Callable[[str], None]] = None):
        self.progress = progress or (lambda _message: None)
        self.text_style_name = "ARIAL_REGULAR"
        # AutoCAD puede crear un texto y después lanzar una excepción al
        # ejecutar Update(). Este registro impide que un reintento duplique
        # textos que ya existen físicamente en el dibujo.
        self._created_text_keys: Set[Tuple[object, ...]] = set()
        # Registro independiente de los puntos físicos creados. Los atributos
        # TIPO/OT no prueban que el círculo de color exista en el bloque.
        self._created_marker_keys: Set[Tuple[object, ...]] = set()
        try:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore
            from win32com.client import VARIANT  # type: ignore
        except ImportError as error:
            raise RuntimeError(
                "Falta pywin32. Instálalo con: pip install pywin32"
            ) from error
        self.pythoncom = pythoncom
        self.win32 = win32com.client
        self.VARIANT = VARIANT

    def _ensure_arial_regular(self, document) -> None:
        """Crea y activa un estilo Arial sin negrita ni cursiva."""
        styles = self._com_retry(
            lambda: document.TextStyles,
            "No se pudieron abrir los estilos de texto",
            attempts=6,
            base_delay=0.10,
        )
        try:
            style = self._com_retry(
                lambda: styles.Item(self.text_style_name),
                "No se pudo abrir el estilo Arial",
                attempts=3,
                base_delay=0.08,
            )
        except Exception:
            style = self._com_retry(
                lambda: styles.Add(self.text_style_name),
                "No se pudo crear el estilo Arial",
                attempts=6,
                base_delay=0.10,
            )

        def configure() -> None:
            try:
                style.SetFont("Arial", False, False, 0, 34)
            except Exception:
                style.FontFile = "arial.ttf"
            try:
                style.BigFontFile = ""
            except Exception:
                pass
            document.ActiveTextStyle = style

        self._com_retry(
            configure,
            "AutoCAD no pudo activar Arial Regular",
            attempts=6,
            base_delay=0.10,
        )

    def _apply_text_style(self, entity) -> bool:
        """Aplica Arial Regular a textos existentes, incluidos atributos."""
        try:
            self._com_retry(
                lambda: setattr(entity, "StyleName", self.text_style_name),
                "AutoCAD no pudo aplicar Arial Regular",
                attempts=5,
                base_delay=0.08,
            )
            return True
        except Exception:
            return False

    def point(self, x: float, y: float, z: float = 0.0):
        return self.VARIANT(
            self.pythoncom.VT_ARRAY | self.pythoncom.VT_R8,
            (float(x), float(y), float(z)),
        )

    def _com_retry(
        self,
        action: Callable[[], object],
        description: str,
        attempts: int = 8,
        base_delay: float = 0.20,
    ):
        """Reintenta llamadas COM rechazadas mientras AutoCAD está ocupado."""
        last_error: Optional[Exception] = None
        for attempt in range(attempts):
            try:
                return action()
            except Exception as error:
                last_error = error
                if attempt + 1 >= attempts:
                    break
                try:
                    self.pythoncom.PumpWaitingMessages()
                except Exception:
                    pass
                time.sleep(min(1.25, base_delay * (attempt + 1)))
        raise RuntimeError(f"{description}: {last_error}") from last_error

    def _yield_to_autocad(self, delay: float = 0.0) -> None:
        """Entrega mensajes pendientes a AutoCAD y aplica una pausa corta."""
        try:
            self.pythoncom.PumpWaitingMessages()
        except Exception:
            pass
        if delay > 0.0:
            time.sleep(delay)

    def _run_graphic_passes(
        self,
        items: Sequence[object],
        draw: Callable[[object], bool],
        label: str,
        passes: int = 6,
        batch_size: int = 1,
    ) -> Tuple[int, List[object]]:
        """Dibuja elementos y reintenta solamente los que hayan fallado.

        Las funciones gráficas deben limpiar sus entidades parciales cuando
        devuelven False. Así una segunda pasada nunca duplica un elemento que
        ya fue creado correctamente.
        """
        pending = list(items)
        completed = 0
        total_passes = max(1, int(passes))
        pace = max(1, int(batch_size))

        for pass_index in range(total_passes):
            if not pending:
                break
            if pass_index > 0:
                self.progress(
                    f"Reintentando {len(pending)} {label} "
                    f"(pasada {pass_index + 1}/{total_passes})..."
                )
                self._yield_to_autocad(0.75 * pass_index)

            failed: List[object] = []
            for item_index, item in enumerate(pending, start=1):
                try:
                    created = bool(draw(item))
                except Exception:
                    created = False
                if created:
                    completed += 1
                else:
                    failed.append(item)

                # Y960A puede superar 400 entidades. Esta pausa evita que la
                # cola COM quede ocupada y rechace todos los elementos restantes.
                if item_index % pace == 0:
                    self._yield_to_autocad(0.12 if created else 0.24)
            pending = failed

        return completed, pending

    def _delete_graphic_entities(self, entities: Sequence[object]) -> None:
        """Elimina de forma segura una creación parcial antes del reintento."""
        for entity in reversed(list(entities)):
            try:
                self._com_retry(
                    lambda target=entity: target.Delete(),
                    "No se pudo limpiar una entidad gráfica parcial",
                    attempts=4,
                    base_delay=0.08,
                )
            except Exception:
                pass
        self._yield_to_autocad(0.03)

    def _get_property(
        self, entity, name: str, default: object = "", attempts: int = 4
    ):
        try:
            return self._com_retry(
                lambda: getattr(entity, name),
                f"No se pudo leer {name}",
                attempts=attempts,
                base_delay=0.10,
            )
        except Exception:
            return default

    def _object_name(self, entity) -> str:
        try:
            return str(self._get_property(entity, "ObjectName"))
        except Exception:
            return ""

    def _block_name(self, entity) -> str:
        for attribute in ("EffectiveName", "Name"):
            value = self._get_property(entity, attribute)
            if value:
                return str(value)
        return ""

    def _attributes(self, block_reference) -> List[object]:
        try:
            attributes = self._com_retry(
                lambda: block_reference.GetAttributes(),
                "No se pudieron leer los atributos del bloque",
                attempts=6,
                base_delay=0.12,
            )
            if attributes is None:
                return []
            if isinstance(attributes, tuple):
                return list(attributes)
            return list(attributes)
        except Exception:
            return []

    def _set_attribute(self, attribute, value: str) -> bool:
        def update() -> None:
            try:
                attribute.StyleName = self.text_style_name
            except Exception:
                pass
            attribute.TextString = str(value)
            attribute.Update()

        try:
            self._com_retry(
                update,
                "AutoCAD no aceptó la actualización del atributo",
                attempts=8,
                base_delay=0.12,
            )
            return True
        except Exception:
            current = str(self._get_property(attribute, "TextString", ""))
            return current == str(value)

    def _set_multiline_attribute(
        self, attribute, lines: Sequence[str]
    ) -> bool:
        """Actualiza un ATTRIB que ya usa MTEXT sin cambiar su tipo."""
        clean_lines = [str(value) for value in lines if str(value)]
        mtext_content = "\\P".join(clean_lines)
        if not mtext_content:
            return False

        def update_mtext() -> None:
            try:
                attribute.StyleName = self.text_style_name
            except Exception:
                pass
            attribute.MTextAttributeContent = mtext_content
            attribute.UpdateMTextAttribute()
            attribute.Update()

        try:
            self._com_retry(
                update_mtext,
                "AutoCAD no aceptó el número de plug en dos líneas",
                attempts=8,
                base_delay=0.12,
            )
            current = str(
                self._get_property(attribute, "MTextAttributeContent", "")
            )
            return all(value in current for value in clean_lines)
        except Exception:
            # UpdateMTextAttribute puede guardar el contenido y aun así lanzar
            # una excepción COM. Se comprueba el valor real antes de crear un
            # MTEXT de respaldo.
            current = str(
                self._get_property(attribute, "MTextAttributeContent", "")
            )
            return bool(current) and all(value in current for value in clean_lines)

    @staticmethod
    def _normalized_cad_text(value: object) -> str:
        """Normaliza saltos y formato básico para comparar texto visible."""
        text = str(value or "").upper()
        text = text.replace("\\P", "|").replace("\r", "|").replace("\n", "|")
        text = re.sub(r"\{\\[^;{}]+;", "", text)
        text = text.replace("{", "").replace("}", "")
        return re.sub(r"\s+", "", text)

    def _attribute_contains_lines(
        self, attribute, lines: Sequence[str]
    ) -> bool:
        expected = [
            self._normalized_cad_text(value)
            for value in lines
            if str(value)
        ]
        if not expected:
            return False
        candidates = (
            self._get_property(attribute, "MTextAttributeContent", ""),
            self._get_property(attribute, "TextString", ""),
        )
        for candidate in candidates:
            current = self._normalized_cad_text(candidate)
            if current and all(value in current for value in expected):
                return True
        return False

    def _block_contains_plug_text(
        self,
        block,
        insertion: Sequence[float],
        lines: Sequence[str],
        tolerance: float = 5.0,
    ) -> bool:
        """Confirma que un MTEXT de respaldo existe junto a la cavidad."""
        expected = [
            self._normalized_cad_text(value)
            for value in lines
            if str(value)
        ]
        if not expected:
            return False
        try:
            count = int(self._get_property(block, "Count", 0, attempts=6) or 0)
        except (TypeError, ValueError):
            return False
        target_x = float(insertion[0])
        target_y = float(insertion[1])
        for index in range(count - 1, -1, -1):
            try:
                entity = self._com_retry(
                    lambda item_index=index: block.Item(item_index),
                    "No se pudo verificar un texto de plug",
                    attempts=3,
                    base_delay=0.05,
                )
            except Exception:
                continue
            if "MText" not in self._object_name(entity):
                continue
            content = self._normalized_cad_text(
                self._get_property(entity, "TextString", "")
            )
            if not content or not all(value in content for value in expected):
                continue
            point = self._get_property(entity, "InsertionPoint", None)
            if point is None:
                continue
            try:
                if (
                    abs(float(point[0]) - target_x) <= tolerance
                    and abs(float(point[1]) - target_y) <= tolerance
                ):
                    return True
            except (TypeError, ValueError, IndexError):
                continue
        return False

    def _add_plug_text_fallback(
        self,
        block,
        circuit_attribute,
        lines: Sequence[str],
    ) -> bool:
        """Escribe el plug como MTEXT cuando AutoCAD rechaza editar el ATTRIB."""
        clean_lines = [str(value) for value in lines if str(value)]
        if not clean_lines:
            return False
        insertion = self._get_property(
            circuit_attribute, "TextAlignmentPoint", None
        )
        if insertion is None:
            insertion = self._get_property(
                circuit_attribute, "InsertionPoint", None
            )
        if insertion is None:
            return False
        height = float(self._get_property(circuit_attribute, "Height", 1.5))
        rotation = float(self._get_property(circuit_attribute, "Rotation", 0.0))
        width = float(
            self._get_property(circuit_attribute, "MTextWidth", 6.0)
        )
        if width <= 0.0:
            width = max(6.0, height * 4.0)
        content = "\\P".join(clean_lines)
        block_name = str(self._get_property(block, "Name", f"ID:{id(block)}"))
        text_key = (
            "PLUG_FALLBACK",
            block_name,
            round(float(insertion[0]), 4),
            round(float(insertion[1]), 4),
            content,
        )
        if text_key in self._created_text_keys:
            return True

        label = None
        # AddMText puede crear la entidad y aun así devolver una excepción COM.
        # Se llama una sola vez por intento y después se busca físicamente el
        # texto antes de repetir; así nunca se crean copias superpuestas.
        for attempt in range(10):
            try:
                label = block.AddMText(
                    self.point(
                        float(insertion[0]),
                        float(insertion[1]),
                        float(insertion[2]) if len(insertion) > 2 else 0.0,
                    ),
                    width,
                    content,
                )
            except Exception:
                label = None
                self._yield_to_autocad(0.18 + attempt * 0.10)
                if self._block_contains_plug_text(block, insertion, clean_lines):
                    self._created_text_keys.add(text_key)
                    return True
                continue
            break

        if label is None:
            return False

        # AddMText ya creó la entidad con el contenido. Se registra antes de
        # configurar su estilo porque Update() puede lanzar una excepción aun
        # cuando el texto quedó guardado correctamente.
        self._created_text_keys.add(text_key)

        def configure() -> None:
            label.Height = height
            label.Rotation = rotation
            label.AttachmentPoint = 5
            label.Color = 7
            try:
                label.StyleName = self.text_style_name
            except Exception:
                pass
            label.Update()

        try:
            self._com_retry(
                configure,
                "AutoCAD no pudo terminar de configurar el plug",
                attempts=8,
                base_delay=0.12,
            )
        except Exception:
            # La entidad existe y contiene el número. No se repite ni se marca
            # como faltante solamente porque Update() respondió tarde.
            pass
        self._yield_to_autocad(0.05)
        return self._block_contains_plug_text(
            block, insertion, clean_lines
        ) or text_key in self._created_text_keys

    def _write_plug_verified(
        self,
        block,
        circuit_attribute,
        lines: Sequence[str],
    ) -> bool:
        """Inserta un plug y comprueba su existencia antes de continuar."""
        clean_lines = [str(value) for value in lines if str(value)]
        if not clean_lines:
            return False
        if self._attribute_contains_lines(circuit_attribute, clean_lines):
            return True

        # El MTEXT independiente es más estable que editar un ATTRIB-MTEXT
        # repetidamente dentro de bloques grandes como Y900A_2.
        if self._add_plug_text_fallback(
            block, circuit_attribute, clean_lines
        ):
            return True

        # Última ruta: intentar el atributo original y volver a validarlo.
        if self._set_multiline_attribute(circuit_attribute, clean_lines):
            return True
        self._yield_to_autocad(0.35)
        return self._attribute_contains_lines(circuit_attribute, clean_lines)

    def _apply_wire_color(self, entity, color_code: str) -> bool:
        """Aplica TrueColor RGB; conserva ACI únicamente como respaldo."""
        code = str(color_code).strip().upper()
        red, green, blue = WIRE_RGB.get(code, (255, 255, 255))

        def apply_true_color() -> None:
            color = entity.TrueColor
            color.SetRGB(int(red), int(green), int(blue))
            entity.TrueColor = color
            entity.Update()

        try:
            self._com_retry(
                apply_true_color,
                f"AutoCAD no pudo aplicar TrueColor {code}",
                attempts=5,
                base_delay=0.10,
            )
            return True
        except Exception:
            try:
                color_index = int(WIRE_ACI.get(code, 7))

                def apply_aci() -> None:
                    entity.Color = color_index
                    entity.Update()

                self._com_retry(
                    apply_aci,
                    f"AutoCAD no pudo aplicar el color ACI {code}",
                    attempts=5,
                    base_delay=0.10,
                )
                return True
            except Exception:
                return False

    def _apply_rgb_color(
        self,
        entity,
        rgb: Tuple[int, int, int],
        fallback_aci: int,
    ) -> bool:
        """Aplica un RGB explícito y usa ACI solamente como respaldo."""
        red, green, blue = rgb

        def apply_true_color() -> None:
            color = entity.TrueColor
            color.SetRGB(int(red), int(green), int(blue))
            entity.TrueColor = color
            entity.Update()

        try:
            self._com_retry(
                apply_true_color,
                "AutoCAD no pudo aplicar el color RGB",
                attempts=5,
                base_delay=0.08,
            )
            return True
        except Exception:
            try:
                entity.Color = int(fallback_aci)
                entity.Update()
                return True
            except Exception:
                return False

    def _add_color_fill(self, block, cavity_reference, wire_color: str) -> int:
        """Rellena el cuadro inferior: base sólida y franja diagonal secundaria."""
        codes = [part.strip().upper() for part in wire_color.split("/") if part.strip()]
        if not codes:
            return 0
        try:
            origin = self._get_property(cavity_reference, "InsertionPoint", None)
            if origin is None:
                return 0
            rotation = float(self._get_property(cavity_reference, "Rotation", 0.0))
            scale_x = float(
                self._get_property(cavity_reference, "XScaleFactor", 1.0)
            )
            scale_y = float(
                self._get_property(cavity_reference, "YScaleFactor", 1.0)
            )
        except Exception:
            return 0

        def transform(local_x: float, local_y: float) -> Tuple[float, float]:
            lx = local_x * scale_x
            ly = local_y * scale_y
            return (
                float(origin[0]) + lx * math.cos(rotation) - ly * math.sin(rotation),
                float(origin[1]) + lx * math.sin(rotation) + ly * math.cos(rotation),
            )

        created_solids: List[object] = []

        def triangle(points: Sequence[Tuple[float, float]], color_code: str) -> bool:
            transformed = [transform(x, y) for x, y in points]
            solid = None
            try:
                solid = self._com_retry(
                    lambda: block.AddSolid(
                        self.point(*transformed[0]),
                        self.point(*transformed[1]),
                        self.point(*transformed[2]),
                        self.point(*transformed[2]),
                    ),
                    "AutoCAD no pudo crear el sólido de color",
                    attempts=8,
                    base_delay=0.12,
                )
                if not self._apply_wire_color(solid, color_code):
                    raise RuntimeError(
                        f"AutoCAD no pudo aplicar el color {color_code} al sólido"
                    )
                created_solids.append(solid)
                return True
            except Exception:
                try:
                    solid.Delete()
                except Exception:
                    pass
                return False

        added = 0
        # Interior del cuadro 10 x 10 de la cavidad, debajo de los textos.
        base_triangles = (
            ((0.35, 0.35), (9.65, 0.35), (0.35, 3.85)),
            ((9.65, 0.35), (9.65, 3.85), (0.35, 3.85)),
        )
        for points in base_triangles:
            added += int(triangle(points, codes[0]))

        if len(codes) > 1:
            # Banda ancha ascendente como la referencia del conector analizado.
            band_triangles = (
                ((0.35, 0.45), (0.35, 1.55), (9.65, 2.65)),
                ((0.35, 1.55), (9.65, 3.75), (9.65, 2.65)),
            )
            for points in band_triangles:
                added += int(triangle(points, codes[1]))
        expected = 4 if len(codes) > 1 else 2
        if added != expected:
            # No se reporta una cavidad como coloreada si quedó incompleta.
            for solid in created_solids:
                try:
                    self._com_retry(
                        lambda target=solid: target.Delete(),
                        "No se pudo limpiar un sólido parcial",
                        attempts=3,
                        base_delay=0.08,
                    )
                except Exception:
                    pass
            return 0
        return added

    def _clear_blocked_cavity_attribute(self, attribute) -> bool:
        """Vacía el texto simple y el MTEXT incrustado; conserva el número CAV."""
        mode = self._get_property(attribute, "MTextAttribute", None)
        if mode is None:
            return False
        multiline = bool(mode)

        def clear():
            attribute.TextString = ""
            if multiline:
                attribute.MTextAttributeContent = ""
                attribute.UpdateMTextAttribute()
            attribute.Update()

        try:
            self._com_retry(clear, "No se pudo vaciar el texto de la cavidad bloqueada",
                            attempts=7, base_delay=0.12)
        except Exception:
            pass
        plain = self._get_property(attribute, "TextString", None)
        embedded = self._get_property(attribute, "MTextAttributeContent", None) if multiline else ""
        return (plain is not None and embedded is not None
                and not self._normalized_cad_text(plain).strip("|")
                and not self._normalized_cad_text(embedded).strip("|"))

    def _blocked_cavity_local_diagonal(self, document, cavity_reference):
        """Mide los bordes reales del bloque de cavidad sin incluir atributos."""
        name = str(self._get_property(cavity_reference, "Name", ""))
        if document is None or not name:
            return blocked_cavity_diagonal([])
        cache = getattr(self, "_blocked_diagonal_cache", None)
        if cache is None:
            self._blocked_diagonal_cache = cache = {}
        key = (id(document), name)
        if key in cache:
            return cache[key]
        segments = []
        try:
            definition = document.Blocks.Item(name)
            for index in range(int(definition.Count)):
                entity = definition.Item(index)
                object_name = self._object_name(entity)
                if object_name == "AcDbLine":
                    start = self._get_property(entity, "StartPoint", None)
                    end = self._get_property(entity, "EndPoint", None)
                    if start is not None and end is not None:
                        segments.append(((float(start[0]), float(start[1])),
                                         (float(end[0]), float(end[1]))))
                elif object_name in {"AcDbPolyline", "AcDb2dPolyline"}:
                    coordinates = self._get_property(entity, "Coordinates", ())
                    stride = 2 if object_name == "AcDbPolyline" else 3
                    vertices = [
                        (float(coordinates[i]), float(coordinates[i + 1]))
                        for i in range(0, len(coordinates) - stride + 1, stride)
                    ]
                    segments.extend(zip(vertices, vertices[1:]))
                    if vertices and self._get_property(entity, "Closed", False):
                        segments.append((vertices[-1], vertices[0]))
        except Exception:
            return blocked_cavity_diagonal([])
        result = blocked_cavity_diagonal(segments)
        if segments:
            cache[key] = result
        return result

    def _add_blocked_cavity_mark(self, block, cavity_reference, document=None) -> bool:
        """Dibuja la diagonal de una cavidad bloqueada sin texto ni relleno."""
        origin = self._get_property(cavity_reference, "InsertionPoint", None)
        if origin is None:
            return False
        rotation = float(self._get_property(cavity_reference, "Rotation", 0.0))
        scale_x = float(
            self._get_property(cavity_reference, "XScaleFactor", 1.0)
        )
        scale_y = float(
            self._get_property(cavity_reference, "YScaleFactor", 1.0)
        )

        def transform(local_x: float, local_y: float) -> Tuple[float, float]:
            lx = local_x * scale_x
            ly = local_y * scale_y
            return (
                float(origin[0]) + lx * math.cos(rotation) - ly * math.sin(rotation),
                float(origin[1]) + lx * math.sin(rotation) + ly * math.cos(rotation),
            )

        local_start, local_end = self._blocked_cavity_local_diagonal(document, cavity_reference)
        start = transform(*local_start)
        end = transform(*local_end)

        def same_point(first: object, second: Tuple[float, float]) -> bool:
            try:
                return (
                    abs(float(first[0]) - second[0]) <= 0.05
                    and abs(float(first[1]) - second[1]) <= 0.05
                )
            except Exception:
                return False

        def is_expected_line(entity) -> bool:
            if "Line" not in self._object_name(entity):
                return False
            line_start = self._get_property(entity, "StartPoint", None)
            line_end = self._get_property(entity, "EndPoint", None)
            return (
                same_point(line_start, start) and same_point(line_end, end)
            ) or (
                same_point(line_start, end) and same_point(line_end, start)
            )

        def find_line():
            try:
                count = int(self._get_property(block, "Count", 0))
            except Exception:
                return None
            for index in range(count):
                try:
                    entity = block.Item(index)
                    if is_expected_line(entity):
                        return entity
                except Exception:
                    continue
            return None

        # Se comprueba la geometría antes de cada AddLine. Una excepción de
        # AutoCAD puede llegar después de crear la entidad o durante Update.
        # Repetir AddLine sin comprobarla produciría diagonales duplicadas.
        for attempt in range(7):
            line = find_line()
            if line is None:
                try:
                    line = block.AddLine(self.point(*start), self.point(*end))
                except Exception:
                    line = find_line()
            if line is not None:
                try:
                    line.Color = 7
                    line.Update()
                except Exception:
                    pass
                if is_expected_line(line):
                    return True
            if find_line() is not None:
                return True
            if attempt < 6:
                self._yield_to_autocad(0.12 * (attempt + 1))
        return False

    def _add_application_marker(
        self,
        block,
        cavity_reference,
        marker_color: int,
        marker_slot: int = 0,
        marker_count: int = 1,
    ) -> int:
        """Dibuja un punto sólido dentro de la cavidad usando el color de la leyenda."""
        origin = self._get_property(cavity_reference, "InsertionPoint", None)
        if origin is None:
            return 0
        block_name = str(self._get_property(block, "Name", f"ID:{id(block)}"))
        marker_key = (
            "APPLICATION_MARKER",
            block_name,
            round(float(origin[0]), 4),
            round(float(origin[1]), 4),
            int(marker_color),
            int(marker_slot),
        )
        if marker_key in self._created_marker_keys:
            return 10
        rotation = float(self._get_property(cavity_reference, "Rotation", 0.0))
        scale_x = float(
            self._get_property(cavity_reference, "XScaleFactor", 1.0)
        )
        scale_y = float(
            self._get_property(cavity_reference, "YScaleFactor", 1.0)
        )

        def transform(local_x: float, local_y: float) -> Tuple[float, float]:
            lx = local_x * scale_x
            ly = local_y * scale_y
            return (
                float(origin[0]) + lx * math.cos(rotation) - ly * math.sin(rotation),
                float(origin[1]) + lx * math.sin(rotation) + ly * math.cos(rotation),
            )

        # Una cavidad puede tener varias condiciones (la principal y plugs o
        # circuitos alternativos). Se distribuyen sin cubrirse entre sí.
        center = (
            8.25 - 1.15 * (int(marker_slot) % 3),
            6.15 - 1.15 * (int(marker_slot) // 3),
        )
        radius = 0.55 if int(marker_count) <= 1 else 0.42
        created: List[object] = []
        segments = 10
        try:
            for index in range(segments):
                angle_a = 2.0 * math.pi * index / segments
                angle_b = 2.0 * math.pi * (index + 1) / segments
                local_points = (
                    center,
                    (
                        center[0] + radius * math.cos(angle_a),
                        center[1] + radius * math.sin(angle_a),
                    ),
                    (
                        center[0] + radius * math.cos(angle_b),
                        center[1] + radius * math.sin(angle_b),
                    ),
                )
                points = [transform(*point) for point in local_points]
                solid = self._com_retry(
                    lambda values=points: block.AddSolid(
                        self.point(*values[0]),
                        self.point(*values[1]),
                        self.point(*values[2]),
                        self.point(*values[2]),
                    ),
                    "AutoCAD no pudo crear el punto de aplicación",
                    attempts=7,
                    base_delay=0.10,
                )
                created.append(solid)

                def apply_marker_color(target=solid) -> None:
                    target.Color = int(marker_color)
                    target.Update()

                try:
                    self._com_retry(
                        apply_marker_color,
                        "AutoCAD no pudo colorear el punto de aplicación",
                        attempts=5,
                        base_delay=0.08,
                    )
                except Exception:
                    try:
                        solid.Color = int(marker_color)
                    except Exception:
                        pass
            if len(created) == segments:
                self._created_marker_keys.add(marker_key)
                return segments
            return 0
        except Exception:
            for solid in created:
                try:
                    solid.Delete()
                except Exception:
                    pass
            return 0

    def _add_application_notes(
        self,
        model,
        panel_x: float,
        panel_y: float,
        notes: Sequence[Tuple[str, Sequence[str]]],
        part_colors: Dict[str, str],
        rows_per_column: int = 4,
        column_width: float = 160.0,
        marker_colors: Optional[Dict[Tuple[str, Tuple[str, ...]], int]] = None,
    ) -> Tuple[bool, float]:
        if not notes:
            return False, panel_y
        requests: List[object] = []
        column_y: Dict[int, float] = defaultdict(lambda: panel_y)
        occupied_bottom = panel_y
        rows_per_column = max(1, int(rows_per_column))
        for index, (note_type, values) in enumerate(notes):
            normalized_values = tuple(str(value) for value in values)
            marker_color = (
                marker_colors.get((str(note_type).upper(), normalized_values), 7)
                if marker_colors is not None
                else APPLICATION_NOTE_ACI[index % len(APPLICATION_NOTE_ACI)]
            )
            column = index // rows_per_column
            request_y = column_y[column]
            requests.append(
                (
                    panel_x + column * column_width,
                    request_y,
                    str(note_type).upper(),
                    normalized_values,
                    marker_color,
                )
            )
            line_count = max(
                1, math.ceil(len(values) / APPLICATIONS_PER_LINE)
            )
            label_bottom = request_y + 5.0 - (
                10.5 + (line_count - 1) * 6.5
            )
            occupied_bottom = min(
                occupied_bottom, label_bottom, request_y - 3.2
            )
            column_y[column] = request_y - (
                17.0 + (line_count - 1) * 6.5 + 4.0
            )

        def draw_request(item: object) -> bool:
            request_x, request_y, note_type, values, marker_color = item
            return self._add_variation_header(
                model,
                float(request_x),
                float(request_y),
                str(note_type),
                tuple(str(value) for value in values),
                int(marker_color),
                part_colors,
            )

        _created, failed = self._run_graphic_passes(
            requests,
            draw_request,
            "etiquetas SOLO/N/A",
            passes=6,
            batch_size=1,
        )
        return not failed, occupied_bottom

    def _entity_bounds(
        self,
        entity,
        fallback_x: float,
        fallback_y: float,
    ) -> Tuple[float, float, float, float]:
        """Devuelve límites WCS confiables aun si AutoCAD responde tarde."""
        try:
            minimum, maximum = self._com_retry(
                lambda: entity.GetBoundingBox(),
                "No se pudieron calcular los límites gráficos",
                attempts=8,
                base_delay=0.10,
            )
            return (
                float(minimum[0]),
                float(minimum[1]),
                float(maximum[0]),
                float(maximum[1]),
            )
        except Exception:
            return (
                float(fallback_x),
                float(fallback_y) - 90.0,
                float(fallback_x) + 220.0,
                float(fallback_y) + 90.0,
            )

    def _align_connector_to_cell(
        self,
        block_reference,
        insertion_x: float,
        insertion_y: float,
        cell_left: float,
        row_top: float,
    ) -> Tuple[float, float, Tuple[float, float, float, float]]:
        """Alinea el borde superior/izquierdo real, no el origen del DWG."""
        try:
            block_reference.Update()
        except Exception:
            pass
        self._yield_to_autocad(0.04)
        minimum_x, _minimum_y, _maximum_x, maximum_y = self._entity_bounds(
            block_reference, insertion_x, insertion_y
        )
        target_left = float(cell_left) + 20.0
        delta_x = target_left - minimum_x
        delta_y = float(row_top) - maximum_y
        if abs(delta_x) > 0.001 or abs(delta_y) > 0.001:
            try:
                self._com_retry(
                    lambda: block_reference.Move(
                        self.point(0.0, 0.0, 0.0),
                        self.point(delta_x, delta_y, 0.0),
                    ),
                    "No se pudo alinear el conector en su celda",
                    attempts=8,
                    base_delay=0.12,
                )
                try:
                    block_reference.Update()
                except Exception:
                    pass
                self._yield_to_autocad(0.05)
            except Exception:
                delta_x = 0.0
                delta_y = 0.0
        aligned_x = insertion_x + delta_x
        aligned_y = insertion_y + delta_y
        bounds = self._entity_bounds(block_reference, aligned_x, aligned_y)
        return aligned_x, aligned_y, bounds

    def _right_application_anchor(
        self,
        block_reference,
        fallback_x: float,
        fallback_y: float,
    ) -> Tuple[float, float]:
        """Ubica SOLO/N/A a la derecha del límite gráfico del conector."""
        try:
            minimum, maximum = self._com_retry(
                lambda: block_reference.GetBoundingBox(),
                "No se pudieron calcular los límites del conector",
                attempts=6,
                base_delay=0.10,
            )
            min_y = float(minimum[1])
            max_x = float(maximum[0])
            max_y = float(maximum[1])
            height = max(1.0, max_y - min_y)
            # Cerca del tercio superior, igual que el ejemplo entregado.
            return max_x + 14.0, max_y - min(24.0, height * 0.28)
        except Exception:
            return fallback_x + 95.0, fallback_y - 78.0

    def _add_global_marker(
        self, model, center_x: float, center_y: float, marker_color: int
    ) -> bool:
        created: List[object] = []
        radius = 3.2
        segments = 10
        try:
            for index in range(segments):
                angle_a = 2.0 * math.pi * index / segments
                angle_b = 2.0 * math.pi * (index + 1) / segments
                points = (
                    (center_x, center_y),
                    (
                        center_x + radius * math.cos(angle_a),
                        center_y + radius * math.sin(angle_a),
                    ),
                    (
                        center_x + radius * math.cos(angle_b),
                        center_y + radius * math.sin(angle_b),
                    ),
                )
                solid = self._com_retry(
                    lambda values=points: model.AddSolid(
                        self.point(*values[0]),
                        self.point(*values[1]),
                        self.point(*values[2]),
                        self.point(*values[2]),
                    ),
                    "No se pudo crear el punto de la varianza",
                    attempts=9,
                    base_delay=0.10,
                )
                created.append(solid)

                def apply_color(target=solid) -> None:
                    target.Color = int(marker_color)
                    target.Update()

                try:
                    self._com_retry(
                        apply_color,
                        "No se pudo colorear el punto de la varianza",
                        attempts=5,
                        base_delay=0.08,
                    )
                except Exception:
                    # La entidad ya existe y normalmente Color quedó aplicado
                    # antes de que Update respondiera con retraso.
                    try:
                        solid.Color = int(marker_color)
                    except Exception:
                        pass
            return len(created) == segments
        except Exception:
            self._delete_graphic_entities(created)
            return False

    def _add_variation_header(
        self,
        model,
        x: float,
        y: float,
        note_type: str,
        values: Sequence[str],
        marker_color: int,
        part_colors: Dict[str, str],
    ) -> bool:
        """Crea punto y etiqueta amarilla, con cuatro aplicaciones por línea."""
        text, line_count, longest_visible = format_colored_application_mtext(
            note_type,
            values,
            part_colors,
        )
        label_x = x + 11.0
        label_width = min(142.0, max(72.0, longest_visible * 3.25 + 8.0))
        top = y + 5.0
        bottom = top - (10.5 + (line_count - 1) * 6.5)
        created: List[object] = []
        triangles = (
            ((label_x, top), (label_x + label_width, top), (label_x, bottom)),
            (
                (label_x + label_width, top),
                (label_x + label_width, bottom),
                (label_x, bottom),
            ),
        )
        try:
            for points in triangles:
                solid = self._com_retry(
                    lambda values=points: model.AddSolid(
                        self.point(*values[0]),
                        self.point(*values[1]),
                        self.point(*values[2]),
                        self.point(*values[2]),
                    ),
                    "No se pudo crear el fondo amarillo de la varianza",
                    attempts=9,
                    base_delay=0.10,
                )
                created.append(solid)
                if not self._apply_rgb_color(
                    solid, BRIGHT_YELLOW_RGB, fallback_aci=2
                ):
                    # El sólido ya existe. Se intenta conservar al menos el
                    # amarillo ACI aunque AutoCAD haya respondido tarde al RGB.
                    try:
                        solid.Color = 2
                    except Exception:
                        pass
            label = self._com_retry(
                lambda: model.AddMText(
                    self.point(label_x + 2.0, y + 3.2),
                    label_width - 4.0,
                    text,
                ),
                "No se pudo escribir la condición de la varianza",
                attempts=10,
                base_delay=0.10,
            )
            created.append(label)

            def configure_label() -> None:
                label.Height = 5.3
                label.Color = 250
                try:
                    label.StyleName = self.text_style_name
                except Exception:
                    pass
                label.Update()

            try:
                self._com_retry(
                    configure_label,
                    "AutoCAD no pudo terminar la condición de la varianza",
                    attempts=6,
                    base_delay=0.10,
                )
            except Exception:
                # AddMText ya creó el contenido; StyleName/Update no deben
                # convertir una etiqueta válida en una omisión.
                pass
            marker_ok = self._add_global_marker(model, x + 4.0, y, marker_color)
            if not marker_ok:
                raise RuntimeError("No se pudo crear el punto de la varianza")
            return True
        except Exception:
            self._delete_graphic_entities(created)
            return False

    def _draw_variant_cavity(
        self,
        model,
        x: float,
        y: float,
        option: Dict[str, object],
    ) -> bool:
        """Dibuja una cavidad alternativa con número, textos y color real."""
        created: List[object] = []

        def add_line(start: Tuple[float, float], end: Tuple[float, float]) -> None:
            line = self._com_retry(
                lambda: model.AddLine(self.point(*start), self.point(*end)),
                "No se pudo dibujar el cuadro de la varianza",
                attempts=10,
                base_delay=0.10,
            )
            created.append(line)
            try:
                line.Color = 7
                line.Update()
            except Exception:
                # AddLine ya creó el borde; un Update tardío no debe borrar
                # la cavidad completa.
                pass

        def add_text(
            text_x: float,
            text_y: float,
            width: float,
            content: str,
            height: float,
        ) -> None:
            label = self._com_retry(
                lambda: model.AddMText(
                    self.point(text_x, text_y), width, content
                ),
                "No se pudo escribir la cavidad de la varianza",
                attempts=10,
                base_delay=0.10,
            )
            created.append(label)
            try:
                label.Height = height
                label.Color = 7
                try:
                    label.StyleName = self.text_style_name
                except Exception:
                    pass
                label.Update()
            except Exception:
                # El contenido ya quedó creado por AddMText.
                pass

        def add_triangle(
            points: Sequence[Tuple[float, float]], color_code: str
        ) -> None:
            solid = self._com_retry(
                lambda: model.AddSolid(
                    self.point(*points[0]),
                    self.point(*points[1]),
                    self.point(*points[2]),
                    self.point(*points[2]),
                ),
                "No se pudo colorear la cavidad de la varianza",
                attempts=10,
                base_delay=0.10,
            )
            created.append(solid)
            if not self._apply_wire_color(solid, color_code):
                try:
                    solid.Color = int(WIRE_ACI.get(color_code, 7))
                except Exception:
                    pass

        try:
            width = 24.0
            height = 24.0
            kind = str(option.get("KIND", "WIRE")).upper()
            wire_color = str(option.get("COLOR", "")).upper()
            visible_wire_color = spanish_wire_color(wire_color)
            codes = [
                value.strip()
                for value in wire_color.split("/")
                if value.strip()
            ]
            # Los sólidos se crean primero para que líneas y textos queden encima.
            if kind == "WIRE" and codes:
                base = (
                    ((x + 0.3, y - 23.7), (x + 23.7, y - 23.7), (x + 0.3, y - 14.2)),
                    ((x + 23.7, y - 23.7), (x + 23.7, y - 14.2), (x + 0.3, y - 14.2)),
                )
                for points in base:
                    add_triangle(points, codes[0])
                if len(codes) > 1:
                    band = (
                        ((x + 0.3, y - 23.1), (x + 0.3, y - 20.2), (x + 23.7, y - 15.2)),
                        ((x + 0.3, y - 20.2), (x + 23.7, y - 14.4), (x + 23.7, y - 15.2)),
                    )
                    for points in band:
                        add_triangle(points, codes[1])

            add_line((x, y), (x + width, y))
            add_line((x + width, y), (x + width, y - height))
            add_line((x + width, y - height), (x, y - height))
            add_line((x, y - height), (x, y))
            add_line((x + 17.0, y), (x + 17.0, y - 7.0))
            add_line((x + 17.0, y - 7.0), (x + width, y - 7.0))

            cavity = str(option.get("CAV", ""))
            add_text(x + 18.0, y - 0.5, 5.5, cavity, 3.2)
            if kind == "BLOCKED":
                # Cuadro vacío con diagonal, igual que dentro del conector.
                add_line((x, y - height), (x + 17.0, y - 7.0))
            elif kind == "PLUG":
                first, second = split_plug_number(option.get("PLUG", ""))
                add_text(
                    x + 0.7,
                    y - 2.0,
                    16.0,
                    "\\P".join(value for value in (first, second) if value),
                    3.3,
                )
            else:
                gauge = str(option.get("CALIBRE", ""))
                circuit = str(option.get("CIRCUITO", ""))
                add_text(
                    x + 0.7,
                    y - 0.7,
                    16.0,
                    "\\P".join((gauge, circuit, visible_wire_color)),
                    2.8,
                )
            return True
        except Exception:
            self._delete_graphic_entities(created)
            return False

    def _add_variation_cavities(
        self,
        model,
        x: float,
        y: float,
        cavity_data: Dict[str, Dict[str, object]],
        notes: Sequence[Tuple[str, Sequence[str]]],
        part_colors: Dict[str, str],
        embedded_option_counts: Optional[Dict[str, int]] = None,
        panel_x: Optional[float] = None,
        panel_y: Optional[float] = None,
        marker_colors: Optional[
            Dict[Tuple[str, Tuple[str, ...]], int]
        ] = None,
    ) -> Tuple[int, int, int, float]:
        """Agrupa alternativas y reintenta solo los elementos fallidos."""
        items = cavity_variation_items(
            cavity_data, embedded_option_counts=embedded_option_counts
        )
        if not items:
            return 0, 0, 0, float(panel_y if panel_y is not None else y)
        if marker_colors is None:
            marker_colors = application_marker_colors(notes)
        grouped: Dict[Tuple[str, Tuple[str, ...]], List[Dict[str, object]]] = defaultdict(list)
        for item in items:
            note_type = str(item.get("TIPO", "")).upper()
            values = tuple(
                value for value in str(item.get("OT", "")).split(",") if value
            )
            grouped[(note_type, values)].append(item)

        note_order = {
            (str(note_type).upper(), tuple(str(value) for value in values)): index
            for index, (note_type, values) in enumerate(notes)
        }
        groups = sorted(
            grouped.items(),
            key=lambda item: (
                note_order.get(item[0], 999),
                item[0][0],
                item[0][1],
            ),
        )
        panel_x = float(panel_x if panel_x is not None else x - 100.0)
        panel_y = float(panel_y if panel_y is not None else y - 132.0)
        group_width = 160.0
        max_group_columns = 3
        boxes_per_row = 5
        header_requests: List[object] = []
        box_requests: List[object] = []
        occupied_bottom = panel_y

        for row_start in range(0, len(groups), max_group_columns):
            group_row = groups[row_start: row_start + max_group_columns]
            group_heights = [
                34.0
                + (
                    max(1, math.ceil(len(group_key[1]) / APPLICATIONS_PER_LINE))
                    - 1
                ) * 6.5
                + math.ceil(len(group_items) / boxes_per_row) * 28.0
                for group_key, group_items in group_row
            ]
            for column, ((note_type, values), group_items) in enumerate(group_row):
                group_x = panel_x + column * group_width
                marker_color = marker_colors.get((note_type, values), 7)
                header_requests.append(
                    (
                        group_x,
                        panel_y,
                        note_type,
                        values,
                        marker_color,
                    )
                )
                header_lines = max(
                    1, math.ceil(len(values) / APPLICATIONS_PER_LINE)
                )
                header_bottom = panel_y + 5.0 - (
                    10.5 + (header_lines - 1) * 6.5
                )
                occupied_bottom = min(occupied_bottom, header_bottom)
                header_drop = 17.0 + (header_lines - 1) * 6.5
                for index, option in enumerate(group_items):
                    box_column = index % boxes_per_row
                    box_row = index // boxes_per_row
                    box_x = group_x + 10.0 + box_column * 27.0
                    box_y = panel_y - header_drop - box_row * 28.0
                    box_requests.append((box_x, box_y, option))
                    occupied_bottom = min(occupied_bottom, box_y - 24.0)
            panel_y -= max(group_heights) + 10.0

        def draw_header(item: object) -> bool:
            header_x, header_y, note_type, values, marker_color = item
            return self._add_variation_header(
                model,
                float(header_x),
                float(header_y),
                str(note_type),
                tuple(str(value) for value in values),
                int(marker_color),
                part_colors,
            )

        def draw_box(item: object) -> bool:
            box_x, box_y, option = item
            return self._draw_variant_cavity(
                model,
                float(box_x),
                float(box_y),
                option,
            )

        _headers_created, failed_headers = self._run_graphic_passes(
            header_requests,
            draw_header,
            "encabezados de cavidades alternativas",
            passes=6,
            batch_size=1,
        )
        created_count, _failed_boxes = self._run_graphic_passes(
            box_requests,
            draw_box,
            "cavidades alternativas",
            passes=6,
            batch_size=1,
        )
        return created_count, len(items), len(failed_headers), occupied_bottom

    def _add_missing_placeholder(
        self,
        model,
        x: float,
        y: float,
        connector: ConnectorRecord,
        group_name: str,
        reason: str,
    ) -> bool:
        text = (
            f"{{\\C1;CONECTOR FALTANTE}}\\P"
            f"{connector.name}\\P"
            f"CON:{connector.component}\\P"
            f"PLANTILLA:{group_name}\\P"
            f"{reason}"
        )
        try:
            marker = self._com_retry(
                lambda: model.AddMText(
                    self.point(x - 80.0, y + 45.0),
                    180.0,
                    text,
                ),
                f"No se pudo marcar el faltante {connector.name}",
                attempts=8,
                base_delay=0.12,
            )
            marker.Height = 6.0
            marker.Color = 1
            self._apply_text_style(marker)
            marker.Update()
            return True
        except Exception:
            return False

    def _add_component_overflow(
        self,
        model,
        block_reference,
        lines: Sequence[str],
    ) -> int:
        """Agrega líneas debajo de los tres atributos disponibles del bloque."""
        clean_lines = [str(value) for value in lines if str(value)]
        if not clean_lines:
            return 0
        attributes = {
            str(self._get_property(attribute, "TagString")).strip().upper(): attribute
            for attribute in self._attributes(block_reference)
        }
        anchor = attributes.get("ACCESORIOS")
        if anchor is None:
            return 0
        insertion = self._get_property(anchor, "InsertionPoint", None)
        if insertion is None:
            return 0
        height = float(self._get_property(anchor, "Height", 4.5) or 4.5)
        rotation = float(self._get_property(anchor, "Rotation", 0.0) or 0.0)
        line_spacing = max(6.5, height * 1.55)
        created_count = 0
        reference_name = str(
            self._get_property(block_reference, "Name", f"ID:{id(block_reference)}")
        )
        for index, content in enumerate(clean_lines, start=1):
            offset = line_spacing * index
            point_x = float(insertion[0]) + offset * math.sin(rotation)
            point_y = float(insertion[1]) - offset * math.cos(rotation)
            point_z = float(insertion[2]) if len(insertion) > 2 else 0.0
            text_key = (
                "ACCESSORY_OVERFLOW",
                reference_name,
                round(point_x, 4),
                round(point_y, 4),
                content,
            )
            if text_key in self._created_text_keys:
                created_count += 1
                continue
            try:
                label = self._com_retry(
                    lambda text=content, px=point_x, py=point_y, pz=point_z: model.AddText(
                        text,
                        self.point(px, py, pz),
                        height,
                    ),
                    "No se pudo agregar una línea de accesorio",
                    attempts=10,
                    base_delay=0.12,
                )
            except Exception:
                continue

            self._created_text_keys.add(text_key)
            created_count += 1

            def configure(target=label) -> None:
                target.Rotation = rotation
                target.Color = 7
                try:
                    target.StyleName = self.text_style_name
                except Exception:
                    pass
                target.Update()

            try:
                self._com_retry(
                    configure,
                    "AutoCAD no pudo terminar una línea de accesorio",
                    attempts=8,
                    base_delay=0.12,
                )
            except Exception:
                # AddText ya devolvió una entidad válida con el contenido.
                pass
            self._yield_to_autocad(0.08)
        return created_count

    def _add_exterior_connector_label(
        self,
        model,
        block_reference,
        x: float,
        y: float,
        connector: ConnectorRecord,
        component_labels: Sequence[str],
    ) -> Tuple[bool, float]:
        """Crea una etiqueta compacta justo debajo del borde real."""
        label_lines = [
            connector.name,
            f"CON:{connector.component}",
            *(str(value) for value in component_labels if str(value)),
        ]
        content = "\\P".join(label_lines)
        minimum_x, minimum_y, _maximum_x, _maximum_y = self._entity_bounds(
            block_reference, x, y
        )
        point_x = minimum_x
        point_y = minimum_y - 10.0
        occupied_bottom = point_y - max(1, len(label_lines)) * 8.0
        text_key = (
            "EXTERIOR_CONNECTOR_LABEL",
            round(point_x, 4),
            round(point_y, 4),
            content,
        )
        if text_key in self._created_text_keys:
            return True, occupied_bottom
        try:
            label = self._com_retry(
                lambda: model.AddMText(
                    self.point(point_x, point_y),
                    180.0,
                    content,
                ),
                f"No se pudo crear la etiqueta de {connector.name}",
                attempts=12,
                base_delay=0.14,
            )
        except Exception:
            return False, occupied_bottom

        self._created_text_keys.add(text_key)

        def configure() -> None:
            label.Height = 5.0
            label.Color = 7
            try:
                label.StyleName = self.text_style_name
            except Exception:
                pass
            label.Update()

        try:
            self._com_retry(
                configure,
                f"AutoCAD no pudo terminar la etiqueta de {connector.name}",
                attempts=8,
                base_delay=0.12,
            )
        except Exception:
            # La etiqueta ya existe: no se informa como ausente por un fallo
            # posterior de StyleName o Update().
            pass
        return True, occupied_bottom

    def _fill_connector_block(
        self,
        document,
        block_reference,
        connector: ConnectorRecord,
        group_name: str,
        cavity_data: Dict[str, Dict[str, object]],
        notes: Sequence[Tuple[str, Sequence[str]]],
        component_labels: Sequence[str],
        position_index: int,
        marker_colors: Optional[
            Dict[Tuple[str, Tuple[str, ...]], int]
        ] = None,
    ) -> Dict[str, object]:
        # La información se crea siempre como una sola etiqueta exterior
        # calculada desde el borde real. Se limpian los atributos originales,
        # que en varios DWG están decenas de unidades lejos del conector.
        label_values = {
            "SALIDA": "",
            "CONECTOR": "",
            "APLICACION": "",
            "INDICE": str(position_index),
            "DESCRIPCION": "",
            "COLOR": "",
            "ACCESORIOS": "",
        }
        for attribute in self._attributes(block_reference):
            self._apply_text_style(attribute)
            tag = str(self._get_property(attribute, "TagString")).strip().upper()
            if not tag:
                continue
            if tag in label_values:
                self._set_attribute(attribute, label_values[tag])

        try:
            reference_name = str(self._get_property(block_reference, "Name"))
            definition = self._com_retry(
                lambda: document.Blocks.Item(reference_name),
                f"No se pudo abrir el bloque insertado {reference_name}",
                attempts=10,
                base_delay=0.18,
            )
        except Exception as error:
            raise RuntimeError(
                f"No se pudo abrir la definición insertada de {connector.name}"
            ) from error

        stats: Dict[str, object] = {
            "cavidades_detectadas": 0,
            "cavidades_completadas": 0,
            "cavidades_que_requieren_color": 0,
            "cavidades_con_plug": 0,
            "colores_insertados": 0,
            "puntos_aplicacion": 0,
            "puntos_requeridos": 0,
            "puntos_faltantes": [],
            "plugs_escritos": 0,
            "plugs_requeridos": 0,
            "plugs_faltantes": [],
            "cavidades_bloqueadas": 0,
            "bloqueos_dibujados": 0,
            "bloqueos_faltantes": [],
            "coincidencias_cavr": 0,
            "repeticiones_cavr_omitidas": 0,
            "opciones_distribuidas_cavr": {},
            "etiqueta_principal": False,
        }
        expected_plug_counts: Counter[Tuple[str, str]] = Counter(
            (
                normalize_cavity(cavity),
                str(data.get("PLUG", "")),
            )
            for cavity, data in cavity_data.items()
            if str(data.get("KIND", "")).upper() == "PLUG"
            and str(data.get("PLUG", ""))
        )
        encountered_plug_counts: Counter[Tuple[str, str]] = Counter()
        verified_plug_counts: Counter[Tuple[str, str]] = Counter()
        expected_blocked_counts: Counter[str] = Counter(
            normalize_cavity(cavity)
            for cavity, data in cavity_data.items()
            if str(data.get("KIND", "")).upper() == "BLOCKED"
        )
        encountered_blocked_counts: Counter[str] = Counter()
        verified_blocked_counts: Counter[str] = Counter()
        if marker_colors is None:
            marker_colors = application_marker_colors(notes)
        visited: Set[str] = set()
        cavr_occurrences: Dict[str, int] = {}

        def walk(current_definition, depth: int = 0) -> None:
            definition_name = str(
                self._get_property(current_definition, "Name")
            ).upper()
            if not definition_name:
                definition_name = f"ID:{id(current_definition)}"
            if definition_name in visited or depth > 5:
                return
            visited.add(definition_name)

            count = int(
                self._com_retry(
                    lambda: current_definition.Count,
                    f"No se pudo contar el contenido de {definition_name}",
                    attempts=8,
                    base_delay=0.12,
                )
            )
            for index in range(count):
                try:
                    entity = self._com_retry(
                        lambda item_index=index: current_definition.Item(item_index),
                        f"No se pudo leer la entidad {index} de {definition_name}",
                        attempts=6,
                        base_delay=0.10,
                    )
                except Exception:
                    continue
                object_name = self._object_name(entity)
                if any(
                    text_type in object_name
                    for text_type in ("Text", "MText", "Attribute")
                ):
                    self._apply_text_style(entity)
                if "BlockReference" not in object_name:
                    continue

                attributes = self._attributes(entity)
                tags: Dict[str, object] = {}
                for attribute in attributes:
                    self._apply_text_style(attribute)
                    tag = str(
                        self._get_property(attribute, "TagString")
                    ).strip().upper()
                    if tag:
                        tags[tag] = attribute

                if "CAV" in tags or "CAVR" in tags:
                    stats["cavidades_detectadas"] = int(stats["cavidades_detectadas"]) + 1
                    cavity, data, matched_by_cavr, option_index = (
                        select_cavity_option_for_block(
                            cavity_data,
                            self._get_property(
                                tags.get("CAV"), "TextString"
                            ) if tags.get("CAV") is not None else "",
                            self._get_property(
                                tags.get("CAVR"), "TextString"
                            ) if tags.get("CAVR") is not None else "",
                            cavr_occurrences,
                        )
                    )
                    if matched_by_cavr:
                        stats["coincidencias_cavr"] = int(
                            stats["coincidencias_cavr"]
                        ) + 1
                        if not data:
                            stats["repeticiones_cavr_omitidas"] = int(
                                stats["repeticiones_cavr_omitidas"]
                            ) + 1
                        elif option_index > 0:
                            distributed = stats["opciones_distribuidas_cavr"]
                            if isinstance(distributed, dict):
                                distributed[cavity] = max(
                                    int(distributed.get(cavity, 0)),
                                    option_index,
                                )
                    if not data:
                        continue
                    kind = str(data.get("KIND", "WIRE")).upper()
                    if kind == "BLOCKED":
                        blocked_key = normalize_cavity(cavity)
                        encountered_blocked_counts[blocked_key] += 1
                        attributes_empty = True
                        for tag in (
                            "CIRCUITO", "CALIBRE", "COLOR", "TIPO", "OT", "PLUG"
                        ):
                            if tag in tags:
                                cleared = self._clear_blocked_cavity_attribute(tags[tag])
                                attributes_empty = cleared and attributes_empty
                        diagonal_verified = self._add_blocked_cavity_mark(current_definition, entity, document)
                        if diagonal_verified and attributes_empty:
                            verified_blocked_counts[blocked_key] += 1
                            stats["cavidades_completadas"] = int(
                                stats["cavidades_completadas"]
                            ) + 1
                        continue
                    if kind == "PLUG":
                        plug_number = str(data.get("PLUG", ""))
                        first, second = split_plug_number(plug_number)
                        plug_key = (normalize_cavity(cavity), plug_number)
                        encountered_plug_counts[plug_key] += 1
                        plug_written = False
                        if "CIRCUITO" in tags:
                            plug_written = self._write_plug_verified(
                                current_definition,
                                tags["CIRCUITO"],
                                (first, second),
                            )
                        if plug_written:
                            verified_plug_counts[plug_key] += 1
                        if "CALIBRE" in tags:
                            self._set_attribute(tags["CALIBRE"], "")
                    else:
                        for tag in ("CIRCUITO", "CALIBRE"):
                            if tag in tags:
                                self._set_attribute(
                                    tags[tag], str(data.get(tag, ""))
                                )

                    note_type = str(data.get("TIPO", "")).upper()
                    note_values = tuple(
                        value for value in str(data.get("OT", "")).split(",") if value
                    )
                    marker_conditions: List[Tuple[str, Tuple[str, ...]]] = []
                    marker_sources = (
                        [data]
                        if str(data.get("KIND", "")).upper() == "WIRE"
                        else []
                    )
                    variations = data.get("VARIANTES", [])
                    if isinstance(variations, list):
                        marker_sources.extend(
                            option
                            for option in variations
                            if isinstance(option, dict)
                            and str(option.get("KIND", "")).upper() == "WIRE"
                        )
                    for marker_source in marker_sources:
                        source_type = str(marker_source.get("TIPO", "")).upper()
                        source_values = tuple(
                            value
                            for value in str(marker_source.get("OT", "")).split(",")
                            if value
                        )
                        condition = (source_type, source_values)
                        if source_type and source_values and condition not in marker_conditions:
                            marker_conditions.append(condition)

                    marker_required = bool(marker_conditions)
                    stats["puntos_requeridos"] = (
                        int(stats["puntos_requeridos"]) + len(marker_conditions)
                    )
                    completed_conditions: List[Tuple[str, Tuple[str, ...]]] = []
                    for marker_slot, condition in enumerate(marker_conditions):
                        marker_color = marker_colors.get(condition, 7)
                        marker_complete = self._add_application_marker(
                            current_definition,
                            entity,
                            marker_color,
                            marker_slot=marker_slot,
                            marker_count=len(marker_conditions),
                        ) >= 10
                        if marker_complete:
                            completed_conditions.append(condition)
                            stats["puntos_aplicacion"] = (
                                int(stats["puntos_aplicacion"]) + 1
                            )
                        else:
                            missing_points = stats["puntos_faltantes"]
                            if isinstance(missing_points, list):
                                missing_points.append(
                                    f"cavidad {normalize_cavity(cavity)} "
                                    f"({condition[0]}:{','.join(condition[1])}, "
                                    f"color ACI {marker_color})"
                                )

                    if not marker_required or len(completed_conditions) == len(marker_conditions):
                        for tag, value in (
                            ("TIPO", note_type),
                            ("OT", ",".join(note_values)),
                        ):
                            if tag in tags:
                                self._set_attribute(tags[tag], value)

                    if kind == "PLUG":
                        if "COLOR" in tags:
                            self._set_attribute(tags["COLOR"], "")
                        stats["cavidades_con_plug"] = int(stats["cavidades_con_plug"]) + 1
                        stats["cavidades_completadas"] = int(stats["cavidades_completadas"]) + 1
                        continue

                    wire_color = str(data.get("COLOR", ""))
                    visible_wire_color = spanish_wire_color(wire_color)
                    color_parts = [
                        part for part in wire_color.split("/") if part.strip()
                    ]
                    if color_parts:
                        stats["cavidades_que_requieren_color"] = int(
                            stats["cavidades_que_requieren_color"]
                        ) + 1
                    expected_solids = 4 if len(color_parts) > 1 else 2
                    existing_color = str(
                        self._get_property(tags.get("COLOR"), "TextString")
                        if tags.get("COLOR") is not None
                        else ""
                    ).strip().upper()
                    color_complete = bool(
                        color_parts
                        and existing_color == visible_wire_color.upper()
                    )
                    if color_parts and not color_complete:
                        color_solids = self._add_color_fill(
                            current_definition, entity, wire_color
                        )
                        color_complete = color_solids >= expected_solids
                        if color_complete and "COLOR" in tags:
                            color_complete = self._set_attribute(
                                tags["COLOR"], visible_wire_color
                            )
                    if color_complete:
                        stats["colores_insertados"] = int(stats["colores_insertados"]) + 1
                    stats["cavidades_completadas"] = int(stats["cavidades_completadas"]) + 1
                    continue

                # Algunos archivos guardan las cavidades un nivel más adentro.
                try:
                    nested_name = str(self._get_property(entity, "Name"))
                    nested = self._com_retry(
                        lambda: document.Blocks.Item(nested_name),
                        f"No se pudo abrir el bloque anidado {nested_name}",
                        attempts=5,
                        base_delay=0.10,
                    )
                except Exception:
                    continue
                walk(nested, depth + 1)

        walk(definition)
        for plug_key, count in encountered_plug_counts.items():
            expected_plug_counts[plug_key] = max(
                expected_plug_counts[plug_key], count
            )
        stats["plugs_requeridos"] = sum(expected_plug_counts.values())
        stats["plugs_escritos"] = sum(
            min(count, verified_plug_counts.get(plug_key, 0))
            for plug_key, count in expected_plug_counts.items()
        )
        missing_plugs: List[str] = []
        for (cavity, plug_number), count in sorted(
            expected_plug_counts.items(),
            key=lambda item: (
                0 if item[0][0].isdigit() else 1,
                int(item[0][0]) if item[0][0].isdigit() else item[0][0],
                item[0][1],
            ),
        ):
            missing_count = count - verified_plug_counts.get(
                (cavity, plug_number), 0
            )
            missing_plugs.extend(
                f"{cavity}:{plug_number}" for _ in range(max(0, missing_count))
            )
        stats["plugs_faltantes"] = missing_plugs
        # La obligación viene del reporte, incluso si el DWG no contiene un
        # cuadro coincidente. Contar solo los cuadros encontrados ocultaba la
        # omisión como un falso resultado de cero bloqueos requeridos.
        for cavity, count in encountered_blocked_counts.items():
            expected_blocked_counts[cavity] = max(expected_blocked_counts[cavity], count)
        stats["cavidades_bloqueadas"] = sum(expected_blocked_counts.values())
        stats["bloqueos_dibujados"] = sum(
            min(count, verified_blocked_counts.get(cavity, 0))
            for cavity, count in expected_blocked_counts.items()
        )
        stats["bloqueos_faltantes"] = [
            cavity
            for cavity in sorted(expected_blocked_counts, key=configuration_sort_key)
            for _ in range(max(0, expected_blocked_counts[cavity] - verified_blocked_counts[cavity]))
        ]
        return stats

    @staticmethod
    def _dxf_is_complete(path: Path) -> bool:
        try:
            if not path.is_file() or path.stat().st_size < 512:
                return False
            with path.open("rb") as handle:
                handle.seek(max(0, path.stat().st_size - 4096))
                tail = handle.read()
            return bool(re.search(rb"(?:^|[\r\n])EOF(?:[\r\n]|$)", tail))
        except OSError:
            return False

    @staticmethod
    def _available_output_path(requested: Path) -> Path:
        """Nunca intenta sobrescribir un DXF que puede estar abierto en AutoCAD."""
        if not requested.exists():
            return requested
        for index in range(1, 1000):
            candidate = requested.with_name(
                f"{requested.stem}_NUEVO_{index}{requested.suffix}"
            )
            if not candidate.exists():
                return candidate
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return requested.with_name(
            f"{requested.stem}_NUEVO_{timestamp}{requested.suffix}"
        )

    def _save_dxf(self, document, requested: Path) -> Tuple[Path, bool]:
        """Guarda una vez y acepta el archivo si AutoCAD lo completó antes del error COM."""
        target = self._available_output_path(requested)
        save_error: Optional[Exception] = None
        try:
            document.SaveAs(str(target), 65)
        except Exception as error:
            save_error = error

        for _ in range(20):
            if self._dxf_is_complete(target):
                return target, save_error is not None
            try:
                self.pythoncom.PumpWaitingMessages()
            except Exception:
                pass
            time.sleep(0.25)

        if save_error is not None:
            raise RuntimeError(
                "AutoCAD no pudo guardar un DXF completo. Cierra cualquier DXF "
                f"de salida abierto e inténtalo otra vez. Detalle: {save_error}"
            ) from save_error
        raise RuntimeError(
            f"AutoCAD terminó SaveAs, pero el archivo {target.name} no contiene EOF"
        )

    def assemble(
        self,
        output_path: Path,
        grouped_instances: Dict[str, List[Dict[str, object]]],
        columns: int = 4,
        cell_width: float = 650.0,
        cell_height: float = 390.0,
        *,
        close_document: bool = False,
    ) -> Dict[str, object]:
        if os.name != "nt":
            raise RuntimeError("La generación con AutoCAD solamente se ejecuta en Windows")
        self.pythoncom.CoInitialize()
        temp_folder = Path(tempfile.mkdtemp(prefix="conectores_dxf_"))
        log = {
            "posiciones_requeridas": sum(len(items) for items in grouped_instances.values()),
            "insertados": 0,
            "faltantes_fuente": 0,
            "fallos_insercion": 0,
            "marcadores_faltantes": 0,
            "cavidades_detectadas": 0,
            "cavidades_completadas": 0,
            "cavidades_con_plug": 0,
            "colores_insertados": 0,
            "puntos_aplicacion": 0,
            "plugs_escritos": 0,
            "plugs_verificados": 0,
            "cavidades_bloqueadas": 0,
            "bloqueos_dibujados": 0,
            "etiquetas_componentes": 0,
            "coincidencias_cavr": 0,
            "repeticiones_cavr_omitidas": 0,
            "opciones_distribuidas_cavr": 0,
            "variantes_graficas": 0,
            "advertencias": [],
            "errores": [],
        }
        critical_plug_failures: List[str] = []
        critical_marker_failures: List[str] = []
        critical_blocked_failures: List[str] = []
        pending_quality_recoveries: List[Dict[str, object]] = []
        document = None
        try:
            self.progress("Conectando con AutoCAD...")
            try:
                acad = self._com_retry(
                    lambda: self.win32.Dispatch("AutoCAD.Application"),
                    "No se pudo iniciar AutoCAD",
                    attempts=5,
                    base_delay=0.35,
                )
            except Exception as error:
                raise RuntimeError("No se pudo iniciar AutoCAD") from error
            acad.Visible = True
            document = self._com_retry(
                lambda: acad.Documents.Add(),
                "No se pudo crear el dibujo de salida",
                attempts=8,
                base_delay=0.25,
            )
            self._ensure_arial_regular(document)
            model = self._com_retry(
                lambda: document.ModelSpace,
                "No se pudo abrir ModelSpace",
                attempts=8,
                base_delay=0.20,
            )

            base_y = 0.0
            sequence = 0
            for group_name, instances in grouped_instances.items():
                self.progress(f"Insertando grupo {group_name}: {len(instances)} conectores")
                try:
                    title = self._com_retry(
                        lambda: model.AddMText(
                            self.point(-100.0, base_y + 125.0),
                            columns * cell_width,
                            f"PLANTILLA {group_name}",
                        ),
                        f"No se pudo crear el título {group_name}",
                        attempts=6,
                        base_delay=0.12,
                    )
                    title.Height = 16.0
                    title.Color = 2
                    self._apply_text_style(title)
                    title.Update()
                except Exception:
                    pass

                row_top = base_y + 80.0
                row_bottom = row_top
                for index, instance in enumerate(instances):
                    connector = instance["connector"]
                    source_value = instance.get("source")
                    cavity_data = instance["cavity_data"]
                    notes = instance.get("application_notes", [])
                    visible_notes = instance.get("visible_application_notes", notes)
                    visible_marker_colors = application_marker_colors(visible_notes)
                    part_colors = instance.get("part_colors", {})
                    component_labels = instance.get("component_labels", [])
                    column = index % columns
                    if column == 0:
                        row_bottom = row_top
                    end_of_row = (
                        column == columns - 1 or index == len(instances) - 1
                    )
                    cell_left = column * cell_width
                    x = cell_left + 20.0
                    y = row_top
                    occupied_bottom = y
                    sequence += 1
                    self.progress(f"{group_name}: {connector.name} / {connector.component}")

                    if not source_value:
                        log["faltantes_fuente"] += 1
                        reason = str(
                            instance.get(
                                "missing_reason", "DWG/DXF no encontrado en la carpeta"
                            )
                        )
                        if self._add_missing_placeholder(
                            model, x, y, connector, group_name, reason
                        ):
                            log["marcadores_faltantes"] += 1
                        log["advertencias"].append(
                            f"{group_name}: {connector.name} / {connector.component}: {reason}"
                        )
                        occupied_bottom = y - max(
                            100.0, min(160.0, cell_height * 0.30)
                        )
                        row_bottom = min(row_bottom, occupied_bottom)
                        if end_of_row:
                            row_top = row_bottom - 55.0
                        continue

                    source_path = Path(source_value)
                    safe_name = re.sub(
                        r"[^A-Za-z0-9_]+",
                        "_",
                        f"{sequence}_{connector.name}_{group_name}_{connector.component}",
                    )[:120]
                    temporary_source = temp_folder / f"{safe_name}{source_path.suffix.lower()}"
                    try:
                        shutil.copy2(source_path, temporary_source)
                        block_reference = self._com_retry(
                            lambda: model.InsertBlock(
                                self.point(x, y),
                                str(temporary_source),
                                1.0,
                                1.0,
                                1.0,
                                0.0,
                            ),
                            f"AutoCAD no pudo insertar {connector.name}",
                            attempts=12,
                            base_delay=0.25,
                        )
                        log["insertados"] += 1
                    except Exception as error:
                        log["fallos_insercion"] += 1
                        message = (
                            f"{group_name}: {connector.name} / {connector.component}: {error}"
                        )
                        log["errores"].append(message)
                        if self._add_missing_placeholder(
                            model,
                            x,
                            y,
                            connector,
                            group_name,
                            "ERROR AL INSERTAR EL ARCHIVO",
                        ):
                            log["marcadores_faltantes"] += 1
                        occupied_bottom = y - max(
                            100.0, min(160.0, cell_height * 0.30)
                        )
                        row_bottom = min(row_bottom, occupied_bottom)
                        if end_of_row:
                            row_top = row_bottom - 55.0
                        continue

                    try:
                        block_stats = None
                        last_fill_error: Optional[Exception] = None
                        for fill_attempt in range(5):
                            try:
                                candidate = self._fill_connector_block(
                                    document,
                                    block_reference,
                                    connector,
                                    group_name,
                                    cavity_data,
                                    notes,
                                    component_labels,
                                    index + 1,
                                    marker_colors=visible_marker_colors,
                                )
                                block_stats = candidate
                                expected = len(cavity_data)
                                completed_now = int(candidate["cavidades_completadas"])
                                colored_now = int(candidate["colores_insertados"])
                                wire_now = int(
                                    candidate["cavidades_que_requieren_color"]
                                )
                                points_now = int(candidate["puntos_aplicacion"])
                                points_required = int(candidate["puntos_requeridos"])
                                plugs_now = int(candidate["plugs_escritos"])
                                plugs_required = int(candidate["plugs_requeridos"])
                                blocked_now = int(candidate["bloqueos_dibujados"])
                                blocked_required = int(
                                    candidate["cavidades_bloqueadas"]
                                )
                                if not expected or (
                                    completed_now >= expected
                                    and colored_now >= wire_now
                                    and points_now >= points_required
                                    and plugs_now >= plugs_required
                                    and blocked_now >= blocked_required
                                ):
                                    break
                            except Exception as error:
                                last_fill_error = error
                            if fill_attempt < 4:
                                try:
                                    self.pythoncom.PumpWaitingMessages()
                                except Exception:
                                    pass
                                time.sleep(0.45 * (fill_attempt + 1))

                        if block_stats is None:
                            raise RuntimeError(
                                f"No se pudo analizar el bloque después de 5 intentos: "
                                f"{last_fill_error}"
                            )

                        # Los atributos lejanos ya fueron limpiados. Ahora se
                        # mide y alinea únicamente la geometría realmente visible.
                        x, y, block_bounds = self._align_connector_to_cell(
                            block_reference,
                            x,
                            y,
                            cell_left,
                            row_top,
                        )
                        occupied_bottom = min(
                            occupied_bottom, float(block_bounds[1])
                        )

                        log["cavidades_detectadas"] += int(
                            block_stats["cavidades_detectadas"]
                        )
                        log["cavidades_completadas"] += int(
                            block_stats["cavidades_completadas"]
                        )
                        log["colores_insertados"] += int(
                            block_stats["colores_insertados"]
                        )
                        log["cavidades_con_plug"] += int(
                            block_stats["cavidades_con_plug"]
                        )
                        log["puntos_aplicacion"] += int(
                            block_stats["puntos_aplicacion"]
                        )
                        log["plugs_escritos"] += int(
                            block_stats["plugs_escritos"]
                        )
                        log["plugs_verificados"] += int(
                            block_stats["plugs_escritos"]
                        )
                        log["cavidades_bloqueadas"] += int(
                            block_stats["cavidades_bloqueadas"]
                        )
                        log["bloqueos_dibujados"] += int(
                            block_stats["bloqueos_dibujados"]
                        )
                        log["coincidencias_cavr"] += int(
                            block_stats["coincidencias_cavr"]
                        )
                        log["repeticiones_cavr_omitidas"] += int(
                            block_stats["repeticiones_cavr_omitidas"]
                        )
                        distributed_options = block_stats.get(
                            "opciones_distribuidas_cavr", {}
                        )
                        if not isinstance(distributed_options, dict):
                            distributed_options = {}
                        log["opciones_distribuidas_cavr"] += sum(
                            int(value) for value in distributed_options.values()
                        )
                        expected = len(cavity_data)
                        detected = int(block_stats["cavidades_detectadas"])
                        completed = int(block_stats["cavidades_completadas"])
                        colored = int(block_stats["colores_insertados"])
                        wire_required = int(
                            block_stats["cavidades_que_requieren_color"]
                        )
                        points = int(block_stats["puntos_aplicacion"])
                        points_required = int(block_stats["puntos_requeridos"])
                        plugs_written = int(block_stats["plugs_escritos"])
                        plugs_required = int(block_stats["plugs_requeridos"])
                        missing_plugs = block_stats.get("plugs_faltantes", [])
                        if not isinstance(missing_plugs, list):
                            missing_plugs = []
                        blocked_drawn = int(block_stats["bloqueos_dibujados"])
                        blocked_required = int(
                            block_stats["cavidades_bloqueadas"]
                        )
                        if expected and not detected:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                "no se detectó ningún bloque con atributo CAV"
                            )
                        elif expected and not completed and not (
                            blocked_required or plugs_required or points_required
                        ):
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                f"se detectaron {detected} cavidades, pero ninguna coincidió"
                            )
                        elif wire_required and colored < wire_required:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                f"solo se colorearon {colored} de {wire_required} cavidades con cable"
                            )
                        missing_points = block_stats.get(
                            "puntos_faltantes", []
                        )
                        if not isinstance(missing_points, list):
                            missing_points = []
                        if (
                            points < points_required
                            or plugs_written < plugs_required
                            or blocked_drawn < blocked_required
                        ):
                            # No se declara el fallo todavía. AutoCAD puede estar
                            # ocupado mientras se construyen los conectores y
                            # aceptar la misma operación cuando termina el dibujo.
                            pending_quality_recoveries.append(
                                {
                                    "group_name": group_name,
                                    "block_reference": block_reference,
                                    "connector": connector,
                                    "cavity_data": cavity_data,
                                    "notes": notes,
                                    "component_labels": component_labels,
                                    "position_index": index + 1,
                                    "marker_colors": visible_marker_colors,
                                    "initial_points": points,
                                    "points_required": points_required,
                                    "initial_missing_points": list(missing_points),
                                    "initial_plugs": plugs_written,
                                    "plugs_required": plugs_required,
                                    "initial_missing_plugs": list(missing_plugs),
                                    "initial_blocked": blocked_drawn,
                                    "blocked_required": blocked_required,
                                    "initial_missing_blocked": list(
                                        block_stats.get("bloqueos_faltantes", [])
                                    ),
                                }
                            )

                        label_added, label_bottom = (
                            self._add_exterior_connector_label(
                                model,
                                block_reference,
                                x,
                                y,
                                connector,
                                component_labels,
                            )
                        )
                        occupied_bottom = min(occupied_bottom, label_bottom)
                        if label_added:
                            # Nombre + CON: + todas las líneas de componentes.
                            log["etiquetas_componentes"] += len(
                                component_labels
                            )
                        else:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                "no se pudo crear la etiqueta exterior"
                            )
                        self._yield_to_autocad(0.18)
                        notes_x, notes_y = self._right_application_anchor(
                            block_reference, x, y
                        )
                        display_notes = top_application_notes(
                            visible_notes,
                            cavity_data,
                            embedded_option_counts=distributed_options,
                        )
                        available_note_width = max(
                            153.0,
                            cell_left + cell_width - notes_x,
                        )
                        desired_note_columns = max(
                            1, math.ceil(len(display_notes) / 4)
                        )
                        fitting_note_columns = max(
                            1,
                            int((available_note_width - 153.0) // 160.0) + 1,
                        )
                        note_columns = min(
                            desired_note_columns, fitting_note_columns
                        )
                        note_rows = max(
                            1, math.ceil(len(display_notes) / note_columns)
                        )
                        notes_added, notes_bottom = self._add_application_notes(
                            model,
                            notes_x,
                            notes_y,
                            display_notes,
                            part_colors,
                            rows_per_column=note_rows,
                            column_width=160.0,
                            marker_colors=visible_marker_colors,
                        )
                        if display_notes:
                            occupied_bottom = min(
                                occupied_bottom, notes_bottom
                            )
                        if display_notes and not notes_added:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                f"no se pudo mostrar {format_application_notes(display_notes)}"
                            )
                        self._yield_to_autocad(0.20)
                        (
                            variations_added,
                            variations_expected,
                            variation_headers_failed,
                            variations_bottom,
                        ) = (
                            self._add_variation_cavities(
                                model,
                                x,
                                y,
                                cavity_data,
                                notes,
                                part_colors,
                                embedded_option_counts=distributed_options,
                                panel_x=cell_left + 20.0,
                                panel_y=occupied_bottom - 18.0,
                                marker_colors=visible_marker_colors,
                            )
                        )
                        if variations_expected:
                            occupied_bottom = min(
                                occupied_bottom, variations_bottom
                            )
                        log["variantes_graficas"] += variations_added
                        if variations_added < variations_expected:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                f"solo se dibujaron {variations_added} de "
                                f"{variations_expected} cavidades alternativas"
                            )
                        if variation_headers_failed:
                            log["advertencias"].append(
                                f"{connector.name} / {connector.component}: "
                                f"no se pudieron crear {variation_headers_failed} "
                                "encabezados de cavidades alternativas"
                            )
                        self._yield_to_autocad(0.14)
                    except Exception as error:
                        log["errores"].append(
                            f"{group_name}: {connector.name} / {connector.component}: {error}"
                        )
                        # El DWG ya fue insertado: no debe marcarse como faltante.
                        # El detalle queda en el resumen final para diagnóstico.
                        occupied_bottom = min(
                            occupied_bottom,
                            self._entity_bounds(block_reference, x, y)[1],
                        )
                    row_bottom = min(row_bottom, occupied_bottom)
                    if end_of_row:
                        row_top = row_bottom - 55.0

                # El siguiente grupo comienza después de la huella real de la
                # última fila, no después de una cantidad fija de celdas.
                base_y = row_bottom - 180.0

            if pending_quality_recoveries:
                self.progress(
                    "Control final: recuperando cavidades no confirmadas..."
                )
                # Se deja terminar la cola COM antes de volver a tocar los
                # bloques pendientes. Esto evita reiniciar toda la generación
                # por una respuesta tardía de AutoCAD.
                self._yield_to_autocad(0.90)
                try:
                    self._com_retry(
                        lambda: document.Regen(1),
                        "AutoCAD no pudo preparar la recuperación final",
                        attempts=5,
                        base_delay=0.20,
                    )
                except Exception:
                    pass

                for recovery_index, recovery in enumerate(
                    pending_quality_recoveries, start=1
                ):
                    connector = recovery["connector"]
                    group_name = str(recovery["group_name"])
                    final_stats: Optional[Dict[str, object]] = None
                    last_error: Optional[Exception] = None
                    for recovery_attempt in range(8):
                        self.progress(
                            "Control final "
                            f"{recovery_index}/{len(pending_quality_recoveries)}: "
                            f"{connector.name}, intento {recovery_attempt + 1}/8"
                        )
                        try:
                            final_stats = self._fill_connector_block(
                                document,
                                recovery["block_reference"],
                                connector,
                                group_name,
                                recovery["cavity_data"],
                                recovery["notes"],
                                recovery["component_labels"],
                                int(recovery["position_index"]),
                                marker_colors=recovery["marker_colors"],
                            )
                            final_points = int(
                                final_stats["puntos_aplicacion"]
                            )
                            final_points_required = int(
                                final_stats["puntos_requeridos"]
                            )
                            final_plugs = int(final_stats["plugs_escritos"])
                            final_plugs_required = int(
                                final_stats["plugs_requeridos"]
                            )
                            final_blocked = int(final_stats["bloqueos_dibujados"])
                            final_blocked_required = int(final_stats["cavidades_bloqueadas"])
                            if (
                                final_points >= final_points_required
                                and final_plugs >= final_plugs_required
                                and final_blocked >= final_blocked_required
                            ):
                                break
                        except Exception as error:
                            last_error = error
                            final_stats = None
                        try:
                            document.Regen(1)
                        except Exception:
                            pass
                        self._yield_to_autocad(
                            0.65 + recovery_attempt * 0.18
                        )

                    if final_stats is None:
                        final_points = int(recovery["initial_points"])
                        final_points_required = int(
                            recovery["points_required"]
                        )
                        final_plugs = int(recovery["initial_plugs"])
                        final_plugs_required = int(
                            recovery["plugs_required"]
                        )
                        final_missing_points = recovery[
                            "initial_missing_points"
                        ]
                        final_missing_plugs = recovery[
                            "initial_missing_plugs"
                        ]
                        final_blocked = int(recovery["initial_blocked"])
                        final_blocked_required = int(recovery["blocked_required"])
                        final_missing_blocked = recovery["initial_missing_blocked"]
                    else:
                        final_points = int(final_stats["puntos_aplicacion"])
                        final_points_required = int(
                            final_stats["puntos_requeridos"]
                        )
                        final_plugs = int(final_stats["plugs_escritos"])
                        final_plugs_required = int(
                            final_stats["plugs_requeridos"]
                        )
                        final_missing_points = final_stats.get(
                            "puntos_faltantes", []
                        )
                        final_missing_plugs = final_stats.get(
                            "plugs_faltantes", []
                        )
                        final_blocked = int(final_stats["bloqueos_dibujados"])
                        final_blocked_required = int(final_stats["cavidades_bloqueadas"])
                        final_missing_blocked = final_stats.get("bloqueos_faltantes", [])

                    if not isinstance(final_missing_points, list):
                        final_missing_points = []
                    if not isinstance(final_missing_plugs, list):
                        final_missing_plugs = []
                    if not isinstance(final_missing_blocked, list):
                        final_missing_blocked = []

                    recovered_points = max(
                        0,
                        final_points - int(recovery["initial_points"]),
                    )
                    recovered_plugs = max(
                        0,
                        final_plugs - int(recovery["initial_plugs"]),
                    )
                    recovered_blocked = max(
                        0, final_blocked - int(recovery["initial_blocked"])
                    )
                    log["puntos_aplicacion"] += recovered_points
                    log["plugs_escritos"] += recovered_plugs
                    log["plugs_verificados"] += recovered_plugs
                    log["bloqueos_dibujados"] += recovered_blocked
                    log["cavidades_completadas"] += recovered_blocked

                    if final_points < final_points_required:
                        detail = "; ".join(
                            str(value) for value in final_missing_points
                        )
                        if not detail and last_error is not None:
                            detail = str(last_error)
                        critical_marker_failures.append(
                            f"{group_name}: {connector.name} / "
                            f"{connector.component}: "
                            f"{detail or 'punto no identificado'}"
                        )
                    if final_plugs < final_plugs_required:
                        detail = ", ".join(
                            str(value) for value in final_missing_plugs
                        )
                        if not detail and last_error is not None:
                            detail = str(last_error)
                        critical_plug_failures.append(
                            f"{group_name}: {connector.name} / "
                            f"{connector.component}: "
                            f"{detail or 'cavidad no identificada'}"
                        )
                    if final_blocked < final_blocked_required:
                        detail = ", ".join(f"cavidad {value}" for value in final_missing_blocked)
                        if not detail and last_error is not None:
                            detail = str(last_error)
                        critical_blocked_failures.append(
                            f"{group_name}: {connector.name} / {connector.component}: "
                            f"{detail or 'diagonal no confirmada'}"
                        )

            if critical_plug_failures or critical_marker_failures or critical_blocked_failures:
                quality_failures: List[str] = []
                quality_failures.extend(
                    f"PLUG — {value}" for value in critical_plug_failures
                )
                quality_failures.extend(
                    f"PUNTO — {value}" for value in critical_marker_failures
                )
                quality_failures.extend(
                    f"BLOQUEO — {value}" for value in critical_blocked_failures
                )
                raise RuntimeError(
                    "CONTROL DE CALIDAD: no se guardó el DXF porque faltan "
                    "elementos verificados. AutoCAD no confirmó estas cavidades:\n- "
                    + "\n- ".join(quality_failures)
                    + "\n\nCierra otros procesos de AutoCAD y vuelve a generar. "
                    "El programa no entregará un archivo incompleto."
                )

            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            self.progress("Guardando DXF...")
            self._com_retry(
                lambda: document.Regen(1),
                "AutoCAD no pudo regenerar el dibujo",
                attempts=8,
                base_delay=0.20,
            )
            # 65 = AutoCAD 2018 DXF en la enumeración AcSaveAsType.
            saved_path, recovered_after_com_error = self._save_dxf(
                document, output_path
            )
            if saved_path != output_path:
                log["advertencias"].append(
                    f"El archivo {output_path.name} ya existía o estaba abierto; "
                    f"se guardó como {saved_path.name}"
                )
            if recovered_after_com_error:
                log["advertencias"].append(
                    "AutoCAD informó un error COM después de escribir el DXF, "
                    "pero el archivo se verificó completo mediante su marca EOF"
                )
            log["archivo_guardado"] = str(saved_path)
            self.progress(f"DXF guardado: {saved_path}")
            return log
        finally:
            if close_document and document is not None:
                try:
                    self._com_retry(
                        lambda: document.Close(False),
                        "No se pudo cerrar el dibujo de salida",
                        attempts=5,
                        base_delay=0.2,
                    )
                except Exception:
                    pass
            shutil.rmtree(temp_folder, ignore_errors=True)
            try:
                self.pythoncom.CoUninitialize()
            except Exception:
                pass
