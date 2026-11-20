"""Prepara la información de los conectores seleccionados sin usar AutoCAD."""
from __future__ import annotations

from pathlib import Path
from typing import Dict, List, Optional, Set
from analizador.core.engine import (
    ReportData, GroupRequirement, application_part_colors, cavity_analysis,
    complete_application_notes, application_notes, visible_circuit_notes,
    connector_component_labels,
)
from analizador.services.selection import connector_selection_key


def build_generation_instances(
    report: ReportData,
    requirements: Dict[str, GroupRequirement],
    matches: Dict[str, List[Path]],
    assignments: Dict[str, str],
    selected_keys: Optional[Set[str]] = None,
) -> Dict[str, List[Dict[str, object]]]:
    grouped = {}
    for group_name, requirement in requirements.items():
        instances = []
        part_colors = application_part_colors(requirement.part_numbers, assignments)
        for connector in requirement.connectors:
            if selected_keys is not None and connector_selection_key(
                group_name, connector
            ) not in selected_keys:
                continue
            paths = matches.get(connector.component, [])
            cavity_data, _ = cavity_analysis(report, connector, requirement.part_numbers)
            notes = complete_application_notes(
                application_notes(report, connector, requirement.part_numbers), cavity_data
            )
            instances.append({
                "connector": connector,
                "source": paths[0] if paths else None,
                "cavity_data": cavity_data,
                "application_notes": notes,
                "visible_application_notes": visible_circuit_notes(notes, cavity_data),
                "part_colors": part_colors,
                "component_labels": connector_component_labels(
                    report, connector, requirement.part_numbers
                ),
                "missing_reason": (
                    "DWG/DXF no encontrado en la carpeta seleccionada" if not paths else ""
                ),
            })
        if instances:
            grouped[group_name] = instances
    return grouped
