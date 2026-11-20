"""Análisis de la versión web con las reglas del motor 0.7.4."""

from pathlib import Path

from analizador.core.engine import (
    adapt_group_expression,
    assignments_from_family,
    build_database_json,
    calculate_requirements,
    cavity_analysis,
    match_connector_files,
    normalize_color,
    parse_group_expression,
)
from analizador.services.selection import connector_selection_key


def plan_report(report, assignments, expression):
    missing = [part for part in report.part_numbers if not assignments.get(part)]
    if missing:
        raise ValueError("Hay números sin color: " + ", ".join(missing))
    groups = parse_group_expression(expression, assignments)
    return groups, calculate_requirements(report, assignments, groups)


def apply_family(report, database, family, expression):
    colors, conflicts = assignments_from_family(database, family)
    assignments = {part: colors.get(part, "") for part in report.part_numbers}
    missing = [part for part, color in assignments.items() if not color]
    return {
        "assignments": assignments,
        "groups": adapt_group_expression(expression, assignments.values()),
        "missing": missing,
        "conflicts": {part: conflicts[part] for part in missing if part in conflicts},
    }


def apply_colors(report, submitted, expression):
    assignments = {
        part: normalize_color(submitted.get(part, ""))
        for part in report.part_numbers
    }
    missing = [part for part, color in assignments.items() if not color]
    if missing:
        raise ValueError("Todavía hay números / códigos sin color: " + ", ".join(missing))
    return {
        "assignments": assignments,
        "groups": adapt_group_expression(expression, assignments.values()),
    }


def analyze_report(report, assignments, expression, drawings, export_family=None):
    groups, requirements = plan_report(report, assignments, expression)
    required = {
        component for requirement in requirements.values()
        for component in requirement.unique_components
    }
    scanned = drawings is not None
    matches, extras = match_connector_files(
        ((Path(item["path"]), item["modified"]) for item in (drawings or [])),
        required,
    )
    summary = []
    connectors = []
    missing_rows = []
    variations = set()
    blocked = []
    for group, requirement in requirements.items():
        unique = requirement.unique_components
        found = sum(component in matches for component in unique)
        summary.append([
            group, " + ".join(requirement.colors), len(requirement.part_numbers),
            len(requirement.connectors), len(unique),
            found if scanned else "—", len(unique) - found if scanned else "—",
        ])
        for connector in requirement.connectors:
            paths = matches.get(connector.component, [])
            missing = scanned and not paths
            connectors.append({
                "key": connector_selection_key(group, connector),
                "group": group,
                "name": connector.name,
                "component": connector.component,
                "kind": connector.kind,
                "drawing": paths[0].as_posix() if paths else None,
                "missing": missing,
            })
            if missing:
                missing_rows.append((group, connector.name, connector.component))
            cavity_data, messages = cavity_analysis(report, connector, requirement.part_numbers)
            variations.update(f"{group}: {message}" for message in messages)
            blocked.extend(
                f"[{group}] {connector.name} / {connector.component} — cavidad {cavity}"
                for cavity, data in cavity_data.items() if data.get("KIND") == "BLOCKED"
            )

    lines = [f"TOTAL DE POSICIONES: {len(connectors)}", f"COMPONENTES REQUERIDOS: {len(required)}"]
    if scanned:
        duplicates = sorted(
            f"{component}: " + " | ".join(path.as_posix() for path in paths)
            for component, paths in matches.items() if len(paths) > 1
        )
        lines.extend([
            f"ARCHIVOS EXTRA EN CARPETA: {len(extras)}", "",
            "FALTANTES (PLANTILLA | CONECTOR | NÚMERO DE PARTE):",
            "\n".join(f"[{group}] {name} — {component}"
                      for group, name, component in sorted(set(missing_rows))) or "Ninguno", "",
            "DUPLICADOS:", "\n".join(duplicates) or "Ninguno", "",
        ])
    lines.extend([
        f"CAVIDADES BLOQUEADAS A MARCAR: {len(blocked)}",
        "\n".join(sorted(set(blocked))) or "Ninguna indicada en el reporte", "",
        "CAVIDADES CON VARIACIONES:",
        "\n".join(sorted(variations)) or "Ninguna",
    ])

    result = {
        "summary": summary,
        "connectors": connectors,
        "missing": sorted(set(missing_rows)),
        "missingCount": len(missing_rows),
        "missingDrawings": len({component for _, _, component in missing_rows}),
        "validation": "\n".join(lines),
    }
    if export_family is not None:
        result["database"] = build_database_json(
            export_family, report, assignments, groups, requirements
        )
    return result
