from __future__ import annotations
import json
from pathlib import Path
import tempfile
import unittest
from analizador.core.engine import (
    InitialReportReader, configuration_from_header, normalize_configuration,
    short_configuration, assignments_from_family, load_family_database,
    matching_families, cavity_analysis, parse_group_expression,
    calculate_requirements, application_condition, application_part_colors,
    format_colored_application_mtext, scan_connector_folder,
    normalize_part_number, build_database_json, cavity_variation_items,
)
from analizador.services.selection import connector_selection_key
from analizador.services.generation import build_generation_instances
from workbook_factory import make_report


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)
        self.report = InitialReportReader(make_report(self.folder / "report.xlsx")).read()
        self.assignments = dict.fromkeys(self.report.part_numbers, "AZUL")
        self.groups = parse_group_expression("AZUL", self.assignments)
        self.requirements = calculate_requirements(self.report, self.assignments, self.groups)

    def test_alpha_headers_and_missing_article(self):
        self.assertEqual(self.report.part_numbers, ["AAA", "ABA", "ERA"])
        self.assertEqual(configuration_from_header("w3kt-14401-aaa A01 F0015930201"), "AAA")
        self.assertEqual(configuration_from_header("W3KT-14401-ERA A01"), "ERA")
        self.assertEqual(short_configuration("AAA"), "AAA")
        self.assertEqual(short_configuration("AB12"), "AB12")
        self.assertEqual(configuration_from_header("F0015930201"), "")
        self.assertEqual(configuration_from_header("BOM Inclusion"), "")

    def test_numeric_initial_report_still_reads_circuits_plugs_and_accessories(self):
        headers = ["68745494AA A03 F0016134500", "68745495AA A03 F0016134600", "68745496AA A03 F0016134700"]
        report = InitialReportReader(make_report(self.folder / "numeric.xlsx", headers)).read()
        self.assertEqual(report.part_numbers, ["68745494", "68745495", "68745496"])
        self.assertEqual(short_configuration(report.part_numbers[0]), "494")
        self.assertEqual(report.plugs[0].applicable_parts, {"68745496"})
        self.assertEqual(report.accessories[0].applicable_parts, set(report.part_numbers))
        self.assertEqual(configuration_from_header("1234567890AA A01 F0012345678"), "1234567890")

    def test_json_family_assignments_and_round_trip(self):
        path = self.folder / "family.json"
        path.write_text(json.dumps({"IP 736": {"azul": ["aaa", "ABA", "ERA"]}}))
        database = load_family_database(path)
        assignments, conflicts = assignments_from_family(database, "IP 736")
        self.assertEqual(assignments, self.assignments)
        self.assertFalse(conflicts)
        self.assertEqual(matching_families(self.report, database), ["IP 736"])
        output = build_database_json("IP 736", self.report, assignments, self.groups, self.requirements)
        path.write_text(json.dumps(output))
        self.assertEqual(load_family_database(path)["IP 736"]["AZUL"], ["AAA", "ABA", "ERA"])

    def test_conflicts_and_ambiguous_families_are_not_auto_chosen(self):
        database = {"IP 736": {"AZUL": ["AAA", "ABA", "ERA"], "VERDE": ["ERA"]}}
        assignments, conflicts = assignments_from_family(database, "IP 736")
        self.assertNotIn("ERA", assignments)
        self.assertEqual(conflicts["ERA"], ["AZUL", "VERDE"])
        self.assertEqual(matching_families(self.report, database), [])
        database = {name: {"AZUL": self.report.part_numbers} for name in ("F1", "F2")}
        self.assertEqual(matching_families(self.report, database), ["F1", "F2"])

    def test_configuration_is_not_article_or_numeric_suffix(self):
        self.assertEqual(normalize_configuration("W3KT-14401-AAA A01 F0015930201"), "AAA")
        database = {"incorrect": {"AZUL": ["F0015930201", "494"]}}
        self.assertEqual(matching_families(self.report, database), [])

    def test_shared_component_does_not_share_circuits(self):
        left, right = self.report.connectors
        left_data, _ = cavity_analysis(self.report, left, self.report.part_numbers)
        right_data, _ = cavity_analysis(self.report, right, self.report.part_numbers)
        self.assertEqual(left_data["1"]["CIRCUITO"], "LEFT")
        self.assertEqual(right_data["1"]["CIRCUITO"], "RIGHT")
        self.assertNotIn("6", left_data)
        self.assertEqual(right_data["6"]["CIRCUITOS_UNIDOS"], ["Z910AF", "Z910AM"])
        self.assertEqual(right_data["6"]["TIPO"], "N/A")
        self.assertEqual(right_data["6"]["OT"], "ERA")

    def test_plugs_and_accessories_keep_alpha_applicability(self):
        self.assertEqual(self.report.plugs[0].connector_name, "X322A")
        self.assertEqual(self.report.plugs[0].applicable_parts, {"ERA"})
        self.assertEqual(self.report.accessories[0].connector_name, "X321A")
        self.assertEqual(self.report.accessories[0].applicable_parts, {"AAA", "ABA", "ERA"})
        cavities, _ = cavity_analysis(self.report, self.report.connectors[1], self.report.part_numbers)
        self.assertEqual(cavities["2"]["PLUG"], "7158331550")
        self.assertEqual(cavities["2"]["OT"], "ERA")
        self.assertFalse(cavity_variation_items({"2": cavities["2"]}))

    def test_selection_limits_generation_to_requested_connector(self):
        left, right = self.report.connectors
        self.assertNotEqual(connector_selection_key("AZUL", left), connector_selection_key("AZUL", right))
        chosen = {connector_selection_key("AZUL", right)}
        instances = build_generation_instances(self.report, self.requirements, {}, self.assignments, chosen)
        self.assertEqual([i["connector"].name for i in instances["AZUL"]], ["X322A"])
        self.assertEqual(instances["AZUL"][0]["cavity_data"]["1"]["CIRCUITO"], "RIGHT")
        self.assertEqual(build_generation_instances(self.report, self.requirements, {}, self.assignments, set()), {})

    def test_letters_are_colored_and_wrapped_four_per_line(self):
        parts = ["AAA", "ABA", "ACA", "ADA", "ERA"]
        colors = application_part_colors(parts, dict.fromkeys(parts, "AZUL"))
        content, lines, _ = format_colored_application_mtext("SOLO", parts, colors)
        self.assertEqual(lines, 2)
        self.assertEqual(content.count("\\P"), 1)
        self.assertIn("ERA", content)
        self.assertNotIn("F001", content)
        self.assertEqual(set(colors), set(parts))

    def test_mixed_configuration_sorting_is_safe(self):
        kind, values = application_condition({"AAA", "68745494"}, {"AAA", "68745494", "ABA", "ERA"})
        self.assertEqual(kind, "SOLO")
        self.assertEqual(values, ["494", "AAA"])

    def test_full_component_names_and_leading_zeroes(self):
        self.assertEqual(normalize_part_number("TEMP03265603"), "03265603")
        self.assertEqual(normalize_part_number("X32657A649"), "X32657A649")
        self.assertEqual(normalize_part_number("9U5T-14489-TA"), "9U5T-14489-TA")
        for file in ("03265603.dwg", "9U5T-14489-TA (1).dwg", "X32657A649.dxf", "other.dwg"):
            (self.folder / file).touch()
        matches, extras = scan_connector_folder(self.folder, {"03265603", "9U5T-14489-TA", "X32657A649"})
        self.assertEqual(set(matches), {"03265603", "9U5T-14489-TA", "X32657A649"})
        self.assertEqual([p.name for p in extras], ["other.dwg"])

    def test_duplicate_configurations_are_reported(self):
        headers = ["W3KT-14401-AAA A01 F0015930201", "W3KT-14401-AAA A02 F0015930202", "W3KT-14401-ERA A01"]
        with self.assertRaisesRegex(ValueError, "AAA.*repetida"):
            InitialReportReader(make_report(self.folder / "duplicate.xlsx", headers)).read()


if __name__ == "__main__":
    unittest.main()
