"""Bloqueos explícitos del reporte y control de diagonales COM simulado."""
from __future__ import annotations
from collections import Counter
import json
import math
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import patch
from analizador.core import engine
from analizador.core.engine import (
    AutoCADAssembler, InitialReportReader, ConnectorRecord, cavity_analysis,
    is_blocked_cavity, select_cavity_option_for_block,
    blocked_cavity_diagonal,
)
from workbook_factory import make_report
from validar_reporte import validate


class FakeLine:
    ObjectName = "AcDbLine"

    def __init__(self, start, end, update_error=False):
        self.StartPoint, self.EndPoint = start, end
        self.update_error = update_error

    def Update(self):
        if self.update_error:
            raise RuntimeError("AutoCAD ocupado después de crear la línea")


class FakeBlock:
    Name = "ROOT"

    def __init__(self, entities=None, *, failures=0, error_after_create=False, update_error=False):
        self.entities = list(entities or [])
        self.failures = failures
        self.error_after_create = error_after_create
        self.update_error = update_error
        self.add_calls = 0

    @property
    def Count(self):
        return len(self.entities)

    def Item(self, index):
        return self.entities[index]

    def AddLine(self, start, end):
        self.add_calls += 1
        if self.add_calls <= self.failures:
            raise RuntimeError("AutoCAD ocupado")
        line = FakeLine(start, end, self.update_error)
        self.entities.append(line)
        if self.error_after_create:
            raise RuntimeError("AutoCAD creó la entidad y devolvió error COM")
        return line


class FakeReference:
    ObjectName = "AcDbBlockReference"
    Name = "ROOT"
    InsertionPoint = (0.0, 0.0, 0.0)
    Rotation = 0.0
    XScaleFactor = 1.0
    YScaleFactor = 1.0

    def __init__(self, cavity=None):
        self.attributes = [] if cavity is None else [SimpleNamespace(TagString="CAV", TextString=cavity)]
        self.HasAttributes = bool(self.attributes)

    def GetAttributes(self):
        return self.attributes


class FakeAssembler(AutoCADAssembler):
    """Ejecuta el flujo real de recuperación y guardado con un documento falso."""
    def __init__(self, block):
        self.block = block
        self.messages = []
        self.progress = self.messages.append
        self.saved = False
        self.fill_calls = 0
        self.pythoncom = SimpleNamespace(CoInitialize=lambda: None, CoUninitialize=lambda: None,
                                         PumpWaitingMessages=lambda: None)
        reference = FakeReference()
        model = SimpleNamespace(
            AddMText=lambda *args: SimpleNamespace(Update=lambda: None),
            InsertBlock=lambda *args: reference,
        )
        document = SimpleNamespace(
            Blocks=SimpleNamespace(Item=lambda name: self.block), ModelSpace=model,
            Regen=lambda mode: None,
        )
        app = SimpleNamespace(Documents=SimpleNamespace(Add=lambda: document))
        self.win32 = SimpleNamespace(Dispatch=lambda name: app)

    def point(self, x, y, z=0):
        return (x, y, z)

    def _com_retry(self, action, *args, **kwargs):
        return action()

    def _yield_to_autocad(self, *args):
        pass

    def _ensure_arial_regular(self, document):
        pass

    def _apply_text_style(self, entity):
        return True

    def _fill_connector_block(self, *args, **kwargs):
        self.fill_calls += 1
        return super()._fill_connector_block(*args, **kwargs)

    def _align_connector_to_cell(self, *args):
        return 0, 0, (0, -20, 20, 0)

    def _add_exterior_connector_label(self, *args):
        return True, -30

    def _right_application_anchor(self, *args):
        return 30, 0

    def _add_application_notes(self, *args, **kwargs):
        return True, -30

    def _add_variation_cavities(self, *args, **kwargs):
        return 0, 0, 0, -30

    def _save_dxf(self, document, path):
        self.saved = True
        return path, False


class BlockedCavityTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.folder = Path(self.temp.name)

    def test_blocked_description_and_component_in_both_formats(self):
        formats = [None, ["68745494AA A03 F0016134500", "68745495AA A03 F0016134600", "68745496AA A03 F0016134700"]]
        for headers in formats:
            with self.subTest(format="letters" if headers is None else "numeric"):
                rows = [
                    ["X321A", "", "BLOCKED CAVITY", "03, 4", "1", "X", "-", "-"],
                    ["X322A", "BLOCKED CAVITY", "", "5", "1", "-", "-", "X"],
                    ["X322A", "12345678", "blocked cavity", "7", "1", "X", "-", "-"],
                    ["X322A", "blocked_cavity", "CAVITY PLUG", "8", "1", "X", "-", "-"],
                    ["X321A", "", "BLOCKED CAVITY", "9", "1", "-", "-", "-"],
                ]
                report = InitialReportReader(make_report(self.folder / "report.xlsx", headers, rows)).read()
                blocked = [p for p in report.plugs if is_blocked_cavity(p.part_number)]
                self.assertEqual(len(blocked), 6)
                a, _, last = report.part_numbers
                left, right = report.connectors
                left_data, _ = cavity_analysis(report, left, [a])
                right_data, _ = cavity_analysis(report, right, [a])
                self.assertEqual(left_data["3"]["KIND"], "BLOCKED")
                self.assertEqual(left_data["4"]["KIND"], "BLOCKED")
                self.assertNotIn("3", right_data)
                self.assertNotIn("5", right_data)
                self.assertEqual(right_data["7"]["KIND"], "BLOCKED")
                self.assertEqual(right_data["8"]["KIND"], "BLOCKED")
                self.assertNotIn("9", left_data)
                last_data, _ = cavity_analysis(report, right, [last])
                self.assertEqual(last_data["5"]["KIND"], "BLOCKED")
                self.assertFalse(any(a.part_number == "BLOCKED CAVITY" for a in report.accessories))
                database = self.folder / "family.json"
                database.write_text(json.dumps({"FAMILY": {"AZUL": report.part_numbers}}))
                self.assertEqual(validate(report.source_path, database)["registros_bloqueo"], 6)

    def test_unlisted_empty_cavity_is_not_assumed_blocked(self):
        report = InitialReportReader(make_report(self.folder / "report.xlsx")).read()
        self.assertFalse(any(is_blocked_cavity(p.part_number) for p in report.plugs))
        data, _ = cavity_analysis(report, report.connectors[0], report.part_numbers)
        self.assertNotIn("3", data)

    def test_cavr_can_locate_blocked_cavity(self):
        data = {"3": {"KIND": "BLOCKED"}}
        cavity, selected, fallback, occurrence = select_cavity_option_for_block(data, "99", "03", {})
        self.assertEqual((cavity, selected["KIND"], fallback, occurrence), ("3", "BLOCKED", True, 0))

    def test_line_is_not_duplicated_when_addline_or_update_reports_error(self):
        for mode in ("normal", "addline", "update"):
            with self.subTest(mode=mode):
                block = FakeBlock(error_after_create=mode == "addline", update_error=mode == "update")
                assembler = FakeAssembler(block)
                self.assertTrue(assembler._add_blocked_cavity_mark(block, FakeReference("3")))
                self.assertTrue(assembler._add_blocked_cavity_mark(block, FakeReference("3")))
                self.assertEqual(block.add_calls, 1)
                self.assertEqual(len(block.entities), 1)

    def test_line_respects_rotation_and_scale(self):
        block = FakeBlock()
        reference = FakeReference("3")
        reference.InsertionPoint = (10, 20, 0)
        reference.Rotation = math.pi / 2
        reference.XScaleFactor = 2
        reference.YScaleFactor = 3
        self.assertTrue(FakeAssembler(block)._add_blocked_cavity_mark(block, reference))
        line = block.entities[0]
        self.assertAlmostEqual(line.StartPoint[0], 10)
        self.assertAlmostEqual(line.StartPoint[1], 20)
        self.assertAlmostEqual(line.EndPoint[0], 10 - 7.5 * 3)
        self.assertAlmostEqual(line.EndPoint[1], 20 + 7 * 2)

    def test_diagonal_touches_both_corners_from_reference_frame(self):
        # Coordenadas de los bloques Cavidad de los conectores de referencia.
        corner = (7.0, 7.486564212435496)
        segments = [
            ((0, 0), (10, 0)), ((10, 0), (10, 9.999996)),
            ((10, 9.999996), (0, 9.999996)), ((0, 9.999996), (0, 0)),
            ((7, 9.999996), corner), (corner, (10, corner[1])),
            ((0.35, 0.35), (6.85, 6.85)),
        ]
        self.assertEqual(blocked_cavity_diagonal(segments), ((0, 0), corner))
        definition = FakeBlock([FakeLine((*a, 0), (*b, 0)) for a, b in segments])
        document = SimpleNamespace(Blocks=SimpleNamespace(Item=lambda name: definition))
        parent = FakeBlock()
        assembler = FakeAssembler(parent)
        self.assertTrue(assembler._add_blocked_cavity_mark(parent, FakeReference("17"), document))
        line = parent.entities[0]
        self.assertEqual(line.StartPoint, (0, 0, 0))
        self.assertEqual(line.EndPoint, (*corner, 0))

    def test_geometry_can_be_read_from_polylines(self):
        outer = SimpleNamespace(ObjectName="AcDbPolyline", Closed=True,
                                Coordinates=(0, 0, 20, 0, 20, 30, 0, 30))
        number = SimpleNamespace(ObjectName="AcDbPolyline", Closed=True,
                                 Coordinates=(16, 25, 20, 25, 20, 30, 16, 30))
        definition = FakeBlock([outer, number])
        document = SimpleNamespace(Blocks=SimpleNamespace(Item=lambda name: definition))
        assembler = FakeAssembler(FakeBlock())
        self.assertEqual(assembler._blocked_cavity_local_diagonal(document, FakeReference("18")),
                         ((0, 0), (16, 25)))

    def test_blocked_cell_keeps_number_and_clears_embedded_text(self):
        reference = FakeReference("17")
        text = SimpleNamespace(TagString="CIRCUITO", TextString="Z910AF", MTextAttribute=True,
                               MTextAttributeContent="{\\FArial;Z910AF}",
                               Update=lambda: None, UpdateMTextAttribute=lambda: None)
        reference.attributes.append(text)
        block = FakeBlock([reference])
        assembler = FakeAssembler(block)
        document = SimpleNamespace(Blocks=SimpleNamespace(Item=lambda name: block))
        result = assembler._fill_connector_block(
            document, FakeReference(), ConnectorRecord("X321A", "7289656830", "Connector"),
            "AZUL", {"17": {"KIND": "BLOCKED"}}, [], [], 1,
        )
        self.assertEqual(reference.attributes[0].TextString, "17")
        self.assertEqual((text.TextString, text.MTextAttributeContent), ("", ""))
        self.assertEqual(result["bloqueos_dibujados"], 1)

    def _run_generation(self, block):
        source = self.folder / "source.dwg"
        source.write_bytes(b"fixture")
        connector = ConnectorRecord("X321A", "7289656830", "Connector", {"AAA"})
        instance = {"connector": connector, "source": source,
                    "cavity_data": {"3": {"KIND": "BLOCKED"}}}
        assembler = FakeAssembler(block)
        with patch.object(engine, "os", SimpleNamespace(name="nt")), patch.object(engine.time, "sleep"):
            result = assembler.assemble(self.folder / "out.dxf", {"AZUL": [instance]})
        return assembler, result

    def test_missing_cavity_is_reported_and_prevents_save(self):
        block = FakeBlock()
        source = self.folder / "source.dwg"
        source.touch()
        connector = ConnectorRecord("X321A", "7289656830", "Connector", {"AAA"})
        assembler = FakeAssembler(block)
        payload = {"AZUL": [{"connector": connector, "source": source,
                             "cavity_data": {"3": {"KIND": "BLOCKED"}}}]}
        with patch.object(engine, "os", SimpleNamespace(name="nt")), patch.object(engine.time, "sleep"):
            with self.assertRaisesRegex(RuntimeError, "BLOQUEO.*X321A.*cavidad 3"):
                assembler.assemble(self.folder / "out.dxf", payload)
        self.assertFalse(assembler.saved)
        self.assertEqual(assembler.fill_calls, 13)

    def test_failed_diagonal_is_recovered_before_save(self):
        # Cinco pasadas de siete intentos fallan. La recuperación final crea
        # la diagonal una vez que AutoCAD acepta AddLine.
        block = FakeBlock([FakeReference("3")], failures=35)
        assembler, result = self._run_generation(block)
        self.assertTrue(assembler.saved)
        self.assertEqual(assembler.fill_calls, 6)
        self.assertEqual(result["cavidades_bloqueadas"], 1)
        self.assertEqual(result["bloqueos_dibujados"], 1)
        self.assertEqual(result["cavidades_completadas"], 1)
        self.assertFalse(result["errores"])
        self.assertFalse(result["advertencias"])
        self.assertEqual(Counter(e.ObjectName for e in block.entities)["AcDbLine"], 1)


if __name__ == "__main__":
    unittest.main()
