from io import BytesIO
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from analizador.analysis import analyze_report
from analizador.browser_bridge import dispatch
from analizador.core.engine import InitialReportReader
from windows_service import create_app, METRICS
from workbook_factory import make_report


ASSIGNMENTS = {"AAA": "ROSA", "ABA": "VERDE", "ERA": "AZUL"}
GROUPS = "ROSA+VERDE;AZUL"


class WebTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.report_path = make_report(
            self.root / "Prueba.xlsx",
            extra_accessories=[["X321A", "BLOCKED CAVITY", "", "9", "1", "X", "X", "X"]],
        )
        self.report = InitialReportReader(self.report_path).read()

    def test_web_analysis_reports_missing_and_uses_original_drawing_priority(self):
        drawings = [
            {"path": "DB/7289656830.dxf", "modified": 100},
            {"path": "DB/7289656830.dwg", "modified": 15},
            {"path": "DB/Histórico/7289656830 (1).DWG", "modified": 20},
        ]
        result = analyze_report(self.report, ASSIGNMENTS, GROUPS, drawings, "Ejemplo")
        self.assertEqual(result["summary"][0], ["ROSA_VERDE", "ROSA + VERDE", 2, 2, 1, 1, 0])
        self.assertEqual(result["connectors"][0]["drawing"], "DB/Histórico/7289656830 (1).DWG")
        self.assertIn("CAVIDADES BLOQUEADAS A MARCAR: 2", result["validation"])
        self.assertEqual(result["database"]["Ejemplo"]["azul"], ["ERA"])
        self.assertEqual(result["missingCount"], 0)
        missing = analyze_report(self.report, ASSIGNMENTS, GROUPS, [])
        self.assertEqual(missing["missingCount"], 4)
        self.assertEqual(missing["missingDrawings"], 1)
        self.assertIn("[ROSA_VERDE] X322A — 7289656830", missing["validation"])

    def test_web_bridge_applies_unique_family_and_preserves_original_filename(self):
        database = self.root / "familias.json"
        database.write_text(json.dumps({"Ejemplo": {
            "rosa": ["AAA"], "verde": ["ABA"], "azul": ["ERA"],
        }}), encoding="utf-8")
        dispatch("families", "{}", str(database))
        payload = json.dumps({"name": "Mi reporte original.xlsx"})
        details = json.loads(dispatch("report", payload, str(self.report_path)))
        self.assertEqual(details["matching"], ["Ejemplo"])
        applied = json.loads(dispatch("apply_family", json.dumps({
            "family": "Ejemplo", "groups": GROUPS,
        }), None))
        self.assertEqual(applied["assignments"], ASSIGNMENTS)
        result = json.loads(dispatch("analyze", json.dumps({
            **applied, "drawings": None, "exportFamily": "Ejemplo",
        }), None))
        self.assertEqual(result["database"]["_origen"]["archivo"], "Mi reporte original.xlsx")

    def test_directory_handles_need_dates_only_for_duplicate_required_drawings(self):
        paths = [
            "DB/7289656830.dxf", "DB/7289656830.dwg", "DB/Otros/7289656830 (1).DWG",
            "DB/11111111.dwg", "DB/Otros/11111111 (1).DWG",
        ]
        drawings = [{"path": path, "modified": None} for path in paths]
        pending = analyze_report(self.report, ASSIGNMENTS, GROUPS, drawings)
        self.assertEqual(pending, {"drawingsNeedingDates": sorted(paths[1:3])})

        drawings[1]["modified"] = 100
        drawings[2]["modified"] = 220
        result = analyze_report(self.report, ASSIGNMENTS, GROUPS, drawings)
        self.assertEqual(result["connectors"][0]["drawing"], paths[2])
        self.assertEqual(result["missingCount"], 0)
        self.assertIn("ARCHIVOS EXTRA EN CARPETA: 2", result["validation"])

    def test_job_downloads_the_selected_drawing_with_its_report(self):
        calls = []

        class TestAssembler:
            def __init__(self, progress):
                self.progress = progress

            def assemble(self, output, instances, *, close_document):
                calls.append((instances, close_document))
                self.progress("Generando dibujo de prueba…")
                output.parent.mkdir()
                output.write_bytes(b"0\nEOF\n")
                return dict.fromkeys(METRICS, 1) | {
                    "archivo_guardado": str(output), "advertencias": ["Aviso de prueba"], "errores": [],
                }

        app, pool = create_app("token-local-de-prueba", self.root, TestAssembler, cad_available=True)
        self.addCleanup(pool.shutdown)
        client = app.test_client()
        headers = {"Authorization": "Bearer token-local-de-prueba"}
        self.assertEqual(client.get("/api/health").status_code, 401)
        result = analyze_report(self.report, ASSIGNMENTS, GROUPS, None)
        row = result["connectors"][0]
        options = json.dumps({
            "assignments": ASSIGNMENTS, "groups": GROUPS, "output": "Análisis.dxf",
            "selected": [row["key"]], "allowMissing": False,
        })
        accepted = client.post("/api/jobs", headers=headers, data={
            "report": (BytesIO(self.report_path.read_bytes()), "Prueba.xlsx"),
            "drawings": (BytesIO(b"0\nEOF\n"), "7289656830.dxf"), "options": options,
        })
        self.assertEqual(accepted.status_code, 202)
        route = "/api/jobs/" + accepted.json["id"]
        pool.shutdown(wait=True)
        job = client.get(route, headers=headers).json
        self.assertEqual(job["status"], "completed", job)
        self.assertTrue(job["hasDetail"])
        self.assertEqual(len(calls), 1)
        self.assertTrue(calls[0][1])
        self.assertEqual([item["connector"].name for item in calls[0][0]["ROSA_VERDE"]], [row["name"]])
        download = client.get(route + "/download", headers=headers)
        try:
            self.assertEqual(download.data, b"0\nEOF\n")
            self.assertIn("attachment;", download.headers["Content-Disposition"])
        finally:
            download.close()
        detail = client.get(route + "/report", headers=headers)
        try:
            self.assertIn(b"Aviso de prueba", detail.data)
        finally:
            detail.close()


if __name__ == "__main__":
    unittest.main()
