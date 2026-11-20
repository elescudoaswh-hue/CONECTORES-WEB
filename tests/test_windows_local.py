from io import BytesIO
import json
from pathlib import Path
import sys
from tempfile import TemporaryDirectory
from types import ModuleType
import unittest
from unittest.mock import MagicMock, patch

from analizador.analysis import analyze_report
from analizador.core.engine import InitialReportReader
from windows_local import HOST, LOCAL_HEADER, SCHEME, VERSION, create_local_app, install, site_origin
from windows_service import METRICS
from workbook_factory import make_report


ORIGIN = "https://conectores.example"
LOCAL_URL = f"http://{HOST}"
REQUEST_HEADERS = {"Origin": ORIGIN, LOCAL_HEADER: "1"}


class LocalAppTests(unittest.TestCase):
    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_origin_is_the_exact_https_address_entered_at_installation(self):
        self.assertEqual(site_origin(" https://Conectores.Example:443/ "), ORIGIN)
        for value in (
            "http://conectores.example", "https://user@conectores.example",
            ORIGIN + "/path", ORIGIN + "?query", ORIGIN + ":invalid",
        ):
            with self.subTest(value=value), self.assertRaises(ValueError):
                site_origin(value)

    def test_local_listener_checks_host_origin_methods_and_preflight_headers(self):
        app, pool = create_local_app(ORIGIN, self.root, cad_available=True)
        self.addCleanup(pool.shutdown)
        client = app.test_client()
        path = "/api/health"
        for address, headers in (
            (LOCAL_URL, {"Origin": "https://otra-pagina.example", LOCAL_HEADER: "1"}),
            (LOCAL_URL, {"Origin": "null", LOCAL_HEADER: "1"}),
            (LOCAL_URL, {LOCAL_HEADER: "1"}),
            ("http://other.example:8766", REQUEST_HEADERS),
        ):
            denied = client.get(path, base_url=address, headers=headers)
            self.assertEqual(denied.status_code, 403)
            self.assertNotIn("Access-Control-Allow-Origin", denied.headers)

        rejected = client.get(path, base_url=LOCAL_URL, headers={"Origin": ORIGIN})
        self.assertEqual(rejected.status_code, 403)
        preflight = {
            "Origin": ORIGIN, "Access-Control-Request-Method": "POST",
            "Access-Control-Request-Headers": "content-type,x-conectores-local",
        }
        allowed = client.options("/api/jobs", base_url=LOCAL_URL, headers=preflight)
        self.assertEqual(allowed.status_code, 204)
        self.assertEqual(allowed.headers["Access-Control-Allow-Origin"], ORIGIN)
        self.assertIn(LOCAL_HEADER, allowed.headers["Access-Control-Allow-Headers"])
        self.assertNotIn("Access-Control-Allow-Credentials", allowed.headers)
        preflight["Access-Control-Request-Headers"] = "x-conectores-local,authorization"
        self.assertEqual(client.options("/api/jobs", base_url=LOCAL_URL, headers=preflight).status_code, 403)
        healthy = client.get(path, base_url=LOCAL_URL, headers=REQUEST_HEADERS)
        self.assertTrue(healthy.json["autocad"])
        self.assertEqual(healthy.headers[LOCAL_HEADER], "1")

    def test_local_generator_receives_and_returns_a_drawing_using_the_same_engine(self):
        class CadDePrueba:
            def __init__(self, progress):
                self.progress = progress

            def assemble(self, output, instances, *, close_document):
                item = next(iter(instances.values()))[0]
                output.parent.mkdir()
                output.write_bytes(item["source"].read_bytes())
                return dict.fromkeys(METRICS, 1) | {
                    "archivo_guardado": str(output), "advertencias": ["Aviso de prueba"], "errores": [],
                }

        report = make_report(self.root / "Reporte.xlsx")
        colors = {"AAA": "ROSA", "ABA": "VERDE", "ERA": "AZUL"}
        groups = "ROSA+VERDE;AZUL"
        analysis = analyze_report(InitialReportReader(report).read(), colors, groups, None)
        item = analysis["connectors"][0]
        options = {
            "assignments": colors, "groups": groups, "output": "Resultado.dxf",
            "selected": [item["key"]], "allowMissing": False,
        }
        app, pool = create_local_app(
            ORIGIN, self.root, assembler_factory=CadDePrueba, cad_available=True,
        )
        self.addCleanup(pool.shutdown)
        client = app.test_client()
        data = b"0\nEOF\n"
        started = client.post("/api/jobs", base_url=LOCAL_URL, headers=REQUEST_HEADERS, data={
            "options": json.dumps(options),
            "report": (BytesIO(report.read_bytes()), report.name),
            "drawings": (BytesIO(data), item["component"] + ".dxf"),
        })
        self.assertEqual(started.status_code, 202)
        pool.shutdown(wait=True)
        path = "/api/jobs/" + started.json["id"]
        finished = client.get(path, base_url=LOCAL_URL, headers=REQUEST_HEADERS)
        self.assertEqual(finished.json["status"], "completed", finished.json)
        self.assertTrue(finished.json["hasDetail"])
        self.assertEqual(client.get(path + "/download", base_url=LOCAL_URL).status_code, 403)
        with client.get(path + "/download", base_url=LOCAL_URL, headers=REQUEST_HEADERS) as drawing:
            self.assertEqual(drawing.data, data)
        with client.get(path + "/report", base_url=LOCAL_URL, headers=REQUEST_HEADERS) as detail:
            self.assertIn(b"Aviso de prueba", detail.data)

    def test_installer_copies_to_profile_and_registers_only_with_current_user(self):
        source = self.root / "Descargas"
        (source / "runtime").mkdir(parents=True)
        (source / "runtime" / "python.exe").touch()
        (source / "windows_local.py").touch()
        profile = self.root / "Perfil local"
        registry = MagicMock()
        win32com = ModuleType("win32com")
        win32com.client = ModuleType("win32com.client")
        with patch.dict(sys.modules, {
            "winreg": registry, "win32com": win32com, "win32com.client": win32com.client,
        }), patch("windows_local.package_folder", return_value=source), \
            patch("builtins.input", return_value=ORIGIN), patch("builtins.print"), \
            patch.dict("os.environ", {"LOCALAPPDATA": str(profile)}):
            install()
        for call in registry.CreateKey.call_args_list:
            hive, path = call.args
            self.assertIs(hive, registry.HKEY_CURRENT_USER)
            self.assertTrue(path.startswith(rf"Software\Classes\{SCHEME}"))
        installed = profile / "ConectoresYazaki" / "GeneradorLocal" / VERSION
        self.assertEqual(json.loads((installed / "sitio.json").read_text()), {"origin": ORIGIN})
        self.assertEqual(
            registry.SetValueEx.call_args.args[-1],
            f'"{installed / "runtime" / "python.exe"}" "{installed / "windows_local.py"}" "%1"',
        )


if __name__ == "__main__":
    unittest.main()
