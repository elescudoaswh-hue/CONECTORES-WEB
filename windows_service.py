"""Servicio DXF para ejecutar en la sesión de Windows que tiene AutoCAD."""

from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from getpass import getpass
from importlib.util import find_spec
import json
import os
from pathlib import Path
import re
from secrets import compare_digest, token_urlsafe
import shutil
import sys
from tempfile import TemporaryDirectory
from threading import Lock
import time
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile, ZipFile

from flask import Flask, jsonify, request, send_file

from analizador.analysis import apply_colors, plan_report
from analizador.core.engine import APP_VERSION, AutoCADAssembler, InitialReportReader, scan_connector_folder
from analizador.services.generation import build_generation_instances
from analizador.services.selection import connector_selection_key

UPLOAD_LIMIT = 95 * 1024 * 1024
RESULT_LIFETIME = 2 * 60 * 60
METRICS = {
    "posiciones_requeridas": "Posiciones requeridas",
    "insertados": "Conectores insertados",
    "faltantes_fuente": "DWG/DXF faltantes",
    "fallos_insercion": "Fallos de inserción",
    "marcadores_faltantes": "Marcadores rojos",
    "cavidades_detectadas": "Cavidades detectadas",
    "cavidades_completadas": "Cavidades completadas",
    "cavidades_con_plug": "Cavidades con plug",
    "plugs_verificados": "Plugs verificados en cavidades",
    "cavidades_bloqueadas": "Cavidades bloqueadas",
    "bloqueos_dibujados": "Diagonales de bloqueo",
    "etiquetas_componentes": "Líneas PLUG/accesorios",
    "coincidencias_cavr": "Cuadros reconocidos por CAVR",
    "opciones_distribuidas_cavr": "Opciones distribuidas por CAVR",
    "repeticiones_cavr_omitidas": "Repeticiones CAVR omitidas",
    "colores_insertados": "Cuadros coloreados",
    "puntos_aplicacion": "Puntos de aplicación",
    "variantes_graficas": "Cavidades alternativas dibujadas",
}


def valid_name(name, suffixes):
    if not isinstance(name, str):
        raise ValueError("Se requiere un nombre de archivo.")
    name = name.strip()
    reserved = re.fullmatch(r"(?i)(CON|PRN|AUX|NUL|COM[1-9¹²³]|LPT[1-9¹²³])", name.split(".")[0].rstrip())
    if (
        not name or len(name) > 200 or reserved
        or re.search(r'[<>:"/\\|?*\x00-\x1f]', name)
        or Path(name).suffix.lower() not in suffixes
    ):
        raise ValueError("Nombre de archivo no válido: " + name[:200])
    return name


def read_options():
    try:
        options = json.loads(request.form.get("options", ""))
    except ValueError as error:
        raise ValueError("La configuración de generación no es válida.") from error
    if not isinstance(options, dict):
        raise ValueError("La configuración de generación no es válida.")
    assignments = options.get("assignments")
    selected = options.get("selected")
    if not (
        isinstance(assignments, dict) and all(isinstance(value, str) for value in assignments.values())
        and isinstance(selected, list) and selected
        and all(isinstance(key, str) and re.fullmatch(r"[0-9a-f]{64}", key) for key in selected)
        and isinstance(options.get("groups"), str)
        and type(options.get("allowMissing")) is bool
    ):
        raise ValueError("Revisa los colores y selecciona al menos un conector.")
    options["output"] = valid_name(options.get("output"), {".dxf"})
    return options


@dataclass
class Job:
    id: str
    folder: Path
    output: str
    status: str = "queued"
    message: str = "En cola para generar con AutoCAD…"
    result: dict = field(default_factory=dict)
    completed_at: float = 0
    drawing: Path | None = None
    detail: Path | None = None

    def snapshot(self):
        return {
            "id": self.id, "status": self.status, "message": self.message,
            "output": self.output, **self.result,
        }


def generation_summary(result, filename, workspace):
    def clean(message):
        return str(message).replace(str(workspace) + os.sep, "")

    warnings = [clean(item) for item in result["advertencias"]]
    errors = [clean(item) for item in result["errores"]]
    lines = [f"Archivo: {filename}"]
    lines.extend(f"{label}: {result[key]}" for key, label in METRICS.items())
    lines.extend([f"Advertencias: {len(warnings)}", f"Errores: {len(errors)}"])
    return "\n".join(lines), warnings + errors


def create_app(token, workspace, assembler_factory=AutoCADAssembler, cad_available=None):
    if not token:
        raise ValueError("Configura DXF_API_TOKEN antes de iniciar el generador.")
    if cad_available is None:
        cad_available = os.name == "nt" and find_spec("win32com") is not None
    app = Flask(__name__, static_folder=None)
    app.config.update(
        MAX_CONTENT_LENGTH=UPLOAD_LIMIT,
        MAX_FORM_MEMORY_SIZE=4 * 1024 * 1024,
        MAX_FORM_PARTS=10000,
    )
    pool = ThreadPoolExecutor(max_workers=1, thread_name_prefix="autocad")
    lock = Lock()
    jobs = {}
    workspace = Path(workspace)

    def update(job, **changes):
        with lock:
            for name, value in changes.items():
                setattr(job, name, value)

    def generate(job, report_file, options):
        update(job, status="running", message="Preparando generación con AutoCAD…")
        try:
            try:
                with ZipFile(report_file) as archive:
                    if sum(entry.file_size for entry in archive.infolist()) > 512 * 1024 * 1024:
                        raise ValueError("El Initial Report es demasiado grande al descomprimirlo.")
                report = InitialReportReader(report_file).read()
            except (BadZipFile, ParseError, KeyError) as error:
                raise ValueError("El archivo no es un Initial Report .xlsx válido.") from error
            # El navegador conserva el orden de las plantillas de la expresión original.
            assignments = apply_colors(report, options["assignments"], options["groups"])["assignments"]
            _, requirements = plan_report(report, assignments, options["groups"])
            selected = set(options["selected"])
            required = {
                connector.component
                for group, requirement in requirements.items()
                for connector in requirement.connectors
                if connector_selection_key(group, connector) in selected
            }
            matches, _ = scan_connector_folder(job.folder / "drawings", required)
            instances = build_generation_instances(report, requirements, matches, assignments, selected)
            if sum(map(len, instances.values())) != len(selected):
                raise ValueError("La selección cambió. Vuelve a analizar y selecciona los conectores.")
            if not options["allowMissing"] and any(
                not item["source"] for group in instances.values() for item in group
            ):
                raise ValueError("Faltan dibujos seleccionados. Revisa la carpeta DWG y vuelve a generar.")
            assembler = assembler_factory(progress=lambda message: update(job, message=message))
            result = assembler.assemble(job.folder / "output" / job.output, instances, close_document=True)
            drawing = Path(result["archivo_guardado"])
            if not drawing.is_file():
                raise ValueError("AutoCAD no dejó un archivo DXF descargable.")
            summary, messages = generation_summary(result, drawing.name, job.folder)
            detail = None
            if messages:
                try:
                    detail = drawing.with_name(f"{drawing.stem}_REPORTE_GENERACION.txt")
                    detail.write_text(
                        summary + "\n\nDETALLE COMPLETO\n" + "=" * 40 + "\n" + "\n".join(messages),
                        encoding="utf-8",
                    )
                except OSError:
                    app.logger.exception("No se pudo guardar el reporte de generación.")
                    detail = None
            update(
                job, status="completed", message=f"DXF generado: {drawing.name}",
                output=drawing.name, drawing=drawing, detail=detail,
                result={
                    "summary": summary, "messages": messages[:15],
                    "extraMessages": max(0, len(messages) - 15), "hasDetail": detail is not None,
                },
            )
        except Exception as error:
            message = str(error).replace(str(job.folder) + os.sep, "")
            update(job, status="failed", message=message or "No se pudo generar el DXF.")
        finally:
            shutil.rmtree(job.folder / "drawings", ignore_errors=True)
            shutil.rmtree(job.folder / "report", ignore_errors=True)
            update(job, completed_at=time.monotonic())

    def remove_expired():
        cutoff = time.monotonic() - RESULT_LIFETIME
        with lock:
            expired = [job for job in jobs.values() if 0 < job.completed_at < cutoff]
            for job in expired:
                shutil.rmtree(job.folder, ignore_errors=True)
                if not job.folder.exists():
                    jobs.pop(job.id)

    @app.before_request
    def authenticate():
        actual = request.headers.get("Authorization", "").encode("utf-8")
        if not compare_digest(actual, f"Bearer {token}".encode("utf-8")):
            return jsonify(message="Acceso al generador no autorizado."), 401
        remove_expired()

    @app.after_request
    def prevent_cache(response):
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.errorhandler(413)
    def upload_too_large(_error):
        return jsonify(message="Los archivos de generación superan el límite de carga (95 MiB)."), 413

    @app.get("/api/health")
    def health():
        return jsonify(autocad=cad_available, version=APP_VERSION)

    @app.post("/api/jobs")
    def create_job():
        if not cad_available:
            return jsonify(message="El generador necesita Windows, AutoCAD y pywin32."), 503
        try:
            options = read_options()
            report = request.files.get("report")
            if report is None:
                raise ValueError("Incluye el Initial Report.")
            report_name = valid_name(report.filename, {".xlsx"})
            drawings = [
                (file, valid_name(file.filename, {".dwg", ".dxf"}))
                for file in request.files.getlist("drawings")
            ]
        except ValueError as error:
            return jsonify(message=str(error)), 400

        with lock:
            if sum(job.completed_at == 0 for job in jobs.values()) >= 4:
                return jsonify(message="El generador está ocupado. Inténtalo en unos minutos."), 429
            job_id = token_urlsafe(24)
            job = jobs[job_id] = Job(job_id, workspace / job_id, options["output"])
        try:
            report_path = job.folder / "report" / report_name
            report_path.parent.mkdir(parents=True)
            report.save(report_path)
            for number, (file, name) in enumerate(drawings):
                folder = job.folder / "drawings" / str(number)
                folder.mkdir(parents=True)
                file.save(folder / name)
            pool.submit(generate, job, report_path, options)
        except Exception:
            with lock:
                jobs.pop(job_id, None)
            shutil.rmtree(job.folder, ignore_errors=True)
            app.logger.exception("No se pudo admitir el trabajo DXF.")
            return jsonify(message="No se pudieron preparar los archivos de generación."), 500
        with lock:
            return jsonify(job.snapshot()), 202

    @app.get("/api/jobs/<job_id>")
    def job_status(job_id):
        with lock:
            job = jobs.get(job_id)
            if job is None:
                return jsonify(message="El trabajo ya no está disponible."), 404
            return jsonify(job.snapshot())

    def send_result(attribute, mime_type):
        def download(job_id):
            with lock:
                job = jobs.get(job_id)
                path = getattr(job, attribute) if job else None
                if path is None:
                    return jsonify(message="El archivo aún no está disponible o ya caducó."), 404
                return send_file(path, mimetype=mime_type, as_attachment=True)
        return download

    app.add_url_rule(
        "/api/jobs/<job_id>/download", "draw", send_result("drawing", "application/dxf")
    )
    app.add_url_rule(
        "/api/jobs/<job_id>/report", "detail", send_result("detail", "text/plain")
    )
    return app, pool


def main():
    from waitress import serve

    token = os.environ.get("DXF_API_TOKEN", "").strip()
    if not token and sys.stdin.isatty():
        token = getpass("Token del generador DXF: ").strip()
    if not token:
        raise SystemExit("Configura DXF_API_TOKEN antes de iniciar el generador.")
    with TemporaryDirectory(prefix="conectores_web_") as workspace:
        app, pool = create_app(token, workspace)
        print("Generador DXF escuchando en http://127.0.0.1:8765", flush=True)
        try:
            serve(app, host="127.0.0.1", port=8765, max_request_body_size=UPLOAD_LIMIT)
        finally:
            pool.shutdown(wait=True)


if __name__ == "__main__":
    main()
