"""Lanzador del generador DXF para la sesión de Windows del usuario."""

import errno
import json
import os
from pathlib import Path
import re
from secrets import token_urlsafe
import shutil
import sys
from tempfile import TemporaryDirectory
from urllib.parse import urlsplit

from waitress import create_server
from werkzeug.wrappers import Response

from windows_service import UPLOAD_LIMIT, create_app

VERSION = "1.1.0"
SCHEME = "conectores-dxf"
NATIVE_HOST = "com.conectores_yazaki.generador"
EXTENSION_ORIGIN = "chrome-extension://aclgejkdghnapkoegdmciammdjgpolpn/"
NATIVE_EXECUTABLE = "ConectoresNative.exe"
PORT = 8766
HOST = f"127.0.0.1:{PORT}"
LOCAL_HEADER = "X-Conectores-Local"
JOB_ROUTE = re.compile(r"/api/jobs/[a-zA-Z0-9_-]{32}(?:/download|/report)?")
HOSTNAME = re.compile(r"[a-z0-9]+(?:[a-z0-9-]*[a-z0-9])?(?:\.[a-z0-9]+(?:[a-z0-9-]*[a-z0-9])?)*")


def site_origin(value):
    try:
        url = urlsplit(value.strip())
        host = url.hostname
        port = url.port
        if (
            url.scheme == "https" and host and HOSTNAME.fullmatch(host)
            and url.username is None and url.password is None
            and url.path in ("", "/") and not (url.query or url.fragment)
            and (port is None or port > 0)
        ):
            return f"https://{host}{':' + str(port) if port and port != 443 else ''}"
    except ValueError:
        pass
    raise ValueError("Pega la dirección HTTPS de la página, sin ruta ni parámetros.")


class LocalConnection:
    """Autoriza solicitudes del sitio que se eligió en la instalación."""

    def __init__(self, application, origin, token):
        self.application = application
        self.origin = site_origin(origin)
        self.authorization = f"Bearer {token}"

    def __call__(self, environ, start_response):
        if environ.get("HTTP_HOST") != HOST or environ.get("HTTP_ORIGIN") != self.origin:
            return Response("Origen sin permiso.", status=403)(environ, start_response)

        def reply(status, headers, exc_info=None):
            headers.extend([
                ("Access-Control-Allow-Origin", self.origin),
                ("Access-Control-Expose-Headers", LOCAL_HEADER),
                ("Vary", "Origin"),
                (LOCAL_HEADER, "1"),
            ])
            return start_response(status, headers, exc_info)

        path = environ.get("PATH_INFO", "")
        method = environ["REQUEST_METHOD"]
        if path == "/api/jobs":
            allowed = "POST"
        elif path == "/api/health" or JOB_ROUTE.fullmatch(path):
            allowed = "GET"
        else:
            return Response("Ruta desconocida.", status=404)(environ, reply)

        if method == "OPTIONS":
            headers = {
                name.strip().lower()
                for name in environ.get("HTTP_ACCESS_CONTROL_REQUEST_HEADERS", "").split(",")
                if name.strip()
            }
            if (
                environ.get("HTTP_ACCESS_CONTROL_REQUEST_METHOD") != allowed
                or LOCAL_HEADER.lower() not in headers
                or headers - {LOCAL_HEADER.lower(), "content-type"}
            ):
                return Response("Solicitud sin permiso.", status=403)(environ, reply)
            response = Response(status=204, headers={
                "Access-Control-Allow-Methods": allowed,
                "Access-Control-Allow-Headers": f"{LOCAL_HEADER}, Content-Type",
                "Vary": "Access-Control-Request-Method, Access-Control-Request-Headers",
            })
            return response(environ, reply)

        if method != allowed or environ.get("HTTP_X_CONECTORES_LOCAL") != "1":
            return Response("Solicitud sin permiso.", status=403)(environ, reply)
        environ["HTTP_AUTHORIZATION"] = self.authorization
        return self.application(environ, reply)


def create_local_app(origin, workspace, **app_options):
    token = token_urlsafe(32)
    app, pool = create_app(token, workspace, **app_options)
    app.wsgi_app = LocalConnection(app.wsgi_app, origin, token)
    return app, pool


def package_folder():
    return Path(__file__).resolve().parent


def register_native_host(winreg, destination):
    manifest = destination / f"{NATIVE_HOST}.json"
    manifest.write_text(json.dumps({
        "name": NATIVE_HOST,
        "description": "Generador de Conectores Yazaki",
        "path": str(destination / NATIVE_EXECUTABLE),
        "type": "stdio",
        "allowed_origins": [EXTENSION_ORIGIN],
    }), encoding="utf-8")
    for browser in (r"Google\Chrome", r"Microsoft\Edge"):
        key_path = rf"Software\{browser}\NativeMessagingHosts\{NATIVE_HOST}"
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "", 0, winreg.REG_SZ, str(manifest))


def install():
    import win32com.client
    import winreg

    origin = site_origin(input("Pega la dirección del Analizador de Conectores y pulsa Enter\n> "))
    base = Path(os.environ["LOCALAPPDATA"]) / "ConectoresYazaki" / "GeneradorLocal"
    base.mkdir(parents=True, exist_ok=True)
    destination = base / VERSION
    source = package_folder()
    if not (source / "runtime" / "python.exe").is_file() or not (source / NATIVE_EXECUTABLE).is_file():
        raise ValueError("Extrae todos los archivos del ZIP del lanzador antes de instalar.")
    if not destination.exists():
        with TemporaryDirectory(prefix="instalacion-", dir=base) as temp:
            staged = Path(temp) / VERSION
            shutil.copytree(source, staged, ignore=shutil.ignore_patterns("sitio.json", "__pycache__"))
            staged.rename(destination)
    command = f'"{destination / "runtime" / "python.exe"}" "{destination / "windows_local.py"}" "%1"'
    configuration = destination / "sitio.json"
    with TemporaryDirectory(prefix="sitio-", dir=destination) as temp:
        staged = Path(temp) / configuration.name
        staged.write_text(json.dumps({"origin": origin}), encoding="utf-8")
        staged.replace(configuration)
    key_path = rf"Software\Classes\{SCHEME}"
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Generador local de Conectores Yazaki")
        winreg.SetValueEx(key, "URL Protocol", 0, winreg.REG_SZ, "")
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r"\shell\open\command") as key:
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, command)
    register_native_host(winreg, destination)
    print(f"Instalación lista para {origin}. Vuelve a la web para generar el DXF.")


def configured_origin():
    configuration = json.loads((package_folder() / "sitio.json").read_text(encoding="utf-8"))
    return site_origin(configuration["origin"])


def get_local_server(origin, workspace):
    app, pool = create_local_app(origin, workspace)
    try:
        server = create_server(app, host="127.0.0.1", port=PORT, max_request_body_size=UPLOAD_LIMIT)
    except Exception:
        pool.shutdown(wait=True)
        raise
    return server, pool


def port_in_use(error):
    return error.errno == errno.EADDRINUSE or getattr(error, "winerror", None) == 10048


def run_local():
    origin = configured_origin()
    with TemporaryDirectory(prefix="conectores_local_") as workspace:
        try:
            server, pool = get_local_server(origin, workspace)
        except OSError as error:
            if port_in_use(error):
                return
            raise
        print("Generador local de Conectores Yazaki", flush=True)
        print(f"Listo para recibir trabajos de {origin}", flush=True)
        print("Mantén esta ventana abierta mientras generas el DXF.", flush=True)
        try:
            server.run()
        finally:
            pool.shutdown(wait=True)


def main():
    if os.name != "nt":
        raise SystemExit("El lanzador se ejecuta en Windows.")
    arguments = sys.argv[1:]
    if arguments == ["--install"]:
        try:
            install()
        except (ValueError, OSError, ImportError) as error:
            raise SystemExit(f"No se pudo instalar el lanzador. {error}") from error
    elif arguments in ([f"{SCHEME}://start"], [f"{SCHEME}://start/"]):
        run_local()


if __name__ == "__main__":
    main()
