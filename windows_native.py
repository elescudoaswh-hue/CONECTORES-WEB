"""Conexión de la extensión de Chrome o Edge con el lanzador local."""

from http.client import HTTPConnection, HTTPException
import json
import os
import re
import sys
from tempfile import TemporaryDirectory
from threading import Thread

from windows_local import (
    EXTENSION_ORIGIN, LOCAL_HEADER, PORT, configured_origin, get_local_server, port_in_use,
)

MESSAGE_LIMIT = 4096
MESSAGE_ID = re.compile(r"[a-f0-9-]{36}")


def read_message(stream):
    header = stream.read(4)
    if not header:
        return None
    size = int.from_bytes(header, "little")
    if len(header) != 4 or not 0 < size <= MESSAGE_LIMIT:
        raise ValueError("Tamaño de mensaje no válido.")
    payload = stream.read(size)
    if len(payload) != size:
        raise ValueError("Mensaje incompleto.")
    message = json.loads(payload)
    if not isinstance(message, dict):
        raise ValueError("Mensaje no válido.")
    return message


def write_message(stream, response):
    payload = json.dumps(response, ensure_ascii=False).encode("utf-8")
    stream.write(len(payload).to_bytes(4, "little") + payload)
    stream.flush()


def launcher_is_running(origin):
    connection = HTTPConnection("127.0.0.1", PORT, timeout=1)
    try:
        connection.request("GET", "/api/health", headers={"Origin": origin, LOCAL_HEADER: "1"})
        reply = connection.getresponse()
        return (
            reply.status == 200 and reply.getheader(LOCAL_HEADER) == "1"
            and reply.getheader("Access-Control-Allow-Origin") == origin
        )
    except (OSError, HTTPException):
        return False
    finally:
        connection.close()


def serve_native(input_stream, output_stream, origin):
    server = pool = thread = None
    with TemporaryDirectory(prefix="conectores_extension_") as workspace:
        try:
            while (message := read_message(input_stream)) is not None:
                request_id = message.get("id")
                if not isinstance(request_id, str) or not MESSAGE_ID.fullmatch(request_id):
                    continue
                response = {"id": request_id, "ok": True}
                if message.get("command") != "start" or set(message) != {"id", "command", "origin"}:
                    response.update(ok=False, message="La extensión solo puede iniciar el generador de Conectores.")
                elif message.get("origin") != origin:
                    response.update(ok=False, message="El lanzador se configuró para otra dirección web.")
                elif server is None and not launcher_is_running(origin):
                    try:
                        server, pool = get_local_server(origin, workspace)
                    except OSError as error:
                        if not (port_in_use(error) and launcher_is_running(origin)):
                            response.update(ok=False, message="No se pudo abrir el generador local en esta PC.")
                            write_message(output_stream, response)
                            break
                    else:
                        thread = Thread(target=server.run, name="conectores-local", daemon=True)
                        thread.start()
                write_message(output_stream, response)
        finally:
            if server is not None:
                server.close()
                server.task_dispatcher.shutdown()
                thread.join(timeout=1)
                pool.shutdown(wait=True)


def main():
    caller = sys.argv[1:]
    if os.name != "nt" or caller not in ([EXTENSION_ORIGIN], [EXTENSION_ORIGIN.rstrip("/")]):
        raise SystemExit("Esta conexión está reservada para la extensión de Conectores en Windows.")
    try:
        serve_native(sys.stdin.buffer, sys.stdout.buffer, configured_origin())
    except (ValueError, OSError, KeyError) as error:
        print(f"Conexión de Conectores cerrada. {error}", file=sys.stderr)


if __name__ == "__main__":
    main()
