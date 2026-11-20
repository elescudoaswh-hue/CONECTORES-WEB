"""Empaqueta el lanzador para Windows de 64 bits con un Python portátil."""

import argparse
from hashlib import file_digest
import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
from urllib.request import urlopen
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
PYTHON_VERSION = "3.13.14"
PYTHON_ZIP = f"python-{PYTHON_VERSION}-embed-amd64.zip"
PYTHON_SHA256 = "90b4e5b9898b72d744650524bff92377c367f44bd5fbd09e3148656c080ad907"
WHEEL_LOCK = ROOT / "scripts" / "launcher-wheels.txt"


def sha256(path):
    with path.open("rb") as file:
        return file_digest(file, "sha256").hexdigest()


def compile_native_host(package_root, cache, go):
    environment = os.environ | {
        "GOOS": "windows", "GOARCH": "amd64", "CGO_ENABLED": "0", "GOCACHE": str(cache / "go-build"),
    }
    subprocess.run([
        go, "build", "-trimpath", "-buildvcs=false", "-ldflags", "-s -w -H=windowsgui",
        "-o", str(package_root / "ConectoresNative.exe"), str(ROOT / "windows_native_host" / "main.go"),
    ], env=environment, check=True)
    go_root = subprocess.check_output([go, "env", "GOROOT"], text=True).strip()
    shutil.copyfile(Path(go_root) / "LICENSE", package_root / "runtime" / "Go-LICENSE.txt")


def package(destination, cache, go):
    cache.mkdir(parents=True, exist_ok=True)
    python_zip = cache / PYTHON_ZIP
    if not python_zip.exists():
        with urlopen(f"https://www.python.org/ftp/python/{PYTHON_VERSION}/{PYTHON_ZIP}", timeout=30) as stream:
            with python_zip.open("wb") as file:
                shutil.copyfileobj(stream, file)
    if sha256(python_zip) != PYTHON_SHA256:
        raise ValueError(f"El SHA-256 de {python_zip} difiere del valor esperado.")
    lock = {
        line.split("--hash=sha256:")[1] for line in WHEEL_LOCK.read_text(encoding="utf-8").splitlines() if line
    }
    with TemporaryDirectory() as temp:
        package_root = Path(temp) / "generador"
        runtime = package_root / "runtime"
        runtime.mkdir(parents=True)
        with ZipFile(python_zip) as archive:
            archive.extractall(runtime)
        (runtime / "python313._pth").write_text(
            "python313.zip\n.\n..\nLib\\site-packages\nimport site\n",
            encoding="utf-8",
        )
        wheel_dir = Path(temp) / "wheels"
        subprocess.run([
            sys.executable, "-m", "pip", "download", "--require-hashes", "--no-deps",
            "--no-cache-dir", "--only-binary=:all:", "--platform", "win_amd64",
            "--implementation", "cp", "--abi", "cp313", "--python-version", "3.13",
            "--find-links", str(cache), "--dest", str(wheel_dir), "-r", str(WHEEL_LOCK),
        ], check=True)
        wheels = list(wheel_dir.glob("*.whl"))
        if {sha256(wheel) for wheel in wheels} != lock:
            raise ValueError("Las dependencias descargadas difieren del listado del lanzador.")
        site_packages = runtime / "Lib" / "site-packages"
        for wheel in wheels:
            with ZipFile(wheel) as archive:
                archive.extractall(site_packages)
        for name in ("windows_local.py", "windows_native.py", "windows_service.py"):
            shutil.copyfile(ROOT / name, package_root / name)
        compile_native_host(package_root, cache.resolve(), go)
        shutil.copytree(ROOT / "analizador", package_root / "analizador", ignore=shutil.ignore_patterns("__pycache__"))
        script = "INSTALAR_GENERADOR_LOCAL.bat"
        (package_root / script).write_text(
            (ROOT / script).read_text(encoding="utf-8"), encoding="utf-8", newline="\r\n",
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        with ZipFile(destination, "w", compression=ZIP_DEFLATED, compresslevel=5) as archive:
            for file in sorted(package_root.rglob("*")):
                if file.is_file():
                    archive.write(file, file.relative_to(package_root))


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("output", type=Path)
    parser.add_argument("--cache", type=Path, required=True)
    parser.add_argument("--go", default="go", help="Ruta al compilador Go")
    arguments = parser.parse_args()
    package(arguments.output, arguments.cache, arguments.go)
