/* global loadPyodide */
importScripts("./runtime/pyodide.js");

const modules = [
  "__init__.py", "analysis.py", "browser_bridge.py",
  "core/__init__.py", "core/engine.py", "services/__init__.py", "services/selection.py",
];

async function initialize() {
  const runtime = await loadPyodide({ indexURL: new URL("./runtime/", self.location).href });
  const packagePath = "/analizador";
  runtime.FS.mkdirTree(`${packagePath}/core`);
  runtime.FS.mkdirTree(`${packagePath}/services`);
  await Promise.all(modules.map(async (filename) => {
    const response = await fetch(`./motor/analizador/${filename}`);
    if (!response.ok) throw new Error(`No se pudo cargar el motor de análisis (${response.status}).`);
    runtime.FS.writeFile(`${packagePath}/${filename}`, new Uint8Array(await response.arrayBuffer()));
  }));
  const dispatch = runtime.runPython(
    "import sys\nsys.path.insert(0, '/')\nfrom analizador.browser_bridge import dispatch\ndispatch",
  );
  return { runtime, dispatch };
}

let initialization;
let pending = Promise.resolve();
self.onmessage = ({ data: { id, action, payload, file } }) => {
  pending = pending.then(async () => {
    try {
      initialization ||= initialize().catch((error) => {
        initialization = undefined;
        throw error;
      });
      const { runtime, dispatch } = await initialization;
      const path = file ? "/tmp/entrada_analizador" : null;
      if (path) runtime.FS.writeFile(path, new Uint8Array(await file.arrayBuffer()));
      try {
        const result = dispatch(action, JSON.stringify(payload), path);
        self.postMessage({ id, result: JSON.parse(result) });
      } finally {
        if (path) runtime.FS.unlink(path);
      }
    } catch (error) {
      const message = String(error.message || error).trim();
      const pythonMessage = message.match(/(?:ValueError|OSError|MessageError): ([\s\S]+)$/);
      self.postMessage({ id, error: pythonMessage?.[1] || message.split("\n").at(-1) });
    }
  });
};
