import { cp, mkdir, rm, stat } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
const dist = join(root, "dist");
const maxAsset = 25 * 1024 * 1024;
const runtime = [
  "pyodide.js", "pyodide.asm.js", "pyodide.asm.wasm",
  "python_stdlib.zip", "pyodide-lock.json",
];

await rm(dist, { recursive: true, force: true });
await cp(join(root, "web"), dist, { recursive: true });
await mkdir(join(dist, "runtime"));
await Promise.all(runtime.map(async (filename) => {
  const source = join(root, "node_modules", "pyodide", filename);
  if ((await stat(source)).size > maxAsset) {
    throw new Error(`${filename} supera el tamaño máximo de un asset de Cloudflare.`);
  }
  await cp(source, join(dist, "runtime", filename));
}));
await cp(join(root, "analizador"), join(dist, "motor", "analizador"), {
  recursive: true,
  filter: (source) => !source.endsWith("__pycache__") && !source.endsWith(".pyc"),
});
console.log("Interfaz y motor Python listos en dist/.");
