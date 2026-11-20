import { dxfApi, localMode } from "./dxf.js";
import { startWithExtension } from "./extension-bridge.js";

export async function localHealth(signal) {
  const limit = AbortSignal.timeout(3000);
  return dxfApi(localMode, "/api/health", {
    signal: signal ? AbortSignal.any([limit, signal]) : limit,
  });
}

export function showLocalGenerator({ forJob = false, message } = {}) {
  const dialog = document.getElementById("local-generator-dialog");
  const site = document.getElementById("local-site-address");
  const status = document.getElementById("local-generator-status");
  const start = document.getElementById("open-local-generator");
  const check = document.getElementById("check-local-generator");
  const copy = document.getElementById("copy-local-site");
  const controller = new AbortController();
  let checking = false;

  site.value = window.location.origin;
  status.textContent = message || (forJob ? "Abre el lanzador para continuar con este DXF." : "");

  function update(message) {
    if (!controller.signal.aborted) status.textContent = message;
  }

  async function checkConnection(attempts) {
    if (checking) return;
    checking = true;
    check.disabled = true;
    update("Buscando el lanzador en esta PC…");
    try {
      for (let attempt = 0; attempt < attempts && !controller.signal.aborted; attempt += 1) {
        try {
          const health = await localHealth(controller.signal);
          if (controller.signal.aborted) return;
          if (!health.autocad) {
            update("El lanzador respondió. Para generar, se necesita AutoCAD en este equipo Windows.");
          } else if (forJob) {
            dialog.close("ready");
          } else {
            update("Lanzador conectado. Ya puedes pulsar Generar DXF.");
          }
          return;
        } catch {
          if (attempt + 1 < attempts && !controller.signal.aborted) {
            await new Promise((resolve) => setTimeout(resolve, 1000));
          }
        }
      }
      update("Lanzador sin conexión. Comprueba que esté abierto y que el navegador permita el acceso a la red local.");
    } finally {
      check.disabled = false;
      checking = false;
    }
  }

  start.onclick = () => {
    setTimeout(() => {
      if (!controller.signal.aborted) void checkConnection(12);
    }, 600);
  };
  check.onclick = () => void checkConnection(1);
  copy.onclick = async () => {
    try {
      await navigator.clipboard.writeText(site.value);
      update("Dirección copiada.");
    } catch {
      site.select();
      update("Copia la dirección seleccionada y pégala en el instalador.");
    }
  };

  dialog.returnValue = "";
  return new Promise((resolve) => {
    dialog.addEventListener("close", () => {
      controller.abort();
      start.onclick = null;
      check.onclick = null;
      copy.onclick = null;
      resolve(dialog.returnValue === "ready");
    }, { once: true });
    dialog.showModal();
  });
}

export async function ensureLocalGenerator() {
  let health;
  try {
    health = await localHealth();
  } catch {
    const extension = await startWithExtension();
    if (!extension?.ok) return showLocalGenerator({ forJob: true, message: extension?.message });
    try {
      health = await localHealth();
    } catch {
      return showLocalGenerator({
        forJob: true, message: "La extensión inició el lanzador. Confirma el acceso local y pulsa Comprobar conexión.",
      });
    }
  }
  if (!health.autocad) throw new Error("Para generar DXF se necesita AutoCAD en este equipo Windows.");
  return true;
}
