const button = document.getElementById("link-site");
const status = document.getElementById("status");
const site = document.getElementById("site");
const scriptId = "conectores-yazaki";

function match(url) {
  return `${url.protocol}//${url.hostname}/*`;
}

async function prepare() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.url || tab.id === undefined) return false;
  const url = new URL(tab.url);
  if (url.protocol !== "https:") return false;
  const [check] = await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => Boolean(document.querySelector('meta[name="conectores-yazaki"][content="web"]')),
  });
  if (!check?.result) return false;

  const { linkedOrigin } = await chrome.storage.local.get("linkedOrigin");
  site.textContent = url.origin;
  site.hidden = false;
  status.textContent = linkedOrigin === url.origin
    ? "Sitio vinculado. Usa Generar DXF en la página."
    : "Autoriza esta dirección para conectar la página con AutoCAD.";
  button.textContent = linkedOrigin === url.origin ? "Volver a vincular" : "Vincular este sitio";
  button.disabled = false;

  button.addEventListener("click", async () => {
    button.disabled = true;
    const pattern = match(url);
    try {
      if (!(await chrome.permissions.request({ origins: [pattern] }))) {
        status.textContent = "La extensión necesita tu permiso para vincular este sitio.";
        return;
      }
      const settings = {
        id: scriptId, matches: [pattern], js: ["page-bridge.js"],
        runAt: "document_idle", persistAcrossSessions: true,
      };
      const existing = await chrome.scripting.getRegisteredContentScripts({ ids: [scriptId] });
      if (existing.length) await chrome.scripting.updateContentScripts([settings]);
      else await chrome.scripting.registerContentScripts([settings]);

      const previous = (await chrome.storage.local.get("linkedOrigin")).linkedOrigin;
      await chrome.storage.local.set({ linkedOrigin: url.origin });
      if (previous && match(new URL(previous)) !== pattern) {
        await chrome.permissions.remove({ origins: [match(new URL(previous))] });
      }
      try {
        await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ["page-bridge.js"] });
        status.textContent = "Sitio vinculado. Usa Generar DXF en la página.";
      } catch {
        status.textContent = "Sitio vinculado. Recarga la página para usar Generar DXF.";
      }
    } catch {
      status.textContent = "No se pudo vincular el sitio. Abre tu web de Conectores y vuelve a intentarlo.";
    } finally {
      button.disabled = false;
    }
  });
  return true;
}

prepare().catch(() => false).then((ready) => {
  if (!ready) status.textContent = "Abre tu web de Conectores y vuelve a pulsar el icono de la extensión.";
});
