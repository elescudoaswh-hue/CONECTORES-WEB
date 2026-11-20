const hostName = "com.conectores_yazaki.generador";
const unavailable = "Actualiza el lanzador local de Conectores para conectar la extensión.";
const pending = new Map();
let nativePort;

function finish(id, response) {
  const request = pending.get(id);
  if (!request) return;
  clearTimeout(request.timer);
  pending.delete(id);
  request.resolve(response);
}

function getNativePort() {
  if (nativePort) return nativePort;
  const port = chrome.runtime.connectNative(hostName);
  nativePort = port;
  port.onMessage.addListener((message) => {
    if (typeof message?.id !== "string") return;
    finish(message.id, message.ok === true
      ? { ok: true }
      : { ok: false, message: String(message.message || unavailable).slice(0, 250) });
  });
  port.onDisconnect.addListener(() => {
    void chrome.runtime.lastError;
    if (nativePort !== port) return;
    nativePort = undefined;
    for (const id of pending.keys()) finish(id, { ok: false, message: unavailable });
  });
  return port;
}

function startNative(origin) {
  return new Promise((resolve) => {
    const id = crypto.randomUUID();
    const timer = setTimeout(() => finish(id, { ok: false, message: unavailable }), 8000);
    pending.set(id, { timer, resolve });
    try {
      getNativePort().postMessage({ id, command: "start", origin });
    } catch {
      finish(id, { ok: false, message: unavailable });
    }
  });
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type !== "start-local-generator" || sender.frameId !== 0 || sender.tab?.id === undefined) {
    return false;
  }
  let origin;
  try {
    const url = new URL(sender.url);
    if (url.protocol !== "https:") return false;
    origin = url.origin;
  } catch {
    return false;
  }
  void chrome.storage.local.get("linkedOrigin").then(({ linkedOrigin }) => {
    if (origin !== linkedOrigin) return { ok: false, message: "Vincula esta página desde la extensión." };
    return startNative(origin);
  }).then(sendResponse).catch(() => sendResponse({ ok: false, message: unavailable }));
  return true;
});
