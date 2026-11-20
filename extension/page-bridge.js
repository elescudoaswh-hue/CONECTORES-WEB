(() => {
  if (globalThis.conectoresPageBridge || !document.querySelector('meta[name="conectores-yazaki"][content="web"]')) return;
  globalThis.conectoresPageBridge = true;

  window.addEventListener("message", async (event) => {
    if (event.source !== window || event.origin !== window.location.origin) return;
    const message = event.data;
    if (message?.source !== "conectores-web" || message.type !== "start-local-generator") return;
    if (typeof message.id !== "string" || !/^[a-f0-9-]{36}$/.test(message.id)) return;

    const reply = (stage, result = {}) => window.postMessage({
      source: "conectores-extension", id: message.id, stage,
      ok: result.ok === true, message: result.message,
    }, window.location.origin);
    reply("received");
    try {
      reply("finished", await chrome.runtime.sendMessage({ type: "start-local-generator" }));
    } catch {
      reply("finished", { message: "Abre la extensión y vincula de nuevo esta página." });
    }
  });
})();
