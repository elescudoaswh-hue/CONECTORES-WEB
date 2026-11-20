export function startWithExtension() {
  if (window.location.protocol !== "https:") return Promise.resolve(null);
  const id = crypto.randomUUID();
  const origin = window.location.origin;
  return new Promise((resolve) => {
    let acknowledged = false;
    let timer = setTimeout(() => finish(null), 500);
    function finish(result) {
      clearTimeout(timer);
      window.removeEventListener("message", receive);
      resolve(result);
    }
    function receive(event) {
      if (event.source !== window || event.origin !== origin) return;
      const { data } = event;
      if (data?.source !== "conectores-extension" || data.id !== id) return;
      if (data.stage === "received" && !acknowledged) {
        acknowledged = true;
        clearTimeout(timer);
        timer = setTimeout(() => finish({
          ok: false, message: "La extensión tardó en conectar con el lanzador local.",
        }), 9000);
      }
      if (data.stage === "finished" && acknowledged) {
        finish({ ok: data.ok === true, message: typeof data.message === "string" ? data.message : "" });
      }
    }
    window.addEventListener("message", receive);
    window.postMessage({ source: "conectores-web", type: "start-local-generator", id }, origin);
  });
}
