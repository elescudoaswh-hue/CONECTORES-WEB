const jobPath = /^\/api\/jobs\/[\w-]{32}(?:\/(download|report))?$/;

function error(message, status) {
  return Response.json({ message }, { status, headers: { "Cache-Control": "no-store" } });
}

function backendUrl(origin) {
  try {
    const url = new URL(origin);
    const local = url.protocol === "http:" && ["127.0.0.1", "localhost", "[::1]"].includes(url.hostname);
    if (url.username || url.password || (!local && url.protocol !== "https:")) return null;
    return url;
  } catch {
    return null;
  }
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!url.pathname.startsWith("/api/")) return env.ASSETS.fetch(request);
    if (url.pathname === "/api/jobs") {
      if (request.method !== "POST") return error("Método no admitido.", 405);
      const origin = request.headers.get("Origin");
      if (origin && origin !== url.origin) return error("Solicitud de otro sitio.", 403);
    } else {
      if (url.pathname !== "/api/health" && !jobPath.test(url.pathname)) {
        return error("Ruta del generador no disponible.", 404);
      }
      if (!["GET", "HEAD"].includes(request.method)) return error("Método no admitido.", 405);
    }

    const target = backendUrl(env.DXF_API_ORIGIN);
    if (!target || !env.DXF_API_TOKEN || target.origin === url.origin) {
      return error("El generador DXF aún no está conectado a este sitio.", 503);
    }
    target.pathname = url.pathname;
    target.search = "";
    const requestHeaders = new Headers({ Authorization: `Bearer ${env.DXF_API_TOKEN}` });
    for (const name of ["Accept", "Content-Type", "Range", "If-Range"]) {
      if (request.headers.has(name)) requestHeaders.set(name, request.headers.get(name));
    }
    if (env.CF_ACCESS_CLIENT_ID && env.CF_ACCESS_CLIENT_SECRET) {
      requestHeaders.set("CF-Access-Client-Id", env.CF_ACCESS_CLIENT_ID);
      requestHeaders.set("CF-Access-Client-Secret", env.CF_ACCESS_CLIENT_SECRET);
    }

    try {
      const upstream = await fetch(target, {
        method: request.method,
        headers: requestHeaders,
        body: request.body,
        redirect: "manual",
      });
      if (upstream.status >= 300 && upstream.status < 400) {
        return error("El enlace con el generador devolvió una redirección. Revisa el túnel y su acceso.", 502);
      }
      if (upstream.status === 401 || upstream.status === 403) {
        return error("El sitio no pudo autenticarse con el generador DXF.", 502);
      }
      if (upstream.status === 413) return error("Los archivos de generación superan el límite de carga.", 413);
      if (upstream.status >= 500 && !upstream.headers.get("Content-Type")?.includes("application/json")) {
        return error("No se pudo conectar con el generador DXF de Windows.", 502);
      }
      const responseHeaders = new Headers({ "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" });
      for (const name of ["Content-Type", "Content-Disposition", "Accept-Ranges", "Content-Range"]) {
        if (upstream.headers.has(name)) responseHeaders.set(name, upstream.headers.get(name));
      }
      return new Response(upstream.body, { status: upstream.status, headers: responseHeaders });
    } catch {
      return error("No se pudo conectar con el generador DXF de Windows.", 502);
    }
  },
};
