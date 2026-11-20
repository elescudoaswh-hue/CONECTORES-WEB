export const localMode = "local";
export const serverMode = "server";
const localOrigin = "http://127.0.0.1:8766";
const localHeader = "X-Conectores-Local";
const jobId = /^[a-zA-Z0-9_-]{32}$/;

export function storedJob(value) {
  if (!value) return null;
  if (jobId.test(value)) return { id: value, mode: serverMode };
  try {
    const job = JSON.parse(value);
    if (typeof job?.id === "string" && jobId.test(job.id) && [localMode, serverMode].includes(job.mode)) {
      return job;
    }
  } catch {
    return null;
  }
  return null;
}

export async function dxfFetch(mode, path, options) {
  const request = { cache: "no-store", ...options };
  if (mode === serverMode) return fetch(path, request);
  const headers = new Headers(request.headers);
  headers.set(localHeader, "1");
  const response = await fetch(localOrigin + path, {
    ...request, headers, mode: "cors", credentials: "omit", redirect: "error", targetAddressSpace: "loopback",
  });
  if (response.headers.get(localHeader) !== "1") {
    throw new Error("La respuesta no corresponde al lanzador local de Conectores.");
  }
  return response;
}

export async function dxfApi(mode, path, options) {
  const response = await dxfFetch(mode, path, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !Object.keys(payload).length) {
    const error = new Error(payload.message || `El generador DXF no pudo procesar la solicitud (HTTP ${response.status}).`);
    error.status = response.status;
    throw error;
  }
  return payload;
}
