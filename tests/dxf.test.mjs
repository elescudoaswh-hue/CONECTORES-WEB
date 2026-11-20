import assert from "node:assert/strict";
import test from "node:test";
import { dxfApi, dxfFetch, localMode, serverMode, storedJob } from "../web/dxf.js";

test("jobs remember which generator owns them, including saved jobs from older versions", () => {
  const id = "a".repeat(32);
  assert.deepEqual(storedJob(id), { id, mode: serverMode });
  const localJob = { id, mode: localMode };
  assert.deepEqual(storedJob(JSON.stringify(localJob)), localJob);
  assert.equal(storedJob(JSON.stringify({ id, mode: "unknown" })), null);
  assert.equal(storedJob("../bravo"), null);
});

test("the local generator receives drawing data only on the loopback address", async (t) => {
  const body = new FormData();
  body.append("drawings", new File(["drawing"], "F-02.dwg"));
  const mock = t.mock.method(globalThis, "fetch", async (path, request) => {
    assert.equal(path, "http://127.0.0.1:8766/api/jobs");
    assert.equal(request.credentials, "omit");
    assert.equal(request.redirect, "error");
    assert.equal(request.targetAddressSpace, "loopback");
    assert.equal(request.headers.get("X-Conectores-Local"), "1");
    assert.equal(request.body, body);
    return Response.json({ status: "queued" }, { status: 202, headers: { "X-Conectores-Local": "1" } });
  });
  assert.deepEqual(await dxfApi(localMode, "/api/jobs", { method: "POST", body }), { status: "queued" });
  assert.equal(mock.mock.callCount(), 1);
});

test("the network generator keeps using the Cloudflare Worker API", async (t) => {
  t.mock.method(globalThis, "fetch", async (path, request) => {
    assert.equal(path, "/api/health");
    assert.equal(request.credentials, undefined);
    return Response.json({ autocad: true });
  });
  assert.equal((await dxfApi(serverMode, "/api/health")).autocad, true);
});

test("an unrelated service on the local port cannot be used as the generator", async (t) => {
  t.mock.method(globalThis, "fetch", async () => Response.json({ autocad: true }));
  await assert.rejects(dxfFetch(localMode, "/api/health"), /no corresponde al lanzador local/);
});
