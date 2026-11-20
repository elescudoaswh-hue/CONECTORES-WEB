import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import { runInNewContext } from "node:vm";
import { startWithExtension } from "../web/extension-bridge.js";

const origin = "https://conectores.example";
const bridge = await readFile(new URL("../extension/page-bridge.js", import.meta.url), "utf8");
const worker = await readFile(new URL("../extension/background.js", import.meta.url), "utf8");

test("the page can request an on-demand start through the approved tab and receives an acknowledgement", async (t) => {
  const handlers = new Set();
  const browserWindow = {
    location: new URL(origin),
    addEventListener: (type, handler) => handlers.add(handler),
    removeEventListener: (type, handler) => handlers.delete(handler),
    postMessage(data, target) {
      assert.equal(target, origin);
      queueMicrotask(() => {
        for (const handler of handlers) handler({ source: browserWindow, origin, data });
      });
    },
  };
  let receive, onNativeMessage;
  const calls = [];
  const port = {
    onMessage: { addListener: (handler) => { onNativeMessage = handler; } },
    onDisconnect: { addListener() {} },
    postMessage(data) {
      calls.push(data);
      queueMicrotask(() => onNativeMessage({ id: data.id, ok: true }));
    },
  };
  const mockChrome = {
    storage: { local: { async get() { return { linkedOrigin: origin }; } } },
    runtime: {
      connectNative(name) {
        assert.equal(name, "com.conectores_yazaki.generador");
        return port;
      },
      onMessage: { addListener: (listener) => { receive = listener; } },
      sendMessage(message) {
        return new Promise((resolve) => receive(message, { frameId: 0, tab: { id: 1 }, url: origin }, resolve));
      },
    },
  };
  Object.defineProperty(globalThis, "window", { value: browserWindow, configurable: true });
  t.after(() => { delete globalThis.window; });
  runInNewContext(worker, { chrome: mockChrome, crypto, URL, setTimeout, clearTimeout });
  runInNewContext(bridge, {
    window: browserWindow, chrome: mockChrome,
    document: { querySelector: () => true },
  });
  assert.deepEqual(await startWithExtension(), { ok: true, message: "" });
  assert.deepEqual(Object.keys(calls[0]).sort(), ["command", "id", "origin"]);
  assert.equal(calls[0].command, "start");
  assert.equal(calls[0].origin, origin);
  assert.equal(handlers.size, 1);

  const outside = await new Promise((resolve) => receive(
    { type: "start-local-generator", origin }, { tab: { id: 1 }, frameId: 0, url: "https://distinta.example" }, resolve,
  ));
  assert.equal(outside.ok, false);
  assert.equal(calls.length, 1);
  assert.equal(receive({ type: "start-local-generator" }, { tab: { id: 1 }, frameId: 1, url: origin }, () => {}), false);
});
