import assert from "node:assert/strict";
import test from "node:test";
import { readDrawingFiles, scanDirectory, scanFileList } from "../web/drawings.js";

test("a large folder shows progress and opens only drawings requested later", async () => {
  const opened = [];
  const entries = Array.from({ length: 13_654 }, (_, index) => {
    const name = index % 2 ? `${index}.DWG` : `${index}.txt`;
    return {
      kind: "file", name,
      async getFile() {
        opened.push(name);
        return new File(["drawing"], name, { lastModified: index });
      },
    };
  });
  const folder = {
    name: "Conectores",
    *values() {
      yield* entries.slice(0, 4000);
      yield { name: "Respaldo", kind: "directory", values: () => entries.slice(4000).values() };
    },
  };
  const progress = [];
  const { checked, root, drawings } = await scanDirectory(folder, (count) => progress.push(count));
  assert.equal(checked, 13_654);
  assert.equal(root, "Conectores");
  assert.equal(drawings.size, 6827);
  assert.equal(progress[0], 250);
  assert.equal(progress.at(-1), 13500);
  assert.deepEqual(opened, []);
  const [file] = await readDrawingFiles(drawings, ["Conectores/Respaldo/13653.DWG"]);
  assert.equal(file.lastModified, 13653);
  assert.deepEqual(opened, ["13653.DWG"]);
});

test("the browser file selector retains nested paths and the files it already opened", async () => {
  const first = new File(["first"], "00001234.dxf", { lastModified: 20 });
  const second = new File(["second"], "00001234.DXF", { lastModified: 30 });
  Object.defineProperty(first, "webkitRelativePath", { value: "Conectores/00001234.dxf" });
  Object.defineProperty(second, "webkitRelativePath", { value: "Conectores/Otros/00001234.DXF" });
  const files = [first, ...Array.from({ length: 249 }, () => ({ name: "notas.txt" })), second];
  const progress = [];
  const folder = await scanFileList(files, (...counts) => progress.push(counts));
  assert.equal(folder.root, "Conectores");
  assert.equal(folder.checked, 251);
  assert.deepEqual(progress, [[250, 251]]);
  const fetched = await readDrawingFiles(folder.drawings, [
    "Conectores/Otros/00001234.DXF", "Conectores/00001234.dxf",
  ]);
  assert.deepEqual(fetched, [second, first]);
});
