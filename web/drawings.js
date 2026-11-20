const drawingExtension = /\.(dwg|dxf)$/i;
const scanBatchSize = 250;
const fileBatchSize = 24;
const yieldToBrowser = () => new Promise((resolve) => setTimeout(resolve, 0));

async function* directoryFiles(directory, parent = directory.name) {
  for await (const entry of directory.values()) {
    const path = `${parent}/${entry.name}`;
    if (entry.kind === "directory") yield* directoryFiles(entry, path);
    if (entry.kind === "file") yield [path, entry];
  }
}

export async function scanDirectory(directory, onProgress) {
  const drawings = new Map();
  let checked = 0;
  for await (const [path, handle] of directoryFiles(directory)) {
    if (drawingExtension.test(handle.name)) drawings.set(path, handle);
    checked += 1;
    if (checked % scanBatchSize === 0) {
      onProgress(checked);
      await yieldToBrowser();
    }
  }
  return { root: directory.name, drawings, checked };
}

export async function scanFileList(files, onProgress) {
  const drawings = new Map();
  const total = files.length;
  const firstPath = files[0]?.webkitRelativePath;
  const root = firstPath ? firstPath.split("/")[0] : "Archivos seleccionados";
  for (let index = 0; index < total; index += 1) {
    const file = files[index];
    if (drawingExtension.test(file.name)) drawings.set(file.webkitRelativePath || file.name, file);
    if ((index + 1) % scanBatchSize === 0) {
      onProgress(index + 1, total);
      await yieldToBrowser();
    }
  }
  return { root, drawings, checked: total };
}

export async function readDrawingFiles(drawings, paths) {
  const files = [];
  for (let offset = 0; offset < paths.length; offset += fileBatchSize) {
    const batch = paths.slice(offset, offset + fileBatchSize);
    const opened = await Promise.all(batch.map(async (path) => {
      try {
        const entry = drawings.get(path);
        if (!entry) throw new Error("Dibujo ausente de la selección.");
        return typeof entry.getFile === "function" ? await entry.getFile() : entry;
      } catch (error) {
        throw new Error(`No se pudo leer «${path}». Revisa que el archivo esté disponible y vuelve a seleccionar la carpeta.`, { cause: error });
      }
    }));
    files.push(...opened);
  }
  return files;
}
