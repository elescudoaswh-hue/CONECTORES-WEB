import { loadSettings, saveSettings } from "./settings.js";
import { scanDirectory, scanFileList, readDrawingFiles } from "./drawings.js";
import { dxfApi, dxfFetch, localMode, serverMode, storedJob } from "./dxf.js";
import { ensureLocalGenerator, showLocalGenerator } from "./local-generator.js";

const element = (id) => document.getElementById(id);
const familyInput = element("family-name");
const groupsInput = element("groups");
const outputInput = element("output-name");
const colorsDialog = element("colors-dialog");
const noticeDialog = element("notice-dialog");
const generateButton = element("generate");
const generatorMode = element("generator-mode");
const configureLocal = element("configure-local-generator");
const drawingsInput = element("drawings-folder");
const drawingsName = element("drawings-name");
const alternatePicker = element("alternate-drawings-picker");
const tabButtons = [...document.querySelectorAll('[role="tab"]')];
const taskControls = [...document.querySelectorAll(".files-panel button, .files-panel input, .files-panel select, .task-actions button")];
const jobStorageKey = "conectores-web-generacion";
const maxUpload = 94 * 1024 * 1024;

let engine;
let nextRequest = 0;
const analysisRequests = new Map();
let report;
let reportFile;
let databaseFile;
let families = [];
let assignments = {};
let drawings = null;
let folderLabel = "";
let connectorRows = [];
let selectedKeys = new Set();
let highlightedKeys = new Set();
let highlightAnchor = null;
let selectionInitialized = false;
let colorDraft = {};
let colorRows = [];
let busy = false;
let folderSelectionQueued = false;
let activeJob;
try {
  activeJob = storedJob(sessionStorage.getItem(jobStorageKey));
} catch {
  activeJob = null;
}

function runEngine(action, payload = {}, file) {
  if (!engine) {
    engine = new Worker("./analysis-worker.js");
    engine.addEventListener("message", ({ data }) => {
      const request = analysisRequests.get(data.id);
      if (!request) return;
      analysisRequests.delete(data.id);
      if (data.error) request.reject(new Error(data.error));
      else request.resolve(data.result);
    });
    engine.addEventListener("error", () => {
      const error = new Error("No se pudo iniciar el motor de análisis. Vuelve a intentarlo.");
      for (const request of analysisRequests.values()) request.reject(error);
      analysisRequests.clear();
      engine.terminate();
      engine = undefined;
    });
  }
  const id = ++nextRequest;
  return new Promise((resolve, reject) => {
    analysisRequests.set(id, { resolve, reject });
    engine.postMessage({ id, action, payload, file });
  });
}

function status(message) {
  element("status").textContent = message;
}

function setBusy(value) {
  busy = value;
  for (const control of taskControls) control.disabled = value;
  if (!value && folderSelectionQueued) {
    folderSelectionQueued = false;
    queueMicrotask(readDrawingSelection);
  }
}

function notice(title, message, { confirm, job, mode } = {}) {
  element("notice-title").textContent = title;
  element("notice-text").textContent = message;
  element("notice-cancel").hidden = !confirm;
  element("notice-ok").textContent = confirm || "Aceptar";
  element("notice-downloads").hidden = !job;
  if (job) {
    const route = `/api/jobs/${encodeURIComponent(job.id)}`;
    element("download-status").textContent = "";
    const download = element("download-dxf");
    download.href = `${route}/download`;
    download.download = job.output;
    const detail = element("download-detail");
    detail.href = `${route}/report`;
    detail.removeAttribute("download");
    detail.hidden = !job.hasDetail;
    download.onclick = detail.onclick = null;
    if (mode === localMode) {
      prepareLocalDownload(download, `${route}/download`, job.output);
      prepareLocalDownload(detail, `${route}/report`, job.output.replace(/\.dxf$/i, "") + "_REPORTE_GENERACION.txt");
    }
  }
  noticeDialog.returnValue = "";
  return new Promise((resolve) => {
    noticeDialog.addEventListener("close", () => resolve(noticeDialog.returnValue === "accept"), { once: true });
    noticeDialog.showModal();
    if (confirm) element("notice-cancel").focus();
  });
}

async function runTask(title, task) {
  if (busy) return;
  setBusy(true);
  try {
    await task();
  } catch (error) {
    const message = error.message || "No se pudo completar la operación.";
    status(message);
    await notice(title, message);
  } finally {
    setBusy(false);
  }
}

function remember() {
  return saveSettings({
    reportFile, databaseFile, family: familyInput.value.trim(),
    groups: groupsInput.value, output: outputInput.value.trim(), generatorMode: generatorMode.value,
  });
}

function activateTab(tab, focus = false) {
  for (const button of tabButtons) {
    const active = button === tab;
    button.tabIndex = active ? 0 : -1;
    button.setAttribute("aria-selected", String(active));
    element(button.getAttribute("aria-controls")).hidden = !active;
  }
  if (focus) tab.focus();
}

function renderCells(row, values) {
  for (const value of values) row.insertCell().textContent = value;
}

function refreshSelection() {
  for (const row of connectorRows) {
    row.checkbox.checked = selectedKeys.has(row.key);
    row.node.classList.toggle("highlighted", highlightedKeys.has(row.key));
  }
  element("selection-count").textContent = `${selectedKeys.size} de ${connectorRows.length} seleccionados`;
  element("select-all").disabled = connectorRows.length === 0;
  element("select-none").disabled = connectorRows.length === 0;
  element("select-highlighted").disabled = highlightedKeys.size === 0;
}

function resetAnalysis() {
  connectorRows = [];
  selectedKeys.clear();
  highlightedKeys.clear();
  highlightAnchor = null;
  selectionInitialized = false;
  element("summary-rows").replaceChildren();
  element("connector-rows").replaceChildren();
  element("summary-empty").hidden = false;
  element("connectors-empty").hidden = false;
  element("validation-log").textContent = "Carga el reporte y ejecuta Analizar y contar para ver la validación.";
  refreshSelection();
}

function displayFamilyList() {
  element("family-options").replaceChildren(...families.map((name) => {
    const option = document.createElement("option");
    option.value = name;
    return option;
  }));
}

function reportStem() {
  return reportFile?.name.replace(/\.xlsx$/i, "").split(/[-_]InitialReport/i)[0] || "Conectores";
}

function fileStem(name, fallback) {
  return name.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_").replace(/[. ]+$/g, "").slice(0, 175) || fallback;
}

function defaultOutput() {
  return `${fileStem(reportStem(), "Conectores")}_ANALIZADO.dxf`;
}

async function applySelectedFamily(showDetails = false) {
  if (!report) throw new Error("Primero carga el Initial Report.");
  if (!databaseFile) throw new Error("Primero selecciona una base JSON.");
  const family = familyInput.value.trim();
  if (!family) throw new Error("Selecciona una familia de la base JSON.");
  const result = await runEngine("apply_family", { family, groups: groupsInput.value });
  assignments = result.assignments;
  groupsInput.value = result.groups;
  resetAnalysis();
  status(`Familia ${family}: ${report.parts.length - result.missing.length} de ${report.parts.length} números asignados desde la base`);
  void remember();
  if (!showDetails) return;
  if (result.missing.length) {
    const conflicts = Object.entries(result.conflicts).map(([part, values]) => `${part}: ${values.join(" / ")}`);
    const details = conflicts.length ? `\n\nConflictos encontrados\n${conflicts.join("\n")}` : "";
    await notice("Asignación incompleta", `Pendientes: ${result.missing.join(", ")}${details}\n\nUsa Configurar colores para completar los pendientes.`);
  } else {
    await notice("Familia aplicada", `Se aplicaron todos los colores de ${family}.\n\nPlantillas: ${result.groups}`);
  }
}

async function suggestFamily(matches, preferred) {
  const family = families.includes(preferred) ? preferred : matches.length === 1 ? matches[0] : "";
  if (family) {
    familyInput.value = family;
    await applySelectedFamily();
  } else if (databaseFile) {
    status(matches.length > 1
      ? "Varias familias coinciden. Selecciona una y aplícala."
      : "Selecciona una familia o configura los colores pendientes.");
  }
}

async function openReport(file, { restore = false, preferred } = {}) {
  status("Leyendo Initial Report…");
  const parsed = await runEngine("report", { name: file.name }, file);
  report = parsed;
  reportFile = file;
  element("report-name").value = file.name;
  assignments = parsed.assignments;
  resetAnalysis();
  if (!databaseFile && !restore) familyInput.value = reportStem();
  if (!outputInput.value.trim()) outputInput.value = defaultOutput();
  status(`Reporte cargado: ${report.parts.length} configuraciones, ${report.connectorCount} registros de componentes`);
  await suggestFamily(parsed.matching, preferred);
  void remember();
  if (!databaseFile && !restore) openColors();
}

async function openDatabase(file, { restore = false } = {}) {
  status("Leyendo base de familias…");
  const parsed = await runEngine("families", {}, file);
  databaseFile = file;
  element("database-name").value = file.name;
  families = parsed.families;
  displayFamilyList();
  if (!families.includes(familyInput.value)) familyInput.value = "";
  status(`Base cargada: ${families.length} familias. Selecciona una y presiona Aplicar familia.`);
  if (report) await suggestFamily(parsed.matching);
  if (!restore) void remember();
}

function folderProgress(checked, total) {
  const count = checked.toLocaleString();
  const progress = total === undefined ? count : `${count} de ${total.toLocaleString()}`;
  drawingsName.value = `Leyendo carpeta… ${progress} archivos revisados`;
  status(`Leyendo carpeta de conectores. Archivos revisados: ${progress}.`);
}

function saveDrawings(folder) {
  drawings = folder.drawings;
  folderLabel = `${folder.root} · ${drawings.size.toLocaleString()} dibujos DWG / DXF`;
  drawingsName.value = folderLabel;
  if (drawings.size) {
    status(`Carpeta lista: ${drawings.size.toLocaleString()} dibujos en ${folder.checked.toLocaleString()} archivos revisados. Pulsa Analizar y contar.`);
  } else {
    status(`Se revisaron ${folder.checked.toLocaleString()} archivos. Elige una carpeta que contenga dibujos .dwg o .dxf.`);
  }
}

function pickDrawingsWithFileInput() {
  if (busy) return;
  drawingsInput.value = "";
  drawingsName.value = "Esperando la confirmación del navegador…";
  status("Selecciona la carpeta de conectores y confirma Cargar.");
  drawingsInput.click();
}

function readDrawingSelection() {
  if (busy) {
    folderSelectionQueued = true;
    return;
  }
  const files = drawingsInput.files;
  void runTask("Carpeta DWG", async () => {
    try {
      if (!files.length) throw new Error("El navegador devolvió una selección vacía. Elige la carpeta que contiene los dibujos.");
      folderProgress(0, files.length);
      saveDrawings(await scanFileList(files, folderProgress));
    } catch (error) {
      drawingsName.value = folderLabel;
      throw error;
    } finally {
      drawingsInput.value = "";
    }
  });
}

function pickDrawings() {
  if (!window.isSecureContext || window.top !== window || typeof window.showDirectoryPicker !== "function") {
    pickDrawingsWithFileInput();
    return;
  }
  void runTask("Carpeta DWG", async () => {
    drawingsName.value = "Esperando la carpeta…";
    status("Selecciona la carpeta donde están los dibujos de los conectores.");
    let directory;
    try {
      directory = await window.showDirectoryPicker({ mode: "read", id: "conectores-dwg" });
      folderProgress(0);
      saveDrawings(await scanDirectory(directory, folderProgress));
      alternatePicker.hidden = true;
    } catch (error) {
      drawingsName.value = folderLabel;
      alternatePicker.hidden = false;
      if (!directory && error.name === "AbortError") {
        status("La selección se cerró. Puedes elegir la carpeta de nuevo o usar Selector alternativo.");
        return;
      }
      throw new Error("El navegador no pudo leer la carpeta de conectores. Comprueba que esté accesible o usa Selector alternativo.", { cause: error });
    }
  });
}

function openColors() {
  if (!report) {
    void notice("Initial Report", "Primero carga el Initial Report.");
    return;
  }
  colorDraft = { ...assignments };
  element("all-parts").checked = false;
  element("all-parts").indeterminate = false;
  element("custom-color").hidden = true;
  element("custom-color-name").value = "";
  const body = element("color-rows");
  body.replaceChildren();
  colorRows = report.parts.map((part) => {
    const row = body.insertRow();
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.setAttribute("aria-label", `Seleccionar ${part}`);
    checkbox.addEventListener("change", refreshColorSelections);
    row.insertCell().append(checkbox);
    renderCells(row, [part, report.headers[part] || part]);
    const color = row.insertCell();
    color.textContent = colorDraft[part] || "—";
    return { part, checkbox, color };
  });
  const palette = element("palette");
  palette.replaceChildren();
  for (const [name, [background, foreground]] of Object.entries(report.palette)) {
    const button = paletteButton(name, () => paintSelected(name), "palette-color");
    button.style.setProperty("--swatch", background);
    button.style.setProperty("--swatch-text", foreground);
    palette.append(button);
  }
  palette.append(
    paletteButton("Nuevo color", () => {
      if (!colorRows.some((row) => row.checkbox.checked)) {
        colorStatus("Selecciona al menos un número o código.", true);
        return;
      }
      element("custom-color").hidden = false;
      element("custom-color-name").focus();
    }),
    paletteButton("Limpiar", () => paintSelected(""), "danger"),
  );
  updateColorCount();
  colorsDialog.showModal();
}

function paletteButton(label, action, modifier = "") {
  const button = document.createElement("button");
  button.type = "button";
  button.className = `button ${modifier}`.trim();
  button.textContent = label;
  button.addEventListener("click", action);
  return button;
}

function refreshColorSelections() {
  const count = colorRows.filter((row) => row.checkbox.checked).length;
  const selectAll = element("all-parts");
  selectAll.checked = count === colorRows.length;
  selectAll.indeterminate = count > 0 && count < colorRows.length;
}

function colorStatus(message, error = false) {
  const target = element("color-status");
  target.textContent = message;
  target.toggleAttribute("data-error", error);
}

function updateColorCount() {
  colorStatus(`Asignados: ${report.parts.filter((part) => colorDraft[part]).length} de ${report.parts.length}`);
}

function paintSelected(color) {
  const checked = colorRows.filter((row) => row.checkbox.checked);
  if (!checked.length) {
    colorStatus("Selecciona al menos un número o código.", true);
    return;
  }
  for (const row of checked) {
    colorDraft[row.part] = color;
    row.color.textContent = color || "—";
  }
  updateColorCount();
}

async function saveColors() {
  const button = element("save-colors");
  button.disabled = true;
  try {
    const result = await runEngine("apply_colors", { assignments: colorDraft, groups: groupsInput.value });
    assignments = result.assignments;
    groupsInput.value = result.groups;
    resetAnalysis();
    status(`Colores configurados: ${[...new Set(Object.values(assignments))].sort().join(", ")}`);
    colorsDialog.close();
    void remember();
  } catch (error) {
    colorStatus(error.message, true);
  } finally {
    button.disabled = false;
  }
}

function highlight(row, event) {
  if (event.shiftKey && highlightAnchor !== null) {
    const [start, end] = [highlightAnchor, row.index].sort((a, b) => a - b);
    if (!event.ctrlKey && !event.metaKey) highlightedKeys.clear();
    for (const item of connectorRows.slice(start, end + 1)) highlightedKeys.add(item.key);
  } else if (event.ctrlKey || event.metaKey) {
    if (!highlightedKeys.delete(row.key)) highlightedKeys.add(row.key);
    highlightAnchor = row.index;
  } else {
    highlightedKeys = new Set([row.key]);
    highlightAnchor = row.index;
  }
  refreshSelection();
}

function renderAnalysis(result) {
  const summary = element("summary-rows");
  summary.replaceChildren();
  for (const values of result.summary) renderCells(summary.insertRow(), values);
  element("summary-empty").hidden = result.summary.length !== 0;

  const body = element("connector-rows");
  body.replaceChildren();
  const keys = new Set(result.connectors.map((row) => row.key));
  selectedKeys = selectionInitialized ? new Set([...selectedKeys].filter((key) => keys.has(key))) : new Set(keys);
  highlightedKeys = new Set([...highlightedKeys].filter((key) => keys.has(key)));
  highlightAnchor = null;
  selectionInitialized = true;
  connectorRows = result.connectors.map((row, index) => {
    const node = body.insertRow();
    node.classList.toggle("missing", row.missing);
    node.tabIndex = -1;
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.setAttribute("aria-label", `Generar ${row.name} en ${row.group}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) selectedKeys.add(row.key);
      else selectedKeys.delete(row.key);
      refreshSelection();
    });
    node.insertCell().append(checkbox);
    renderCells(node, [row.group, row.name, row.component, row.kind, row.drawing || (row.missing ? "FALTANTE" : "Sin escanear")]);
    const item = { ...row, node, checkbox, index };
    node.addEventListener("click", (event) => {
      if (event.target === checkbox) return;
      node.focus({ preventScroll: true });
      highlight(item, event);
    });
    node.addEventListener("keydown", (event) => {
      if (event.target !== node || event.key !== " ") return;
      event.preventDefault();
      const keysToToggle = highlightedKeys.size ? highlightedKeys : new Set([row.key]);
      for (const key of keysToToggle) {
        if (!selectedKeys.delete(key)) selectedKeys.add(key);
      }
      refreshSelection();
    });
    return item;
  });
  element("connectors-empty").hidden = connectorRows.length !== 0;
  element("validation-log").textContent = result.validation;
  refreshSelection();
  status(`Análisis listo: ${connectorRows.length} posiciones; ${result.missingCount} conectores faltantes (${result.missingDrawings} dibujos únicos)`);
}

function missingPreview(rows, count) {
  const lines = rows.slice(0, count).map(([group, name, component]) => `[${group}] ${name} — ${component}`);
  if (rows.length > count) lines.push(`… y ${rows.length - count} conectores más.`);
  return lines.join("\n");
}

async function analyze({ notify = false, exportFamily } = {}) {
  if (!report) throw new Error("Primero carga el Initial Report.");
  status("Analizando y contando conectores…");
  const payload = {
    assignments, groups: groupsInput.value, exportFamily,
    drawings: drawings ? [...drawings].map(([path, entry]) => ({ path, modified: entry.lastModified ?? null })) : null,
  };
  let result = await runEngine("analyze", payload);
  if (result.drawingsNeedingDates) {
    const paths = result.drawingsNeedingDates;
    status("Comparando las fechas de los dibujos duplicados…");
    const files = await readDrawingFiles(drawings, paths);
    const dates = new Map(paths.map((path, index) => [path, files[index].lastModified]));
    for (const drawing of payload.drawings) {
      if (dates.has(drawing.path)) drawing.modified = dates.get(drawing.path);
    }
    result = await runEngine("analyze", payload);
  }
  renderAnalysis(result);
  void remember();
  if (notify && drawings) {
    if (result.missingCount) {
      const reportText = `Faltan ${result.missingCount} conectores (${result.missingDrawings} dibujos únicos).\n\n`
        + missingPreview(result.missing, 35) + "\n\nLa lista completa está en la pestaña Validación.";
      await notice("Conectores faltantes", reportText);
    } else {
      await notice("Análisis completo", "Se encontraron todos los conectores requeridos.");
    }
  }
  return result;
}

async function exportDatabase() {
  const family = familyInput.value.trim() || "FAMILIA";
  const result = await analyze({ exportFamily: family });
  const blob = new Blob([JSON.stringify(result.database, null, 2) + "\n"], { type: "application/json" });
  const filename = `Base_${fileStem(family, "FAMILIA")}.json`;
  downloadBlob(blob, filename);
  status(`Base guardada: ${filename}`);
}

function downloadBlob(blob, filename) {
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  link.click();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function prepareLocalDownload(link, path, filename) {
  link.href = "#";
  link.removeAttribute("download");
  link.onclick = async (event) => {
    event.preventDefault();
    if (link.getAttribute("aria-disabled") === "true") return;
    const progress = element("download-status");
    link.setAttribute("aria-disabled", "true");
    progress.textContent = "Preparando la descarga…";
    try {
      const response = await dxfFetch(localMode, path);
      if (!response.ok) throw new Error("El archivo ya no está disponible en el lanzador.");
      downloadBlob(await response.blob(), filename);
      progress.textContent = "";
    } catch {
      progress.textContent = `No se pudo descargar ${filename}. Comprueba que el lanzador esté abierto.`;
    } finally {
      link.removeAttribute("aria-disabled");
    }
  };
}

function setActiveJob(job) {
  activeJob = job;
  generateButton.textContent = job ? "Consultar DXF" : "3. Generar DXF";
  try {
    if (job) sessionStorage.setItem(jobStorageKey, JSON.stringify(job));
    else sessionStorage.removeItem(jobStorageKey);
  } catch {
    // Se conserva el trabajo durante la visita.
  }
}

async function watchJob(job, mode) {
  let retries = 0;
  while (job.status === "queued" || job.status === "running") {
    status(job.message);
    await new Promise((resolve) => setTimeout(resolve, retries ? 3000 : 1200));
    try {
      job = await dxfApi(mode, `/api/jobs/${encodeURIComponent(job.id)}`);
      retries = 0;
    } catch (error) {
      if (error.status === 404) {
        setActiveJob(null);
        throw error;
      }
      if (++retries === 3) throw new Error("Se perdió la conexión con el generador. Pulsa Consultar DXF para retomar la espera.");
    }
  }
  setActiveJob(null);
  if (job.status !== "completed") throw new Error(job.message);
  status(job.message);
  const reportText = [job.summary, ...job.messages];
  if (job.extraMessages) reportText.push(`… y ${job.extraMessages} avisos más en el reporte.`);
  await notice("DXF generado", reportText.join("\n"), { job, mode });
}

async function generateDxf() {
  if (activeJob) {
    const { id, mode } = activeJob;
    if (mode === localMode && !(await ensureLocalGenerator())) return;
    await watchJob({ id, status: "running", message: "Consultando el trabajo de AutoCAD…" }, mode);
    return;
  }
  const mode = generatorMode.value;
  if (!drawings) throw new Error("Selecciona la carpeta con los conectores.");
  await analyze();
  const selected = connectorRows.filter((row) => selectedKeys.has(row.key));
  if (!selected.length) {
    activateTab(element("tab-connectors"));
    throw new Error("Selecciona al menos un conector para generar el DXF.");
  }
  const missing = selected.filter((row) => row.missing);
  if (missing.length) {
    const rows = missing.map((row) => [row.group, row.name, row.component]);
    const approved = await notice("Conectores faltantes",
      `Faltan ${missing.length} conectores. Si continúas, quedarán marcados en rojo en el DXF.\n\n${missingPreview(rows, 20)}`,
      { confirm: "Generar DXF" });
    if (!approved) return;
  }
  status("Preparando los dibujos seleccionados…");
  const paths = [...new Set(selected.map((row) => row.drawing).filter(Boolean))];
  const files = await readDrawingFiles(drawings, paths);
  if (reportFile.size + files.reduce((total, file) => total + file.size, 0) > maxUpload) {
    throw new Error("Los archivos superan el límite del generador. Divide la selección de conectores en varias generaciones.");
  }
  if (mode === localMode) {
    if (!(await ensureLocalGenerator())) {
      status("Generación DXF cancelada.");
      return;
    }
  } else {
    const health = await dxfApi(mode, "/api/health");
    if (!health.autocad) throw new Error("El generador en red requiere Windows con AutoCAD y pywin32.");
  }
  let output = outputInput.value.trim() || defaultOutput();
  if (!output.toLowerCase().endsWith(".dxf")) output += ".dxf";
  outputInput.value = output;
  const body = new FormData();
  body.append("report", reportFile);
  body.append("options", JSON.stringify({
    assignments, groups: groupsInput.value, output,
    selected: selected.map((row) => row.key), allowMissing: missing.length > 0,
  }));
  for (const file of files) body.append("drawings", file, file.name);
  status(mode === localMode ? "Iniciando AutoCAD en esta PC…" : "Enviando los dibujos al generador en red…");
  const job = await dxfApi(mode, "/api/jobs", { method: "POST", body });
  setActiveJob({ id: job.id, mode });
  void remember();
  await watchJob(job, mode);
}

for (const tab of tabButtons) {
  tab.addEventListener("click", () => activateTab(tab));
  tab.addEventListener("keydown", (event) => {
    const current = tabButtons.indexOf(tab);
    const destination = {
      ArrowLeft: (current + tabButtons.length - 1) % tabButtons.length,
      ArrowRight: (current + 1) % tabButtons.length, Home: 0, End: tabButtons.length - 1,
    }[event.key];
    if (destination === undefined) return;
    event.preventDefault();
    activateTab(tabButtons[destination], true);
  });
}

for (const button of document.querySelectorAll("[data-picker]")) {
  button.addEventListener("click", () => element(button.dataset.picker).click());
}
element("report-file").addEventListener("change", (event) => {
  const file = event.target.files[0];
  event.target.value = "";
  if (file) void runTask("Error al leer reporte", () => openReport(file));
});
element("database-file").addEventListener("change", (event) => {
  const file = event.target.files[0];
  event.target.value = "";
  if (file) void runTask("Base de familias", () => openDatabase(file));
});
element("select-drawings").addEventListener("click", pickDrawings);
alternatePicker.addEventListener("click", pickDrawingsWithFileInput);
drawingsInput.addEventListener("change", readDrawingSelection);
drawingsInput.addEventListener("cancel", () => {
  folderSelectionQueued = false;
  drawingsName.value = folderLabel;
  status("El navegador cerró la selección de carpeta. Puedes volver a intentarlo.");
});
element("reset-output").addEventListener("click", () => {
  outputInput.value = defaultOutput();
  void remember();
});
for (const input of [familyInput, groupsInput, outputInput]) input.addEventListener("change", remember);
generatorMode.addEventListener("change", () => {
  configureLocal.hidden = generatorMode.value !== localMode;
  void remember();
});
configureLocal.addEventListener("click", () => void showLocalGenerator());

element("apply-family").addEventListener("click", () => void runTask("Base de familias", () => applySelectedFamily(true)));
element("configure-colors").addEventListener("click", openColors);
element("analyze").addEventListener("click", () => void runTask("Configuración incompleta", () => analyze({ notify: true })));
element("export-json").addEventListener("click", () => void runTask("Guardar base JSON", exportDatabase));
generateButton.addEventListener("click", () => void runTask("Generación DXF", generateDxf));
element("select-all").addEventListener("click", () => {
  selectedKeys = new Set(connectorRows.map((row) => row.key));
  refreshSelection();
});
element("select-none").addEventListener("click", () => {
  selectedKeys.clear();
  refreshSelection();
});
element("select-highlighted").addEventListener("click", () => {
  selectedKeys = new Set(highlightedKeys);
  refreshSelection();
});
element("all-parts").addEventListener("change", (event) => {
  for (const row of colorRows) row.checkbox.checked = event.target.checked;
});
element("save-colors").addEventListener("click", saveColors);
element("colors-form").addEventListener("submit", (event) => {
  if (event.submitter?.value === "cancel") return;
  event.preventDefault();
});
element("assign-custom-color").addEventListener("click", async () => {
  try {
    const color = await runEngine("normalize_color", { color: element("custom-color-name").value });
    if (!color || /[+;-]/.test(color)) {
      colorStatus("Usa un nombre de color sin +, - ni punto y coma.", true);
      return;
    }
    paintSelected(color);
    element("custom-color").hidden = true;
  } catch (error) {
    colorStatus(error.message, true);
  }
});

async function restore() {
  setBusy(true);
  try {
    const saved = await loadSettings();
    familyInput.value = saved.family || "";
    groupsInput.value = saved.groups ?? groupsInput.value;
    outputInput.value = saved.output || "";
    generatorMode.value = saved.generatorMode === serverMode ? serverMode : localMode;
    configureLocal.hidden = generatorMode.value !== localMode;
    if (saved.databaseFile) {
      try {
        await openDatabase(saved.databaseFile, { restore: true });
      } catch (error) {
        status(`No se pudo restaurar la base de familias. ${error.message}`);
      }
    }
    if (saved.reportFile) {
      try {
        await openReport(saved.reportFile, { restore: true, preferred: saved.family });
      } catch (error) {
        status(`No se pudo restaurar el Initial Report. ${error.message}`);
      }
    }
  } finally {
    setBusy(false);
  }
  setActiveJob(activeJob);
  if (activeJob?.mode === serverMode) {
    void runTask("Generación DXF", generateDxf);
  }
}

void restore();
