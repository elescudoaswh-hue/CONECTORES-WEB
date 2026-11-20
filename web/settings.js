let connection;

function openDatabase() {
  connection ||= new Promise((resolve, reject) => {
    const request = indexedDB.open("conectores-web", 1);
    request.onupgradeneeded = () => request.result.createObjectStore("preferencias");
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("El almacenamiento local está ocupado."));
  });
  return connection;
}

export async function loadSettings() {
  try {
    const db = await openDatabase();
    return await new Promise((resolve, reject) => {
      const request = db.transaction("preferencias").objectStore("preferencias").get("archivos");
      request.onsuccess = () => resolve(request.result || {});
      request.onerror = () => reject(request.error);
    });
  } catch {
    return {};
  }
}

export async function saveSettings(settings) {
  try {
    const db = await openDatabase();
    await new Promise((resolve, reject) => {
      const transaction = db.transaction("preferencias", "readwrite");
      transaction.objectStore("preferencias").put(settings, "archivos");
      transaction.oncomplete = resolve;
      transaction.onabort = () => reject(transaction.error);
    });
  } catch {
    // Los archivos siguen disponibles durante la visita.
  }
}
