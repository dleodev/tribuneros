// Shim de window.storage: reemplaza la API clave-valor del artefacto original
// por llamadas a nuestro backend (/api/kv/:key). Respeta EXACTAMENTE la interfaz
// que usan los helpers load/save del artefacto:
//   - get(key)  -> devuelve { value: string } | null
//   - set(key, value) -> guarda (value siempre es string)
// Al importar este módulo se define window.storage. No se toca la lógica de App.

async function get(key) {
  const r = await fetch(`/api/kv/${encodeURIComponent(key)}`, {
    credentials: "same-origin",
  });
  if (r.status === 401) {
    // Sesión no válida: recargar para mostrar el muro de contraseña.
    location.reload();
    return null;
  }
  if (!r.ok) return null; // 404 incluido -> load usa el valor por defecto
  return await r.json(); // { value: "<string>" } | null
}

async function set(key, value) {
  await fetch(`/api/kv/${encodeURIComponent(key)}`, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    credentials: "same-origin",
    body: JSON.stringify({ value }),
  });
}

window.storage = { get, set };
