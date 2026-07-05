import { useState, useEffect } from "react";

// Muro de contraseña mínimo. Consulta GET /api/me:
//  - { auth: true }  -> muestra la app (children)
//  - 401             -> muestra el formulario de login
// Si el server no tiene APP_PASSWORD configurada, /api/me siempre da auth:true.

const V = "#7C5CFC";
const VD = "#5B3FD6";
const FONT = "'Inter', system-ui, -apple-system, sans-serif";

export default function Gate({ children }) {
  const [status, setStatus] = useState("loading"); // loading | ok | login
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);

  const check = async () => {
    try {
      const r = await fetch("/api/me", { credentials: "same-origin" });
      setStatus(r.ok ? "ok" : "login");
    } catch {
      setStatus("login");
    }
  };

  useEffect(() => {
    check();
  }, []);

  const submit = async (e) => {
    e.preventDefault();
    if (busy) return;
    setBusy(true);
    setError("");
    try {
      const r = await fetch("/api/login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        credentials: "same-origin",
        body: JSON.stringify({ password }),
      });
      if (r.ok) {
        setStatus("ok");
      } else {
        setError("Contraseña incorrecta.");
      }
    } catch {
      setError("No se pudo conectar. Probá de nuevo.");
    } finally {
      setBusy(false);
    }
  };

  if (status === "loading") {
    return (
      <div style={wrap}>
        <div style={{ color: V, fontWeight: 700 }}>Cargando…</div>
      </div>
    );
  }

  if (status === "login") {
    return (
      <div style={wrap}>
        <form onSubmit={submit} style={card}>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
            <div style={logo}>T</div>
            <div style={{ fontWeight: 800, fontSize: 20, color: "#1A1A2E" }}>Tribuneros</div>
          </div>
          <label style={{ fontSize: 12, color: "#6B6B80", fontWeight: 600, display: "block", marginBottom: 6 }}>
            Contraseña
          </label>
          <input
            type="password"
            autoFocus
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            placeholder="Ingresá la contraseña"
            style={input}
          />
          {error && <div style={{ color: "#EB5757", fontSize: 12.5, marginTop: 8 }}>{error}</div>}
          <button type="submit" disabled={busy} style={{ ...btn, opacity: busy ? 0.6 : 1 }}>
            {busy ? "Entrando…" : "Entrar"}
          </button>
        </form>
      </div>
    );
  }

  return children;
}

const wrap = {
  minHeight: "100vh",
  background: "#F5F5FA",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  fontFamily: FONT,
  padding: 16,
};
const card = {
  background: "#fff",
  borderRadius: 18,
  border: "1px solid #ECECF3",
  boxShadow: "0 8px 30px rgba(20,20,50,0.08)",
  padding: 28,
  width: "100%",
  maxWidth: 360,
};
const logo = {
  width: 34,
  height: 34,
  borderRadius: 10,
  background: "linear-gradient(135deg,#7C5CFC,#B9A6FF)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  color: "#fff",
  fontWeight: 800,
};
const input = {
  width: "100%",
  boxSizing: "border-box",
  background: "#FAFAFD",
  border: "1px solid #ECECF3",
  borderRadius: 10,
  color: "#1A1A2E",
  fontSize: 14,
  padding: "10px 12px",
  fontFamily: FONT,
  outline: "none",
};
const btn = {
  marginTop: 16,
  width: "100%",
  background: VD,
  color: "#fff",
  border: "none",
  borderRadius: 10,
  fontSize: 14,
  fontWeight: 700,
  padding: "11px 16px",
  cursor: "pointer",
  fontFamily: FONT,
};
