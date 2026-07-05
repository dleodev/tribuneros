import express from "express";
import cookieParser from "cookie-parser";
import crypto from "node:crypto";
import path from "node:path";
import { fileURLToPath } from "node:url";
import pg from "pg";

const { Pool } = pg;
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DIST = path.join(__dirname, "..", "dist");

const PORT = process.env.PORT || 3000;
const APP_PASSWORD = process.env.APP_PASSWORD || "";
const AUTH_ENABLED = APP_PASSWORD.length > 0;
const COOKIE = "tbx_auth";

// Token determinístico derivado de la contraseña. Si cambia APP_PASSWORD,
// las sesiones viejas dejan de valer.
const TOKEN = AUTH_ENABLED
  ? crypto.createHash("sha256").update(APP_PASSWORD).digest("hex")
  : "";

if (!process.env.DATABASE_URL) {
  console.error("Falta DATABASE_URL");
  process.exit(1);
}

const pool = new Pool({ connectionString: process.env.DATABASE_URL });

async function init() {
  await pool.query(
    "CREATE TABLE IF NOT EXISTS kv (key text PRIMARY KEY, value jsonb NOT NULL)"
  );
}

const app = express();
app.use(express.json({ limit: "5mb" }));
app.use(cookieParser());

// ---------- Auth ----------
function isAuthed(req) {
  if (!AUTH_ENABLED) return true;
  return req.cookies?.[COOKIE] === TOKEN;
}

app.post("/api/login", (req, res) => {
  if (!AUTH_ENABLED) return res.json({ ok: true });
  const { password } = req.body || {};
  if (password === APP_PASSWORD) {
    res.cookie(COOKIE, TOKEN, {
      httpOnly: true,
      sameSite: "lax",
      secure: req.secure || req.headers["x-forwarded-proto"] === "https",
      maxAge: 180 * 24 * 60 * 60 * 1000, // ~180 días
    });
    return res.json({ ok: true });
  }
  return res.status(401).json({ error: "bad password" });
});

app.get("/api/me", (req, res) => {
  if (isAuthed(req)) return res.json({ auth: true });
  return res.status(401).json({ auth: false });
});

// Middleware: protege toda la API KV.
app.use("/api/kv", (req, res, next) => {
  if (isAuthed(req)) return next();
  return res.status(401).json({ error: "unauthorized" });
});

// ---------- KV ----------
app.get("/api/kv/:key", async (req, res) => {
  try {
    const { rows } = await pool.query(
      "SELECT value #>> '{}' AS value FROM kv WHERE key = $1",
      [req.params.key]
    );
    if (!rows.length) return res.status(404).json(null);
    return res.json({ value: rows[0].value });
  } catch (e) {
    console.error("GET kv error", e);
    return res.status(500).json({ error: "db error" });
  }
});

app.put("/api/kv/:key", async (req, res) => {
  try {
    const value = req.body?.value;
    if (typeof value !== "string") {
      return res.status(400).json({ error: "value must be string" });
    }
    await pool.query(
      `INSERT INTO kv (key, value) VALUES ($1, to_jsonb($2::text))
       ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value`,
      [req.params.key, value]
    );
    return res.json({ ok: true });
  } catch (e) {
    console.error("PUT kv error", e);
    return res.status(500).json({ error: "db error" });
  }
});

// ---------- Estáticos + SPA fallback ----------
app.use(express.static(DIST));
app.get("*", (req, res) => {
  res.sendFile(path.join(DIST, "index.html"));
});

init()
  .then(() => {
    app.listen(PORT, () => {
      console.log(`Tribuneros escuchando en :${PORT} (auth ${AUTH_ENABLED ? "ON" : "OFF"})`);
    });
  })
  .catch((e) => {
    console.error("No se pudo inicializar la DB", e);
    process.exit(1);
  });
