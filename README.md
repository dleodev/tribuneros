# Tribuneros

Panel de gestión de e-commerce (ventas, egresos, proveedores, stock, clientes).
Frontend React (Vite) + backend Express con almacenamiento clave-valor en
PostgreSQL. Datos **persistentes** y compartidos entre dispositivos.

## Arquitectura

- **Frontend**: React + Vite (`src/`). El panel es el artefacto original sin
  cambios de lógica; `src/storage.js` define `window.storage` y lo conecta a la API.
- **Backend**: Express (`server/index.js`). Sirve el build y expone:
  - `GET /api/kv/:key` → `{ value }` | `404`
  - `PUT /api/kv/:key` con body `{ value: "<string>" }`
  - `POST /api/login`, `GET /api/me` (login mínimo con `APP_PASSWORD`)
- **DB**: PostgreSQL, tabla única `kv (key text primary key, value jsonb)`. Se crea
  sola al arrancar.

## Variables de entorno

| Variable       | Descripción                                   |
| -------------- | --------------------------------------------- |
| `DATABASE_URL` | Conexión a PostgreSQL                         |
| `PORT`         | Puerto del server (default `3000`)            |
| `APP_PASSWORD` | Contraseña única. Vacía = app abierta sin login |

## Desarrollo local

```bash
npm install
cp .env.example .env   # completá DATABASE_URL y APP_PASSWORD
# Terminal 1: backend
npm run dev:server
# Terminal 2: frontend (Vite con proxy /api -> :3000)
npm run dev
```

## Producción (Docker)

```bash
docker build -t tribuneros .
docker run -p 3000:3000 \
  -e DATABASE_URL=postgres://... \
  -e APP_PASSWORD=... \
  tribuneros
```

El deploy real se hace en Easypanel (proyecto propio + Postgres 1-clic + build por
Dockerfile). Ver la guía del plan.
