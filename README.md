# Sistema de Inventario DPT

Base mínima para la migración web. La aplicación legacy Python/Textual permanece sin modificaciones.

## Requisitos

- Node.js 20 o superior
- Docker y Docker Compose

## Base de datos

Copiá `backend/.env.example` como `backend/.env` sólo si necesitás cambiar los valores por defecto. Iniciá PostgreSQL desde la raíz:

```bash
docker compose up -d postgres
```

## Backend

```bash
cd backend
npm install
npm run dev
```

La API queda disponible en `http://localhost:3000`. PostgreSQL puede estar apagado: el proceso inicia igual y `/api/health` informa `database: "disconnected"`.

## Frontend

En otra terminal:

```bash
cd frontend
npm install
npm run dev
```

La pantalla queda disponible en `http://localhost:5173`. Para cambiar la URL de la API, copiá `frontend/.env.example` como `frontend/.env` y ajustá `VITE_BACKEND_URL`.

## Health checks

- Backend: `http://localhost:3000/api/health`
- Frontend: `http://localhost:5173`
