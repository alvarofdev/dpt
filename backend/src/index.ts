import cors from "cors";
import dotenv from "dotenv";
import express from "express";
import { Pool } from "pg";

dotenv.config();

const port = Number(process.env.PORT ?? 3000);
const frontendUrl = process.env.FRONTEND_URL ?? "http://localhost:5173";
const databaseUrl = process.env.DATABASE_URL ?? "postgresql://inventario:inventario@localhost:5432/inventario";

const pool = new Pool({ connectionString: databaseUrl });
const app = express();

app.use(cors({ origin: frontendUrl }));
app.use(express.json());

app.get("/api/health", async (_request, response) => {
  let database: "connected" | "disconnected" = "disconnected";

  try {
    await pool.query("SELECT 1");
    database = "connected";
  } catch {
    database = "disconnected";
  }

  response.json({ service: "backend", status: "ok", database });
});

const server = app.listen(port, () => {
  console.log(`Backend escuchando en http://localhost:${port}`);
});

const shutdown = async (): Promise<void> => {
  await pool.end();
  server.close();
};

process.on("SIGINT", () => void shutdown());
process.on("SIGTERM", () => void shutdown());
