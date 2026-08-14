import { useEffect, useState, type ReactElement } from "react";

type HealthResponse = {
  service: string;
  status: string;
  database: "connected" | "disconnected";
};

const backendUrl = import.meta.env.VITE_BACKEND_URL ?? "http://localhost:3000";

export function App(): ReactElement {
  const [health, setHealth] = useState<HealthResponse | null>(null);

  useEffect(() => {
    const loadHealth = async (): Promise<void> => {
      try {
        const response = await fetch(`${backendUrl}/api/health`);
        if (!response.ok) return;
        setHealth((await response.json()) as HealthResponse);
      } catch {
        setHealth(null);
      }
    };

    void loadHealth();
  }, []);

  return (
    <main className="shell">
      <section className="card" aria-labelledby="title">
        <p className="eyebrow">DPT / INVENTARIO</p>
        <h1 id="title">La nueva base web está lista.</h1>
        <p className="description">
          React, Express y PostgreSQL preparados para construir el sistema de inventario.
        </p>
        <div className="status" role="status">
          <span className={`indicator ${health ? "online" : "pending"}`} />
          {health ? `Backend conectado · Base de datos ${health.database}` : "Esperando al backend..."}
        </div>
      </section>
    </main>
  );
}
