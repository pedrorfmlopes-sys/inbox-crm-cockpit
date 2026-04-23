import { runGroupsV1BrowserValidation } from "./runtimeValidation";

declare global {
  interface Window {
    __GROUPS_VALIDATION_RESULT__?: {
      done: boolean;
      ok: boolean;
      report?: Awaited<ReturnType<typeof runGroupsV1BrowserValidation>>;
      error?: string;
    };
  }
}

function renderState(title: string, body: string, tone: "ok" | "error" | "running") {
  const root = document.getElementById("root");
  if (!root) return;
  root.innerHTML = `
    <main style="font-family: Segoe UI, Arial, sans-serif; padding: 24px; color: #243244; background: #f6f8fb; min-height: 100vh;">
      <div style="max-width: 1080px; margin: 0 auto; display: grid; gap: 16px;">
        <header style="display: grid; gap: 6px;">
          <div style="font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.08em; color: #64748b;">Groups v1 validation</div>
          <h1 style="margin: 0; font-size: 26px;">${title}</h1>
        </header>
        <section style="border-radius: 16px; padding: 16px 18px; background: ${
          tone === "error" ? "rgba(254,242,242,0.96)" : tone === "ok" ? "rgba(239,246,255,0.96)" : "rgba(255,255,255,0.96)"
        }; border: 1px solid ${
          tone === "error" ? "rgba(220,38,38,0.18)" : tone === "ok" ? "rgba(59,130,246,0.18)" : "rgba(148,163,184,0.18)"
        };">
          <pre style="white-space: pre-wrap; margin: 0; font-size: 13px; line-height: 1.5;">${body}</pre>
        </section>
      </div>
    </main>
  `;
}

async function boot() {
  renderState("A correr", "A executar a matriz deterministica do modulo Groups...", "running");
  try {
    const report = await runGroupsV1BrowserValidation();
    const failedScenarios = report.scenarios.filter((scenario) => scenario.status === "failed");
    window.__GROUPS_VALIDATION_RESULT__ = {
      done: true,
      ok: failedScenarios.length === 0,
      report,
    };
    const summary = [
      `Gerado: ${report.generatedAtIso}`,
      `Passou: ${report.passed}`,
      `Falhou: ${report.failed}`,
      "",
      ...report.scenarios.map((scenario) => `[${scenario.status.toUpperCase()}] ${scenario.id} — ${scenario.details}`),
    ].join("\n");
    renderState(
      failedScenarios.length ? "Validacao com falhas" : "Validacao concluida",
      summary,
      failedScenarios.length ? "error" : "ok"
    );
  } catch (error) {
    const message = error instanceof Error ? `${error.message}\n\n${error.stack || ""}` : String(error || "Erro desconhecido");
    window.__GROUPS_VALIDATION_RESULT__ = {
      done: true,
      ok: false,
      error: message,
    };
    renderState("Falha na validacao", message, "error");
  }
}

void boot();
