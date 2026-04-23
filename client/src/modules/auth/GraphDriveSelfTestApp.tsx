import React, { useEffect, useMemo, useRef, useState } from "react";
import { getSettings } from "@/settings";
import {
  GRAPH_DRIVE_SELF_TEST_SCOPES,
  requestCockpitHostAction,
  runGraphDriveWriteSelfTest,
  type GraphDriveSelfTestResult,
} from "@/office";
import { applySkin } from "@/ui/skins";
import { PanelState } from "@/ui/PanelState";
import "../../global.css";

type ViewState = {
  running: boolean;
  result: GraphDriveSelfTestResult | null;
  technicalError: string;
};

declare global {
  interface Window {
    __GRAPH_DRIVE_SELF_TEST_RESULT__?: {
      done: boolean;
      started: boolean;
      result: GraphDriveSelfTestResult | null;
      technicalError: string;
    };
  }
}

function writeWindowResult(state: ViewState) {
  window.__GRAPH_DRIVE_SELF_TEST_RESULT__ = {
    done: !state.running && Boolean(state.result || state.technicalError),
    started: state.running || Boolean(state.result || state.technicalError),
    result: state.result,
    technicalError: state.technicalError,
  };
}

function formatJson(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value ?? "");
  }
}

function toneFromConclusion(result: GraphDriveSelfTestResult | null): "info" | "success" | "error" {
  if (!result) return "info";
  if (result.conclusion === "tenant_allows_user_write") return "success";
  return "error";
}

export default function GraphDriveSelfTestApp(): JSX.Element {
  const scopes = useMemo(() => [...GRAPH_DRIVE_SELF_TEST_SCOPES], []);
  const [state, setState] = useState<ViewState>({
    running: false,
    result: null,
    technicalError: "",
  });
  const runCounterRef = useRef(0);

  useEffect(() => {
    let alive = true;
    void (async () => {
      try {
        const settings = await getSettings();
        if (!alive) return;
        applySkin(settings.skinId || "classic");
      } catch {
        applySkin("classic");
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  useEffect(() => {
    writeWindowResult(state);
  }, [state]);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (closed) return;
    window.close();
  }

  async function handleRun() {
    const runId = runCounterRef.current + 1;
    runCounterRef.current = runId;
    const nextRunningState: ViewState = {
      running: true,
      result: null,
      technicalError: "",
    };
    setState(nextRunningState);

    try {
      const result = await runGraphDriveWriteSelfTest();
      if (runCounterRef.current !== runId) return;
      setState({
        running: false,
        result,
        technicalError: "",
      });
    } catch (error) {
      if (runCounterRef.current !== runId) return;
      setState({
        running: false,
        result: null,
        technicalError: error instanceof Error ? error.message : String(error || "Erro tecnico desconhecido."),
      });
    }
  }

  return (
    <div style={styles.root} data-testid="graph-drive-self-test-root">
      <div style={styles.headerCard}>
        <div style={styles.headerCopy}>
          <div style={styles.kicker}>Graph Self-Test</div>
          <div style={styles.title}>Teste minimo de escrita em OneDrive</div>
          <div style={styles.subtitle}>
            Este self-test pede os scopes configurados, tenta `GET /me/drive`, cria a pasta de teste
            e limpa no fim. Sem Graph productization, sem mexer no resto da app.
          </div>
        </div>
        <div style={styles.headerActions}>
          <button type="button" style={styles.secondaryBtn} onClick={() => void handleClose()}>
            Fechar
          </button>
          <button
            type="button"
            style={styles.primaryBtn}
            onClick={() => void handleRun()}
            disabled={state.running}
            data-testid="graph-drive-self-test-run"
          >
            {state.running ? "A executar..." : "Executar self-test"}
          </button>
        </div>
      </div>

      <div style={styles.scopeCard}>
        <div style={styles.sectionTitle}>Scopes pedidos</div>
        <div style={styles.scopeWrap}>
          {scopes.map((scope) => (
            <span key={scope} style={styles.scopePill}>{scope}</span>
          ))}
        </div>
      </div>

      {state.running ? (
        <PanelState
          compact
          tone="loading"
          title="Self-test em execucao"
          description="A pedir consentimento/token e a testar o drive do utilizador."
        />
      ) : null}

      {state.technicalError ? (
        <PanelState compact tone="error" title="Falha tecnica no self-test" description={state.technicalError} />
      ) : null}

      {state.result ? (
        <>
          <PanelState
            compact
            tone={toneFromConclusion(state.result)}
            title={`Conclusao: ${state.result.conclusion}`}
            description={state.result.conclusionMessage}
          />

          <div style={styles.grid}>
            <section style={styles.card}>
              <div style={styles.sectionTitle}>Consentimento / Token</div>
              <dl style={styles.definitionList}>
                <div style={styles.definitionRow}><dt>Modo</dt><dd>{state.result.authMode}</dd></div>
                <div style={styles.definitionRow}><dt>Resultado</dt><dd>{state.result.consent.result}</dd></div>
                <div style={styles.definitionRow}><dt>Conta</dt><dd>{state.result.consent.account || "--"}</dd></div>
                <div style={styles.definitionRow}><dt>Erro</dt><dd>{state.result.consent.errorMessage || "--"}</dd></div>
                <div style={styles.definitionRow}><dt>Codigo</dt><dd>{state.result.consent.errorCode || "--"}</dd></div>
              </dl>
            </section>

            <section style={styles.card}>
              <div style={styles.sectionTitle}>Graph /me/drive</div>
              <dl style={styles.definitionList}>
                <div style={styles.definitionRow}><dt>Status</dt><dd>{String(state.result.meDrive.status ?? "--")}</dd></div>
                <div style={styles.definitionRow}><dt>OK</dt><dd>{state.result.meDrive.ok ? "sim" : "nao"}</dd></div>
                <div style={styles.definitionRow}><dt>Erro</dt><dd>{state.result.meDrive.errorMessage || "--"}</dd></div>
              </dl>
              <pre style={styles.pre}>{formatJson(state.result.meDrive.response)}</pre>
            </section>

            <section style={styles.card}>
              <div style={styles.sectionTitle}>Criacao da pasta de teste</div>
              <dl style={styles.definitionList}>
                <div style={styles.definitionRow}><dt>Status</dt><dd>{String(state.result.createFolder.status ?? "--")}</dd></div>
                <div style={styles.definitionRow}><dt>OK</dt><dd>{state.result.createFolder.ok ? "sim" : "nao"}</dd></div>
                <div style={styles.definitionRow}><dt>Pasta</dt><dd>{state.result.createFolder.folderName || "--"}</dd></div>
                <div style={styles.definitionRow}><dt>Folder ID</dt><dd>{state.result.createFolder.folderId || "--"}</dd></div>
                <div style={styles.definitionRow}><dt>Erro</dt><dd>{state.result.createFolder.errorMessage || "--"}</dd></div>
              </dl>
              <pre style={styles.pre}>{formatJson(state.result.createFolder.response)}</pre>
            </section>

            <section style={styles.card}>
              <div style={styles.sectionTitle}>Limpeza</div>
              <dl style={styles.definitionList}>
                <div style={styles.definitionRow}><dt>Attempted</dt><dd>{state.result.cleanup.attempted ? "sim" : "nao"}</dd></div>
                <div style={styles.definitionRow}><dt>OK</dt><dd>{state.result.cleanup.ok ? "sim" : "nao"}</dd></div>
                <div style={styles.definitionRow}><dt>Status</dt><dd>{String(state.result.cleanup.status ?? "--")}</dd></div>
                <div style={styles.definitionRow}><dt>Detalhe</dt><dd>{state.result.cleanup.detail || state.result.cleanup.errorMessage || "--"}</dd></div>
              </dl>
            </section>
          </div>
        </>
      ) : (
        <PanelState
          compact
          tone="info"
          title="Aguardando execucao"
          description="O self-test ainda nao foi corrido neste snapshot."
        />
      )}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: {
    minHeight: "100vh",
    boxSizing: "border-box",
    padding: "20px",
    display: "grid",
    gap: "14px",
    background: "var(--iccc-bg)",
    color: "var(--iccc-text)",
    fontFamily: "var(--iccc-font, 'Segoe UI', sans-serif)",
  },
  headerCard: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: "16px",
    padding: "16px 18px",
    borderRadius: "18px",
    border: "1px solid rgba(15,23,42,0.08)",
    background: "rgba(255,255,255,0.92)",
    boxShadow: "0 12px 30px rgba(15,23,42,0.06)",
  },
  headerCopy: {
    display: "grid",
    gap: "4px",
    minWidth: 0,
  },
  kicker: {
    fontSize: "10px",
    fontWeight: 800,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    color: "#64748b",
  },
  title: {
    fontSize: "24px",
    fontWeight: 800,
    color: "#0f172a",
  },
  subtitle: {
    maxWidth: "860px",
    fontSize: "13px",
    lineHeight: 1.5,
    color: "#475569",
  },
  headerActions: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
  },
  primaryBtn: {
    borderRadius: "999px",
    border: "1px solid rgba(37, 99, 235, 0.28)",
    background: "linear-gradient(180deg, rgba(59,130,246,0.96) 0%, rgba(29,78,216,0.9) 100%)",
    color: "#ffffff",
    fontSize: "12px",
    fontWeight: 800,
    padding: "10px 16px",
    cursor: "pointer",
  },
  secondaryBtn: {
    borderRadius: "999px",
    border: "1px solid rgba(15,23,42,0.12)",
    background: "#ffffff",
    color: "#0f172a",
    fontSize: "12px",
    fontWeight: 700,
    padding: "9px 14px",
    cursor: "pointer",
  },
  scopeCard: {
    padding: "14px 16px",
    borderRadius: "16px",
    border: "1px solid rgba(15,23,42,0.08)",
    background: "rgba(255,255,255,0.9)",
    display: "grid",
    gap: "10px",
  },
  sectionTitle: {
    fontSize: "15px",
    fontWeight: 800,
    color: "#0f172a",
  },
  scopeWrap: {
    display: "flex",
    flexWrap: "wrap",
    gap: "8px",
  },
  scopePill: {
    borderRadius: "999px",
    padding: "6px 10px",
    fontSize: "11px",
    fontWeight: 700,
    color: "#0f172a",
    background: "rgba(15,23,42,0.05)",
    border: "1px solid rgba(15,23,42,0.08)",
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
    gap: "14px",
  },
  card: {
    padding: "14px 16px",
    borderRadius: "16px",
    border: "1px solid rgba(15,23,42,0.08)",
    background: "rgba(255,255,255,0.92)",
    boxShadow: "0 10px 26px rgba(15,23,42,0.05)",
    display: "grid",
    gap: "12px",
    minWidth: 0,
  },
  definitionList: {
    display: "grid",
    gap: "8px",
    margin: 0,
  },
  definitionRow: {
    display: "grid",
    gridTemplateColumns: "110px minmax(0, 1fr)",
    gap: "10px",
    fontSize: "12px",
    lineHeight: 1.45,
  },
  pre: {
    margin: 0,
    padding: "10px 12px",
    borderRadius: "12px",
    background: "rgba(15,23,42,0.04)",
    border: "1px solid rgba(15,23,42,0.06)",
    fontSize: "11px",
    lineHeight: 1.5,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    overflowX: "auto",
  },
};
