import React, { useState } from "react";
import { SCENARIOS } from "./testScenarios";
import type { OutlookMessageContext } from "../office";
import type { LinkEntry, OdooMeta } from "../api";

export default function DebugPanel({
  ctx,
  links,
  meta,
  compact,
}: {
  ctx?: OutlookMessageContext;
  links?: LinkEntry[];
  meta?: OdooMeta | null;
  compact?: boolean;
}) {
  const { report } = useTestSuite();

  return (
    <div style={S.wrap}>
      <details style={S.details} open={false}>
        <summary style={S.summary} title="Ver detalhes técnicos (debug)">
          Debug {report && "(TEST FAILURES DETECTED)"}
        </summary>

        {/* Simulation Suite buttons removed for production - usable via internal triggers */}

        {report && (
          <div style={S.reportBlock}>
            <div style={{ color: "#d32f2f", fontWeight: 800, fontSize: 11, marginBottom: 4 }}>DETAILED ERROR REPORT</div>
            <pre style={S.reportPre}>{report}</pre>
          </div>
        )}

        <div style={S.block}>
          <div style={S.h}>Contexto do email</div>
          <pre style={{ ...S.pre, maxHeight: compact ? 120 : 220 }}>
            {JSON.stringify(ctx ?? {}, null, 2)}
          </pre>
        </div>

        <div style={S.block}>
          <div style={S.h}>Links</div>
          <pre style={{ ...S.pre, maxHeight: compact ? 120 : 220 }}>
            {JSON.stringify(links ?? [], null, 2)}
          </pre>
        </div>

        <div style={S.block}>
          <div style={S.h}>Odoo meta</div>
          <pre style={{ ...S.pre, maxHeight: compact ? 120 : 220 }}>
            {JSON.stringify(meta ?? {}, null, 2)}
          </pre>
        </div>
      </details>
    </div>
  );
}

function useTestSuite() {
  const [report, setReport] = useState<string | null>(null);

  const runSuite = async () => {
    setReport(null);
    let logs: string[] = [`[${new Date().toLocaleTimeString()}] Starting Full Test Suite...\n`];

    for (const scenario of SCENARIOS) {
      logs.push(`Running ${scenario.id}: ${scenario.name}...`);

      try {
        // 1. Inject mock data into bridge (localStorage)
        localStorage.setItem("ic_bridge_body", scenario.bodyText || "");
        localStorage.setItem("ic_bridge_atts", JSON.stringify(scenario.attachments || []));

        // 2. Trigger Scenario Change in DialogApp
        window.dispatchEvent(new CustomEvent("iccc:run-scenario", { detail: scenario }));

        // 3. Wait for UI to react
        await new Promise(r => setTimeout(r, 2000));

        // 4. Basic automated checks
        const hasHorizontalScroll = document.documentElement.scrollWidth > document.documentElement.clientWidth;
        if (hasHorizontalScroll) logs.push(`> [FAIL] Horizontal Overflow detected (Scrollbar visible).`);

        // Check for pill width adherence (94px)
        const pills = Array.from(document.querySelectorAll('button')).filter(b => b.innerText.length > 0 && b.offsetWidth > 0);
        const oversizedPill = pills.find(p => p.offsetWidth > 100); // Allow some margin for small browsers, but 94px is target
        if (oversizedPill) logs.push(`> [WARN] Pill "${oversizedPill.innerText.substring(0, 10)}..." exceeds 94px limit (${oversizedPill.offsetWidth}px).`);

        // Mocking checks (in real life we'd check DOM elements)
        const text = document.body.innerText.toUpperCase();
        const hasAssistant = text.includes("ASSISTENTE IA") || text.includes("A ANALISAR");
        if (!hasAssistant && scenario.expectedResults.aiTriggers) {
          logs.push(`> [FAIL] AI Assistant did not trigger automatically.`);
        }

        logs.push(`> ${scenario.id} sequence finished.\n`);
      } catch (e: any) {
        logs.push(`> [ERROR] ${e.message}`);
      }
    }

    logs.push(`Test Suite Completed.`);
    setReport(logs.join("\n"));
  };

  return { report, setReport, runSuite };
}

const S: Record<string, React.CSSProperties> = {
  wrap: { marginTop: 10 },
  details: {
    border: "1px solid rgba(11,45,107,0.12)",
    borderRadius: 12,
    background: "rgba(255,255,255,0.85)",
    padding: 8,
  },
  summary: {
    cursor: "pointer",
    fontWeight: 700,
    fontSize: 11,
    color: "#0b2d6b",
    userSelect: "none",
  },
  testBtn: {
    padding: "4px 10px",
    background: "#0b2d6b",
    color: "#fff",
    border: "none",
    borderRadius: 8,
    fontSize: 10,
    fontWeight: 700,
    cursor: "pointer",
  },
  reportBlock: {
    marginTop: 8,
    padding: 8,
    background: "#FFEBEE",
    borderRadius: 8,
    border: "1px solid #FFCDD2"
  },
  reportPre: {
    margin: 0,
    fontSize: 10,
    whiteSpace: "pre-wrap",
    maxHeight: 200,
    overflow: "auto",
    color: "#B71C1C"
  },
  block: { marginTop: 10 },
  h: { fontWeight: 700, fontSize: 11, marginBottom: 6, color: "rgba(11,45,107,0.75)" },
  pre: {
    margin: 0,
    padding: 8,
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.10)",
    background: "#fff",
    color: "#111",
    fontSize: 10,
    overflow: "auto",
  },
};
