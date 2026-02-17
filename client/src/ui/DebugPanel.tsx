import React from "react";
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
  return (
    <div style={S.wrap}>
      <details style={S.details} open={false}>
        <summary style={S.summary} title="Ver detalhes técnicos (debug)">
          Debug
        </summary>

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
