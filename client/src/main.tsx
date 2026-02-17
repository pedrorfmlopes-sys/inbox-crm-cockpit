import React from "react";
import ReactDOM from "react-dom/client";
import App from "./ui/App";
import DialogApp from "./ui/DialogApp";

// Decide which UI to render based on URL param:
// - taskpane: main sidebar
// - dialog: Office Dialog UI (Create/Add/Edit)
function getView(): string {
  const p = new URLSearchParams(window.location.search);
  return (p.get("view") || "taskpane").toLowerCase();
}

// Tell Preflight (index.html) that React mounted, so it can auto-hide.
// This matters because Outlook sometimes reports generic "Script error." from office.js (cross-origin),
// even when our UI is 100% OK.
function markMounted() {
  try {
    document.documentElement.dataset.icccMounted = "1";
    window.dispatchEvent(new Event("iccc:mounted"));
  } catch {
    // ignore
  }
}

// Boot wrapper: if something crashes in Compose (Outlook Classic) we show the error
// instead of a white screen.
function Boot() {
  const [fatal, setFatal] = React.useState<string | null>(null);

  React.useEffect(() => {
    const onErr = (e: any) => {
      try {
        const err = e?.error || e?.reason || e;
        const msg =
          err && (err.stack || err.message)
            ? String(err.stack || err.message)
            : String(err || "Erro desconhecido");
        // Store first fatal only (keeps UI stable)
        setFatal((prev) => prev || msg);
      } catch {
        setFatal((prev) => prev || "Erro desconhecido");
      }
    };

    window.addEventListener("error", onErr);
    window.addEventListener("unhandledrejection", onErr);
    return () => {
      window.removeEventListener("error", onErr);
      window.removeEventListener("unhandledrejection", onErr);
    };
  }, []);

  if (fatal) {
    return (
      <div style={{ padding: 12, fontFamily: "system-ui, Segoe UI, Arial" }}>
        <div style={{ fontWeight: 800, marginBottom: 8 }}>⚠️ O add-in falhou ao iniciar</div>
        <div style={{ color: "#444", marginBottom: 8 }}>
          Isto acontece mais vezes no Outlook Classic em modo de resposta (Compose). Copia o erro abaixo e envia-me.
        </div>
        <pre
          style={{
            whiteSpace: "pre-wrap",
            fontSize: 12,
            background: "rgba(0,0,0,0.04)",
            padding: 10,
            borderRadius: 10,
            maxHeight: 260,
            overflow: "auto",
          }}
        >
          {fatal}
        </pre>
      </div>
    );
  }

  const view = getView();
  return view === "dialog" ? <DialogApp /> : <App />;
}

const rootEl = document.getElementById("root");

if (!rootEl) {
  throw new Error("Root element #root não existe.");
}

ReactDOM.createRoot(rootEl).render(
  <React.StrictMode>
    <Boot />
  </React.StrictMode>
);

// Next tick so DOM exists
setTimeout(markMounted, 0);
