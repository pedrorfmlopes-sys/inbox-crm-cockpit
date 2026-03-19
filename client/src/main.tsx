import React from "react";
import ReactDOM from "react-dom/client";
import UniversalApp from "@/ui/UniversalApp";
import DialogApp from "@/ui/DialogApp";
import GroupExplorerApp from "@/modules/crm/GroupExplorerApp";
import GroupManagerApp from "@/modules/crm/GroupManagerApp";

const WARM_BOOT_STORAGE_KEY = "iccc_warm_boot_v1";

// Decide which UI to render based on URL param:
// - taskpane: main sidebar
// - dialog: Office Dialog UI (Create/Add/Edit)
function getView(): string {
  const p = new URLSearchParams(window.location.search);
  return (p.get("view") || "taskpane").toLowerCase();
}

// Tell Preflight (index.html) that React mounted, so it can auto-hide.
function markMounted() {
  try {
    document.documentElement.dataset.icccMounted = "1";
    sessionStorage.setItem(WARM_BOOT_STORAGE_KEY, String(Date.now()));
    window.dispatchEvent(new Event("iccc:mounted"));
  } catch {
    // ignore
  }
}

// Boot wrapper
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
        <div style={{ fontWeight: 800, marginBottom: 8 }}>⚠️ O add-in falhou ao iniciar (v9.6)</div>
        <div style={{ color: "#444", marginBottom: 8 }}>
          Tenta carregar no botão direito e "Recarregar".
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
  if (view === "dialog") return <DialogApp />;
  if (view === "group-explorer") return <GroupExplorerApp />;
  if (view === "group-manager") return <GroupManagerApp />;
  return <UniversalApp />;
}

const rootEl = document.getElementById("root");
if (!rootEl) throw new Error("Root element #root não existe.");

const root = ReactDOM.createRoot(rootEl);

function renderApp() {
  if ((window as any).__ICCC_BOOTED__) return;
  (window as any).__ICCC_BOOTED__ = true;
  console.log("[main] Rendering React App...");
  root.render(
    <React.StrictMode>
      <Boot />
    </React.StrictMode>
  );
  setTimeout(markMounted, 0);
}

// Safety: wait for Office.js handshake. 
// We NO LONGER use a force-boot timeout because accessing Office APIs before ready crashes the host.
const OfficeAny = (window as any).Office;

if (OfficeAny) {
  console.log("[main] Office found, waiting for onReady...");

  // FAILSAFE: If Office.onReady hangs (common in some Outlook versions), 
  // we MUST boot React anyway so the user sees the app (even if limited).
  // Otherwise they are stuck on "Preflight".
  const bootTimer = setTimeout(() => {
    console.warn("[main] Office.onReady took too long (>5s). Forcing boot.");
    const statusEl = document.getElementById("pf-status");
    if (statusEl) statusEl.textContent = "Office.onReady demorou. A forçar arranque...";
    renderApp();
  }, 5000);

  // Also support legacy initialize via a global callback (from index.html)
  (window as any).__ICCC_SET_READY__ = () => {
    console.log("[main] Legacy initialize callback triggered.");
    clearTimeout(bootTimer);
    renderApp();
  };

  OfficeAny.onReady((info: any) => {
    clearTimeout(bootTimer);
    console.log("[main] Office.onReady resolved. Host:", info?.host);
    renderApp();
  });
} else {
  console.log("[main] No Office object (probably browser). Booting automatically.");
  renderApp();
}
