import React, { useEffect, useRef, useState } from "react";
import * as pdfjsLib from "pdfjs-dist";
import { type RelatedEmailEntry } from "@/api";
import { type PreviewMode, type AttachmentPreviewState } from "../types";
import { dataUrlToUint8Array } from "../previewUtils";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

export interface PreviewPaneProps {
  previewShellStyle: React.CSSProperties;
  previewMode: PreviewMode;
  setPreviewMode: (mode: PreviewMode) => void;
  previewHtml: string;
  previewHasDocument: boolean;
  selectedEmail: RelatedEmailEntry | null;
  selectedAttachmentPreview: any;
  selectedAttachmentDocumentPreview: AttachmentPreviewState | null;
  selectedAttachmentPreviewRemoteStatus: string;
  selectedAttachmentPreviewMode: string;
  handlePreviewReply: () => Promise<void>;
  handlePreviewForward: () => Promise<void>;
}

export function StudioPdfPreview({ dataUrl, title }: { dataUrl: string; title: string }) {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const [status, setStatus] = useState<"loading" | "ready" | "error">("loading");

  useEffect(() => {
    let cancelled = false;
    const host = hostRef.current;
    if (!host || !dataUrl) {
      setStatus("error");
      return;
    }

    host.innerHTML = "";
    setStatus("loading");

    (async () => {
      try {
        const loadingTask = pdfjsLib.getDocument({ data: dataUrlToUint8Array(dataUrl) });
        const pdf = await loadingTask.promise;
        if (cancelled) {
          void loadingTask.destroy();
          return;
        }

        const nextPageCount = pdf.numPages || 0;

        for (let pageNumber = 1; pageNumber <= nextPageCount; pageNumber += 1) {
          if (cancelled) break;
          const page = await pdf.getPage(pageNumber);
          const viewport = page.getViewport({ scale: 1.15 });
          const canvas = document.createElement("canvas");
          canvas.style.display = "block";
          canvas.style.width = "100%";
          canvas.style.maxWidth = `${Math.ceil(viewport.width)}px`;
          canvas.style.height = "auto";
          canvas.style.margin = pageNumber === nextPageCount ? "0 auto" : "0 auto 12px auto";
          canvas.style.background = "#fff";
          canvas.style.borderRadius = "8px";
          canvas.style.boxShadow = "0 6px 16px rgba(15,23,42,0.08)";
          const context = canvas.getContext("2d", { alpha: false });
          if (!context) continue;
          canvas.width = Math.ceil(viewport.width);
          canvas.height = Math.ceil(viewport.height);
          host.appendChild(canvas);
          await page.render({
            canvasContext: context,
            viewport,
            canvas: canvas as any,
          }).promise;
        }

        if (!cancelled) setStatus("ready");
      } catch (error) {
        console.warn("[classification-studio] pdf preview failed", error);
        if (!cancelled) setStatus("error");
      }
    })();

    return () => {
      cancelled = true;
      if (hostRef.current) hostRef.current.innerHTML = "";
    };
  }, [dataUrl]);

  if (status === "error") {
    return <div style={S.attachmentPreviewEmpty}>Este PDF foi detetado, mas nao foi possivel renderiza-lo dentro do add-in.</div>;
  }

  return (
    <div style={S.attachmentPdfPreviewShell} aria-label={title}>
      {status === "loading" ? (
        <div style={S.attachmentPdfPreviewLoading}>A carregar PDF...</div>
      ) : null}
      <div
        ref={hostRef}
        style={{
          ...S.attachmentPdfPreviewCanvasHost,
          display: status === "loading" ? "none" : S.attachmentPdfPreviewCanvasHost.display,
        }}
      />
    </div>
  );
}

const PreviewPane: React.FC<PreviewPaneProps & React.HTMLAttributes<HTMLElement>> = (props) => {
  const {
    previewShellStyle,
    previewMode,
    setPreviewMode,
    previewHtml,
    previewHasDocument,
    selectedEmail,
    selectedAttachmentPreview,
    selectedAttachmentDocumentPreview,
    selectedAttachmentPreviewRemoteStatus,
    selectedAttachmentPreviewMode,
    handlePreviewReply,
    handlePreviewForward,
    ...rest
  } = props;

  return (
    <section style={previewShellStyle} {...rest}>
      <div style={S.previewToolbar}>
        <button type="button" style={previewMode === "email" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("email")} disabled={!previewHtml}>Email</button>
        <button type="button" style={previewMode === "document" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("document")} disabled={!previewHasDocument}>Documento</button>
        <button type="button" style={previewMode === "reply" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("reply")} disabled={!selectedEmail}>Responder</button>
        <button type="button" style={previewMode === "forward" ? S.previewTabOn : S.previewTab} onClick={() => setPreviewMode("forward")} disabled={!selectedEmail}>Reencaminhar</button>
      </div>
      <div style={S.previewBody}>
        {previewMode === "email" ? (
          previewHtml ? (
            <div style={S.previewHtml} dangerouslySetInnerHTML={{ __html: previewHtml }} />
          ) : (
            <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado suficiente para preview." />
          )
        ) : null}
        {previewMode === "document" ? (
          selectedAttachmentPreview ? (
            <div style={S.documentPreviewShell}>
              {selectedAttachmentDocumentPreview?.kind === "image" ? (
                <div style={S.documentPreviewFrame}>
                  <img src={selectedAttachmentDocumentPreview.src!} alt={selectedAttachmentPreview?.name || "Imagem"} style={S.attachmentPreviewImage} />
                </div>
              ) : null}
              {selectedAttachmentDocumentPreview?.kind === "pdf" ? (
                <div style={S.documentPreviewFrame}>
                  {selectedAttachmentDocumentPreview.src!.startsWith("data:")
                    ? <StudioPdfPreview dataUrl={selectedAttachmentDocumentPreview.src!} title={selectedAttachmentPreview?.name || "PDF"} />
                    : <iframe title={selectedAttachmentPreview?.name || "PDF"} src={selectedAttachmentDocumentPreview.src!} style={S.documentPreviewIframe} />}
                </div>
              ) : null}
              {selectedAttachmentDocumentPreview?.kind === "office" ? (
                <div style={S.documentPreviewFrame}>
                  <iframe title={selectedAttachmentPreview?.name || "Documento"} src={selectedAttachmentDocumentPreview.url!} style={S.documentPreviewIframe} />
                </div>
              ) : null}
              {selectedAttachmentDocumentPreview?.kind === "text" ? (
                <div style={S.documentPreviewFrame}>
                  <pre style={S.attachmentPreviewText}>{selectedAttachmentDocumentPreview.text!}</pre>
                </div>
              ) : null}
              {!selectedAttachmentDocumentPreview && selectedAttachmentPreviewRemoteStatus === "loading" ? (
                <PanelState compact tone="loading" title="A carregar documento" description="A preparar o preview do documento selecionado." />
              ) : null}
              {selectedAttachmentDocumentPreview?.kind === "unsupported" ? (
                <PanelState compact tone="info" title="Preview nao disponivel" description="Este documento pode exigir download ou URL publica para preview." />
              ) : null}
              {!selectedAttachmentDocumentPreview && selectedAttachmentPreviewRemoteStatus !== "loading" && selectedAttachmentPreviewMode !== "none" ? (
                <PanelState compact tone="info" title="Preview nao disponivel" description="Nao foi possivel abrir este documento com a mesma base de preview da aba Grupos." />
              ) : null}
              {selectedAttachmentPreviewMode === "none" ? (
                <PanelState compact tone="info" title="Escolhe um documento" description="Seleciona um documento rapido para abrir o preview." />
              ) : null}
            </div>
          ) : (
            <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe primeiro um documento rapido para abrir o preview." />
          )
        ) : null}
        {previewMode === "reply" ? (
          <div style={S.previewPlaceholder}>
            <div style={S.cardTitle}>Responder</div>
            <div style={S.cardMeta}>Estrutura pronta para editor, IA e selecao de anexos numa fase seguinte.</div>
            <button type="button" style={S.primaryBtn} onClick={() => void handlePreviewReply()} disabled={!selectedEmail}>
              <Icons.MessageSquare size={12} />
              Abrir resposta
            </button>
          </div>
        ) : null}
        {previewMode === "forward" ? (
          <div style={S.previewPlaceholder}>
            <div style={S.cardTitle}>Reencaminhar</div>
            <div style={S.cardMeta}>Estrutura pronta para editor, IA e composicao de envio numa fase seguinte.</div>
            <button type="button" style={S.primaryBtn} onClick={() => void handlePreviewForward()} disabled={!selectedEmail}>
              <Icons.ExternalLink size={12} />
              Abrir reencaminhamento
            </button>
          </div>
        ) : null}
      </div>
    </section>
  );
};

const S: Record<string, React.CSSProperties> = {
  previewToolbar: { display: "flex", gap: 1, background: "rgba(148,163,184,0.12)", padding: 4, borderRadius: 12, marginBottom: 8 },
  previewTab: { flex: 1, height: 28, border: "none", background: "transparent", color: "var(--iccc-muted)", fontSize: 10.5, fontWeight: 600, borderRadius: 8, cursor: "pointer", transition: "all 120ms ease" },
  previewTabOn: { flex: 1, height: 28, border: "none", background: "#fff", color: "var(--iccc-accent)", fontSize: 10.5, fontWeight: 700, borderRadius: 8, cursor: "pointer", boxShadow: "0 2px 8px rgba(15,23,42,0.06)" },
  previewBody: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column", background: "rgba(255,255,255,0.6)", borderRadius: 12, border: "1px solid rgba(148,163,184,0.12)", overflow: "hidden" },
  previewHtml: { flex: 1, minHeight: 0, overflowY: "auto", background: "#fff" },
  documentPreviewShell: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column", overflow: "hidden" },
  documentPreviewFrame: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column", overflow: "hidden" },
  attachmentPreviewImage: { maxWidth: "100%", maxHeight: "100%", objectFit: "contain", margin: "auto" },
  documentPreviewIframe: { border: "none", width: "100%", height: "100%", background: "#fff" },
  attachmentPreviewText: { margin: 0, padding: 18, color: "#172b4d", background: "#fff", font: "14px/1.55 'Segoe UI',sans-serif", whiteSpace: "pre-wrap", wordBreak: "break-word", overflowY: "auto", flex: 1 },
  attachmentPreviewEmpty: { padding: 40, textAlign: "center", fontSize: 13, color: "var(--iccc-muted)", fontStyle: "italic" },
  attachmentPdfPreviewShell: { flex: 1, minHeight: 0, display: "flex", flexDirection: "column", overflowY: "auto", background: "#f8fafc", padding: 20 },
  attachmentPdfPreviewLoading: { padding: 20, textAlign: "center", fontSize: 13, color: "var(--iccc-muted)" },
  attachmentPdfPreviewCanvasHost: { display: "flex", flexDirection: "column", gap: 12, alignItems: "center" },
  previewPlaceholder: { flex: 1, display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", gap: 12, padding: 40, textAlign: "center" },
  cardTitle: { fontSize: 13, fontWeight: 650, color: "var(--iccc-text)" },
  cardMeta: { fontSize: 10.5, lineHeight: 1.3, color: "var(--iccc-muted)", maxWidth: 320 },
  primaryBtn: { height: 30, padding: "0 11px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "linear-gradient(180deg,#3b82f6 0%, #2563eb 100%)", color: "#fff", fontSize: 10.5, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", boxShadow: "0 4px 10px rgba(37,99,235,0.14)" },
};

export default PreviewPane;
