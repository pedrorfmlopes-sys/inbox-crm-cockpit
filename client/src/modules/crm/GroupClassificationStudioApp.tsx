import React, { useEffect, useMemo, useState } from "react";
import { CockpitProvider, useCockpit } from "@/components/shell/CockpitProvider";
import { getRelatedEmailContext, listLinkGroups, listGroupTicketSeries, searchKnownEmails, type GroupTicketEntry, type GroupTicketSeriesEntry, type LinkGroupEntry, type RelatedEmailEntry } from "@/api";
import { requestCockpitHostAction } from "@/office";
import { getSettings } from "@/settings";
import { PanelState } from "@/ui/PanelState";
import { applySkin } from "@/ui/skins";
import * as Icons from "@/ui/icons";
import "../../global.css";

type SectionId = "emails" | "classification" | "labels" | "filters" | "summary";
type ScopeMode = "related" | "all";
type LabelDraft = { categorize: boolean; hasStatus: boolean };

const MENU: Array<{ id: SectionId; label: string; icon: React.ReactNode; help: string }> = [
  { id: "emails", label: "Emails", icon: <Icons.MessageSquare size={15} />, help: "Lista e preview base do caso." },
  { id: "classification", label: "Classificacao", icon: <Icons.Target size={15} />, help: "Grupo principal, referencias e ticket." },
  { id: "labels", label: "Etiquetas", icon: <Icons.Star size={15} />, help: "Etiquetas e futuras categorias Outlook." },
  { id: "filters", label: "Filtros", icon: <Icons.Search size={15} />, help: "Reducao da lista e testes de vista." },
  { id: "summary", label: "Resumo", icon: <Icons.Clipboard size={15} />, help: "Fotografia do que esta preparado." },
];

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.emailKey || email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function dedupeEmails(emails: RelatedEmailEntry[]): RelatedEmailEntry[] {
  const seen = new Set<string>();
  return emails.filter((email) => {
    const key = makeEmailKey(email);
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function htmlToPlainText(html: string): string {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;|&#039;/gi, "'")
    .replace(/&quot;/gi, "\"")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function escapeHtml(value: string): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildEmailPreviewHtml(email: RelatedEmailEntry | null): string {
  const html = String(email?.bodyHtml || "").trim();
  if (html) {
    return `<!doctype html><html><head><meta charset="utf-8" /><style>html,body{margin:0;padding:0;background:#fff;color:#172b4d;font:14px/1.5 'Segoe UI',sans-serif}body{padding:18px}img{max-width:100%;height:auto}table{max-width:100%}blockquote{margin-left:0;padding-left:12px;border-left:3px solid #dbeafe;color:#475569}pre{white-space:pre-wrap;word-break:break-word}</style></head><body>${html}</body></html>`;
  }
  const text = String(email?.bodyText || "").trim();
  if (!text) return "";
  return `<!doctype html><html><head><meta charset="utf-8" /><style>html,body{margin:0;padding:0;background:#fff;color:#172b4d;font:14px/1.55 'Segoe UI',sans-serif}body{padding:18px}pre{margin:0;white-space:pre-wrap;word-break:break-word;font:inherit}</style></head><body><pre>${escapeHtml(text)}</pre></body></html>`;
}

function buildSnippet(email: RelatedEmailEntry): string {
  const source = String(email.bodyText || "").trim() || htmlToPlainText(String(email.bodyHtml || ""));
  return source.length > 180 ? `${source.slice(0, 177).trim()}...` : source;
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
}

function isExternalEmail(email: RelatedEmailEntry): boolean {
  const from = String(email.fromEmail || "").toLowerCase();
  return from ? !from.endsWith("@divitek.pt") : true;
}

function StudioInner() {
  const { ctx, attachments } = useCockpit();
  const [section, setSection] = useState<SectionId>("emails");
  const [scopeMode, setScopeMode] = useState<ScopeMode>("related");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [status, setStatus] = useState("");
  const [groupFilterId, setGroupFilterId] = useState("");
  const [emailSearch, setEmailSearch] = useState("");
  const [onlyExternal, setOnlyExternal] = useState(false);
  const [onlyWithAttachments, setOnlyWithAttachments] = useState(false);
  const [allGroups, setAllGroups] = useState<LinkGroupEntry[]>([]);
  const [ticketSeries, setTicketSeries] = useState<GroupTicketSeriesEntry[]>([]);
  const [relatedTickets, setRelatedTickets] = useState<GroupTicketEntry[]>([]);
  const [relatedEmails, setRelatedEmails] = useState<RelatedEmailEntry[]>([]);
  const [knownEmails, setKnownEmails] = useState<RelatedEmailEntry[]>([]);
  const [selectedEmailKey, setSelectedEmailKey] = useState("");
  const [principalGroupId, setPrincipalGroupId] = useState("");
  const [referenceGroupIds, setReferenceGroupIds] = useState<string[]>([]);
  const [selectedSeriesId, setSelectedSeriesId] = useState("");
  const [labelInput, setLabelInput] = useState("");
  const [selectedLabels, setSelectedLabels] = useState<string[]>([]);
  const [labelDrafts, setLabelDrafts] = useState<Record<string, LabelDraft>>({});

  useEffect(() => {
    void (async () => {
      try {
        const settings = await getSettings();
        applySkin(settings.skinId || "soft");
      } catch {
        applySkin("soft");
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoading(true);
      setError("");
      try {
        const payload = {
          conversationId: ctx.conversationId,
          internetMessageId: ctx.internetMessageId,
          itemId: ctx.itemId,
          subject: ctx.subject,
          fromEmail: ctx.fromEmail,
          fromName: ctx.fromName,
          receivedAtIso: ctx.receivedDateTimeIso,
        };
        const [related, groups, emails, series] = await Promise.all([
          getRelatedEmailContext(payload),
          listLinkGroups(""),
          searchKnownEmails("", { limit: 120 }),
          listGroupTicketSeries(),
        ]);
        if (cancelled) return;
        const mergedGroups = [...groups, ...related.groups].reduce<LinkGroupEntry[]>((acc, group) => {
          if (!group?.id || acc.some((entry) => entry.id === group.id)) return acc;
          acc.push(group);
          return acc;
        }, []);
        const mergedEmails = dedupeEmails([...(related.emails || []), ...(emails || [])]);
        setAllGroups(mergedGroups);
        setTicketSeries(Array.isArray(series) ? series : []);
        setRelatedTickets(Array.isArray(related.tickets) ? related.tickets : []);
        setRelatedEmails(Array.isArray(related.emails) ? related.emails : []);
        setKnownEmails(Array.isArray(emails) ? emails : []);
        setSelectedEmailKey((current) => {
          if (current && mergedEmails.some((email) => makeEmailKey(email) === current)) return current;
          const currentItem = mergedEmails.find((email) => String(email.itemId || "").trim() === String(ctx.itemId || "").trim());
          return makeEmailKey(currentItem || mergedEmails[0] || {});
        });
        setStatus("Janela base pronta. Nesta fase ainda nao altera o sistema atual; prepara apenas a futura UX completa.");
      } catch (fetchError: any) {
        if (!cancelled) setError(String(fetchError?.message || fetchError || "Falha a preparar o studio de classificacao."));
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject]);

  const groupMap = useMemo(() => new Map(allGroups.map((group) => [group.id, group])), [allGroups]);
  const emailPool = useMemo(() => (scopeMode === "related" ? dedupeEmails(relatedEmails) : dedupeEmails([...relatedEmails, ...knownEmails])), [knownEmails, relatedEmails, scopeMode]);

  const visibleEmails = useMemo(() => {
    const q = String(emailSearch || "").trim().toLowerCase();
    return [...emailPool]
      .sort((a, b) => String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || "")))
      .filter((email) => {
        if (onlyExternal && !isExternalEmail(email)) return false;
        if (onlyWithAttachments && !(Array.isArray(email.attachments) && email.attachments.length)) return false;
        if (groupFilterId) {
          const relatedGroupIds = new Set([email.groupId, ...(email.relatedGroups || []).map((entry) => entry.id)].filter(Boolean));
          if (!relatedGroupIds.has(groupFilterId)) return false;
        }
        if (!q) return true;
        const haystack = [email.subject, email.fromName, email.fromEmail, buildSnippet(email)].join(" ").toLowerCase();
        return haystack.includes(q);
      });
  }, [emailPool, emailSearch, groupFilterId, onlyExternal, onlyWithAttachments]);

  const selectedEmail = useMemo(
    () => visibleEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || emailPool.find((email) => makeEmailKey(email) === selectedEmailKey) || visibleEmails[0] || emailPool[0] || null,
    [emailPool, selectedEmailKey, visibleEmails]
  );

  const selectedEmailGroups = useMemo(() => {
    if (!selectedEmail) return [];
    const list = [...(selectedEmail.relatedGroups || []), ...(selectedEmail.groupId ? [{ id: selectedEmail.groupId, name: selectedEmail.groupName, relationKind: selectedEmail.membershipKind }] : [])];
    return list.reduce<Array<{ id: string; name?: string; relationKind?: string }>>((acc, row) => {
      if (!row?.id || acc.some((entry) => entry.id === row.id)) return acc;
      acc.push(row);
      return acc;
    }, []);
  }, [selectedEmail]);

  useEffect(() => {
    if (!selectedEmail) return;
    setPrincipalGroupId((current) => {
      if (current) return current;
      const principal = selectedEmailGroups.find((group) => String(group.relationKind || "").toLowerCase() === "principal");
      return principal?.id || "";
    });
    setReferenceGroupIds((current) => {
      if (current.length) return current;
      return selectedEmailGroups.filter((group) => String(group.relationKind || "").toLowerCase() !== "principal").map((group) => group.id);
    });
  }, [selectedEmail, selectedEmailGroups]);

  const previewHtml = useMemo(() => buildEmailPreviewHtml(selectedEmail), [selectedEmail]);
  const labelCatalog = useMemo(() => {
    const values = new Set<string>();
    allGroups.forEach((group) => (group.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    relatedTickets.forEach((ticket) => (ticket.labels || []).forEach((label) => String(label || "").trim() && values.add(String(label).trim())));
    selectedLabels.forEach((label) => values.add(label));
    return Array.from(values).sort((a, b) => a.localeCompare(b, "pt"));
  }, [allGroups, relatedTickets, selectedLabels]);
  const filteredLabelCatalog = useMemo(() => {
    const q = String(labelInput || "").trim().toLowerCase();
    return q ? labelCatalog.filter((label) => label.toLowerCase().includes(q)) : labelCatalog;
  }, [labelCatalog, labelInput]);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (!closed) window.close();
  }

  function toggleReferenceGroup(groupId: string) {
    setReferenceGroupIds((current) => current.includes(groupId) ? current.filter((entry) => entry !== groupId) : [...current, groupId]);
  }

  function addLabel(label: string) {
    const value = String(label || "").trim();
    if (!value) return;
    setSelectedLabels((current) => current.includes(value) ? current : [...current, value]);
    setLabelDrafts((current) => current[value] ? current : { ...current, [value]: { categorize: false, hasStatus: false } });
    setLabelInput("");
  }

  function updateLabelDraft(label: string, patch: Partial<LabelDraft>) {
    setLabelDrafts((current) => ({ ...current, [label]: { categorize: current[label]?.categorize ?? false, hasStatus: current[label]?.hasStatus ?? false, ...patch } }));
  }

  function removeLabel(label: string) {
    setSelectedLabels((current) => current.filter((entry) => entry !== label));
  }

  function renderWorkspace() {
    if (loading) return <PanelState compact tone="loading" title="A preparar a janela" description="A carregar emails, grupos e series para o novo studio." />;
    if (error) return <PanelState compact tone="error" title="Falha a preparar o studio" description={error} />;

    if (section === "emails") {
      if (!selectedEmail) return <PanelState compact tone="info" title="Sem email selecionado" description="Escolhe um email na coluna do meio." />;
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.titleRow}>
              <div>
                <div style={S.cardTitle}>Preview grande</div>
                <div style={S.cardMeta}>{selectedEmail.subject || "(sem assunto)"}</div>
              </div>
              {(selectedEmail.itemId || selectedEmail.emailWebLink) ? (
                <button type="button" style={S.secondaryBtn} onClick={() => void requestCockpitHostAction({ type: "open-email", itemId: selectedEmail.itemId, emailWebLink: selectedEmail.emailWebLink })}>
                  <Icons.ExternalLink size={12} />
                  Abrir no Outlook
                </button>
              ) : null}
            </div>
            <div style={S.metaLine}>
              <span>{selectedEmail.fromName || selectedEmail.fromEmail || "--"}</span>
              <span>{formatDate(selectedEmail.messageDateIso || selectedEmail.receivedAtIso) || "--"}</span>
              <span>{Array.isArray(selectedEmail.attachments) ? `${selectedEmail.attachments.length} anexo(s)` : "Sem anexos"}</span>
            </div>
            {selectedEmailGroups.length ? <div style={S.chips}>{selectedEmailGroups.map((group) => <span key={group.id} style={S.groupChip}>{group.name || groupMap.get(group.id)?.name || group.id}</span>)}</div> : null}
            {previewHtml ? <iframe title={selectedEmail.subject || "Preview"} srcDoc={previewHtml} style={S.preview} sandbox="" /> : <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado suficiente para preview." />}
          </div>
        </div>
      );
    }

    if (section === "classification") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Classificacao base</div>
            <div style={S.cardMeta}>Estrutura funcional local, sem gravacao nesta fase.</div>
            <div style={S.grid2}>
              <label style={S.field}><span style={S.label}>Grupo principal</span><select style={S.select} value={principalGroupId} onChange={(event) => setPrincipalGroupId(event.target.value)}><option value="">Sem grupo principal</option>{allGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}</select></label>
              <label style={S.field}><span style={S.label}>Serie de ticket</span><select style={S.select} value={selectedSeriesId} onChange={(event) => setSelectedSeriesId(event.target.value)}><option value="">Sem ticket/caso</option>{ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}</select></label>
            </div>
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Grupos referencia</div>
            <div style={S.chips}>{allGroups.filter((group) => group.id !== principalGroupId).map((group) => <button key={group.id} type="button" style={referenceGroupIds.includes(group.id) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => toggleReferenceGroup(group.id)}>{group.name}</button>)}</div>
          </div>
        </div>
      );
    }

    if (section === "labels") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas estruturadas</div>
            <div style={S.inline}>
              <input style={S.input} value={labelInput} onChange={(event) => setLabelInput(event.target.value)} placeholder="Pesquisar ou criar etiqueta" />
              <button type="button" style={S.secondaryBtn} onClick={() => addLabel(labelInput)} disabled={!String(labelInput || "").trim()}><Icons.Plus size={12} />Adicionar</button>
            </div>
            {filteredLabelCatalog.length ? <div style={S.chips}>{filteredLabelCatalog.slice(0, 24).map((label) => <button key={label} type="button" style={selectedLabels.includes(label) ? S.groupChipBtnOn : S.groupChipBtn} onClick={() => addLabel(label)}>{label}</button>)}</div> : null}
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Etiquetas selecionadas</div>
            {selectedLabels.length ? selectedLabels.map((label) => {
              const draft = labelDrafts[label] || { categorize: false, hasStatus: false };
              return (
                <div key={label} style={S.labelRow}>
                  <div style={S.labelHead}><strong>{label}</strong><button type="button" style={S.linkBtn} onClick={() => removeLabel(label)}>Remover</button></div>
                  <label style={S.check}><input type="checkbox" checked={draft.categorize} onChange={(event) => updateLabelDraft(label, { categorize: event.target.checked })} /><span>Virar categoria Outlook</span></label>
                  <label style={S.check}><input type="checkbox" checked={draft.hasStatus} onChange={(event) => updateLabelDraft(label, { hasStatus: event.target.checked })} /><span>Tem estado associado</span></label>
                </div>
              );
            }) : <PanelState compact tone="info" title="Sem etiquetas ainda" description="Vai adicionando etiquetas para testar esta estrutura nova." />}
          </div>
        </div>
      );
    }

    if (section === "filters") {
      return (
        <div style={S.stack}>
          <div style={S.card}>
            <div style={S.cardTitle}>Filtros da janela</div>
            <div style={S.grid2}>
              <label style={S.field}><span style={S.label}>Fonte da lista</span><select style={S.select} value={scopeMode} onChange={(event) => setScopeMode(event.target.value as ScopeMode)}><option value="related">So emails relacionados</option><option value="all">Todos os emails conhecidos</option></select></label>
              <label style={S.field}><span style={S.label}>Filtrar por grupo</span><select style={S.select} value={groupFilterId} onChange={(event) => setGroupFilterId(event.target.value)}><option value="">Sem filtro</option>{allGroups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}</select></label>
            </div>
            <div style={S.inlineChecks}>
              <label style={S.check}><input type="checkbox" checked={onlyExternal} onChange={(event) => setOnlyExternal(event.target.checked)} /><span>So emails externos</span></label>
              <label style={S.check}><input type="checkbox" checked={onlyWithAttachments} onChange={(event) => setOnlyWithAttachments(event.target.checked)} /><span>So emails com anexos</span></label>
            </div>
          </div>
          <div style={S.card}>
            <div style={S.cardTitle}>Resultado atual</div>
            <div style={S.summaryRow}><span>Emails visiveis</span><strong>{visibleEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Emails relacionados</span><strong>{relatedEmails.length}</strong></div>
            <div style={S.summaryRow}><span>Total conhecido</span><strong>{dedupeEmails([...relatedEmails, ...knownEmails]).length}</strong></div>
            <div style={S.summaryRow}><span>Tickets do caso</span><strong>{relatedTickets.length}</strong></div>
          </div>
        </div>
      );
    }

    return (
      <div style={S.stack}>
        <div style={S.card}>
          <div style={S.cardTitle}>Resumo da estrutura</div>
          <div style={S.summaryRow}><span>Email selecionado</span><strong>{selectedEmail?.subject || "--"}</strong></div>
          <div style={S.summaryRow}><span>Grupo principal</span><strong>{principalGroupId ? groupMap.get(principalGroupId)?.name || principalGroupId : "--"}</strong></div>
          <div style={S.summaryRow}><span>Grupos referencia</span><strong>{referenceGroupIds.length}</strong></div>
          <div style={S.summaryRow}><span>Serie de ticket</span><strong>{selectedSeriesId ? ticketSeries.find((entry) => entry.id === selectedSeriesId)?.prefix || selectedSeriesId : "--"}</strong></div>
          <div style={S.summaryRow}><span>Etiquetas</span><strong>{selectedLabels.length}</strong></div>
          <div style={S.summaryRow}><span>Anexos do email atual</span><strong>{attachments.length}</strong></div>
        </div>
        <div style={S.note}>Janela nova criada sem alterar o fluxo atual dos grupos. O proximo passo sera ligar estas escolhas ao sistema real de classificacao e categorias.</div>
      </div>
    );
  }

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.mainTitle}>Studio de classificacao</div>
          <div style={S.mainMeta}>Janela nova e isolada para desenhar a futura atribuicao completa de grupos, tickets, etiquetas e filtros.</div>
        </div>
        <button type="button" style={S.secondaryBtn} onClick={handleClose}>Fechar</button>
      </div>

      <div style={S.context}>
        <div><div style={S.kicker}>Email atual</div><div style={S.contextTitle}>{ctx.subject || "(sem assunto)"}</div></div>
        <div style={S.badges}><span style={S.badge}>{attachments.length} anexo(s)</span><span style={S.badge}>{relatedTickets.length} ticket(s)</span><span style={S.badge}>{relatedEmails.length} relacionados</span></div>
      </div>

      {status ? <div style={S.notice}>{status}</div> : null}

      <div style={S.shell}>
        <aside style={S.sidebar}>
          {MENU.map((item) => (
            <button key={item.id} type="button" style={section === item.id ? S.menuOn : S.menu} onClick={() => setSection(item.id)}>
              <span>{item.icon}</span>
              <span style={{ display: "grid", gap: 2, textAlign: "left" }}><strong>{item.label}</strong><small>{item.help}</small></span>
            </button>
          ))}
        </aside>

        <section style={S.listCol}>
          <div style={S.colTitle}>Emails</div>
          <input style={S.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Pesquisar por assunto, remetente ou texto..." />
          <div style={S.listBody}>
            {loading ? <PanelState compact tone="loading" title="A carregar emails" description="A preparar a lista desta nova janela." /> : null}
            {!loading && !visibleEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Ajusta os filtros ou muda a fonte da lista." /> : null}
            {!loading && visibleEmails.map((email) => (
              <button key={makeEmailKey(email)} type="button" style={makeEmailKey(email) === makeEmailKey(selectedEmail || {}) ? S.emailOn : S.email} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                <div style={S.emailTop}><strong>{email.subject || "(sem assunto)"}</strong>{Array.isArray(email.attachments) && email.attachments.length ? <span style={S.counter}>{email.attachments.length}</span> : null}</div>
                <div style={S.emailMeta}>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</div>
                <div style={S.emailSnippet}>{buildSnippet(email) || "Sem preview curto disponivel."}</div>
              </button>
            ))}
          </div>
        </section>

        <main style={S.workCol}>{renderWorkspace()}</main>
      </div>
    </div>
  );
}

export default function GroupClassificationStudioApp(): JSX.Element {
  return <CockpitProvider><StudioInner /></CockpitProvider>;
}

const S: Record<string, React.CSSProperties> = {
  root: { height: "100vh", boxSizing: "border-box", padding: 18, display: "grid", gridTemplateRows: "auto auto auto minmax(0,1fr)", gap: 12, background: "var(--iccc-bg)", color: "var(--iccc-text)", fontFamily: "var(--iccc-font)", overflow: "hidden" },
  header: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 16, padding: "14px 16px", borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)" },
  kicker: { fontSize: 10, fontWeight: 700, letterSpacing: "0.08em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  mainTitle: { fontSize: 24, fontWeight: 800, color: "var(--iccc-text)" },
  mainMeta: { fontSize: 13, lineHeight: 1.45, color: "var(--iccc-muted)", maxWidth: 820 },
  secondaryBtn: { height: 34, padding: "0 12px", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.88)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 8, cursor: "pointer" },
  context: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "12px 14px", borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.8)" },
  contextTitle: { fontSize: 15, fontWeight: 700, color: "var(--iccc-text)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: 780 },
  badges: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  badge: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, background: "rgba(30,64,175,0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  notice: { padding: "10px 12px", borderRadius: 12, border: "1px solid #bfdbfe", background: "#eff6ff", color: "#1d4ed8", fontSize: 12, lineHeight: 1.45 },
  shell: { minHeight: 0, display: "grid", gridTemplateColumns: "220px 320px minmax(0,1fr)", gap: 12 },
  sidebar: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gap: 8, alignContent: "start", overflowY: "auto" },
  menu: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(148,163,184,0.2)", background: "rgba(255,255,255,0.78)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  menuOn: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.9)", padding: "10px 12px", display: "grid", gridTemplateColumns: "auto minmax(0,1fr)", gap: 10, cursor: "pointer" },
  listCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, display: "grid", gridTemplateRows: "auto auto minmax(0,1fr)", gap: 10, overflow: "hidden" },
  colTitle: { fontSize: 17, fontWeight: 800, color: "var(--iccc-text)" },
  input: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  select: { width: "100%", height: 38, boxSizing: "border-box", borderRadius: 12, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.92)", padding: "0 12px", fontSize: 13, color: "var(--iccc-text)", outline: "none" },
  listBody: { minHeight: 0, display: "grid", gap: 8, overflowY: "auto", paddingRight: 2 },
  email: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(148,163,184,0.2)", background: "rgba(255,255,255,0.78)", padding: "10px 12px", display: "grid", gap: 6, cursor: "pointer" },
  emailOn: { width: "100%", textAlign: "left", borderRadius: 14, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", padding: "10px 12px", display: "grid", gap: 6, cursor: "pointer" },
  emailTop: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 },
  emailMeta: { fontSize: 11, color: "var(--iccc-muted)" },
  emailSnippet: { fontSize: 12, lineHeight: 1.45, color: "var(--iccc-text-soft, #334155)" },
  counter: { minWidth: 22, height: 22, borderRadius: 999, display: "inline-flex", alignItems: "center", justifyContent: "center", background: "rgba(15,23,42,0.06)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 700 },
  workCol: { minHeight: 0, borderRadius: 18, border: "1px solid var(--iccc-border)", background: "var(--iccc-panel)", boxShadow: "var(--iccc-shadow)", padding: 12, overflow: "hidden" },
  stack: { height: "100%", minHeight: 0, display: "grid", gap: 12, alignContent: "start", overflowY: "auto", paddingRight: 2 },
  card: { borderRadius: 16, border: "1px solid var(--iccc-border)", background: "rgba(255,255,255,0.74)", padding: 14, display: "grid", gap: 12 },
  titleRow: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12 },
  cardTitle: { fontSize: 16, fontWeight: 800, color: "var(--iccc-text)" },
  cardMeta: { fontSize: 12, lineHeight: 1.45, color: "var(--iccc-muted)" },
  metaLine: { display: "flex", gap: 12, flexWrap: "wrap", fontSize: 11, color: "var(--iccc-muted)" },
  chips: { display: "flex", flexWrap: "wrap", gap: 8 },
  groupChip: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, background: "rgba(29,78,216,0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  groupChipBtn: { borderRadius: 999, border: "1px solid rgba(148,163,184,0.24)", background: "rgba(255,255,255,0.92)", color: "var(--iccc-text)", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  groupChipBtnOn: { borderRadius: 999, border: "1px solid rgba(37,99,235,0.24)", background: "rgba(219,234,254,0.92)", color: "#1d4ed8", fontSize: 12, fontWeight: 700, padding: "8px 12px", cursor: "pointer" },
  preview: { width: "100%", minHeight: 520, borderRadius: 14, overflow: "hidden", border: "1px solid rgba(148,163,184,0.24)", background: "#fff" },
  grid2: { display: "grid", gridTemplateColumns: "repeat(2,minmax(0,1fr))", gap: 12 },
  field: { display: "grid", gap: 6 },
  label: { fontSize: 11, fontWeight: 700, letterSpacing: "0.06em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  inline: { display: "flex", alignItems: "center", gap: 8 },
  labelRow: { borderRadius: 14, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", padding: 12, display: "grid", gap: 8 },
  labelHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  linkBtn: { border: "none", background: "transparent", color: "#2563eb", fontSize: 12, fontWeight: 700, cursor: "pointer", padding: 0 },
  check: { display: "inline-flex", alignItems: "center", gap: 8, fontSize: 12, color: "var(--iccc-text)" },
  inlineChecks: { display: "flex", gap: 16, flexWrap: "wrap" },
  summaryRow: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "10px 12px", borderRadius: 12, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.76)", fontSize: 13, color: "var(--iccc-text)" },
  note: { padding: "12px 14px", borderRadius: 14, border: "1px solid rgba(191,219,254,0.8)", background: "#eff6ff", color: "#1d4ed8", fontSize: 13, lineHeight: 1.5 },
};
