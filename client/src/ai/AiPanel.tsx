import React, { useEffect, useMemo, useState, useRef } from "react";
import type { OutlookMessageContext } from "../office";
import { getEmailBodyText } from "../office";
import { aiGenerate, type AiLocale, type AiMode, type AiTone } from "./aiClient";
import { getSettings, saveSettings, type CockpitSettingsV1, type LangOption, type AppLocale } from "../settings";

type Action = "summarize" | "reply" | "tasks" | "rewrite";

type RecipientPreset = "reply" | "replyAll" | "custom";
type RecipientRole = "to" | "cc" | "bcc";
type RecipientOrigin = "from" | "to" | "cc" | "body" | "manual";

type RecipientRow = {
  email: string;
  name?: string;
  origins: RecipientOrigin[];
  include: boolean;
  role: RecipientRole;
};


type SnippetTemplate = { id: string; name: string; body: string };

const TPL_KEY = "crmCockpit.templates.v1";

// Local history of generated outputs (per thread) — kept for 3 days
type AiHistoryEntry = {
  id: string;
  ts: number;
  // key that uniquely identifies the email item (thread + item)
  emailKey: string;
  conversationId: string;
  subject?: string;
  html?: string;
  text?: string;
};

const languageLine =
  effectiveLocale === "auto"
    ? `Responde no mesmo idioma em que o email está escrito.
`
    : `Escreve em ${lang}.
`;

const AI_HISTORY_KEY_V1 = "icc.aiHistory.v1";
const AI_HISTORY_KEY_V2 = "icc.aiHistory.v2";
const AI_BLOCK_START = "<!--ICC_AI_START-->";
const AI_BLOCK_END = "<!--ICC_AI_END-->";
const AI_HISTORY_KEEP_MS = 5 * 24 * 60 * 60 * 1000;

function loadAiHistory(): AiHistoryEntry[] {
  try {
    // Prefer v2. If missing, try v1 and migrate.
    const rawV2 = localStorage.getItem(AI_HISTORY_KEY_V2);
    const rawV1 = rawV2 ? null : localStorage.getItem(AI_HISTORY_KEY_V1);
    const raw = rawV2 || rawV1;
    const arr = raw ? (JSON.parse(raw) as any[]) : [];
    const now = Date.now();
    const pruned = Array.isArray(arr)
      ? arr
          .filter((x) => x && typeof x.ts === "number" && now - x.ts <= AI_HISTORY_KEEP_MS)
          .map((x) => {
            // v1 entries might not have emailKey; keep them but tag with conversationId.
            const cid = String((x as any).conversationId || "");
            const emailKey = String((x as any).emailKey || (cid ? `cid:${cid}` : ""));
            return { ...(x as any), conversationId: cid, emailKey } as AiHistoryEntry;
          })
      : [];
    // persist prune to keep storage clean
    localStorage.setItem(AI_HISTORY_KEY_V2, JSON.stringify(pruned));
    if (rawV1) localStorage.removeItem(AI_HISTORY_KEY_V1);
    return pruned;
  } catch {
    return [];
  }
}

function saveAiHistory(entries: AiHistoryEntry[]) {
  try {
    const now = Date.now();
    const pruned = entries.filter((x) => x && typeof x.ts === "number" && now - x.ts <= AI_HISTORY_KEEP_MS);
    localStorage.setItem(AI_HISTORY_KEY_V2, JSON.stringify(pruned.slice(-300)));
  } catch {
    // ignore
  }
}

// Workspace per email item (kept for a few days so you can resume work after switching emails)
type EmailWorkspace = {
  key: string;
  ts: number;
  updatedAt: number;
  conversationId?: string;
  subject?: string;
  summary?: string;
  // compose/rewrite state
  tplPickId?: string;
  composeNotes?: string;
  replyAll?: boolean;
  rewriteText?: string;
  recipientPreset?: RecipientPreset;
  includeBodyEmails?: boolean;
  attachOriginalItem?: boolean;
  // latest outputs
  // 3 result slots (Opção 1/2/3)
  activeOption?: number;
  results?: Array<{ html?: string; text?: string; ts?: number }>;
  // legacy single output (kept for backward compatibility)
  htmlOut?: string;
  textOut?: string;
  resultView?: "html" | "text";
};

const WORKSPACE_KEY = "icc.workspace.v1";
const WORKSPACE_KEEP_MS = 5 * 24 * 60 * 60 * 1000;

function normStr(s: any): string {
  return String(s || "").trim();
}

function normEmail(s: any): string {
  return normStr(s).toLowerCase();
}

function fnv1a(str: string): string {
  let h = 2166136261;
  for (let i = 0; i < str.length; i++) {
    h ^= str.charCodeAt(i);
    // 32-bit FNV-1a
    h += (h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24);
  }
  // >>>0 converts to uint32
  return (h >>> 0).toString(16);
}

function makeEmailKey(ctx: OutlookMessageContext): string {
  const cid = normStr((ctx as any).conversationId);
  const imid = normStr((ctx as any).internetMessageId);
  const itemId = normStr((ctx as any).itemId || (ctx as any).id);
  if (cid && imid) return `${cid}::${imid}`;
  if (cid && itemId) return `${cid}::${itemId}`;
  // Fallback: stable-ish hash from visible metadata
  const subject = normStr((ctx as any).subject).replace(/\s+/g, " ").toLowerCase();
  const from = normEmail((ctx as any).fromEmail || (ctx as any).from);
  const to = normEmail((ctx as any).toEmail || "");
  return `nocid::h${fnv1a([subject, from, to].join("|"))}`;
}

function loadWorkspaces(): Record<string, EmailWorkspace> {
  try {
    const raw = localStorage.getItem(WORKSPACE_KEY);
    const data = raw ? (JSON.parse(raw) as Record<string, EmailWorkspace>) : {};
    const now = Date.now();
    const out: Record<string, EmailWorkspace> = {};
    for (const [k, v] of Object.entries(data || {})) {
      if (!v || typeof v !== "object") continue;
      const updatedAt = typeof (v as any).updatedAt === "number" ? (v as any).updatedAt : (v as any).ts;
      if (!updatedAt || now - updatedAt > WORKSPACE_KEEP_MS) continue;
      out[k] = { ...(v as any), key: k, updatedAt } as EmailWorkspace;
    }
    if (Object.keys(out).length !== Object.keys(data || {}).length) {
      localStorage.setItem(WORKSPACE_KEY, JSON.stringify(out));
    }
    return out;
  } catch {
    return {};
  }
}

function saveWorkspace(ws: EmailWorkspace) {
  try {
    const all = loadWorkspaces();
    all[ws.key] = ws;
    localStorage.setItem(WORKSPACE_KEY, JSON.stringify(all));
  } catch {
    // ignore
  }
}

function loadWorkspace(key: string): EmailWorkspace {
  const all = loadWorkspaces();
  const hit = all[key];
  if (hit) return hit;
  const now = Date.now();
  return { key, ts: now, updatedAt: now };
}

const SUMMARY_KEY_V2 = "icc.summary.v2";

function loadSummary(emailKey: string): string {
  try {
    const raw = localStorage.getItem(SUMMARY_KEY_V2);
    const map = raw ? (JSON.parse(raw) as Record<string, { ts: number; text: string }>) : {};
    const row = map[emailKey];
    if (!row) return "";
    if (typeof row.ts === "number" && Date.now() - row.ts > WORKSPACE_KEEP_MS) return "";
    return String(row.text || "");
  } catch {
    return "";
  }
}

function upsertSummary(emailKey: string, text: string, conversationId: string) {
  try {
    const raw = localStorage.getItem(SUMMARY_KEY_V2);
    const map = raw ? (JSON.parse(raw) as Record<string, { ts: number; text: string; conversationId?: string }>) : {};
    map[emailKey] = { ts: Date.now(), text, conversationId };
    // prune
    const now = Date.now();
    for (const k of Object.keys(map)) {
      const ts = (map[k] as any)?.ts;
      if (!ts || now - ts > WORKSPACE_KEEP_MS) delete map[k];
    }
    localStorage.setItem(SUMMARY_KEY_V2, JSON.stringify(map));
    window.dispatchEvent(new CustomEvent("icc-summary-updated", { detail: { emailKey, conversationId } }));
  } catch {
    // ignore
  }
}

function defaultTemplates(): SnippetTemplate[] {
  return [
    {
      id: "tpl-followup",
      name: "Pedido de informação (follow-up)",
      body: `Olá {{nome}},\n\nObrigado pelo seu email.\nPara avançarmos, pode confirmar por favor:\n- (ponto 1)\n- (ponto 2)\n\nFico a aguardar.\n\nCumprimentos,`,
    },
    {
      id: "tpl-orcamento",
      name: "Pedido de orçamento (curto)",
      body: `Olá {{nome}},\n\nObrigado pelo contacto.\nPara preparar o orçamento, preciso de confirmar:\n- referência / modelo\n- acabamento\n- quantidades\n- prazo pretendido\n\nAssim que tiver estes dados envio a proposta.\n\nCumprimentos,`,
    },
    {
      id: "tpl-atraso",
      name: "Atualização de prazo / atraso",
      body: `Olá {{nome}},\n\nSó para atualizar: estamos a acompanhar o processo e assim que tivermos confirmação de data/prazo voltamos a contactar.\n\nObrigado pela compreensão.\n\nCumprimentos,`,
    },
  ];
}

function loadTemplates(): SnippetTemplate[] {
  try {
    const raw = localStorage.getItem(TPL_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) {
        const clean = parsed
          .map((t: any) => ({
            id: String(t?.id || "").trim() || `tpl-${Math.random()}`,
            name: String(t?.name || "").trim() || "Sem nome",
            body: String(t?.body || ""),
          }))
          .filter((t: any) => t.id && t.name);
        if (clean.length) return clean;
      }
    }
  } catch {
    // ignore
  }
  return defaultTemplates();
}

function saveTemplates(list: SnippetTemplate[]) {
  try {
    localStorage.setItem(TPL_KEY, JSON.stringify(list || []));
  } catch {
    // ignore
  }
}

function applyTemplateVars(body: string, ctx: OutlookMessageContext): string {
  const name = String((ctx as any)?.fromName || "").trim();
  const subject = String((ctx as any)?.subject || "").trim();
  return String(body || "")
    .replace(/\{\{\s*nome\s*\}\}/gi, name || "")
    .replace(/\{\{\s*assunto\s*\}\}/gi, subject || "");
}


function trimEmailBody(raw: string) {
  if (!raw) return "";
  let s = String(raw);

  const markers = [
    /^From:\s.+$/im,
    /^Sent:\s.+$/im,
    /^De:\s.+$/im,
    /^Enviado:\s.+$/im,
    /^On\s.+wrote:\s*$/im,
    /^Em\s.+escreveu:\s*$/im,
    /^-----Original Message-----$/im,
    /^-----Mensagem original-----$/im,
  ];

  // Many forwarded emails contain a mini-header inside the body ("De:", "Enviada:", ...)
  // right at the top. If we cut there, we lose the actual message content below.
  // So we only consider these markers as a "quote boundary" if they appear sufficiently
  // later in the body.
  const MIN_QUOTE_INDEX = 220;

  let cutAt = s.length;
  for (const rx of markers) {
    const m = s.match(rx);
    if (!m || m.index == null) continue;
    const idx = m.index as number;
    if (idx < MIN_QUOTE_INDEX) continue;
    cutAt = Math.min(cutAt, idx);
  }
  s = s.slice(0, cutAt);

  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);

  const MAX = 4500;
  if (s.length > MAX) s = s.slice(0, MAX);

  return s.trim();
}

function trimEmailBodyFull(raw: string) {
  if (!raw) return "";
  let s = String(raw);
  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);
  const MAX = 9000;
  if (s.length > MAX) s = s.slice(0, MAX);
  return s.trim();
}

function sanitizeAiHtml(html: string) {
  // defense-in-depth (server already instructs the model)
  const allowed = new Set(["P", "BR", "UL", "OL", "LI", "STRONG", "EM", "A", "H3", "H4", "CODE"]);
  const div = document.createElement("div");
  div.innerHTML = html || "";

  const walk = (node: Element) => {
    const children = Array.from(node.children);
    for (const el of children) {
      if (!allowed.has(el.tagName)) {
        const text = document.createTextNode(el.textContent || "");
        el.replaceWith(text);
        continue;
      }
      if (el.tagName === "A") {
        const href = el.getAttribute("href") || "";
        if (!href.startsWith("http://") && !href.startsWith("https://") && !href.startsWith("mailto:")) {
          el.removeAttribute("href");
        } else {
          el.setAttribute("target", "_blank");
          el.setAttribute("rel", "noreferrer");
        }
      }
      for (const attr of Array.from(el.attributes)) {
        if (el.tagName === "A" && ["href", "target", "rel"].includes(attr.name)) continue;
        el.removeAttribute(attr.name);
      }
      walk(el);
    }
  };

  walk(div);
  return div.innerHTML;
}

async function copyToClipboard(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // ignore
  }
}

function getMyEmailLower(): string {
  try {
    const e = String((window as any)?.Office?.context?.mailbox?.userProfile?.emailAddress || "").trim();
    return e.toLowerCase();
  } catch {
    return "";
  }
}

function normalizeEmailCandidate(raw: string) {
  let s = String(raw || "").trim();
  s = s.replace(/^[<\(\[]+/, "").replace(/[>\)\],.;:\s]+$/g, "");
  s = s.trim().toLowerCase();
  if (!s.includes("@") || s.length < 5) return "";
  return s;
}

function extractEmailsFromText(text: string): string[] {
  const s = String(text || "");
  const out: string[] = [];

  const mailtoRx = /mailto:([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})/gi;
  let m: RegExpExecArray | null;
  while ((m = mailtoRx.exec(s))) out.push(m[1]);

  const emailRx = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  while ((m = emailRx.exec(s))) out.push(m[0]);

  const seen = new Set<string>();
  const clean: string[] = [];
  for (const x of out) {
    const e = normalizeEmailCandidate(x);
    if (!e) continue;
    if (seen.has(e)) continue;
    seen.add(e);
    clean.push(e);
  }
  return clean;
}

function escapeHtml(s: string) {
  // NOTE: no replaceAll (keeps TS target compatible)
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function makeReplySubject(subject: string) {
  const s = String(subject || "").trim();
  if (!s) return "";
  if (/^\s*re:\s*/i.test(s)) return s;
  return `RE: ${s}`;
}

function makeForwardSubject(subject: string) {
  const s = String(subject || "").trim();
  if (!s) return "";
  if (/^\s*fw:\s*/i.test(s) || /^\s*fwd:\s*/i.test(s)) return s;
  return `FW: ${s}`;
}

function mergeOrigins(a: RecipientOrigin[] = [], b: RecipientOrigin[] = []) {
  const s = new Set<RecipientOrigin>([...a, ...b]);
  return Array.from(s.values());
}

function originLabel(row: RecipientRow): { text: string; tone: "hi" | "mid" | "low" } {
  const o = new Set(row.origins);
  if (o.has("from")) return { text: "Remetente", tone: "hi" };
  if (o.has("to")) return { text: "To", tone: "mid" };
  if (o.has("cc")) return { text: "Cc", tone: "mid" };
  if (o.has("body")) return { text: "Corpo", tone: "low" };
  return { text: "Manual", tone: "low" };
}

function buildHeaderRows(ctx: OutlookMessageContext, myEmailLower: string): RecipientRow[] {
  const rows: RecipientRow[] = [];

  const push = (email?: string, name?: string, origin?: RecipientOrigin) => {
    const e = normalizeEmailCandidate(email || "");
    if (!e) return;
    if (myEmailLower && e === myEmailLower) return;

    const existing = rows.find((r) => r.email === e);
    if (existing) {
      existing.origins = mergeOrigins(existing.origins, origin ? [origin] : []);
      if (!existing.name && name) existing.name = name;
      return;
    }

    rows.push({
      email: e,
      name,
      origins: origin ? [origin] : [],
      include: false,
      role: "cc",
    });
  };

  // From (alvo natural de resposta)
  push(ctx.fromEmail, ctx.fromName, "from");

  for (const r of ctx.toRecipients || []) push((r as any)?.email, (r as any)?.name, "to");
  for (const r of ctx.ccRecipients || []) push((r as any)?.email, (r as any)?.name, "cc");

  return rows;
}

function applyPreset(rows: RecipientRow[], preset: RecipientPreset, ctx: OutlookMessageContext, myEmailLower: string) {
  const from = normalizeEmailCandidate(ctx.fromEmail || "");

  for (const r of rows) {
    r.include = false;
    r.role = "cc";
  }

  if (preset === "reply" || preset === "replyAll") {
    if (from && (!myEmailLower || from !== myEmailLower)) {
      const rr = rows.find((x) => x.email === from);
      if (rr) {
        rr.include = true;
        rr.role = "to";
      } else {
        rows.unshift({ email: from, name: ctx.fromName, origins: ["from"], include: true, role: "to" });
      }
    }
  }

  if (preset === "replyAll") {
    const skip = new Set<string>();
    if (myEmailLower) skip.add(myEmailLower);
    if (from) skip.add(from);

    for (const r of rows) {
      if (skip.has(r.email)) continue;
      const isHeader = r.origins.some((o) => o === "to" || o === "cc");
      if (!isHeader) continue;
      r.include = true;
      r.role = "cc";
    }
  }
}

function BottomSheet({
  open,
  title,
  onClose,
  children,
}: {
  open: boolean;
  title: string;
  onClose: () => void;
  children: React.ReactNode;
}) {
  if (!open) return null;
  return (
    <div style={S.sheetOverlay} onClick={onClose}>
      <div style={S.sheet} onClick={(e) => e.stopPropagation()}>
        <div style={S.sheetHeader}>
          <div style={S.sheetTitle}>{title}</div>
          <button style={S.iconBtnHeader} onClick={onClose} title="Fechar">
            ✕
          </button>
        </div>
        {children}
      </div>
    </div>
  );
}


const LOCALE_LABEL: Record<AppLocale, string> = {
  "pt-PT": "Português (PT)",
  "es-ES": "Espanhol (ES)",
  "en-GB": "Inglês (UK)",
  "it-IT": "Italiano (IT)",
  "de-DE": "Alemão (DE)",
};

const LOCALE_SHORT: Record<AppLocale, string> = {
  "pt-PT": "PT",
  "es-ES": "ES",
  "en-GB": "EN",
  "it-IT": "IT",
  "de-DE": "DE",
};

const ALL_LOCALES: AppLocale[] = ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"];

const ALL_APP_LOCALES: AppLocale[] = ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"];

const LANG_OPTIONS: Array<{ value: LangOption; label: string }> = [
  { value: "auto", label: "Auto" },
  { value: "pt-PT", label: LOCALE_LABEL["pt-PT"] },
  { value: "es-ES", label: LOCALE_LABEL["es-ES"] },
  { value: "en-GB", label: LOCALE_LABEL["en-GB"] },
  { value: "it-IT", label: LOCALE_LABEL["it-IT"] },
  { value: "de-DE", label: LOCALE_LABEL["de-DE"] },
];

function stripDiacritics(s: string): string {
  try {
    return s.normalize("NFD").replace(/[̀-ͯ]/g, "");
  } catch {
    return s;
  }
}

function detectLocaleFromText(text: string): AppLocale {
  const t = stripDiacritics(String(text || "").toLowerCase());

  const score = (phrases: string[]) =>
    phrases.reduce((acc, p) => (t.includes(p) ? acc + 1 : acc), 0);

  const scores: Record<AppLocale, number> = {
    "pt-PT": score([
      " bom dia",
      " obrigado",
      " por favor",
      " cumprimentos",
      " nao ",
      " pode ",
      " preciso ",
    ]),
    "es-ES": score([
      " hola",
      " gracias",
      " por favor",
      " saludos",
      " necesito",
      " presupuesto",
      " podria",
    ]),
    "en-GB": score([
      " hello",
      " thank",
      " thanks",
      " please",
      " regards",
      " best regards",
      " could you",
      " quotation",
    ]),
    "it-IT": score([
      " buongiorno",
      " grazie",
      " per favore",
      " cordiali",
      " saluti",
      " preventivo",
      " potresti",
    ]),
    "de-DE": score([
      " hallo",
      " danke",
      " bitte",
      " mit freundlichen",
      " angebot",
      " konnten sie",
      " grusse",
      " gruessen",
    ]),
  };

  let best: AppLocale = "pt-PT";
  let bestScore = -1;
  (Object.keys(scores) as AppLocale[]).forEach((k) => {
    if (scores[k] > bestScore) {
      best = k;
      bestScore = scores[k];
    }
  });

  // Se não apanharmos nada, cai para PT.
  return best;
}

function resolveLocale(option: LangOption, textForDetection: string, fallback: AppLocale): AppLocale {
  if (option && option !== "auto") return option;
  const detected = detectLocaleFromText(textForDetection);
  return detected || fallback;
}

export default function AiPanel({ ctx }: { ctx: OutlookMessageContext }) {
  const [tone, setTone] = useState<AiTone>("neutro");
  const [mode, setMode] = useState<AiMode>("fast");
  const [locale, setLocale] = useState<AiLocale>("pt-PT");
  const [readingLang, setReadingLang] = useState<LangOption>("auto");
  const [replyLang, setReplyLang] = useState<LangOption>("auto");
  const [settings, setSettings] = useState<CockpitSettingsV1 | null>(null);

  // Languages shown in the quick picker (bottom bar). Controlled from Settings.
  const [enabledLangs, setEnabledLangs] = useState<AppLocale[]>(ALL_LOCALES);
  const [langMenuOpen, setLangMenuOpen] = useState(false);

  // Preview selector (evita duplicar HTML+Texto e rebentar o layout)
  const [resultView, setResultView] = useState<"html" | "text">("html");

  const [rawBody, setRawBody] = useState<string>("");
  const [body, setBody] = useState<string>("");
  const [fullBody, setFullBody] = useState<string>("");
  const [bodyScope, setBodyScope] = useState<"main" | "full">("main");
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState<string>("");
  const [notice, setNotice] = useState<string>("");

  const [reminderDate, setReminderDate] = useState<string>("");
  const [reminderCreateEvent, setReminderCreateEvent] = useState<boolean>(true);
  const REMINDER_CATEGORY = "CRM: Follow-up";

  type ResultSlot = { html: string; text: string; ts: number };
  const makeEmptySlots = (): ResultSlot[] =>
    Array.from({ length: 3 }, () => ({ html: "", text: "", ts: 0 }));

  const [activeOption, setActiveOption] = useState<number>(0);
  const [resultSlots, setResultSlots] = useState<ResultSlot[]>(makeEmptySlots);

  // Derived current outputs (selected tab)
  const htmlOut = resultSlots[activeOption]?.html || "";
  const textOut = resultSlots[activeOption]?.text || "";

  const setHtmlOut = (v: string) =>
    setResultSlots((prev) => {
      const next = prev.length ? [...prev] : makeEmptySlots();
      const cur = next[activeOption] || { html: "", text: "", ts: 0 };
      next[activeOption] = { ...cur, html: v, ts: Date.now() };
      return next;
    });

  const setTextOut = (v: string) =>
    setResultSlots((prev) => {
      const next = prev.length ? [...prev] : makeEmptySlots();
      const cur = next[activeOption] || { html: "", text: "", ts: 0 };
      next[activeOption] = { ...cur, text: v, ts: Date.now() };
      return next;
    });

  const [aiHistory, setAiHistory] = useState<AiHistoryEntry[]>(() => (typeof window === "undefined" ? [] : loadAiHistory()));
  const [historyPickId, setHistoryPickId] = useState<string>("");
  const restoringRef = React.useRef<boolean>(false);

  // Unique key per email item (thread + message id). Used for summary + workspace persistence.
  const emailKey = useMemo(
    () => makeEmailKey(ctx),
    [ctx.conversationId, (ctx as any).internetMessageId, (ctx as any).itemId, (ctx as any).id, (ctx as any).subject]
  );

  // Keep local history (5 days) per email item
  useEffect(() => {
    if (typeof window === "undefined") return;
    if (!emailKey) return;
    // reset picker when switching email
    setHistoryPickId("");
    setAiHistory(loadAiHistory());
  }, [emailKey]);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const conv = ctx.conversationId || "";
    if (!conv) return;
    if (restoringRef.current) return;
    if (!htmlOut && !textOut) return;

    const entry: AiHistoryEntry = {
      id: `${emailKey}-${Date.now()}`,
      ts: Date.now(),
      emailKey,
      conversationId: conv,
      subject: ctx.subject || "",
      html: htmlOut || undefined,
      text: textOut || undefined,
    };

    setAiHistory((prev) => {
      const next = [...prev, entry];
      saveAiHistory(next);
      return next;
    });
  }, [htmlOut, textOut, ctx.conversationId, emailKey]);

  const [sheet, setSheet] = useState<"" | "options" | "compose" | "rewrite" | "recipients" | "context" | "reminder">("");

  // Templates/Snippets

  type SigMode = "off" | "text" | "html" | "image";
  const SIG_KEY_MODE = "icc.sig.mode";
  const SIG_KEY_TEXT = "icc.sig.text";
  const SIG_KEY_HTML = "icc.sig.html";
  const SIG_KEY_IMG = "icc.sig.img";
  const [sigMode, setSigMode] = useState<SigMode>("off");
  const [sigText, setSigText] = useState<string>("");
  const [sigHtml, setSigHtml] = useState<string>("");
  const [sigImgUrl, setSigImgUrl] = useState<string>("");
  const SIG_KEY_IMG_DATA = "icc.sig.img.data";
  const SIG_KEY_IMG_W = "icc.sig.img.w";
  const [sigImgDataUrl, setSigImgDataUrl] = useState<string>("");
  const [sigImgMaxW, setSigImgMaxW] = useState<string>("260");

  const persistSig = (k: string, v: string) => {
    try {
      localStorage.setItem(k, v);
    } catch {}
  };

  const loadSig = () => {
    try {
      const m = (localStorage.getItem(SIG_KEY_MODE) as SigMode) || "off";
      setSigMode(m);
      setSigText(localStorage.getItem(SIG_KEY_TEXT) || "");
      setSigHtml(localStorage.getItem(SIG_KEY_HTML) || "");
      setSigImgUrl(localStorage.getItem(SIG_KEY_IMG) || "");
      setSigImgDataUrl(localStorage.getItem(SIG_KEY_IMG_DATA) || "");
      setSigImgMaxW(localStorage.getItem(SIG_KEY_IMG_W) || "260");
    } catch {}
  };

  const buildSignatureHtml = () => {
    if (sigMode === "off") return "";

    if (sigMode === "html") {
      const h = (sigHtml || "").trim();
      return h ? `<div class="icc-sig">${h}</div>` : "";
    }

    if (sigMode === "image") {
      const data = (sigImgDataUrl || "").trim();
      const url = (sigImgUrl || "").trim();
      const src = data || url;
      if (!src) return "";
      const w = Math.max(120, Math.min(600, parseInt(sigImgMaxW || "260", 10) || 260));
      const safeSrc = src.startsWith("data:") ? src : escapeHtml(src);
      return `<div class="icc-sig"><br/><img src="${safeSrc}" alt="" style="max-width:${w}px;height:auto;display:block;"/></div>`;
    }

    // text
    const t = (sigText || "").trim();
    if (!t) return "";
    return `<div class="icc-sig" style="white-space:pre-wrap"><br/>${escapeHtml(t)}</div>`;
  };
  const [templates, setTemplates] = useState<SnippetTemplate[]>(() => loadTemplates());
  const [tplPickId, setTplPickId] = useState<string>("");
  const [tplEdit, setTplEdit] = useState<SnippetTemplate | null>(null);
  const [tplName, setTplName] = useState<string>("");
  const [tplBody, setTplBody] = useState<string>("");

  useEffect(() => {
    saveTemplates(templates);
  }, [templates]);


  const [composeNotes, setComposeNotes] = useState<string>("");
  const [rewriteText, setRewriteText] = useState<string>("");

  const [recipientPreset, setRecipientPreset] = useState<RecipientPreset>("reply");
  const [replyAll, setReplyAll] = useState<boolean>(false);
  const myEmailLower = useMemo(() => getMyEmailLower(), []);

  const [includeBodyEmails, setIncludeBodyEmails] = useState<boolean>(true);

  // Quando ativo, ao criar Reply/New, anexa o email original como .msg
  const [attachOriginalItem, setAttachOriginalItem] = useState<boolean>(false);

  const [recipientRows, setRecipientRows] = useState<RecipientRow[]>([]);

  const canRun = Boolean(ctx.conversationId);

  // Persist + restore per-email "workspace" so switching emails doesn't mix state
  const restoringWsRef = useRef(false);

  useEffect(() => {
    if (!emailKey) return;
    restoringWsRef.current = true;
    const ws = loadWorkspace(emailKey);

    setTplPickId(ws.tplPickId || "");
    setComposeNotes(ws.composeNotes || "");
    setRewriteText(ws.rewriteText || "");
    if (ws.recipientPreset) setRecipientPreset(ws.recipientPreset);
    setReplyAll(Boolean(ws.replyAll));
    if (typeof ws.includeBodyEmails === "boolean") setIncludeBodyEmails(ws.includeBodyEmails);
    setAttachOriginalItem(Boolean(ws.attachOriginalItem));
    // Restore 3 result slots (backwards compatible with legacy htmlOut/textOut)
    const slots = makeEmptySlots();
    const wsSlots = Array.isArray(ws.results) ? ws.results : [];
    for (let i = 0; i < 3; i++) {
      const v = wsSlots[i] || undefined;
      if (v && (v.html || v.text)) {
        slots[i] = { html: v.html || "", text: v.text || "", ts: v.ts || 0 };
      }
    }
    if (!slots.some((x) => x.html || x.text)) {
      // legacy
      slots[0] = { html: ws.htmlOut || "", text: ws.textOut || "", ts: ws.ts || 0 };
    }
    setResultSlots(slots);
    setActiveOption(typeof ws.activeOption === "number" ? Math.max(0, Math.min(2, ws.activeOption)) : 0);

    restoringWsRef.current = false;
  }, [emailKey]);

  useEffect(() => {
    if (!emailKey) return;
    if (restoringWsRef.current) return;
    const t = window.setTimeout(() => {
      saveWorkspace({
        key: emailKey,
        ts: Date.now(),
        updatedAt: Date.now(),
        conversationId: ctx.conversationId,
        subject: (ctx as any).subject,
        summary: loadSummary(emailKey) || undefined,
        tplPickId: tplPickId || undefined,
        composeNotes: composeNotes || undefined,
        rewriteText: rewriteText || undefined,
        recipientPreset,
        replyAll,
        includeBodyEmails,
        attachOriginalItem,
        activeOption,
        results: resultSlots.map((s) => ({ html: s.html || undefined, text: s.text || undefined, ts: s.ts || undefined })),
        // legacy (keep for downgrade compatibility)
        htmlOut: htmlOut || undefined,
        textOut: textOut || undefined,
      });
    }, 250);
    return () => window.clearTimeout(t);
  }, [
    emailKey,
    ctx.conversationId,
    (ctx as any).subject,
    tplPickId,
    composeNotes,
    rewriteText,
    recipientPreset,
    replyAll,
    includeBodyEmails,
    attachOriginalItem,
    activeOption,
    resultSlots,
    htmlOut,
    textOut,
  ]);

  const [addEmail, setAddEmail] = useState("");
  const [addRole, setAddRole] = useState<RecipientRole>("bcc");


// Track the currently selected emailKey so delayed tasks can cancel correctly
const currentEmailKeyRef = useRef<string>("");
useEffect(() => {
  currentEmailKeyRef.current = emailKey || "";
}, [emailKey]);

// Auto-summary state per email (throttle + in-flight guard). We only mark success when a summary is actually saved.
const autoSummaryStateRef = useRef<Record<string, { inflight?: boolean; lastAttempt?: number }>>({});

function stripHtml(html: string): string {
  return (html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

useEffect(() => {
  if (!ctx.conversationId) return;
  if (!emailKey) return;

  // If we already have a cached summary for this email, nothing to do.
  const cached = loadSummary(emailKey);
  if (cached) return;

  // Need a body to summarize (avoid using previous email body).
  const bodyNow = (rawBody || body || "").trim();
  if (!bodyNow) return;

  const st = (autoSummaryStateRef.current[emailKey] ||= {});
  const now = Date.now();

  // Don't run multiple times in parallel for the same email.
  if (st.inflight) return;

  // Throttle retries (important when the API is down; otherwise we 'give up forever').
  if (st.lastAttempt && now - st.lastAttempt < 30000) return;

  st.inflight = true;
  st.lastAttempt = now;

  const keyAtSchedule = emailKey;

  const t = window.setTimeout(async () => {
    try {
      // If user switched emails meanwhile, cancel.
      if (currentEmailKeyRef.current !== keyAtSchedule) return;

      // Run summarize. Summary language is always pt-PT (Portuguese Portugal).
      await run("summarize");
    } finally {
      const st2 = (autoSummaryStateRef.current[keyAtSchedule] ||= {});
      st2.inflight = false;
    }
  }, 350);

  return () => {
    window.clearTimeout(t);
    const st2 = (autoSummaryStateRef.current[keyAtSchedule] ||= {});
    st2.inflight = false;
  };
  // eslint-disable-next-line react-hooks/exhaustive-deps
}, [ctx.conversationId, emailKey, rawBody, body]);


  async function ensureMasterCategory(displayName: string): Promise<void> {
    const OfficeAny: any = (window as any).Office;
    if (!OfficeAny?.context?.mailbox?.masterCategories) return;
    await new Promise<void>((resolve) => {
      try {
        OfficeAny.context.mailbox.masterCategories.getAsync((res: any) => {
          if (res.status !== OfficeAny.AsyncResultStatus.Succeeded) return resolve();
          const list = res.value || [];
          const exists = list.some((c: any) => (c.displayName || c.name) === displayName);
          if (exists) return resolve();
          const color = OfficeAny.MailboxEnums?.CategoryColor?.Preset0;
          OfficeAny.context.mailbox.masterCategories.addAsync([{ displayName, color }], (r2: any) => resolve());
        });
      } catch {
        resolve();
      }
    });
  }

  async function addCategoryToItem(displayName: string): Promise<void> {
    const OfficeAny: any = (window as any).Office;
    if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) return;
    await new Promise<void>((resolve) => {
      try {
        OfficeAny.context.mailbox.item.categories.addAsync([displayName], (_: any) => resolve());
      } catch {
        resolve();
      }
    });
  }

  function makeDueDate(offsetDays: number): Date {
    const d = new Date();
    d.setDate(d.getDate() + offsetDays);
    d.setHours(9, 0, 0, 0);
    return d;
  }

  function parseCustomDate(isoDate: string): Date | null {
    const s = String(isoDate || "").trim();
    if (!s) return null;
    // expected yyyy-mm-dd
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (!m) return null;
    const y = Number(m[1]);
    const mo = Number(m[2]) - 1;
    const da = Number(m[3]);
    const d = new Date(y, mo, da, 9, 0, 0, 0);
    if (Number.isNaN(d.getTime())) return null;
    return d;
  }

  function openCalendarReminder(due: Date) {
    const OfficeAny: any = (window as any).Office;
    const mb = OfficeAny?.context?.mailbox;
    if (!mb?.displayNewAppointmentForm) return;
    const subject = ctx.subject ? `Follow-up: ${ctx.subject}` : "Follow-up";
    const bodyTxt = `Follow-up do email\n\nAssunto: ${ctx.subject || ""}\nDe: ${ctx.fromName || ""} <${ctx.fromEmail || ""}>\n`;
    const start = due;
    const end = new Date(due.getTime() + 15 * 60 * 1000);
    try {
      mb.displayNewAppointmentForm({ subject, body: bodyTxt, start, end });
    } catch {
      // ignore
    }
  }

  async function applyReminder(due: Date) {
    if (!ctx.subject) {
      setNotice("Seleciona um email primeiro.");
      return;
    }
    setBusy(true);
    setErr("");
    try {
      await ensureMasterCategory(REMINDER_CATEGORY);
      await addCategoryToItem(REMINDER_CATEGORY);
      if (reminderCreateEvent) openCalendarReminder(due);
      setNotice(`Lembrete marcado para ${due.toLocaleDateString()} (categoria aplicada).`);
      setSheet("");
    } catch (e: any) {
      setErr(e?.message ? String(e.message) : "Não foi possível aplicar o lembrete.");
    } finally {
      setBusy(false);
    }
  }


  // Load defaults from Settings (best effort)
  useEffect(() => {
    (async () => {
      try {
        const s = await getSettings();
        if (s) setSettings(s);
        if (s?.tone) setTone(s.tone);
        if (s?.readingLanguage) setReadingLang(s.readingLanguage);
        if (s?.replyLanguage) setReplyLang(s.replyLanguage);
        if (s?.enabledLanguages && Array.isArray(s.enabledLanguages) && s.enabledLanguages.length > 0) {
          // Keep only known locales, preserve order.
          const filtered = s.enabledLanguages.filter((l) => (ALL_LOCALES as any).includes(l));
          if (filtered.length > 0) setEnabledLangs(filtered as AppLocale[]);
        }
        // Mantém a UI alinhada com a língua da app (fallback).
        if (s?.appLanguage) setLocale(s.appLanguage as AiLocale);
      } catch {
        // ignore
      }
    })();
  }, []);

  // Init recipients when message changes.
  // NOTE: using emailKey (not only conversationId) so switching within the same thread refreshes correctly.
  useEffect(() => {
    const base = buildHeaderRows(ctx, myEmailLower);
    applyPreset(base, "reply", ctx, myEmailLower);
    setRecipientPreset("reply");
    setRecipientRows(base);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [emailKey]);

  // When switching email, clear body state immediately to avoid leaking content between emails.
  useEffect(() => {
    setRawBody("");
    setBody("");
    setFullBody("");
    setBodyScope("main");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [emailKey]);

  // Auto-load body text (best effort)
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const raw = await getEmailBodyText();
        const rawStr = raw || "";
        const clean = trimEmailBody(rawStr);
        if (!cancelled) {
          setRawBody(rawStr);
          setBody(clean);
          setFullBody(rawStr);
        }
      } catch {
        // ignore
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [emailKey]);

  // Extract emails from body and add as suggestions
  useEffect(() => {
    if (!includeBodyEmails) return;
    const src = rawBody || body;
    if (!src) return;

    const found = extractEmailsFromText(src);
    if (!found.length) return;

    setRecipientRows((prev) => {
      const rows = prev.map((r) => ({ ...r, origins: [...r.origins] }));
      for (const email of found) {
        if (myEmailLower && email === myEmailLower) continue;
        const existing = rows.find((r) => r.email === email);
        if (existing) {
          existing.origins = mergeOrigins(existing.origins, ["body"]);
        } else {
          rows.push({ email, name: undefined, origins: ["body"], include: false, role: "cc" });
        }
      }
      return rows;
    });
  }, [rawBody, body, includeBodyEmails, myEmailLower]);

  function applyRecipientPreset(preset: RecipientPreset) {
    setRecipientRows((prev) => {
      const rows = prev.map((r) => ({ ...r, origins: [...r.origins] }));
      applyPreset(rows, preset, ctx, myEmailLower);
      return rows;
    });
    setRecipientPreset(preset);
  }

  function groupedRecipientsFromRows(rows: RecipientRow[]) {
    const to: string[] = [];
    const cc: string[] = [];
    const bcc: string[] = [];

    const seen = new Set<string>();
    for (const r of rows) {
      if (!r.include) continue;
      const e = normalizeEmailCandidate(r.email);
      if (!e) continue;

      const key = `${r.role}:${e}`;
      if (seen.has(key)) continue;
      seen.add(key);

      if (r.role === "to") to.push(e);
      else if (r.role === "cc") cc.push(e);
      else bcc.push(e);
    }

    return { to, cc, bcc };
  }

  const email = useMemo(() => {
    const { to, cc, bcc } = groupedRecipientsFromRows(recipientRows);

    const selectedBody = bodyScope === "full" ? (fullBody || rawBody || body) : body;

    return {
      subject: ctx.subject || "",
      from: ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail || ""}>` : ctx.fromEmail || "",
      to,
      cc,
      bcc,
      bodyScope,
      bodyText: selectedBody,
    };
  }, [ctx, body, rawBody, fullBody, bodyScope, recipientRows]);

  async function run(action: Action) {
    try {
      setErr("");
      setNotice("");
      setBusy(true);

      // Capture the email identity at the moment the action starts.
      // This prevents results (especially summary) from being saved to the wrong email
      // if the user switches the selected message while the AI call is running.
      const runEmailKey = emailKey;
      const runConversationId = ctx.conversationId || "";

      // For non-summary actions, we clear the previous output before generating again.
      // (Summary should live in the Summary card only, so we don't duplicate it in Result.)
      if (action !== "summarize") {
        setHtmlOut("");
        setTextOut("");
      }

      const emailForAi = action !== "rewrite" ? { ...email } : undefined;
      if (action === "reply" && composeNotes.trim()) {
        emailForAi!.bodyText = `${email.bodyText || ""}\n\n---\nINSTRUÇÕES DO UTILIZADOR (não copiar literalmente):\n${composeNotes.trim()}\n`;
      }

      
      const referenceText = action === "rewrite"
        ? rewriteText
        : `${email.subject || ""}
${email.bodyText || ""}
${composeNotes || ""}`;
      // Language rules:
      // - Summary: always Portuguese (Portugal)
      // - Other actions: selected reply language; if Auto, keep the email's original language (server handles)
      let effectiveLocale: AiLocale = (locale as any) as AiLocale;
      if (action === "summarize") {
        effectiveLocale = "pt-PT";
      } else {
        effectiveLocale = (replyLang === "auto" ? ("auto" as any) : (replyLang as any)) as AiLocale;
      }

      const r = await aiGenerate({
        action,
        mode,
        locale: effectiveLocale,
        tone,
        email: action === "rewrite" ? undefined : (emailForAi as any),
        inputText: action === "rewrite" ? rewriteText : undefined,
      });

      const safeHtml = sanitizeAiHtml((r as any).html || "");
      const plainTxt = ((r as any).text || "").trim() || stripHtml(safeHtml);

      // Summary: store only in the Summary cache/card (per email), not in Result output.
      if (action === "summarize") {
        if (runEmailKey && plainTxt) {
          upsertSummary(runEmailKey, plainTxt, runConversationId);
          setNotice("Resumo atualizado.");
        }
        setSheet("");
        return;
      }

      // For other actions, only apply to the currently open email.
      // If the user switched emails mid-run, don't pollute the current view.
      if (runEmailKey !== emailKey) {
        // Persist the result to the workspace of the email where the action started.
        // When the user returns to that email, it will be restored.
        try {
          saveWorkspace({
            key: runEmailKey,
            ts: Date.now(),
            updatedAt: Date.now(),
            conversationId: runConversationId,
            subject: (ctx as any).subject,
            summary: loadSummary(runEmailKey) || undefined,
            tplPickId: tplPickId || undefined,
            composeNotes: composeNotes || undefined,
            rewriteText: rewriteText || undefined,
            recipientPreset,
            replyAll,
            includeBodyEmails,
            attachOriginalItem,
            htmlOut: safeHtml || undefined,
            textOut: plainTxt || undefined,
          });
        } catch {
          // ignore
        }
        setNotice("O email mudou durante a geração; o resultado ficou guardado no email anterior.");
        setSheet("");
        return;
      }

      setHtmlOut(safeHtml);
      setTextOut(plainTxt);
      setResultView(safeHtml ? "html" : "text");

      setSheet("");
    } catch (e: any) {
      setErr(e?.message ?? String(e));
    } finally {
      setBusy(false);
    }
  }

  function buildHtmlForInsert() {
    return htmlOut
      ? htmlOut
      : textOut
        ? `<pre style="white-space:pre-wrap;font-family:inherit">${escapeHtml(textOut)}</pre>`
        : "";
  }

  function wrapAiBlock(innerHtml: string) {
    // Marker-based block so the next insertion can replace only our content
    return `${AI_BLOCK_START}<div data-icc-ai="1">${innerHtml}</div>${AI_BLOCK_END}`;
  }

  async function replaceOrInsertAiBlock(item: any, htmlBlock: string): Promise<boolean> {
    const OfficeAny = (window as any)?.Office;
    const body = item?.body;
    if (!OfficeAny || !body || !body.getAsync || !body.setAsync) return false;

    return await new Promise((resolve) => {
      try {
        body.getAsync(OfficeAny.CoercionType.Html, (r: any) => {
          try {
            if (!r || r.status !== OfficeAny.AsyncResultStatus.Succeeded) return resolve(false);
            const current = String(r.value || "");

            const doSet = (nextHtml: string) => {
              body.setAsync(nextHtml, { coercionType: OfficeAny.CoercionType.Html }, (r2: any) => {
                resolve(!!r2 && r2.status === OfficeAny.AsyncResultStatus.Succeeded);
              });
            };

            if (current.includes(AI_BLOCK_START) && current.includes(AI_BLOCK_END)) {
              const before = current.split(AI_BLOCK_START)[0] || "";
              const after = current.split(AI_BLOCK_END)[1] || "";
              return doSet(before + htmlBlock + after);
            }

            if (typeof body.prependAsync === "function") {
              body.prependAsync(htmlBlock, { coercionType: OfficeAny.CoercionType.Html }, (r3: any) => {
                resolve(!!r3 && r3.status === OfficeAny.AsyncResultStatus.Succeeded);
              });
              return;
            }

            // Fallback: best-effort replace body (may remove signature in alguns Outlooks)
            return doSet(htmlBlock + current);
          } catch (e) {
            resolve(false);
          }
        });
      } catch (e) {
        resolve(false);
      }
    });
  }

  /**

   * Inserts generated content into the CURRENT compose draft (best-effort):
   * 1) cursor (setSelectedDataAsync)
   * 2) prependAsync
   * 3) appendAsync
   * 4) setAsync (replace) as last resort
   */
  async function insertIntoCurrentDraft() {
    try {
      setErr("");
      setNotice("");

      const OfficeAny = (window as any)?.Office;
      const item = OfficeAny?.context?.mailbox?.item;
      if (!item) throw new Error("Esta ação só funciona dentro do Outlook (email aberto).");

      const html = buildHtmlForInsert();
      if (!html) return;

      const sig = buildSignatureHtml();
      const htmlWithSig = sig ? `${html}${sig}` : html;

      // Prefer marker-based replacement to avoid duplicating content on consecutive inserts
      const htmlBlock = wrapAiBlock(htmlWithSig);
      const didReplace = await replaceOrInsertAiBlock(item, htmlBlock);
      if (didReplace) {
        setNotice("Conteúdo inserido (substituído) no rascunho.");
        return;
      }

      const body = item.body;
      if (!body) {
        throw new Error("Este Outlook não expôs a API do corpo (body). Usa ‘Inserir thread’ ou ‘Nova msg’. ");
      }

      const callAsync = (fn: any, args: any[] = []) =>
        new Promise<void>((resolve, reject) => {
          try {
            fn(
              ...args,
              (asyncResult: any) => {
                if (!asyncResult || asyncResult.status === "succeeded") resolve();
                else reject(asyncResult.error?.message || "Falha ao inserir no rascunho");
              }
            );
          } catch (e: any) {
            reject(e?.message ?? String(e));
          }
        });

      // 1) Insert at cursor/selection (compose mode)
      if (typeof body.setSelectedDataAsync === "function") {
        await callAsync(body.setSelectedDataAsync.bind(body), [htmlWithSig, { coercionType: "html" }]);
        setNotice("Inserido no rascunho (no cursor).");
        return;
      }

      // 2) Prepend
      if (typeof body.prependAsync === "function") {
        await callAsync(body.prependAsync.bind(body), [htmlWithSig, { coercionType: "html" }]);
        setNotice("Inserido no rascunho (no topo).");
        return;
      }

      // 3) Append
      if (typeof body.appendAsync === "function") {
        await callAsync(body.appendAsync.bind(body), [htmlWithSig, { coercionType: "html" }]);
        setNotice("Inserido no rascunho (no fim).");
        return;
      }

      // 4) Replace (last resort)
      if (typeof body.setAsync === "function") {
        await callAsync(body.setAsync.bind(body), [htmlWithSig, { coercionType: "html" }]);
        setNotice("Inserido no rascunho (substituiu o corpo, porque este Outlook não permite inserir no cursor).");
        return;
      }

      throw new Error("Este Outlook não permite inserir no rascunho via add-in (API indisponível). Usa ‘Inserir thread’ ou ‘Nova msg’. ");
    } catch (e: any) {
      setErr(e?.message ?? String(e));
    }
  }


async function insertThread() {
  try {
    setErr("");
    setNotice("");

    const OfficeAny = (window as any)?.Office;
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) throw new Error("Esta ação só funciona dentro do Outlook (email aberto).");

    const html = buildHtmlForInsert();
    if (!html) return;

    const sig = buildSignatureHtml();
    const htmlWithSig = sig ? `${html}${sig}` : html;

    if (recipientPreset === "custom") {
      // Não bloqueamos: abrimos a thread na mesma, mas avisamos.
      // (O Outlook decide To/Cc no Reply/ReplyAll em Read mode.)
      setNotice("Na thread, o Outlook decide To/Cc. Se precisares de controlar destinatários, usa “Nova msg”.");
    }

    const OfficeAny2 = (window as any)?.Office;

    const makeOriginalItemAttachment = () => {
      if (!attachOriginalItem) return null;
      const itemId = item?.itemId;
      if (!itemId) return null;
      const t = OfficeAny2?.MailboxEnums?.AttachmentType?.Item ?? "item";
      const safe = (ctx.subject || "email").replace(/[^a-z0-9\-_. ]+/gi, "_").slice(0, 64);
      return { itemId, name: `${safe}.msg`, type: t };
    };

    // Compatibilidade: alguns Outlook só aceitam string (htmlBody) em vez de object { htmlBody, attachments }
    const tryCall = (fn: any) => {
      const original = makeOriginalItemAttachment();

      // 1) tentar com object (para suportar anexos)
      try {
        fn(original ? { htmlBody: htmlWithSig, attachments: [original] } : { htmlBody: htmlWithSig });
        return true;
      } catch {}

      // 2) fallback: só html (alguns clientes não aceitam attachments no form)
      try {
        fn(htmlWithSig);
        setNotice("Nota: este Outlook não suportou anexar o original na resposta. Abri a resposta sem anexo.");
        return true;
      } catch {}

      return false;
    };

    if (recipientPreset === "replyAll" && typeof item.displayReplyAllForm === "function") {
      if (tryCall(item.displayReplyAllForm.bind(item))) return;
    }

    if (typeof item.displayReplyForm === "function") {
      if (tryCall(item.displayReplyForm.bind(item))) return;
    }

    throw new Error("Não foi possível abrir uma resposta na thread (API indisponível).");
  } catch (e: any) {
    setErr(e?.message ?? String(e));
  }
}




async function insertNewMessage() {
  try {
    setErr("");
    setNotice("");

    const OfficeAny = (window as any)?.Office;
    const mailbox = OfficeAny?.context?.mailbox;
    const item = mailbox?.item;
    if (!item) throw new Error("Esta ação só funciona dentro do Outlook (email aberto).");

    const html = buildHtmlForInsert();
    if (!html) return;

    const sig = buildSignatureHtml();
    const htmlWithSig = sig ? `${html}${sig}` : html;

    const { to, cc, bcc } = groupedRecipientsFromRows(recipientRows);
    if (!to.length && !cc.length && !bcc.length) {
      throw new Error("Escolhe pelo menos 1 destinatário (To/Cc/Bcc).");
    }

    const formData: any = {
      subject: makeReplySubject(ctx.subject || ""),
      htmlBody: htmlWithSig,
    };

    // Opcional: anexar o email original como .msg
    if (attachOriginalItem && item?.itemId) {
      const type = OfficeAny?.MailboxEnums?.AttachmentType?.Item ?? "item";
      formData.attachments = [
        {
          itemId: item.itemId,
          name: `${(ctx.subject || "email").slice(0, 60).replace(/[\\/:*?"<>|]/g, "-")}.msg`,
          type,
        },
      ];
    }
    if (to.length) formData.toRecipients = to;
    if (cc.length) formData.ccRecipients = cc;
    if (bcc.length) formData.bccRecipients = bcc;

    // Nota: em muitos Outlook, o método está no mailbox (não no item)
    if (mailbox && typeof mailbox.displayNewMessageForm === "function") {
      try {
        mailbox.displayNewMessageForm(formData);
        return;
      } catch (e) {
        if (formData.attachments) {
          const copy: any = { ...formData };
          delete copy.attachments;
          try {
            mailbox.displayNewMessageForm(copy);
            setNotice("Nota: este Outlook não suportou anexar o original na nova mensagem. Abri a mensagem sem anexo.");
            return;
          } catch {}
        }
        throw e;
      }
    }

    // fallback (algumas builds expõem no item)
    if (typeof item.displayNewMessageForm === "function") {
      item.displayNewMessageForm(formData);
      return;
    }

    throw new Error("Este Outlook não permite criar uma nova mensagem via add-in (API indisponível).");
  } catch (e: any) {
    setErr(e?.message ?? String(e));
  }
}


async function insertForward() {
  try {
    setErr("");
    setNotice("");

    const OfficeAny = (window as any)?.Office;
    const mailbox = OfficeAny?.context?.mailbox;
    const item = mailbox?.item;
    if (!item) throw new Error("Esta ação só funciona dentro do Outlook (email aberto).");

    const html = buildHtmlForInsert();
    if (!html) return;

    const sig = buildSignatureHtml();
    const htmlWithSig = sig ? `${html}${sig}` : html;

    // Forward também não é um reply real, mas dá controlo de destinatários e mantém o original no corpo.
    // Tentamos preencher To/Cc quando o Outlook aceita formData. Bcc raramente é suportado.
    const { to, cc, bcc } = groupedRecipientsFromRows(recipientRows);

    const tryCall = (fn: any) => {
      // 1) tentar com object (melhor: respeita destinatários)
      try {
        fn({ htmlBody: htmlWithSig, toRecipients: to, ccRecipients: cc });
        return true;
      } catch {}
      // 2) fallback: string (compat)
      try {
        fn(htmlWithSig);
        if ((to && to.length) || (cc && cc.length) || (bcc && bcc.length)) {
          setNotice("Nota: este Outlook não suportou preencher destinatários no Forward. Abri o Forward e tens de preencher manualmente.");
        }
        return true;
      } catch {}
      // 3) object sem recipients
      try {
        fn({ htmlBody: htmlWithSig });
        return true;
      } catch {}
      return false;
    };

    if (typeof item.displayForwardForm === "function") {
      if (tryCall(item.displayForwardForm.bind(item))) return;
    }

    // Fallback: alguns Outlook não expõem displayForwardForm.
    // Abrimos uma "Nova msg" com assunto FW: e incluímos o original no corpo.
    if (mailbox && typeof mailbox.displayNewMessageForm === "function") {
      const original = fullBody || rawBody || body || "";
      const originalHtml = original
        ? `<hr style="border:none;border-top:1px solid #ddd;margin:16px 0"/>` +
          `<div style="color:#666;font-size:12px;margin-bottom:6px"><strong>Mensagem original</strong></div>` +
          `<div style="white-space:pre-wrap;font-family:inherit;font-size:12px">${escapeHtml(original)}</div>`
        : "";

      const formData: any = {
        subject: makeForwardSubject(ctx.subject || ""),
        htmlBody: `${html}${originalHtml}`,
      };
      if (to.length) formData.toRecipients = to;
      if (cc.length) formData.ccRecipients = cc;
      if (bcc.length) formData.bccRecipients = bcc;

      mailbox.displayNewMessageForm(formData);
      setNotice("O Outlook não expôs a API de Forward. Abri uma nova mensagem do tipo FW (com o original no corpo).");
      return;
    }

    throw new Error("Não foi possível abrir um forward (API indisponível).");
  } catch (e: any) {
    setErr(e?.message ?? String(e));
  }
}



  function onToggleInclude(emailAddr: string, checked: boolean) {
    setRecipientPreset("custom");
    setRecipientRows((prev) => prev.map((r) => (r.email === emailAddr ? { ...r, include: checked } : r)));
  }

  function onChangeRole(emailAddr: string, role: RecipientRole) {
    setRecipientPreset("custom");
    setRecipientRows((prev) => prev.map((r) => (r.email === emailAddr ? { ...r, role } : r)));
  }

  function addManualRecipient() {
    const e = normalizeEmailCandidate(addEmail);
    if (!e) return;

    setRecipientPreset("custom");
    setRecipientRows((prev) => {
      const rows = prev.map((r) => ({ ...r, origins: [...r.origins] }));
      const existing = rows.find((r) => r.email === e);
      if (existing) {
        existing.origins = mergeOrigins(existing.origins, ["manual"]);
        existing.include = true;
        existing.role = addRole;
        return rows;
      }
      rows.push({ email: e, origins: ["manual"], include: true, role: addRole });
      return rows;
    });

    setAddEmail("");
    setAddRole("bcc");
  }

  const visibleRecipientRows = useMemo(() => {
    if (includeBodyEmails) return recipientRows;
    return recipientRows.filter((r) => !(r.origins.length === 1 && r.origins[0] === "body"));
  }, [recipientRows, includeBodyEmails]);

  return (
    <div style={S.aiShell}>
      <div style={S.aiHeader}>
        <div style={{ flex: 1 }} />
        <div style={S.aiHeaderBtns}>
          <button style={S.iconBtnHeader} onClick={() => setSheet("options")} title="Opções (modo/tom)">
            ⚙
          </button>
          <button style={S.iconBtnHeader} onClick={() => setSheet("recipients")} title="Destinatários (To/Cc/Bcc)">
            👥
          </button>
          <button style={S.iconBtnHeader} onClick={() => setSheet("context")} title="Contexto do email">
            ✉
          </button>
          <button style={S.iconBtnHeader} onClick={() => setSheet("reminder")} title="Lembrete / Follow-up">
            ⏰
          </button>
        </div>
      </div>

      <div style={S.resultCard}>
        <div style={S.resultTopRow}>
        <div style={S.resultLabel}>Resultado</div>

        <button
          style={resultView === "html" ? S.pillActive : S.pill}
          onClick={() => setResultView("html")}
          title="Ver HTML"
        >
          HTML
        </button>
        <button
          style={resultView === "text" ? S.pillActive : S.pill}
          onClick={() => setResultView("text")}
          title="Ver texto"
        >
          Txt
        </button>

        <button style={S.iconBtn} onClick={() => copyToClipboard(resultView === "html" ? buildHtmlForInsert() : textOut)} title="Copiar">
          📋
        </button>

        <div style={S.optionTabs}>
          {[0, 1, 2].map((i) => (
            <button
              key={i}
              style={activeOption === i ? S.optionTabActive : S.optionTab}
              onClick={() => setActiveOption(i)}
              title={`Opção ${i + 1}`}
            >
              {i + 1}
            </button>
          ))}
        </div>

        <button style={S.iconBtn} onClick={() => setSheet("options")} title="Histórico">
          🕘
        </button>
      </div>

      {!(htmlOut || textOut) && !err ? (
          <div style={S.placeholder}>{busy ? "A gerar…" : "Escolhe uma ação no rodapé para gerar conteúdo."}</div>
        ) : (
          <>
            {resultView === "html" ? (
              <div style={S.preview} dangerouslySetInnerHTML={{ __html: htmlOut || buildHtmlForInsert() }} />
            ) : (
              <pre style={S.pre} title="Texto simples">
                {textOut || ""}
              </pre>
            )}
          </>
        )}

        {notice && <div style={S.notice}>{notice}</div>}
        {err && <div style={S.err}>{err}</div>}
      </div>

      {/* bottom bar */}
      <div style={S.bottomBar}>
        {/* Quick language picker (drop-up) */}
        <div style={S.langQuickWrap}>
          <button
            type="button"
            style={S.langQuickBtn}
            onClick={() => setLangMenuOpen((v) => !v)}
            title={`Idioma (escrita): ${replyLang === "auto" ? "Auto" : (LOCALE_LABEL[replyLang as AppLocale] ?? String(replyLang))}`}
          >
            <span style={S.langQuickCode}>
              {replyLang === "auto"
                ? "🪄"
                : ((LOCALE_SHORT[replyLang as AppLocale] ?? String(replyLang).slice(0, 2).toUpperCase()).slice(0, 2))}
            </span>
          </button>

          {langMenuOpen && (
            <>
              <div style={S.langOverlay} onClick={() => setLangMenuOpen(false)} />
              <div style={S.langMenu}>
                <div style={S.langPillGrid}>
                  {["auto", ...enabledLangs].map((opt) => {
                    const active = replyLang === (opt as any);
                    const label = opt === "auto" ? "Auto" : (LOCALE_LABEL[opt as AppLocale] ?? String(opt));
                    const short = opt === "auto" ? "🪄" : ((LOCALE_SHORT[opt as AppLocale] ?? String(opt).slice(0, 2).toUpperCase()).slice(0, 2));
                    return (
                      <button
                        key={`reply-${opt}`}
                        type="button"
                        title={label}
                        style={{ ...S.langPillBtn, ...(active ? S.langPillBtnActive : {}) }}
                        onClick={async () => {
                          setReplyLang(opt as LangOption);
                          await saveSettings({ replyLanguage: opt as LangOption });
                          setLangMenuOpen(false);
                        }}
                      >
                        {short}
                      </button>
                    );
                  })}
                </div>
              </div>
            </>
          )}
        </div>
<button
          style={busy || !canRun ? S.navBtnDisabled : S.navBtn}
          disabled={busy || !canRun}
          onClick={() => setSheet("compose")}
          title="Resposta sugerida (com notas)"
        >
          <span style={S.navIcon}>✍️</span>
          <div style={S.navTxt}>Responder</div>
        </button>

        <button
          style={busy || !canRun ? S.navBtnDisabled : S.navBtn}
          disabled={busy || !canRun}
          onClick={() => run("tasks")}
          title="Extrair tarefas e dependências"
        >
          <span style={S.navIcon}>✅</span>
          <div style={S.navTxt}>Tarefas</div>
        </button>

        <button
          style={busy ? S.navBtnDisabled : S.navBtn}
          disabled={busy}
          onClick={() => setSheet("rewrite")}
          title="Reescrever um texto"
        >
          <span style={S.navIcon}>♻️</span>
          <div style={S.navTxt}>Reescrever</div>
        </button>
      </div>

      {/* sheets */}
      <BottomSheet open={sheet === "options"} title="Opções" onClose={() => setSheet("")}> 
        <div style={S.sheetBody}>
          <label style={S.fieldLabel}>Modo</label>
          <select style={S.select} value={mode} onChange={(e) => setMode(e.target.value as AiMode)}>
            <option value="fast">Rápido</option>
            <option value="quality">Qualidade</option>
          </select>

          <label style={S.fieldLabel}>Tom</label>
          <select style={S.select} value={tone} onChange={(e) => setTone(e.target.value as AiTone)}>
            <option value="neutro">Neutro</option>
            <option value="formal">Formal</option>
            <option value="direto">Direto</option>
            <option value="simpático">Profissional simpático</option>
            <option value="curto">Curto</option>
          </select>

          <label style={S.fieldLabel}>Idioma (resumo)</label>
          <select
            style={S.select}
            value={readingLang}
            onChange={(e) => {
              const v = e.currentTarget.value as LangOption;
              setReadingLang(v);
              void saveSettings({ readingLanguage: v });
            }}
          >
            {LANG_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.value === "auto" ? "Auto (detetar)" : o.label}
              </option>
            ))}
          </select>

          <label style={S.fieldLabel}>Idioma (resposta)</label>
          <select
            style={S.select}
            value={replyLang}
            onChange={(e) => {
              const v = e.currentTarget.value as LangOption;
              setReplyLang(v);
              void saveSettings({ replyLanguage: v });
            }}
          >
            {LANG_OPTIONS.map((o) => (
              <option key={o.value} value={o.value}>
                {o.value === "auto" ? "Auto (detetar)" : o.label}
              </option>
            ))}
          
          <div style={{ marginTop: 16, paddingTop: 12, borderTop: "1px solid rgba(11,45,107,0.10)" }}>
            <div style={{ fontSize: 12, fontWeight: 800, color: "#0b2d6b", marginBottom: 8 }}>Assinatura</div>

            <label style={S.fieldLabel}>Modo</label>
            <select
              style={S.select}
              value={sigMode}
              onChange={(e) => {
                const v = e.currentTarget.value as SigMode;
                setSigMode(v);
                persistSig(SIG_KEY_MODE, v);
              }}
            >
              <option value="off">Desligada</option>
              <option value="text">Texto</option>
              <option value="html">HTML</option>
              <option value="image">Imagem (upload/URL)</option>
            </select>

            {sigMode === "text" && (
              <>
                <label style={S.fieldLabel}>Texto da assinatura</label>
                <textarea
                  style={S.textarea}
                  rows={4}
                  value={sigText}
                  onChange={(e) => {
                    const v = e.currentTarget.value;
                    setSigText(v);
                    persistSig(SIG_KEY_TEXT, v);
                  }}
                  placeholder="Ex.:
Pedro Lopes
Divitek
+351 ..."
                />
              </>
            )}

            {sigMode === "html" && (
              <>
                <label style={S.fieldLabel}>HTML da assinatura</label>
                <textarea
                  style={S.textarea}
                  rows={4}
                  value={sigHtml}
                  onChange={(e) => {
                    const v = e.currentTarget.value;
                    setSigHtml(v);
                    persistSig(SIG_KEY_HTML, v);
                  }}
                  placeholder='Ex.: <div><b>Pedro Lopes</b><br/>Divitek</div>'
                />
              </>
            )}

            {sigMode === "image" && (
              <>
                <label style={S.fieldLabel}>Imagem (upload)</label>
                <input
                  style={S.input}
                  type="file"
                  accept="image/*"
                  onChange={(e) => {
                    const f = e.currentTarget.files?.[0];
                    if (!f) return;
                    const reader = new FileReader();
                    reader.onload = () => {
                      const v = String(reader.result || "");
                      setSigImgDataUrl(v);
                      persistSig(SIG_KEY_IMG_DATA, v);
                    };
                    reader.readAsDataURL(f);
                  }}
                />
                <div style={{ fontSize: 10, opacity: 0.75, marginTop: 4 }}>
                  Dica: upload guarda a imagem localmente (data URL) e funciona mesmo sem link público.
                </div>

                <label style={S.fieldLabel}>Largura máx. (px)</label>
                <input
                  style={S.input}
                  value={sigImgMaxW}
                  onChange={(e) => {
                    const v = e.currentTarget.value.replace(/[^0-9]/g, "");
                    setSigImgMaxW(v);
                    persistSig(SIG_KEY_IMG_W, v);
                  }}
                  placeholder="260"
                />

                <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                  <button
                    style={S.secondaryBtn}
                    onClick={() => {
                      setSigImgDataUrl("");
                      persistSig(SIG_KEY_IMG_DATA, "");
                    }}
                  >
                    Limpar upload
                  </button>
                </div>

                <div style={{ marginTop: 10, paddingTop: 10, borderTop: "1px solid rgba(11,45,107,0.10)" }}>
                  <label style={S.fieldLabel}>OU URL da imagem</label>
                  <input
                    style={S.input}
                    value={sigImgUrl}
                    onChange={(e) => {
                      const v = e.currentTarget.value;
                      setSigImgUrl(v);
                      persistSig(SIG_KEY_IMG, v);
                    }}
                    placeholder="https://.../assinatura.png"
                  />
                </div>
              </>
            )}
          </div>

</select>

          <div style={S.sheetHint}>Estas opções afetam o estilo e o nível de detalhe.</div>
        </div>
      
        <div style={S.sectionTitle}>Templates</div>
        <div style={S.sectionCard}>
          <div style={S.smallHint}>
            Cria textos reutilizáveis para respostas rápidas. Suporta <code>{"{{nome}}"}</code> e <code>{"{{assunto}}"}</code>.
          </div>

          <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
            <button
              style={S.secondaryBtn}
              onClick={() => {
                const t: SnippetTemplate = { id: `tpl-${Date.now()}`, name: "Novo template", body: "" };
                setTplEdit(t);
                setTplName(t.name);
                setTplBody(t.body);
              }}
            >
              + Novo
            </button>

            <button
              style={S.secondaryBtn}
              onClick={() => {
                setTemplates(defaultTemplates());
              }}
              title="Repor templates base"
            >
              Repor
            </button>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 10 }}>
            {templates.map((t) => (
              <div key={t.id} style={S.tplRow}>
                <div style={{ minWidth: 0 }}>
                  <div style={S.tplName} title={t.name}>
                    {t.name}
                  </div>
                  <div style={S.tplMeta}>{(t.body || "").trim().slice(0, 90) || "—"}</div>
                </div>

                <div style={{ display: "flex", gap: 6 }}>
                  <button
                    style={S.smallBtn}
                    onClick={() => {
                      setTplEdit(t);
                      setTplName(t.name);
                      setTplBody(t.body);
                    }}
                  >
                    Editar
                  </button>
                  <button style={S.smallBtn} onClick={() => setTemplates((prev) => prev.filter((x) => x.id !== t.id))}>
                    🗑
                  </button>
                </div>
              </div>
            ))}
          </div>

          {tplEdit && (
            <div style={S.tplEditor}>
              <div style={S.tplEditorTitle}>Editar template</div>

              <label style={S.fieldLabel}>Nome</label>
              <input style={S.input} value={tplName} onChange={(e) => setTplName(e.currentTarget.value)} />

              <label style={S.fieldLabel}>Corpo</label>
              <textarea
                style={S.textarea}
                value={tplBody}
                onChange={(e) => setTplBody(e.currentTarget.value)}
                rows={9}
                placeholder="Escreve aqui o texto do template…"
              />

              <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
                <button
                  style={S.primaryBtn}
                  onClick={() => {
                    const next: SnippetTemplate = { id: tplEdit.id, name: (tplName || "Sem nome").trim(), body: tplBody || "" };
                    setTemplates((prev) => {
                      const idx = prev.findIndex((x) => x.id === next.id);
                      if (idx >= 0) {
                        const copy = [...prev];
                        copy[idx] = next;
                        return copy;
                      }
                      return [next, ...prev];
                    });
                    setTplEdit(null);
                    setTplName("");
                    setTplBody("");
                  }}
                >
                  Guardar
                </button>
                <button
                  style={S.secondaryBtn}
                  onClick={() => {
                    setTplEdit(null);
                    setTplName("");
                    setTplBody("");
                  }}
                >
                  Cancelar
                </button>
              </div>
            </div>
          )}

          <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px solid rgba(11,45,107,0.10)" }}>
            <label style={S.fieldLabel}>Histórico (5 dias)</label>
            <div style={S.muted}>Guarda outputs por email (por item) para poderes retomar mais tarde.</div>

            <select
              style={{ ...S.select, marginTop: 8 }}
              value={historyPickId}
              onChange={(e) => {
                const id = e.currentTarget.value;
                setHistoryPickId(id);
                const item = aiHistory.find((x) => x.id === id);
                if (!item) return;
                restoringRef.current = true;
                setHtmlOut(item.html || "");
                setTextOut(item.text || "");
                window.setTimeout(() => {
                  restoringRef.current = false;
                }, 0);
              }}
            >
              <option value="">— Selecionar —</option>
              {aiHistory
                .filter((x) => x.emailKey === emailKey)
                .slice(-20)
                .reverse()
                .map((x) => {
                  const d = new Date(x.ts);
                  const label = `${d.toLocaleDateString()} ${d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })} — ${
                    (x.subject || "").slice(0, 40) || "Resposta"
                  }`;
                  return (
                    <option key={x.id} value={x.id}>
                      {label}
                    </option>
                  );
                })}
            </select>
          </div>

        </div>

</BottomSheet>

      <BottomSheet open={sheet === "context"} title="Contexto do email" onClose={() => setSheet("")}> 
        <div style={S.sheetBody}>
          <label style={S.fieldLabel}>Corpo do email (para a IA)</label>

          <div style={S.scopeRow}>
            <label style={S.scopeOpt} title="Menos ruído: normalmente chega para responder.">
              <input
                type="radio"
                name="bodyScope"
                checked={bodyScope === "main"}
                onChange={() => setBodyScope("main")}
              />
              <span>Mensagem principal</span>
            </label>
            <label style={S.scopeOpt} title="Inclui reencaminhados/citações (mais completo, pode ser mais longo).">
              <input
                type="radio"
                name="bodyScope"
                checked={bodyScope === "full"}
                onChange={() => setBodyScope("full")}
              />
              <span>Corpo completo</span>
            </label>
          </div>

          <textarea
            style={S.textarea}
            value={bodyScope === "full" ? fullBody : body}
            onChange={(e) => (bodyScope === "full" ? setFullBody(e.target.value) : setBody(e.target.value))}
            placeholder="(não consegui ler o corpo automaticamente — podes colar aqui)"
            rows={10}
          />
          <div style={S.sheetHint}>
            {(bodyScope === "full" ? fullBody.length : body.length)} caracteres ·{" "}
            <a
              style={S.inlineLink}
              href="#"
              onClick={(e) => {
                e.preventDefault();
                if (bodyScope === "full") setFullBody(trimEmailBodyFull(fullBody));
                else setBody(trimEmailBody(body));
              }}
            >
              limpar
            </a>
          </div>
        </div>
      </BottomSheet>

      
      <BottomSheet open={sheet === "reminder"} title="Lembrete" onClose={() => setSheet("")}>
        <div style={S.sheetBody}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
            <div style={{ fontWeight: 500, color: "#0b2d6b" }}>Marcar follow-up</div>
            <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, opacity: 0.9 }}>
              <input
                type="checkbox"
                checked={reminderCreateEvent}
                onChange={(e) => setReminderCreateEvent(e.currentTarget.checked)}
              />
              <span>Criar evento no calendário</span>
            </label>
          </div>

          <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 10 }}>
            <button style={busy ? S.smallBtnDisabled : S.smallBtn} disabled={busy} onClick={() => applyReminder(makeDueDate(0))}>
              Hoje
            </button>
            <button style={busy ? S.smallBtnDisabled : S.smallBtn} disabled={busy} onClick={() => applyReminder(makeDueDate(1))}>
              Amanhã
            </button>
            <button style={busy ? S.smallBtnDisabled : S.smallBtn} disabled={busy} onClick={() => applyReminder(makeDueDate(2))}>
              2 dias
            </button>
            <button style={busy ? S.smallBtnDisabled : S.smallBtn} disabled={busy} onClick={() => applyReminder(makeDueDate(7))}>
              7 dias
            </button>
          </div>

          <div style={{ marginTop: 12, display: "flex", gap: 10, alignItems: "center" }}>
            <label style={{ fontSize: 12, fontWeight: 600, opacity: 0.9 }}>Data:</label>
            <input
              type="date"
              value={reminderDate}
              onChange={(e) => setReminderDate(e.currentTarget.value)}
              style={{ padding: "8px 10px", borderRadius: 10, border: "1px solid rgba(11,45,107,0.18)" }}
            />
            <button
              style={busy ? S.primaryBtnDisabled : S.primaryBtn}
              disabled={busy}
              onClick={() => {
                const d = parseCustomDate(reminderDate);
                if (!d) {
                  setNotice("Escolhe uma data válida.");
                  return;
                }
                void applyReminder(d);
              }}
            >
              Aplicar
            </button>
          </div>

          <div style={S.sheetHint}>
            Isto aplica a categoria <b>{REMINDER_CATEGORY}</b> ao email. Se a opção estiver ativa, abre também um evento no calendário (15 min, 09:00).
          </div>
        </div>
      </BottomSheet>


      <BottomSheet open={sheet === "compose"} title="Responder" onClose={() => setSheet("")}> 
        <div style={S.sheetBody}>
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 10 }}>
          <select
            style={S.select}
            value={tplPickId}
            onChange={(e) => {
              const id = e.currentTarget.value;
              setTplPickId(id);
              const t = templates.find((x) => x.id === id);
              if (t) {
                const filled = applyTemplateVars(t.body, ctx);
                setComposeNotes(filled);
              }
            }}
          >
            <option value="">— Template —</option>
            {templates.map((t) => (
              <option key={t.id} value={t.id}>
                {t.name}
              </option>
            ))}
          </select>

          <button style={S.secondaryBtn} onClick={() => setSheet("options")} title="Gerir templates nas opções">
            Gerir
          </button>
        </div>

<label style={S.fieldLabel}>O que queres transmitir?</label>
          <textarea
            style={S.textarea}
            value={composeNotes}
            onChange={(e) => setComposeNotes(e.target.value)}
            placeholder="Ex.: agradecer, pedir confirmação de medidas, indicar prazos, anexar proposta, etc."
            rows={7}
          />

          <div style={S.sheetHint}>A IA vai usar o email + estas notas para criar uma resposta.</div>

          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <button
              style={busy || !canRun ? S.primaryBtnDisabled : S.primaryBtn}
              disabled={busy || !canRun}
              onClick={() => { setRecipientPreset(replyAll ? "replyAll" : "reply"); run("reply"); }}
              title="Gerar resposta sugerida"
            >
              {busy ? "A gerar…" : "Gerar resposta"}
            </button>

            <button style={S.secondaryBtn} onClick={() => setSheet("recipients")} title="Escolher To/Cc/Bcc">
              Destinatários
            </button>


          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 10 }}>
            <input
              id="replyAllToggle"
              type="checkbox"
              checked={replyAll}
              onChange={(e) => setReplyAll(e.currentTarget.checked)}
            />
            <label htmlFor="replyAllToggle" style={{ fontSize: 12, fontWeight: 500, color: "#0b2d6b" }}>
              Responder a todos
            </label>
          </div>
          </div>
        </div>
      </BottomSheet>

      <BottomSheet open={sheet === "rewrite"} title="Reescrever" onClose={() => setSheet("")}> 
        <div style={S.sheetBody}>
          <label style={S.fieldLabel}>Texto</label>
          <textarea
            style={S.textarea}
            value={rewriteText}
            onChange={(e) => setRewriteText(e.target.value)}
            placeholder="Cola aqui uma frase/parágrafo para reescrever…"
            rows={8}
          />
          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <button
              style={busy || !rewriteText.trim() ? S.primaryBtnDisabled : S.primaryBtn}
              disabled={busy || !rewriteText.trim()}
              onClick={() => run("rewrite")}
              title="Gerar nova versão"
            >
              {busy ? "A gerar…" : "Reescrever"}
            </button>
            <div style={{ flex: 1 }} />
            <div style={S.sheetHint}>Dica: seleciona “Curto” em Opções para encurtar.</div>
          </div>
        </div>
      </BottomSheet>

      <BottomSheet open={sheet === "recipients"} title="Destinatários" onClose={() => setSheet("")}> 
        <div style={S.sheetBody}>
          <div style={S.sheetHint}>
            “Inserir thread” usa Responder/Responder a todos (o Outlook pode mexer em To/Cc). Para controlo total,
            usa “Nova msg”.
          </div>

          <div style={S.presetRow}>
            <button
              style={recipientPreset === "reply" ? S.presetBtnActive : S.presetBtn}
              onClick={() => applyRecipientPreset("reply")}
              title="Responder apenas ao remetente"
            >
              Responder
            </button>
            <button
              style={recipientPreset === "replyAll" ? S.presetBtnActive : S.presetBtn}
              onClick={() => applyRecipientPreset("replyAll")}
              title="Responder a todos (remetente + restantes em Cc)"
            >
              Responder a todos
            </button>
            <button
              style={recipientPreset === "custom" ? S.presetBtnActive : S.presetBtn}
              onClick={() => setRecipientPreset("custom")}
              title="Escolha manual"
            >
              Manual
            </button>

            <div style={{ flex: 1 }} />

            <label style={S.toggleRow} title="Mostra/oculta emails encontrados no corpo do email">
              <input
                type="checkbox"
                checked={includeBodyEmails}
                onChange={(e) => setIncludeBodyEmails(e.target.checked)}
              />
              <span>Emails do corpo</span>
            </label>

            <label style={S.toggleRow} title="Ao criar Reply/Nova msg, anexa o email original como .msg">
              <input
                type="checkbox"
                checked={attachOriginalItem}
                onChange={(e) => setAttachOriginalItem(e.target.checked)}
              />
              <span>Anexar original (.msg)</span>
            </label>
          </div>

          <div style={S.recList}>
            {visibleRecipientRows.length ? (
              visibleRecipientRows
                .slice()
                .sort((a, b) => {
                  const aIsFrom = a.origins.includes("from") ? 0 : 1;
                  const bIsFrom = b.origins.includes("from") ? 0 : 1;
                  if (aIsFrom !== bIsFrom) return aIsFrom - bIsFrom;
                  if (a.include !== b.include) return a.include ? -1 : 1;
                  return a.email.localeCompare(b.email);
                })
                .map((r) => {
                  const label = r.name ? `${r.name} <${r.email}>` : r.email;
                  const src = originLabel(r);
                  const badgeStyle = src.tone === "hi" ? S.badgeHi : src.tone === "mid" ? S.badgeMid : S.badgeLow;

                  return (
                    <div key={r.email} style={S.recLine}>
                      <input
                        type="checkbox"
                        checked={!!r.include}
                        onChange={(ev) => onToggleInclude(r.email, ev.target.checked)}
                        title="Incluir/Excluir"
                      />

                      <span style={{ ...S.badge, ...badgeStyle }}>{src.text}</span>

                      <div style={S.recMain} title={label}>
                        <div style={S.recEmail}>{label}</div>
                      </div>

                      <select
                        style={S.roleSelect}
                        value={r.role}
                        onChange={(e) => onChangeRole(r.email, e.target.value as RecipientRole)}
                        disabled={!r.include}
                        title="Enviar como"
                      >
                        <option value="to">To</option>
                        <option value="cc">Cc</option>
                        <option value="bcc">Bcc</option>
                      </select>
                    </div>
                  );
                })
            ) : (
              <div style={S.muted}>—</div>
            )}
          </div>

          <div style={S.addRow}>
            <input
              style={S.addInput}
              value={addEmail}
              onChange={(e) => setAddEmail(e.target.value)}
              placeholder="Adicionar email (ex.: nome@empresa.pt)"
            />
            <select style={S.addSelect} value={addRole} onChange={(e) => setAddRole(e.target.value as RecipientRole)}>
              <option value="to">To</option>
              <option value="cc">Cc</option>
              <option value="bcc">Bcc</option>
            </select>
            <button
              style={!normalizeEmailCandidate(addEmail) ? S.smallBtnDisabled : S.smallBtn}
              disabled={!normalizeEmailCandidate(addEmail)}
              onClick={addManualRecipient}
              title="Adicionar"
            >
              + Adicionar
            </button>
          </div>

          <div style={{ display: "flex", gap: 8, marginTop: 10 }}>
            <button style={S.secondaryBtn} onClick={() => setSheet("compose")} title="Voltar">
              Voltar
            </button>
            <div style={{ flex: 1 }} />
            <button style={S.primaryBtn} onClick={() => setSheet("")} title="Fechar">
              OK
            </button>
          </div>
        </div>
      </BottomSheet>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  aiShell: {
    width: "100%",
    maxWidth: "100%",
    margin: 0,
    paddingBottom: 108, // space for bottom bar (evita sobreposição)
  },

  aiHeader: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 8,
    padding: "6px 2px 10px 2px",
  },
  aiTitle: { fontWeight: 400, fontSize: 12, lineHeight: "14px", color: "#0b2d6b" },
  aiSub: {
    fontWeight: 500,
    fontSize: 10,
    color: "rgba(11,45,107,0.70)",
    marginTop: 2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  aiHeaderBtns: { display: "flex", gap: 6 },

  iconBtnHeader: {
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.85)",
    borderRadius: 10,
    padding: "3px 6px",
    cursor: "pointer",
    fontSize: 11,
    color: "#0b2d6b",
    lineHeight: 1,
  },

  resultCard: {
    borderRadius: 16,
    padding: 10,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.88)",
    boxShadow: "0 1px 10px rgba(11,45,107,0.06)",
    boxSizing: "border-box",
  },
  resultTopRow: { display: "flex", gap: 6, alignItems: "center", flexWrap: "nowrap", overflowX: "auto" },
  resultLabel: { fontWeight: 400, fontSize: 12, color: "rgba(11,45,107,0.90)" },

  smallBtn: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.16)",
    background: "rgba(239,246,255,0.75)",
    color: "#0b2d6b",
    fontWeight: 400,
    cursor: "pointer",
    fontSize: 10,
  },
  smallBtnActive: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.22)",
    background: "rgba(11,45,107,0.92)",
    color: "white",
    fontWeight: 800,
    cursor: "pointer",
    fontSize: 10,
  },
  smallBtnDisabled: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.10)",
    background: "rgba(11,45,107,0.05)",
    color: "rgba(11,45,107,0.45)",
    fontWeight: 400,
    cursor: "not-allowed",
    fontSize: 10,
  },
  optTabs: { display: "inline-flex", gap: 4, alignItems: "center" },
  optTab: {
    padding: "1px 6px",
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.70)",
    color: "#0b2d6b",
    fontSize: 10,
    fontWeight: 400,
    lineHeight: "12px",
    borderRadius: 999,
    cursor: "pointer",
  },
  optTabActive: {
    border: "1px solid rgba(11,45,107,0.28)",
    background: "rgba(11,45,107,0.10)",
  },


  iconBtn: {
    borderRadius: 999,
    padding: "3px 8px",
    width: 28,
    minWidth: 28,
    border: "1px solid rgba(11,45,107,0.16)",
    background: "rgba(255,255,255,0.70)",
    color: "#0b2d6b",
    fontWeight: 400,
    cursor: "pointer",
    fontSize: 11,
    lineHeight: "12px",
  },
  iconBtnDisabled: {
    borderRadius: 999,
    padding: "3px 8px",
    width: 28,
    minWidth: 28,
    border: "1px solid rgba(11,45,107,0.10)",
    background: "rgba(11,45,107,0.05)",
    color: "rgba(11,45,107,0.45)",
    fontWeight: 400,
    cursor: "not-allowed",
    fontSize: 11,
    lineHeight: "12px",
  },

  placeholder: {
    marginTop: 10,
    borderRadius: 14,
    padding: 12,
    border: "1px dashed rgba(11,45,107,0.18)",
    background: "rgba(239,246,255,0.55)",
    color: "rgba(11,45,107,0.75)",
    fontSize: 10,
  },

  preview: {
    marginTop: 10,
    borderRadius: 14,
    padding: 10,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.96)",
    fontSize: 10,
    lineHeight: 1.35,
    color: "#0b2d6b",
    overflowWrap: "anywhere",
    maxHeight: 220,
    overflow: "auto",
  },

  pre: {
    marginTop: 10,
    marginBottom: 0,
    whiteSpace: "pre-wrap",
    borderRadius: 14,
    padding: 10,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.96)",
    fontSize: 10,
    lineHeight: 1.35,
    color: "#0b2d6b",
    overflowWrap: "anywhere",
    maxHeight: 220,
    overflow: "auto",
  },

  notice: {
    marginTop: 10,
    borderRadius: 14,
    padding: 10,
    border: "1px solid rgba(206, 149, 0, 0.35)",
    background: "rgba(255, 235, 170, 0.35)",
    color: "rgba(94, 64, 0, 0.95)",
    fontSize: 10,
  },

  err: {
    marginTop: 10,
    borderRadius: 14,
    padding: 10,
    border: "1px solid rgba(176, 40, 40, 0.25)",
    background: "rgba(176, 40, 40, 0.06)",
    color: "rgba(176, 40, 40, 0.95)",
    fontSize: 10,
  },
  bottomBar: {
    position: "fixed",
    left: 10,
    right: 10,
    bottom: 10,
    zIndex: 50,
    margin: 0,
    padding: "6px 6px",
    borderRadius: 18,
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.68)",
    backdropFilter: "blur(10px)",
    WebkitBackdropFilter: "blur(10px)",
    boxShadow: "0 8px 26px rgba(11,45,107,0.12)",
    display: "grid",
    gridTemplateColumns: "68px repeat(4, minmax(0, 1fr))",
    gap: 6,
    alignItems: "stretch",
    justifyItems: "stretch",
    boxSizing: "border-box",
    height: 64,
  },

  langMenuOverlay: {
    position: "fixed",
    inset: 0,
    background: "transparent",
    zIndex: 40,
  },
  langOverlay: { position: "fixed", inset: 0, background: "transparent", zIndex: 40 },
  langQuickWrap: {
    position: "relative",
    width: 64,
    flex: "0 0 64px",
  },
  langQuickBtn: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "3px 6px",
    borderRadius: 999,
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.92)",
    cursor: "pointer",
  },
  langQuickTop: {
    display: "inline-flex",
    alignItems: "center",
    gap: 6,
    fontSize: 10,
    fontWeight: 500,
    letterSpacing: 0.2,
    color: "#0b2d6b",
    lineHeight: "12px",
  },
  langQuickCode: {
    fontSize: 10,
    fontWeight: 800,
    color: "#0b2d6b",
    letterSpacing: 0.6,
  },
  langQuickCaret: {
    fontSize: 10,
    opacity: 0.65,
  },
  langMenu: {
    position: "absolute",
    left: 0,
    bottom: 54,
    zIndex: 60,
    borderRadius: 14,
    padding: 8,
    background: "rgba(255,255,255,0.98)",
    boxShadow: "0 10px 30px rgba(0,0,0,0.15)",
    border: "1px solid rgba(11,45,107,0.10)",
  },
  langPillGrid: {
    display: "flex",
    flexDirection: "column",
    gap: 4,
    alignItems: "flex-start",
  },
  langPillBtn: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    borderRadius: 999,
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.96)",
    padding: "3px 6px",
    fontSize: 10,
    fontWeight: 400,
    letterSpacing: 0.1,
    color: "#0b2d6b",
    cursor: "pointer",
    userSelect: "none",
    lineHeight: 1,
  },
  langPillBtnActive: {
    border: "1px solid rgba(11,45,107,0.28)",
    background: "rgba(11,45,107,0.08)",
  },

  langMenuList: {
    display: "flex",
    flexDirection: "column",
    gap: 6,
    maxHeight: 240,
    overflowY: "auto",
  },
  langMenuItem: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    width: "100%",
    textAlign: "left",
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.10)",
    background: "rgba(255,255,255,0.96)",
    padding: "9px 10px",
    cursor: "pointer",
    userSelect: "none",
  },
  langMenuItemActive: {
    border: "1px solid rgba(11,45,107,0.22)",
    background: "rgba(11,45,107,0.06)",
  },
  langMenuItemCode: {
    flex: "0 0 auto",
    fontSize: 10,
    fontWeight: 800,
    letterSpacing: 0.6,
    color: "#0b2d6b",
    borderRadius: 8,
    padding: "4px 6px",
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(11,45,107,0.04)",
    lineHeight: 1,
  },
  langMenuItemLabel: {
    flex: "1 1 auto",
    fontSize: 11,
    fontWeight: 600,
    color: "#0b2d6b",
    opacity: 0.9,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  langMenuSection: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  langMenuRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
  },
  langMenuLabel: {
    width: 38,
    fontSize: 10,
    fontWeight: 800,
    color: "rgba(11,45,107,0.70)",
    textTransform: "uppercase",
    letterSpacing: 0.6,
  },
  langMenuPills: {
    display: "flex",
    flexWrap: "wrap",
    gap: 6,
  },
  langPill: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(11,45,107,0.04)",
    color: "#0b2d6b",
    fontSize: 10,
    fontWeight: 800,
    cursor: "pointer",
    userSelect: "none",
  },
  langPillActive: {
    border: "1px solid rgba(11,45,107,0.22)",
    background: "rgba(11,45,107,0.12)",
  },
  navBtn: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: 1,
    padding: "5px 4px",
    minHeight: 50,
    borderRadius: 14,
    border: "1px solid transparent",
    background: "transparent",
    cursor: "pointer",
    userSelect: "none",
    boxSizing: "border-box",
    width: "100%",
    minWidth: 0,
    overflow: "hidden",
  },
  navBtnDisabled: {
    flex: 1,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    gap: 1,
    padding: "5px 4px",
    minHeight: 50,
    borderRadius: 14,
    border: "1px solid rgba(11,45,107,0.10)",
    background: "rgba(11,45,107,0.06)",
    color: "rgba(11,45,107,0.35)",
    cursor: "not-allowed",
    userSelect: "none",
    boxSizing: "border-box",
    width: "100%",
    minWidth: 0,
    overflow: "hidden",
  },
  navIcon: { fontSize: 16, lineHeight: 1, marginBottom: 0 },
  navTxt: { fontSize: 9, lineHeight: 1.05, fontWeight: 800, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", maxWidth: "100%" },

  sheetOverlay: {
    position: "fixed",
    inset: 0,
    background: "rgba(11,45,107,0.18)",
    display: "flex",
    alignItems: "flex-end",
    justifyContent: "center",
    zIndex: 9999,
  },
  sheet: {
    width: "calc(100% - 20px)",
    maxWidth: "100%",
    margin: "0 auto 10px auto",
    borderRadius: 20,
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.96)",
    boxShadow: "0 16px 50px rgba(11,45,107,0.18)",
    overflow: "hidden",
    maxHeight: "72vh",
    display: "flex",
    flexDirection: "column",
  },
  sheetHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 8,
    padding: "10px 10px 8px 12px",
    borderBottom: "1px solid rgba(11,45,107,0.10)",
  },
  sheetTitle: { fontWeight: 750, fontSize: 12, color: "#0b2d6b" },
  sheetBody: { padding: 12, overflow: "auto" },

  fieldLabel: { display: "block", fontSize: 10, fontWeight: 500, color: "rgba(11,45,107,0.85)", marginBottom: 6 },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.16)",
    padding: "8px 10px",
    fontSize: 11,
    outline: "none",
    marginBottom: 10,
  },
  textarea: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.16)",
    padding: "10px 10px",
    fontSize: 11,
    outline: "none",
    resize: "vertical",
    boxSizing: "border-box",
    maxHeight: 160,
  },
  sheetHint: { marginTop: 8, fontSize: 10, color: "rgba(11,45,107,0.65)" },
  inlineLink: { color: "#0b2d6b", fontWeight: 500, textDecoration: "none" },

  primaryBtn: {
    borderRadius: 14,
    padding: "9px 12px",
    border: "1px solid rgba(11,45,107,0.18)",
    background: "rgba(11,45,107,0.92)",
    color: "white",
    fontWeight: 750,
    cursor: "pointer",
  },

  sectionTitle: {
    marginTop: 14,
    marginBottom: 8,
    fontSize: 11,
    fontWeight: 800,
    color: "#0b2d6b",
    letterSpacing: 0.2,
  },
  sectionCard: {
    border: "1px solid rgba(11,45,107,0.10)",
    borderRadius: 16,
    padding: 12,
    background: "rgba(255,255,255,0.92)",
  },
  smallHint: {
    fontSize: 10,
    color: "#0b2d6b",
    opacity: 0.7,
    lineHeight: "15px",
  },
  tplRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    padding: 10,
    border: "1px solid rgba(11,45,107,0.10)",
    borderRadius: 14,
    background: "rgba(255,255,255,0.9)",
  },
  tplName: {
    fontWeight: 800,
    fontSize: 13,
    color: "#0b2d6b",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  tplMeta: {
    fontSize: 10,
    color: "#0b2d6b",
    opacity: 0.6,
    marginTop: 2,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  tplEditor: {
    marginTop: 12,
    padding: 12,
    borderRadius: 16,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(11,45,107,0.03)",
  },
  tplEditorTitle: {
    fontWeight: 900,
    color: "#0b2d6b",
    marginBottom: 8,
  },
  primaryBtnDisabled: {
    borderRadius: 14,
    padding: "9px 12px",
    border: "1px solid rgba(11,45,107,0.10)",
    background: "rgba(11,45,107,0.20)",
    color: "rgba(255,255,255,0.75)",
    fontWeight: 750,
    cursor: "not-allowed",
  },
  secondaryBtn: {
    borderRadius: 14,
    padding: "9px 12px",
    border: "1px solid rgba(11,45,107,0.18)",
    background: "rgba(239,246,255,0.80)",
    color: "#0b2d6b",
    fontWeight: 750,
    cursor: "pointer",
  },

  presetRow: { display: "flex", gap: 6, alignItems: "center", marginTop: 10, marginBottom: 10, flexWrap: "wrap" },
  presetBtn: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(239,246,255,0.55)",
    color: "#0b2d6b",
    fontWeight: 750,
    cursor: "pointer",
    fontSize: 10,
  },
  presetBtnActive: {
    borderRadius: 999,
    padding: "3px 8px",
    border: "1px solid rgba(11,45,107,0.18)",
    background: "rgba(11,45,107,0.92)",
    color: "white",
    fontWeight: 750,
    cursor: "pointer",
    fontSize: 10,
  },

  toggleRow: { display: "flex", alignItems: "center", gap: 6, fontSize: 10, color: "rgba(11,45,107,0.85)" },

  scopeRow: { display: "flex", gap: 12, alignItems: "center", marginBottom: 8, flexWrap: "wrap" },
  scopeOpt: { display: "flex", gap: 6, alignItems: "center", fontSize: 10, color: "rgba(11,45,107,0.90)" },

  recList: {
    borderRadius: 14,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(239,246,255,0.35)",
    padding: 8,
    maxHeight: 240,
    overflow: "auto",
  },
  recLine: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "6px 6px",
    borderRadius: 10,
    background: "rgba(255,255,255,0.85)",
    border: "1px solid rgba(11,45,107,0.08)",
    marginBottom: 6,
  },
  recMain: { minWidth: 0, flex: 1 },
  recEmail: {
    fontSize: 10,
    color: "#0b2d6b",
    fontWeight: 400,
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
  },
  roleSelect: {
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.16)",
    padding: "3px 6px",
    fontSize: 10,
    outline: "none",
    background: "white",
  },

  badge: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    minWidth: 62,
    padding: "4px 8px",
    borderRadius: 999,
    fontSize: 10,
    fontWeight: 800,
    border: "1px solid rgba(11,45,107,0.12)",
    userSelect: "none",
  },
  badgeHi: { background: "rgba(11,45,107,0.12)", color: "#0b2d6b" },
  badgeMid: { background: "rgba(11,45,107,0.07)", color: "rgba(11,45,107,0.85)" },
  badgeLow: { background: "rgba(11,45,107,0.04)", color: "rgba(11,45,107,0.70)" },

  addRow: { display: "flex", gap: 6, alignItems: "center", marginTop: 10 },
  addInput: {
    flex: 1,
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.16)",
    padding: "8px 10px",
    fontSize: 11,
    outline: "none",
    minWidth: 0,
  },
  addSelect: {
    borderRadius: 10,
    border: "1px solid rgba(11,45,107,0.16)",
    padding: "8px 8px",
    fontSize: 11,
    outline: "none",
    background: "white",
  },

  muted: { fontSize: 10, color: "rgba(11,45,107,0.55)" },
};
