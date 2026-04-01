import { clientLog } from "./logger";
import { getLinks, getRelatedEmailContext } from "./api";
import { getSettings } from "./settings";
import {
  buildOutlookCategorySourceFromRelatedContext,
  GROUP_CATEGORY_PREFIX,
  REFERENCE_CATEGORY_PREFIX,
  TICKET_CATEGORY_PREFIX,
  LEGACY_STATUS_CATEGORY_PREFIX,
  GROUP_STATUS_CATEGORY_PREFIX,
  TICKET_STATUS_CATEGORY_PREFIX,
  LABEL_STATUS_CATEGORY_PREFIX,
  LEGACY_TICKET_CATEGORY_PREFIX,
  LEGACY_LABEL_CATEGORY_PREFIX,
  ODOO_LINKED_CATEGORY,
  buildOutlookCategoryPlan,
  buildOutlookCategorySourceFromLegacyInput,
  isManagedCategoryFamilyName,
  isReservedOutlookCategoryName,
  normalizeUniqueCategoryValues,
  type LegacyManagedOutlookCategoryInput,
  type OutlookCategoryPlan,
  type OutlookCategorySource,
} from "./outlookCategories";



export type Recipient = { name: string; email: string };

export type OutlookMessageContext = {
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  internetMessageId?: string;
  conversationId?: string;
  itemId?: string;
  receivedDateTimeIso?: string;

  toRecipients?: Recipient[];
  ccRecipients?: Recipient[];
  isCompose?: boolean;
};

export type OutlookAttachment = {
  id?: string;
  name: string;
  contentType: string;
  content: string;
  size?: number;
  isInline?: boolean;
  contentId?: string;
};

const GRAPH_NAA_CLIENT_ID = "67f42759-4576-461a-b87b-c78332f7a1e7";
const GRAPH_NAA_TENANT_ID = "7eff32de-43f1-447b-af3e-af8d2939e93d";
const GRAPH_NAA_AUTHORITY = `https://login.microsoftonline.com/${GRAPH_NAA_TENANT_ID}`;
const GRAPH_ATTACHMENT_SCOPES = ["Mail.Read", "User.Read"];
const GRAPH_PEOPLE_SCOPES = ["People.Read", "User.Read"];
const GRAPH_NAA_REDIRECT_URI = `${window.location.origin}/`;
const GRAPH_RUNTIME_ENABLED = false;

let nestableMsalPromise: Promise<any> | null = null;
export const OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT = "iccc-outlook-category-context-invalidated";

function sleep(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
}

function dispatchOutlookCategoryContextInvalidated() {
  if (typeof window === "undefined" || typeof window.dispatchEvent !== "function") return;
  try {
    window.dispatchEvent(new CustomEvent(OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT));
  } catch {
    // best effort only
  }
}

function collectKnownOutlookCategoryLabelNames(input: {
  settings: Awaited<ReturnType<typeof getSettings>> | null;
  email: any;
  groups: any[];
  tickets: any[];
}): string[] {
  return Array.from(new Set([
    ...(Array.isArray(input.settings?.groupLabelCatalog)
      ? input.settings.groupLabelCatalog.map((entry) => String(entry?.label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.labels)
      ? input.email.labels.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.removedInheritedLabels)
      ? input.email.removedInheritedLabels.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.email?.classificationMeta?.categorizedLabelNames)
      ? input.email.classificationMeta.categorizedLabelNames.map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.groups)
      ? input.groups.flatMap((group: any) => Array.isArray(group?.labels) ? group.labels : []).map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
    ...(Array.isArray(input.tickets)
      ? input.tickets.flatMap((ticket: any) => Array.isArray(ticket?.labels) ? ticket.labels : []).map((label: unknown) => String(label || "").trim()).filter(Boolean)
      : []),
  ]));
}

async function waitForOffice(maxWaitMs = 5000): Promise<any> {
  const start = Date.now();
  while (true) {
    const OfficeAny = (window as any).Office;
    if (OfficeAny) return OfficeAny;
    if (Date.now() - start > maxWaitMs) return null;
    await sleep(50);
  }
}

async function withTimeout<T>(promise: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: any;
  const timeoutPromise = new Promise<T>((resolve) => {
    timer = setTimeout(() => resolve(fallback), ms);
  });
  return Promise.race([
    promise.then(v => { clearTimeout(timer); return v; }),
    timeoutPromise
  ]);
}

async function ensureOfficeReady(): Promise<any> {
  const OfficeAny = await waitForOffice(5000);
  if (!OfficeAny) throw new Error("Office.js não está disponível (Script não carregou).");

  // Bypass: If legacy initialize already fired and we have context, don't wait for onReady (it might hang)
  if (OfficeAny.context) {
    return OfficeAny;
  }

  clientLog.log("[office] waiting for onReady...");
  // Modern Office.onReady returns a promise. We wait up to 10s for the handshake.
  const ready = await withTimeout(OfficeAny.onReady(), 10000, null);

  if (!ready) {
    // Final check before throwing: maybe context appeared meanwhile?
    if (OfficeAny.context) return OfficeAny;
    clientLog.error("[office] onReady timeout (10s)");
    throw new Error("O Office.js demorou demasiado tempo a carregar (Handshake timeout). Reabre o Cockpit.");
  }

  clientLog.log("[office] onReady finished");
  return OfficeAny;
}

async function getNestableMsalInstance(): Promise<any | null> {
  const OfficeAny = await ensureOfficeReady();
  const supportsNaa = Boolean(OfficeAny?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1"));
  if (!supportsNaa) return null;
  if (!nestableMsalPromise) {
    nestableMsalPromise = (async () => {
      const { createNestablePublicClientApplication } = await import("@azure/msal-browser");
      return await createNestablePublicClientApplication({
        auth: {
          clientId: GRAPH_NAA_CLIENT_ID,
          authority: GRAPH_NAA_AUTHORITY,
          redirectUri: GRAPH_NAA_REDIRECT_URI,
        },
        cache: {
          cacheLocation: "localStorage",
        },
      });
    })().catch((error) => {
      nestableMsalPromise = null;
      throw error;
    });
  }
  return await nestableMsalPromise;
}

async function acquireGraphTokenWithNaa(scopes: string[], options?: { allowPrompts?: boolean }): Promise<string> {
  try {
    const msal = await getNestableMsalInstance();
    if (!msal) return "";
    const { InteractionRequiredAuthError } = await import("@azure/msal-browser");
    const tokenRequest = { scopes };
    try {
      const authResult = await msal.acquireTokenSilent(tokenRequest);
      return String(authResult?.accessToken || "").trim();
    } catch (error: any) {
      const interactionRequired =
        error instanceof InteractionRequiredAuthError
        || String(error?.name || "").trim() === "InteractionRequiredAuthError"
        || /interaction_required|consent_required|login_required/i.test(String(error?.errorCode || error?.message || ""));
      if (!interactionRequired) throw error;
      if (options?.allowPrompts === false) return "";
      const authResult = await msal.acquireTokenPopup(tokenRequest);
      return String(authResult?.accessToken || "").trim();
    }
  } catch (error) {
    clientLog.warn("[office] NAA Graph token acquisition failed", error);
    return "";
  }
}

async function getGraphAccessToken(scopes: string[], options?: { allowPrompts?: boolean }): Promise<string> {
  if (!GRAPH_RUNTIME_ENABLED) return "";
  return await acquireGraphTokenWithNaa(scopes, options);
}

async function getCurrentMessageRestId(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const mailbox = OfficeAny?.context?.mailbox;
    const item = mailbox?.item;
    const itemId = String(item?.itemId || "").trim();
    if (!itemId) return "";
    try {
      if (typeof mailbox?.convertToRestId === "function" && OfficeAny?.MailboxEnums?.RestVersion?.v2_0) {
        return String(mailbox.convertToRestId(itemId, OfficeAny.MailboxEnums.RestVersion.v2_0) || "").trim() || itemId;
      }
    } catch (error) {
      clientLog.warn("[office] convertToRestId failed", error);
    }
    return itemId;
  } catch (error) {
    clientLog.warn("[office] getCurrentMessageRestId failed", error);
    return "";
  }
}

function mergeOutlookAttachments(primary: OutlookAttachment[], fallback: OutlookAttachment[]): OutlookAttachment[] {
  const byKey = new Map<string, OutlookAttachment>();
  const makeKey = (attachment: Partial<OutlookAttachment>) =>
    String(attachment?.id || attachment?.contentId || attachment?.name || "").trim().toLowerCase();
  for (const attachment of fallback) {
    const key = makeKey(attachment);
    if (!key) continue;
    byKey.set(key, { ...attachment });
  }
  for (const attachment of primary) {
    const key = makeKey(attachment);
    if (!key) continue;
    const existing = byKey.get(key);
    byKey.set(key, {
      ...(existing || {}),
      ...attachment,
      content: String(attachment.content || existing?.content || "").trim(),
    });
  }
  return Array.from(byKey.values()).filter((attachment) => String(attachment.name || "").trim());
}

async function getAttachmentsViaGraphForCurrentItem(): Promise<OutlookAttachment[]> {
  const messageId = await getCurrentMessageRestId();
  if (!messageId) return [];
  const token = await getGraphAccessToken(GRAPH_ATTACHMENT_SCOPES, { allowPrompts: true });
  if (!token) return [];
  const headers = { Authorization: `Bearer ${token}` };
  try {
    const listRes = await fetch(
      `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments?$select=id,name,contentType,size,isInline,contentId`,
      { headers }
    );
    if (!listRes.ok) {
      clientLog.warn("[office] Graph attachment list failed", { status: listRes.status });
      return [];
    }
    const listBody: any = await listRes.json();
    const rawItems = Array.isArray(listBody?.value) ? listBody.value : [];
    const results: OutlookAttachment[] = [];
    for (const summary of rawItems) {
      const attachmentId = String(summary?.id || "").trim();
      const attachmentName = String(summary?.name || "").trim();
      if (!attachmentId || !attachmentName) continue;
      try {
        const detailRes = await fetch(
          `https://graph.microsoft.com/v1.0/me/messages/${encodeURIComponent(messageId)}/attachments/${encodeURIComponent(attachmentId)}`,
          { headers }
        );
        if (!detailRes.ok) {
          clientLog.warn("[office] Graph attachment detail failed", { attachmentId, status: detailRes.status });
          continue;
        }
        const detail: any = await detailRes.json();
        results.push({
          id: attachmentId || undefined,
          name: attachmentName,
          contentType: String(detail?.contentType || summary?.contentType || "application/octet-stream").trim(),
          size: Number(detail?.size || summary?.size || 0) || undefined,
          isInline: Boolean(detail?.isInline ?? summary?.isInline),
          contentId: String(detail?.contentId || summary?.contentId || "").trim() || undefined,
          content: String(detail?.contentBytes || "").trim(),
        });
      } catch (error) {
        clientLog.warn("[office] Graph attachment detail exception", { attachmentId, error });
      }
    }
    return results.filter((attachment) => attachment.name);
  } catch (error) {
    clientLog.warn("[office] Graph attachment fallback failed", error);
    return [];
  }
}

function normalizeRecipients(arr: any): Recipient[] {
  if (!Array.isArray(arr)) return [];
  return arr
    .map((r) => ({
      name: String(r?.displayName || "").trim(),
      email: String(r?.emailAddress || "").trim(),
    }))
    .filter((r) => r.email);
}

export async function getOutlookContext(): Promise<OutlookMessageContext> {
  try {
    const OfficeAny = await ensureOfficeReady();

    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) {
      clientLog.warn("[office] mailbox.item is empty");
      return {};
    }

    const getAsyncValue = async (obj: any, coercer: (v: any) => string): Promise<string> => {
      if (!obj?.getAsync) return "";
      const p = new Promise<string>((resolve) => {
        try {
          obj.getAsync((r: any) => {
            try {
              if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(coercer(r.value));
              else resolve("");
            } catch { resolve(""); }
          });
        } catch { resolve(""); }
      });
      return await withTimeout(p, 2000, "");
    };

    const getSubject = async (): Promise<string> => {
      const s = item.subject;
      if (typeof s === "string") return s;
      // Compose: subject is an object with getAsync/setAsync
      return await getAsyncValue(s, (v) => String(v ?? ""));
    };

    const getRecipients = async (recips: any): Promise<Recipient[]> => {
      if (Array.isArray(recips)) return normalizeRecipients(recips);
      // Compose: recipients are an object with getAsync/addAsync
      if (recips?.getAsync) {
        const raw = await new Promise<any[]>((resolve) => {
          try {
            recips.getAsync((r: any) => {
              try {
                if (r?.status === OfficeAny.AsyncResultStatus.Succeeded && Array.isArray(r.value)) resolve(r.value);
                else resolve([]);
              } catch {
                resolve([]);
              }
            });
          } catch {
            resolve([]);
          }
        });
        return normalizeRecipients(raw);
      }
      return [];
    };

    const subject = await getSubject();

    // From is only reliable in Read. In Compose it may be missing/unsupported.
    const from = item.from;
    const fromEmail = from?.emailAddress ? String(from.emailAddress) : "";
    const fromName = from?.displayName ? String(from.displayName) : "";

    const conversationId = typeof item.conversationId === "string" ? item.conversationId : "";
    const internetMessageId = typeof item.internetMessageId === "string" ? item.internetMessageId : "";

    const itemId = typeof item.itemId === "string" ? item.itemId : "";

    const receivedDateTimeIso = item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : "";

    const toRecipients = await getRecipients(item.to);
    const ccRecipients = await getRecipients(item.cc);

    let isCompose = false;
    try {
      isCompose = await isComposeMode();
    } catch {
      isCompose = false;
    }

    return {
      subject,
      fromEmail,
      fromName,
      conversationId,
      internetMessageId,
      itemId,
      receivedDateTimeIso,
      toRecipients,
      ccRecipients,
      isCompose,
    };
  } catch (error) {
    clientLog.error("[office] getOutlookContext error", error);
    return {};
  }
}


// Backwards-compat with older UI code
export const getSelectedMessageContext = getOutlookContext;

const ODOO_LINKED_NOTICE = "iccc-odoo-linked";

function firstCategoryColor(colors: any, candidates: string[]): any {
  for (const candidate of candidates) {
    if (colors?.[candidate]) return colors[candidate];
  }
  return colors?.Preset0;
}

function hashCategorySeed(value: string): number {
  let hash = 0;
  for (const char of String(value || "")) {
    hash = ((hash << 5) - hash + char.charCodeAt(0)) | 0;
  }
  return Math.abs(hash);
}

function extractTicketSeriesKey(ticketCode: string): string {
  const raw = String(ticketCode || "").trim().replace(/^Ticket:\s*/i, "").replace(/^TK:\s*/i, "");
  return String(raw.split(/[-/_\s]/)[0] || "").trim().toUpperCase();
}

function isReservedManagedCategoryName(name: string): boolean {
  return isReservedOutlookCategoryName(name);
}

function resolveStatusCategoryColor(label: string, colors: any): any {
  const normalized = String(label || "").trim().toLowerCase();
  if (normalized === "em analise") return firstCategoryColor(colors, ["Preset3", "Preset1", "Preset0"]);
  if (normalized === "em progresso") return firstCategoryColor(colors, ["Preset1", "Preset5", "Preset0"]);
  if (normalized === "concluido") return firstCategoryColor(colors, ["Preset4", "Preset14", "Preset0"]);
  if (normalized === "aberto") return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  if (normalized === "fechado") return firstCategoryColor(colors, ["Preset4", "Preset14", "Preset0"]);
  return firstCategoryColor(colors, ["Preset12", "Preset0"]);
}

function resolveManagedCategoryColor(displayName: string, colors: any): any {
  const label = String(displayName || "").trim();
  if (!label || !colors) return colors?.Preset0;

  if (label === ODOO_LINKED_CATEGORY) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(GROUP_CATEGORY_PREFIX)) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(REFERENCE_CATEGORY_PREFIX)) {
    return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
  }

  if (label.startsWith(LEGACY_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(LEGACY_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(GROUP_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(GROUP_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(TICKET_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(TICKET_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(LABEL_STATUS_CATEGORY_PREFIX)) {
    return resolveStatusCategoryColor(label.slice(LABEL_STATUS_CATEGORY_PREFIX.length).trim(), colors);
  }

  if (label.startsWith(TICKET_CATEGORY_PREFIX)) {
    const seriesKey = extractTicketSeriesKey(label.slice(TICKET_CATEGORY_PREFIX.length));
    if (/^(RCL|REC|CLA|RET|RMA)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset6", "Preset9", "Preset0"]);
    }
    if (/^(EDD|ENC|ORD|PED|PO)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
    }
    if (/^(SUP|TCK|INC|SRV|TEC)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset5", "Preset8", "Preset0"]);
    }
    const palette = ["Preset22", "Preset5", "Preset8", "Preset4", "Preset1", "Preset6", "Preset9", "Preset14"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seriesKey) % palette.length], "Preset0"]);
  }

  if (label.startsWith(LEGACY_TICKET_CATEGORY_PREFIX)) {
    const seriesKey = extractTicketSeriesKey(label.slice(LEGACY_TICKET_CATEGORY_PREFIX.length));
    if (/^(RCL|REC|CLA|RET|RMA)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset6", "Preset9", "Preset0"]);
    }
    if (/^(EDD|ENC|ORD|PED|PO)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset22", "Preset7", "Preset0"]);
    }
    if (/^(SUP|TCK|INC|SRV|TEC)$/i.test(seriesKey)) {
      return firstCategoryColor(colors, ["Preset5", "Preset8", "Preset0"]);
    }
    const palette = ["Preset22", "Preset5", "Preset8", "Preset4", "Preset1", "Preset6", "Preset9", "Preset14"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seriesKey) % palette.length], "Preset0"]);
  }

  if (label.startsWith(LEGACY_LABEL_CATEGORY_PREFIX)) {
    const seed = label.slice(LEGACY_LABEL_CATEGORY_PREFIX.length).trim();
    const palette = ["Preset12", "Preset10", "Preset11", "Preset13", "Preset15", "Preset16", "Preset17", "Preset18"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(seed) % palette.length], "Preset0"]);
  }

  if (!isReservedManagedCategoryName(label)) {
    const palette = ["Preset12", "Preset10", "Preset11", "Preset13", "Preset15", "Preset16", "Preset17", "Preset18"];
    return firstCategoryColor(colors, [palette[hashCategorySeed(label) % palette.length], "Preset0"]);
  }

  return colors?.Preset0;
}


export type OutlookContactSuggestion = {
  name?: string;
  company?: string;
  jobTitle?: string;
  phones?: string[];
  email?: string;
};

export async function getOutlookContactSuggestionByEmail(emailRaw: string): Promise<OutlookContactSuggestion | null> {
  const email = String(emailRaw || "").trim().toLowerCase();
  if (!email) return null;
  if (!GRAPH_RUNTIME_ENABLED) return null;

  try {
    const token = await getGraphAccessToken(GRAPH_PEOPLE_SCOPES, { allowPrompts: false });
    if (!token) return null;

    const q = encodeURIComponent(`"${email}"`);
    const url = `https://graph.microsoft.com/v1.0/me/people?$search=${q}&$top=10`;
    const res = await fetch(url, {
      headers: {
        Authorization: `Bearer ${token}`,
        ConsistencyLevel: "eventual",
      },
    });
    if (!res.ok) return null;

    const body: any = await res.json();
    const arr = Array.isArray(body?.value) ? body.value : [];

    const exact = arr.find((p: any) => {
      const emails = Array.isArray(p?.scoredEmailAddresses) ? p.scoredEmailAddresses.map((x: any) => String(x?.address || "").trim().toLowerCase()) : [];
      return emails.includes(email);
    });

    if (!exact) return null;

    const phones = [
      ...(Array.isArray(exact?.businessPhones) ? exact.businessPhones : []),
      exact?.mobilePhone,
    ].map((x: any) => String(x || "").trim()).filter(Boolean);

    return {
      name: String(exact?.displayName || "").trim() || undefined,
      company: String(exact?.companyName || "").trim() || undefined,
      jobTitle: String(exact?.jobTitle || "").trim() || undefined,
      phones,
      email,
    };
  } catch {
    return null;
  }
}

// Ler corpo do email (texto simples) — usado pela IA
export async function getEmailBodyText(): Promise<string> {
  clientLog.log("[office] getEmailBodyText start");
  try {
    const OfficeAny = await ensureOfficeReady();
    const item: any = OfficeAny?.context?.mailbox?.item;
    if (!item?.body?.getAsync) return "";

    const p = new Promise<string>((resolve) => {
      item.body.getAsync("text", (r: any) => {
        try {
          if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(String(r.value ?? ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    });

    const result = await withTimeout(p, 3000, "");
    clientLog.log("[office] getEmailBodyText end");
    return result;
  } catch (e) {
    clientLog.error("[office] getEmailBodyText error", e);
    return "";
  }
}

export async function getEmailBodyHtml(): Promise<string> {
  clientLog.log("[office] getEmailBodyHtml start");
  try {
    const OfficeAny = await ensureOfficeReady();
    const item: any = OfficeAny?.context?.mailbox?.item;
    if (!item?.body?.getAsync) return "";

    const p = new Promise<string>((resolve) => {
      item.body.getAsync("html", (r: any) => {
        try {
          if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(String(r.value ?? ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    });

    const result = await withTimeout(p, 4000, "");
    clientLog.log("[office] getEmailBodyHtml end");
    return result;
  } catch (e) {
    clientLog.error("[office] getEmailBodyHtml error", e);
    return "";
  }
}


// Token barato para detetar mudanca de email (para polling fallback ao ItemChanged)
export async function getCurrentItemToken(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) return "";

    // "Context Poke": Read a basic property to force some hosts (Outlook Desktop) 
    // to refresh the internal state of the proxy object.
    void item.itemId;

    const cid = typeof item.conversationId === "string" ? item.conversationId : "";
    const imid = typeof item.internetMessageId === "string" ? item.internetMessageId : "";
    const itemId = typeof item.itemId === "string" ? item.itemId : "";
    const created = item.dateTimeCreated ? String(item.dateTimeCreated) : "";
    const subj = typeof item.subject === "string" ? item.subject : "";

    // Using a more structured token for better comparison
    return [cid, imid, itemId, created, subj].filter(Boolean).join("|");
  } catch {
    return "";
  }
}

async function ensureMasterCategory(displayName: string): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.masterCategories) return;
  const categoryColor = resolveManagedCategoryColor(displayName, OfficeAny.MailboxEnums?.CategoryColor);

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.masterCategories.getAsync((res: any) => {
        if (res.status !== OfficeAny.AsyncResultStatus.Succeeded) return resolve();
        const list = Array.isArray(res.value) ? res.value : [];
        const existing = list.find((c: any) => (c.displayName || c.name) === displayName);
        const addCategory = () =>
          OfficeAny.context.mailbox.masterCategories.addAsync([{ displayName, color: categoryColor }], (addResult: any) => {
            if (addResult?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              clientLog.warn("[office] masterCategories.addAsync failed", {
                displayName,
                error: addResult?.error?.message || addResult?.error?.code || "unknown",
              });
            }
            resolve();
          });

        if (existing) {
          if (!categoryColor || String(existing?.color || "") === String(categoryColor)) return resolve();
          if (typeof OfficeAny.context.mailbox.masterCategories.removeAsync !== "function") return resolve();
          return OfficeAny.context.mailbox.masterCategories.removeAsync([displayName], (removeResult: any) => {
            if (removeResult?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              clientLog.warn("[office] masterCategories.removeAsync failed", {
                displayName,
                error: removeResult?.error?.message || removeResult?.error?.code || "unknown",
              });
              return resolve();
            }
            addCategory();
          });
        }

        addCategory();
      });
    } catch {
      resolve();
    }
  });
}

async function addCategoryToCurrentItem(displayName: string): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.addAsync([displayName], () => resolve());
    } catch {
      resolve();
    }
  });
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
  let binary = "";
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    const slice = bytes.subarray(index, index + chunkSize);
    binary += String.fromCharCode(...slice);
  }
  return globalThis.btoa(binary);
}

function uint8ArraysToBase64(chunks: Uint8Array[]): string {
  const totalLength = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(totalLength);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return arrayBufferToBase64(merged.buffer);
}

async function getCurrentMessageAsEmlBase64(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item?.getAsFileAsync) return "";

    const file: any = await new Promise((resolve, reject) => {
      try {
        item.getAsFileAsync((result: any) => {
          if (result?.status === OfficeAny.AsyncResultStatus.Succeeded && result?.value) resolve(result.value);
          else reject(new Error(result?.error?.message || "getAsFileAsync failed"));
        });
      } catch (error) {
        reject(error);
      }
    });

    const sliceCount = Number(file?.sliceCount || 0);
    if (!sliceCount || !file?.getSliceAsync) return "";

    const chunks: Uint8Array[] = [];
    try {
      for (let index = 0; index < sliceCount; index += 1) {
        const sliceResult: any = await new Promise((resolve, reject) => {
          try {
            file.getSliceAsync(index, (result: any) => {
              if (result?.status === OfficeAny.AsyncResultStatus.Succeeded && result?.value) resolve(result.value);
              else reject(new Error(result?.error?.message || `getSliceAsync failed @${index}`));
            });
          } catch (error) {
            reject(error);
          }
        });

        const data = sliceResult?.data;
        if (Array.isArray(data)) {
          chunks.push(Uint8Array.from(data));
        } else if (data instanceof ArrayBuffer) {
          chunks.push(new Uint8Array(data));
        } else if (ArrayBuffer.isView(data)) {
          chunks.push(new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
        } else if (typeof data === "string" && data) {
          chunks.push(Uint8Array.from(Array.from(data).map((char) => char.charCodeAt(0) & 0xff)));
        }
      }
    } finally {
      try {
        if (typeof file?.closeAsync === "function") {
          file.closeAsync(() => {});
        }
      } catch {
        // noop
      }
    }

    if (!chunks.length) return "";
    return uint8ArraysToBase64(chunks);
  } catch (error) {
    clientLog.warn("[office] getCurrentMessageAsEmlBase64 failed", error);
    return "";
  }
}

async function getAttachmentsViaEmlForCurrentItem(): Promise<OutlookAttachment[]> {
  try {
    const emlBase64 = await getCurrentMessageAsEmlBase64();
    if (!emlBase64) return [];

    const response = await fetch("/api/links/eml/extract", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ emlBase64 }),
    });
    if (!response.ok) {
      clientLog.warn("[office] EML attachment extract failed", { status: response.status });
      return [];
    }
    const payload: any = await response.json();
    const rawAttachments = Array.isArray(payload?.attachments) ? payload.attachments : [];
    return rawAttachments
      .map((attachment: any) => ({
        id: String(attachment?.id || "").trim() || undefined,
        name: String(attachment?.name || "").trim(),
        contentType: String(attachment?.contentType || "application/octet-stream").trim(),
        size: Number(attachment?.size || 0) || undefined,
        isInline: Boolean(attachment?.isInline),
        contentId: String(attachment?.contentId || "").trim() || undefined,
        content: String(attachment?.content || "").trim(),
      }))
      .filter((attachment: OutlookAttachment) => attachment.name && attachment.content);
  } catch (error) {
    clientLog.warn("[office] EML attachment fallback failed", error);
    return [];
  }
}

async function addCategoriesToCurrentItem(displayNames: string[]): Promise<void> {
  const uniqueNames = Array.from(new Set((displayNames || []).map((name) => String(name || "").trim()).filter(Boolean)));
  if (!uniqueNames.length) return;
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.addAsync(uniqueNames, (result: any) => {
        if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.warn("[office] item.categories.addAsync failed", {
            categories: uniqueNames,
            error: result?.error?.message || result?.error?.code || "unknown",
          });
        }
        resolve();
      });
    } catch {
      resolve();
    }
  });
}

async function removeCategoryFromCurrentItem(displayName: string): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.removeAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.removeAsync([displayName], () => resolve());
    } catch {
      resolve();
    }
  });
}

async function removeCategoriesFromCurrentItem(displayNames: string[]): Promise<void> {
  const uniqueNames = Array.from(new Set((displayNames || []).map((name) => String(name || "").trim()).filter(Boolean)));
  if (!uniqueNames.length) return;
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.removeAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.removeAsync(uniqueNames, (result: any) => {
        if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.warn("[office] item.categories.removeAsync failed", {
            categories: uniqueNames,
            error: result?.error?.message || result?.error?.code || "unknown",
          });
        }
        resolve();
      });
    } catch {
      resolve();
    }
  });
}

async function getCurrentItemCategoryNames(): Promise<string[]> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  const categoriesApi = OfficeAny?.context?.mailbox?.item?.categories;
  if (!categoriesApi) return [];

  if (Array.isArray(categoriesApi)) {
    return categoriesApi
      .map((entry: any) => String(entry?.displayName || entry?.name || entry || "").trim())
      .filter(Boolean);
  }

  if (typeof categoriesApi.getAsync === "function") {
    return await new Promise<string[]>((resolve) => {
      try {
        categoriesApi.getAsync((result: any) => {
          if (result?.status !== OfficeAny.AsyncResultStatus.Succeeded) return resolve([]);
          const value = Array.isArray(result.value) ? result.value : [];
          resolve(
            value
              .map((entry: any) => String(entry?.displayName || entry?.name || entry || "").trim())
              .filter(Boolean)
          );
        });
      } catch {
        resolve([]);
      }
    });
  }

  return [];
}

async function hasExpectedCurrentItemToken(expectedItemToken?: string): Promise<boolean> {
  if (!expectedItemToken) return true;
  const currentToken = await getCurrentItemToken().catch(() => "");
  if (!currentToken) {
    clientLog.warn("[office] current item token unavailable during category sync; continuing without strict guard", {
      expectedItemToken,
    });
    return true;
  }
  return currentToken === expectedItemToken;
}

function getCurrentManagedCategoryNames(currentCategories: string[], plan: OutlookCategoryPlan): string[] {
  const managedLabelSet = new Set(plan.managedLabelNames.map((label) => label.toLowerCase()));
  const managedSpecialSet = new Set(plan.managedSpecialCategories.map((label) => label.toLowerCase()));
  return currentCategories.filter((name) => {
    const normalized = String(name || "").trim().toLowerCase();
    return Boolean(
      (plan.manageClassificationFamilies && isManagedCategoryFamilyName(name))
      || managedLabelSet.has(normalized)
      || managedSpecialSet.has(normalized)
    );
  });
}

export async function applyOutlookCategoryPlan(
  plan: OutlookCategoryPlan,
  options?: { expectedItemToken?: string }
): Promise<boolean> {
  const expectedItemToken = String(options?.expectedItemToken || "").trim();
  if (!(await hasExpectedCurrentItemToken(expectedItemToken))) return false;

  const currentCategories = await getCurrentItemCategoryNames();
  const desiredCategories = normalizeUniqueCategoryValues(plan.desiredCategories);
  const currentManagedCategories = getCurrentManagedCategoryNames(currentCategories, plan);
  const desiredCategorySet = new Set(desiredCategories);
  const toAdd = desiredCategories.filter((name) => !currentCategories.includes(name));
  const toRemove = Array.from(new Set(currentManagedCategories.filter((name) => !desiredCategorySet.has(name))));

  for (const categoryName of desiredCategories) {
    await ensureMasterCategory(categoryName);
  }

  if (!(await hasExpectedCurrentItemToken(expectedItemToken))) {
    clientLog.warn("[office] applyOutlookCategoryPlan aborted after category preparation because the item changed", {
      expectedItemToken,
    });
    return false;
  }

  await addCategoriesToCurrentItem(toAdd);

  if (!(await hasExpectedCurrentItemToken(expectedItemToken))) {
    clientLog.warn("[office] applyOutlookCategoryPlan skipped removals because the item changed", {
      expectedItemToken,
    });
    return false;
  }

  await removeCategoriesFromCurrentItem(toRemove);
  return true;
}

export async function syncOutlookCategorySource(
  source: Partial<OutlookCategorySource> | null | undefined,
  options?: { expectedItemToken?: string; manageClassificationFamilies?: boolean }
): Promise<boolean> {
  return await applyOutlookCategoryPlan(
    buildOutlookCategoryPlan(source, { manageClassificationFamilies: options?.manageClassificationFamilies }),
    { expectedItemToken: options?.expectedItemToken }
  );
}

export async function syncCurrentItemOutlookCategoriesFromContext(
  options?: { expectedItemToken?: string }
): Promise<boolean> {
  const currentContext = await getSelectedMessageContext().catch(() => ({} as OutlookMessageContext));
  const hasCurrentIdentity = Boolean(
    String(currentContext.itemId || "").trim()
    || String(currentContext.internetMessageId || "").trim()
    || String(currentContext.conversationId || "").trim()
  );
  if (!hasCurrentIdentity) return false;

  const expectedItemToken = String(options?.expectedItemToken || "").trim() || await getCurrentItemToken().catch(() => "");
  const payload = {
    itemId: String(currentContext.itemId || "").trim() || undefined,
    internetMessageId: String(currentContext.internetMessageId || "").trim() || undefined,
    conversationId: String(currentContext.conversationId || "").trim() || undefined,
    subject: String(currentContext.subject || "").trim() || undefined,
    fromEmail: String(currentContext.fromEmail || "").trim() || undefined,
    fromName: String(currentContext.fromName || "").trim() || undefined,
    receivedAtIso: String(currentContext.receivedDateTimeIso || "").trim() || undefined,
    messageDateIso: String(currentContext.receivedDateTimeIso || "").trim() || undefined,
  };
  const [settings, related, links] = await Promise.all([
    getSettings().catch(() => null),
    getRelatedEmailContext(payload).catch(() => null),
    getLinks(payload.conversationId, payload.internetMessageId, payload.itemId).catch(() => []),
  ]);
  const knownLabelNames = collectKnownOutlookCategoryLabelNames({
    settings,
    email: related?.email || null,
    groups: Array.isArray(related?.groups) ? related.groups : [],
    tickets: Array.isArray(related?.tickets) ? related.tickets : [],
  });
  const snapshot = await getManagedOutlookCategorySnapshot(knownLabelNames).catch(() => null);
  const applied = await syncOutlookCategorySource(buildOutlookCategorySourceFromRelatedContext({
    email: related?.email || null,
    groups: Array.isArray(related?.groups) ? related.groups : [],
    tickets: Array.isArray(related?.tickets) ? related.tickets : [],
    settings,
    currentOutlookLabelNames: snapshot?.labelNames || [],
    specialCategories: Array.isArray(links) && links.length ? [ODOO_LINKED_CATEGORY] : [],
    managedSpecialCategories: [ODOO_LINKED_CATEGORY],
  }), {
    expectedItemToken,
  });
  if (applied) dispatchOutlookCategoryContextInvalidated();
  return applied;
}

export async function syncOdooLinkedCategory(hasLinks: boolean): Promise<void> {
  await syncOutlookCategorySource(
    {
      specialCategories: hasLinks ? [ODOO_LINKED_CATEGORY] : [],
      managedSpecialCategories: [ODOO_LINKED_CATEGORY],
    },
    { manageClassificationFamilies: false }
  );
}

export async function syncManagedOutlookCategories(input: LegacyManagedOutlookCategoryInput): Promise<void> {
  await syncOutlookCategorySource(buildOutlookCategorySourceFromLegacyInput(input));
}

export async function getManagedOutlookCategorySnapshot(): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}>;
export async function getManagedOutlookCategorySnapshot(knownLabelNames: string[]): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}>;
export async function getManagedOutlookCategorySnapshot(knownLabelNames?: string[]): Promise<{
  principalGroupNames: string[];
  referenceGroupNames: string[];
  groupNames: string[];
  ticketCodes: string[];
  statuses: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  labelNames: string[];
}> {
  const currentCategories = await getCurrentItemCategoryNames();
  const knownLabelSet = new Set(
    normalizeUniqueCategoryValues(knownLabelNames).map((label) => String(label || "").trim().toLowerCase())
  );
  const principalGroupNames = currentCategories
    .filter((name) => name.startsWith(GROUP_CATEGORY_PREFIX))
    .map((name) => name.slice(GROUP_CATEGORY_PREFIX.length).trim())
    .filter(Boolean);
  const referenceGroupNames = currentCategories
    .filter((name) => name.startsWith(REFERENCE_CATEGORY_PREFIX))
    .map((name) => name.slice(REFERENCE_CATEGORY_PREFIX.length).trim())
    .filter(Boolean);
  return {
    principalGroupNames,
    referenceGroupNames,
    groupNames: [...principalGroupNames, ...referenceGroupNames],
    ticketCodes: currentCategories
      .filter((name) => name.startsWith(TICKET_CATEGORY_PREFIX) || name.startsWith(LEGACY_TICKET_CATEGORY_PREFIX))
      .map((name) => name.startsWith(TICKET_CATEGORY_PREFIX) ? name.slice(TICKET_CATEGORY_PREFIX.length).trim() : name.slice(LEGACY_TICKET_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    statuses: currentCategories
      .filter((name) => name.startsWith(LEGACY_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(LEGACY_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    groupStatuses: currentCategories
      .filter((name) => name.startsWith(GROUP_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(GROUP_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    ticketStatuses: currentCategories
      .filter((name) => name.startsWith(TICKET_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(TICKET_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    labelStatuses: currentCategories
      .filter((name) => name.startsWith(LABEL_STATUS_CATEGORY_PREFIX))
      .map((name) => name.slice(LABEL_STATUS_CATEGORY_PREFIX.length).trim())
      .filter(Boolean),
    labelNames: currentCategories
      .filter((name) =>
        name.startsWith(LEGACY_LABEL_CATEGORY_PREFIX)
        || (!isReservedManagedCategoryName(name) && knownLabelSet.has(String(name || "").trim().toLowerCase()))
      )
      .map((name) => name.startsWith(LEGACY_LABEL_CATEGORY_PREFIX) ? name.slice(LEGACY_LABEL_CATEGORY_PREFIX.length).trim() : String(name || "").trim())
      .filter(Boolean),
  };
}

export async function syncManualGroupCategories(groupNames: string[]): Promise<void> {
  await syncManagedOutlookCategories({ principalGroupNames: groupNames });
}

export async function syncLinkCategoriesToComposeDraft(
  input: (Partial<OutlookCategorySource> | (LegacyManagedOutlookCategoryInput & { hasOdooLinks?: boolean })) | null | undefined,
  options?: { attempts?: number; delayMs?: number }
): Promise<void> {
  const attempts = Math.max(1, Number(options?.attempts || 12));
  const delayMs = Math.max(150, Number(options?.delayMs || 450));
  const source = "specialCategories" in (input || {}) || "managedSpecialCategories" in (input || {})
    ? buildOutlookCategoryPlan(input as Partial<OutlookCategorySource>).source
    : buildOutlookCategorySourceFromLegacyInput({
        ...(input as LegacyManagedOutlookCategoryInput & { hasOdooLinks?: boolean }),
        specialCategories: (input as { hasOdooLinks?: boolean } | null | undefined)?.hasOdooLinks ? [ODOO_LINKED_CATEGORY] : [],
        managedSpecialCategories: typeof (input as { hasOdooLinks?: boolean } | null | undefined)?.hasOdooLinks === "boolean" ? [ODOO_LINKED_CATEGORY] : [],
      });
  const hasManagedCategories = Boolean(
    source.principalGroupNames.length
    || source.referenceGroupNames.length
    || source.ticketCodes.length
    || source.groupStatuses.length
    || source.ticketStatuses.length
    || source.labelStatuses.length
    || source.labelNames.length
    || source.managedLabelNames.length
    || source.specialCategories.length
    || source.managedSpecialCategories.length
  );
  const hasOdooLinks = input?.hasOdooLinks === true;

  if (!hasManagedCategories && !hasOdooLinks) return;

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (attempt > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }

    const composeReady = await isComposeMode().catch(() => false);
    if (!composeReady) continue;

    await syncOutlookCategorySource(source).catch(() => {
      // best-effort
    });
    return;
  }

  clientLog.warn("[office] syncLinkCategoriesToComposeDraft: compose draft not ready in time");
}

export async function setSubjectInComposeDraft(subject: string, options?: { attempts?: number; delayMs?: number }): Promise<void> {
  const desiredSubject = String(subject || "").trim();
  if (!desiredSubject) return;

  const attempts = Math.max(1, Number(options?.attempts || 12));
  const delayMs = Math.max(150, Number(options?.delayMs || 450));

  for (let attempt = 0; attempt < attempts; attempt += 1) {
    if (attempt > 0) {
      await new Promise((resolve) => setTimeout(resolve, delayMs));
    }

    const composeReady = await isComposeMode().catch(() => false);
    if (!composeReady) continue;

    const OfficeAny = await ensureOfficeReady().catch(() => null);
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item?.subject?.setAsync) continue;

    await new Promise<void>((resolve, reject) => {
      item.subject.setAsync(desiredSubject, (result: any) => {
        if (result?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result?.error?.message || "Erro ao definir assunto"));
      });
    });
    return;
  }

  clientLog.warn("[office] setSubjectInComposeDraft: compose draft subject not ready in time");
}

export async function syncOdooLinkedNotification(hasLinks: boolean, count = 0): Promise<void> {
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  const notifications = OfficeAny?.context?.mailbox?.item?.notificationMessages;
  if (!notifications?.replaceAsync || !notifications?.removeAsync) return;

  await new Promise<void>((resolve) => {
    try {
      if (!hasLinks) {
        notifications.removeAsync(ODOO_LINKED_NOTICE, () => resolve());
        return;
      }

      notifications.replaceAsync(
        ODOO_LINKED_NOTICE,
        {
          type: OfficeAny.MailboxEnums?.ItemNotificationMessageType?.InformationalMessage,
          message: count > 0
            ? `Este email tem ${count} ligacao(oes) ativa(s) ao Odoo.`
            : "Este email tem ligacoes ativas ao Odoo.",
          icon: "Icon.80x80",
          persistent: false,
        },
        () => resolve()
      );
    } catch {
      resolve();
    }
  });
}

let activeDialog: any = null;

export type AiReplyTargetSelection = {
  emailKey: string;
  itemId?: string;
  emailWebLink?: string;
  internetMessageId?: string;
  conversationId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  messageDateIso?: string;
  receivedAtIso?: string;
  bodyText?: string;
  bodyHtml?: string;
};

type CockpitHostAction =
  | { type: "close" }
  | { type: "open-email"; itemId?: string; emailWebLink?: string }
  | { type: "reply-current" }
  | { type: "forward-current" }
  | { type: "sync-current-item-categories" }
  | {
      type: "sync-managed-categories";
      payload: (LegacyManagedOutlookCategoryInput & Partial<OutlookCategorySource>);
    };

async function executeCockpitHostAction(action: CockpitHostAction): Promise<void> {
  if (action.type === "close") {
    try {
      if (activeDialog) activeDialog.close();
    } catch { }
    activeDialog = null;
    return;
  }

  if (action.type === "open-email") {
    await openLinkedOutlookEmail({ itemId: action.itemId, emailWebLink: action.emailWebLink });
    return;
  }

  if (action.type === "reply-current") {
    await displayReplyForm("", true);
    return;
  }

  if (action.type === "forward-current") {
    await displayForwardForm("", true);
    return;
  }

  if (action.type === "sync-current-item-categories") {
    await syncCurrentItemOutlookCategoriesFromContext();
    return;
  }

  if (action.type === "sync-managed-categories") {
    if ("groupNames" in (action.payload || {}) || "statuses" in (action.payload || {})) {
      await syncManagedOutlookCategories(action.payload || {});
      dispatchOutlookCategoryContextInvalidated();
      return;
    }
    await syncOutlookCategorySource(action.payload || {});
    dispatchOutlookCategoryContextInvalidated();
    return;
  }
}

function tryParseCockpitHostMessage(rawMessage: any): CockpitHostAction | null {
  const text = String(rawMessage || "").trim();
  if (!text) return null;
  if (text === "close") return { type: "close" };
  try {
    const parsed = JSON.parse(text);
    if (parsed?.type !== "host-action" || !parsed?.action?.type) return null;
    return parsed.action as CockpitHostAction;
  } catch {
    return null;
  }
}

function buildCockpitViewUrl(view: string, params: Record<string, string>) {
  const url = new URL(window.location.origin);
  url.searchParams.set("view", view);
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));
  return url;
}

function tryOpenStandaloneWindow(url: URL, name: string, features: string): boolean {
  try {
    const popup = window.open(url.toString(), name, features);
    if (popup) {
      try {
        popup.focus();
      } catch {
        // best effort
      }
      return true;
    }
  } catch (error) {
    clientLog.warn("[office] standalone window fallback failed", error);
  }
  return false;
}

/**
 * Opens a separate window using Office Dialog API.
 * Guard: only one dialog at a time (evita "já existe uma dialog ativa").
 */
async function openCockpitView<T = void>(view: string, params: Record<string, string>, options?: { height?: number; width?: number; displayInIframe?: boolean; timeoutMs?: number }) {
  const OfficeAny = await ensureOfficeReady();
  const url = buildCockpitViewUrl(view, params);

  clientLog.log(`[office] openDialog ${url.toString()}`);

  // close previous if any
  try {
    if (activeDialog) activeDialog.close();
  } catch { }
  activeDialog = null;

  return await new Promise<T | null>((resolve, reject) => {
    let settled = false;
    const resolveOnce = (value: T | null = null) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };
    const rejectOnce = (error: Error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };
    const timer = setTimeout(() => {
      clientLog.warn(`[office] displayDialogAsync timeout for ${view}`);
      rejectOnce(new Error("A abertura da janela demorou demasiado tempo."));
    }, Math.max(2000, Number(options?.timeoutMs || 4000)));

    OfficeAny.context.ui.displayDialogAsync(
      url.toString(),
      { height: options?.height || 65, width: options?.width || 40, displayInIframe: Boolean(options?.displayInIframe) },
      (result: any) => {
        clearTimeout(timer);
        if (result.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.error(`[office] displayDialogAsync failed: ${result.error?.message || "unknown"}`);
          rejectOnce(new Error(result.error?.message || "Falha ao abrir janela (Dialog)."));
          return;
        }
        const dialog = result.value;
        activeDialog = dialog;

        dialog.addEventHandler(OfficeAny.EventType.DialogMessageReceived, (arg: any) => {
          const action = tryParseCockpitHostMessage(arg?.message);
          if (action) {
            void executeCockpitHostAction(action)
              .catch((error) => clientLog.error("[office] host action failed", error))
              .finally(() => {
                if (action.type === "close") {
                  resolveOnce();
                }
              });
            return;
          }

          const resultPayload = tryParseCockpitDialogResult<T>(arg?.message);
          if (typeof resultPayload !== "undefined") {
            try {
              if (activeDialog) activeDialog.close();
            } catch { }
            activeDialog = null;
            resolveOnce(resultPayload);
          }
        });

        dialog.addEventHandler(OfficeAny.EventType.DialogEventReceived, () => {
          activeDialog = null;
          resolveOnce();
        });
      }
    );
  });
}

export async function openCockpitDialog(params: Record<string, string>) {
  return await openCockpitView("dialog", params, { height: 65, width: 40, displayInIframe: false });
}

export async function openGroupExplorer(params: Record<string, string>) {
  try {
    return await openCockpitView("group-explorer", params, { height: 78, width: 52, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-explorer", params);
    clientLog.warn("[office] group explorer fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

function tryParseCockpitDialogResult<T>(rawMessage: any): T | undefined {
  const text = String(rawMessage || "").trim();
  if (!text) return undefined;
  try {
    const parsed = JSON.parse(text);
    if (parsed?.type !== "dialog-result") return undefined;
    return parsed.result as T;
  } catch {
    return undefined;
  }
}

export async function openGroupManager(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("group-manager", params, { height: 82, width: 58, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-manager", params);
    clientLog.warn("[office] group manager fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openAiSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("ai-settings", params, { height: 84, width: 60, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("ai-settings", params);
    clientLog.warn("[office] ai settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openGroupSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("group-settings", params, { height: 84, width: 60, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-settings", params);
    clientLog.warn("[office] group settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openGroupClassificationStudio(params: Record<string, string> = {}) {
  const url = buildCockpitViewUrl("group-classification-studio", params);
  try {
    return await openCockpitView("group-classification-studio", params, {
      height: 84,
      width: 74,
      displayInIframe: false,
      timeoutMs: 16000,
    });
  } catch (error) {
    clientLog.warn("[office] group classification studio dialog failed; retrying standalone dialog", error);
  }

  await sleep(450);

  try {
    return await openCockpitView("group-classification-studio", params, {
      height: 84,
      width: 74,
      displayInIframe: false,
      timeoutMs: 16000,
    });
  } catch (error) {
    clientLog.warn("[office] group classification studio retry failed; trying popup fallback", error);
  }

  const popupOpened = tryOpenStandaloneWindow(
    url,
    "iccc-group-classification-studio",
    "popup=yes,width=1520,height=980,resizable=yes,scrollbars=yes"
  );

  if (popupOpened) return null;

  throw new Error("Nao foi possivel abrir a janela externa do Classificar.");
}

export async function openAppSettings(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("app-settings", params, { height: 86, width: 64, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("app-settings", params);
    clientLog.warn("[office] app settings fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
}

export async function openAiReplyTargetPicker(params: Record<string, string> = {}) {
  return await openCockpitView<AiReplyTargetSelection>("ai-reply-target-picker", params, { height: 84, width: 74, displayInIframe: true });
}

export async function requestCockpitHostAction(action: CockpitHostAction): Promise<boolean> {
  try {
    const OfficeAny = await ensureOfficeReady();
    if (typeof OfficeAny?.context?.ui?.messageParent === "function") {
      OfficeAny.context.ui.messageParent(JSON.stringify({ type: "host-action", action }));
      return true;
    }
  } catch {
    // fall through to local execution
  }

  try {
    await executeCockpitHostAction(action);
    return true;
  } catch {
    return false;
  }
}

/**
 * Subscribe to selection change (when user clicks a different email).
 * IMPORTANT: This must NEVER open dialogs. Only refresh the taskpane state.
 */
export async function subscribeToItemChanges(onChanged: () => void): Promise<() => void> {
  const OfficeAny = await ensureOfficeReady();

  const handler = () => {
    try {
      onChanged();
    } catch (e) {
      clientLog.error("[office] ItemChanged handler error", e);
    }
  };

  try {
    if (OfficeAny?.context?.mailbox?.addHandlerAsync) {
      OfficeAny.context.mailbox.addHandlerAsync(OfficeAny.EventType.ItemChanged, handler);
      clientLog.log("[office] subscribed ItemChanged");
      return () => {
        try {
          OfficeAny.context.mailbox.removeHandlerAsync(OfficeAny.EventType.ItemChanged, { handler });
          clientLog.log("[office] unsubscribed ItemChanged");
        } catch { }
      };
    }
  } catch (e) {
    clientLog.warn("[office] ItemChanged not supported here", e);
  }

  return () => { };
}

/**
 * Checks if the current item is in compose mode (editable).
 */
export async function isComposeMode(): Promise<boolean> {
  const OfficeAny = await ensureOfficeReady();
  return Boolean(OfficeAny?.context?.mailbox?.item?.body?.setSelectedDataAsync);
}

/**
 * Inserts HTML or Text into the message body at the current cursor position.
 * Only works in Compose mode.
 */
export async function insertTextToBody(content: string, isHtml = true): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  clientLog.log(`[office] insertTextToBody called. Item exists: ${!!item}, setSelectedDataAsync exists: ${!!item?.body?.setSelectedDataAsync}`);

  if (!item?.body?.setSelectedDataAsync) {
    clientLog.warn("[office] insertTextToBody: Not in compose mode or not supported (setSelectedDataAsync missing).");
    throw new Error("Não é possível inserir texto: o item não está em modo de edição ou a funcionalidade não é suportada.");
  }

  // Formatting fix: Convert newlines to <br> if HTML
  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;

  return await new Promise<void>((resolve, reject) => {
    item.body.setSelectedDataAsync(
      finalContent,
      { coercionType: isHtml ? OfficeAny.CoercionType.Html : OfficeAny.CoercionType.Text },
      (result: any) => {
        if (result.status === OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.log("[office] insertTextToBody: Success");
          resolve();
        } else {
          clientLog.error(`[office] setSelectedDataAsync failed: ${result.error?.message || "unknown"}`);
          reject(new Error(result.error?.message || "Falha ao inserir no email."));
        }
      }
    );
  });
}

/**
 * Opens a new Reply form with pre-filled content.
 */
export async function displayReplyForm(content: string, isHtml = true, options?: { replyAll?: boolean }): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  if (!item?.displayReplyAllForm && !item?.displayReplyForm) {
    throw new Error("Funcionalidade de resposta não disponível neste item.");
  }

  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;
  const replyAll = options?.replyAll !== false;

  if (replyAll && typeof item.displayReplyAllForm === "function") {
    if (isHtml) item.displayReplyAllForm({ htmlBody: finalContent });
    else item.displayReplyAllForm(finalContent);
    return;
  }

  if (typeof item.displayReplyForm === "function") {
    if (isHtml) item.displayReplyForm({ htmlBody: finalContent });
    else item.displayReplyForm(finalContent);
    return;
  }

  if (isHtml) item.displayReplyAllForm({ htmlBody: finalContent });
  else item.displayReplyAllForm(finalContent);
}

/**
 * Opens a new Forward form with pre-filled content.
 */
export async function displayForwardForm(content: string, isHtml = true): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  if (!item?.displayForwardForm) {
    throw new Error("Funcionalidade de reenvio não disponível neste item.");
  }

  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;

  item.displayForwardForm({ htmlBody: isHtml ? finalContent : undefined, textBody: !isHtml ? finalContent : undefined });
}

/**
 * Opens a brand new email form with pre-filled recipients, subject and body.
 */
export async function displayNewMessageForm(params: {
  toRecipients?: string[];
  ccRecipients?: string[];
  bccRecipients?: string[];
  subject?: string;
  body?: string;
  isHtml?: boolean;
}): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const mailbox = OfficeAny?.context?.mailbox;
  if (!mailbox?.displayNewMessageForm) {
    throw new Error("Funcionalidade de nova mensagem nÃ£o disponÃ­vel neste ambiente.");
  }

  const isHtml = params.isHtml !== false;
  mailbox.displayNewMessageForm({
    toRecipients: Array.isArray(params.toRecipients) ? params.toRecipients : [],
    ccRecipients: Array.isArray(params.ccRecipients) ? params.ccRecipients : [],
    bccRecipients: Array.isArray(params.bccRecipients) ? params.bccRecipients : [],
    subject: params.subject,
    htmlBody: isHtml ? String(params.body || "") : undefined,
    body: !isHtml ? String(params.body || "") : undefined,
  });
}

/**
 * Opens the New Appointment form with pre-filled details.
 */
export async function displayNewMeetingForm(params: {
  subject?: string;
  body?: string;
  location?: string;
  start?: Date;
  end?: Date;
  requiredAttendees?: string[];
}) {
  const OfficeAny = await ensureOfficeReady();
  const mailbox = OfficeAny?.context?.mailbox;

  if (!mailbox?.displayNewAppointmentForm) {
    throw new Error("Calendário não suportado neste ambiente.");
  }

  mailbox.displayNewAppointmentForm({
    subject: params.subject,
    body: params.body,
    location: params.location,
    start: params.start,
    end: params.end,
    requiredAttendees: params.requiredAttendees,
  });
}

/**
 * Sets recipients in a compose item.
 * @param type 'to' | 'cc' | 'bcc'
 * @param recipients Array of email strings or Recipient objects
 */
export async function setRecipients(type: 'to' | 'cc' | 'bcc', recipients: (string | Recipient)[]): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  const target = item?.[type];

  if (!target?.setAsync) {
    clientLog.warn(`[office] setRecipients: target ${type} does not support setAsync`);
    return;
  }

  const formatted = recipients.map(r => typeof r === 'string' ? r : r.email);

  return await new Promise<void>((resolve, reject) => {
    target.setAsync(formatted, (result: any) => {
      if (result.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || `Erro ao definir destinatários ${type}`));
    });
  });
}

/**
 * Sets the subject of a compose item.
 */
export async function setSubject(subject: string): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  if (!item?.subject?.setAsync) {
    clientLog.warn("[office] setSubject: item.subject does not support setAsync");
    return;
  }

  return await new Promise<void>((resolve, reject) => {
    item.subject.setAsync(subject, (result: any) => {
      if (result.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Erro ao definir assunto"));
    });
  });
}

export async function addBase64AttachmentToCompose(name: string, contentBase64: string): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;

  if (!item?.addFileAttachmentFromBase64Async) {
    throw new Error("Abre uma mensagem em modo de edição para anexar documentos.");
  }

  const safeName = String(name || "").trim() || "documento";
  const base64 = String(contentBase64 || "").trim().replace(/^data:[^,]+,/, "");
  if (!base64) {
    throw new Error("O documento não tem conteúdo disponível para anexar.");
  }

  await new Promise<void>((resolve, reject) => {
    try {
      item.addFileAttachmentFromBase64Async(base64, safeName, { isInline: false }, (result: any) => {
        if (result?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve();
        else reject(new Error(result?.error?.message || "Falha ao anexar o documento."));
      });
    } catch (error: any) {
      reject(error);
    }
  });
}

/**
 * Fetch attachments from the current item.
 * Returns array of attachment metadata plus content when available.
 */
export async function getAttachments(): Promise<OutlookAttachment[]> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;

    if (!item?.attachments) return [];

    const attachments = item.attachments;
    const fileAttachments = Array.from(attachments).filter((att: any) => att?.attachmentType === "file");
    const results: OutlookAttachment[] = [];

    for (const att of fileAttachments) {
      // Only process file attachments
      try {
        const content = await new Promise<string>((resolve, reject) => {
          item.getAttachmentContentAsync(att.id, async (result: any) => {
            if (result.status !== OfficeAny.AsyncResultStatus.Succeeded) {
              reject(new Error(result.error?.message));
              return;
            }

            const format = String(result?.value?.format || "").trim().toLowerCase();
            const rawContent = String(result?.value?.content || "").trim();
            if (!rawContent) {
              resolve("");
              return;
            }

            if (!format || format === "base64") {
              resolve(rawContent);
              return;
            }

            if (format === "url") {
              try {
                const response = await fetch(rawContent);
                if (!response.ok) {
                  reject(new Error(`Falha ao descarregar conteudo do anexo (${response.status})`));
                  return;
                }
                const buffer = await response.arrayBuffer();
                resolve(arrayBufferToBase64(buffer));
                return;
              } catch (error: any) {
                reject(error);
                return;
              }
            }

            resolve(rawContent);
          });
        });

        results.push({
          id: String(att.id || "").trim() || undefined,
          name: String(att.name || "").trim(),
          contentType: String(att.contentType || "").trim(),
          size: Number(att.size || 0) || undefined,
          isInline: Boolean((att as any).isInline),
          contentId: String((att as any).contentId || "").trim() || undefined,
          content: content,
        });
      } catch (e) {
        clientLog.error(`[office] Failed to download attachment ${att.name}`, e);
      }
    }

    let mergedResults = results;

    const needsEmlFallback =
      fileAttachments.length > 0 &&
      (mergedResults.length < fileAttachments.length || !mergedResults.some((attachment) => String(attachment.content || "").trim()));
    if (needsEmlFallback) {
      const emlResults = await getAttachmentsViaEmlForCurrentItem();
      if (emlResults.length) {
        mergedResults = mergeOutlookAttachments(mergedResults, emlResults);
      }
    }

    const needsGraphFallback =
      fileAttachments.length > 0 &&
      (mergedResults.length < fileAttachments.length || !mergedResults.some((attachment) => String(attachment.content || "").trim()));
    if (GRAPH_RUNTIME_ENABLED && needsGraphFallback) {
      const graphResults = await getAttachmentsViaGraphForCurrentItem();
      if (graphResults.length) {
        return mergeOutlookAttachments(mergedResults, graphResults);
      }
    }

    return mergedResults;
  } catch (error) {
    clientLog.error("[office] getAttachments error", error);
    return [];
  }
}

export async function openLinkedOutlookEmail(target: { itemId?: string; emailWebLink?: string }): Promise<boolean> {
  const itemId = String(target?.itemId || "").trim();
  if (itemId) {
    const OfficeAny: any = await ensureOfficeReady().catch(() => null);
    const mailbox = OfficeAny?.context?.mailbox;
    if (typeof mailbox?.displayMessageForm === "function") {
      mailbox.displayMessageForm(itemId);
      return true;
    }
  }

  const emailWebLink = String(target?.emailWebLink || "").trim();
  if (emailWebLink) {
    window.open(emailWebLink, "_blank", "noopener,noreferrer");
    return true;
  }

  return false;
}
