import { clientLog } from "./logger";



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

function sleep(ms: number) {
  return new Promise((r) => setTimeout(r, ms));
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

const ODOO_LINKED_CATEGORY = "Odoo Linked";
const GROUP_CATEGORY_PREFIX = "Grupo: ";
const ODOO_LINKED_NOTICE = "iccc-odoo-linked";


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

  try {
    const ort: any = (window as any).OfficeRuntime;
    const auth = ort?.auth;
    if (!auth?.getAccessToken) return null;

    const token = await auth.getAccessToken({
      allowSignInPrompt: false,
      allowConsentPrompt: false,
      forMSGraphAccess: true,
    });

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

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.masterCategories.getAsync((res: any) => {
        if (res.status !== OfficeAny.AsyncResultStatus.Succeeded) return resolve();
        const list = Array.isArray(res.value) ? res.value : [];
        const exists = list.some((c: any) => (c.displayName || c.name) === displayName);
        if (exists) return resolve();
        const color = OfficeAny.MailboxEnums?.CategoryColor?.Preset0;
        OfficeAny.context.mailbox.masterCategories.addAsync([{ displayName, color }], () => resolve());
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

async function addCategoriesToCurrentItem(displayNames: string[]): Promise<void> {
  const uniqueNames = Array.from(new Set((displayNames || []).map((name) => String(name || "").trim()).filter(Boolean)));
  if (!uniqueNames.length) return;
  const OfficeAny: any = await ensureOfficeReady().catch(() => null);
  if (!OfficeAny?.context?.mailbox?.item?.categories?.addAsync) return;

  await new Promise<void>((resolve) => {
    try {
      OfficeAny.context.mailbox.item.categories.addAsync(uniqueNames, () => resolve());
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
      OfficeAny.context.mailbox.item.categories.removeAsync(uniqueNames, () => resolve());
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

export async function syncOdooLinkedCategory(hasLinks: boolean): Promise<void> {
  if (hasLinks) {
    await ensureMasterCategory(ODOO_LINKED_CATEGORY);
    await addCategoryToCurrentItem(ODOO_LINKED_CATEGORY);
    return;
  }
  await removeCategoryFromCurrentItem(ODOO_LINKED_CATEGORY);
}

export async function syncManualGroupCategories(groupNames: string[]): Promise<void> {
  const uniqueNames = Array.from(
    new Set(
      (groupNames || [])
        .map((name) => String(name || "").trim())
        .filter(Boolean)
    )
  );
  const desiredCategories = uniqueNames.map((name) => `${GROUP_CATEGORY_PREFIX}${name}`);
  const currentCategories = await getCurrentItemCategoryNames();
  const currentGroupCategories = currentCategories.filter((name) => name.startsWith(GROUP_CATEGORY_PREFIX));
  const toRemove = currentGroupCategories.filter((name) => !desiredCategories.includes(name));
  const toAdd = desiredCategories.filter((name) => !currentCategories.includes(name));

  for (const categoryName of toAdd) {
    await ensureMasterCategory(categoryName);
  }
  await addCategoriesToCurrentItem(toAdd);
  await removeCategoriesFromCurrentItem(toRemove);
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

type CockpitHostAction =
  | { type: "close" }
  | { type: "open-email"; itemId?: string; emailWebLink?: string }
  | { type: "reply-current" }
  | { type: "forward-current" };

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

/**
 * Opens a separate window using Office Dialog API.
 * Guard: only one dialog at a time (evita "já existe uma dialog ativa").
 */
async function openCockpitView(view: string, params: Record<string, string>, options?: { height?: number; width?: number; displayInIframe?: boolean }) {
  const OfficeAny = await ensureOfficeReady();
  const url = buildCockpitViewUrl(view, params);

  clientLog.log(`[office] openDialog ${url.toString()}`);

  // close previous if any
  try {
    if (activeDialog) activeDialog.close();
  } catch { }
  activeDialog = null;

  return await new Promise<void>((resolve, reject) => {
    let settled = false;
    const resolveOnce = () => {
      if (settled) return;
      settled = true;
      resolve();
    };
    const rejectOnce = (error: Error) => {
      if (settled) return;
      settled = true;
      reject(error);
    };
    const timer = setTimeout(() => {
      clientLog.warn(`[office] displayDialogAsync timeout for ${view}`);
      rejectOnce(new Error("A abertura da janela demorou demasiado tempo."));
    }, 4000);

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
          if (!action) return;
          void executeCockpitHostAction(action)
            .catch((error) => clientLog.error("[office] host action failed", error))
            .finally(() => {
              if (action.type === "close") {
                resolveOnce();
              }
            });
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

export async function openGroupManager(params: Record<string, string> = {}) {
  try {
    return await openCockpitView("group-manager", params, { height: 82, width: 58, displayInIframe: true });
  } catch (error) {
    const url = buildCockpitViewUrl("group-manager", params);
    clientLog.warn("[office] group manager fallback to same-window navigation", error);
    window.location.assign(url.toString());
  }
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
export async function displayReplyForm(content: string, isHtml = true): Promise<void> {
  const OfficeAny = await ensureOfficeReady();
  const item = OfficeAny?.context?.mailbox?.item;
  if (!item?.displayReplyAllForm && !item?.displayReplyForm) {
    throw new Error("Funcionalidade de resposta não disponível neste item.");
  }

  const finalContent = isHtml ? content.replace(/\n/g, "<br/>") : content;

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
    const results: OutlookAttachment[] = [];

    for (const att of attachments) {
      // Only process file attachments
      if (att.attachmentType === "file") {
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
    }

    return results;
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
