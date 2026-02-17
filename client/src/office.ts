import { clientLog } from "./logger";

declare const Office: any;

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

async function ensureOfficeReady() {
  const OfficeAny = await waitForOffice(8000);
  if (!OfficeAny) throw new Error("Office.js não está disponível (o taskpane não correu dentro do Outlook?).");
  await new Promise<void>((resolve) => OfficeAny.onReady(() => resolve()));
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
  const OfficeAny = await ensureOfficeReady();

  const item = OfficeAny?.context?.mailbox?.item;
  if (!item) {
    clientLog.warn("[office] mailbox.item is empty");
    return {};
  }

  const getAsyncValue = async (obj: any, coercer: (v: any) => string): Promise<string> => {
    if (!obj?.getAsync) return "";
    return await new Promise<string>((resolve) => {
      try {
        obj.getAsync((r: any) => {
          try {
            if (r?.status === OfficeAny.AsyncResultStatus.Succeeded) resolve(coercer(r.value));
            else resolve("");
          } catch {
            resolve("");
          }
        });
      } catch {
        resolve("");
      }
    });
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
  };
}


// Backwards-compat with older UI code
export const getSelectedMessageContext = getOutlookContext;

// Ler corpo do email (texto simples) — usado pela IA
export async function getEmailBodyText(): Promise<string> {
  try {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item: any = (Office as any)?.context?.mailbox?.item;
    if (!item?.body?.getAsync) return "";
    return await new Promise<string>((resolve) => {
      item.body.getAsync("text", (r: any) => {
        try {
          if (r?.status === (Office as any).AsyncResultStatus.Succeeded) resolve(String(r.value ?? ""));
          else resolve("");
        } catch {
          resolve("");
        }
      });
    });
  } catch {
    return "";
  }
}


// Token barato para detetar mudanca de email (para polling fallback ao ItemChanged)
export async function getCurrentItemToken(): Promise<string> {
  try {
    const OfficeAny = await ensureOfficeReady();
    const item = OfficeAny?.context?.mailbox?.item;
    if (!item) return "";
    const cid = typeof item.conversationId === "string" ? item.conversationId : "";
    const imid = typeof item.internetMessageId === "string" ? item.internetMessageId : "";
    const itemId = typeof item.itemId === "string" ? item.itemId : "";
    const created = item.dateTimeCreated ? String(item.dateTimeCreated) : "";
    const subj = typeof item.subject === "string" ? item.subject : "";
    return [cid, imid, itemId, created, subj].filter(Boolean).join("|");
  } catch {
    return "";
  }
}

let activeDialog: any = null;

/**
 * Opens a separate window using Office Dialog API.
 * Guard: only one dialog at a time (evita "já existe uma dialog ativa").
 */
export async function openCockpitDialog(params: Record<string, string>) {
  const OfficeAny = await ensureOfficeReady();

  const url = new URL(window.location.origin);
  url.searchParams.set("view", "dialog");
  Object.entries(params || {}).forEach(([k, v]) => url.searchParams.set(k, v));

  clientLog.log(`[office] openDialog ${url.toString()}`);

  // close previous if any
  try {
    if (activeDialog) activeDialog.close();
  } catch {}
  activeDialog = null;

  return await new Promise<void>((resolve, reject) => {
    OfficeAny.context.ui.displayDialogAsync(
      url.toString(),
      { height: 65, width: 40, displayInIframe: false },
      (result: any) => {
        if (result.status !== OfficeAny.AsyncResultStatus.Succeeded) {
          clientLog.error(`[office] displayDialogAsync failed: ${result.error?.message || "unknown"}`);
          reject(new Error(result.error?.message || "Falha ao abrir janela (Dialog)."));
          return;
        }
        const dialog = result.value;
        activeDialog = dialog;

        dialog.addEventHandler(OfficeAny.EventType.DialogMessageReceived, (arg: any) => {
          if (arg?.message === "close") {
            try {
              dialog.close();
            } catch {}
            activeDialog = null;
            resolve();
          }
        });

        dialog.addEventHandler(OfficeAny.EventType.DialogEventReceived, () => {
          activeDialog = null;
          resolve();
        });
      }
    );
  });
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
        } catch {}
      };
    }
  } catch (e) {
    clientLog.warn("[office] ItemChanged not supported here", e);
  }

  return () => {};
}
