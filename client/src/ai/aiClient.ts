// client/src/ai/aiClient.ts
export type AiAction = "summarize" | "reply" | "rewrite" | "forward" | "tasks" | "refine" | "intent_proposals" | "extract_contacts" | "extract_tasks_json";
export type AiMode = "fast" | "quality";
export type AiTone = "neutro" | "formal" | "curto" | "direto" | "simpático";
export type AiLocale = "pt-PT" | "es-ES" | "en-GB" | "it-IT" | "de-DE" | "auto";

export type AiEmailContext = {
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  /**
   * "main": apenas a mensagem principal (pode cortar citações/reencaminhados)
   * "full": inclui todo o corpo disponível (incl. reencaminhados/citações)
   */
  bodyScope?: "main" | "full";
  bodyText: string;
};

export type AiGenerateRequest = {
  action: AiAction;
  mode: AiMode;
  locale: AiLocale;
  tone: AiTone;
  email?: AiEmailContext;
  inputText?: string;
  knowledge?: string[];
  files?: any[]; // optional: raw files
  history?: Array<{ role: "user" | "assistant"; content: string }>; // NEW: Chat history
  filesContext?: string; // optional: pre-extracted text context
  persona?: any; // NEW: User persona / style
  briefing?: string | null; // NEW: Automated contextual briefing
  contactAliases?: Array<{ id: string; name: string; email: string }>; // NEW: Contact aliases
  customModels?: {
    openaiModelFast?: string;
    openaiModelQuality?: string;
    geminiModel?: string;
    openaiApiKey?: string;
    geminiApiKey?: string;
  };
};

export type AiGenerateResponse =
  | {
    ok: true;
    html: string;
    text: string;
    suggestedRecipients?: { to: string[]; cc: string[] };
    suggestedSubject?: string;
  }
  | { ok: false; error: string };

async function requestJSON<T>(url: string, init?: RequestInit): Promise<T> {
  const res = await fetch(url, {
    ...init,
    headers: { "Content-Type": "application/json", ...(init?.headers || {}) },
  });
  const text = await res.text();
  if (!res.ok) throw new Error(`HTTP ${res.status}: ${text.slice(0, 400)}`);
  try {
    return JSON.parse(text) as T;
  } catch {
    throw new Error(`JSON inválido: ${text.slice(0, 400)}`);
  }
}

export async function aiGenerate(payload: AiGenerateRequest): Promise<AiGenerateResponse> {
  return requestJSON<AiGenerateResponse>("/api/ai/generate", {
    method: "POST",
    body: JSON.stringify(payload),
  });
}
