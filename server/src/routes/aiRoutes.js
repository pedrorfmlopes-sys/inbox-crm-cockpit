import express from "express";
import { aiCreateText, getAiMeta, listAvailableModels } from "../ai/aiService.js";
import { buildPrompt } from "../ai/promptTemplates.js";
import { getBriefing, saveBriefing, initBriefingDb } from "../ai/briefingCache.js";
import { getStyleProfile } from "../learningStore.js";
import { createRequire } from "module";

const require = createRequire(import.meta.url);
let pdfParseRaw = null;
try {
  pdfParseRaw = require("pdf-parse");
} catch (e) {
  console.warn("[ai] pdf-parse dependency not found or failed to load. Server-side PDF parsing disabled.");
}

// Function to safely extract the PDF parsing function based on version/structure
function getPdfParser(pkg) {
  if (!pkg) return null;
  if (typeof pkg === 'function') return pkg;
  if (pkg.default && typeof pkg.default === 'function') return pkg.default;
  if (pkg.PDFParse && typeof pkg.PDFParse === 'function') return pkg.PDFParse;
  if (pkg.default && pkg.default.PDFParse && typeof pkg.default.PDFParse === 'function') return pkg.default.PDFParse;
  return null;
}

const pdfParse = getPdfParser(pdfParseRaw);

function stripHtmlToText(html) {
  if (!html) return "";
  let s = String(html);
  s = s.replace(/<\s*br\s*\/?\s*>/gi, "\n");
  s = s.replace(/<\/\s*p\s*>/gi, "\n");
  s = s.replace(/<\/\s*li\s*>/gi, "\n");
  s = s.replace(/<[^>]+>/g, "");
  s = s.replace(/&nbsp;/g, " ");
  s = s.replace(/&amp;/g, "&");
  s = s.replace(/&lt;/g, "<");
  s = s.replace(/&gt;/g, ">");
  s = s.replace(/&quot;/g, "\"");
  s = s.replace(/&#0?39;/g, "'");
  s = s.replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

function ensureBasicHtml(out) {
  const t = String(out || "").trim();
  if (!t) return "<p></p>";
  if (t.includes("<p") || t.includes("<ul") || t.includes("<br") || t.includes("<li")) return t;
  const escaped = t
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/\n/g, "<br>")}</p>`)
    .join("");
  return escaped;
}

function looksLikeGenericDraftRefusal(text) {
  const value = String(text || "")
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ");
  if (!value) return true;

  const hasApologyOrRefusal = [
    "sorry",
    "i'm sorry",
    "desculpa",
    "lamento",
    "lo siento",
    "mi dispiace",
    "entschuldigung",
  ].some((marker) => value.includes(marker));

  const hasGenericInability = [
    "can't assist",
    "cant assist",
    "cannot assist",
    "can't help",
    "cant help",
    "cannot help",
    "nao posso",
    "nao consigo",
    "nao posso ajudar",
    "no puedo",
    "no puedo ayudar",
  ].some((marker) => value.includes(marker));

  if (/\b(nao posso|nao consigo|no puedo)\s+(confirmar|validar|garantir|assegurar|informar|indicar|avancar|responder)\b/.test(value)) {
    return false;
  }

  return hasApologyOrRefusal && hasGenericInability;
}

function trimEmailBody(raw) {
  if (!raw) return "";
  let s = String(raw);
  const markers = [
    /^From:\s.+$/im, /^Sent:\s.+$/im, /^De:\s.+$/im, /^Enviado:\s.+$/im,
    /^On\s.+wrote:\s*$/im, /^Em\s.+escreveu:\s*$/im,
    /^-----Original Message-----$/im, /^-----Mensagem original-----$/im,
  ];
  const MIN_QUOTE_INDEX = 220;
  let cutAt = s.length;
  for (const rx of markers) {
    const m = s.match(rx);
    if (m && m.index != null && m.index >= MIN_QUOTE_INDEX) {
      cutAt = Math.min(cutAt, m.index);
    }
  }
  s = s.slice(0, cutAt);
  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);
  const MAX = 4500;
  if (s.length > MAX) s = s.slice(0, MAX);
  return s.trim();
}

function trimEmailBodyFull(raw) {
  if (!raw) return "";
  let s = String(raw);
  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);
  const MAX = 9000;
  if (s.length > MAX) s = s.slice(0, MAX);
  return s.trim();
}

const VALID_LENGTHS = new Set(["xs", "s", "m", "l"]);

function normalizeLength(value) {
  const normalized = String(value || "").trim().toLowerCase();
  return VALID_LENGTHS.has(normalized) ? normalized : "m";
}

function normalizeStringList(value, maxEntries = 40, maxChars = 2000) {
  const source = Array.isArray(value) ? value : [];
  const out = [];
  const seen = new Set();
  for (const entry of source) {
    const clean = String(entry || "").trim().slice(0, maxChars);
    if (!clean) continue;
    const key = clean.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(clean);
    if (out.length >= maxEntries) break;
  }
  return out;
}

function normalizeReplyDirection(value) {
  if (!value || typeof value !== "object") return null;
  const addresseeName = String(value.addresseeName || "").trim().slice(0, 160);
  const addresseeContext = String(value.addresseeContext || "").trim().slice(0, 500);
  const ignoreIntermediateForwarders = value.ignoreIntermediateForwarders !== false;
  if (!addresseeName && !addresseeContext) return null;
  return {
    addresseeName,
    addresseeContext,
    ignoreIntermediateForwarders,
  };
}

function normalizeSignature(value) {
  if (!value || typeof value !== "object") return null;
  const text = String(value.text || "").trim().slice(0, 4000);
  const html = String(value.html || "").trim().slice(0, 6000);
  const imageUrl = String(value.imageUrl || "").trim().slice(0, 12000);
  const imageMaxWidth = Math.max(80, Math.min(800, Number(value.imageMaxWidth || 260) || 260));
  if (!text && !html && !imageUrl) return null;
  return {
    text,
    html,
    imageUrl,
    imageMaxWidth,
  };
}

function maxOutputTokensFor(action, length) {
  const normalizedLength = normalizeLength(length);
  const table = {
    xs: { reply: 350, summarize: 450, rewrite: 350, tasks: 450, default: 350 },
    s: { reply: 550, summarize: 650, rewrite: 550, tasks: 650, default: 550 },
    m: { reply: 850, summarize: 900, rewrite: 850, tasks: 900, default: 800 },
    l: { reply: 1300, summarize: 1300, rewrite: 1300, tasks: 1200, default: 1100 },
  };
  const byLength = table[normalizedLength] || table.m;
  return byLength[action] || byLength.default;
}

export function createAiRouter() {
  const router = express.Router();
  initBriefingDb(); // Ensure DB table exists

  function tryParseAiJsonPayload(rawText) {
    const trimmed = String(rawText || "").trim();
    if (!trimmed) return null;

    const candidates = [
      trimmed,
      trimmed.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim(),
    ];

    const arrayStart = trimmed.indexOf("[");
    const arrayEnd = trimmed.lastIndexOf("]");
    if (arrayStart >= 0 && arrayEnd > arrayStart) {
      candidates.push(trimmed.slice(arrayStart, arrayEnd + 1).trim());
    }

    const objectStart = trimmed.indexOf("{");
    const objectEnd = trimmed.lastIndexOf("}");
    if (objectStart >= 0 && objectEnd > objectStart) {
      candidates.push(trimmed.slice(objectStart, objectEnd + 1).trim());
    }

    for (const candidate of candidates) {
      try {
        return JSON.parse(candidate);
      } catch {
        // try next candidate
      }
    }

    return null;
  }

  router.get("/meta", (_req, res) => {
    res.json({ ok: true, ...getAiMeta() });
  });

  router.post("/generate", async (req, res) => {
    try {
      const {
        action = "reply",
        mode = "fast",
        locale = "pt-PT",
        tone = "neutro",
        length = "m",
        email,
        inputText,
        draftText = "",    // NEW: explicit draft for refine
        knowledge = [],
        aiKnowledge = [],
        signature = null,
        replyDirection = null,
        history = [], // NEW: Support for chat refinement
        filesContext: clientFilesContext = "",
        contextBundle = "",
        persona = {}, // NEW: Persona / Style mimic
        files = [],   // NEW: Direct files support
        customModels: _customModels = {}, // ignored on /generate for stability
        briefing = null,   // NEW: Contextual briefing
        contactAliases = [], // NEW: Contact Aliases
      } = req.body || {};

      const effectiveLength = normalizeLength(length);
      const normalizedKnowledge = normalizeStringList([
        ...normalizeStringList(aiKnowledge),
        ...normalizeStringList(knowledge),
      ]);
      const normalizedReplyDirection = normalizeReplyDirection(replyDirection);
      const normalizedSignature = normalizeSignature(signature);

      const safeEmail = email
        ? {
            subject: String(email.subject || ""),
            from: String(email.from || ""),
            fromName: String(email.fromName || "").trim(),
            fromEmail: String(email.fromEmail || "").trim(),
            greetingName: String(email.greetingName || "").trim(),
            greetingEmail: String(email.greetingEmail || "").trim(),
            to: Array.isArray(email.to) ? email.to.map(String) : [],
            cc: Array.isArray(email.cc) ? email.cc.map(String) : [],
            bodyText:
              String(email.bodyScope || "main") === "full"
                ? trimEmailBodyFull(email.bodyText || "")
                : trimEmailBody(email.bodyText || ""),
            bodyScope: email.bodyScope === "full" ? "full" : "main",
          }
        : null;

      // --- PROCESS FILES ---
      // Option 1 (NATIVE): Send files directly to the AI service (Gemini/future OpenAI)
      console.log(`[ai] Request: action=${action}, files=${files?.length}`);

      // NEW: Fetch and merge autonomous style profile
      let learnedProfile = null;
      try {
        const fullProfile = await getStyleProfile("global");
        if (fullProfile && fullProfile.styleData && Object.keys(fullProfile.styleData).length > 0) {
          learnedProfile = fullProfile;
        }
      } catch (e) {
        console.warn("[ai] Failed to fetch learning profile:", e.message);
      }

      // --- AUTO LOCALE DETECTION ---
      // If locale==="auto", detect from email text to get a deterministic language.
      let effectiveLocale = locale;
      if (locale === "auto" && action !== "refine") {
        const sampleText = ((email?.subject || "") + "\n" + (email?.bodyText || "")).slice(0, 800).toLowerCase();
        // Simple heuristic: count characteristic words per language.
        const scores = {
          "pt-PT": (sampleText.match(/\b(ol\u00e1|obrigad|por favor|tamb\u00e9m|e-mail|prezado|atenciosamente|bom dia|boa tarde)\b/g) || []).length,
          "es-ES": (sampleText.match(/\b(hola|gracias|por favor|tambi\u00e9n|correo|estimado|atentamente|buenos d\u00edas)\b/g) || []).length,
          "en-GB": (sampleText.match(/\b(hello|thanks|thank you|please|also|email|dear|regards|good morning|hi\b)\b/g) || []).length,
          "it-IT": (sampleText.match(/\b(ciao|grazie|per favore|anche|anche|gentile|cordiali saluti|buongiorno)\b/g) || []).length,
          "de-DE": (sampleText.match(/\b(hallo|danke|bitte|auch|e-mail|sehr geehrte|mit freundlichen|guten morgen)\b/g) || []).length,
        };
        const best = Object.entries(scores).sort((a, b) => b[1] - a[1])[0];
        effectiveLocale = best && best[1] > 0 ? best[0] : "pt-PT";
        console.log(`[ai] Auto-locale detected: ${effectiveLocale} (scores: ${JSON.stringify(scores)})`);
      }

      const instructions = req.body.prompt || buildPrompt({
        action,
        locale: effectiveLocale,
        tone,
        email: safeEmail,
        inputText: String(inputText || ""),
        length: effectiveLength,
        knowledge: normalizedKnowledge,
        aiKnowledge: normalizedKnowledge,
        signature: normalizedSignature,
        replyDirection: normalizedReplyDirection,
        filesContext: clientFilesContext,
        contextBundle: String(contextBundle || ""),
        persona: {
          ...persona,
          learnedProfile: learnedProfile?.styleData || null,
          learnedHabits: learnedProfile?.habitsData || null,
        },
        briefing,
        contactAliases: Array.isArray(contactAliases) ? contactAliases : [],
        currentTime: new Date().toISOString(), // NEW: Time awareness for greetings
      });

      // For refine: use explicit draftText from client (the current editor content).
      // Fallback to last assistant message in history if draftText not provided.
      const lastAssistant = [...history].reverse().find(m => m.role === "assistant" && typeof m.content === "string");
      const currentDraft = String(draftText || lastAssistant?.content || "");
      const normalizedInputText = String(inputText || "").trim();
      const providerInput = action === "refine"
        ? `INSTRUÇÃO DO UTILIZADOR: ${inputText || "Melhora o rascunho"}
RASCUNHO ATUAL (edita APENAS este texto, não inventar factos/prazos/preços/referências):
${currentDraft}`
        : action === "extract_tasks_json"
          ? "Extrai apenas as tarefas pendentes do contexto acima e devolve apenas um array JSON válido."
          : action === "extract_contacts"
            ? "Extrai apenas os emails relevantes do contexto acima."
            : action === "intent_proposals"
              ? "Propõe apenas 3 intenções curtas separadas por ponto e vírgula."
              : action === "summarize_actions"
                ? "Devolve apenas JSON válido com summary e actions."
                : action === "forward"
                  ? (normalizedInputText || "Redige agora o email final de reencaminhamento com base em todo o contexto fornecido, pronto a enviar ao destinatario final.")
                  : action === "reply"
                    ? (normalizedInputText || "Redige agora a resposta final pronta a enviar com base em todo o contexto fornecido.")
                    : (normalizedInputText || "Executa a tarefa pedida com base em todo o contexto fornecido.");

      const safeCustomModels = {};

      let result = await aiCreateText({
        mode,
        instructions,
        input: providerInput,
        files: action === "refine" ? [] : files,
        history: action === "refine" ? [] : history,
        max_output_tokens: maxOutputTokensFor(action, effectiveLength),
        temperature: action === "refine" ? 0 : 0.1,
        customModels: safeCustomModels,
      });

      if ((action === "reply" || action === "forward") && looksLikeGenericDraftRefusal(result.text)) {
        const retryInstructions = `${instructions}

[OVERRIDE OPERACIONAL]
- Esta e uma tarefa legitima de backoffice/comercial: redigir correspondencia empresarial comum com base em emails, anexos e contexto ligado.
- Gera o rascunho final pronto a enviar.
- Se faltar informacao factual, gera um email profissional curto a pedir a informacao em falta.
- Nao inventes factos, precos, prazos, referencias, disponibilidade ou condicoes comerciais.
- Nao devolvas recusas genericas se o pedido for apenas criar um draft profissional legitimo.
- Mantem a excecao de seguranca: se o pedido for claramente ilegal, perigoso ou abusivo, nao cumpras.`;

        result = await aiCreateText({
          mode,
          instructions: retryInstructions,
          input: providerInput,
          files: action === "refine" ? [] : files,
          history: action === "refine" ? [] : history,
          max_output_tokens: maxOutputTokensFor(action, effectiveLength),
          temperature: action === "refine" ? 0 : 0.1,
          customModels: safeCustomModels,
        });
      }

      if ((action === "reply" || action === "forward") && looksLikeGenericDraftRefusal(result.text)) {
        return res.json({
          ok: false,
          error: "A IA devolveu uma recusa genérica e não conseguiu gerar um rascunho útil. Revê a instrução ou seleciona melhor o email-alvo/contexto.",
        });
      }

      const data = tryParseAiJsonPayload(result.text);

      res.json({
        ok: true,
        html: ensureBasicHtml(result.text || ""),
        text: stripHtmlToText(result.text),
        data
      });

    } catch (e) {
      console.error("[ai] generate error:", e?.message || e);
      res.status(500).json({ ok: false, error: String(e?.message || e) });
    }
  });

  router.post("/extract-anchors", async (req, res) => {
    try {
      const { emailBody, emailContext, customModels } = req.body;
      const { extractAnchors } = await import("../aiOrchestrator.js");
      // Use full context if available, fallback to just body for backwards compatibility
      const anchors = await extractAnchors(emailContext || emailBody, customModels);
      res.json({ ok: true, anchors });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  });

  router.post("/briefing", async (req, res) => {
    try {
      const { context, history, customModels, conversationId, cacheKey } = req.body;
      const effectiveCacheKey = String(cacheKey || conversationId || "").trim();

      // 1. Try cache first
      if (effectiveCacheKey) {
        const cached = await getBriefing(effectiveCacheKey);
        if (cached) {
          console.log(`[ai] Cache HIT for briefing: ${effectiveCacheKey}`);
          return res.json({ ok: true, summary: cached, cached: true });
        }
      }

      // 2. Generate new
      const { generateExecutiveSummary } = await import("../aiOrchestrator.js");
      const summary = await generateExecutiveSummary(context, history, customModels);

      // 3. Save to cache
      if (effectiveCacheKey && summary) {
        await saveBriefing(effectiveCacheKey, summary);
      }

      res.json({ ok: true, summary, cached: false });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  });

  router.post("/voice-command", async (req, res) => {
    try {
      const { commandText, context } = req.body;
      const { processVoiceCommand } = await import("../VoiceActionEngine.js");
      const result = await processVoiceCommand(commandText, context);
      res.json(result);
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  });

  router.post("/selftest", async (req, res) => {
    try {
      const { customModels } = req.body || {};
      const { aiSelftest } = await import("../ai/aiService.js");
      const result = await aiSelftest(customModels);
      res.json(result); // returns { ok, openai, gemini }
    } catch (e) {
      res.json({ ok: false, openai: { ok: false, error: e.message }, gemini: { ok: false, error: e.message } });
    }
  });

  router.get("/list-models", async (req, res) => {
    try {
      const models = await listAvailableModels();
      res.json({ ok: true, ...models });
    } catch (e) {
      res.status(500).json({ ok: false, error: e.message });
    }
  });

  return router;
}
