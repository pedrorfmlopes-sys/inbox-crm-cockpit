import express from "express";
import { aiCreateText, getAiMeta, listAvailableModels } from "../ai/aiService.js";
import { buildPrompt } from "../ai/promptTemplates.js";
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

export function createAiRouter() {
  const router = express.Router();

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
        email,
        inputText,
        knowledge = [],
        history = [], // NEW: Support for chat refinement
        filesContext: clientFilesContext = "",
        persona = {}, // NEW: Persona / Style mimic
        files = [],   // NEW: Direct files support
        customModels = {}, // NEW: Custom models from client
      } = req.body || {};

      const safeEmail = email
        ? {
          subject: String(email.subject || ""),
          from: String(email.from || ""),
          to: Array.isArray(email.to) ? email.to.map(String) : [],
          cc: Array.isArray(email.cc) ? email.cc.map(String) : [],
          bodyText:
            String(email.bodyScope || "main") === "full"
              ? trimEmailBodyFull(email.bodyText || "")
              : trimEmailBody(email.bodyText || ""),
        }
        : null;

      // --- PROCESS FILES ---
      // Option 1 (NATIVE): Send files directly to the AI service (Gemini/future OpenAI)
      console.log(`[ai] Request: action=${action}, files=${files?.length}`);

      const instructions = req.body.prompt || buildPrompt({
        action,
        locale,
        tone,
        email: safeEmail,
        inputText: String(inputText || ""),
        knowledge: Array.isArray(knowledge) ? knowledge.map(String) : [],
        filesContext: clientFilesContext,
        persona,
      });

      const result = await aiCreateText({
        mode,
        instructions,
        input: action === "refine" ? (inputText || "Refinar") : "ok",
        files,
        history,
        max_output_tokens: action === "summarize" || action === "tasks" || action === "summarize_actions" ? 800 : 700,
        temperature: 0.1,
        customModels,
      });

      let data = null;
      if (result.text && (result.text.includes("{") || action === "summarize_actions")) {
        try {
          const jsonStr = result.text.substring(result.text.indexOf("{"), result.text.lastIndexOf("}") + 1);
          data = JSON.parse(jsonStr);
        } catch (e) {
          // ignore parsing error
        }
      }

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
      const { context, history, customModels } = req.body;
      const { generateExecutiveSummary } = await import("../aiOrchestrator.js");
      const summary = await generateExecutiveSummary(context, history, customModels);
      res.json({ ok: true, summary });
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
      res.json({ ok: false, openai: false, gemini: false, error: e.message });
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
