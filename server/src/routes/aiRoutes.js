// server/src/routes/aiRoutes.js
import express from "express";
import { aiCreateText, getAiMeta } from "../ai/aiService.js";
import { buildPrompt } from "../ai/promptTemplates.js";

function stripHtmlToText(html) {
  if (!html) return "";
  let s = String(html);

  // Normalize breaks / paragraphs to newlines
  s = s.replace(/<\s*br\s*\/?\s*>/gi, "\n");
  s = s.replace(/<\/\s*p\s*>/gi, "\n");
  s = s.replace(/<\/\s*li\s*>/gi, "\n");

  // Remove all tags
  s = s.replace(/<[^>]+>/g, "");

  // Decode basic entities
  s = s.replace(/&nbsp;/g, " ");
  s = s.replace(/&amp;/g, "&");
  s = s.replace(/&lt;/g, "<");
  s = s.replace(/&gt;/g, ">");
  s = s.replace(/&quot;/g, "\"");
  s = s.replace(/&#0?39;/g, "'");

  // Cleanup
  s = s.replace(/\n{3,}/g, "\n\n").trim();
  return s;
}

function ensureBasicHtml(out) {
  const t = String(out || "").trim();
  if (!t) return "<p></p>";
  // If it already looks like HTML, return as-is
  if (t.includes("<p") || t.includes("<ul") || t.includes("<br") || t.includes("<li")) return t;
  // Otherwise, wrap paragraphs
  const escaped = t
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/\n/g, "<br>")}</p>`)
    .join("");
  return escaped;
}

// Basic email trimming server-side (client also trims, but keep defense-in-depth)
function trimEmailBody(raw) {
  if (!raw) return "";
  let s = String(raw);

  // Common "reply" separators
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

  // Many forwarded emails include a mini-header inside the body ("De:", "Enviada:", etc.)
  // right at the top. If we cut at the first marker, we lose the real message content.
  // So we only treat these markers as a "quote boundary" if they appear sufficiently
  // later in the body.
  const MIN_QUOTE_INDEX = 220;

  let cutAt = s.length;
  for (const rx of markers) {
    const m = s.match(rx);
    if (!m || m.index == null) continue;
    const idx = m.index;
    if (idx < MIN_QUOTE_INDEX) continue;
    cutAt = Math.min(cutAt, idx);
  }
  s = s.slice(0, cutAt);

  // Signature delimiter
  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);

  // Limit length hard (cost control)
  const MAX = 4500; // chars
  if (s.length > MAX) s = s.slice(0, MAX);

  return s.trim();
}

// "Full" mode: keep the whole body (incl. forwarded/quoted content), only remove signature
// and cap size for cost control.
function trimEmailBodyFull(raw) {
  if (!raw) return "";
  let s = String(raw);

  const sig = s.indexOf("\n-- \n");
  if (sig > 0) s = s.slice(0, sig);

  // Slightly higher cap because "full" is used on demand.
  const MAX = 9000;
  if (s.length > MAX) s = s.slice(0, MAX);

  return s.trim();
}

export function createAiRouter() {
  const router = express.Router();

  router.get("/meta", (_req, res) => {
    res.json({ ok: true, ...getAiMeta() });
  });

  router.get("/ping", (_req, res) => {
    res.json({ ok: true, ...getAiMeta(), now: new Date().toISOString() });
  });

  const selftestHandler = async (_req, res) => {
    try {
      console.log("[ai] selftest called");
      const result = await aiCreateText({
        mode: "fast",
        instructions: "És um healthcheck. Responde apenas com 'OK'.",
        input: "ping",
        max_output_tokens: 16,
        temperature: 0,
      });
      res.json({ ok: true, text: result.text || "" });
    } catch (e) {
      console.error("[ai] selftest error:", e?.status || "", e?.message || e);
      res.status(e?.status || 500).json({ ok: false, error: String(e?.message || e) });
    }
  };

  router.get("/selftest", selftestHandler);
  router.post("/selftest", selftestHandler);

  /**
   * POST /api/ai/generate
   * body:
   * {
   *   action: "reply"|"summarize"|"rewrite"|"tasks",
   *   mode: "fast"|"quality",
   *   locale: "pt-PT",
   *   tone: "neutro"|"formal"|"curto"|"direto"|"simpático",
   *   email: { subject, from, to:[], cc:[], bodyText },
   *   inputText?: string
   * }
   *
   * returns: { ok:true, html:"...", text:"..." }
   */
  router.post("/generate", async (req, res) => {
    try {
      const {
        action = "reply",
        mode = "fast",
        locale = "pt-PT",
        tone = "neutro",
        email,
        inputText,
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

      const instructions = buildPrompt({
        action,
        locale,
        tone,
        email: safeEmail,
        inputText: String(inputText || ""),
      });

      // Use empty input; most of the task is in instructions (reduces prompt duplication)
      const result = await aiCreateText({
        mode,
        instructions,
        input: "ok",
        max_output_tokens:
          action === "summarize" || action === "tasks" ? 600 : action === "reply" ? 700 : 500,
        temperature: 0.25,
      });

      const html = ensureBasicHtml(result.text || "");
      const text = stripHtmlToText(html);

      res.json({ ok: true, html, text });
    } catch (e) {
      console.error("[ai] generate error:", e?.status || "", e?.message || e);
      res.status(e?.status || 500).json({ ok: false, error: String(e?.message || e) });
    }
  });

  return router;
}
