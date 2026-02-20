// server/src/ai/openaiProvider.js
// Uses native fetch (Node 18+) to avoid axios dependency issues.
// Returns standard chat completions for GPT-4o-mini etc.

export async function openaiCreateResponse({
  apiKey,
  model = "gpt-4o-mini",
  instructions,
  input,
  files = [], // NEW: Support for images
  history = [], // NEW: Support for conversation history
  max_output_tokens = 512,
  temperature = 0.2,
  timeout_ms = 60000,
}) {
  if (!apiKey) throw Object.assign(new Error("OPENAI_API_KEY em falta"), { status: 400 });

  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeout_ms);

  const messages = [];
  if (instructions) {
    messages.push({ role: "system", content: instructions });
  }

  // Add history messages
  for (const h of history) {
    messages.push({ role: h.role, content: h.content });
  }

  const userContent = [];
  if (input) {
    userContent.push({ type: "text", text: input });
  }

  // Support images in OpenAI (png, jpg, jpeg, webp)
  for (const f of files) {
    if (f.type && f.type.startsWith("image/")) {
      const base64Data = f.content.includes(",") ? f.content.split(",")[1] : f.content;
      userContent.push({
        type: "image_url",
        image_url: { url: `data:${f.type};base64,${base64Data}` },
      });
    } else if (f.name) {
      // For non-images (like PDFs), we just mention the filename to OpenAI in text
      userContent.push({ type: "text", text: `[Nota: Ficheiro anexo "${f.name}" (${f.type}) foi ignorado por falta de suporte visual neste motor.]` });
    }
  }

  messages.push({ role: "user", content: userContent });

  try {
    const res = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        messages,
        max_tokens: max_output_tokens,
        temperature,
      }),
      signal: controller.signal,
    });

    const text = await res.text();
    let data;
    try {
      data = text ? JSON.parse(text) : null;
    } catch {
      data = { raw_text: text };
    }

    if (!res.ok) {
      const msg =
        data?.error?.message ||
        (typeof data?.message === "string" ? data.message : "") ||
        `OpenAI HTTP ${res.status}`;
      const err = new Error(msg);
      err.status = res.status;
      err.details = data;
      throw err;
    }

    const outputText = data?.choices?.[0]?.message?.content || "";
    return { raw: data, text: outputText.trim() };
  } catch (e) {
    if (e?.name === "AbortError") {
      const err = new Error("OpenAI timeout (abort)");
      err.status = 504;
      throw err;
    }
    if (e?.code === "ECONNRESET") {
      const err = new Error("OpenAI network ECONNRESET");
      err.status = 502;
      throw err;
    }
    throw e;
  } finally {
    clearTimeout(t);
  }
}
