// server/src/ai/openaiProvider.js
// Uses native fetch (Node 18+) to avoid axios dependency issues.
// Returns structured errors (never crashes the process).

function extractTextFromResponses(data) {
  if (!data) return "";
  if (typeof data.output_text === "string") return data.output_text;

  const output = Array.isArray(data.output) ? data.output : [];
  const chunks = [];
  for (const item of output) {
    if (item?.type === "message" && Array.isArray(item?.content)) {
      for (const c of item.content) {
        if (c?.type === "output_text" && typeof c?.text === "string") chunks.push(c.text);
        if (c?.type === "text" && typeof c?.text === "string") chunks.push(c.text);
      }
    }
  }
  return chunks.join("\n").trim();
}

export async function openaiCreateResponse({
  apiKey,
  model,
  instructions,
  input,
  max_output_tokens = 256,
  temperature = 0.2,
  timeout_ms = 60000,
}) {
  if (!apiKey) throw Object.assign(new Error("OPENAI_API_KEY em falta"), { status: 400 });
  if (!model) throw Object.assign(new Error("model em falta"), { status: 400 });
  if (input == null || input === "") throw Object.assign(new Error("input em falta"), { status: 400 });

  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeout_ms);

  try {
    const res = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model,
        instructions: instructions || undefined,
        input,
        max_output_tokens,
        temperature,
      }),
      signal: controller.signal,
    });

    const text = await res.text();
    let data;
    try { data = text ? JSON.parse(text) : null; } catch { data = { raw_text: text }; }

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

    return { raw: data, text: extractTextFromResponses(data) };
  } catch (e) {
    // Normalize common errors
    if (e?.name === "AbortError") {
      const err = new Error("OpenAI timeout (abort)");
      err.status = 504;
      throw err;
    }
    // Some network errors can be ECONNRESET; we still return a clean JSON error to the client.
    if (e?.code === "ECONNRESET") {
      const err = new Error("OpenAI network ECONNRESET (ligação foi reiniciada)");
      err.status = 502;
      throw err;
    }
    throw e;
  } finally {
    clearTimeout(t);
  }
}
