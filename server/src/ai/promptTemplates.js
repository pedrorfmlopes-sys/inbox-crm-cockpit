// server/src/ai/promptTemplates.js
// Centralized prompt templates for the "MailMaestro-like" features.
// Keep these versioned and isolated from Odoo/CRM code.

export function buildPrompt({ action, locale = "pt-PT", tone = "neutro", email, inputText }) {
  const LOCALE_HUMAN = {
    "pt-PT": "Português (Portugal)",
    "es-ES": "Espanhol",
    "en-GB": "Inglês",
    "it-IT": "Italiano",
    "de-DE": "Alemão",
    auto: "Auto",
  };

  // ---- Language rules ----
  // Summaries are ALWAYS in Portuguese (Portugal) regardless of user selection.
  const isSummary = action === "summarize";
  const effectiveLocale = isSummary ? "pt-PT" : (locale || "pt-PT");

  // Human label (only for fixed languages)
  const lang = LOCALE_HUMAN[effectiveLocale] || effectiveLocale;

  // IMPORTANT:
  // - For "auto" replies, instruct model to answer in the same language as the email
  // - For non-auto, force the requested language
  const languageLine =
    effectiveLocale === "auto"
      ? `Responde no mesmo idioma em que o email está escrito.
Mantém tom profissional e objetivo.
`
      : `Escreve em ${lang}.
`;

  const rules =
    languageLine +
    `NUNCA inventes factos, números, prazos, preços ou compromissos. Se faltar informação, faz perguntas curtas.
` +
    `Devolve HTML simples e seguro: usa apenas <p>, <br>, <ul>, <ol>, <li>, <strong>, <em>, <a>.
` +
    `Sem CSS, sem estilos inline, sem classes, sem scripts.
` +
    `Evita linhas enormes: parágrafos curtos.
`;

  const toneLine = `Tom: ${tone}.`;

  const emailBlock = email
    ? `

CONTEXTO DO EMAIL:
Assunto: ${email.subject || ""}
De: ${email.from || ""}
Para: ${(email.to || []).join("; ")}
Cc: ${(email.cc || []).join("; ")}
Corpo (texto limpo):
${email.bodyText || ""}
`
    : "";

  if (action === "summarize") {
    return (
      rules +
      toneLine +
      `

TAREFA: Resume o email em 5–8 bullets e propõe 3–6 próximos passos (bullets).
Estrutura obrigatória:
<p><strong>Resumo</strong></p><ul>...</ul>
<p><strong>Próximos passos</strong></p><ul>...</ul>
<p><strong>Perguntas (se necessário)</strong></p><ul>...</ul>` +
      emailBlock
    );
  }

  if (action === "reply") {
    return (
      rules +
      toneLine +
      `

TAREFA: Cria uma resposta sugerida ao email.
Regras extra:
- Mantém o assunto implícito (não repitas "Re:").
- Usa uma saudação adequada.
- Se for preciso, faz 1–3 perguntas objetivas.
- Termina com fecho profissional.` +
      emailBlock
    );
  }

  if (action === "rewrite") {
    return (
      rules +
      toneLine +
      `

TAREFA: Reescreve o texto abaixo mantendo o significado, mas ajustando ao tom.
Texto para reescrever:
${inputText || ""}
`
    );
  }

  if (action === "tasks") {
    return (
      rules +
      toneLine +
      `

TAREFA: Extrai tarefas/ações do email (checklist) e identifica responsáveis (se possível) e prazos (se explícitos).
Estrutura:
<p><strong>Tarefas</strong></p><ul>...</ul>
<p><strong>Riscos/Dependências</strong></p><ul>...</ul>` +
      emailBlock
    );
  }

  // default (safe)
  return rules + toneLine + emailBlock;
}
