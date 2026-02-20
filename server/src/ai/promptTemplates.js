// server/src/ai/promptTemplates.js
// Centralized prompt templates for the "MailMaestro-like" features.
// Keep these versioned and isolated from Odoo/CRM code.

export function buildPrompt({ action, locale = "pt-PT", tone = "neutro", email, inputText, knowledge = [], filesContext = "" }) {
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
      ? `Responde no mesmo idioma em que o email está escrito.\nMantém tom profissional e objetivo.\n`
      : `Escreve em ${lang}.\n`;

  const baseRules = languageLine +
    `NUNCA inventes factos, números, prazos, preços ou compromissos. Se faltar informação, faz perguntas curtas.\n` +
    `Devolve HTML simples e seguro: usa apenas <p>, <br>, <ul>, <ol>, <li>, <strong>, <em>, <a>.\n` +
    `Sem CSS, sem estilos inline, sem classes, sem scripts.\n` +
    `Evita linhas enormes: parágrafos curtos.\n`;

  // Inject user knowledge if present
  const knowledgeBlock = knowledge.length > 0
    ? `\nREGRAS/CONHECIMENTO EXTRA DO UTILIZADOR:\n${knowledge.map(k => `- ${k}`).join('\n')}\n`
    : "";

  // Inject file content if present
  const filesBlock = filesContext
    ? `\nCONTEÚDO DOS FICHEIROS ANEXOS (Para tua análise):\n${filesContext}\n`
    : "";

  const finalRules = baseRules + knowledgeBlock + filesBlock;
  const toneLine = `Tom: ${tone}.`;

  const emailBlock = email
    ? `\n\nCONTEXTO DO EMAIL:\nAssunto: ${email.subject || ""}\nDe: ${email.from || ""}\nPara: ${(email.to || []).join("; ")}\nCc: ${(email.cc || []).join("; ")}\nCorpo (texto limpo):\n${email.bodyText || ""}\n`
    : "";

  if (action === "summarize") {
    // Check if user is asking for a detailed file analysis
    if (inputText && inputText.includes("ANÁLISE DETALHADA")) {
      return (
        finalRules +
        toneLine +
        `\n\nTAREFA: O utilizador pediu uma ANÁLISE DETALHADA do anexo. NÃO faças um resumo genérico.\n` +
        `1. Identifica o tipo de documento (Fatura, Extrato, Proposta, etc).\n` +
        `2. EXTRAI DADOS CONCRETOS: Datas, Números de Fatura, Valores Totais, Tabelas de itens.\n` +
        `3. Lista os itens/movimentos principais encontrados.\n` +
        `4. Se for um extrato de conta, lista os movimentos pendentes ou saldos.\n` +
        `\nINSTRUÇÃO DO UTILIZADOR: "${inputText}"\n` +
        emailBlock
      );
    }

    const userInstruction = inputText ? `\n\nINSTRUÇÃO ESPECÍFICA DO UTILIZADOR: "${inputText}" (Prioriza esta instrução sobre o resumo genérico).` : "";
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Resume o email (e os anexos, se houver) em 5–8 bullets e propõe 3–6 próximos passos (bullets).\n` +
      `Estrutura obrigatória:\n<p><strong>Resumo</strong></p><ul>...</ul>\n<p><strong>Próximos passos</strong></p><ul>...</ul>\n<p><strong>Perguntas (se necessário)</strong></p><ul>...</ul>` +
      userInstruction +
      emailBlock
    );
  }

  if (action === "reply") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Cria uma resposta sugerida ao email (considerando anexos se relevante).\nRegras extra:\n- Mantém o assunto implícito (não repitas "Re:").\n- Usa uma saudação adequada.\n- Se for preciso, faz 1–3 perguntas objetivas.\n- Termina com fecho profissional.` +
      emailBlock
    );
  }

  if (action === "rewrite") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Reescreve o texto fornecido pelo utilizador em 'inputText' (se existir) ou a resposta anterior, mantendo o significado mas aplicando o tom pedido.\nTexto para reescrever:\n${inputText || ""}\n`
    );
  }

  if (action === "refine") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: O utilizador enviou uma instrução para REFINAR a tua resposta anterior.\n` +
      `REGRAS CRÍTICAS:\n` +
      `1. Aplica a instrução do utilizador ao conteúdo da tua última resposta.\n` +
      `2. Devolve APENAS o conteúdo final alterado. Proibido usar "Aqui está", "Entendido" ou qualquer introdução.\n` +
      `3. Mantém a estrutura HTML pedida.\n` +
      `4. Se a instrução for uma tradução, traduz todo o bloco anterior.\n`
    );
  }

  if (action === "forward") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Escreve um rascunho de email para REENVIAR a uma terceira entidade (não o remetente original) com base no assunto em contexto.\nRegras extra:\n- O rascunho deve começar com "[Rascunho para Reenvio]".\n- Mantém o tom profissional.\n- Explica o contexto do email original se necessário.\n- Se o utilizador forneceu instruções em 'inputText', segue-as rigorosamente: "${inputText || ""}"\n- Devolve apenas o corpo do email.` +
      emailBlock
    );
  }

  if (action === "tasks") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Extrai tarefas/ações do email e anexos (checklist) e identifica responsáveis (se possível) e prazos (se explícitos).\nEstrutura:\n<p><strong>Tarefas</strong></p><ul>...</ul>\n<p><strong>Riscos/Dependências</strong></p><ul>...</ul>` +
      emailBlock
    );
  }

  // default (safe)
  return finalRules + toneLine + emailBlock;
}
