// server/src/ai/promptTemplates.js
// Centralized prompt templates for the "MailMaestro-like" features.
// Keep these versioned and isolated from Odoo/CRM code.

export function buildPrompt({ action, locale = "pt-PT", tone = "neutro", email, inputText, knowledge = [], filesContext = "", persona = {}, briefing = null, contactAliases = [] }) {
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

  // REGRAS DE IDIOMA (User feedback: Strict enforcement)
  const languageEnforcement =
    effectiveLocale === "auto"
      ? `[REGRA DE IDIOMA]: DEVES DETETAR O IDIOMA DO EMAIL ORIGINAL E RESPONDER NO MESMO IDIOMA.\n`
      : `[REGRA DE IDIOMA]: DEVES ESCREVER ABSOLUTAMENTE TUDO EM ${lang.toUpperCase()}. IGNORA QUALQUER OUTRA INSTRUÇÃO EM PORTUGUÊS QUE POSSA SUGERIR O CONTRÁRIO.\n`;

  // REGRAS DE PRAGMATISMO (User feedback: No echo in replies, context in forwards)
  const isReply = action === "reply" || action === "refine";
  const pragmatismRules = isReply
    ? `- PROIBIDO REPETIR FACTOS: Se o remetente diz que "X está pronto", NÃO respondas "Confirmo que X está pronto". Apenas agradece ou passa ao próximo passo.\n` +
    `- EVITA O "EFEITO ESPELHO": Não resumas o email do remetente na tua resposta. Ele já sabe o que escreveu.\n` +
    `- BRIEFING MÍNIMO: Sê o mais curto possível. Se um "Obrigado, fico a aguardar" resolve, usa apenas isso.\n`
    : `- CONTEXTO DE REENVIO: Como estás a reencaminhar para um terceiro, deves resumir os factos principais do email original para que o destinatário perceba o contexto.\n`;

  const baseRules = languageEnforcement +
    `REGRAS CRÍTICAS (ESTILO PEDRO):\n` +
    pragmatismRules +
    `- PROIBIDO: "Aqui está a sua resposta", "Espero que este email...", "Como assistente de IA...", ou introduções vazias.\n` +
    `- COMEÇA IMEDIATAMENTE com o corpo do email.\n` +
    `- NUNCA inventes factos. Se faltar informação, faz perguntas curtas.\n` +
    `- Devolve HTML simples: <p>, <br>, <ul>, <li>, <strong>, <em>, <a>.\n`;

  // Inject user knowledge if present
  const knowledgeBlock = knowledge.length > 0
    ? `\nREGRAS/CONHECIMENTO EXTRA DO UTILIZADOR:\n${knowledge.map(k => `- ${k}`).join('\n')}\n`
    : "";

  // Inject file content if present
  const filesBlock = filesContext
    ? `\nCONTEÚDO DOS FICHEIROS ANEXOS (Para tua análise):\n${filesContext}\n`
    : "";

  // Inject Persona / User Style mimic (The "Pedro" Standard)
  const learnedStyleLine = (persona.learnedProfile && typeof persona.learnedProfile === 'string')
    ? `\nESTILO APRENDIDO (HISTÓRICO):\n${persona.learnedProfile}\n`
    : "";

  const learnedHabitsLine = (persona.learnedHabits && typeof persona.learnedHabits === 'string')
    ? `\nHÁBITOS IDENTIFICADOS:\n${persona.learnedHabits}\n`
    : "";

  const personaBlock = `
ESTÁS A ATUAR COMO: Pedro, um profissional altamente pragmático, estruturado e orientado a resultados.
PERFIL DE COMUNICAÇÃO:
- Direto, claro e profissional, mas humano (sem formalismo exagerado).
- Foco absoluto em precisão, completude e utilidade prática.
- Evita floreados, generalidades ("espero que este email...") e "linguagem de IA" artificial.
- Mantém o tom cordial e confiante.
- Escreve sempre no idioma solicitado (${lang}), respeitando as normas gramaticais e de negócio locais.${learnedStyleLine}${learnedHabitsLine}
`;

  const briefingBlock = briefing
    ? `\nCONTEXTO DO THREAD (30-Second Briefing):\n${briefing}\n`
    : "";

  // Inject Contact Aliases (Forward shortcuts)
  const contactBlock = contactAliases.length > 0
    ? `\nTABELA DE ATALHOS DE CONTACTOS (Resolve estes nomes para os respetivos emails se o utilizador os mencionar):\n${contactAliases.map(c => `- ${c.name}: ${c.email}`).join('\n')}\n`
    : "";

  const finalRules = baseRules + knowledgeBlock + filesBlock + personaBlock + briefingBlock + contactBlock;
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
      `\n\nTAREFA: Resume o email (e os anexos) com foco executivo.\n` +
      `Estrutura obrigatória:\n` +
      `<p><strong>Resumo</strong></p><ul>... (máximo 6 pontos claros)</ul>\n` +
      `<p><strong>Próximos passos / Ações</strong></p><ul>... (objetivos e acionáveis)</ul>\n` +
      `<p><strong>Perguntas / Riscos</strong></p><ul>... (se aplicável)</ul>` +
      userInstruction +
      emailBlock
    );
  }

  if (action === "reply") {
    return (
      finalRules +
      toneLine +
      `\n\nTAREFA: Cria uma resposta profissional sugerida ao email.\n` +
      `ESTRUTURA (Pragmatismo Pedro):\n` +
      `1. Agradecimento ou confirmação de receção (curto).\n` +
      `2. Próximo passo ou decisão (se necessário).\n` +
      `3. Fecho cordial.\n\n` +
      `- NÃO repitas o que o remetente acabou de dizer.\n` +
      `- NÃO uses listas se o texto couber num parágrafo curto.\n` +
      `- Garante que é uma resposta "pronta a enviar".` +
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

  if (action === "intent_proposals") {
    return (
      finalRules +
      `\n\nTAREFA: Analisa o email e propõe 3 intenções de resposta curta (Smart Replies).\n` +
      `As propostas devem ser dinâmicas e contextuais (ex: se for convite -> Aceitar, Recusar, Propor nova hora).\n` +
      `REGRAS:\n` +
      `- Devolve APENAS as 3 frases curtas separadas por ponto e vírgula.\n` +
      `- Máximo 4 palavras por frase.\n` +
      `- Exemplo de output: Aceitar convite;Recusar educadamente;Pedir mais informações\n` +
      emailBlock
    );
  }

  if (action === "summarize_actions") {
    return (
      finalRules +
      `Tom: informativo.\n` +
      `TAREFA: Analisa o email e os anexos e devolve APENAS um JSON válido.\n` +
      `REGRAS:\n` +
      `- "summary": lista de exatamente 3 pontos chave (curtos e diretos). Sem saudações.\n` +
      `- "actions": lista de tarefas concretas (ex: "Enviar proposta").\n` +
      `Exemplo de output: {"summary": ["Ponto 1", "Ponto 2", "Ponto 3"], "actions": ["Ligar cliente"]}\n` +
      emailBlock
    );
  }

  if (action === "extract_contacts") {
    return (
      finalRules +
      `TAREFA: Analisa o corpo do email e extrai TODOS os endereços de email mencionados que pareçam ser potenciais destinatários ou pessoas a contactar.\n` +
      `REGRAS:\n` +
      `- Devolve APENAS os emails separados por ponto e vírgula.\n` +
      `- Não incluas o remetente original se for óbvio.\n` +
      `- Se não houver emails, devolve uma string vazia.\n` +
      `- Exemplo: joao@exemplo.com;maria@empresa.pt\n` +
      emailBlock
    );
  }

  if (action === "extract_tasks_json") {
    return (
      finalRules +
      `TAREFA: Analisa o email e os anexos e extrai apenas tarefas/ações CONCRETAS e PENDENTES.\n` +
      `REGRAS CRÍTICAS:\n` +
      `- DEVOLVE APENAS UM ARRAY JSON VÁLIDO.\n` +
      `- Ignora informação geral ou factos (ex: "O meu NIF é X") - foca em AÇÕES (ex: "Enviar fatura").\n` +
      `- Cada objeto deve ter: "title" (descrição curta da ação), "dueDate" (YYYY-MM-DD), "owner" (quem deve fazer).\n` +
      `- Se não houver ações claras para o utilizador ou destinatário, devolve [].\n` +
      `- NÃO incluas texto fora do JSON.\n` +
      `- Exemplo: [{"title": "Ligar ao transitário", "dueDate": "2024-03-01", "owner": "Pedro"}]\n` +
      emailBlock
    );
  }

  // final enforcement
  const finalPrompt = finalRules + toneLine + emailBlock;
  return finalPrompt + `\n\nLEMBRETE FINAL: RESPONDER EM ${lang.toUpperCase()}.`;
}
