# HANDOFF

## HOTFIX IA/OUTLOOK: contexto temporario a partir do corpo do rascunho nativo (Maio 2026)
- **Causa raiz**:
  - em `mode=compose`, quando o Outlook abre um rascunho nativo de `Responder`/`Reencaminhar`, o add-in pode receber `ctx.conversationId`, `ctx.itemId` e `ctx.internetMessageId` vazios;
  - nesse estado, a app nao devia procurar contexto antigo/errado nem depender cegamente do ultimo email lido;
  - alem disso, `setAiState(...)` nao persistia alteracoes sem `conversationId`, impedindo a escolha de modo IA nesse contexto temporario.
- **Solucao implementada**:
  - `office.ts` passou a expor leitura segura do corpo do compose via `Office.context.mailbox.item.body.getAsync(...)` em texto e HTML;
  - `CockpitProvider` cria um contexto temporario `compose-draft-body-context` quando o item ativo e compose sem identidade legivel;
  - nesse contexto, a IA usa o subject/destinatarios disponiveis e o corpo do proprio rascunho aberto, sem registar email na BD, sem sincronizar categorias e sem consultar contexto relacionado;
  - `setAiState(...)` aplica alteracoes em memoria quando nao existe `conversationId`, sem gravar cache persistente.
- **Ficheiros alterados**:
  - `client/src/office.ts`
  - `client/src/components/shell/CockpitProvider.tsx`
  - `docs/HANDOFF.md`
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx src/components/shell/CockpitProvider.tsx src/office.ts` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Riscos remanescentes**:
  - a leitura do corpo em compose depende do host Outlook disponibilizar `item.body.getAsync`; se o corpo ainda nao estiver pronto, a UI mostra aviso e deve recuperar quando o Outlook disponibilizar o item/corpo;
  - validacao completa exige teste no Outlook real/Render com rascunho nativo aberto.
- **Fora de scope confirmado**:
  - sem alteracoes em backend IA, anexos, Para/CC/Bcc, Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook fora da pausa de sync em compose sem identidade, manifest, permissoes ou package scripts.

## DIAGNOSTICO IA/OUTLOOK: trace temporario da selecao Reply/Forward (Maio 2026)
- **Objetivo**:
  - recolher factos no Outlook real sobre porque o clique em `Forward` ainda nao fica ativo/respeitado no `AiCockpit` durante compose/reencaminhamento nativo.
- **Hipotese a validar**:
  - no momento do clique, o contexto exposto ao `AiCockpit` pode estar sem `ctx.conversationId`;
  - nesse caso, o `CockpitProvider.setAiState(...)` ignora o update e a acao nao passa para `forward`.
- **Diagnostico adicionado**:
  - prefixo unico de console: `[AI_ACTION_DIAG]`;
  - `AiCockpit` regista o clique em `Reply`/`Forward`, estado antes do clique, contexto, `emailKey`, prompt/output e tamanho do historico;
  - `AiCockpit` agenda um log curto pos-clique para ver se `aiState.action`/`selectedAction` mudaram;
  - `CockpitProvider` regista updates de `setAiState` com `update.action`, se vai aplicar ou ignorar, e a chave de cache usada;
  - `CockpitProvider` regista a preservacao do email ancora quando entra em `compose-without-readable-message-identity`.
- **Instrucoes de teste pos-deploy**:
  - abrir DevTools/Console, abrir email, esperar barra verde, abrir reencaminhamento nativo do Outlook, clicar `Forward` na app;
  - copiar todos os logs com `[AI_ACTION_DIAG]`;
  - verificar especialmente `ctxConversationId`, `willApply`, `selectedAction`, `aiStateAction` e se aparece `setAiState ignored: missing conversationId`.
- **Ficheiros alterados**:
  - `client/src/modules/ai/AiCockpit.tsx`
  - `client/src/components/shell/CockpitProvider.tsx`
  - `docs/HANDOFF.md`
- **Confirmacao**:
  - esta ronda nao implementa correcao funcional nem altera o fluxo pretendido; adiciona apenas logs temporarios, seguros e removiveis.

## HOTFIX IA/OUTLOOK: Reply e Forward totalmente controlados pelo utilizador (Maio 2026)
- **Causa raiz**:
  - a ronda anterior ainda deixava `composeIntent` influenciar `selectedAction` no `AiCockpit`;
  - mesmo como sugestao inicial, isso continuava a criar estados em que a app podia voltar a `Reply`/`Forward` sem decisao explicita do utilizador.
- **Solucao implementada**:
  - removida do `AiCockpit` toda a logica que aplicava `ctx.composeIntent` como `action`;
  - removido o efeito que fazia `setAiState({ action: ctx.composeIntent })`;
  - removida a logica de `defaultAction` baseada em `composeIntent`;
  - os botoes `Reply` e `Forward` continuam sempre clicaveis e a acao passa a ser apenas a que existe no estado IA ou a escolhida manualmente pelo utilizador.
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Fora de scope confirmado**:
  - sem alteracoes em IA backend, anexos, Para/CC/Bcc, Odoo, Grupos, categorias Outlook, manifest, permissoes ou package scripts.

## HOTFIX IA/OUTLOOK: permitir override manual do modo Forward (Maio 2026)
- **Causa raiz**:
  - `composeIntent` estava a ser tratado como ordem permanente no `AiCockpit`;
  - quando o Outlook sinalizava `composeIntent = reply`, clicar manualmente em `Forward` atualizava a acao, mas o `useEffect` seguinte voltava a aplicar `reply`.
- **Solucao implementada**:
  - o `AiCockpit` passa a manter `generationActionTouched` e uma chave de aplicacao unica por contexto/intent;
  - cliques manuais em `Reply` ou `Forward` marcam override manual e passam a prevalecer;
  - `composeIntent` continua a poder sugerir o modo inicial quando o compose nativo entra sem identidade, mas so enquanto nao houver prompt/output/history e enquanto o utilizador nao tiver tocado no modo;
  - `composeIntent = unknown` nao forca nenhuma acao.
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Fora de scope confirmado**:
  - sem alteracoes em IA backend, anexos, Para/CC/Bcc, Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes ou package scripts.

## HOTFIX IA/OUTLOOK: alinhar compose nativo Forward com modo FORWARD (Maio 2026)
- **Causa raiz**:
  - a hotfix anterior preservava corretamente o email ancora quando o Outlook abria um compose nativo sem `itemId`/`internetMessageId`;
  - esse estado ficava visualmente laranja, mas o contexto preservado perdia a intencao do compose (`Reply` vs `Forward`);
  - como o `AiCockpit` continuava a ver apenas o email ancora e o estado IA anterior, podia manter `REPLY` selecionado e gerar/inserir uma resposta quando o utilizador estava numa janela nativa de reencaminhamento.
- **Solucao implementada**:
  - `office.ts` passa a inferir `composeIntent` em compose por prefixos de assunto conhecidos (`FW`, `FWD`, `ENC`, `RV`, `TR`, `WG` para forward; `RE`, `RES`, `RESP`, `AW`, `SV` para reply);
  - `CockpitProvider` preserva o email ancora confirmado, mas junta o `composeIntent` e o estado controlado `compose-without-readable-message-identity` ao contexto exposto;
  - `AiCockpit` usa esse metadado para selecionar automaticamente `FORWARD` ou `REPLY` quando o Outlook esta num compose nativo sem identidade legivel;
  - compose sem intencao clara fica como `unknown` e nao força `FORWARD`.
- **Impacto esperado**:
  - ao abrir `Reencaminhar` nativo depois de uma leitura verde, a linha laranja passa a estar alinhada com o estado funcional e `FORWARD` fica selecionado;
  - gerar email nesse estado usa o fluxo `forward`/body-only ja existente;
  - `Reply` nativo continua a alinhar para `REPLY`.
- **Ficheiros alterados**:
  - `client/src/office.ts`
  - `client/src/components/shell/CockpitProvider.tsx`
  - `client/src/modules/ai/AiCockpit.tsx`
  - `docs/HANDOFF.md`
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx src/components/shell/CockpitProvider.tsx src/office.ts` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Riscos remanescentes**:
  - a inferencia depende do prefixo de assunto que o host Outlook aplica ao compose nativo; hosts/idiomas com prefixos diferentes podem ficar como `unknown`;
  - validacao final de Reply/Forward precisa de teste manual no Outlook real.
- **Fora de scope confirmado**:
  - sem alteracoes em IA/prompt backend, destinatarios Para/CC/Bcc, dropdowns, anexos, Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes ou package scripts.

## HOTFIX Outlook: preservar email ancora ao abrir compose nativo sem identidade (Maio 2026)
- **Causa raiz**:
  - quando o utilizador abria `Responder`/`Reencaminhar` nativo do Outlook, o add-in podia receber um `mailbox.item` existente mas em modo compose;
  - esse item podia nao ter `itemId` nem `internetMessageId`, mas ainda assim era tratado como contexto novo/incompleto;
  - com isso, o Provider podia substituir o email lido com sucesso por dados parciais do rascunho e a barra de leitura passava a vermelho.
- **Solucao implementada**:
  - `office.ts` passa a devolver `itemUnavailableReason: "compose-without-readable-message-identity"` quando o item ativo esta em compose sem identidade de mensagem legivel;
  - `waitForStableSelectedMessageContext(...)` devolve imediatamente esse estado controlado, em vez de o deixar cair como selecao instavel comum;
  - `CockpitProvider` passa a manter um `lastConfirmedEmailAnchorRef`, atualizado apenas quando a leitura real do email terminou em verde;
  - quando entra em compose sem identidade, o Provider preserva esse email ancora confirmado e nao limpa `ctx`, corpo, HTML ou anexos.
- **Impacto esperado**:
  - ao abrir resposta/reencaminhamento nativo depois de uma leitura verde, a app mantem o email original em memoria;
  - a barra deixa de passar a erro fatal por causa do rascunho transitorio;
  - syncs indevidos continuam bloqueados por `outlookItemUnavailableReasonRef`.
- **Ficheiros alterados**:
  - `client/src/office.ts`
  - `client/src/components/shell/CockpitProvider.tsx`
  - `docs/HANDOFF.md`
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/components/shell/CockpitProvider.tsx src/office.ts` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Riscos remanescentes**:
  - validacao real depende do host Outlook, porque o estado compose/identity varia entre Outlook desktop, web e novo Outlook;
  - se o utilizador arrancar a app ja dentro de um compose sem email ancora previo, a app continua a pedir para voltar ao email em leitura.
- **Fora de scope confirmado**:
  - sem alteracoes em IA, geracao de texto, Para/CC/Bcc, dropdowns, anexos, Odoo, Grupos, `Preparar`, `Classificar`, manifest, permissoes ou package scripts.

## HOTFIX IA/DRAFTS: anexos sem duplicados, dropdown visivel e attach com espera (Maio 2026)
- **Causa raiz da duplicacao**:
  - o mesmo anexo podia chegar por `persistedEmailAttachments` e `liveAttachments` com chaves diferentes;
  - a deduplicacao anterior ainda dependia demasiado da chave de origem, permitindo duas opcoes para o mesmo ficheiro.
- **Causa raiz da dropdown escondida/cortada**:
  - o campo `Anexos` reutilizava a dropdown absoluta dos campos de contactos dentro do card `Detalhes do Rascunho`;
  - com chips removidos e em sidebar estreita, o overlay podia ficar invisivel ou cortado pelo card.
- **Causa raiz provavel da falha de anexacao**:
  - a app chamava `addBase64AttachmentToCompose(...)` imediatamente apos `displayNewMessageForm(...)` / `displayReplyForm(...)`;
  - em alguns hosts Outlook, o compose ainda nao disponibilizou `addFileAttachmentFromBase64Async` nesse instante.
- **Solucao implementada**:
  - anexos passam a ter deduplicacao canonica por `contentId`, `id/attachmentId`, `nome+tamanho`, `nome+tipo` e nome;
  - se o mesmo anexo existir em persisted/live, a UI mostra uma unica opcao e prefere a versao com conteudo, dando prioridade a `liveAttachments` quando aplicavel;
  - a dropdown de `Anexos` passa a ser bloco inline dentro do card, sem alterar dropdowns de contactos;
  - a anexacao espera o compose ficar pronto e tenta cada anexo com retry antes de reportar falha.
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Riscos remanescentes**:
  - a anexacao automatica continua dependente do host Outlook e da disponibilidade de base64/conteudo para cada anexo;
  - se o host nunca disponibilizar compose com API de anexacao, a app avisa e o utilizador pode ter de anexar manualmente.
- **Fora de scope confirmado**:
  - sem alteracoes em Para/CC/Bcc, touched state de destinatarios ou dropdowns de contactos;
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes, package scripts ou backend;
  - sem alteracoes na regra `FORWARD body-only`.

## HOTFIX IA/DRAFTS: selecao de anexos do rascunho isolada e estavel (Maio 2026)
- **Causa raiz identificada**:
  - a selecao de anexos a incluir no rascunho estava acoplada a `fileUsage.forward`, estado tambem usado pelos controlos de ficheiros/anexos;
  - no modo `FORWARD`, a auto-selecao por defeito podia voltar a marcar anexos depois de o utilizador os desmarcar;
  - a resolucao para anexacao dependia demasiado do nome do ficheiro, em vez de usar identidade estavel por `key`, `id`, `contentId` ou `nome+tamanho`.
- **Solucao implementada**:
  - `AiCockpit` passa a manter estado proprio para anexos a incluir no rascunho final;
  - foi adicionado controlo de `draftAttachmentsTouched` para impedir que os defaults do `FORWARD` sejam reaplicados apos intervencao do utilizador;
  - a resolucao de anexos selecionados passa a deduplicar e resolver por chave estavel, com fallback por nome/tamanho;
  - o campo `Anexos` continua dentro de `Detalhes do Rascunho`, com chips e dropdown compacto.
- **Regras por modo**:
  - em `REPLY`, anexos aparecem disponiveis mas nao ficam selecionados por defeito; so os selecionados sao tentados no compose de resposta;
  - em `FORWARD`, anexos reais nao-inline do email aberto ficam selecionados por defeito; se o utilizador desmarcar, a app respeita a selecao e usa o caminho controlado quando necessario.
- **Separacao Analisar vs Reenviar/Incluir**:
  - `Analisar` continua a controlar apenas conteudo enviado a IA;
  - `Incluir/Reenviar` controla apenas anexacao ao rascunho final;
  - selecionar anexos para incluir nao ativa `Analisar` nem envia conteudo/base64 para IA.
- **Ficheiros alterados**:
  - `client/src/modules/ai/AiCockpit.tsx`
  - `docs/HANDOFF.md`
- **Validacoes desta ronda**:
  - `npm.cmd -w client run build` passou; manteve avisos antigos do Vite sobre imports dinamicos/chunk grande;
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx` passou sem erros; manteve warnings antigos fora de scope;
  - `git diff --check` passou.
- **Riscos remanescentes**:
  - a anexacao automatica continua dependente do suporte do host Outlook/WebView;
  - anexos sem conteudo/base64 disponivel podem exigir anexacao manual pelo utilizador.
- **Fora de scope / nao alterado**:
  - sem alteracoes em Para/CC/Bcc, touched state de destinatarios ou dropdowns de contactos;
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes, package scripts ou backend;
  - sem alteracoes na regra `FORWARD body-only`: IA escreve o corpo; a app trata destinatarios, assunto, anexos e Outlook.

## HOTFIX IA/OUTLOOK: dropdowns de destinatarios e item vazio controlado (Maio 2026)
- **Causa raiz da lista de sugestoes fora dos campos**:
  - a caixa `Detalhes do Rascunho` tinha uma lista separada de sugestoes por baixo de `Para/Cc/Bcc`, o que duplicava o fluxo e parecia uma area independente
  - contactos da thread e contactos guardados podiam surgir juntos fora do campo ativo, sem contexto claro sobre a origem da sugestao
- **Causa raiz do erro `displayForwardForm` indisponivel**:
  - o caminho `FORWARD` assumia que `item.displayForwardForm` existia quando havia anexos a reenviar
  - alguns hosts/itens Outlook nao disponibilizam esse metodo, deixando apenas erro no console e sem feedback util na UI
- **Causa raiz do loop `mailbox.item is empty`**:
  - ao abrir o reencaminhamento nativo do Outlook, o add-in podia passar por um estado compose/empty transitorio
  - esse estado era tratado como contexto fatal vazio, limpando o email valido anterior e mantendo tentativas/syncs indevidos
- **Solucao implementada**:
  - `Para`, `CC` e `Bcc` agora usam campos multi-email compactos com chips e dropdown por campo
  - a dropdown separa `Thread atual` de `Contactos guardados`; contactos guardados aparecem quando ha pesquisa
  - a lista externa de sugestoes foi removida da UI
  - `FORWARD` tenta o reencaminhamento nativo quando aplicavel, mas faz fallback para nova mensagem com `Para/Cc/Bcc`, assunto, corpo HTML e anexos best-effort
  - `office.ts` sinaliza `itemUnavailableReason` e limita logs repetidos de `mailbox.item is empty`
  - `CockpitProvider` preserva o ultimo email valido em memoria e sai cedo de leituras/syncs quando o Outlook esta em compose/empty
- **Ficheiros alterados**:
  - `client/src/modules/ai/AiCockpit.tsx`
  - `client/src/components/shell/CockpitProvider.tsx`
  - `client/src/office.ts`
  - `docs/HANDOFF.md`
- **Validacoes realizadas**:
  - `npm.cmd -w client run build` passou
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx src/components/shell/CockpitProvider.tsx src/office.ts` passou sem erros; ficaram warnings antigos fora de scope
  - `git diff --check` passou
- **Riscos remanescentes**:
  - comportamento real de `displayForwardForm`, Bcc e anexacao automatica depende do host Outlook/WebView
  - o estado compose/empty precisa de confirmacao manual no Outlook depois do deploy Render
- **Fora de scope / nao alterado**:
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, manifest, permissoes ou package scripts
  - sem alteracoes funcionais na geracao IA body-only do `FORWARD`

## HOTFIX IA/FORWARD: IA gera so corpo; app gere rascunho Outlook (Maio 2026)
- **Causa raiz conceptual**:
  - o modo `FORWARD` ainda misturava a geracao do corpo do email com destinatarios/metadados do rascunho
  - o prompt podia conter `Para/Cc/Bcc`, email-alvo, destinatarios finais, aliases/contactos ou contexto consolidado que fazia a IA pensar que precisava resolver o envio completo
  - isso aumentava a probabilidade de recusa generica em pedidos simples como `reenviar pedido de cotacao para Giuseppe`
- **Nova regra funcional**:
  - a IA escreve apenas o corpo HTML do email
  - a app trata de `Para`, `Cc`, `Bcc`, assunto, sugestoes de contactos, anexos e criacao/insercao no Outlook
  - `FORWARD` significa criar um email novo baseado no email aberto, sem exigir destinatario real resolvido
- **Regra de idioma**:
  - idioma fixo continua a mandar integralmente no output
  - em `AUTO`, o idioma predominante do email base aberto deve mandar sobre a instrucao curta do utilizador
  - o prompt reforca para ignorar assinaturas, disclaimers, historico citado e mensagens reenviadas antigas sempre que possivel
- **Regra de anexos**:
  - no `FORWARD`, anexos do email aberto ficam marcados para `Reenviar` por defeito
  - `Analisar` e separado de `Reenviar`: conteudo/base64 so vai para IA quando marcado para analisar
  - nomes dos anexos marcados para reenvio podem ir para a IA apenas para permitir texto como `segue em anexo`
  - anexacao real ao rascunho continua responsabilidade da app/Outlook
- **Ficheiros alterados**:
  - `client/src/modules/ai/AiCockpit.tsx`
  - `server/src/routes/aiRoutes.js`
  - `server/src/ai/promptTemplates.js`
  - `docs/HANDOFF.md`
- **Validacoes realizadas**:
  - `npm.cmd -w client run build` passou
  - `node --check server/src/routes/aiRoutes.js` passou
  - `node --check server/src/ai/promptTemplates.js` passou
  - `npx.cmd eslint src/modules/ai/AiCockpit.tsx` passou sem erros; ficaram warnings antigos fora de scope
- **Riscos remanescentes**:
  - teste ponta-a-ponta continua dependente do Outlook/Render publicado e do provider IA real
  - anexacao automatica continua best-effort conforme suporte do host Outlook/WebView
- **Fora de scope / nao alterado**:
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes ou package scripts

## HOTFIX IA: cache com TTL, historico visivel e destinatarios de rascunho controlados (Maio 2026)
- **Causa raiz da cache IA**:
  - `iccc_ai_cache_v1` / `icc_ai_cache_v1` guardavam estados por conversa sem `updatedAtMs`, `cachedAtMs` ou `expiresAt`
  - a limpeza limitava tamanho e numero de conversas, mas nao removia por idade real
  - quando o `localStorage` entrava em `QuotaExceededError`, a chave podia ser removida mas o estado em memoria continuava a tentar persistir, repetindo o aviso em emails seguintes
- **Causa raiz do historico vazio**:
  - `icc.ai_history.v1` era guardado newest-first, mas a UI usava `history.slice(1)`, escondendo a unica entrada existente
  - `saveHistory()` usava `slice(-100)`, que podia preservar entradas antigas e cortar as mais recentes
  - nao existia teto real de tamanho JSON para o historico local
- **Causa raiz da caixa de rascunho que desaparecia**:
  - `showDraftPreview` controlava ao mesmo tempo a existencia do painel e o estado expandido/recolhido
  - ao fechar, o painel era desmontado e deixava de haver cabecalho para reabrir
- **Nova regra funcional de REPLY/FORWARD**:
  - `REPLY` passa a usar por defeito o remetente original como `Para`, e deixa de usar `toRecipients` do email original como destino principal
  - `FORWARD` passa a significar "novo email baseado no email aberto": nao exige email-alvo, nao inventa destinatarios e deixa `Para/Cc/Bcc` vazios salvo selecao/edicao do utilizador
  - `Para`, `Cc` e `Bcc` aceitam edicao manual com `;` ou `,`, normalizam espacos, removem duplicados e mostram sugestoes clicaveis sem colocar todos os contactos automaticamente em copia
- **Ficheiros alterados**:
  - `client/src/components/shell/CockpitProvider.tsx`
  - `client/src/modules/ai/AiCockpit.tsx`
  - `docs/HANDOFF.md`
- **Validacoes realizadas**:
  - `npm.cmd -w client run build` passou
  - `npx.cmd eslint src/components/shell/CockpitProvider.tsx src/modules/ai/AiCockpit.tsx` passou sem erros; ficaram warnings antigos ja existentes/fora de scope
- **Riscos remanescentes**:
  - comportamento final de Bcc depende do suporte do host Outlook/WebView, mas a chamada e best-effort e nao deve quebrar o fluxo
  - a validacao ponta-a-ponta de recipients e limpeza de cache continua a depender de teste no Outlook/Render publicado
- **Fora de scope / nao alterado**:
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest, permissoes, settings globais ou package scripts

## HOTFIX IA: recusa generica deixa de ser mostrada como rascunho em reply/forward (Maio 2026)
- **Causa raiz**:
  - `server/src/routes/aiRoutes.js` ja tinha retry para recusas genericas em `reply`/`forward`, mas o detector nao apanhava respostas como `I'm sorry,\n\nI can't assist with that.`
  - a regex antiga nao normalizava quebras de linha e cobria `can't help` / `cannot help`, mas nao `can't assist` / `cannot assist`
  - quando a recusa escapava ao detector, o endpoint devolvia `ok:true` e o frontend mostrava a recusa como se fosse email gerado
- **Ficheiros alterados**:
  - `server/src/routes/aiRoutes.js`
  - `server/src/ai/promptTemplates.js`
  - `docs/HANDOFF.md`
- **Impacto esperado**:
  - recusas genericas em `reply`/`forward` disparam o retry operacional com instrucoes mais claras de backoffice/comercial
  - se a IA insistir na recusa generica, o endpoint devolve `ok:false` com erro amigavel em vez de entregar `I'm sorry...` como rascunho
  - o prompt de `forward` passa a orientar melhor reencaminhamentos legitimos, anexos/contexto e pedidos de informacao em falta, sem inventar factos comerciais
- **Validacoes realizadas**:
  - `node --check server/src/routes/aiRoutes.js` passou
  - `node --check server/src/ai/promptTemplates.js` passou
  - `npm.cmd -w client run build` passou
  - teste isolado do detector confirmou `I'm sorry,\n\nI can't assist with that.` como recusa generica
  - teste isolado confirmou que `Lamento, nao posso confirmar o prazo sem validacao.` nao e tratado como recusa generica
  - chamada real a `/api/ai/generate` nao executada nesta shell porque `AI_ENABLED`, `OPENAI_API_KEY` e `GEMINI_API_KEY` nao estavam definidos
- **Fora de scope / nao alterado**:
  - sem alteracoes em Odoo, Grupos, `Preparar`, `Classificar`, categorias Outlook, manifest ou settings globais

## IA: cache local deixa de rebentar quota do `localStorage` (Maio 2026)
- **Erro tratado nesta ronda**:
  - o console mostrava `QuotaExceededError` ao gravar o cache IA em `localStorage`
  - `main` ja tinha compactacao do cache atual `iccc_ai_cache_v1`, mas nao lia nem limpava a chave legacy reportada pelo bundle (`icc_ai_cache_v1`)
- **O que foi alterado**:
  - `CockpitProvider` passa a carregar o cache IA pela chave atual e pela chave legacy
  - ao persistir o cache atual, a chave legacy e removida
  - se a quota continuar cheia depois da compactacao, o runtime limpa as chaves de cache IA e segue sem bloquear o add-in
- **Impacto esperado**:
  - a IA pode perder historico local antigo quando o browser estiver sem quota
  - a falha deixa de bloquear o fluxo principal e passa a ser uma limpeza controlada de cache
- **Validacao desta ronda**:
  - `npm.cmd -w client run build` passou
  - `npx.cmd eslint src/components/shell/CockpitProvider.tsx` passou sem erros; mantem warnings antigos do ficheiro

## Hotfix: `Classificar` deixa de cair no bootstrap por import em falta de `resolveClassificationIntermediateCase` (Abril 2026)
- **Causa raiz confirmada**:
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` chamava `resolveClassificationIntermediateCase(...)` no bootstrap do studio sem importar o helper
  - isso abria um `ReferenceError: resolveClassificationIntermediateCase is not defined` no caminho ativo de `Classificar`
- **Correcao feita**:
  - import explicito de `resolveClassificationIntermediateCase` a partir de `groups-v1/storage/resolveClassificationIntermediateCase`
- **Validacao desta ronda**:
  - `build` passou
  - `eslint` do ficheiro tocado passou sem erros
  - `git diff --check` passou
  - smoke browser de `?view=group-classification-studio` passou a carregar o studio sem `ReferenceError` nem `pageerror` desse bootstrap
- **Nota honesta**:
  - no smoke local sem backend dedicado apareceram `500` de recursos/API, mas nao surgiu novo blocker imediato de bootstrap do `Classificar` no cliente
  - esta ronda ficou deliberadamente limitada ao blocker confirmado do import em falta

## Hotfix de estabilizacao da publicacao: aba Groups + quota da cache IA (Abril 2026)
- **Causa raiz do branco na aba Groups**:
  - `GroupsPrepareCockpit` usava `intermediateCaseBinarySources` num `useEffect` antes da propria declaracao do `useMemo`
  - isto criava uma TDZ real em runtime (`Cannot access 'intermediateCaseBinarySources' before initialization`) e deixava a tab totalmente branca ao abrir `Grupos`
- **Correcao aplicada**:
  - `intermediateCaseBinarySources` passa a ser resolvido antes do efeito que persiste `prepareIntermediateCase`
  - a tab `Grupos` volta a abrir sem erro fatal de inicializacao
- **Erro adicional do mesmo release encontrado e corrigido**:
  - `makeEmailKey(...)` em `GroupsPrepareCockpit` devolvia `|||` quando nao havia identidade suficiente do email
  - isso fazia a tab tentar ler o workset `groups_v1_workset:|||` e gerar 404 no arranque sem ancora real
  - a fallback key agora so e construida quando existe pelo menos um segmento util; caso contrario devolve string vazia
- **Causa raiz da quota em `iccc_ai_cache_v1`**:
  - `CockpitProvider` serializava o objeto completo `aiCache` para `localStorage` em cada alteracao, sem limites por conversa, sem poda de historico e sem teto total de payload
  - isto fazia o release rebentar quota quando havia outputs/historicos grandes acumulados
- **Correcao aplicada na cache IA**:
  - leitura passa a normalizar e podar o cache persistido
  - escrita passa a limitar numero de conversas, tamanho de `prompt`/`output`, historico, smart replies, recipients e tamanho total do JSON persistido
  - em quota pressure, a persistencia tenta uma versao compacta antes de remover a chave e falhar de forma controlada
- **Validacao objetiva desta ronda**:
  - reproduzido o branco da aba `Grupos` em browser local antes da correcao
  - confirmada a abertura da tab `Grupos` sem fatal error apos a correcao
  - injectado payload artificialmente grande em `iccc_ai_cache_v1` (~429 KB) e confirmado re-shrink automatico para ~23 KB apos reload, sem erros de consola
- **Fora de scope mantido**:
  - sem novas features
  - sem redesign
  - sem refactor estrutural grande do apply/storage

## Grupos v1: fundacao sem Graph/admin fica finalmente fechada, com URL web explicitamente fora do runtime executavel (Abril 2026)
- **O que ficou fechado nesta ronda**:
  - o perimetro desta frente passa a ser explicitamente **sem Graph e sem permissoes admin**
  - `local_device`, `chosen_folder`, `hybrid`, `local_indexeddb` e `cloud` ficam como base executavel final desta fase
  - `document_library` / URL web de OneDrive/SharePoint deixa de contaminar o runtime como se fosse submodo meio suportado
- **Correcao estrutural em codigo**:
  - `resolveGroupStorageRuntime(...)` passa a expor `projectSupport`
  - quando a configuracao pede `document_library` ou URL web:
    - `projectSupport.supported = false`
    - worksets deixam de ser persistidos como se o modo file-backed estivesse valido
    - o `legacyBridge` deixa de propagar path/provider web para anexos/worksets desta frente
  - `supportsPrimaryGroupWorksetPersistence(...)` deixa de decidir so por `mode`; passa a exigir runtime suportado dentro deste perimetro
  - `validateGroupStorageTarget(...)` no servidor passa a bloquear tanto URL web como `document_library` explicito, mesmo que apareca com path fisico legado
  - `buildGroupWorksetMirrorFileLocation(...)` deixa de aceitar `document_library` como mirror file-backed
- **Settings alinhados com a engine**:
  - o selector de destino deixa de apresentar `document_library` como tipo ativo
  - `chosen_folder` e `hybrid` mostram apenas "Pasta fisica / sincronizada"
  - continua a existir prova tecnica do bloqueio web, mas agora como bloqueio fora de fase, nao como promessa por fechar nesta base
- **Resultado operacional desta frente**:
  - a fundacao de storage/settings pode ser dada como fechada **no perimetro sem Graph/admin**
  - OneDrive/SharePoint por URL web continua explicitamente bloqueado e fora desta fase
- **Fora do scope mantido**:
  - `Explorar`
  - `Gestor do Grupo`
  - Graph
  - permissoes admin
  - redesign geral da UI

## Grupos v1: fundacao de storage/settings passa a usar validacao real de destinos, worksets reativos e manutencao executavel do intermédio (Abril 2026)
- **O que ficou fechado nesta ronda**:
  - `local_device`, `chosen_folder` e `hybrid` deixam de ficar apenas bloqueados por texto e passam a depender de **validacao real** do destino no servidor
  - `Preparar` deixa de ter worksets artificialmente desligados para estes modos; a persistencia principal do workset volta a correr em todos os modos e o servidor tenta espelhar o manifesto para o destino file-backed quando ele e valido
  - a shell de settings da aba `Groups` deixa de ter migracao/limpeza puramente decorativas e passa a executar:
    - migracao real do `IntermediateCase` entre namespaces de `IndexedDB`
    - limpeza real do `IntermediateCase` por regras de retention
- **Backend real fechado nesta ronda**:
  - `POST /api/links/groups/storage/validate`
    - valida de forma real se o path file-backed e acessivel ao processo do servidor
    - escreve/le um ficheiro probe
    - bloqueia explicitamente URL web de OneDrive/SharePoint
  - `POST /api/links/groups/worksets/migrate`
    - migra/regrava o manifesto de workset para o destino alvo
  - `groupWorksetStore` passa a:
    - ler manifestos do store central e de mirror file-backed
    - gravar mirror JSON do workset em destino local validado quando aplicavel
- **Guardas tecnicas confirmadas no repo**:
  - `local_device` nesta arquitetura significa **path local/UNC acessivel ao host do servidor**, nao “o disco do utilizador” por magia
  - picker verdadeiro de pasta continua bloqueado pelo host atual; a alternativa real fechada nesta ronda e `path manual + validacao real`
  - OneDrive/SharePoint por URL web continua bloqueado com prova tecnica:
    - `linkStore` escreve binario via filesystem
    - sem Graph/SharePoint API nao ha escrita real por URL web
- **Politica executavel desta ronda**:
  - intermedio:
    - `IndexedDB` namespaced por `baseFolderPath`
    - migracao e limpeza reais do intermédio na shell da aba
  - final:
    - persistencia classificada continua central em `/api/links/*`
    - file-backed modes passam a ser destinos reais para mirror/binario apenas quando o path e validado
  - sessao/cache:
    - `prepareSession` e fallback em memoria continuam fora da persistencia funcional final
- **Politica de anexos alinhada com o codigo**:
  - metadata sobe sempre
  - binario real so em `cloud` ou path local/sincronizado/UNC validado
  - URL web continua bloqueada
  - `replaceAttachments: false` continua a preservar payload existente em updates parciais
- **O que continua bloqueado por arquitetura/host**:
  - picker nativo que entregue path reutilizavel ao backend
  - OneDrive/SharePoint por URL web
  - migracao historica total dos binarios ja promovidos no `linkStore` sem job backend dedicado
- **Fora do scope mantido**:
  - `Explorar`
  - `Gestor do Grupo`
  - redesign geral da UI
  - backend novo gigante

## Grupos v1: fecho operacional da politica executavel de gravacao alinha shell, storage intermedio e persistencia final real (Abril 2026)
- **O que ficou fechado nesta ronda**:
  - a shell de Settings da aba `Groups` deixa de vender `OneDrive / SharePoint`, migracao, limpeza e `Explorar` como se fossem capacidades prontas nesta fase
  - o storage intermedio desta frente fica explicitamente reduzido a:
    - `IndexedDB` local do add-in quando existe `baseFolderPath` configurado como namespace logico
    - memoria quando o namespace nao existe ou quando o modo esta `disabled`
  - os settings globais de `groupStorage` passam a tratar como modos executaveis apenas:
    - `Cockpit Cloud`
    - `Pasta local / sincronizada`
  - `local_device`, `hybrid` e caminhos web de OneDrive/SharePoint ficam marcados como indisponiveis nesta fase
- **O que o codigo passa a deixar explicito**:
  - `GroupsSettingsPanel.tsx` foi reescrito para separar:
    - intermedio real
    - shell herdada nao executavel
    - fora de scope desta fase
  - `groupsTabSettings.ts` passa a normalizar o modo intermedio canonico como `local_indexeddb | disabled` e a diagnosticar o namespace do `IndexedDB` em vez de fingir uma pasta real
  - `GroupsPrepareCockpit.tsx` passa a mostrar `Namespace` em vez de `Localizacao`
- **Politica executavel desta fase**:
  - intermedio:
    - `IntermediateCase` em `IndexedDB` local do host quando existe namespace
    - fallback em memoria sem namespace
  - final:
    - persistencia classificada via `/api/links/*` e `linkStore`
    - apply continua por email alvo e por scope
  - sessao/cache:
    - `prepareSession`, seeds temporarias e memoria transitiva do add-in
    - nao contam como persistencia final
- **Politica executavel para anexos**:
  - metadata sobe sempre quando o payload final inclui anexos
  - binario real so e tentado quando o provider atual o suporta de verdade:
    - `cloud`
    - caminho local/sincronizado/UNC real
  - URL web de OneDrive/SharePoint continua fora do caminho suportado
  - `replaceAttachments: false` preserva anexos anteriores em payload parcial
- **O que ficou explicitamente fora nesta ronda**:
  - `Explorar`
  - `Gestor do Grupo`
  - migracao real de storage
  - limpeza real do intermedio
  - backend novo grande

## Grupos v1: fecho da estrutura de escrita e armazenamento deixa explicita a fronteira entre intermadio e final (Abril 2026)
- **O que ficou fechado nesta ronda**:
  - o `IntermediateCase` passa a ficar explicitamente fechado como camada **intermedia** de draft, continuidade de sessao, reidratacao controlada e ponte `Preparar -> Classificar`
  - a persistencia **final** desta fase fica explicitamente assumida como a promovida pelo pipeline `/api/links/*` sobre `server/src/linkStore.js`
  - a etapa final do apply deixa de ficar ambiguamente “bem gravada no caso” mas ainda indefinida no store principal
- **O que ja existia e foi confirmado no repo**:
  - `resolveIntermediateCaseStorage(...)` continua a usar `IndexedDB` namespaced por `baseFolderPath` quando o storage intermadio esta `ready`; `missing_location` e `disabled` ficam em memoria
  - `IntermediateCaseRepository` persiste `case.json` e blobs locais do caso
  - `registerRelevantEmail(...)` / `upsertEmail(...)` guardam identidade do email, assunto, remetente, datas, corpo, labels, `removedInheritedLabels`, `labelStates`, `classificationMeta` e anexos
  - `addEmailToLinkGroup(...)` / `removeEmailFromLinkGroup(...)` guardam memberships finais por grupo
  - `createGroupTicket(...)`, `updateGroupTicket(...)` e `linkEmailToGroupTicket(...)` guardam ligacao operacional a tickets
  - `saveGroupDocuments(...)` guarda documentos finais do grupo
- **Risco real encontrado e corrigido**:
  - a ligacao do email ao ticket ainda podia reimpor `membershipKind` unico a todos os `groupIds` finais do ticket, o que abria risco de degradar referencias para principal na persistencia final
  - a correcao passa a fazer o apply final de grupos/referencias explicitamente por `addEmailToLinkGroup(...)` e a usar `linkEmailToGroupTicket(...)` apenas para ligar email + ticket e atualizar `ticket.groupIds`, sem reclassificar memberships
- **Politica final desta fase para anexos**:
  - no `IntermediateCase`, anexos continuam com papel de draft/local (`storageDecision`, `localRef`, `serverRef`, `previewReady`)
  - na persistencia final por email:
    - metadata do anexo fica sempre no store final quando o anexo entra no payload
    - provider `cloud`: o conteudo pode continuar no proprio store final atual
    - provider `local` / `onedrive`: o backend tenta gravar binario real para caminho local/sincronizado; quando consegue, persiste refs (`storageBasePath`, `storagePathHint`) e limpa `content`
    - payload parcial nao limpa anexos antigos sem `replaceAttachments: true`
  - em documentos do grupo, `saveGroupDocuments(...)` segue a mesma separacao: metadata sempre, binario real apenas quando o provider/path suportam escrita segura
- **O que fica intermadio vs final**:
  - intermadio:
    - `IntermediateCase`
    - sessao/seeds de continuidade
    - estados de decisao de anexos ainda nao promovidos
  - final:
    - email classificado e respetiva classificacao por email
    - memberships finais de grupo principal/referencias
    - ligacoes a ticket
    - documentos do grupo
- **O que ainda fica limitado pelo host/contrato atual**:
  - `to` / `cc` ainda nao entram no contrato atual de `RelevantEmailPayload` / `RelatedEmailEntry`, por isso nao ficaram fechados nesta ronda sem abrir API nova
  - URLs web de OneDrive/SharePoint continuam fora do caminho final suportado; o backend so fecha escrita real para pasta sincronizada local / UNC
  - esta ronda nao abriu `Explorar`, `Explorador de Grupos` nem `Gestor do Grupo`; apenas deixou a base final mais pronta para eles
  - nao houve redesign de UI nem nova frente funcional paralela

## Grupos v1: smoke/regression do pipeline de apply confirma a cadeia principal e corrige regressao no link do ticket base quando havia update de estado (Abril 2026)
- **O que foi validado nesta ronda**:
  - **Cenario A**: apply sobre um unico email com grupo principal, referencias, labels e ticket existente sem criacao de ticket novo
    - validado por fluxo/codigo: `resolvedApplySelection -> remoteApplyPlan -> executeLegacyBaseTicketApply(...) -> executeLegacyRemoteApplyForTarget(...)`
    - confirmado que grupos, referencias e labels continuam a ser aplicados por email alvo
  - **Cenario B**: apply sobre um unico email com criacao de ticket novo, projecao local e persistencia/reidratacao
    - validado por fluxo/codigo: `executeLegacyBaseTicketApply(...) -> projectApplyIntoIntermediateCase(...) -> persistAndRefreshClassificationCase(...)`
  - **Cenario C**: apply sobre varios emails alvo no mesmo scope
    - validado por fluxo/codigo que `remoteApplyPlan.targetPlans` e `projectApplyIntoIntermediateCase(... targetEmails)` continuam por email, sem virar apply global cego do caso
  - **Cenario D**: apply com alteracao de anexos/documentState/isHidden
    - validado por fluxo/codigo que `applyClassificationToIntermediateCase(...)` continua a projetar anexos por `emailKey` dono do anexo
  - **Cenario E**: apply com current target incluido e camada Outlook/categorias ativa
    - validado por fluxo/codigo que `beginApplyOutlookCategoryOperation(...)`, `executePostApplyOutlookCategorySync(...)` e o fecho operacional continuam encadeados sem reabrir o handler
  - **Cenario F**: erro/degradacao a meio
    - validado por fluxo/codigo que `finalizeFailedApplyOperation(...)` continua a distinguir erro total vs `Guardado com avisos`, com fecho seguro da operacao Outlook
- **Regressao real encontrada e corrigida**:
  - quando existia `ticket` ja selecionado e havia apenas `update` do estado do ticket, o `base target` podia saltar indevidamente o `linkEmailToGroupTicket(...)`
  - a causa era `skipTicketLink` depender de `finalTicket`, o que confundia `ticket criado` com `ticket existente atualizado`
  - a correcao passou `skipTicketLink` a depender apenas de `ticketExecution.createdTicket` no email base
- **O que nao foi possivel validar com total seguranca**:
  - comportamento runtime real do host Outlook/Office.js e confirmacao de categorias fora do browser local
  - confirmacao ponta-a-ponta com Odoo real e efeitos remotos em ambiente de producao
- **Guardas mantidas**:
  - sem nova frente funcional
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: hardening do pipeline de apply no `Classificar` afina contratos, resultados e tipagem sem abrir nova frente funcional (Abril 2026)
- **O que foi afinado nesta ronda**:
  - `handleApplyClassification()` passa a declarar explicitamente `ApplyOperationResult` como resultado operacional comum
  - `beginApplyOutlookCategoryOperation(...)` e `executePostApplyOutlookCategorySync(...)` passam a expor tipos de resultado nomeados, em vez de contratos inline mais soltos
  - `persistAndRefreshClassificationCase(...)` passa a expor tipos nomeados para sync options e resultado de persist/reidratacao
  - `executeLegacyRemoteApplyForTarget(...)` deixa de aceitar `attachmentStorageOptions` como `Record<string, unknown>` e passa a usar um contrato mais estreito e explicito
  - `outlookCategoryApply.ts` passa a reutilizar o extrator comum de mensagens de erro do fecho operacional e deixa de depender de `any` evitavel nesse ponto
  - o fallback de `selectedEmail` no fluxo Outlook deixa de depender de um check redundante (`selectedEmailKey === selectedEmailKey`) e fica mais claro como guarda de reidratacao do email atual
  - `applyResolution.ts` reduz casts evitaveis em anexos e grupos relacionados, aproximando a projecao de payload dos tipos reais de `RelatedEmailEntry` e `RelevantEmailPayload`
- **Pontos ambiguos corrigidos**:
  - resultados de sucesso/erro/falha degradada ficam mais alinhados entre helpers
  - contratos de persistencia e sync deixam de depender tanto de tipos anonimos inline
  - a camada Outlook usa a mesma normalizacao de mensagem de erro do fecho operacional
- **Divida tecnica pequena que ainda fica**:
  - warnings antigos do `GroupClassificationStudioApp.tsx` continuam fora desta ronda
  - o `finally` minimo do handler continua inline por ainda pertencer ao ciclo de vida do componente
  - nao houve nova extracao estrutural; esta ronda foi apenas de robustez e regressao

## Grupos v1: fecho operacional do apply no `Classificar` sai do handler principal e passa para helper proprio (Abril 2026)
- **O que foi extraido nesta ronda**:
  - normalizacao do resultado final de sucesso do apply
  - normalizacao do resultado final de erro/degradado
  - fecho seguro da operacao Outlook quando uma excecao escapa depois de abrir a operacao
  - o studio passa a delegar esta etapa para `finalizeSuccessfulApplyOperation(...)`, `finalizeFailedApplyOperation(...)` e `closeApplyOutlookOperationSafely(...)`
- **O que ainda fica no handler principal**:
  - mensagens/status intermédios de cada etapa do pipeline
  - orquestracao entre resolucao comum, plano remoto, ticket base, execucao por target, projecao local, persistencia/reidratacao e Outlook
  - `finally` minimo para libertar `actionBusy` e `applyInProgressRef`
- **Acoplamentos reduzidos**:
  - o `catch` deixa de fechar inline a operacao Outlook e de reconstruir manualmente o resultado final do apply
  - o caminho de sucesso/erro fica mais uniforme e legivel no fim do handler
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: camada Outlook/categorias do pos-apply sai do handler principal e passa para helper proprio no `Classificar` (Abril 2026)
- **O que foi extraido nesta ronda**:
  - abertura da operacao Outlook com fase inicial de `saving`
  - construcao do fallback do email atual para projecao Outlook
  - construcao do `source` e `plan` de categorias
  - `enqueue` / `requestCockpitHostAction` / `waitForOutlookCategorySyncResult`
  - `completeOutlookCategoryOperation(...)` nos cenarios de sucesso, timeout e falha
  - o studio passa a delegar esta etapa para `beginApplyOutlookCategoryOperation(...)` e `executePostApplyOutlookCategorySync(...)`
- **O que ainda fica no handler principal**:
  - mensagens/status gerais do apply
  - orquestracao entre plano remoto, ticket base, execucao por target, projecao local e persistencia
  - catch/final fallback para fechar a operacao se alguma excecao escapar antes do helper a completar
- **Acoplamentos reduzidos**:
  - o handler deixa de concentrar inline a maior parte da logica host-specific de Outlook/categorias
  - fica mais claro o corte entre pipeline local de apply e camada Outlook
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: persistencia local e reidratacao pos-apply do `Classificar` saem do handler principal e passam para helper proprio (Abril 2026)
- **O que foi extraido nesta ronda**:
  - `resolveClassificationIntermediateCase(...)`
  - `writeCase(...)`
  - sync imediato do caso no studio
  - refresh/reidratacao pos-apply
  - sync do caso novamente apos o refresh
  - o studio passa a delegar esta etapa para `persistAndRefreshClassificationCase(...)`
- **O que ainda fica no handler principal**:
  - mensagens/status da operacao
  - orquestracao entre ticket base, execucao por target e projecao local
  - sincronizacao Outlook/categorias
- **Acoplamentos reduzidos**:
  - o handler deixa de concentrar inline a persistencia do caso e a reidratacao basica pos-apply
  - fica mais claro o corte entre projecao local, persist/sync/reidratacao e logica de Outlook
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - a ancora continua a ser preservada por `preferredSelectedEmailKey`
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: projecao local no `IntermediateCase` sai do handler principal e passa para helper proprio no `Classificar` (Abril 2026)
- **O que foi extraido nesta ronda**:
  - a construcao do `localClassificationDraft`
  - a resolucao de `localClassificationState`
  - a aplicacao da classificacao local por email no `IntermediateCase`, incluindo a projecao canonica de anexos ja suportada por `applyClassificationToIntermediateCase(...)`
  - o studio passa a delegar esta etapa para `projectApplyIntoIntermediateCase(...)`
- **O que ainda fica no handler principal**:
  - `writeCase(...)`
  - mensagens/status da operacao
  - sync imediato do caso no studio
  - refresh/reidratacao
  - sincronizacao Outlook/categorias
- **Acoplamentos reduzidos**:
  - o handler deixa de montar inline o draft local e a projecao do caso
  - fica mais claro o corte entre plano remoto, ticket base, execucao por target, projecao local e persistencia
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: execucao do ticket base sai do handler principal e passa para helper proprio no `Classificar` (Abril 2026)
- **O que foi extraido nesta ronda**:
  - a decisao/executacao de criar ticket novo e atualizar estado do ticket existente sai do `handleApplyClassification()`
  - o studio passa a delegar essa etapa para `executeLegacyBaseTicketApply(...)`, alimentado por `resolvedApplySelection`, `remoteApplyPlan`, `currentContext` e `currentOutlookTicket`
- **O que ainda fica no handler principal**:
  - mensagens/status da operacao
  - execucao remota por target via `executeLegacyRemoteApplyForTarget(...)`
  - projecao local no `IntermediateCase`
  - refresh/reidratacao
  - sincronizacao Outlook/categorias
- **Acoplamentos reduzidos**:
  - o handler deixa de carregar inline a logica de `createGroupTicket` / `updateGroupTicket`
  - fica mais claro o corte entre plano remoto, ticket base, execucao por target e projecao local
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: execucao remota legacy por email sai do loop inline do `Classificar` e passa para helper proprio (Abril 2026)
- **O que foi extraido nesta ronda**:
  - a execucao remota por email alvo (`removeEmailFromLinkGroup`, `addEmailToLinkGroup`, `unlinkEmailFromGroupTicket`, `registerRelevantEmail`, `linkEmailToGroupTicket`) sai do loop inline do `handleApplyClassification()`
  - o studio passa a delegar essa parte para `executeLegacyRemoteApplyForTarget(...)`, alimentado por `resolvedApplySelection` + `remoteApplyPlan`
- **O que ainda fica no handler principal**:
  - criacao/atualizacao do ticket base (`createGroupTicket` / `updateGroupTicket`)
  - projecao local no `IntermediateCase`
  - refresh/reidratacao do studio
  - sincronizacao Outlook/categorias
- **Acoplamentos reduzidos**:
  - o loop principal deixa de repetir payloads e operacoes remotas por target no meio da orquestracao
  - fica mais claro o corte entre resolucao comum, plano remoto, execucao por target e projecao local
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: execucao remota legacy do apply no `Classificar` fica mais explicita e menos espalhada (Abril 2026)
- **O que foi isolado nesta ronda**:
  - o studio passa a gerar um `remoteApplyPlan` explicito para a execucao remota legacy, com emails alvo, payloads por email, grupos a remover, ticketIds a desligar, base target para ticket e contexto do email atual para categorias Outlook
  - o fallback do email atual usado na projecao Outlook/categorias passa a ser construido por helper proprio, em vez de ficar montado inline dentro do handler
- **Duplicacoes reduzidas**:
  - a execucao remota deixa de recalcular no meio do handler a mesma informacao de targets, payload base e payload classificado
  - a ordem remota fica mais legivel: operacao Outlook -> ticket create/update -> loop por email alvo -> projecao local no caso -> refresh/reidratacao -> categorias Outlook
- **O que continua acoplado nesta ronda**:
  - as chamadas remotas legacy (`createGroupTicket`, `updateGroupTicket`, grupos, tickets, `registerRelevantEmail`) continuam no `GroupClassificationStudioApp`
  - a projecao Outlook/categorias continua no handler porque depende diretamente do host e do ciclo de operacao atual
- **Guardas mantidas**:
  - `resolvedApplySelection` continua a ser a base comum
  - o apply continua por email alvo e por scope
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: pipeline local de apply do `Classificar` fica mais coerente e menos duplicado (Abril 2026)
- **O que foi consolidado nesta ronda**:
  - o studio passa a ter uma resolucao comum de apply (`resolvedApplySelection`) para targets, grupo principal, referencias, labels, ticket, metadados e semantica de scope
  - a construcao de payload por email deixa de ser duplicada em varios pontos do `handleApplyClassification()` e passa a sair de helpers dedicados
  - a projecao no `IntermediateCase` deixa de reconstruir a classificacao local a partir de logica paralela e passa a beber da mesma resolucao comum do apply
- **Duplicacoes reduzidas**:
  - selecao de emails alvo e respetivas chaves
  - payload base por email vs payload classificado por email
  - draft local para `IntermediateCase`
  - criterios de `canApplyClassification`
- **O que ainda continua separado por seguranca**:
  - chamadas remotas legacy (`registerRelevantEmail`, grupos, tickets, Outlook categories) continuam separadas do helper puro de resolucao
  - drafts editoriais ricos (`classificationMetaDraft`, pesquisa, previews, planos locais) continuam locais e nao foram forcados para canonicidade falsa
- **Guardas mantidas**:
  - o apply continua por email alvo e por scope
  - o email ancora continua protegido; a resolucao comum nao transforma o caso num apply global cego
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: bloco de anexos/documentos do `Classificar` fica mais coerente entre estado canonico do caso e estado editorial/local (Abril 2026)
- **O que passa a ficar mais claramente canonico**:
  - a lista base de anexos do email selecionado passa a nascer primeiro do `IntermediateCase` / email canonico, em vez de depender de chaves editoriais demasiado globais
  - quick docs, preview ativo e reidratacao passam a usar chave composta por `emailKey + attachmentKey`, evitando mistura silenciosa entre anexos de emails diferentes
  - `documentState` e `isHidden` do anexo continuam a ser a verdade funcional e passam a mandar mais claramente na reidratacao do preview e dos quick docs
- **O que continua editorial/local**:
  - `showHiddenQuickDocuments`
  - `expandedQuickDocumentKeys`
  - preview aberto/fechado e estado remoto de preview
  - `attachmentPlan` (`analyze` / `save` / `forward`) enquanto plano local de trabalho
- **Reducao de ambiguidade nesta ronda**:
  - o studio separa melhor a colecao canonica de anexos da camada visual filtrada
  - a selecao de preview deixa de depender apenas de `selectedEmail` e passa a conseguir respeitar o par `email + anexo`
  - quick docs e picker de anexos deixam de tratar anexos de emails diferentes como se partilhassem sempre a mesma chave local
- **Guardas mantidas**:
  - alteracoes continuam a ser feitas por email alvo; nao ha alteracao global cega de anexos do caso inteiro
  - preview, expand/collapse e filtros temporarios continuam fora do caso canonico
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: bloco de grupos e referencias do `Classificar` fica mais coerente entre base canonica e pesquisa editorial (Abril 2026)
- **O que passa a ficar mais claramente canonico**:
  - `effectivePrincipalGroupId` e `effectiveReferenceGroupIds` passam a representar a verdade funcional do email quando o caso canonico ja existe
  - o grupo principal e as referencias usados em resumos, chips ativos, favoritos, apply e reidratacao deixam de depender apenas do estado bruto do editor e passam a respeitar primeiro a selecao canonica do email/caso
  - a reidratacao do editor volta a limpar `principalSearch` e `referenceSearch`, para que a pesquisa nao fique a parecer o estado aplicado
- **O que continua editorial/local**:
  - `principalSearch` e `referenceSearch` continuam como texto de pesquisa/criacao
  - resultados temporarios, sugestoes e drafts ricos de `classificationMetaDraft` continuam locais
- **Reducao de ambiguidade nesta ronda**:
  - pesquisar grupo ou referencia deixa de funcionar como espelho implicito do grupo/referencia canonicos
  - `hasPendingClassificationChanges` deixa de contar o texto de pesquisa como alteracao funcional pendente
  - o apply e os resumos passam a usar os ids efetivos do caso, sem reabrir classificacao global cega
- **Guardas mantidas**:
  - o email selecionado/ancora continua protegido pela `preferred key` / ancora do caso
  - o apply por scope continua por email alvo
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: bloco de tickets do `Classificar` fica mais coerente entre ticket canonico e pesquisa editorial (Abril 2026)
- **O que passa a ficar mais claramente canonico**:
  - o ticket real do email passa a ser resolvido primeiro a partir do proprio email/caso (`classificationMeta.ticketId` e fallback dos `ticketIds` do contexto)
  - `selectedTicket` deixa de nascer de uma mistura vaga entre resultados de pesquisa e tickets do caso; primeiro resolve o ticket canonico, e so cai em resultados de pesquisa quando o utilizador esta mesmo a editar o ticket manualmente
  - a reidratacao de email e o pos-apply continuam a limpar `ticketSearch` / `ticketSearchResults`, para o ticket canonico do caso voltar a mandar logo que exista
- **O que continua editorial/local**:
  - `ticketSearch` e `ticketSearchResults` ficam apenas como ferramenta de procura
  - `createTicketTitle` e `ticketStatusDraft` continuam drafts locais de edicao
- **Reducao de duplicacao/ambiguidade nesta ronda**:
  - `canonicalTicketChoices` passa a representar tickets reais/contextuais
  - `ticketPickerChoices` passa a representar a combinacao usada no picker visual, incluindo pesquisa, sem contaminar a verdade principal
  - sugestoes e preservacao do ticket selecionado deixam de depender da lista que mistura pesquisa com contexto canonico
- **Guardas mantidas**:
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio
  - o apply por scope continua por email alvo, sem virar classificacao global cega do caso

## Grupos v1: reducao de estados paralelos editoriais no `Classificar` (Abril 2026)
- **O que passou a depender menos de estado paralelo**:
  - a mudanca de email selecionado deixa de fazer apenas uma limpeza parcial do editor e passa a reidratar grupo principal, referencias, labels canonicas, `ticketId` canonico e `classificationMetaDraft` derivavel diretamente do `IntermediateCase`
  - a pesquisa de ticket deixa de contar como sinal de alteracao funcional pendente; `ticketSearch` continua a existir, mas fica explicitamente como ferramenta editorial
  - a reidratacao canonica passa tambem a limpar `ticketSearch` / `ticketSearchResults`, para que o ticket canonico do email volte a mandar logo que o caso atualizado esta disponivel
- **Separacao mais clara entre canonico e editorial**:
  - canonico/derivavel do caso: email selecionado, grupo principal, referencias, labels canonicas, `ticketId` canonico e metadados de classificacao por email
  - editorial/transitorio: texto de pesquisa de ticket, resultados de pesquisa, `createTicketTitle`, `ticketStatusDraft` e partes ainda ricas de `classificationMetaDraft`
- **O que ainda continua paralelo nesta ronda**:
  - `ticketSearch` / `ticketSearchResults` continuam locais como UX de apoio
  - `createTicketTitle`, `ticketStatusDraft` e toggles ricos de `classificationMetaDraft` continuam editoriais porque ainda nao representam contrato canonico final por email
- **Guardas mantidas**:
  - a identidade do email selecionado continua a ser reidratada pela `preferred key` / ancora do caso, sem regressar ao padrao do primeiro item da lista
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio

## Grupos v1: reidratacao local pos-apply no `Classificar` passa a voltar explicitamente ao `IntermediateCase` (Abril 2026)
- **O que passou a ser reidratado diretamente do caso apos apply**:
  - email selecionado/ancora preservado por `emailKey`, sem cair no primeiro da lista
  - emails do conjunto canonico voltam a sobrepor os equivalentes legacy em `relatedEmails` / `knownEmails`
  - etiquetas selecionadas, removidas herdadas, estados por etiqueta e categorizadas voltam a semear o editor a partir do caso
  - grupo principal, referencias e `ticketId` canonico do email selecionado voltam a ser refletidos logo apos o apply
  - selecao de emails alvo do apply volta a ser reconciliada com o conjunto do caso atualizado
- **Estados stale reduzidos nesta ronda**:
  - o studio deixa de depender apenas do `refreshSelectedEmailContext()` para refletir o caso atualizado
  - o `IntermediateCase` atualizado passa a ser sincronizado imediatamente apos o `writeCase(...)` e reaplicado novamente depois do refresh legacy, para que a base canonica volte a mandar
  - listas case-backed deixam de ser reinjetadas por merge vago quando o mesmo email ja existe no caso canonico
- **O que ainda fica paralelo/legacy nesta ronda**:
  - `ticketSearch`, `ticketSearchResults` e parte dos toggles ricos de `classificationMetaDraft`
  - tickets/grupos carregados do servidor continuam a enriquecer o studio para nomes, estados e contexto, mas deixam de mandar na identidade do email e na classificacao canonica ja persistida
- **Fora do scope mantido**:
  - sem promocao final nova para servidor
  - sem limpeza real do intermedio
  - sem endpoints novos

## Grupos v1: etiquetas locais do `Classificar` passam a reabrir com fidelidade a partir do `IntermediateCase` (Abril 2026)
- **O que passa a ficar canonicamente guardado por email no caso**:
  - `labels` deixam de representar apenas labels "owned" do apply local e passam a guardar a lista final de etiquetas ativas do email
  - `removedInheritedLabels` passa a guardar as etiquetas herdadas removidas explicitamente
  - `labelStates` passa a guardar o estado por etiqueta ja estabilizado no studio
  - `categorizedLabelNames` passa a guardar as etiquetas marcadas para categorizacao quando essa parte ja esta estavel
- **Como isto e escrito no apply local**:
  - o apply continua a respeitar o scope por email do studio
  - por cada email alvo, o `IntermediateCase` passa a receber a lista final de etiquetas selecionadas, removidas herdadas, estados por etiqueta e etiquetas categorizadas
  - o caso atualizado e regravado no storage intermédio da frente
- **Como o `Classificar` reabre isto**:
  - `hydrateIntermediateCaseEmailsToRelatedEntries(...)` passa a devolver `labels`, `removedInheritedLabels`, `labelStates` e `categorizedLabelNames` diretamente do caso canonico
  - o studio volta a semear `selectedLabels` e a parte canonica de `labelDrafts` a partir destes campos antes de cair em heuristicas legacy
- **O que continua draft local puro nesta micro-ronda**:
  - os toggles mais ricos de `classificationMetaDraft` que ainda nao tem contrato canonico fechado
  - o objeto completo de `labelDrafts` enquanto editor rico; o caso guarda apenas a parte estavel necessaria para reabertura fiel
- **Legado que continua**:
  - fallback legacy continua a existir para cenarios antigos sem `IntermediateCase` ou sem estes campos ainda persistidos
  - quando o caso canonico ja traz a informacao de etiquetas, ele passa a ser a base principal desta parte do studio

## Grupos v1: classificacao local do studio passa a ser projetada por email no `IntermediateCase` (Abril 2026)
- **O que passa a ser escrito no caso canonico**:
  - grupo principal e nome do grupo principal
  - grupos de referencia
  - labels proprias do email
  - `ticketIds` / `ticketCodes` quando ja existem no fluxo atual
  - `status` e `state` locais derivados do apply atual
  - `classifiedAt` e `classifiedSource`
  - `documentState` e `isHidden` dos anexos do email alvo
- **Como o scope e aplicado**:
  - o apply continua a respeitar `current`, `selected` e os outros scopes ja existentes no studio
  - o `IntermediateCase` e atualizado por email alvo, nunca por caso inteiro de forma cega
  - quando ha caso canonico, ele e regravado no storage intermédio e o bootstrap do studio e atualizado a partir desse caso
- **O que ainda fica draft local puro nesta ronda**:
  - `labelDrafts`
  - `classificationMetaDraft` completo
  - `selectedSeriesId`
  - `ticketStatusDraft` enquanto draft de edicao
  - `attachmentPlan` (`analyze` / `save` / `forward`)
- **Legado que continua**:
  - o apply legacy/remoto continua a coexistir para nao partir o fluxo atual
  - o `IntermediateCase` passa a ser a verdade local principal onde ja existe, e o legado fica como compatibilidade/transicao
- **Fora do scope mantido**:
  - sem promocao final nova para servidor
  - sem limpeza real do intermédio
  - sem endpoints novos

## Grupos v1: `Classificar` passa a hidratar-se internamente a partir do `IntermediateCase` (Abril 2026)
- **O que passou a nascer diretamente do caso canonico**:
  - `classificationCase` como base interna explicita do studio
  - email ancora a partir do caso canonico, respeitando `anchorEmailKey`
  - emails do conjunto/contexto do studio a partir dos emails do caso canonico
  - total conhecido do studio a partir do caso canonico, com enriquecimento legacy apenas como complemento
- **Reducao de reconstrucao vaga**:
  - `emailPool`, `caseScopeEmails`, quick documents e listas de contexto deixam de depender primeiro de `relatedEmails` / `knownEmails` quando o caso canonico existe
  - a escolha do email ancora deixa de cair apenas em heuristicas da lista visivel e passa a respeitar a ancora do caso
  - mutacoes locais de anexos/visibilidade e refresh contextual passam a reconciliar tambem o bootstrap canonico quando ele existe
- **O que continua legado nesta ronda**:
  - `seedKey` e `prepareSeedKey` continuam como fallback para cenarios sem `IntermediateCase`
  - leituras de servidor e listas legacy continuam como enriquecimento/fallback para nao partir o fluxo atual
- **Fora do scope mantido**:
  - sem promocao real para servidor
  - sem limpeza real do intermédio
  - sem refactor profundo total do `Classificar`

## Data
- 2026-04-06

## Estado atual resumido
- A app está funcionalmente orientada para Outlook + Odoo + IA + documentos, mas a base técnica ainda está mais madura do lado funcional do que do lado de segurança e coerência arquitetural.
- A arquitetura observada no repositório é uma app unificada com frontend React/Vite em `client`, backend Node/Express em `server`, manifestos Outlook em `manifest` e deploy previsto para Render.
- A prioridade recomendada continua a ser segurança e coerência arquitetural antes de novas features.

## Resumo técnico do que a app é hoje
- Outlook add-in com task pane e comandos read/compose.
- Frontend que lê contexto do email atual, corpo, HTML, anexos, categorias e ações de compose através de Office.js.
- Backend que expõe API própria, integra com Odoo por JSON-RPC, orquestra IA com OpenAI/Gemini e mantém uma camada própria de persistência para links, grupos, tickets, documentos e caches.
- Persistência híbrida: Odoo para entidades de negócio, Postgres opcional ou JSON local para a camada própria, e `localStorage` / Office RoamingSettings para settings e caches do cliente.

## Confirmado pelo repositório
- Existe um monorepo com workspaces `client` e `server`.
- O backend usa Express, `pg`, Odoo por JSON-RPC, providers OpenAI/Gemini e serve também o frontend compilado.
- O frontend usa React, Office.js e vários módulos operacionais: AI, CRM, groups, related, files e settings.
- O manifest declara `ReadWriteMailbox` e requirement set `Mailbox 1.14`, com superfícies de read e compose.
- Há uma camada própria de persistência (`linkStore`) para links, grupos, tickets, documentos e emails relevantes.
- Há sinais reais de dívida técnica: ficheiros muito grandes, coexistência de fluxos legados/duplicados e múltiplas zonas de polling/recarregamento.
- O repositório prova uso opcional de Postgres e compatibilidade explícita com Postgres alojado na Supabase, mas não prova uso de `supabase-js`, Auth, Storage, Realtime ou Edge Functions.

## Provável mas não confirmado em produção
- Se `DATABASE_URL` estiver ativo em produção, a camada própria de dados está a usar Postgres em vez de apenas ficheiro local.
- Se a produção seguir o mesmo padrão do ambiente local analisado, o Postgres poderá estar alojado na Supabase.
- O tráfego repetido do add-in pode estar a contribuir materialmente para custos de DB/egress, mas isso não fica provado sem métricas reais de produção.
- Podem existir diferenças entre o código no repositório e a app atualmente em produção, incluindo variáveis de ambiente, manifest instalado e deploy ativo.

## Audit de Segurança e Fluxo de Dados (Abril 2026)
- **Documento Detalhado**: [docs/audits/security-and-data-flow-audit.md](docs/audits/security-and-data-flow-audit.md)
- **Fronteira de Segurança**: Confirmado CORS aberto (`cors()`) e gestão de sessões apenas em memória.
- **Segredos**: Confirmado que `odooPassword` e API Keys (OpenAI/Gemini) são persistidos no cliente (`settings.ts`). **Prioridade Crítica para correção**.
- **Fontes de Verdade**: Mapeadas entre Odoo (SSOT para negócio), linkStore (Auxiliar/Studio) e RoamingSettings (Settings/Secrets).
- **Riscos Estruturais**: Identificados 5 pontos críticos, incluindo o gigantismo dos ficheiros `linkStore.js` (4.5k) e `GroupClassificationStudioApp.tsx` (6.5k).

## Principais riscos atuais (Auditados)
- **Crítico**: Exposição de credenciais Odoo e chaves AI no armazenamento local do browser/RoamingSettings.
- **Alto**: Backend com CORS aberto e fallback de autenticação global via variáveis de ambiente em `getOdooCached`.
- **Alto**: In-memory sessions causam UX pobre em restarts de servidor e dependência de re-auth automática via segredos no client.
- **Médio**: Verdade arquitetural fragmentada e risco de divergência entre Postgres e Odoo.
- **Médio**: Rigidez e risco de regressão devido a ficheiros monolíticos massivos.

## Prioridades recomendadas (Atualizadas)
1. **Segurança (Vaulting)**: Migrar segredos do cliente para o servidor (vaulting).
2. **Segurança (Perímetro)**: Restringir CORS e endurecer `sessionManager`.
3. **Coerência de Dados**: Reduzir fallback global de Odoo e consolidar durabilidade de links/sessões em DB.
4. **Refactoring**: Sharding de módulos massivos para reduzir risco de manutenção.
5. **Funcional**: Evoluir features apenas após estabilização do perímetro.

## Frentes em aberto
- Segurança do backend e política de autenticação/autorização
- Confirmação da arquitetura real em produção: Render, `DATABASE_URL`, host de Postgres, pooler e manifest ativo
- Consolidação de persistência e clarificação de ownership dos dados
- Revisão de polling, refreshes e reingestão de emails
- Consolidação dos módulos AI/CRM e redução de duplicação
- Observabilidade, testes e processo de code review

## Unificação de Estados e Cores (Abril 2026)
- **Centralização UI**: Criado `client/src/statusUtils.ts` como SSOT para visualização de estados.
- **Matriz de Cores**: Implementada matriz unificada (4 grupos: Azul/Analise, Amarelo/Progresso, Verde/Concluido, Vermelho/Fechado).
- **Abordagem Não Destrutiva**: O backend (`linkStore.js`) foi tornado permissivo para aceitar e manter aliases legados (ex: "Aberto", "Aguarda", "Bloqueado") sem reescrita automática, garantindo unificação visual sem migração de base de dados.
- **Sincronização Outlook**: As categorias do Outlook geradas em `outlookCategories.ts` agora seguem os mesmos labels amigáveis da UI.

## Hotfix Final do Modal "Aplicar a..." (Abril 2026)
- **Densidade dos Cards**: Corrigidas as regras CSS (`S.email` e `S.quickDocList`) para garantir que os cards ocupem apenas uma linha compacta (28-30px de altura). Os itens usam `flex` puro, as legendas cortam com `ellipsis`, e bloqueou-se o `alignContent: "stretch"` para impedir que o flex grid estique os cards verticalmente quando há poucos itens ou o espaço é grande.
- **Ecrã Branco Modo Avançado (React #130)**: A causa raiz do log em produção foi identificada. O componente tentava renderizar `<S.StatusLegendContainer>`, que era um styled-component inexistente na diretoria inline `S`, o que devolvia `undefined`. O `renderOutlookColorLegend` foi reescrito para utilizar nós DOM nativos (`div` e `span`) com estilos CSS corretos, sanando o formidável crash.
- **Projeção Outlook Sync**: O fluxo que gravava categorias no Outlook deixava de executar caso a UI pedisse para processar "Todos os emails do caso". O cálculo de `includesCurrentTarget` (que aprova a adição da gravação Outlook para o item master aberto no painel) foi reescrito. Agora confere corretamente não as seleções da UI, mas se o array `effectiveTargetEmails` engloba o `currentContext.itemId`. A persistência Outlook passa a executar-se corretamente nos 3 scopes: "Só este email", "Emails selecionados", e "Todos os emails do caso".
- **Nota Limitação Outlook Sync**: O sync em tempo real pela Add-in API impõe projeção imediata apenas sobre o email currente aberto no ecra (`context.mailbox.item`). Emails não abertos recebem guardado no Odoo, mas só propagarão para o Exchange Server em flows background extra ou polling.

## Áreas sensíveis onde não convém mexer sem revisão
- `server/src/index.js`
- `server/src/linkStore.js`
- `server/src/odoo.js`
- `server/src/routes/aiRoutes.js`
- `client/src/office.ts`
- `client/src/components/shell/CockpitProvider.tsx`
- `client/src/modules/ai/AiCockpit.tsx`
- `client/src/modules/crm/GroupClassificationStudioApp.tsx` (Módulo Crítico)
- `manifest/`

...
(rest of the handoff)
...

## Refactoring Estrutural da Classificação (Abril 2026 - Ronda 1)
- **Extração de Lógica Apátrida**: O módulo `GroupClassificationStudioApp.tsx` (anteriormente com 6.5k linhas) iniciou o processo de sharding.
- **Novos Módulos CRM**: Criada a diretoria `client/src/modules/crm/group-classification/` contendo:
  - `types.ts`: Definições de interfaces e tipos específicos do domínio de classificação.
  - `constants.ts`: Configurações de UI, opções de status e valores padrão.
  - `documentUtils.ts`: Biblioteca de funções puras para processamento de emails, anexos, referências e metadados.
- **Desacoplamento UI/Lógica**: O componente principal agora importa helpers e constantes destes módulos, reduzindo a duplicação e facilitando a manutenção futura sem alterar o comportamento funcional (functional parity).
- **Legendas Unificadas**: Integrado o `UNIFIED_STATUS_LEGEND` em `statusUtils.ts` para garantir consistência visual no Estúdio.

## Refactoring Estrutural da Classificação (Abril 2026 - Ronda 2)
- **Extração de Componentes UI**: O componente monolítico `GroupClassificationStudioApp.tsx` foi fragmentado com a extração de 3 componentes visuais para `client/src/modules/crm/group-classification/components/`:
  - `EmailsCard.tsx`: Gere a listagem, pesquisa e seleção múltipla de emails.
  - `QuickDocumentsCard.tsx`: Gere a listagem de anexos persistidos e controlos de visibilidade/preview.
  - `StatusLegend.tsx`: Componente partilhado para a legenda de estados e cores do Outlook/Odoo.
- **Consolidação de Tipos**: O ficheiro `types.ts` foi expandido para incluir todos os tipos de domínio (SectionId, ReadingSuggestionChip, StudioParams, TicketEditorMode, etc.), assegurando que o novo módulo é auto-suficiente e tipado corretamente.
- **Integridade Funcional**: A extração seguiu o princípio de paridade funcional estrita. Nenhuma lógica de negócio ou estado global foi alterado; a orquestração permanece no componente pai (Studio App), mas a superfície de markup e estilos locais foi significativamente reduzida.
- **Validação de Build**: Confirmado que o projeto compila (`npm run build`) e passa no check de tipos (`tsc`) para os ficheiros modificados, garantindo que a modularização não quebrou referências internas.

## Refactoring Estrutural da Classificação (Abril 2026 - Ronda 3)
- **Extração de Componentes UI Estruturais**: Continuando a modularização do `GroupClassificationStudioApp.tsx`, extraíram-se dois componentes fundamentais para `client/src/modules/crm/group-classification/components/`:
  - `ClassificationEditor.tsx`: Gere a interface de edição de classificação em todos os seus modos (Grupo Principal, Etiquetas, Ticket e Referências). Inclui o rendering de sugestões, pesquisa, criação e opções avançadas.
  - `ApplyDialog.tsx`: Gere a interface do modal "Aplicar a...", permitindo a escolha do âmbito de aplicação (este email, selecionados, ou todos os emails do caso) com preview de emails alvo.
- **Redução do Monólito**: O componente principal foi aliviado de toda a camada de desenho visual destes fluxos, mantendo-se como o orquestrador de estado e lógica de negócio.
- **Manutenção de Lógica**: Funções de persistência, `handleApplyClassification`, sincronização Outlook e lógica de orquestração de estado permanecem no Studio App para garantir paridade funcional e evitar efeitos secundários não previstos.
- **Validação de Build e Tipos**: Confirmado que o projeto compila (`npm run build`) e que os ficheiros do módulo `group-classification` e o Studio App passam no check de tipos (`tsc`), resolvendo dependências de tipos entre o pai e o filho.

## Refactoring Estrutural da Classificação (Abril 2026 - Ronda 4)
- **Extração de Camada de Preview**: O módulo `GroupClassificationStudioApp.tsx` foi novamente reduzido com a extração da lógica de preview para dois novos ficheiros:
  - `previewUtils.ts`: Contém helpers puros para sanitização de HTML (`sanitizeEmailPreviewHtml`), escape, descodificação base64 e utilitários do Office Web Viewer.
  - `PreviewPane.tsx`: Componente que encapsula o painel inferior de preview, incluindo modos de email, documento, e placeholders de resposta/reencaminhamento.
- **Preview de PDF Local**: O componente `StudioPdfPreview` foi movido para o módulo de preview, sendo agora exportado para ser usado tanto no painel inferior como no preview de anexos do workspace principal.
- **Isolamento de Estilos**: Os estilos CSS específicos do preview foram movidos para o componente `PreviewPane.tsx`, simplificando o objeto de estilos do Studio App.
- **Integridade Funcional**: Mantida a paridade funcional estrita. Toda a orquestração de qual documento ou email mostrar permanece no componente pai, que passa os dados e callbacks para a nova camada de preview.

## Refactoring Estrutural da Classificação (Abril 2026 - Ronda 5)
- **Limpeza do Shell Orchestrator**: Concluída a transição do `GroupClassificationStudioApp.tsx` para uma arquitetura de "shell" pura, removendo aproximadamente 600 linhas de código de rendering legacy.
- **Novos Componentes Extraídos**:
  - `ClassificationEditorHeader.tsx`: Encapsula a lógica de cabeçalho do editor de classificação, com botões de navegação e ação (Aplicar).
  - `ClassificationSummaryTiles.tsx`: Encapsula a visualização resumida dos dados atribuídos (Grupo, Referências, Etiquetas, Ticket) quando o editor está fechado.
- **Remoção de Dead Code**: Eliminadas as funções de rendering inline (`renderPrincipalEditor`, `renderLabelsEditor`, `renderSuggestionTrayLegacy`, etc.) que agora são responsabilidade total do componente `ClassificationEditor`.
- **Integridade de Estado e Fluxos**: Estrita manutenção de todos os estados funcionais (`attachmentPlan`, `manageableGroups`, etc.) e handlers de negócio no shell. O shell atua agora estritamente como orquestrador, passando estado e callbacks para os sub-componentes.
- **Validação de Build**: Confirmada a compilação do projeto e integridade das referências entre o shell e os novos componentes modulares.

## Hotfix: Crash de Lançamento da Ronda 5 (Abril 2026)
- **Causa Raiz**: Após a modularização do Studio (Ronda 5), o array `quickDocumentAttachments` passou a ser fornecido como uma lista plana de anexos (`attachment[]`), enquanto o componente `QuickDocumentsCard.tsx` esperava pares `{ email, attachment }`. A tentativa de aceder a `email.attachments` sem validação resultava em crash de runtime.
- **Correção Aplicada**: 
  - No `GroupClassificationStudioApp.tsx`, a coleção foi refatorada para reconstruir a estrutura de pares `{ email, attachment }` percorrendo todos os emails relacionados.
  - No `QuickDocumentsCard.tsx`, foram adicionados guards defensivos rigorosos e optional chaining para evitar falhas com entradas nulas ou incompletas.
  - Restaurada a compatibilidade do estado `selectedEmailAttachments` para manter fluxos legados funcionais.
- **Ficheiros Alterados**:
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx`
  - `client/src/modules/crm/group-classification/components/QuickDocumentsCard.tsx`
- **Validação Executada**: 
  - Build bem-sucedido (`npm run build`).
  - Playwright Tests (6/6 passing) confirmam que o Studio abre sem erros e os cards carregam corretamente.

## Hotfix: Ocultação de Decorativos e Visibilidade do Modal (Abril 2026)
- **Causa dos Problemas**: 
  1. **Documentos Rápidos**: A lógica de filtragem automática de anexos decorativos (assinaturas, ícones sociais, imagens < 15KB) foi perdida ou desativada durante o sharding do Studio, poluindo a lista de documentos úteis.
  2. **Modal Invisível**: O componente `ApplyDialog.tsx` dependia exclusivamente de variáveis de CSS (`--skin-bg-main`) que, em certos contextos (Outlook sem tema injetado ou browser sem cockpit ativo), resultavam em cores transparentes ou indefinidas, tornando o modal ilegível.
- **Correções Aplicadas**:
  - **Heurística Restaurada**: Atualizada a função `isStudioAttachmentHiddenInQuickDocs` em `documentUtils.ts` para integrar a verificação `isLikelyDecorativeAttachment`. Agora, anexos irrelevantes são escondidos por defeito (mas acessíveis via "Ver silenciados").
  - **Fallbacks de UI**: Adicionados fallbacks de cor e border explícitos (ex: `#ffffff`, `#000000`, `#e5e7eb`) em todos os estilos do `ApplyDialog.tsx`, garantindo legibilidade universal.
- **Ficheiros Alterados**:
  - `client/src/modules/crm/group-classification/documentUtils.ts`
  - `client/src/modules/crm/group-classification/components/ApplyDialog.tsx`
- **Validação**: Build bem-sucedido e verificação de tipos garantida.

## Hotfix: Estabilização do Studio de Classificação (Abril 2026)
- **Deduplicação e Idempotência**: Implementada a assinatura de classificação (`getClassificationSignature`) e o `lastAppliedSignatureRef` para evitar gravações redundantes tanto no Odoo como no Outlook. Se o estado não mudar, o fluxo de aplicação é abortado silenciosamente com sucesso.
- **In-flight Guard**: Adicionado o `applyInProgressRef` para garantir que apenas uma operação de classificação está ativa por vez, bloqueando cliques repetidos no botão "Confirmar".
- **Sincronização Outlook (Hotfix)**: Resolvido o problema de falha na escrita de categorias do Outlook. O fluxo agora garante a reidratação do email final, obtém as definições de armazenamento atualizadas (`getSettings`) e executa o `applyOutlookCategoryPlan` de forma síncrona com o estado do servidor.
- **Tipagem Outlook (Office.js)**: Corrigidos erros de linting em `client/src/office.ts` relacionados com a tipagem `unknown` de anexos, utilizando casting explícito para garantir acesso seguro a propriedades e métodos das APIs Microsoft Office.
- **Ficheiros Alterados**:
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` (Reimplementação robusta do `handleApplyClassification`)
  - `client/src/office.ts` (Correção de tipagem de anexos)
  - `client/src/outlookCategories.ts` (Exposição de helpers de assinatura)
- **Validação**: Run `tsc --noEmit` bem-sucedido (sem erros reportados no módulo CRM) e merge limpo para `main`.

## Hotfix: Outlook Categories Host Sync (Abril 2026)
- **Âmbito Delimitado**: Intervenção apenas no pipeline cliente que projeta categorias para o item aberto no host Outlook a partir do Classification Studio. Não houve alterações em backend, settings/quota, Vaulting, modal ou Documentos Rápidos.
- **Causa Raiz Confirmada no Repositório**: O writer do Outlook (`applyOutlookCategoryPlan`) aceitava falhas reais do host como best-effort. As chamadas `masterCategories.addAsync`, `item.categories.addAsync` e `item.categories.removeAsync` faziam log de warning mas resolviam sem erro, e o fluxo podia terminar com `success` mesmo sem o item ficar com as categorias desejadas.
- **Correção Aplicada**:
  - `client/src/office.ts`: o pipeline de escrita passou a devolver resultado estruturado (`success` / `noop` / `failed` / `stale` / `item-mismatch`) em vez de bool cego.
  - `client/src/office.ts`: foram adicionados diffs de categorias geridas, verificação pós-write e retry curto para lidar com latência de propagação das master categories no host Outlook.
  - `client/src/office.ts`: falhas de `addAsync` / `removeAsync` / `masterCategories` deixam de ser engolidas silenciosamente; passam a produzir `detail` concreto quando o estado final não converge.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx`: o Studio agora expõe o detalhe devolvido pelo writer quando o Outlook não confirma a aplicação.
- **Ficheiros Alterados**:
  - `client/src/office.ts`
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx`
- **Validação Executada**:
  - `npm -w client run build` bem-sucedido.
  - `tsc --noEmit -p client/tsconfig.json` continua com erros pré-existentes noutros módulos não tocados (`AiCockpit`, `AiReplyTargetPickerApp`, `CrmCockpit2`, `GroupExplorerApp`, `GroupManagerCockpit`, `GroupsCockpit`, `FileCockpit`, `DialogApp`, `SettingsPanel`), pelo que não serve como validação limpa desta ronda.
- **Validação Ainda em Falta Fora do Repo**:
  - Confirmar no Outlook real (novo Outlook e clássico, se aplicável) que a criação/atribuição de master categories propaga e aparece visualmente no item aberto.
  - Confirmar se os `detail` novos apontam algum erro específico de host/API em produção caso a escrita continue a falhar.

## Hotfix: Outlook Categories Apply Readback (Abril 2026)
- **Âmbito Delimitado**: Ajuste estrito do pipeline pós-guardar para confirmar categorias Outlook no host através de readback antes de marcar sucesso funcional. Não houve mudanças em backend, Odoo, preview, documentos, split estrutural ou UX fora das mensagens já existentes.
- **Causa Encontrada no Repositório**: O readback de categorias aceitava qualquer fonte disponível (`item.categories` array local ou fallback), o que podia confundir estado em memória com confirmação real do host. Além disso, a confirmação de equivalência podia acontecer sem uma leitura explícita via `item.categories.getAsync`.
- **Correção Aplicada**:
  - `client/src/office.ts`: o readback passou a preferir `Office.context.mailbox.item.categories.getAsync`, com fallback explícito e degradado apenas quando a leitura confirmada do host não está disponível.
  - `client/src/office.ts`: o sucesso/no-op/equivalência agora só contam quando o readback vem de `getAsync` e confirma exatamente as categorias geridas pedidas.
  - `client/src/office.ts`: mantido retry curto com delay e readback repetido após `addAsync`/`removeAsync`.
  - `client/src/office.ts`: adicionados logs temporários mínimos (`[TEMP][outlook-category-apply]`) com item id, categorias pedidas, resultado bruto do apply, categorias lidas e motivo de fallback.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx`: adicionados logs temporários de timeout e resultado degradado quando o writer não confirma a tempo.
- **Validação Executada**:
  - `npm -w client run build` bem-sucedido.
  - Continua em falta validação real no staging/Outlook para confirmar que o host devolve `getAsync` consistente após a aplicação.

## Grupos v1: Sync Documental + Fase 0/1 (Abril 2026)
- **Âmbito Delimitado**: Apenas sincronização documental para GitHub, congelamento da baseline (Fase 0) e contratos/fundações de baixo risco (Fase 1). Não houve abertura de Fase 2+ nem UI pesada nova.
- **Baseline Canonica Introduzida**:
  - `docs/plano_implementacao_grupos_v1.md` ficou como documento-base com nome canónico.
  - Criado `docs/grupos_v1_index.md` para apontar a ordem de precedência entre plano, mockups/exportações mais recentes e docs de suporte.
  - Criado `docs/grupos_v1_fase1_contratos.md` para fixar semântica, cardinalidade, contrato de mudança de grupo, tarefas mínimas e política de persistência/cache.
  - Os relatórios `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-09_com_Explorador_e_Screenshots.md` e `docs/Relatorio_Gestor_do_Grupo_Mockup_2026-04-09_v2_embedded.md` foram alinhados com a guarda conceptual: o Explorador consulta e abre o Gestor; o Gestor é o único editor rico.
  - Criada a pasta `docs/groups_report_assets/` com screenshots extraídos dos HTML aprovados para que os `.md` renderizem no GitHub sem links partidos.
- **Contratos de Código Introduzidos**:
  - Novo módulo `client/src/modules/crm/groups-v1/contracts.ts` com:
    - semântica `principal` / `referencia`
    - helpers para manter `1 email = 0 ou 1 grupo principal`
    - contrato mínimo de mudança de grupo por email
    - contrato mínimo de tarefas
    - contrato de persistência/cache (`GROUPS_PERSISTENCE_CONTRACT`)
  - `client/src/api.ts` passou a expor os tipos partilhados e a aceitar tarefas opcionais em `LinkGroupEntry`.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` deixou de gerir a exclusão de principal/referências em vários pontos ad hoc e passou a reutilizar os helpers partilhados.
  - `client/src/modules/crm/GroupManagerCockpit.tsx` passou a reutilizar os mesmos helpers no fluxo de quick-link.
- **Validação Executada**:
  - `npm.cmd -w client run build` bem-sucedido.
  - `git diff --check` sem erros de whitespace; apenas avisos de LF/CRLF no Windows.
- **Fora do Scope / Não Tocado Propositadamente**:
  - Fases 2+ (`Preparar`, `Explorar` add-in, `Explorador de Grupos`, `Gestor do Grupo` em UI completa).
  - Backend `server/src/linkStore.js`, `server/src/index.js`, `client/src/office.ts` e restantes zonas sensíveis.
  - Nova área grande de settings.
- **Resíduos / Riscos**:
  - `docs/Grupos_08042026.html` existe localmente mas é duplicado byte-a-byte de `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-08.html`; não foi promovido para evitar redundância.
  - A enforcement runtime total do contrato `1 email = 0 ou 1 grupo principal` no backend continua futura; nesta ronda ficou fechado o contrato partilhado e a sua adoção nos pontos cliente já existentes.

## Grupos v1: Fase 2 - Preparar (Abril 2026)
- **Base da Ronda**: A branch desta ronda partiu de `origin/codex/groups-v1-phase0-phase1` porque o commit `1c5540460f50e72b680d4664969adce3cb4cc55f` ainda não estava mergeado em `main`.
- **Revisão Obrigatória da Ronda Anterior**:
  - `client/src/api.ts` ficou limitado a contratos/tipos partilhados e plumbing de export.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` reutiliza helpers de contratos para principal/referências sem alterar layout ou abrir novos fluxos.
  - `client/src/modules/crm/GroupManagerCockpit.tsx` reutiliza os mesmos helpers no quick-link, sem drift visual e sem duplicar responsabilidades.
  - Conclusão: a ronda anterior ficou aprovada como base segura para avançar para Fase 2.
- **Implementação Introduzida em Fase 2**:
  - A aba `Groups` no task pane passou a renderizar uma superfície dedicada a `Preparar` através de `client/src/modules/crm/GroupsPrepareCockpit.tsx`.
  - O shell foi ajustado em `client/src/components/shell/CockpitShell.tsx` para apontar a tab `groups` para essa nova superfície compacta.
  - Foi criado `client/src/modules/crm/groups-v1/prepareSession.ts` para guardar progresso de sessão local de `Preparar` e um seed mínimo de passagem para `Classificar`, sem abrir persistência remota pesada.
  - `Preparar` ficou com:
    - card de Email Âncora
    - switches compactos `Grupo` e `Filtros`, ambos OFF por defeito
    - sub-vistas `Lista`, `Anexos` e `Resumo`
    - lista de emails selecionáveis com cards expansíveis
    - painel de grupo em trabalho, sem edição rica
    - painel de filtros de pesquisa, sem classificação final
    - preparação local de anexos
    - resumo antes da passagem a `Classificar`
- **Guardas Mantidas**:
  - `Preparar` não substitui `Classificar`.
  - Não foi criado viewer novo.
  - Não foi aberta UI pesada de `Explorar`, `Explorador de Grupos` ou `Gestor do Grupo`.
  - A semântica `grupo` / `referência` / `ticket` / `etiqueta` e a regra `1 email = 0 ou 1 grupo principal` mantiveram-se intactas.
- **Scaffolding Deliberadamente Mínimo**:
  - A ponte para `Classificar` escreve apenas um seed local com seleção, grupo em trabalho, anexos preparados e filtros ativos.
  - A passagem integral do conjunto preparado e a persistência remota aprofundada continuam reservadas para fases seguintes.
- **Validação Executada**:
  - `npm.cmd -w client run build`
  - `git diff --check`
- **Fora do Scope / Não Tocado Propositadamente**:
  - Fase 3 completa de persistência.
  - Fase 4 completa de integração entre `Preparar` e `Classificar`.
  - `Explorar` do add-in, `Explorador de Grupos`, `Gestor do Grupo` e a aba principal `Tarefas`.
  - `client/src/modules/crm/GroupManagerCockpit.tsx`, `client/src/modules/crm/GroupClassificationStudioApp.tsx` e `client/src/api.ts` nesta ronda, para evitar scope drift depois da revisão da base.
- **Resíduos / Riscos**:
  - O seed local para `Classificar` nesta fase é scaffolding e ainda não representa uma transferência integral de estado.
  - Continua pendente validação funcional integrada quando a Fase 4 ligar a passagem completa para o fluxo de classificação.

## Grupos v1: Fase 3 - Persistencia segura / cache de sessao / save before exit (Abril 2026)
- **Base da Ronda**: Esta branch partiu de `origin/codex/groups-v1-phase2-preparar` porque nem `1c5540460f50e72b680d4664969adce3cb4cc55f` nem `f426ec7c026f7e66264a94137737373d3115b281` estavam mergeados em `main`.
- **Auditoria Obrigatoria da Fase 2**:
  - `client/src/modules/crm/groups-v1/prepareSession.ts` estava limitado a progresso de sessao local, mas ainda sem politica explicita suficiente e sem disciplina de flush em saidas/context switches.
  - `client/src/components/shell/CockpitShell.tsx` continuou limpo: a tab `groups` abre `GroupsPrepareCockpit` e nao houve impacto funcional observado nas restantes tabs.
  - `client/src/modules/crm/GroupsPrepareCockpit.tsx` manteve o papel de `Preparar` e nao derivou para editor rico nem para `Classificar 2`, mas dependia de um debounce curto para quase todo o save.
  - Conclusao: Fase 2 ficou aprovada como base segura; a unica correcao necessaria para Fase 3 era fechar a politica de sessao e os flushes de saida.
- **Politica de Sessao Fechada**:
  - Criado `docs/grupos_v1_fase3_sessao_cache.md` como referencia curta da politica.
  - `prepareSession.ts` passou a distinguir explicitamente:
    - sessao local de `Preparar` em `sessionStorage`
    - seed temporario de ponte para `Classificar` em `localStorage`
  - O record de sessao passou a ser auto-descritivo (`kind`, `version`, `storage`, `isCanonical`, `lastReason`) para evitar deriva para pseudo-truth-store.
  - O seed local para `Classificar` ganhou TTL e limpeza de seeds stale, para continuar a ser apenas bridge temporaria.
- **Save before exit / context change**:
  - `GroupsPrepareCockpit.tsx` passou a fazer flush local:
    - antes de mudar de sub-vista relevante
    - antes de mudar de email/contexto ancora
    - antes de abrir `Classificar`
    - ao sair da superficie (`unmount`)
    - em `pagehide`, `beforeunload` e `visibilitychange(hidden)`
  - O save diferido durante edicao continua a existir, mas ficou controlado por `sessionScopeKey` + assinatura de snapshot, para nao gravar estado do email errado nem comportar-se como autosave histerico.
- **Rehidratacao e limites**:
  - A sessao continua a reidratar apenas dados de trabalho local: selecao, grupo em trabalho, filtros, anexos preparados e sub-vista.
  - Nao sobe HTML/binarios/resultados de pesquisa/catalogos para sessao.
  - Nao ha promocao remota nova nesta fase.
- **Ficheiros Tocados**:
  - `client/src/modules/crm/groups-v1/prepareSession.ts`
  - `client/src/modules/crm/GroupsPrepareCockpit.tsx`
  - `docs/grupos_v1_fase3_sessao_cache.md`
  - `docs/HANDOFF.md`
  - `docs/DECISIONS.md`
- **Validacao Executada**:
  - `npm.cmd -w client run build`
  - `git diff --check`
  - `npx.cmd eslint src/modules/crm/GroupsPrepareCockpit.tsx src/modules/crm/groups-v1/prepareSession.ts`
- **Fora do Scope / Nao Tocado Propositadamente**:
  - Fase 4 completa de integracao `Preparar -> Classificar`
  - qualquer persistencia remota pesada / backend novo
  - `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` e aba principal `Tarefas`
  - `client/src/components/shell/CockpitShell.tsx` nesta ronda, para nao reabrir shell sensivel sem necessidade
- **Residuos / Riscos**:
  - O lint global do client continua historicamente ruidoso fora do scope; a validacao limpa desta ronda ficou direcionada aos ficheiros tocados.
  - A sessao de `Preparar` continua deliberadamente local e nao resolve ainda a promocao remota/final do conjunto preparado.
  - O consumo do seed de `Preparar` por `Classificar` continua para Fase 4; nesta ronda ficou apenas mais seguro e mais bem delimitado.

## Grupos v1: Fase 4 minima - ligacao limpa `Preparar -> Classificar` (Abril 2026)
- **Base da Ronda**: Esta branch partiu de `origin/codex/groups-v1-phase3-session-cache` porque os commits `1c5540460f50e72b680d4664969adce3cb4cc55f`, `f426ec7c026f7e66264a94137737373d3115b281` e `2fa369a72d9516d4480b2e35ab71cbdd8dbccc9a` ainda nao estavam mergeados em `main`.
- **Auditoria Obrigatoria da Fase 3**:
  - `client/src/modules/crm/groups-v1/prepareSession.ts` continuou limitado a progresso de sessao e seed temporario, com TTL e cleanup adequados.
  - `client/src/modules/crm/GroupsPrepareCockpit.tsx` continuou a usar o seed apenas como ponte de arranque, sem persistencia canonica nem sync continuo.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` era o ponto certo para consumir o seed, no bootstrap inicial do Studio, sem reabrir o fluxo inteiro.
  - Conclusao: Fase 3 ficou aprovada como base segura para a ponte minima desta ronda.
- **Implementacao Introduzida em Fase 4**:
  - `client/src/modules/crm/group-classification/types.ts` e `documentUtils.ts` passaram a reconhecer o parametro `prepareSeedKey` na abertura do Studio.
  - `client/src/modules/crm/GroupClassificationStudioApp.tsx` passou a:
    - ler o seed de `Preparar` de forma controlada
    - validar conteudo/TTL com fallback limpo
    - bootstrapar selecao de emails, `applyScopeMode`, grupo em trabalho e filtro textual relevante
    - promover anexos preparados para `attachmentPlan.save` sem criar storage canonico novo
    - limpar o seed depois de bootstrap valido ou mismatch claro, evitando reconsumo fantasma
  - `client/src/modules/crm/groups-v1/prepareSession.ts` ganhou helper explicito para limpeza do seed, mantendo o modulo leve e focado em sessao/bridge temporaria.
- **Boundaries Mantidas**:
  - O seed continua a ser apenas bootstrap temporario e unidirecional.
  - `Classificar` consome o contexto e segue com o seu estado proprio; nao fica "ligado" ao seed.
  - Nao houve persistencia remota nova, backend novo, viewer novo ou abertura de `Explorar` / `Explorador` / `Gestor`.
- **Validacao Executada**:
  - `npm.cmd -w client run build`
  - `git diff --check`
  - `npx.cmd eslint src/modules/crm/GroupClassificationStudioApp.tsx src/modules/crm/group-classification/documentUtils.ts src/modules/crm/group-classification/types.ts src/modules/crm/groups-v1/prepareSession.ts`
- **Fora do Scope / Nao Tocado Propositadamente**:
  - integracao final completa de persistencia/promocao remota
  - qualquer sync bidirecional `Preparar <-> Classificar`
  - `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` e aba principal `Tarefas`
- **Residuos / Riscos**:
  - o seed continua limitado ao bootstrap local; nao cobre ainda promocao final para backend nem transferencia integral de todos os estados futuros
  - o lint global do client pode continuar vermelho por divida antiga fora do scope, pelo que a validacao limpa desta ronda deve continuar direcionada

## Última orientação estratégica
- Segurança e coerência arquitetural antes de novas features.

## Como retomar o projeto numa nova ronda
1. Ler `AGENTS.md`, este `docs/HANDOFF.md` e `docs/DECISIONS.md`
2. Confirmar o estado atual do repo com `git status`
3. Delimitar o âmbito da ronda em termos concretos
4. Confirmar no código o que é facto e marcar como hipótese tudo o que dependa de produção
5. Só depois planear e intervir

## Antes de codar
- Confirmar que a tarefa respeita as prioridades atuais do projeto
- Identificar zonas sensíveis e dependências cruzadas
- Separar explicitamente:
  - confirmado pelo repositório
  - provável mas não confirmado em produção
- Definir validações mínimas antes de tocar em código
- Garantir que a mudança é delimitada e não transversal por defeito

## Depois de codar
- Validar comportamento, risco e impacto lateral
- Rever segurança, persistência, Odoo, IA e Outlook add-in na zona tocada
- Atualizar `docs/HANDOFF.md` com o novo estado operacional

## Grupos v1: stack Fase 0-4 integrada em `main` (Abril 2026)
- **Integracao Executada**: A stack `Grupos v1` ate Fase 4 foi integrada em `main` por merges `--no-ff`, sem squash e sem rebase, preservando o historico por fase.
- **Checkpoint Pre-Merge**: criada a tag `pre-merge-groups-v1-phase4-stack-2026-04-09` antes de tocar em `main`.
- **Estado Atual em `main`**:
  - baseline documental e contratos de Fase 0/1 ja estao em `main`
  - `Preparar` da Fase 2 ja esta em `main`
  - politica de sessao/cache e save-before-exit da Fase 3 ja estao em `main`
  - ponte minima `Preparar -> Classificar` da Fase 4 ja esta em `main`
- **Guardas Mantidas na Integracao**:
  - nao entrou implementacao de `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` nem aba principal `Tarefas`
  - nao entrou viewer novo
  - nao entrou persistencia remota pesada
  - o seed de `Preparar -> Classificar` continua temporario, unidirecional e nao canonico
- **Proximo Gate Obrigatorio**:
  - o proximo passo ja depende de teste real no host Outlook com utilizador real
  - validar fluxo completo `Preparar -> Classificar`
  - validar comportamento de sessao/cache em contexto real do add-in
  - confirmar que nao ha regressao de UX operacional antes de abrir fases seguintes
- **Nota Operacional**:
  - se os testes reais encontrarem problema, o rollback preferido continua a ser `git revert -m 1 <merge_commit>` pela ordem inversa dos merges desta stack
  - nao usar `reset --hard` em `main`
- **Ronda Visual Mais Recente**:
  - hotfix de compactacao e fidelidade ao mockup em `client/src/modules/crm/GroupsPrepareCockpit.tsx`
  - sem novas features, sem reabrir seed/cache/persistencia e sem tocar em `Explorar` / `Explorador` / `Gestor`
  - switches `Grupo` e `Filtros` redesenhados como controlo compacto real; tabs, cards, badges e footer compactados
  - ajuste visual seguinte apertou ainda mais header, rails, card `Email ancora`, lista e footer para aproximar `Groups > Preparar` do mockup aprovado, sem mexer em logica

## Grupos v1: arquitetura de storage / settings / promocao (Abril 2026)
- **Objetivo desta ronda**:
  - fechar a base tecnica de storage que faltava desde o inicio
  - separar sessao temporaria, persistencia principal e promocao remota
  - reduzir risco de egress prematuro para Supabase
- **Auditoria inicial confirmada**:
  - a sessao/cache atual de `Preparar` continua em `client/src/modules/crm/groups-v1/prepareSession.ts`
  - o repo ainda tinha pressupostos antigos `cloud/local/onedrive` espalhados pelo client
  - o storage principal ainda nao estava descrito de forma canonica nem modular
  - o backend atual (`server/src/linkStore.js`) continua a aceitar sobretudo `cloud/local/onedrive`, pelo que esta ronda introduziu uma bridge de compatibilidade no client sem reabrir o backend
- **Implementacao feita**:
  - nova camada modular em `client/src/modules/crm/groups-v1/storage/`
  - contratos/tipos para:
    - storage mode
    - storage settings
    - session draft
    - workset manifest
    - promotion policy
    - attachment policy
    - storage locations / pointers
  - providers pequenos separados:
    - `supabaseProvider.ts`
    - `localDeviceProvider.ts`
    - `chosenFolderProvider.ts`
    - `hybridProvider.ts`
  - `client/src/settings.ts` passou a normalizar `groupStorage` pelo modelo canonico novo, mantendo compatibilidade com o modelo antigo
  - os fluxos atuais passaram a resolver `attachmentStorageProvider` / `attachmentStorageBasePath` pelo resolver central em vez de ler diretamente `groupStorage.provider/baseFolderPath`
- **Politica fechada nesta fase**:
  - cache do add-in = sessao / rascunho
  - persistencia principal = destino ativo escolhido pelo utilizador
  - Supabase = promocao remota separada
  - binarios grandes pedem decisao
  - promocao binaria automatica para Supabase fica desligada por defeito
  - save before context change / exit continua local; nao foi promovido a sync remoto
- **Doc canonico novo**:
  - `docs/grupos_v1_storage_architecture.md`
- **Limites atuais assumidos**:
  - `local_device` e `chosen_folder` ainda dependem de caminho base explicito; ainda nao existe picker final nesta fase
  - URLs web de OneDrive/SharePoint nao funcionam como destino final no `linkStore`; e preciso pasta sincronizada local ou UNC
  - a promocao final completa de worksets para Supabase continua para fase posterior
- Atualizar `docs/DECISIONS.md` se alguma norma ou decisão mudou
- Registar no output final: alterações, riscos, validações e próximos passos
## Grupos v1: primeira persistencia principal funcional de worksets (Abril 2026)
- **Objetivo desta ronda**:
  - tirar `Preparar` da dependencia exclusiva da sessao local
  - fechar a primeira gravacao principal real do workset
  - manter metadata/manifests first e evitar promocao binaria prematura
- **Implementacao feita**:
  - novo backend pequeno dedicado:
    - `server/src/groupWorksetManifest.js`
    - `server/src/groupWorksetStore.js`
    - rotas `/api/links/groups/worksets`
  - novo conjunto modular no client para save/load e construcao do manifesto:
    - `client/src/modules/crm/groups-v1/storage/buildPrepareWorksetManifest.ts`
    - `client/src/modules/crm/groups-v1/storage/guards.ts`
    - `client/src/modules/crm/groups-v1/storage/loadWorkset.ts`
    - `client/src/modules/crm/groups-v1/storage/mergeWorksetPayload.ts`
    - `client/src/modules/crm/groups-v1/storage/saveWorkset.ts`
    - `client/src/modules/crm/groups-v1/storage/worksetApi.ts`
  - `GroupsPrepareCockpit` passou a:
    - salvar checkpoint principal do workset quando o modo ativo e `supabase` ou `hybrid`
    - reabrir o workset persistido como fallback quando nao existe sessao local
    - manter a sessao local como draft e precedencia quando existe rascunho valido
- **O que ja funciona**:
  - `supabase`
    - save/load real do manifesto persistido
    - metadata/manifests first
    - sem promocao binaria automatica
  - `hybrid`
    - save/load real do manifesto persistido
    - pointers/localizacao principal e politica reference-first de anexos ficam gravados
    - continua sem escrita local final completa do destino principal
- **O que continua parcial**:
  - `local_device`
    - continua scaffold/contrato; falta picker/caminho final e escrita principal completa fora do manifesto
  - `chosen_folder`
    - continua scaffold/contrato; falta picker/caminho final e escrita principal completa fora do manifesto
- **Guardas mantidas**:
  - sessao local continua nao canonica
  - seed `Preparar -> Classificar` nao foi promovido a base de verdade
  - anexos grandes continuam a gerar `requiresDecision` no manifesto e sem promocao binaria automatica
  - payloads pobres passam por merge conservador para nao apagarem manifestos melhores
  - `storage/saveWorkset.ts` e `storage/loadWorkset.ts` usam uma folha HTTP propria (`worksetApi.ts`) e nao importam o hub global `client/src/api.ts`
- **Proximo passo recomendado**:
  - fechar a escrita principal executora para `local_device` / `chosen_folder`
  - depois disso, rever a promocao remota controlada por policy sem abrir egress agressivo

## Grupos v1: storage + worksets integrados em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #14 (`codex/groups-v1-storage-architecture`) integrada em `main` por merge commit explicito
  - PR #15 (`codex/groups-v1-workset-persistence-supabase-hybrid`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-storage-worksets-2026-04-10`
- **O que passou a estar em `main`**:
  - arquitetura canonica de storage de Grupos v1
  - primeira persistencia principal funcional de worksets para modos `supabase` e `hybrid`
  - `local_device` e `chosen_folder` continuam parciais/scaffold, sem promessa de completude
- **Gate seguinte**:
  - teste real em host Outlook/app com foco em `Preparar`, save/load de worksets e regressao de arranque/health
- **Rollback preferido**:
  - reverter primeiro o merge da PR #15 se o problema estiver na persistencia funcional
  - reverter depois o merge da PR #14 se o problema estiver na arquitetura base
  - usar sempre `git revert -m 1 <merge_commit_hash>`, nunca `reset --hard` em `main`

## Grupos v1: hotfix Preparar settings/search/estados/leitura (Abril 2026)
- **Objetivo desta ronda**:
  - corrigir regressao real no teste de `Groups > Preparar`
  - resolver erro de quota em `cockpitSettingsV1`
  - simplificar pesquisa de grupos, repor estados visuais e aproximar contraste/leitura da qualidade de `Classificar`
- **Settings**:
  - `client/src/settings.ts` passou a serializar apenas chaves conhecidas do contrato `CockpitSettingsV1`
  - chaves top-level desconhecidas deixam de ser preservadas por `mergeSettings`, impedindo que worksets/caches/listas/anexos acidentais fiquem presos nas settings
  - o espelho em `localStorage` e best-effort quando existe Office roaming settings; em fallback local, remove a copia antiga antes de tentar regravar a versao compacta
  - `cockpitSettingsV1` continua reservado a preferencias pequenas: modo storage, limites, toggles, paths curtos e escolhas de comportamento
- **Preparar**:
  - pesquisa de grupo em trabalho deixou de carregar uma lista inicial de grupos
  - mostra o grupo principal atual do email quando existir; sem grupo, mostra apenas uma nota compacta
  - sugestoes aparecem so apos pesquisa com pelo menos 2 caracteres, em dropdown simples, sem mini-cards
  - a auto-selecao do grupo principal atual acontece so uma vez por email, para nao bloquear a pesquisa manual
  - toggles `Grupo` e `Filtros` usam cor semantica: verde quando ON, vermelho quando OFF
  - estados visuais pequenos foram repostos para storage (`Local`/`Remoto`/`Hibrido`) e workset (`Sessao`/`Draft`/`Pendente`/`Persistido`)
  - contraste, pesos e badges foram suavizados para leitura mais proxima de `Classificar`, sem mudar o papel funcional de `Preparar`
- **Correcao minima de arranque em Preparar**:
  - durante a validacao visual foi detetado um TDZ local em `GroupsPrepareCockpit.tsx`: efeitos que dependiam de `worksetManifest/worksetSignature` estavam declarados antes da criacao desses valores
  - os efeitos foram apenas movidos para depois do manifesto, sem alterar a arquitetura de storage nem o contrato de persistencia
- **Guardas mantidas**:
  - sem nova UX pesada
  - sem mexer em `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` ou `Tarefas`
  - sem mexer no backend
  - sem transformar sessao/cache/seed em persistencia canonica
- **Proximo passo recomendado**:
  - testar no Outlook real: save de settings apos historico antigo possivelmente poluido, abertura de `Preparar`, pesquisa de grupos e leitura dos estados visuais

## Grupos v1: afinação fina de `Preparar` (Abril 2026)
- **Objetivo desta ronda**:
  - remover poluicao visual remanescente em `Groups > Preparar`
  - corrigir a duplicacao visual do email ancora
  - alinhar estados visiveis e leitura com a semantica aprovada
- **Ajustes feitos**:
  - o email ancora continua a entrar internamente no conjunto de trabalho, mas deixou de aparecer como card normal na Lista
  - o card fechado da Lista mostra apenas assunto, remetente e data/hora, com o indicador de origem reduzido a um ponto discreto junto da checkbox
  - informacao extra como grupo, referencia, ticket, anexos, estado e localizacao ficou limitada ao card expandido
  - os toggles `Grupo` e `Filtros` mantem verde/vermelho apenas no switch; o texto voltou a ficar neutro
  - `Hibrido`, `Remoto`, `Draft`, `Sessao`, `Pendente` e `Persistido` deixaram de ser badges visiveis de utilizador em `Preparar`
  - estados visiveis de localizacao ficaram limitados a `Rascunho`, `Local` e `Servidor`
  - assunto, iconografia primaria e botoes foram suavizados para aproximar a leitura da aba `Classificar`
- **Guardas mantidas**:
  - sem backend
  - sem mexer na arquitetura de storage
  - sem novas features ou novas superficies de Grupos
  - sem transformar `Preparar` em `Classificar 2`
- **Proximo passo recomendado**:
  - validar no Outlook real se a Lista ja deixa de duplicar o ancora e se a leitura dos cards fechados ficou limpa no task pane estreito

## Grupos v1: stack de correções de `Preparar` integrada em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #16 (`codex/groups-v1-workset-api-cycle-hotfix`) integrada em `main` por merge commit explicito
  - PR #17 (`codex/groups-prepare-settings-search-visual-fix`) integrada em `main` por merge commit explicito
  - PR #18 (`codex/groups-prepare-fine-tune-cleanup`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-prepare-stack-2026-04-10`
- **O que passou a estar em `main`**:
  - fronteira HTTP de worksets isolada em `groups-v1/storage/worksetApi.ts`, evitando o ciclo anterior com `client/src/api.ts`
  - settings compactas: `cockpitSettingsV1` volta a guardar apenas preferencias pequenas e conhecidas
  - pesquisa de grupo em `Preparar` simplificada para campo + sugestoes compactas
  - leitura/contraste de `Preparar` ajustados sem abrir nova UX
  - email ancora removido da lista visual duplicada
  - cards fechados limpos e estados visiveis reduzidos a `Rascunho`, `Local` e `Servidor`
- **Guardas mantidas**:
  - sem novas features
  - sem reabrir `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` ou `Tarefas`
  - sem redesenho adicional de `Preparar`
  - sem reabrir arquitetura de storage alem do hotfix de ciclo/imports ja aprovado
- **Gate seguinte**:
  - teste real em host Outlook com foco em arranque sem crash, settings, pesquisa de grupos, lista sem duplicacao do ancora e leitura dos cards
- **Rollback preferido**:
  - reverter primeiro o merge da PR #18 se o problema for visual/semantico no cleanup
  - reverter depois o merge da PR #17 se o problema estiver em settings/search/visual
  - reverter por ultimo o merge da PR #16 se o problema estiver na camada workset/api
  - usar sempre `git revert -m 1 <merge_commit_hash>`, nunca `reset --hard` em `main`

## Grupos v1: hotfix semantica visivel de storage em `Preparar` (Abril 2026)
- **Problema corrigido**:
  - `Preparar` estava a mostrar `Servidor` quando existia workset/manifesto persistido em modo `supabase` ou `hybrid`
  - isto confundia infraestrutura de retoma/progresso com persistencia funcional final do email/classificacao
- **Regra operacional**:
  - `Rascunho` = informacao vinda de Outlook/sessao/preparacao, sem checkpoint local nem sinal funcional persistido
  - `Local` = progresso/checkpoint local ou workset de retoma, ainda sem prova de classificacao final no Supabase
  - `Servidor` = apenas quando o payload do email traz sinais funcionais persistidos, como grupo principal, referencia, motivo de grupo, etiquetas ou status final nao-draft
- **Guardas mantidas**:
  - sem backend
  - sem nova arquitetura de storage
  - sem novas superficies de Grupos
  - sem novos estados visiveis alem de `Rascunho`, `Local` e `Servidor`
- **Proximo passo recomendado**:
  - validar no Outlook real que abrir/preparar ja nao promove visualmente emails para `Servidor` apenas por existir workset salvo

## Grupos v1: semantica visivel de storage integrada em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #19 (`codex/groups-prepare-storage-state-semantics`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-storage-state-semantics-2026-04-14`
- **Gate seguinte**:
  - teste real em host Outlook para confirmar que `Servidor` nao aparece cedo demais e que `Local` cobre apenas workset/checkpoint sem persistencia funcional final

## Grupos v1: correcao semantica do email atual em `Preparar` (Abril 2026)
- **Problema corrigido**:
  - historico relacionado da conversa podia contaminar o email atual e fazer aparecer `Servidor` mesmo quando o email novo ainda nao tinha classificacao propria
  - o aviso de mudanca de grupo podia aparecer por causa de grupos principais em emails relacionados, mesmo quando o email atual nao tinha grupo
- **Regra operacional**:
  - `Servidor` no email ancora depende apenas de sinais funcionais do proprio email atual, identificados por `itemId` ou `internetMessageId` quando existirem
  - historico/conversa/sugestoes/worksets continuam disponiveis como contexto, mas nao contam como classificacao final do email atual
  - o aviso de grupo principal diferente usa apenas o grupo principal real do email atual comparado com o grupo em trabalho
- **Guardas mantidas**:
  - sem backend
  - sem mexer na arquitetura de storage
  - sem nova UX ou novas superficies de Grupos
- **Proximo passo recomendado**:
  - testar no Outlook real um email novo numa conversa com historico ja agrupado e confirmar que fica `Rascunho`/`Local`, sem aviso de mudanca, ate o proprio email ter grupo final

## Grupos v1: estabilidade de `Preparar` antes de merge (Abril 2026)
- **Problema corrigido**:
  - o warning React `Maximum update depth exceeded` vinha da sincronizacao `workingGroupId -> activeGroupSelection`
  - `setActiveGroupForCurrentEmail` era recriada a cada render no `CockpitProvider` e escrevia sempre um novo objeto, criando uma cascata provider/consumer/effect
- **Correcao aplicada**:
  - `setActiveGroupForCurrentEmail` ficou memoizada com `useCallback`
  - a escrita em `activeGroupSelection` passou a ser no-op quando `emailKey` e `groupId` ja estao iguais
- **Auditoria curta**:
  - `Servidor` continua dependente apenas do payload proprio do email atual
  - o aviso de mudanca de grupo continua limitado ao grupo principal real do email atual
  - nao foram abertas novas superficies, backend, storage architecture ou UX
- **Proximo passo recomendado**:
  - se o PR empilhado passar review, integrar a frente e testar no Outlook real com email novo numa conversa com historico agrupado

## Grupos v1: correcao de raiz do email atual e escrita prematura em `Preparar` (Abril 2026)
- **Problemas corrigidos**:
  - a semantica do email ancora ainda podia usar `relatedGroups` / `relatedReasons` vindos do historico relacionado
  - a resolucao do email atual ainda tentava registar automaticamente o email no servidor quando nao encontrava payload completo
  - `Preparar` ainda tinha caminho ativo para persistir worksets remotos em modos `supabase` / `hybrid`
- **Regra operacional atualizada**:
  - o email atual usa helpers diretos para grupo principal, referencia e estado `Servidor`, baseados apenas em sinais proprios (`groupId`, `groupName`, `membershipKind`, labels/status do proprio email)
  - contexto relacionado continua auxiliar para lista/historico, mas nao define estado visual nem aviso de mudanca do ancora
  - abrir/preparar nao chama `registerRelevantEmail` nem pelo cockpit provider nem pela vista, e nao faz flush remoto de workset; `Preparar` guarda apenas sessao/rascunho local e seed temporario para `Classificar`
- **Guardas mantidas**:
  - sem backend novo
  - sem UX nova
  - sem mexer em `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` ou `Tarefas`
- **Proximo passo recomendado**:
  - testar em Outlook real que email novo com historico agrupado fica `Rascunho`/`Local`, sem aviso de mudanca, e que nao ha escrita remota ao abrir/preparar

## Grupos v1: correcao estrutural de `Preparar` integrada em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #22 (`codex/groups-prepare-root-semantics-persistence-fix`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-root-semantics-persistence-fix-2026-04-14`
- **Gate seguinte**:
  - teste real em host Outlook para confirmar que o email atual nao herda historico, abrir/preparar nao grava no servidor, `Servidor` so aparece com sinal funcional final real e o aviso de mudanca de grupo so surge quando o email atual ja tem grupo principal real diferente

## Grupos v1: fecho estrutural `Preparar` + `Classificar` (Abril 2026)
- **Correcoes aplicadas nesta ronda**:
  - a lista de `Preparar` passou a usar helpers diretos de email para grupo principal, referencias e estado `Servidor`; `relatedGroups` / `relatedReasons` deixam de contar como classificacao real do card
  - `Classificar` deixa de chamar `registerRelevantEmail` no carregamento inicial; abrir a vista passa a ser leitura/bootstrap e a persistencia fica no apply final
  - o servidor passa a resolver `email` atual apenas por identidade direta forte (`itemId`, `internetMessageId`, `emailKey`/fingerprint), mantendo historico/conversa em `emails`/`groups`
  - foram removidas chamadas a `persistRelatedEmailsToServer`, que era um no-op com nome enganador
- **Regra operacional reforcada**:
  - email atual, emails relacionados e historico de conversa sao camadas separadas
  - `Preparar` prepara; `Classificar` fecha; carregar contexto nao e classificacao final
- **Proximo passo recomendado**:
  - testar no Outlook real os cenarios de email novo sem grupo, email novo com historico agrupado, abertura de `Classificar` sem escrita remota e classificacao final com persistencia controlada

## Grupos v1: fecho estrutural `Preparar` + `Classificar` integrado em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #23 (`codex/groups-prepare-classify-structural-fix`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-prepare-classify-structural-fix-2026-04-15`
- **Gate seguinte**:
  - teste real em host Outlook para confirmar que historico nao contamina o email atual, `Preparar`/`Classificar` nao gravam cedo demais e a lista mostra apenas grupo real do email

## Grupos v1: correcao cirurgica da frente 1 de `Preparar` (Abril 2026)
- **Correcoes aplicadas nesta ronda**:
  - `Filtros` OFF passa a significar zero filtragem: pesquisa, modo de anexos e modo de grupo ficam inativos ate o painel ser ligado
  - a selecao inicial de trabalho passa a acompanhar o conjunto visivel/ativo de `Preparar` ate haver escolha manual do utilizador, evitando que a aba `Anexos` pareca vazia quando a lista mostra emails relacionados com anexos
  - a lista de relacionados deixa de excluir linhas por match contextual largo de conversa/assunto/remetente; apenas a identidade direta do ancora e removida da lista
- **Guardas mantidas**:
  - sem mexer em `Classificar`, preview de anexos, backend ou storage architecture
  - sem nova UX, `Explorar`, `Explorador de Grupos`, `Gestor do Grupo` ou `Tarefas`
- **Proximo passo recomendado**:
  - validar em Outlook real uma conversa com dois emails relacionados, alternando entre `Filtros` OFF/ON e confirmando que a aba `Anexos` reflete os emails efetivamente selecionados no conjunto ativo

## Grupos v1: frente 1 de `Preparar` integrada em `main` (Abril 2026)
- **Estado pos-merge**:
  - PR #24 (`codex/groups-prepare-filters-attachments-related-fix`) integrada em `main` por merge commit explicito
  - checkpoint pre-merge publicado em `pre-merge-groups-prepare-front1-2026-04-20`
- **Gate seguinte**:
  - teste real em host Outlook para validar `Filtros` OFF/ON, coerencia da aba `Anexos` em `Preparar` e consistencia local da lista de relacionados
  - a frente 2 de anexos/preview em `Classificar` permanece separada e fora deste merge

## IA: settings efetivas e direcao explicita de resposta (Abril 2026)
- **Objetivo desta ronda**:
  - fechar o circuito tecnico entre settings do utilizador e geracao IA
  - separar a pessoa a quem o texto e dirigido dos destinatarios reais do Outlook
- **Implementacao feita**:
  - o contrato `aiGenerate` passou a suportar `length`, `aiKnowledge`, `signature` e `replyDirection`
  - `AiCockpit` e a superficie legada `AiPanel` releem `getSettings()` no momento de gerar, evitando depender de settings carregadas no arranque do painel
  - `length` entra no prompt e tambem regula `max_output_tokens` no servidor
  - `aiKnowledge` entra no prompt como regras fixas do utilizador com prioridade alta
  - `replyDirection.addresseeName/addresseeContext` entra no prompt de `reply` como instrucao de escrita, sem mexer em To/Cc/Bcc
  - respostas podem ser dirigidas explicitamente a uma pessoa indicada mesmo em cadeias de reencaminhamento
  - assinatura oficial vem de `cockpitSettingsV1` e e aplicada no output de reply pelo client; `icc.sig.*` deixa de ser fonte ativa e passa por migracao best-effort para settings oficiais
- **Decisao de superficie IA**:
  - `client/src/modules/ai/AiCockpit.tsx` e a superficie principal do modulo IA no shell atual
  - `client/src/ai/AiPanel.tsx` fica classificado como superficie legada/secundaria; deve manter compatibilidade basica, mas as novas capacidades completas desta ronda nao devem ser vendidas como garantidas por esse painel
  - novas correcoes funcionais do modulo IA devem partir do `AiCockpit` salvo pedido explicito para reativar/consolidar `AiPanel`
- **Guardas mantidas**:
  - sem auto-preenchimento de `To:` a partir de "Dirigir resposta a"
  - sem procura automatica de email da pessoa indicada
  - sem alteracoes em Odoo, manifest ou modulo de Grupos
  - fluxos de copy/insert/new message mantidos sobre o mesmo output final
- **Validacao feita nesta ronda**:
  - build client bem-sucedido (`npm -w client run build`, executado via `npm.cmd` por bloqueio de PowerShell aos shims `.ps1`)
  - build server bem-sucedido (`npm -w server run build`, executado via `npm.cmd`)
  - typecheck client foi executado (`npx tsc -p client/tsconfig.json --noEmit`, via `npx.cmd`) e continua bloqueado por erros pre-existentes fora desta ronda; os erros introduzidos no `AiCockpit` foram corrigidos
  - teste de prompt simulado confirmou que `reply` inclui length, aiKnowledge, assinatura e interlocutor principal `Sr. X`; `summarize` inclui length e aiKnowledge
- **Limites atuais**:
  - nao houve teste em Outlook real nesta execucao; a validacao de host real continua recomendada para confirmar UX e APIs do Outlook
  - a superficie legada `client/src/ai/AiPanel.tsx` continua existente, mas deixa de escrever `icc.sig.*`; o cockpit ativo e `client/src/modules/ai/AiCockpit.tsx`

## IA: MODS oficiais em `responsePresets` (Abril 2026)
- **Decisao final**:
  - MODS passam a ter uma unica fonte oficial: `cockpitSettingsV1.responsePresets`
  - o editor oficial fica em AI Settings, na seccao `MODS / Response Presets` de `client/src/modules/ai/AiSettingsApp.tsx`; `client/src/ui/SettingsPanel.tsx` tambem edita a mesma fonte quando aberto na seccao IA
  - `client/src/modules/ai/AiCockpit.tsx` e a superficie ativa e consome apenas `settings.responsePresets`
  - `client/src/ai/AiPanel.tsx` permanece legado/secundario, mas deixa de editar ou gravar `crmCockpit.templates.v1`; quando mostra templates/MODS, le apenas a fonte oficial
- **Migracao**:
  - `client/src/settings.ts` migra uma vez `crmCockpit.templates.v1` para `responsePresets` quando o legado existe e a lista oficial esta vazia ou apenas com defaults
  - a migracao deduplica por nome/prompt, marca `migrations.legacyResponsePresetsV1` e remove o storage legado quando `getSettings()` persiste a migracao
  - depois da migracao, `crmCockpit.templates.v1` deixa de ser lido como fonte funcional ativa
- **Comportamento no cockpit**:
  - o menu MODS filtra apenas `settings.responsePresets`
  - ao selecionar um MOD, o prompt do MOD entra como instrucao obrigatoria da geracao atual, mantendo o pipeline de `reply`/`forward`, contexto do email, assinatura e normalizacao de output
- **Validacao recomendada fora do repo**:
  - criar, editar, duplicar, reordenar e apagar MODS em Settings > IA
  - confirmar no Outlook que o MOD aparece/desaparece no menu MODS do `AiCockpit`
  - gerar resposta com um MOD e confirmar que a instrucao influencia o texto sem virar texto fixo, salvo quando o proprio MOD for texto fechado
- **Groups settings shell (UI only)**:
  - `client/src/modules/crm/GroupsPrepareCockpit.tsx` passou a expor um icon pequeno de engrenagem no cabecalho de `Groups`, abrindo um modal compacto de settings dentro da propria aba
  - `client/src/modules/crm/groups-v1/settings/GroupsSettingsPanel.tsx` implementa a shell visual com menu lateral, uma secao ativa de cada vez, tooltips discretos em hover e botoes `Fechar` / `Guardar`
  - scope desta ronda fica limitado a interface/estrutura: `General`, `Armazenamento intermedio`, `Anexos`, `Limpeza`, `Avisos`, `Migracao`, `Manutencao`, `Explorar` e `Sobre`
  - campos, toggles e acoes estao stubados/local-state only; nao foi ligada logica pesada de storage real, servidor, migracao, backup, reset, `Preparar` profundo ou `Classificar`
  - afinacao seguinte da mesma shell corrige apenas linguagem e defaults user-facing: `Modo de armazenamento` mostra so `OneDrive / SharePoint` e `Desativado`; `Estrategia de armazenamento` mostra `Todos no servidor`, `Todos fora do servidor` e `Por tamanho`; defaults de `Limpeza`, `Avisos`, `Anexos`, `Migracao` e `Explorar` passam a bater com o contrato fechado, sem alterar a estrutura nem ligar logica real
  - afinacao final da shell: `General` passa a mostrar `Aba Grupos ativa` como controlo visual real; campos de localizacao deixam de usar input livre como controlo principal e passam a leitura + acoes; textos user-facing ficam em PT-PT legivel com acentos, sem redesenhar a shell nem mexer na logica pesada
  - wiring seguinte da shell: `client/src/settings.ts` passa a persistir um bloco proprio `groupsTabSettings`; `client/src/modules/crm/groups-v1/settings/groupsTabSettings.ts` define defaults/normalizacao; `GroupsSettingsPanel` abre com valores reais, edita draft local e `Guardar` persiste via `saveSettings`
  - `locationStatus` e `quickDiagnostic` continuam leves e derivados do proprio bloco de settings; `groupsVersion` fica estatico nesta ronda
  - continuam explicitamente fora do scope: validacao real de localizacao, OneDrive/SharePoint real, migracao real, backup/reset reais, limpeza real, avisos reais, storage final, servidor, `Preparar` profundo e `Classificar` profundo

## Grupos v1: efeitos reais leves dos `groupsTabSettings` na aba `Groups` (Abril 2026)
- **O que passou a ter efeito real nesta ronda**:
  - `groupsTabEnabled` passa a bloquear o uso local da aba `Groups` quando desligado; a vista deixa de expor o fluxo de `Preparar` e mostra um estado claro de modulo desativado
  - `storageMode = disabled` passa a bloquear localmente o fluxo de `Preparar`; a aba deixa de agir como se a base intermedia estivesse ativa e mostra um estado coerente com o storage desligado
  - `locationStatus`, `baseFolderPath` e `quickDiagnostic` passam a aparecer de forma visivel na propria aba, como resumo leve do estado configurado
  - os Settings da aba `Groups` continuam acessiveis mesmo quando o modulo ou o storage estao limitados
- **Guardas mantidas**:
  - sem validacao real de pasta
  - sem OneDrive / SharePoint real
  - sem migracao, backup, reset, limpeza ou avisos reais
  - sem refactor profundo de `Preparar` ou `Classificar`
- **Comportamento deliberadamente fora desta ronda**:
  - `validateLocationOnOpen`, `warnIfUnavailable`, `autoRetryValidation`, `cleanup*`, `warning*`, `attachment*`, `migration*` e `maintenance` continuam apenas persistidos, sem motor real por baixo
- **Proximo passo recomendado**:
  - validar em Outlook real a combinacao de `groupsTabEnabled` e `storageMode`, confirmando que a aba mostra bloqueio claro e reversivel sem fingir capacidades de storage que ainda nao existem

## Grupos v1: alinhamento local de `Preparar` com `groupsTabSettings` (Abril 2026)
- **O que passou a depender de `groupsTabSettings` como fonte principal**:
  - gating leve do modulo (`groupsTabEnabled`)
  - gating leve de armazenamento ativo/desativado (`storageMode`)
  - resumo visual do estado (`locationStatus`, `baseFolderPath`, `quickDiagnostic`)
  - mensagens locais de bloqueio e disponibilidade do fluxo de `Preparar`
- **O que deixou de depender diretamente de `groupStorage` no cockpit**:
  - a montagem local de anexos deixou de ler `settings.groupStorage.ignoreInlineAttachments` diretamente; o cockpit passa a usar apenas um runtime tecnico legado encapsulado
- **O que ainda ficou pendurado no legado tecnico (`groupStorage`)**:
  - resolucao de `legacyStorageRuntime` para workset/storage
  - politica tecnica de ignorar anexos inline, atraves de `legacyStorageRuntime.attachmentPolicy.ignoreInlineAttachments`
- **Porque ficou assim nesta ronda**:
  - estes pontos pertencem ao contrato tecnico da frente de storage/anexos e ainda nao tem equivalente funcional fechado em `groupsTabSettings`; mover agora sem storage real abriria semantica falsa
- **Fora do scope mantido**:
  - storage real, filesystem real, migracao real, limpeza/avisos reais, politica real de anexos, refactor profundo de `Preparar` e `Classificar`

## Grupos v1: modelo canonico da base intermedia por caso (Abril 2026)
- **Objetivo fechado nesta ronda**:
  - criar o contrato canonico do caso intermedio que vai servir de fonte de verdade para `Preparar` e, depois, para `Classificar`, sem abrir ainda storage real ou refactor profundo do fluxo
- **Estrutura do modelo**:
  - `IntermediateCase` passa a separar explicitamente:
    - top-level do caso (`schemaVersion`, `caseId`, `anchorEmailKey`, `conversationId`, `createdAt`, `updatedAt`, `lastAccessedAt`)
    - lista de emails (`IntermediateCaseEmail[]`)
    - classificacao por email (`IntermediateEmailClassification`)
    - anexos por email (`IntermediateCaseAttachment[]`)
    - resumos derivados (`sourceSummary`, `classificationSummary`, `retentionSummary`, `diagnosticSummary`)
- **Casos mistos suportados no contrato**:
  - emails novos vindos do Outlook coexistem com historico ja no servidor
  - alguns emails podem estar classificados e outros nao
  - anexos podem ter decisoes distintas (`local`, `server`, `hybrid`, `metadata_only`, `pending`)
  - o caso pode ficar `local_only`, `mixed` ou `promoted` para preparacao futura da limpeza segura
- **Estrutura fisica alvo da base intermedia**:
  - `Groups/cases/<caseId>/case.json`
  - `Groups/cases/<caseId>/attachments/<emailKey>/...`
  - o `case.json` passa a ser o manifesto canonico do caso; anexos ficam referenciados por `localRef` / `serverRef`, sem obrigar ainda a filesystem real
  - apagar um caso intermedio significa apagar a arvore completa `Groups/cases/<caseId>/...`, nao apenas o `case.json`, para nao deixar anexos orfaos
- **Helpers novos criados**:
  - `createEmptyIntermediateCase`
  - `normalizeIntermediateCase`
  - `buildIntermediateCaseFromSeed`
  - `mergeEmailIntoIntermediateCase`
  - `mergeAttachmentsIntoIntermediateCase`
  - `touchIntermediateCaseAccess`
  - `buildIntermediateCaseSummary`
  - `serializeIntermediateCase`
  - `parseIntermediateCase`
  - repositorio abstrato com `readCase`, `writeCase`, `deleteCase`, `listCases`, `findCaseByEmailKey`
- **O que ficou stub / deliberadamente fora**:
  - OneDrive / SharePoint reais
  - leitura/escrita em pasta real
  - refactor ponta-a-ponta de `Preparar`
  - refactor ponta-a-ponta de `Classificar`
  - endpoints / servidor
  - migracao total do workset antigo
- **Compatibilidade mantida**:
  - workset/storage antigo continua a coexistir
  - o novo modelo canonico fica definido como alvo da frente, sem partir o fluxo atual

## Grupos v1: `Preparar` ligado ao `IntermediateCase` canonico (Abril 2026)
- **O que passou a usar `IntermediateCase` de verdade**:
  - `GroupsPrepareCockpit` passa a montar um `IntermediateCase` explicito a partir do email atual, relacionados visiveis e anexos conhecidos
  - a lista ativa de emails de `Preparar` passa a ser derivada dos emails do `IntermediateCase`, em vez de trabalhar apenas sobre arrays soltos
  - a lista de anexos preparada passa a nascer dos anexos dos emails do caso canonico
  - o caso e escrito num repositorio abstrato em memoria durante a sessao, como ponte para a futura base intermedia real
- **Como o caso e montado nesta ronda**:
  - `caseId` usa `conversationId` quando existe; caso contrario usa o `anchorEmailKey`
  - o email atual entra como ancora do caso e continua a usar apenas sinais diretos proprios
  - emails relacionados entram como emails distintos do caso, com classificacao propria e anexos proprios
  - os anexos selecionados em `Preparar` passam a refletir-se no caso como decisoes locais/pending sem abrir ainda storage real
- **Micro-correcao estrutural seguinte**:
  - o `IntermediateCase` deixa de nascer da lista filtrada/visivel da UI
  - os filtros de `Preparar` passam a afetar apenas a projecao visivel (`visibleEmails` / `visibleListEmails`)
  - o conjunto canonico do caso passa a ser montado a partir do email atual, relacionados conhecidos e emails ja preservados no caso existente
  - emails ja integrados no caso deixam de ser removidos automaticamente so por estarem escondidos por filtros
- **O que ainda ficou legado / ponte temporaria**:
  - o workset antigo continua a existir para draft de selecao, filtros, grupo em trabalho e seed para `Classificar`
  - `legacyStorageRuntime` continua a suportar o gate tecnico de anexos inline e o contrato antigo de workset, ate a ronda de storage real
- **Fora do scope mantido**:
  - sem OneDrive / SharePoint reais
  - sem promocao remota real para servidor
  - sem refactor profundo de `Classificar`
  - sem reescrita total do fluxo `Servidor -> Intermedio -> Outlook`

## Grupos v1: primeira camada de storage intermédio real do `IntermediateCase` (Abril 2026)
- **Auditoria tecnica fechada nesta ronda**:
  - o repo ja tinha prova de uso real de `indexedDB` no cliente (`client/src/modules/crm/excelProvider.ts`)
  - nao existe ainda no repo uma bridge real para escrever diretamente em `groupsTabSettings.baseFolderPath`
  - nao existe ainda integracao real com OneDrive / SharePoint ou com a pasta escolhida pelo utilizador
  - por isso, a camada real desta ronda fecha apenas o que o host atual suporta de forma segura: storage persistente no browser via IndexedDB
- **Adapter real implementado**:
  - `client/src/modules/crm/groups-v1/storage/intermediateCaseIndexedDbAdapter.ts`
  - operacoes reais: `readText`, `writeText`, `deleteTree`, `listPaths`, `readBinary`, `writeBinary`
  - o storage fisico continua a respeitar o contrato logico aprovado:
    - `Groups/cases/<caseId>/case.json`
    - `Groups/cases/<caseId>/attachments/<emailKey>/...`
  - a `baseFolderPath` configurada passa a ser usada como namespace logico do adapter, nao como pasta real validada no host
- **O que ja ficou funcional de verdade**:
  - `case.json` pode ser lido/escrito de verdade
  - `listCases` e `findCaseByEmailKey` passam a funcionar sobre a base persistida no IndexedDB
  - `GroupsPrepareCockpit` ja consegue reabrir um caso persistido dessa base quando `storageMode` esta ativo e existe `baseFolderPath` configurada
- **Como ficaram os anexos nesta ronda**:
  - `case.json` referencia anexos usando o path canonico de `attachments/<emailKey>/...`
  - blobs reais sao escritos apenas quando o host ja tem conteudo do anexo em memoria (`attachment.content`)
  - anexos sem binario disponivel continuam metadata-only; nao se finge preview real nem ficheiro persistido
- **Limitacoes honestas mantidas**:
  - sem escrita direta na localizacao escolhida pelo utilizador
  - sem validacao real de `baseFolderPath`
  - sem OneDrive / SharePoint reais
  - sem promocao para servidor
  - sem refactor profundo de `Classificar`

## Grupos v1: UI de `Preparar` alinhada com o resolver real de storage intermédio (Abril 2026)
- `GroupsPrepareCockpit` passa a usar `intermediateCaseStorage.availability` como verdade local para `ready`, `missing_location` e `disabled`
- `missing_location` deixa de se comportar como storage real pronto; a vista mostra estado explicito de configuracao incompleta e assume apenas modo transitorio em memoria
- o cartao de estado passa a distinguir `Estado real` do storage resolvido e `Configuracao` derivada dos settings
- fora do scope mantido: adapter IndexedDB, OneDrive / SharePoint reais, refactor profundo de `Classificar`

## Grupos v1: precedência real de abertura em `Preparar` (Abril 2026)
- `GroupsPrepareCockpit` passa a abrir o caso por precedência explícita `Servidor -> Intermédio -> Outlook`
- **Servidor nesta frente**:
  - `getRelatedEmailContext(...)` para email atual e históricos relacionados já conhecidos no backend
  - `searchKnownEmails(...)` continua apenas como pesquisa auxiliar da vista; não passa a ser a fonte canónica de abertura do caso
- **Intermédio nesta frente**:
  - `readCase(caseId)` e `findCaseByEmailKey(emailKey)` sobre o storage intermédio real desta ronda
- **Outlook nesta frente**:
  - contexto do email aberto, corpo atual e anexos já carregados no host
- O caso final continua a ser um `IntermediateCase`, mas passa a ser montado por batches de fonte:
  - Outlook como fallback do âncora e dos campos em falta
  - Intermédio para preservar rascunhos e dados locais úteis já existentes
  - Servidor como fonte mais forte quando já há dados persistidos para o email atual e/ou históricos relacionados
- Casos mistos ficam suportados de forma explícita:
  - o `primarySource` do caso pode ser `server`, `intermediate` ou `outlook`
  - cada email mantém `sourceOrigin` próprio
  - o email âncora continua limpo e não é redefinido pelo histórico
- Fora do scope mantido:
  - sem endpoints novos
  - sem promoção real para servidor
  - sem refactor profundo de `Classificar`

## Grupos v1: handoff `Preparar -> Classificar` passa a priorizar o `IntermediateCase` (Abril 2026)
- **Nova regra do handoff**:
  - `Preparar` persiste primeiro o `IntermediateCase` corrente e so depois abre `Classificar`
  - o handoff passa a transportar identidade explicita do caso: `caseId` e `anchorEmailKey`
  - `Classificar` tenta abrir primeiro o caso canonico por `readCase(caseId)` e, se preciso, por `findCaseByEmailKey(anchorEmailKey)`
- **O que `Classificar` ja le do `IntermediateCase`**:
  - email ancora
  - emails relacionados
  - anexos
  - classificacao por email ja existente
  - `sourceSummary` / origem principal do caso
- **Fallback legado que continua nesta ronda**:
  - `seedKey` e `prepareSeedKey` continuam como ponte temporaria para nao partir cenarios onde o caso canonico ainda nao exista
  - o seed legado deixa de ser a verdade principal quando o `IntermediateCase` esta disponivel
- **Garantias mantidas**:
  - o email ancora continua limpo e nao e redefinido pelo historico
  - a abertura a partir do caso canonico nao implica promocao real para servidor
  - limpeza real do intermédio continua fora desta ronda

## Grupos v1: contrato minimo de persistencia final por email ficou fechado em codigo (Abril 2026)
- **Persistencia final agora garantida em codigo para cada email classificado**:
  - identidade forte: `itemId`, `internetMessageId`, `conversationId`
  - contexto de leitura: `subject`, `fromEmail`, `fromName`, `emailWebLink`
  - datas: `messageDateIso`, `receivedAtIso`, `sentAtIso`
  - destinatarios: `toRecipients`, `ccRecipients`
  - corpo/metadados de consulta: `bodyText`, `bodyHtml`
  - classificacao auxiliar: `status`, `labels`, `removedInheritedLabels`, `labelStates`, `classificationMeta`
  - anexos: metadata + refs de storage + `documentState` + `isHidden`
- **Onde ficou fechado**:
  - cliente:
    - `client/src/api.ts`
    - `client/src/modules/crm/group-classification/applyResolution.ts`
    - `client/src/modules/crm/group-classification/documentUtils.ts`
    - `client/src/modules/crm/group-classification/legacyRemoteApply.ts`
    - `client/src/modules/crm/GroupClassificationStudioApp.tsx`
    - `client/src/modules/crm/GroupsPrepareCockpit.tsx`
    - bridges do intermédio em `intermediateCaseAdapters.ts` e `intermediateCaseClassification.ts`
  - servidor:
    - `server/src/linkStore.js`
- **Persistencia final duravel desta fase**:
  - `crm_custom_group_members` passa a guardar tambem `to_recipients_json` e `cc_recipients_json`
  - `buildEmailListEntry(...)` e `mapDbGroupMemberRow(...)` passam a devolver recipients no contrato final
  - o store JSON em memoria/disco tambem passa a reter recipients no email canonico
- **Politica efetiva de anexos fechada nesta fase**:
  - o payload final do email classificado passa a subir explicitamente com `replaceAttachments: false`
  - `registerRelevantEmail(...)` continua a promover anexos com `attachmentStorageProvider` / `attachmentStorageBasePath`
  - `createGroupTicket(...)` passa a receber o mesmo `attachmentStorageOptions` no email base, para nao cair em persistencia parcial diferente quando o ticket e criado nessa operacao
  - binario remoto/final continua fechado apenas onde o provider atual suporta escrita real (`cloud`, `local`, `onedrive` via pasta sincronizada/local); quando nao ha binario ou path valido, fica metadata + refs sem promessas falsas
- **O que continua limitado pelo host/contrato atual**:
  - o contexto Outlook atual ainda nao fornece `emailWebLink` nem `sentAtIso` de forma universal para o email aberto; esses campos continuam best-effort quando ja existem no servidor/intermedio/seed
  - o `IntermediateCase` continua a guardar apenas `to` / `cc` como emails simples, sem nomes; a persistencia final passa a guardar `email + name`
  - URLs web reais de OneDrive / SharePoint continuam fora; a escrita final suportada continua a depender de caminho local/sincronizado quando o provider nao e `cloud`
# HANDOFF

## Hotfix Groups host/auth: `Classificar` sem TDZ e Settings da aba Groups apenas por janela/dialog externa (Abril 2026)
- **`Classificar`**:
  - o bloco de `resolvedApplySelection` deixou de ler `inheritedLabels`, `selectedEmailStoredLabels`, `summaryLabels` e `selectedLabelStates` antes da inicializacao
  - a cadeia imediata de TDZ no studio foi fechada tambem para:
    - `canApplyClassification`
    - `rehydrateClassificationEditorFromCaseEmail`
  - `getEmailGroupRelations(...)` e `rehydrateClassificationEditorFromCaseEmail(...)` passaram a ficar definidos depois de `groupMap` e com `useCallback` estavel, para nao reabrir TDZ nem entrar em loop de `Maximum update depth exceeded`
- **Settings da aba Groups**:
  - `GroupsPrepareCockpit` deixa de renderizar `GroupsSettingsPanel` embebido no taskpane
  - a engrenagem passa a abrir apenas o caminho externo `openGroupsTabSettings(...) -> openGroupSettings(...) -> view=group-settings&surface=groups-tab`
  - `GroupSettingsApp` passa a servir duas superficies:
    - `surface=groups-tab` para os settings reais da aba Groups
    - `surface=manager` para o fluxo legado do gestor
  - o caminho ativo de settings de Groups deixa de importar estaticamente `GroupManagerCockpit`, evitando arrastar `window.confirm` para esta funcionalidade
- **Save / close da janela/dialog**:
  - `Guardar` persiste `groupsTabSettings` via `saveSettings(...)` e mostra estado inline de sucesso/erro
  - `Fechar` fecha via host action quando existe dialog real; no fallback browser/same-window volta para a rota base sem depender de modal embebido
- **Validacao desta ronda**:
  - build passou
  - lint dos ficheiros tocados passou sem erros
  - `git diff --check` passou
  - smoke em browser:
    - `?view=group-classification-studio` abre sem `ReferenceError` nem loop infinito
    - `Groups -> engrenagem` navega para `?view=group-settings&surface=groups-tab`
    - `Guardar` funciona e `Fechar` sai da vista de settings

## Hotfix: estabilizacao de host/settings e semantica de auth no `main` publicado (Abril 2026)
- **Causa raiz corrigida do erro `window.prompt is not supported`**:
  - a aba `Groups` publicada ainda usava `window.prompt` em `client/src/modules/crm/groups-v1/settings/GroupsSettingsPanel.tsx` para editar `baseFolderPath` e `migrationTarget`
  - esse caminho e incompatível com o host do add-in e rebentava ao abrir/usar os settings
  - o painel passa agora a usar um editor inline dentro do proprio modal:
    - `PathEditorState`
    - `openPathEditor(...)`
    - `applyPathEditor(...)`
  - deixamos de depender de `window.prompt`, `alert` ou `confirm` no caminho real de `Groups/settings`
- **Causa raiz corrigida de `Unknown Odoo error` no arranque sem sessao**:
  - o backend devolvia `GET /api/auth/check -> { ok:false }` para ausencia normal de sessao
  - `client/src/api.ts#getJsonErrorMessage(...)` tratava qualquer `ok:false` sem detalhes como `"Unknown Odoo error"`
  - `CockpitProvider` chamava `apiCheckAuth()` no arranque e apanhava esse caso como erro generico em vez de estado normal de nao autenticado
  - o contrato foi corrigido para:
    - backend: `/api/auth/check -> { ok:true, authenticated:false, reason:"no_session" }`
    - frontend: `AuthCheckResponse`
    - `CockpitProvider` passa a distinguir `ok && authenticated` de ausencia normal de sessao
- **Sweep curto de APIs modais do browser nesta frente**:
  - dentro de `client/src/modules/crm/groups-v1/**` nao ficam usos ativos de `window.prompt`, `window.alert` ou `window.confirm`
  - continuam a existir usos fora desta frente, por exemplo em `AiCockpit`, `DialogApp`, `GroupsCockpit` e `GroupManagerCockpit`, mas ficaram fora desta ronda por nao pertencerem ao caminho real pedido
- **Validacao desta ronda**:
  - `npm.cmd install` no worktree limpo para disponibilizar dependencias locais
  - `npm.cmd -w client exec -- eslint src/modules/crm/groups-v1/settings/GroupsSettingsPanel.tsx src/components/shell/CockpitProvider.tsx src/api.ts`
  - `npm.cmd -w client run build`
  - `node --check server/src/index.js`
  - `git diff --check`
  - validacao funcional:
    - browser/Playwright em `https://localhost:5173`
    - `Groups` abre sem erro fatal
    - settings abre em taskpane estreito
    - `Definir namespace` abre editor inline, sem `window.prompt`
    - `GET /api/auth/check` sem sessao responde `200 {"ok":true,"authenticated":false,"reason":"no_session"}`

## Hotfix: estabilizacao imediata da aba Groups publicada (Abril 2026)
- **404 de workset no arranque**:
  - a causa raiz nao era chave invalida nem falha de derivacao
  - o cliente pedia legitimamente o workset do email ancora e o servidor respondia `404 group_workset_not_found` quando ainda nao existia manifesto persistido
  - isso era um miss normal de bootstrap, mas aparecia como erro vermelho destrutivo no browser/taskpane
  - `server/src/index.js` passa a devolver `200` com `exists: false` e `manifest: null` para workset inexistente, preservando `500` apenas para falha real de carregamento
- **Cenario de 2 emails com o mesmo assunto**:
  - a auditoria concluiu que mostrar apenas 1 email nao e, por si so, regressao
  - o conjunto de trabalho em `GroupsPrepareCockpit` nao usa "mesmo assunto" como regra de relacao
  - os emails adicionais entram por identidade forte, relacoes conhecidas do servidor, conversa persistida, grupo/ticket/entidades ou pesquisa/filtro auxiliar
  - por isso, abrir dois emails com o mesmo assunto sem outra relacao conhecida pode continuar a mostrar apenas o email ancora
- **Settings da aba Groups em taskpane estreito**:
  - a causa raiz era estrutural no layout de `GroupsSettingsPanel.tsx`
  - o modal usava sempre grelha fixa de duas colunas (`168px + 1fr`) e rows com segunda coluna fixa (`180px-240px`), o que esmagava sidebar, labels e controls no Outlook estreito
  - o painel passa a colapsar para modo compacto quando a largura da janela e reduzida:
    - header em stack
    - navegacao das secoes em faixa horizontal no topo
    - content com padding reduzido
    - rows com grelha elastica `auto-fit`, sem esmagar controlos
    - actions de path alinhadas a esquerda no modo estreito
- **Fecho tecnico adicional**:
  - `GroupsSettingsPanel.tsx` ficou com `ActionRow` local explicito; deixamos de depender de uma referencia JSX solta dentro da secao de manutencao
- **Validacao desta ronda**:
  - `npm.cmd -w client exec -- eslint src/modules/crm/groups-v1/settings/GroupsSettingsPanel.tsx src/modules/crm/GroupsPrepareCockpit.tsx`
  - `node --check server/src/index.js`
  - `npm.cmd -w client run build`
  - `git diff --check`
  - validacao browser com Playwright em `340x760`: aba `Grupos` abre e o painel de settings fica utilizavel sem erro fatal
  - validacao HTTP direta ao endpoint de workset inexistente: `200 {"ok":true,"exists":false,"manifest":null}`

## Grupos v1: picker/path fica fechado por fluxo manual validado; URL web fica formalmente provada como bloqueio arquitetural (Abril 2026)
- **Picker/path real**:
  - fica assumido como fechado nesta arquitetura por `path manual + normalizacao + validacao real no servidor`
  - nao depende de picker nativo; essa parte continua bloqueada pelo host, mas deixa de ser requisito em aberto porque a alternativa executavel fica assumida como solucao oficial da v1
- **OneDrive/SharePoint por URL web**:
  - continua bloqueado, agora com prova tecnica mais especifica no proprio contrato de validacao
  - o runtime passa a devolver:
    - `architecturalBlocker = web_document_library_requires_graph_backend`
    - `requiredChange` com a mudanca minima necessaria
  - factos confirmados no repo:
    - manifests do add-in continuam apenas com `ReadWriteMailbox`
    - `client/src/office.ts` continua a pedir apenas `Mail.Read`, `User.Read` e `People.Read`
    - o backend escreve binario por filesystem e nao existe uploader Graph/SharePoint por URL web
- **Conclusao desta frente**:
  - `picker/path real`: fechado
  - `OneDrive/SharePoint por URL web`: nao fechado
  - enquanto URL web se mantiver requisito obrigatorio, a fundacao ainda nao pode ser dada como totalmente encerrada

## Pre-verificacao antes de promover `codex/groups-v1-host-auth-stability` para `main` (Abril 2026)
- **Sweep curto do caminho ativo**:
  - no caminho real de arranque + `Groups` + settings de `Groups` ja nao ficam usos ativos de `window.prompt`, `window.alert` ou `window.confirm`
  - os usos restantes encontrados no repo ficam fora desta frente e fora do caminho minimo validado
- **Auth/startup**:
  - `/api/auth/check` sem sessao continua alinhado com o contrato `200 { ok:true, authenticated:false, reason:"no_session" }`
  - `CockpitProvider` continua a distinguir ausencia normal de sessao de erro real
- **Blocker real encontrado no smoke minimo**:
  - `GroupClassificationStudioApp.tsx` ainda tinha uma cadeia de TDZ no arranque do studio
  - causas confirmadas nesta ronda:
    - `selectedEmailRef.current = selectedEmail` antes da declaracao de `selectedEmail`
    - dependencia prematura entre `rehydrateClassificationEditorFromCaseEmail` e `getEmailGroupRelations`
    - uso prematuro de labels derivadas (`categorizableLabels`) na construcao do bloco de apply
  - correcoes feitas nesta ronda:
    - sincronizacao de `selectedEmailRef` movida para depois da inicializacao de `selectedEmail`
    - `getEmailGroupRelations` simplificado para helper local sem dependencia prematura de hook
    - `categorizedLabelNames` no resolved apply deixou de depender da inicializacao adiantada de `categorizableLabels`
  - **estado final desta pre-verificacao**:
    - o studio continua **nao pronto** para promover para `main`
    - permanece um blocker do mesmo tipo no arranque de `Classificar`: `ReferenceError: Cannot access 'inheritedLabels' before initialization`
    - isto prova que ainda existe pelo menos mais uma dependencia cruzada/ordem de inicializacao errada em `GroupClassificationStudioApp.tsx`, pelo que a branch nao deve ser promovida sem fechar essa cadeia primeiro
- **Validacao desta pre-verificacao**:
  - browser/Playwright:
    - app arranca
    - `Groups` abre
    - settings de `Groups` abrem
    - editor inline abre e fecha sem APIs modais incompatíveis
    - regressar ao taskpane principal e voltar a `Groups` nao rebenta o painel
    - `group-classification-studio` continua a falhar no arranque com TDZ residual (`inheritedLabels`)
  - tecnico:
    - `npm.cmd -w client exec -- eslint src/modules/crm/GroupClassificationStudioApp.tsx src/modules/crm/groups-v1/settings/GroupsSettingsPanel.tsx src/components/shell/CockpitProvider.tsx src/api.ts`
    - `node --check server/src/index.js`
    - `npm.cmd -w client run build`
    - `git diff --check`
