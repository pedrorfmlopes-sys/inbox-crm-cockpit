# HANDOFF

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
- Atualizar `docs/DECISIONS.md` se alguma norma ou decisão mudou
- Registar no output final: alterações, riscos, validações e próximos passos
