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
