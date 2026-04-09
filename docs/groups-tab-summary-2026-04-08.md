# Resumo da Aba Grupos

Data: 2026-04-08

## Nota de precedência (2026-04-09)

- Este documento mantém valor como snapshot técnico/histórico.
- Para baseline ativa de Grupos v1, usar primeiro:
  - `docs/grupos_v1_index.md`
  - `docs/plano_implementacao_grupos_v1.md`
  - `docs/grupos_v1_fase1_contratos.md`
  - mockups/exportações de 2026-04-09

## Objetivo deste ficheiro

- Consolidar o que ficou decidido sobre a aba `Grupos`.
- Separar claramente:
  - o que foi decidido nesta thread
  - o que está confirmado no repositório
  - o que continua por validar fora do repo
- Evitar que a leitura futura da área de grupos dependa de memória informal.

## Fontes usadas

- Thread atual até 2026-04-08
- `AGENTS.md`
- `docs/HANDOFF.md`
- `docs/DECISIONS.md`
- `client/src/components/shell/CockpitShell.tsx`
- `client/src/main.tsx`
- `client/src/modules/crm/GroupManagerCockpit.tsx`
- `client/src/modules/crm/GroupExplorerApp.tsx`
- `client/src/modules/crm/GroupsCockpit.tsx`
- `client/src/ui/SettingsPanel.tsx`
- `server/src/index.js`

## 1. O que ficou decidido nesta thread

- As rondas recentes desta thread focaram-se no `Classification Studio` e no pipeline de categorias Outlook.
- Foi pedido explicitamente para não abrir novos temas e para não criar scope drift.
- Nessas rondas, a área de `Grupos` foi tratada como fora de scope.
- Portanto, nesta thread, não ficou decidida nenhuma nova feature nem nenhum refactor específico da aba `Grupos`.
- A decisão operacional foi proteger a aba `Grupos` de alterações enquanto se corrigia o sync Outlook do Studio.
- A promoção recente para `main` serviu para permitir deploy e teste do hotfix Outlook no Render, não para declarar a aba `Grupos` como concluída.

## 2. Decisão arquitetural já refletida no repositório

### 2.1 Superfície principal da aba

- Confirmado no repo: a aba `Grupos` do shell principal renderiza `GroupManagerCockpit`.
- Confirmado no repo: `GroupsCockpit.tsx` continua a existir, mas não é a superfície principal ligada à navegação do shell.
- Confirmado no repo: existem também views standalone para:
  - `group-manager`
  - `group-explorer`
  - `group-settings`

### 2.2 Dependência de Odoo

- Confirmado no repo: a aba `Grupos` pode ser usada mesmo sem autenticação Odoo ativa.
- Confirmado no repo: o gate do shell só bloqueia `crm`, `crm2` e `related`; não bloqueia `groups`.

### 2.3 Fonte de verdade dos grupos

- Confirmado no repo: os grupos vivem na camada própria de links/custom groups da app.
- Confirmado no repo: a aba consome endpoints próprios em `/api/links/groups`.
- Confirmado no repo: isto aponta para `linkStore` e respetiva persistência interna, não para Odoo como SSOT desta área.

## 3. O que está assumido/implementado para a aba Grupos

### 3.1 Lista central de grupos

- Pesquisa de grupos por nome.
- Criação rápida de grupo quando não existe match exato.
- Filtros por estado, arquivo e etiquetas.
- Favoritos e ordenação por relevância/recência.
- Entrada no detalhe de um grupo a partir da lista.

### 3.2 Detalhe do grupo

- Edição de:
  - nome
  - descrição
  - estado
  - etiquetas
  - `documentsEnabled`
  - `isArchived`
- Associação do email atual ou da thread atual ao grupo.
- Distinção entre ligação `principal` e `referencia`.
- Gestão de emails já ligados ao grupo.
- Entrada para biblioteca de emails registados.
- Entrada para explorer do grupo.
- Entrada para operações de tickets.

### 3.3 Ligação rápida do email atual

- Existe uma view própria de `Ligacao rapida`.
- O fluxo permite, numa só operação:
  - escolher grupo principal
  - escolher grupos secundários
  - aplicar etiquetas
  - ligar a ticket existente
  - criar ticket novo
  - guardar a ligação da thread atual
- O fluxo foi desenhado para atuar sobre a thread atual, não apenas sobre um email isolado.

### 3.4 Gestor de etiquetas dentro dos grupos

- Existe catálogo central de etiquetas.
- É possível:
  - criar
  - renomear globalmente
  - eliminar globalmente
  - configurar se a etiqueta vira categoria Outlook
  - configurar se a etiqueta tem estado
- A ativação global desta área é feita em `Settings > Grupos`.

### 3.5 Tickets dos grupos

- Existe extra opcional de tickets dentro da aba `Grupos`.
- O módulo suporta:
  - séries
  - prefixos
  - contadores
  - comportamento de auto-link
  - drafts sugeridos
  - drafts com IA
  - inclusão do código de ticket no assunto
- A ativação global desta área também é feita em `Settings > Grupos`.

### 3.6 Categorias Outlook ligadas aos grupos

- Existe configuração para escrever categorias Outlook com base em:
  - grupos
  - tickets
  - estados
  - etiquetas
- Esta capacidade é opcional e governada por `Settings > Grupos`.
- A aba `Grupos` chama sync Outlook após várias ações que alteram o contexto do email atual.

### 3.7 Documentos e anexos dos grupos

- A aba permite guardar anexos do email atual como documentos do grupo.
- O fluxo usa configuração global de armazenamento dos grupos:
  - provider `cloud`
  - provider `local`
  - provider `onedrive`
- O utilizador pode definir base path/localização base.
- Há suporte para ignorar imagens inline e para preparação de pastas automáticas.

### 3.8 Explorer do grupo

- Existe um explorer dedicado para consultar:
  - emails ligados ao grupo
  - documentos do grupo
  - previews
  - downloads
  - remoção de documentos
  - reabertura de email ligado
  - envio de documento para compose

## 4. Como isto está implementado hoje

### 4.1 Entrada principal no shell

- `CockpitShell.tsx` liga `tab === "groups"` a `GroupManagerCockpit`.
- Isto confirma que a experiência principal da aba `Grupos` hoje é o `GroupManagerCockpit`.

### 4.2 Orquestrador principal

- `GroupManagerCockpit.tsx` é o shell funcional da área.
- As views internas confirmadas no repo são:
  - `groups`
  - `detail`
  - `library`
  - `quicklink`
  - `settings`
  - `labels`
  - `tickets`
- O componente acumula a orquestração de estado, chamadas API e ligação a Office.js.

### 4.3 Fluxo de criação e edição de grupo

- Criar grupo:
  - pesquisa atual não encontra match exato
  - chama `createLinkGroup`
  - seleciona o grupo novo
  - entra na view de detalhe
- Editar grupo:
  - recolhe draft local
  - normaliza etiquetas
  - chama `updateLinkGroup`
  - refresca lista e detalhe

### 4.4 Fluxo de ligação do email atual

- O cockpit recolhe a thread atual com `collectCurrentThreadEmails`.
- A ligação é gravada com `addEmailToLinkGroup`.
- A ligação distingue `principal` de `referencia`.
- Depois de guardar, faz refresh do estado e tenta sync de categorias Outlook para o item atual.

### 4.5 Fluxo de ligação rápida

- A `Ligacao rapida` junta grupos, etiquetas e ticket numa única operação.
- Sequência confirmada no repo:
  - resolve grupo principal e secundários
  - garante etiquetas no catálogo, se o gestor estiver ativo
  - adiciona emails da thread aos grupos
  - cria ou atualiza ticket, se aplicável
  - liga emails da thread ao ticket
  - refresca dados
  - tenta sync Outlook
  - opcionalmente abre draft de resposta se a configuração o pedir

### 4.6 Fluxo de documentos

- Os anexos selecionados do email atual são preparados no cliente.
- Quando falta conteúdo local, o cockpit tenta carregar o base64 do anexo persistido.
- O save usa batches para evitar cargas demasiado grandes num único pedido.
- O persist é feito com `saveGroupDocuments`.

### 4.7 Explorer

- O `GroupExplorerApp` carrega emails e documentos do grupo em paralelo.
- Faz preview de:
  - imagem
  - PDF
  - Office/WebViewer
  - texto
- Também suporta:
  - remover email do grupo
  - apagar documento
  - descarregar documento
  - anexar documento a compose
  - reabrir email ligado no Outlook

### 4.8 Settings dos grupos

- `SettingsPanel.tsx` concentra a configuração global da área.
- O que está confirmado no repo:
  - toggle para gestor de etiquetas
  - toggle para tickets
  - toggle para categorias Outlook
  - configuração da origem de armazenamento
  - base folder path
  - criação automática de pasta
  - opções ligadas ao tratamento documental

### 4.9 Backend

- `server/src/index.js` expõe endpoints para:
  - listar grupos
  - criar grupo
  - atualizar grupo
  - apagar grupo
  - listar emails de um grupo
  - adicionar emails a grupo
  - remover emails de grupo
  - listar documentos
  - obter conteúdo de documento
  - guardar flags de anexos
  - guardar documentos
  - apagar documento

## 5. Leitura operacional do que foi realmente decidido implementar

Se eu reduzir tudo ao essencial, o que está decidido e refletido hoje no repo para a aba `Grupos` é isto:

- A aba `Grupos` é um cockpit próprio, centrado na camada interna de grupos e não dependente de Odoo para abrir.
- A superfície principal escolhida para essa aba é `GroupManagerCockpit`.
- O fluxo principal da área é:
  - encontrar/criar grupo
  - ligar o email ou a thread atual
  - distinguir grupo principal vs referência
  - gerir etiquetas
  - opcionalmente gerir tickets
  - opcionalmente projetar categorias Outlook
  - guardar e explorar documentos do grupo
- A configuração global da área fica em `Settings > Grupos`.
- O explorer existe como superfície auxiliar de consulta e preview.

## 6. O que esta thread NÃO decidiu

- Não decidiu um novo redesign da aba `Grupos`.
- Não decidiu mexer no backend de grupos.
- Não decidiu substituir `linkStore` por Odoo nesta área.
- Não decidiu remover `GroupsCockpit.tsx`.
- Não decidiu novo refactor estrutural do `GroupManagerCockpit`.
- Não fechou funcionalmente a integração Outlook da aba `Grupos` com teste real.

## 7. Impacto indireto das últimas hotfixes Outlook

- Confirmado no repo: `GroupManagerCockpit` chama `syncCurrentItemOutlookCategoriesFromContext()` através de `syncCurrentEmailOutlookCategories()`.
- Confirmado no repo: as últimas hotfixes desta thread mexeram no pipeline de confirmação/readback dentro de `client/src/office.ts`.
- Hipótese razoável, mas ainda não validada fora do repo:
  - como a aba `Grupos` usa o mesmo entrypoint de sync Outlook, a robustez extra do pipeline poderá beneficiar também ações disparadas a partir da aba `Grupos`
- Isto não foi validado funcionalmente em Outlook real nesta thread.

## 8. Riscos e pontos em aberto

- `GroupManagerCockpit.tsx` continua a ser um ficheiro grande e sensível.
- Existe coexistência de `GroupManagerCockpit` e `GroupsCockpit`, o que sugere dívida/legado e pede cuidado antes de novas rondas.
- A persistência documental externa (`local`/`onedrive`) precisa sempre de validação real por ambiente.
- O comportamento real do sync Outlook continua dependente do host Outlook e deve ser validado fora do repo.

## 9. Conclusão curta

- Nesta thread, a principal decisão sobre a aba `Grupos` foi não a tocar durante os hotfixes do Studio/Outlook.
- No repositório, a direção já materializada é clara:
  - `GroupManagerCockpit` como aba principal
  - gestão central de grupos
  - ligação rápida da thread
  - etiquetas e tickets opcionais
  - documentos por grupo
  - integração Outlook opcional via settings
- Se a próxima ronda for para mexer mesmo na aba `Grupos`, o ponto de partida certo é este cockpit e não os hotfixes recentes do Studio.
