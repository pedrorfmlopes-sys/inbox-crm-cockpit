# Grupos v1 — Contratos de Fase 1

Data: 2026-04-09

## Objetivo

- Fechar nomes, invariantes e helpers de baixo risco antes de UI mais pesada.
- Dar ao repo um contrato tecnico reutilizavel para as fases seguintes.

## Semantica final

- `grupo`: pertença principal do email.
- `referencia`: ligacao auxiliar.
- `ticket`: ligacao operacional.
- `etiqueta`: classificacao auxiliar.

## Cardinalidade

- `1 email = 0 ou 1 grupo principal`.
- referencias nunca substituem semanticamente o grupo principal.
- qualquer lista de referencias deve excluir o grupo principal atual.

## Contrato de mudanca de grupo

- a mudanca de grupo principal e explicita.
- a mudanca exige aviso claro ao utilizador.
- pode existir opcao para manter o grupo anterior como referencia.
- essa opcao e facultativa.
- essa conversao so vale para aquele email.
- o contrato tecnico minimo foi fixado em `client/src/modules/crm/groups-v1/contracts.ts`.

## Contrato minimo de Tarefas

- campos minimos:
  - `title`
  - `status`
  - `priority`
  - `dueDate`
  - `owner`
  - `originContext`
  - `notes`
- estados minimos canonicos:
  - `por_fazer`
  - `em_curso`
  - `concluida`
  - `bloqueada`
  - `adiada`
- esta ronda fecha o contrato, nao a UI principal de `Tarefas`.

## Persistencia e cache

- progresso de sessao e persistencia remota sao camadas separadas.
- escrita remota nao deve ser prematura.
- promocao para persistencia remota deve acontecer antes de sair de contexto ou por save explicito.
- a politica minima desta ronda foi fixada em `GROUPS_PERSISTENCE_CONTRACT` e `shouldPromoteGroupsSessionProgress(...)`.

## Codigo ligado a estes contratos

- contrato partilhado: `client/src/modules/crm/groups-v1/contracts.ts`
- tipos de API: `client/src/api.ts`
- orquestracao atual do Studio: `client/src/modules/crm/GroupClassificationStudioApp.tsx`
- orquestracao atual do Group Manager / quick link: `client/src/modules/crm/GroupManagerCockpit.tsx`

## Settings e toggles nesta fase

- nao foi aberta uma nova area grande de settings.
- os contratos ficaram fixados em codigo e documentacao, para evitar scope drift antes de existir necessidade real de UI/configuracao dedicada.
