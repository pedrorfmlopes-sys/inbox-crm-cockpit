# Grupos v1 - Fase 3: Sessao e Cache

Data: 2026-04-09

## Objetivo

- Proteger o progresso de `Preparar` sem transformar a sessao numa fonte de verdade canonica.
- Definir claramente o limite entre cache local de trabalho e persistencia remota futura.

## O que e sessao local

- storage principal: `sessionStorage`
- chave: `iccc_groups_prepare_session_v1:<anchorEmailKey>`
- ownership: apenas `Preparar`
- duracao alvo: a sessao ativa do host/task pane

## O que pode viver na sessao

- sub-vista ativa (`list`, `attachments`, `summary`)
- toggles visuais (`showGroupPanel`, `showFiltersPanel`)
- grupo em trabalho (`workingGroupId`, `workingGroupQuery`)
- filtros ativos (`filterQuery`, `attachmentMode`, `groupMode`)
- selecao de emails
- emails expandidos
- selecao de anexos
- metadata de save local (`updatedAtIso`, `lastReason`)

## O que nao pode viver na sessao

- HTML/corpo canonico de emails como verdade final
- binarios de anexos como cache canonica
- resultados completos de pesquisa ou catalogos de backend
- estado remoto de grupos/tickets/documentos
- payload final de classificacao
- qualquer write remoto implicito para Odoo, Postgres ou Supabase

## Save before exit

`Preparar` deve fazer flush local da sessao nestes pontos:

- antes de mudar de sub-vista relevante
- antes de mudar o email/contexto ancora
- antes de sair da superficie `Preparar`
- em `pagehide` / `beforeunload` / `visibilitychange(hidden)`
- antes de abrir `Classificar`

Durante edicao normal, existe apenas um save diferido curto para reduzir perda acidental sem criar um autosave agressivo a cada clique.

## Seed local para Fase 4

- o bridge `Preparar -> Classificar` usa um seed local separado em `localStorage`
- esse seed e temporario, com TTL, e nao substitui persistencia remota
- o seed existe apenas para a futura integracao completa da Fase 4

## Consumo minimo em Fase 4

- `Preparar` escreve o seed e passa apenas a chave por URL
- `Classificar` le o seed de forma controlada, valida conteudo e expiracao, e arranca com bootstrap local quando o seed e valido
- o consumo e unidirecional: nao existe sync continuo entre `Preparar` e `Classificar`
- depois de bootstrap bem-sucedido, o seed e limpo para evitar reconsumos fantasmas
- se o seed estiver expirado, invalido ou nao corresponder ao conjunto atual, `Classificar` cai para o fluxo normal

## Fora desta fase

- persistencia remota pesada
- sincronizacao final com `Classificar`
- qualquer backend novo para promover automaticamente a sessao a fonte de verdade
