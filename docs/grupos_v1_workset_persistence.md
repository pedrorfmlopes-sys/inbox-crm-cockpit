# Grupos v1: primeira persistencia principal funcional de worksets

## Objetivo desta ronda
- Deixar de depender apenas da sessao local de `Preparar`.
- Introduzir uma gravacao principal real do workset, com scope apertado e metadata/manifests first.

## O que ficou funcional

### `supabase`
- O workset passa a ser gravado e relido por manifesto persistido.
- O manifesto inclui:
  - email ancora
  - emails incluidos
  - grupo em trabalho
  - filtros ativos relevantes
  - anexos preparados como metadata/estado/disposition
  - storage mode
  - localizacao principal
  - localizacao de promocao remota
  - estado de promocao
- A persistencia e feita pelo endpoint `/api/links/groups/worksets`.
- O foco continua a ser metadata/manifests first.
- Nao entrou promocao binaria automatica de anexos.

### `hybrid`
- O workset tambem passa a ter manifesto persistido.
- O manifesto guarda:
  - pointers/localizacao principal local ou pasta escolhida
  - politica de anexos reference-first
  - promocao remota separada e controlada
- Nesta fase, o valor funcional do modo `hybrid` e:
  - o workset deixa de viver apenas em sessao
  - a intencao/localizacao principal fica persistida
  - a promocao remota continua limitada por policy
- A escrita final completa para `local_device` / `chosen_folder` continua fase posterior.

## O que continua parcial

### `local_device`
- Continua com contrato/scaffolding valido.
- Continua a faltar picker/caminho final e fluxo completo de persistencia principal fora do manifesto.

### `chosen_folder`
- Continua com contrato/scaffolding valido.
- Continua a faltar picker/caminho final e fluxo completo para destino final de pasta.
- URLs web de OneDrive/SharePoint continuam fora do caminho final suportado.

## Sessao vs persistencia principal
- `prepareSession` continua a ser apenas rascunho local de `Preparar`.
- O manifesto persistido passa a ser checkpoint principal do workset nos modos `supabase` e `hybrid`.
- A sessao continua com precedencia quando existe rascunho local valido.
- O manifesto persistido serve como:
  - gravacao principal
  - fallback de reabertura
  - base para fases seguintes

## Tratamento de anexos nesta ronda
- O workset persiste apenas metadata/estado dos anexos preparados.
- Cada anexo passa com:
  - `selection`
  - `storageDisposition`
  - `requiresDecision`
- O limiar configuravel continua em `attachmentPromptThresholdMb`.
- Anexos grandes continuam sem promocao binaria automatica.
- Em `hybrid`, o comportamento fica explicitamente reference-first.

## Guardas mantidas
- Nao foi promovida sessao/cache a verdade canonica.
- Nao entrou upload remoto prematuro.
- Nao entrou UX nova pesada.
- Nao entrou falsa promessa de completude para `local_device` e `chosen_folder`.

## Proximo passo recomendado
- Fechar a persistencia principal executora para os destinos locais reais (`local_device` / `chosen_folder`) sem perder esta separacao:
  - sessao local
  - manifesto persistido
  - promocao remota separada
