# Grupos v1: arquitetura de storage

## Fecho de fase: `Preparar` + `Classificar` (Abril 2026)

### Fronteira canonica desta fase
- `IntermediateCase`
  - e **intermedio**
  - vive no host local (`IndexedDB` namespaced por `baseFolderPath` quando a localizacao existe; memoria quando o storage intermadio esta `missing_location`/`disabled`)
  - serve para:
    - draft
    - continuidade de sessao
    - reidratacao controlada
    - ponte `Preparar -> Classificar`
- Persistencia final classificada
  - vive no storage principal atual da frente via `client/src/api.ts` + `server/src/linkStore.js`
  - serve para:
    - emails classificados
    - memberships finais de grupo principal/referencias
    - ligacoes a ticket
    - documentos do grupo
    - base futura para `Explorar` e `Gestor do Grupo`

### O que fica intermadio
- `IntermediateCase.case.json`
- blobs locais do caso quando o adapter suporta binario
- seeds/sessao de handoff e reabertura controlada
- estados de decisao de anexos ainda nao promovidos (`storageDecision`, `requiresDecision`, etc.)

### O que fica final nesta fase
- Por email classificado, via `registerRelevantEmail(...)` / `upsertEmail(...)`:
  - identidade forte (`itemId`, `internetMessageId`, `conversationId`, `emailKey` derivavel)
  - assunto, remetente, datas, corpo e metadata de leitura futura
  - labels, `removedInheritedLabels`, `labelStates`, `classificationMeta`
  - anexos do email com `documentState`, `isHidden`, `hasContent`, `storageProvider`, `storageBasePath`, `storagePathHint`
- Por pertença relacional:
  - grupo principal e referencias via `addEmailToLinkGroup(...)`
  - ticket via `createGroupTicket(...)` / `updateGroupTicket(...)` / `linkEmailToGroupTicket(...)`
  - documentos de grupo via `saveGroupDocuments(...)`

### Regra de promocao desta fase
1. `Preparar`/`Classificar` trabalham no `IntermediateCase`
2. o apply continua por email alvo / por scope
3. a classificacao e promovida para persistencia final pelo pipeline `/api/links/*`
4. o `IntermediateCase` volta a ser escrito apenas como draft/sessao coerente pos-apply

### Politica final desta fase para anexos
- No `IntermediateCase`
  - o anexo pode manter estado intermadio (`storageDecision`, `localRef`, `serverRef`, `previewReady`)
  - serve para continuidade local, preview e reabertura controlada
- Na persistencia final por email (`registerRelevantEmail`)
  - se o payload trouxer `attachments` com conteudo:
    - `cloud`: o store final guarda metadata + conteudo/base64 no proprio store atual
    - `local` / `onedrive`: o backend tenta gravar binario para a pasta configurada; quando consegue, limpa `content` e fica com `storageBasePath` + `storagePathHint` + `hasContent`
  - se o payload final nao trouxer conteudo mas ja existir anexo persistido:
    - o backend preserva storage refs/binario anterior
  - payload parcial nao apaga anexos antigos, salvo `replaceAttachments: true`
- Em documentos do grupo (`saveGroupDocuments`)
  - a regra e paralela: metadata sempre, binario real apenas quando o provider/path suportam escrita segura
- `requiresDecision`
  - continua a ser regra intermadio/sessao, nao contrato do storage final

### Limitacoes reais que continuam nesta fase
- `to` / `cc` ainda nao fazem parte do contrato atual de `RelevantEmailPayload` / `RelatedEmailEntry`; ficam fora da escrita final desta ronda para nao abrir contrato de API novo
- `OneDrive/SharePoint` por URL web continua nao suportado como destino final de ficheiros; o host atual so fecha a escrita real para caminho sincronizado local / UNC
- o `IntermediateCase` continua local ao host e nao substitui a persistencia final

## Objetivo
- Tornar explicita a separacao entre:
  - sessao temporaria do add-in
  - persistencia principal do conjunto de trabalho
  - promocao remota para Supabase
- Reduzir escrita prematura e egress desnecessario.

## Camadas canonicas

### 1. Sessao temporaria
- Vive no add-in.
- Serve apenas para:
  - filtros
  - selecoes
  - sub-vista
  - grupo em trabalho
  - anexos preparados
  - progresso local / rascunho
- Nao e persistencia final.
- Nao substitui a base principal.
- Nao sobe automaticamente para Supabase.

### 2. Persistencia principal
- E a base real de trabalho escolhida pelo utilizador.
- E onde devem viver emails e anexos de trabalho quando saem do estado de rascunho.
- Modos suportados:
  - `supabase`
  - `local_device`
  - `chosen_folder`
  - `hybrid`

### 3. Promocao remota
- E uma fase separada da sessao e da persistencia principal.
- Supabase recebe apenas o que for promovido para la.
- Regra geral:
  - manifestos e metadata podem ser promovidos por politica
  - binarios nao devem ser promovidos automaticamente por defeito
  - anexos grandes pedem decisao
  - payloads pobres nao podem apagar dados bons

## Modos

### Tudo no Supabase
- `mode: "supabase"`
- Persistencia principal remota.
- Continua a existir sessao local de rascunho.
- Promocao binaria automatica fica desligada por defeito.

### Local neste PC
- `mode: "local_device"`
- Persistencia principal local ao dispositivo.
- Supabase fica fora do caminho normal, salvo promocao posterior.
- Limite atual:
  - continua a precisar de caminho base configurado; ainda nao existe picker dedicado nesta fase

### Local em pasta escolhida
- `mode: "chosen_folder"`
- Persistencia principal numa pasta definida pelo utilizador.
- Pode representar filesystem local ou biblioteca sincronizada.
- Limite atual:
  - URLs web de OneDrive/SharePoint nao sao suportadas como destino final pelo `linkStore`; e preciso caminho sincronizado local ou UNC

### Hibrido
- `mode: "hybrid"`
- Persistencia principal local/pasta escolhida.
- Supabase fica disponivel como camada de promocao controlada.
- Mantem separacao explicita entre gravacao principal e promocao remota.

## Politica de anexos
- Limiar configuravel em `attachmentPromptThresholdMb`.
- Acima do limiar:
  - o binario deve pedir decisao do utilizador
  - nao segue por promocao binaria automatica
- Por defeito:
  - `ignoreInlineAttachments = true`
  - promocao binaria automatica para Supabase = desligada
- Em modos locais/hibridos:
  - estrategia principal tende a `store_reference`
- Em `supabase`:
  - a estrategia principal pode aceitar binario, mas a promocao remota binaria continua desligada por defeito

## Manifest / workset
- O modelo tecnico do conjunto de trabalho fica em `client/src/modules/crm/groups-v1/storage/worksetManifest.ts`.
- O manifesto contem:
  - email ancora
  - emails incluidos
  - grupo em trabalho
  - filtros
  - anexos preparados
  - modo de storage
  - localizacao principal
  - localizacao de promocao remota
  - estado de promocao

## Save semantics
- Sessao local:
  - rascunho / progresso
- Persistencia principal:
  - gravacao real do workset
- Promocao remota:
  - fase separada, com politica e controlo proprio

## Implementacao nesta ronda
- Contratos e tipos:
  - `client/src/modules/crm/groups-v1/storage/types.ts`
- Defaults e normalizacao:
  - `client/src/modules/crm/groups-v1/storage/settings.ts`
- Resolucao central do modo ativo:
  - `client/src/modules/crm/groups-v1/storage/resolveStorageMode.ts`
- Politica de anexos:
  - `client/src/modules/crm/groups-v1/storage/attachmentPolicy.ts`
- Politica de promocao:
  - `client/src/modules/crm/groups-v1/storage/promotionPolicy.ts`
- Manifesto:
  - `client/src/modules/crm/groups-v1/storage/worksetManifest.ts`
- Sessao draft:
  - `client/src/modules/crm/groups-v1/storage/sessionDraft.ts`
- Providers/adapters pequenos:
  - `providers/supabaseProvider.ts`
  - `providers/localDeviceProvider.ts`
  - `providers/chosenFolderProvider.ts`
  - `providers/hybridProvider.ts`

## Integracao minima feita
- `settings.ts` passou a normalizar `groupStorage` pelo modelo canonico novo.
- Os fluxos atuais de `Preparar`, `Classificar`, `Gestor`, `GroupsCockpit`, `AI` e bootstrap do cockpit passaram a ler o destino de anexos atraves do resolver central.
- O backend nao foi reaberto nesta ronda.

## Estado funcional atual
- `supabase`
  - ja consegue servir como persistencia principal funcional do workset
  - o manifesto e salvo/carregado por backend pequeno dedicado
  - continua sem promocao binaria agressiva
- `hybrid`
  - ja consegue servir como persistencia principal funcional do manifesto/workset
  - o manifesto guarda pointers locais e politica remote-first/controlada
  - a execucao final do destino local continua parcial
- `local_device`
  - continua parcial
  - falta fechar picker/caminho final e escrita principal fora do manifesto
- `chosen_folder`
  - continua parcial
  - falta fechar picker/caminho final e escrita principal fora do manifesto

## Persistencia principal minima introduzida
- O manifesto de workset passou a ter save/load real via:
  - `server/src/groupWorksetStore.js`
  - `server/src/groupWorksetManifest.js`
  - `/api/links/groups/worksets`
- Em `Preparar`, a sessao continua a ser o draft local.
- Quando o modo ativo e `supabase` ou `hybrid`, o workset passa a ter checkpoint principal persistido.
- A reabertura de `Preparar` pode rehidratar do workset persistido quando nao existe sessao local.

## Fora do scope
- Picker final de pasta
- sincronizacao remota pesada
- nova UX de `Preparar` / `Explorar` / `Gestor`
- aba principal `Tarefas`
- promocao final completa de worksets para Supabase
