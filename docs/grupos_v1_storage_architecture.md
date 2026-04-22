# Grupos v1: fundacao executavel de storage/settings

## Objetivo desta fase
- fechar o que `Preparar` e `Classificar` gravam de verdade
- distinguir claramente intermedio, final e sessao/cache
- alinhar settings com capacidades reais do host e do backend atual
- deixar a base pronta para teste real antes de abrir `Explorar` ou `Gestor do Grupo`

## Fronteira canonica

### 1. Intermedio
- Base: `IntermediateCase`
- Onde grava:
  - `IndexedDB` local do add-in quando `groupsTabSettings.baseFolderPath` existe como namespace logico
  - memoria quando o modo esta `disabled` ou quando nao existe namespace
- Papel:
  - draft
  - continuidade de sessao
  - reidratacao controlada
  - ponte `Preparar -> Classificar`
- Operacoes reais desta ronda:
  - migracao entre namespaces de `IndexedDB`
  - limpeza real de casos promovidos/abandonados no namespace atual

### 2. Persistencia final
- Base: pipeline atual da app via `/api/links/*`
- Backend: `server/src/linkStore.js`
- O que sobe de forma final/coerente:
  - email classificado
  - memberships finais de grupo principal e referencias
  - tickets ligados ao email
  - metadata e refs finais de anexos/documentos
- O que continua central mesmo nos modos file-backed:
  - metadata canonica
  - ligacoes grupo/ticket/email
  - reabertura funcional futura

### 3. Sessao / cache
- `prepareSession`
- seeds temporarias de bootstrap
- fallback em memoria quando o intermedio nao tem namespace
- nao conta como persistencia funcional final

## Politica executavel desta fase

### Intermedio
- `local_indexeddb`
  - com namespace: `IntermediateCase` persistido em `IndexedDB`
  - sem namespace: fallback em memoria
- `disabled`
  - sem persistencia intermedia

### Final
- o apply continua por email alvo e por scope
- o `IntermediateCase` e projetado localmente e depois promovido para a persistencia final atual
- worksets de `Preparar` deixam de ficar desligados a forceps:
  - passam a ser persistidos via `/api/links/groups/worksets`
  - quando o modo principal e file-backed e o caminho e valido, o servidor tambem escreve um mirror JSON do workset nesse destino

## Modos e estado real

| Modo | Estado | Como grava | Onde grava | Limitacoes reais |
| --- | --- | --- | --- | --- |
| `supabase` | Suportado | Persistencia final central da app; binario cloud quando o payload traz conteudo | `/api/links/*` + store central | Nao cria copia local adicional por si so |
| `local_device` | Suportado com validacao | Metadata final continua central; worksets e binario tentam usar path file-backed validado | Caminho local/UNC acessivel ao processo do servidor | Nao representa automaticamente o disco do utilizador sem bridge nativa |
| `chosen_folder` | Suportado com validacao | Metadata final continua central; worksets/binario usam a pasta configurada quando o caminho e fisico | Pasta local/sincronizada/UNC validada no servidor | URL web de OneDrive/SharePoint continua bloqueada |
| `hybrid` | Suportado com validacao | Persistencia final central + mirror local de worksets/binario no destino primario | App central + path local validado | Continua a exigir path fisico acessivel ao servidor |
| OneDrive/SharePoint por URL web | Nao suportado | Sem escrita real nesta arquitetura | n/a | O backend atual grava binario via filesystem; sem Graph/SharePoint API nao ha escrita por URL web |

## Picker/path real

- O repo nao tem bridge nativa que entregue ao backend um caminho local do utilizador atraves de um picker verdadeiro.
- Nesta arquitetura, o caminho suportado nesta fase e:
  - input manual
  - validacao real no servidor
  - bloqueio explicito quando o destino nao e acessivel
- O settings deixa isto explicito; nao vende picker falso.
- Portanto, o requisito funcional de `picker/path real` fica fechado por esta alternativa executavel:
  - path manual
  - normalizacao
  - probe de escrita/leitura real no servidor

## Politica executavel para anexos

### Regra geral
- metadata do anexo sobe sempre quando o payload final inclui o anexo
- `replaceAttachments: false` preserva anexos anteriores quando o payload e parcial

### Binario real
- `cloud`
  - o store atual pode manter metadata + conteudo
- `local` / `onedrive`
  - o backend tenta escrever binario apenas para caminho local/sincronizado/UNC realmente acessivel
  - quando consegue, ficam `storageBasePath`, `storagePathHint` e refs finais
- sem path/provider real
  - fica metadata + referencia
  - nao ha promessa de escrita binaria

### Intermedio
- pode manter `storageDecision`, `localRef`, `serverRef`, `previewReady`
- estes campos continuam com papel de draft/sessao

## Migracao real desta fase

### Ja executavel
- migracao real do `IntermediateCase` entre namespaces de `IndexedDB`
  - copia/move `case.json`
  - copia blobs de anexos locais
  - pode remover a origem em modo `move`
- migracao de workset
  - existe endpoint para migrar o manifesto e reescreve-lo com o destino novo
  - quando o destino e file-backed valido, o mirror JSON passa a ser regravado nesse destino

### Ainda bloqueado
- migracao historica da persistencia final central (`linkStore`) para novo provider file-backed
- mover todos os anexos/documentos ja promovidos para um novo destino sem job backend dedicado

## Limpeza real desta fase

### Ja executavel
- limpeza manual real do intermedio na shell da aba `Groups`
- regras:
  - `promoted` com idade acima de `cleanupClosedCaseDays` -> apaga
  - `local_only` com idade acima de `cleanupAbandonedCaseDays` -> apaga
  - `mixed` so apaga quando `neverDeleteMixedSilently = false`
- apagar um caso remove a respetiva arvore `Groups/cases/<caseId>/...` no namespace do `IndexedDB`

### Ainda fora
- limpeza automatica total agendada
- cleanup de storage final central/historico remoto

## Settings alinhados com a realidade desta fase

### Aba Groups
- continua a representar o intermedio:
  - namespace
  - migracao real de namespace
  - limpeza real do intermédio
- deixa de fingir validacao de pasta cloud ou migracao final total

### Settings globais
- `groupStorage` deixa de tratar `local_device` e `hybrid` como meras shells
- o utilizador pode:
  - escolher o modo
  - definir o path
  - validar o destino no servidor
- se o destino file-backed falhar:
  - o save fica bloqueado
  - o bloqueio tecnico fica visivel

## Bloqueios tecnicos reais desta arquitetura

### OneDrive / SharePoint por URL web
- bloqueado porque o backend atual faz escrita binaria por filesystem
- manifests do add-in no repo continuam apenas com `ReadWriteMailbox`
- `client/src/office.ts` so pede `Mail.Read`, `User.Read` e `People.Read`
- nao existe uploader Graph/SharePoint dedicado no backend atual
- para fechar URL web de verdade seria necessaria integracao dedicada com Graph/SharePoint, autenticacao associada e upload/download por API em vez de filesystem

## Estado de fecho desta frente
- `picker/path real`: fechado pela via manual validada
- `OneDrive/SharePoint por URL web`: nao fechado na arquitetura atual
- enquanto URL web se mantiver requisito obrigatorio, a fundacao de storage/settings nao pode ser dada como totalmente encerrada

### Picker real de pasta
- bloqueado porque o host atual nao expõe ao backend um caminho local do utilizador atraves de picker browser reutilizavel
- a alternativa real desta fase e path manual + validacao real no servidor

### Migracao final total de binarios historicos
- bloqueada sem job backend dedicado para enumerar, copiar e reescrever refs dos documentos/anexos ja persistidos

## O que fica preparado para a fase seguinte
- worksets deixam de estar artificialmente desligados em `local_device` / `chosen_folder` / `hybrid`
- validacao de destino passa a ser real
- settings deixam de prometer capacidades sem prova
- o intermedio ja tem motor real de migracao/limpeza

## Fora de scope mantido
- `Explorar`
- `Gestor do Grupo`
- backend novo gigante
- integracao Graph/SharePoint
- redesign geral da UI
