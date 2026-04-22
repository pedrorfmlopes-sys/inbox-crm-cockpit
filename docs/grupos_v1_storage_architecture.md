# Grupos v1: politica executavel de storage desta fase

## Objetivo desta fase
- fechar o que `Preparar` e `Classificar` gravam de verdade
- distinguir claramente intermedio, final e sessao/cache
- alinhar settings com capacidades reais do host e do backend atual
- deixar a base pronta para teste real antes de qualquer fase de `Explorar` ou `Gestor do Grupo`

## Fronteira canonica

### 1. Intermedio
- Base: `IntermediateCase`
- Onde grava:
  - `IndexedDB` local do host quando existe `baseFolderPath` configurado em `groupsTabSettings`
  - memoria quando o modo esta `disabled` ou quando nao existe namespace configurado
- Como e resolvido:
  - `client/src/modules/crm/groups-v1/storage/resolveIntermediateCaseStorage.ts`
- Papel nesta fase:
  - draft
  - continuidade de sessao
  - reidratacao controlada
  - ponte `Preparar -> Classificar`
- O que nao e:
  - nao e storage final
  - nao e OneDrive/SharePoint real
  - nao e pasta fisica validada pelo utilizador

### 2. Persistencia final
- Base: pipeline atual da app via `client/src/api.ts` + `/api/links/*`
- Backend: `server/src/linkStore.js`
- O que sobe de forma final/coerente:
  - email classificado
  - memberships finais de grupo principal e referencias
  - tickets ligados ao email
  - metadata e refs finais de anexos/documentos
- Funcoes principais:
  - `registerRelevantEmail(...)`
  - `addEmailToLinkGroup(...)`
  - `removeEmailFromLinkGroup(...)`
  - `createGroupTicket(...)`
  - `updateGroupTicket(...)`
  - `linkEmailToGroupTicket(...)`
  - `saveGroupDocuments(...)`

### 3. Sessao / cache
- `prepareSession`
- seeds temporarias de bootstrap
- estado em memoria quando o intermedio nao tem namespace
- serve apenas para continuidade leve da UI
- nao conta como persistencia final

## Politica executavel desta fase

### Intermedio
- `groupsTabSettings.storageMode = "local_indexeddb"`
  - com `baseFolderPath`: usa `IndexedDB` local namespaced por essa chave logica
  - sem `baseFolderPath`: fallback em memoria
- `groupsTabSettings.storageMode = "disabled"`
  - sem persistencia intermedia
  - o cockpit continua apenas com estado local transitario

### Final
- O apply continua por email alvo e por scope
- O `IntermediateCase` e projetado localmente e depois promovido para a persistencia final atual
- O que sobe sempre no payload final por email:
  - identidade forte (`itemId`, `internetMessageId`, `conversationId`)
  - assunto
  - remetente
  - datas relevantes disponiveis
  - corpo/metadados de consulta futura
  - `toRecipients` e `ccRecipients` quando disponiveis
  - labels, `removedInheritedLabels`, `labelStates`, `classificationMeta`
  - grupo principal e referencias via memberships
  - tickets ligados
  - anexos com metadata, estado e refs
- O que continua best-effort:
  - alguns metadados Outlook que o host nem sempre expõe no item aberto
  - binario real de anexos/documentos fora do modo `cloud`

## Modos oficialmente suportados nesta fase

### Intermedio
- `local_indexeddb`
- `disabled`

### Persistencia final de documentos/anexos na shell global
- `supabase` (`Cockpit Cloud`)
- `chosen_folder` (`Pasta local / sincronizada`)

## Modos ou destinos que ficam fora nesta fase
- `local_device` como opcao executavel completa
- `hybrid` como opcao executavel completa
- OneDrive/SharePoint por URL web
- picker real de pasta cloud
- migracao real de storage
- limpeza automatica do intermedio

## Politica executavel para anexos

### Regra geral
- metadata do anexo sobe sempre quando o payload final inclui o anexo
- `replaceAttachments: false` preserva anexos anteriores quando o payload e parcial

### Binario real
- `cloud`
  - o store atual pode manter metadata + conteudo
- `local` / `onedrive`
  - o backend tenta escrever binario apenas para caminho local/sincronizado/UNC realmente acessivel
  - quando consegue, fica com `storageBasePath`, `storagePathHint` e refs finais
- sem path/provider real
  - fica metadata + referencia
  - nao ha promessa de escrita binaria

### Intermedio
- pode manter `storageDecision`, `localRef`, `serverRef`, `previewReady`
- estes campos continuam com papel de draft/sessao
- nao sao o contrato final do storage persistido

## Settings alinhados com a realidade desta fase

### Aba Groups
- o painel da aba Groups passa a descrever:
  - storage intermedio local do add-in (IndexedDB)
  - namespace logico
  - shell nao executavel para migracao, limpeza e explorar
- deixa de apresentar:
  - OneDrive/SharePoint como storage intermedio real
  - validacao de pasta real
  - acoes de migracao/manutencao como se estivessem prontas

### Settings globais
- o bloco `groupStorage` passa a apresentar como executaveis apenas:
  - `Cockpit Cloud`
  - `Pasta local / sincronizada`
- `local_device` e `hybrid` ficam marcados como indisponiveis nesta fase
- caminhos web de OneDrive/SharePoint ficam explicitamente fora

## O que esta pronto para a fase seguinte
- payload final do email classificado mais completo
- recipients persistidos no store final atual
- base suficiente para consultas futuras por identidade, assunto, remetente, datas, labels, grupos, tickets e estado de anexos

## O que continua fora de scope
- `Explorar`
- `Gestor do Grupo`
- novo backend grande
- promocao final nova para servidor
- limpeza real do intermedio
