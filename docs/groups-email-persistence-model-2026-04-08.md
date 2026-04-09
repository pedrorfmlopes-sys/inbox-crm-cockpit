# Modelo de Persistência de Emails da Aba Grupos

Data: 2026-04-08

## Nota de precedência (2026-04-09)

- Este documento mantém valor como suporte técnico ao modelo de persistência.
- Para baseline ativa de Grupos v1, usar primeiro:
  - `docs/grupos_v1_index.md`
  - `docs/plano_implementacao_grupos_v1.md`
  - `docs/grupos_v1_fase1_contratos.md`

## Objetivo deste ficheiro

- Fixar o modelo que ficou definido para os dados dos emails na aba `Grupos`.
- Separar claramente:
  - o que acontece quando um email é aberto
  - para onde vão os anexos
  - qual é a fonte de verdade
  - qual é o plano de implementação acordado

## Princípio base

O modelo fechado para a aba `Grupos` é:

- a app decide
- a base de dados guarda
- o Outlook só reflete

Isto significa:

- a verdade funcional não vive no Outlook
- a verdade funcional não vive no estado momentâneo da UI
- a classificação e o contexto do caso devem nascer do estado persistido

## O que acontece quando um email é aberto na aba Grupos

Quando o utilizador abre um email com a aba `Grupos` ativa, o fluxo esperado é:

1. o cliente lê o email atual no Outlook
2. tenta obter:
   - identidade do email
   - assunto
   - remetente
   - datas
   - corpo
   - anexos
3. esses dados são enviados para o backend
4. o backend persiste o email
5. a aba `Grupos` e o `Classificar` passam a trabalhar sobre esse estado persistido
6. no fim, o Outlook apenas projeta o resultado, por exemplo com categorias

## O que acontece aos dados do email

Os dados do email ficam divididos em duas partes:

- metadados do email:
  - identidade
  - assunto
  - remetente
  - datas
  - corpo
  - labels/estados/classification meta quando aplicável
- anexos:
  - metadados do anexo
  - conteúdo/binário do anexo

### Metadados do email

Os metadados do email ficam persistidos no backend/store da app.

Isto inclui:

- `itemId`
- `internetMessageId`
- `conversationId`
- `subject`
- `fromEmail`
- `fromName`
- `receivedAtIso`
- `messageDateIso`
- `bodyText`
- `bodyHtml`
- estado de classificação associado ao email

### Anexos

O modelo definido foi:

- a base de dados/store guarda a referência e os metadados dos anexos
- o conteúdo/binário dos anexos vai para o storage configurado da app

Os providers previstos para esse storage ficaram definidos como:

- `cloud`
- `local`
- `onedrive`

Ou seja:

- a BD/store sabe que o anexo existe e como o localizar
- o conteúdo do anexo fica no storage configurado
- depois a UI reusa isso a partir do backend, sem depender do Outlook em tempo real

## Regras importantes que ficaram definidas

### 1. O email aberto é a fonte de ingestão direta

O email aberto é o que a app consegue ler de forma direta e fiável naquele momento.

Logo:

- esse email deve ser o primeiro a ser persistido
- os anexos desse email devem tentar ser persistidos logo na ingestão

### 2. A thread não fica automaticamente toda ingerida

Os outros emails da thread:

- podem aparecer como relacionados
- podem existir na BD
- podem ser reidratados depois

Mas não ficou definido que todos os emails da thread são lidos ao vivo só por pertencerem à mesma conversa.

### 3. Payload parcial não pode destruir dados já persistidos

Ficou definido que:

- um save parcial não pode apagar dados bons já guardados
- `attachments: []` por defeito não pode limpar anexos antigos
- só uma intenção explícita de replace pode substituir/limpar anexos

### 4. O Classificar não deve regravar emails pobres de forma destrutiva

Se o `Classificar` rehidratar um email reduzido:

- não deve regravar esse email de forma a perder anexos já persistidos
- a persistência deve ser conservadora

### 5. O plano final nasce do estado persistido

A classificação final deve seguir esta ordem:

1. guardar na BD
2. reler o contexto persistido
3. construir o plano final
4. refletir no Outlook

## Onde isto ficou implementado / discutido no projeto

### Cliente

- `client/src/components/shell/CockpitProvider.tsx`
- `client/src/modules/crm/GroupClassificationStudioApp.tsx`
- `client/src/modules/crm/group-classification/documentUtils.ts`
- `client/src/api.ts`
- `client/src/office.ts`

### Servidor

- `server/src/index.js`
- `server/src/linkStore.js`

## Plano de implementação que ficou definido

O plano consolidado ficou assim:

### Fase 1. Ingestão do email atual

- ler o email aberto
- ler corpo
- ler anexos
- persistir esse email no backend

### Fase 2. Persistência segura

- guardar metadados do email no store/BD
- guardar metadados dos anexos
- guardar o conteúdo dos anexos no storage configurado
- impedir que payloads parciais apaguem dados anteriores

### Fase 3. Reidratação

- reler o contexto do email/caso a partir do backend
- montar emails relacionados, grupos, tickets e documentos sobre dados persistidos
- evitar depender do Outlook como fonte de verdade

### Fase 4. Classificação

- editar grupo principal
- editar referências
- editar etiquetas
- editar ticket
- guardar tudo no backend

### Fase 5. Projeção Outlook

- depois do estado persistido final estar estável
- construir o plano final de categorias
- aplicar no Outlook apenas como projeção

## O que ficou como critério de qualidade

O modelo só está correto se:

- o email aberto ficar persistido com os dados relevantes
- os anexos não desaparecerem por saves parciais
- o `Classificar` trabalhar sobre estado persistido
- a UI não tratar o Outlook como fonte de verdade
- o Outlook só refletir o resultado final

## Resumo em uma frase

Na aba `Grupos`, o email aberto é ingerido primeiro, os dados ficam persistidos no backend, os anexos ficam referenciados na BD/store e guardados no storage configurado, e toda a classificação posterior deve nascer desse estado persistido antes de qualquer projeção para o Outlook.
