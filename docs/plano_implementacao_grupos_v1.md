# Plano de Implementação — Grupos v1

**Projeto:** Inbox CRM Cockpit  
**Data base:** 2026-04-09  
**Estado:** baseline operacional para implementação faseada  
**Objetivo:** servir como mapa de construção, fonte única de continuidade entre agents e proteção contra regressões conceptuais.

## 1. Finalidade deste documento

Este documento define:

- a ordem de implementação aprovada;
- as guardas conceptuais que não podem ser violadas;
- o que entra e o que não entra em cada fase;
- os critérios de fecho de cada fase;
- o protocolo de continuidade quando houver troca de agent;
- a relação entre Grupos, Explorador, Gestor do Grupo e Tarefas.

Este plano não substitui prompts por fase.  
Ele funciona como **mapa-mãe de construção**, para que qualquer agent trabalhe sobre a mesma arquitetura.

---

## 2. Fonte de verdade e prioridade

### 2.1. Fonte de verdade v1
A fonte de verdade visual/funcional v1 é o estado **mais recente** destes mockups no canvas:

1. `Groups-tab-mockups`
2. `Groups-explorer-mockups`
3. `Group-manager-mockups`

### 2.2. Regra crítica
Não usar apenas relatórios ou handoffs em texto como baseline final quando houver divergência com o canvas.

### 2.3. Ordem de precedência
Quando existir conflito, seguir esta ordem:

1. mockups mais recentes no canvas;
2. decisões conceptuais fechadas nos handoffs mais recentes;
3. relatórios de implementação/mockup mais recentes;
4. versões intermédias antigas.

### 2.4. Regra de desenho
**Não redesenhar o que já foi aprovado.**  
A implementação deve traduzir a arquitetura aprovada, não reinventá-la.

---

## 3. Guardas conceptuais obrigatórias

## 3.1. Separação semântica
Manter separação rígida entre:

- **Grupo** = pertença principal;
- **Referência** = ligação auxiliar;
- **Ticket** = ligação operacional;
- **Etiqueta** = classificação auxiliar.

### Regra dura
**1 email = 0 ou 1 grupo principal.**

Não voltar a permitir que referências apareçam como grupo.

## 3.2. Mudança de grupo
Quando um email muda de grupo:

- a mudança tem de ser explícita;
- tem de existir aviso claro;
- pode existir opção para o grupo antigo passar a referência;
- essa opção é facultativa;
- essa conversão afeta apenas aquele email.

## 3.3. Preparar não substitui Classificar
`Preparar` serve para:

- montar seleção;
- escolher grupo em trabalho;
- aplicar filtros de procura;
- preparar anexos;
- enviar para `Classificar`.

`Classificar` continua a ser o local onde se fecha a classificação final.

## 3.4. Explorar dentro de Grupos
A sub-aba `Explorar` dentro da aba `Grupos` é:

- global;
- independente;
- sem `Email Âncora`;
- sem contexto herdado do email atual.

Não reintroduzir contexto do email atual nessa sub-aba.

## 3.5. Três níveis diferentes
Não confundir:

- **Explorar dentro da aba Grupos** = consulta rápida no add-in;
- **Explorador de Grupos** = vista maior de consulta e navegação;
- **Gestor do Grupo** = editor rico único.

Resumo funcional:

- o add-in prepara e consulta rápido;
- o Explorador consulta com mais profundidade;
- o Gestor edita.

## 3.6. Gestor do Grupo = editor rico único
A edição rica e completa do grupo vive apenas no `Gestor do Grupo`.

Pode ser chamado a partir de:

- `Classificar`;
- `Explorador de Grupos`;
- futuramente, também da aba `Grupos`.

Mas não deve existir edição rica duplicada noutras vistas.

## 3.7. Regra específica do mockup atual do Explorador
No mockup atual do `Groups-explorer-mockups` existe um bloco intermédio de “gestor do grupo”.

Esse bloco deve ser lido como:

- ponte conceptual;
- ponto de transição;
- entrada para abrir o Gestor.

Não deve ser interpretado como autorização para construir um segundo editor rico dentro do Explorador.

### Regra final
- o Explorador consulta e abre o Gestor;
- o Gestor é que edita;
- não duplicar edição rica dentro do Explorador.

## 3.8. Viewer / preview
Regra fixa:

- não criar viewer novo paralelo;
- reutilizar a lógica/direção do viewer do `Classificar`;
- preview em baixo;
- preview full-width;
- aplicar tanto ao `Explorador de Grupos` como ao `Gestor do Grupo`.

O mockup define a direção visual.  
O alvo técnico é **reuso**, não duplicação.

## 3.9. Persistência / egress
Não assumir como correto enviar tudo logo para Supabase.

Direção aprovada:

- reduzir escrita prematura;
- usar cache/progresso de sessão;
- gravar antes de sair do contexto;
- selecionar melhor o que sobe para persistência remota.

## 3.10. Estado do hotfix Outlook
O hotfix das categorias Outlook **não** está funcionalmente fechado só porque foi mergeado em `main`.

Só deve ser tratado como fechado após:

- teste real em host;
- confirmação por readback real do host.

---

## 4. Contrato mínimo de Tarefas nesta fase

O módulo principal `Tarefas` fica para fase futura separada.

Nesta frente atual, Tarefas entram apenas em modo leve, para não bloquear Grupos.

### 4.1. O que fica definido já
Contrato mínimo:

- título;
- estado;
- prioridade;
- prazo;
- responsável;
- origem/contexto;
- notas curtas associadas, se necessário.

### 4.2. Estados mínimos
Estados mínimos recomendados:

- Por fazer;
- Em curso;
- Concluída;
- Bloqueada;
- Adiada.

### 4.3. Onde aparecem agora
No contexto atual:

- em `Explorar` no add-in = visão curta;
- no `Explorador de Grupos` = visão curta/média;
- no `Gestor do Grupo` = secção própria, mas ainda dentro do contexto de grupo.

### 4.4. O que não entra agora
Não desenhar ainda:

- gestor principal transversal de tarefas;
- vista global completa de tarefas;
- dashboards de tarefas;
- automações extensas entre múltiplas fontes.

Isso fica para a futura aba principal `Tarefas`.

---

## 5. Estratégia global de implementação

A implementação deve seguir uma lógica de baixo risco:

1. fixar invariantes;
2. construir o fluxo mínimo útil;
3. validar navegação e persistência;
4. só depois abrir as vistas maiores;
5. só depois enriquecer edição e extensões.

### Regra operacional
Uma ronda = um scope.  
Não misturar duas frentes grandes na mesma ronda.

---

## 6. Fases de implementação

## Fase 0 — Congelamento do baseline

### Objetivo
Congelar a baseline funcional/visual v1 antes de mexer no código.

### Inclui
- registar este plano no projeto;
- referenciar os 3 mockups do canvas como fonte de verdade v1;
- registar as guardas conceptuais;
- alinhar a terminologia oficial.

### Não inclui
- alterações de UI;
- alterações de comportamento.

### Critério de fecho
Existe um documento único no projeto com:

- baseline;
- guardas;
- ordem de fases;
- distinção entre Explorar / Explorador / Gestor.

---

## Fase 1 — Modelo, contratos e Settings mínimos

### Objetivo
Fechar o contrato de dados e comportamento antes da UI mais pesada.

### Inclui
- semântica final de grupo/referência/ticket/etiqueta;
- regra de 0 ou 1 grupo principal por email;
- comportamento de mudança de grupo;
- contrato mínimo de Tarefas;
- definição dos modos de persistência/cache relevantes;
- toggles e flags necessários em Settings, se já forem necessários para suportar a arquitetura.

### Não inclui
- gestor completo de tarefas;
- explorador rico;
- redesign de Classificar.

### Dependências
- Fase 0 fechada.

### Critério de fecho
Os contratos e nomes usados no código/UI ficam estáveis e alinhados com o plano.

---

## Fase 2 — Implementar `Preparar` na aba Grupos

### Objetivo
Construir a entrada de preparação do conjunto de trabalho.

### Inclui
- `Email Âncora`;
- switches compactos `Grupo` e `Filtros`;
- sub-vistas `Lista`, `Anexos`, `Resumo`;
- seleção de emails;
- preparação de anexos;
- escolha de grupo em trabalho;
- filtros de procura;
- UI fiel ao mockup aprovado.

### Não inclui
- classificação final;
- lógica rica de edição de grupo;
- substituição de `Classificar`.

### Dependências
- Fase 1 fechada.

### Critério de fecho
A navegação e estrutura de `Preparar` estão fiéis ao baseline aprovado e não invadem o papel de `Classificar`.

---

## Fase 3 — Persistência segura e cache de sessão

### Objetivo
Garantir que o trabalho preparado não se perde e não provoca escrita remota prematura.

### Inclui
- cache/progresso de sessão;
- guardado antes de sair do contexto;
- proteção contra payload pobre destruir dados bons;
- definição prática do que sobe e do que não sobe para persistência remota.

### Não inclui
- otimização final de todos os modos de storage;
- arquitetura final de storage externo avançado, se ainda não for necessária.

### Dependências
- Fase 2 funcional.

### Critério de fecho
Não há perda de trabalho ao mudar de email/aba/contexto e não há promoção prematura desnecessária para Supabase.

---

## Fase 4 — Ligação de `Preparar` ao `Classificar`

### Objetivo
Passar o conjunto preparado para o local certo de fecho da classificação.

### Inclui
- passagem de emails selecionados;
- passagem de anexos preparados;
- passagem do grupo em trabalho;
- passagem de filtros/contexto relevante;
- abertura do `Classificar` com o contexto correto.

### Não inclui
- recriação do `Classificar`;
- editor rico do grupo dentro desta ponte.

### Dependências
- Fases 2 e 3 fechadas.

### Critério de fecho
O `Classificar` recebe o contexto certo, sem ruído nem duplicação de responsabilidades.

---

## Fase 5 — Implementar `Explorar` dentro da aba Grupos

### Objetivo
Criar a consulta rápida global no add-in.

### Inclui
- pesquisa global leve;
- resultados;
- detalhe curto;
- notas leves;
- tarefas leves;
- botão para abrir o `Explorador de Grupos`.

### Não inclui
- Email Âncora;
- contexto herdado do email atual;
- edição rica do grupo;
- exploração pesada tipo workspace completo.

### Dependências
- Fase 1 fechada.

### Critério de fecho
`Explorar` funciona como consulta rápida independente e não mistura o seu papel com `Preparar` nem com o `Explorador` maior.

---

## Fase 6 — Implementar o `Explorador de Grupos`

### Objetivo
Criar a vista maior de consulta, navegação e acompanhamento.

### Inclui
- coluna esquerda com pesquisa/filtros/resultados;
- área de detalhe;
- emails/documentos/notas/tarefas em modo de consulta;
- preview em baixo, full-width;
- reuso da lógica/direção do viewer do `Classificar`.

### Não inclui
- segundo editor rico de grupo;
- duplicação do Gestor do Grupo;
- viewer novo paralelo.

### Dependências
- Fase 5 fechada ou suficientemente estável;
- regra de viewer fixada.

### Critério de fecho
O Explorador consulta bem e encaminha bem, mas não duplica edição rica.

---

## Fase 7 — Implementar o `Gestor do Grupo`

### Objetivo
Criar o editor rico único do grupo.

### Inclui
- estrutura aprovada do mockup;
- secções como Ficha, Pessoas, Emails, Documentos, Notas, Tarefas e Tabelas;
- preview em baixo, full-width;
- reuso da lógica/direção do viewer do `Classificar`;
- abertura a partir de `Classificar`;
- abertura a partir do `Explorador de Grupos`.

### Não inclui
- duplicação da mesma edição noutras vistas;
- criação de um editor rico concorrente dentro do Explorador.

### Dependências
- Fase 6 suficientemente estável;
- guardas conceptuais respeitadas.

### Critério de fecho
Toda a edição rica relevante do grupo acontece no Gestor e as outras vistas limitam-se a abrir/chamar esse editor.

---

## Fase 8 — Integração leve de Tarefas no contexto atual

### Objetivo
Consolidar o contrato mínimo de tarefas dentro do contexto de grupo sem abrir ainda o módulo principal.

### Inclui
- criação rápida de tarefa;
- atualização rápida de estado;
- apresentação coerente em `Explorar`, `Explorador` e `Gestor`;
- persistência mínima alinhada com o contrato definido.

### Não inclui
- aba principal `Tarefas`;
- exploração global transversal de tarefas;
- integração avançada com fontes externas.

### Dependências
- Fase 1 fechada;
- Fases 5 a 7 suficientemente maduras.

### Critério de fecho
Tarefas funcionam no contexto atual sem desviar o projeto para uma frente maior ainda não aprovada.

---

## Fase 9 — Módulo principal `Tarefas` (futuro)

### Objetivo
Abrir a frente transversal de tarefas fora do contexto exclusivo de Grupos.

### Estado
Fora do scope atual.

### Nota
Só deve avançar depois de Grupos v1, Explorador e Gestor estarem suficientemente estáveis.

---

## 7. Critérios gerais de fecho por fase

Cada fase só pode ser dada como fechada quando cumprir todos estes pontos:

1. respeita o baseline do canvas;
2. respeita as guardas conceptuais;
3. não invade responsabilidades de outra vista;
4. não duplica viewer ou edição rica;
5. foi validada pelo utilizador;
6. o handoff/documentação do projeto foi atualizado.

---

## 8. Regra de documentação contínua

Depois de cada fase concluída, atualizar no projeto:

- o estado da fase;
- o que entrou;
- o que ficou fora;
- riscos/resíduos conhecidos;
- próximos passos.

### Objetivo
Permitir mudança de agent sem regressão conceptual nem perda de contexto.

---

## 9. Protocolo de trabalho para qualquer agent

Antes de mexer no código, o agent deve:

1. ler `AGENTS.md`;
2. ler `docs/HANDOFF.md`;
3. ler `docs/CODE_REVIEW.md`;
4. ler este plano;
5. consultar o estado mais recente dos 3 mockups no canvas;
6. confirmar a fase em curso;
7. limitar-se ao scope dessa fase.

### Regra de ouro
Não reinventar arquitetura já fechada.

---

## 10. Recomendação de localização no projeto

Guardar este documento numa zona estável e fácil de encontrar, por exemplo:

- `docs/implementation/GRUPOS_V1_IMPLEMENTATION_PLAN.md`

ou, se preferirem concentrar tudo em roadmap funcional:

- `docs/roadmap/GRUPOS_V1_BUILD_MAP.md`

---

## 11. Resumo executivo

A ordem recomendada é:

1. congelar baseline e contratos;
2. implementar `Preparar`;
3. proteger persistência/cache;
4. ligar ao `Classificar`;
5. implementar `Explorar` no add-in;
6. implementar `Explorador de Grupos`;
7. implementar `Gestor do Grupo`;
8. integrar Tarefas em modo leve;
9. deixar módulo principal `Tarefas` para frente futura.

### Princípio central
O projeto avança já para implementação, mas sem abrir agora uma segunda frente grande de desenho do gestor principal de Tarefas.

### Regra final para agents
**Continua a partir daqui, mas trata os 3 mockups atuais do canvas como baseline v1. No Explorador, não transformar o bloco intermédio de “gestor” num segundo editor rico: o editor completo único é o Gestor do Grupo. O Explorador consulta e abre o Gestor; não duplica a edição.**