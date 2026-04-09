# Relatório base de implementação — Aba Grupos

**Projeto:** Inbox CRM Cockpit  
**Data:** 2026-04-08

## Leitura obrigatória para agents
- `AGENTS.md`
- `docs/HANDOFF.md`
- `docs/CODE_REVIEW.md`

## Finalidade
Fixar a estrutura funcional e visual da aba **Grupos**, com regras de armazenamento, fluxo, mockups aprovados, etapas e critérios de fecho.

## Princípios fechados
- A app decide, a base de dados guarda, o Outlook só reflete.
- Grupo, referência, ticket e etiqueta são coisas diferentes.
- 1 email = 1 grupo principal.
- Cache do add-in = só sessão e retoma de trabalho.
- Grupos prepara e organiza; Classificar fecha.

## Armazenamento
- **Cache do add-in:** seleções, filtros, estado da sessão, rascunho da operação.
- **Base principal escolhida pelo utilizador:** emails e anexos reais de trabalho.
- **Supabase:** só o que for promovido para lá.

### Modos previstos
1. Tudo no Supabase
2. Local neste PC
3. Local em pasta escolhida
4. Híbrido

### Regras
- Abrir um email não deve significar enviar logo tudo para o Supabase.
- Anexos acima de um limite configurável (exemplo de trabalho: 5 MB) pedem decisão.
- Payloads pobres não podem apagar dados bons.
- Ao mudar de aba, abrir outro email ou sair do contexto, a app deve gravar o progresso da sessão antes de mudar.

## Estrutura da aba Grupos
### Preparar
- Tem Email Âncora.
- Tem switches ultra compactos **Grupo** e **Filtros**, ambos off por defeito.
- Sub-vistas:
  - Lista
  - Anexos
  - Resumo

### Explorar
- Não mostra Email Âncora.
- Não mostra contexto vindo de Preparar.
- É consulta global.
- Cada card abre detalhe curto no próprio add-in.
- No fundo existe botão para abrir o explorador completo.

## Regras de UI
- Cada card de email pode expandir/retrair.
- Em retraído mostra só assunto, remetente, data e estado mínimo.
- Em expandido mostra labels, grupo, referências, ticket e anexos.
- Referência nunca aparece como grupo.
- Se um email mudar de grupo principal, deve haver aviso claro e opção de converter o grupo antigo em referência **só nesse email**.

## Notas e tarefas
- No detalhe curto de Explorar aparecem Notas e Tarefas.
- No add-in isto fica leve.
- Gestão pesada de tarefas fica para módulo próprio.

## Etapas
- E1 — Fechar modelo e settings
- E2 — Implementar Preparar
- E3 — Persistência segura
- E4 — Ligação ao Classificar
- E5 — Implementar Explorar
- E6 — Explorador completo externo

## Regra de fecho
Cada ronda só passa a fechada depois de aprovação final explícita do utilizador.

## Ficheiro principal para revisão visual
- `Relatorio_Aba_Grupos_Implementacao_2026-04-08.html`
