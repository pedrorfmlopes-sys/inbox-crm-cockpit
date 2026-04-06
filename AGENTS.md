# AGENTS.md

## Propósito prático do projeto
- Este repositório mantém um Outlook add-in com painel lateral que liga emails, contexto operacional, Odoo, documentos/anexos e assistentes de IA.
- O objetivo imediato não é acelerar novas features; é estabilizar segurança, verdade arquitetural e manutenção sem regressões.

## Papéis operacionais
- ChatGPT atua como gestor técnico/orientador: define direção, enquadra prioridades, valida coerência e decide o próximo passo.
- Agents externos atuam como executores por turno: analisam, implementam, validam e documentam dentro de um âmbito delimitado.
- Git é a fonte de verdade. Conversas ajudam, mas não substituem o estado real do repositório.

## Regras obrigatórias
- Antes de qualquer nova implementação, ler `AGENTS.md`, `docs/HANDOFF.md` e `docs/DECISIONS.md`.
- Nunca continuar com base apenas em contexto antigo de conversa.
- Confirmar sempre no repositório o que é facto, o que é hipótese e o que falta validar fora do repo.
- Não fazer mudanças amplas sem delimitar claramente o âmbito.
- Evitar regressões acima de tudo.
- Não mexer em zonas sensíveis sem necessidade explícita.
- Qualquer ronda com alterações relevantes deve atualizar `docs/HANDOFF.md` e, se aplicável, `docs/DECISIONS.md`.

## Prioridades atuais do projeto
1. Segurança
2. Verdade arquitetural / fontes de verdade
3. Consolidação estrutural
4. Performance / custos
5. Novas features / UX refinada

## Ordem de execução
1. Analisar primeiro
2. Planear segundo
3. Implementar terceiro
4. Validar quarto
5. Resumir e registar no handoff no fim

## Regras de saída
- Indicar o que foi alterado
- Indicar riscos
- Indicar validações realizadas
- Indicar próximos passos sugeridos
