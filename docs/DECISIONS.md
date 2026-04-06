# DECISIONS

| Decisão | Estado | Motivo | Impacto | Quando rever |
| --- | --- | --- | --- | --- |
| ChatGPT atua como gestor técnico/orientador | Ativa | Garantir continuidade estratégica, priorização e coerência entre rondas | A direção técnica não depende do contexto parcial de um executor isolado | Rever se o modelo de coordenação mudar |
| Agents externos trabalham por turno | Ativa | Reduzir conflito de contexto e tornar cada intervenção auditável | Cada ronda deve ter âmbito claro, validação própria e handoff explícito | Rever se houver workflow multi-agent concorrente formalizado |
| Git é a fonte de verdade | Ativa | Conversas e memórias antigas podem divergir do estado real do projeto | Toda a análise e implementação deve partir do repo e não de contexto informal | Nunca, salvo mudança total de processo |
| Antes de implementar, ler `AGENTS.md`, `docs/HANDOFF.md` e `docs/DECISIONS.md` | Ativa | Forçar alinhamento mínimo entre agentes | Reduz regressões por contexto incompleto | Rever apenas se estes ficheiros forem substituídos por outro mecanismo canónico |
| Cada ronda relevante deve atualizar o handoff | Ativa | Evitar perda de contexto operacional entre turnos | O estado do projeto fica documentado no próprio repo | Rever se surgir outra prática equivalente e melhor |
| Segurança e arquitetura vêm antes de features | Ativa | O diagnóstico atual mostra maior risco estrutural do que falta funcional | Prioriza estabilização antes de expansão | Rever quando a base técnica estiver consolidada |
| Nunca usar contexto antigo como verdade sem validar no repo | Ativa | Há histórico longo e risco real de informação desatualizada | Obriga a distinguir facto, hipótese e memória | Nunca |
| Qualquer hipótese sobre produção deve ser marcada como hipótese, não como facto | Ativa | O repositório não prova sozinho a configuração real de Render/Supabase/manifest | Melhora rigor técnico e evita decisões erradas | Rever se passar a existir acesso direto e estável aos ambientes |
| Evitar regressões e mudanças transversais não delimitadas | Ativa | O projeto tem ficheiros sensíveis, monolíticos e acoplamento elevado | Cada alteração deve ser cirúrgica, com raio de impacto controlado | Nunca |
| Não mexer em zonas sensíveis sem necessidade explícita | Ativa | `office.ts`, `CockpitProvider`, `linkStore`, `index.js`, Odoo e manifest concentram risco alto | Aumenta prudência em áreas que podem quebrar múltiplos fluxos | Rever quando essas zonas forem consolidadas |
| Confirmar primeiro a verdade arquitetural antes de otimizações profundas | Provisória | Ainda há dúvidas relevantes sobre produção, Postgres/Supabase e fontes de verdade | Evita otimizar ou refatorar com premissas erradas | Rever após validação de produção |
| Alterações documentais também devem refletir a estratégia atual do projeto | Ativa | A documentação passa a ser parte do workflow operacional | Handoff e checklist tornam-se instrumentos de trabalho, não anexos decorativos | Rever se o processo documental deixar de ser canónico |
