# CODE REVIEW

## Objetivo
- Usar este checklist em qualquer PR, diff ou commit relevante.
- A revisão deve ser prática e orientada a risco.
- Distinguir sempre:
  - confirmado pelo repositório
  - provável mas não confirmado em produção

## Checklist geral
- [ ] O âmbito da alteração está delimitado e coerente com a tarefa?
- [ ] A mudança evita regressões óbvias e não toca em áreas sensíveis sem necessidade?
- [ ] O output final explica alterações, riscos, validações e próximos passos?
- [ ] `docs/HANDOFF.md` foi atualizado se a ronda mudou o estado operacional?
- [ ] `docs/DECISIONS.md` foi atualizado se a ronda alterou normas, decisões ou direção?

## Segurança backend
- [ ] Há novas rotas ou mudanças de rotas? Estão protegidas de forma consistente?
- [ ] O impacto em CORS foi revisto?
- [ ] O impacto em autenticação e autorização foi revisto?
- [ ] Há risco novo de abuso de endpoint, exfiltração ou escalada de privilégio?
- [ ] O tamanho de payloads, uploads e parsing foi revisto?

## Segredos e credenciais
- [ ] Há segredos, tokens, passwords ou chaves novas no cliente?
- [ ] Há logs, debug output ou ficheiros auxiliares a registar dados sensíveis?
- [ ] Alguma variável de ambiente sensível ficou exposta em código, bundle ou output?

## Impacto em Odoo
- [ ] A alteração mexe em autenticação, sessão, rotas, métodos ou modelos permitidos?
- [ ] A alteração muda payloads, schema filtering, writes ou ligação de emails a registos?
- [ ] O risco de acesso indevido, duplicação ou inconsistência com Odoo foi avaliado?
- [ ] Se houver hipótese sobre comportamento em produção/Odoo real, está marcada como hipótese?

## Impacto em Outlook add-in
- [ ] A alteração toca em manifest, permissões, command surfaces ou task pane?
- [ ] Há impacto em read mode, compose mode, categorias, anexos ou `ItemChanged`?
- [ ] A compatibilidade com novo Outlook vs clássico foi pelo menos considerada?
- [ ] O requirement set usado continua coerente com as APIs chamadas?

## Persistência e fontes de verdade
- [ ] A alteração cria ou muda uma fonte de verdade?
- [ ] O impacto em Postgres, JSON local, storage local/OneDrive, `localStorage` ou RoamingSettings foi revisto?
- [ ] Há duplicação nova de dados, reingestão desnecessária ou risco de divergência?
- [ ] Está claro o que é persistido, o que é cache e o que é apenas derivado?

## Tráfego, polling e custos
- [ ] A alteração aumenta polling, refreshes, retries ou fetches redundantes?
- [ ] Há impacto em payloads grandes, anexos, previews ou downloads binários?
- [ ] O impacto provável em Render, OpenAI, Gemini, Odoo ou Postgres/Supabase foi considerado?
- [ ] Se houver suspeita de custo/egress, isso foi identificado como hipótese e não como facto?

## IA, prompts e dados sensíveis
- [ ] A alteração mexe em prompts, contexto, briefings, learning, modelos ou providers?
- [ ] O impacto em dados sensíveis enviados para IA foi avaliado?
- [ ] Há mudança de comportamento automático que possa aumentar custo, latência ou ruído?
- [ ] O fluxo principal e os fluxos legados continuam coerentes entre si?

## UX e fluxo operacional
- [ ] A alteração preserva o fluxo principal do utilizador?
- [ ] Há regressões em navegação, estado atual do email, compose, ligação a registos ou documentos?
- [ ] A experiência continua previsível em cenários com email sem anexos, múltiplos anexos e contexto parcial?

## Fecho da revisão
- [ ] O diff final continua dentro do âmbito combinado?
- [ ] O que está confirmado no repo foi separado do que depende de validação em produção?
- [ ] A recomendação final está alinhada com a prioridade atual: segurança e coerência arquitetural antes de novas features?
