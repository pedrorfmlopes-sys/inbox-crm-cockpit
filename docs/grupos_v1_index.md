# Grupos v1 — Baseline Canonica

Data: 2026-04-09

## Objetivo

- Fixar a baseline documental de Grupos v1 que deve ser usada por qualquer agent antes de abrir Fase 2+.
- Tornar navegavel no GitHub a combinacao de plano, mockups/exportacoes e contratos minimos desta frente.

## Ordem de precedencia

1. `docs/plano_implementacao_grupos_v1.md`
2. mockups/exportacoes mais recentes de 2026-04-09:
   - `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-09_com_Explorador_e_Screenshots.md`
   - `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-09_com_Explorador_e_Screenshots.html`
   - `docs/Relatorio_Gestor_do_Grupo_Mockup_2026-04-09_v2_embedded.md`
   - `docs/Relatorio_Gestor_do_Grupo_Mockup_2026-04-09_v2_embedded.html`
3. `docs/grupos_v1_fase1_contratos.md`
4. docs de suporte 2026-04-08, quando nao houver conflito:
   - `docs/groups-tab-summary-2026-04-08.md`
   - `docs/groups-email-persistence-model-2026-04-08.md`
   - `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-08.md`
   - `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-08.html`

## Distincao obrigatoria

- `Explorar` dentro da aba `Grupos`: consulta rapida global no add-in, sem `Email Ancora` e sem contexto herdado do email atual.
- `Explorador de Grupos`: vista maior de consulta e acompanhamento; consulta e abre o Gestor.
- `Gestor do Grupo`: editor rico unico.
- O bloco intermédio tipo "gestor" no mockup do Explorador e ponte conceptual/ponto de entrada, nao um segundo editor rico.
- O preview/viewer reutiliza a direcao do `Classificar`: em baixo, full-width, tanto no Explorador como no Gestor.

## Documentos canonicos desta ronda

- Plano: `docs/plano_implementacao_grupos_v1.md`
- Contratos de Fase 1: `docs/grupos_v1_fase1_contratos.md`
- Relatorio visual Grupos + Explorador: `docs/Relatorio_Aba_Grupos_Implementacao_2026-04-09_com_Explorador_e_Screenshots.md`
- Relatorio visual Gestor: `docs/Relatorio_Gestor_do_Grupo_Mockup_2026-04-09_v2_embedded.md`
- Assets de screenshots para GitHub: `docs/groups_report_assets/`

## Regras conceptuais que esta baseline protege

- `Preparar` nao substitui `Classificar`.
- `Explorar` em `Grupos` e global e independente.
- `Explorador de Grupos` nao duplica edicao rica.
- `Gestor do Grupo` e o unico editor completo.
- `grupo`, `referencia`, `ticket` e `etiqueta` mantem semanticas separadas.
- `1 email = 0 ou 1 grupo principal`.
- cache/progresso de sessao e persistencia remota sao camadas diferentes.

## Scopes fora desta ronda

- Fase 2+ de UI pesada.
- segundo editor rico no Explorador.
- novo viewer paralelo.
- aba principal `Tarefas`.
