# Auditoria de Segurança e Fluxo de Dados

Este documento detalha o estado técnico do projeto InboxCockpit à data de Abril de 2026, focando-se em fronteiras de segurança, soberania de dados e riscos estruturais.

## 1. Superfície de Segurança (Backend)

### 1.1 CORS e Exposição
- **Configuração**: `app.use(cors())` em `server/src/index.js`.
- **Estado**: **Aberto (High Risk)**. Aceita pedidos de qualquer origem.
- **Impacto**: Suscetível a ataques de Cross-Site Request Forgery (CSRF) se não houver proteção adicional rígida nas rotas.

### 1.2 Autenticação e Sessões
- **Mecanismo**: `SessionManager` (in-memory).
- **Tokens**: Strings aleatórias geradas no login, sem expiração explícita ou rotação (apenas validade de uptime do servidor).
- **Risco**: Reinícios de servidor invalidam todas as sessões, forçando re-auth.

### 1.3 Gestão de Segredos e Credenciais
- **Client-Side (Critical Risk)**: O ficheiro `client/src/settings.ts` confirma que `odooPassword`, `geminiApiKey` e `openaiApiKey` são persistidos em `RoamingSettings` ou `localStorage`.
- **Server-Side**: Depende de `.env` para o Odoo global. O `getOdooCached` faz fallback para a credencial global se a sessão falhar, o que é um risco de elevação de privilégios.

## 2. Matriz de Dados e Fontes de Verdade

| Tipo de Dado | Origem (SSOT) | Persistência Adicional | Risco de Divergência | Sensibilidade |
| :--- | :--- | :--- | :--- | :--- |
| **Entidades Odoo** | Odoo | Cache em Memória (Server) | Média (Sync manual) | Alta |
| **Links Email-Record** | linkStore | Postgres / `links.json` | Baixa | Média |
| **Grupos & Docs Studio** | linkStore | Postgres / Disco Local | Média (FS vs DB) | Alta |
| **Settings & Secrets** | Cliente | `RoamingSettings` | N/A (Client-only) | **Crítica** |
| **Perfis IA / Estilo** | learningStore | Postgres / `learning.json` | Baixa | Média |
| **Categorias Outlook** | Outlook | `localStorage` (Sync) | Alta (Race conditions) | Média |

## 3. Top 5 Riscos Técnicos Imediatos

| Risco | Localização | Impacto | Prioridade |
| :--- | :--- | :--- | :--- |
| **Exposição de Segredos** | `client/src/settings.ts` | Compromisso total de Odoo/AI | **Urgente** |
| **CORS Hiper-Permissivo** | `server/src/index.js:61` | Ataques cross-origin | Alta |
| **Fallback Global Odoo** | `server/src/index.js:272` | Acesso não autorizado via Env | Alta |
| **Gigantismo (Studio)** | `GroupClassificationStudioApp.tsx` | Regressões em UX crítica | Média |
| **Persistência Volátil** | `sessionManager.js` | Perda de sessões e UX pobre | Baixa |

## 4. Ordem Recomendada de Intervenção (Sem Regressões)

1. **Fase 1: Vaulting (Segurança)**: Mudar a persistência de segredos (API Keys, Passwords) para o Server-side (Postgres/Supabase) e usar apenas tokens de sessão no client.
2. **Fase 2: Perímetro**: Restringir CORS no backend apenas ao domínio de produção e localhost.
3. **Fase 3: Refactor de Persistência**: Migrar obrigatoriamente para PostgreSQL em produção para garantir durabilidade de sessões e links.
4. **Fase 4: Sharding de Código**: Dividir `linkStore.js` e `GroupClassificationStudioApp.tsx` em submódulos funcionais menores.

---
*Documento gerado como deliverable da ronda de auditoria técnica.*
