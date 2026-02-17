# Divitek Outlook Odoo Add-in (MVP)

Este projeto é um *Outlook web add-in* (Office.js) com painel lateral (task pane) tipo “JIRA”, focado em:
- Ler o email selecionado no Outlook (assunto + remetente)
- Procurar o contacto no Odoo por email
- Criar uma lead no Odoo com 1 clique (MVP)

> Nota: Não consigo testar aqui no teu Outlook, mas o projeto está estruturado para funcionar em Outlook na Web e no New Outlook no Windows.

## Requisitos
- Node.js 18+ (recomendado 20+)
- Outlook na Web (recomendado para primeiro teste) ou New Outlook no Windows
- Acesso ao teu Odoo via URL (self-host ou Odoo Online)

## 1) Instalação
Na raiz do projeto:

```bash
npm install
```

Cria `.env` na pasta `server` (há um exemplo em `server/.env.example`).

## 2) Arrancar em DEV (HTTPS)
```bash
npm run dev
```

Isto levanta:
- UI (Vite/React) em https://localhost:5173
- API (Express) em http://localhost:7071

## 3) Sideload no Outlook (DEV)
Segue o guia oficial (aba “XML manifest”):  
https://learn.microsoft.com/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing

Usa o ficheiro:
- `manifest/manifest.xml`

Dica: começa por testar **Outlook na Web** (é o mais previsível em dev).

## 4) O que já faz (MVP)
- Mostra Subject + From
- Botão “Procurar no Odoo”
- Se não existir, botão “Criar Lead” (crm.lead) com o email e assunto

## 5) Próximos passos (quando quiseres)
- Editar lead/contacto (CRUD completo)
- Pipeline stages, atividades (mail.activity), anexar email ao chatter
- Autenticação por utilizador (SSO/OAuth) em vez de credenciais em env
- “Modo MailMaestro” (AI): sumarizar thread + rascunhos + reescrita + inserir no compose

## Referências
- Build your first Outlook add-in (Yo/VS): https://learn.microsoft.com/office/dev/add-ins/quickstarts/outlook-quickstart-yo
- Office.Mailbox API: https://learn.microsoft.com/javascript/api/outlook/office.mailbox
- Odoo External API (JSON-RPC): https://www.odoo.com/documentation
