# Divitek Outlook Odoo Add-in (InboxCockpit)

Este projeto é um *Outlook web add-in* (Office.js) com painel lateral (task pane) focado na integração entre Outlook e Odoo.

## 🚀 Ambiente Staging (Render)
Staging is hosted on Render as a unified service (API + UI).
- **URL**: `https://inbox-cockpit-staging.onrender.com`
- **Health Check**: `https://inbox-cockpit-staging.onrender.com/health`
- **Manifest**: Use `manifest/manifest.staging.xml`.

### Health Check (Unified)
```bash
# Verify API and UI are live
curl https://inbox-cockpit-staging.onrender.com/health
```

### Configuração no Render
No painel do Render, deves configurar as seguintes Environment Variables:
- `AI_ENABLED`: `0` (default) ou `1`.
- `OPENAI_API_KEY`: A tua chave (só necessária se `AI_ENABLED=1`).
- `ODOO_URL`, `ODOO_DB`, `ODOO_USER`, `ODOO_PASS`: Credenciais do Odoo.

## 🛠️ Desenvolvimento Local

### 1) Instalação
Na raiz do projeto:
```bash
npm install
```
Cria `.env` na pasta `server` (copia de `server/.env.example`).

Se o browser/Outlook mostrar `ERR_CERT_DATE_INVALID` ou um aviso de privacidade para `https://localhost:5173`, renova os certificados locais:
```bash
npm run certs:verify
npm run certs:install
```

### 2) Execução (HTTPS Local)
```bash
npm run dev
```
- **UI**: https://localhost:5173
- **API**: http://localhost:7071

### 3) Manifestos
Existem dois manifestos principais na pasta `manifest/`:
- `manifest.dev.xml`: Aponta para `https://localhost:5173` (para desenvolvimento).
- `manifest.staging.xml`: Aponta para o URL do Render (para testes reais).

## 📥 Instalação (Sideload)

### Outlook na Web / Novo Outlook
1. Abre o Outlook e vai a "Get Add-ins" ou "Manage Add-ins".
2. Escolhe "My add-ins" -> "Add a custom add-in" -> "Add from file...".
3. Seleciona o manifesto pretendido (`dev` ou `staging`).

### Outlook Classic (Desktop)
1. Segue o mesmo processo via Outlook na Web (a conta sincroniza o add-in para o desktop).
2. Se necessário, usa o botão "Sideload" no separador "File" -> "Manage Add-ins".

## ✅ Validação Rápida
- **Check Health**: `curl http://localhost:7071/health`
- **Check AI**: `curl http://localhost:7071/api/ai/selftest` (se `AI_ENABLED=1`)

---
### Referências
- [Office.js API](https://learn.microsoft.com/javascript/api/outlook/office.mailbox)
- [Odoo External API](https://www.odoo.com/documentation)
