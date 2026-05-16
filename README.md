# Automação de Emails — Cronograma de NF

Script Python para envio automático de cronograma de pagamento de Nota Fiscal via **Microsoft Graph API (Office 365)**.

---

## Pré-requisitos

### Antes de iniciar

```powershell
winget install Microsoft.VisualStudioCode
winget install --id Git.Git -e --source winget
winget install -e --id Python.Python.3.12
winget install GnuWin32.Make

# Após instalar os requisitos para clone do repositório, execute em um novo terminal e encerre o antigo
git clone https://github.com/Gustavmk/t-sg-cronograma-nfe
```

### 1. Python 3.12+

```bash
python --version
python -m venv .venv
.venv\Scripts\activate

```

### 2. Instalar dependências

```bash
pip install -r requirements.txt
```

### 3. Registrar aplicativo no Azure AD

1. Acesse o [Azure Portal](https://portal.azure.com) → **Azure Active Directory** → **Registros de aplicativo** → **Novo registro**
2. Nome: `Cronograma NF Bot` (ou qualquer nome)
3. Tipo de conta: **Somente contas neste diretório organizacional**
4. URI de Redirecionamento: deixe em branco (usaremos device code flow)
5. Após criar, anote o **Application (client) ID** e o **Directory (tenant) ID**
6. No menu lateral: **Permissões de API** → **Adicionar permissão** → **Microsoft Graph** → **Permissões delegadas**
   - Adicionar: `Mail.Send` e `User.Read`
7. Clique em **Conceder consentimento do administrador**

10. Libere isFallBackPublicClient

Utilizando azure shell: 
```bash
ID='<Valor do ApplicationID>'
az ad app show \
  --id $ID \
  --query id -o tsv

# Libere a propriedade
az ad app update \
  --id $ID \
  --set isFallbackPublicClient=true

# Valide se está ativo
az ad app show \
  --id $ID \
  --query isFallbackPublicClient
```

### 4. Configurar o script

```bash
cp config.example.json config.json
```

Edite `config.json`:

```json
{
  "tenant_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "cc_emails": [
    "financeiro@suaempresa.com"
  ],
  "notification_email": "financeiro@suaempresa.com"
}
```

| Campo | Obrigatório | Descrição |
|---|---|---|
| `tenant_id` | ✅ | ID do diretório Azure AD |
| `client_id` | ✅ | ID do aplicativo registrado |
| `cc_emails` | ✅ | Lista de emails sempre em cópia |
| `notification_email` | ➖ | Recebe o relatório pós-envio. Omita para não enviar. |

---

## Uso

```bash
python send_cronograma.py --csv destinatarios.csv
```

### Opções disponíveis

| Opção | Descrição | Exemplo |
|---|---|---|
| `--csv` | **(obrigatório)** Arquivo CSV com os destinatários | `--csv lista_maio.csv` |
| `--assunto` | Assunto do email (perguntado se omitido) | `--assunto "Cronograma Maio/2026"` |
| `--mes` | Mês de referência, 1–12 (padrão: mês atual) | `--mes 5` |
| `--ano` | Ano de referência (padrão: ano atual) | `--ano 2026` |
| `--dry-run` | Preview sem autenticar ou enviar | `--dry-run` |
| `--force` | Reenvia mesmo emails já enviados no mês | `--force` |

### Fluxo de execução

1. Lê e valida o CSV (formato, colunas, emails)
2. Verifica **registros MX** dos domínios de email via DNS
3. Consulta o log (`send_log.jsonl`) para identificar emails **já enviados** no mesmo mês
4. Exibe **preview** com status de cada destinatário (a enviar / já enviado)
5. Pede **confirmação** antes de enviar
6. Na primeira execução, abre o **device code flow** para autenticação (browser)
7. Token é cacheado — execuções seguintes não pedem autenticação novamente
8. Envia com **retry automático** (3 tentativas, backoff exponencial)
9. Salva log em `send_log.jsonl`
10. Envia **email de relatório** para `notification_email` com resumo de sucesso/falha

---

## Formato do CSV

Delimitador: `;` (ponto e vírgula)

```csv
Nome;Email;ValorNF
João Silva;joao.silva@company.co;R$ 10.000,00
Maria Pereira;mariap@company.co;R$ 15.000,00
```

---

## Template de email

O corpo do email é definido em [`email_template.html`](email_template.html). Edite este arquivo para ajustar o texto sem precisar alterar o código Python.

### Placeholders disponíveis

| Placeholder | Descrição | Exemplo |
|---|---|---|
| `{{MesAtual}}` | Mês de referência por extenso | `Maio` |
| `{{AnoVigente}}` | Ano de referência | `2026` |
| `{{MesSeguinte}}` | Mês seguinte por extenso | `Junho` |
| `{{AnoMesSeguinte}}` | Ano do mês seguinte | `2026` |
| `{{QuintoDiaUtil}}` | 5º dia útil do mês seguinte (feriados BR) | `07` |
| `{{ValorNF}}` | Valor da NF do destinatário (do CSV) | `R$ 10.000,00` |

---

## Idempotência

O script verifica `send_log.jsonl` antes de cada envio. Se um email já foi enviado com sucesso para o mesmo destinatário no mesmo mês de referência, ele é **automaticamente pulado** para evitar duplicatas.

Para forçar reenvio:

```bash
python send_cronograma.py --csv lista.csv --force
```

---

## Validação de MX

Antes de enviar, o script resolve o registro DNS MX de cada domínio de email. Domínios sem MX recebem um aviso no terminal. O envio não é abortado — o email é tentado mesmo assim — mas o aviso permite detectar typos antes de disparar o envio real.

---

## Arquivos gerados

| Arquivo | Descrição |
|---|---|
| `.token_cache.bin` | Cache do token de autenticação (não commitar) |
| `send_log.jsonl` | Log de todos os envios com status e timestamp |

---
## Segurança

- **Não commitar** `config.json` nem `.token_cache.bin` — já incluídos no `.gitignore`
- O token é armazenado localmente com escopo restrito (`Mail.Send`, `User.Read`)
- Emails são enviados **em nome do usuário autenticado** (permissão delegada)