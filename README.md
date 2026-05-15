# Automação de Emails — Cronograma de NF

Script Python para envio automático de cronograma de pagamento de Nota Fiscal via **Microsoft Graph API (Office 365)**.

---

## Pré-requisitos

### 1. Python 3.11+

```bash
python --version
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
  ]
}
```

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

### Fluxo de execução

1. Lê e valida o CSV
2. Exibe **preview** com todos os emails e datas calculadas
3. Pede **confirmação** antes de enviar
4. Na primeira execução, abre o **device code flow** para autenticação (visitar URL + inserir código no browser)
5. Token é cacheado — nas próximas execuções, não é necessário autenticar novamente
6. Envia com **retry automático** (3 tentativas com backoff exponencial)
7. Salva log em `send_log.jsonl`

---

## Formato do CSV

Delimitador: `;` (ponto e vírgula)

```csv
Nome;Email;ValorNF
João Silva;joao.silva@company.co;R$ 10.000,00
Maria Pereira;mariap@company.co;R$ 15.000,00
```

---

## Arquivos gerados

| Arquivo | Descrição |
|---|---|
| `.token_cache.bin` | Cache do token de autenticação (não commitar) |
| `send_log.jsonl` | Log de todos os envios com status e timestamp |

---

## Segurança

- **Não commitar** `config.json` nem `.token_cache.bin` — adicione ao `.gitignore`
- O token é armazenado localmente com escopo restrito (`Mail.Send`, `User.Read`)
- Emails são enviados **em nome do usuário autenticado** (permissão delegada)


---

## Automação

### make setup

1. Verifica se Python 3.12+ está instalado (imprime mensagem com o comando de instalação se não estiver: winget install Python.Python.3.12)
2. Verifica se o venv já existe — reutiliza se sim, cria se não
3. Instala/atualiza todas as dependências do requirements.txt

### make dryrun CSV=arquivo.csv

1. Carrega o venv e executa o script com --dry-run
2. Exibe o preview completo (destinatários, datas calculadas, assunto) sem autenticar nem enviar nada
3. Bloqueia se o venv não existir, pedindo para rodar make setup

### make run CSV=arquivo.csv

1. Carrega o venv e executa o envio real
2. Fluxo normal: preview → confirmação manual → autenticação → envio com retry

#### Parâmetros opcionais disponíveis

make run CSV=lista.csv MES=5 ANO=2026
make run CSV=lista.csv ASSUNTO="Cronograma Maio/2026"

#### Instalação do Make no Windows

winget install GnuWin32.Make
