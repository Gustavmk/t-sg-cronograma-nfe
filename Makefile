# =============================================================================
# Automação de Emails — Cronograma de NF
# =============================================================================
#
# Windows: instale o Make via winget antes de usar este arquivo
#   winget install GnuWin32.Make
#
# Linux/macOS: make já está disponível nativamente
#
# Uso básico:
#   make setup
#   make dryrun  CSV=destinatarios.csv
#   make run     CSV=destinatarios.csv
#
# Parâmetros opcionais (sobrescrevem padrões):
#   CSV=arquivo.csv   Arquivo de destinatários  (padrão: destinatarios_exemplo.csv)
#   MES=5             Mês de referência 1-12    (padrão: mês atual)
#   ANO=2026          Ano de referência         (padrão: ano atual)
#   ASSUNTO="..."     Assunto do email          (padrão: solicitado interativamente)
# =============================================================================

.DEFAULT_GOAL := help

VENV     := venv
CSV      ?= destinatarios_exemplo.csv
MES      ?=
ANO      ?=
ASSUNTO  ?=

# ── Caminhos: Windows vs Unix ─────────────────────────────────────────────────
ifeq ($(OS),Windows_NT)
    PY          := python
    PYTHON      := $(VENV)\Scripts\python.exe
    PIP         := $(VENV)\Scripts\pip.exe
    VENV_MARKER := $(VENV)\Scripts\python.exe
    RM_VENV     := rmdir /s /q $(VENV)
else
    PY          := python3
    PYTHON      := $(VENV)/bin/python
    PIP         := $(VENV)/bin/pip
    VENV_MARKER := $(VENV)/bin/python
    RM_VENV     := rm -rf $(VENV)
endif

# ── Flags opcionais passadas ao script ───────────────────────────────────────
SCRIPT_FLAGS :=
ifneq ($(MES),)
    SCRIPT_FLAGS += --mes $(MES)
endif
ifneq ($(ANO),)
    SCRIPT_FLAGS += --ano $(ANO)
endif
ifneq ($(ASSUNTO),)
    SCRIPT_FLAGS += --assunto "$(ASSUNTO)"
endif

# =============================================================================
# Etapa 1 — Setup
# =============================================================================

.PHONY: setup
setup: ## [Etapa 1] Valida Python 3.12, cria/valida venv e instala dependências
	@echo ""
	@echo "============================================================"
	@echo " ETAPA 1 — SETUP"
	@echo "============================================================"
	@echo ""
	@echo "--> [1/3] Verificando Python 3.12..."
	@$(PY) --version
	@$(PY) -c "import sys; v=sys.version_info; sys.exit(0) if (v.major,v.minor)>=(3,12) else (print('[ERRO] Python 3.12+ obrigatório. Instale via:'), print('       Windows: winget install Python.Python.3.12'), print('       Linux:   sudo apt install python3.12'), sys.exit(1))"
	@echo ""
	@echo "--> [2/3] Verificando venv..."
	@$(PY) -c "\
import os, subprocess, sys; \
exists = os.path.exists(r'$(VENV_MARKER)'); \
print('    Venv já existe. Reutilizando.' if exists else '    Criando venv...'); \
subprocess.run([sys.executable, '-m', 'venv', '$(VENV)'], check=True) if not exists else None \
"
	@echo ""
	@echo "--> [3/3] Instalando/atualizando dependências..."
	@$(PIP) install --upgrade pip --quiet
	@$(PIP) install -r requirements.txt --quiet
	@echo "    Dependências instaladas."
	@echo ""
	@echo "============================================================"
	@echo " Setup concluído!"
	@echo " Próximo passo: make dryrun CSV=$(CSV)"
	@echo "============================================================"
	@echo ""

# =============================================================================
# Etapa 2 — Dry-run
# =============================================================================

.PHONY: dryrun
dryrun: _check-venv ## [Etapa 2] Preview sem enviar (não autentica, não envia)
	@echo ""
	@echo "============================================================"
	@echo " ETAPA 2 — DRY-RUN"
	@echo "============================================================"
	@echo ""
	@echo "--> Carregando venv e executando preview..."
	@echo ""
	$(PYTHON) send_cronograma.py --csv $(CSV) $(SCRIPT_FLAGS) --dry-run
	@echo ""
	@echo "============================================================"
	@echo " Dry-run concluído."
	@echo " Se o preview estiver correto: make run CSV=$(CSV)"
	@echo "============================================================"
	@echo ""

# =============================================================================
# Etapa 3 — Execução real
# =============================================================================

.PHONY: run
run: _check-venv ## [Etapa 3] Executa o envio real (autenticação + confirmação)
	@echo ""
	@echo "============================================================"
	@echo " ETAPA 3 — ENVIO REAL"
	@echo "============================================================"
	@echo ""
	@echo "--> Carregando venv e iniciando script de envio..."
	@echo ""
	$(PYTHON) send_cronograma.py --csv $(CSV) $(SCRIPT_FLAGS)
	@echo ""
	@echo "============================================================"
	@echo " Execução finalizada. Confira o log em send_log.jsonl"
	@echo "============================================================"
	@echo ""

# =============================================================================
# Utilitários
# =============================================================================

.PHONY: clean
clean: ## Remove venv e arquivos de cache
	@echo "--> Removendo venv..."
	@$(RM_VENV) 2>/dev/null || true
	@find . -type d -name __pycache__ -exec rm -rf {} + 2>/dev/null || true
	@find . -name "*.pyc" -delete 2>/dev/null || true
	@echo "    Limpo."

.PHONY: help
help: ## Exibe esta ajuda
	@echo ""
	@echo "Automação de Emails — Cronograma de NF"
	@echo ""
	@echo "Uso: make <alvo> [CSV=arquivo.csv] [MES=5] [ANO=2026] [ASSUNTO='...']"
	@echo ""
	@echo "Alvos:"
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) \
		| grep -v '^_' \
		| awk 'BEGIN {FS=":.*?## "}; {printf "  \033[36m%-10s\033[0m %s\n", $$1, $$2}'
	@echo ""
	@echo "Exemplos:"
	@echo "  make setup"
	@echo "  make dryrun  CSV=lista_maio.csv"
	@echo "  make run     CSV=lista_maio.csv"
	@echo "  make run     CSV=lista_maio.csv  MES=5  ANO=2026"
	@echo "  make run     CSV=lista_maio.csv  ASSUNTO=\"Cronograma Maio/2026\""
	@echo ""

# ── Guardas internas ─────────────────────────────────────────────────────────

.PHONY: _check-venv
_check-venv:
	@$(PY) -c "import os, sys; \
e=os.path.exists(r'$(VENV_MARKER)'); \
(print(''), print('[ERRO] Venv não encontrado. Execute primeiro: make setup'), print('')) \
if not e else None; sys.exit(0 if e else 1)"
