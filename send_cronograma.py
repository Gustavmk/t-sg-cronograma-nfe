#!/usr/bin/env python3
"""
Automação de envio de cronograma de NF via Microsoft Graph API.

Uso:
    python send_cronograma.py --csv lista.csv
    python send_cronograma.py --csv lista.csv --assunto "Meu Assunto" --mes 5 --ano 2026
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import holidays
import msal
import requests

CONFIG_PATH = Path(__file__).parent / "config.json"
TOKEN_CACHE_PATH = Path(__file__).parent / ".token_cache.bin"
LOG_PATH = Path(__file__).parent / "send_log.jsonl"

MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

GRAPH_SCOPES = ["https://graph.microsoft.com/Mail.Send", "https://graph.microsoft.com/User.Read"]


# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        print(f"[ERRO] config.json não encontrado em {CONFIG_PATH}")
        print("       Copie config.example.json para config.json e preencha os valores.")
        sys.exit(1)
    with open(CONFIG_PATH, encoding="utf-8") as f:
        cfg = json.load(f)
    required = ["tenant_id", "client_id"]
    missing = [k for k in required if not cfg.get(k)]
    if missing:
        print(f"[ERRO] Campos obrigatórios ausentes em config.json: {missing}")
        sys.exit(1)
    return cfg


# ---------------------------------------------------------------------------
# Authentication (Microsoft Graph — delegated, device code flow)
# ---------------------------------------------------------------------------

def get_access_token(config: dict) -> str:
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_PATH.exists():
        cache.deserialize(TOKEN_CACHE_PATH.read_text(encoding="utf-8"))

    app = msal.PublicClientApplication(
        client_id=config["client_id"],
        authority=f"https://login.microsoftonline.com/{config['tenant_id']}",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
            return result["access_token"]

    # First auth or token expired — use device code flow
    flow = app.initiate_device_flow(scopes=GRAPH_SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Falha ao iniciar device code flow: {flow}")

    print("\n" + "=" * 65)
    print(flow["message"])
    print("=" * 65 + "\n")

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"Falha na autenticação: {result.get('error_description', result)}")

    TOKEN_CACHE_PATH.write_text(cache.serialize(), encoding="utf-8")
    return result["access_token"]


def get_sender_email(token: str) -> str:
    resp = requests.get(
        "https://graph.microsoft.com/v1.0/me",
        headers={"Authorization": f"Bearer {token}"},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    return data.get("mail") or data.get("userPrincipalName")


# ---------------------------------------------------------------------------
# Date calculation
# ---------------------------------------------------------------------------

def nth_business_day(year: int, month: int, n: int) -> date:
    """Return the nth business day of a given month, skipping Brazilian holidays."""
    br_holidays = holidays.Brazil(years=year)
    d = date(year, month, 1)
    count = 0
    while True:
        if d.weekday() < 5 and d not in br_holidays:
            count += 1
            if count == n:
                return d
        d += timedelta(days=1)


def compute_dates(ref_date: date) -> dict:
    mes = ref_date.month
    ano = ref_date.year
    mes_seg = 1 if mes == 12 else mes + 1
    ano_seg = ano + 1 if mes == 12 else ano
    quinto_du = nth_business_day(ano_seg, mes_seg, 5)
    return {
        "MesAtual": MONTHS_PT[mes],
        "AnoVigente": str(ano),
        "MesSeguinte": MONTHS_PT[mes_seg],
        "AnoMesSeguinte": str(ano_seg),
        "QuintoDiaUtil": quinto_du.strftime("%d"),
    }


# ---------------------------------------------------------------------------
# CSV reading
# ---------------------------------------------------------------------------

def read_csv(csv_path: str) -> list[dict]:
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"Arquivo CSV não encontrado: {csv_path}")

    with open(path, encoding="utf-8-sig") as f:
        reader = csv.DictReader(f, delimiter=";")
        fieldnames = [c.strip() for c in (reader.fieldnames or [])]
        required = {"Nome", "Email", "ValorNF"}
        missing_cols = required - set(fieldnames)
        if missing_cols:
            raise ValueError(f"Colunas obrigatórias ausentes no CSV: {missing_cols}")

        recipients = []
        for i, row in enumerate(reader, start=2):
            row = {k.strip(): (v or "").strip() for k, v in row.items()}
            missing_vals = [k for k in required if not row.get(k)]
            if missing_vals:
                raise ValueError(f"Linha {i}: campos obrigatórios vazios: {missing_vals}")
            if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", row["Email"]):
                raise ValueError(f"Linha {i}: endereço de email inválido: '{row['Email']}'")
            recipients.append(row)

    if not recipients:
        raise ValueError("CSV não contém dados válidos (verifique se há linhas após o cabeçalho)")

    return recipients


# ---------------------------------------------------------------------------
# Email rendering
# ---------------------------------------------------------------------------

def render_html(recipient: dict, dates: dict) -> str:
    d = dates
    pagamento = f"{d['QuintoDiaUtil']}/{d['MesSeguinte']}/{d['AnoMesSeguinte']}"
    valor_nf = recipient["ValorNF"]
    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<body style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:680px;line-height:1.6;">

<p>Boa tarde,</p>

<p>Segue cronograma referente ao mês de <strong>{d['MesAtual']}</strong> de <strong>{d['AnoVigente']}</strong>.</p>

<p>Valor da NF <strong>{valor_nf}</strong></p>

<p>Enviar NF e documentos para <a href="mailto:dp@twygo.com">dp@twygo.com</a></p>
<ul>
  <li>Cópia do pró-labore (MEI está isento).</li>
  <li>Comprovante de pagamento dos tributos (INSS pago) + Certidão Negativa de Débitos Federais</li>
</ul>

<table border="1" cellpadding="10" cellspacing="0"
       style="border-collapse:collapse;width:100%;max-width:520px;border-color:#ccc;">
  <thead>
    <tr style="background-color:#f0f0f0;">
      <th style="text-align:left;">Cronograma Ref. {d['MesAtual']}/{d['AnoVigente']}</th>
    </tr>
  </thead>
  <tbody>
    <tr><td>Envio da Nota Fiscal até: 22/{d['MesAtual']}/{d['AnoVigente']} – 18:00</td></tr>
    <tr><td>Envio de documentos até: 22/{d['MesAtual']}/{d['AnoVigente']} – 18:00</td></tr>
    <tr><td>Data de pagamento: {pagamento}</td></tr>
  </tbody>
</table>

<br>
<p>Reforçamos a importância de que o envio da Nota Fiscal e dos demais documentos siga
rigorosamente o cronograma estabelecido acima.</p>

<p>Conforme disposto na cláusula <strong>2.7</strong> do contrato, destacamos que:</p>

<blockquote style="border-left:3px solid #ccc;margin:0 0 0 16px;padding:8px 16px;
                   color:#555;font-style:italic;">
  Ocorrendo atraso, pela CONTRATADA, na entrega da fatura e respectiva Nota Fiscal nos termos
  do item "h" do Preâmbulo, caso a emissão da nota fiscal ocorra a partir do dia 21 do mês da
  prestação de serviço, a data de pagamento será dia 20 do mês subsequente à prestação de serviço.
</blockquote>

<br>
<p>Portanto, o não cumprimento dos prazos acarretará o replanejamento da data de pagamento
conforme previsto contratualmente.</p>

<p>Atenciosamente,</p>

</body>
</html>"""


# ---------------------------------------------------------------------------
# Email sending
# ---------------------------------------------------------------------------

def send_email(token: str, to_email: str, cc_list: list[str], subject: str, html_body: str) -> None:
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html_body},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
            "ccRecipients": [{"emailAddress": {"address": cc}} for cc in cc_list if cc],
        },
        "saveToSentItems": True,
    }
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    resp = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        json=payload,
        headers=headers,
        timeout=30,
    )
    resp.raise_for_status()


def send_with_retry(
    token: str,
    to_email: str,
    cc_list: list[str],
    subject: str,
    html_body: str,
    retries: int = 3,
) -> None:
    for attempt in range(1, retries + 1):
        try:
            send_email(token, to_email, cc_list, subject, html_body)
            return
        except requests.HTTPError as exc:
            if attempt == retries:
                raise
            wait = 2 ** attempt
            print(f"\n    [aviso] tentativa {attempt}/{retries} falhou ({exc}). Aguardando {wait}s...")
            time.sleep(wait)


# ---------------------------------------------------------------------------
# Preview & UX
# ---------------------------------------------------------------------------

def print_preview(recipients: list[dict], dates: dict, subject: str, cc_list: list[str]) -> None:
    sep = "=" * 72
    print(f"\n{sep}")
    print("  PREVIEW — EMAILS A SEREM ENVIADOS")
    print(sep)
    print(f"  Assunto    : {subject}")
    print(f"  Ref.       : {dates['MesAtual']}/{dates['AnoVigente']}")
    print(f"  Pagamento  : {dates['QuintoDiaUtil']}/{dates['MesSeguinte']}/{dates['AnoMesSeguinte']}")
    if cc_list:
        print(f"  CC (todos) : {', '.join(cc_list)}")
    print(f"\n  {'#':<4} {'Nome':<26} {'Email':<36} {'Valor NF'}")
    print("  " + "-" * 72)
    for i, r in enumerate(recipients, 1):
        print(f"  {i:<4} {r['Nome']:<26} {r['Email']:<36} {r['ValorNF']}")
    print(sep)
    print(f"\n  Total: {len(recipients)} email(s)\n")


def ask_subject(dates: dict) -> str:
    default = f"Cronograma de Pagamento - {dates['MesAtual']}/{dates['AnoVigente']}"
    print(f"Assunto do email (Enter para: \"{default}\"):")
    entered = input("  > ").strip()
    return entered or default


def confirm_send() -> bool:
    print("Confirma o envio dos emails? [s/N] ", end="", flush=True)
    return input().strip().lower() == "s"


def append_log(entry: dict) -> None:
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(json.dumps(entry, ensure_ascii=False) + "\n")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(
        description="Envio de cronograma de NF via Microsoft Graph API (Office 365)"
    )
    parser.add_argument("--csv", required=True, metavar="ARQUIVO.csv",
                        help="Arquivo CSV com colunas Nome;Email;ValorNF")
    parser.add_argument("--assunto", metavar="TEXTO",
                        help="Assunto do email (solicitado interativamente se omitido)")
    parser.add_argument("--mes", type=int, choices=range(1, 13), metavar="1-12",
                        help="Mês de referência (padrão: mês atual)")
    parser.add_argument("--ano", type=int, metavar="AAAA",
                        help="Ano de referência (padrão: ano atual)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Exibe preview sem autenticar nem enviar emails")
    args = parser.parse_args()

    config = load_config()
    cc_list: list[str] = config.get("cc_emails", [])

    today = date.today()
    ref_date = date(args.ano or today.year, args.mes or today.month, 1)
    dates = compute_dates(ref_date)

    # --- Ler CSV ---
    print(f"\nLendo arquivo CSV: {args.csv}")
    try:
        recipients = read_csv(args.csv)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[ERRO] {exc}")
        sys.exit(1)
    print(f"  {len(recipients)} destinatário(s) encontrado(s).")

    # --- Assunto ---
    subject = args.assunto or ask_subject(dates)

    # --- Preview ---
    print_preview(recipients, dates, subject, cc_list)

    if args.dry_run:
        print("  [DRY-RUN] Nenhum email foi enviado. Remova --dry-run para enviar.")
        print("=" * 72)
        sys.exit(0)

    if not confirm_send():
        print("Envio cancelado.")
        sys.exit(0)

    # --- Autenticar ---
    print("\nAutenticando com Microsoft Graph API...")
    try:
        token = get_access_token(config)
        sender = get_sender_email(token)
        print(f"  Remetente: {sender}\n")
    except Exception as exc:
        print(f"[ERRO] Falha na autenticação: {exc}")
        sys.exit(1)

    # --- Enviar ---
    print(f"Enviando {len(recipients)} email(s)...")
    results = []
    for i, recipient in enumerate(recipients, 1):
        label = f"[{i}/{len(recipients)}] {recipient['Nome']} <{recipient['Email']}>"
        print(f"  {label} ... ", end="", flush=True)
        html_body = render_html(recipient, dates)
        status, error = "ok", None
        try:
            send_with_retry(token, recipient["Email"], cc_list, subject, html_body)
            print("OK")
        except Exception as exc:
            error = str(exc)
            status = "error"
            print(f"FALHOU")
            print(f"    Erro: {error}")

        results.append({"recipient": recipient, "status": status, "error": error})
        append_log({
            "timestamp": datetime.now().isoformat(),
            "subject": subject,
            "mes_ref": f"{dates['MesAtual']}/{dates['AnoVigente']}",
            "sender": sender,
            "recipient_name": recipient["Nome"],
            "recipient_email": recipient["Email"],
            "status": status,
            "error": error,
        })

    # --- Resumo ---
    ok = sum(1 for r in results if r["status"] == "ok")
    failed = len(results) - ok
    print(f"\n{'=' * 72}")
    print(f"  RESULTADO: {ok} enviado(s) com sucesso, {failed} falha(s)")
    if failed:
        print("\n  Falhas:")
        for r in results:
            if r["status"] == "error":
                print(f"    - {r['recipient']['Nome']} <{r['recipient']['Email']}>: {r['error']}")
    print(f"\n  Log salvo em: {LOG_PATH}")
    print("=" * 72)


if __name__ == "__main__":
    main()
