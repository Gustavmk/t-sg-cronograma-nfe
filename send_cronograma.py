#!/usr/bin/env python3
"""
Automação de envio de cronograma de NF via Microsoft Graph API.

Uso:
    python send_cronograma.py --csv lista.csv
    python send_cronograma.py --csv lista.csv --assunto "Meu Assunto" --mes 5 --ano 2026
    python send_cronograma.py --csv lista.csv --dry-run
    python send_cronograma.py --csv lista.csv --force   # ignora idempotência
"""

import argparse
import csv
import json
import re
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import dns.resolver
import holidays
import msal
import requests

CONFIG_PATH   = Path(__file__).parent / "config.json"
TEMPLATE_PATH = Path(__file__).parent / "email_template.html"
TOKEN_CACHE_PATH = Path(__file__).parent / ".token_cache.bin"
LOG_PATH      = Path(__file__).parent / "send_log.jsonl"

MONTHS_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

GRAPH_SCOPES = [
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/User.Read",
]


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
        "MesAtual":       MONTHS_PT[mes],
        "AnoVigente":     str(ano),
        "MesSeguinte":    MONTHS_PT[mes_seg],
        "AnoMesSeguinte": str(ano_seg),
        "QuintoDiaUtil":  quinto_du.strftime("%d"),
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
# MX validation
# ---------------------------------------------------------------------------

def validate_mx(recipients: list[dict]) -> list[dict]:
    """
    Check DNS MX records for each unique email domain.
    Returns list of recipients whose domain has no MX record.
    Warnings are printed; execution is not aborted.
    """
    print("\nValidando registros MX dos domínios de email...")
    domains_checked: dict[str, bool] = {}
    invalid: list[dict] = []

    for r in recipients:
        domain = r["Email"].split("@")[1]
        if domain not in domains_checked:
            try:
                dns.resolver.resolve(domain, "MX", lifetime=5)
                domains_checked[domain] = True
            except (dns.resolver.NXDOMAIN, dns.resolver.NoAnswer,
                    dns.resolver.NoNameservers, dns.exception.Timeout):
                domains_checked[domain] = False

        if not domains_checked[domain]:
            print(f"  [AVISO] Domínio sem registro MX: {domain} ({r['Email']})")
            invalid.append(r)

    if not invalid:
        print(f"  Todos os {len(recipients)} domínio(s) validados com sucesso.")

    return invalid


# ---------------------------------------------------------------------------
# Idempotency check
# ---------------------------------------------------------------------------

def already_sent_emails(mes_ref: str) -> set[str]:
    """Return set of emails that were successfully sent for the given mes_ref."""
    if not LOG_PATH.exists():
        return set()
    sent = set()
    with open(LOG_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
                if entry.get("mes_ref") == mes_ref and entry.get("status") == "ok":
                    sent.add(entry["recipient_email"])
            except json.JSONDecodeError:
                continue
    return sent


# ---------------------------------------------------------------------------
# Email template & rendering
# ---------------------------------------------------------------------------

def load_template() -> str:
    if not TEMPLATE_PATH.exists():
        print(f"[ERRO] Template de email não encontrado: {TEMPLATE_PATH}")
        sys.exit(1)
    return TEMPLATE_PATH.read_text(encoding="utf-8")


def render_html(template: str, recipient: dict, dates: dict) -> str:
    variables = {**dates, "ValorNF": recipient["ValorNF"]}
    body = template
    for key, value in variables.items():
        body = body.replace("{{" + key + "}}", value)
    return body


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
# Summary notification email
# ---------------------------------------------------------------------------

def send_summary_notification(
    token: str,
    to_email: str,
    subject_ref: str,
    results: list[dict],
    mes_ref: str,
    sender: str,
) -> None:
    ok      = [r for r in results if r["status"] == "ok"]
    failed  = [r for r in results if r["status"] == "error"]
    skipped = [r for r in results if r["status"] == "skipped"]

    rows_ok = "".join(
        f"<tr><td style='padding:6px 10px;border:1px solid #ddd;'>✅</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;'>{r['recipient']['Nome']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;'>{r['recipient']['Email']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;'></td></tr>"
        for r in ok
    )
    rows_failed = "".join(
        f"<tr><td style='padding:6px 10px;border:1px solid #ddd;color:#c00;'>❌</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;'>{r['recipient']['Nome']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;'>{r['recipient']['Email']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;color:#c00;font-size:12px;'>{r['error']}</td></tr>"
        for r in failed
    )
    rows_skipped = "".join(
        f"<tr><td style='padding:6px 10px;border:1px solid #ddd;color:#888;'>⏭️</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;color:#888;'>{r['recipient']['Nome']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;color:#888;'>{r['recipient']['Email']}</td>"
        f"<td style='padding:6px 10px;border:1px solid #ddd;color:#888;'>Já enviado anteriormente</td></tr>"
        for r in skipped
    )

    status_geral = "✅ Concluído com sucesso" if not failed else f"⚠️ Concluído com {len(failed)} falha(s)"
    timestamp    = datetime.now().strftime("%d/%m/%Y %H:%M")

    html = f"""<!DOCTYPE html><html lang="pt-BR">
<body style="font-family:Arial,sans-serif;font-size:14px;color:#333;max-width:720px;">
<h2 style="color:#2c5f8a;">Relatório de Envio — {mes_ref}</h2>
<p><strong>Status:</strong> {status_geral}<br>
   <strong>Executado em:</strong> {timestamp}<br>
   <strong>Remetente:</strong> {sender}<br>
   <strong>Assunto enviado:</strong> {subject_ref}</p>

<table style="width:100%;border-collapse:collapse;margin-top:16px;">
  <thead>
    <tr style="background:#f0f0f0;">
      <th style="padding:8px 10px;border:1px solid #ddd;text-align:left;width:40px;"></th>
      <th style="padding:8px 10px;border:1px solid #ddd;text-align:left;">Nome</th>
      <th style="padding:8px 10px;border:1px solid #ddd;text-align:left;">Email</th>
      <th style="padding:8px 10px;border:1px solid #ddd;text-align:left;">Obs.</th>
    </tr>
  </thead>
  <tbody>
    {rows_ok}{rows_failed}{rows_skipped}
  </tbody>
</table>

<p style="margin-top:24px;font-size:12px;color:#888;">
  Enviado automaticamente por <em>send_cronograma.py</em>.
</p>
</body></html>"""

    payload = {
        "message": {
            "subject": f"[Relatório] Envio de Cronograma — {mes_ref}",
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
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


# ---------------------------------------------------------------------------
# Preview & UX
# ---------------------------------------------------------------------------

def print_preview(
    recipients: list[dict],
    dates: dict,
    subject: str,
    cc_list: list[str],
    already_sent: set[str],
) -> None:
    sep = "=" * 72
    print(f"\n{sep}")
    print("  PREVIEW — EMAILS A SEREM ENVIADOS")
    print(sep)
    print(f"  Assunto    : {subject}")
    print(f"  Ref.       : {dates['MesAtual']}/{dates['AnoVigente']}")
    print(f"  Pagamento  : {dates['QuintoDiaUtil']}/{dates['MesSeguinte']}/{dates['AnoMesSeguinte']}")
    if cc_list:
        print(f"  CC (todos) : {', '.join(cc_list)}")
    print(f"\n  {'#':<4} {'S':<3} {'Nome':<24} {'Email':<34} {'Valor NF'}")
    print("  " + "-" * 72)
    for i, r in enumerate(recipients, 1):
        flag = "[já enviado]" if r["Email"] in already_sent else ""
        mark = "⏭" if flag else "→"
        print(f"  {i:<4} {mark:<3} {r['Nome']:<24} {r['Email']:<34} {r['ValorNF']} {flag}")
    print(sep)
    new_count = sum(1 for r in recipients if r["Email"] not in already_sent)
    skip_count = len(recipients) - new_count
    print(f"\n  Total: {len(recipients)} destinatário(s)  |  A enviar: {new_count}  |  Já enviados (pulados): {skip_count}\n")


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
    parser.add_argument("--force", action="store_true",
                        help="Ignora checagem de idempotência e reenvia mesmo emails já enviados")
    args = parser.parse_args()

    config   = load_config()
    cc_list: list[str] = config.get("cc_emails", [])
    notif_to: str      = config.get("notification_email", "")

    today    = date.today()
    ref_date = date(args.ano or today.year, args.mes or today.month, 1)
    dates    = compute_dates(ref_date)
    mes_ref  = f"{dates['MesAtual']}/{dates['AnoVigente']}"

    # --- Carregar template ---
    template = load_template()

    # --- Ler CSV ---
    print(f"\nLendo arquivo CSV: {args.csv}")
    try:
        recipients = read_csv(args.csv)
    except (FileNotFoundError, ValueError) as exc:
        print(f"[ERRO] {exc}")
        sys.exit(1)
    print(f"  {len(recipients)} destinatário(s) encontrado(s).")

    # --- Validar MX (skip em dry-run para agilizar) ---
    mx_invalid: list[dict] = []
    if not args.dry_run:
        mx_invalid = validate_mx(recipients)

    # --- Idempotência ---
    sent_before: set[str] = set() if args.force else already_sent_emails(mes_ref)
    if sent_before and not args.force:
        print(f"\n  [Idempotência] {len(sent_before)} email(s) já enviado(s) este mês serão pulados.")
        print("                 Use --force para reenviar mesmo assim.")

    # --- Assunto ---
    subject = args.assunto or ask_subject(dates)

    # --- Preview ---
    print_preview(recipients, dates, subject, cc_list, sent_before)

    if args.dry_run:
        if mx_invalid:
            print(f"  [DRY-RUN] {len(mx_invalid)} domínio(s) sem registro MX detectado(s) — verifique acima.")
        print("  [DRY-RUN] Nenhum email foi enviado. Remova --dry-run para enviar.")
        print("=" * 72)
        sys.exit(0)

    # --- Avisar sobre domínios inválidos antes de confirmar ---
    if mx_invalid:
        print(f"  [AVISO] {len(mx_invalid)} email(s) com domínio sem MX. Eles serão tentados mesmo assim.")

    if not confirm_send():
        print("Envio cancelado.")
        sys.exit(0)

    # --- Autenticar ---
    print("\nAutenticando com Microsoft Graph API...")
    try:
        token  = get_access_token(config)
        sender = get_sender_email(token)
        print(f"  Remetente: {sender}\n")
    except Exception as exc:
        print(f"[ERRO] Falha na autenticação: {exc}")
        sys.exit(1)

    # --- Enviar ---
    print(f"Enviando emails para {mes_ref}...")
    results = []

    for i, recipient in enumerate(recipients, 1):
        label = f"[{i}/{len(recipients)}] {recipient['Nome']} <{recipient['Email']}>"

        # Idempotência: pular já enviados
        if recipient["Email"] in sent_before:
            print(f"  {label} ... PULADO (já enviado)")
            results.append({"recipient": recipient, "status": "skipped", "error": None})
            continue

        print(f"  {label} ... ", end="", flush=True)
        html_body = render_html(template, recipient, dates)
        status, error = "ok", None

        try:
            send_with_retry(token, recipient["Email"], cc_list, subject, html_body)
            print("OK")
        except Exception as exc:
            error  = str(exc)
            status = "error"
            print("FALHOU")
            print(f"    Erro: {error}")

        results.append({"recipient": recipient, "status": status, "error": error})
        if status != "skipped":
            append_log({
                "timestamp":       datetime.now().isoformat(),
                "subject":         subject,
                "mes_ref":         mes_ref,
                "sender":          sender,
                "recipient_name":  recipient["Nome"],
                "recipient_email": recipient["Email"],
                "status":          status,
                "error":           error,
            })

    # --- Resumo ---
    ok      = sum(1 for r in results if r["status"] == "ok")
    failed  = sum(1 for r in results if r["status"] == "error")
    skipped = sum(1 for r in results if r["status"] == "skipped")

    print(f"\n{'=' * 72}")
    print(f"  RESULTADO: {ok} enviado(s) | {failed} falha(s) | {skipped} pulado(s)")
    if failed:
        print("\n  Falhas:")
        for r in results:
            if r["status"] == "error":
                print(f"    - {r['recipient']['Nome']} <{r['recipient']['Email']}>: {r['error']}")
    print(f"\n  Log salvo em: {LOG_PATH}")
    print("=" * 72)

    # --- Notificação pós-envio ---
    if notif_to and ok + failed > 0:
        print(f"\nEnviando relatório para {notif_to} ... ", end="", flush=True)
        try:
            send_summary_notification(token, notif_to, subject, results, mes_ref, sender)
            print("OK")
        except Exception as exc:
            print(f"FALHOU ({exc})")


if __name__ == "__main__":
    main()
