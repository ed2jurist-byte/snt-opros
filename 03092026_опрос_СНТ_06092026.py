# -*- coding: utf-8 -*-
"""
╔══════════════════════════════════════════════════════════════════╗
║        ООО «ЭК ЕД2» — Единая система email-рассылки              ║
║        Версия 12.2.2 — Бюллетень № 3: Собрание СНТ 06.09.2026   ║
╚══════════════════════════════════════════════════════════════════╝
"""

import smtplib
import time
import random
import re
import csv
import sys
import hashlib
import argparse
import logging
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("❌ Установите зависимости: pip install pandas openpyxl")
    sys.exit(1)

# ════════════════════════════════════════════════════════════════════
# ██  БЛОК НАСТРОЕК (ПРОВЕРЬТЕ ПЕРЕД ЗАПУСКОМ!)
# ════════════════════════════════════════════════════════════════════

GMAIL_USER   = "ed2.jurist@gmail.com"
GMAIL_PASS   = "xbxy pimm xrck swia"
SMTP_SERVER  = "smtp.gmail.com"
SMTP_PORT    = 465
REPLY_TO     = "ed2.jurist@ya.ru"

# ⚠️ ЗАМЕНИТЕ после деплоя на Render (см. инструкцию)
SURVEY_URL   = "https://ВАШ-СЕРВИС.onrender.com"

EXCEL_PATH   = Path(r"C:\Users\SM\OneDrive\2025 - 2026\ЭК\реестр домовладений\Сводный_реестр_сверка.xlsx")
REPORT_PATH  = Path(r"C:\Users\SM\OneDrive\2025 - 2026\ЭК\реестр домовладений\ОТЧЕТ_РАССЫЛКА.csv")
CASE_FOLDER = Path(r"C:\Users\SM\OneDrive\Файлы загрузки\Загрузки\СНТ 22.08.2026\документы")
ATTACHMENT_PATH = Path(r"C:\Users\SM\OneDrive\Файлы загрузки\Загрузки\СНТ 22.08.2026\документы\возражения\Дополнения_к_Возражениям_31_08_26.pdf")

TEST_TARGETS = [
    {"email": "ed2.jurist@ya.ru", "fio": "Иванов Иван Иванович", "uch": "000-Яндекс"}
]

SUBJECT_TEMPLATES = [
    "🚨 ВАЖНО: Электронный опрос жителей перед собранием СНТ 06.09.2026 (на 17 уч.)",
    "🚨 Опрос жителей перед собранием СНТ 06.09.2026 (на 17 уч.) — ВАЖНО"
]

TIMINGS = {
    "micro_delay_min": 2, "micro_delay_max": 5,
    "batch_size_min": 5, "batch_size_max": 8,
    "between_batch_min": 240, "between_batch_max": 360,
    "after_error": 180, "smtp_timeout": 45,
}

MAILING_ID = "ED2_POLL_SNT_MEETING_06092026_v1"

# ════════════════════════════════════════════════════════════════════
# ██  ЛОГИРОВАНИЕ
# ════════════════════════════════════════════════════════════════════

logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(levelname)s: %(message)s",
                    datefmt="%H:%M:%S", handlers=[logging.StreamHandler(sys.stdout)])
log = logging.getLogger("ed2_mailer")

def resolve_attachment_path() -> Path | None:
    if ATTACHMENT_PATH.is_file():
        return ATTACHMENT_PATH
    target_dir = ATTACHMENT_PATH.parent
    if target_dir.is_dir():
        files = list(target_dir.glob("*.pdf")) + list(target_dir.glob("*.docx"))
        if files:
            log.warning(f"⚠️ Точный файл не найден, выбран: {files[0].name}")
            return files[0]
    elif CASE_FOLDER.is_dir():
        files = list(CASE_FOLDER.glob("*.pdf")) + list(CASE_FOLDER.glob("*.docx"))
        if files:
            log.warning(f"⚠️ Точный файл не найден, выбран: {files[0].name}")
            return files[0]
    log.error("❌ Файл вложения не обнаружен.")
    return None

def countdown(seconds: int, label: str = "⏳ Ожидание"):
    for remaining in range(int(seconds), 0, -1):
        m, s = divmod(remaining, 60)
        sys.stdout.write(f"\r{label}: {m:02d}:{s:02d} осталось   ")
        sys.stdout.flush()
        time.sleep(1)
    sys.stdout.write("\r" + " " * 60 + "\r")
    sys.stdout.flush()

def build_html_message(fio: str, uch: str, email: str) -> str:
    return f"""<!DOCTYPE html>
<html lang="ru">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Электронный опрос жителей перед собранием СНТ</title>
</head>
<body style="margin:0;padding:0;background-color:#0f172a;font-family:'Segoe UI',Roboto,Helvetica,Arial,sans-serif;-webkit-font-smoothing:antialiased;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#0f172a" style="background-color:#0f172a;padding:30px 10px;">
<tr><td align="center">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff" style="background-color:#ffffff;border-radius:16px;max-width:600px;overflow:hidden;box-shadow:0 20px 25px -5px rgba(0,0,0,0.5),0 10px 10px -5px rgba(0,0,0,0.04);">
<tr>
  <td align="center" bgcolor="#1e293b" style="background:linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0284c7 100%);padding:36px 28px 28px 28px;text-align:center;">
    <table role="presentation" cellpadding="0" cellspacing="0" border="0" align="center" style="margin-bottom:16px;">
      <tr><td bgcolor="#ef4444" style="background-color:#ef4444;border-radius:20px;padding:6px 16px;">
        <span style="color:#ffffff;font-weight:800;font-size:12px;letter-spacing:1.5px;text-transform:uppercase;">🚨 ВАЖНО 🚨</span>
      </td></tr>
    </table>
    <div style="font-size:20px;line-height:1.3;font-weight:800;color:#ffffff;letter-spacing:-0.5px;">
      Электронный опрос жителей перед собранием СНТ 06.09.2026 (на 17 уч.)
    </div>
    <div style="margin-top:14px;font-size:14px;color:#cbd5e1;font-weight:600;">
      <span style="color:#38bdf8;font-weight:600;">Участок № {uch}</span>
    </div>
  </td>
</tr>
<tr>
  <td style="padding:30px 28px 10px 28px;">
    <p style="font-size:16px;line-height:1.6;color:#334155;margin:0 0 10px 0;font-weight:600;">Уважаемый(ая) {fio}!</p>
    <p style="font-size:15px;line-height:1.65;color:#334155;margin:0 0 16px 0;">
      Напоминаем, что на <b>06 сентября 2026 года в 12:00</b> на территории 17 строительного участка правлением СНТ назначено общее собрание.
    </p>
    <p style="font-size:15px;line-height:1.65;color:#334155;margin:0 0 20px 0;">
      Прохождение электронного опроса критически важно для выработки единой позиции собственников перед собранием. Позиция ООО «ЭК ЕД2» остается неизменной: более 76% сметы (12,4 млн руб. из 16,2 млн руб.) не имеют первичных подтверждающих документов, а коммунальные услуги и охрана территории уже в полном объеме оказываются нашей компанией по прямым договорам.
    </p>

    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#eff6ff" style="background-color:#eff6ff;border:2px dashed #0284c7;border-radius:12px;margin-bottom:24px;">
      <tr>
        <td style="padding:20px;text-align:left;">
          <div style="font-size:16px;font-weight:700;color:#0369a1;margin-bottom:12px;text-align:center;">📌 Как принять участие в опросе:</div>
          <ol style="margin:0;padding-left:20px;font-size:14.5px;line-height:1.7;color:#334155;">
            <li style="margin-bottom:8px;">
              Нажмите (кликните) на синюю кнопку <b>«Перейти к голосованию»</b> внизу данного письма.
            </li>
            <li style="margin-bottom:8px;">
              При входе в опрос укажите строго следующие данные:
              <ul style="margin-top:6px;padding-left:18px;">
                <li><b>E-mail:</b> используйте только тот адрес электронной почты, на который вы получили данное письмо — <span style="color:#0284c7;font-weight:700;">{email}</span>;</li>
                <li><b>Номер участка:</b> используйте только тот номер строительного участка, который указан в теме данного письма — <span style="color:#0284c7;font-weight:700;">{uch}</span>.</li>
              </ul>
            </li>
            <li style="margin-bottom:0;">
              Ответьте на вопросы анкеты и нажмите кнопку <b>«Отправить голос»</b>.
            </li>
          </ol>
        </td>
      </tr>
    </table>

    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">
      <tr><td align="center">
        <a href="{SURVEY_URL}" target="_blank" style="background-color:#0284c7;color:#ffffff;display:inline-block;padding:14px 28px;font-size:16px;font-weight:700;text-decoration:none;border-radius:8px;box-shadow:0 4px 6px -1px rgba(2,132,199,0.4);">👉 Перейти к голосованию</a>
      </td></tr>
    </table>

    <p style="font-size:14.5px;line-height:1.65;color:#334155;margin:0 0 16px 0;">
      💡 <b>Ваше участие поможет</b> сформировать консолидированные возражения, защитить интересы жителей и не допустить навязывания необоснованных взносов на собрании на 17 строительном участке.
    </p>
    <p style="font-size:13.5px;line-height:1.6;color:#64748b;margin:0 0 24px 0;background-color:#f8fafc;padding:12px;border-radius:8px;border:1px solid #e2e8f0;">
      🔒 <b>Конфиденциальность:</b> Голосование является тайным. Персональные данные используются исключительно для сверки с реестром.
    </p>
  </td>
</tr>
<tr>
  <td bgcolor="#f8fafc" style="background-color:#f8fafc;padding:24px 28px;border-top:1px solid #e2e8f0;text-align:center;">
    <div style="font-size:14px;font-weight:700;color:#0f172a;margin-bottom:4px;">С уважением,</div>
    <div style="font-size:15px;font-weight:800;color:#0284c7;margin-bottom:4px;">Администрация ООО «ЭК ЕД2»</div>
  </td>
</tr>
</table>
</td></tr>
</table>
</body>
</html>"""

def compute_content_hash() -> str:
    return hashlib.md5(MAILING_ID.encode("utf-8")).hexdigest()[:8]

def validate_email(raw: str) -> list:
    if not raw or str(raw).strip().lower() in ("nan", "none", "", "-"): return []
    pattern = re.compile(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")
    result = []
    cleaned_raw = str(raw).replace("\xa0", " ").replace(";", " ").replace(",", " ")
    for token in cleaned_raw.split():
        token = token.strip().lower()
        if "(" in token or ")" in token:
            continue
        if token and pattern.match(token): 
            result.append(token)
    return result

def open_smtp_connection():
    try:
        srv = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=TIMINGS["smtp_timeout"])
        srv.ehlo()
        try:
            srv.login(GMAIL_USER, GMAIL_PASS)
        except smtplib.SMTPAuthenticationError as auth_err:
            srv.close()
            if "535" in str(auth_err) or "Invalid user or password" in str(auth_err):
                raise RuntimeError(
                    "SMTP аутентификация провалилась (ошибка 535).\n"
                    "Причина: пароль приложения устарел, неверен или не сгенерирован.\n"
                    "Решение: Google Аккаунт → Безопасность → Двухэтапная аутентификация → Пароли приложений → Почта.\n"
                    "Скопируйте 16-значный код (без пробелов) в переменную GMAIL_PASS."
                )
            raise RuntimeError(f"SMTP аутентификация провалилась: {auth_err}")
        return srv
    except Exception as e:
        raise RuntimeError(f"Ошибка при установлении SMTP-сессии: {e}")

CSV_HEADERS = ["ID_рассылки", "ФИО", "Участок", "Email", "Тема", "Вложения", "Время", "Статус"]

def init_report_csv():
    if not REPORT_PATH.exists():
        REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(REPORT_PATH, "w", encoding="utf-8-sig", newline="") as f:
            csv.writer(f, delimiter=";").writerow(CSV_HEADERS)

def append_csv(row: list, retries: int = 3):
    for attempt in range(retries):
        try:
            with open(REPORT_PATH, "a", encoding="utf-8-sig", newline="") as f:
                csv.writer(f, delimiter=";").writerow(row)
            return
        except PermissionError:
            time.sleep(2)

def append_session_marker(label: str, content_hash: str):
    ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    append_csv([f"── {label} [{content_hash}] {ts} ──", "", "", "", "", "", "", ""])

def get_already_sent(content_hash: str) -> set:
    sent = set()
    if not REPORT_PATH.exists(): return sent
    try:
        with open(REPORT_PATH, "r", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=";")
            next(reader, None)
            for row in reader:
                if len(row) >= 8:
                    if row[0].strip() == content_hash and row[7].strip() == "OK":
                        sent.add(row[3].strip().lower())
    except Exception as e:
        log.warning(f"Ошибка чтения лог-файла: {e}")
    return sent

def load_recipients() -> list[dict]:
    if not EXCEL_PATH.exists():
        log.error(f"Реестр не найден: {EXCEL_PATH}")
        return []
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=0, dtype=str, engine="openpyxl")
        df.columns = df.columns.str.strip().str.lower()
    except Exception as e:
        log.error(f"Ошибка чтения Excel: {e}")
        return []

    col_fio = next((c for c in df.columns if any(k in c for k in ("фио", "ф.и.о", "фамилия", "собственник"))), None)
    col_uch = next((c for c in df.columns if any(k in c for k in ("участ", "номер", "уч.", "№"))), None)
    col_email = next((c for c in df.columns if any(k in c for k in ("email", "e-mail", "mail", "почта", "@"))), None)

    if not col_email:
        log.error("Столбец Email не найден!")
        return []

    recipients = []
    seen = set()
    for _, row in df.iterrows():
        fio = str(row.get(col_fio, "")).strip() if pd.notna(row.get(col_fio)) else "Собственник поселка"
        uch = str(row.get(col_uch, "")).strip().replace(".0", "") if pd.notna(row.get(col_uch)) else "—"
        email_raw = row.get(col_email, "")
        if pd.isna(email_raw): continue
        for email in validate_email(email_raw):
            if email not in seen:
                seen.add(email)
                recipients.append({"fio": fio or "Собственник поселка", "uch": uch or "—", "email": email})
    return recipients

def send_one_via_session(srv, target: dict, content_hash: str, file_path: Path, dry_run: bool = False, is_test: bool = False) -> bool:
    fio, uch, email = target["fio"], target["uch"], target["email"]
    subject = random.choice(SUBJECT_TEMPLATES)
    html_body = build_html_message(fio, uch, email)
    attachment_name = file_path.name if file_path else "Без приложения"

    if dry_run:
        log.info(f"[DRY-RUN] → {fio} | {email} | уч. {uch} | Вложение: {attachment_name} | Тема: {subject}")
        return True

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"ООО «ЭК «Европейская Долина-2» <{GMAIL_USER}>"
    msg["To"] = email
    msg["Reply-To"] = REPLY_TO
    msg.set_content(re.sub('<[^<]+?>', '', html_body))
    msg.add_alternative(html_body, subtype='html')

    if file_path and file_path.is_file():
        try:
            ext = file_path.suffix.lower()
            if ext == ".docx":
                maintype, subtype = "application", "vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif ext == ".pdf":
                maintype, subtype = "application", "pdf"
            else:
                maintype, subtype = "application", "octet-stream"
            with open(file_path, "rb") as f:
                msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=file_path.name)
        except Exception as e:
            log.error(f"❌ Ошибка чтения вложения: {e}")
            return False

    srv.send_message(msg)
    ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    log.info(f"✅ {ts}  {fio} | {email} | уч. {uch} | 📎 {attachment_name}")
    if not is_test:
        append_csv([content_hash, fio, uch, email, subject, attachment_name, ts, "OK"])
    return True

def mode_test():
    log.info("🧪 ЗАПУСК КОНТРОЛЬНОГО ТЕСТА...")
    if "ВАШ-СЕРВИС" in SURVEY_URL:
        log.error("🛑 Сначала замените SURVEY_URL на реальный URL после деплоя!")
        return
    file_path = resolve_attachment_path()
    if not file_path:
        log.error("🛑 ТЕСТ ОСТАНОВЛЕН: вложение не найдено.")
        return
    content_hash = compute_content_hash()
    try:
        srv = open_smtp_connection()
        for target in TEST_TARGETS:
            send_one_via_session(srv, target, content_hash, file_path=file_path, dry_run=False, is_test=True)
        srv.quit()
        log.info("🏁 Тест выполнен! Проверьте ящик ed2.jurist@ya.ru.")
    except Exception as e:
        log.error(f"Ошибка теста: {e}")

def mode_release(dry_run: bool = False, limit: int = 0):
    log.info("═" * 60)
    mode_label = "DRY-RUN (без отправки)" if dry_run else "МАССОВЫЙ ПУСК"
    log.info(f"🚀 РЕЖИМ: {mode_label}")
    log.info("═" * 60)
    if "ВАШ-СЕРВИС" in SURVEY_URL:
        log.error("🛑 Замените SURVEY_URL на реальный URL после деплоя!")
        return
    file_path = resolve_attachment_path()
    if not file_path:
        log.error("🛑 РАССЫЛКА ОСТАНОВЛЕНА: вложение не найдено.")
        return
    content_hash = compute_content_hash()
    init_report_csv()
    all_recipients = load_recipients()
    if not all_recipients: return
    already_sent = get_already_sent(content_hash)
    targets = [r for r in all_recipients if r["email"] not in already_sent]
    if limit > 0: targets = targets[:limit]
    total = len(targets)
    log.info(f"🔄 Всего: {len(all_recipients)}. Уже отправлено: {len(already_sent)}. К отправке: {total}")
    if total == 0:
        log.info("🏁 Все адреса уже обработаны.")
        return
    if not dry_run: append_session_marker("СЕССИЯ НАЧАТА", content_hash)
    idx = 0
    try:
        while idx < len(targets):
            batch = targets[idx : idx + random.randint(TIMINGS["batch_size_min"], TIMINGS["batch_size_max"])]
            srv = None
            if not dry_run:
                log.info(f"\n🔑 Подключение к Gmail для пачки из {len(batch)} писем...")
                try: srv = open_smtp_connection()
                except Exception as e:
                    log.error(f"Сбой подключения: {e}")
                    countdown(TIMINGS["after_error"], "😴 Пауза перед переподключением")
                    continue
            for target in batch:
                success = False
                for attempt in range(1, 4):
                    try:
                        success = send_one_via_session(srv, target, content_hash, file_path=file_path, dry_run=dry_run, is_test=False)
                        break
                    except smtplib.SMTPResponseException as smtp_err:
                        err_code = smtp_err.smtp_code
                        err_msg = str(smtp_err.smtp_error).lower()
                        if err_code == 550 or "limit exceeded" in err_msg or "daily" in err_msg:
                            log.error("\n🛑 Превышен суточный лимит Gmail!")
                            sys.exit(0)
                    except Exception as e:
                        log.error(f"❌ [Попытка {attempt}/3] {target['email']}: {e}")
                        if attempt == 3:
                            ts = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
                            append_csv([content_hash, target['fio'], target['uch'], target['email'], "Сбой", file_path.name, ts, f"ОШИБКА: {e}"])
                idx += 1
                if idx < len(targets) and target != batch[-1] and not dry_run and success:
                    countdown(random.randint(TIMINGS["micro_delay_min"], TIMINGS["micro_delay_max"]), "⚡ Задержка")
            if srv:
                try: srv.quit()
                except: pass
            if idx < len(targets):
                log.info(f"📊 Отправлено {idx}/{total}")
                countdown(random.randint(TIMINGS["between_batch_min"], TIMINGS["between_batch_max"]), "☕ Anti-spam пауза")
    finally:
        if not dry_run: append_session_marker("СЕССИЯ ЗАВЕРШЕНА", content_hash)
        log.info(f"🏁 Готово. Лог: {REPORT_PATH}")

def mode_status():
    if not REPORT_PATH.exists():
        log.info("Файл отчетов отсутствует.")
        return
    stats = {}
    with open(REPORT_PATH, "r", encoding="utf-8-sig") as f:
        reader = csv.reader(f, delimiter=";")
        next(reader, None)
        for row in reader:
            if len(row) < 8 or row[0].startswith("──"): continue
            h = row[0].strip()
            if h not in stats: stats[h] = {"ok": 0, "fail": 0}
            if row[7].strip() == "OK": stats[h]["ok"] += 1
            else: stats[h]["fail"] += 1
    current_hash = compute_content_hash()
    current_stat = stats.get(current_hash, {"ok": 0, "fail": 0})
    all_rec = load_recipients()
    already = get_already_sent(current_hash)
    pending = len([r for r in all_rec if r["email"] not in already])
    print(f"\n📊 СТАТУС [{current_hash}]:")
    print(f"   Успешно: {current_stat['ok']}")
    print(f"   Осталось: {pending}")

def main():
    parser = argparse.ArgumentParser(description="ООО «ЭК ЕД2» — Mailer")
    parser.add_argument("--mode", required=True, choices=["test", "release", "resume", "status"])
    parser.add_argument("--dry-run", action="store_true", default=False)
    parser.add_argument("--limit", type=int, default=0)
    args = parser.parse_args()
    if args.mode == "test": mode_test()
    elif args.mode in ("release", "resume"): mode_release(dry_run=args.dry_run, limit=args.limit)
    elif args.mode == "status": mode_status()

if __name__ == "__main__":
    main()
