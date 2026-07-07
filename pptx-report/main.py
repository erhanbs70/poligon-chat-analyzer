"""
main.py
Daily orchestration: fetch metrics -> build pptx -> email as attachment.

Run manually:
    python3 main.py 2026-07-06

Or via GitHub Actions (see .github/workflows/daily-pptx.yml), which passes
"yesterday" (Europe/Sofia) by default.

Required environment variables (set as GitHub Secrets):
    MR_SITE_ID, MR_API_KEY, MR_EMAIL, MR_CS_DEPT   (Comm100 API credentials)
    GMAIL_ADDRESS, GMAIL_APP_PASSWORD               (sender account)
    MAIL_TO                                         (comma-separated recipients)
"""
import os
import sys
import smtplib
import datetime as dt
from email.message import EmailMessage

from fetch_metrics import build_metrics
from build_pptx import build_pptx

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.pptx")


def sofia_yesterday():
    # Europe/Sofia is UTC+2/+3; using UTC+3 upper bound is safe enough for
    # picking "yesterday" in a daily-morning cron context.
    now_utc = dt.datetime.utcnow()
    sofia_now = now_utc + dt.timedelta(hours=3)
    return (sofia_now.date() - dt.timedelta(days=1)).isoformat()


def send_email(pptx_path, date_label, metrics):
    gmail_address = os.environ["GMAIL_ADDRESS"]
    gmail_password = os.environ["GMAIL_APP_PASSWORD"]
    mail_to = [addr.strip() for addr in os.environ.get("MAIL_TO", "").split(",") if addr.strip()]
    if not mail_to:
        raise RuntimeError("MAIL_TO is empty — set it as a GitHub secret (comma-separated addresses).")

    msg = EmailMessage()
    total = metrics["totalChats"]
    acc = metrics["acceptancePct"]
    msg["Subject"] = f"CS Team Daily Chat Stats | {date_label} | {total} chat | {acc:.2f}% acc"
    msg["From"] = gmail_address
    msg["To"] = ", ".join(mail_to)
    msg.set_content(f"CS Team Daily Chat Stats — {date_label}\n\nEkte pptx raporu bulabilirsiniz.")

    with open(pptx_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.presentationml.presentation",
            filename=os.path.basename(pptx_path),
        )

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(gmail_address, gmail_password)
        smtp.send_message(msg)


def main():
    date_str = sys.argv[1].strip() if len(sys.argv) > 1 and sys.argv[1].strip() else sofia_yesterday()
    date_label = dt.datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")

    print(f"[1/3] Fetching metrics for {date_str} ...")
    metrics = build_metrics(date_str)

    print("[2/3] Building pptx ...")
    out_name = f"CS_Team_daily_chat_stats_{date_label.replace('.', '_')}.pptx"
    out_path = build_pptx(metrics, TEMPLATE_PATH, out_name, date_label)

    print("[3/3] Sending email ...")
    send_email(out_path, date_label, metrics)
    print("Done ->", out_path)


if __name__ == "__main__":
    main()
