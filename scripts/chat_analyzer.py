#!/usr/bin/env python3
# ============================================================
# POLIGON COMM100 CHAT ANALYZER — v2
# Sadece CS departmanı, sadece şikayet/sorun konuları
# ============================================================

import os
import json
import base64
import requests
import smtplib
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID       = os.environ["COMM100_SITE_ID"]
API_KEY       = os.environ["COMM100_API_KEY"]
COMM100_EMAIL = os.environ["COMM100_EMAIL"]
GEMINI_KEYS   = json.loads(os.environ["GEMINI_KEYS"])
CLAUDE_KEY    = os.environ["CLAUDE_KEY"]
GROQ_KEYS     = json.loads(os.environ["GROQ_KEYS"])
GMAIL_USER    = os.environ["GMAIL_USER"]
GMAIL_PASS    = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS = [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

CS_DEPT_ID = "29098b6e-0acb-40cc-8cd7-ebf662defc42"

CAMPAIGNS = {
    "SB": "3cd6532a-0d60-4d7d-9331-5dd5be71f527",
    "BS": "2b71ad44-6a12-4073-bc6c-8edb22eabcc6",
    "TB": "6712311a-2268-408c-a803-b338a1308010"
}

BASE_URL = f"https://dash15.lively-chat.com/api/LiveChat/chats:search?siteId={SITE_ID}"
gemini_key_index = 0

# ── AUTH ─────────────────────────────────────────────────────
def get_auth():
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

# ── DATE ─────────────────────────────────────────────────────
def get_yesterday_sofia():
    import pytz
    sofia_tz = pytz.timezone("Europe/Sofia")
    yesterday = datetime.now(sofia_tz) - timedelta(days=1)
    return yesterday.strftime("%Y-%m-%d")

def sofia_to_utc_range(date_str):
    import pytz
    sofia_tz = pytz.timezone("Europe/Sofia")
    start = sofia_tz.localize(datetime.strptime(date_str + " 00:00:00", "%Y-%m-%d %H:%M:%S"))
    end   = sofia_tz.localize(datetime.strptime(date_str + " 23:59:59", "%Y-%m-%d %H:%M:%S"))
    return (
        start.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        end.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    )

# ── COMM100 FETCH — sadece CS dept ───────────────────────────
def fetch_cs_chats(date_str):
    start_time, end_time = sofia_to_utc_range(date_str)
    auth = get_auth()
    all_chats = []
    page = 1
    seen_ids = set()

    print(f"[FETCH] CS chatleri çekiliyor: {date_str}")

    while page <= 40:
        url = (f"{BASE_URL}&pageIndex={page}&pageSize=500"
               f"&include=chatAgent&include=chatWrapup&include=postChatSurvey"
               f"&sortBy=startTime&sortOrder=asc")
        try:
            r = requests.post(
                url,
                headers={"Authorization": auth, "Content-Type": "application/json"},
                json={
                    "startTime": start_time,
                    "endTime": end_time,
                    "departmentId": CS_DEPT_ID   # sadece CS departmanı
                },
                timeout=30
            )
            if r.status_code != 200:
                print(f"[FETCH] HTTP {r.status_code} sayfa {page}: {r.text[:200]}")
                break

            chats = r.json().get("list", [])
            if not chats:
                break

            for c in chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                if cid and cid not in seen_ids:
                    seen_ids.add(cid)
                    all_chats.append(c)

            print(f"[FETCH] Sayfa {page}: {len(chats)} chat ({len(all_chats)} toplam)")
            if len(chats) < 500:
                break
            page += 1
            time.sleep(0.3)

        except Exception as e:
            print(f"[FETCH] Hata sayfa {page}: {e}")
            break

    print(f"[FETCH] Toplam {len(all_chats)} CS chat çekildi")
    return all_chats

def fetch_chat_messages(chat_id):
    url = f"https://dash15.lively-chat.com/api/LiveChat/chats/{chat_id}?include=messages&siteId={SITE_ID}"
    try:
        r = requests.get(url, headers={"Authorization": get_auth()}, timeout=15)
        if r.status_code == 200:
            return r.json().get("messages", [])
    except Exception as e:
        print(f"[MSG] {chat_id} hata: {e}")
    return []

# ── HELPERS ──────────────────────────────────────────────────
def get_agent_name(chat):
    agents = chat.get("chatAgents", [])
    if agents:
        names = [
            (a.get("agent") or {}).get("displayName") or a.get("displayName") or ""
            for a in agents if not a.get("botId")
        ]
        return ", ".join(filter(None, names))
    return chat.get("agentName", "")

def get_brand(chat):
    cid = (chat.get("campaignId") or "").lower()
    for brand, camp_id in CAMPAIGNS.items():
        if cid == camp_id.lower():
            return brand
    url = chat.get("requestingPageURL", "").lower()
    if "superbetin" in url: return "SB"
    if "betsat" in url: return "BS"
    if "turkbet" in url: return "TB"
    return "?"

def is_bot_only(chat):
    return chat.get("chatType") == "chatBotOnly"

def get_visitor_messages(messages):
    """Ziyaretçi mesajlarını birleştir, max 8 mesaj"""
    msgs = []
    for m in messages:
        if m.get("senderType") == "visitor":
            text = (m.get("message") or "").strip()
            if text and len(text) > 3:
                msgs.append(text)
    return " | ".join(msgs[:8])

# ── PROCESS: sadece agent chatler, mesajları çek ─────────────
def process_chats(chats):
    agent_chats = [c for c in chats if not is_bot_only(c)]
    print(f"[PROCESS] {len(chats)} CS chat → {len(agent_chats)} agent chat")

    chat_texts = []
    for i, chat in enumerate(agent_chats):
        chat_id = str(chat.get("id") or chat.get("chatId") or "")
        brand   = get_brand(chat)

        messages = fetch_chat_messages(chat_id)
        visitor_msgs = get_visitor_messages(messages)

        if not visitor_msgs or len(visitor_msgs) < 10:
            continue

        chat_texts.append({
            "chat_id":     chat_id,
            "brand":       brand,
            "visitor_msgs": visitor_msgs
        })

        if (i + 1) % 100 == 0:
            print(f"[PROCESS] {i+1}/{len(agent_chats)} işlendi ({len(chat_texts)} mesaj var)")
        time.sleep(0.08)

    print(f"[PROCESS] {len(chat_texts)} chat mesajı AI için hazır")
    return agent_chats, chat_texts

# ── AI ────────────────────────────────────────────────────────
def build_prompt(chat_texts):
    lines = [f"{i+1}. [{c['brand']}] {c['visitor_msgs']}"
             for i, c in enumerate(chat_texts)]

    schema = json.dumps({
        "categories": [
            {
                "name": "Spesifik şikayet/sorun adı — max 6 kelime",
                "count": 0,
                "brand_breakdown": [{"brand": "SB", "count": 0}],
                "short_note": "1 cümle özet"
            }
        ],
        "summary": "max 2 cümle"
    }, ensure_ascii=False)

    return f"""Comm100 müşteri destek chat analisti olarak aşağıdaki CS departmanı ziyaretçi mesajlarını analiz et.

Tarih: {datetime.now().strftime('%d.%m.%Y')} | Toplam: {len(chat_texts)} agent chat

--- ZİYARETÇİ MESAJLARI ---
{chr(10).join(lines)}
--- ---

KURAL 1 — SADECE ŞİKAYET/SORUN/ÇÖZÜLMEZ TALEPLERİ GRUPLA:
✅ DAHİL ET: Para çekimi gelmedi, yatırım yansımadı, bonus verilmedi, hesap açılmıyor, teknik sorun, şikayet
❌ HARIÇ TUT: "Canlı destek var mı?", "bakiye sorgulama", "bahis oranı sorma", genel bilgi soruları

KURAL 2 — KONU ADI SPESİFİK OLMALI:
YANLIŞ: "Çekim Sorunları", "Diğer", "Genel Sorun"
DOĞRU: "Onaylı Havale Çekimi Hesaba Geçmedi", "Papara Yatırımı Yansımadı", "Bonus Aktivasyon Yapılmıyor"

KURAL 3 — "Diğer" kategorisi YASAK. Her konu kendi spesifik adıyla gruplandırılmalı.

KURAL 4 — brand_breakdown: sadece o konuda hangi brand kaç chat var.

GÖREV: Şikayet/sorun içeren mesajları spesifik konulara göre grupla, büyükten küçüğe sırala.

SB=Superbetin | BS=Betsat | TB=Turkbet
Sadece JSON döndür:
{schema}"""


def try_gemini(prompt):
    global gemini_key_index
    models = ["gemini-2.5-flash", "gemini-2.0-flash"]
    for ki in range(len(GEMINI_KEYS)):
        key = GEMINI_KEYS[(gemini_key_index + ki) % len(GEMINI_KEYS)]
        for model in models:
            try:
                r = requests.post(
                    f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}",
                    json={
                        "contents": [{"parts": [{"text": prompt}]}],
                        "generationConfig": {"responseMimeType": "application/json"},
                        "safetySettings": [
                            {"category": "HARM_CATEGORY_HARASSMENT",        "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_HATE_SPEECH",       "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
                        ]
                    }, timeout=60
                )
                if r.status_code in [429, 503]: break
                if r.status_code == 200:
                    res = r.json()
                    cands = res.get("candidates", [])
                    if cands and cands[0].get("content"):
                        text = cands[0]["content"]["parts"][0]["text"].strip()
                        gemini_key_index = (gemini_key_index + ki + 1) % len(GEMINI_KEYS)
                        model_name = "Gemini 2.5 Flash" if "2.5" in model else "Gemini 2.0 Flash"
                        print(f"[AI] {model_name} başarılı")
                        return {"success": True, "text": text, "model": model_name}
            except Exception as e:
                print(f"[GEMINI] {model} key#{ki}: {e}")
    return {"success": False}

def try_claude(prompt):
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": CLAUDE_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 2000,
                  "messages": [{"role": "user", "content": prompt}]},
            timeout=60
        )
        if r.status_code == 200:
            text = r.json()["content"][0]["text"].strip()
            print("[AI] Claude Haiku başarılı")
            return {"success": True, "text": text, "model": "Claude Haiku 4.5"}
    except Exception as e:
        print(f"[CLAUDE] {e}")
    return {"success": False}

def try_groq(prompt):
    for i, key in enumerate(GROQ_KEYS):
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={"model": "meta-llama/llama-4-scout-17b-16e-instruct",
                      "messages": [{"role": "user", "content": prompt}],
                      "temperature": 0.2, "response_format": {"type": "json_object"}},
                timeout=60
            )
            if r.status_code == 200:
                text = r.json()["choices"][0]["message"]["content"].strip()
                print(f"[AI] Groq başarılı (key#{i})")
                return {"success": True, "text": text, "model": "Groq Llama-4 Scout"}
        except Exception as e:
            print(f"[GROQ] key#{i}: {e}")
    return {"success": False}

def parse_json(text):
    try:
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except:
        return None

def analyze_with_ai(chat_texts):
    if not chat_texts:
        return None, "—"
    prompt = build_prompt(chat_texts)
    for fn in [try_gemini, try_claude, try_groq]:
        result = fn(prompt)
        if result["success"]:
            data = parse_json(result["text"])
            if data:
                data["model_used"] = result["model"]
                return data, result["model"]
    return None, "Hata"

# ── HTML ──────────────────────────────────────────────────────
def line_color(cnt):
    if cnt >= 8: return "#2F1555"
    if cnt >= 4: return "#662D91"
    if cnt >= 2: return "#B28ABF"
    return "#d8b4fe"

def brand_badge(brand, count=None):
    styles = {
        "SB": "background:#1d4ed8;color:#fff;",
        "BS": "background:#FFE600;color:#2F1555;",
        "TB": "background:#E30613;color:#fff;",
    }
    s = styles.get(brand, "background:#666;color:#fff;")
    badge = f'<span style="display:inline-block;padding:1px 6px;{s}border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{brand}</span>'
    if count is not None:
        badge += f'<strong style="font-size:12px;color:#662D91;margin-left:3px;font-family:Montserrat,Arial,sans-serif;">{count}</strong>'
    return badge

def build_html(ai_data, stats, date_str, model_used):
    total        = stats["total"]
    agent_count  = stats["agent_count"]
    bot_count    = stats["bot_count"]
    brand_counts = stats["brands"]

    # Brand pill'leri
    brand_html = ""
    for b, cnt in brand_counts.items():
        if cnt:
            bg = {"SB": "background:#1d4ed8;color:#fff;",
                  "BS": "background:#FFE600;color:#2F1555;",
                  "TB": "background:#E30613;color:#fff;"}.get(b, "")
            brand_html += (
                f'<span style="margin-right:16px;font-family:Montserrat,Arial,sans-serif;">'
                f'<strong style="color:#FFE600;font-size:15px;">{cnt}</strong>'
                f'<span style="{bg}padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700;margin-left:4px;">{b}</span>'
                f'</span>'
            )

    # Konu satırları
    topics_html = ""
    if ai_data and ai_data.get("categories"):
        cats = sorted(ai_data["categories"], key=lambda x: x.get("count", 0), reverse=True)
        for idx, cat in enumerate(cats):
            cnt   = cat.get("count", 0)
            lclr  = line_color(cnt)
            border = "none" if idx == len(cats) - 1 else "1px solid #f0e8ff"

            bd_html = ""
            if cat.get("brand_breakdown"):
                items = "".join(
                    f'<span style="display:inline-block;margin-right:8px;">'
                    f'{brand_badge(bd["brand"])}'
                    f'<strong style="font-size:12px;color:#662D91;margin-left:3px;font-family:Montserrat,Arial,sans-serif;">{bd["count"]}</strong>'
                    f'</span>'
                    for bd in cat["brand_breakdown"]
                )
                bd_html = (
                    f'<div style="margin:5px 0;">'
                    f'<span style="font-size:9px;font-weight:700;color:#7c3aed;text-transform:uppercase;'
                    f'letter-spacing:.08em;font-family:Montserrat,Arial,sans-serif;">Marka: </span>{items}'
                    f'</div>'
                )

            note_html = ""
            if cat.get("short_note"):
                note_html = (
                    f'<p style="margin:6px 0 0;font-size:11px;color:#662D91;font-style:italic;'
                    f'font-family:Montserrat,Arial,sans-serif;line-height:1.5;">{cat["short_note"]}</p>'
                )

            topics_html += f'''
            <table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:{border};">
              <tr>
                <td style="width:4px;background-color:{lclr};font-size:0;" bgcolor="{lclr}">&nbsp;</td>
                <td style="padding:14px 16px;background-color:#ffffff;">
                  <table width="100%" cellpadding="0" cellspacing="0"><tr>
                    <td><span style="font-size:13px;font-weight:700;color:#2F1555;font-family:Montserrat,Arial,sans-serif;">{cat.get("name","Konu")}</span></td>
                    <td style="text-align:right;white-space:nowrap;">
                      <span style="font-size:26px;font-weight:800;color:{lclr};line-height:1;font-family:Montserrat,Arial,sans-serif;">{cnt}</span>
                    </td>
                  </tr></table>
                  {bd_html}{note_html}
                </td>
              </tr>
            </table>'''

    summary_html = ""
    if ai_data and ai_data.get("summary"):
        summary_html = f'''
        <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;border-top:1px solid #f0e8ff;">
          <tr><td style="padding-top:14px;">
            <p style="margin:0 0 5px;font-size:10px;font-weight:800;color:#662D91;letter-spacing:.12em;
               text-transform:capitalize;font-family:Montserrat,Arial,sans-serif;">Özet</p>
            <p style="margin:0;font-size:13px;color:#4b5563;line-height:1.7;font-family:Montserrat,Arial,sans-serif;">{ai_data["summary"]}</p>
          </td></tr>
        </table>'''

    model_badge_html = (
        f'<span style="display:inline-block;padding:2px 9px;background:#f3e8ff;color:#662D91;'
        f'border:1px solid #B28ABF;border-radius:3px;font-size:10px;font-weight:700;'
        f'font-family:Montserrat,Arial,sans-serif;">{model_used}</span>'
    )

    no_data_html = '<p style="color:#B28ABF;font-size:13px;font-family:Montserrat,Arial,sans-serif;">Şikayet/sorun tespit edilemedi.</p>'

    return f'''<!DOCTYPE html><html>
<head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background-color:#f0f0f0;" bgcolor="#f0f0f0">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f0f0">
<tr><td align="center" style="padding:32px 16px;">
<table width="620" cellpadding="0" cellspacing="0"
  style="max-width:620px;width:100%;background-color:#ffffff;border-radius:16px;overflow:hidden;
         box-shadow:0 16px 60px rgba(47,21,85,0.35),0 4px 20px rgba(102,45,145,0.25);"
  bgcolor="#ffffff">
<tr><td>

<!-- HERO -->
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td style="background-color:#2F1555;padding:28px 28px 24px;text-align:center;" bgcolor="#2F1555">
  <p style="margin:0 0 8px;font-size:9px;color:#B28ABF;text-transform:uppercase;letter-spacing:.2em;
     font-weight:700;font-family:Montserrat,Arial,sans-serif;">Poligon &middot; CS Ops &middot; Günlük Şikayet Analizi</p>
  <p style="margin:0;font-size:64px;font-weight:900;color:#FFE600;line-height:1;font-family:Montserrat,Arial,sans-serif;">{agent_count}</p>
  <p style="margin:3px 0 2px;font-size:10px;color:#B28ABF;text-transform:uppercase;letter-spacing:.22em;
     font-weight:700;font-family:Montserrat,Arial,sans-serif;">CS AGENT CHAT</p>
  <p style="margin:0 0 16px;font-size:13px;color:#FFE600;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{date_str}</p>
  <table cellpadding="0" cellspacing="0" align="center"><tr>
    <td style="background-color:rgba(255,255,255,0.1);border-radius:20px;padding:7px 20px;
               border:1px solid rgba(178,138,191,0.3);">{brand_html}</td>
  </tr></table>
  <p style="margin:10px 0 0;">
    <span style="font-size:10px;color:#e9d5ff;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Toplam: {total}</span>
    <span style="color:rgba(178,138,191,0.4);margin:0 8px;">&middot;</span>
    <span style="font-size:10px;color:#e9d5ff;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Bot: {bot_count}</span>
    <span style="color:rgba(178,138,191,0.4);margin:0 8px;">&middot;</span>
    <span style="font-size:10px;color:#e9d5ff;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Agent: {agent_count}</span>
  </p>
</td></tr>
</table>

<!-- HR -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:#f0e8ff;font-size:0;" bgcolor="#f0e8ff">&nbsp;</td>
</tr></table>

<!-- AI ANALİZİ -->
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td style="padding:20px 28px;" bgcolor="#ffffff">
  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr>
    <td style="font-size:11px;font-weight:800;color:#2F1555;text-transform:uppercase;
       letter-spacing:.14em;font-family:Montserrat,Arial,sans-serif;">Şikayet Analizi</td>
    <td style="text-align:right;">{model_badge_html}</td>
  </tr></table>

  {'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e9d5ff;border-radius:8px;overflow:hidden;">' + topics_html + '</table>' if topics_html else no_data_html}

  {summary_html}
</td></tr>
</table>

<!-- FOOTER -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:#2F1555;" bgcolor="#2F1555">
  <p style="margin:0;font-size:11px;color:#FFE600;font-weight:700;text-align:center;
     font-family:Montserrat,Arial,sans-serif;">POLIGON Chat Analyzer &nbsp;&middot;&nbsp; CS Günlük</p>
  <p style="margin:4px 0 0;font-size:10px;color:#B28ABF;text-align:center;
     font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>

</td></tr></table>
</td></tr></table>
</body></html>'''

# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, date_str, stats):
    subject = f"CS Şikayet Analizi | {date_str} | {stats['agent_count']} Agent Chat"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])
    msg.attach(MIMEText(html, "html", "utf-8"))
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
    print(f"[EMAIL] Gönderildi: {', '.join(REPORT_EMAILS)}")

# ── MAIN ──────────────────────────────────────────────────────
def main():
    date_str = get_yesterday_sofia()
    print(f"\n{'='*50}\nPOLIGON CS CHAT ANALYZER — {date_str}\n{'='*50}\n")

    # 1. CS chatlerini çek
    chats = fetch_cs_chats(date_str)
    if not chats:
        print("[MAIN] CS chat bulunamadı")
        return

    # 2. İstatistikler
    bot_chats   = [c for c in chats if is_bot_only(c)]
    agent_chats = [c for c in chats if not is_bot_only(c)]
    brand_counts = {"SB": 0, "BS": 0, "TB": 0}
    for c in agent_chats:
        b = get_brand(c)
        if b in brand_counts:
            brand_counts[b] += 1

    stats = {
        "total":       len(chats),
        "agent_count": len(agent_chats),
        "bot_count":   len(bot_chats),
        "brands":      brand_counts
    }
    print(f"[STATS] Toplam: {stats['total']} | Agent: {stats['agent_count']} | Bot: {stats['bot_count']}")
    print(f"[STATS] Brands: {brand_counts}")

    # 3. Mesajları çek
    agent_chats_list, chat_texts = process_chats(chats)

    # 4. AI analizi
    ai_data, model_used = None, "—"
    if chat_texts:
        print(f"\n[AI] {len(chat_texts)} chat analiz ediliyor...")
        ai_data, model_used = analyze_with_ai(chat_texts)
        if ai_data:
            print(f"[AI] {len(ai_data.get('categories', []))} şikayet kategorisi ({model_used})")
        else:
            print("[AI] Analiz başarısız")

    # 5. Rapor gönder
    html = build_html(ai_data, stats, date_str, model_used)
    send_email(html, date_str, stats)
    print(f"\n[DONE] Tamamlandı — {date_str}")

if __name__ == "__main__":
    main()
