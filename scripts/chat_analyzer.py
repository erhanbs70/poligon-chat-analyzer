#!/usr/bin/env python3
# ============================================================
# POLIGON COMM100 CHAT ANALYZER
# Günlük tüm chatleri çekip AI ile analiz eder
# Ziyaretçi mesajlarından konu/şikayet gruplar
# Email raporu gönderir
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
from collections import defaultdict

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID      = os.environ["COMM100_SITE_ID"]
API_KEY      = os.environ["COMM100_API_KEY"]
COMM100_EMAIL= os.environ["COMM100_EMAIL"]
GEMINI_KEYS  = json.loads(os.environ["GEMINI_KEYS"])
CLAUDE_KEY   = os.environ["CLAUDE_KEY"]
GROQ_KEYS    = json.loads(os.environ["GROQ_KEYS"])
GMAIL_USER   = os.environ["GMAIL_USER"]
GMAIL_PASS   = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS= [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

CAMPAIGNS = {
    "SB": "3cd6532a-0d60-4d7d-9331-5dd5be71f527",
    "BS": "2b71ad44-6a12-4073-bc6c-8edb22eabcc6",
    "TB": "6712311a-2268-408c-a803-b338a1308010"
}

BASE_URL = f"https://dash15.lively-chat.com/api/LiveChat/chats:search?siteId={SITE_ID}"

gemini_key_index = 0

# ── AUTH ─────────────────────────────────────────────────────
def get_auth():
    creds = f"{COMM100_EMAIL}:{API_KEY}"
    return "Basic " + base64.b64encode(creds.encode()).decode()

# ── DATE HELPERS ─────────────────────────────────────────────
def get_yesterday_sofia():
    """Sofia timezone'da dünün tarihini döndür"""
    from datetime import timezone
    import pytz
    sofia_tz = pytz.timezone("Europe/Sofia")
    now_sofia = datetime.now(sofia_tz)
    yesterday = now_sofia - timedelta(days=1)
    return yesterday.strftime("%Y-%m-%d")

def sofia_to_utc_range(date_str):
    """Sofia tarihini UTC start/end'e çevir"""
    import pytz
    sofia_tz = pytz.timezone("Europe/Sofia")
    start = sofia_tz.localize(datetime.strptime(date_str + " 00:00:00", "%Y-%m-%d %H:%M:%S"))
    end   = sofia_tz.localize(datetime.strptime(date_str + " 23:59:59", "%Y-%m-%d %H:%M:%S"))
    return (
        start.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        end.astimezone(pytz.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    )

# ── COMM100 FETCH ─────────────────────────────────────────────
def fetch_all_chats(date_str):
    """Günün tüm chatlerini çek (sayfalı)"""
    start_time, end_time = sofia_to_utc_range(date_str)
    auth = get_auth()
    all_chats = []
    page = 1
    seen_ids = set()

    print(f"[FETCH] Chatler çekiliyor: {date_str}")

    while page <= 40:
        url = f"{BASE_URL}&pageIndex={page}&pageSize=500&include=chatAgent&include=chatWrapup&include=postChatSurvey&sortBy=startTime&sortOrder=asc"
        try:
            r = requests.post(
                url,
                headers={"Authorization": auth, "Content-Type": "application/json"},
                json={"startTime": start_time, "endTime": end_time},
                timeout=30
            )
            if r.status_code != 200:
                print(f"[FETCH] HTTP {r.status_code} sayfa {page}")
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

    print(f"[FETCH] Toplam {len(all_chats)} chat çekildi")
    return all_chats

def fetch_chat_messages(chat_id):
    """Bir chatin mesajlarını çek"""
    url = f"https://dash15.lively-chat.com/api/LiveChat/chats/{chat_id}?include=messages&siteId={SITE_ID}"
    try:
        r = requests.get(
            url,
            headers={"Authorization": get_auth()},
            timeout=15
        )
        if r.status_code == 200:
            return r.json().get("messages", [])
    except Exception as e:
        print(f"[MSG] Chat {chat_id} mesaj hatası: {e}")
    return []

# ── CHAT HELPERS ─────────────────────────────────────────────
def get_agent_name(chat):
    agents = chat.get("chatAgents", [])
    if agents:
        names = [
            (a.get("agent", {}) or {}).get("displayName") or a.get("displayName") or a.get("name", "")
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
    """Sadece visitor mesajlarını çıkar, birleştir"""
    visitor_msgs = []
    for m in messages:
        if m.get("senderType") == "visitor":
            text = (m.get("message") or "").strip()
            if text and len(text) > 2:
                visitor_msgs.append(text)
    return " | ".join(visitor_msgs[:10])  # max 10 mesaj

# ── AI ANALİZ ─────────────────────────────────────────────────
def build_prompt(chat_texts):
    """Toplu analiz için prompt oluştur"""
    lines = []
    for i, item in enumerate(chat_texts):
        lines.append(f"{i+1}. [{item['brand']}] {item['visitor_msgs']}")

    schema = json.dumps({
        "categories": [
            {
                "name": "Spesifik konu - max 5 kelime (örn: Onaylı Çekim Hesaba Geçmedi)",
                "count": 0,
                "brand_breakdown": [{"brand": "SB", "count": 0}],
                "users": [{"chat_id": "str", "brand": "SB"}],
                "short_note": "1 cümle özet"
            }
        ],
        "summary": "max 2 cümle genel özet"
    }, ensure_ascii=False)

    return f"""Comm100 müşteri destek chat analisti olarak aşağıdaki ziyaretçi mesajlarını analiz et.

Dönem: {datetime.now().strftime('%d.%m.%Y')} | Toplam: {len(chat_texts)} chat

--- MESAJLAR ---
{chr(10).join(lines)}
--- ---

KONU ADI KURALLARI:
YANLIŞ: "Çekim Sorunları", "Teknik Sorunlar"
DOĞRU: "Onaylı Havale Çekimi Hesaba Geçmedi", "SMS Kodu Ulaşmıyor"
Aynı kategoride farklı sebepler varsa farklı grupla.

GÖREV:
1. Ziyaretçi mesajlarından asıl konuyu/şikayeti çıkar
2. Benzer konuları grupla, spesifik isim ver
3. brand_breakdown: her brand kaç chat, sadece varsa ekle
4. users: her konudaki chat_id ve brand bilgisi
5. categories listesini büyükten küçüğe sırala

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
                url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}"
                r = requests.post(url, json={
                    "contents": [{"parts": [{"text": prompt}]}],
                    "generationConfig": {"responseMimeType": "application/json"},
                    "safetySettings": [
                        {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
                    ]
                }, timeout=45)

                if r.status_code in [429, 503]:
                    break
                if r.status_code == 200:
                    res = r.json()
                    candidates = res.get("candidates", [])
                    if candidates and candidates[0].get("content"):
                        text = candidates[0]["content"]["parts"][0]["text"].strip()
                        gemini_key_index = (gemini_key_index + ki + 1) % len(GEMINI_KEYS)
                        model_name = "Gemini 2.5 Flash" if "2.5" in model else "Gemini 2.0 Flash"
                        print(f"[AI] {model_name} başarılı")
                        return {"success": True, "text": text, "model": model_name}
            except Exception as e:
                print(f"[GEMINI] {model} key#{ki} hata: {e}")

    return {"success": False}

def try_claude(prompt):
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": CLAUDE_KEY,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 2000, "messages": [{"role": "user", "content": prompt}]},
            timeout=45
        )
        if r.status_code == 200:
            text = r.json()["content"][0]["text"].strip()
            print("[AI] Claude Haiku başarılı")
            return {"success": True, "text": text, "model": "Claude Haiku 4.5"}
    except Exception as e:
        print(f"[CLAUDE] Hata: {e}")
    return {"success": False}

def try_groq(prompt):
    for i, key in enumerate(GROQ_KEYS):
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": "meta-llama/llama-4-scout-17b-16e-instruct",
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.2,
                    "response_format": {"type": "json_object"}
                },
                timeout=45
            )
            if r.status_code == 200:
                text = r.json()["choices"][0]["message"]["content"].strip()
                print(f"[AI] Groq Llama-4 Scout başarılı (key#{i})")
                return {"success": True, "text": text, "model": "Groq Llama-4 Scout"}
        except Exception as e:
            print(f"[GROQ] key#{i} hata: {e}")
    return {"success": False}

def parse_json(text):
    try:
        clean = text.replace("```json", "").replace("```", "").strip()
        return json.loads(clean)
    except:
        return None

def analyze_with_ai(chat_texts):
    if not chat_texts:
        return None, "Gemini"

    prompt = build_prompt(chat_texts)

    for fn in [try_gemini, try_claude, try_groq]:
        result = fn(prompt)
        if result["success"]:
            data = parse_json(result["text"])
            if data:
                data["model_used"] = result["model"]
                return data, result["model"]

    return None, "Hata"

# ── BATCH İŞLEME ─────────────────────────────────────────────
def process_chats(chats, date_str):
    """Chatleri filtrele ve ziyaretçi mesajlarını çek"""
    chat_texts = []
    total = len(chats)

    # Sadece agent chatleri (bot-only hariç), en fazla 300 chat analiz et
    agent_chats = [c for c in chats if not is_bot_only(c)]
    sample = agent_chats[:300]

    print(f"[PROCESS] {total} chat → {len(agent_chats)} agent chat → {len(sample)} analiz edilecek")

    for i, chat in enumerate(sample):
        chat_id = str(chat.get("id") or chat.get("chatId") or "")
        brand = get_brand(chat)

        messages = fetch_chat_messages(chat_id)
        visitor_msgs = get_visitor_messages(messages)

        if not visitor_msgs or len(visitor_msgs) < 10:
            continue

        chat_texts.append({
            "chat_id": chat_id,
            "brand": brand,
            "agent": get_agent_name(chat),
            "visitor_msgs": visitor_msgs
        })

        if (i + 1) % 50 == 0:
            print(f"[PROCESS] {i+1}/{len(sample)} chat işlendi")
        time.sleep(0.1)

    print(f"[PROCESS] {len(chat_texts)} chat mesajı analiz için hazır")
    return chat_texts

# ── HTML RAPOR ────────────────────────────────────────────────
def line_color(cnt):
    if cnt >= 8: return "#2F1555"
    if cnt >= 4: return "#662D91"
    if cnt >= 2: return "#B28ABF"
    return "#d8b4fe"

def brand_style(brand):
    styles = {
        "SB": ("background:#1d4ed8;color:#fff;", "#1d4ed8"),
        "BS": ("background:#FFE600;color:#2F1555;", "#662D91"),
        "TB": ("background:#E30613;color:#fff;", "#E30613"),
    }
    return styles.get(brand, ("background:#666;color:#fff;", "#666"))

def build_html(ai_data, stats, date_str, model_used):
    total = stats["total"]
    agent_count = stats["agent_chats"]
    bot_count = stats["bot_chats"]
    brand_counts = stats["brands"]

    # Brand bar
    brand_html = ""
    for b, cnt in brand_counts.items():
        if cnt:
            bg, _ = brand_style(b)
            brand_html += f'<span style="margin-right:16px;font-family:Montserrat,Arial,sans-serif;"><strong style="color:#FFE600;font-size:15px;">{cnt}</strong><span style="{bg}padding:1px 6px;border-radius:3px;font-size:10px;font-weight:700;margin-left:4px;">{b}</span></span>'

    # Konu satırları
    topics_html = ""
    if ai_data and ai_data.get("categories"):
        cats = sorted(ai_data["categories"], key=lambda x: x.get("count", 0), reverse=True)
        for idx, cat in enumerate(cats):
            cnt = cat.get("count", 0)
            lclr = line_color(cnt)
            is_last = idx == len(cats) - 1

            bd_html = ""
            if cat.get("brand_breakdown"):
                bd_items = ""
                for bd in cat["brand_breakdown"]:
                    bg, fg = brand_style(bd["brand"])
                    bd_items += f'<span style="display:inline-block;margin-right:8px;"><span style="display:inline-block;padding:1px 6px;{bg}border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{bd["brand"]}</span><strong style="font-size:12px;color:{fg};margin-left:3px;font-family:Montserrat,Arial,sans-serif;">{bd["count"]}</strong></span>'
                bd_html = f'<div style="margin:5px 0;"><span style="font-size:9px;font-weight:700;color:#7c3aed;text-transform:uppercase;letter-spacing:.08em;font-family:Montserrat,Arial,sans-serif;">Marka: </span>{bd_items}</div>'

            note_html = ""
            if cat.get("short_note"):
                note_html = f'<p style="margin:6px 0 0;font-size:11px;color:#662D91;font-style:italic;font-family:Montserrat,Arial,sans-serif;line-height:1.5;">{cat["short_note"]}</p>'

            border = "none" if is_last else "1px solid #f0e8ff"
            topics_html += f'''
            <table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:{border};">
              <tr>
                <td style="width:4px;background-color:{lclr};font-size:0;" bgcolor="{lclr}">&nbsp;</td>
                <td style="padding:14px 16px;background-color:#ffffff;">
                  <table width="100%" cellpadding="0" cellspacing="0"><tr>
                    <td><span style="font-size:13px;font-weight:700;color:#2F1555;font-family:Montserrat,Arial,sans-serif;">{cat.get("name","Konu")}</span></td>
                    <td style="text-align:right;white-space:nowrap;"><span style="font-size:24px;font-weight:800;color:{lclr};line-height:1;font-family:Montserrat,Arial,sans-serif;">{cnt}</span></td>
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
            <p style="margin:0 0 5px;font-size:10px;font-weight:800;color:#662D91;letter-spacing:.12em;text-transform:capitalize;font-family:Montserrat,Arial,sans-serif;">Özet</p>
            <p style="margin:0;font-size:13px;color:#4b5563;line-height:1.7;font-family:Montserrat,Arial,sans-serif;">{ai_data["summary"]}</p>
          </td></tr>
        </table>'''

    model_badge = f'<span style="display:inline-block;padding:2px 9px;background:#f3e8ff;color:#662D91;border:1px solid #B28ABF;border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{model_used}</span>'

    return f'''<!DOCTYPE html><html><head><meta charset="UTF-8">
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
  <p style="margin:0 0 8px;font-size:9px;color:#B28ABF;text-transform:uppercase;letter-spacing:.2em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">Poligon &middot; CS Ops &middot; Günlük Chat Analizi</p>
  <p style="margin:0;font-size:64px;font-weight:900;color:#FFE600;line-height:1;font-family:Montserrat,Arial,sans-serif;">{agent_count}</p>
  <p style="margin:3px 0 2px;font-size:10px;color:#B28ABF;text-transform:uppercase;letter-spacing:.22em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">AGENT CHAT</p>
  <p style="margin:0 0 16px;font-size:13px;color:#FFE600;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{date_str}</p>
  <table cellpadding="0" cellspacing="0" align="center"><tr>
    <td style="background-color:rgba(255,255,255,0.1);border-radius:20px;padding:7px 20px;border:1px solid rgba(178,138,191,0.3);">
      {brand_html}
    </td>
  </tr></table>
  <p style="margin:10px 0 0;font-family:Montserrat,Arial,sans-serif;">
    <span style="font-size:10px;color:#FFE600;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Toplam: {total}</span>
    <span style="color:rgba(178,138,191,0.5);margin:0 8px;">&middot;</span>
    <span style="font-size:10px;color:#FFE600;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Bot: {bot_count}</span>
    <span style="color:rgba(178,138,191,0.5);margin:0 8px;">&middot;</span>
    <span style="font-size:10px;color:#FFE600;font-weight:600;font-family:Montserrat,Arial,sans-serif;">Agent: {agent_count}</span>
  </p>
</td></tr>
</table>

<!-- HR -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:#f0e8ff;font-size:0;" bgcolor="#f0e8ff">&nbsp;</td>
</tr></table>

<!-- AI ANALİZİ -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:20px 28px;" bgcolor="#ffffff">

  <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr>
    <td style="font-size:11px;font-weight:800;color:#2F1555;text-transform:uppercase;letter-spacing:.14em;font-family:Montserrat,Arial,sans-serif;">AI Analizi</td>
    <td style="text-align:right;">{model_badge}</td>
  </tr></table>

  {'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e9d5ff;border-radius:8px;overflow:hidden;">' + topics_html + '</table>' if topics_html else '<p style="color:#B28ABF;font-size:13px;font-family:Montserrat,Arial,sans-serif;">Analiz yapılamadı.</p>'}

  {summary_html}

</td></tr></table>

<!-- FOOTER -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:#2F1555;" bgcolor="#2F1555">
  <p style="margin:0;font-size:11px;color:#FFE600;font-weight:700;text-align:center;font-family:Montserrat,Arial,sans-serif;">POLIGON Chat Analyzer &nbsp;&middot;&nbsp; Günlük</p>
  <p style="margin:4px 0 0;font-size:10px;color:#B28ABF;text-align:center;font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>

</td></tr></table>
</td></tr></table>
</body></html>'''

# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, date_str, stats, model_used):
    subject = f"CS Chat Analizi | {date_str} | {stats['agent_chats']} Agent Chat"

    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])

    msg.attach(MIMEText(html, "html", "utf-8"))

    all_recipients = REPORT_EMAILS

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_PASS)
            server.sendmail(GMAIL_USER, all_recipients, msg.as_string())
        print(f"[EMAIL] Gönderildi: {', '.join(all_recipients)}")
    except Exception as e:
        print(f"[EMAIL] Hata: {e}")
        raise

# ── MAIN ──────────────────────────────────────────────────────
def main():
    date_str = get_yesterday_sofia()
    print(f"\n{'='*50}")
    print(f"POLIGON CHAT ANALYZER — {date_str}")
    print(f"{'='*50}\n")

    # 1. Chatleri çek
    chats = fetch_all_chats(date_str)
    if not chats:
        print("[MAIN] Chat bulunamadı, çıkılıyor")
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
        "agent_chats": len(agent_chats),
        "bot_chats":   len(bot_chats),
        "brands":      brand_counts
    }
    print(f"[STATS] Toplam: {stats['total']} | Agent: {stats['agent_chats']} | Bot: {stats['bot_chats']}")
    print(f"[STATS] Brands: {brand_counts}")

    # 3. Mesajları çek ve hazırla
    chat_texts = process_chats(chats, date_str)

    # 4. AI analizi
    ai_data, model_used = None, "—"
    if chat_texts:
        print(f"\n[AI] {len(chat_texts)} chat analiz ediliyor...")
        ai_data, model_used = analyze_with_ai(chat_texts)
        if ai_data:
            cats = ai_data.get("categories", [])
            print(f"[AI] {len(cats)} konu grubu tespit edildi ({model_used})")
        else:
            print("[AI] Analiz başarısız")
    else:
        print("[AI] Analiz için yeterli mesaj yok")

    # 5. HTML rapor oluştur
    html = build_html(ai_data, stats, date_str, model_used)

    # 6. Email gönder
    send_email(html, date_str, stats, model_used)

    print(f"\n[DONE] Tamamlandı — {date_str}")

if __name__ == "__main__":
    main()
