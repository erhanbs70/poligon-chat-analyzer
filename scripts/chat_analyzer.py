#!/usr/bin/env python3
# ============================================================
# POLIGON COMM100 CHAT ANALYZER — v4
# fix: mesaj API çağrısı kaldırıldı (timeout'a yol açıyordu)
# fix: departmentId query string'e taşındı
# fix: Claude max_tokens 8000, Groq max_tokens eklendi
# Şikayet tespiti: rating 1-2 + QA tag'li chatler
# (zaten search API'den gelen veriler yeterli, ayrı mesaj çağrısı yok)
# ============================================================

import os
import sys
import json
import base64
import requests
import smtplib
import time
import io
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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

# QA tag'leri — bu tag'leri taşıyan chatler şikayet/sorun içerir
QA_TAGS = [
    "game_fairness", "ac_closure_request", "casino_cashback_query",
    "sport_cashback_query", "deposit_issue", "deposit_missing",
    "betting_rules_query", "deposit_query"
]

gemini_key_index = 0

# ── AUTH ─────────────────────────────────────────────────────
def get_auth():
    return "Basic " + base64.b64encode(f"{COMM100_EMAIL}:{API_KEY}".encode()).decode()

# ── DATE ─────────────────────────────────────────────────────
def get_yesterday_sofia():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return (datetime.now(tz) - timedelta(days=1)).strftime("%Y-%m-%d")

def sofia_to_utc_range(date_str):
    import pytz
    tz    = pytz.timezone("Europe/Sofia")
    start = tz.localize(datetime.strptime(date_str + " 00:00:00", "%Y-%m-%d %H:%M:%S"))
    end   = tz.localize(datetime.strptime(date_str + " 23:59:59", "%Y-%m-%d %H:%M:%S"))
    utc   = pytz.utc
    return (
        start.astimezone(utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        end.astimezone(utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    )

def in_sofia_range(ts_str, date_str):
    if not ts_str:
        return True
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        return dt.astimezone(tz).strftime("%Y-%m-%d") == date_str
    except Exception:
        return True

# ── COMM100 FETCH ─────────────────────────────────────────────
def fetch_cs_chats(date_str):
    """
    CS departmanı chatlerini çek.
    departmentId query string'de — POST body filtresi çalışmıyor.
    include=chatWrapup: tag bilgisi için
    include=postChatSurvey: rating bilgisi için
    include=messages YOK — ayrı API çağrısına gerek yok
    """
    start_time, end_time = sofia_to_utc_range(date_str)
    auth     = get_auth()
    result   = []
    seen_ids = set()
    page     = 1

    print(f"[FETCH] CS chatleri çekiliyor: {date_str}")

    while page <= 40:
        url = (
            f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
            f"?siteId={SITE_ID}"
            f"&pageIndex={page}&pageSize=500"
            f"&include=chatAgent&include=chatWrapup&include=postChatSurvey"
            f"&sortBy=startTime&sortOrder=asc"
            f"&departmentId={CS_DEPT_ID}"
        )
        try:
            r = requests.post(
                url,
                headers={"Authorization": auth, "Content-Type": "application/json"},
                json={"startTime": start_time, "endTime": end_time},
                timeout=90,
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
                    if in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
                        result.append(c)

            print(f"[FETCH] Sayfa {page}: {len(chats)} chat ({len(result)} toplam)")
            if len(chats) < 500:
                break
            page += 1
            time.sleep(0.3)

        except Exception as e:
            print(f"[FETCH] Hata sayfa {page}: {e}")
            break

    print(f"[FETCH] Toplam {len(result)} CS chat çekildi")
    return result

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
    url = (chat.get("requestingPageURL") or "").lower()
    if "superbetin" in url: return "SB"
    if "betsat"     in url: return "BS"
    if "turkbet"    in url: return "TB"
    return "?"

def is_bot_only(chat):
    return chat.get("chatType") == "chatBotOnly"

def get_tag(chat):
    """chatWrapup.categoriesName'den ana tag'i çıkar."""
    import re
    cat = (chat.get("chatWrapup") or {}).get("categoriesName") or ""
    if not cat:
        return ""
    inner = [m.group(1).strip() for m in re.finditer(r"\(([^)]+)\)", cat)
             if not re.search(r"VIP", m.group(1), re.I)]
    for t in inner:
        if "_" in t:
            return t
    if inner:
        return inner[0]
    parts = [p.strip() for p in re.sub(r"\([^)]*\)", "", cat).split(",") if p.strip()]
    for p in parts:
        if "_" in p:
            return p
    return parts[0] if parts else ""

def get_rating(chat):
    s = chat.get("postChatSurvey") or {}
    g = s.get("ratingGrade")
    return int(g) if g is not None else None

def get_rating_comment(chat):
    s = chat.get("postChatSurvey") or {}
    return (s.get("ratingComment") or "").strip()

def is_complaint_chat(chat):
    """
    Şikayet/sorun göstergesi:
    1. Rating 1 veya 2
    2. QA tag'li (deposit_issue, game_fairness vb.)
    3. Rating comment varsa
    """
    rating  = get_rating(chat)
    tag     = get_tag(chat).lower()
    comment = get_rating_comment(chat)

    if rating in (1, 2):
        return True
    if any(qa in tag for qa in QA_TAGS):
        return True
    if comment and len(comment) > 10:
        return True
    return False

# ── PROCESS ──────────────────────────────────────────────────
def process_chats(chats):
    """
    Şikayet içeren chatleri tespit et.
    Mesaj API çağrısı YOK — tag + rating zaten mevcut.
    """
    agent_chats     = [c for c in chats if not is_bot_only(c)]
    complaint_chats = [c for c in agent_chats if is_complaint_chat(c)]

    print(f"[PROCESS] {len(chats)} CS chat → {len(agent_chats)} agent → {len(complaint_chats)} şikayet")

    # AI için chat özetleri oluştur
    chat_texts = []
    for c in complaint_chats:
        brand   = get_brand(c)
        tag     = get_tag(c) or "—"
        rating  = get_rating(c)
        comment = get_rating_comment(c)
        agent   = get_agent_name(c)

        # AI'a gönderilecek özet satır
        parts = [f"[{brand}]", f"Tag:{tag}"]
        if rating:
            parts.append(f"Rating:{rating}")
        if comment:
            parts.append(f"Yorum:{comment[:120]}")
        # Agent bilgisi kasıtlı çıkarıldı — AI müşteri ile karıştırıyor

        chat_texts.append({
            "brand":   brand,
            "tag":     tag,
            "rating":  rating,
            "comment": comment,
            "summary": " | ".join(parts)
        })

    print(f"[PROCESS] {len(chat_texts)} chat AI için hazır")
    return agent_chats, complaint_chats, chat_texts

# ── AI ────────────────────────────────────────────────────────
def build_prompt(chat_texts, date_str):
    lines = [f"{i+1}. {c['summary']}" for i, c in enumerate(chat_texts)]

    schema = json.dumps({
        "categories": [{
            "name":            "Spesifik şikayet/sorun adı — max 6 kelime",
            "count":           0,
            "brand_breakdown": [{"brand": "SB", "count": 0}],
            "short_note":      "1 cümle somut özet — rakam/detay içersin"
        }],
        "critical": [{
            "username": "kullanıcı adı",
            "brand":    "SB/BS/TB",
            "reason":   "Neden kritik — tutar, tehdit veya aciliyet — 1 cümle"
        }],
        "action_items": [
            "Departman + yapılacak aksiyon — 1 cümle"
        ],
        "summary": "3-4 cümle genel değerlendirme — dominant sorun, trend, önemli kullanıcı adları"
    }, ensure_ascii=False)

    return f"""Comm100 CS departmanı şikayet analisti olarak aşağıdaki verileri analiz et.

Tarih: {date_str} | Toplam şikayet/sorun: {len(chat_texts)} chat

--- VERİLER (brand | tag | rating | kullanıcı yorumu) ---
NOT: Verilerdeki isimler ŞİKAYET EDEN MÜŞTERİ adlarıdır, CS temsilcisi değil.
{chr(10).join(lines)}
--- ---

KURAL 1 — KONU ADI SPESİFİK OLMALI:
YANLIŞ: "Çekim Sorunları", "Diğer", "Genel Sorun"
DOĞRU: "Onaylı Havale Çekimi Hesaba Geçmedi", "Papara Yatırımı Yansımadı", "Bonus Aktivasyon Yapılmıyor"

KURAL 2 — Aynı sorunu farklı tag/yorumla ifade edenler TEK kategori altında toplan.
KURAL 3 — "Diğer" kategorisi YASAK.
KURAL 4 — brand_breakdown: sadece o konuda hangi brand kaç chat var.
KURAL 5 — short_note: somut, spesifik, rakam/detay içersin.
KURAL 6 — critical: yüksek tutar (5000 TL+), hesap kapatma/silinme tehdidi, acil çözüm bekleyen, Rating 1 veren MÜŞTERİLER. username alanına ŞİKAYET EDEN MÜŞTERİNİN adını yaz, CS temsilcisi adını ASLA yazma. Yoksa boş liste.
KURAL 7 — action_items: en az 2, en fazla 5. Her biri "Finans Departmanı: ..." formatında hangi ekip ne yapmalı.
KURAL 8 — summary: 3-4 cümle. Dominant sorun, brand dağılımı, dikkat çeken trend ve önemli kullanıcı varsa isim yaz.

GÖREV: Şikayet/sorunları konulara göre grupla, büyükten küçüğe sırala.
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
                        "contents":         [{"parts": [{"text": prompt}]}],
                        "generationConfig": {"responseMimeType": "application/json"},
                        "safetySettings": [
                            {"category": "HARM_CATEGORY_HARASSMENT",        "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_HATE_SPEECH",       "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"}
                        ]
                    }, timeout=60
                )
                if r.status_code in [429, 503]:
                    continue
                if r.status_code == 200:
                    res   = r.json()
                    cands = res.get("candidates", [])
                    if cands and cands[0].get("content"):
                        text = cands[0]["content"]["parts"][0]["text"].strip()
                        gemini_key_index = (gemini_key_index + ki + 1) % len(GEMINI_KEYS)
                        model_name = "Gemini 2.5 Flash" if "2.5" in model else "Gemini 2.0 Flash"
                        in_t  = res.get("usageMetadata", {}).get("promptTokenCount", 0)
                        out_t = res.get("usageMetadata", {}).get("candidatesTokenCount", 0)
                        cost  = (in_t * 0.075 / 1_000_000) + (out_t * 0.30 / 1_000_000)
                        print(f"[AI] {model_name} — {in_t}+{out_t} token | ${cost:.6f}")
                        return {"success": True, "text": text, "model": model_name,
                                "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
            except Exception as e:
                print(f"[GEMINI] {model} key#{ki}: {e}")
    return {"success": False}

def try_claude(prompt):
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": CLAUDE_KEY, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model":      "claude-haiku-4-5-20251001",
                  "max_tokens": 8000,
                  "messages":   [{"role": "user", "content": prompt}]},
            timeout=60
        )
        if r.status_code == 200:
            res   = r.json()
            text  = res["content"][0]["text"].strip()
            in_t  = res.get("usage", {}).get("input_tokens", 0)
            out_t = res.get("usage", {}).get("output_tokens", 0)
            cost  = (in_t * 0.25 / 1_000_000) + (out_t * 1.25 / 1_000_000)
            print(f"[AI] Claude Haiku — {in_t}+{out_t} token | ${cost:.6f}")
            return {"success": True, "text": text, "model": "Claude Haiku 4.5",
                    "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
    except Exception as e:
        print(f"[CLAUDE] {e}")
    return {"success": False}

def try_groq(prompt):
    for i, key in enumerate(GROQ_KEYS):
        try:
            r = requests.post(
                "https://api.groq.com/openai/v1/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={"model":           "meta-llama/llama-4-scout-17b-16e-instruct",
                      "messages":        [{"role": "user", "content": prompt}],
                      "max_tokens":      8000,
                      "temperature":     0.2,
                      "response_format": {"type": "json_object"}},
                timeout=60
            )
            if r.status_code == 200:
                res   = r.json()
                text  = res["choices"][0]["message"]["content"].strip()
                in_t  = res.get("usage", {}).get("prompt_tokens", 0)
                out_t = res.get("usage", {}).get("completion_tokens", 0)
                cost  = (in_t * 0.11 / 1_000_000) + (out_t * 0.34 / 1_000_000)
                print(f"[AI] Groq — {in_t}+{out_t} token | ${cost:.6f}")
                return {"success": True, "text": text, "model": "Groq Llama-4 Scout",
                        "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
        except Exception as e:
            print(f"[GROQ] key#{i}: {e}")
    return {"success": False}

def parse_json(text):
    try:
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception:
        return None

def analyze_with_ai(chat_texts, date_str):
    if not chat_texts:
        return None, "—", {}
    prompt = build_prompt(chat_texts, date_str)
    for fn in [try_gemini, try_claude, try_groq]:
        result = fn(prompt)
        if result["success"]:
            data = parse_json(result["text"])
            if data:
                data["model_used"] = result["model"]
                usage = {"model":     result["model"],
                         "in_tokens":  result.get("in_tokens", 0),
                         "out_tokens": result.get("out_tokens", 0),
                         "cost":       result.get("cost", 0)}
                return data, result["model"], usage
    return None, "Hata", {}

# ── HTML ──────────────────────────────────────────────────────
def line_color(cnt):
    if cnt >= 8: return "#2F1555"
    if cnt >= 4: return "#662D91"
    if cnt >= 2: return "#B28ABF"
    return "#d8b4fe"

def brand_badge(brand):
    styles = {
        "SB": "background:#1d4ed8;color:#fff;",
        "BS": "background:#FFE600;color:#2F1555;",
        "TB": "background:#E30613;color:#fff;",
    }
    s = styles.get(brand, "background:#666;color:#fff;")
    return f'<span style="display:inline-block;padding:1px 6px;{s}border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{brand}</span>'

def build_html(ai_data, stats, date_str, model_used, ai_usage=None):
    total          = stats["total"]
    agent_count    = stats["agent_count"]
    bot_count      = stats["bot_count"]
    complaint_count= stats["complaint_count"]
    brand_counts   = stats["brands"]

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

    topics_html = ""
    if ai_data and ai_data.get("categories"):
        cats = sorted(ai_data["categories"], key=lambda x: x.get("count", 0), reverse=True)
        for idx, cat in enumerate(cats):
            cnt    = cat.get("count", 0)
            lclr   = line_color(cnt)
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
            <p style="margin:0;font-size:13px;color:#4b5563;line-height:1.7;
               font-family:Montserrat,Arial,sans-serif;">{ai_data["summary"]}</p>
          </td></tr>
        </table>'''

    model_badge_html = (
        f'<span style="display:inline-block;padding:2px 9px;background:#f3e8ff;color:#662D91;'
        f'border:1px solid #B28ABF;border-radius:3px;font-size:10px;font-weight:700;'
        f'font-family:Montserrat,Arial,sans-serif;">{model_used}</span>'
    )

    # ── Kritik kullanıcılar ──
    critical_html = ""
    if ai_data and ai_data.get("critical"):
        crits = ai_data["critical"]
        if crits:
            rows = ""
            for u in crits:
                bt_styles = {"SB": "background:#1d4ed8;color:#fff;",
                             "BS": "background:#FFE600;color:#2F1555;",
                             "TB": "background:#E30613;color:#fff;"}
                bs = bt_styles.get(u.get("brand",""), "background:#666;color:#fff;")
                rows += (
                    f'<tr style="border-bottom:1px solid #f0e8ff;"><td style="padding:9px 14px;white-space:nowrap;font-size:13px;font-weight:700;color:#2F1555;font-family:Montserrat,Arial,sans-serif;">{u.get("username","-")}</td><td style="padding:9px 14px;"><span style="display:inline-block;padding:1px 7px;{bs}border-radius:3px;font-size:10px;font-weight:700;">{u.get("brand","-")}</span></td><td style="padding:9px 14px;font-size:12px;color:#662D91;line-height:1.5;font-family:Montserrat,Arial,sans-serif;">{u.get("reason","-")}</td></tr>'
                )
            critical_html = (
                f'<p style="margin:22px 0 10px;font-size:11px;font-weight:800;color:#2F1555;text-transform:uppercase;letter-spacing:.12em;font-family:Montserrat,Arial,sans-serif;">🚨 Kritik Kullanıcılar</p><table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e9d5ff;border-top:3px solid #FFE600;border-radius:0 0 8px 8px;">{rows}</table>'
            )

    # ── Aksiyon önerileri ──
    actions_html = ""
    if ai_data and ai_data.get("action_items"):
        items = ai_data["action_items"]
        if items:
            rows = "".join(
                f'<table cellpadding="0" cellspacing="0" style="margin-bottom:9px;width:100%;"><tr><td style="vertical-align:top;width:24px;padding-right:10px;"><span style="display:inline-block;width:20px;height:20px;background:#2F1555;color:#FFE600;border-radius:50%;font-size:10px;font-weight:700;text-align:center;line-height:20px;font-family:Montserrat,Arial,sans-serif;">{i+1}</span></td><td style="font-size:13px;color:#4b5563;line-height:1.6;vertical-align:top;font-family:Montserrat,Arial,sans-serif;">{a}</td></tr></table>'
                for i, a in enumerate(items)
            )
            actions_html = (
                f'<p style="margin:22px 0 10px;font-size:11px;font-weight:800;color:#2F1555;text-transform:uppercase;letter-spacing:.12em;font-family:Montserrat,Arial,sans-serif;">✅ Önerilen Aksiyonlar</p>{rows}'
            )

    no_data_html = '<p style="color:#B28ABF;font-size:13px;font-family:Montserrat,Arial,sans-serif;">Şikayet/sorun tespit edilemedi.</p>'

    if ai_usage and ai_usage.get("in_tokens"):
        u = ai_usage
        token_line = (
            f'<p style="margin:5px 0 0;font-size:10px;color:#B28ABF;text-align:center;'
            f'font-family:Montserrat,Arial,sans-serif;">'
            f'{u["model"]} &nbsp;&middot;&nbsp; '
            f'{u["in_tokens"]:,} input + {u["out_tokens"]:,} output token &nbsp;&middot;&nbsp; '
            f'${u["cost"]:.6f}</p>'
        )
    else:
        token_line = ""

    complaint_pct = f"{complaint_count/agent_count*100:.1f}%" if agent_count else "—"

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
    <span style="color:rgba(178,138,191,0.4);margin:0 8px;">&middot;</span>
    <span style="font-size:10px;color:#FFE600;font-weight:700;font-family:Montserrat,Arial,sans-serif;">Şikayet: {complaint_count} ({complaint_pct})</span>
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
  {critical_html}
  {actions_html}
  {summary_html}
</td></tr>
</table>

<!-- FOOTER -->
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:#2F1555;" bgcolor="#2F1555">
  <p style="margin:0;font-size:11px;color:#FFE600;font-weight:700;text-align:center;
     font-family:Montserrat,Arial,sans-serif;">POLIGON Chat Analyzer &nbsp;&middot;&nbsp; CS Günlük</p>
  {token_line}
  <p style="margin:4px 0 0;font-size:10px;color:#B28ABF;text-align:center;
     font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>

</td></tr></table>
</td></tr></table>
</body></html>'''

# ── EXCEL EKİ ────────────────────────────────────────────────
def build_excel(complaint_chats, date_str):
    """Şikayet listesini Excel olarak oluştur (openpyxl)."""
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Şikayetler {date_str}"

    # Renkler
    HDR_BG  = "2F1555"
    HDR_FG  = "FFE600"
    ROW_ALT = "F3F0FF"
    RED_BG  = "F28B82"
    YLW_BG  = "FBBC04"

    headers = ["#", "Kullanıcı Adı", "Brand", "Agent", "Tag", "Rating", "Yorum", "Tarih", "Chat Linki"]
    col_widths = [5, 22, 8, 22, 30, 8, 60, 20, 55]

    # Header satırı
    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = Font(bold=True, color=HDR_FG, name="Calibri", size=10)
        cell.fill      = PatternFill("solid", fgColor=HDR_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    thin = Side(style="thin", color="E0D8F0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    portal = "https://dash15.lively-chat.com/ui/90005373/livechat/history/chats/transcriptdetail"

    for idx, c in enumerate(complaint_chats, 1):
        row = idx + 1
        bg  = ROW_ALT if idx % 2 == 0 else "FFFFFF"

        username = c.get("preChatName") or c.get("name") or c.get("visitorName") or ""
        brand    = get_brand(c)
        agent    = get_agent_name(c)
        tag      = get_tag(c) or "—"
        rating   = get_rating(c) or ""
        comment  = get_rating_comment(c)
        ts       = c.get("startTime") or c.get("start_time") or ""
        cid      = str(c.get("id") or c.get("chatId") or "")
        link     = f"{portal}?chatId={cid}" if cid else ""

        # Tarih formatla
        try:
            import pytz
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
            ts_fmt = dt.astimezone(pytz.timezone("Europe/Sofia")).strftime("%d/%m/%Y %H:%M")
        except Exception:
            ts_fmt = ts

        values = [idx, username, brand, agent, tag, rating, comment, ts_fmt, link]

        # Rating'e göre satır rengi
        if rating == 1:
            bg = RED_BG
        elif rating == 2:
            bg = YLW_BG

        for ci, val in enumerate(values, 1):
            cell = ws.cell(row=row, column=ci, value=val)
            cell.fill      = PatternFill("solid", fgColor=bg)
            cell.font      = Font(name="Calibri", size=9)
            cell.alignment = Alignment(vertical="center", wrap_text=(ci == 7))
            cell.border    = border
            if ci == 9 and val:  # Link sütunu
                cell.hyperlink = val
                cell.value     = "Chati Aç"
                cell.font      = Font(name="Calibri", size=9, color="1155CC", underline="single")

        ws.row_dimensions[row].height = 18

    # Auto-filter
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, date_str, stats, complaint_chats=None):
    subject = (f"CS Şikayet Analizi | {date_str} | "
               f"{stats['agent_count']} Agent | {stats['complaint_count']} Şikayet")
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])

    # HTML body
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    # Excel eki
    if complaint_chats:
        xlsx_bytes = build_excel(complaint_chats, date_str)
        if xlsx_bytes:
            attachment = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            attachment.set_payload(xlsx_bytes)
            encoders.encode_base64(attachment)
            attachment.add_header("Content-Disposition", "attachment",
                                  filename=f"sikayet_{date_str}.xlsx")
            msg.attach(attachment)
            print(f"[EMAIL] Excel eki hazır: {len(xlsx_bytes)//1024}KB")

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
    print(f"[EMAIL] Gönderildi: {', '.join(REPORT_EMAILS)}")

# ── MAIN ──────────────────────────────────────────────────────
def main():
    date_str = get_yesterday_sofia()
    print(f"\n{'='*55}\nPOLIGON CS CHAT ANALYZER — {date_str}\n{'='*55}\n")

    chats = fetch_cs_chats(date_str)
    if not chats:
        print("[MAIN] CS chat bulunamadı")
        return

    bot_chats    = [c for c in chats if is_bot_only(c)]
    agent_chats  = [c for c in chats if not is_bot_only(c)]
    brand_counts = {"SB": 0, "BS": 0, "TB": 0}
    for c in agent_chats:
        b = get_brand(c)
        if b in brand_counts:
            brand_counts[b] += 1

    _, complaint_chats, chat_texts = process_chats(chats)

    stats = {
        "total":           len(chats),
        "agent_count":     len(agent_chats),
        "bot_count":       len(bot_chats),
        "complaint_count": len(complaint_chats),
        "brands":          brand_counts
    }
    print(f"[STATS] Toplam:{stats['total']} | Agent:{stats['agent_count']} | "
          f"Bot:{stats['bot_count']} | Şikayet:{stats['complaint_count']}")

    ai_data, model_used, ai_usage = None, "—", {}
    if chat_texts:
        print(f"\n[AI] {len(chat_texts)} şikayet analiz ediliyor...")
        ai_data, model_used, ai_usage = analyze_with_ai(chat_texts, date_str)
        if ai_data:
            print(f"[AI] {len(ai_data.get('categories', []))} kategori ({model_used})")
        else:
            print("[AI] Analiz başarısız")

    html = build_html(ai_data, stats, date_str, model_used, ai_usage)
    send_email(html, date_str, stats, complaint_chats=complaint_chats)
    print(f"\n[DONE] Tamamlandı — {date_str}")

if __name__ == "__main__":
    main()
