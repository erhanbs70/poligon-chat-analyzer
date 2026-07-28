#!/usr/bin/env python3
# ============================================================
# POLIGON COMM100 CHAT ANALYZER — v6
# fix: parse_json hata loglama + Claude JSON prefill + Groq boş
#      key guard + AI analizi başarısız olduğunda email'de görünür uyarı
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

QA_TAGS = [
    "game_fairness",
    "ac_closure_request",
    "deposit_issue",
    "deposit_missing",
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
    desc sırayla çek (en yeni önce).
    Bugün (sonra) → dün (hedef) → önceki günler (önce/dur).
    """
    import pytz
    start_time, end_time = sofia_to_utc_range(date_str)
    auth     = get_auth()
    result   = []
    seen_ids = set()
    page     = 1
    sofia_tz = pytz.timezone("Europe/Sofia")

    print(f"[FETCH] CS chatleri çekiliyor: {date_str}")

    while page <= 40:
        url = (
            f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
            f"?siteId={SITE_ID}"
            f"&pageIndex={page}&pageSize=500"
            f"&include=chatAgent&include=chatWrapup&include=postChatSurvey"
            f"&sortBy=startTime&sortOrder=desc"
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

            page_in     = 0  # hedef tarih
            page_after  = 0  # hedef tarihten SONRA (henüz gelmedik)
            page_before = 0  # hedef tarihten ÖNCE (geçtik, dur)

            for c in chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                ts  = c.get("startTime") or c.get("start_time") or ""

                sofia_date = None
                if ts:
                    try:
                        dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
                        sofia_date = dt.astimezone(sofia_tz).strftime("%Y-%m-%d")
                    except Exception:
                        pass

                if sofia_date == date_str:
                    page_in += 1
                elif sofia_date and sofia_date > date_str:
                    page_after += 1
                elif sofia_date and sofia_date < date_str:
                    page_before += 1

                if cid and cid not in seen_ids:
                    seen_ids.add(cid)
                    if not sofia_date or sofia_date == date_str:
                        result.append(c)

            print(f"[FETCH] Sayfa {page}: {len(chats)} chat | "
                  f"hedef:{page_in} sonra:{page_after} once:{page_before} ({len(result)} toplam)")

            # desc sıralamada:
            # sonra>0 ve hedef=0 → henüz dünün chatlerine gelmedik, devam
            # hedef>0            → dünün chatlerini buluyoruz, devam
            # once>0 ve hedef=0 → dünü geçtik, dur
            if page_before > 0 and page_in == 0:
                print("[FETCH] Hedef tarih geçildi, durduruluyor.")
                break

            if len(chats) < 500:
                break

            page += 1
            time.sleep(0.3)

        except Exception as e:
            print(f"[FETCH] Hata sayfa {page}: {e}")
            break

    print(f"[FETCH] Toplam {len(result)} CS chat çekildi")
    return result

# ── MESAJ FETCH ──────────────────────────────────────────────
def fetch_messages_batch(chat_ids):
    auth    = get_auth()
    results = {}
    total   = len(chat_ids)

    for i, cid in enumerate(chat_ids):
        url = (
            f"https://dash15.lively-chat.com/api/LiveChat/chats/{cid}"
            f"?siteId={SITE_ID}&include=messages"
        )
        try:
            r = requests.get(url, headers={"Authorization": auth}, timeout=20)
            if r.status_code == 200:
                messages = r.json().get("messages", [])
                visitor_msgs = []
                for m in messages:
                    if m.get("senderType") in ("visitor", "Visitor"):
                        text = (m.get("message") or m.get("body") or "").strip()
                        if text and len(text) > 5:
                            visitor_msgs.append(text[:200])
                    if len(visitor_msgs) >= 5:
                        break
                results[cid] = " | ".join(visitor_msgs)
        except Exception:
            results[cid] = ""

        if (i + 1) % 10 == 0:
            time.sleep(0.5)
        if (i + 1) % 100 == 0:
            print(f"[MSG] {i+1}/{total} mesaj çekildi...")

    print(f"[MSG] Toplam {len(results)} chat mesajı çekildi")
    return results

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

COMPLAINT_KEYWORDS = [
    "şikayet", "mağdur", "dolandırıcı", "sahtekâr", "rezalet", "berbat",
    "korkunç", "çok kötü", "iğrenç", "skandal", "mahkeme", "avukat",
    "paranı ver", "paramı ver", "param nerede", "param kayboldu",
    "yatırım gelmedi", "yatırım yok", "para yok", "para gelmiyor",
    "çekim gelmiyor", "çekim yok", "para çıkmıyor", "ödeme yok",
    "ödeme gelmiyor", "ödeme yapılmadı", "hesaba geçmedi", "yansımadı",
    "hile", "hileli", "manipüle", "oyun hileliydi", "kazandım ama",
    "açılmıyor", "açılmıyo", "girilmiyor", "giremiyorum", "giriş yapamıyorum",
    "donuyor", "dondu", "takıldı", "takılıyor", "kasıyor", "kasılıyor",
    "yavaş", "çok yavaş", "site yavaş", "uygulama yavaş",
    "çöküyor", "çöktü", "kapanıyor", "kapandı", "hata veriyor",
    "hata aldım", "error", "bağlanamıyorum", "bağlantı yok",
    "giriş yapamıyorum", "şifre çalışmıyor", "sms gelmiyor",
    "doğrulama gelmiyor", "kod gelmiyor",
    "neden hâlâ", "neden hala", "hala çözülmedi", "çözülmedi",
    "çözüm yok", "ilgilenmiyor", "ilgilenilmiyor", "cevap yok",
    "cevap vermedi", "cevap verilmiyor", "bekletiyorsunuz",
    "saatlerdir", "günlerdir", "haftadır", "bekliyorum",
    "mağdur ettiniz", "zarar gördüm", "zarar ettim",
    "yanlış hesaplandı", "hatalı", "eksik yatırıldı",
    "bonus verilmedi", "bonus gelmedi", "bonus yok",
    "sinir bozucu", "berbat site", "kötü site", "rezil",
    "şikayetvar", "sikayetvar", "twitter", "sosyal medya",
    "btcm", "şikayet edeceğim", "şikayet açacağım",
    "mahkemeye vereceğim", "avukata vereceğim",
    "hesabımı kapat", "hesabı sil", "üyeliğimi iptal",
    "küfür", "hakaret", "terbiyesiz", "saygısız",
    "orospu", "siktir", "amk", "bok", "göt", "piç",
]

def is_complaint_chat(chat, visitor_msgs=""):
    rating  = get_rating(chat)
    tag     = get_tag(chat).lower()
    comment = get_rating_comment(chat)

    if rating in (1, 2):
        return True
    if any(qa in tag for qa in QA_TAGS):
        return True
    if comment and len(comment) > 30:
        return True
    if visitor_msgs:
        msgs_lower = visitor_msgs.lower()
        if any(kw in msgs_lower for kw in COMPLAINT_KEYWORDS):
            return True
    return False

# ── PROCESS ──────────────────────────────────────────────────
def process_chats(chats):
    agent_chats = [c for c in chats if not is_bot_only(c)]
    print(f"[PROCESS] {len(chats)} CS chat → {len(agent_chats)} agent chat")

    quick_complaints = set()
    remaining_ids    = []
    for c in agent_chats:
        cid = str(c.get("id") or c.get("chatId") or "")
        if is_complaint_chat(c, ""):
            quick_complaints.add(cid)
        else:
            remaining_ids.append(cid)

    print(f"[PROCESS] Hızlı filtre: {len(quick_complaints)} şikayet, {len(remaining_ids)} mesaj kontrolü bekliyor")

    msg_map = {}
    if remaining_ids:
        print(f"[MSG] {len(remaining_ids)} chat için mesajlar çekiliyor...")
        msg_map = fetch_messages_batch(remaining_ids)

    complaint_chats = []
    for c in agent_chats:
        cid          = str(c.get("id") or c.get("chatId") or "")
        visitor_msgs = msg_map.get(cid, "")
        if cid in quick_complaints or is_complaint_chat(c, visitor_msgs):
            c["_visitor_msgs"] = visitor_msgs
            complaint_chats.append(c)

    print(f"[PROCESS] Toplam {len(complaint_chats)} şikayet tespit edildi")

    chat_texts = []
    for c in complaint_chats:
        brand        = get_brand(c)
        tag          = get_tag(c) or "—"
        rating       = get_rating(c)
        comment      = get_rating_comment(c)
        visitor_msgs = c.get("_visitor_msgs", "")

        visitor = (c.get("preChatName") or c.get("name") or
                   c.get("visitorName") or c.get("visitor_name") or "—")

        parts = [f"Uye:{visitor}", f"[{brand}]", f"Tag:{tag}"]
        if rating:
            parts.append(f"Rating:{rating}")
        if comment:
            parts.append(f"Yorum:{comment[:120]}")
        elif visitor_msgs:
            parts.append(f"Mesaj:{visitor_msgs[:150]}")

        chat_texts.append({
            "brand":    brand,
            "tag":      tag,
            "rating":   rating,
            "username": visitor,
            "comment":  comment or visitor_msgs[:120],
            "summary":  " | ".join(parts)
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
            "username": "satır başındaki Uye:XXX değeri — TAM kopyala",
            "brand":    "SB/BS/TB",
            "reason":   "Neden kritik — tutar, tehdit veya aciliyet — 1 cümle"
        }],
        "summary": "3-4 cümle genel değerlendirme — dominant sorun, trend, önemli kullanıcı adları"
    }, ensure_ascii=False)

    return f"""Comm100 CS departmanı şikayet analisti olarak aşağıdaki verileri analiz et.

Tarih: {date_str} | Toplam şikayet/sorun: {len(chat_texts)} chat

--- VERİLER (Uye:KULLANICI_ADI | brand | tag | rating | yorum/mesaj) ---
NOT: "Uye:" ile başlayan isimler ŞİKAYET EDEN MÜŞTERİ kullanıcı adlarıdır.
CS temsilcisi adları bu veride YOKTUR.
{chr(10).join(lines)}
--- ---

KURAL 1 — KONU ADI SPESİFİK OLMALI:
YANLIŞ: "Çekim Sorunları", "Diğer", "Genel Sorun"
DOĞRU: "Onaylı Havale Çekimi Hesaba Geçmedi", "Papara Yatırımı Yansımadı", "Bonus Aktivasyon Yapılmıyor"

KURAL 2 — Aynı sorunu farklı tag/yorumla ifade edenler TEK kategori altında toplan.
KURAL 3 — "Diğer" kategorisi YASAK.
KURAL 4 — brand_breakdown: sadece o konuda hangi brand kaç chat var.
KURAL 5 — short_note: somut, spesifik, rakam/detay içersin.
KURAL 6 — critical: SADECE şu 3 durumdan biri varsa ekle:
  a) 5000 TL+ tutar kaybı/çekim sorunu belirtilmişse
  b) Hesap silme/kapatma tehdidi + Rating 1 birlikte varsa
  c) Açık tehdit, ölüm tehdidi veya hukuki süreç başlatacağını belirten mesaj varsa
  Genel "memnun değilim" veya sadece Rating 1 olan KRİTİK DEĞİLDİR.
  username = satır başındaki "Uye:XXX" kısmındaki XXX değeri — TAM OLARAK kopyala.
  CS temsilcisi adını ASLA yazma. Kritik yoksa boş liste: "critical": []
KURAL 7 — summary: 3-4 cümle. Dominant sorun, brand dağılımı, dikkat çeken trend.
  ÖNEMLİ: Özette CS temsilcisi adlarını (agent) ASLA kullanıcı gibi gösterme.
  Eğer küfür/hakaret içeren yorumlar varsa "bazı müşteriler sert dil kullandı" şeklinde yaz.

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
                    print(f"[GEMINI] {model} key#{ki}: HTTP {r.status_code} (rate limit/kullanılamıyor)")
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
                    else:
                        print(f"[GEMINI] {model} key#{ki}: HTTP 200 ama candidate boş döndü")
                else:
                    print(f"[GEMINI] {model} key#{ki}: HTTP {r.status_code}: {r.text[:200]}")
            except Exception as e:
                print(f"[GEMINI] {model} key#{ki}: {e}")
    return {"success": False}

def try_claude(prompt):
    """
    NOT: JSON prefill kullanılıyor — assistant mesajını "{" ile başlatarak
    Claude'un yanıtın başına açıklama/markdown eklemesini engelliyoruz.
    Böylece parse_json() daha güvenilir çalışıyor.
    """
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": CLAUDE_KEY, "anthropic-version": "2023-06-01",
                     "content-type": "application/json"},
            json={"model":      "claude-haiku-4-5-20251001",
                  "max_tokens": 8000,
                  "messages":   [
                      {"role": "user", "content": prompt},
                      {"role": "assistant", "content": "{"}
                  ]},
            timeout=60
        )
        if r.status_code == 200:
            res   = r.json()
            text  = "{" + res["content"][0]["text"].strip()
            in_t  = res.get("usage", {}).get("input_tokens", 0)
            out_t = res.get("usage", {}).get("output_tokens", 0)
            cost  = (in_t * 0.25 / 1_000_000) + (out_t * 1.25 / 1_000_000)
            print(f"[AI] Claude Haiku — {in_t}+{out_t} token | ${cost:.6f}")
            return {"success": True, "text": text, "model": "Claude Haiku 4.5",
                    "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
        else:
            print(f"[CLAUDE] HTTP {r.status_code}: {r.text[:200]}")
    except Exception as e:
        print(f"[CLAUDE] {e}")
    return {"success": False}

def try_groq(prompt):
    if not GROQ_KEYS:
        print("[GROQ] GROQ_KEYS boş/tanımsız, atlanıyor")
        return {"success": False}
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
            else:
                print(f"[GROQ] key#{i} HTTP {r.status_code}: {r.text[:200]}")
        except Exception as e:
            print(f"[GROQ] key#{i}: {e}")
    return {"success": False}

def parse_json(text):
    try:
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except Exception as e:
        print(f"[PARSE] Hata: {e} | text[:300]={text[:300]!r}")
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
            else:
                print(f"[AI] {result['model']} yanıt verdi ama JSON parse edilemedi, sıradaki modele geçiliyor")
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

def build_html(ai_data, stats, date_str, model_used, ai_usage=None, ai_failed=False):
    total           = stats["total"]
    agent_count     = stats["agent_count"]
    bot_count       = stats["bot_count"]
    complaint_count = stats["complaint_count"]
    brand_counts    = stats["brands"]

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
                    f'<tr style="border-bottom:1px solid #f0e8ff;">'
                    f'<td style="padding:9px 14px;white-space:nowrap;font-size:13px;font-weight:700;color:#2F1555;font-family:Montserrat,Arial,sans-serif;">{u.get("username","-")}</td>'
                    f'<td style="padding:9px 14px;"><span style="display:inline-block;padding:1px 7px;{bs}border-radius:3px;font-size:10px;font-weight:700;">{u.get("brand","-")}</span></td>'
                    f'<td style="padding:9px 14px;font-size:12px;color:#662D91;line-height:1.5;font-family:Montserrat,Arial,sans-serif;">{u.get("reason","-")}</td>'
                    f'</tr>'
                )
            critical_html = (
                f'<p style="margin:22px 0 10px;font-size:11px;font-weight:800;color:#2F1555;text-transform:uppercase;letter-spacing:.12em;font-family:Montserrat,Arial,sans-serif;">🚨 Kritik Kullanıcılar</p>'
                f'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e9d5ff;border-top:3px solid #FFE600;border-radius:0 0 8px 8px;">{rows}</table>'
            )

    # ── AI analizi başarısız oldu mu, yoksa gerçekten şikayet yok mu ayrımı ──
    if ai_failed:
        no_data_html = (
            '<table width="100%" cellpadding="0" cellspacing="0" '
            'style="background:#FEF2F2;border:1px solid #FCA5A5;border-radius:8px;">'
            '<tr><td style="padding:14px 16px;">'
            '<p style="margin:0;font-size:12px;font-weight:800;color:#B91C1C;'
            'font-family:Montserrat,Arial,sans-serif;">⚠️ AI ANALİZİ BAŞARISIZ</p>'
            f'<p style="margin:6px 0 0;font-size:11px;color:#991B1B;line-height:1.5;'
            f'font-family:Montserrat,Arial,sans-serif;">Gemini, Claude ve Groq sırayla denendi, '
            f'hiçbiri geçerli/parse edilebilir bir sonuç üretemedi. {complaint_count} şikayet tespit '
            f'edildi ancak kategorize edilemedi — ham veriler Excel ekinde, GitHub Actions loglarını kontrol et.</p>'
            '</td></tr></table>'
        )
    else:
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
def build_excel(complaint_chats, date_str, ai_data=None):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return None

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"Şikayetler {date_str}"

    HDR_BG  = "2F1555"
    HDR_FG  = "FFE600"
    ROW_ALT = "F3F0FF"
    RED_BG  = "F28B82"
    YLW_BG  = "FBBC04"

    headers    = ["#", "Kullanıcı Adı", "Brand", "Agent", "Tag", "Rating", "Yorum", "Tarih", "Chat Linki"]
    col_widths = [5, 22, 8, 22, 30, 8, 60, 20, 55]

    for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font      = Font(bold=True, color=HDR_FG, name="Calibri", size=10)
        cell.fill      = PatternFill("solid", fgColor=HDR_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=False)
        ws.column_dimensions[get_column_letter(ci)].width = w

    ws.row_dimensions[1].height = 22
    ws.freeze_panes = "A2"

    thin   = Side(style="thin", color="E0D8F0")
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

        try:
            import pytz
            dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
            ts_fmt = dt.astimezone(pytz.timezone("Europe/Sofia")).strftime("%d/%m/%Y %H:%M")
        except Exception:
            ts_fmt = ts

        values = [idx, username, brand, agent, tag, rating, comment, ts_fmt, link]

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
            if ci == 9 and val:
                cell.hyperlink = val
                cell.value     = "Chati Aç"
                cell.font      = Font(name="Calibri", size=9, color="1155CC", underline="single")

        ws.row_dimensions[row].height = 18

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

    # ── SEKME 2: AI Kategori Grupları ──
    if ai_data and ai_data.get("categories"):
        cats = sorted(ai_data["categories"], key=lambda x: x.get("count", 0), reverse=True)
        ws2  = wb.create_sheet(title="Kategori Grupları")

        ws2_cols   = ["Kategori", "Brand", "Kullanıcı Adı", "Tag", "Rating", "Yorum", "Tarih", "Chat Linki"]
        ws2_widths = [35, 8, 22, 30, 8, 60, 20, 55]
        for ci, (h, w) in enumerate(zip(ws2_cols, ws2_widths), 1):
            cell = ws2.cell(row=1, column=ci, value=h)
            cell.font      = Font(bold=True, color=HDR_FG, name="Calibri", size=10)
            cell.fill      = PatternFill("solid", fgColor=HDR_BG)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws2.column_dimensions[get_column_letter(ci)].width = w
        ws2.row_dimensions[1].height = 22
        ws2.freeze_panes = "A2"
        ws2.auto_filter.ref = f"A1:{get_column_letter(len(ws2_cols))}1"

        import re as _re

        def cat_keywords(cat_name):
            name = cat_name.lower()
            kws  = []
            if any(w in name for w in ["yatırım", "deposit", "para yatır", "yansımadı", "geçmedi", "eksikliği"]):
                kws += ["deposit_missing", "deposit_issue", "deposit_query"]
            if any(w in name for w in ["cashback", "casino cashback"]):
                kws += ["casino_cashback_query"]
            if any(w in name for w in ["spor cashback", "sport cashback"]):
                kws += ["sport_cashback_query"]
            if any(w in name for w in ["bahis kural", "betting", "kural"]):
                kws += ["betting_rules_query"]
            if any(w in name for w in ["oyun adil", "dolandırıcı", "game_fairness", "adil"]):
                kws += ["game_fairness"]
            if any(w in name for w in ["hesap kapat", "hesap sil", "ac_closure", "kapatma"]):
                kws += ["ac_closure_request"]
            if any(w in name for w in ["çekim", "para çek", "ödeme", "withdrawal"]):
                kws += ["withdrawal", "wd"]
            if any(w in name for w in ["teknik", "hata", "sistem", "error"]):
                kws += ["deposit_issue", "deposit_missing"]
            if any(w in name for w in ["bonus", "deneme", "goodwill", "iyi niyet"]):
                kws += ["goodwill_query", "bonus", "deneme"]
            return kws if kws else []

        assigned    = set()
        current_row = 2
        CAT_COLORS  = [
            "1a237e", "283593", "303f9f", "3949ab", "3f51b5",
            "5c6bc0", "7986cb", "512da8", "673ab7", "7b1fa2",
        ]

        for cat_idx, cat in enumerate(cats):
            cat_name  = cat.get("name", "")
            kws       = cat_keywords(cat_name)
            matched   = []

            for c in complaint_chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                if cid in assigned:
                    continue
                tag     = get_tag(c).lower()
                comment = get_rating_comment(c).lower()
                match   = False
                if kws:
                    for kw in kws:
                        if kw and (kw in tag or kw in comment):
                            match = True
                            break
                if match:
                    matched.append(c)

            if not matched:
                continue

            cat_color  = CAT_COLORS[cat_idx % len(CAT_COLORS)]
            merge_range = f"A{current_row}:{get_column_letter(len(ws2_cols))}{current_row}"
            ws2.merge_cells(merge_range)
            title_cell = ws2.cell(row=current_row, column=1,
                                  value=f"  {cat_name}  ({len(matched)} şikayet)")
            title_cell.font      = Font(bold=True, color="FFFFFF", name="Calibri", size=10)
            title_cell.fill      = PatternFill("solid", fgColor=cat_color)
            title_cell.alignment = Alignment(vertical="center")
            ws2.row_dimensions[current_row].height = 20
            current_row += 1

            for c in matched:
                cid     = str(c.get("id") or c.get("chatId") or "")
                brand   = get_brand(c)
                visitor = c.get("preChatName") or c.get("name") or c.get("visitorName") or ""
                tag_raw = get_tag(c) or "—"
                rating  = get_rating(c) or ""
                comment = get_rating_comment(c)
                ts      = c.get("startTime") or c.get("start_time") or ""
                link    = f"{portal}?chatId={cid}" if cid else ""

                try:
                    import pytz as _pytz2
                    dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
                    ts_fmt = dt.astimezone(_pytz2.timezone("Europe/Sofia")).strftime("%d/%m/%Y %H:%M")
                except Exception:
                    ts_fmt = ts

                bg = ROW_ALT if (current_row % 2 == 0) else "FFFFFF"
                if rating == 1:
                    bg = RED_BG
                elif rating == 2:
                    bg = YLW_BG

                values = [cat_name, brand, visitor, tag_raw, rating, comment, ts_fmt, link]
                for ci, val in enumerate(values, 1):
                    cell = ws2.cell(row=current_row, column=ci, value=val)
                    cell.fill      = PatternFill("solid", fgColor=bg)
                    cell.font      = Font(name="Calibri", size=9)
                    cell.alignment = Alignment(vertical="center", wrap_text=(ci == 6))
                    cell.border    = border
                    if ci == 8 and val:
                        cell.hyperlink = val
                        cell.value     = "Chati Aç"
                        cell.font      = Font(name="Calibri", size=9, color="1155CC", underline="single")

                ws2.row_dimensions[current_row].height = 18
                assigned.add(cid)
                current_row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()

# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, date_str, stats, complaint_chats=None, ai_data=None):
    subject = (f"CS Şikayet Analizi | {date_str} | "
               f"{stats['agent_count']} Agent | {stats['complaint_count']} Şikayet")
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    if complaint_chats:
        xlsx_bytes = build_excel(complaint_chats, date_str, ai_data=ai_data)
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
    ai_failed = False
    if chat_texts:
        print(f"\n[AI] {len(chat_texts)} şikayet analiz ediliyor...")
        ai_data, model_used, ai_usage = analyze_with_ai(chat_texts, date_str)
        if ai_data:
            print(f"[AI] {len(ai_data.get('categories', []))} kategori ({model_used})")
        else:
            print("[AI] Analiz başarısız")
            ai_failed = True

    html = build_html(ai_data, stats, date_str, model_used, ai_usage, ai_failed=ai_failed)
    send_email(html, date_str, stats, complaint_chats=complaint_chats, ai_data=ai_data)
    print(f"\n[DONE] Tamamlandı — {date_str}")

if __name__ == "__main__":
    main()
