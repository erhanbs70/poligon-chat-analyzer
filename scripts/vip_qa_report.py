#!/usr/bin/env python3
# ============================================================
# POLIGON VIP QA & AI REPORT — GitHub Actions Port (v1)
# GAS "Poligon VIP QA Report V69" scriptinin Python/Actions karşılığı.
#
# Neden taşındı:
#   Apps Script'in 6 dakikalık sert zaman aşımı limiti try/catch ile
#   yakalanamıyordu (platform script'i dışarıdan öldürüyordu), bu yüzden
#   AI_LOGS'a hiçbir kayıt düşmeden çalışma sessizce kayboluyordu.
#   GitHub Actions'ta bu sorun yapısal olarak yok — job başına 45 dk
#   bütçe var ve normal çalışma süresi (~1-2 dk) bunun çok altında.
#
# Repo'nun mevcut pattern'lerini takip eder:
#   - Comm100 auth / Sofia DST tarih yardımcıları: commqa.py ile aynı
#   - Google Sheets log: commqa.py'nin ensure_sheet/batchUpdate deseni
#   - Gemini → Claude → Groq fallback: chat_analyzer.py ile aynı yapı
#   - Hata durumunda hem sheet log hem uyarı maili: commqa.py deseni
# ============================================================
import os
import sys
import json
import base64
import time
import traceback
import requests
import smtplib
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID   = os.environ["MR_SITE_ID"]
API_KEY   = os.environ["MR_API_KEY"]
CFG_EMAIL = os.environ["MR_EMAIL"]

# commqa.py'deki CS_DEPT_ID gibi sabit — hassas veri değil, secret'a gerek yok
VIP_DEPT_ID = "433cb80f-4e0d-4e6d-96d3-bff700296a6f"

GEMINI_KEYS   = json.loads(os.environ["GEMINI_KEYS"])
CLAUDE_KEY    = os.environ["CLAUDE_KEY"]
GROQ_KEYS     = json.loads(os.environ["GROQ_KEYS"])
GMAIL_USER    = os.environ["GMAIL_USER"]
GMAIL_PASS    = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS = [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

# YENİ secret — bkz. teslimat notu. Boş bir Google Sheet ID'si yeterli,
# "AI_LOGS" tab'ı otomatik oluşturulur.
SPREADSHEET_ID = os.environ["VIP_SPREADSHEET_ID"]
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]

PORTAL_BASE   = f"https://dash15.lively-chat.com/ui/{SITE_ID}/livechat/history/chats/transcriptdetail"
AI_LOGS_SHEET = "AI_LOGS"
MAX_ANALYSES  = 20     # bir çalıştırmada en fazla kaç kritik chat AI ile analiz edilsin
SLEEP_BETWEEN_ANALYSES = 2.0  # Comm100 + AI API'lerini yormamak için

TR_MONTHS = ["Ocak","Subat","Mart","Nisan","Mayis","Haziran",
             "Temmuz","Agustos","Eylul","Ekim","Kasim","Aralik"]

gemini_key_index = 0

# ── DATE HELPERS (commqa.py ile aynı) ────────────────────────
def sofia_yesterday():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return (datetime.now(tz) - timedelta(days=1)).strftime("%Y-%m-%d")
def sofia_now_hhmm():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return datetime.now(tz).strftime("%H:%M")
def sofia_to_utc_range(date_str):
    import pytz
    tz    = pytz.timezone("Europe/Sofia")
    start = tz.localize(datetime.strptime(date_str + " 00:00:00", "%Y-%m-%d %H:%M:%S"))
    end   = tz.localize(datetime.strptime(date_str + " 23:59:59", "%Y-%m-%d %H:%M:%S"))
    utc   = pytz.utc
    return (
        start.astimezone(utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        end.astimezone(utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
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
def date_tr(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{d.day} {TR_MONTHS[d.month-1]} {d.year}"

# ── AUTH ─────────────────────────────────────────────────────
def get_auth():
    return "Basic " + base64.b64encode(f"{CFG_EMAIL}:{API_KEY}".encode()).decode()

# ── COMM100 FETCH — VIP chatleri ──────────────────────────────
def fetch_vip_chats(date_str):
    """
    Orijinal GAS scripti gibi departmentId'yi query param olarak GÖNDERMİYORUZ —
    tüm chatleri çekip client-side'da departmentId == VIP_DEPT_ID filtreliyoruz.
    (commqa.py server-side &departmentId= kullanıyor; bu script proven-working
    orijinal davranışı koruyor, riski minimize etmek için.)
    """
    start_time, end_time = sofia_to_utc_range(date_str)
    auth   = get_auth()
    result = []
    seen   = set()
    page   = 1
    print(f"[FETCH] VIP chatleri çekiliyor: {date_str}")
    while page <= 20:
        url = (
            f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
            f"?siteId={SITE_ID}&pageIndex={page}&pageSize=500"
            f"&include=chatAgent&include=chatWrapup&include=postChatSurvey&include=visitor"
        )
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
            if not cid or cid in seen:
                continue
            seen.add(cid)
            if c.get("departmentId") == VIP_DEPT_ID and in_sofia_range(c.get("startTime"), date_str):
                result.append(c)
        print(f"[FETCH] Sayfa {page}: {len(chats)} chat tarandı ({len(result)} VIP eşleşti)")
        if len(chats) < 500:
            break
        page += 1
        time.sleep(0.2)
    print(f"[FETCH] Toplam {len(result)} VIP chat")
    return result

# ── FIELD HELPERS ─────────────────────────────────────────────
def get_agent(c):
    agents = c.get("chatAgents") or []
    if agents:
        names = [
            (a.get("agent") or {}).get("displayName") or a.get("displayName") or a.get("name") or ""
            for a in agents if not a.get("botId")
        ]
        names = [n for n in names if n]
        if names:
            return ", ".join(names)
    return c.get("agentName") or "VIP Agent"
def get_visitor(c):
    return (c.get("preChatName") or c.get("name") or c.get("visitorName")
            or (c.get("visitor") or {}).get("name") or "VIP Üye")
def get_rating(c):
    s = c.get("postChatSurvey") or {}
    g = s.get("ratingGrade")
    return int(g) if g is not None else 0
def get_raw_tags(c):
    return (c.get("chatWrapup") or {}).get("categoriesName") or ""

# ── TRANSCRIPT ────────────────────────────────────────────────
def get_chat_transcript(chat_id):
    if not chat_id:
        return ""
    url = f"https://dash15.lively-chat.com/api/LiveChat/chats/{chat_id}?include=messages&siteId={SITE_ID}"
    try:
        r = requests.get(url, headers={"Authorization": get_auth()}, timeout=30)
        if r.status_code == 200:
            msgs = r.json().get("messages", [])
            lines = []
            for m in msgs:
                sender_type = m.get("senderType")
                sender = m.get("senderName") or (
                    "Agent" if sender_type == "agent" else
                    "Visitor" if sender_type == "visitor" else "System"
                )
                text = m.get("message") or ""
                if sender_type == "system" and "dosya" not in text.lower():
                    continue
                if text or sender_type != "system":
                    lines.append(f"{sender}: {text}")
            if lines:
                return "\n".join(lines)
    except Exception as e:
        print(f"[TRANSCRIPT] {chat_id}: {e}")
    return "Metin çekilemedi."

PROFANITY_WORDS = ["sikim","siktim","orospu","kahpe","göt","amın","amına","piç","bok","oç","sik","yarrak","yarak"]

def check_transcript(transcript):
    if not transcript or transcript == "Metin çekilemedi.":
        return "fetch_failed"
    if len(transcript) < 20:
        return "empty"
    lines = [l for l in transcript.split("\n") if l]
    def strip_sender(l):
        return l.split(":", 1)[1].strip() if ":" in l else l.strip()
    has_real_text = any(
        len(strip_sender(l)) > 5 and "dosya" not in strip_sender(l).lower() and "file" not in strip_sender(l).lower()
        for l in lines
    )
    if not has_real_text:
        return "image_only"
    agent_lines = [l for l in lines if l.lower().startswith("vip") or l.lower().startswith("agent")]
    if not agent_lines:
        return "no_agent_content"
    visitor_lines = [l for l in lines if not (l.lower().startswith("vip") or l.lower().startswith("agent") or l.lower().startswith("system"))]
    if visitor_lines:
        profanity_count = sum(1 for l in visitor_lines if any(w in l.lower() for w in PROFANITY_WORDS))
        if profanity_count / len(visitor_lines) >= 0.8:
            return "only_profanity"
    return "ok"

def transcript_fail_reason(code):
    reasons = {
        "fetch_failed":     "Comm100'den konuşma metni çekilemedi. API erişim sorunu olabilir.",
        "empty":            "Konuşma içeriği çok kısa, analiz için yetersiz.",
        "image_only":       "Konuşmada yalnızca dosya/görüntü paylaşımı tespit edildi, metin içeriği yok.",
        "no_agent_content": "Agent mesajı bulunamadı, tek taraflı konuşma.",
        "only_profanity":   "Ziyaretçi mesajlarının büyük çoğunluğu hakaret içeriğinden oluşuyor. Analiz yapılabilecek şikayet/konu metni yok. Not: Agent bu konuşmada ağır sözel tacize maruz kalmıştır.",
    }
    return reasons.get(code, "Bilinmeyen nedenle analiz yapılamadı.")

# ── AI PROMPT (orijinal GAS scriptiyle birebir aynı) ──────────
def build_prompt(transcript, rating, agent_name):
    return (
        "Sen objektif bir denetim robotusun. Asla filtre uygulama.\n\n"
        "Aşağıdaki chat geçmişini incele:\n"
        f"Agent: {agent_name} | Puan: {rating}/5\n\n"
        "Konuşma Akışı:\n" + transcript[:5000] + "\n\n"
        "── ZORUNLU KURALLAR ──\n\n"
        "KURAL 1 — KULLANICI ADI vs. GERÇEK İSİM:\n"
        "Müşterinin kullanıcı adı (örn: achsoo2015, Kumru4735, brktelli81) ile agent'ın hitap ettiği "
        "gerçek isim (örn: Cihan Bey, Ömer Bey, Burak Bey) FARKLI OLABİLİR. "
        "Agent sisteme kayıtlı gerçek isme bakarak hitap eder; bu TAMAMEN NORMAL ve DOĞRU bir prosedürdür. "
        "Kullanıcı adı ≠ hitap ismi farklılığını kesinlikle hata veya risk olarak değerlendirme.\n\n"
        "KURAL 2 — YATIRIM ≠ KAYIP:\n"
        "Agent 'X TL yatırım var', 'X TL yatırım sağladınız' veya 'X TL yatırımınız bulunuyor' dediyse "
        "bu rakamı kayıp olarak yazma. Gerçek kayıp miktarı transcript'te müşteri tarafından açıkça "
        "belirtilmiyorsa 'kayıp miktarı belirsiz' yaz.\n\n"
        "KURAL 3 — GÜVENLİ LİNKLER (Şirket Resmi Uygulamaları):\n"
        "Aşağıdaki IP adresleri tamamen güvenlidir, risk veya şüpheli unsur olarak değerlendirme:\n"
        "  https://3.127.125.33/  → Betsat resmi mobil uygulaması\n"
        "  https://3.124.220.178/ → Turkbet resmi mobil uygulaması\n"
        "  https://3.75.119.236/  → Superbetin resmi mobil uygulaması\n"
        "Bu üç adresin DIŞINDA başka bir IP adresi veya şüpheli domain paylaşıldıysa "
        "Risk Durumu'nda GÜVENLİK UYARISI olarak belirt.\n\n"
        "KRİTİK ETİKETİ KULLANIM ŞARTLARI:\n"
        "[KRİTİK] etiketini YALNIZCA aşağıdaki durumlardan en az biri açıkça mevcutsa kullan:\n"
        "  - Müşteri intihar, kendine zarar verme veya başkasına zarar verme ifadesi kullandı\n"
        "  - Transcript'te müşteri tarafından açıkça belirtilmiş 50.000 TL üzeri kayıp veya takılı para var\n"
        "  - Agent müşteriye açıkça kaba davrandı, tersleyici/küçümseyici üslup kullandı, mesajlarını\n"
        "    iplemedi, yanlış veya yanıltıcı bilgi verdi ya da açık bir ihmal yaptı\n"
        "DÜŞÜK PUAN, müşteri şikayeti, bonus reddi veya genel memnuniyetsizlik TEK BAŞINA [KRİTİK] için YETERLİ DEĞİLDİR.\n"
        "Agent'ın VIP üyeyi ikna etmeye çalışması, hesabını açık tutmasını önermesi veya oynamaya devam\n"
        "etmesini teşvik etmesi tamamen şirket politikasına uygundur; bunu hata, risk veya KRİTİK olarak değerlendirme.\n"
        "Hesap kapatma taleplerinin mail üzerinden işleme alınması standarttır; agent'ın müşteriyi\n"
        "mail atmaya yönlendirmesi doğru prosedürdür, ihmal değildir.\n\n"
        "ÇIKTI FORMATI — ÖNEMLİ:\n"
        "Analizinde iç kural referanslarına (örn. KURAL 1, KURAL 4, madde a/b/c) KESİNLİKLE yer verme.\n"
        "Sadece tespitini ve gerekçeni kendi cümlelerinle yaz.\n"
        "[KRİTİK] koşullarından hiçbiri yoksa Risk Durumu'nu normal şekilde yaz; agent'ın olumlu davranışlarını belirt.\n"
        "'ACİL MÜDAHALE' ifadesini YALNIZCA [KRİTİK] etiketiyle birlikte kullan.\n"
        "[KRİTİK] yokken meta-açıklama cümleleri KESİNLİKLE yazma.\n\n"
        "── TALİMAT ──\n"
        "Aşağıdaki 3 sorunun cevabını HTML, Markdown veya Emoji KULLANMADAN alt alta yaz. "
        "Başlıkları tam olarak belirttiğim gibi kullan:\n\n"
        "Kök Neden: [Müşterinin asıl sorunu veya şikayetinin kısa özeti]\n"
        "Agent Üslubu: [Agent'ın iletişim kalitesi, empati düzeyi, prosedür uyumu]\n"
        "Risk Durumu: [Genel risk değerlendirmesi. Müşteri konuşmanın SONUNDA sakin mi ayrıldı, "
        "ikna oldu mu, yoksa hâlâ çözümsüz/öfkeli mi ayrıldı? Bunu mutlaka belirt. "
        "Yukarıdaki KRİTİK koşullarından biri geçerliyse cevabının EN BAŞINA tam olarak [KRİTİK] kelimesini ekle.]"
    )

# ── AI FALLBACK: Gemini → Claude → Groq ───────────────────────
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
                        "safetySettings": [
                            {"category": "HARM_CATEGORY_HARASSMENT",        "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_HATE_SPEECH",       "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
                            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
                        ],
                    },
                    timeout=60,
                )
                if r.status_code in (429, 503):
                    print(f"[GEMINI] {model} key#{ki}: HTTP {r.status_code}, sıradaki deneniyor")
                    break
                if r.status_code == 200:
                    res = r.json()
                    cands = res.get("candidates", [])
                    if cands and cands[0].get("content"):
                        text = cands[0]["content"]["parts"][0]["text"].strip()
                        gemini_key_index = (gemini_key_index + ki + 1) % len(GEMINI_KEYS)
                        usage = res.get("usageMetadata", {})
                        in_t, out_t = usage.get("promptTokenCount", 0), usage.get("candidatesTokenCount", 0)
                        cost = (in_t * 0.075 / 1_000_000) + (out_t * 0.30 / 1_000_000)
                        model_name = "Gemini 2.5 Flash" if "2.5" in model else "Gemini 2.0 Flash"
                        print(f"[AI] {model_name} — {in_t}+{out_t} token | ${cost:.6f}")
                        return {"success": True, "text": text, "model": model_name,
                                "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
                    else:
                        reason = (cands[0].get("finishReason") if cands else "UNKNOWN")
                        print(f"[GEMINI] {model} key#{ki}: içerik filtresi ({reason})")
                else:
                    print(f"[GEMINI] {model} key#{ki}: HTTP {r.status_code}: {r.text[:150]}")
            except Exception as e:
                print(f"[GEMINI] {model} key#{ki}: {e}")
    return {"success": False}

def try_claude(prompt):
    try:
        r = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={"x-api-key": CLAUDE_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json"},
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 1000,
                  "messages": [{"role": "user", "content": prompt}]},
            timeout=60,
        )
        if r.status_code == 200:
            res = r.json()
            in_t  = res.get("usage", {}).get("input_tokens", 0)
            out_t = res.get("usage", {}).get("output_tokens", 0)
            cost  = (in_t * 0.25 / 1_000_000) + (out_t * 1.25 / 1_000_000)
            print(f"[AI] Claude Haiku 4.5 — {in_t}+{out_t} token | ${cost:.6f}")
            return {"success": True, "text": res["content"][0]["text"].strip(), "model": "Claude Haiku 4.5",
                    "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
        print(f"[CLAUDE] HTTP {r.status_code}: {r.text[:150]}")
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
                      "messages": [{"role": "user", "content": prompt}], "temperature": 0.4},
                timeout=60,
            )
            if r.status_code == 200:
                res = r.json()
                in_t  = res.get("usage", {}).get("prompt_tokens", 0)
                out_t = res.get("usage", {}).get("completion_tokens", 0)
                cost  = (in_t * 0.59 / 1_000_000) + (out_t * 0.79 / 1_000_000)
                print(f"[AI] Groq Llama-4 Scout — {in_t}+{out_t} token | ${cost:.6f}")
                return {"success": True, "text": res["choices"][0]["message"]["content"].strip(),
                        "model": "Groq Llama-4 Scout", "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
            print(f"[GROQ] key#{i} HTTP {r.status_code}: {r.text[:150]}")
        except Exception as e:
            print(f"[GROQ] key#{i}: {e}")
    return {"success": False}

def analyze_with_ai(transcript, rating, tags, agent_name):
    ret = {"text": "⚠️ Analiz yapılamadı.", "inT": 0, "outT": 0, "modelUsed": "", "cost": 0}
    check = check_transcript(transcript)
    if check != "ok":
        ret["text"] = "⚠️ " + transcript_fail_reason(check)
        return ret
    prompt = build_prompt(transcript, rating, agent_name)
    for fn in (try_gemini, try_claude, try_groq):
        result = fn(prompt)
        if result["success"]:
            ret.update({
                "text": result["text"], "inT": result["in_tokens"], "outT": result["out_tokens"],
                "modelUsed": result["model"], "cost": result["cost"],
            })
            return ret
    ret["text"] = "⚠️ Tüm AI motorları yanıt vermedi (Gemini/Claude/Groq)."
    return ret

# ── PROCESS ────────────────────────────────────────────────────
def process_vip_chats(chats):
    m = {"totalScanned": len(chats), "aiAnalyses": [], "inT": 0, "outT": 0,
         "totalVirtualCost": 0.0, "totalCritical": 0}
    for c in chats:
        rating   = get_rating(c)
        raw_tags = get_raw_tags(c)
        is_critical = (0 < rating <= 2) or ("game_fairness" in raw_tags) or ("ac_closure_request" in raw_tags)
        agent = get_agent(c)
        if is_critical and "VIP" in agent.upper():
            m["totalCritical"] += 1
            if len(m["aiAnalyses"]) < MAX_ANALYSES:
                chat_id = c.get("id") or c.get("chatId")
                transcript = get_chat_transcript(chat_id)
                ai_res = analyze_with_ai(transcript, rating, raw_tags, agent)
                m["inT"] += ai_res["inT"]
                m["outT"] += ai_res["outT"]
                m["totalVirtualCost"] += ai_res["cost"]
                clean_tag = "Belirtilmemiş"
                if raw_tags:
                    import re
                    first_tag = raw_tags.split(",")[0].split("\n")[0]
                    match = re.search(r"\(([^)]+)\)", first_tag)
                    clean_tag = match.group(1) if match else first_tag.replace("(", "").replace(")", "")
                m["aiAnalyses"].append({
                    "agent": agent, "visitor": get_visitor(c), "chatId": chat_id,
                    "rating": rating, "tags": clean_tag, "summary": ai_res["text"],
                    "model": ai_res["modelUsed"],
                })
                print(f"[AI] {get_visitor(c)} / {agent} → {ai_res['modelUsed'] or 'BAŞARISIZ'}")
                time.sleep(SLEEP_BETWEEN_ANALYSES)
    return m

# ── HTML BUILD (orijinal GAS scriptiyle aynı görünüm) ─────────
def build_html(m, date_label):
    import re
    now = sofia_now_hhmm()
    if m["aiAnalyses"]:
        rows_html = ""
        for a in m["aiAnalyses"]:
            clr = "#dc2626" if (0 < a["rating"] <= 2) else "#ea580c"
            is_critical_risk = "[KRİTİK]" in a["summary"]
            is_warning = a["summary"].startswith("⚠️")
            raw_summary = a["summary"].replace("[KRİTİK]", "").replace("**", "")
            raw_summary = re.sub(r"\[KRİTİK\]\s*", "", a["summary"]).replace("**", "")
            if is_warning:
                formatted = (
                    "<div style='background:#fffbeb; border-left:4px solid #f59e0b; padding:14px 18px; "
                    "border-radius:6px; color:#92400e; font-size:14px; line-height:1.6;'>"
                    f"<b>&#9888; Otomatik Analiz Yapılamadı</b><br><br>{raw_summary.replace('⚠️', '').strip()}</div>"
                )
            else:
                formatted = re.sub(
                    r"Kök Neden\s?:",
                    "<div style='margin-bottom:12px; padding-bottom:12px; border-bottom:1px dashed #e2e8f0;'>"
                    "<b style='color:#b91c1c; font-size:15px;'>&#127919; Kök Neden:</b><br>",
                    raw_summary, flags=re.IGNORECASE,
                )
                formatted = re.sub(
                    r"Agent Üslubu\s?:",
                    "</div><div style='margin-bottom:12px; padding-bottom:12px; border-bottom:1px dashed #e2e8f0;'>"
                    "<b style='color:#0369a1; font-size:15px;'>&#128172; Agent Üslubu:</b><br>",
                    formatted, flags=re.IGNORECASE,
                )
                if is_critical_risk:
                    formatted = re.sub(
                        r"Risk Durumu\s?:",
                        "</div><div style='background:#fef2f2; border-left:4px solid #ef4444; padding:12px; "
                        "border-radius:6px; margin-top:10px;'>"
                        "<b style='color:#dc2626; font-size:15px;'>&#128680; ACİL MÜDAHALE (KRİTİK):</b><br>"
                        "<span style='color:#991b1b; font-weight:bold;'>",
                        formatted, flags=re.IGNORECASE,
                    )
                    if "&#127919;" in formatted:
                        formatted += "</span></div>"
                elif "Risk Durumu" in formatted:
                    formatted = re.sub(
                        r"Risk Durumu\s?:",
                        "</div><div><b style='color:#b45309; font-size:15px;'>&#128680; Risk Durumu:</b><br>",
                        formatted, flags=re.IGNORECASE,
                    )
                    if "&#127919;" in formatted:
                        formatted += "</div>"
                formatted = re.sub(r"<br>\s*[Rr]isk\s*[Dd]urumu\s*:?\s*", "<br>", formatted)
            rating_badge = f"{a['rating']} Yıldız Puan" if a["rating"] else f"Tag: {a['tags']}"
            model_line = f"<div style='font-size:11px; color:#6b7280; margin-top:12px; font-weight:bold;'>&#129416; AI Motoru: {a['model']}</div>" if a["model"] else ""
            rows_html += f"""<tr>
      <td style='padding:25px; border-bottom:1px solid #fee2e2; background:#fff8f8; vertical-align:top; width:260px; border-right:1px solid #fee2e2;'>
        <div style='font-weight:900; font-size:16px; color:#111827;'>Müşteri: {a['visitor']}</div>
        <div style='font-size:14px; color:#4b5563; margin-top:10px;'>Agent: <b style='color:#1d4ed8;'>{a['agent']}</b></div>
        <div style='font-size:13px; color:#6b7280; margin-top:10px;'>Chat ID: <a href="{PORTAL_BASE}?chatId={a['chatId']}" target="_blank" style="color:#2563eb;text-decoration:none;font-weight:bold;">{a['chatId']} &#128279;</a></div>
        <div style='margin-top:15px;'><span style='background:{clr}; color:#fff; padding:6px 12px; border-radius:6px; font-size:13px; font-weight:bold; letter-spacing:0.5px;'>{rating_badge}</span></div>
        {model_line}
      </td>
      <td style='padding:25px; border-bottom:1px solid #fee2e2; background:#fff; font-size:15px; line-height:1.7; color:#1f2937; word-break:break-word;'>
        {formatted}
      </td>
    </tr>"""
    else:
        rows_html = "<tr><td colspan='2' style='padding:30px;text-align:center;'>Riskli bir durum tespit edilmedi.</td></tr>"

    if m["totalCritical"] > len(m["aiAnalyses"]):
        critical_note = f"<b>{m['totalCritical']}</b> kritik chat tespit edildi, ilk <b>{len(m['aiAnalyses'])}</b> tanesi AI ile analiz edildi."
    else:
        critical_note = f"<b>{m['totalCritical']}</b> kritik chat tespit edildi ve tamamı analiz edildi."

    return f"""<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Arial,sans-serif;">
  <table width="100%" cellpadding="0" cellspacing="0"><tr><td align="center" style="padding:40px 15px;">
    <table width="950" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 15px 35px rgba(0,0,0,0.05);">
      <tr><td style="background:linear-gradient(135deg, #0f1e36 0%, #1e3a8a 100%); padding:40px 30px; text-align:center;">
        <div style="color:#fff; font-size:28px; font-weight:900; letter-spacing:1px;">VIP Quality Assurance & AI Alert</div>
        <div style="color:#bfdbfe; font-size:15px; margin-top:10px; font-weight:500;">{date_label} | Sofia Time: {now}</div>
      </td></tr>
      <tr><td style="padding:40px;">
        <div style="background:#eff6ff; border-left:4px solid #3b82f6; padding:18px 25px; border-radius:0 8px 8px 0; margin-bottom:35px; font-size:15px; color:#1e3a8a; line-height:1.6;">
          Sistem bugün toplam <b>{m['totalScanned']} adet</b> VIP chat taradı. {critical_note}
        </div>
        <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #fee2e2; border-radius:10px; overflow:hidden; border-collapse:collapse;">
          {rows_html}
        </table>
        <div style="text-align:center; font-size:13px; color:#9ca3af; margin-top:45px; padding-top:25px; border-top:1px dashed #e5e7eb;">
          <strong>POLIGON VIP QA SYSTEM (GitHub Actions)</strong><br>Developed by Erhan
        </div>
      </td></tr>
    </table>
  </td></tr></table></body></html>"""

# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, date_label, m):
    alert_count = len(m["aiAnalyses"])
    subject = f"VIP AI Alert | {date_label} | " + (f"{alert_count} Kritik Durum Tespit Edildi ⚠️" if alert_count else "Risk Yok ✅")
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

def send_error_email(date_str, error_msg, tb_str):
    subject = f"⚠️ VIP QA Report HATA | {date_str}"
    html = f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:20px;">
<h2 style="color:#c0392b;">VIP QA & AI Report — Hata Raporu</h2>
<p><strong>Tarih:</strong> {date_str}</p>
<p><strong>Hata:</strong> {error_msg}</p>
<pre style="background:#f8f9fa;padding:16px;border-radius:6px;font-size:12px;white-space:pre-wrap;word-break:break-all;">{tb_str}</pre>
<p style="color:#666;font-size:11px;">CS Operations &amp; Technology / Poligon Entertainment N.V.</p>
</body></html>"""
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])
    msg.attach(MIMEText(html, "html", "utf-8"))
    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(GMAIL_USER, GMAIL_PASS)
            server.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
        print("[EMAIL] Hata maili gönderildi")
    except Exception as e:
        print(f"[EMAIL] Hata maili gönderilemedi: {e}")

# ── GOOGLE SHEETS LOG (commqa.py deseni) ──────────────────────
def sheets_service():
    creds = service_account.Credentials.from_service_account_info(
        json.loads(GCP_CREDS_JSON),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)

def get_sheet_id(svc, sheet_name):
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    for s in meta.get("sheets", []):
        if s["properties"]["title"] == sheet_name:
            return s["properties"]["sheetId"]
    return None

def ensure_sheet(svc, sheet_name):
    sid = get_sheet_id(svc, sheet_name)
    if sid is not None:
        return sid
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]

def rgb(hex_str):
    h = hex_str.lstrip("#")
    return {"red": int(h[0:2], 16) / 255, "green": int(h[2:4], 16) / 255, "blue": int(h[4:6], 16) / 255}

def log_to_sheet(date_str, m, duration_sec, status):
    try:
        svc = sheets_service()
        sid = ensure_sheet(svc, AI_LOGS_SHEET)
        header_row_result = svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=f"'{AI_LOGS_SHEET}'!A1:I1"
        ).execute().get("values", [])
        header_row = header_row_result[0] if header_row_result else []
        if not header_row:
            # Sheet tamamen boş — tam 9 kolonlu header'ı sıfırdan yaz
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"'{AI_LOGS_SHEET}'!A1",
                valueInputOption="RAW",
                body={"values": [["Tarih","Taranan Chat","Kritik Toplam","Analiz Edilen",
                                   "Süre (Saniye)","Okunan Token","Yazılan Token",
                                   "Sanal Maliyet (Tasarruf $)","Durum"]]},
            ).execute()
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": [{
                    "repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 0, "endColumnIndex": 9},
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": rgb("#0f1e36"),
                            "textFormat": {"bold": True, "foregroundColor": rgb("#ffffff")},
                        }},
                        "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
                    }
                }]},
            ).execute()
        elif len(header_row) < 9 or "Durum" not in header_row:
            # [FIX] Eski (GAS'tan kalma) 8 kolonlu header — mevcut geçmiş veriye
            # dokunmadan sadece I1'e "Durum" başlığını ekle
            svc.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID, range=f"'{AI_LOGS_SHEET}'!I1",
                valueInputOption="RAW", body={"values": [["Durum"]]},
            ).execute()
            svc.spreadsheets().batchUpdate(
                spreadsheetId=SPREADSHEET_ID,
                body={"requests": [{
                    "repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 8, "endColumnIndex": 9},
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": rgb("#0f1e36"),
                            "textFormat": {"bold": True, "foregroundColor": rgb("#ffffff")},
                        }},
                        "fields": "userEnteredFormat.backgroundColor,userEnteredFormat.textFormat",
                    }
                }]},
            ).execute()
            print("[LOG] Eski header'a 'Durum' kolonu eklendi (I1)")
        cost_str = f"${float(m.get('totalVirtualCost') or 0):.5f}"
        svc.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID, range=f"'{AI_LOGS_SHEET}'!A1",
            valueInputOption="RAW", insertDataOption="INSERT_ROWS",
            body={"values": [[date_str, m.get("totalScanned"), m.get("totalCritical"),
                               len(m.get("aiAnalyses", [])), duration_sec,
                               m.get("inT"), m.get("outT"), cost_str, status]]},
        ).execute()
        print(f"[LOG] AI_LOGS'a yazıldı — Durum: {status}")
    except Exception as e:
        print(f"[LOG] Log hatası: {e}")

# ── MAIN ───────────────────────────────────────────────────────
def main():
    t0 = time.time()
    date_arg = sys.argv[1] if len(sys.argv) > 1 and sys.argv[1].strip() else None
    date_str = date_arg or sofia_yesterday()
    date_label = date_tr(date_str)
    print(f"\n{'='*55}\nPOLIGON VIP QA & AI REPORT — {date_str}\n{'='*55}\n")
    try:
        chats = fetch_vip_chats(date_str)
        metrics = process_vip_chats(chats)
        html = build_html(metrics, date_label)
        send_email(html, date_label, metrics)
        duration = round(time.time() - t0, 1)
        log_to_sheet(date_str, metrics, duration, "Tamamlandı")
        print(f"\n[DONE] {date_str} — {metrics['totalCritical']} kritik / {len(metrics['aiAnalyses'])} analiz edildi / {duration}s")
    except Exception as e:
        duration = round(time.time() - t0, 1)
        tb_str = traceback.format_exc()
        print(f"[ERROR] {e}\n{tb_str}")
        try:
            log_to_sheet(date_str, {"totalScanned": "-", "totalCritical": "-", "aiAnalyses": [],
                                     "inT": "-", "outT": "-", "totalVirtualCost": 0},
                         duration, f"HATA: {e}")
        except Exception as log_err:
            print(f"[LOG] Log da yazılamadı: {log_err}")
        send_error_email(date_str, str(e), tb_str)
        sys.exit(1)

if __name__ == "__main__":
    main()
