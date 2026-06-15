#!/usr/bin/env python3
# ============================================================
# POLIGON CS FEEDBACK AI REPORT — GitHub Actions
# Google Sheets'ten feedback verilerini çekip AI ile analiz eder
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
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── CONFIG ───────────────────────────────────────────────────
SPREADSHEET_ID = os.environ["SPREADSHEET_ID"]
SHEET_PREFIX   = "CS ALL"
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]
GEMINI_KEYS    = json.loads(os.environ["GEMINI_KEYS"])
CLAUDE_KEY     = os.environ["CLAUDE_KEY"]
GROQ_KEYS      = json.loads(os.environ["GROQ_KEYS"])
GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASS     = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS  = [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

TZ_SOFIA = "Europe/Sofia"
BRANDS   = ["SB", "BS", "TB"]
gemini_key_index = 0

TR_MONTHS = ["Ocak","Subat","Mart","Nisan","Mayis","Haziran",
             "Temmuz","Agustos","Eylul","Ekim","Kasim","Aralik"]

# ── DATE HELPERS ─────────────────────────────────────────────
def get_sofia_today():
    import pytz
    sofia = pytz.timezone("Europe/Sofia")
    return datetime.now(sofia).strftime("%Y-%m-%d")

def get_sofia_yesterday():
    import pytz
    sofia = pytz.timezone("Europe/Sofia")
    return (datetime.now(sofia) - timedelta(days=1)).strftime("%Y-%m-%d")

def date_tr(date_str):
    d = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{d.day} {TR_MONTHS[d.month-1]} {d.year}"

def week_label_tr(start_str, end_str):
    return f"{date_tr(start_str)} \u2013 {date_tr(end_str)}"

# ── GOOGLE SHEETS ─────────────────────────────────────────────
def get_sheets_service():
    creds_data = json.loads(GCP_CREDS_JSON)
    creds = service_account.Credentials.from_service_account_info(
        creds_data,
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    return build("sheets", "v4", credentials=creds)

def load_rows(date_from, date_to):
    service = get_sheets_service()
    year = int(date_from[:4])

    # Önce o yılın sheet'ini dene
    sheet_name = f"{SHEET_PREFIX} {year}"
    try:
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{sheet_name}'!A:H"
        ).execute()
        values = result.get("values", [])
    except Exception:
        # Bulamazsa ilk uygun sheet'i dene
        meta = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
        sheets = meta.get("sheets", [])
        cs_sheets = [s["properties"]["title"] for s in sheets
                     if SHEET_PREFIX in s["properties"]["title"]]
        if not cs_sheets:
            print("[SHEETS] CS ALL sheet bulunamadı")
            return []
        cs_sheets.sort(reverse=True)
        result = service.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{cs_sheets[0]}'!A:H"
        ).execute()
        values = result.get("values", [])

    rows = []
    for row in values:
        if len(row) < 6:
            continue
        brand_raw = str(row[0]).strip()
        date_val  = str(row[1]).strip() if len(row) > 1 else ""
        username  = str(row[2]).strip() if len(row) > 2 else ""
        employee  = str(row[3]).strip() if len(row) > 3 else ""
        category  = str(row[4]).strip() if len(row) > 4 else ""
        feedback  = str(row[5]).strip() if len(row) > 5 else ""
        assigned  = str(row[6]).strip() if len(row) > 6 else ""
        result_   = str(row[7]).strip() if len(row) > 7 else ""

        if not brand_raw or brand_raw.upper() in ["BRAND", "SINAN92X"]:
            continue
        if not category or category.upper() == "MAIN CATEGORY":
            continue
        if not date_val:
            continue

        brand = norm_brand(brand_raw)
        if not brand:
            continue

        # Tarih parse
        date_str = parse_date(date_val)
        if not date_str:
            continue
        if date_str < date_from or date_str > date_to:
            continue

        rows.append({
            "brand":    brand,
            "date":     date_str,
            "username": username,
            "employee": employee,
            "category": norm_cat(category),
            "feedback": feedback,
            "assigned": assigned,
            "result":   result_
        })

    print(f"[SHEETS] {len(rows)} kayıt yüklendi ({date_from} - {date_to})")
    return rows

def parse_date(val):
    """Tarih string'ini YYYY-MM-DD formatına çevir"""
    import re
    val = str(val).strip()

    # Excel serial number
    try:
        n = float(val)
        if 30000 < n < 60000:
            from datetime import date
            d = date(1899, 12, 30) + timedelta(days=int(n))
            return d.strftime("%Y-%m-%d")
    except:
        pass

    # YYYY-MM-DD
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", val)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"

    # DD.MM.YYYY veya DD/MM/YYYY
    m = re.match(r"^(\d{1,2})[./](\d{1,2})[./](\d{4})", val)
    if m:
        return f"{m.group(3)}-{m.group(2).zfill(2)}-{m.group(1).zfill(2)}"

    return None

def norm_brand(r):
    x = r.upper().strip()
    if x in ["SB", "SUPERBETIN"]: return "SB"
    if x in ["BS", "BETSAT"]:     return "BS"
    if x in ["TB", "TURKBET"]:    return "TB"
    return ""

def norm_cat(r):
    x = r.upper().strip()
    if x in ["WD", "WITHDRAWAL"]:    return "WD"
    if x == "DEPOSIT":               return "Deposit"
    if x in ["PROMOTION", "BONUS"]:  return "Promotion"
    if x == "TECHNICAL":             return "Technical"
    if x == "FAKE":                  return "Fake"
    if x == "GENERAL":               return "General"
    if x == "SPORT":                 return "Sport"
    if x == "CASINO":                return "Casino"
    if x == "SUPPORT":               return "Support"
    return r.strip() or "Diger"

# ── METRİKLER ─────────────────────────────────────────────────
def compute_metrics(rows):
    m = {
        "total": len(rows),
        "by_brand": {"SB": 0, "BS": 0, "TB": 0},
        "by_category": {},
        "repeat_users": [],
        "category_ranked": []
    }
    uc = {}
    for r in rows:
        if r["brand"] in m["by_brand"]:
            m["by_brand"][r["brand"]] += 1
        cat = r["category"]
        m["by_category"][cat] = m["by_category"].get(cat, 0) + 1
        k = f"{r['brand']}|{r['username'].lower()}"
        uc[k] = uc.get(k, 0) + 1

    for k, cnt in uc.items():
        if cnt >= 2:
            brand, username = k.split("|", 1)
            m["repeat_users"].append({"brand": brand, "username": username, "count": cnt})
    m["repeat_users"].sort(key=lambda x: -x["count"])

    m["category_ranked"] = sorted(
        [{"cat": c, "cnt": n} for c, n in m["by_category"].items()],
        key=lambda x: -x["cnt"]
    )
    return m

def daily_breakdown(rows, start_str, end_str):
    from datetime import date
    s = datetime.strptime(start_str, "%Y-%m-%d").date()
    e = datetime.strptime(end_str,   "%Y-%m-%d").date()
    days = []
    cur = s
    while cur <= e:
        ds  = cur.strftime("%Y-%m-%d")
        lbl = cur.strftime("%a %d.%m")
        days.append({"label": lbl, "count": sum(1 for r in rows if r["date"] == ds)})
        cur += timedelta(days=1)
    return days

# ── AI ────────────────────────────────────────────────────────
def build_prompt(rows, mode, lbl):
    weekly = (mode == "weekly")
    lines = []
    for i, r in enumerate(rows):
        fb = str(r["feedback"] or "-") if mode == "daily" else str(r["feedback"] or "-")[:150]
        lines.append(f"{i+1}. [{r['brand']}] [{r['category']}] {r['username'] or '?'} | {fb}")

    schema = json.dumps({
        "categories": [{
            "name":           "Spesifik konu — feedback metninden cikart, max 6 kelime",
            "count":          0,
            "brandBreakdown": [{"brand": "SB", "count": 0}],
            "users":          [{"username": "str", "brand": "SB"}],
            "shortNote":      "Gercek sorunu 1-2 cumle ile acikla — spesifik ol, somut bilgi ver"
        }],
        "critical":    [{"username": "str", "brand": "str", "reason": "Tutar veya aciliyet — 1 cumle, rakam varsa yaz"}],
        "actionItems": ["Kime ne yapilmali — departman + aksiyon — 1 cumle"],
        "summary":     "Genel durumu 2 cumle ile ozetle, onemli kullanici adi varsa yaz"
    }, ensure_ascii=False)

    return (
        f"CS feedback analistisisin. Asagidaki listeyi analiz et.\n\n"
        f"Donem: {lbl} | Toplam: {len(rows)} kayit\n\n"
        f"--- KAYITLAR ---\n" + "\n".join(lines) + "\n--- ---\n\n"
        "KONU ADI KURALLARI:\n"
        "YANLIS: \"Cekim Sorunlari\", \"Yatirim Talepleri\"\n"
        "DOGRU: \"Onayli Havale Cekimi Hesaba Gecmedi\", \"OTH Aktiflestirilmesi Talebi\"\n"
        "Ayni kategoride farkli sebepler varsa farkli grupla.\n\n"
        "BENZER KONULARI BIRLESTIR: Ayni sorunu farkli kelimelerle ifade eden kayitlar TEK kategori altinda toplanmali.\n"
        "Ornek: \"cekim gelmedi\", \"havale cekimi hesaba gecmedi\", \"onayli cekim yansimadi\" → hepsi = \"Onayli Cekim Hesaba Gecmedi\"\n"
        "Ornek: \"BBL kaldirilmasi talebi\", \"BBL durumu gozden gecirme\" → hepsi = \"BBL Kaldirilmasi Talebi\"\n"
        "Suphe durumunda birlestir, ayirma.\n\n"
        "SHORT NOTE KURALI: Her kategori icin shortNote ZORUNLU. Genel soz degil — somut sorun yaz.\n"
        "KOTU: \"Cekim sorunlari mevcut\"\n"
        "IYI: \"Onayli havale cekimleri 1-3 is gunu icinde hesaplara yansimamis, musteri maduriyet bildiriyor\"\n\n"
        "KRITIK KURAL: Yuksek tutar (5000 TL+), hesap kapatma tehdidi, acil cozum gerektiren durumlar kritik olarak isaretle.\n\n"
        "AKSIYON KURALI: Her aksiyon icin hangi departman ne yapmali yaz. Ornek: \"Finans departmani onayli cekim listesini kontrol etmeli.\"\n\n"
        "GOREV:\n"
        "1. Spesifik konu altinda grupla\n"
        "2. brandBreakdown: her brand kac kayit var, SADECE varsa ekle\n"
        "3. users: her uyenin username ve brand bilgisini ekle — hicbirini atla\n"
        "4. categories listesini buyukten kucuge sirala\n"
        "5. Kritik olanlari isaretle (yuksek tutar, acil, cozumsuz, tehdit)\n"
        "6. actionItems: en az 2, en fazla 5 aksiyon yaz\n"
        + ("7. Haftalik seyri ozetle — artan/azalan konulari belirt\n" if weekly else "")
        + "\nBBL=Bonus Black List | SB=Superbetin | BS=Betsat | TB=Turkbet\n"
        "OZET KURALI: Onemli durum varsa ilgili kullanici adini yaz. Genel ifadelerden kac.\n"
        f"Sadece JSON dondur:\n{schema}"
    )

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
                if r.status_code in [429, 503]:
                    break
                if r.status_code == 200:
                    res   = r.json()
                    cands = res.get("candidates", [])
                    if cands and cands[0].get("content"):
                        text = cands[0]["content"]["parts"][0]["text"].strip()
                        gemini_key_index = (gemini_key_index + ki + 1) % len(GEMINI_KEYS)
                        usage  = res.get("usageMetadata", {})
                        in_t   = usage.get("promptTokenCount", 0)
                        out_t  = usage.get("candidatesTokenCount", 0)
                        cost   = (in_t * 0.075 / 1_000_000) + (out_t * 0.30 / 1_000_000)
                        model_name = "Gemini 2.5 Flash" if "2.5" in model else "Gemini 2.0 Flash"
                        print(f"[AI] {model_name} | {in_t}+{out_t} token | ${cost:.6f}")
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
            json={"model": "claude-haiku-4-5-20251001", "max_tokens": 8000,
                  "messages": [{"role": "user", "content": prompt}]},
            timeout=60
        )
        if r.status_code == 200:
            res  = r.json()
            text = res["content"][0]["text"].strip()
            in_t  = res.get("usage", {}).get("input_tokens",  0)
            out_t = res.get("usage", {}).get("output_tokens", 0)
            cost  = (in_t * 0.25 / 1_000_000) + (out_t * 1.25 / 1_000_000)
            print(f"[AI] Claude Haiku | {in_t}+{out_t} token | ${cost:.6f}")
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
                json={"model": "meta-llama/llama-4-scout-17b-16e-instruct",
                      "messages": [{"role": "user", "content": prompt}],
                      "temperature": 0.2, "response_format": {"type": "json_object"}},
                timeout=60
            )
            if r.status_code == 200:
                res  = r.json()
                text = res["choices"][0]["message"]["content"].strip()
                in_t  = res.get("usage", {}).get("prompt_tokens",     0)
                out_t = res.get("usage", {}).get("completion_tokens", 0)
                cost  = (in_t * 0.11 / 1_000_000) + (out_t * 0.34 / 1_000_000)
                print(f"[AI] Groq | {in_t}+{out_t} token | ${cost:.6f}")
                return {"success": True, "text": text, "model": "Groq Llama-4 Scout",
                        "in_tokens": in_t, "out_tokens": out_t, "cost": cost}
        except Exception as e:
            print(f"[GROQ] key#{i}: {e}")
    return {"success": False}

def parse_json(text):
    try:
        return json.loads(text.replace("```json", "").replace("```", "").strip())
    except:
        return None

def analyze_with_ai(rows, mode, lbl):
    if not rows:
        return None, "—", {}
    prompt = build_prompt(rows, mode, lbl)
    for fn in [try_gemini, try_claude, try_groq]:
        result = fn(prompt)
        if result["success"]:
            data = parse_json(result["text"])
            if data:
                data["modelUsed"] = result["model"]
                usage = {
                    "model":      result["model"],
                    "in_tokens":  result.get("in_tokens",  0),
                    "out_tokens": result.get("out_tokens", 0),
                    "cost":       result.get("cost", 0)
                }
                return data, result["model"], usage
    return None, "Hata", {}

# ── HTML ──────────────────────────────────────────────────────
C = {
    "dark":    "#2F1555",
    "main":    "#662D91",
    "yellow":  "#FFE600",
    "lavanta": "#B28ABF",
    "subtle":  "#7c3aed",
    "sb":      "#1d4ed8",
    "tb":      "#E30613",
    "white":   "#ffffff",
    "bgBody":  "#ffffff",
    "divider": "#f0e8ff",
    "border":  "#e9d5ff"
}

def brand_bg_text(b):
    if b == "BS": return (C["yellow"], C["dark"])
    if b == "TB": return (C["tb"],     C["white"])
    return (C["sb"], C["white"])

def line_color(cnt):
    if cnt >= 8: return C["dark"]
    if cnt >= 4: return C["main"]
    if cnt >= 2: return C["lavanta"]
    return "#d8b4fe"

def brand_color(b):
    return {"SB": C["sb"], "BS": C["main"], "TB": C["tb"]}.get(b, C["dark"])

def topic_rows(categories):
    if not categories:
        return ""
    cats = sorted(categories, key=lambda x: x.get("count", 0), reverse=True)
    html = ""
    for idx, cat in enumerate(cats):
        cnt   = cat.get("count", 0)
        lclr  = line_color(cnt)
        is_last = idx == len(cats) - 1
        border  = "none" if is_last else f"1px solid {C['divider']}"

        bd_html = ""
        if cat.get("brandBreakdown"):
            items = ""
            for bd in cat["brandBreakdown"]:
                bg, fg = brand_bg_text(bd["brand"])
                items += (
                    f'<span style="display:inline-block;margin-right:8px;">'
                    f'<span style="display:inline-block;padding:1px 6px;background:{bg};color:{fg};border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{bd["brand"]}</span>'
                    f'<strong style="font-size:12px;color:{brand_color(bd["brand"])};margin-left:3px;font-family:Montserrat,Arial,sans-serif;">{bd["count"]}</strong>'
                    f'</span>'
                )
            bd_html = (
                f'<table cellpadding="0" cellspacing="0" style="margin:5px 0 7px;"><tr>'
                f'<td style="vertical-align:middle;padding-right:7px;white-space:nowrap;">'
                f'<span style="font-size:9px;font-weight:700;color:{C["subtle"]};text-transform:uppercase;letter-spacing:.08em;font-family:Montserrat,Arial,sans-serif;">Marka:</span></td>'
                f'<td>{items}</td></tr></table>'
            )

        user_html = ""
        if cat.get("users"):
            users = cat["users"]
            show  = users[:5]
            extra = len(users) - 5 if len(users) > 5 else 0
            pills = ""
            for u in show:
                bg, fg = brand_bg_text(u.get("brand", "SB"))
                pills += (
                    f'<span style="display:inline-block;padding:2px 8px;background:{bg};color:{fg};'
                    f'border-radius:4px;font-size:10px;font-weight:700;margin:2px 3px 2px 0;'
                    f'font-family:Montserrat,Arial,sans-serif;">{u["username"]}</span>'
                )
            if extra > 0:
                pills += (
                    f'<span style="display:inline-block;padding:2px 8px;background:#f0e8ff;color:{C["subtle"]};'
                    f'border-radius:4px;font-size:10px;font-weight:600;margin:2px 3px 2px 0;'
                    f'font-family:Montserrat,Arial,sans-serif;">ve {extra} kişi daha</span>'
                )
            user_html = (
                f'<table cellpadding="0" cellspacing="0" style="margin-top:5px;"><tr>'
                f'<td style="vertical-align:top;padding-right:7px;white-space:nowrap;padding-top:4px;">'
                f'<span style="font-size:9px;font-weight:700;color:{C["subtle"]};text-transform:uppercase;letter-spacing:.08em;font-family:Montserrat,Arial,sans-serif;">Uyeler:</span></td>'
                f'<td>{pills}</td></tr></table>'
            )

        note_html = ""
        if cat.get("shortNote"):
            note_html = (
                f'<p style="margin:7px 0 0;font-size:11px;color:{C["main"]};font-style:italic;'
                f'font-family:Montserrat,Arial,sans-serif;line-height:1.5;">{cat["shortNote"]}</p>'
            )

        html += f'''
        <table width="100%" cellpadding="0" cellspacing="0" style="border-bottom:{border};">
          <tr>
            <td style="width:4px;background-color:{lclr};font-size:0;" bgcolor="{lclr}">&nbsp;</td>
            <td style="padding:14px 20px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">
              <table width="100%" cellpadding="0" cellspacing="0"><tr>
                <td><span style="font-size:13px;font-weight:700;color:{C["dark"]};font-family:Montserrat,Arial,sans-serif;">{cat.get("name","Konu")}</span></td>
                <td style="text-align:right;white-space:nowrap;padding-left:10px;">
                  <span style="font-size:26px;font-weight:800;color:{lclr};line-height:1;font-family:Montserrat,Arial,sans-serif;">{cnt}</span>
                </td>
              </tr></table>
              {bd_html}{user_html}{note_html}
            </td>
          </tr>
        </table>'''
    return html

def render_ai(ai, ai_usage):
    if not ai:
        return f'<p style="margin:0;font-size:13px;color:{C["lavanta"]};font-family:Montserrat,Arial,sans-serif;">Bu donem icin kayit bulunamadi.</p>'
    if ai.get("error") and not ai.get("categories"):
        return f'<p style="margin:0;font-size:13px;color:{C["lavanta"]};font-family:Montserrat,Arial,sans-serif;">AI analizi alinamadi.</p>'

    model_badge = ""
    if ai.get("modelUsed"):
        model_badge = (
            f'<span style="display:inline-block;padding:2px 9px;background:#f3e8ff;color:{C["main"]};'
            f'border:1px solid {C["lavanta"]};border-radius:3px;font-size:10px;font-weight:700;'
            f'font-family:Montserrat,Arial,sans-serif;">{ai["modelUsed"]}</span>'
        )

    html = (
        f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:16px;"><tr>'
        f'<td style="font-size:11px;font-weight:800;color:{C["dark"]};text-transform:uppercase;'
        f'letter-spacing:.14em;font-family:Montserrat,Arial,sans-serif;">AI Analizi</td>'
        f'<td style="text-align:right;">{model_badge}</td>'
        f'</tr></table>'
    )

    if ai.get("categories"):
        rows_html = topic_rows(ai["categories"])
        html += f'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {C["border"]};border-radius:8px;overflow:hidden;">{rows_html}</table>'

    if ai.get("critical"):
        html += f'<p style="margin:22px 0 10px;font-size:12px;font-weight:800;color:{C["main"]};letter-spacing:.06em;text-transform:capitalize;font-family:Montserrat,Arial,sans-serif;">Kritik Kay\u0131tlar</p>'
        crit_rows = ""
        for idx, u in enumerate(ai["critical"]):
            is_last = idx == len(ai["critical"]) - 1
            border  = "none" if is_last else f"1px solid {C['divider']}"
            bg, fg  = brand_bg_text(u.get("brand","SB"))
            crit_rows += (
                f'<tr style="border-bottom:{border};"><td style="padding:10px 14px;vertical-align:middle;white-space:nowrap;width:1%;">'
                f'<table cellpadding="0" cellspacing="0"><tr>'
                f'<td style="padding-right:8px;vertical-align:middle;"><span style="font-size:13px;font-weight:700;color:{C["dark"]};font-family:Montserrat,Arial,sans-serif;white-space:nowrap;">{u.get("username","-")}</span></td>'
                f'<td style="vertical-align:middle;"><span style="display:inline-block;padding:1px 7px;background:{bg};color:{fg};border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;white-space:nowrap;">{u.get("brand","-")}</span></td>'
                f'</tr></table></td>'
                f'<td style="padding:10px 14px;font-size:12px;color:{C["main"]};line-height:1.5;font-family:Montserrat,Arial,sans-serif;">{u.get("reason","-")}</td></tr>'
            )
        html += f'<table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid {C["border"]};border-top:3px solid {C["yellow"]};border-radius:0 0 8px 8px;margin-bottom:18px;" bgcolor="#faf7ff">{crit_rows}</table>'

    if ai.get("actionItems"):
        html += f'<p style="margin:18px 0 10px;font-size:12px;font-weight:800;color:{C["main"]};letter-spacing:.06em;text-transform:capitalize;font-family:Montserrat,Arial,sans-serif;">\u00d6nerilen Aksiyonlar</p>'
        for i, a in enumerate(ai["actionItems"]):
            html += (
                f'<table cellpadding="0" cellspacing="0" style="margin-bottom:9px;width:100%;"><tr>'
                f'<td style="vertical-align:top;width:24px;padding-right:10px;">'
                f'<span style="display:inline-block;width:20px;height:20px;background:{C["dark"]};color:{C["yellow"]};border-radius:50%;font-size:10px;font-weight:700;text-align:center;line-height:20px;font-family:Montserrat,Arial,sans-serif;">{i+1}</span>'
                f'</td><td style="font-size:13px;color:#4b5563;line-height:1.6;vertical-align:top;font-family:Montserrat,Arial,sans-serif;">{a}</td>'
                f'</tr></table>'
            )

    if ai.get("summary"):
        html += (
            f'<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;border-top:1px solid {C["divider"]};"><tr>'
            f'<td style="padding-top:14px;">'
            f'<p style="margin:0 0 5px;font-size:12px;font-weight:800;color:{C["main"]};letter-spacing:.06em;text-transform:capitalize;font-family:Montserrat,Arial,sans-serif;">\u00d6zet</p>'
            f'<p style="margin:0;font-size:13px;color:#4b5563;line-height:1.7;font-family:Montserrat,Arial,sans-serif;">{ai["summary"]}</p>'
            f'</td></tr></table>'
        )

    return html

def build_daily_html(metrics, ai, date_label, ai_usage):
    brand_parts = ""
    for b in BRANDS:
        cnt = metrics["by_brand"].get(b, 0)
        if cnt:
            brand_parts += f'<span style="font-size:12px;color:#ffffff;font-weight:700;margin:0 8px;font-family:Montserrat,Arial,sans-serif;">{cnt} <span style="color:#e9d5ff;font-weight:600;font-size:11px;">{b}</span></span>'

    cat_parts = '<span style="color:rgba(233,213,255,0.5);margin:0 6px;">&middot;</span>'.join(
        f'<span style="font-size:10px;color:{C["yellow"]};font-weight:600;font-family:Montserrat,Arial,sans-serif;">{item["cat"]} {item["cnt"]}</span>'
        for item in metrics["category_ranked"][:5]
    )

    token_line = ""
    if ai_usage and ai_usage.get("in_tokens"):
        u = ai_usage
        token_line = (
            f'<p style="margin:5px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">'
            f'{u["model"]} &nbsp;&middot;&nbsp; {u["in_tokens"]:,} input + {u["out_tokens"]:,} output token &nbsp;&middot;&nbsp; ${u["cost"]:.6f}'
            f'</p>'
        )

    ai_section = render_ai(ai, ai_usage)
    ai_wrap = f'<table width="100%" cellpadding="0" cellspacing="0"><tr><td style="padding:20px 28px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">{ai_section}</td></tr></table>'

    return f'''<!DOCTYPE html><html>
<head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background-color:#f0f0f0;" bgcolor="#f0f0f0">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f0f0">
<tr><td align="center" style="padding:32px 16px;">
<table width="600" cellpadding="0" cellspacing="0"
  style="max-width:600px;width:100%;background-color:#ffffff;border-radius:16px;overflow:hidden;
         box-shadow:0 16px 60px rgba(47,21,85,0.35),0 4px 20px rgba(102,45,145,0.25);"
  bgcolor="#ffffff">
<tr><td>

<table width="100%" cellpadding="0" cellspacing="0">
<tr><td style="background-color:{C["dark"]};padding:28px 28px 24px;text-align:center;" bgcolor="{C["dark"]}">
  <p style="margin:0 0 8px;font-size:9px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.2em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">Poligon &middot; CS Ops &middot; Gunluk Rapor</p>
  <p style="margin:0;font-size:64px;font-weight:900;color:{C["yellow"]};line-height:1;font-family:Montserrat,Arial,sans-serif;">{metrics["total"]}</p>
  <p style="margin:3px 0 2px;font-size:10px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.22em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">KAYIT</p>
  <p style="margin:0 0 16px;font-size:13px;color:{C["yellow"]};font-weight:700;font-family:Montserrat,Arial,sans-serif;">{date_label}</p>
  <table cellpadding="0" cellspacing="0" align="center"><tr>
    <td style="background-color:rgba(255,255,255,0.1);border-radius:20px;padding:7px 20px;border:1px solid rgba(178,138,191,0.3);">{brand_parts}</td>
  </tr></table>
  <p style="margin:10px 0 0;font-family:Montserrat,Arial,sans-serif;">{cat_parts}</p>
</td></tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td>
</tr></table>

{ai_wrap}

<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:{C["dark"]};" bgcolor="{C["dark"]}">
  <p style="margin:0;font-size:11px;color:{C["yellow"]};font-weight:700;text-align:center;font-family:Montserrat,Arial,sans-serif;">POLIGON CS Feedback AI Report &nbsp;&middot;&nbsp; Gunluk</p>
  {token_line}
  <p style="margin:4px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>

</td></tr></table>
</td></tr></table>
</body></html>'''

def build_weekly_html(metrics, ai, week_label, breakdown, ai_usage):
    brand_parts = ""
    for b in BRANDS:
        cnt = metrics["by_brand"].get(b, 0)
        if cnt:
            brand_parts += f'<span style="font-size:12px;color:#ffffff;font-weight:700;margin:0 8px;font-family:Montserrat,Arial,sans-serif;">{cnt} <span style="color:#e9d5ff;font-weight:600;font-size:11px;">{b}</span></span>'

    cat_parts = '<span style="color:rgba(233,213,255,0.5);margin:0 6px;">&middot;</span>'.join(
        f'<span style="font-size:10px;color:{C["yellow"]};font-weight:600;font-family:Montserrat,Arial,sans-serif;">{item["cat"]} {item["cnt"]}</span>'
        for item in metrics["category_ranked"][:5]
    )

    max_day = max((d["count"] for d in breakdown), default=1) or 1
    bars = ""
    for d in breakdown:
        h   = max(4, round(d["count"] / max_day * 44)) if d["count"] else 2
        bgc = C["main"] if d["count"] else C["divider"]
        cnt_html = (f'<p style="margin:0 0 4px;font-size:11px;font-weight:700;color:{C["main"]};font-family:Montserrat,Arial,sans-serif;">{d["count"]}</p>'
                    if d["count"] else f'<p style="margin:0 0 4px;font-size:11px;color:{C["divider"]};font-family:Montserrat,Arial,sans-serif;">&nbsp;</p>')
        bars += (
            f'<td style="text-align:center;vertical-align:bottom;padding:0 4px;">{cnt_html}'
            f'<table cellpadding="0" cellspacing="0" align="center"><tr>'
            f'<td style="background-color:{bgc};height:{h}px;width:24px;" bgcolor="{bgc}"><p style="margin:0;font-size:0;line-height:0;">&nbsp;</p></td>'
            f'</tr></table>'
            f'<p style="margin:5px 0 0;font-size:9px;color:{C["lavanta"]};white-space:nowrap;font-family:Montserrat,Arial,sans-serif;">{d["label"]}</p>'
            f'</td>'
        )

    repeat_html = ""
    if metrics["repeat_users"]:
        rows_html = ""
        for idx, u in enumerate(metrics["repeat_users"][:10]):
            is_last = idx == min(9, len(metrics["repeat_users"]) - 1)
            border  = "none" if is_last else f"1px solid {C['divider']}"
            bg, fg  = brand_bg_text(u["brand"])
            rows_html += (
                f'<tr style="border-bottom:{border};">'
                f'<td style="padding:9px 0;font-size:13px;font-weight:700;color:{C["dark"]};font-family:Montserrat,Arial,sans-serif;">{u["username"]}</td>'
                f'<td style="padding:9px 0;"><span style="display:inline-block;padding:1px 7px;background:{bg};color:{fg};border-radius:3px;font-size:10px;font-weight:700;font-family:Montserrat,Arial,sans-serif;">{u["brand"]}</span></td>'
                f'<td style="padding:9px 0;text-align:right;"><span style="font-size:12px;color:{C["main"]};font-weight:700;font-family:Montserrat,Arial,sans-serif;">{u["count"]}x kayit</span></td>'
                f'</tr>'
            )
        repeat_html = (
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td></tr></table>'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="padding:20px 28px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">'
            f'<p style="margin:0 0 12px;font-size:10px;font-weight:700;color:{C["lavanta"]};letter-spacing:.14em;text-transform:uppercase;font-family:Montserrat,Arial,sans-serif;">Tekrar Eden Uyeler</p>'
            f'<table width="100%" cellpadding="0" cellspacing="0">{rows_html}</table>'
            f'</td></tr></table>'
        )

    token_line = ""
    if ai_usage and ai_usage.get("in_tokens"):
        u = ai_usage
        token_line = (
            f'<p style="margin:5px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">'
            f'{u["model"]} &nbsp;&middot;&nbsp; {u["in_tokens"]:,} input + {u["out_tokens"]:,} output token &nbsp;&middot;&nbsp; ${u["cost"]:.6f}'
            f'</p>'
        )

    ai_section = render_ai(ai, ai_usage)
    ai_wrap    = f'<table width="100%" cellpadding="0" cellspacing="0"><tr><td style="padding:20px 28px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">{ai_section}</td></tr></table>'

    return f'''<!DOCTYPE html><html>
<head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap" rel="stylesheet">
</head>
<body style="margin:0;padding:0;background-color:#f0f0f0;" bgcolor="#f0f0f0">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f0f0">
<tr><td align="center" style="padding:32px 16px;">
<table width="600" cellpadding="0" cellspacing="0"
  style="max-width:600px;width:100%;background-color:#ffffff;border-radius:16px;overflow:hidden;
         box-shadow:0 16px 60px rgba(47,21,85,0.35),0 4px 20px rgba(102,45,145,0.25);"
  bgcolor="#ffffff">
<tr><td>

<table width="100%" cellpadding="0" cellspacing="0">
<tr><td style="background-color:{C["dark"]};padding:28px 28px 24px;text-align:center;" bgcolor="{C["dark"]}">
  <p style="margin:0 0 8px;font-size:9px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.2em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">Poligon &middot; CS Ops &middot; Haftalik Rapor</p>
  <p style="margin:0;font-size:64px;font-weight:900;color:{C["yellow"]};line-height:1;font-family:Montserrat,Arial,sans-serif;">{metrics["total"]}</p>
  <p style="margin:3px 0 2px;font-size:10px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.22em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">KAYIT</p>
  <p style="margin:0 0 16px;font-size:13px;color:{C["yellow"]};font-weight:700;font-family:Montserrat,Arial,sans-serif;">{week_label}</p>
  <table cellpadding="0" cellspacing="0" align="center"><tr>
    <td style="background-color:rgba(255,255,255,0.1);border-radius:20px;padding:7px 20px;border:1px solid rgba(178,138,191,0.3);">{brand_parts}</td>
  </tr></table>
  <p style="margin:10px 0 0;font-family:Montserrat,Arial,sans-serif;">{cat_parts}</p>
</td></tr>
</table>

<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td>
</tr></table>

<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:20px 28px 22px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">
  <p style="margin:0 0 14px;font-size:10px;font-weight:700;color:{C["lavanta"]};letter-spacing:.14em;text-transform:uppercase;font-family:Montserrat,Arial,sans-serif;">Gunluk Dagilim</p>
  <table cellpadding="0" cellspacing="0"><tr style="vertical-align:bottom;">{bars}</tr></table>
</td></tr></table>

<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td>
</tr></table>

{ai_wrap}
{repeat_html}

<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:{C["dark"]};" bgcolor="{C["dark"]}">
  <p style="margin:0;font-size:11px;color:{C["yellow"]};font-weight:700;text-align:center;font-family:Montserrat,Arial,sans-serif;">POLIGON CS Feedback AI Report &nbsp;&middot;&nbsp; Haftalik</p>
  {token_line}
  <p style="margin:4px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>

</td></tr></table>
</td></tr></table>
</body></html>'''

# ── EMAIL ─────────────────────────────────────────────────────
def send_email(html, subject):
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

def send_email_with_attachment(html, subject, attachment_path, attachment_name):
    from email.mime.base import MIMEBase
    from email import encoders
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])
    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)
    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f'attachment; filename="{attachment_name}"')
    msg.attach(part)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
    print(f"[EMAIL] Attachment ile gönderildi: {', '.join(REPORT_EMAILS)}")

# ── MAIN ──────────────────────────────────────────────────────
def run_daily(date_str=None):
    if not date_str:
        date_str = get_sofia_yesterday()
    lbl = date_tr(date_str)
    print(f"\n[DAILY] {date_str}")

    rows    = load_rows(date_str, date_str)
    metrics = compute_metrics(rows)
    ai, model_used, ai_usage = analyze_with_ai(rows, "daily", lbl) if rows else (None, "—", {})

    html    = build_daily_html(metrics, ai, lbl, ai_usage)
    subject = f"CS Feedback | {lbl}" + (f" | {metrics['total']} Kayit" if metrics["total"] else " | Kayit Yok")
    send_email(html, subject)
    print(f"[DAILY] Tamamlandı — {metrics['total']} kayıt")

def run_weekly(start_str=None, end_str=None):
    import os, pytz
    if not end_str:
        sofia = pytz.timezone("Europe/Sofia")
        now   = datetime.now(sofia)
        dow   = now.weekday()
        end   = now - timedelta(days=dow + 1)
        start = end - timedelta(days=6)
        end_str   = end.strftime("%Y-%m-%d")
        start_str = start.strftime("%Y-%m-%d")

    lbl = week_label_tr(start_str, end_str)
    print(f"\n[WEEKLY] {start_str} - {end_str}")

    rows      = load_rows(start_str, end_str)
    metrics   = compute_metrics(rows)
    breakdown = daily_breakdown(rows, start_str, end_str)
    ai, model_used, ai_usage = analyze_with_ai(rows, "weekly", lbl) if rows else (None, "—", {})

    html      = build_weekly_html(metrics, ai, lbl, breakdown, ai_usage)
    word_name = f"CS_Feedback_Haftalik_{start_str}_{end_str}.docx"
    subject   = f"CS Feedback Haftalik | {lbl} | {metrics['total']} Kayit"

    # Word attachment oluştur
    word_path = build_word_doc(ai, metrics, lbl, breakdown, ai_usage)
    send_email_with_attachment(html, subject, word_path, word_name)
    os.unlink(word_path)
    print(f"[WEEKLY] Tamamlandi — {metrics['total']} kayit")

def main():
    import sys
    mode = sys.argv[1] if len(sys.argv) > 1 else "daily"
    if mode == "weekly":
        run_weekly()
    else:
        run_daily()

if __name__ == "__main__":
    main()

# ── WORD ATTACHMENT ───────────────────────────────────────────
def build_word_doc(ai, metrics, lbl, breakdown, ai_usage):
    from docx import Document
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

    # Başlık
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("CS Feedback Haftalik Rapor")
    r.bold = True; r.font.size = Pt(18)
    r.font.color.rgb = RGBColor(47, 21, 85)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r2 = sub.add_run(lbl)
    r2.font.size = Pt(12); r2.font.color.rgb = RGBColor(102, 45, 145)
    doc.add_paragraph()

    # Genel bakış
    h = doc.add_heading("Genel Bakis", level=1)
    for run in h.runs:
        run.font.color.rgb = RGBColor(47, 21, 85)
    tp = doc.add_paragraph()
    tp.add_run(f"Toplam: {metrics['total']} kayit   |   ")
    for b in ["SB", "BS", "TB"]:
        cnt = metrics["by_brand"].get(b, 0)
        if cnt:
            tp.add_run(f"{b}: {cnt}   ")

    # Kategori tablosu
    doc.add_paragraph()
    h2 = doc.add_heading("Konu Dagilimi", level=2)
    for run in h2.runs:
        run.font.color.rgb = RGBColor(102, 45, 145)
    if metrics.get("category_ranked"):
        tbl = doc.add_table(rows=1, cols=2)
        tbl.style = "Table Grid"
        hdr = tbl.rows[0].cells
        hdr[0].text = "Kategori"; hdr[1].text = "Adet"
        for cell in hdr:
            for run in cell.paragraphs[0].runs:
                run.bold = True
        for item in metrics["category_ranked"]:
            row = tbl.add_row().cells
            row[0].text = item["cat"]
            row[1].text = str(item["cnt"])
    doc.add_paragraph()

    # AI kategorileri
    if ai and ai.get("categories"):
        h3 = doc.add_heading("AI Analizi — Detayli Konular", level=1)
        for run in h3.runs:
            run.font.color.rgb = RGBColor(47, 21, 85)
        cats = sorted(ai["categories"], key=lambda x: x.get("count", 0), reverse=True)
        for cat in cats:
            cnt = cat.get("count", 0)
            p = doc.add_paragraph()
            r = p.add_run(f"{cat.get('name','Konu')}  —  {cnt}")
            r.bold = True; r.font.size = Pt(13)
            r.font.color.rgb = RGBColor(47, 21, 85)

            if cat.get("brandBreakdown"):
                bd_text = "  |  ".join([f"{bd['brand']}: {bd['count']}" for bd in cat["brandBreakdown"]])
                bp = doc.add_paragraph()
                br = bp.add_run(f"Marka: {bd_text}")
                br.font.size = Pt(10); br.font.color.rgb = RGBColor(124, 58, 237)

            if cat.get("users"):
                users_text = ",  ".join([u["username"] for u in cat["users"]])
                up = doc.add_paragraph()
                ur = up.add_run(f"Uyeler: {users_text}")
                ur.font.size = Pt(10); ur.italic = True

            if cat.get("shortNote"):
                np2 = doc.add_paragraph()
                nr = np2.add_run(cat["shortNote"])
                nr.font.size = Pt(11); nr.italic = True
                nr.font.color.rgb = RGBColor(102, 45, 145)
            doc.add_paragraph()

    # Kritik
    if ai and ai.get("critical"):
        h4 = doc.add_heading("Kritik Kayitlar", level=1)
        for run in h4.runs:
            run.font.color.rgb = RGBColor(47, 21, 85)
        for u in ai["critical"]:
            p = doc.add_paragraph()
            r1 = p.add_run(f"{u.get('username','-')} ({u.get('brand','-')}): ")
            r1.bold = True; r1.font.color.rgb = RGBColor(47, 21, 85)
            r2 = p.add_run(u.get("reason", "-"))
            r2.font.color.rgb = RGBColor(102, 45, 145)
        doc.add_paragraph()

    # Aksiyonlar
    if ai and ai.get("actionItems"):
        h5 = doc.add_heading("Onerilen Aksiyonlar", level=1)
        for run in h5.runs:
            run.font.color.rgb = RGBColor(47, 21, 85)
        for i, a in enumerate(ai["actionItems"]):
            p = doc.add_paragraph()
            p.add_run(f"{i+1}. {a}").font.size = Pt(11)
        doc.add_paragraph()

    # Özet
    if ai and ai.get("summary"):
        h6 = doc.add_heading("Ozet", level=1)
        for run in h6.runs:
            run.font.color.rgb = RGBColor(47, 21, 85)
        doc.add_paragraph(ai["summary"])
        doc.add_paragraph()

    # Tekrar eden üyeler
    if metrics.get("repeat_users"):
        h7 = doc.add_heading("Tekrar Eden Uyeler", level=1)
        for run in h7.runs:
            run.font.color.rgb = RGBColor(47, 21, 85)
        tbl2 = doc.add_table(rows=1, cols=3)
        tbl2.style = "Table Grid"
        hdr2 = tbl2.rows[0].cells
        hdr2[0].text = "Kullanici"; hdr2[1].text = "Marka"; hdr2[2].text = "Kayit"
        for cell in hdr2:
            for run in cell.paragraphs[0].runs:
                run.bold = True
        for u in metrics["repeat_users"][:20]:
            row = tbl2.add_row().cells
            row[0].text = u["username"]
            row[1].text = u["brand"]
            row[2].text = f"{u['count']}x"

    # Footer
    doc.add_paragraph()
    fp = doc.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fr = fp.add_run("POLIGON CS Feedback AI Report — Haftalik")
    fr.font.size = Pt(9); fr.font.color.rgb = RGBColor(178, 138, 191)
    if ai_usage and ai_usage.get("model"):
        u = ai_usage
        fp2 = doc.add_paragraph()
        fp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fr2 = fp2.add_run(f"{u['model']}  |  {u.get('in_tokens',0):,}+{u.get('out_tokens',0):,} token  |  ${u.get('cost',0):.6f}")
        fr2.font.size = Pt(9); fr2.font.color.rgb = RGBColor(178, 138, 191)

    import tempfile
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(tmp.name)
    return tmp.name

# ── KISA HAFTALIK HTML (email body) ───────────────────────────
def build_weekly_short_html(metrics, ai, week_label, breakdown, ai_usage):
    brand_parts = ""
    for b in BRANDS:
        cnt = metrics["by_brand"].get(b, 0)
        if cnt:
            brand_parts += f'<span style="font-size:12px;color:#ffffff;font-weight:700;margin:0 8px;font-family:Montserrat,Arial,sans-serif;">{cnt} <span style="color:#e9d5ff;font-weight:600;font-size:11px;">{b}</span></span>'

    cat_parts = '<span style="color:rgba(233,213,255,0.5);margin:0 6px;">&middot;</span>'.join(
        f'<span style="font-size:10px;color:{C["yellow"]};font-weight:600;font-family:Montserrat,Arial,sans-serif;">{item["cat"]} {item["cnt"]}</span>'
        for item in metrics["category_ranked"][:5]
    )

    max_day = max((d["count"] for d in breakdown), default=1) or 1
    bars = ""
    for d in breakdown:
        h   = max(4, round(d["count"] / max_day * 44)) if d["count"] else 2
        bgc = C["main"] if d["count"] else C["divider"]
        cnt_h = (f'<p style="margin:0 0 4px;font-size:11px;font-weight:700;color:{C["main"]};font-family:Montserrat,Arial,sans-serif;">{d["count"]}</p>'
                 if d["count"] else f'<p style="margin:0 0 4px;font-size:11px;color:{C["divider"]};font-family:Montserrat,Arial,sans-serif;">&nbsp;</p>')
        bars += (f'<td style="text-align:center;vertical-align:bottom;padding:0 4px;">{cnt_h}'
                 f'<table cellpadding="0" cellspacing="0" align="center"><tr>'
                 f'<td style="background-color:{bgc};height:{h}px;width:24px;" bgcolor="{bgc}"><p style="margin:0;font-size:0;line-height:0;">&nbsp;</p></td>'
                 f'</tr></table>'
                 f'<p style="margin:5px 0 0;font-size:9px;color:{C["lavanta"]};white-space:nowrap;font-family:Montserrat,Arial,sans-serif;">{d["label"]}</p>'
                 f'</td>')

    summary_html = ""
    if ai and ai.get("summary"):
        summary_html = (
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td></tr></table>'
            f'<table width="100%" cellpadding="0" cellspacing="0"><tr>'
            f'<td style="padding:20px 28px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">'
            f'<p style="margin:0 0 8px;font-size:11px;font-weight:800;color:{C["dark"]};text-transform:uppercase;letter-spacing:.14em;font-family:Montserrat,Arial,sans-serif;">Ozet</p>'
            f'<p style="margin:0;font-size:13px;color:#4b5563;line-height:1.7;font-family:Montserrat,Arial,sans-serif;">{ai["summary"]}</p>'
            f'<p style="margin:14px 0 0;font-size:12px;color:{C["main"]};font-style:italic;font-family:Montserrat,Arial,sans-serif;">'
            f'Tam analiz ekte Word dosyasi olarak gonderilmistir.</p>'
            f'</td></tr></table>'
        )

    token_line = ""
    if ai_usage and ai_usage.get("in_tokens"):
        u = ai_usage
        token_line = (f'<p style="margin:5px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">'
                      f'{u["model"]} &nbsp;&middot;&nbsp; {u["in_tokens"]:,} input + {u["out_tokens"]:,} output token &nbsp;&middot;&nbsp; ${u["cost"]:.6f}</p>')

    return f'''<!DOCTYPE html><html>
<head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800;900&display=swap" rel="stylesheet">
</head><body style="margin:0;padding:0;background-color:#f0f0f0;" bgcolor="#f0f0f0">
<table width="100%" cellpadding="0" cellspacing="0" bgcolor="#f0f0f0">
<tr><td align="center" style="padding:32px 16px;">
<table width="600" cellpadding="0" cellspacing="0"
  style="max-width:600px;width:100%;background-color:#ffffff;border-radius:16px;overflow:hidden;
         box-shadow:0 16px 60px rgba(47,21,85,0.35),0 4px 20px rgba(102,45,145,0.25);" bgcolor="#ffffff">
<tr><td>
<table width="100%" cellpadding="0" cellspacing="0">
<tr><td style="background-color:{C["dark"]};padding:28px 28px 24px;text-align:center;" bgcolor="{C["dark"]}">
  <p style="margin:0 0 8px;font-size:9px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.2em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">Poligon &middot; CS Ops &middot; Haftalik Rapor</p>
  <p style="margin:0;font-size:64px;font-weight:900;color:{C["yellow"]};line-height:1;font-family:Montserrat,Arial,sans-serif;">{metrics["total"]}</p>
  <p style="margin:3px 0 2px;font-size:10px;color:{C["lavanta"]};text-transform:uppercase;letter-spacing:.22em;font-weight:700;font-family:Montserrat,Arial,sans-serif;">KAYIT</p>
  <p style="margin:0 0 16px;font-size:13px;color:{C["yellow"]};font-weight:700;font-family:Montserrat,Arial,sans-serif;">{week_label}</p>
  <table cellpadding="0" cellspacing="0" align="center"><tr>
    <td style="background-color:rgba(255,255,255,0.1);border-radius:20px;padding:7px 20px;border:1px solid rgba(178,138,191,0.3);">{brand_parts}</td>
  </tr></table>
  <p style="margin:10px 0 0;font-family:Montserrat,Arial,sans-serif;">{cat_parts}</p>
</td></tr></table>
<table width="100%" cellpadding="0" cellspacing="0"><tr>
  <td style="height:1px;background-color:{C["divider"]};font-size:0;" bgcolor="{C["divider"]}">&nbsp;</td>
</tr></table>
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:20px 28px 22px;background-color:{C["bgBody"]};" bgcolor="{C["bgBody"]}">
  <p style="margin:0 0 14px;font-size:10px;font-weight:700;color:{C["lavanta"]};letter-spacing:.14em;text-transform:uppercase;font-family:Montserrat,Arial,sans-serif;">Gunluk Dagilim</p>
  <table cellpadding="0" cellspacing="0"><tr style="vertical-align:bottom;">{bars}</tr></table>
</td></tr></table>
{summary_html}
<table width="100%" cellpadding="0" cellspacing="0"><tr>
<td style="padding:14px 28px;background-color:{C["dark"]};" bgcolor="{C["dark"]}">
  <p style="margin:0;font-size:11px;color:{C["yellow"]};font-weight:700;text-align:center;font-family:Montserrat,Arial,sans-serif;">POLIGON CS Feedback AI Report &nbsp;&middot;&nbsp; Haftalik</p>
  {token_line}
  <p style="margin:4px 0 0;font-size:10px;color:{C["lavanta"]};text-align:center;font-family:Montserrat,Arial,sans-serif;">Developed by Erhan</p>
</td></tr></table>
</td></tr></table>
</td></tr></table>
</body></html>'''

# ── EMAIL (attachment destekli) ───────────────────────────────
def send_email_v2(html, subject, attachment_path=None, attachment_name=None):
    from email.mime.base import MIMEBase
    from email import encoders

    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = GMAIL_USER
    msg["To"]      = REPORT_EMAILS[0]
    if len(REPORT_EMAILS) > 1:
        msg["Cc"] = ", ".join(REPORT_EMAILS[1:])

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(html, "html", "utf-8"))
    msg.attach(alt)

    if attachment_path and attachment_name:
        with open(attachment_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{attachment_name}"')
        msg.attach(part)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_USER, GMAIL_PASS)
        server.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
    print(f"[EMAIL] Gönderildi: {', '.join(REPORT_EMAILS)}")

# ── RUN WEEKLY (yeni versiyon) ────────────────────────────────
def run_weekly_v2(start_str=None, end_str=None):
    import os, pytz
    if not end_str:
        sofia = pytz.timezone("Europe/Sofia")
        now   = datetime.now(sofia)
        dow   = now.weekday()
        end   = now - timedelta(days=dow + 1)
        start = end - timedelta(days=6)
        end_str   = end.strftime("%Y-%m-%d")
        start_str = start.strftime("%Y-%m-%d")

    lbl = week_label_tr(start_str, end_str)
    print(f"\n[WEEKLY] {start_str} - {end_str}")

    rows      = load_rows(start_str, end_str)
    metrics   = compute_metrics(rows)
    breakdown = daily_breakdown(rows, start_str, end_str)
    ai, model_used, ai_usage = analyze_with_ai(rows, "weekly", lbl) if rows else (None, "—", {})

    # Kısa email body
    html = build_weekly_short_html(metrics, ai, lbl, breakdown, ai_usage)

    # Word attachment
    word_path = build_word_doc(ai, metrics, lbl, breakdown, ai_usage)
    word_name = f"CS_Feedback_Haftalik_{start_str}_{end_str}.docx"

    subject = f"CS Feedback Haftalik | {lbl} | {metrics['total']} Kayit"
    send_email_v2(html, subject, word_path, word_name)

    os.unlink(word_path)
    print(f"[WEEKLY] Tamamlandi — {metrics['total']} kayit")
