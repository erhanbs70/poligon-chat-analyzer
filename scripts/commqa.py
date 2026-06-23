#!/usr/bin/env python3
# ============================================================
# POLIGON COMMQA — v1 (GitHub Actions)
# Rating 1-2-3 + Comment sheet  &  QA Tag Listesi sheet
# Google Sheets API ile direkt yazar — GAS timeout yok
# ============================================================

import os
import sys
import json
import base64
import requests
import smtplib
import time
import traceback
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ── CONFIG ───────────────────────────────────────────────────
SITE_ID        = os.environ["COMM100_SITE_ID"]
API_KEY        = os.environ["COMM100_API_KEY"]
COMM100_EMAIL  = os.environ["COMM100_EMAIL"]
SPREADSHEET_ID = os.environ["COMMQA_SPREADSHEET_ID"]
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]
GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASS     = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS  = [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

PORTAL_BASE = (
    "https://dash15.lively-chat.com/ui/90005373"
    "/livechat/history/chats/transcriptdetail"
)

QA_TAGS = [
    "game_fairness",
    "ac_closure_request",
    "casino_cashback_query",
    "sport_cashback_query",
    "deposit_issue",
    "deposit_missing",
    "betting_rules_query",
    "deposit_query",
]
VIP_TAGS = ["VIP TIER3", "VIP"]

# Sheet adları — mevcut GAS çıktısıyla aynı
RATING_SHEET = "Rating & Comment"
TAG_SHEET    = "🏷️ Tag Listesi"
LOG_SHEET    = "__log__"

# ── AUTH ─────────────────────────────────────────────────────
def comm100_auth():
    return "Basic " + base64.b64encode(
        f"{COMM100_EMAIL}:{API_KEY}".encode()
    ).decode()

# ── DATE — DST-aware Sofia ───────────────────────────────────
def sofia_yesterday():
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
        end.astimezone(utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )

def in_sofia_range(ts_str, date_str):
    """Timestamp'i Sofia'ya çevir, date_str ile karşılaştır."""
    if not ts_str:
        return True
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        sofia_date = dt.astimezone(tz).strftime("%Y-%m-%d")
        return sofia_date == date_str
    except Exception:
        return True

def format_sofia(ts_str):
    """API timestamp → Sofia dd/MM/yyyy HH:mm:ss"""
    if not ts_str:
        return ""
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        return dt.astimezone(tz).strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return ts_str

# ── GOOGLE SHEETS SERVICE ────────────────────────────────────
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
    """Sheet yoksa oluştur, varsa ID döner."""
    sid = get_sheet_id(svc, sheet_name)
    if sid is not None:
        return sid
    body = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]

def clear_sheet(svc, sheet_name):
    svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'",
        body={}
    ).execute()

def write_values(svc, sheet_name, values, start="A1"):
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start}",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

def batch_format(svc, sheet_id, requests_list):
    if not requests_list:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests_list}
    ).execute()

# ── HELPERS: renk nesnesi ────────────────────────────────────
def rgb(hex_str):
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return {"red": r / 255, "green": g / 255, "blue": b / 255}

RATING_COLORS = {
    1: {"bg": "#f28b82", "fg": "#7c0000"},
    2: {"bg": "#fbbc04", "fg": "#5a3a00"},
    3: {"bg": "#fff475", "fg": "#5a4a00"},
    4: {"bg": "#ccff90", "fg": "#1a5c00"},
    5: {"bg": "#a8f0c6", "fg": "#0d4020"},
}

def cell_fmt(bg_hex, fg_hex=None, bold=False):
    fmt = {
        "backgroundColor": rgb(bg_hex),
        "textFormat": {"bold": bold},
    }
    if fg_hex:
        fmt["textFormat"]["foregroundColor"] = rgb(fg_hex)
    return fmt

def range_req(sheet_id, r1, c1, r2, c2):
    """0-indexed inclusive range for Sheets API."""
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    r1,
        "endRowIndex":      r2 + 1,
        "startColumnIndex": c1,
        "endColumnIndex":   c2 + 1,
    }

# ── COMM100 FETCH ────────────────────────────────────────────
def fetch_all_chats(date_str):
    """Tüm chatleri çek (sayfalı). Timeout yok — GH Actions'da sınır 60dk."""
    start_time, end_time = sofia_to_utc_range(date_str)
    auth    = comm100_auth()
    result  = []
    seen    = set()
    page    = 1
    base    = (
        f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
        f"?siteId={SITE_ID}"
        f"&include=chatWrapupCategory&include=postChatSurvey&include=chatAgent"
        f"&sortBy=startTime&sortOrder=asc"
    )

    print(f"[FETCH] {date_str} chatleri çekiliyor...")
    while page <= 40:
        url = f"{base}&pageIndex={page}&pageSize=500"
        try:
            r = requests.post(
                url,
                headers={"Authorization": auth, "Content-Type": "application/json"},
                json={"startTime": start_time, "endTime": end_time},
                timeout=90,
            )
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code} sayfa {page}: {r.text[:200]}")
            chats = r.json().get("list", [])
            if not chats:
                break
            for c in chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                if cid and cid not in seen:
                    seen.add(cid)
                    result.append(c)
            print(f"[FETCH] Sayfa {page}: {len(chats)} chat ({len(result)} toplam)")
            if len(chats) < 500:
                break
            page += 1
            time.sleep(0.3)
        except Exception as e:
            raise RuntimeError(f"fetch_all_chats sayfa {page}: {e}") from e

    print(f"[FETCH] Toplam {len(result)} chat çekildi")
    return result

# ── FIELD HELPERS ─────────────────────────────────────────────
def api_agent(c):
    agents = c.get("chatAgents") or []
    if agents:
        names = [
            (a.get("agent") or {}).get("displayName") or a.get("displayName") or a.get("name") or ""
            for a in agents if not a.get("botId")
        ]
        return ", ".join(filter(None, names))
    return c.get("agentName") or c.get("agent_name") or ""

def api_visitor(c):
    return c.get("preChatName") or c.get("name") or c.get("visitorName") or c.get("visitor_name") or ""

def api_site(c):
    try:
        url = (c.get("requestingPageURL")
               or (c.get("visitor") or {}).get("currentBrowsing")
               or c.get("requestPage") or "")
        import re
        m = re.match(r"https?://([^/]+)", url)
        return m.group(1) if m else url
    except Exception:
        return ""

def api_duration(c):
    try:
        ts1 = c.get("startTime") or c.get("start_time") or ""
        ts2 = c.get("endTime")   or c.get("end_time")   or ""
        if not ts1 or not ts2:
            return ""
        dt1 = datetime.fromisoformat(ts1.replace("Z", "+00:00"))
        dt2 = datetime.fromisoformat(ts2.replace("Z", "+00:00"))
        secs = int((dt2 - dt1).total_seconds())
        h  = secs // 3600
        mn = (secs % 3600) // 60
        s  = secs % 60
        return (f"{h}s " if h else "") + f"{mn}dk {s}sn"
    except Exception:
        return ""

def api_cat(c):
    return (
        (c.get("chatWrapup") or {}).get("categoriesName")
        or (c.get("wrapUp")  or {}).get("category")
        or ""
    )

def extract_main_tag(category):
    if not category:
        return ""
    import re
    inner = [m.group(1).strip() for m in re.finditer(r"\(([^)]+)\)", category)
             if not re.search(r"VIP", m.group(1), re.I)]
    for t in inner:
        if "_" in t:
            return t
    if inner:
        return inner[0]

    SKIP = re.compile(
        r"^(Wild card|Promotions|Deposit|Casino|Sport|General|WD|VIP TIER\d*|VIP)$", re.I
    )
    parts = [p.strip() for p in re.sub(r"\([^)]*\)", "", category).split(",") if p.strip()]
    for p in parts:
        if "_" in p and not SKIP.match(p):
            return p
    for p in parts:
        if not SKIP.match(p):
            return p
    if re.search(r"VIP", category, re.I):
        return "⚠️ TAG EKSİK"
    return parts[0] if parts else ""

def extract_vip_tag(category):
    if not category:
        return ""
    import re
    m = re.search(r"VIP TIER(\d+)", category, re.I)
    if m:
        return "VIP TIER" + m.group(1)
    for v in VIP_TAGS:
        if v in category:
            return v
    return ""

def all_tag_parts(cat):
    import re
    inner = [m.group(1).strip().lower() for m in re.finditer(r"\(([^)]+)\)", cat)]
    outer = [p.strip().lower() for p in re.sub(r"\([^)]*\)", "", cat).split(",") if p.strip()]
    return inner + outer

# ── FİLTRE: Rating ───────────────────────────────────────────
def filter_rating_chats(chats, date_str):
    seen   = set()
    result = []
    for c in chats:
        cid = str(c.get("id") or c.get("chatId") or "")
        if not cid or cid in seen:
            continue
        seen.add(cid)
        if not in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
            continue
        if not api_agent(c).strip():
            continue
        s = c.get("postChatSurvey") or {}
        g = s.get("ratingGrade")
        g = int(g) if g is not None else None
        cmt = (s.get("ratingComment") or "").strip()
        if (g is not None and 1 <= g <= 3) or cmt:
            result.append(c)
    print(f"[FILTER] Rating: {len(result)} chat")
    return result

# ── FİLTRE: Tag ──────────────────────────────────────────────
def filter_tag_chats(chats, date_str, exclude_ids=None):
    excl   = set(exclude_ids or [])
    seen   = set()
    result = []
    for c in chats:
        cid = str(c.get("id") or c.get("chatId") or "")
        if not cid or cid in seen or cid in excl:
            continue
        seen.add(cid)
        if not in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
            continue
        if not api_agent(c).strip():
            continue
        cat = api_cat(c)
        if not cat:
            continue
        if any(t in all_tag_parts(cat) for t in QA_TAGS):
            result.append(c)
    print(f"[FILTER] Tag: {len(result)} chat (excl: {len(excl)})")
    return result

# ── SHEETS YAZ: Rating & Comment ─────────────────────────────
def write_rating_sheet(svc, chats, date_str):
    print(f"[SHEET] Rating & Comment yazılıyor ({len(chats)} chat)...")
    sheet_id = ensure_sheet(svc, RATING_SHEET)
    clear_sheet(svc, RATING_SHEET)

    # Başlık + header
    header_row  = [["⭐ Rating 1-2-3 & Commentli Chatler   |   " + date_str + " → " + date_str]]
    col_headers = [["Temsilci İsmi", "Visitor İsmi", "Sohbet Tarihi", "Kullanılan TAG",
                    "VIP Tag", "Site Adı", "Chat Süresi", "Rating Score",
                    "User Comment", "Chat Linki"]]

    rows = []
    for c in chats:
        agent   = api_agent(c)
        visitor = api_visitor(c)
        ts      = c.get("startTime") or c.get("start_time") or ""
        cat     = api_cat(c)
        tag     = extract_main_tag(cat)
        vip     = extract_vip_tag(cat)
        site    = api_site(c)
        dur     = api_duration(c)
        s       = c.get("postChatSurvey") or {}
        g       = s.get("ratingGrade")
        rating  = int(g) if g is not None else ""
        comment = (s.get("ratingComment") or "").strip()
        cid     = str(c.get("id") or c.get("chatId") or "")
        link    = f"{PORTAL_BASE}?chatId={cid}" if cid else ""
        rows.append([agent, visitor, format_sofia(ts), tag, vip, site, dur,
                     rating, comment, link])

    all_values = header_row + col_headers + rows
    write_values(svc, RATING_SHEET, all_values)

    # ── Formatting ──
    fmt_reqs = []

    # Başlık satırı (row 0) birleştir + renk
    fmt_reqs.append({
        "mergeCells": {
            "range": range_req(sheet_id, 0, 0, 0, 9),
            "mergeType": "MERGE_ALL"
        }
    })
    fmt_reqs.append({
        "repeatCell": {
            "range": range_req(sheet_id, 0, 0, 0, 9),
            "cell": {"userEnteredFormat": {
                **cell_fmt("#1a1a2e", "#f0c040", bold=True),
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": 11,
                               "foregroundColor": rgb("#f0c040")},
                "backgroundColor": rgb("#1a1a2e"),
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })

    # Header row (row 1)
    fmt_reqs.append({
        "repeatCell": {
            "range": range_req(sheet_id, 1, 0, 1, 9),
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb("#1a73e8"),
                "textFormat": {"bold": True, "fontSize": 10,
                               "foregroundColor": rgb("#ffffff")},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }
    })

    # Freeze 2 satır
    fmt_reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"
        }
    })

    # Sütun genişlikleri (piksel)
    col_widths = [155, 130, 160, 180, 110, 160, 90, 80, 280, 200]
    for i, w in enumerate(col_widths):
        fmt_reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i + 1},
                "properties": {"pixelSize": w},
                "fields": "pixelSize"
            }
        })

    # Veri satırları — satır bazlı renklendirme
    for idx, c in enumerate(chats):
        row_i   = idx + 2  # 0-indexed, başlık+header geçtik
        cat     = api_cat(c)
        vip     = extract_vip_tag(cat)
        s       = c.get("postChatSurvey") or {}
        g       = s.get("ratingGrade")
        rating  = int(g) if g is not None else None

        default_bg = "#fff8e1" if vip else ("#f8f9fa" if idx % 2 == 0 else "#ffffff")

        # Tüm satır arka planı
        fmt_reqs.append({
            "repeatCell": {
                "range": range_req(sheet_id, row_i, 0, row_i, 9),
                "cell": {"userEnteredFormat": {"backgroundColor": rgb(default_bg)}},
                "fields": "userEnteredFormat.backgroundColor"
            }
        })

        # VIP sütunu (col 4) sarı + bold
        if vip:
            fmt_reqs.append({
                "repeatCell": {
                    "range": range_req(sheet_id, row_i, 4, row_i, 4),
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": rgb("#ffd666"),
                        "textFormat": {"bold": True}
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

        # Rating sütunu (col 7) renk
        if rating and rating in RATING_COLORS:
            rc = RATING_COLORS[rating]
            fmt_reqs.append({
                "repeatCell": {
                    "range": range_req(sheet_id, row_i, 7, row_i, 7),
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": rgb(rc["bg"]),
                        "textFormat": {"bold": True,
                                       "foregroundColor": rgb(rc["fg"])},
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

    batch_format(svc, sheet_id, fmt_reqs)
    print(f"[SHEET] Rating & Comment ✅ ({len(chats)} satır)")

# ── SHEETS YAZ: Tag Listesi ───────────────────────────────────
def write_tag_sheet(svc, chats, date_str):
    print(f"[SHEET] Tag Listesi yazılıyor ({len(chats)} chat)...")
    sheet_id = ensure_sheet(svc, TAG_SHEET)
    clear_sheet(svc, TAG_SHEET)

    header_row  = [[f"🏷️ QA Tag Listesi   |   {date_str} → {date_str}   |   {len(chats)} chat"]]
    col_headers = [["Temsilci İsmi", "Visitor İsmi", "Sohbet Tarihi", "Kullanılan Tag",
                    "VIP Tag", "Site Adı", "Chat Süresi", "Rating Score",
                    "User Comment", "Chat Linki"]]

    rows = []
    for c in chats:
        agent   = api_agent(c)
        visitor = api_visitor(c)
        ts      = c.get("startTime") or c.get("start_time") or ""
        cat     = api_cat(c)
        tag     = extract_main_tag(cat)
        vip     = extract_vip_tag(cat)
        site    = api_site(c)
        dur     = api_duration(c)
        s       = c.get("postChatSurvey") or {}
        g       = s.get("ratingGrade")
        rating  = int(g) if g is not None else ""
        comment = (s.get("ratingComment") or "").strip()
        cid     = str(c.get("id") or c.get("chatId") or "")
        link    = f"{PORTAL_BASE}?chatId={cid}" if cid else ""
        rows.append([agent, visitor, format_sofia(ts), tag, vip, site, dur,
                     rating, comment, link])

    all_values = header_row + col_headers + rows
    write_values(svc, TAG_SHEET, all_values)

    # ── Formatting ──
    fmt_reqs = []

    fmt_reqs.append({
        "mergeCells": {
            "range": range_req(sheet_id, 0, 0, 0, 9),
            "mergeType": "MERGE_ALL"
        }
    })
    fmt_reqs.append({
        "repeatCell": {
            "range": range_req(sheet_id, 0, 0, 0, 9),
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb("#1a1a2e"),
                "horizontalAlignment": "CENTER",
                "textFormat": {"bold": True, "fontSize": 11,
                               "foregroundColor": rgb("#f0c040")},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)"
        }
    })
    fmt_reqs.append({
        "repeatCell": {
            "range": range_req(sheet_id, 1, 0, 1, 9),
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb("#1a73e8"),
                "textFormat": {"bold": True, "fontSize": 10,
                               "foregroundColor": rgb("#ffffff")},
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat)"
        }
    })
    fmt_reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": 2}},
            "fields": "gridProperties.frozenRowCount"
        }
    })

    col_widths = [160, 140, 160, 180, 110, 160, 100, 80, 250, 200]
    for i, w in enumerate(col_widths):
        fmt_reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i + 1},
                "properties": {"pixelSize": w},
                "fields": "pixelSize"
            }
        })

    for idx, c in enumerate(chats):
        row_i  = idx + 2
        cat    = api_cat(c)
        tag    = extract_main_tag(cat)
        vip    = extract_vip_tag(cat)
        s      = c.get("postChatSurvey") or {}
        g      = s.get("ratingGrade")
        rating = int(g) if g is not None else None

        default_bg = "#fff8e1" if vip else ("#f8f9fa" if idx % 2 == 0 else "#ffffff")
        fmt_reqs.append({
            "repeatCell": {
                "range": range_req(sheet_id, row_i, 0, row_i, 9),
                "cell": {"userEnteredFormat": {"backgroundColor": rgb(default_bg)}},
                "fields": "userEnteredFormat.backgroundColor"
            }
        })

        if vip:
            fmt_reqs.append({
                "repeatCell": {
                    "range": range_req(sheet_id, row_i, 4, row_i, 4),
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": rgb("#ffd666"),
                        "textFormat": {"bold": True}
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

        if tag == "⚠️ TAG EKSİK":
            fmt_reqs.append({
                "repeatCell": {
                    "range": range_req(sheet_id, row_i, 3, row_i, 3),
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": rgb("#f28b82"),
                        "textFormat": {"bold": True,
                                       "foregroundColor": rgb("#7c0000")},
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

        if rating and rating in RATING_COLORS:
            rc = RATING_COLORS[rating]
            fmt_reqs.append({
                "repeatCell": {
                    "range": range_req(sheet_id, row_i, 7, row_i, 7),
                    "cell": {"userEnteredFormat": {
                        "backgroundColor": rgb(rc["bg"]),
                        "textFormat": {"bold": True,
                                       "foregroundColor": rgb(rc["fg"])},
                    }},
                    "fields": "userEnteredFormat(backgroundColor,textFormat)"
                }
            })

    batch_format(svc, sheet_id, fmt_reqs)
    print(f"[SHEET] Tag Listesi ✅ ({len(chats)} satır)")

# ── SHEETS YAZ: Log ───────────────────────────────────────────
def write_log(svc, log_type, date_str, duration_s, detail):
    sheet_id = ensure_sheet(svc, LOG_SHEET)
    # Header yoksa ekle
    existing = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A1"
    ).execute().get("values", [])
    if not existing:
        write_values(svc, LOG_SHEET,
                     [["Zaman", "Tür", "Tarih", "Süre (sn)", "Detay"]])
        fmt_reqs = [{
            "repeatCell": {
                "range": range_req(sheet_id, 0, 0, 0, 4),
                "cell": {"userEnteredFormat": {
                    "backgroundColor": rgb("#1a1a2e"),
                    "textFormat": {"bold": True,
                                   "foregroundColor": rgb("#f0c040")},
                }},
                "fields": "userEnteredFormat(backgroundColor,textFormat)"
            }
        }]
        batch_format(svc, sheet_id, fmt_reqs)

    now_str = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    svc.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [[now_str, log_type, date_str, duration_s, str(detail)]]}
    ).execute()

    # Renk — son satırı bul
    data = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A:A"
    ).execute().get("values", [])
    last_row = len(data) - 1  # 0-indexed

    is_error = "ERROR" in log_type
    is_warn  = "WARN"  in log_type
    bg = "#f28b82" if is_error else "#fff3cd" if is_warn else (
        "#f8f9fa" if last_row % 2 == 0 else "#ffffff"
    )
    fg = "#7c0000" if is_error else "#856404" if is_warn else "#000000"

    fmt_reqs = [{
        "repeatCell": {
            "range": range_req(sheet_id, last_row, 0, last_row, 4),
            "cell": {"userEnteredFormat": {
                "backgroundColor": rgb(bg),
                "textFormat": {"bold": is_error or is_warn,
                               "foregroundColor": rgb(fg)},
                "wrapStrategy": "WRAP",
            }},
            "fields": "userEnteredFormat(backgroundColor,textFormat,wrapStrategy)"
        }
    }]
    batch_format(svc, sheet_id, fmt_reqs)

# ── HATA MAİLİ ───────────────────────────────────────────────
def send_error_email(date_str, error_msg, tb_str):
    subject = f"⚠️ CommQA HATA | {date_str}"
    html = f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:20px;">
<h2 style="color:#c0392b;">CommQA — Hata Raporu</h2>
<p><strong>Tarih:</strong> {date_str}</p>
<p><strong>Hata:</strong> {error_msg}</p>
<pre style="background:#f8f9fa;padding:16px;border-radius:6px;font-size:12px;
     white-space:pre-wrap;word-break:break-all;">{tb_str}</pre>
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
        print(f"[EMAIL] Hata maili gönderildi: {', '.join(REPORT_EMAILS)}")
    except Exception as e:
        print(f"[EMAIL] Hata maili gönderilemedi: {e}")

# ── MAIN ─────────────────────────────────────────────────────
def main():
    t0       = time.time()
    date_str = sofia_yesterday()
    print(f"\n{'='*55}\nCOMMQA — {date_str}\n{'='*55}\n")

    try:
        svc = sheets_service()

        # 1. Tüm chatleri çek
        all_chats = fetch_all_chats(date_str)
        if not all_chats:
            print("[MAIN] Chat bulunamadı, çıkılıyor.")
            write_log(svc, "INFO", date_str, "0", "Hiç chat bulunamadı")
            return

        # 2. Rating & Comment
        rating_chats = filter_rating_chats(all_chats, date_str)
        rating_ids   = {str(c.get("id") or c.get("chatId") or "") for c in rating_chats}
        write_rating_sheet(svc, rating_chats, date_str)

        # 3. Tag Listesi (Rating chatlerini hariç tut)
        tag_chats = filter_tag_chats(all_chats, date_str, exclude_ids=rating_ids)
        write_tag_sheet(svc, tag_chats, date_str)

        # 4. Log
        elapsed = round(time.time() - t0, 1)
        detail  = (f"{len(rating_chats)} rating chat | "
                   f"{len(tag_chats)} tag chat | "
                   f"toplam fetch: {len(all_chats)} | "
                   f"süre: {elapsed}s")
        write_log(svc, "COMMQA ✅", date_str, elapsed, detail)
        print(f"\n[DONE] {detail}")

    except Exception as e:
        elapsed = round(time.time() - t0, 1)
        tb_str  = traceback.format_exc()
        print(f"[ERROR] {e}\n{tb_str}")

        # Sheets log (mümkünse)
        try:
            svc = sheets_service()
            write_log(svc, "COMMQA ERROR", date_str, elapsed,
                      f"❌ {e}\n{tb_str}")
        except Exception as log_err:
            print(f"[LOG] Log yazılamadı: {log_err}")

        # Hata maili
        send_error_email(date_str, str(e), tb_str)

        sys.exit(1)

if __name__ == "__main__":
    main()
