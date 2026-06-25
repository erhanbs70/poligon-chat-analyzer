#!/usr/bin/env python3
# ============================================================
# POLIGON COMMQA — v2 (GitHub Actions)
# Rating 1-2-3 + Comment sheet  &  QA Tag Listesi sheet
# Google Sheets API ile direkt yazar — GAS timeout yok
# v2: mergeCells kaldırıldı (clear sonrası merge hatası fix)
#     batch format optimize — satır başına tek request yerine
#     renk gruplarına göre toplu request
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
    if not ts_str:
        return True
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        return dt.astimezone(tz).strftime("%Y-%m-%d") == date_str
    except Exception:
        return True

def format_sofia(ts_str):
    if not ts_str:
        return ""
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    try:
        dt = datetime.fromisoformat(ts_str.replace("Z", "+00:00"))
        return dt.astimezone(tz).strftime("%d/%m/%Y %H:%M:%S")
    except Exception:
        return ts_str

# ── GOOGLE SHEETS ────────────────────────────────────────────
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
    body = {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]}
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]

def clear_sheet(svc, sheet_id, sheet_name):
    """İçeriği temizle + tüm merge'leri kaldır."""
    svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'",
        body={}
    ).execute()
    # Mevcut merge'leri kaldır — yoksa yeni merge hata verir
    try:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{
                "unmergeCells": {
                    "range": {
                        "sheetId":          sheet_id,
                        "startRowIndex":    0,
                        "endRowIndex":      1,
                        "startColumnIndex": 0,
                        "endColumnIndex":   10,
                    }
                }
            }]}
        ).execute()
    except Exception:
        pass  # Merge yoksa hata fırlatır, ignore

def write_values(svc, sheet_name, values, start="A1"):
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start}",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

def apply_hyperlinks(svc, sheet_id, links, data_start_row, link_col=9):
    """links: URL listesi. Her satıra Chati Aç + tıklanabilir link yazar."""
    reqs = []
    for idx, url in enumerate(links):
        if not url:
            continue
        row_i = data_start_row + idx
        reqs.append({
            "updateCells": {
                "range": rng(sheet_id, row_i, link_col, row_i, link_col),
                "rows": [{
                    "values": [{
                        "userEnteredValue": {"stringValue": "Chati Ac"},
                        "userEnteredFormat": {
                            "horizontalAlignment": "CENTER",
                            "textFormat": {
                                "link":            {"uri": url},
                                "foregroundColor": rgb("#1155cc"),
                                "underline":       True,
                            }
                        }
                    }]
                }],
                "fields": "userEnteredValue,userEnteredFormat.horizontalAlignment,userEnteredFormat.textFormat"
            }
        })
    for i in range(0, len(reqs), 500):
        batch_format(svc, reqs[i:i + 500])

def batch_format(svc, requests_list):
    if not requests_list:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests_list}
    ).execute()

# ── RENK ─────────────────────────────────────────────────────
def rgb(hex_str):
    h = hex_str.lstrip("#")
    return {
        "red":   int(h[0:2], 16) / 255,
        "green": int(h[2:4], 16) / 255,
        "blue":  int(h[4:6], 16) / 255,
    }

RATING_COLORS = {
    1: {"bg": "#f28b82", "fg": "#7c0000"},
    2: {"bg": "#fbbc04", "fg": "#5a3a00"},
    3: {"bg": "#fff475", "fg": "#5a4a00"},
    4: {"bg": "#ccff90", "fg": "#1a5c00"},
    5: {"bg": "#a8f0c6", "fg": "#0d4020"},
}

def rng(sheet_id, r1, c1, r2, c2):
    """0-indexed inclusive GridRange."""
    return {
        "sheetId":          sheet_id,
        "startRowIndex":    r1,
        "endRowIndex":      r2 + 1,
        "startColumnIndex": c1,
        "endColumnIndex":   c2 + 1,
    }

def repeat_cell(sheet_id, r1, c1, r2, c2, bg=None, fg=None,
                bold=False, font_size=None, h_align=None, wrap=False):
    fmt = {}
    if bg:
        fmt["backgroundColor"] = rgb(bg)
    tf = {}
    if bold:
        tf["bold"] = True
    if fg:
        tf["foregroundColor"] = rgb(fg)
    if font_size:
        tf["fontSize"] = font_size
    if tf:
        fmt["textFormat"] = tf
    if h_align:
        fmt["horizontalAlignment"] = h_align
    if wrap:
        fmt["wrapStrategy"] = "WRAP"

    fields = []
    if bg:
        fields.append("userEnteredFormat.backgroundColor")
    if tf:
        fields.append("userEnteredFormat.textFormat")
    if h_align:
        fields.append("userEnteredFormat.horizontalAlignment")
    if wrap:
        fields.append("userEnteredFormat.wrapStrategy")

    return {
        "repeatCell": {
            "range": rng(sheet_id, r1, c1, r2, c2),
            "cell":  {"userEnteredFormat": fmt},
            "fields": ",".join(fields),
        }
    }

# ── COMM100 FETCH ────────────────────────────────────────────
def fetch_all_chats(date_str):
    start_time, end_time = sofia_to_utc_range(date_str)
    auth   = comm100_auth()
    result = []
    seen   = set()
    page   = 1
    base   = (
        f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
        f"?siteId={SITE_ID}"
        f"&include=chatWrapupCategory&include=postChatSurvey&include=chatAgent"
        f"&sortBy=startTime&sortOrder=asc"
    )

    import pytz as _pytz
    _target_tz = _pytz.timezone("Europe/Sofia")

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

            page_in  = 0
            page_out = 0
            for c in chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                ts  = c.get("startTime") or c.get("start_time") or ""
                if not cid or cid in seen:
                    continue
                if in_sofia_range(ts, date_str):
                    seen.add(cid)
                    result.append(c)
                    page_in += 1
                elif ts:
                    # Sonraki güne geçtik mi kontrol et
                    try:
                        dt = datetime.fromisoformat(ts.replace("Z", "+00:00"))
                        if dt.astimezone(_target_tz).strftime("%Y-%m-%d") > date_str:
                            page_out += 1
                    except Exception:
                        pass

            print(f"[FETCH] Sayfa {page}: {len(chats)} chat | bugun:{page_in} dis:{page_out} ({len(result)} toplam)")

            # Sayfanin tamami hedef tarih sonrasindaysa dur
            if page_out > 0 and page_in == 0:
                print("[FETCH] Hedef tarih geçildi, durduruluyor.")
                break
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
    return (c.get("preChatName") or c.get("name") or
            c.get("visitorName") or c.get("visitor_name") or "")

def api_site(c):
    import re
    try:
        url = (c.get("requestingPageURL")
               or (c.get("visitor") or {}).get("currentBrowsing")
               or c.get("requestPage") or "")
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
        dt1  = datetime.fromisoformat(ts1.replace("Z", "+00:00"))
        dt2  = datetime.fromisoformat(ts2.replace("Z", "+00:00"))
        secs = int((dt2 - dt1).total_seconds())
        h, rem = divmod(secs, 3600)
        mn, s  = divmod(rem, 60)
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
    import re
    if not category:
        return ""
    inner = [m.group(1).strip() for m in re.finditer(r"\(([^)]+)\)", category)
             if not re.search(r"VIP", m.group(1), re.I)]
    for t in inner:
        if "_" in t:
            return t
    if inner:
        return inner[0]
    SKIP  = re.compile(r"^(Wild card|Promotions|Deposit|Casino|Sport|General|WD|VIP TIER\d*|VIP)$", re.I)
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
    import re
    if not category:
        return ""
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

# ── FİLTRE ───────────────────────────────────────────────────
def filter_rating_chats(chats, date_str):
    seen, result = set(), []
    for c in chats:
        cid = str(c.get("id") or c.get("chatId") or "")
        if not cid or cid in seen:
            continue
        seen.add(cid)
        if not in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
            continue
        if not api_agent(c).strip():
            continue
        s   = c.get("postChatSurvey") or {}
        g   = s.get("ratingGrade")
        g   = int(g) if g is not None else None
        cmt = (s.get("ratingComment") or "").strip()
        if (g is not None and 1 <= g <= 3) or cmt:
            result.append(c)
    print(f"[FILTER] Rating: {len(result)} chat")
    return result

def filter_tag_chats(chats, date_str, exclude_ids=None):
    excl = set(exclude_ids or [])
    seen, result = set(), []
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

# ── SHEET YAZ: ortak header format ───────────────────────────
def _header_fmt_reqs(sheet_id, n_cols, title_bg, title_fg, header_bg, header_fg,
                     col_widths, frozen_rows=2):
    reqs = []

    # Başlık satırı — merge YOK, sadece renk + center
    # (merge önceki çalışmadan kalmış olabilir, unmerge clear_sheet'te yapılıyor)
    reqs.append(repeat_cell(sheet_id, 0, 0, 0, n_cols - 1,
                             bg=title_bg, fg=title_fg, bold=True,
                             font_size=11, h_align="CENTER"))

    # Kolon başlıkları
    reqs.append(repeat_cell(sheet_id, 1, 0, 1, n_cols - 1,
                             bg=header_bg, fg=header_fg, bold=True, font_size=10))

    # Freeze
    reqs.append({
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id,
                           "gridProperties": {"frozenRowCount": frozen_rows}},
            "fields": "gridProperties.frozenRowCount"
        }
    })

    # Sütun genişlikleri
    for i, w in enumerate(col_widths):
        reqs.append({
            "updateDimensionProperties": {
                "range": {"sheetId": sheet_id, "dimension": "COLUMNS",
                          "startIndex": i, "endIndex": i + 1},
                "properties": {"pixelSize": w},
                "fields": "pixelSize"
            }
        })

    return reqs

def _row_fmt_reqs(sheet_id, chats, data_start_row, include_tag_missing=False):
    """
    Tüm veri satırları için renk request'lerini üret.
    Satır başına ayrı request yerine renk gruplarına göre toplu.
    """
    reqs = []

    # Önce tümüne zebra arka plan (tek tek yerine grup bazlı)
    even_rows, odd_rows, vip_rows = [], [], []
    for idx, c in enumerate(chats):
        row_i = data_start_row + idx
        vip   = extract_vip_tag(api_cat(c))
        if vip:
            vip_rows.append(row_i)
        elif idx % 2 == 0:
            even_rows.append(row_i)
        else:
            odd_rows.append(row_i)

    # Toplu zebra — ardışık aralıkları birleştir
    for rows, bg in [(even_rows, "#f8f9fa"), (odd_rows, "#ffffff"), (vip_rows, "#fff8e1")]:
        for row_i in rows:
            reqs.append(repeat_cell(sheet_id, row_i, 0, row_i, 9, bg=bg))

    # Hücre bazlı özel renkler (VIP, rating, tag eksik)
    for idx, c in enumerate(chats):
        row_i  = data_start_row + idx
        cat    = api_cat(c)
        vip    = extract_vip_tag(cat)
        tag    = extract_main_tag(cat)
        s      = c.get("postChatSurvey") or {}
        g      = s.get("ratingGrade")
        rating = int(g) if g is not None else None

        if vip:
            reqs.append(repeat_cell(sheet_id, row_i, 4, row_i, 4,
                                    bg="#ffd666", bold=True))

        if include_tag_missing and tag == "⚠️ TAG EKSİK":
            reqs.append(repeat_cell(sheet_id, row_i, 3, row_i, 3,
                                    bg="#f28b82", fg="#7c0000", bold=True))

        if rating and rating in RATING_COLORS:
            rc = RATING_COLORS[rating]
            reqs.append(repeat_cell(sheet_id, row_i, 7, row_i, 7,
                                    bg=rc["bg"], fg=rc["fg"], bold=True))

    return reqs

# ── SHEETS YAZ: Rating & Comment ─────────────────────────────
def write_rating_sheet(svc, chats, date_str):
    print(f"[SHEET] Rating & Comment yazılıyor ({len(chats)} chat)...")
    sheet_id = ensure_sheet(svc, RATING_SHEET)
    clear_sheet(svc, sheet_id, RATING_SHEET)

    header_row  = [["⭐ Rating 1-2-3 & Commentli Chatler   |   " + date_str + " → " + date_str]]
    col_headers = [["Temsilci İsmi", "Visitor İsmi", "Sohbet Tarihi", "Kullanılan TAG",
                    "VIP Tag", "Site Adı", "Chat Süresi", "Rating Score",
                    "User Comment", "Chat Linki"]]
    rows = []
    for c in chats:
        s       = c.get("postChatSurvey") or {}
        g       = s.get("ratingGrade")
        cid = str(c.get("id") or c.get("chatId") or "")
        rows.append([
            api_agent(c),
            api_visitor(c),
            format_sofia(c.get("startTime") or c.get("start_time") or ""),
            extract_main_tag(api_cat(c)),
            extract_vip_tag(api_cat(c)),
            api_site(c),
            api_duration(c),
            int(g) if g is not None else "",
            (s.get("ratingComment") or "").strip(),
            "",  # link sütunu — apply_hyperlinks ile doldurulacak
        ])

    write_values(svc, RATING_SHEET, header_row + col_headers + rows)

    links = [
        f"{PORTAL_BASE}?chatId={str(c.get('id') or c.get('chatId') or '')}"
        if (c.get('id') or c.get('chatId')) else ""
        for c in chats
    ]
    fmt_reqs = _header_fmt_reqs(
        sheet_id, n_cols=10,
        title_bg="#1a1a2e", title_fg="#f0c040",
        header_bg="#1a73e8", header_fg="#ffffff",
        col_widths=[155, 130, 160, 180, 110, 160, 90, 80, 280, 200],
    )
    fmt_reqs += _row_fmt_reqs(sheet_id, chats, data_start_row=2,
                               include_tag_missing=False)
    batch_format(svc, fmt_reqs)
    apply_hyperlinks(svc, sheet_id, links, data_start_row=2)
    print(f"[SHEET] Rating & Comment ✅ ({len(chats)} satır)")

# ── SHEETS YAZ: Tag Listesi ───────────────────────────────────
def write_tag_sheet(svc, chats, date_str):
    print(f"[SHEET] Tag Listesi yazılıyor ({len(chats)} chat)...")
    sheet_id = ensure_sheet(svc, TAG_SHEET)
    clear_sheet(svc, sheet_id, TAG_SHEET)

    header_row  = [[f"🏷️ QA Tag Listesi   |   {date_str} → {date_str}   |   {len(chats)} chat"]]
    col_headers = [["Temsilci İsmi", "Visitor İsmi", "Sohbet Tarihi", "Kullanılan Tag",
                    "VIP Tag", "Site Adı", "Chat Süresi", "Rating Score",
                    "User Comment", "Chat Linki"]]
    rows = []
    for c in chats:
        s   = c.get("postChatSurvey") or {}
        g   = s.get("ratingGrade")
        cid = str(c.get("id") or c.get("chatId") or "")
        rows.append([
            api_agent(c),
            api_visitor(c),
            format_sofia(c.get("startTime") or c.get("start_time") or ""),
            extract_main_tag(api_cat(c)),
            extract_vip_tag(api_cat(c)),
            api_site(c),
            api_duration(c),
            int(g) if g is not None else "",
            (s.get("ratingComment") or "").strip(),
            "",  # link sütunu — apply_hyperlinks ile doldurulacak
        ])

    write_values(svc, TAG_SHEET, header_row + col_headers + rows)

    links = [
        f"{PORTAL_BASE}?chatId={str(c.get('id') or c.get('chatId') or '')}"
        if (c.get('id') or c.get('chatId')) else ""
        for c in chats
    ]
    fmt_reqs = _header_fmt_reqs(
        sheet_id, n_cols=10,
        title_bg="#1a1a2e", title_fg="#f0c040",
        header_bg="#1a73e8", header_fg="#ffffff",
        col_widths=[160, 140, 160, 180, 110, 160, 100, 80, 250, 200],
    )
    fmt_reqs += _row_fmt_reqs(sheet_id, chats, data_start_row=2,
                               include_tag_missing=True)
    batch_format(svc, fmt_reqs)
    apply_hyperlinks(svc, sheet_id, links, data_start_row=2)
    print(f"[SHEET] Tag Listesi ✅ ({len(chats)} satır)")

# ── SHEETS YAZ: Log ───────────────────────────────────────────
def write_log(svc, log_type, date_str, duration_s, detail):
    sheet_id = ensure_sheet(svc, LOG_SHEET)

    existing = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A1"
    ).execute().get("values", [])

    if not existing:
        write_values(svc, LOG_SHEET,
                     [["Zaman", "Tür", "Tarih", "Süre (sn)", "Detay"]])
        batch_format(svc, [repeat_cell(sheet_id, 0, 0, 0, 4,
                                       bg="#1a1a2e", fg="#f0c040", bold=True)])

    now_str = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
    svc.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{LOG_SHEET}'!A1",
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body={"values": [[now_str, log_type, date_str, str(duration_s), str(detail)]]}
    ).execute()

    data     = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=f"'{LOG_SHEET}'!A:A"
    ).execute().get("values", [])
    last_row = len(data) - 1

    is_error = "ERROR" in log_type
    is_warn  = "WARN"  in log_type
    bg = "#f28b82" if is_error else "#fff3cd" if is_warn else (
         "#f8f9fa" if last_row % 2 == 0 else "#ffffff")
    fg = "#7c0000" if is_error else "#856404" if is_warn else "#000000"

    batch_format(svc, [repeat_cell(sheet_id, last_row, 0, last_row, 4,
                                   bg=bg, fg=fg,
                                   bold=(is_error or is_warn), wrap=True)])

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

        all_chats = fetch_all_chats(date_str)
        if not all_chats:
            print("[MAIN] Chat bulunamadı, çıkılıyor.")
            write_log(svc, "INFO", date_str, "0", "Hiç chat bulunamadı")
            return

        rating_chats = filter_rating_chats(all_chats, date_str)
        rating_ids   = {str(c.get("id") or c.get("chatId") or "") for c in rating_chats}
        write_rating_sheet(svc, rating_chats, date_str)

        tag_chats = filter_tag_chats(all_chats, date_str, exclude_ids=rating_ids)
        write_tag_sheet(svc, tag_chats, date_str)

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
        try:
            svc = sheets_service()
            write_log(svc, "COMMQA ERROR", date_str, elapsed, f"❌ {e}\n{tb_str}")
        except Exception as log_err:
            print(f"[LOG] Log yazılamadı: {log_err}")
        send_error_email(date_str, str(e), tb_str)
        sys.exit(1)

if __name__ == "__main__":
    main()
