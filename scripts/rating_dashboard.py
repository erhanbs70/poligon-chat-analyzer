#!/usr/bin/env python3
# ============================================================
# POLIGON RATING DASHBOARD — v1 (GitHub Actions)
# Günlük: Rating 1 & 2 pivot — tag + agent breakdown
# Haftalık: 7 günlük özet + trend — her Pazartesi
# Google Sheets API ile yazar — GAS timeout yok
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
SPREADSHEET_ID = os.environ["RATING_SPREADSHEET_ID"]
GCP_CREDS_JSON = os.environ["GCP_CREDENTIALS"]
GMAIL_USER     = os.environ["GMAIL_USER"]
GMAIL_PASS     = os.environ["GMAIL_APP_PASSWORD"]
REPORT_EMAILS  = [e.strip() for e in os.environ["REPORT_EMAILS"].split(",")]

PORTAL_BASE = (
    "https://dash15.lively-chat.com/ui/90005373"
    "/livechat/history/chats/transcriptdetail"
)

EN_MONTHS = ["January","February","March","April","May","June",
             "July","August","September","October","November","December"]

# ── AUTH ─────────────────────────────────────────────────────
def comm100_auth():
    return "Basic " + base64.b64encode(
        f"{COMM100_EMAIL}:{API_KEY}".encode()
    ).decode()

# ── DATE ─────────────────────────────────────────────────────
def sofia_yesterday():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return (datetime.now(tz) - timedelta(days=1)).strftime("%Y-%m-%d")

def sofia_today():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return datetime.now(tz).strftime("%Y-%m-%d")

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

def sheet_name_for(date_str):
    """2026-06-22 → '22 June'"""
    dt = datetime.strptime(date_str, "%Y-%m-%d")
    return f"{dt.day} {EN_MONTHS[dt.month - 1]}"

def week_range_back(n_weeks=0):
    """
    n_weeks=0 → geçen hafta (Pzt-Paz)
    n_weeks=1 → iki hafta önce
    """
    import pytz
    tz  = pytz.timezone("Europe/Sofia")
    now = datetime.now(tz)
    # Bu haftanın Pazartesi'si
    this_monday = now - timedelta(days=now.weekday())
    end_dt   = this_monday - timedelta(days=1 + n_weeks * 7)          # önceki Pazar
    start_dt = end_dt - timedelta(days=6)                              # önceki Pazartesi
    return start_dt.strftime("%Y-%m-%d"), end_dt.strftime("%Y-%m-%d")

def is_monday():
    import pytz
    tz = pytz.timezone("Europe/Sofia")
    return datetime.now(tz).weekday() == 0  # 0 = Pazartesi

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

def ensure_sheet(svc, sheet_name, index=None):
    sid = get_sheet_id(svc, sheet_name)
    if sid is not None:
        return sid
    props = {"title": sheet_name}
    if index is not None:
        props["index"] = index
    body = {"requests": [{"addSheet": {"properties": props}}]}
    resp = svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID, body=body
    ).execute()
    return resp["replies"][0]["addSheet"]["properties"]["sheetId"]

def delete_sheet_if_exists(svc, sheet_name):
    sid = get_sheet_id(svc, sheet_name)
    if sid is None:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"deleteSheet": {"sheetId": sid}}]}
    ).execute()

def write_values(svc, sheet_name, values, start="A1"):
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start}",
        valueInputOption="RAW",
        body={"values": values},
    ).execute()

def batch_fmt(svc, reqs):
    if not reqs:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": reqs}
    ).execute()

# ── RENK ─────────────────────────────────────────────────────
def rgb(hex_str):
    h = hex_str.lstrip("#")
    return {"red": int(h[0:2],16)/255, "green": int(h[2:4],16)/255, "blue": int(h[4:6],16)/255}

def rng(sid, r1, c1, r2, c2):
    return {"sheetId": sid,
            "startRowIndex": r1, "endRowIndex": r2+1,
            "startColumnIndex": c1, "endColumnIndex": c2+1}

def rcell(sid, r1, c1, r2, c2, bg=None, fg=None, bold=False,
          fs=None, halign=None, wrap=False, italic=False):
    fmt, tf, fields = {}, {}, []
    if bg:
        fmt["backgroundColor"] = rgb(bg); fields.append("userEnteredFormat.backgroundColor")
    if bold:
        tf["bold"] = True
    if fg:
        tf["foregroundColor"] = rgb(fg)
    if fs:
        tf["fontSize"] = fs
    if italic:
        tf["italic"] = True
    if tf:
        fmt["textFormat"] = tf; fields.append("userEnteredFormat.textFormat")
    if halign:
        fmt["horizontalAlignment"] = halign; fields.append("userEnteredFormat.horizontalAlignment")
    if wrap:
        fmt["wrapStrategy"] = "WRAP"; fields.append("userEnteredFormat.wrapStrategy")
    return {"repeatCell": {
        "range": rng(sid, r1, c1, r2, c2),
        "cell":  {"userEnteredFormat": fmt},
        "fields": ",".join(fields)
    }}

def col_width(sid, col_idx, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS",
                  "startIndex": col_idx, "endIndex": col_idx+1},
        "properties": {"pixelSize": px},
        "fields": "pixelSize"
    }}

def row_height(sid, row_idx, px):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "ROWS",
                  "startIndex": row_idx, "endIndex": row_idx+1},
        "properties": {"pixelSize": px},
        "fields": "pixelSize"
    }}

def freeze(sid, rows=0, cols=0):
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid,
                       "gridProperties": {"frozenRowCount": rows, "frozenColumnCount": cols}},
        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount"
    }}

def merge(sid, r1, c1, r2, c2):
    return {"mergeCells": {
        "range": rng(sid, r1, c1, r2, c2),
        "mergeType": "MERGE_ALL"
    }}

# ── COMM100 FETCH ─────────────────────────────────────────────
def fetch_low_rating_chats(date_str):
    """Rating 1 ve 2 olan chatleri çek."""
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
    print(f"[FETCH] {date_str} low rating chatleri çekiliyor...")
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
            if len(chats) < 500:
                break
            page += 1
            time.sleep(0.2)
        except Exception as e:
            raise RuntimeError(f"fetch_low_rating sayfa {page}: {e}") from e

    # Filtrele: sadece rating 1 & 2, Sofia tarih aralığında
    filtered = []
    for c in result:
        if not in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
            continue
        s = c.get("postChatSurvey") or {}
        g = s.get("ratingGrade")
        if g is not None and int(g) in (1, 2):
            filtered.append(c)

    print(f"[FETCH] {date_str}: {len(filtered)} low rating chat")
    return filtered

def fetch_all_chats_count(date_str):
    """Toplam chat sayısını çek (agent bazlı hesap için)."""
    start_time, end_time = sofia_to_utc_range(date_str)
    auth   = comm100_auth()
    result = []
    seen   = set()
    page   = 1
    base   = (
        f"https://dash15.lively-chat.com/api/LiveChat/chats:search"
        f"?siteId={SITE_ID}&include=chatAgent"
        f"&sortBy=startTime&sortOrder=asc"
    )
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
                break
            chats = r.json().get("list", [])
            if not chats:
                break
            for c in chats:
                cid = str(c.get("id") or c.get("chatId") or "")
                if cid and cid not in seen:
                    seen.add(cid)
                    if in_sofia_range(c.get("startTime") or c.get("start_time"), date_str):
                        result.append(c)
            if len(chats) < 500:
                break
            page += 1
            time.sleep(0.2)
        except Exception:
            break
    print(f"[FETCH] {date_str}: {len(result)} toplam chat")
    return result

# ── FIELD HELPERS ─────────────────────────────────────────────
def get_agent(c):
    agents = c.get("chatAgents") or []
    if agents:
        names = [
            (a.get("agent") or {}).get("displayName") or a.get("displayName") or a.get("name") or ""
            for a in agents if not a.get("botId")
        ]
        return ", ".join(filter(None, names))
    return c.get("agentName") or c.get("agent_name") or ""

def get_tag(c):
    import re
    cat = (c.get("chatWrapup") or {}).get("categoriesName") or ""
    if not cat:
        return "—"
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
    return parts[0] if parts else "—"

def get_grade(c):
    s = c.get("postChatSurvey") or {}
    g = s.get("ratingGrade")
    return int(g) if g is not None else None

# ── PIVOT HESAPLA ─────────────────────────────────────────────
def build_pivots(low_chats, all_chats):
    # Tag pivot
    tag_map = {}
    for c in low_chats:
        tag = get_tag(c)
        g   = get_grade(c)
        if tag not in tag_map:
            tag_map[tag] = {"r1": 0, "r2": 0}
        if g == 1: tag_map[tag]["r1"] += 1
        elif g == 2: tag_map[tag]["r2"] += 1

    total_low = len(low_chats)
    tag_rows = sorted(
        [[tag, v["r1"], v["r2"], v["r1"]+v["r2"],
          f'{(v["r1"]+v["r2"])/total_low*100:.1f}%' if total_low else "—"]
         for tag, v in tag_map.items()],
        key=lambda x: -x[3]
    )

    # Agent pivot
    agent_total = {}
    for c in all_chats:
        for a in get_agent(c).split(","):
            a = a.strip()
            if a:
                agent_total[a] = agent_total.get(a, 0) + 1

    agent_map = {}
    for c in low_chats:
        g = get_grade(c)
        for a in get_agent(c).split(","):
            a = a.strip()
            if not a:
                continue
            if a not in agent_map:
                agent_map[a] = {"r1": 0, "r2": 0}
            if g == 1: agent_map[a]["r1"] += 1
            elif g == 2: agent_map[a]["r2"] += 1

    agent_rows = []
    for agent, v in agent_map.items():
        tot      = v["r1"] + v["r2"]
        tot_chat = agent_total.get(agent, tot)
        pct      = f"{tot/tot_chat*100:.1f}%" if tot_chat else "—"
        agent_rows.append([agent, v["r1"], v["r2"], tot, tot_chat, pct])
    agent_rows.sort(key=lambda x: -x[3])

    return tag_rows, agent_rows, total_low, len(all_chats)

# ── GÜNLÜK SHEET YAZ ─────────────────────────────────────────
def write_daily_sheet(svc, date_str, tag_rows, agent_rows, total_low, total_all):
    sname = sheet_name_for(date_str)
    print(f"[SHEET] {sname} yazılıyor...")

    delete_sheet_if_exists(svc, sname)
    sid = ensure_sheet(svc, sname, index=0)

    # ── Veri ──
    col_w = [220, 90, 90, 100, 110, 160]
    r1_tot = sum(r[1] for r in tag_rows)
    r2_tot = sum(r[2] for r in tag_rows)

    rows = []
    # Başlık
    rows.append([f"📊 Rating 1 & 2 Dashboard   |   {sname}   |   Low: {total_low} / {total_all} chat", "", "", "", "", ""])
    # Tag bölümü
    rows.append(["🏷️ BREAKDOWN BY TAG", "", "", "", "", ""])
    rows.append(["Tag", "Rating 1 ❌", "Rating 2 🟠", "Toplam", "Pay (%)", ""])
    for tr in tag_rows:
        rows.append(tr + [""])
    rows.append(["TOTAL", r1_tot, r2_tot, r1_tot+r2_tot, "100%", ""])
    rows.append(["", "", "", "", "", ""])
    # Agent bölümü
    rows.append(["👤 BREAKDOWN BY AGENT", "", "", "", "", ""])
    rows.append(["Agent", "Rating 1 ❌", "Rating 2 🟠", "Low Toplam", "Toplam Chat", "% Oranı"])
    for ar in agent_rows:
        rows.append(ar)
    ag_r1 = sum(r[1] for r in agent_rows)
    ag_r2 = sum(r[2] for r in agent_rows)
    ag_tc = sum(r[4] for r in agent_rows)
    rows.append(["TOTAL", ag_r1, ag_r2, ag_r1+ag_r2, ag_tc, ""])

    write_values(svc, sname, rows)

    # ── Format ──
    reqs = []
    reqs.append(freeze(sid, rows=3))
    for i, w in enumerate(col_w):
        reqs.append(col_width(sid, i, w))
    reqs.append(row_height(sid, 0, 36))

    # Başlık satırı
    reqs.append(rcell(sid, 0, 0, 0, 5, bg="#1a1a2e", fg="#f0c040", bold=True, fs=11, halign="CENTER"))

    tag_header_row = 1
    tag_col_row    = 2
    tag_data_start = 3
    tag_data_end   = tag_data_start + len(tag_rows) - 1
    tag_total_row  = tag_data_end + 1
    gap_row        = tag_total_row + 1
    agent_header_row = gap_row + 1
    agent_col_row    = agent_header_row + 1
    agent_data_start = agent_col_row + 1
    agent_data_end   = agent_data_start + len(agent_rows) - 1
    agent_total_row  = agent_data_end + 1

    # Tag bölüm başlıkları
    reqs.append(rcell(sid, tag_header_row, 0, tag_header_row, 5,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, tag_col_row, 0, tag_col_row, 4,
                      bg="#2c5f8a", fg="#ffffff", bold=True, fs=10))

    # Tag veri satırları
    for i in range(len(tag_rows)):
        ri = tag_data_start + i
        bg = "#f0f4ff" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 4, bg=bg))
        reqs.append(rcell(sid, ri, 1, ri, 1, bg="#fde8e8", fg="#7c0000"))
        reqs.append(rcell(sid, ri, 2, ri, 2, bg="#fff3cd", fg="#5a3a00"))
        reqs.append(rcell(sid, ri, 3, ri, 3, bold=True))

    reqs.append(rcell(sid, tag_total_row, 0, tag_total_row, 4,
                      bg="#1a3a5c", fg="#f0c040", bold=True))

    # Agent bölüm başlıkları
    reqs.append(rcell(sid, agent_header_row, 0, agent_header_row, 5,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, agent_col_row, 0, agent_col_row, 5,
                      bg="#2c5f8a", fg="#ffffff", bold=True, fs=10))

    # Agent veri satırları
    for i, ar in enumerate(agent_rows):
        ri  = agent_data_start + i
        bg  = "#f0f4ff" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 5, bg=bg))
        reqs.append(rcell(sid, ri, 1, ri, 1, bg="#fde8e8", fg="#7c0000"))
        reqs.append(rcell(sid, ri, 2, ri, 2, bg="#fff3cd", fg="#5a3a00"))
        reqs.append(rcell(sid, ri, 3, ri, 3, bold=True))
        reqs.append(rcell(sid, ri, 4, ri, 4, fg="#444444"))
        # % oranı rengi
        try:
            pct_val = float(str(ar[5]).replace("%", ""))
            pct_bg  = "#f28b82" if pct_val > 30 else "#fbbc04" if pct_val > 15 else "#fff475" if pct_val > 5 else "#d4edda"
            pct_fg  = "#7c0000" if pct_val > 30 else "#5a3a00" if pct_val > 15 else "#5a4a00" if pct_val > 5 else "#155724"
        except Exception:
            pct_bg, pct_fg = "#ffffff", "#000000"
        reqs.append(rcell(sid, ri, 5, ri, 5, bg=pct_bg, fg=pct_fg, bold=True))

    reqs.append(rcell(sid, agent_total_row, 0, agent_total_row, 5,
                      bg="#1a3a5c", fg="#f0c040", bold=True))

    batch_fmt(svc, reqs)
    print(f"[SHEET] {sname} ✅ (low: {total_low}, all: {total_all})")
    return total_low, total_all

# ── HAFTALIK SHEET YAZ ───────────────────────────────────────
def write_weekly_sheet(svc, week_data, start_str, end_str, prev_week_data=None):
    """
    week_data: list of dict — her gün için
      {date, low_chats, all_count, tag_rows, agent_rows, total_low, total_all}
    prev_week_data: aynı yapı — trend için (None ise trend gösterilmez)
    """
    sname = "📊 Weekly"
    print(f"[SHEET] Weekly sheet yazılıyor ({start_str} - {end_str})...")

    # Sheet varsa sil, yeniden oluştur
    delete_sheet_if_exists(svc, sname)
    sid = ensure_sheet(svc, sname, index=0)

    # ── Haftalık toplam pivotları hesapla ──
    weekly_tag   = {}
    weekly_agent = {}
    weekly_agent_total = {}
    daily_totals = []  # trend için

    for d in week_data:
        daily_totals.append({
            "date":      d["date"],
            "low":       d["total_low"],
            "all":       d["total_all"],
            "pct":       f'{d["total_low"]/d["total_all"]*100:.1f}%' if d["total_all"] else "—",
        })
        for tr in d["tag_rows"]:
            tag = tr[0]
            if tag not in weekly_tag:
                weekly_tag[tag] = {"r1": 0, "r2": 0}
            weekly_tag[tag]["r1"] += tr[1]
            weekly_tag[tag]["r2"] += tr[2]
        for ar in d["agent_rows"]:
            agent = ar[0]
            if agent not in weekly_agent:
                weekly_agent[agent] = {"r1": 0, "r2": 0}
            weekly_agent[agent]["r1"] += ar[1]
            weekly_agent[agent]["r2"] += ar[2]
            weekly_agent_total[agent] = weekly_agent_total.get(agent, 0) + ar[4]

    total_low_week = sum(d["total_low"] for d in week_data)
    total_all_week = sum(d["total_all"] for d in week_data)
    low_pct_week   = f"{total_low_week/total_all_week*100:.1f}%" if total_all_week else "—"

    # Trend vs önceki hafta
    prev_low = sum(d["total_low"] for d in prev_week_data) if prev_week_data else None
    prev_all = sum(d["total_all"] for d in prev_week_data) if prev_week_data else None
    if prev_low is not None and prev_all:
        trend_low = total_low_week - prev_low
        trend_pct_prev = f"{prev_low/prev_all*100:.1f}%"
        trend_str = (f"▲ +{trend_low}" if trend_low > 0 else f"▼ {trend_low}") + f" vs geçen hafta ({prev_low} low)"
    else:
        trend_str      = "—"
        trend_pct_prev = "—"

    # En kötü gün
    worst_day = max(daily_totals, key=lambda x: x["low"]) if daily_totals else None
    # En çok low alan agent
    worst_agent = max(weekly_agent.items(), key=lambda x: x[1]["r1"]+x[1]["r2"]) if weekly_agent else None
    # En sorunlu tag
    worst_tag = max(weekly_tag.items(), key=lambda x: x[1]["r1"]+x[1]["r2"]) if weekly_tag else None

    # Tag sıralama
    weekly_tag_rows = sorted(
        [[tag, v["r1"], v["r2"], v["r1"]+v["r2"],
          f'{(v["r1"]+v["r2"])/total_low_week*100:.1f}%' if total_low_week else "—"]
         for tag, v in weekly_tag.items()],
        key=lambda x: -x[3]
    )
    # Agent sıralama
    weekly_agent_rows = []
    for agent, v in weekly_agent.items():
        tot      = v["r1"] + v["r2"]
        tot_chat = weekly_agent_total.get(agent, tot)
        pct      = f"{tot/tot_chat*100:.1f}%" if tot_chat else "—"
        weekly_agent_rows.append([agent, v["r1"], v["r2"], tot, tot_chat, pct])
    weekly_agent_rows.sort(key=lambda x: -x[3])

    # ── Veri satırları ──
    dt_start = datetime.strptime(start_str, "%Y-%m-%d")
    dt_end   = datetime.strptime(end_str,   "%Y-%m-%d")
    week_label = f"{dt_start.day} {EN_MONTHS[dt_start.month-1]} – {dt_end.day} {EN_MONTHS[dt_end.month-1]}"

    rows = []
    row_meta = {}  # satır adı → index (format için)

    def add(row, label=None):
        idx = len(rows)
        if label:
            row_meta[label] = idx
        rows.append(row)

    NCOLS = 7  # A..G

    # ── 1. ÖZET BÖLÜMÜ ──
    add([f"📊 Rating Dashboard — Haftalık Özet   |   {week_label}", "","","","","",""], "title")
    add(["","","","","","",""])  # boşluk

    add(["📌 HAFTALIK ÖZET", "","","","","",""], "summary_header")
    add(["Metrik", "Değer", "Trend", "","","",""], "summary_col")
    add(["Toplam Low Rating (1+2)", total_low_week, trend_str, "","","",""])
    add(["Toplam Chat", total_all_week, "", "","","",""])
    add(["Low Rating Oranı", low_pct_week,
         f"Geçen hafta: {trend_pct_prev}" if trend_pct_prev != "—" else "—",
         "","","",""])
    if worst_day:
        add(["En Kötü Gün", sheet_name_for(worst_day["date"]),
             f"{worst_day['low']} low / {worst_day['all']} chat ({worst_day['pct']})",
             "","","",""])
    if worst_tag:
        wt_tot = worst_tag[1]["r1"] + worst_tag[1]["r2"]
        add(["En Sorunlu Tag", worst_tag[0], f"{wt_tot} low rating", "","","",""])
    if worst_agent:
        wa_tot = worst_agent[1]["r1"] + worst_agent[1]["r2"]
        add(["En Çok Low Alan Agent", worst_agent[0], f"{wa_tot} low rating", "","","",""])

    add(["","","","","","",""])  # boşluk

    # ── 2. GÜNLÜK DAĞILIM ──
    add(["📅 GÜNLÜK DAĞILIM", "","","","","",""], "daily_header")
    add(["Gün", "Tarih", "Low (1+2)", "Toplam Chat", "Low %", "Rating 1", "Rating 2"], "daily_col")
    row_meta["daily_data_start"] = len(rows)
    for d in daily_totals:
        r1_cnt = sum(c[1] for c in d.get("tag_rows_raw", []))  # fallback
        r2_cnt = sum(c[2] for c in d.get("tag_rows_raw", []))
        # tag_rows_raw yok — week_data'dan çek
        wd = next((x for x in week_data if x["date"] == d["date"]), None)
        if wd:
            r1_cnt = sum(tr[1] for tr in wd["tag_rows"])
            r2_cnt = sum(tr[2] for tr in wd["tag_rows"])
        dt = datetime.strptime(d["date"], "%Y-%m-%d")
        day_name = ["Pzt","Sal","Çar","Per","Cum","Cmt","Paz"][dt.weekday()]
        rows.append([day_name, sheet_name_for(d["date"]), d["low"], d["all"], d["pct"], r1_cnt, r2_cnt])
    row_meta["daily_data_end"] = len(rows) - 1

    # Haftalık toplam
    total_r1 = sum(tr[1] for tr in weekly_tag_rows)
    total_r2 = sum(tr[2] for tr in weekly_tag_rows)
    rows.append(["TOPLAM", "", total_low_week, total_all_week, low_pct_week, total_r1, total_r2])
    row_meta["daily_total"] = len(rows) - 1

    add(["","","","","","",""])

    # ── 3. TAG PIVOT ──
    add(["🏷️ TAG BREAKDOWN — HAFTALIK", "","","","","",""], "tag_header")
    add(["Tag", "Rating 1 ❌", "Rating 2 🟠", "Toplam", "Pay (%)", "", ""], "tag_col")
    row_meta["tag_data_start"] = len(rows)
    for tr in weekly_tag_rows:
        rows.append(tr + ["", ""])
    row_meta["tag_data_end"] = len(rows) - 1
    wtr1 = sum(r[1] for r in weekly_tag_rows)
    wtr2 = sum(r[2] for r in weekly_tag_rows)
    rows.append(["TOTAL", wtr1, wtr2, wtr1+wtr2, "100%", "", ""])
    row_meta["tag_total"] = len(rows) - 1

    add(["","","","","","",""])

    # ── 4. AGENT PIVOT ──
    add(["👤 AGENT BREAKDOWN — HAFTALIK", "","","","","",""], "agent_header")
    add(["Agent", "Rating 1 ❌", "Rating 2 🟠", "Low Toplam", "Toplam Chat", "% Oranı", ""], "agent_col")
    row_meta["agent_data_start"] = len(rows)
    for ar in weekly_agent_rows:
        rows.append(ar + [""])
    row_meta["agent_data_end"] = len(rows) - 1
    war1 = sum(r[1] for r in weekly_agent_rows)
    war2 = sum(r[2] for r in weekly_agent_rows)
    watc = sum(r[4] for r in weekly_agent_rows)
    rows.append(["TOTAL", war1, war2, war1+war2, watc, "", ""])
    row_meta["agent_total"] = len(rows) - 1

    write_values(svc, sname, rows)

    # ── Format ──
    reqs = []
    reqs.append(freeze(sid, rows=2))

    col_widths = [160, 140, 110, 120, 100, 100, 100]
    for i, w in enumerate(col_widths):
        reqs.append(col_width(sid, i, w))
    reqs.append(row_height(sid, row_meta["title"], 40))

    # Başlık
    reqs.append(rcell(sid, row_meta["title"], 0, row_meta["title"], 6,
                      bg="#1a1a2e", fg="#f0c040", bold=True, fs=12, halign="CENTER"))

    # Özet bölümü
    reqs.append(rcell(sid, row_meta["summary_header"], 0, row_meta["summary_header"], 6,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, row_meta["summary_col"], 0, row_meta["summary_col"], 2,
                      bg="#2c5f8a", fg="#ffffff", bold=True))

    # Özet veri satırları — zebra
    summary_data_start = row_meta["summary_col"] + 1
    summary_data_end   = row_meta["daily_header"] - 2
    for i in range(summary_data_end - summary_data_start + 1):
        ri = summary_data_start + i
        bg = "#f8f9fa" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 0, bold=True, fg="#1a3a5c"))
        reqs.append(rcell(sid, ri, 1, ri, 2, bg=bg))

    # Günlük dağılım
    reqs.append(rcell(sid, row_meta["daily_header"], 0, row_meta["daily_header"], 6,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, row_meta["daily_col"], 0, row_meta["daily_col"], 6,
                      bg="#2c5f8a", fg="#ffffff", bold=True))

    for i in range(row_meta["daily_data_end"] - row_meta["daily_data_start"] + 1):
        ri  = row_meta["daily_data_start"] + i
        bg  = "#f0f4ff" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 6, bg=bg))
        reqs.append(rcell(sid, ri, 2, ri, 2, bold=True))
        # Low % rengi
        wd = week_data[i] if i < len(week_data) else None
        if wd and wd["total_all"]:
            pv = wd["total_low"] / wd["total_all"] * 100
            pb = "#f28b82" if pv > 10 else "#fff475" if pv > 5 else "#d4edda"
            pf = "#7c0000" if pv > 10 else "#5a4a00" if pv > 5 else "#155724"
            reqs.append(rcell(sid, ri, 4, ri, 4, bg=pb, fg=pf, bold=True))

    reqs.append(rcell(sid, row_meta["daily_total"], 0, row_meta["daily_total"], 6,
                      bg="#1a3a5c", fg="#f0c040", bold=True))

    # Tag pivot
    reqs.append(rcell(sid, row_meta["tag_header"], 0, row_meta["tag_header"], 6,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, row_meta["tag_col"], 0, row_meta["tag_col"], 4,
                      bg="#2c5f8a", fg="#ffffff", bold=True))

    for i in range(row_meta["tag_data_end"] - row_meta["tag_data_start"] + 1):
        ri = row_meta["tag_data_start"] + i
        bg = "#f0f4ff" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 4, bg=bg))
        reqs.append(rcell(sid, ri, 1, ri, 1, bg="#fde8e8", fg="#7c0000"))
        reqs.append(rcell(sid, ri, 2, ri, 2, bg="#fff3cd", fg="#5a3a00"))
        reqs.append(rcell(sid, ri, 3, ri, 3, bold=True))

    reqs.append(rcell(sid, row_meta["tag_total"], 0, row_meta["tag_total"], 4,
                      bg="#1a3a5c", fg="#f0c040", bold=True))

    # Agent pivot
    reqs.append(rcell(sid, row_meta["agent_header"], 0, row_meta["agent_header"], 6,
                      bg="#1a3a5c", fg="#ffffff", bold=True, fs=10))
    reqs.append(rcell(sid, row_meta["agent_col"], 0, row_meta["agent_col"], 5,
                      bg="#2c5f8a", fg="#ffffff", bold=True))

    for i, ar in enumerate(weekly_agent_rows):
        ri  = row_meta["agent_data_start"] + i
        bg  = "#f0f4ff" if i % 2 == 0 else "#ffffff"
        reqs.append(rcell(sid, ri, 0, ri, 5, bg=bg))
        reqs.append(rcell(sid, ri, 1, ri, 1, bg="#fde8e8", fg="#7c0000"))
        reqs.append(rcell(sid, ri, 2, ri, 2, bg="#fff3cd", fg="#5a3a00"))
        reqs.append(rcell(sid, ri, 3, ri, 3, bold=True))
        try:
            pv  = float(str(ar[5]).replace("%",""))
            pb  = "#f28b82" if pv>30 else "#fbbc04" if pv>15 else "#fff475" if pv>5 else "#d4edda"
            pf  = "#7c0000" if pv>30 else "#5a3a00" if pv>15 else "#5a4a00" if pv>5 else "#155724"
        except Exception:
            pb, pf = "#ffffff", "#000000"
        reqs.append(rcell(sid, ri, 5, ri, 5, bg=pb, fg=pf, bold=True))

    reqs.append(rcell(sid, row_meta["agent_total"], 0, row_meta["agent_total"], 5,
                      bg="#1a3a5c", fg="#f0c040", bold=True))

    batch_fmt(svc, reqs)
    print(f"[SHEET] Weekly ✅ ({start_str} - {end_str})")

# ── HATA MAİLİ ───────────────────────────────────────────────
def send_error_email(date_str, error_msg, tb_str):
    subject = f"⚠️ Rating Dashboard HATA | {date_str}"
    html = f"""<!DOCTYPE html><html><body style="font-family:monospace;padding:20px;">
<h2 style="color:#c0392b;">Rating Dashboard — Hata Raporu</h2>
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
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(GMAIL_USER, GMAIL_PASS)
            s.sendmail(GMAIL_USER, REPORT_EMAILS, msg.as_string())
        print(f"[EMAIL] Hata maili gönderildi")
    except Exception as e:
        print(f"[EMAIL] Gönderilemedi: {e}")

# ── MAIN ─────────────────────────────────────────────────────
def main():
    t0       = time.time()
    date_str = sofia_yesterday()
    mode     = sys.argv[1] if len(sys.argv) > 1 else "daily"
    print(f"\n{'='*55}\nRATING DASHBOARD — {date_str} | mode={mode}\n{'='*55}\n")

    try:
        svc = sheets_service()

        # ── Günlük ──
        low_chats = fetch_low_rating_chats(date_str)
        all_chats = fetch_all_chats_count(date_str)
        tag_rows, agent_rows, total_low, total_all = build_pivots(low_chats, all_chats)

        if total_low > 0:
            write_daily_sheet(svc, date_str, tag_rows, agent_rows, total_low, total_all)
        else:
            print(f"[MAIN] {date_str} için low rating yok, sheet yazılmadı.")

        # ── Haftalık — sadece Pazartesi ──
        if is_monday() or mode == "weekly":
            print("[MAIN] Pazartesi — haftalık rapor hazırlanıyor...")
            start_str, end_str = week_range_back(0)

            # Geçen haftanın her günü için veri çek
            week_data   = []
            prev_totals = []

            cur = datetime.strptime(start_str, "%Y-%m-%d")
            end = datetime.strptime(end_str,   "%Y-%m-%d")
            while cur <= end:
                ds = cur.strftime("%Y-%m-%d")
                try:
                    lc = fetch_low_rating_chats(ds)
                    ac = fetch_all_chats_count(ds)
                    tr, ar, tl, ta = build_pivots(lc, ac)
                    week_data.append({
                        "date": ds, "low_chats": lc, "all_count": ac,
                        "tag_rows": tr, "agent_rows": ar,
                        "total_low": tl, "total_all": ta,
                    })
                except Exception as e:
                    print(f"[WEEKLY] {ds} atlandı: {e}")
                    week_data.append({
                        "date": ds, "low_chats": [], "all_count": [],
                        "tag_rows": [], "agent_rows": [],
                        "total_low": 0, "total_all": 0,
                    })
                cur += timedelta(days=1)
                time.sleep(0.5)

            # Önceki hafta toplam low sayısı (trend için)
            prev_start, prev_end = week_range_back(1)
            prev_week_data = []
            cur = datetime.strptime(prev_start, "%Y-%m-%d")
            pend = datetime.strptime(prev_end, "%Y-%m-%d")
            while cur <= pend:
                ds = cur.strftime("%Y-%m-%d")
                try:
                    lc = fetch_low_rating_chats(ds)
                    ac = fetch_all_chats_count(ds)
                    tr, ar, tl, ta = build_pivots(lc, ac)
                    prev_week_data.append({
                        "date": ds, "total_low": tl, "total_all": ta,
                        "tag_rows": tr, "agent_rows": ar,
                    })
                except Exception:
                    prev_week_data.append({"date": ds, "total_low": 0, "total_all": 0,
                                           "tag_rows": [], "agent_rows": []})
                cur += timedelta(days=1)
                time.sleep(0.3)

            write_weekly_sheet(svc, week_data, start_str, end_str, prev_week_data)

        elapsed = round(time.time() - t0, 1)
        print(f"\n[DONE] Tamamlandı — {elapsed}s")

    except Exception as e:
        elapsed = round(time.time() - t0, 1)
        tb_str  = traceback.format_exc()
        print(f"[ERROR] {e}\n{tb_str}")
        send_error_email(date_str, str(e), tb_str)
        sys.exit(1)

if __name__ == "__main__":
    main()
