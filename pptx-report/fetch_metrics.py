"""
fetch_metrics.py
Python port of the Poligon Daily Chat Report (v8/v9 GAS script) metrics
pipeline. Talks to the same Comm100 endpoints, produces the same shape of
data, but without Apps Script's 6-minute execution limit.

Credentials/config are read from environment variables so this can run
safely inside GitHub Actions (see .github/workflows/daily-pptx.yml).
"""
import os
import time
import json
import datetime as dt
from concurrent.futures import ThreadPoolExecutor
import requests
import base64

# ============================================================
# CONFIG
# ============================================================
SITE_ID = os.environ.get("MR_SITE_ID", "90005373")
API_KEY = os.environ.get("MR_API_KEY", "")
EMAIL = os.environ.get("MR_EMAIL", "erhan@premilogic.com")
CS_DEPT = os.environ.get("MR_CS_DEPT", "29098b6e-0acb-40cc-8cd7-ebf662defc42")

CAMPAIGNS = {
    "superbetin": "3cd6532a-0d60-4d7d-9331-5dd5be71f527",
    "betsat": "2b71ad44-6a12-4073-bc6c-8edb22eabcc6",
    "turkbet": "6712311a-2268-408c-a803-b338a1308010",
}
BRANDS = ["superbetin", "betsat", "turkbet"]
TIER_SEGMENT_NAMES = ["VIP", "VIP2", "VIP3"]

BASE = f"https://dash15.lively-chat.com/api"
SEARCH_URL = f"{BASE}/LiveChat/chats:search?siteId={SITE_ID}"
REPORT_URL = f"{BASE}/reportingquery/reports/query?siteId={SITE_ID}"
TZ = "FLE Standard Time"


def _auth_header():
    token = base64.b64encode(f"{EMAIL}:{API_KEY}".encode()).decode()
    return {"Authorization": f"Basic {token}", "Content-Type": "application/json"}


# ============================================================
# TIME / TIMEZONE HELPERS (Europe/Sofia, DST-aware, mirrors GAS logic)
# ============================================================
def _sofia_offset_minutes(date_str):
    """Returns -180 (summer, UTC+3) or -120 (winter, UTC+2) matching the
    original GAS _mrSofiaTimezone / offsetHours convention (kept as sign
    for parity, only the magnitude is used below)."""
    d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    year = d.year
    last_march = dt.date(year, 3, 31)
    last_march -= dt.timedelta(days=(last_march.isoweekday() % 7))
    last_oct = dt.date(year, 10, 31)
    last_oct -= dt.timedelta(days=(last_oct.isoweekday() % 7))
    return -180 if (last_march <= d.date() < last_oct) else -120


def _offset_hours(date_str):
    return 3 if _sofia_offset_minutes(date_str) == -180 else 2


def sofia_to_utc_range(date_from, date_to):
    off = _offset_hours(date_from)
    start = dt.datetime.strptime(date_from + "T00:00:00", "%Y-%m-%dT%H:%M:%S") - dt.timedelta(hours=off)
    end = dt.datetime.strptime(date_to + "T23:59:59", "%Y-%m-%dT%H:%M:%S") - dt.timedelta(hours=off)
    return start.strftime("%Y-%m-%dT%H:%M:%S") + "Z", end.strftime("%Y-%m-%dT%H:%M:%S") + "Z"


def _parse_chat_ts(ts):
    if not ts:
        return None
    if isinstance(ts, str) and "/Date(" in ts:
        raw = int(ts.replace("/Date(", "").replace(")/", ""))
        return dt.datetime.utcfromtimestamp(raw / 1000)
    # ISO 8601, e.g. "2026-07-07T10:30:41.05Z"
    ts = ts.replace("Z", "")
    if "." in ts:
        ts = ts.split(".")[0]
    return dt.datetime.strptime(ts, "%Y-%m-%dT%H:%M:%S")


def in_sofia_range(ts, date_from, date_to):
    parsed = _parse_chat_ts(ts)
    if parsed is None:
        return True
    off = _offset_hours(date_from)
    local = parsed + dt.timedelta(hours=off)
    d = local.strftime("%Y-%m-%d")
    return date_from <= d <= date_to


def sofia_hour(ts, date_str):
    parsed = _parse_chat_ts(ts)
    if parsed is None:
        return -1
    off = _offset_hours(date_str)
    local = parsed + dt.timedelta(hours=off)
    return local.hour


# ============================================================
# SMALL HELPERS (agent name, tag parsing, hms parsing) — mirror GAS 1:1
# ============================================================
def agent_name(c):
    agents = c.get("chatAgents") or []
    names = []
    for a in agents:
        if a.get("botId"):
            continue
        n = (a.get("agent") or {}).get("displayName") or a.get("displayName") or a.get("name") or ""
        if n:
            names.append(n)
    if names:
        return ", ".join(names)
    return c.get("agentName") or ""


def tag_from_wrapup(c):
    import re
    cat = (c.get("chatWrapup") or {}).get("categoriesName") or ""
    if not cat:
        return "—"
    inner = []
    for m in re.finditer(r"\(([^)]+)\)", cat):
        t = m.group(1).strip()
        if "vip" not in t.lower():
            inner.append(t)
    for t in inner:
        if "_" in t:
            return t
    if inner:
        return inner[0]
    parts = [p.strip() for p in re.split(r"[,\n\r]+", re.sub(r"\([^)]*\)", "", cat)) if p.strip()]
    for p in parts:
        if "_" in p:
            return p
    return parts[0] if parts else "—"


def parse_hms(s):
    if not s or not isinstance(s, str):
        return 0
    parts = s.split(":")
    if len(parts) < 2:
        return 0
    h = int(parts[0] or 0)
    m = int(parts[1] or 0)
    sec = int(parts[2] or 0) if len(parts) > 2 else 0
    return h * 3600 + m * 60 + sec


def brand_from_url(url):
    if not url:
        return "other"
    u = url.lower()
    if "superbetin" in u:
        return "superbetin"
    if "betsat" in u:
        return "betsat"
    if "turkbet" in u:
        return "turkbet"
    return "other"


def campaign_map():
    return {v.lower(): k for k, v in CAMPAIGNS.items()}


def brand_of(c, camp_map):
    cid = (c.get("campaignId") or c.get("campaign_id") or c.get("campaignID") or "").lower()
    if cid and cid in camp_map:
        return camp_map[cid]
    return brand_from_url(c.get("requestingPageURL") or "")


# ============================================================
# RAW CHAT HISTORY (paginated, mirrors _mrFetchAll)
# ============================================================
def fetch_all_chats(date_from, date_to):
    start, end = sofia_to_utc_range(date_from, date_to)
    headers = _auth_header()
    result = []
    page = 1
    while True:
        url = (
            f"{SEARCH_URL}&pageIndex={page}&pageSize=500"
            "&include=chatAgent&include=chatWrapup&include=postChatSurvey"
            "&include=customVariable&sortBy=startTime&sortOrder=asc"
        )
        try:
            r = requests.post(url, headers=headers, json={"startTime": start, "endTime": end}, timeout=60)
        except requests.RequestException:
            break
        if r.status_code != 200:
            break
        chats = (r.json() or {}).get("list") or []
        if not chats:
            break
        result.extend(chats)
        if len(chats) < 500 or page >= 40:
            break
        page += 1
        time.sleep(0.3)

    seen = set()
    out = []
    for c in result:
        cid = str(c.get("id") or c.get("chatId") or "")
        if not cid or cid in seen:
            continue
        seen.add(cid)
        if in_sofia_range(c.get("startTime"), date_from, date_to):
            out.append(c)
    return out


# ============================================================
# REPORTING API — generic POST helper + parallel batch
# ============================================================
def _report_query(payload):
    r = requests.post(REPORT_URL, headers=_auth_header(), json=payload, timeout=60)
    if r.status_code != 200:
        return {}
    try:
        return r.json()
    except ValueError:
        return {}


def _report_query_total(payload):
    return (_report_query(payload) or {}).get("total") or {}


def _parallel(payloads):
    """Runs a list of report payloads concurrently, returns list of raw
    json responses in the same order (mirrors UrlFetchApp.fetchAll)."""
    with ThreadPoolExecutor(max_workers=min(12, len(payloads))) as ex:
        return list(ex.map(_report_query, payloads))


# ============================================================
# BRAND VOLUMES + EFFICIENCY (6 parallel requests)
# ============================================================
def fetch_brand_volumes_and_visits(date_str):
    d = date_str.replace("-", "/")
    payloads = []
    for brand in BRANDS:
        camp_id = CAMPAIGNS[brand]
        payloads.append({
            "cubeEntities": [
                {
                    "name": "Chat",
                    "fields": [
                        {"name": "chats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "Status0 & Duration>0", "conditionMatchType": "all",
                         "conditions": [
                             {"name": "Status0", "fieldName": "Status", "operate": "equals", "values": ["0"]},
                             {"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]},
                         ]},
                        {"name": "missedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "missedChatStatus", "conditionMatchType": "all",
                         "conditions": [{"name": "missedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["2", "3"]}]},
                        {"name": "refusedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "refusedChatStatus", "conditionMatchType": "all",
                         "conditions": [{"name": "refusedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["1"]}]},
                        {"name": "chatRequests", "calculationType": "expression", "valueType": "int",
                         "expression": "chats + missedChats + refusedChats"},
                        {"name": "chatFromBotToAgentAccepted", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "Status=0&ChatType=2", "conditionMatchType": "all",
                         "conditions": [
                             {"name": "Status=0", "fieldName": "Status", "operate": "equals", "values": ["0"]},
                             {"name": "ChatType=2", "fieldName": "ChatType", "operate": "equals", "values": ["2"]},
                         ]},
                        {"name": "chatFromBotToAgentrefused", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "Status=1&ChatType=1", "conditionMatchType": "all",
                         "conditions": [
                             {"name": "Status=1", "fieldName": "Status", "operate": "equals", "values": ["1"]},
                             {"name": "ChatType=1", "fieldName": "ChatType", "operate": "equals", "values": ["1"]},
                         ]},
                        {"name": "chatFromBotToAgentVisitorCloseWindow", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "Status=2&ChatType=1", "conditionMatchType": "all",
                         "conditions": [
                             {"name": "Status=2", "fieldName": "Status", "operate": "equals", "values": ["2"]},
                             {"name": "ChatType=1", "fieldName": "ChatType", "operate": "equals", "values": ["1"]},
                         ]},
                        {"name": "chatFromBotToAgentVisitorLeaveMsg", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                         "conditionExpression": "Status=2&ChatType=4", "conditionMatchType": "all",
                         "conditions": [
                             {"name": "Status=2", "fieldName": "Status", "operate": "equals", "values": ["2"]},
                             {"name": "ChatType=4", "fieldName": "ChatType", "operate": "equals", "values": ["4"]},
                         ]},
                        {"name": "chatFromBotToAgentRequest", "calculationType": "expression", "valueType": "int",
                         "expression": "chatFromBotToAgentAccepted+chatFromBotToAgentrefused+chatFromBotToAgentVisitorCloseWindow+chatFromBotToAgentVisitorLeaveMsg"},
                    ],
                    "filters": [
                        {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
                        {"fieldName": "CampaignId", "matchType": "equals", "value": [camp_id]},
                    ],
                    "rowGroups": [{"name": "time", "fieldName": "RequestedTime", "isFull": True, "timeDisplayType": "day"}],
                },
                {
                    "name": "VisitsStatistics",
                    "fields": [{"name": "visits", "calculationType": "sum", "valueType": "int", "fieldName": "VisitCount"}],
                    "filters": [
                        {"fieldName": "LogTime", "matchType": "between", "value": [d, d]},
                        {"fieldName": "CampaignId", "matchType": "equals", "value": [camp_id]},
                    ],
                    "rowGroups": [{"name": "time", "fieldName": "LogTime", "isFull": True, "timeDisplayType": "day"}],
                },
            ],
            "mergeType": "column", "timezone": TZ,
        })
    for brand in BRANDS:
        camp_id = CAMPAIGNS[brand]
        payloads.append({
            "cubeEntities": [{
                "name": "Chat",
                "fields": [
                    {"name": "avgChatTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentDuration"},
                    {"name": "avgAgentResponseTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AvgResponseTime"},
                ],
                "filters": [
                    {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
                    {"fieldName": "CampaignId", "matchType": "equals", "value": [camp_id]},
                    {"fieldName": "Status", "matchType": "equals", "value": ["0"]},
                    {"fieldName": "Duration", "matchType": "greaterThan", "value": ["0"]},
                    {"fieldName": "ChatType", "matchType": "equals", "value": ["0", "2"]},
                ],
                "rowGroups": [{"name": "time", "fieldName": "RequestedTime", "isFull": True, "timeDisplayType": "day"}],
            }],
            "mergeType": "column", "timezone": TZ,
        })

    responses = _parallel(payloads)
    brand_data = {}
    for i, brand in enumerate(BRANDS):
        total = (responses[i] or {}).get("total") or {}
        brand_data[brand] = {
            "totalChats": int(total.get("chatRequests") or 0),
            "served": int(total.get("chats") or 0),
            "missed": int(total.get("missedChats") or 0),
            "botToAgent": int(total.get("chatFromBotToAgentRequest") or 0),
            "visits": int(total.get("visits") or 0),
            "avgDur": 0, "avgResp": 0,
        }
    for i, brand in enumerate(BRANDS):
        total = (responses[len(BRANDS) + i] or {}).get("total") or {}
        brand_data[brand]["avgDur"] = float(total.get("avgChatTime") or 0)
        brand_data[brand]["avgResp"] = float(total.get("avgAgentResponseTime") or 0)
    return brand_data


# ============================================================
# CORE REPORTING (6 parallel requests)
# ============================================================
def fetch_core_reporting(date_str):
    d = date_str.replace("-", "/")

    wait_payload = {"cubeEntities": [{"name": "Chat", "fields": [
        {"name": "avgWaitingTime", "calculationType": "average", "valueType": "timespan", "fieldName": "WaitingTime",
         "conditionExpression": "WaitingTime>0", "conditionMatchType": "all",
         "conditions": [{"name": "WaitingTime>0", "fieldName": "WaitingTime", "operate": "notEquals", "values": ["0"]}]},
        {"name": "avgWaitingTimeOfMissedChats", "calculationType": "average", "valueType": "timespan", "fieldName": "WaitingTime",
         "conditionExpression": "Status2 & WaitingTime>0", "conditionMatchType": "all",
         "conditions": [
             {"name": "Status2", "fieldName": "Status", "operate": "equals", "values": ["2"]},
             {"name": "WaitingTime>0", "fieldName": "WaitingTime", "operate": "notEquals", "values": ["0"]},
         ]},
    ], "filters": [
        {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
        {"fieldName": "DepartmentId", "matchType": "equals", "value": [CS_DEPT]},
    ]}], "mergeType": "column", "timezone": TZ}

    efficiency_payload = {"cubeEntities": [{"name": "Chat", "fields": [
        {"name": "avgAgentResponseTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AvgResponseTime"},
        {"name": "avgChatTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentDuration"},
        {"name": "avgFirstResponseTime", "calculationType": "average", "valueType": "timespan", "fieldName": "AgentFirstResponseTime"},
    ], "filters": [
        {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
        {"fieldName": "DepartmentId", "matchType": "equals", "value": [CS_DEPT]},
        {"fieldName": "Duration", "matchType": "greaterThan", "value": ["0"]},
        {"fieldName": "ChatType", "matchType": "equals", "value": ["0", "2"]},
        {"fieldName": "Status", "matchType": "equals", "value": ["0"]},
    ]}], "mergeType": "column", "timezone": TZ}

    volume_payload = {"cubeEntities": [{"name": "Chat", "fields": [
        {"name": "chats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
         "conditionExpression": "Status0 & Duration>0", "conditionMatchType": "all",
         "conditions": [
             {"name": "Status0", "fieldName": "Status", "operate": "equals", "values": ["0"]},
             {"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]},
         ]},
        {"name": "missedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
         "conditionExpression": "missedChatStatus", "conditionMatchType": "all",
         "conditions": [{"name": "missedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["2", "3"]}]},
        {"name": "refusedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
         "conditionExpression": "refusedChatStatus", "conditionMatchType": "all",
         "conditions": [{"name": "refusedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["1"]}]},
        {"name": "chatRequests", "calculationType": "expression", "valueType": "int",
         "expression": "chats + missedChats + refusedChats"},
    ], "filters": [{"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]}]}],
        "mergeType": "column", "timezone": TZ}

    rating_payload = {"cubeEntities": [{"name": "Chat", "fields": [
        {"name": "ratingTimes", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.ChatId"},
        {"name": "score5", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.RatingGrade",
         "conditionExpression": "RatingGrade5", "conditionMatchType": "all",
         "conditions": [{"name": "RatingGrade5", "fieldName": "PostChatSurvey.RatingGrade", "operate": "equals", "values": ["5"]}]},
        {"name": "score4", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.RatingGrade",
         "conditionExpression": "RatingGrade4", "conditionMatchType": "all",
         "conditions": [{"name": "RatingGrade4", "fieldName": "PostChatSurvey.RatingGrade", "operate": "equals", "values": ["4"]}]},
        {"name": "score3", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.RatingGrade",
         "conditionExpression": "RatingGrade3", "conditionMatchType": "all",
         "conditions": [{"name": "RatingGrade3", "fieldName": "PostChatSurvey.RatingGrade", "operate": "equals", "values": ["3"]}]},
        {"name": "score2", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.RatingGrade",
         "conditionExpression": "RatingGrade2", "conditionMatchType": "all",
         "conditions": [{"name": "RatingGrade2", "fieldName": "PostChatSurvey.RatingGrade", "operate": "equals", "values": ["2"]}]},
        {"name": "score1", "calculationType": "count", "valueType": "int", "fieldName": "PostChatSurvey.RatingGrade",
         "conditionExpression": "RatingGrade1", "conditionMatchType": "all",
         "conditions": [{"name": "RatingGrade1", "fieldName": "PostChatSurvey.RatingGrade", "operate": "equals", "values": ["1"]}]},
        {"name": "avgScore", "calculationType": "average", "valueType": "decimal", "fieldName": "PostChatSurvey.RatingGrade"},
    ], "filters": [
        {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
        {"fieldName": "PostChatSurvey.RatingGrade", "matchType": "greaterThan", "value": ["0"]},
        {"fieldName": "ChatType", "matchType": "equals", "value": ["0", "2"]},
        {"fieldName": "Status", "matchType": "equals", "value": ["0"]},
    ]}], "mergeType": "column", "timezone": TZ}

    bot_payload = {"cubeEntities": [{"name": "Chat", "fields": [
        {"name": "botOnlyChats", "calculationType": "count", "valueType": "int", "fieldName": "Id",
         "conditionExpression": "Status & ChatType", "conditionMatchType": "all",
         "conditions": [
             {"name": "Status", "fieldName": "Status", "operate": "equals", "values": ["0", "1", "2"]},
             {"name": "ChatType", "fieldName": "ChatType", "operate": "equals", "values": ["1"]},
         ]},
        {"name": "chatsFromBotToOnlineAgent", "calculationType": "count", "valueType": "int", "fieldName": "Id",
         "conditionExpression": "Status & ChatType", "conditionMatchType": "all",
         "conditions": [
             {"name": "Status", "fieldName": "Status", "operate": "equals", "values": ["0"]},
             {"name": "ChatType", "fieldName": "ChatType", "operate": "equals", "values": ["2"]},
         ]},
    ], "filters": [{"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]}]}],
        "mergeType": "column", "timezone": TZ}

    queue_payload = {"cubeEntities": [
        {"name": "Chat", "fields": [
            {"name": "chatsFromQueue", "calculationType": "count", "valueType": "int", "fieldName": "Status",
             "conditionExpression": "Status0", "conditionMatchType": "all",
             "conditions": [{"name": "Status0", "fieldName": "Status", "operate": "equals", "values": ["0"]}]},
            {"name": "switchedToMessaged", "calculationType": "count", "valueType": "int", "fieldName": "Status",
             "conditionExpression": "Status2 & ChatType4or8or16", "conditionMatchType": "all",
             "conditions": [
                 {"name": "Status2", "fieldName": "Status", "operate": "equals", "values": ["2"]},
                 {"name": "ChatType4or8or16", "fieldName": "ChatType", "operate": "equals", "values": ["4", "8", "16"]},
             ]},
            {"name": "abandonedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
             "conditionExpression": "Status2 & ChatTypeNot4or8or16", "conditionMatchType": "all",
             "conditions": [
                 {"name": "Status2", "fieldName": "Status", "operate": "equals", "values": ["2"]},
                 {"name": "ChatTypeNot4or8or16", "fieldName": "ChatType", "operate": "equals", "values": ["0", "1", "2", "32"]},
             ]},
            {"name": "refusedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
             "conditionExpression": "Status1", "conditionMatchType": "all",
             "conditions": [{"name": "Status1", "fieldName": "Status", "operate": "equals", "values": ["1"]}]},
            {"name": "queuedChatRequests", "calculationType": "expression", "valueType": "int",
             "expression": "chatsFromQueue+switchedToMessaged+abandonedChats+refusedChats"},
        ], "filters": [
            {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
            {"fieldName": "DepartmentId", "matchType": "equals", "value": [CS_DEPT]},
            {"fieldName": "WaitingTime", "matchType": "notEquals", "value": ["0"]},
        ]},
        {"name": "QueueStatistics", "fields": [
            {"name": "maxQueueSize", "calculationType": "max", "valueType": "int", "fieldName": "QueueLength"}
        ], "filters": [
            {"fieldName": "Time", "matchType": "between", "value": [d, d]},
            {"fieldName": "DepartmentId", "matchType": "equals", "value": [CS_DEPT]},
        ]},
    ], "mergeType": "column", "timezone": TZ}

    responses = _parallel([wait_payload, efficiency_payload, volume_payload, rating_payload, bot_payload, queue_payload])
    w = (responses[0] or {}).get("total") or {}
    e = (responses[1] or {}).get("total") or {}
    v = (responses[2] or {}).get("total") or {}
    ra = (responses[3] or {}).get("total") or {}
    b = (responses[4] or {}).get("total") or {}
    q = (responses[5] or {}).get("total") or {}

    return {
        "avgWait": float(w.get("avgWaitingTime") or 0),
        "avgWaitMissed": float(w.get("avgWaitingTimeOfMissedChats") or 0),
        "avgResp": float(e.get("avgAgentResponseTime") or 0),
        "avgFirstResp": float(e.get("avgFirstResponseTime") or 0),
        "avgDur": float(e.get("avgChatTime") or 0),
        "totChats": int(v.get("chatRequests") or 0),
        "missed": int(v.get("missedChats") or 0),
        "rT": int(ra.get("ratingTimes") or 0),
        "r5": int(ra.get("score5") or 0), "r4": int(ra.get("score4") or 0),
        "r3": int(ra.get("score3") or 0), "r2": int(ra.get("score2") or 0), "r1": int(ra.get("score1") or 0),
        "rAvg": ra.get("avgScore"),
        "botOnly": int(b.get("botOnlyChats") or 0),
        "botToAgent": int(b.get("chatsFromBotToOnlineAgent") or 0),
        "qReq": int(q.get("queuedChatRequests") or 0),
        "qSrv": int(q.get("chatsFromQueue") or 0),
        "qMax": int(q.get("maxQueueSize") or 0),
    }


# ============================================================
# BRAND x VIP TIER SEGMENT BREAKDOWN (3 parallel requests)
# ============================================================
def fetch_brand_segment_breakdown(date_str):
    d = date_str.replace("-", "/")
    payloads = []
    for brand in BRANDS:
        camp_id = CAMPAIGNS[brand]
        payloads.append({
            "cubeEntities": [{
                "name": "Chat",
                "fields": [
                    {"name": "chats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                     "conditionExpression": "Status0 & Duration>0", "conditionMatchType": "all",
                     "conditions": [
                         {"name": "Status0", "fieldName": "Status", "operate": "equals", "values": ["0"]},
                         {"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]},
                     ]},
                    {"name": "segmentId", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatSegmentResult.Segment.Id"},
                    {"name": "segmentName", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatSegmentResult.Segment.Name"},
                ],
                "filters": [
                    {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
                    {"fieldName": "CampaignId", "matchType": "equals", "value": [camp_id]},
                ],
                "rowGroups": [{"name": "segmentId", "fieldName": "ChatSegmentResult.SegmentId", "isFull": True}],
            }],
            "mergeType": "column", "timezone": TZ,
        })
    responses = _parallel(payloads)
    brand_tier = {b: {"VIP": 0, "VIP2": 0, "VIP3": 0} for b in BRANDS}
    for i, brand in enumerate(BRANDS):
        series = (responses[i] or {}).get("series") or []
        for s in series:
            name = s.get("segmentName")
            if name in TIER_SEGMENT_NAMES:
                brand_tier[brand][name] = int(s.get("chats") or 0)
    return brand_tier


# ============================================================
# HOURLY VOLUME (genel + brand bazlı, 4 paralel istek)
# Tek istekte chats/missedChats/refusedChats/acceptanceRate/visits — saat
# bazlı — reporting API'nin timeDisplayType="hour" desteğiyle. Bu, artık
# Slide 2 ve Slide 5/6/7'nin saatlik grafik/kutularını TAM olarak
# doldurmamızı sağlıyor (önceki "N/A" alanları burada çözülüyor).
# ============================================================
def _hourly_volume_payload(date_str, campaign_id=None):
    d = date_str.replace("-", "/")
    chat_filters = [{"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]}]
    visits_filters = [{"fieldName": "LogTime", "matchType": "between", "value": [d, d]}]
    if campaign_id:
        chat_filters.append({"fieldName": "CampaignId", "matchType": "equals", "value": [campaign_id]})
        visits_filters.append({"fieldName": "CampaignId", "matchType": "equals", "value": [campaign_id]})
    return {
        "cubeEntities": [
            {
                "name": "Chat",
                "fields": [
                    {"name": "chats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                     "conditionExpression": "Status0 & Duration>0", "conditionMatchType": "all",
                     "conditions": [
                         {"name": "Status0", "fieldName": "Status", "operate": "equals", "values": ["0"]},
                         {"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]},
                     ]},
                    {"name": "missedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                     "conditionExpression": "missedChatStatus", "conditionMatchType": "all",
                     "conditions": [{"name": "missedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["2", "3"]}]},
                    {"name": "refusedChats", "calculationType": "count", "valueType": "int", "fieldName": "Status",
                     "conditionExpression": "refusedChatStatus", "conditionMatchType": "all",
                     "conditions": [{"name": "refusedChatStatus", "fieldName": "Status", "operate": "equals", "values": ["1"]}]},
                    {"name": "chatRequests", "calculationType": "expression", "valueType": "int",
                     "expression": "chats + missedChats + refusedChats"},
                    {"name": "chatAcceptanceRate", "calculationType": "expression", "valueType": "percent",
                     "expression": "chats / chatRequests"},
                    {"name": "time", "calculationType": "originalValue", "valueType": "string"},
                ],
                "filters": chat_filters,
                "rowGroups": [{"name": "time", "fieldName": "RequestedTime", "isFull": True, "timeDisplayType": "hour"}],
            },
            {
                "name": "VisitsStatistics",
                "fields": [{"name": "visits", "calculationType": "sum", "valueType": "int", "fieldName": "VisitCount"}],
                "filters": visits_filters,
                "rowGroups": [{"name": "time", "fieldName": "LogTime", "isFull": True, "timeDisplayType": "hour"}],
            },
        ],
        "mergeType": "column", "timezone": TZ,
    }


def _parse_hourly_series(resp):
    series = (resp or {}).get("series") or []
    rows = []
    for s in series:
        time_str = s.get("time", "")
        # "2026-07-06 00:00~01:00" -> hour = 0
        try:
            hour = int(time_str.split(" ")[1].split(":")[0])
        except (IndexError, ValueError):
            continue
        rows.append({
            "hour": hour,
            "total": int(s.get("chatRequests") or 0),
            "served": int(s.get("chats") or 0),
            "missed": int(s.get("missedChats") or 0),
            "acceptancePct": round(float(s.get("chatAcceptanceRate") or 0), 2),
            "visits": int(s.get("visits") or 0),
        })
    rows.sort(key=lambda r: r["hour"])
    return rows


def fetch_hourly_volume_all(date_str):
    """Returns {"overall": [...], "superbetin": [...], "betsat": [...], "turkbet": [...]}
    each a 24-row (or fewer, sparse hours omitted by the API) list of
    {hour, total, served, missed, acceptancePct, visits}."""
    keys = ["overall"] + BRANDS
    payloads = [_hourly_volume_payload(date_str)] + [_hourly_volume_payload(date_str, CAMPAIGNS[b]) for b in BRANDS]
    responses = _parallel(payloads)
    return {key: _parse_hourly_series(resp) for key, resp in zip(keys, responses)}
# ============================================================
# BRAND x TAG BREAKDOWN (Slide 8'in 3 donut'u için — tek istek)
# Wrap-up survey reporting API'sinden direkt campaignId + categoryOptionName
# kırılımı geliyor, ham chatWrapup.categoriesName parse etmeye gerek yok.
# ============================================================
_CAMPAIGN_NAME_TO_BRAND = {"SUPERBETIN": "superbetin", "BetSat": "betsat", "Turkbet": "turkbet"}


def fetch_brand_tag_breakdown(date_str):
    d = date_str.replace("-", "/")
    payload = {
        "cubeEntities": [{
            "name": "Chat",
            "fields": [
                {"name": "campaignId", "calculationType": "originalValue", "valueType": "string", "fieldName": "Campaign.Id"},
                {"name": "campaignName", "calculationType": "originalValue", "valueType": "string", "fieldName": "Campaign.Name"},
                {"name": "count", "calculationType": "count", "valueType": "int", "fieldName": "Id",
                 "conditionExpression": "Duration>0", "conditionMatchType": "all",
                 "conditions": [{"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]}]},
                {"name": "categoryOptionId", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatWrapupCategory.CategoryOptionId"},
                {"name": "categoryOptionName", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatWrapupCategory.CategoryOption.Name"},
            ],
            "filters": [
                {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
                {"fieldName": "Duration", "matchType": "notEquals", "value": ["0"]},
            ],
            "rowGroups": [
                {"name": "campaignId", "fieldName": "CampaignId", "isFull": False},
                {"name": "optionId", "fieldName": "ChatWrapupCategory.CategoryOptionId", "isFull": False},
            ],
        }],
        "mergeType": "row", "timezone": TZ,
    }
    resp = _report_query(payload)
    series = resp.get("series") or []
    camp_map = campaign_map()
    brand_tags = {b: {} for b in BRANDS}
    for row in series:
        name = row.get("categoryOptionName") or ""
        if not name:
            continue  # tag atanmamış chat'ler (empty categoryOptionName) hariç
        brand = _CAMPAIGN_NAME_TO_BRAND.get(row.get("campaignName") or "")
        if not brand:
            brand = camp_map.get((row.get("campaignId") or "").lower())
        if not brand or brand not in brand_tags:
            continue
        brand_tags[brand][name] = brand_tags[brand].get(name, 0) + int(row.get("count") or 0)
    return {b: sorted(tags.items(), key=lambda kv: -kv[1])[:10] for b, tags in brand_tags.items()}


# ============================================================
# AI AGENT ACTION USAGE (Slide 4 — tek istek)
# ============================================================
def fetch_ai_action_usage(date_str):
    d = date_str.replace("-", "/")
    payload = {
        "cubeEntities": [{
            "name": "AIReplyRecord",
            "fields": [
                {"name": "time", "calculationType": "originalValue", "valueType": "string"},
                {"name": "usedAIReplies", "calculationType": "sum", "valueType": "decimal", "fieldName": "UsedAIReplies"},
            ],
            "filters": [
                {"fieldName": "CreatedTime", "matchType": "between", "value": [d, d]},
                {"fieldName": "UsedBy", "matchType": "equals", "value": ["0"]},
            ],
            "rowGroups": [{"name": "time", "fieldName": "CreatedTime", "isFull": True, "timeDisplayType": "day"}],
        }],
        "mergeType": "column", "timezone": TZ,
    }
    total = (_report_query(payload) or {}).get("total") or {}
    return float(total.get("usedAIReplies") or 0)


def fetch_tag_breakdown_overall(date_str):
    """Slide 8'in alt tablosu (genel Top 10 Tags) için — brand filtresi
    olmadan, aynı Wrap-up reporting API'sinden. Ham chatWrapup.categoriesName
    regex parse etmekten (tag_from_wrapup) çok daha güvenilir — o yöntem
    bazı düz (parantezsiz) "VIP TIER3" gibi kategorileri kaçırıyordu."""
    d = date_str.replace("-", "/")
    payload = {
        "cubeEntities": [{
            "name": "Chat",
            "fields": [
                {"name": "count", "calculationType": "count", "valueType": "int", "fieldName": "Id",
                 "conditionExpression": "Duration>0", "conditionMatchType": "all",
                 "conditions": [{"name": "Duration>0", "fieldName": "Duration", "operate": "notEquals", "values": ["0"]}]},
                {"name": "categoryOptionId", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatWrapupCategory.CategoryOptionId"},
                {"name": "categoryOptionName", "calculationType": "originalValue", "valueType": "string", "fieldName": "ChatWrapupCategory.CategoryOption.Name"},
            ],
            "filters": [
                {"fieldName": "RequestedTime", "matchType": "between", "value": [d, d]},
                {"fieldName": "Duration", "matchType": "notEquals", "value": ["0"]},
            ],
            "rowGroups": [{"name": "optionId", "fieldName": "ChatWrapupCategory.CategoryOptionId", "isFull": False}],
        }],
        "mergeType": "row", "timezone": TZ,
    }
    resp = _report_query(payload)
    series = resp.get("series") or []
    tags = {}
    for row in series:
        name = row.get("categoryOptionName") or ""
        if not name:
            continue
        tags[name] = tags.get(name, 0) + int(row.get("count") or 0)
    return sorted(tags.items(), key=lambda kv: -kv[1])[:10]


# ============================================================
# MAIN METRICS BUILDER — combines raw chat history + reporting API
# ============================================================
def build_metrics(date_str):
    chats = fetch_all_chats(date_str, date_str)
    camp_map = campaign_map()

    agent_names = {}
    tag_map = {}
    hour_map = {h: 0 for h in range(24)}
    hour_agent_map = {h: set() for h in range(24)}
    bm = {b: {"botOnly": 0, "agentOnly": 0, "ratingSum": 0, "ratedCount": 0} for b in BRANDS + ["other"]}

    for c in chats:
        ct = c.get("chatType") or ""
        is_bot_only = (ct == "chatBotOnly")
        a_name = agent_name(c)
        if a_name:
            agent_names[a_name] = agent_names.get(a_name, 0) + 1
        hr = sofia_hour(c.get("startTime"), date_str)
        if 0 <= hr <= 23:
            hour_map[hr] += 1
            if a_name:
                hour_agent_map[hr].add(a_name)
        tag = tag_from_wrapup(c)
        if tag and tag != "—":
            tag_map[tag] = tag_map.get(tag, 0) + 1
        brand = brand_of(c, camp_map)
        bmb = bm.get(brand, bm["other"])
        if is_bot_only:
            bmb["botOnly"] += 1
        elif ct != "fromBotToAgent":
            bmb["agentOnly"] += 1
        survey = c.get("postChatSurvey") or {}
        rating = survey.get("ratingGrade")
        rating = int(rating) if rating is not None else 0
        if 1 <= rating <= 5 and not is_bot_only:
            bmb["ratingSum"] += rating
            bmb["ratedCount"] += 1

    api_brand = fetch_brand_volumes_and_visits(date_str)
    api_core = fetch_core_reporting(date_str)
    api_segment = fetch_brand_segment_breakdown(date_str)
    hourly_all = fetch_hourly_volume_all(date_str)
    brand_tags = fetch_brand_tag_breakdown(date_str)
    action_usage = fetch_ai_action_usage(date_str)
    overall_tags = fetch_tag_breakdown_overall(date_str)

    m = {}
    m["totalChats"] = api_core["totChats"] if api_core["totChats"] > 0 else len(chats)
    m["missed"] = api_core["missed"] if api_core["totChats"] > 0 else 0
    served = m["totalChats"] - m["missed"]

    m["brandMetrics"] = {}
    total_uniq = 0
    for b in BRANDS + ["other"]:
        d_local = bm[b]
        api_d = api_brand.get(b) or {"totalChats": 0, "served": 0, "missed": 0, "botToAgent": 0, "visits": 0, "avgDur": 0, "avgResp": 0}
        if b == "other":
            api_d = {"totalChats": 0, "served": 0, "missed": 0, "botToAgent": 0, "visits": 0, "avgDur": 0, "avgResp": 0}
        total_uniq += api_d["visits"]
        m["brandMetrics"][b] = {
            "total": api_d["totalChats"],
            "served": api_d["served"],
            "missed": api_d["missed"],
            "botToAgent": api_d["botToAgent"],
            "accPct": round(api_d["served"] / api_d["totalChats"] * 100, 2) if api_d["totalChats"] > 0 else 0,
            "unique": api_d["visits"],
            "botOnly": d_local["botOnly"],
            "avgDur": api_d["avgDur"] or 0,
            "avgResp": api_d["avgResp"] or 0,
            "sat": round(d_local["ratingSum"] / d_local["ratedCount"], 2) if d_local["ratedCount"] > 0 else None,
        }

    m["agentOnly"] = sum(bm[b]["agentOnly"] for b in BRANDS + ["other"])
    m["botOnly"] = api_core["botOnly"]
    m["botToAgent"] = api_core["botToAgent"]
    m["actionUsage"] = action_usage
    m["actionPerChat"] = round(action_usage / m["botToAgent"], 2) if m["botToAgent"] else 0
    m["uniqueVisitors"] = total_uniq
    m["activeAgents"] = len(agent_names)
    m["botVolumePct"] = round(m["botOnly"] / m["totalChats"] * 100, 2) if m["totalChats"] > 0 else 0
    m["acceptancePct"] = round(served / m["totalChats"] * 100, 2) if m["totalChats"] > 0 else 0
    m["missedPct"] = round(m["missed"] / m["totalChats"] * 100, 2) if m["totalChats"] > 0 else 0
    m["avgDuration"] = api_core["avgDur"]
    m["avgResponse"] = api_core["avgResp"]
    m["avgFirstResponse"] = api_core["avgFirstResp"]
    m["avgWaitServed"] = api_core["avgWait"]
    m["avgWaitMissed"] = api_core["avgWaitMissed"]
    m["rated"] = api_core["rT"]
    m["ratedPct"] = round(m["rated"] / m["totalChats"] * 100, 2) if m["totalChats"] > 0 else 0
    m["r5"], m["r4"], m["r3"], m["r2"], m["r1"] = api_core["r5"], api_core["r4"], api_core["r3"], api_core["r2"], api_core["r1"]
    m["lowRated"] = m["r1"] + m["r2"]
    m["lowRatePct"] = round(m["lowRated"] / served * 100, 2) if served > 0 else 0
    m["goodRated"] = m["r4"] + m["r5"]
    m["goodRatePct"] = round(m["goodRated"] / served * 100, 2) if served > 0 else 0
    m["satisfaction"] = round(float(api_core["rAvg"]), 2) if api_core["rAvg"] else None
    m["qReq"], m["qSrv"], m["qMax"] = api_core["qReq"], api_core["qSrv"], api_core["qMax"]
    m["avgChatsPerAgent"] = round(m["totalChats"] / m["activeAgents"]) if m["activeAgents"] > 0 else 0

    m["topTags"] = overall_tags  # reporting API'den (regex parse değil — Bug fix: VIP TIER3 kaçırılıyordu)
    m["topAgents"] = sorted(agent_names.items(), key=lambda kv: -kv[1])[:10]

    # Saatlik dağılım artık reporting API'den (served/missed/acceptance/visits
    # dahil, tam veri) — ham chat fetch'ten sadece "o saatte aktif agent
    # sayısı" için hour_agent_map kullanıyoruz (reporting API bunu vermiyor).
    hourly_report = hourly_all["overall"]
    hourly_by_hour = {h["hour"]: h for h in hourly_report}
    m["hourlyDist"] = [
        {
            "hour": h,
            "count": hourly_by_hour.get(h, {}).get("total", hour_map[h]),
            "served": hourly_by_hour.get(h, {}).get("served", 0),
            "missed": hourly_by_hour.get(h, {}).get("missed", 0),
            "acceptancePct": hourly_by_hour.get(h, {}).get("acceptancePct", 0),
            "agents": len(hour_agent_map[h]),
        }
        for h in range(24)
    ]
    # En kötü 3 saat artık MISSED sayısına göre seçiliyor (orijinal deck'in
    # mantığıyla birebir örtüşüyor — "hangi saatte en çok chat kaçtı").
    missed_sorted = sorted(m["hourlyDist"], key=lambda h: -h["missed"])
    m["peakHours"] = [h for h in missed_sorted if h["missed"] > 0][:3]
    if not m["peakHours"]:
        # Missed hiç yoksa (iyi bir gün!) en yoğun 3 saati göster
        m["peakHours"] = sorted(m["hourlyDist"], key=lambda h: -h["count"])[:3]

    m["brandTierMap"] = api_segment
    m["brandHourly"] = {b: hourly_all[b] for b in BRANDS}
    m["brandTagMap"] = brand_tags
    return m


if __name__ == "__main__":
    import sys
    date_str = sys.argv[1] if len(sys.argv) > 1 else dt.date.today().isoformat()
    metrics = build_metrics(date_str)
    print(json.dumps(metrics, indent=2, default=str))
