"""Quick test: builds a pptx from FAKE metrics (no live API calls) so we
can verify the template-filling + chart-generation pipeline end to end."""
from build_pptx import build_pptx

mock_metrics = {
    "totalChats": 5486,
    "missed": 138,
    "acceptancePct": 97.48,
    "avgResponse": 41,
    "avgFirstResponse": 35.98,
    "avgDuration": 405,
    "avgWaitServed": 92,
    "avgWaitMissed": 103,
    "botOnly": 1112,
    "botToAgent": 3619,
    "actionUsage": 21171,
    "actionPerChat": 5.85,
    "botVolumePct": 23.50,
    "uniqueVisitors": 59215,
    "activeAgents": 24,
    "rated": 476,
    "ratedPct": 8.67,
    "satisfaction": 4.06,
    "r5": 300, "r4": 100, "r3": 40, "r2": 20, "r1": 16,
    "lowRated": 36, "lowRatePct": 0.67,
    "goodRated": 400, "goodRatePct": 7.48,
    "qReq": 5486, "qSrv": 5300, "qMax": 42,
    "topTags": [
        ("VIP TIER3", 731), ("wd_check", 679), ("deposit_check", 542),
        ("goodwill_query", 478), ("deposit_missing", 432), ("deposit_issue", 242),
        ("sport_cashback_query", 184), ("wd_missing", 144),
        ("casino_cashback_query", 141), ("free_spin_query", 140),
    ],
    "topAgents": [("Ecem", 300), ("Bensu", 280)],
    "hourlyDist": [
        {
            "hour": h,
            "count": (h * 13) % 400 + 50,
            "served": ((h * 13) % 400 + 50) - (39 if h == 1 else 35 if h == 20 else 28 if h == 21 else 0),
            "missed": 39 if h == 1 else 35 if h == 20 else 28 if h == 21 else 0,
            "acceptancePct": 89.01 if h == 1 else 90.98 if h == 20 else 93.56 if h == 21 else 100.0,
            "agents": max(1, h % 12),
        }
        for h in range(24)
    ],
    "peakHours": [
        {"hour": 1, "count": 356, "served": 317, "missed": 39, "acceptancePct": 89.01, "agents": 8},
        {"hour": 20, "count": 388, "served": 353, "missed": 35, "acceptancePct": 90.98, "agents": 11},
        {"hour": 21, "count": 435, "served": 407, "missed": 28, "acceptancePct": 93.56, "agents": 12},
    ],
    "brandMetrics": {
        "superbetin": {"total": 1356, "served": 1356, "missed": 0, "botToAgent": 889, "accPct": 100.0,
                       "unique": 33841, "botOnly": 240, "avgDur": 396, "avgResp": 41, "sat": 3.87},
        "betsat": {"total": 804, "served": 802, "missed": 2, "botToAgent": 600, "accPct": 99.75,
                   "unique": 20283, "botOnly": 141, "avgDur": 356, "avgResp": 37, "sat": 3.63},
        "turkbet": {"total": 132, "served": 132, "missed": 0, "botToAgent": 87, "accPct": 100.0,
                    "unique": 5091, "botOnly": 30, "avgDur": 386, "avgResp": 40, "sat": 4.56},
        "other": {"total": 0, "served": 0, "missed": 0, "botToAgent": 0, "accPct": 0,
                  "unique": 0, "botOnly": 0, "avgDur": 0, "avgResp": 0, "sat": None},
    },
    "brandTierMap": {
        "superbetin": {"VIP": 20, "VIP2": 54, "VIP3": 60},
        "betsat": {"VIP": 2, "VIP2": 15, "VIP3": 22},
        "turkbet": {"VIP": 2, "VIP2": 3, "VIP3": 1},
    },
    "brandHourly": {
        "superbetin": [{"hour": h, "total": (h*7)%150+20, "served": (h*7)%150+18, "missed": 2 if h in (1,20,21) else 0, "acceptancePct": 95.0} for h in range(24)],
        "betsat": [{"hour": h, "total": (h*5)%90+10, "served": (h*5)%90+9, "missed": 1 if h in (1,20,21) else 0, "acceptancePct": 96.0} for h in range(24)],
        "turkbet": [{"hour": h, "total": (h*2)%20+2, "served": (h*2)%20+2, "missed": 0, "acceptancePct": 100.0} for h in range(24)],
    },
    "brandTagMap": {
        "superbetin": [("VIP TIER3", 495), ("wd_check", 401), ("deposit_check", 328), ("goodwill_query", 295),
                       ("deposit_missing", 280), ("deposit_issue", 179), ("sport_cashback_query", 113),
                       ("free_spin_query", 102), ("wd_missing", 101), ("no_answer", 83)],
        "betsat": [("wd_check", 239), ("VIP TIER3", 214), ("deposit_check", 195), ("goodwill_query", 163),
                   ("deposit_missing", 134), ("sport_cashback_query", 68), ("casino_cashback_query", 60),
                   ("deposit_issue", 59), ("no_answer", 43), ("wd_missing", 35)],
        "turkbet": [("wd_check", 39), ("casino_instant_cashback", 30), ("VIP TIER3", 22), ("goodwill_query", 20),
                    ("deposit_check", 19), ("wr_query", 12), ("fake_site", 11), ("no_answer", 8),
                    ("free_spin_query", 10), ("casino_cashback_query", 9)],
    },
}

out = build_pptx(mock_metrics, "template.pptx", "test_output.pptx", "06.07.2026")
print("Built:", out)
