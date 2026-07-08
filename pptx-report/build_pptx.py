"""
build_pptx.py
Fills template.pptx with a metrics dict (as produced by fetch_metrics.build_metrics)
and produces the final daily deck.

All 7 content slides (2-8) are now fully data-driven — no more "N/A"
placeholders except Slide 3's First Response Time and Slide 4's chatbot
Action usage/per chat, which need separate data sources not yet wired in.
"""
import io
import copy
from pptx import Presentation
from pptx.util import Inches, Pt
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

BRAND_COLORS = {
    "superbetin": "#2A3990",
    "betsat": "#F2C230",
    "turkbet": "#D71920",
}
BRAND_LABELS = {"superbetin": "Superbetin", "betsat": "Betsat", "turkbet": "Turkbet"}


# ============================================================
# FORMATTING HELPERS (mirror the GAS _mrFmt* helpers)
# ============================================================
def fmt_time(seconds):
    if not seconds or seconds <= 0:
        return "—"
    seconds = round(seconds)
    h, rem = divmod(seconds, 3600)
    m, s = divmod(rem, 60)
    if h > 0:
        return f"{h:02d}:{m:02d}:{s:02d}"
    return f"{m:02d}:{s:02d}"


def fmt_min_sec(seconds):
    """'1m 32s' style, matching the original deck's wording."""
    if not seconds or seconds <= 0:
        return "0s"
    seconds = round(seconds)
    m, s = divmod(seconds, 60)
    return f"{m}m {s}s" if m > 0 else f"{s}s"


def fmt_pct(value):
    if value is None:
        return "—"
    return f"{value:.2f}".replace(".", ",")


def fmt_n(n):
    return f"{round(n):,}".replace(",", ".")


# ============================================================
# TEXT REPLACEMENT (token -> value, run-level exact match)
# ============================================================
def replace_all_tokens(prs, tokens):
    for slide in prs.slides:
        _replace_in_shapes(slide.shapes, tokens)


def _replace_in_shapes(shapes, tokens):
    for shape in shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                for run in para.runs:
                    for token, value in tokens.items():
                        if token in run.text:
                            run.text = run.text.replace(token, str(value))
        if getattr(shape, "shape_type", None) is not None and shape.shape_type == 6:  # GROUP
            _replace_in_shapes(shape.shapes, tokens)


# ============================================================
# CHART GENERATION (matplotlib -> in-memory PNG)
# ============================================================
def _save_png(fig, width_in, height_in):
    """KRİTİK: figür zaten doğru width_in x height_in ile oluşturulmuş
    olmalı (plt.subplots(figsize=(...))). Burada bbox_inches='tight'
    KULLANMIYORUZ — çünkü o, içeriğe göre kırpma yapıp PNG'nin gerçek
    piksel oranını figsize'dan saptırıyor, bu da PowerPoint'te görseli
    slayt kutusuna sığdırırken dikey gerilip metnin "şişmiş/büyük"
    görünmesine sebep oluyordu. Sabit boyutla kaydedince oran birebir
    korunuyor, hiç gerilme olmuyor."""
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150)
    plt.close(fig)
    buf.seek(0)
    return buf


def build_hourly_volume_chart(hourly_dist, title="Chats by Hour", figsize=(10.63, 3.57)):
    """Served (mavi) + Missed (kırmızı, üstte stack) bar + Acceptance Rate
    (yeşil çizgi, ikincil eksen) — orijinal Comm100 dashboard görünümüne
    yakın, gerçek saatlik served/missed verisiyle."""
    hours = [f'{h["hour"]:02d}:00' for h in hourly_dist]
    served = [h.get("served", 0) for h in hourly_dist]
    missed = [h.get("missed", 0) for h in hourly_dist]
    acc = [h.get("acceptancePct", 0) for h in hourly_dist]

    fig, ax1 = plt.subplots(figsize=figsize)
    ax1.bar(hours, served, color="#2E6DA4", label="Served")
    ax1.bar(hours, missed, bottom=served, color="#D9534F", label="Missed")
    ax1.set_ylabel("Chats", fontsize=8)
    ax1.tick_params(axis="x", rotation=90, labelsize=6)
    ax1.tick_params(axis="y", labelsize=7)

    ax2 = ax1.twinx()
    ax2.plot(hours, acc, color="#5CB85C", marker="o", markersize=2.5, linewidth=1.2, label="Acceptance %")
    ax2.set_ylim(0, 105)
    ax2.set_ylabel("Acceptance %", fontsize=8)
    ax2.tick_params(axis="y", labelsize=7)

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", fontsize=6.5)
    ax1.set_title(title, fontsize=10)
    fig.tight_layout(pad=0.6)
    return _save_png(fig, *figsize)


def build_brand_tag_donuts(brand_tag_map, figsize=(11.10, 2.98)):
    """Slide 8 — 3 marka için ayrı donut (orijinal deck'teki gibi).
    Dar dikey alana (2.98") sığması için: en fazla 6 tag + "Diğer" grubu,
    küçük ve kompakt legend."""
    fig, axes = plt.subplots(1, 3, figsize=figsize)
    order = ["betsat", "superbetin", "turkbet"]  # orijinal deck'teki soldan sağa sıralama
    max_slices = 6
    for ax, brand in zip(axes, order):
        tags = brand_tag_map.get(brand, [])
        if not tags:
            ax.axis("off")
            continue
        top = tags[:max_slices]
        rest_sum = sum(c for _, c in tags[max_slices:])
        labels = [t[0] for t in top]
        values = [t[1] for t in top]
        if rest_sum > 0:
            labels.append("Diğer")
            values.append(rest_sum)
        ax.pie(values, wedgeprops=dict(width=0.45), startangle=90,
               colors=plt.cm.tab20.colors[:len(values)])
        ax.set_title(BRAND_LABELS[brand], fontsize=9, pad=2)
        ax.legend(labels, loc="center left", bbox_to_anchor=(0.95, 0.5),
                  fontsize=5.5, frameon=False, labelspacing=0.3, handlelength=1)
    fig.tight_layout(pad=0.4, w_pad=2.2)
    return _save_png(fig, *figsize)


def build_top_tags_chart(top_tags, figsize=(11.10, 2.98)):
    """Genel (tüm markalar birleşik) Top 10 tag bar grafiği — artık
    build_brand_tag_donuts ile değiştirildiği için slide 8'de kullanılmıyor,
    ama başka bir yerde ihtiyaç olursa diye tutuluyor."""
    tags = [t[0] for t in top_tags][::-1]
    counts = [t[1] for t in top_tags][::-1]
    fig, ax = plt.subplots(figsize=figsize)
    ax.barh(tags, counts, color="#2E6DA4")
    ax.set_title("Top 10 Chat Tags", fontsize=10)
    ax.tick_params(axis="both", labelsize=7)
    for i, v in enumerate(counts):
        ax.text(v, i, f" {v}", va="center", fontsize=7)
    fig.tight_layout(pad=0.6)
    return _save_png(fig, *figsize)


def build_chatbot_chart(bot_only, bot_to_agent, figsize=(10.67, 3.67)):
    """Slide 4 — simple bot-only vs bot-to-agent comparison bar."""
    fig, ax = plt.subplots(figsize=figsize)
    labels = ["Chatbot Only", "Bot \u2192 Agent"]
    values = [bot_only, bot_to_agent]
    ax.bar(labels, values, color=["#F2A93B", "#2E6DA4"], width=0.5)
    for i, v in enumerate(values):
        ax.text(i, v, fmt_n(v), ha="center", va="bottom", fontsize=9)
    ax.set_title("Chatbot Volume", fontsize=10)
    ax.tick_params(axis="both", labelsize=8)
    fig.tight_layout(pad=0.6)
    return _save_png(fig, *figsize)


# ============================================================
# IMAGE PLACEMENT (delete old screenshot, insert new chart at same box)
# ============================================================
def replace_picture(slide, shape_name, png_stream, left_in, top_in, width_in, height_in):
    target = None
    for shape in slide.shapes:
        if shape.name == shape_name:
            target = shape
            break
    if target is not None:
        target._element.getparent().remove(target._element)
    slide.shapes.add_picture(png_stream, Inches(left_in), Inches(top_in),
                              width=Inches(width_in), height=Inches(height_in))


# ============================================================
# TABLE FILL (Slide 8 — Top 10 Tags table)
# ============================================================
def fill_top_tags_table(slide, top_tags):
    table_shape = None
    for shape in slide.shapes:
        if shape.has_table:
            table_shape = shape
            break
    if table_shape is None:
        return
    table = table_shape.table
    for i, (tag, count) in enumerate(top_tags):
        if i >= len(table.rows):
            break
        table.cell(i, 0).text = tag
        table.cell(i, 1).text = str(count)


# ============================================================
# MAIN BUILD FUNCTION
# ============================================================
def build_pptx(metrics, template_path, out_path, date_label):
    prs = Presentation(template_path)
    m = metrics
    bmap = m["brandMetrics"]
    served = m["totalChats"] - m["missed"]

    # ---- text tokens ----
    peak = m["peakHours"][0] if m["peakHours"] else {"hour": 0, "count": 0, "missed": 0, "acceptancePct": 0}
    if peak.get("missed", 0) > 0:
        note = (f'En yüksek missed chat sayısı {peak["hour"]:02d}:00–{(peak["hour"]+1)%24:02d}:00 '
                f'arasında görüldü ({fmt_n(peak["missed"])} chat kaçtı, kabul oranı %{fmt_pct(peak["acceptancePct"])}). '
                f'Olayın sebebini buraya ekleyin.')
    else:
        note = (f'En yoğun saat aralığı {peak["hour"]:02d}:00–{(peak["hour"]+1)%24:02d}:00 '
                f'oldu ({fmt_n(peak["count"])} chat). Missed chat görülmedi.')

    tokens = {
        "{{REPORT_DATE}}": date_label,
        "{{S2_TOTAL}}": fmt_n(m["totalChats"]),
        "{{S2_SERVED}}": fmt_n(served),
        "{{S2_MISSED}}": fmt_n(m["missed"]),
        "{{S2_SUCCESS}}": fmt_pct(m["acceptancePct"]),
        "{{S2_NOTE}}": note,
        "{{S3_RESPONSE}}": str(round(m["avgResponse"])),
        "{{S3_FIRST_RESPONSE}}": f'{round(m["avgFirstResponse"])}s',
        "{{S3_WAIT}}": fmt_min_sec(m["avgWaitServed"]),
        "{{S3_DURATION}}": fmt_min_sec(m["avgDuration"]),
        "{{S4_BOT_ONLY}}": fmt_n(m["botOnly"]),
        "{{S4_BOT_TO_AGENT}}": fmt_n(m["botToAgent"]),
        "{{S4_BOT_PCT}}": fmt_pct(
            round(m["botOnly"] / (m["botOnly"] + m["botToAgent"]) * 100, 2)
            if (m["botOnly"] + m["botToAgent"]) else 0
        ),
        "{{S4_ACTION_USAGE}}": fmt_n(m["actionUsage"]),
        "{{S4_ACTION_PER_CHAT}}": f'{m["actionPerChat"]:.2f}'.replace(".", ","),
    }

    for i, brand in enumerate(["superbetin", "betsat", "turkbet"]):
        prefix = f"S{5+i}"
        d = bmap[brand]
        tokens[f"{{{{{prefix}_TOTAL}}}}"] = fmt_n(d["total"])
        tokens[f"{{{{{prefix}_SERVED}}}}"] = fmt_n(d["served"])
        tokens[f"{{{{{prefix}_MISSED}}}}"] = fmt_n(d["missed"])
        tokens[f"{{{{{prefix}_SUCCESS}}}}"] = fmt_pct(d["accPct"])
        tokens[f"{{{{{prefix}_BOT_ONLY}}}}"] = fmt_n(d["botOnly"])
        tokens[f"{{{{{prefix}_BOT_TO_AGENT}}}}"] = fmt_n(d["botToAgent"])
        tokens[f"{{{{{prefix}_BOT_PCT}}}}"] = fmt_pct(
            round(d["botOnly"] / (d["botOnly"] + d["botToAgent"]) * 100, 2)
            if (d["botOnly"] + d["botToAgent"]) else 0
        )

    # ---- Slide 2 hour boxes: en kötü 3 saat, artık tam veriyle ----
    top3 = m["peakHours"][:3] if m["peakHours"] else []
    while len(top3) < 3:
        top3.append({"hour": 0, "count": 0, "served": 0, "missed": 0, "acceptancePct": 0, "agents": 0})
    for i, h in enumerate(top3, start=1):
        prefix = f"S2_H{i}"
        cpa = round(h["count"] / h["agents"]) if h.get("agents") else 0
        tokens[f"{{{{{prefix}_HSTART}}}}"] = f'{h["hour"]:02d}'
        tokens[f"{{{{{prefix}_HEND}}}}"] = f'{(h["hour"]+1)%24:02d}'
        tokens[f"{{{{{prefix}_TOTAL}}}}"] = fmt_n(h["count"])
        tokens[f"{{{{{prefix}_SERVED}}}}"] = fmt_n(h.get("served", 0))
        tokens[f"{{{{{prefix}_MISSED}}}}"] = fmt_n(h.get("missed", 0))
        tokens[f"{{{{{prefix}_SUCCESS}}}}"] = fmt_pct(h.get("acceptancePct", 0))
        tokens[f"{{{{{prefix}_AGENTS}}}}"] = str(h.get("agents", 0))
        tokens[f"{{{{{prefix}_CPA}}}}"] = str(cpa)

    replace_all_tokens(prs, tokens)

    # ---- Slide 8: Top 10 Tags table ----
    fill_top_tags_table(prs.slides[7], m["topTags"])

    # ---- Charts ----
    replace_picture(prs.slides[1], "Picture 5",
                     build_hourly_volume_chart(m["hourlyDist"], "Chats by Hour", figsize=(10.63, 3.57)),
                     0.01, 1.52, 10.63, 3.57)
    replace_picture(prs.slides[3], "Picture 8",
                     build_chatbot_chart(m["botOnly"], m["botToAgent"], figsize=(10.67, 3.67)),
                     0.00, 1.90, 10.67, 3.67)
    replace_picture(prs.slides[7], "Picture 4",
                     build_brand_tag_donuts(m.get("brandTagMap", {}), figsize=(11.10, 2.98)),
                     0.85, 1.30, 11.10, 2.98)

    brand_hourly = m.get("brandHourly", {})
    replace_picture(prs.slides[4], "Picture 10",
                     build_hourly_volume_chart(brand_hourly.get("superbetin", []), "Superbetin — Chats by Hour", figsize=(10.71, 3.57)),
                     0.00, 1.37, 10.71, 3.57)
    replace_picture(prs.slides[5], "Picture 4",
                     build_hourly_volume_chart(brand_hourly.get("betsat", []), "Betsat — Chats by Hour", figsize=(10.47, 3.54)),
                     -0.00, 1.31, 10.47, 3.54)
    replace_picture(prs.slides[6], "Picture 5",
                     build_hourly_volume_chart(brand_hourly.get("turkbet", []), "Turkbet — Chats by Hour", figsize=(10.34, 3.50)),
                     0.00, 1.27, 10.34, 3.50)

    # ---- Slide 9 (email/call stats) dropped from this deck for now ----
    xml_slides = prs.slides._sldIdLst
    slides = list(xml_slides)
    if len(slides) >= 9:
        xml_slides.remove(slides[8])

    prs.save(out_path)
    return out_path


if __name__ == "__main__":
    import sys
    import datetime as dt
    sys.path.insert(0, ".")
    from fetch_metrics import build_metrics

    date_str = sys.argv[1] if len(sys.argv) > 1 else dt.date.today().isoformat()
    metrics = build_metrics(date_str)
    date_label = dt.datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")
    out = build_pptx(metrics, "template.pptx", f"CS_Team_daily_chat_stats_{date_label.replace('.', '_')}.pptx", date_label)
    print("Saved:", out)
