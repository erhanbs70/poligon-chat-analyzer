"""
make_template.py
Converts the original CS_Team_daily_chat_stats pptx into a reusable template
with {{TOKEN}} placeholders, ready for automated daily filling.

Run once (or whenever the visual design changes):
    python3 make_template.py template_source.pptx template.pptx
"""
import sys
from pptx import Presentation


def set_paragraph_text(para, new_text):
    """Put the whole new_text into the paragraph's first run, blank the rest.
    Used for paragraphs we fully own (narrative notes, title date)."""
    if not para.runs:
        return
    para.runs[0].text = new_text
    for r in para.runs[1:]:
        r.text = ""


def replace_exact_runs(shape, mapping):
    """mapping: {exact_original_run_text: new_run_text}
    Only replaces runs whose text matches EXACTLY (safe, no partial-substring
    ambiguity across different facts living in the same shape)."""
    if not shape.has_text_frame:
        return
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if run.text in mapping:
                run.text = mapping[run.text]


def replace_run_by_index(shape, para_idx, run_idx, new_text):
    para = shape.text_frame.paragraphs[para_idx]
    para.runs[run_idx].text = new_text


def get_shape(slide, name):
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    raise KeyError(f"Shape '{name}' not found on slide")


def build_template(src_path, out_path):
    prs = Presentation(src_path)
    slides = prs.slides

    # ---------------- SLIDE 1 : Title ----------------
    s1 = slides[0]
    subtitle = get_shape(s1, "Subtitle 2")
    set_paragraph_text(subtitle.text_frame.paragraphs[0], "CS TEAM_DAILY_CHAT_STATS_{{REPORT_DATE}}")

    # ---------------- SLIDE 2 : Missed Chats (genel) ----------------
    s2 = slides[1]

    tb1 = get_shape(s2, "TextBox 1")
    replace_exact_runs(tb1, {
        "      - Total Chats: 5486": "      - Total Chats: {{S2_TOTAL}}",
        "      - Total Served Chats: 5348": "      - Total Served Chats: {{S2_SERVED}}",
        "138": "{{S2_MISSED}}",
        ": 97,48": ": {{S2_SUCCESS}}",
    })
    # narrative paragraph (index 4) -> single auto-generated note token
    set_paragraph_text(tb1.text_frame.paragraphs[4], "{{S2_NOTE}}")

    # Hour breakdown boxes -> index-based (run positions verified stable)
    def fill_hour_box(shape, prefix):
        runs = shape.text_frame.paragraphs[0].runs
        n = len(runs)
        runs[1].text = "{{" + prefix + "_HSTART}}"
        runs[3].text = "{{" + prefix + "_HEND}}"
        runs[6].text = " {{" + prefix + "_TOTAL}}"
        runs[8].text = "{{" + prefix + "_SERVED}}"
        runs[10].text = "{{" + prefix + "_MISSED}}"
        runs[12].text = "{{" + prefix + "_SUCCESS}}"
        runs[14].text = "{{" + prefix + "_AGENTS}}"
        # last field (avg chat/agent) may be split across 1 or 2 trailing runs
        # depending on whether the original decimal value got its own run
        if n >= 19:
            runs[17].text = "{{" + prefix + "_CPA}}"
            for r in runs[18:]:
                r.text = ""
        else:
            runs[-1].text = " {{" + prefix + "_CPA}}"

    fill_hour_box(get_shape(s2, "TextBox 3"), "S2_H1")
    fill_hour_box(get_shape(s2, "TextBox 12"), "S2_H2")
    fill_hour_box(get_shape(s2, "TextBox 15"), "S2_H3")

    # ---------------- SLIDE 3 : Response Time & Chat Duration ----------------
    s3 = slides[2]
    tb = get_shape(s3, "TextBox 12")
    replace_exact_runs(tb, {
        ". response time: 41": ". response time: {{S3_RESPONSE}}",
        "Avg. First response time: 35s": "Avg. First response time: {{S3_FIRST_RESPONSE}}",
        "Avg. Wait Time: 1m 32s": "Avg. Wait Time: {{S3_WAIT}}",
        "Avg. Chat Duration: 6m 45s": "Avg. Chat Duration: {{S3_DURATION}}",
    })

    # ---------------- SLIDE 4 : Chatbot Statistics (genel) ----------------
    s4 = slides[3]
    tb = get_shape(s4, "TextBox 5")
    replace_exact_runs(tb, {
        "Chatbot only chats : 1112": "Chatbot only chats : {{S4_BOT_ONLY}}",
        "From chatbot to agent : 3619": "From chatbot to agent : {{S4_BOT_TO_AGENT}}",
        "Percentage of chatbot : 23,50": "Percentage of chatbot : {{S4_BOT_PCT}}",
        "Action usage : 21.171": "Action usage : {{S4_ACTION_USAGE}}",
        "Action per chat : 5.84": "Action per chat : {{S4_ACTION_PER_CHAT}}",
    })

    # ---------------- SLIDE 5 : Superbetin ----------------
    s5 = slides[4]
    tb = get_shape(s5, "TextBox 2")
    replace_exact_runs(tb, {
        "- Total Chats: 3324": "- Total Chats: {{S5_TOTAL}}",
        "- Total served chats: 3230": "- Total served chats: {{S5_SERVED}}",
        "- Total Missed chats: 37": "- Total Missed chats: {{S5_MISSED}}",
        "97,17": "{{S5_SUCCESS}}",
    })
    tb = get_shape(s5, "TextBox 18")
    replace_exact_runs(tb, {
        "Chatbot only chats : 663": "Chatbot only chats : {{S5_BOT_ONLY}}",
        "2071": "{{S5_BOT_TO_AGENT}}",
        "Percentage of chatbot : 24,25": "Percentage of chatbot : {{S5_BOT_PCT}}",
    })

    # ---------------- SLIDE 6 : Betsat ----------------
    s6 = slides[5]
    tb = get_shape(s6, "TextBox 1")
    replace_exact_runs(tb, {
        "- Total Chats: 1848": "- Total Chats: {{S6_TOTAL}}",
        " Total served chats: 1813": " Total served chats: {{S6_SERVED}}",
        "- Total Missed chats: 35": "- Total Missed chats: {{S6_MISSED}}",
        "Succes: 98,11%": "Succes: {{S6_SUCCESS}}%",
    })
    tb = get_shape(s6, "TextBox 3")
    replace_exact_runs(tb, {
        "Chatbot only chats : 393": "Chatbot only chats : {{S6_BOT_ONLY}}",
        "- From chatbot to agent : 1325": "- From chatbot to agent : {{S6_BOT_TO_AGENT}}",
        "Percentage of chatbot: 22,88%": "Percentage of chatbot: {{S6_BOT_PCT}}%",
    })

    # ---------------- SLIDE 7 : Turkbet ----------------
    s7 = slides[6]
    tb = get_shape(s7, "TextBox 1")
    replace_exact_runs(tb, {
        "- Total Chats: 314": "- Total Chats: {{S7_TOTAL}}",
        "- Total served chats: 305": "- Total served chats: {{S7_SERVED}}",
        "- Total Missed chats: 9": "- Total Missed chats: {{S7_MISSED}}",
        "Succes: 97,13%": "Succes: {{S7_SUCCESS}}%",
    })
    tb = get_shape(s7, "TextBox 11")
    replace_exact_runs(tb, {
        "Chatbot only chats : 56": "Chatbot only chats : {{S7_BOT_ONLY}}",
        "From chatbot to agent : 223": "From chatbot to agent : {{S7_BOT_TO_AGENT}}",
        "Percentage of chatbot : 20,07": "Percentage of chatbot : {{S7_BOT_PCT}}",
    })

    # ---------------- SLIDE 8 : Top 10 Tags ----------------
    # Table cells are filled directly in build_pptx.py at run time (not tokens),
    # since the number of rows/tags is dynamic. Nothing to templatize here.

    # ---------------- SLIDE 9 : dropped from automation for now ----------------
    # (kept as-is in the template; simply not filled / can be deleted later)

    prs.save(out_path)
    print(f"Template saved -> {out_path}")


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "template_source.pptx"
    out = sys.argv[2] if len(sys.argv) > 2 else "template.pptx"
    build_template(src, out)
