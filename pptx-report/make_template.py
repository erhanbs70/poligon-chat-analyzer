"""
make_template.py
Converts the original CS_Team_daily_chat_stats pptx into a reusable template
with {{TOKEN}} placeholders, ready for automated daily filling.

Run once (or whenever the visual design changes):
    python3 make_template.py template_source.pptx template.pptx
"""
import sys
import os
import io
import zipfile
from pptx import Presentation
from pptx.util import Emu, Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from PIL import Image

ASSETS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")


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


def _identify_brand_color(shape):
    """Shape'in içindeki görselin baskın rengine bakarak markayı tespit eder
    (kırmızı=Turkbet, sarı=Betsat, mavi=Superbetin)."""
    try:
        blob = shape.image.blob
    except Exception:
        return None
    img = Image.open(io.BytesIO(blob)).convert("RGBA")
    w, h = img.size
    counts = {"red": 0, "yellow": 0, "blue": 0}
    for x in range(0, w, max(1, w // 20)):
        for y in range(0, h, max(1, h // 6)):
            r, g, b, a = img.getpixel((x, y))
            if a < 50:
                continue
            if r > 180 and g < 80 and b < 80:
                counts["red"] += 1
            elif r > 180 and g > 150 and b < 80:
                counts["yellow"] += 1
            elif b > 120 and r < 100:
                counts["blue"] += 1
    if not any(counts.values()):
        return None
    return max(counts, key=counts.get)


# Marka logolarının GERÇEK dosya en-boy oranları (bkz. _replace_logo_images'taki
# hedef boyutlar) — tüm slaytlarda aynı YÜKSEKLİKTE görünmeleri için kullanılıyor.
_BRAND_ASPECT = {"red": 1326 / 317, "yellow": 400 / 103, "blue": 359 / 115}
_LOGO_HEIGHT_IN = 0.55
_LOGO_TOP_IN = 0.30
_LOGO_MARGIN_X_IN = 0.60
_SLIDE_WIDTH_IN = 13.33


def standardize_header_logos(slides):
    """Slide 2-9'daki marka logolarını (Betsat/Superbetin/Turkbet) tek bir
    standart boyut ve konum şemasına çeker — önceden her slaytta farklı
    boyut/pozisyondaydılar (orijinal insan yapımı deck'ten miras kalma
    tutarsızlık), artık hepsi aynı yükseklikte ve aynı hizada."""
    for i in range(1, 9):  # Slide 2..9 (0-index 1..8)
        slide = slides[i]
        header_shapes = []
        for s in slide.shapes:
            if s.shape_type == 13 and s.top is not None:
                top_in = s.top / 914400
                h_in = s.height / 914400
                if top_in < 1.1 and h_in < 1.4:
                    brand = _identify_brand_color(s)
                    if brand:
                        header_shapes.append((s, brand))

        multi = len(set(b for _, b in header_shapes)) > 1
        for shape, brand in header_shapes:
            aspect = _BRAND_ASPECT[brand]
            w_in = _LOGO_HEIGHT_IN * aspect
            shape.height = Inches(_LOGO_HEIGHT_IN)
            shape.width = Inches(w_in)
            shape.top = Inches(_LOGO_TOP_IN)
            if multi:
                if brand == "yellow":       # Betsat -> sol
                    shape.left = Inches(_LOGO_MARGIN_X_IN)
                elif brand == "blue":       # Superbetin -> orta
                    shape.left = Inches((_SLIDE_WIDTH_IN - w_in) / 2)
                elif brand == "red":        # Turkbet -> sağ
                    shape.left = Inches(_SLIDE_WIDTH_IN - _LOGO_MARGIN_X_IN - w_in)
            else:
                # Tek marka gösteren slayt (5/6/7) -> ortala
                shape.left = Inches((_SLIDE_WIDTH_IN - w_in) / 2)


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
    # Toplam Chats/Served/Missed/Success artık build_pptx.py'de KPI kart
    # olarak ekleniyor — bu metin kutusundaki eski madde işaretli satırları
    # (header + 3 sayısal satır) siliyoruz, sadece anlatı notu kalıyor.
    paras = tb1.text_frame.paragraphs
    for idx in (3, 2, 1, 0):  # ters sırada sil, index kaymasın
        paras[idx]._p.getparent().remove(paras[idx]._p)
    # not paragrafı artık index 0
    set_paragraph_text(tb1.text_frame.paragraphs[0], "{{S2_NOTE}}")
    # Kartlara yer açmak için metin kutusunu aşağı kaydırıp küçültüyoruz
    tb1.top = Emu(int(4.97 * 914400))
    tb1.height = Emu(int(0.89 * 914400))

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
    # Küçük tekrar logoları (büyük marka logolarının hemen altındaki küçük
    # kopyaları) kalıcı olarak siliniyor — kalabalık/gereksiz görünüyorlardı.
    s8 = slides[7]
    for shape_name in ("Picture 9", "Picture 10", "Picture 15"):
        try:
            shp = get_shape(s8, shape_name)
            shp._element.getparent().remove(shp._element)
        except KeyError:
            pass

    # ---------------- SLIDE 9 : E-mail/Call (manuel) + Yesterday's Values (otomatik) ----------------
    s9 = slides[8]

    # "Yesterday's Values" — bu veri zaten Slide 3 ile aynı kaynaktan geliyor,
    # otomatik doldurulabilir.
    tb14 = get_shape(s9, "TextBox 14")
    replace_run_by_index(tb14, 1, 1, "{{S9_RESPONSE}}")
    replace_run_by_index(tb14, 2, 1, "{{S9_DURATION}}")
    replace_run_by_index(tb14, 3, 1, " {{S9_SATISFACTION}}")
    replace_run_by_index(tb14, 4, 2, " {{S9_WAIT_SERVED}}")
    replace_run_by_index(tb14, 5, 2, "{{S9_WAIT_MISSED}}")
    replace_run_by_index(tb14, 5, 3, "")
    replace_run_by_index(tb14, 6, 1, "{{S9_ACCEPTANCE}}")
    replace_run_by_index(tb14, 7, 1, " {{S9_RATED_PCT}}")

    # "Solved E-mails" — Zendesk kaynaklı, elimizde yok. Manuel doldurulacak (boş/—).
    email_shape = get_shape(s9, "Content Placeholder 2")
    replace_run_by_index(email_shape, 0, 5, "{{S9_EMAIL_SB_COUNT}}")
    replace_run_by_index(email_shape, 0, 10, " {{S9_EMAIL_SB_TIME}}")
    replace_run_by_index(email_shape, 0, 13,
                          ": {{S9_EMAIL_BS_COUNT}} / First resolution time median: {{S9_EMAIL_BS_TIME}}")
    replace_run_by_index(email_shape, 0, 18, " {{S9_EMAIL_TB_COUNT}} ")
    replace_run_by_index(email_shape, 0, 20, " : {{S9_EMAIL_TB_TIME}}")

    # "Call Statistics" — call panel kaynaklı, elimizde yok. Manuel doldurulacak (boş/—).
    tb_bs = get_shape(s9, "TextBox 4")  # BS = Betsat
    replace_run_by_index(tb_bs, 0, 6, " {{S9_CALL_BS_ATTEMPTS}}")
    replace_run_by_index(tb_bs, 0, 10, " {{S9_CALL_BS_REACHED}} ")
    replace_run_by_index(tb_bs, 0, 12, "{{S9_CALL_BS_REACHED_PCT}}")
    replace_run_by_index(tb_bs, 0, 13, ")")
    replace_run_by_index(tb_bs, 0, 18, "{{S9_CALL_BS_NOTREACHED}}")
    replace_run_by_index(tb_bs, 0, 20, "{{S9_CALL_BS_NOTREACHED_PCT}}")
    replace_run_by_index(tb_bs, 0, 21, ")")

    tb_sb = get_shape(s9, "TextBox 15")  # SB = Superbetin
    replace_run_by_index(tb_sb, 0, 6, " {{S9_CALL_SB_ATTEMPTS}}")
    replace_run_by_index(tb_sb, 1, 3, "{{S9_CALL_SB_REACHED}}")
    replace_run_by_index(tb_sb, 1, 4, " ({{S9_CALL_SB_REACHED_PCT}})")
    replace_run_by_index(tb_sb, 1, 9, "{{S9_CALL_SB_NOTREACHED}}")
    replace_run_by_index(tb_sb, 1, 10, " ({{S9_CALL_SB_NOTREACHED_PCT}})")

    tb_tb = get_shape(s9, "TextBox 16")  # TB = Turkbet
    replace_run_by_index(tb_tb, 0, 6, " {{S9_CALL_TB_ATTEMPTS}}")
    replace_run_by_index(tb_tb, 1, 2, " {{S9_CALL_TB_REACHED}} (")
    replace_run_by_index(tb_tb, 1, 3, "{{S9_CALL_TB_REACHED_PCT}}")
    replace_run_by_index(tb_tb, 1, 4, ")")
    replace_run_by_index(tb_tb, 1, 10, "{{S9_CALL_TB_NOTREACHED}}")
    replace_run_by_index(tb_tb, 1, 11, " ({{S9_CALL_TB_NOTREACHED_PCT}})")

    # Kaynağı belirsiz, tek başına duran "207 " kutusu — temizliyoruz.
    rect1 = get_shape(s9, "Rectangle 1")
    rect1.text_frame.paragraphs[0].runs[0].text = ""

    # ---------------- TASARIM B: koyu lacivert header bandı ----------------
    # Slide 2-9 (içerik slaytları) — eski gri-krem gradient yerine düz beyaz
    # zemin + üstte koyu lacivert bir header bandı (logolar bandın üstünde
    # kalıyor, konumlarına dokunulmuyor).
    NAVY = RGBColor(0x1B, 0x24, 0x36)
    for i in range(1, 9):
        slide = slides[i]
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.33), Inches(1.15))
        band.fill.solid()
        band.fill.fore_color.rgb = NAVY
        band.line.fill.background()
        band.shadow.inherit = False
        band.text_frame.paragraphs[0].text = ""
        # en arkaya gönder (logoların altına) — spTree'de ilk çocuk yap
        sp_tree = band._element.getparent()
        sp_tree.remove(band._element)
        sp_tree.insert(2, band._element)  # ilk 2 eleman nvGrpSpPr/grpSpPr, sonrası shape'ler

    # Layout seviyesindeki kalın turuncu alt bar -> ince lacivert çizgiye çeviriliyor
    # (tüm slaytları tek noktadan etkiler).
    layout = slides[1].slide_layout
    for shp in layout.shapes:
        if shp.name in ("Rectangle 4", "Rectangle 5"):
            shp.fill.solid()
            shp.fill.fore_color.rgb = NAVY
    for shp in layout.shapes:
        if shp.name == "Rectangle 4":
            shp.top = Inches(7.42)
            shp.height = Inches(0.06)
        elif shp.name == "Rectangle 5":
            shp.top = Inches(7.36)
            shp.height = Inches(0.06)

    # ---------------- SLIDE 1: diğerleriyle tutarlı navy tasarım ----------------
    s1_slide = slides[0]
    s1_slide.background.fill.solid()
    s1_slide.background.fill.fore_color.rgb = NAVY
    # Alt başlık (tarih) beyaza çevriliyor (navy zeminde okunur olsun diye)
    subtitle = get_shape(s1_slide, "Subtitle 2")
    for para in subtitle.text_frame.paragraphs:
        for run in para.runs:
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # ---------------- SLIDE 2: yorum/not alanı (asıl istenen yer burasıydı) ----------------
    note_box = s2.shapes.add_textbox(Inches(10.75), Inches(5.90), Inches(2.47), Inches(0.42))
    note_tf = note_box.text_frame
    note_tf.word_wrap = True
    note_p = note_tf.paragraphs[0]
    note_p.alignment = PP_ALIGN.CENTER
    note_r = note_p.add_run()
    note_r.text = "(Yorum eklemek için tıklayın)"
    note_r.font.size = Pt(9)
    note_r.font.italic = True
    note_r.font.color.rgb = RGBColor(0x9C, 0xA3, 0xAF)

    # ---------------- Header logolarını tüm slaytlarda standart boyut/konuma çekiyoruz ----------------
    standardize_header_logos(slides)

    prs.save(out_path)

    # ---------------- Marka logoları — daha modern/okunur varyantlarla değiştiriliyor ----------------
    # image3/image9 = Turkbet (kırmızı kutu + artık BEYAZ "BET" yazısı, navy zeminde okunur)
    # image4/image8 = Betsat (sarı wordmark, eski mor-kutulu versiyon yerine)
    # image5        = Superbetin (ikon aynen korunuyor, sadece "superbetin" yazısı beyaza boyandı)
    _replace_logo_images(out_path)
    print(f"Template saved -> {out_path}")


def _fit_and_pad(src_path, target_w, target_h):
    """Kaynak görseli, oranını bozmadan hedef kutuya sığdırıp ortalar
    (şeffaf dolgu ile) — marka logosunun gerilip deforme olmaması için."""
    src = Image.open(src_path).convert("RGBA")
    sw, sh = src.size
    scale = min(target_w / sw, target_h / sh)
    nw, nh = max(1, round(sw * scale)), max(1, round(sh * scale))
    resized = src.resize((nw, nh), Image.LANCZOS)
    canvas = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
    canvas.paste(resized, ((target_w - nw) // 2, (target_h - nh) // 2), resized)
    buf = io.BytesIO()
    canvas.save(buf, format="PNG")
    return buf.getvalue()


def _replace_logo_images(pptx_path):
    """template.pptx içindeki eski marka logosu media dosyalarını, aynı
    dosya adlarını koruyarak yeni logo asset'leriyle değiştirir (böylece
    o dosyayı referans eden TÜM slaytlar otomatik güncellenir)."""
    targets = {
        "ppt/media/image3.png": (os.path.join(ASSETS_DIR, "logo_turkbet.png"), None),
        "ppt/media/image9.png": (os.path.join(ASSETS_DIR, "logo_turkbet.png"), (667, 160)),
        "ppt/media/image4.png": (os.path.join(ASSETS_DIR, "logo_betsat.png"), (400, 103)),
        "ppt/media/image8.png": (os.path.join(ASSETS_DIR, "logo_betsat.png"), (182, 65)),
        "ppt/media/image5.png": (os.path.join(ASSETS_DIR, "logo_superbetin.png"), (359, 115)),
    }
    new_bytes = {}
    for arcname, (src_path, target_size) in targets.items():
        if target_size is None:
            with open(src_path, "rb") as f:
                new_bytes[arcname] = f.read()
        else:
            new_bytes[arcname] = _fit_and_pad(src_path, *target_size)

    tmp_path = pptx_path + ".tmp"
    with zipfile.ZipFile(pptx_path, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = new_bytes.get(item.filename, zin.read(item.filename))
            zout.writestr(item, data)
    os.replace(tmp_path, pptx_path)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "template_source.pptx"
    out = sys.argv[2] if len(sys.argv) > 2 else "template.pptx"
    build_template(src, out)
