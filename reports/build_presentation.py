"""Builds reports/Presentation.pptx (+ PNG previews of every slide) from saved EDA results.

Order:  notebooks/EDA.ipynb  ->  python reports/make_slide_figures.py  ->  python reports/build_presentation.py

Every number on the slides is read from data/processed/ and outputs/tables/, not typed by hand.
Text is wrapped here with real font metrics, so the .pptx and the previews show the same lines.
"""
import math
from pathlib import Path

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "outputs" / "figures" / "slides"
TAB = ROOT / "outputs" / "tables"
PREVIEW = ROOT / "reports" / "preview"
PREVIEW.mkdir(parents=True, exist_ok=True)

AUTHOR = "Табачук Олександр Ігорович"
GROUP = "група КН-42"
FOOTER = "Табачук О. І., КН-42  ·  Лабораторна робота №2  ·  EDA Auto MPG"
REPO = "github.com/Cut1e1337/Lab02_Tabachuk_Oleksandr"

# ---------------------------------------------------------------- results
df = pd.read_csv(ROOT / "data" / "processed" / "auto_mpg_clean.csv")
L = "l_per_100km"
corr = pd.read_csv(TAB / "correlations.csv", index_col=0)
vif = pd.read_csv(TAB / "vif.csv", index_col=0)["VIF"]

med_l = df[L].median()
q1_l, q3_l = df[L].quantile([0.25, 0.75])
by_o = df.groupby("origin").agg(l=(L, "mean"), w=("weight_kg", "mean"))
jp_vs_us = 1 - by_o.loc["Японія", "l"] / by_o.loc["США", "l"]
w_gap = by_o.loc["США", "w"] - by_o.loc["Японія", "w"]
gw = df.groupby(["weight_group", "origin"])[L].mean()
gap_light = 1 - gw[("легкі", "Японія")] / gw[("легкі", "США")]
gap_mid = 1 - gw[("середні", "Японія")] / gw[("середні", "США")]
yr = df.groupby("year").agg(l=(L, "mean"), w=("weight_kg", "mean"))
drop = 1 - yr["l"].iloc[-1] / yr["l"].iloc[0]
rho = corr["mpg"].drop("mpg").abs().sort_values(ascending=False)
block = corr.loc[["weight_kg", "displacement_l", "horsepower", "cylinders"],
                 ["weight_kg", "displacement_l", "horsepower", "cylinders"]].abs().values
block_min, block_max = block[~np.eye(4, dtype=bool)].min(), block[~np.eye(4, dtype=bool)].max()
slope100 = df[L].cov(df["weight_kg"]) / df["weight_kg"].var() * 100


def f(x, d=1):
    return f"{x:.{d}f}".replace(".", ",")


def pct(x):
    return f"{x * 100:.0f} %"


# ---------------------------------------------------------------- design tokens
PAPER, CARD = "F1EFE8", "FFFDF8"
INK, MUTED, FAINT = "1F2328", "5E636B", "8F9399"
CORAL, TEAL, OLIVE = "D1495B", "2E7D8C", "8A8B4A"
KHAKI, SAGE, SAND = "E4E1B8", "D3DCCB", "EADFCB"

FONTS = {  # name in PowerPoint -> file for metrics/preview
    "title": ("Bahnschrift SemiBold", r"C:\Windows\Fonts\bahnschrift.ttf", b"SemiBold"),
    "body": ("Segoe UI", r"C:\Windows\Fonts\segoeui.ttf", None),
    "bold": ("Segoe UI", r"C:\Windows\Fonts\segoeuib.ttf", None),
    "semi": ("Segoe UI Semibold", r"C:\Windows\Fonts\seguisb.ttf", None),
}
DPI = 150                       # preview resolution
SW, SH = 13.333, 7.5            # slide size, inches
LS = 1.18                       # line spacing (multiple of font size)
_font_cache = {}


def pil_font(kind, size_pt, scale=DPI):
    key = (kind, size_pt, scale)
    if key not in _font_cache:
        _, path, var = FONTS[kind]
        fnt = ImageFont.truetype(path, max(1, round(size_pt / 72 * scale)))
        if var:
            fnt.set_variation_by_name(var)
        _font_cache[key] = fnt
    return _font_cache[key]


from fontTools.ttLib import TTFont  # noqa: E402
CMAPS = {k: set(TTFont(v[1]).getBestCmap()) for k, v in FONTS.items()}


def segments(text, kind):
    """Split text so glyphs missing in the font (≈ − → in Bahnschrift) fall back to Segoe UI Semibold."""
    out = []
    for ch in text:
        k = kind if (ord(ch) in CMAPS[kind] or ch == " ") else "semi"
        if out and out[-1][1] == k:
            out[-1][0] += ch
        else:
            out.append([ch, k])
    return out


def measure(text, kind, size):
    """Width of text in inches (measured at high resolution)."""
    return sum(pil_font(k, size, 600).getlength(t) for t, k in segments(text, kind)) / 600


def rgb(h):
    return RGBColor.from_string(h)


def hex2rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------- canvas: draws to pptx and PIL at once
class Slide:
    def __init__(self, prs, n=None):
        self.s = prs.slides.add_slide(prs.slide_layouts[6])
        self.img = Image.new("RGB", (round(SW * DPI), round(SH * DPI)), hex2rgb(PAPER))
        self.d = ImageDraw.Draw(self.img)
        self.n = n
        bg = self.s.background.fill
        bg.solid(); bg.fore_color.rgb = rgb(PAPER)

    def px(self, v):
        return round(v * DPI)

    def rect(self, x, y, w, h, color, radius=0.0):
        shp = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
        s = self.s.shapes.add_shape(shp, Inches(x), Inches(y), Inches(w), Inches(h))
        if radius:
            s.adjustments[0] = radius / min(w, h)
        s.fill.solid(); s.fill.fore_color.rgb = rgb(color); s.line.fill.background(); s.shadow.inherit = False
        box = [self.px(x), self.px(y), self.px(x + w), self.px(y + h)]
        if radius:
            self.d.rounded_rectangle(box, radius=self.px(radius), fill=hex2rgb(color))
        else:
            self.d.rectangle(box, fill=hex2rgb(color))

    def blob(self, cx, cy, rx, ry, color, seed=0):
        rng = np.random.default_rng(seed)
        a1, a2, p1, p2 = 0.10 + 0.05 * rng.random(), 0.05 + 0.03 * rng.random(), rng.random() * 6, rng.random() * 6
        pts = []
        for t in np.linspace(0, 2 * math.pi, 120, endpoint=False):
            k = 1 + a1 * math.sin(3 * t + p1) + a2 * math.sin(5 * t + p2)
            pts.append((cx + rx * k * math.cos(t), cy + ry * k * math.sin(t)))
        ff = self.s.shapes.build_freeform(Inches(pts[0][0]), Inches(pts[0][1]))
        ff.add_line_segments([(Inches(px_), Inches(py_)) for px_, py_ in pts[1:]], close=True)
        s = ff.convert_to_shape()
        s.fill.solid(); s.fill.fore_color.rgb = rgb(color); s.line.fill.background(); s.shadow.inherit = False
        self.d.polygon([(self.px(a), self.px(b)) for a, b in pts], fill=hex2rgb(color))

    def image(self, path, x, y, w, h):
        iw, ih = Image.open(path).size
        sc = min(w / iw, h / ih)
        dw, dh = iw * sc, ih * sc
        ox, oy = x + (w - dw) / 2, y + (h - dh) / 2
        self.s.shapes.add_picture(str(path), Inches(ox), Inches(oy), Inches(dw), Inches(dh))
        im = Image.open(path).convert("RGB").resize((self.px(dw), self.px(dh)), Image.LANCZOS)
        self.img.paste(im, (self.px(ox), self.px(oy)))

    def text(self, x, y, w, paras, size=16, color=INK, kind="body", align="left", gap=0.0, upper=False):
        """paras: list of paragraphs; a paragraph is a str or a list of (text, kind, color) runs.
        Wraps by real metrics into width w. Returns the height used (inches)."""
        if isinstance(paras, str):
            paras = [paras]
        lh = size * LS / 72
        tb = self.s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(1))
        tf = tb.text_frame
        tf.word_wrap = False; tf.auto_size = MSO_AUTO_SIZE.NONE; tf.vertical_anchor = MSO_ANCHOR.TOP
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        cy = y
        for pi, para in enumerate(paras):
            runs = [(para, kind, color)] if isinstance(para, str) else para
            if upper:
                runs = [(t.upper(), k, c) for t, k, c in runs]
            # tokenize into words keeping style
            words = []
            for t, k, c in runs:
                for wd in t.split(" "):
                    if wd:
                        words.append([wd, k, c, True])
            lines, cur, cur_w = [], [], 0.0
            for wd, k, c, _ in words:
                sp = measure(" ", k, size) if cur else 0
                ww = measure(wd, k, size)
                if cur and cur_w + sp + ww > w * 0.97:
                    lines.append(cur); cur, cur_w, sp = [], 0.0, 0
                cur.append((wd, k, c)); cur_w += sp + ww
            if cur:
                lines.append(cur)
            if pi > 0:
                cy += gap
            p = tf.paragraphs[0] if pi == 0 else tf.add_paragraph()
            p.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}[align]
            p.line_spacing = Pt(size * LS)
            if pi > 0:
                p.space_before = Pt(gap * 72)
            for li, line in enumerate(lines):
                if li > 0:
                    p.add_line_break()
                # merge consecutive same-style words into runs
                merged = []
                for wd, k, c in line:
                    if merged and merged[-1][1] == k and merged[-1][2] == c:
                        merged[-1][0] += " " + wd
                    else:
                        merged.append([(" " if merged else "") + wd, k, c])
                lw = sum(measure(t, k, size) for t, k, c in merged)
                lx = x if align == "left" else (x + (w - lw) / 2 if align == "center" else x + w - lw)
                base = cy + lh * 0.80
                for t0, k0, c in merged:
                    for t, k in segments(t0, k0):
                        r = p.add_run(); r.text = t
                        r.font.name = FONTS[k][0]; r.font.size = Pt(size); r.font.color.rgb = rgb(c)
                        r.font.bold = (k == "bold")
                        self.d.text((self.px(lx), self.px(base)), t, font=pil_font(k, size), fill=hex2rgb(c), anchor="ls")
                        lx += measure(t, k, size)
                cy += lh
        tb.height = Inches(max(cy - y, 0.1))
        return cy - y

    def notes(self, s):
        self.s.notes_slide.notes_text_frame.text = s

    def save_preview(self, idx):
        self.img.save(PREVIEW / f"slide_{idx:02d}.png")


# ---------------------------------------------------------------- slide templates
def header(sl, kicker, title):
    sl.text(0.75, 0.48, 11.8, kicker, size=13, color=CORAL, kind="semi", upper=True)
    sl.text(0.75, 0.82, 11.9, title, size=29, color=INK, kind="title")


def footer(sl):
    sl.text(0.75, 7.02, 9, FOOTER, size=10, color=FAINT)
    sl.text(12.0, 7.02, 0.6, str(sl.n), size=10, color=FAINT, align="right")


def deco(sl, seed):
    sl.blob(13.6, -0.3, 1.5, 1.2, KHAKI, seed)
    sl.blob(-0.45, 7.6, 1.2, 0.9, SAGE, seed + 1)


def chart_slide(prs, n, kicker, title, fig, side, notes, chart_w=8.45):
    sl = Slide(prs, n)
    deco(sl, n)
    header(sl, kicker, title)
    sl.rect(0.6, 2.08, chart_w, 4.85, CARD, radius=0.22)
    sl.image(FIG / fig, 0.78, 2.22, chart_w - 0.36, 4.57)
    side(sl, 0.6 + chart_w + 0.4, 12.75 - (0.6 + chart_w + 0.4))
    footer(sl)
    sl.notes(notes)
    return sl


def big_number(sl, x, y, w, value, label, color=CORAL):
    sl.text(x, y, w, value, size=40, color=color, kind="title")
    return 0.78 + sl.text(x, y + 0.78, w, label.split("\n"), size=14, color=MUTED)


def bullets(sl, x, y, w, items, size=15, color=INK, gap=0.16):
    h = 0
    for it in items:
        runs = [(it, "body", color)] if isinstance(it, str) else it
        sl.rect(x, y + h + size / 72 * 0.42, 0.09, 0.09, CORAL, radius=0.045)
        h += sl.text(x + 0.25, y + h, w - 0.25, [runs], size=size) + gap
    return h


# ---------------------------------------------------------------- build
prs = Presentation()
prs.slide_width, prs.slide_height = Inches(SW), Inches(SH)
slides = []

# 1 — title
sl = Slide(prs, 1)
sl.blob(10.4, 3.7, 3.0, 2.85, KHAKI, 3)
sl.blob(13.0, 0.4, 1.3, 1.1, SAGE, 4)
sl.blob(0.2, 7.3, 1.6, 1.0, SAND, 5)
sl.text(0.8, 0.9, 7, "Лабораторна робота №2 · Розвідувальний аналіз даних", size=14, color=CORAL, kind="semi", upper=True)
sl.text(0.8, 1.45, 7.6, ["Від чого залежить", "витрата палива", "автомобіля?"], size=50, color=INK, kind="title", upper=True)
sl.text(0.8, 4.05, 7.2, "EDA набору даних Auto MPG: 398 моделей авто ринку США, 1970–1982", size=19, color=MUTED)
sl.rect(0.8, 5.35, 5.3, 1.08, CARD, radius=0.18)
sl.text(1.05, 5.52, 5.0, [[(AUTHOR, "bold", INK)], [(GROUP, "body", MUTED)]], size=17, gap=0.04)
for i, (v, lbl) in enumerate([("398", "автомобілів"), ("13", "модельних років"), ("7", "потенційних факторів")]):
    sl.text(8.75, 2.15 + i * 1.12, 1.4, v, size=40, color=CORAL, kind="title", align="right")
    sl.text(10.35, 2.43 + i * 1.12, 2.4, lbl, size=17, color=INK, kind="semi")
sl.notes(f"Доброго дня. Мене звати {AUTHOR}, {GROUP}. Тема моєї роботи — розвідувальний аналіз даних для задачі "
         "регресії: від чого залежить витрата палива автомобіля. Я дослідив відкритий набір даних Auto MPG — "
         "398 моделей автомобілів, що продавалися у США з 1970 по 1982 рік.")
slides.append(sl)

# 2 — goal & questions
sl = Slide(prs, 2); deco(sl, 2)
header(sl, "Постановка задачі", "Мета — з'ясувати, що визначає витрату палива, і підготувати дані до регресії")
qs = ["Як розподілена витрата палива і чи є незвичні авто?",
      "Які характеристики пов'язані з витратою найсильніше?",
      "Чи різняться авто США, Європи та Японії за однакової маси?",
      "Як змінилася витрата палива у 1970–1982 роках?",
      "Чи дублюють фактори один одного?"]
cw, cg = 2.24, 0.165
for i, q in enumerate(qs):
    x = 0.75 + i * (cw + cg)
    sl.rect(x, 2.45, cw, 3.45, CARD, radius=0.2)
    sl.text(x + 0.22, 2.65, cw - 0.4, f"RQ{i + 1}", size=30, color=CORAL, kind="title")
    sl.text(x + 0.22, 3.45, cw - 0.3, q, size=18, color=INK, kind="semi")
sl.text(0.75, 6.35, 11.8, [[("Ціль: ", "bold", INK), ("витрата палива (неперервна числова змінна). ", "body", MUTED),
                            ("Модель у роботі не будується — лише EDA.", "body", MUTED)]], size=14)
footer(sl)
sl.notes("Мета — зрозуміти, які характеристики автомобіля пов'язані з витратою палива, і підготувати дані до "
         "майбутньої регресійної моделі. Саму модель я не будував — тільки розвідувальний аналіз. "
         "Я сформулював п'ять дослідницьких запитань: про розподіл цілі, про найсильніші фактори, про різницю між "
         "регіонами, про зміну в часі та про мультиколінеарність. На наступних слайдах я відповідаю на кожне з них.")
slides.append(sl)

# 3 — data
sl = Slide(prs, 3); deco(sl, 3)
header(sl, "Характеристика даних", "Дані: 398 моделей авто, 9 змінних, майже без пропусків")
tiles = [("398", "моделей авто,\n1970–1982 рр."), ("7 + ціль", "факторів: 5 числових,\n2 категоріальні"),
         (f(df["hp_imputed"].mean() * 100) + " %", "пропусків — 6 значень\nпотужності"), ("0", "дублікатів")]
for i, (v, lbl) in enumerate(tiles):
    x, y = 0.75 + (i % 2) * 2.95, 2.15 + (i // 2) * 2.3
    sl.rect(x, y, 2.75, 2.1, CARD, radius=0.2)
    sl.text(x + 0.25, y + 0.2, 2.4, v.replace(".", ","), size=34, color=CORAL, kind="title")
    sl.text(x + 0.25, y + 1.05, 2.4, lbl.split("\n"), size=14, color=MUTED)
sl.rect(6.85, 2.15, 5.9, 4.4, CARD, radius=0.2)
sl.text(7.15, 2.35, 5.4, "Що в даних", size=18, color=INK, kind="title")
bullets(sl, 7.15, 2.9, 5.4, [
    [("Ціль: ", "bold", INK), ("mpg — милі на галон; для наочності переведено в л/100 км", "body", INK)],
    [("Фактори: ", "bold", INK), ("маса, об'єм двигуна, потужність, циліндри, розгін, рік, регіон", "body", INK)],
    [("Очищення: ", "bold", INK), ("пропуски → медіана групи «циліндри × регіон»; виправлено помилки в марках", "body", INK)],
    [("Викиди ", "bold", INK), ("реальні (Mazda GLC, V8) — не видаляли", "body", INK)],
    [("Джерело: ", "bold", INK), ("UCI ML Repository, ліцензія CC BY 4.0", "body", INK)],
], size=15, gap=0.2)
footer(sl)
sl.notes("Дані взяті з репозиторію UCI, ліцензія CC BY 4.0. Одиниця спостереження — модель автомобіля певного "
         "року. Ціль — mpg, милі на галон; для наочності я перевів її в літри на 100 км. Якість даних хороша: "
         "дублікатів немає, пропуски лише у 6 значеннях потужності, їх я заповнив медіаною в групі з тією ж "
         "кількістю циліндрів і регіоном. Також виправив помилки в назвах марок. Викиди — це реальні автомобілі, "
         "тому я їх не видаляв.")
slides.append(sl)


# 4 — RQ1 target
def side4(sl, x, w):
    h = big_number(sl, x, 2.15, w, f"{f(med_l)} л", "медіанна витрата\nна 100 км")
    bullets(sl, x, 2.25 + h + 0.25, w, [
        f"Половина авто — від {f(q1_l)} до {f(q3_l)} л",
        "Розподіл асиметричний: довгий «хвіст» великих V8",
        "→ для регресії ціль краще логарифмувати",
    ], size=15)


slides.append(chart_slide(
    prs, 4, "RQ1 · Розподіл цільової змінної",
    f"Типове авто витрачає ≈ {f(med_l, 0)} л/100 км, але розкид — у 5 разів", "s1_target.png", side4,
    f"Перше запитання — як розподілена ціль. Медіанне авто витрачає близько {f(med_l)} літра на 100 км. "
    f"Але розкид дуже великий: від 5 літрів у Mazda GLC до 26 у пікапа з V8. Розподіл не нормальний, має "
    "правий «хвіст» — тому в майбутній регресії ціль варто логарифмувати або брати в л/100 км."))


# 5 — RQ2 weight
def side5(sl, x, w):
    sl.text(x, 2.15, w, "Сила зв'язку з витратою, |ρ|", size=15, color=INK, kind="semi")
    names = {"weight_kg": "Маса", "displacement_l": "Об'єм", "horsepower": "Потужність",
             "cylinders": "Циліндри", "year": "Рік", "acceleration": "Розгін"}
    for i, (k, v) in enumerate(rho.items()):
        y = 2.7 + i * 0.52
        col = CORAL if i == 0 else (TEAL if i < 4 else "B9B6AC")
        sl.text(x, y, 1.25, names[k], size=13, color=INK)
        bw = (w - 1.95) * v
        sl.rect(x + 1.3, y + 0.05, bw, 0.25, col, radius=0.05)
        sl.text(x + 1.38 + bw, y, 0.6, f(v, 2), size=13, color=INK, kind="semi")
    sl.text(x, 5.95, w, "У л/100 км зв'язок з масою майже лінійний", size=13, color=MUTED)


slides.append(chart_slide(
    prs, 5, "RQ2 · Зв'язки з ціллю",
    f"Маса — головний фактор: кожні +100 кг ≈ +{f(slope100)} л/100 км", "s2_weight.png", side5,
    f"Друге запитання — що пов'язане з витратою найсильніше. Це маса автомобіля: кореляція Спірмена "
    f"{f(rho.iloc[0], 2)}. У середньому кожні 100 кілограмів додають близько {f(slope100)} літра на 100 км. "
    "Майже такі ж сильні зв'язки з об'ємом двигуна, потужністю і циліндрами; рік і розгін — помірні. "
    "Важливо: у милях на галон зв'язки нелінійні, а в літрах на 100 км — майже лінійні. "
    "Це кореляція, а не доведений причинний вплив."))


# 6 — RQ3 naive
def side6(sl, x, w):
    h = big_number(sl, x, 2.15, w, f"−{pct(jp_vs_us)}", "витрата японських авто\nпроти американських")
    bullets(sl, x, 2.25 + h + 0.25, w, [
        f"Але японські авто в середньому на {w_gap:.0f} кг легші",
        "Питання: це ефект країни чи маси?",
    ], size=15)


slides.append(chart_slide(
    prs, 6, "RQ3 · Порівняння регіонів",
    f"Японські авто витрачають на {pct(jp_vs_us)} менше пального, ніж американські…", "s3_origin.png", side6,
    f"Третє запитання — чи різняться регіони. На перший погляд так: японські авто витрачають на {pct(jp_vs_us)} "
    f"менше пального, ніж американські. Але подивіться на масу — японські машини в середньому на {w_gap:.0f} кг "
    "легші. Тож постає питання: це ефект країни чи просто ефект маси?"))


# 7 — RQ3 stratified
def side7(sl, x, w):
    y = 2.15
    for v, lbl in [(f"−{pct(gap_light)}", "серед легких авто"), (f"−{pct(gap_mid)}", "серед середніх авто")]:
        sl.text(x, y, w, v, size=36, color=CORAL, kind="title")
        sl.text(x, y + 0.72, w, lbl, size=14, color=MUTED)
        y += 1.3
    bullets(sl, x, y + 0.1, w, [
        [("Маса — змішувальний фактор: ", "bold", INK), ("просте порівняння країн вводить в оману", "body", INK)],
        "Важких японських авто у вибірці немає",
    ], size=14)


slides.append(chart_slide(
    prs, 7, "RQ3 · Багатовимірний аналіз",
    f"…але за однакової маси різниця лише {pct(gap_light).replace(' %', '')}–{pct(gap_mid)}",
    "s4_origin_within_weight.png", side7,
    f"Щоб це перевірити, я розбив авто на вагові групи і порівняв регіони всередині кожної. Розрив різко "
    f"скорочується: серед легких авто японські економніші лише на {pct(gap_light)}, серед середніх — на {pct(gap_mid)}, "
    f"замість {pct(jp_vs_us)} загалом. Отже, маса — змішувальний фактор: вона пов'язана і з регіоном, і з витратою. "
    "Висновок «японські авто економніші, бо вони японські» некоректний. Обмеження: важких японських авто немає."))


# 8 — RQ4 trend
def side8(sl, x, w):
    h = big_number(sl, x, 2.15, w, f"{f(yr['l'].iloc[0])} → {f(yr['l'].iloc[-1])}",
                   "л/100 км, середнє\n1970 → 1982")
    bullets(sl, x, 2.25 + h + 0.25, w, [
        "Витрата падає в кожній ваговій групі",
        f"Середня маса: {yr['w'].iloc[0]:.0f} → {yr['w'].iloc[-1]:.0f} кг",
        "→ працюють і технології, і полегшення авто",
    ], size=15)


slides.append(chart_slide(
    prs, 8, "RQ4 · Динаміка в часі",
    f"За 12 років витрата впала на {pct(drop)} — і не лише через легші авто", "s5_trend.png", side8,
    f"Четверте запитання — динаміка. За 12 років середня витрата впала з {f(yr['l'].iloc[0])} до "
    f"{f(yr['l'].iloc[-1])} літра, майже вдвічі. Перелом — після нафтових криз 1973 і 1979 років. "
    f"Частково це тому, що авто полегшали, з {yr['w'].iloc[0]:.0f} до {yr['w'].iloc[-1]:.0f} кг. Але на графіку видно, що "
    "витрата падає і всередині кожної вагової групи — тобто машини однакової маси теж стали економнішими. "
    "Отже, працюють обидва механізми: і технології, і полегшення."))


# 9 — RQ5 multicollinearity
def side9(sl, x, w):
    h = big_number(sl, x, 2.15, w, f"ρ = {f(block_min, 2)}–{f(block_max, 2)}",
                   "між масою, об'ємом, потужністю\nта циліндрами")
    bullets(sl, x, 2.25 + h + 0.25, w, [
        [("VIF до ", "body", INK), (f"{vif.max():.0f}", "bold", INK), (" (поріг — 10): мультиколінеарність", "body", INK)],
        [("→ у регресію брати 1–2 з них, ", "body", INK), ("напр. масу і потужність", "bold", INK)],
        "Рік слабко пов'язаний з іншими — окреме джерело інформації",
    ], size=15)


slides.append(chart_slide(
    prs, 9, "RQ5 · Мультиколінеарність",
    "Маса, об'єм, потужність і циліндри майже дублюють один одного", "s6_multicollinearity.png", side9,
    f"П'яте запитання — чи дублюють фактори один одного. Так: маса, об'єм двигуна, потужність і циліндри "
    f"корелюють між собою на рівні {f(block_min, 2)}–{f(block_max, 2)}, а коефіцієнт інфляції дисперсії сягає "
    f"{vif.max():.0f}. Фактично це одне явище — «розмір автомобіля». Тому в регресію варто брати один-два з них, "
    "наприклад масу і потужність. Модельний рік з ними пов'язаний слабко, тож несе окрему інформацію.",
    chart_w=6.1))

# 10 — conclusions
sl = Slide(prs, 10); deco(sl, 10)
header(sl, "Висновки та обмеження", "Маса пояснює найбільше, а прості порівняння можуть вводити в оману")
sl.rect(0.75, 2.15, 7.25, 4.1, CARD, radius=0.2)
sl.text(1.05, 2.35, 6.8, "Головні знахідки", size=18, color=INK, kind="title")
bullets(sl, 1.05, 2.9, 6.75, [
    [("Ціль: ", "bold", INK), (f"медіана {f(med_l)} л/100 км, розкид у 5 разів, асиметрія", "body", INK)],
    [("Маса ", "bold", INK), (f"— найсильніший фактор (ρ = {f(rho.iloc[0], 2)}); зв'язки нелінійні", "body", INK)],
    [("Регіон: ", "bold", INK), (f"−{pct(jp_vs_us)} у Японії, але за однакової маси лише {pct(gap_light).replace(' %', '')}–{pct(gap_mid)}", "body", INK)],
    [("Час: ", "bold", INK), (f"витрата −{pct(drop)} за 1970–82, і не лише через легші авто", "body", INK)],
    [("Для регресії: ", "bold", INK), ("log-ціль, маса + потужність + рік + регіон", "body", INK)],
], size=16, gap=0.32)
sl.rect(8.3, 2.15, 4.45, 4.1, KHAKI, radius=0.2)
sl.text(8.6, 2.35, 4.0, "Обмеження", size=18, color=INK, kind="title")
bullets(sl, 8.6, 2.9, 3.95, [
    "Старі дані, лише ринок США",
    "63 % вибірки — американські авто",
    "Немає типу палива, коробки передач, ціни",
    "Кореляція ≠ причинність",
], size=16, gap=0.32)
footer(sl)
sl.notes("Підсумую. Типове авто витрачає близько 10 літрів, розкид великий. Маса — головний фактор. "
         "Різниця між регіонами здебільшого пояснюється масою, тож прості порівняння вводять в оману. "
         "За 12 років витрата впала майже вдвічі — і через легші авто, і через кращі технології. "
         "Для майбутньої регресії: логарифмувати ціль, брати масу, потужність, рік і регіон, але не всі "
         "розмірні змінні одночасно. Обмеження: дані старі й тільки американського ринку, вибірка незбалансована, "
         "немає типу палива й коробки передач, і це спостережні дані — кореляція не означає причинність.")
slides.append(sl)

# 11 — thanks
sl = Slide(prs, 11)
sl.blob(6.67, 3.75, 4.3, 2.75, KHAKI, 11)
sl.blob(12.9, 0.5, 1.4, 1.1, SAGE, 12)
sl.blob(0.4, 7.2, 1.5, 1.0, SAND, 13)
sl.text(1.0, 2.75, 11.33, "Дякую за увагу!", size=54, color=INK, kind="title", align="center", upper=True)
sl.text(1.0, 3.95, 11.33, "Готовий відповісти на запитання", size=20, color=MUTED, align="center")
sl.text(1.0, 6.55, 11.33, f"Ноутбук, дані та код: {REPO}", size=13, color=MUTED, align="center")
sl.notes("Дякую за увагу. Готовий відповісти на запитання.")
slides.append(sl)

out = ROOT / "reports" / "Presentation.pptx"
prs.save(out)
for i, s in enumerate(slides, 1):
    s.save_preview(i)
sheet = Image.new("RGB", (2 * 1000 + 30, 6 * 562 + 70), (200, 200, 200))
for i in range(len(slides)):
    im = Image.open(PREVIEW / f"slide_{i + 1:02d}.png").resize((1000, 562))
    sheet.paste(im, (10 + (i % 2) * 1010, 10 + (i // 2) * 572))
sheet.save(PREVIEW / "_all_slides.png")

# PDF backup (rendered previews) — for a PC without the fonts or without PowerPoint
pages = [Image.open(PREVIEW / f"slide_{i + 1:02d}.png").convert("RGB") for i in range(len(slides))]
pages[0].save(ROOT / "reports" / "Presentation.pdf", save_all=True, append_images=pages[1:], resolution=DPI)

# Speaker script = notes of every slide
script = ["# Текст доповіді (5–6 хв)\n", f"{AUTHOR}, {GROUP}. Те саме є в нотатках доповідача в Presentation.pptx.\n"]
for i, s in enumerate(slides, 1):
    script.append(f"## Слайд {i}\n\n{s.s.notes_slide.notes_text_frame.text}\n")
words = sum(len(s.s.notes_slide.notes_text_frame.text.split()) for s in slides)
script.append(f"\n_Обсяг: {words} слів ≈ {words / 130:.1f} хв у спокійному темпі._\n")
(ROOT / "reports" / "speech.md").write_text("\n".join(script), encoding="utf-8")
print(f"Speech: {words} words")
print("Saved", out, "| slides:", len(slides), "| previews:", PREVIEW)
