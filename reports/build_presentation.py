"""Builds reports/Presentation.pptx from saved EDA results (outputs/ and data/processed/).

Run after notebooks/EDA.ipynb:  python reports/build_presentation.py
"""
from pathlib import Path

import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches, Pt

ROOT = Path(__file__).resolve().parents[1]
FIG = ROOT / "outputs" / "figures"
TAB = ROOT / "outputs" / "tables"
AUTHOR = "Табачук Олександр, група КН-42"

# ---------- results used on the slides (read from files, not typed by hand) ----------
df = pd.read_csv(ROOT / "data" / "processed" / "auto_mpg_clean.csv")
desc = pd.read_csv(TAB / "descriptive_statistics.csv", index_col=0)
tcorr = pd.read_csv(TAB / "target_correlations.csv", index_col=0)
vif = pd.read_csv(TAB / "vif.csv", index_col=0)["VIF"]
gc = pd.read_csv(TAB / "group_comparison.csv")
yearly = pd.read_csv(TAB / "yearly_trend.csv", index_col=0)

med = desc.loc["mpg", "50%"]
by_origin = df.groupby("origin")["mpg"].mean()
gap_all = by_origin["Японія"] - by_origin["США"]
light = gc[gc["weight_group"] == "легкі"].set_index("origin")["mean"]
mid = gc[gc["weight_group"] == "середні"].set_index("origin")["mean"]
gap_light, gap_mid = light["Японія"] - light["США"], mid["Японія"] - mid["США"]
y0, y1 = yearly["mpg_mean"].iloc[0], yearly["mpg_mean"].iloc[-1]

# ---------- design ----------
NAVY = RGBColor(0x1F, 0x2A, 0x44)
ACCENT = RGBColor(0xC4, 0x4E, 0x52)
GREY = RGBColor(0x5F, 0x66, 0x73)
LIGHT = RGBColor(0xF3, 0xF4, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FONT = "Segoe UI"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
W, H = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def text(slide, x, y, w, h, content, size=18, color=NAVY, bold=False, align=None):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    lines = content if isinstance(content, list) else [content]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.space_after = Pt(8)
        if align is not None:
            p.alignment = align
        for r in p.runs:
            r.font.size, r.font.color.rgb, r.font.bold, r.font.name = Pt(size), color, bold, FONT
    return tb


def rect(slide, x, y, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, h)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()
    return s


def new_slide(title, kicker=None, n=None):
    s = prs.slides.add_slide(BLANK)
    rect(s, 0, 0, Inches(0.18), H, ACCENT)
    if kicker:
        text(s, Inches(0.6), Inches(0.35), Inches(12), Inches(0.4), kicker.upper(), 12, ACCENT, True)
    text(s, Inches(0.6), Inches(0.65), Inches(12.2), Inches(1.0), title, 28, NAVY, True)
    if n:
        text(s, Inches(12.3), Inches(6.95), Inches(0.8), Inches(0.4), str(n), 11, GREY)
    return s


def picture(slide, path, x, y, max_w, max_h):
    iw, ih = Image.open(path).size
    scale = min(max_w / iw, max_h / ih)
    w, h = int(iw * scale), int(ih * scale)
    slide.shapes.add_picture(str(path), x + (max_w - w) // 2, y + (max_h - h) // 2, w, h)


def stat(slide, x, y, value, label, color=ACCENT):
    rect(slide, x, y, Inches(2.8), Inches(1.45), LIGHT)
    text(slide, x + Inches(0.2), y + Inches(0.1), Inches(2.5), Inches(0.75), value, 30, color, True)
    text(slide, x + Inches(0.2), y + Inches(0.85), Inches(2.5), Inches(0.6), label, 13, GREY)


def notes(slide, s):
    slide.notes_slide.notes_text_frame.text = s


# 1. Title
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, NAVY)
rect(s, Inches(0.8), Inches(2.3), Inches(1.2), Inches(0.08), ACCENT)
text(s, Inches(0.8), Inches(2.6), Inches(11.5), Inches(1.6), "Від чого залежить витрата палива автомобіля?", 40, WHITE, True)
text(s, Inches(0.8), Inches(4.0), Inches(11.5), Inches(0.8),
     "Розвідувальний аналіз даних Auto MPG (UCI): 398 авто, 1970–1982", 20, RGBColor(0xC9, 0xCF, 0xDB))
text(s, Inches(0.8), Inches(6.3), Inches(11.5), Inches(0.5), f"Лабораторна робота №2 · {AUTHOR}", 14, RGBColor(0xC9, 0xCF, 0xDB))

# 2. Goal & questions
s = new_slide("Мета: зрозуміти, що визначає економність авто, перед побудовою регресії", "Постановка задачі", 2)
text(s, Inches(0.6), Inches(1.9), Inches(12), Inches(5), [
    "1.  Який вигляд має розподіл витрати палива (mpg)?",
    "2.  Які характеристики найсильніше пов'язані з mpg і чи лінійні ці зв'язки?",
    "3.  Чи відрізняються авто з США, Європи та Японії, якщо врахувати масу?",
    "4.  Як змінилася економність у 1970–1982: менші авто чи кращі технології?",
    "5.  Наскільки фактори дублюють один одного (мультиколінеарність)?",
], 20, NAVY)
notes(s, "Ціль — mpg, милі на галон; більше = економніше. Для інтерпретації перераховано в л/100 км. Модель не будуємо — лише EDA.")

# 3. Data
s = new_slide("Дані: 398 моделей авто ринку США, ціль — mpg (милі/галон)", "Характеристика даних", 3)
stat(s, Inches(0.6), Inches(1.9), f"{len(df)}", "автомобілів, 1970–1982")
stat(s, Inches(3.6), Inches(1.9), "7 + 2", "числових + категоріальних факторів")
stat(s, Inches(6.6), Inches(1.9), f"{int(df['hp_imputed'].sum())} (1,5%)", "пропусків потужності → медіана групи")
stat(s, Inches(9.6), Inches(1.9), "0", "дублікатів, викиди залишено")
text(s, Inches(0.6), Inches(3.8), Inches(12), Inches(3), [
    "Джерело: UCI Machine Learning Repository, ліцензія CC BY 4.0",
    "Фактори: маса, об'єм двигуна, потужність, циліндри, розгін, модельний рік, регіон (США/Європа/Японія)",
    "Очищення: виправлено помилки в марках (chevroelt, toyouta, vokswagen…), додано метричні одиниці",
    "Особливість: «3-циліндрові» авто — це роторні Mazda RX; 63 % вибірки — американські авто",
], 17, GREY)

# 4. Target
s = new_slide(f"Типове авто: медіана {med:.0f} mpg ≈ {235.215 / med:.0f} л/100 км", "RQ1 · Розподіл цілі", 4)
picture(s, FIG / "01_target_distribution.png", Inches(0.5), Inches(1.8), Inches(12.3), Inches(3.6))
text(s, Inches(0.6), Inches(5.6), Inches(12), Inches(1.5), [
    f"Діапазон {desc.loc['mpg', 'min']:.0f}–{desc.loc['mpg', 'max']:.1f} mpg: економність відрізняється у 5 разів; правий «хвіст» — японські та європейські малолітражки",
    "Розподіл не нормальний (Шапіро–Вілк p < 0,001) → для регресії краще log(mpg) або л/100 км",
], 16, GREY)

# 5. Weight
r_w = tcorr.loc["weight_kg", "Pearson r (л/100км)"]
slope = df["l_per_100km"].cov(df["weight_kg"]) / df["weight_kg"].var() * 100
s = new_slide(f"Маса — головний фактор: кожні +100 кг ≈ +{slope:.1f} л/100 км".replace(".", ","), "RQ2 · Двовимірний аналіз", 5)
picture(s, FIG / "05_l100_vs_weight.png", Inches(0.4), Inches(1.7), Inches(7.6), Inches(5.3))
text(s, Inches(8.3), Inches(2.0), Inches(4.7), Inches(5), [
    f"Зв'язок з mpg (ρ Спірмена):",
    *[f"  • {n}: {tcorr.loc[k, 'Spearman ρ (mpg)']:.2f}" for k, n in
      [("weight_kg", "маса"), ("displacement_l", "об'єм двигуна"), ("horsepower", "потужність"),
       ("year", "модельний рік"), ("acceleration", "розгін")]],
    f"У mpg зв'язки нелінійні, а в л/100 км майже лінійні (r = {r_w:.2f})",
], 16, NAVY)
notes(s, f"Нахил лінії ≈ {slope:.2f} л/100 км на 100 кг. Це кореляція, а не доведений причинний ефект.")

# 6. Multicollinearity
s = new_slide("Розмір, потужність і циліндри дублюють один одного", "RQ5 · Багатовимірний аналіз", 6)
picture(s, FIG / "08_correlation_heatmap.png", Inches(0.4), Inches(1.6), Inches(6.8), Inches(5.6))
text(s, Inches(7.5), Inches(2.0), Inches(5.5), Inches(5), [
    "Маса, об'єм, потужність, циліндри: ρ = 0,85–0,95",
    "VIF: " + ", ".join(f"{n} {vif[k]:.0f}" for k, n in
                        [("displacement_l", "об'єм"), ("weight_kg", "маса"), ("cylinders", "циліндри"), ("horsepower", "потужність")]),
    "→ у регресію брати 1–2 з них (напр., масу й потужність)",
    f"Рік слабко пов'язаний з іншими (VIF {vif['year']:.1f}) — окреме джерело інформації",
], 17, NAVY)

# 7. Origin confounding
s = new_slide(f"«Японська перевага» {gap_all:.0f} mpg майже зникає за однакової маси", "RQ3 · Змішувальний фактор", 7)
picture(s, FIG / "10_origin_within_weight.png", Inches(0.4), Inches(1.7), Inches(7.4), Inches(5.2))
text(s, Inches(8.1), Inches(2.0), Inches(4.9), Inches(5), [
    f"Уся вибірка: Японія {by_origin['Японія']:.1f} vs США {by_origin['США']:.1f} mpg",
    f"Легкі авто: різниця лише {gap_light:.1f} mpg",
    f"Середні авто: різниця {gap_mid:.1f} mpg",
    "Японські авто в середньому на ≈ 500 кг легші; важких японських авто у вибірці немає",
    "Маса — змішувальний фактор: просте порівняння за країною вводить в оману",
], 17, NAVY)

# 8. Time trend
s = new_slide(f"За 12 років економність зросла майже вдвічі: {y0:.1f} → {y1:.1f} mpg", "RQ4 · Динаміка", 8)
picture(s, FIG / "07_mpg_by_year.png", Inches(0.5), Inches(1.7), Inches(12.3), Inches(4.4))
text(s, Inches(0.6), Inches(6.2), Inches(12), Inches(1), [
    f"Ріст почався після нафтової кризи 1973 р., найбільший стрибок — 1980 р. Середня маса впала з "
    f"{yearly['weight_kg'].iloc[0]:.0f} до {yearly['weight_kg'].iloc[-1]:.0f} кг",
], 16, GREY)

# 9. Within-group progress
s = new_slide("Прогрес є навіть серед авто однакової маси", "RQ4 · Багатовимірний аналіз", 9)
picture(s, FIG / "11_year_trend_by_weight.png", Inches(0.4), Inches(1.7), Inches(7.6), Inches(5.2))
text(s, Inches(8.3), Inches(2.0), Inches(4.7), Inches(5), [
    "У кожній ваговій групі mpg зростає ≈ 0,7–1,1 на рік",
    "Загальний тренд ≈ 1,2 mpg/рік — більший, бо автопарк ще й полегшав",
    "→ працюють обидва механізми: менші авто + кращі технології",
    "→ рік і маса — взаємодоповнювальні фактори для регресії",
], 17, NAVY)

# 10. Conclusions
s = new_slide("Висновки та обмеження", "Підсумок", 10)
text(s, Inches(0.6), Inches(1.8), Inches(6.2), Inches(5), [
    "Що з'ясували",
    "• Маса — найсильніший фактор економності",
    "• Зв'язки нелінійні; л/100 км або log(mpg) — краща ціль",
    "• Різниця між країнами здебільшого пояснюється масою",
    "• 1970–82: майже ×2 економності — і полегшення, і технології",
    "• Маса/об'єм/потужність/циліндри мультиколінеарні",
], 17, NAVY)
text(s, Inches(7.2), Inches(1.8), Inches(5.8), Inches(5), [
    "Обмеження",
    "• Старі дані, лише ринок США, 398 авто",
    "• 63 % — американські авто; немає важких японських",
    "• Немає типу палива, трансмісії, ціни, продажів",
    "• Спостережні дані: кореляція ≠ причинність",
    "Далі: регресія log(mpg) ~ маса + потужність + рік + регіон",
], 17, GREY)

# 11. Thanks
s = prs.slides.add_slide(BLANK)
rect(s, 0, 0, W, H, NAVY)
text(s, Inches(0.8), Inches(3.0), Inches(11.5), Inches(1.2), "Дякую за увагу!", 44, WHITE, True)
text(s, Inches(0.8), Inches(4.2), Inches(11.5), Inches(0.6), "Запитання?", 22, RGBColor(0xC9, 0xCF, 0xDB))

out = ROOT / "reports" / "Presentation.pptx"
prs.save(out)
print("Saved", out, "slides:", len(prs.slides))
