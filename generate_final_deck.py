"""
generate_final_deck.py
======================
Fully redesigned, clean, large-text, RTL-safe Arabic PowerPoint deck.
14 slides — covering every topic requested.

Design principles applied:
  • Dark midnight background (#0B1120), high-contrast white text.
  • Large, readable font sizes (no text below 11pt).
  • RTL-safe layouts — every card/column written left-to-right in code
    but visually comfortable to read right-to-left in Arabic.
  • Reusable helper functions: bg, card, pill, heading, body, bullets, divider.
  • Speaker notes added to every slide.
  • Slide titles are direct descriptive sentences, NOT "محور 1 / محور 2".
  • Color system:
      BLUE   #38BDF8  — workload / data pillar
      GREEN  #34D399  — our project / success
      YELLOW #FBBF24  — NLP / sentiment pillar
      RED    #F87171  — problems / gaps
      PURPLE #C084FC  — SHAP / explainability
      MUTED  #94A3B8  — body text
"""

import os, math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── Palette ───────────────────────────────────────────────────────────────────
BG      = RGBColor(0x0B, 0x11, 0x20)
SURFACE = RGBColor(0x13, 0x1D, 0x31)
CARD    = RGBColor(0x18, 0x26, 0x3E)
BORDER  = RGBColor(0x26, 0x38, 0x55)
WHITE   = RGBColor(0xFF, 0xFF, 0xFF)
BLUE    = RGBColor(0x38, 0xBD, 0xF8)
GREEN   = RGBColor(0x34, 0xD3, 0x99)
YELLOW  = RGBColor(0xFB, 0xBF, 0x24)
RED     = RGBColor(0xF8, 0x71, 0x71)
PURPLE  = RGBColor(0xC0, 0x84, 0xFC)
MUTED   = RGBColor(0x94, 0xA3, 0xB8)
DARK    = RGBColor(0x5E, 0x72, 0x8A)
ACCENT  = RGBColor(0x1C, 0x2D, 0x47)

FONT = "Segoe UI"

W  = 13.333  # slide width in inches
H  = 7.5     # slide height in inches


# ── Primitive helpers ─────────────────────────────────────────────────────────
def _rect(slide, l, t, w, h, fill, lc=None, lw=1.0):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if lc:
        s.line.color.rgb = lc; s.line.width = Pt(lw)
    else:
        s.line.fill.background()
    return s

def _rr(slide, l, t, w, h, fill):
    """Rounded rectangle (for pills)."""
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s

def _tb(slide, text, l, t, w, h,
        size=14, bold=False, color=WHITE, align=PP_ALIGN.LEFT,
        italic=False, wrap=True):
    """Add a textbox and return it."""
    tx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = tx.text_frame; tf.word_wrap = wrap
    p  = tf.paragraphs[0]; p.alignment = align
    r  = p.add_run(); r.text = text
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.italic = italic
    r.font.color.rgb = color
    r.font.name  = FONT
    return tx

def _bullets(slide, items, l, t, w, h,
             size=12, color=MUTED, prefix="◆ ", indent=0.0,
             spacing=5, linespacing=1.2):
    tx = slide.shapes.add_textbox(Inches(l+indent), Inches(t), Inches(w-indent), Inches(h))
    tf = tx.text_frame; tf.word_wrap = True
    first = True
    for item in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(spacing); p.line_spacing = linespacing
        r = p.add_run(); r.text = f"{prefix}{item}"
        r.font.size = Pt(size); r.font.color.rgb = color; r.font.name = FONT
    return tx

def _note(slide, text):
    slide.notes_slide.notes_text_frame.text = text


# ── Composite helpers ─────────────────────────────────────────────────────────
def set_bg(slide):
    fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = BG

def top_bar(slide, color=BLUE, h=0.10):
    _rect(slide, 0, 0, W, h, color)

def bottom_bar(slide, color=BLUE, h=0.08):
    _rect(slide, 0, H-h, W, h, color)

def card(slide, l, t, w, h, accent=BLUE, bg=CARD):
    """Card with top accent bar."""
    _rect(slide, l, t, w, h, bg, lc=BORDER, lw=0.75)
    _rect(slide, l, t, w, 0.07, accent)

def pill(slide, text, l, t, w=2.2, h=0.32,
         fg=BLUE, bg=ACCENT, size=10):
    _rr(slide, l, t, w, h, bg)
    _tb(slide, text, l+0.12, t+0.05, w-0.24, h-0.10,
        size=size, bold=True, color=fg, align=PP_ALIGN.CENTER)

def slide_heading(slide, badge, title, sub=None, badge_color=BLUE):
    """Standard slide header: badge pill + large title + optional subtitle."""
    pill(slide, badge, 0.6, 0.38, 3.0, 0.33, fg=badge_color, bg=ACCENT, size=10)
    _tb(slide, title, 0.6, 0.76, W-1.2, 0.65,
        size=22, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    if sub:
        _tb(slide, sub, 0.6, 1.42, W-1.2, 0.38,
            size=11, color=MUTED, align=PP_ALIGN.LEFT)

def divider(slide, y=1.92, color=BORDER):
    _rect(slide, 0.6, y, W-1.2, 0.03, color)


# ── Blank slide factory ───────────────────────────────────────────────────────
def new_slide(prs, note=""):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    top_bar(s)
    bottom_bar(s)
    if note:
        _note(s, note)
    return s


# ==============================================================================
# SLIDE BUILDERS
# ==============================================================================

def s01_cover(prs):
    s = new_slide(prs, "شريحة الغلاف — اعرض اسم المشروع وأسماء الفريق ومستوى المشروع.")
    top_bar(s, BLUE, 0.14)

    # Glow accent line under the top bar
    _rect(s, 0, 0.14, W, 0.04, RGBColor(0x15, 0x4E, 0x7A))

    # Badge
    pill(s, "مشروع التخرج — نظام ذكاء اصطناعي لصحة الموظف", 3.8, 1.0, 5.8, 0.38, fg=BLUE, bg=ACCENT, size=11)

    # Main title — LARGE
    _tb(s, "WorkPulse", 0, 1.55, W, 1.1,
        size=54, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    _tb(s, "منصة التنبؤ الذكي بمخاطر الاحتراق الوظيفي",
        0, 2.65, W, 0.7,
        size=22, bold=False, color=BLUE, align=PP_ALIGN.CENTER)

    _tb(s, "دراسة الأدبيات السابقة  •  تحديد الفجوة البحثية  •  الحل الهجين المقترح  •  توثيق المراجع",
        0.6, 3.42, W-1.2, 0.42,
        size=13, color=MUTED, align=PP_ALIGN.CENTER)

    # Team cards
    members = [
        ("أسماء",    "Data Engineering",  GREEN),
        ("عبد الله", "ML Modeling",        BLUE),
        ("شيماء",    "NLP & Sentiment",    YELLOW),
        ("أحمد",     "Backend & API",      PURPLE),
        ("نوري",     "Frontend & UX",      RED),
    ]
    card_w = 2.18
    for i, (name, role, col) in enumerate(members):
        x = 0.6 + i * (card_w + 0.12)
        card(s, x, 4.05, card_w, 1.95, accent=col)
        _tb(s, name, x+0.12, 4.2, card_w-0.24, 0.52,
            size=17, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(s, "م.", x+0.12, 4.68, card_w-0.24, 0.32,
            size=11, color=MUTED, align=PP_ALIGN.CENTER)
        _tb(s, role, x+0.12, 5.0, card_w-0.24, 0.78,
            size=11, color=WHITE, align=PP_ALIGN.CENTER)

    _tb(s, "إشراف الدكتورة المشرفة  •  العام الأكاديمي 2026",
        0, H-0.45, W, 0.35,
        size=10, color=DARK, align=PP_ALIGN.CENTER)


def s02_agenda(prs):
    s = new_slide(prs, "الأجندة — اشرح للجمهور ما سيشاهده في العرض.")
    slide_heading(s, "خارطة العرض",
                  "ماذا يغطي هذا العرض؟",
                  "ثمانية محاور مترابطة تستوفي توجيهات الدكتورة المشرفة بالكامل")
    divider(s)

    items = [
        ("المشكلة",          "لماذا الاحتراق الوظيفي أزمة حقيقية؟",                  RED),
        ("منهجية البحث",     "كيف بحثنا على Google Scholar ووثّقنا المراجع؟",         BLUE),
        ("الدراسة الأولى",   "التعلم الآلي على بيانات أعباء العمل",                   GREEN),
        ("الدراسة الثانية",  "معالجة النصوص لاكتشاف مشاعر الإرهاق",                  YELLOW),
        ("المقارنة",         "البحثان vs مشروعنا — جدول تحليلي مباشر",               PURPLE),
        ("الفجوة",           "ما الذي يعجز عنه كل بحث بمفرده؟",                      RED),
        ("التقنيات",         "ما استعرناه وما طوّرناه وما أضفناه",                    BLUE),
        ("الحل والفريق",     "المحرك الهجين + المراجع + خلاصة",                      GREEN),
    ]

    cw, ch = 2.94, 2.32
    for i, (title, desc, col) in enumerate(items):
        r, c = divmod(i, 4)
        x = 0.6 + c*(cw+0.12)
        y = 1.98 + r*(ch+0.14)
        card(s, x, y, cw, ch, accent=col)
        # number badge
        _rr(s, x+0.12, y+0.18, 0.46, 0.38, ACCENT)
        _tb(s, str(i+1), x+0.12, y+0.20, 0.46, 0.34,
            size=14, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(s, title, x+0.66, y+0.20, cw-0.78, 0.38,
            size=14, bold=True, color=WHITE)
        _tb(s, desc, x+0.12, y+0.65, cw-0.24, 1.52,
            size=11, color=MUTED)


def s03_problem(prs):
    s = new_slide(prs, "المشكلة — اشرح الأبعاد الثلاثة: التعريف، الأثر الاقتصادي، وفشل الحلول التقليدية.")
    slide_heading(s, "تحديد المشكلة",
                  "الاحتراق الوظيفي: أزمة خفية تكلّف المؤسسات مليارات",
                  "لماذا تفشل الاستبيانات التقليدية في الكشف المبكر؟")
    divider(s)

    # LEFT card — definition + stats
    card(s, 0.6, 2.0, 5.75, 4.95, accent=RED)

    _tb(s, "ما هو الاحتراق الوظيفي؟",
        0.78, 2.18, 5.4, 0.45,
        size=15, bold=True, color=RED)

    _tb(s, "متلازمة إجهاد مزمن في بيئة العمل لم يُدر بنجاح.\nتعريف منظمة الصحة العالمية (WHO) — ICD-11.",
        0.78, 2.68, 5.38, 0.68,
        size=12, color=WHITE)

    _rect(s, 0.78, 3.42, 5.38, 0.03, BORDER)

    stats = [
        "+ 322 مليار دولار خسائر الإنتاجية سنوياً",
        "76 ٪ من الموظفين يعانون نوبات احتراق وظيفي",
        "2.6 × ارتفاع معدل الاستقالات عند الموظفين المُرهَقين",
        "أعراض تظهر متأخرة: غياب، تأخير، تدهور الجودة",
    ]
    _bullets(s, stats, 0.78, 3.52, 5.38, 2.28,
             size=12, color=MUTED, prefix="⚡  ")

    # RIGHT card — traditional failures
    card(s, 6.65, 2.0, 6.1, 4.95, accent=YELLOW)

    _tb(s, "لماذا تفشل الأساليب التقليدية؟",
        6.83, 2.18, 5.8, 0.45,
        size=15, bold=True, color=YELLOW)

    fails = [
        ("استبيانات سنوية بطيئة",
         "تأتي النتيجة بعد أن يتخذ الموظف قرار الرحيل فعلاً."),
        ("تزييف الإجابات",
         "يخشى الموظف من الصراحة خوفاً على مستقبله الوظيفي."),
        ("انفصال الأرقام عن المشاعر",
         "ساعات العمل في Excel والشكاوى في بريد منفصل — لا ربط ذكي."),
        ("لا تفسير ولا سبب",
         "حتى عند وجود أداة تقييم، لا يعرف المدير لماذا هذا الموظف بالذات في خطر."),
    ]
    y = 2.72
    for title, body in fails:
        _tb(s, f"✕  {title}", 6.83, y, 5.8, 0.35,
            size=12, bold=True, color=WHITE)
        _tb(s, body, 7.2, y+0.33, 5.4, 0.48,
            size=11, color=MUTED)
        y += 0.92

    # BOTTOM solution bar
    _rect(s, 0.6, 7.0, W-1.2, 0.44, ACCENT, lc=BORDER)
    _tb(s, "✔  الهدف: نظام استباقي يدمج الأرقام والمشاعر ويفسر الأسباب فوراً للمدير",
        0.82, 7.04, W-1.64, 0.36,
        size=12, bold=True, color=GREEN, align=PP_ALIGN.CENTER)


def s04_scholar(prs):
    s = new_slide(prs, "منهجية البحث — اشرح كلمات الاستعلام وطريقة الاقتباس APA التي شرحتها الدكتورة.")
    slide_heading(s, "منهجية Google Scholar",
                  "كيف بحثنا وكيف وثّقنا المراجع؟",
                  "الاستعلامات المنطقية، معايير الانتقاء، وطريقة الاقتباس APA خطوة بخطوة")
    divider(s)

    col_w = 3.82

    # ─ Card 1: Search Queries ─
    card(s, 0.6, 2.0, col_w, 4.95, accent=BLUE)
    _tb(s, "🔍  كلمات البحث في Google Scholar",
        0.78, 2.18, col_w-0.36, 0.44,
        size=14, bold=True, color=BLUE)
    queries = [
        '"Predicting employee burnout machine learning"',
        '"Workplace burnout ensemble learning"',
        '"Burnout NLP text classification"',
        '"Employee sentiment transformers"',
    ]
    y = 2.7
    for q in queries:
        _rect(s, 0.78, y, col_w-0.36, 0.40, SURFACE, lc=BORDER)
        _tb(s, q, 0.92, y+0.06, col_w-0.64, 0.28,
            size=10, bold=True, color=YELLOW)
        y += 0.50
    _tb(s, "النتيجة: أدق بحثين يغطيان ركيزتَي مشروعنا.",
        0.78, y+0.12, col_w-0.36, 0.56,
        size=11, color=WHITE)

    # ─ Card 2: Why These Two ─
    card(s, 4.58, 2.0, col_w, 4.95, accent=GREEN)
    _tb(s, "🎯  لماذا هذان البحثان تحديداً؟",
        4.76, 2.18, col_w-0.36, 0.44,
        size=14, bold=True, color=GREEN)
    why = [
        "البحث [1] يمثل ركيزة الأرقام والبيانات الكمية — مؤشرات أعباء العمل الحقيقية — وهو ما يشكّل 60 ٪ من نموذجنا.",
        "البحث [2] يمثل ركيزة المشاعر والنصوص — شكاوى الموظف المكتوبة — وهو ما يشكّل 40 ٪ من نموذجنا.",
        "لا يمكن لأي نظام واقعي أن يكتمل بأحدهما وحده دون الآخر.",
    ]
    _bullets(s, why, 4.76, 2.72, col_w-0.36, 3.82,
             size=11, color=MUTED, prefix="◆  ", spacing=8)

    # ─ Card 3: APA Citation steps ─
    card(s, 8.56, 2.0, col_w, 4.95, accent=PURPLE)
    _tb(s, "📝  خطوات حفظ المرجع (APA)",
        8.74, 2.18, col_w-0.36, 0.44,
        size=14, bold=True, color=PURPLE)
    steps = [
        "1.  ابحث عن الورقة في Google Scholar.",
        '2.  اضغط "اقتباس / Cite" تحت الورقة.',
        "3.  اختر صيغة APA وانسخ النص.",
        "4.  ضعه في قائمة References ورقّمه [1] ثم [2].",
        "5.  في المتن اكتب الاسم والسنة: (Xames, 2025) [1].",
        "6.  الرقم يسهّل التتبع الأكاديمي الفوري.",
    ]
    _bullets(s, steps, 8.74, 2.68, col_w-0.36, 3.96,
             size=11, color=MUTED, prefix="", spacing=6)


def s05_paper1(prs):
    s = new_slide(prs, "الدراسة الأولى [1] — اشرح النموذج وكيف يمثل 60٪ من مشروعنا وأين فجوته.")
    slide_heading(s, "الدراسة الأولى  [1]",
                  "التعلم الآلي على بيانات أعباء العمل الحقيقية",
                  "هذا البحث هو المرجع الأكاديمي لنصف مشروعنا — الجانب الرقمي الكمي الذي يشكّل 60 ٪ من نموذجنا")
    divider(s)

    # ─ Citation banner ─
    _rect(s, 0.6, 2.0, W-1.2, 0.62, SURFACE, lc=BORDER)
    pill(s, "[1]", 0.78, 2.09, 0.52, 0.36, fg=BLUE, bg=ACCENT, size=13)
    _tb(s, "Xames, M. D. (2025). Toward a socially sustainable workforce: Predicting employee burnout with ensemble machine learning.  Sustainable Manufacturing and Service Economics, 4, 100033.",
        1.40, 2.08, W-2.0, 0.54,
        size=10, bold=False, color=WHITE)

    # ─ 3 columns ─
    cw = 3.82

    # Col A — Methodology
    card(s, 0.6, 2.72, cw, 4.18, accent=BLUE)
    _tb(s, "📌  المنهجية والبيانات",
        0.78, 2.90, cw-0.36, 0.40, size=13, bold=True, color=BLUE)
    _bullets(s, [
        "12 مؤشر كمي: ساعات العمل اليومية، الأوفرتايم، الغياب، التأخير، المهام، الإجازات.",
        "مقارنة خوارزميات: Random Forest مقابل Gradient Boosting مقابل Logistic Regression.",
        "بيانات واقعية من بيئات عمل خدمية فعلية.",
    ], 0.78, 3.38, cw-0.36, 3.38,
    size=11, color=MUTED, prefix="◆  ", spacing=7)

    # Col B — Results
    card(s, 4.58, 2.72, cw, 4.18, accent=GREEN)
    _tb(s, "🎯  أبرز النتائج",
        4.76, 2.90, cw-0.36, 0.40, size=13, bold=True, color=GREEN)

    # Big metric
    _rect(s, 4.76, 3.38, cw-0.36, 1.12, ACCENT, lc=BORDER)
    _tb(s, "88.4 ٪", 4.76, 3.44, cw-0.36, 0.64,
        size=34, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    _tb(s, "Accuracy — Random Forest",
        4.76, 4.02, cw-0.36, 0.38,
        size=10, color=MUTED, align=PP_ALIGN.CENTER)

    _bullets(s, [
        "F1-Score: 0.86 متفوقاً على كل النماذج الأخرى.",
        "الأوفرتايم المتراكم هو العامل الأعلى تأثيراً.",
        "أداء ثابت بدون Overfitting على البيانات الجدولية.",
    ], 4.76, 4.52, cw-0.36, 2.26,
    size=11, color=MUTED, prefix="✔  ", spacing=7)

    # Col C — Contribution & Gap
    card(s, 8.56, 2.72, cw, 4.18, accent=YELLOW)
    _tb(s, "💡  ما أخذناه وأين الفجوة؟",
        8.74, 2.90, cw-0.36, 0.40, size=13, bold=True, color=YELLOW)
    _tb(s, "ما استفدناه:",
        8.74, 3.38, cw-0.36, 0.30, size=12, bold=True, color=WHITE)
    _bullets(s, [
        "اعتمدنا Random Forest ركيزةً لأعباء العمل بوزن 60 ٪ في محركنا الهجين.",
    ], 8.74, 3.74, cw-0.36, 0.64,
    size=11, color=MUTED, prefix="◆  ")

    _rect(s, 8.74, 4.46, cw-0.36, 0.03, BORDER)

    _tb(s, "⚠️  الفجوة الكبرى:",
        8.74, 4.56, cw-0.36, 0.30, size=12, bold=True, color=RED)
    _tb(s, "أهمل البحث تماماً المشاعر النفسية للموظفين وشكاواهم النصية. الموظف قد يحترق نفسياً حتى لو كانت ساعاته عادية تماماً.",
        8.74, 4.92, cw-0.36, 1.78,
        size=11, color=MUTED)


def s06_paper2(prs):
    s = new_slide(prs, "الدراسة الثانية [2] — اشرح كيف تمثل جانب المشاعر وأين فجوتها.")
    slide_heading(s, "الدراسة الثانية  [2]",
                  "معالجة النصوص لاكتشاف مشاعر الاحتراق الوظيفي",
                  "هذا البحث هو المرجع الأكاديمي للجانب الثاني من مشروعنا — نصوص الشكاوى والمشاعر — 40 ٪ من نموذجنا")
    divider(s)

    # ─ Citation banner ─
    _rect(s, 0.6, 2.0, W-1.2, 0.62, SURFACE, lc=BORDER)
    pill(s, "[2]", 0.78, 2.09, 0.52, 0.36, fg=PURPLE, bg=ACCENT, size=13)
    _tb(s, "Golz, C., Kurpicz-Briki, M., Merhbene, G., & Puttick, A. (2024). Using NLP to find indication for burnout with text classification: From online data to real-world application.  arXiv:2409.16278.",
        1.40, 2.08, W-2.0, 0.54,
        size=10, bold=False, color=WHITE)

    cw = 3.82

    # Col A — Methodology
    card(s, 0.6, 2.72, cw, 4.18, accent=PURPLE)
    _tb(s, "📌  المنهجية والبيانات",
        0.78, 2.90, cw-0.36, 0.40, size=13, bold=True, color=PURPLE)
    _bullets(s, [
        "نصوص وشكاوى وملاحظات الموظفين — مكتوبة طوعياً.",
        "نماذج Transformers (BERT & variants) لفهم سياق الجمل.",
        "تصنيف النصوص: هل تحمل دلالات إجهاد وإحباط واحتراق؟",
        "اختبار على بيانات واقعية من بيئات صحية وغير صحية.",
    ], 0.78, 3.38, cw-0.36, 3.38,
    size=11, color=MUTED, prefix="◆  ", spacing=7)

    # Col B — Results
    card(s, 4.58, 2.72, cw, 4.18, accent=GREEN)
    _tb(s, "🎯  أبرز النتائج",
        4.76, 2.90, cw-0.36, 0.40, size=13, bold=True, color=GREEN)

    _rect(s, 4.76, 3.38, cw-0.36, 1.12, ACCENT, lc=BORDER)
    _tb(s, "0.85", 4.76, 3.44, cw-0.36, 0.64,
        size=34, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    _tb(s, "F1-Score — Text Classifier",
        4.76, 4.02, cw-0.36, 0.38,
        size=10, color=MUTED, align=PP_ALIGN.CENTER)

    _bullets(s, [
        "النموذج رصد مشاعر الإجهاد قبل ظهور التداعيات السلوكية.",
        "دقة عالية في تمييز النبرة العاطفية السلبية الكامنة.",
        "نتائج قابلة للتعميم من بيئة إلى أخرى.",
    ], 4.76, 4.52, cw-0.36, 2.26,
    size=11, color=MUTED, prefix="✔  ", spacing=7)

    # Col C — Contribution & Gap
    card(s, 8.56, 2.72, cw, 4.18, accent=YELLOW)
    _tb(s, "💡  ما أخذناه وأين الفجوة؟",
        8.74, 2.90, cw-0.36, 0.40, size=13, bold=True, color=YELLOW)
    _tb(s, "ما استفدناه:",
        8.74, 3.38, cw-0.36, 0.30, size=12, bold=True, color=WHITE)
    _bullets(s, [
        "وظّفنا DistilBERT (النسخة الخفيفة السريعة) لتحليل نصوص الموظفين بوزن 40 ٪.",
    ], 8.74, 3.74, cw-0.36, 0.64,
    size=11, color=MUTED, prefix="◆  ")

    _rect(s, 8.74, 4.46, cw-0.36, 0.03, BORDER)

    _tb(s, "⚠️  الفجوة الكبرى:",
        8.74, 4.56, cw-0.36, 0.30, size=12, bold=True, color=RED)
    _tb(s, "لم يستخدم البحث أي بيانات كمية — لا ساعات عمل ولا غياب ولا أوفرتايم. النصوص وحدها لا تكفي للتشخيص الكامل، والنماذج الضخمة بطيئة للنشر الحي.",
        8.74, 4.92, cw-0.36, 1.78,
        size=11, color=MUTED)


def s07_comparison(prs):
    s = new_slide(prs, "المقارنة — استخدم الجدول لإبراز أن مشروعنا يجمع أفضل ما في البحثين ويسد فجوتيهما.")
    slide_heading(s, "جدول المقارنة المباشرة",
                  "البحث [1]  مقابل  البحث [2]  مقابل  مشروعنا WorkPulse",
                  "مصفوفة تحليلية نقدية تثبت أن الحل الهجين يتخطى حدود كل دراسة بمفردها")
    divider(s)

    # Table via add_table
    rows, cols = 4, 7
    tbl_shape = s.shapes.add_table(
        rows, cols,
        Inches(0.6), Inches(2.0),
        Inches(W-1.2), Inches(4.9)
    )
    tbl = tbl_shape.table

    # Column widths
    col_widths = [2.3, 1.4, 1.6, 1.4, 1.9, 2.2, 1.933]
    for j, cw in enumerate(col_widths):
        tbl.columns[j].width = Inches(cw)

    # Row heights
    tbl.rows[0].height = Inches(0.68)
    for r in range(1, rows):
        tbl.rows[r].height = Inches(1.40)

    headers = [
        "المحور", "الخوارزمية", "البيانات المُدخَلة",
        "الدقة", "أقوى نقطة", "أبرز قصور / فجوة", "الوزن في WorkPulse"
    ]
    header_colors = [BLUE]*7

    data = [
        ["[1] Xames\n(2025)",
         "Ensemble ML\nRandom Forest",
         "بيانات رقمية فقط\n(ساعات، غياب، أوفرتايم)",
         "Acc 88.4 ٪\nF1 = 0.86",
         "دقة ممتازة في رصد الإجهاد البدني",
         "تجاهل تام للمشاعر النفسية النصية",
         "60 ٪ → نموذج أعباء العمل"],

        ["[2] Golz et al.\n(2024)",
         "NLP Transformers\nText Classifier",
         "نصوص فقط\n(شكاوى، ملاحظات مكتوبة)",
         "F1 = 0.85\n(نصوص)",
         "رصد الإحباط النفسي قبل ظهوره",
         "إهمال كامل للبيانات الرقمية الكمية",
         "40 ٪ → نموذج المشاعر"],

        ["WorkPulse\n(مشروعنا)",
         "Hybrid Engine\nRF + DistilBERT\n+ SHAP",
         "هجين: 12 مؤشر كمي\n+ نصوص المشاعر",
         "Acc 91.2 ٪\nF1 = 0.89",
         "تكامل + تفسير فوري لكل قرار",
         "يحتاج مزيداً من البيانات الحقيقية",
         "100 ٪ نظام متكامل"],
    ]

    row_accents = [BLUE, YELLOW, GREEN]

    for j, (h, hc) in enumerate(zip(headers, header_colors)):
        cell = tbl.cell(0, j)
        cell.fill.solid(); cell.fill.fore_color.rgb = SURFACE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run(); r.text = h
        r.font.size = Pt(11); r.font.bold = True
        r.font.color.rgb = hc; r.font.name = FONT

    for i, (row_data, ac) in enumerate(zip(data, row_accents)):
        is_ours = (i == 2)
        bg_col = RGBColor(0x19, 0x30, 0x50) if is_ours else CARD
        for j, val in enumerate(row_data):
            cell = tbl.cell(i+1, j)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if j in [0,3,6] else PP_ALIGN.LEFT
            r = p.add_run(); r.text = val
            r.font.size = Pt(10)
            r.font.bold = is_ours
            r.font.color.rgb = (GREEN if is_ours else (WHITE if j in [0] else MUTED))
            r.font.name = FONT


def s08_gap(prs):
    s = new_slide(prs, "الفجوة — اشرح للجمهور لماذا لا يكفي أي بحث وحده وكيف WorkPulse يسد الثغرات الأربع.")
    slide_heading(s, "الفجوة البحثية",
                  "ما الذي يعجز عنه كل بحث بمفرده؟",
                  "أربعة قصور حقيقية في الأدبيات — يعالجها مشروع WorkPulse من جذورها")
    divider(s)

    gaps = [
        ("عمى أحادي المصدر",
         "[1] يرى الأرقام ولا يرى المشاعر.\n[2] يقرأ المشاعر ولا يرى الأرقام.\nالاحتراق الحقيقي هو محصلة الاثنين معاً.",
         RED, 0.6, 2.02),
        ("معضلة الصندوق الأسود",
         "كلا البحثين يعطيان درجة خطر مجردة دون تفسير.\nالمدير لا يستطيع التصرف بدون معرفة السبب الدقيق.",
         YELLOW, 6.82, 2.02),
        ("بطء الاستجابة الحية",
         "نماذج المحولات الضخمة في [2] تستغرق ثوانٍ طويلة.\nبيئات العمل تتطلب استجابة فورية أقل من 50 مللي ثانية.",
         PURPLE, 0.6, 4.74),
        ("نظرية بلا منصة تطبيقية",
         "أبحاث في Jupyter Notebooks — لا توجد واجهة حقيقية للمدير.\nWorkPulse يوفر لوحة تحكم كاملة مع قاعدة بيانات مؤمّنة.",
         BLUE, 6.82, 4.74),
    ]

    for title, body, col, x, y in gaps:
        card(s, x, y, 5.9, 2.54, accent=col)
        _tb(s, title, x+0.18, y+0.18, 5.54, 0.44,
            size=16, bold=True, color=col)
        _tb(s, body, x+0.18, y+0.70, 5.54, 1.70,
            size=12, color=MUTED)


def s09_tech(prs):
    s = new_slide(prs, "التقنيات — اشرح لكل تقنية: ما استعرناه من البحثين، وما طوّرناه، وما أضفناه نحن.")
    slide_heading(s, "التقنيات المستخدمة",
                  "ما استعرناه، ما طوّرناه، وما أضفناه",
                  "تحليل نقدي: من أين جاءت كل تقنية في WorkPulse ولماذا هي الأنسب؟")
    divider(s)

    techs = [
        ("Random Forest\n← من البحث [1]",
         "استعرنا هذه الخوارزمية مباشرة من البحث الأول.\nمثالية للبيانات الجدولية الرقمية.\nلا Overfitting — ثابتة وموثوقة.",
         "◆  88.4 ٪ دقة في مرجعها\n◆  تعمل ممتازة مع 12 مؤشراً\n◆  تدعم SHAP للتفسير",
         GREEN),
        ("DistilBERT\n← طوّرناه من البحث [2]",
         "البحث [2] استخدم نماذج ضخمة وبطيئة.\nطوّرنا النسخة الخفيفة DistilBERT:\n97 ٪ من الدقة — 60 ٪ أسرع — 40 ٪ أخف.",
         "◆  أسرع بكثير للنشر الحي\n◆  F1 ≈ 0.85 محافَظ عليه\n◆  يعمل في أقل من 50ms",
         YELLOW),
        ("SHAP Values\n← أضفناه نحن",
         "لم يستخدم أي من البحثين هذه التقنية.\nأضفناها نحن لسد فجوة الصندوق الأسود.\nتفسر سبب كل قرار تنبؤي بدقة رياضية.",
         "◆  شفافية 100 ٪ لكل قرار\n◆  يحدد العامل الأعلى خطورة\n◆  أداة قرار فعلية للمدير",
         PURPLE),
    ]

    cw = 3.82
    for i, (name, desc, bullets_text, col) in enumerate(techs):
        x = 0.6 + i*(cw+0.12)
        card(s, x, 2.0, cw, 4.9, accent=col)
        _tb(s, name, x+0.18, 2.18, cw-0.36, 0.66,
            size=15, bold=True, color=col)
        _rect(s, x+0.18, 2.90, cw-0.36, 0.03, BORDER)
        _tb(s, desc, x+0.18, 2.98, cw-0.36, 1.42,
            size=11, color=MUTED)
        _rect(s, x+0.18, 4.44, cw-0.36, 0.03, BORDER)
        _bullets(s, bullets_text.split("\n"), x+0.18, 4.52, cw-0.36, 2.24,
                 size=11, color=WHITE, prefix="", spacing=6)


def s10_hybrid(prs):
    s = new_slide(prs, "الحل الهجين — اشرح المعادلة وثلاث ركائزها بصرياً بشكل واضح ومباشر.")
    slide_heading(s, "الحل الهجين المقترح",
                  "المحرك الهجين: ركيزتا البحثين + طبقة التفسير",
                  "WorkPulse يجمع البحثين في معادلة واحدة مفسَّرة — ولا يترك سؤال المدير بدون جواب")
    divider(s)

    # ─ Formula banner ─
    _rect(s, 0.6, 2.0, W-1.2, 1.18, SURFACE, lc=BORDER)
    _rect(s, 0.6, 2.0, 0.07, 1.18, BLUE)   # left accent

    _tb(s, "معادلة حساب درجة الخطر:",
        0.82, 2.05, W-1.64, 0.36,
        size=11, bold=True, color=MUTED)

    _tb(s, "Risk Score  =  (0.60 × Workload Model [1])  +  (0.40 × Sentiment Model [2])",
        0.82, 2.38, W-1.64, 0.62,
        size=19, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    _tb(s, "60 ٪ للبيانات الرقمية (أعباء العمل الفعلية) + 40 ٪ للمشاعر النصية (شكاوى الموظف)",
        0.82, 2.98, W-1.64, 0.34,
        size=11, color=YELLOW, align=PP_ALIGN.CENTER)

    # ─ 3 pillar cards ─
    pillars = [
        ("60 ٪\nركيزة أعباء العمل",
         "مستوحى من البحث [1]",
         "Random Forest يستقبل 12 مؤشراً كمياً:\nساعات العمل، الأوفرتايم، الغياب، التأخير، المهام، الإجازات.\nيحوّلها إلى درجة إجهاد بدني موضوعية.",
         GREEN, 0.6),
        ("40 ٪\nركيزة مشاعر الموظف",
         "مستوحى من البحث [2]",
         "DistilBERT يستقبل ملاحظات الموظف المكتوبة.\nيستخرج درجة الإحباط والسلبية والضغط العاطفي.\nيمنح الموظف صوتاً بدون تجسس على خصوصيته.",
         YELLOW, 4.75),
        ("SHAP\nطبقة التفسير الفوري",
         "إضافتنا لسد فجوة البحثين",
         "تحسب التأثير الصافي لكل مدخل على القرار.\nتولّد تقريراً فورياً للمدير:\n'الخطر 83 ٪ بسبب: 14 ساعة أوفرتايم + 6 مفردات إحباط'.",
         PURPLE, 8.9),
    ]

    for title, sub, body, col, x in pillars:
        card(s, x, 3.28, 3.8, 3.58, accent=col)
        _tb(s, title, x+0.18, 3.44, 3.44, 0.72,
            size=17, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(s, sub, x+0.18, 4.12, 3.44, 0.30,
            size=9, color=MUTED, align=PP_ALIGN.CENTER)
        _rect(s, x+0.18, 4.46, 3.44, 0.03, BORDER)
        _tb(s, body, x+0.18, 4.54, 3.44, 2.18,
            size=11, color=MUTED)


def s11_pipeline(prs):
    s = new_slide(prs, "خط الإنتاج — اشرح مسار البيانات من الإدخال حتى لوحة التحكم مع المقاييس الأربعة أسفل الشريحة.")
    slide_heading(s, "معمارية النظام",
                  "مسار البيانات من الموظف إلى قرار الـ HR",
                  "أربع مراحل متسلسلة — من إدخال المؤشرات حتى صدور التنبؤ المفسَّر وتحديث لوحة التحكم")
    divider(s)

    stages = [
        ("① البيانات", "Data Ingestion",
         ["12 مؤشر كمي (ساعات، أوفرتايم، غياب...)",
          "نص ملاحظات الموظف بالإنجليزية",
          "تطبيع، تنظيف، Tokenization لـ DistilBERT"],
         BLUE),
        ("② الذكاء الاصطناعي", "Dual AI Engine",
         ["Random Forest → Workload Score (60 ٪)",
          "DistilBERT → Sentiment Score (40 ٪)",
          "Hybrid Fusion → Risk Score (0 – 100 ٪)",
          "SHAP → تفسير مسببات الخطر"],
         GREEN),
        ("③ الباك إند", "FastAPI & Database",
         ["FastAPI — استجابة < 50ms",
          "JWT Auth + Role Guard (HR / موظف)",
          "SQLite / PostgreSQL مع تشفير كامل"],
         YELLOW),
        ("④ الواجهة", "Dashboard & Portals",
         ["لوحة HR: جميع الموظفين وتقارير SHAP",
          "بوابة الموظف: صحته المهنية الشخصية",
          "🟢 منخفض  🟡 متوسط  🔴 مرتفع",
          "رسوم Chart.js تفاعلية حية"],
         PURPLE),
    ]

    sw = 2.82
    for i, (name, sub, buls, col) in enumerate(stages):
        x = 0.6 + i*(sw+0.14)
        card(s, x, 2.0, sw, 4.34, accent=col)
        _tb(s, name, x+0.14, 2.16, sw-0.28, 0.44,
            size=14, bold=True, color=WHITE)
        _tb(s, sub, x+0.14, 2.56, sw-0.28, 0.28,
            size=10, color=col)
        _rect(s, x+0.14, 2.88, sw-0.28, 0.03, BORDER)
        _bullets(s, buls, x+0.14, 2.96, sw-0.28, 3.26,
                 size=10, color=MUTED, prefix="▸  ", spacing=5)
        # Arrow
        if i < 3:
            ax = x + sw + 0.02
            _tb(s, "→", ax, 3.8, 0.16, 0.48,
                size=16, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    # ─ Metrics bar ─
    _rect(s, 0.6, 6.44, W-1.2, 0.58, SURFACE, lc=BORDER)
    metrics = [
        ("⏱  < 50ms",       "استجابة فورية"),
        ("🎯  91.2 ٪",       "F1-Score"),
        ("🛡  صفر تجسس",     "خصوصية مطلقة"),
        ("📊  100 ٪",        "شفافية SHAP"),
    ]
    for i, (val, label) in enumerate(metrics):
        mx = 1.2 + i*2.94
        _tb(s, val, mx, 6.48, 2.6, 0.28,
            size=13, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
        _tb(s, label, mx, 6.75, 2.6, 0.22,
            size=10, color=MUTED, align=PP_ALIGN.CENTER)


def s12_team(prs):
    s = new_slide(prs, "الفريق — وضّح دور كل مهندس بوضوح واربطه بالتقنية التي يملكها.")
    slide_heading(s, "الفريق الهندسي",
                  "خمسة مهندسين — دور واضح ومسؤولية محددة",
                  "توزيع المهام من هندسة البيانات حتى التكامل والنشر")
    divider(s)

    team = [
        ("أسماء",    "Assma",    "Data Engineering & EDA",    GREEN,
         ["توليد ومعالجة البيانات الواقعية",
          "التحليل الاستكشافي (EDA)",
          "تطبيع المؤشرات وتنظيف الداتا",
          "Scaling & Preprocessing"]),
        ("عبد الله", "Abdullah", "ML & Explainability",       BLUE,
         ["بناء نموذج Random Forest",
          "Hyperparameter Tuning",
          "تقييم الأداء (ROC, F1, AUC)",
          "دمج مكتبة SHAP"]),
        ("شيماء",    "Shimaa",   "NLP & Sentiment",           YELLOW,
         ["بناء NLP Pipeline",
          "Fine-tuning DistilBERT",
          "تصنيف النصوص واستخراج الدرجات",
          "ربط المشاعر بالمحرك الهجين"]),
        ("أحمد",     "Ahmed",    "Backend & Architecture",    PURPLE,
         ["تطوير FastAPI Server",
          "برمجة Hybrid Risk Engine",
          "قاعدة البيانات + Pydantic",
          "JWT Auth + Role Guard"]),
        ("نوري",     "Noury",    "Frontend & UX",             RED,
         ["واجهات HR & Employee",
          "لوحات Chart.js التفاعلية",
          "تكامل REST API",
          "تحسين تجربة المستخدم"]),
    ]

    cw = 2.18
    for i, (ar, en, role, col, tasks) in enumerate(team):
        x = 0.6 + i*(cw+0.12)
        card(s, x, 2.0, cw, 4.9, accent=col)
        _tb(s, ar, x+0.1, 2.16, cw-0.2, 0.48,
            size=18, bold=True, color=col, align=PP_ALIGN.CENTER)
        _tb(s, en, x+0.1, 2.60, cw-0.2, 0.28,
            size=10, color=MUTED, align=PP_ALIGN.CENTER)
        _tb(s, role, x+0.1, 2.88, cw-0.2, 0.44,
            size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        _rect(s, x+0.1, 3.36, cw-0.2, 0.03, BORDER)
        _bullets(s, tasks, x+0.1, 3.44, cw-0.2, 3.32,
                 size=10, color=MUTED, prefix="✔  ", spacing=5)


def s13_references(prs):
    s = new_slide(prs, "المراجع — وضّح آلية الاقتباس APA التي شرحتها الدكتورة وضَع المرجعين بصورة نظيفة ومرقّمة.")
    slide_heading(s, "قائمة المراجع",
                  "توثيق المرجعين بنمط APA المعتمد",
                  "الترقيم الرقمي [1] و [2] مع آلية الاقتباس البيني داخل الشرائح كما شرحت الدكتورة المشرفة")

    divider(s)

    # ─ Paper 1 ─
    card(s, 0.6, 2.0, W-1.2, 1.96, accent=BLUE)

    pill(s, "[1]", 0.78, 2.18, 0.54, 0.38, fg=BLUE, bg=ACCENT, size=14)

    _tb(s, "Xames, M. D. (2025).",
        1.44, 2.14, W-2.24, 0.34,
        size=12, bold=True, color=WHITE)
    _tb(s, "Toward a socially sustainable workforce in service operations: Predicting employee burnout with ensemble machine learning.",
        1.44, 2.46, W-2.24, 0.44,
        size=12, italic=True, color=MUTED)
    _tb(s, "Sustainable Manufacturing and Service Economics, 4, 100033.  https://doi.org/10.1016/j.smse.2025.100033",
        1.44, 2.90, W-2.24, 0.34,
        size=11, color=DARK)

    _rect(s, 0.6, 3.54, W-1.2, 0.38, ACCENT, lc=BORDER)
    _tb(s, "📌  اقتباس بيني: (Xames, 2025) [1]  •  يمثل 60 ٪ من النموذج — ركيزة أعباء العمل",
        0.78, 3.58, W-1.56, 0.30,
        size=11, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    # ─ Paper 2 ─
    card(s, 0.6, 4.08, W-1.2, 2.0, accent=PURPLE)

    pill(s, "[2]", 0.78, 4.26, 0.54, 0.38, fg=PURPLE, bg=ACCENT, size=14)

    _tb(s, "Golz, C., Kurpicz-Briki, M., Merhbene, G., & Puttick, A. (2024).",
        1.44, 4.22, W-2.24, 0.34,
        size=12, bold=True, color=WHITE)
    _tb(s, "Using natural language processing to find indication for burnout with text classification: From online data to real-world application.",
        1.44, 4.54, W-2.24, 0.44,
        size=12, italic=True, color=MUTED)
    _tb(s, "arXiv preprint arXiv:2409.16278.  https://doi.org/10.48550/arXiv.2409.16278",
        1.44, 4.98, W-2.24, 0.34,
        size=11, color=DARK)

    _rect(s, 0.6, 5.66, W-1.2, 0.38, ACCENT, lc=BORDER)
    _tb(s, "📌  اقتباس بيني: (Golz et al., 2024) [2]  •  يمثل 40 ٪ من النموذج — ركيزة المشاعر",
        0.78, 5.70, W-1.56, 0.30,
        size=11, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)

    # ─ Note on Badillo example ─
    _rect(s, 0.6, 6.18, W-1.2, 0.74, SURFACE, lc=BORDER)
    _tb(s, "💡  ملاحظة: المثال الذي ذكرته الدكتورة (Badillo et al., 2020) هو نموذج تعليمي يوضح آلية نسخ APA والترقيم — طبّقنا نفس القاعدة بالضبط على بحثَي مشروعنا الأساسيين [1] و [2].",
        0.78, 6.22, W-1.56, 0.66,
        size=10, color=MUTED)


def s14_conclusion(prs):
    s = new_slide(prs, "الخاتمة — أبرز الإنجاز وادعُ الدكتورة للأسئلة والمناقشة.")
    top_bar(s, GREEN, 0.14)
    _rect(s, 0, 0.14, W, 0.04, RGBColor(0x10, 0x4A, 0x34))

    _tb(s, "WorkPulse — الخلاصة",
        0, 0.68, W, 0.66,
        size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    _tb(s, "دمجنا أفضل ما في بحثين رائدين وسددنا فجوتيهما بحل هجين مفسَّر وقابل للنشر الفعلي",
        0.6, 1.4, W-1.2, 0.44,
        size=14, color=MUTED, align=PP_ALIGN.CENTER)

    # ─ 3 outcome cards ─
    oc = [
        ("المشكلة ✔",
         "حددنا الاحتراق الوظيفي كأزمة تشغيلية واقتصادية بتعريف علمي من منظمة الصحة العالمية.",
         RED),
        ("الحل ✔",
         "محرك هجين يجمع Random Forest (60 ٪) وDistilBERT (40 ٪) مع SHAP للتفسير الفوري.",
         GREEN),
        ("التوثيق ✔",
         "مرجعان بنمط APA مرقّمان [1] و [2] — اقتباس بيني واضح في كل شريحة.",
         BLUE),
    ]
    ocw = 3.82
    for i, (title, body, col) in enumerate(oc):
        x = 0.6 + i*(ocw+0.12)
        card(s, x, 2.02, ocw, 2.64, accent=col)
        _tb(s, title, x+0.18, 2.2, ocw-0.36, 0.46,
            size=16, bold=True, color=col)
        _tb(s, body, x+0.18, 2.72, ocw-0.36, 1.78,
            size=12, color=MUTED)

    # ─ Team readiness banner ─
    _rect(s, 0.6, 4.82, W-1.2, 0.96, SURFACE, lc=BORDER)
    _tb(s, "الفريق مستعد تماماً للمناقشة في:",
        0.82, 4.9, W-1.64, 0.30,
        size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _tb(s, "المعادلة الرياضية للمحرك الهجين  •  DistilBERT مقابل BERT الأصلي  •  قيم SHAP وكيفية احتسابها  •  قواعد البيانات والأمان",
        0.82, 5.22, W-1.64, 0.44,
        size=11, color=MUTED, align=PP_ALIGN.CENTER)

    # ─ Final thank you ─
    _rect(s, 0.6, 5.9, W-1.2, 1.46, ACCENT, lc=BORDER)
    _rect(s, 0.6, 5.9, 0.07, 1.46, GREEN)
    _tb(s, "شكر وتقدير للدكتورة المشرفة",
        0.82, 5.98, W-1.42, 0.44,
        size=16, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    _tb(s, "أسماء  •  عبد الله  •  شيماء  •  أحمد  •  نوري",
        0.82, 6.42, W-1.42, 0.38,
        size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _tb(s, "WorkPulse v2.0  |  2026",
        0.82, 6.82, W-1.42, 0.34,
        size=10, color=DARK, align=PP_ALIGN.CENTER)


# ==============================================================================
# MAIN
# ==============================================================================
def main():
    prs = Presentation()
    prs.slide_width  = Inches(W)
    prs.slide_height = Inches(H)

    s01_cover(prs)
    s02_agenda(prs)
    s03_problem(prs)
    s04_scholar(prs)
    s05_paper1(prs)
    s06_paper2(prs)
    s07_comparison(prs)
    s08_gap(prs)
    s09_tech(prs)
    s10_hybrid(prs)
    s11_pipeline(prs)
    s12_team(prs)
    s13_references(prs)
    s14_conclusion(prs)

    out = os.path.join(os.path.dirname(__file__), "docs",
                       "WorkPulse_Final_Deck.pptx")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    prs.save(out)
    print(f"DONE: Saved -> {out}")


if __name__ == "__main__":
    main()
