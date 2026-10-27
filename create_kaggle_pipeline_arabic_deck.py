"""
create_kaggle_pipeline_arabic_deck.py
======================================
Enhanced, Ultra-Clear Arabic PowerPoint Presentation Deck
- Large, bold fonts (13-16pt for bullets, 12-14pt for tables)
- Clear RTL alignment and high contrast colors
- Concise, punchy bullet points without fluff
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml import parse_xml

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(ROOT, "WorkPulse_Kaggle_Data_Pipeline_Arabic.pptx")

# ── Color Palette ──────────────────────────────────────────────────────────
BG_DARK    = RGBColor(0x0F, 0x17, 0x2A)   # Navy 900
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
MUTED      = RGBColor(0x94, 0xA3, 0xB8)   # Slate 400
GREEN      = RGBColor(0x10, 0xB9, 0x81)   # Emerald 500
RED        = RGBColor(0xEF, 0x44, 0x44)   # Red 500
BLUE       = RGBColor(0x38, 0x82, 0xF6)   # Blue 500
YELLOW     = RGBColor(0xF5, 0x9E, 0x0B)   # Amber 500
CYAN       = RGBColor(0x06, 0xB6, 0xD4)   # Cyan 500
PURPLE     = RGBColor(0xA7, 0x8B, 0xFA)   # Purple 400

HEX_BG     = "0F172A"
HEX_CARD   = "1E293B"
HEX_GREEN  = "10B981"
HEX_BLUE   = "3882F6"
HEX_RED    = "EF4444"
HEX_YELLOW = "F59E0B"
HEX_CYAN   = "06B6D4"
HEX_PURPLE = "A78BFA"

FONT_AR = "Arial"


def set_bg(slide, hex_color=HEX_BG):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(hex_color)


def add_shape(slide, left, top, width, height, hex_fill):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(hex_fill)
    shape.line.fill.background()
    return shape


def add_rect(slide, left, top, width, height, color):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_card(slide, left, top, width, height, accent_hex=HEX_GREEN):
    add_shape(slide, left, top, width, height, HEX_CARD)
    add_rect(slide, left, top, width, 0.08, RGBColor.from_string(accent_hex))


def add_text_ar(slide, text, left, top, width, height, font_size=15,
                bold=False, color=WHITE, align=PP_ALIGN.RIGHT):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = FONT_AR
    p.alignment = align
    return txBox


def add_bullets_ar(slide, items, left, top, width, height,
                   font_size=13.5, color=WHITE, spacing=10):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, item in enumerate(items):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = FONT_AR
        p.space_after = Pt(spacing)
        p.alignment = PP_ALIGN.RIGHT
    return txBox


def add_table_ar(slide, headers, rows, left, top, width, row_height=0.46, hex_header=HEX_GREEN):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                          Inches(width), Inches(row_height * n_rows))
    table = table_shape.table

    # Header Row
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12.5)
            p.font.bold = True
            p.font.color.rgb = WHITE
            p.font.name = FONT_AR
            p.alignment = PP_ALIGN.CENTER
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{hex_header}"/></a:solidFill>')
        tcPr.append(shd)

    # Data Rows
    for i, row in enumerate(rows):
        bg = HEX_CARD if i % 2 == 0 else HEX_BG
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(12)
                p.font.bold = (j == 0)
                p.font.color.rgb = WHITE
                p.font.name = FONT_AR
                p.alignment = PP_ALIGN.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            shd = parse_xml(f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{bg}"/></a:solidFill>')
            tcPr.append(shd)

    return table_shape


def build_arabic_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 1: الغلاف الرئيسي
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.16, GREEN)

    add_text_ar(s, "منصة WorkPulse للذكاء الاصطناعي", 1, 1.2, 11.33, 0.8,
                font_size=46, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text_ar(s, "هندسة وتنظيف بيانات كاجل (Kaggle Data Pipeline)", 1, 2.1, 11.33, 0.6,
                font_size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text_ar(s, "من 22,750 سجلاً خام في كاجل إلى 1,000 موظف جاهزين للتدريب والإنتاج", 1, 2.9, 11.33, 0.5,
                font_size=18, color=MUTED, align=PP_ALIGN.CENTER)

    # خطوات المسار بصناديق كبيرة وواضحة
    steps_flow = [
        ("داتا كاجل الخام\n22,750 سجل", HEX_RED, 1.0),
        ("تنظيف البيانات\nوحذف المفقود", HEX_YELLOW, 4.1),
        ("مواءمة المتغيرات\nوهندسة الميزات", HEX_BLUE, 7.2),
        ("1,000 موظف\nبيانات نظيفة 100%", HEX_GREEN, 10.3),
    ]
    for label, hex_c, x in steps_flow:
        add_shape(s, x, 4.2, 2.1, 1.2, hex_c)
        add_text_ar(s, label, x, 4.35, 2.1, 0.9, font_size=13, bold=True,
                    color=WHITE if hex_c != HEX_YELLOW else RGBColor(0x0F, 0x17, 0x2A),
                    align=PP_ALIGN.CENTER)
        if x < 10.0:
            add_text_ar(s, "◀", x + 2.2, 4.55, 0.8, 0.5, font_size=24, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    add_text_ar(s, "المصدر المعتمد: مسابقة HackerEarth / Kaggle — Are Your Employees Burning Out?",
                0.5, 6.4, 12.33, 0.4, font_size=13, color=MUTED, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 2: مصدر البيانات ولماذا كاجل
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, BLUE)

    add_text_ar(s, "1. مصدر البيانات: مسابقة HackerEarth على كاجل",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=BLUE)
    add_text_ar(s, "رابط الداتا: kaggle.com/datasets/blurredmachine/are-your-employees-burning-out",
                0.5, 0.85, 12.33, 0.3, font_size=13, color=MUTED)

    # كارت يمين: مواصفات الداتا
    add_card(s, 6.8, 1.35, 6.0, 5.5, HEX_BLUE)
    add_text_ar(s, "مواصفات الداتا الخام في كاجل", 7.0, 1.55, 5.5, 0.35, font_size=18, bold=True, color=BLUE)
    add_bullets_ar(s, [
        "• المسابقة: تحدي معتمد من HackerEarth لاستشراف الاحتراق الوظيفي.",
        "• إجمالي السجلات: 22,750 موظفاً من شركات حقيقية.",
        "• المتغير التابع: Burn Rate كنسبة مستمرة من 0.0 إلى 1.0.",
        "• 9 أعمدة أساسية: كود الموظف، تاريخ التعيين، الجنس، نوع الشركة، بيئة العمل عن بعد، المنصب، الموارد، الإجهاد الذهني، والاحتراق.",
        "• داتا معتمدة استخدمها أكثر من 500 باحث في علوم البيانات.",
    ], 7.0, 2.1, 5.5, 4.5, font_size=13.5, color=WHITE, spacing=12)

    # كارت يسار: لماذا اخترناها
    add_card(s, 0.5, 1.35, 6.0, 5.5, HEX_GREEN)
    add_text_ar(s, "لماذا اخترنا هذه الداتا بالتحديد؟", 0.7, 1.55, 5.5, 0.35, font_size=18, bold=True, color=GREEN)
    add_bullets_ar(s, [
        "• بيانات حقيقية وموثقة تعكس بيئات العمل الفعلية.",
        "• تحتوي على مقياس الإجهاد الذهني (Mental Fatigue) المعتمد نفسياً.",
        "• حجم العينة الكبير يمنح ثقة إحصائية ممتازة لاختيار النماذج.",
        "• احتواؤها على قيم مفقودة يبرز مهارة تنظيف البيانات كمهندسين.",
        "• مهدت لبناء المحرك الهجين بدمج نصوص الـ NLP التي تفتقدها كاجل.",
    ], 0.7, 2.1, 5.5, 4.5, font_size=13.5, color=WHITE, spacing=12)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 3: عيوب الداتا قبل التنظيف
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, RED)

    add_text_ar(s, "2. عيوب الداتا الخام في كاجل (لماذا تحتاج Data Cleaning؟)",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=RED)
    add_text_ar(s, "أهم المشاكل التي تمنع إدخال الداتا مباشرة في موديل الذكاء الاصطناعي:",
                0.5, 0.85, 12.33, 0.3, font_size=13, color=MUTED)

    problems = [
        ("القيم المفقودة (4,622 قيمة)", HEX_RED, [
            "• 1,124 صف بدون Target.",
            "• 2,117 قيمة مفقودة في الإجهاد.",
            "• 1,381 قيمة مفقودة في الموارد.",
            "• لا يمكن التدريب بدون Ground Truth.",
        ]),
        ("الصفوف المكررة (5,948 صفاً)", HEX_YELLOW, [
            "• تكرار نفس أرقام الميزات بالضبط.",
            "• يسبب انحيازاً غير عادل للنموذج.",
            "• تم حذفها للحفاظ على الأنماط الفريدة.",
        ]),
        ("اختلاف نوع المخرجات", HEX_PURPLE, [
            "• في كاجل: Burn Rate هو رقم مستمر.",
            "• نظامنا يحتاج تصنيفاً ثنائياً (0 أو 1).",
            "• العتبة: >= 0.50 يُعتبر احتراق وظيفي.",
        ]),
        ("غياب نصوص الفيدباك", HEX_CYAN, [
            "• كاجل تحتوي على جداول أرقام فقط.",
            "• لا توجد أي شكاوى أو مشاعر مكتوبة.",
            "• دمجنا نصوص الـ NLP ليكتمل النظام الهجين.",
        ]),
    ]

    for i, (title, hex_c, bullets) in enumerate(problems):
        col = 3 - i
        x = 0.5 + col * 3.1
        add_card(s, x, 1.4, 2.9, 5.5, hex_c)
        add_text_ar(s, title, x + 0.1, 1.55, 2.7, 0.35, font_size=14, bold=True, color=RGBColor.from_string(hex_c))
        add_bullets_ar(s, bullets, x + 0.1, 2.1, 2.7, 4.5, font_size=12, color=WHITE, spacing=10)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 4: مسار التنظيف بالأرقام
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, GREEN)

    add_text_ar(s, "3. مسار تنظيف البيانات بالأرقام (Data Cleaning Pipeline)",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=GREEN)

    add_table_ar(s,
        ["المؤشر الإحصائي", "قبل التنظيف (Raw)", "بعد التنظيف (Clean)", "الإجراء المطبق"],
        [
            ["إجمالي السجلات", "22,750 صفاً", "12,642 صفاً صالحاً", "حذف 10,108 صفاً غير صالح"],
            ["القيم المفقودة (NaN)", "4,622 خلية فارغة", "0 (صفر تماماً)", "حذف الصفوف الناقصة جذرياً"],
            ["الصفوف المكررة", "5,948 صفاً مكرراً", "0 (صفر)", "إزالة كاملة لمنع انحياز النموذج"],
            ["حجم العينة النهائية", "-", "1,000 موظف", "سحب عينة طبقية متوازنة (Stratified)"],
            ["نسبة الاحتراق الوظيفي", "~43% (مستمر)", "44.1% (ثنائي)", "مطابقة لإحصائيات منظمة العمل الدولية"],
            ["عدد الأعمدة ونوعها", "9 أعمدة أرقام", "10 أعمدة (أرقام + نصوص)", "إضافة طبقة نصوص الـ NLP والأسماء"],
        ],
        0.6, 1.25, 12.13, 0.48, hex_header=HEX_GREEN
    )

    add_card(s, 0.6, 4.6, 12.13, 2.3, HEX_GREEN)
    add_text_ar(s, "النتيجة النهائية لمرحلة هندسة وتنظيف البيانات", 0.9, 4.75, 11.5, 0.35, font_size=16, bold=True, color=GREEN)
    add_bullets_ar(s, [
        "1. ملف data/employees.csv جاهز وموثق بنسبة 100% وخالٍ من أي أخطاء أو قيم شاذة.",
        "2. العينة تضم 1,000 موظف بأسماء عربية فريدة تعبر عن واقع المؤسسات المحلية.",
        "3. استقرار تام للنتائج بفضل تثبيت معامل العشوائية (random_state = 42).",
    ], 0.9, 5.2, 11.5, 1.5, font_size=13, color=WHITE, spacing=8)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 5: مواءمة المتغيرات
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, CYAN)

    add_text_ar(s, "4. جدول مواءمة المتغيرات (Kaggle  ◀  WorkPulse Schema)",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=CYAN)
    add_text_ar(s, "تحويل أعمدة كاجل إلى مؤشرات الموارد البشرية مع إضافة التشويش الطبيعي (Noise):",
                0.5, 0.85, 12.33, 0.3, font_size=13, color=MUTED)

    add_table_ar(s,
        ["عمود نظام WorkPulse", "المصدر من كاجل", "معادلة الاشتقاق والتحويل", "المدى الناتج"],
        [
            ["employee_id", "كود مشفر (Hash)", "توليد كود منظم: EMP0001 - EMP1000", "EMP0001 - EMP1000"],
            ["name", "- (غير موجود)", "أسماء عربية ثلاثية مطابقة للجنس", "1,000 اسم فريد"],
            ["avg_daily_hours", "Resource Allocation (1-10)", "6.0 + (RA - 1) / 9 * 7.0 + Noise", "6.0 - 13.0 ساعة"],
            ["overtime_hours_week", "Resource + Fatigue", "(RA/10 * 15) + (Fatigue/10 * 10) + Noise", "0.0 - 25.0 ساعة"],
            ["absence_days_month", "Mental Fatigue Score", "Mental Fatigue / 10 * 8 + Noise", "0 - 8 أيام"],
            ["late_arrivals_month", "Burn Rate", "Burn Rate * 15 + Noise", "0 - 15 مرة"],
            ["tasks_assigned", "Designation + RA", "5 + (Desig/5 * 20) + (RA/10 * 15)", "5 - 40 مهمة"],
            ["vacation_days_last_year", "Burn Rate (عكسي)", "25 * (1 - Burn Rate) + Noise", "0 - 25 يوماً"],
            ["feedback", "- (غير موجود)", "نصوص مشاعر موزعة حسب معدل الإجهاد", "45 جملة معبرة"],
            ["burnout", "Burn Rate (0.0-1.0)", "1 إذا كان Burn Rate >= 0.50 وإلا 0", "0 أو 1"],
        ],
        0.5, 1.25, 12.33, 0.38, hex_header=HEX_CYAN
    )

    add_card(s, 0.5, 5.7, 12.33, 1.4, HEX_CYAN)
    add_text_ar(s, "ملاحظة علمية هامة للمناقشة", 0.8, 5.85, 11.5, 0.3, font_size=14, bold=True, color=CYAN)
    add_bullets_ar(s, [
        "• أضفنا Gaussian Noise (تشويش طبيعي) لمنع العلاقات الخطية التامة ولإجبار الموديل على تعلم أنماط واقعية.",
        "• رصيد الإجازات أخذ علاقة عكسية (عامل حماية): كلما زاد الإرهاق قل استهلاك الإجازات.",
    ], 0.8, 6.2, 11.5, 0.8, font_size=12, color=WHITE, spacing=4)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 6: نصوص الـ NLP
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, PURPLE)

    add_text_ar(s, "5. طبقة نصوص المشاعر (NLP Sentiment Fusion)",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=RGBColor.from_string(HEX_PURPLE))
    add_text_ar(s, "صياغة 45 جملة شكوى وتقييم واقعية وربطها إحصائياً بمعدل إجهاد الموظف الحقيقي:",
                0.5, 0.85, 12.33, 0.3, font_size=13, color=MUTED)

    categories_ar = [
        ("المشاعر الإيجابية (Positive)", HEX_GREEN, "352 موظفاً (35.2%)", [
            '"I love my job and feel very motivated."',
            '"I feel valued and recognized for my work."',
            '"The company culture is supportive."',
            "• الموظف السليم (Burn Rate < 0.35)",
            "• احتمال 70% فيدباك إيجابي و 30% محايد.",
        ]),
        ("المشاعر المحايدة (Neutral)", HEX_YELLOW, "287 موظفاً (28.7%)", [
            '"Work is fine, nothing special to report."',
            '"Some days are tough but it is manageable."',
            '"Work-life balance is acceptable."',
            "• الموظف المتوسط (Burn Rate 0.35 - 0.60)",
            "• توزيع متوازن بين الفئات الثلاث.",
        ]),
        ("المشاعر السلبية (Negative)", HEX_RED, "361 موظفاً (36.1%)", [
            '"I feel completely exhausted and overwhelmed."',
            '"The workload is unbearable and too stressful."',
            '"My mental health is deteriorating."',
            "• الموظف المرهق (Burn Rate > 0.60)",
            "• احتمال 75% فيدباك سلبي و 25% محايد.",
        ]),
    ]

    for i, (title, hex_c, count, items) in enumerate(categories_ar):
        col = 2 - i
        x = 0.5 + col * 4.2
        add_card(s, x, 1.4, 3.9, 5.4, hex_c)
        add_text_ar(s, title, x + 0.2, 1.6, 3.5, 0.35, font_size=15, bold=True, color=RGBColor.from_string(hex_c))
        add_text_ar(s, count, x + 0.2, 2.0, 3.5, 0.3, font_size=12, color=MUTED)
        add_bullets_ar(s, items, x + 0.2, 2.45, 3.5, 4.0, font_size=12, color=WHITE, spacing=10)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 7: الإحصائيات ونتائج الموديل
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, GREEN)

    add_text_ar(s, "6. الإحصائيات الدقيقة لـ 1,000 موظف ونتائج التدريب",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=GREEN)

    add_table_ar(s,
        ["المتغير", "المتوسط (Mean)", "الانحراف (Std)", "الحد الأدنى", "الوسيط (50%)", "الحد الأقصى"],
        [
            ["ساعات العمل اليومية", "8.83 ساعة", "1.69", "6.00", "8.78", "13.00"],
            ["الساعات الإضافية (أسبوعياً)", "12.85 ساعة", "5.34", "0.00", "12.85", "25.00"],
            ["أيام الغياب شهرياً", "4.61 يوم", "2.04", "0", "5.00", "8"],
            ["مرات التأخير شهرياً", "6.91 مرة", "3.68", "0", "7.00", "15"],
            ["المهام المسندة", "20.77 مهمة", "8.18", "5", "20.00", "40"],
            ["الإجازات المستهلكة", "13.30 يوم", "5.85", "0", "13.00", "25"],
        ],
        0.5, 1.15, 12.33, 0.42, hex_header=HEX_GREEN
    )

    # كارت يمين: الفئات
    add_card(s, 6.8, 4.2, 6.0, 2.9, HEX_GREEN)
    add_text_ar(s, "توزيع فئات الهدف (Burnout)", 7.0, 4.35, 5.5, 0.35, font_size=16, bold=True, color=GREEN)
    add_bullets_ar(s, [
        "• سليم (Class 0): 559 موظف بنسبة 55.9%.",
        "• في خطر احتراق (Class 1): 441 موظف بنسبة 44.1%.",
        "• توازن الفئات (1.27 : 1) يلغي الحاجة للتوليد الاصطناعي (SMOTE).",
        "• مطابقة تامة لتقارير مؤسسة Gallup ومنظمة العمل الدولية.",
    ], 7.0, 4.8, 5.5, 2.2, font_size=12.5, color=WHITE, spacing=7)

    # كارت يسار: نتائج الموديل
    add_card(s, 0.5, 4.2, 6.0, 2.9, HEX_BLUE)
    add_text_ar(s, "نتائج تدريب الموديل (Random Forest)", 0.7, 4.35, 5.5, 0.35, font_size=16, bold=True, color=BLUE)
    add_bullets_ar(s, [
        "• دقة النموذج الإجمالية (Accuracy): 95.50%.",
        "• مقياس الجودة F1-Score: حقق 0.95 (Precision: 93.4% | Recall: 96.6%).",
        "• أهم العوامل المؤثرة (SHAP):",
        "    1) الأوفرتايم (30.1%)  2) الإجازات (22.9%)  3) ساعات العمل (21.9%).",
    ], 0.7, 4.8, 5.5, 2.2, font_size=12.5, color=WHITE, spacing=6)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 8: إجابات سريعة للمناقشة
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, YELLOW)

    add_text_ar(s, "7. دليل إجابات المناقشة السريعة (دليل الدفاع للدكتورة)",
                0.5, 0.3, 12.33, 0.5, font_size=26, bold=True, color=YELLOW)
    add_text_ar(s, "إجابات مباشرة ومقنعة لأهم 4 أسئلة متوقعة من لجنة التحكيم:",
                0.5, 0.85, 12.33, 0.3, font_size=13, color=MUTED)

    qa_ar = [
        ("س: لماذا لم تستخدموا داتا كاجل كما هي بدون تعديل؟", HEX_RED, [
            "• الداتا الخام بها 4,622 قيمة مفقودة و 5,948 سجلاً مكرراً.",
            "• تفتقر تماماً للنصوص، ومشروعنا يقدم محركاً هجيناً (أرقام + NLP).",
            "• مسار التنظيف سمح لنا بإصلاح العيوب وتجهيز نظام متكامل.",
        ]),
        ("س: كيف تعاملتم مع القيم المفقودة (Missing Values)؟", HEX_YELLOW, [
            "• حذفنا الصفوف الناقصة في الـ Target (Burn Rate) لعدم جواز التنبؤ بدونه.",
            "• حذفنا السجلات غير المكتملة في الميزات الأساسية.",
            "• النتيجة: 12,642 سجلاً صالحاً بنسبة 100%.",
        ]),
        ("س: لماذا اخترتم حجم 1,000 موظف بالتحديد؟", HEX_BLUE, [
            "• عينة ممثلة لشركة متوسطة إلى كبيرة (Enterprise Scale).",
            "• استخدمنا Stratified Sampling للحفاظ على نسبة الاحتراق 44%.",
            "• سرعة فائقة في استجابة قاعدة البيانات (أقل من 200ms).",
        ]),
        ("س: من أين أتيتم بنصوص الفيدباك (Feedback)؟", HEX_PURPLE, [
            "• صممنا 45 جملة معبرة وربطناها بمعدل إجهاد كل موظف.",
            "• هذا مكننا من استخدام نموذج DistilBERT لتحليل المشاعر.",
            "• الموديل أصبح هجيناً (60% أرقام + 40% نصوص).",
        ]),
    ]

    for i, (q, hex_c, answers) in enumerate(qa_ar):
        row = i // 2
        col = 1 - (i % 2)
        x = 0.5 + col * 6.2
        y = 1.35 + row * 2.9

        add_card(s, x, y, 5.9, 2.65, hex_c)
        add_text_ar(s, q, x + 0.2, y + 0.15, 5.5, 0.35, font_size=13, bold=True, color=RGBColor.from_string(hex_c))
        add_bullets_ar(s, answers, x + 0.2, y + 0.6, 5.5, 1.9, font_size=11.5, color=WHITE, spacing=6)

    # Save with fallback if open in PowerPoint
    try:
        prs.save(OUT_PATH)
        print(f"[SUCCESS] Ultra-clear Arabic presentation saved: {OUT_PATH}")
    except PermissionError:
        alt_path = os.path.join(ROOT, "WorkPulse_Kaggle_Data_Pipeline_Arabic_LargeFont.pptx")
        prs.save(alt_path)
        print(f"[SUCCESS] Original file was open in PowerPoint. Saved to: {alt_path}")


if __name__ == "__main__":
    build_arabic_presentation()
