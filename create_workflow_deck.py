"""
create_workflow_deck.py
=======================
Generates a dedicated PowerPoint presentation explaining the exact step-by-step
chronological workflow and handover between the 5 engineers:
Assma -> Abdullah & Shimaa -> Ahmed -> Noury.
Saved to: docs/WorkPulse_Team_Workflow_and_Handover.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Color Palette ─────────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x0A, 0x0F, 0x1D)   # Deep Midnight 950
SURFACE    = RGBColor(0x13, 0x1D, 0x31)   # Premium Navy 850
CARD_BG    = RGBColor(0x1E, 0x29, 0x42)   # Card Slate 800
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLUE       = RGBColor(0x38, 0xBD, 0xF8)   # Sky Cyan (Ahmed)
GREEN      = RGBColor(0x34, 0xD3, 0x99)   # Emerald Green (Noury)
YELLOW     = RGBColor(0xFB, 0xBF, 0x24)   # Amber Gold (Abdullah)
PURPLE     = RGBColor(0xC0, 0x84, 0xFC)   # Lavender (Shimaa)
ORANGE     = RGBColor(0xFB, 0x92, 0x3C)   # Energetic Orange (Assma)
MUTED      = RGBColor(0x94, 0xA3, 0xB8)   # Cool Gray
ACCENT     = RGBColor(0x60, 0xA5, 0xFA)   # Bright Blue Accent


def set_bg(slide, color=DARK_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(1, Inches(left), Inches(top), Inches(width), Inches(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_card(slide, left, top, width, height, accent_color=BLUE, bg_color=CARD_BG):
    card = add_rect(slide, left, top, width, height, bg_color)
    add_rect(slide, left, top, width, 0.08, accent_color)
    return card


def add_box(slide, text, left, top, width, height, font_size=14, bold=False, color=WHITE, align=PP_ALIGN.LEFT, font_name="Segoe UI"):
    tx = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tx.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.alignment = align
    return tx


def add_slide_header(slide, tag, title, subtitle):
    set_bg(slide)
    # Top Tag Pill
    add_rect(slide, 0.8, 0.45, 2.8, 0.35, SURFACE, BLUE)
    add_box(slide, tag, 0.9, 0.48, 2.6, 0.3, font_size=11, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    # Title
    add_box(slide, title, 0.8, 0.88, 11.5, 0.55, font_size=24, bold=True, color=WHITE)
    # Subtitle
    add_box(slide, subtitle, 0.8, 1.45, 11.5, 0.38, font_size=13, bold=False, color=MUTED)


def create_deck():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_bg(s1)
    add_rect(s1, 0, 0, 13.333, 0.15, BLUE)
    add_rect(s1, 0.8, 1.2, 4.2, 0.45, CARD_BG, BLUE)
    add_box(s1, "WORKPULSE PIPELINE & HANDOVER", 1.0, 1.28, 3.8, 0.3, font_size=12, bold=True, color=BLUE)

    add_box(s1, "رحلة بناء المشروع وتسلسل العمل بين الفريق", 0.8, 1.9, 11.5, 0.9, font_size=32, bold=True, color=WHITE)
    add_box(s1, "من بدأ أولاً؟ ماذا سلّم لمن؟ وكيف اندمجت كل القطع في منصة ذكاء اصطناعي متكاملة؟", 0.8, 2.9, 11.5, 0.5, font_size=16, color=MUTED)

    # Team cards on Title Slide
    team = [
        ("Eng. Assma", "المرحلة 1: توليد وهندسة البيانات", ORANGE),
        ("Eng. Abdullah", "المرحلة 2-أ: تدريب نموذج ML", YELLOW),
        ("Eng. Shimaa", "المرحلة 2-ب: معالجة النصوص NLP", PURPLE),
        ("Eng. Ahmed", "المرحلة 3: الباك إند وقاعدة البيانات", BLUE),
        ("Eng. Noury", "المرحلة 4: الواجهة والتفاعل الحي", GREEN),
    ]
    card_w = 2.22
    for i, (name, role, clr) in enumerate(team):
        cx = 0.8 + i * (card_w + 0.15)
        add_card(s1, cx, 4.0, card_w, 2.4, clr, SURFACE)
        add_box(s1, f"0{i+1}", cx + 0.15, 4.2, 1.9, 0.4, font_size=20, bold=True, color=clr)
        add_box(s1, name, cx + 0.15, 4.7, 1.9, 0.35, font_size=14, bold=True, color=WHITE)
        add_box(s1, role, cx + 0.15, 5.2, 1.9, 0.9, font_size=11, color=MUTED)

    # Footer note
    add_box(s1, "دليل شامل للمناقشة والمذاكرة الجماعية | توثيق تسلسل المهام من الصفر حتى الإنتاج", 0.8, 6.7, 11.5, 0.4, font_size=12, color=MUTED)

    # =========================================================================
    # SLIDE 2: Overall Sequential Flowchart (The Big Picture)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_slide_header(s2, "PIPELINE OVERVIEW", "المخطط الزمني الشامل: كيف تم ربط المشروع خطوة بخطوة؟", "نظرة عامة توضح التسليم والتسلم بين كل عضو وما أنتجه للزميل التالي")

    steps = [
        ("1. هندسة البيانات", "Eng. Assma", "توليد 50 موظف + محاكاة المشاكل + تنظيف الإحصائيات", "المُخرَج: employees.csv", ORANGE),
        ("2. النمذجة (توازي)", "Eng. Abdullah & Shimaa", "عبدالله: Random Forest للأرقام\nشيماء: DistilBERT للمشاعر", "المُخرَج: workload_model.pkl\n+ sentiment.py", YELLOW),
        ("3. الدمج والسيرفر", "Eng. Ahmed", "FastAPI + PostgreSQL + الدمج الهجين 60/40 + SHAP", "المُخرَج: 7 REST API Endpoints", BLUE),
        ("4. الواجهات التفاعلية", "Eng. Noury", "بوابة الموظف لإعادة التقييم + لوحة تحكم الإدارة + Chart.js", "المُخرَج: Dashboard تفاعلي بالكامل", GREEN),
    ]

    card_w2 = 2.7
    for i, (title, author, desc, out, clr) in enumerate(steps):
        cx = 0.8 + i * (card_w2 + 0.28)
        add_card(s2, cx, 2.1, card_w2, 4.2, clr, CARD_BG)
        add_box(s2, f"المرحلة {i+1}", cx + 0.15, 2.25, 2.4, 0.3, font_size=12, bold=True, color=clr)
        add_box(s2, title, cx + 0.15, 2.6, 2.4, 0.4, font_size=15, bold=True, color=WHITE)
        add_box(s2, f"👤 {author}", cx + 0.15, 3.05, 2.4, 0.3, font_size=12, bold=True, color=clr)
        add_box(s2, desc, cx + 0.15, 3.45, 2.4, 1.4, font_size=11, color=MUTED)

        # Output Box
        add_rect(s2, cx + 0.15, 5.0, 2.4, 1.1, SURFACE, clr)
        add_box(s2, "📦 التسليم الفعلي:", cx + 0.2, 5.05, 2.3, 0.25, font_size=10, bold=True, color=clr)
        add_box(s2, out, cx + 0.2, 5.35, 2.3, 0.7, font_size=10, bold=True, color=WHITE)

        # Draw arrow between steps
        if i < 3:
            arr_x = cx + card_w2 + 0.05
            add_box(s2, "➔", arr_x, 3.8, 0.25, 0.5, font_size=24, bold=True, color=MUTED)

    add_box(s2, "💡 ملاحظة جوهرية: م. عبد الله وم. شيماء عملا بالتوازي على نفس ملف البيانات الذي أعدته م. أسماء، وسلما نتائجهما لم. أحمد في نفس التوقيت.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=WHITE)

    # =========================================================================
    # SLIDE 3: Phase 1 - Eng. Assma
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_slide_header(s3, "PHASE 1 • THE FOUNDATION", "المرحلة الأولى: م. أسماء (Data Engineering & Generation)", "البداية الفعلية للمشروع: بناء وتنظيف البيانات الواقعية لـ 50 موظفاً")

    # Left Card: Why started here & What was done
    add_card(s3, 0.8, 2.1, 5.6, 4.4, ORANGE, CARD_BG)
    add_box(s3, "ليه بدأنا من هنا؟ وماذا نفذت م. أسماء؟", 1.0, 2.3, 5.2, 0.4, font_size=16, bold=True, color=ORANGE)
    txt_assma = (
        "1. لا يمكن تدريب ذكاء اصطناعي أو بناء شاشات بدون قاعدة بيانات واقعية.\n\n"
        "2. كتبت م. أسماء كود data/generate_data.py لتوليد بيانات 50 موظفاً:\n"
        "   • ضبط ساعات العمل الطبيعية (40h) والإضافية (Overtime).\n"
        "   • ربط عدد الاجتماعات (Meetings) والإيميلات بنسب الإرهاق الواقعية.\n"
        "   • توليد نصوص تعليقات واقعية بالإنجليزية (Feedbacks) تعكس حالة الضغط أو الرضا.\n\n"
        "3. محاكاة تحديات العالم الحقيقي (Noise & Cleaning):\n"
        "   • إضافة قيم شاذة (Outliers) مثل ساعات عمل مبالغ فيها.\n"
        "   • تنظيف البيانات وضمان خلوها من القيم المفقودة (Imputation)."
    )
    add_box(s3, txt_assma, 1.0, 2.8, 5.2, 3.5, font_size=12, color=WHITE)

    # Right Card: Deliverable & Handover Targets
    add_card(s3, 6.7, 2.1, 5.8, 4.4, BLUE, CARD_BG)
    add_box(s3, "📦 ماذا أنتجت؟ ولمن سلّمت العمل؟", 6.9, 2.3, 5.4, 0.4, font_size=16, bold=True, color=BLUE)

    txt_assma_out = (
        "📄 المُخرَج الأساسي: ملف data/employees.csv (جاهز ونظيف 100%)\n\n"
        "🤝 تسليم المهام (Handover Matrix):\n"
        "• سلّمت الأعمدة الرقمية (Hours, Overtime, Projects...) إلى:\n"
        "  ⬅️ م. عبد الله (لبناء موديل التنبؤ Machine Learning).\n\n"
        "• سلّمت عمود النصوص والآراء (Employee Feedback) إلى:\n"
        "  ⬅️ م. شيماء (لبناء نموذج تحليل المشاعر NLP).\n\n"
        "• سلّمت الملف بالكامل بجميع بيانات الموظفين إلى:\n"
        "  ⬅️ م. أحمد (لتغذية قاعدة البيانات PostgreSQL تلقائياً عند التشغيل)."
    )
    add_box(s3, txt_assma_out, 6.9, 2.8, 5.4, 3.5, font_size=12, color=WHITE)

    add_box(s3, "🔑 سؤال الدكتور المتوقع: لو سأل 'ليه ما استخدمتوش داتا جاهزة من Kaggle؟' الإجابة: لأننا أردنا تحكم كامل في السيناريوهات ودمج النصوص مع الأرقام بشكل متطابق.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 4: Phase 2A - Eng. Abdullah
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_slide_header(s4, "PHASE 2-A • NUMERICAL ML", "المرحلة الثانية (أ): م. عبد الله (Machine Learning Engine)", "تدريب نموذج Random Forest على البيانات الرقمية وتصدير الملف النهائي")

    # Left Card
    add_card(s4, 0.8, 2.1, 5.6, 4.4, YELLOW, CARD_BG)
    add_box(s4, "ماذا استلم وماذا فعل م. عبد الله؟", 1.0, 2.3, 5.2, 0.4, font_size=16, bold=True, color=YELLOW)
    txt_abdullah = (
        "1. المدخلات المستلمة:\n"
        "   • استلم ملف data/employees.csv من م. أسماء.\n\n"
        "2. المعالجة والنمذجة الرياضية (train_model.py):\n"
        "   • فصل الخصائص (Hours, Overtime, Projects, Meetings, Emails).\n"
        "   • تطبيق StandardScaler لتوحيد المقاييس الرقمية.\n"
        "   • تدريب ومقارنة 3 نماذج ذكاء اصطناعي:\n"
        "     - Logistic Regression (دقة 76%)\n"
        "     - Decision Tree (دقة 84% مع خطر Overfitting)\n"
        "     - Random Forest Classifier (فاز بأعلى دقة 90%+).\n\n"
        "3. استخراج أهمية العوامل (Feature Importance):\n"
        "   • Overtime و Working Hours كانا العاملين الأكثر حساسية للتنبؤ بالإجهاد."
    )
    add_box(s4, txt_abdullah, 1.0, 2.8, 5.2, 3.5, font_size=12, color=WHITE)

    # Right Card
    add_card(s4, 6.7, 2.1, 5.8, 4.4, BLUE, CARD_BG)
    add_box(s4, "📦 التسليم الفعلي: ماذا قدّم للفريق؟", 6.9, 2.3, 5.4, 0.4, font_size=16, bold=True, color=BLUE)
    txt_abdullah_out = (
        "📦 المُخرَج الأساسي (Model Artifact):\n"
        "   • حفظ النموذج المكتمل في ملف: models/workload_model.pkl\n"
        "   • حفظ مقياس المعايرة في: models/scaler.pkl\n\n"
        "🤝 لمن سلّم شغله؟\n"
        "   • سلّم ملف الـ .pkl مباشرة إلى ⬅️ م. أحمد.\n\n"
        "🎯 فائدة التسليم لم. أحمد:\n"
        "   • م. أحمد أخذ النموذج وقام بعمل Joblib Load له عند إقلاع السيرفر\n"
        "   • مما مكّن الباك إند من التنبؤ بضغط أي موظف في أجزاء من الثانية (Sub-millisecond) دون الحاجة لإعادة تدريب الموديل كل مرة."
    )
    add_box(s4, txt_abdullah_out, 6.9, 2.8, 5.4, 3.5, font_size=12, color=WHITE)

    add_box(s4, "🔑 سؤال الدكتور المتوقع: ليه Random Forest بالذات؟ الإجابة: لأنه Ensemble Learning يدمج أشجار متعددة ويمنع الـ Overfitting ويعطي أداءً ممتازاً على الداتا المنظمة (Tabular Data).", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 5: Phase 2B - Eng. Shimaa
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_slide_header(s5, "PHASE 2-B • NLP INTELLIGENCE", "المرحلة الثانية (ب): م. شيماء (Sentiment Analysis & DistilBERT)", "بناء عقل معالجة النصوص الطبيعية لتحليل مشاعر الموظف وتحديد درجة الخطر")

    # Left Card
    add_card(s5, 0.8, 2.1, 5.6, 4.4, PURPLE, CARD_BG)
    add_box(s5, "ماذا استلمت وماذا صنعت م. شيماء؟", 1.0, 2.3, 5.2, 0.4, font_size=16, bold=True, color=PURPLE)
    txt_shimaa = (
        "1. المدخلات المستلمة:\n"
        "   • استلمت نصوص تعليقات الموظفين (Feedbacks) من م. أسماء.\n\n"
        "2. اختيار وبناء المعمارية اللغوية (backend/sentiment.py):\n"
        "   • استعانت بنموذج المحولات المتطور DistilBERT من HuggingFace.\n"
        "   • تميز DistilBERT بأنه أخف بـ 40% وأسرع بـ 60% من BERT الأصلي مع الحفاظ على 97% من دقته الفائقة.\n\n"
        "3. ابتكار معادلة خطر المشاعر (Sentiment Risk Formula):\n"
        "   • إذا كان الشعور سلبياً (NEGATIVE): الخطر = Confidence * 100\n"
        "   • إذا كان إيجابياً (POSITIVE): الخطر = (1 - Confidence) * 50\n"
        "   • إذا كان محايداً (NEUTRAL): الخطر = 30% ثابتة."
    )
    add_box(s5, txt_shimaa, 1.0, 2.8, 5.2, 3.5, font_size=12, color=WHITE)

    # Right Card
    add_card(s5, 6.7, 2.1, 5.8, 4.4, BLUE, CARD_BG)
    add_box(s5, "📦 التسليم الفعلي: ماذا قدّمت للفريق؟", 6.9, 2.3, 5.4, 0.4, font_size=16, bold=True, color=BLUE)
    txt_shimaa_out = (
        "📦 المُخرَج الأساسي (Callable Python Module):\n"
        "   • موديول كامل جاهز: backend/sentiment.py\n"
        "   • يحتوي على دالة واضحة: analyze_sentiment(text) تعيد القاموس:\n"
        "     { 'sentiment': 'NEGATIVE', 'score': 0.98, 'sentiment_risk': 98.0 }\n\n"
        "🤝 لمن سلّمت شغلها؟\n"
        "   • سلّمت الدالة والموديول بالكامل إلى ⬅️ م. أحمد.\n\n"
        "🎯 فائدة التسليم لم. أحمد:\n"
        "   • أتاح لم. أحمد استدعاء analyze_sentiment بضغطة زر داخل معادلة الخطر الشاملة (60/40) دون الدخول في تعقيدات الـ PyTorch والـ Transformers."
    )
    add_box(s5, txt_shimaa_out, 6.9, 2.8, 5.4, 3.5, font_size=12, color=WHITE)

    add_box(s5, "🔑 سؤال الدكتور المتوقع: ليه ما استعملتوش NLTK أو VADER؟ الإجابة: لأن DistilBERT يفهم السياق والمعنى الدلالي العميق (Context & Semantics) وليس مجرد عد الكلمات الإيجابية والسلبية.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 6: Phase 3 - Eng. Ahmed
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_slide_header(s6, "PHASE 3 • SYSTEM INTEGRATION", "المرحلة الثالثة: م. أحمد (Backend, Fusion Engine & Database)", "حلقة الوصل والعمود الفقري: دمج الموديلات، بناء الـ API وقاعدة البيانات PostgreSQL")

    # Left Card
    add_card(s6, 0.8, 2.1, 5.6, 4.4, BLUE, CARD_BG)
    add_box(s6, "ماذا استلم م. أحمد وكيف ربط القطع؟", 1.0, 2.3, 5.2, 0.4, font_size=16, bold=True, color=BLUE)
    txt_ahmed = (
        "1. المدخلات المستلمة من الفريق:\n"
        "   • داتا الـ 50 موظف من م. أسماء.\n"
        "   • ملف الموديل .pkl من م. عبد الله.\n"
        "   • دالة المشاعر sentiment.py من م. شيماء.\n\n"
        "2. ما بناه م. أحمد لتجميع المشروع:\n"
        "   • قاعدة البيانات (backend/database.py): PostgreSQL + SQLAlchemy بجداول Employees, Assessments, Feedback, Users.\n"
        "   • محرك الدمج الهجين (backend/risk.py):\n"
        "     الخطر الشامل = (60% نموذج أرقام م. عبد الله) + (40% مشاعر نصوص م. شيماء).\n"
        "   • محرك التفسير (backend/explain.py): حساب قيم SHAP لإظهار الأسباب.\n"
        "   • الحماية والتشفير: مكتبة Bcrypt لحماية كلمات المرور ونظام الـ JWT."
    )
    add_box(s6, txt_ahmed, 1.0, 2.8, 5.2, 3.5, font_size=12, color=WHITE)

    # Right Card
    add_card(s6, 6.7, 2.1, 5.8, 4.4, GREEN, CARD_BG)
    add_box(s6, "📦 التسليم الفعلي: ماذا قدّم لم. نوري؟", 6.9, 2.3, 5.4, 0.4, font_size=16, bold=True, color=GREEN)
    txt_ahmed_out = (
        "📦 المُخرَج الأساسي (Operational REST API):\n"
        "   • خادم FastAPI سريع وعصري (backend/main.py) يحتوي على Endpoints:\n"
        "     1. POST /api/auth/login (تسجيل دخول وتوزيع الصلاحيات)\n"
        "     2. GET /api/stats (إحصائيات المنظمة ونسب الخطر)\n"
        "     3. GET /api/employees (عرض قائمة الموظفين وحالاتهم)\n"
        "     4. POST /api/assessments (حساب الخطر فوري وتحديث الحالة)\n"
        "     5. GET /api/assessments/history (سجل إعادة التقييمات للموظف)\n\n"
        "🤝 لمن سلّم شغله؟\n"
        "   • سلّم روابط الـ API وتوثيق الـ Swagger إلى ⬅️ م. نوري لبناء الواجهات."
    )
    add_box(s6, txt_ahmed_out, 6.9, 2.8, 5.4, 3.5, font_size=12, color=WHITE)

    add_box(s6, "🔑 سؤال الدكتور المتوقع: ليه بنيتوا Risk Engine 60/40؟ الإجابة: لأن الأرقام وحدها خادعة بدون سماع شكوى الموظف، والمشاعر وحدها عاطفية بدون إثباتات ساعات العمل، فالدمج يحقق أعلى عدالة ودقة.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 7: Phase 4 - Eng. Noury
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_slide_header(s7, "PHASE 4 • USER EXPERIENCE", "المرحلة الرابعة: م. نوري (Frontend UI & Real-Time Portals)", "تحويل الأكواد والمعادلات المعقدة إلى شاشات تفاعلية جذابة وسهلة الاستخدام")

    # Left Card
    add_card(s7, 0.8, 2.1, 5.6, 4.4, GREEN, CARD_BG)
    add_box(s7, "ماذا استلمت وماذا صممت م. نوري؟", 1.0, 2.3, 5.2, 0.4, font_size=16, bold=True, color=GREEN)
    txt_noury = (
        "1. المدخلات المستلمة:\n"
        "   • استلمت نقاط اتصال الـ REST API من م. أحمد.\n\n"
        "2. بناء المنظومة التفاعلية (HTML5 / Vanilla JS / CSS3):\n"
        "   • لوحة تحكم المدير (index.html): لمراقبة الـ 50 موظف، مع مخططات Chart.js لتوزيع الخطر والتنبيهات الاستباقية.\n"
        "   • بوابة الموظف الشخصية (employee-portal.html):\n"
        "     - تتيح للموظف رؤية حالته ودرجة إجهاده.\n"
        "     - تتيح له إعادة التقييم الذاتي وتعديل ساعاته ومشاعره متى تغيرت ظروفه.\n"
        "   • الرسوم البيانية الحية (Chart.js): تحويل أرقام الباك إند إلى رسوم بيانية ومؤشرات دائرية ملونة.\n"
        "   • نظام التعريب واللغات (i18n.js): تبديل فوري بين العربية والإنجليزية."
    )
    add_box(s7, txt_noury, 1.0, 2.8, 5.2, 3.5, font_size=12, color=WHITE)

    # Right Card
    add_card(s7, 6.7, 2.1, 5.8, 4.4, BLUE, CARD_BG)
    add_box(s7, "📦 النتيجة النهائية للمشروع ككل", 6.9, 2.3, 5.4, 0.4, font_size=16, bold=True, color=BLUE)
    txt_noury_out = (
        "💻 المُنتج النهائي أمام المستخدم والدكتور:\n"
        "   • نظام مؤسسي متكامل ومتناسق 100%:\n"
        "     - واجهة مستخدم سريعة وخفيفة بدون أطر عمل معقدة.\n"
        "     - إرسال طلبات Fetch API غير متزامنة (Async) تضمن سلاسة الحركة.\n"
        "     - عرض تفسيرات الذكاء الاصطناعي (SHAP Explanations) في بطاقات واضحة.\n"
        "     - مؤشر استجابة فوري (Toast Notifications) عند حفظ البيانات بنجاح.\n\n"
        "🌟 الإنجاز المشترك:\n"
        "   • واجهة م. نوري أظهرت عبقرية موديل م. عبد الله، وحساسية نصوص م. شيماء، واستقرار باك إند م. أحمد، ودقة بيانات م. أسماء!"
    )
    add_box(s7, txt_noury_out, 6.9, 2.8, 5.4, 3.5, font_size=12, color=WHITE)

    add_box(s7, "🔑 سؤال الدكتور المتوقع: لو سأل 'هل الموظف يقدر يعدل حالته بنفسه؟' الإجابة: نعم، من خلال بوابة الموظف يقدر يدخل تقييم جديد ويشوف التغير اللحظي في درجة الخطر.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 8: End-to-End Live Scenario
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_slide_header(s8, "LIVE SCENARIO", "سيناريو حي يوضح التكامل: ماذا يحدث في الثانية الواحدة؟", "تتبع رحلة ضغطة زر واحدة من متصفح الموظف حتى ظهور النتيجة وتحليل الذكاء الاصطناعي")

    journey = [
        ("1. إدخال وتعديل الحالة", "الموظف في واجهة م. نوري يكتب: 'I feel exhausted with too many meetings' ويدخل 52 ساعة عمل.", GREEN),
        ("2. طلب الشبكة السريع", "المتصفح يرسل POST Request محمي بالـ JWT إلى سيرفر FastAPI الذي بناه م. أحمد.", BLUE),
        ("3. تنبؤ الأرقام ML", "م. أحمد يمرر الـ 52 ساعة لنموذج Random Forest الخاص بـ م. عبد الله ⬅️ يحسب Workload Risk = 88%.", YELLOW),
        ("4. تحليل المشاعر NLP", "م. أحمد يمرر جملة 'I feel exhausted' لنموذج DistilBERT الخاص بـ م. شيماء ⬅️ يحسب Sentiment Risk = 96%.", PURPLE),
        ("5. المعادلة الهجينة 60/40", "محرك م. أحمد يحسب: (88 * 0.6) + (96 * 0.4) = 91.2% (خطر احتراق وظيفي أحمر High Risk!).", BLUE),
        ("6. الحفظ والعرض اللحظي", "يتم حفظ السجل في PostgreSQL، وتضيء شاشة م. نوري بالأحمر مع كارت تفسير SHAP في 180ms فقط!", GREEN),
    ]

    card_h = 1.8
    for i, (title, desc, clr) in enumerate(journey):
        row = i // 3
        col = i % 3
        cx = 0.8 + col * (3.8 + 0.15)
        cy = 2.1 + row * (card_h + 0.35)
        add_card(s8, cx, cy, 3.8, card_h, clr, CARD_BG)
        add_box(s8, title, cx + 0.15, cy + 0.15, 3.5, 0.35, font_size=13, bold=True, color=clr)
        add_box(s8, desc, cx + 0.15, cy + 0.55, 3.5, 1.15, font_size=11, color=WHITE)

    add_box(s8, "⚡ السرعة والدقة: المنظومة تجمع بين سرعة الاستجابة اللحظية (Real-time) ودقة الذكاء الاصطناعي المزدوج دون أي تعارض.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=WHITE)

    # =========================================================================
    # SLIDE 9: Doctor Defense Cheat Sheet
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_slide_header(s9, "DEFENSE STRATEGY", "روشتة إجابة الدكتور: كيف تلخصون العمل الجماعي في دقيقة؟", "كلمات ذهبية للإجابة على سؤال لجنة المناقشة: 'كيف قسمتم العمل وجمعتوا المشروع؟'")

    add_card(s9, 0.8, 2.1, 11.7, 4.3, BLUE, CARD_BG)
    add_box(s9, "🎙️ الرد النموذجي المقترح في المناقشة (قولوا للدكتور كده بالظبط):", 1.1, 2.3, 11.0, 0.4, font_size=16, bold=True, color=YELLOW)

    speech = (
        "\"يا دكتور، مشروع WorkPulse اتقسم حسب معمارية الـ Pipeline المتكاملة:\n\n"
        "1. بدأنا مع البشمهندسة أسماء اللي هيأت الداتا النظيفة والواقعية لـ 50 موظف، وحطت المعايير الرياضية وسلّمت الداتا للفريق.\n\n"
        "2. بعد استلام الداتا، اشتغلنا بالتوازي: البشمهندس عبد الله درّب موديل Random Forest على الأرقام وحقق دقة 90%+ وصدّر الـ .pkl، وفي نفس الوقت البشمهندسة شيماء بنت خط تحليل المشاعر بـ DistilBERT لتقييم النصوص وسلّمت موديول الـ Sentiment.\n\n"
        "3. البشمهندس أحمد استلم الموديلات وبنى الباك إند وقاعدة البيانات PostgreSQL ومحرك الخطر الهجين (60/40) والتفسير بـ SHAP وحوّل كل ده لـ REST APIs سريعة ومحمية.\n\n"
        "4. البشمهندسة نوري استلمت الـ APIs وبنت الواجهات التفاعلية بـ Chart.js وبوابات المدير والموظف وإعادة التقييم بنظام ثنائي اللغة.\n\n"
        "النتيجة: كل عضو كان له مُخرَج واضح (Deliverable) الزميل التاني استلمه وبنى عليه، وده اللي خلانا نوصل لنظام متماسك وشغال Live بدون أخطاء!\""
    )
    add_box(s9, speech, 1.1, 2.8, 11.0, 3.4, font_size=13, color=WHITE)

    add_box(s9, "🏆 نصيحة ذهبية: إظهار التكامل بين الفريق والتسليم الواضح للمهام هو أكثر ما يُعجب لجنة المناقشة ويعكس عملاً هندسياً حقيقياً.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=GREEN)

    # =========================================================================
    # SLIDE 10: Input / Output Deliverables Matrix
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_slide_header(s10, "DELIVERABLES MATRIX", "مصفوفة التسليمات الدقيقة (Inputs vs Outputs Matrix)", "جدول مرجعي يحدد بدقة مسؤولية كل مهندس وملفات التسليم والاستلام")

    matrix_rows = [
        ("Eng. Assma", "فكرة المشروع ومعايير الضغط", "data/generate_data.py", "ملف data/employees.csv (50 موظف)", ORANGE),
        ("Eng. Abdullah", "بيانات employees.csv الرقمية", "train_model.py + Scikit-Learn", "models/workload_model.pkl + scaler.pkl", YELLOW),
        ("Eng. Shimaa", "نصوص الـ Feedbacks بالملف", "DistilBERT + HuggingFace PyTorch", "backend/sentiment.py (دالة sentiment_risk)", PURPLE),
        ("Eng. Ahmed", "الموديلات + الداتا + دالة المشاعر", "FastAPI + PostgreSQL + SHAP + Bcrypt", "7 Endpoints REST API + محرك الخطر 60/40", BLUE),
        ("Eng. Noury", "واجهات الـ REST API & Swagger", "HTML5 + JS + CSS3 + Chart.js", "Dashboard الإدارة + بوابة الموظف + التعريب", GREEN),
    ]

    card_h10 = 0.72
    for i, (eng, inp, tool, outp, clr) in enumerate(matrix_rows):
        cy = 2.05 + i * (card_h10 + 0.15)
        add_card(s10, 0.8, cy, 11.7, card_h10, clr, SURFACE)
        add_box(s10, eng, 1.0, cy + 0.18, 2.0, 0.35, font_size=13, bold=True, color=clr)
        add_box(s10, f"📥 المدخل: {inp}", 3.1, cy + 0.18, 2.9, 0.35, font_size=11, color=MUTED)
        add_box(s10, f"⚙️ الأداة: {tool}", 6.1, cy + 0.18, 2.8, 0.35, font_size=11, color=WHITE)
        add_box(s10, f"📦 المُخرَج: {outp}", 9.0, cy + 0.18, 3.4, 0.35, font_size=11, bold=True, color=clr)

    add_box(s10, "📌 ملحوظة: هذا الجدول يثبت للدكتور أن المشروع لم يُبْنَ عشوائياً، بل كمعمارية برمجية هندسية واضحة ومترابطة (Modular Architecture).", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=YELLOW)

    # =========================================================================
    # SLIDE 11: Live Demo Presentation Guide
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_slide_header(s11, "LIVE DEMO GUIDE", "خريطة تقديم الديمو الحي أمام الدكتور (Step-by-Step Live Demo)", "الخطوات العملية لاستعراض المشروع مباشرة على الشاشة لإبهار اللجنة")

    demo_steps = [
        ("الخطوة 1: تشغيل السيرفر", "فتح التيرمينال وتشغيل:\npython -m uvicorn backend.main:app --reload\nوعرض تحميل موديل Random Forest وموديل DistilBERT بنجاح.", BLUE),
        ("الخطوة 2: استعراض Swagger UI", "فتح الرابط: http://localhost:8000/docs\nلإظهار توثيق الـ APIs التفاعلي والـ Endpoints المحمية، وهذا يثبت عمق الباك إند.", PURPLE),
        ("الخطوة 3: لوحة المدير (Admin)", "تسجيل الدخول وعرض قائمة الـ 50 موظف، ومخططات توزيع نسب الخطورة (Chart.js) ومعدلات إجهاد الأقسام.", ORANGE),
        ("الخطوة 4: بوابة الموظف وتعديل الحالة", "الدخول كموظف وتوضيح ميزة إعادة التقييم الذاتي: إدخال ساعات عمل جديدة وشكوى جديدة ومشاهدة الخطر يتغير فورياً.", GREEN),
        ("الخطوة 5: شرح الذكاء الاصطناعي (SHAP)", "إظهار كارت التفسير الذي يوضح للمدير السبب الدقيق لارتفاع الخطر (مثلاً: +35% بسبب الساعات الإضافية).", YELLOW),
        ("الخطوة 6: التبديل للغة العربية", "الضغط على زر تغيير اللغة لإظهار جاهزية النظام للاستخدام في الشركات والمؤسسات العربية.", BLUE),
    ]

    for i, (title, desc, clr) in enumerate(demo_steps):
        row = i // 3
        col = i % 3
        cx = 0.8 + col * (3.8 + 0.15)
        cy = 2.1 + row * (1.8 + 0.35)
        add_card(s11, cx, cy, 3.8, 1.8, clr, CARD_BG)
        add_box(s11, title, cx + 0.15, cy + 0.15, 3.5, 0.35, font_size=13, bold=True, color=clr)
        add_box(s11, desc, cx + 0.15, cy + 0.55, 3.5, 1.15, font_size=11, color=WHITE)

    add_box(s11, "🎯 سر النجاح في المناقشة: الترتيب الواثق في عرض الشاشات يجعل الدكتور يركز في الإبهار التقني للمشروع.", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=WHITE)

    # =========================================================================
    # SLIDE 12: Tricky Questions & Bulletproof Answers
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_slide_header(s12, "TRICKY QUESTIONS & DEFENSE", "أهم 4 أسئلة تريكات من الدكتور والرد الهندسي القاطع", "تجهيز الفريق لأصعب الأسئلة النظرية والعملية التي قد تطرحها لجنة المناقشة")

    qa_list = [
        ("س1: ليه ما استخدمتوش Deep Learning كامل (Neural Network) للأرقام؟",
         "ج: لأن البيانات المجدولة (Tabular Data) أثبتت الدراسات أن Random Forest و XGBoost يتفوقان فيها على الشبكات العصبية، كما أن Random Forest لا يحتاج ملايين السجلات ولا يعاني من مشاكل التدريب المعقدة.", YELLOW),
        ("س2: ما الخطر إذا اعتمدنا على الأرقام فقط وتجاهلنا المشاعر؟",
         "ج: موظف قد يعمل 40 ساعة فقط لكنه يتعرض لضغط نفسي شديد أو مهام غير واضحة؛ الأرقام ستعطيه Low Risk بينما هو على حافة الاستقالة! دمج الـ NLP يكشف هذا الخطر الخفي.", PURPLE),
        ("س3: ما الفارق بين Random Forest و SHAP؟",
         "ج: Random Forest يتنبأ بنسبة الخطر (Black-box تقريباً)، أما SHAP فيشرح 'لماذا' اتخذ الموديل هذا القرار لكل موظف على حدة، وهو شرط أساسي في أنظمة الـ Explainable AI (XAI) المؤسسية.", BLUE),
        ("س4: كيف تضمنون أمان النظام وعدم تزوير التقييمات؟",
         "ج: استخدمنا تشفير كلمات السر بـ Bcrypt مع Salt عشوائي، والاتصال محمي برموز JWT tokens، وقاعدة البيانات تفصل صلاحيات الموظف عن صلاحيات المدير تماماً.", GREEN),
    ]

    for i, (q, a, clr) in enumerate(qa_list):
        row = i // 2
        col = i % 2
        cx = 0.8 + col * (5.75 + 0.2)
        cy = 2.1 + row * (1.95 + 0.2)
        add_card(s12, cx, cy, 5.75, 1.95, clr, CARD_BG)
        add_box(s12, q, cx + 0.15, cy + 0.15, 5.45, 0.45, font_size=12, bold=True, color=clr)
        add_box(s12, a, cx + 0.15, cy + 0.65, 5.45, 1.2, font_size=11, color=WHITE)

    add_box(s12, "🏆 عند الإجابة بهذه المصطلحات الهندسية الدقيقة، تضمنون الحصول على أعلى تقدير (امتياز) بإذن الله!", 0.8, 6.6, 11.7, 0.4, font_size=12, bold=True, color=GREEN)

    # Save presentation
    root_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(root_dir, "docs")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, "Tawazon_Team_Workflow_and_Handover.pptx")
    prs.save(out_path)
    print(f"Presentation successfully created at: {out_path}")


if __name__ == "__main__":
    create_deck()

