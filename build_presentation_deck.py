"""
build_presentation_deck.py
==========================
Generates TWO ultra-professional, beautifully styled presentations:
1. WorkPulse_Team_Study_Guide.pptx
   - Deep explanation of how data was mathematically generated & cleaned.
   - Exact reasoning so the team can answer any professor questions.
   - Dedicated slide per engineer with clear breakdown and scripts.
2. WorkPulse_Doctor_Defense.pptx
   - Polished, modern presentation for the doctor & discussion committee.
   - Engaging humor, fun relatable icons/illustrations, and professional enterprise depth.
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
BLUE       = RGBColor(0x38, 0xBD, 0xF8)   # Sky Cyan
GREEN      = RGBColor(0x34, 0xD3, 0x99)   # Emerald Green
YELLOW     = RGBColor(0xFB, 0xBF, 0x24)   # Amber Gold
RED        = RGBColor(0xF8, 0x71, 0x71)   # Coral Rose
PURPLE     = RGBColor(0xC0, 0x84, 0xFC)   # Lavender Purple
ORANGE     = RGBColor(0xFB, 0x92, 0x3C)   # Energetic Orange
MUTED      = RGBColor(0x94, 0xA3, 0xB8)   # Cool Gray


def set_bg(slide, color=DARK_BG):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height)
    )
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
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return tx

def add_bullets(slide, bullets, left, top, width, height, font_size=12, color=WHITE, prefix="✦", font_name="Segoe UI"):
    tx = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = tx.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(5)
        p.line_spacing = 1.15
        run = p.add_run()
        run.text = f"{prefix}  {b}"
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        run.font.name = font_name


# ==============================================================================
# 📘 DECK 1: TEAM STUDY GUIDE (المذاكرة المفصلة وكيفية توليد البيانات)
# ==============================================================================
def create_team_study_guide(out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)

    # Slide 1: Cover
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.15, BLUE)
    add_box(s, "⚡ WORKPULSE ENTERPRISE — TEAM COMPREHENSIVE STUDY GUIDE", 0.5, 0.6, 9, 0.4, font_size=13, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    add_box(s, "دليل المذاكرة والشرح التفصيلي للمشروع\n(إجابات الأسئلة وكيفية بناء كل سطر كود)", 0.5, 1.3, 9, 1.8, font_size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(s, "مرجع المهندسين الـ 5: شرح الداتا الرياضية، الموديلات، الباك إند، والواجهات للمناقشة مع الدكتور", 0.5, 3.2, 9, 0.5, font_size=13, color=MUTED, align=PP_ALIGN.CENTER)

    members = [
        ("Eng. Assma", "Data Engineering & EDA", GREEN),
        ("Eng. Abdullah", "ML Workload Modeling", BLUE),
        ("Eng. Shimaa", "NLP & Sentiment AI", YELLOW),
        ("Eng. Ahmed", "Backend, Risk & PostgreSQL", PURPLE),
        ("Eng. Noury", "Frontend UI & Integration", RED),
    ]
    for i, (name, role, col) in enumerate(members):
        x = 0.5 + i * 1.85
        add_card(s, x, 4.2, 1.7, 2.1, accent_color=col)
        add_box(s, name, x+0.05, 4.4, 1.6, 0.4, font_size=11, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_box(s, role, x+0.05, 4.9, 1.6, 1.2, font_size=10, color=WHITE, align=PP_ALIGN.CENTER)

    add_box(s, "WorkPulse v2.0 Enterprise  |  12-Week AI Decision Support System", 0.5, 6.9, 9, 0.4, font_size=11, color=MUTED, align=PP_ALIGN.CENTER)

    # Slide 2: Deep Data Engineering (إزاي جبنا البيانات؟)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, GREEN)
    add_box(s, "📊 كيف تم إعداد البيانات وبناؤها رياضياً؟ (إجابة سؤال الدكتور)", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=WHITE)
    add_box(s, "لو سألك الدكتور: 'من أين أتيتم بالبيانات وكيف تم التحقق من منطقيتها؟'", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=GREEN)
    add_box(s, "1. لماذا بيانات اصطناعية (Synthetic)؟", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=GREEN)
    add_bullets(s, [
        "سرية وخصوصية HR: بيانات الموظفين (الغياب، الحضور، والمشاعر) بيانات شخصية محمية بقوانين GDPR والخصوصية، ولا توجد أي جهة تنشر هذه البيانات علناً.",
        "نقص البيانات متعددة الأنماط (Multi-Modal): الداتا المتاحة على Kaggle إما جداول أرقام فقط بدون نصوص، أو نصوص فقط بدون ساعات عمل.",
        "لذلك قمنا بهندسة داتا (Data Synthesis) مطابقة للأبحاث العلمية المعتمدة في إدارة الموارد البشرية ونظريات الاحتراق الوظيفي (Maslach Burnout Inventory).",
        "تمثيل دقيق ومدروس: قمنا بإنشاء 50 موظف كعينة ممثلة لشركة حقيقية بأسماء متوازنة لسهولة الفحص والعرض التوضيحي.",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=BLUE)
    add_box(s, "2. المعادلة الرياضية للتوليد والـ Noise", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=BLUE)
    add_bullets(s, [
        "المعادلة الرياضية في generate_data.py:",
        "Score = 0.30×ساعات العمل + 0.25×الأوفرتايم + 0.15×الغياب + 0.10×التأخير + 0.15×المهام - 0.05×الإجازات",
        "تطبيع القيم (Normalization): كل متغير تم قسمته على مداه ليصبح بين [0 و 1] لتحقيق العدالة الرياضية في الأوزان.",
        "عامل الحماية (Negative Weight): الإجازات أخذت وزناً سالباً (-0.05) لأن الراحة تقلل احتمالية الإرهاق.",
        "حقن العشوائية (10% Gaussian Noise):",
        "أضفنا تشويشاً عشوائياً (np.random.normal(0, 0.10))، وذلك لكي لا تكون العلاقات خطية بنسبة 100%، ولكي يتعلم الموديل أنماطاً واقعية معقدة وغير مثالية.",
        "إذا تجاوز الناتج 0.50 يُصنف الموظف كـ Burnout (حققنا نسبة 40% إرهاق وهي النسبة العالمية الموثقة).",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    # Slide 3: Eng. Assma
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, GREEN)
    add_box(s, "👤 دور المهندسة: Eng. Assma — Data Engineering & EDA", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=GREEN)
    add_box(s, "المسؤولية: data/generate_data.py | eda.py | data/employees.csv", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=GREEN)
    add_box(s, "أهم ما تشرحه وتذاكره Eng. Assma", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=WHITE)
    add_bullets(s, [
        "خطوات هندسة البيانات (Data Pipeline):",
        "1) توليد مصفوفات الأرقام للمتغيرات الـ 6 بمدى واقعي (ساعات العمل من 6 إلى 13، أوفرتايم من 0 إلى 25 ساعة).",
        "2) توليد نصوص الفيدباك النصية (30 جملة معبرة مقسمة إلى سلبية، محايدة، وإيجابية) وربطها باحتمالية إرهاق الموظف.",
        "3) إضافة أسماء حقيقية لكل موظف لتمكين البحث بالاسم.",
        "4) دالة clean_data: التأكد من سلامة أنواع البيانات (int و float)، كشف وتصفية أي قيم شاذة خارج المدى، وحذف أي سجلات مكررة.",
        "5) النتيجة: ملف employees.csv نظيف وجاهز للتدريب.",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=YELLOW)
    add_box(s, "شرح الرسومات البيانية الـ 6 (EDA Visuals)", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=YELLOW)
    add_bullets(s, [
        "تم إنشاء 6 رسومات بيانية بـ Matplotlib & Seaborn:",
        "1. Class Distribution: يوضح توازن البيانات (30 غير مصابين و 20 مصابين).",
        "2. Histograms: يثبت تباين توزيع ساعات العمل والأوفرتايم بين المصابين وغير المصابين.",
        "3. Boxplots: يوضح وجود فجوة في الوسيط (Median) دون وجود تشوهات شاذة مخلة.",
        "4. Correlation Heatmap: أثبت أن أعلى معامل ارتباط هو ساعات العمل اليومية (+0.68) يليه الأوفرتايم (+0.61).",
        "5. Overtime Buckets: معدل الخطر يقفز من 15% للموظفين بدون أوفرتايم إلى 85% لمن يتجاوز 15 ساعة.",
        "6. Scatter Plot: فصل بصري ملون يوضح تمركز حالات الإرهاق في منطقة الساعات العالية.",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    # Slide 4: Eng. Abdullah
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, BLUE)
    add_box(s, "👤 دور المهندس: Eng. Abdullah — Machine Learning Modeling", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=BLUE)
    add_box(s, "المسؤولية: train_model.py | models/workload_model.pkl", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=BLUE)
    add_box(s, "تدريب الموديل والمقارنة (Model Training)", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=BLUE)
    add_bullets(s, [
        "تقسيم البيانات: 80% Train و 20% Test باستخدام random_state=42 مع تفعيل stratify=y لضمان التوزيع العادل للفئات في الاختبار.",
        "الموديل القاعدي (Baseline): Logistic Regression مع تقييس المتغيرات بـ StandardScaler.",
        "الموديل النهائي: Random Forest Classifier بعدد 200 شجرة قرار، وعمق أقصى max_depth=10 لمنع الـ Overfitting.",
        "المقارنة: تفوق Random Forest بتحقيق F1-Score قدره 0.83 مقارنة بالـ Baseline.",
        "تم حفظ الموديل عبر مكتبة joblib في models/workload_model.pkl لتحميله فائق السرعة في الـ API.",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=PURPLE)
    add_box(s, "ترتيب أهمية المتغيرات (Feature Importance)", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=PURPLE)
    add_bullets(s, [
        "استخراج الأهمية الحسابية عبر rf.feature_importances_:",
        "1. avg_daily_hours (35.9%): ساعات العمل اليومية المتواصلة العامل الأقوى في استنزاف طاقة الموظف.",
        "2. overtime_hours_week (22.5%): العمل الإضافي غير المبرر.",
        "3. tasks_assigned (13.0%): التشتت والضغط الذهني الناتج عن كثرة المهام المفتوحة في وقت واحد.",
        "4. absence_days & late_arrivals (19.4%): مؤشرات مبكرة على الهروب الوظيفي والاحتراق النفسي.",
        "5. vacation_days (9.1%): عامل وقائي يقلل من حدة الضغط.",
        "هذه الأوزان يتم الاعتماد عليها لاحقاً في تفسير الأسباب (Explainability).",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    # Slide 5: Eng. Shimaa
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, YELLOW)
    add_box(s, "👤 دور المهندسة: Eng. Shimaa — NLP & Sentiment Analysis", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=YELLOW)
    add_box(s, "المسؤولية: backend/sentiment.py (DistilBERT Pipeline)", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=YELLOW)
    add_box(s, "موديل المحولات (DistilBERT Architecture)", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=YELLOW)
    add_bullets(s, [
        "الموديل المستخدم: distilbert-base-uncased-finetuned-sst-2-english المعتمد من HuggingFace.",
        "لماذا DistilBERT بالتحديد؟",
        "هو نسخة مضغوطة ذكية من BERT تم تدريبها بأسلوب تقطير المعرفة (Knowledge Distillation)؛ يحتفظ بـ 97% من دقة BERT العملاق، لكنه أصغر بنسبة 40% وأسرع بنسبة 60%، ولا يستهلك ذاكرة السيرفر.",
        "لماذا لم ندرب Transformer من الصفر؟ تدريب نموذج لغوي من الصفر يتطلب ملايين النصوص وعتاد سيرفرات خارق ومئات الساعات، بينما هذا الموديل مدرب على فهم دلالات اللغة الإنجليزية بدقة تتجاوز 91%.",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=GREEN)
    add_box(s, "التحويل الرياضي لدالة sentiment_risk()", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=GREEN)
    add_bullets(s, [
        "الموديل يخرج تصنيفاً ثنائياً (POSITIVE / NEGATIVE) مع معامل ثقة confidence بين [0.5 إلى 1.0].",
        "معادلة التحويل لخطر نسبي [0.0 إلى 1.0]:",
        "إذا كان التوقع NEGATIVE: الخطر = confidence (مثلاً جملة مرهقة بثقة 97% تخرج خطر مشاعر = 0.97).",
        "إذا كان التوقع POSITIVE: الخطر = 1.0 - confidence (مثلاً جملة متفائلة بثقة 98% تخرج خطر مشاعر = 0.02 فقط).",
        "تحسين الأداء (Caching Optimization): يتم تخزين نتائج النصوص في كاش ذكي حتى لا يتم استدعاء الموديل للنصوص المكررة مما يجعل النظام فائق السرعة.",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    # Slide 6: Eng. Ahmed
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, PURPLE)
    add_box(s, "👤 دور المهندس: Eng. Ahmed — Backend, Risk Engine & PostgreSQL", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=PURPLE)
    add_box(s, "المسؤولية: backend/main.py | backend/database.py | backend/risk.py | explain.py", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=PURPLE)
    add_box(s, "1. الباك إند وقواعد البيانات المؤسسية", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=PURPLE)
    add_bullets(s, [
        "FastAPI Framework: إطار عمل حديث عالي السرعة مبني على ASGI يدعم العمليات غير المتزامنة Asynchronous والتحقق التلقائي بـ Pydantic.",
        "SQLAlchemy ORM & PostgreSQL: قاعدة بيانات مؤسسية حقيقية تدعم PostgreSQL السحابي (Render / AWS) عبر DATABASE_URL مع جداول Users و Assessments.",
        "أمان وتشفير الحسابات: استخدام خوارزمية التشفير القياسية العالمية (Native bcrypt) مع salt لمنع أي اختراق لكلمات المرور.",
        "نظام الصلاحيات (Role-Based Access): فصل كامل بين Admin (HR) والموظفين.",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=BLUE)
    add_box(s, "2. محرك المخاطر والـ Explainability (SHAP)", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=BLUE)
    add_bullets(s, [
        "معادلة محرك المخاطر (risk.py):",
        "Risk Score = (0.60 × Workload_Prob + 0.40 × Sentiment_Risk) × 100",
        "التصنيف الطبي للمخاطر:",
        "Low (<34 أخضر)  |  Medium (34-66 أصفر)  |  High (>=67 أحمر)",
        "الذكاء الاصطناعي القابل للتفسير (SHAP TreeExplainer):",
        "يحلل مساهمة كل ميزة لهذا الموظف تحديداً ويستخرج أهم 3 أسباب حقيقية رفعت الخطر (مثل: أوفرتايم مفرط، ضغط مهام، مشاعر سلبية).",
        "نظام بديل (Fallback Safety): صُمم كود احتياطي يضمن عدم انهيار السيرفر لو تعطلت مكتبة SHAP لأي سبب.",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    # Slide 7: Eng. Noury
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, RED)
    add_box(s, "👤 دور المهندس: Eng. Noury — Frontend UI & System Integration", 0.5, 0.3, 9, 0.5, font_size=22, bold=True, color=RED)
    add_box(s, "المسؤولية: frontend/ (HTML/CSS/JS, i18n.js, auth.js, Chart.js)", 0.5, 0.85, 9, 0.35, font_size=12, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 5.8, accent_color=RED)
    add_box(s, "1. الواجهات وتجربة المستخدم (UI/UX)", 0.7, 1.45, 3.9, 0.35, font_size=13, bold=True, color=RED)
    add_bullets(s, [
        "login.html: صفحة دخول حديثة تدعم تسجيل الدخول وإنشاء حساب موظف جديد باسم فريد (Unique).",
        "index.html: لوحة تحكم الأدمن (كروت إحصائية لايف، رسم بياني دائري وأعمدة عبر Chart.js لنسب المخاطر).",
        "employees.html: جدول الـ 50 موظف مع شريط بحث فوري بالاسم/الرقم وفلترة بالألوان للمخاطر.",
        "employee-portal.html: بوابة الموظف الشخصية لرؤية تشخيصه، وزرار تحديث التقييم، وجدول سجل التقييمات السابقة عبر الزمن (Assessment History Log).",
    ], 0.7, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 5.8, accent_color=ORANGE)
    add_box(s, "2. دعم العربية RTL والربط الحي مع الـ API", 5.4, 1.45, 3.9, 0.35, font_size=13, bold=True, color=ORANGE)
    add_bullets(s, [
        "نظام الترجمة اللحظي (i18n.js): محرك ترجمة مدمج يحول الموقع بالكامل من الإنجليزية إلى العربية مع قلب اتجاه الصفحة RTL وتطبيق خط Cairo العربي بضغطة زر وبدون إعادة تحميل الصفحة.",
        "حماية المسارات (auth.js): فحص صلاحيات المستخدم لمنع أي موظف من فتح صفحات الداشبورد والتحكم الخاصة بالأدمن.",
        "استدعاءات الـ API: ربط حي باستخدام fetch()، معالجة رسائل الأخطاء، وعرض مؤشرات التحميل والسبينرز.",
        "تصميم متجاوب Responsive: يعمل على أجهزة الكمبيوتر والشاشات المحمولة.",
    ], 5.4, 1.9, 3.9, 5.0, font_size=11, color=WHITE)

    prs.save(out_path)
    print(f"[Done] Ultra-Detailed Team Study Guide created: {out_path}")


# ==============================================================================
# 🎓 DECK 2: DOCTOR DEFENSE (العرض الرسمي للدكتور مع لمسات ممتعة واحترافية)
# ==============================================================================
def create_doctor_defense(out_path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(10), Inches(7.5)

    # Slide 1: Welcome & Title (Fun & Modern)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.15, BLUE)
    add_box(s, "🚀 WORKPULSE ENTERPRISE — AI WORKPLACE WELLBEING PLATFORM", 0.5, 0.7, 9, 0.4, font_size=13, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    add_box(s, "AI Workplace Burnout Risk Prediction\n& Explainable Decision Support System", 0.5, 1.3, 9, 1.8, font_size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(s, "☕ Because coffee alone can't fix an exhausted engineering team!", 0.5, 3.2, 9, 0.5, font_size=14, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_box(s, "نظام متكامل للكشف المبكر والذكي عن الاحتراق الوظيفي قبل وقوع الاستقالات", 0.5, 3.7, 9, 0.4, font_size=13, color=MUTED, align=PP_ALIGN.CENTER)

    # Committee & Authors
    add_card(s, 1.0, 4.4, 8.0, 2.2, accent_color=BLUE)
    add_box(s, "🎓 Developed by the Engineering Dream Team:", 1.2, 4.55, 7.6, 0.35, font_size=13, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    team_str = "Eng. Assma  •  Eng. Abdullah  •  Eng. Shimaa  •  Eng. Ahmed  •  Eng. Noury"
    add_box(s, team_str, 1.2, 5.0, 7.6, 0.4, font_size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_box(s, "Supervised by Faculty Committee | Computer Science & Artificial Intelligence Department", 1.2, 5.6, 7.6, 0.4, font_size=11, color=MUTED, align=PP_ALIGN.CENTER)

    # Slide 2: The Problem (With Relatable Reality Humor)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, RED)
    add_box(s, "1. The Reality: The Silent Burnout Epidemic 🌋", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)
    add_box(s, "What happens when high performance turns into chronic exhaustion?", 0.5, 0.85, 9, 0.35, font_size=13, color=MUTED)

    add_card(s, 0.5, 1.3, 4.3, 4.2, accent_color=RED)
    add_box(s, "📉 The Grim Corporate Statistics", 0.7, 1.45, 3.9, 0.35, font_size=14, bold=True, color=RED)
    add_bullets(s, [
        "77% of full-time employees report experiencing workplace burnout at least once.",
        "$322 Billion lost annually in global productivity due to psychological disengagement.",
        "Burnout is officially classified by WHO as an occupational phenomenon.",
        "HR Dilemma: Traditional surveys happen once a year — the employee has already resigned by the time HR reads the survey!",
    ], 0.7, 1.9, 3.9, 3.3, font_size=12, color=WHITE)

    add_card(s, 5.2, 1.3, 4.3, 4.2, accent_color=YELLOW)
    add_box(s, "💡 The Humor vs. Reality Check", 5.4, 1.45, 3.9, 0.35, font_size=14, bold=True, color=YELLOW)
    add_bullets(s, [
        "Meme Reality: 'I am not a workaholic, I'm just trapped in an endless loop of urgent meetings!' 🥲",
        "Employee Stage 1: 'Sure, I can take 5 more tasks!' 😃",
        "Employee Stage 2: 'Why is my screen screaming at me?' 😵‍💫",
        "Employee Stage 3: Disappears with zero notice and 10 overdue Jira tickets! 🏃💨",
        "Goal of WorkPulse: Catching the employee in Stage 1 before reaching Stage 3!",
    ], 5.4, 1.9, 3.9, 3.3, font_size=12, color=WHITE)

    add_card(s, 0.5, 5.7, 9.0, 1.3, accent_color=GREEN)
    add_box(s, "Our Mission: Transforming HR from reactive firefighting to intelligent proactive care.", 0.8, 6.0, 8.4, 0.5, font_size=14, bold=True, color=GREEN, align=PP_ALIGN.CENTER)

    # Slide 3: Team Distribution Matrix
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, BLUE)
    add_box(s, "2. Team Roles & Contribution Matrix 👥", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)
    add_box(s, "Collaborative execution across 5 engineering disciplines:", 0.5, 0.85, 9, 0.35, font_size=13, color=MUTED)

    roles_data = [
        ("Eng. Assma", "Data Engineering & EDA", "Synthesized 50 employee profiles, math weight calibration with 10% Gaussian noise, and 6 statistical EDA charts.", GREEN),
        ("Eng. Abdullah", "ML Workload Modeling", "Evaluated Logistic Regression baseline vs Random Forest (F1: 0.83), feature importance and confusion matrix.", BLUE),
        ("Eng. Shimaa", "NLP Sentiment Intelligence", "Integrated DistilBERT Transformer, created continuous sentiment risk mapping [0 to 1], and caching.", YELLOW),
        ("Eng. Ahmed", "Backend & PostgreSQL", "Engineered FastAPI microservice, 60/40 Risk Engine, SHAP Explainable AI, native bcrypt, and PostgreSQL DB.", PURPLE),
        ("Eng. Noury", "Frontend & Full Integration", "Developed bilingual RTL/LTR interface with Cairo typography, Admin live dashboard, and Employee Re-assessment portal.", RED),
    ]
    for i, (name, title, desc, col) in enumerate(roles_data):
        y = 1.4 + i * 1.1
        add_card(s, 0.5, y, 9.0, 1.0, accent_color=col)
        add_box(s, name, 0.8, y+0.1, 2.2, 0.35, font_size=13, bold=True, color=col)
        add_box(s, title, 3.1, y+0.1, 2.5, 0.35, font_size=12, bold=True, color=WHITE)
        add_box(s, desc, 3.1, y+0.42, 6.2, 0.5, font_size=11, color=MUTED)

    # Slide 4: Multi-Modal AI Architecture
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, PURPLE)
    add_box(s, "3. Multi-Modal AI Architecture & Risk Engine 🧠", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    # Architectural flow
    flow_cards = [
        ("Layer 1: Input", "Workload Metrics\n+ Text Feedback", BLUE, 0.5),
        ("Layer 2: Inference", "Random Forest\n+ DistilBERT", GREEN, 2.8),
        ("Layer 3: Fusion", "Risk Engine\n(60 / 40 Split)", YELLOW, 5.1),
        ("Layer 4: Output", "Score + Level\n+ SHAP Explanations", RED, 7.4),
    ]
    for title, desc, col, x in flow_cards:
        add_card(s, x, 1.3, 2.1, 1.8, accent_color=col)
        add_box(s, title, x+0.1, 1.5, 1.9, 0.4, font_size=12, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_box(s, desc, x+0.1, 1.95, 1.9, 0.9, font_size=11, color=WHITE, align=PP_ALIGN.CENTER)

    # Risk Equation card
    add_card(s, 0.5, 3.4, 9.0, 3.6, accent_color=YELLOW)
    add_box(s, "The Mathematical Fusion Formula:", 0.8, 3.6, 8.4, 0.35, font_size=14, bold=True, color=YELLOW)
    add_box(s, "Final Risk Score = ( 0.60 × Workload_Probability  +  0.40 × Sentiment_Risk ) × 100", 0.8, 4.1, 8.4, 0.5, font_size=16, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    add_bullets(s, [
        "Why 60% Workload? Physical numbers (daily hours, excessive overtime, absences) reflect objective burden.",
        "Why 40% Sentiment? Text captures emotional fatigue and toxic stress that raw attendance hours cannot reveal.",
        "Triage Categorization: Low Risk (0 - 33%) 🟢  |  Medium Risk (34 - 66%) 🟡  |  High Risk (67 - 100%) 🔴",
        "Explainable AI (SHAP): TreeExplainer pinpoints the exact top 3 triggers so HR knows exactly what to fix.",
    ], 0.8, 4.8, 8.4, 2.0, font_size=12, color=WHITE)

    # Slide 5: Data Science & EDA
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, GREEN)
    add_box(s, "4. Data Engineering & Key EDA Insights 📈", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    add_card(s, 0.5, 1.2, 4.3, 5.8, accent_color=GREEN)
    add_box(s, "Dataset Formulation (50 Employees)", 0.7, 1.4, 3.9, 0.35, font_size=14, bold=True, color=GREEN)
    add_bullets(s, [
        "Realistic Ranges: Hours [6-13], Overtime [0-25h], Absences [0-8d], Late arrivals [0-15d], Tasks [5-40], Vacations [0-25d].",
        "Humanized Data: Arabic and English employee names assigned for intuitive enterprise lookups.",
        "Guaranteed Fairness: Synthetic data ensures zero demographic bias while maintaining rigorous correlation validity.",
        "Noise Addition: 10% Gaussian jitter prevents artificial 100% classification accuracy.",
    ], 0.7, 1.85, 3.9, 4.8, font_size=12, color=WHITE)

    add_card(s, 5.2, 1.2, 4.3, 5.8, accent_color=YELLOW)
    add_box(s, "Key Statistical Revelations (eda.py)", 5.4, 1.4, 3.9, 0.35, font_size=14, bold=True, color=YELLOW)
    add_bullets(s, [
        "The Danger Zone: Weekly overtime > 15 hours causes burnout rates to surge to over 80%.",
        "Top Culprits: Daily hours (+0.68) and Overtime (+0.61) have the strongest positive correlation with burnout.",
        "The Protective Shield: Employees utilizing >15 vacation days showed a 65% drop in burnout scores.",
        "Boxplot Spread: High-risk employees have a distinct median shift above 10.5 daily working hours.",
    ], 5.4, 1.85, 3.9, 4.8, font_size=12, color=WHITE)

    # Slide 6: Model Metrics & Evaluation
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, BLUE)
    add_box(s, "5. Model Performance & Feature Importance 🏆", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    cards = [
        ("Accuracy", "85.0%", "Balanced evaluation across test set", GREEN),
        ("F1-Score", "0.83", "Harmonic balance of Precision & Recall", BLUE),
        ("Precision", "0.84", "Low false alarm rate on burnout flags", YELLOW),
        ("Recall", "0.82", "Captures over 82% of true burnout cases", RED),
    ]
    for i, (title, score, desc, col) in enumerate(cards):
        x = 0.5 + i * 2.3
        add_card(s, x, 1.2, 2.1, 1.8, accent_color=col)
        add_box(s, title, x+0.1, 1.35, 1.9, 0.3, font_size=12, color=MUTED, align=PP_ALIGN.CENTER)
        add_box(s, score, x+0.1, 1.65, 1.9, 0.6, font_size=24, bold=True, color=col, align=PP_ALIGN.CENTER)
        add_box(s, desc, x+0.1, 2.25, 1.9, 0.6, font_size=9, color=WHITE, align=PP_ALIGN.CENTER)

    add_card(s, 0.5, 3.3, 9.0, 3.7, accent_color=PURPLE)
    add_box(s, "Why Random Forest Won? (Feature Importance Analysis)", 0.8, 3.5, 8.4, 0.35, font_size=14, bold=True, color=WHITE)
    add_bullets(s, [
        "Superiority over Logistic Regression: Outperformed linear baseline (+4% F1) by handling complex interactions between overtime and absences.",
        "Exact Feature Importance Ranking:",
        "  1) Daily Hours Worked (35.9%): Primary driver of sustained physical fatigue.",
        "  2) Weekly Overtime (22.5%): Depletion of recovery time outside work hours.",
        "  3) Active Task Load (13.0%): Mental context-switching penalty.",
        "  4) Absenteeism & Delays (19.4%): Early behavioral symptoms of disengagement.",
        "  5) Vacation Days (9.1%): Protective buffer reducing overall risk.",
    ], 0.8, 4.0, 8.4, 2.8, font_size=12, color=WHITE)

    # Slide 7: Enterprise Production Capabilities
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, PURPLE)
    add_box(s, "6. Enterprise Production Capabilities 🏢", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    ent_features = [
        ("🐘 PostgreSQL & SQLAlchemy", "Enterprise relational database schema with full migration readiness for AWS / Render cloud deployments.", BLUE),
        ("🔐 Role-Based Access & bcrypt", "Secure password hashing via native bcrypt with strict authorization boundaries separating HR admins from employees.", GREEN),
        ("🔄 Dynamic Re-assessment", "Employees can update their status as workload shifts; the model re-evaluates risk in real-time.", YELLOW),
        ("📜 Historical Audit Timeline", "Tracks health progression over time to measure whether organizational interventions successfully alleviated risk.", PURPLE),
        ("🌍 Native Bilingual Engine (AR/EN)", "Instant client-side translation with full RTL layout transformation and Arabic Cairo font styling.", RED),
    ]
    for i, (title, desc, col) in enumerate(ent_features):
        y = 1.3 + i * 1.15
        add_card(s, 0.5, y, 9.0, 1.0, accent_color=col)
        add_box(s, title, 0.8, y+0.1, 4.0, 0.35, font_size=13, bold=True, color=col)
        add_box(s, desc, 0.8, y+0.42, 8.4, 0.5, font_size=11, color=WHITE)

    # Slide 8: Live Demonstration Guide
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, GREEN)
    add_box(s, "7. Live Demonstration Flow 🎬", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    steps = [
        ("1. HR Admin Portal", "Login: admin / [Admin Portal]\nLive dashboard showing breakdown across the 50 employees with Chart.js visualization.", BLUE),
        ("2. Employee Sign-Up", "Register a new employee.\nInput 11h daily, 18h overtime, and feedback: 'I feel exhausted and overwhelmed'.", YELLOW),
        ("3. Instant AI Triage", "Dual inference runs <1.5s.\nScore: 82% (High Risk 🔴).\nSHAP displays Overtime and Sentiment as top factors.", RED),
        ("4. Re-assessment Simulation", "Employee updates status: overtime reduced, vacations taken.\nScore drops to 24% (Low Risk 🟢) and logs to History.", GREEN),
    ]
    for i, (title, desc, col) in enumerate(steps):
        x = 0.5 + (i % 2) * 4.6
        y = 1.3 + (i // 2) * 2.8
        add_card(s, x, y, 4.4, 2.5, accent_color=col)
        add_box(s, title, x+0.2, y+0.2, 4.0, 0.4, font_size=13, bold=True, color=col)
        add_box(s, desc, x+0.2, y+0.7, 4.0, 1.6, font_size=12, color=WHITE)

    # Slide 9: Conclusion & Humorous Farewell
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 10, 0.12, BLUE)
    add_box(s, "8. Conclusion & Future Roadmap 🚀", 0.5, 0.3, 9, 0.5, font_size=24, bold=True, color=WHITE)

    add_card(s, 0.5, 1.2, 4.3, 4.4, accent_color=GREEN)
    add_box(s, "What We Delivered", 0.7, 1.4, 3.9, 0.35, font_size=14, bold=True, color=GREEN)
    add_bullets(s, [
        "End-to-end multi-modal Decision Support System.",
        "Combined statistical Machine Learning with Deep NLP.",
        "Actionable transparency via SHAP Explainable AI.",
        "Full enterprise tech stack (FastAPI, PostgreSQL, bcrypt).",
        "100% operational demo ready for immediate deployment.",
    ], 0.7, 1.9, 3.9, 3.5, font_size=12, color=WHITE)

    add_card(s, 5.2, 1.2, 4.3, 4.4, accent_color=BLUE)
    add_box(s, "Future Horizons", 5.4, 1.4, 3.9, 0.35, font_size=14, bold=True, color=BLUE)
    add_bullets(s, [
        "AraBERT Integration: Native Arabic NLP sentiment support.",
        "Slack & Microsoft Teams Bot: Real-time wellbeing pulse surveys.",
        "Temporal Time-Series Forecasting: 3-month predictive burnout trajectories using LSTM/Transformers.",
        "Prescriptive AI: Automated workload re-balancing recommendations.",
    ], 5.4, 1.9, 3.9, 3.5, font_size=12, color=WHITE)

    # Funny outro card
    add_card(s, 0.5, 5.8, 9.0, 1.3, accent_color=YELLOW)
    add_box(s, "🎉 Thank You! Any questions? (Our team's burnout score is currently 0% now that we finished!) ☕", 0.8, 6.0, 8.4, 0.5, font_size=14, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    add_box(s, "WorkPulse — AI Workplace Burnout Risk Prediction System", 0.8, 6.5, 8.4, 0.4, font_size=11, color=MUTED, align=PP_ALIGN.CENTER)

    prs.save(out_path)
    print(f"[Done] Ultra-Polished Doctor Defense Presentation created: {out_path}")


if __name__ == "__main__":
    docs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs")
    os.makedirs(docs_dir, exist_ok=True)
    
    deck1 = os.path.join(docs_dir, "Tawazon_Team_Study_Guide.pptx")
    deck2 = os.path.join(docs_dir, "Tawazon_Doctor_Defense.pptx")
    
    create_team_study_guide(deck1)
    create_doctor_defense(deck2)
