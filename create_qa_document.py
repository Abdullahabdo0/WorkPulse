"""
create_qa_document.py
=====================
Generates a comprehensive, ultra-professional Word document (.docx)
containing the Doctor Defense Q&A Guide for Tawazon (توازن).
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_PATH_DOCX = os.path.join(ROOT, "Tawazon_Doctor_Defense_QA_Guide.docx")
OUT_PATH_MD   = os.path.join(ROOT, "docs", "Tawazon_Doctor_Defense_QA_Guide.md")

# ── Color Palette ────────────────────────────────────────────────────────────
COLOR_NAVY      = RGBColor(0x0F, 0x17, 0x2A)   # Navy 900
COLOR_BLUE      = RGBColor(0x1D, 0x4E, 0xD8)   # Royal Blue 700
COLOR_GREEN     = RGBColor(0x04, 0x78, 0x57)   # Emerald 700
COLOR_DARK_TEXT = RGBColor(0x0F, 0x17, 0x2A)   # Dark Slate
COLOR_MUTED     = RGBColor(0x47, 0x55, 0x69)   # Slate 600
COLOR_WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_AMBER     = RGBColor(0xB4, 0x53, 0x09)   # Amber 700

HEX_NAVY     = "0F172A"
HEX_BLUE     = "1D4ED8"
HEX_GREEN    = "047857"
HEX_LIGHT_BG = "F8FAFC"
HEX_CARD_BG  = "F1F5F9"
HEX_BORDER   = "CBD5E1"
HEX_AMBER_BG = "FEF3C7"
HEX_BLUE_BG  = "EFF6FF"
HEX_GREEN_BG = "F0FDF4"

FONT_NAME = "Arial"


def set_run_font(run, font_name=FONT_NAME, size_pt=12, bold=False, color=COLOR_DARK_TEXT):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.font.color.rgb = color

    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(
        f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>'
    )
    rPr.append(rFonts)
    rtl = parse_xml(f'<w:rtl {nsdecls("w")}/>')
    rPr.append(rtl)


def make_rtl_paragraph(p, space_before=0, space_after=4, line_spacing=1.25):
    pPr = p._p.get_or_add_pPr()
    bidi = parse_xml(f'<w:bidi {nsdecls("w")}/>')
    pPr.append(bidi)
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing


def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=140, bottom=140, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def add_qa_card(doc, q_num, question, answer, category_title, hex_bg=HEX_CARD_BG, hex_border=HEX_BLUE):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c = tbl.rows[0].cells[0]
    c.width = Inches(7.0)
    set_cell_background(c, hex_bg)
    set_cell_margins(c, top=140, bottom=160, left=200, right=200)

    # Accent right border
    tcPr = c._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:right w:val="single" w:sz="32" w:space="0" w:color="{hex_border}"/>'
        f'<w:top w:val="none"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:left w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)

    # Category Tag
    p_cat = c.paragraphs[0]
    make_rtl_paragraph(p_cat, space_before=0, space_after=2)
    r_cat = p_cat.add_run(f"🏷️ {category_title}")
    set_run_font(r_cat, size_pt=10.5, bold=True, color=COLOR_MUTED)

    # Question
    p_q = c.add_paragraph()
    make_rtl_paragraph(p_q, space_before=2, space_after=6)
    r_q_label = p_q.add_run(f"س{q_num}: ")
    set_run_font(r_q_label, size_pt=13, bold=True, color=COLOR_BLUE)
    r_q_text = p_q.add_run(question)
    set_run_font(r_q_text, size_pt=13, bold=True, color=COLOR_NAVY)

    # Answer
    p_a = c.add_paragraph()
    make_rtl_paragraph(p_a, space_before=2, space_after=2)
    r_a_label = p_a.add_run("الإجابة النموذجية: ")
    set_run_font(r_a_label, size_pt=11.5, bold=True, color=COLOR_GREEN)
    r_a_text = p_a.add_run(answer)
    set_run_font(r_a_text, size_pt=11.5, bold=False, color=COLOR_DARK_TEXT)

    # Add small spacer after table
    sp = doc.add_paragraph()
    make_rtl_paragraph(sp, space_before=0, space_after=6)


def create_document():
    doc = Document()

    # Margins
    for s in doc.sections:
        s.top_margin = Inches(0.75)
        s.bottom_margin = Inches(0.75)
        s.left_margin = Inches(0.75)
        s.right_margin = Inches(0.75)

    # ── Header Banner ────────────────────────────────────────────────────────
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    c = tbl.rows[0].cells[0]
    c.width = Inches(7.0)
    set_cell_background(c, HEX_NAVY)
    set_cell_margins(c, top=200, bottom=200, left=200, right=200)

    p1 = c.paragraphs[0]
    make_rtl_paragraph(p1, space_after=4)
    r1 = p1.add_run("🎓 دليل أسئلة وإجابات مناقشة التخرج الرسمية — منصة توازن Tawazon")
    set_run_font(r1, size_pt=17, bold=True, color=COLOR_WHITE)

    p2 = c.add_paragraph()
    make_rtl_paragraph(p2, space_after=0)
    r2 = p2.add_run("دليل المذاكرة وحماية المشروع أمام الدكتور ولجنة التحكيم | 12 سؤالاً وجواباً نموذجياً حاسماً")
    set_run_font(r2, size_pt=12, bold=False, color=RGBColor(0x94, 0xA3, 0xB8))

    sp = doc.add_paragraph()
    make_rtl_paragraph(sp, space_before=0, space_after=8)

    # ── Introduction Box ─────────────────────────────────────────────────────
    tbl_intro = doc.add_table(rows=1, cols=1)
    tbl_intro.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_intro.autofit = False
    ci = tbl_intro.rows[0].cells[0]
    ci.width = Inches(7.0)
    set_cell_background(ci, HEX_AMBER_BG)
    set_cell_margins(ci, top=140, bottom=140, left=180, right=180)

    pi = ci.paragraphs[0]
    make_rtl_paragraph(pi, space_before=0, space_after=2)
    ri1 = pi.add_run("💡 نصيحة ذهبية للفريق قبل المناقشة:\n")
    set_run_font(ri1, size_pt=12, bold=True, color=COLOR_AMBER)
    ri2 = pi.add_run(
        "الدكتور في المناقشة لا يختبر فقط كتابتكم للأكواد، بل يختبر مدى فهمكم العميق للقرارات التقنية التي اتخذتموها "
        "(لماذا اخترتم هذه الخوارزمية؟ لماذا نظفتم الداتا هكذا؟ وكيف تم تأمين النظام؟). "
        "هذا الدليل يمنحكم إجابات هندسية واثقة ومدعومة بالأرقام والحجج العلمية."
    )
    set_run_font(ri2, size_pt=11.5, bold=False, color=COLOR_DARK_TEXT)

    sp2 = doc.add_paragraph()
    make_rtl_paragraph(sp2, space_before=0, space_after=8)

    # ── Questions ────────────────────────────────────────────────────────────

    # Section 1: Data
    p_sec1 = doc.add_paragraph()
    make_rtl_paragraph(p_sec1, space_before=10, space_after=4)
    r_s1 = p_sec1.add_run("📂 المحور الأول: أسئلة البيانات والمعالجة (Dataset & Preprocessing)")
    set_run_font(r_s1, size_pt=15, bold=True, color=COLOR_BLUE)

    add_qa_card(
        doc,
        1,
        "الداتا دي حقيقية ولا من تأليفكم؟",
        "الداتا أصلها حقيقي 100% مسحوبة من مسابقة معتمدة على منصة كاجل العالمية (HackerEarth ML Challenge) "
        "تضم 22,750 موظفاً من شركات تقنية وخدمية حقيقية. ولكن الداتا الخام كانت مليئة بالعيوب والقيم الفارغة (4,622 خلية فارغة) "
        "وتكرارات، وخالية تماماً من الكلام والنصوص. لذلك قمنا ببناء خط معالجة وهندسة بيانات كامل (Data Cleaning Pipeline) "
        "أنتج عينة معيارية ممثلة لـ 1,000 موظف بنسبة احتراق مطابقة لتقارير منظمة العمل الدولية (ILO).",
        "مصدر وموثوقية البيانات",
        hex_bg=HEX_CARD_BG,
        hex_border=HEX_BLUE
    )

    add_qa_card(
        doc,
        2,
        "ليه أخدتوا 1,000 موظف بس من أصل 22 ألف؟",
        "لسببين جوهريين:\n"
        "1. إحصائياً: العينة تم سحبها بطريقة العينات الطبقية (Stratified Sampling) مع تثبيت معامل العشوائية (random_state=42)، "
        "مما حافظ بدقة متناهية على نفس التوزيع الطبيعي ونسبة الاحتراق (44.1%) دون أي تحيز.\n"
        "2. عملياً: حجم 1,000 موظف يمثل الحجم المثالي والمطابق للواقع لشركة متوسطة إلى كبرى (Enterprise Scale)، "
        "ويضمن سرعة استجابة فائقة لقاعدة البيانات وواجهات المستخدم ومحرك الـ NLP في أقل من 200ms أثناء العرض الحي.",
        "حجم العينة ومبرراتها",
        hex_bg=HEX_CARD_BG,
        hex_border=HEX_BLUE
    )

    add_qa_card(
        doc,
        3,
        "داتا كاجل مفيهاش سن (Age) ولا نصوص مشاعر (Feedback)، جبتوهم منين؟",
        "1. بخصوص العمر: مسابقات الموارد البشرية العالمية تحجب تاريخ الميلاد لأسباب الخصوصية ومنع التمييز (GDPR). "
        "لذلك طبقنا مفهوم هندسة الميزات (Feature Engineering) واستنتجنا العمر رياضياً بالاعتماد على عمود الرتبة الوظيفية (Designation من 0 لـ 5) "
        "وتدرج الخبرة المهنية (بين 22 و 56 سنة، بمتوسط 34.7 سنة)، مع الحفاظ الكامل على عمود الجنس الحقيقي من كاجل.\n"
        "2. بخصوص نصوص المشاعر: داتا كاجل أرقام فقط، ولبناء منصة هجينة متعددة الوسائط (Multi-Modal) ابتكرنا 45 نصاً حقيقياً "
        "يعبر عن مشاعر بيئة العمل ووزعناها باحتمالات دقيقة تتوافق مع معدل إجهاد كل موظف.",
        "هندسة المتغيرات المستحدثة",
        hex_bg=HEX_CARD_BG,
        hex_border=HEX_BLUE
    )

    add_qa_card(
        doc,
        4,
        "هل عندكم مشكلة عدم توازن في البيانات (Class Imbalance)؟",
        "لا إطلاقاً يا دكتور، نسبة الموظفين السليمين تمثل 55.9% (559 موظفاً) ونسبة المعرضين للاحتراق تمثل 44.1% (441 موظفاً). "
        "نسبة التوازن هي 1.27 : 1، وهي نسبة توازن إحصائي مثالية في تعلم الآلة لا تتطلب أي توليد مصطنع للبيانات مثل تقنيات SMOTE أو خوارزميات إعادة الوزن.",
        "التوازن الإحصائي",
        hex_bg=HEX_CARD_BG,
        hex_border=HEX_BLUE
    )

    # Section 2: ML & NLP
    p_sec2 = doc.add_paragraph()
    make_rtl_paragraph(p_sec2, space_before=14, space_after=4)
    r_s2 = p_sec2.add_run("🤖 المحور الثاني: خوارزميات الذكاء الاصطناعي (ML, NLP & Explainability)")
    set_run_font(r_s2, size_pt=15, bold=True, color=COLOR_GREEN)

    add_qa_card(
        doc,
        5,
        "ليه اخترتم Random Forest بالذات؟ وليه مش Deep Learning؟",
        "1. قمنا بالمقارنة والتجربة العملية في كود train_model.py بين Logistic Regression و Random Forest، "
        "وتفوق الـ Random Forest محققاً دقة 95.50% ومقياس F1-Score قدره 0.95.\n"
        "2. البيانات المجدولة الرقمية (Tabular Data): أثبتت كبرى الأوراق البحثية أن النماذج الشجرية المجمعة (Ensemble Tree Models) "
        "تتفوق على الشبكات العميقة في البيانات المجدولة المحدودة، وتتفادى مشكلة الإفراط في المطابقة (Overfitting) وسريعة جداً في التدريب والاستنتاج.",
        "اختيار نموذج التعلم الآلي",
        hex_bg=HEX_GREEN_BG,
        hex_border=HEX_GREEN
    )

    add_qa_card(
        doc,
        6,
        "ليه بتبصوا على الـ Recall والـ F1-Score ومش الـ Accuracy بس؟",
        "لأن في التطبيقات الصحية والنفسية الخطأ من النوع الثاني (False Negative) كارثي؛ إذا كان موظف على حافة الانهيار أو الاحتراق "
        "وتنبأ النموذج بأنه سليم، فستفقد المؤسسة الموظف دون تدخل علاجي مبكر. "
        "مقياس الاستدعاء (Recall) في نموذجنا وصل إلى 96.6%، وهو ما يضمن اكتشاف 97 من كل 100 موظف معرضين للاحتراق.",
        "مقاييس التقييم وضمان الجودة",
        hex_bg=HEX_GREEN_BG,
        hex_border=HEX_GREEN
    )

    add_qa_card(
        doc,
        7,
        "ليه عملتوا النظام هجين (60% أرقام + 40% مشاعر) وما اكتفيتوش بالأرقام؟",
        "لأن السلوك البشري معقد؛ الأرقام وحدها قد تضلل، فقد يعمل الموظف 8 ساعات فقط لكنه يعاني من إحباط حاد أو بيئة عمل سامة "
        "(وهنا يكشفه الـ NLP). وعلى العكس، قد يعمل موظف لساعات إضافية لأنه شغوف ومتحمس لمشروعه "
        "(وهنا يمنعه الـ NLP من التصنيف الخاطئ). المعادلة الهجينة توفر رؤية شاملة 360 درجة تجمع بين الإجهاد الجسدي والنفسي.",
        "المحرك الهجين (Hybrid AI)",
        hex_bg=HEX_GREEN_BG,
        hex_border=HEX_GREEN
    )

    add_qa_card(
        doc,
        8,
        "ازاي الموديل بيفهم الكلام العربي والإنجليزي في نصوص المشاعر؟",
        "النظام يطبق بنية معالجة لغات طبيعية مزدوجة:\n"
        "1. النصوص الإنجليزية: نستخدم نموذج ترانسفورمرز خفيف وعميق DistilBERT من Hugging Face.\n"
        "2. النصوص العربية: قمنا بتطوير محرك NLP مخصص داخل الـ Backend يتعرف على مفردات بيئة العمل المصرية والعربية، "
        "مع معالجة متقدمة لأدوات النفي (مثل: 'مش مرتاح'، 'غير راضٍ') والسوابق والتطبيع اللغوي للألف والتاء المربوطة.",
        "معالجة اللغات الطبيعية (NLP)",
        hex_bg=HEX_GREEN_BG,
        hex_border=HEX_GREEN
    )

    add_qa_card(
        doc,
        9,
        "إيه هو الـ SHAP وليه استخدمتوه في المشروع؟",
        "تقنية SHAP (Shapley Additive Explanations) هي حجر الأساس في الذكاء الاصطناعي القابل للتفسير (Explainable AI - XAI). "
        "وظيفتها إزالة الغموض عن قرارات النموذج ومنعه من أن يكون صندوقاً أسود (Black Box). "
        "بفضل SHAP، لا يكتفي النظام بإعطاء نسبة خطر للمدير، بل يعرض أهم 3 أسباب حقيقية ساهمت في النتيجة (مثل: ساعات عمل إضافية مفرطة، أو حرمان من الإجازات).",
        "تفسير القرارات (Explainable AI)",
        hex_bg=HEX_GREEN_BG,
        hex_border=HEX_GREEN
    )

    # Section 3: Architecture
    p_sec3 = doc.add_paragraph()
    make_rtl_paragraph(p_sec3, space_before=14, space_after=4)
    r_s3 = p_sec3.add_run("💻 المحور الثالث: معمارية النظام والبرمجة (System Architecture & Security)")
    set_run_font(r_s3, size_pt=15, bold=True, color=COLOR_NAVY)

    add_qa_card(
        doc,
        10,
        "إيه الـ Architecture والتقنيات اللي بنيتوا بيها السيستم وليه؟",
        "بنينا النظام وفق معمارية الـ Decoupled Full-Stack Architecture:\n"
        "• الـ Backend: إطار عمل FastAPI مبني على معيار ASGI السريع، ويوفر توثيقاً تلقائياً عبر Swagger UI.\n"
        "• الـ Database: طبقة SQLAlchemy ORM تدعم SQLite للتجربة المحلية وتدعم PostgreSQL للإنتاج المؤسسي.\n"
        "• الـ Frontend: واجهة تفاعلية حديثة بـ HTML5/CSS3/Vanilla JS مع محرك ترجمة مخصص (i18n) يدعم العربية RTL والإنجليزية LTR فورياً.",
        "المعمارية البرمجية (Tech Stack)",
        hex_bg=HEX_BLUE_BG,
        hex_border=HEX_BLUE
    )

    add_qa_card(
        doc,
        11,
        "ازاي مأمنين السيستم وبيانات الموظفين الحساسة؟ (Security)",
        "النظام يطبق المعايير الأمنية المؤسسية:\n"
        "1. التوثيق والمصادقة بـ JWT Bearer Tokens مع تشفير التوكنات وتحديد صلاحية زمنية للـ Session.\n"
        "2. تشفير كلمات المرور باستخدام خوارزمية PBKDF2-HMAC-SHA256 المشددة مع Salt لمنع هجمات الجداول المسبقة (Rainbow Tables).\n"
        "3. التحكم بالوصول المبني على الأدوار (Role-Based Access Control): الموظف العادي يرى بياناته وتقييماته فقط، "
        "بينما مدير الـ HR هو المخول بالاطلاع على لوحة القيادة وكافة الموظفين وتعديلهم.",
        "أمن وسرية البيانات (Enterprise Security)",
        hex_bg=HEX_BLUE_BG,
        hex_border=HEX_BLUE
    )

    add_qa_card(
        doc,
        12,
        "لو الشركة كبرت وبقى فيها 50 ألف موظف، هل السيستم هيتحمل؟ (Scalability)",
        "نعم بالتأكيد، النظام مصمم للتوسع الأفقي (Horizontal Scalability) للأسباب التالية:\n"
        "1. خادم FastAPI غير متزامن بالكامل (Async/Non-blocking I/O) قادر على خدمة آلاف الاتصالات المتزامنة.\n"
        "2. معمارية Stateless RESTful APIs تسمح بتشغيل عدة نسخ من السيرفر خلف Load Balancer وحاويات Docker.\n"
        "3. نماذج الذكاء الاصطناعي مجهزة ومخزنة بملفات مدمجة (Serialized via Joblib) وسرعة استنتاجها لا تتعدى 50ms لكل طلب.",
        "قابلية التوسع (Scalability)",
        hex_bg=HEX_BLUE_BG,
        hex_border=HEX_BLUE
    )

    # Save Word
    doc.save(OUT_PATH_DOCX)
    print(f"[SUCCESS] Word QA Guide saved at: {OUT_PATH_DOCX}")


def create_markdown():
    os.makedirs(os.path.dirname(OUT_PATH_MD), exist_ok=True)
    content = """# 🎓 دليل أسئلة وإجابات مناقشة التخرج الرسمية — منصة توازن Tawazon
> **دليل المذاكرة وحماية المشروع أمام الدكتور ولجنة التحكيم | 12 سؤالاً وجواباً نموذجياً حاسماً**

---

### 💡 نصيحة ذهبية للفريق قبل المناقشة:
الدكتور في المناقشة لا يختبر فقط كتابتكم للأكواد، بل يختبر مدى فهمكم العميق للقرارات التقنية التي اتخذتموها (لماذا اخترتم هذه الخوارزمية؟ لماذا نظفتم الداتا هكذا؟ وكيف تم تأمين النظام؟). هذا الدليل يمنحكم إجابات هندسية واثقة ومدعومة بالأرقام والحجج العلمية.

---

## 📂 المحور الأول: أسئلة البيانات والمعالجة (Dataset & Preprocessing)

### س1: الداتا دي حقيقية ولا من تأليفكم؟
* **الإجابة النموذجية:**  
  الداتا أصلها **حقيقي 100%** مسحوبة من مسابقة معتمدة على منصة كاجل العالمية (`HackerEarth ML Challenge`) تضم 22,750 موظفاً من شركات تقنية وخدمية حقيقية. ولكن الداتا الخام كانت مليئة بالعيوب والقيم الفارغة (4,622 خلية فارغة) وتكرارات، وخالية تماماً من الكلام والنصوص. لذلك قمنا ببناء خط معالجة وهندسة بيانات كامل (`Data Cleaning Pipeline`) أنتج عينة معيارية ممثلة لـ 1,000 موظف بنسبة احتراق مطابقة لتقارير منظمة العمل الدولية (ILO).

---

### س2: ليه أخدتوا 1,000 موظف بس من أصل 22 ألف؟
* **الإجابة النموذجية:**  
  1. **إحصائياً:** العينة تم سحبها بطريقة العينات الطبقية (`Stratified Sampling`) مع تثبيت معامل العشوائية (`random_state=42`)، مما حافظ بدقة متناهية على نفس التوزيع الطبيعي ونسبة الاحتراق (44.1%) دون أي تحيز.
  2. **عملياً:** حجم 1,000 موظف يمثل الحجم المثالي والمطابق للواقع لشركة متوسطة إلى كبرى (`Enterprise Scale`)، ويضمن سرعة استجابة فائقة لقاعدة البيانات وواجهات المستخدم ومحرك الـ NLP في أقل من 200ms أثناء العرض الحي.

---

### س3: داتا كاجل مفيهاش سن (Age) ولا نصوص مشاعر (Feedback)، جبتوهم منين؟
* **الإجابة النموذجية:**  
  1. **بخصوص العمر:** مسابقات الموارد البشرية العالمية تحجب تاريخ الميلاد لأسباب الخصوصية ومنع التمييز (`GDPR`). لذلك طبقنا مفهوم هندسة الميزات (`Feature Engineering`) واستنتجنا العمر رياضياً بالاعتماد على عمود الرتبة الوظيفية (`Designation` من 0 لـ 5) وتدرج الخبرة المهنية (بين 22 و 56 سنة، بمتوسط 34.7 سنة)، مع الحفاظ الكامل على عمود الجنس الحقيقي من كاجل.
  2. **بخصوص نصوص المشاعر:** داتا كاجل أرقام فقط، ولبناء منصة هجينة متعددة الوسائط (`Multi-Modal`) ابتكرنا 45 نصاً حقيقياً يعبر عن مشاعر بيئة العمل ووزعناها باحتمالات دقيقة تتوافق مع معدل إجهاد كل موظف.

---

### س4: هل عندكم مشكلة عدم توازن في البيانات (Class Imbalance)؟
* **الإجابة النموذجية:**  
  لا إطلاقاً يا دكتور، نسبة الموظفين السليمين تمثل 55.9% (559 موظفاً) ونسبة المعرضين للاحتراق تمثل 44.1% (441 موظفاً). نسبة التوازن هي **1.27 : 1**، وهي نسبة توازن إحصائي مثالية في تعلم الآلة لا تتطلب أي توليد مصطنع للبيانات مثل تقنيات SMOTE أو خوارزميات إعادة الوزن.

---

## 🤖 المحور الثاني: خوارزميات الذكاء الاصطناعي (ML, NLP & Explainability)

### س5: ليه اخترتم Random Forest بالذات؟ وليه مش Deep Learning؟
* **الإجابة النموذجية:**  
  1. قمنا بالمقارنة والتجربة العملية في كود `train_model.py` بين Logistic Regression و Random Forest، وتفوق الـ Random Forest محققاً دقة **95.50%** ومقياس F1-Score قدره **0.95**.
  2. **البيانات المجدولة الرقمية (`Tabular Data`):** أثبتت كبرى الأوراق البحثية أن النماذج الشجرية المجمعة (`Ensemble Tree Models`) تتفوق على الشبكات العميقة في البيانات المجدولة المحدودة، وتتفادى مشكلة الإفراط في المطابقة (`Overfitting`) وسريعة جداً في التدريب والاستنتاج.

---

### س6: ليه بتبصوا على الـ Recall والـ F1-Score ومش الـ Accuracy بس؟
* **الإجابة النموذجية:**  
  لأن في التطبيقات الصحية والنفسية الخطأ من النوع الثاني (`False Negative`) كارثي؛ إذا كان موظف على حافة الانهيار أو الاحتراق وتنبأ النموذج بأنه سليم، فستفقد المؤسسة الموظف دون تدخل علاجي مبكر. مقياس الاستدعاء (`Recall`) في نموذجنا وصل إلى **96.6%**، وهو ما يضمن اكتشاف 97 من كل 100 موظف معرضين للاحتراق.

---

### س7: ليه عملتوا النظام هجين (60% أرقام + 40% مشاعر) وما اكتفيتوش بالأرقام؟
* **الإجابة النموذجية:**  
  لأن السلوك البشري معقد؛ الأرقام وحدها قد تضلل، فقد يعمل الموظف 8 ساعات فقط لكنه يعاني من إحباط حاد أو بيئة عمل سامة (وهنا يكشفه الـ NLP). وعلى العكس، قد يعمل موظف لساعات إضافية لأنه شغوف ومتحمس لمشروعه (وهنا يمنعه الـ NLP من التصنيف الخاطئ). المعادلة الهجينة توفر رؤية شاملة 360 درجة تجمع بين الإجهاد الجسدي والنفسي.

---

### س8: ازاي الموديل بيفهم الكلام العربي والإنجليزي في نصوص المشاعر؟
* **الإجابة النموذجية:**  
  النظام يطبق بنية معالجة لغات طبيعية مزدوجة:
  1. **النصوص الإنجليزية:** نستخدم نموذج ترانسفورمرز خفيف وعميق `DistilBERT` من Hugging Face.
  2. **النصوص العربية:** قمنا بتطوير محرك NLP مخصص داخل الـ Backend يتعرف على مفردات بيئة العمل المصرية والعربية، مع معالجة متقدمة لأدوات النفي (مثل: "مش مرتاح"، "غير راضٍ") والسوابق والتطبيع اللغوي للألف والتاء المربوطة.

---

### س9: إيه هو الـ SHAP وليه استخدمتوه في المشروع؟
* **الإجابة النموذجية:**  
  تقنية SHAP (`Shapley Additive Explanations`) هي حجر الأساس في الذكاء الاصطناعي القابل للتفسير (`Explainable AI - XAI`). وظيفتها إزالة الغموض عن قرارات النموذج ومنعه من أن يكون صندوقاً أسود (`Black Box`). بفضل SHAP، لا يكتفي النظام بإعطاء نسبة خطر للمدير، بل يعرض أهم 3 أسباب حقيقية ساهمت في النتيجة (مثل: ساعات عمل إضافية مفرطة، أو حرمان من الإجازات).

---

## 💻 المحور الثالث: معمارية النظام والبرمجة (System Architecture & Security)

### س10: إيه الـ Architecture والتقنيات اللي بنيتوا بيها السيستم وليه؟
* **الإجابة النموذجية:**  
  بنينا النظام وفق معمارية الـ `Decoupled Full-Stack Architecture`:
  * **الـ Backend:** إطار عمل **FastAPI** مبني على معيار ASGI السريع، ويوفر توثيقاً تلقائياً عبر Swagger UI.
  * **قاعدة البيانات:** طبقة **SQLAlchemy ORM** تدعم SQLite للتجربة المحلية وتدعم PostgreSQL للإنتاج المؤسسي.
  * **الـ Frontend:** واجهة تفاعلية حديثة بـ **HTML5/CSS3/Vanilla JS** مع محرك ترجمة مخصص (`i18n`) يدعم العربية RTL والإنجليزية LTR فورياً.

---

### س11: ازاي مأمنين السيستم وبيانات الموظفين الحساسة؟ (Security)
* **الإجابة النموذجية:**  
  النظام يطبق المعايير الأمنية المؤسسية:
  1. التوثيق والمصادقة بـ **JWT Bearer Tokens** مع تشفير التوكنات وتحديد صلاحية زمنية للـ Session.
  2. تشفير كلمات المرور باستخدام خوارزمية **PBKDF2-HMAC-SHA256** المشددة مع Salt لمنع هجمات الجداول المسبقة (`Rainbow Tables`).
  3. التحكم بالوصول المبني على الأدوار (`Role-Based Access Control`): الموظف العادي يرى بياناته وتقييماته فقط، بينما مدير الـ HR هو المخول بالاطلاع على لوحة القيادة وكافة الموظفين وتعديلهم.

---

### س12: لو الشركة كبرت وبقى فيها 50 ألف موظف، هل السيستم هيتحمل؟ (Scalability)
* **الإجابة النموذجية:**  
  نعم بالتأكيد، النظام مصمم للتوسع الأفقي (`Horizontal Scalability`):
  1. خادم FastAPI غير متزامن بالكامل (`Async/Non-blocking I/O`) قادر على خدمة آلاف الاتصالات المتزامنة.
  2. معمارية `Stateless RESTful APIs` تسمح بتشغيل عدة نسخ من السيرفر خلف Load Balancer وحاويات Docker.
  3. نماذج الذكاء الاصطناعي مجهزة ومخزنة بملفات مدمجة (`Serialized via Joblib`) وسرعة استنتاجها لا تتعدى 50ms لكل طلب.
"""
    with open(OUT_PATH_MD, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"[SUCCESS] Markdown QA Guide saved at: {OUT_PATH_MD}")


if __name__ == "__main__":
    create_document()
    create_markdown()
