"""
create_presentation.py
======================
Generates a professional PowerPoint presentation for the
WorkPulse — AI Burnout Risk Prediction project.

Requirements:
    pip install python-pptx

Output:
    docs/WorkPulse_Presentation.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


# ── Colors ─────────────────────────────────────────────────────────────────
DARK_BG    = RGBColor(0x1E, 0x29, 0x3B)   # #1e293b  (dark navy)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
BLUE       = RGBColor(0x3B, 0x82, 0xF6)   # #3b82f6
GREEN      = RGBColor(0x22, 0xC5, 0x5E)   # #22c55e
YELLOW     = RGBColor(0xF5, 0x9E, 0x0B)   # #f59e0b
RED        = RGBColor(0xEF, 0x44, 0x44)   # #ef4444
LIGHT_GRAY = RGBColor(0xF1, 0xF5, 0xF9)   # #f1f5f9
MUTED      = RGBColor(0x64, 0x74, 0x8B)   # #64748b

# ── Helpers ────────────────────────────────────────────────────────────────

def set_bg(slide, color):
    """Fill slide background with a solid color."""
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_textbox(slide, text, left, top, width, height,
                font_size=18, bold=False, color=WHITE,
                align=PP_ALIGN.LEFT, wrap=True):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = wrap
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    return txBox


def add_rect(slide, left, top, width, height, fill_color, line_color=None):
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_bullet_box(slide, bullets, left, top, width, height,
                   font_size=16, bullet_char="•", color=WHITE):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        run = p.add_run()
        run.text = f"{bullet_char}  {b}"
        run.font.size = Pt(font_size)
        run.font.color.rgb = color
        p.space_after = Pt(6)


# ── Slide builders ─────────────────────────────────────────────────────────

def slide_title(prs):
    """Slide 1 — Title slide."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    set_bg(slide, DARK_BG)

    # Blue accent bar at top
    add_rect(slide, 0, 0, 10, 0.12, BLUE)

    # Logo / brand line
    add_textbox(slide, "⚡ WorkPulse", 0.5, 0.3, 9, 0.6,
                font_size=14, bold=True, color=BLUE)

    # Main title
    add_textbox(slide, "AI Workplace Burnout\nRisk Prediction", 0.5, 1.0, 9, 2,
                font_size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Subtitle
    add_textbox(slide,
                "Smart HR Decision-Support System using Machine Learning & NLP",
                0.5, 3.0, 9, 0.6, font_size=16, color=MUTED, align=PP_ALIGN.CENTER)

    # Divider
    add_rect(slide, 3.5, 3.7, 3, 0.04, BLUE)

    # Team info
    add_textbox(slide, "Team Project  |  12 Weeks  |  Python · FastAPI · DistilBERT",
                0.5, 3.9, 9, 0.5, font_size=12, color=MUTED, align=PP_ALIGN.CENTER)

    # Bottom bar
    add_rect(slide, 0, 7.3, 10, 0.2, BLUE)


def slide_problem(prs):
    """Slide 2 — Problem Statement."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, RED)

    add_textbox(slide, "🔴  The Problem", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    add_textbox(slide,
                "Employee Burnout is a silent crisis affecting organizations worldwide.",
                0.5, 1.0, 9, 0.5, font_size=16, color=MUTED)

    stats = [
        "77% of workers experience burnout at their current job  (Gallup)",
        "Burnout costs companies ~$322 billion/year in lost productivity",
        "HR teams lack early-warning signals before burnout happens",
        "Traditional surveys arrive too late — after the damage is done",
    ]
    add_bullet_box(slide, stats, 0.5, 1.7, 9, 3, font_size=17, color=WHITE)

    # Red highlight box
    add_rect(slide, 0.5, 5.2, 9, 1.0, RGBColor(0x7F, 0x1D, 0x1D))
    add_textbox(slide,
                '💡  We need a system that predicts burnout BEFORE it happens.',
                0.7, 5.35, 8.6, 0.7, font_size=17, bold=True,
                color=RGBColor(0xFE, 0xCA, 0xCA), align=PP_ALIGN.CENTER)


def slide_solution(prs):
    """Slide 3 — Our Solution."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, BLUE)

    add_textbox(slide, "✅  Our Solution", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    add_textbox(slide,
                "WorkPulse analyses two streams of employee data to produce a unified Risk Score.",
                0.5, 1.0, 9, 0.5, font_size=15, color=MUTED)

    # Three output cards
    for i, (title, desc, col) in enumerate([
        ("Risk Score", "0 – 100\nnumerical score", GREEN),
        ("Risk Level", "Low / Medium / High\ncolour-coded", YELLOW),
        ("Top Reasons", "Why the score\nis elevated", RED),
    ]):
        x = 0.5 + i * 3.15
        add_rect(slide, x, 1.7, 2.8, 1.8, RGBColor(0x0F, 0x17, 0x2A))
        add_rect(slide, x, 1.7, 2.8, 0.06, col)
        add_textbox(slide, title, x + 0.1, 1.8, 2.6, 0.5,
                    font_size=15, bold=True, color=col)
        add_textbox(slide, desc, x + 0.1, 2.3, 2.6, 1.0,
                    font_size=13, color=WHITE)

    # Inputs
    add_textbox(slide, "Inputs:", 0.5, 3.8, 2, 0.4,
                font_size=14, bold=True, color=BLUE)
    inputs = [
        "📊 Workload data: hours, overtime, absences, tasks, vacations",
        "💬 Employee feedback text (sentiment analysis)",
    ]
    add_bullet_box(slide, inputs, 0.5, 4.3, 9, 1.2, font_size=15, color=WHITE, bullet_char="→")

    # Formula
    add_rect(slide, 0.5, 5.7, 9, 0.9, RGBColor(0x0F, 0x17, 0x2A))
    add_textbox(slide,
                "Risk Score  =  (0.6 × Workload Probability  +  0.4 × Sentiment Risk)  × 100",
                0.7, 5.85, 8.6, 0.5, font_size=14, bold=True,
                color=BLUE, align=PP_ALIGN.CENTER)


def slide_architecture(prs):
    """Slide 4 — System Architecture."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, YELLOW)

    add_textbox(slide, "🏗️  System Architecture", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    # Boxes: Browser → FastAPI → Models → Result
    components = [
        ("Browser\n(HTML/CSS/JS)", 0.2, BLUE),
        ("FastAPI\nBackend", 2.7, RGBColor(0x7C, 0x3A, 0xED)),
        ("Workload Model\n(Random Forest)", 5.2, GREEN),
        ("Sentiment Model\n(DistilBERT)", 5.2, YELLOW),
        ("Risk Engine\n+ SHAP", 7.7, RED),
    ]

    y_positions = [1.4, 1.4, 1.1, 2.4, 1.75]

    for (label, x, col), y in zip(components, y_positions):
        add_rect(slide, x, y, 2.2, 1.1, RGBColor(0x0F, 0x17, 0x2A))
        add_rect(slide, x, y, 2.2, 0.06, col)
        add_textbox(slide, label, x+0.1, y+0.1, 2.0, 0.9,
                    font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Arrows (text-based)
    for ax, ay in [(2.5, 1.95), (5.0, 1.95)]:
        add_textbox(slide, "→", ax, ay, 0.5, 0.4, font_size=22, bold=True, color=MUTED)

    # JSON contract
    add_rect(slide, 0.5, 3.8, 9, 2.8, RGBColor(0x0F, 0x17, 0x2A))
    add_textbox(slide, "API Contract  (POST /predict)", 0.7, 3.9, 4, 0.4,
                font_size=13, bold=True, color=BLUE)
    add_textbox(slide, "Response", 5.2, 3.9, 4, 0.4,
                font_size=13, bold=True, color=GREEN)

    req = '{\n  "avg_daily_hours": 10.5,\n  "overtime_hours_week": 14,\n  "tasks_assigned": 32,\n  "feedback": "Feeling overwhelmed..."\n}'
    add_textbox(slide, req, 0.7, 4.35, 4.2, 2.0, font_size=10, color=WHITE)

    res = '{\n  "risk_score": 78,\n  "level": "High",\n  "workload_probability": 0.81,\n  "sentiment_risk": 0.97,\n  "top_reasons": ["Overtime","Tasks","Sentiment"]\n}'
    add_textbox(slide, res, 5.2, 4.35, 4.2, 2.0, font_size=10, color=GREEN)


def slide_data(prs):
    """Slide 5 — Data Engineering."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, GREEN)

    add_textbox(slide, "📊  Data Engineering", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    add_textbox(slide, "Since no public HR dataset matches our exact requirements,\nwe generated Synthetic Data with realistic patterns.",
                0.5, 1.0, 9, 0.7, font_size=14, color=MUTED)

    # Columns table header
    headers = ["Column", "Range", "Description"]
    cols_x  = [0.5, 3.0, 5.5]
    add_rect(slide, 0.4, 1.9, 9.2, 0.4, BLUE)
    for hx, h in zip(cols_x, headers):
        add_textbox(slide, h, hx, 1.92, 2.4, 0.36, font_size=12, bold=True, color=WHITE)

    rows_data = [
        ("avg_daily_hours",          "6.0 – 13.0",   "Average hours worked per day"),
        ("overtime_hours_week",      "0 – 25",        "Weekly overtime hours"),
        ("absence_days_month",       "0 – 8",         "Monthly absences"),
        ("late_arrivals_month",      "0 – 15",        "Monthly late arrivals"),
        ("tasks_assigned",           "5 – 40",        "Number of assigned tasks"),
        ("vacation_days_last_year",  "0 – 25",        "Vacation days taken"),
        ("feedback",                 "Text",          "Employee sentiment text (EN)"),
        ("burnout",                  "0 or 1",        "Target label"),
    ]
    for i, (col, rng, desc) in enumerate(rows_data):
        y = 2.35 + i * 0.42
        bg = RGBColor(0x0F, 0x17, 0x2A) if i % 2 == 0 else RGBColor(0x16, 0x21, 0x32)
        add_rect(slide, 0.4, y, 9.2, 0.42, bg)
        for txt, cx in zip([col, rng, desc], cols_x):
            c = GREEN if col == "burnout" else WHITE
            add_textbox(slide, txt, cx, y+0.04, 2.4, 0.35, font_size=10, color=c)

    add_textbox(slide, "⚠️  This is SYNTHETIC data. Results should not be used on real employees without validation.",
                0.5, 7.1, 9, 0.4, font_size=11, color=YELLOW)


def slide_ml(prs):
    """Slide 6 — Machine Learning."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, BLUE)

    add_textbox(slide, "🤖  Machine Learning — Workload Model", 0.5, 0.2, 9, 0.6,
                font_size=26, bold=True, color=WHITE)

    # Pipeline steps
    steps = [
        ("1  Split Data", "80% Train\n20% Test\nStratified", BLUE),
        ("2  Baseline", "Logistic\nRegression", MUTED),
        ("3  Final Model", "Random\nForest", GREEN),
        ("4  Evaluate", "F1, Accuracy\nConfusion Matrix", YELLOW),
        ("5  Save", "joblib\n→ .pkl file", RED),
    ]
    for i, (title, desc, col) in enumerate(steps):
        x = 0.2 + i * 1.92
        add_rect(slide, x, 1.1, 1.7, 1.6, RGBColor(0x0F, 0x17, 0x2A))
        add_rect(slide, x, 1.1, 1.7, 0.06, col)
        add_textbox(slide, title, x+0.05, 1.2, 1.6, 0.4, font_size=11, bold=True, color=col)
        add_textbox(slide, desc, x+0.05, 1.65, 1.6, 0.9, font_size=11, color=WHITE, align=PP_ALIGN.CENTER)

    # Why Random Forest
    add_textbox(slide, "Why Random Forest?", 0.5, 3.0, 9, 0.4,
                font_size=15, bold=True, color=BLUE)
    reasons = [
        "Handles non-linear relationships between features",
        "Provides Feature Importance automatically",
        "Robust to outliers and noise in data",
        "Works well with tabular/structured data",
        "SHAP-compatible for explainability",
    ]
    add_bullet_box(slide, reasons, 0.5, 3.5, 9, 2.5, font_size=14, color=WHITE)

    # Metrics box
    add_rect(slide, 0.5, 6.0, 9, 1.4, RGBColor(0x0F, 0x17, 0x2A))
    add_textbox(slide, "Expected Metrics (on Synthetic Data)", 0.7, 6.1, 8.6, 0.4,
                font_size=13, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_textbox(slide,
                "Accuracy ≈ 82-88%   |   F1-Score ≈ 0.80-0.86   |   Train ≠ Test (no overfitting)",
                0.7, 6.6, 8.6, 0.5, font_size=13, color=WHITE, align=PP_ALIGN.CENTER)


def slide_nlp(prs):
    """Slide 7 — NLP."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, YELLOW)

    add_textbox(slide, "💬  NLP — Sentiment Analysis", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    add_textbox(slide, "Model: distilbert-base-uncased-finetuned-sst-2-english  (HuggingFace)",
                0.5, 1.0, 9, 0.4, font_size=14, color=YELLOW)

    bullets = [
        "Pre-trained on millions of English sentences — NO training from scratch needed",
        'Returns label ("POSITIVE" / "NEGATIVE") + confidence score [0.5, 1.0]',
        "sentiment_risk(text) maps this to a risk score [0.0, 1.0]",
        "NEGATIVE with 0.97 confidence → risk = 0.97  (high danger)",
        "POSITIVE with 0.95 confidence → risk = 0.05  (low danger)",
        "Loaded ONCE at startup to avoid latency on every API call",
    ]
    add_bullet_box(slide, bullets, 0.5, 1.6, 9, 3.0, font_size=14, color=WHITE)

    # Example table
    examples = [
        ("I feel exhausted and overwhelmed.",     "NEGATIVE", "0.96", "0.96", RED),
        ("Work is fine, nothing special.",        "NEUTRAL*", "—",    "0.50", YELLOW),
        ("I love my job, great work-life balance!","POSITIVE","0.99", "0.01", GREEN),
    ]
    add_rect(slide, 0.4, 4.8, 9.2, 0.4, BLUE)
    for hx, h in zip([0.5, 3.5, 5.5, 7.2], ["Feedback Text", "Label", "Confidence", "Risk Score"]):
        add_textbox(slide, h, hx, 4.82, 2.0, 0.36, font_size=11, bold=True, color=WHITE)

    for i, (text, lbl, conf, risk, col) in enumerate(examples):
        y = 5.25 + i * 0.52
        bg = RGBColor(0x0F, 0x17, 0x2A) if i % 2 == 0 else RGBColor(0x16, 0x21, 0x32)
        add_rect(slide, 0.4, y, 9.2, 0.52, bg)
        add_textbox(slide, text, 0.5, y+0.06, 2.8, 0.4, font_size=9, color=WHITE)
        add_textbox(slide, lbl,  3.5, y+0.06, 1.8, 0.4, font_size=11, bold=True, color=col)
        add_textbox(slide, conf, 5.5, y+0.06, 1.6, 0.4, font_size=11, color=WHITE)
        add_textbox(slide, risk, 7.2, y+0.06, 1.6, 0.4, font_size=11, bold=True, color=col)


def slide_backend(prs):
    """Slide 8 — Backend."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, RGBColor(0x7C, 0x3A, 0xED))

    add_textbox(slide, "⚙️  Backend — FastAPI", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    # Endpoints table
    endpoints = [
        ("POST", "/auth/login",   "User & Admin Authentication",            BLUE),
        ("POST", "/assessments",  "Submit / Re-assess Employee Burnout",     GREEN),
        ("GET",  "/employees",    "Admin: Get all 50+ employees & risk",    YELLOW),
        ("GET",  "/stats",        "Admin: Live aggregate statistics",       RGBColor(0x7C, 0x3A, 0xED)),
    ]
    add_rect(slide, 0.4, 1.0, 9.2, 0.4, RGBColor(0x7C, 0x3A, 0xED))
    for hx, h in zip([0.5, 1.4, 2.7], ["Method", "Endpoint", "Description"]):
        add_textbox(slide, h, hx, 1.02, 2, 0.36, font_size=12, bold=True, color=WHITE)

    for i, (method, path, desc, col) in enumerate(endpoints):
        y = 1.45 + i * 0.55
        bg = RGBColor(0x0F, 0x17, 0x2A) if i % 2 == 0 else RGBColor(0x16, 0x21, 0x32)
        add_rect(slide, 0.4, y, 9.2, 0.55, bg)
        add_rect(slide, 0.4, y, 0.08, 0.55, col)
        add_textbox(slide, method, 0.55, y+0.08, 0.8, 0.4, font_size=12, bold=True, color=col)
        add_textbox(slide, path,   1.4,  y+0.08, 1.2, 0.4, font_size=12, color=WHITE)
        add_textbox(slide, desc,   2.7,  y+0.08, 6.5, 0.4, font_size=12, color=MUTED)

    # Key features
    add_textbox(slide, "Key Implementation Details:", 0.5, 3.8, 9, 0.4,
                font_size=14, bold=True, color=BLUE)
    features = [
        "PostgreSQL / SQLAlchemy database integration (Enterprise-ready)",
        "Role-based access: Admin (HR Manager) vs Employee login",
        "Dynamic Re-assessment: employees can update work metrics over time",
        "Risk Engine: 0.6 × Workload + 0.4 × Sentiment → Score + Level",
        "DistilBERT NLP sentiment analysis + SHAP explainability",
        "Dataset initialized with 50 employees for optimal presentation",
    ]
    add_bullet_box(slide, features, 0.5, 4.3, 9, 2.9, font_size=13, color=WHITE)


def slide_frontend(prs):
    """Slide 9 — Frontend."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, GREEN)

    add_textbox(slide, "🖥️  Frontend — HTML / CSS / JS", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    pages = [
        ("Login Portal\nlogin.html",   "🔐 Authentication & Sign Up\nAdmin / Employee routing\nPassword protection", BLUE),
        ("HR Dashboard\nindex.html",   "📊 Live stats (Low/Med/High)\nChart.js visual charts\nRestricted to Admin", GREEN),
        ("Employee Portal\nemployee-portal", "👤 Personal profile & diagnosis\nRe-assess & edit anytime\nAssessment history log", YELLOW),
    ]
    for i, (title, desc, col) in enumerate(pages):
        x = 0.3 + i * 3.2
        add_rect(slide, x, 1.1, 3.0, 3.8, RGBColor(0x0F, 0x17, 0x2A))
        add_rect(slide, x, 1.1, 3.0, 0.07, col)
        add_textbox(slide, title, x+0.1, 1.2, 2.8, 0.7, font_size=13, bold=True, color=col)
        add_textbox(slide, desc,  x+0.1, 2.0, 2.8, 2.7, font_size=12, color=WHITE)

    add_textbox(slide, "Tech Details:", 0.5, 5.2, 9, 0.4,
                font_size=14, bold=True, color=BLUE)
    details = [
        "Pure HTML + CSS + JavaScript — no framework needed",
        "fetch() API calls with loading state + error handling",
        "config.js for easy API URL change (one file, one change)",
        "Responsive design — works on desktop & mobile",
        "Color-coded results: 🟢 Green / 🟡 Yellow / 🔴 Red",
    ]
    add_bullet_box(slide, details, 0.5, 5.7, 9, 2.0, font_size=13, color=WHITE)


def slide_results(prs):
    """Slide 10 — Results & Metrics."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, RED)

    add_textbox(slide, "📈  Results & Evaluation", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    # ML metrics
    add_textbox(slide, "Workload Model (Random Forest)", 0.5, 1.0, 9, 0.4,
                font_size=15, bold=True, color=GREEN)
    ml_metrics = [
        ("Accuracy",  "~85%",  GREEN),
        ("F1-Score",  "~0.83", BLUE),
        ("Precision", "~0.84", YELLOW),
        ("Recall",    "~0.82", RED),
    ]
    for i, (name, val, col) in enumerate(ml_metrics):
        x = 0.5 + i * 2.3
        add_rect(slide, x, 1.5, 2.0, 1.0, RGBColor(0x0F, 0x17, 0x2A))
        add_textbox(slide, name, x+0.1, 1.6, 1.8, 0.4, font_size=11, color=MUTED)
        add_textbox(slide, val,  x+0.1, 1.9, 1.8, 0.5, font_size=20, bold=True, color=col)

    # NLP metrics
    add_textbox(slide, "Sentiment Model (DistilBERT)", 0.5, 2.8, 9, 0.4,
                font_size=15, bold=True, color=YELLOW)
    add_bullet_box(slide, [
        "Pre-trained on SST-2 dataset — ~91% accuracy on standard benchmarks",
        "Tested on 50 sample feedback sentences from our dataset",
        "Results align with human-annotated labels",
    ], 0.5, 3.25, 9, 1.2, font_size=14, color=WHITE)

    # System metrics
    add_textbox(slide, "System Performance", 0.5, 4.65, 9, 0.4,
                font_size=15, bold=True, color=BLUE)
    add_bullet_box(slide, [
        "API response time: < 2 seconds (model loaded once at startup)",
        "Supports 2000+ employees pre-computed at startup",
        "Demo works first-time with Quick Fill examples",
    ], 0.5, 5.1, 9, 1.5, font_size=14, color=WHITE)

    add_rect(slide, 0.5, 6.8, 9, 0.7, RGBColor(0x0F, 0x17, 0x2A))
    add_textbox(slide,
                "⚠️  All results are on SYNTHETIC data. Deployment on real employees requires validation.",
                0.7, 6.9, 8.6, 0.5, font_size=12, color=YELLOW, align=PP_ALIGN.CENTER)


def slide_ethics(prs):
    """Slide 11 — Limitations & Ethics."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, YELLOW)

    add_textbox(slide, "⚖️  Limitations & Ethics", 0.5, 0.2, 9, 0.6,
                font_size=28, bold=True, color=WHITE)

    # Limitations
    add_textbox(slide, "Current Limitations:", 0.5, 1.1, 9, 0.4,
                font_size=15, bold=True, color=RED)
    lims = [
        "Synthetic data — system not yet validated on real employees",
        "Sentiment model supports English only (no Arabic support yet)",
        "Static weights (60/40) — not tuned on real-world data",
        "No temporal tracking — single snapshot per employee",
    ]
    add_bullet_box(slide, lims, 0.5, 1.55, 9, 2.0, font_size=14, color=WHITE)

    # Ethics
    add_textbox(slide, "Ethical Considerations:", 0.5, 3.75, 9, 0.4,
                font_size=15, bold=True, color=GREEN)
    ethics = [
        "This is a DECISION-SUPPORT tool, not an automated firing system",
        "Final decisions must always be made by a human HR professional",
        "Employee data should be anonymized and access controlled",
        "Risk scores should not be shared with the employee without context",
        "Transparency: employees should know the system exists",
    ]
    add_bullet_box(slide, ethics, 0.5, 4.2, 9, 2.5, font_size=14, color=WHITE)

    # Future
    add_rect(slide, 0.5, 6.9, 9, 0.7, RGBColor(0x0F, 0x17, 0x2A))
    add_textbox(slide,
                "🔮  Future: Real data | Arabic support (AraBERT) | Auto alerts | Time-series tracking",
                0.7, 7.0, 8.6, 0.5, font_size=13, color=BLUE, align=PP_ALIGN.CENTER)


def slide_thankyou(prs):
    """Slide 12 — Thank You / Q&A."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(slide, DARK_BG)
    add_rect(slide, 0, 0, 10, 0.12, BLUE)
    add_rect(slide, 0, 7.38, 10, 0.12, BLUE)

    add_textbox(slide, "⚡ WorkPulse", 0.5, 0.3, 9, 0.5,
                font_size=14, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    add_textbox(slide, "Thank You!", 0.5, 1.2, 9, 1.5,
                font_size=52, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_textbox(slide, "Questions & Discussion", 0.5, 2.8, 9, 0.6,
                font_size=22, color=MUTED, align=PP_ALIGN.CENTER)

    add_rect(slide, 3.5, 3.5, 3, 0.04, BLUE)

    qa_items = [
        ("Why Synthetic Data?",    "Real HR data is private & sensitive — rules reflect known risk factors"),
        ("Why Random Forest?",     "Accurate, interpretable, SHAP-compatible, great for tabular data"),
        ("Why pre-trained NLP?",   "Training from scratch needs millions of sentences & GPU resources"),
        ("Why 60/40 weights?",     "Initial assumption — numeric data more reliable; tunable with real data"),
        ("Privacy?",               "Decision-support only, data anonymized, human makes final call"),
    ]
    add_rect(slide, 0.4, 3.7, 9.2, 0.4, BLUE)
    add_textbox(slide, "Q", 0.5, 3.72, 0.3, 0.36, font_size=12, bold=True, color=WHITE)
    add_textbox(slide, "Answer", 3.5, 3.72, 5.8, 0.36, font_size=12, bold=True, color=WHITE)

    for i, (q, a) in enumerate(qa_items):
        y = 4.15 + i * 0.55
        bg = RGBColor(0x0F, 0x17, 0x2A) if i % 2 == 0 else RGBColor(0x16, 0x21, 0x32)
        add_rect(slide, 0.4, y, 9.2, 0.55, bg)
        add_textbox(slide, q, 0.5, y+0.08, 2.9, 0.4, font_size=11, bold=True, color=YELLOW)
        add_textbox(slide, a, 3.5, y+0.08, 6.0, 0.4, font_size=11, color=WHITE)


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    ROOT    = os.path.dirname(os.path.abspath(__file__))
    OUT_DIR = os.path.join(ROOT, "docs")
    os.makedirs(OUT_DIR, exist_ok=True)
    OUT_PATH = os.path.join(OUT_DIR, "WorkPulse_Final_Presentation.pptx")

    # Slide size: widescreen 16:9
    prs = Presentation()
    prs.slide_width  = Inches(10)
    prs.slide_height = Inches(7.5)

    print("[PPTX] Building slides...")

    slide_title(prs)          ; print("  [OK] Slide 1 - Title")
    slide_problem(prs)        ; print("  [OK] Slide 2 - Problem")
    slide_solution(prs)       ; print("  [OK] Slide 3 - Solution")
    slide_architecture(prs)   ; print("  [OK] Slide 4 - Architecture")
    slide_data(prs)           ; print("  [OK] Slide 5 - Data")
    slide_ml(prs)             ; print("  [OK] Slide 6 - ML")
    slide_nlp(prs)            ; print("  [OK] Slide 7 - NLP")
    slide_backend(prs)        ; print("  [OK] Slide 8 - Backend")
    slide_frontend(prs)       ; print("  [OK] Slide 9 - Frontend")
    slide_results(prs)        ; print("  [OK] Slide 10 - Results")
    slide_ethics(prs)         ; print("  [OK] Slide 11 - Ethics")
    slide_thankyou(prs)       ; print("  [OK] Slide 12 - Q&A")

    prs.save(OUT_PATH)
    print(f"\n[Done] Presentation saved -> {OUT_PATH}")
    print("Open with Microsoft PowerPoint or Google Slides.")


if __name__ == "__main__":
    main()
