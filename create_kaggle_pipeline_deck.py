"""
create_kaggle_pipeline_deck.py
===============================
PowerPoint Presentation: Kaggle Data Pipeline for WorkPulse
Explains the full journey from raw Kaggle data to clean 1000-employee dataset.

Generates: WorkPulse_Kaggle_Data_Pipeline.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.oxml import parse_xml

# ── Paths ──────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(ROOT, "WorkPulse_Kaggle_Data_Pipeline.pptx")

# ── Color Palette ──────────────────────────────────────────────────────────
BG_DARK    = RGBColor(0x0F, 0x17, 0x2A)   # Dark navy background
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
MUTED      = RGBColor(0x94, 0xA3, 0xB8)   # Slate 400
GREEN      = RGBColor(0x10, 0xB9, 0x81)   # Emerald 500
RED        = RGBColor(0xEF, 0x44, 0x44)   # Red 500
BLUE       = RGBColor(0x38, 0x82, 0xF6)   # Blue 500
YELLOW     = RGBColor(0xF5, 0x9E, 0x0B)   # Amber 500
CYAN       = RGBColor(0x06, 0xB6, 0xD4)   # Cyan 500
PURPLE     = RGBColor(0xA7, 0x8B, 0xFA)   # Purple 400
ORANGE     = RGBColor(0xFB, 0x92, 0x3C)   # Orange 400

HEX_BG     = "0F172A"
HEX_CARD   = "1E293B"
HEX_GREEN  = "10B981"
HEX_BLUE   = "3882F6"
HEX_RED    = "EF4444"
HEX_YELLOW = "F59E0B"
HEX_CYAN   = "06B6D4"
HEX_PURPLE = "A78BFA"

FONT = "Segoe UI"
FONT_AR = "Cairo"


def set_bg(slide, hex_color=HEX_BG):
    """Set slide background color."""
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor.from_string(hex_color)


def add_shape(slide, left, top, width, height, hex_fill, border_hex=None):
    """Add a rounded rectangle shape."""
    shape = slide.shapes.add_shape(
        1,  # MSO_SHAPE.ROUNDED_RECTANGLE
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(hex_fill)
    if border_hex:
        shape.line.color.rgb = RGBColor.from_string(border_hex)
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    return shape


def add_rect(slide, left, top, width, height, color):
    """Add a thin accent rectangle."""
    shape = slide.shapes.add_shape(
        1, Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    return shape


def add_text(slide, text, left, top, width, height, font_size=14,
             bold=False, color=WHITE, align=PP_ALIGN.LEFT, font_name=FONT):
    """Add a text box."""
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = align
    return txBox


def add_bullet_list(slide, items, left, top, width, height,
                    font_size=11, color=WHITE, spacing=8):
    """Add a bulleted list text box."""
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
        p.font.name = FONT
        p.space_after = Pt(spacing)
        # Add bullet
        pPr = p._pPr
        if pPr is None:
            pPr = parse_xml(f'<a:pPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" marL="228600" indent="-228600"/>')
            p._p.insert(0, pPr)
        else:
            pPr.set("marL", "228600")
            pPr.set("indent", "-228600")
    return txBox


def add_card(slide, left, top, width, height, accent_hex=HEX_GREEN):
    """Add a card with accent top border."""
    # Card background
    add_shape(slide, left, top, width, height, HEX_CARD)
    # Accent top bar
    add_rect(slide, left, top, width, 0.06, RGBColor.from_string(accent_hex))


def add_table_slide(slide, headers, rows, left, top, width, row_height=0.4):
    """Add a styled table to the slide."""
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(left), Inches(top),
                                          Inches(width), Inches(row_height * n_rows))
    table = table_shape.table

    # Style header
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = h
        for paragraph in cell.text_frame.paragraphs:
            paragraph.font.size = Pt(10)
            paragraph.font.bold = True
            paragraph.font.color.rgb = WHITE
            paragraph.font.name = FONT
            paragraph.alignment = PP_ALIGN.CENTER
        # Header background
        tcPr = cell._tc.get_or_add_tcPr()
        shd = parse_xml(f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{HEX_GREEN}"/></a:solidFill>')
        tcPr.append(shd)

    # Style rows
    for i, row in enumerate(rows):
        bg = HEX_CARD if i % 2 == 0 else HEX_BG
        for j, val in enumerate(row):
            cell = table.cell(i + 1, j)
            cell.text = str(val)
            for paragraph in cell.text_frame.paragraphs:
                paragraph.font.size = Pt(9)
                paragraph.font.color.rgb = WHITE
                paragraph.font.name = FONT
                paragraph.alignment = PP_ALIGN.CENTER
            tcPr = cell._tc.get_or_add_tcPr()
            shd = parse_xml(f'<a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="{bg}"/></a:solidFill>')
            tcPr.append(shd)

    return table_shape


def build_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 1: Title Slide
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.15, GREEN)

    add_text(s, "WorkPulse", 1, 1.5, 11, 0.8,
             font_size=48, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    add_text(s, "Kaggle Data Pipeline", 1, 2.3, 11, 0.7,
             font_size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_text(s, "From Raw Kaggle Dataset to 1,000 Clean Employee Records", 1, 3.2, 11, 0.5,
             font_size=18, color=MUTED, align=PP_ALIGN.CENTER)

    # Pipeline flow
    add_shape(s, 1.5, 4.5, 2, 0.8, HEX_RED)
    add_text(s, "22,750 Raw\nKaggle Records", 1.5, 4.5, 2, 0.8,
             font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(s, ">>>", 3.7, 4.6, 0.8, 0.5,
             font_size=24, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    add_shape(s, 4.5, 4.5, 2, 0.8, HEX_YELLOW)
    add_text(s, "Data Cleaning\n& Null Handling", 4.5, 4.5, 2, 0.8,
             font_size=12, bold=True, color=RGBColor(0x0F, 0x17, 0x2A), align=PP_ALIGN.CENTER)

    add_text(s, ">>>", 6.7, 4.6, 0.8, 0.5,
             font_size=24, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    add_shape(s, 7.5, 4.5, 2, 0.8, HEX_BLUE)
    add_text(s, "Feature Mapping\n& Transformation", 7.5, 4.5, 2, 0.8,
             font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(s, ">>>", 9.7, 4.6, 0.8, 0.5,
             font_size=24, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

    add_shape(s, 10.5, 4.5, 2, 0.8, HEX_GREEN)
    add_text(s, "1,000 Clean\nEmployees", 10.5, 4.5, 2, 0.8,
             font_size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    add_text(s, "Source: Kaggle HackerEarth  |  Are Your Employees Burning Out?  |  kaggle.com/datasets/blurredmachine/are-your-employees-burning-out",
             0.5, 6.5, 12, 0.4, font_size=10, color=MUTED, align=PP_ALIGN.CENTER)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 2: Data Source & Why Kaggle
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, BLUE)

    add_text(s, "1. Data Source: Kaggle HackerEarth Burnout Dataset",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=BLUE)
    add_text(s, "kaggle.com/datasets/blurredmachine/are-your-employees-burning-out",
             0.5, 0.85, 12, 0.3, font_size=12, color=MUTED)

    # Left card - About the dataset
    add_card(s, 0.5, 1.3, 5.8, 5.5, HEX_BLUE)
    add_text(s, "About the Original Dataset", 0.8, 1.5, 5.2, 0.35,
             font_size=15, bold=True, color=BLUE)
    add_bullet_list(s, [
        "Competition: HackerEarth ML Challenge hosted on Kaggle",
        "Total Records: 22,750 employee entries from real organizations",
        "Target Variable: Burn Rate (continuous 0.0 to 1.0)",
        "9 Original Columns: Employee ID, Date of Joining, Gender,\n  Company Type, WFH Setup, Designation, Resource Allocation,\n  Mental Fatigue Score, Burn Rate",
        "Data Type: Tabular numerical data (no text/NLP feedback)",
        "This is a well-known benchmark dataset cited by 500+ notebooks",
    ], 0.8, 2.0, 5.2, 4.5, font_size=12, color=WHITE)

    # Right card - Why this dataset
    add_card(s, 6.8, 1.3, 5.8, 5.5, HEX_GREEN)
    add_text(s, "Why We Chose This Dataset", 7.1, 1.5, 5.2, 0.35,
             font_size=15, bold=True, color=GREEN)
    add_bullet_list(s, [
        "Real-world data from actual companies (Service & Product sectors)",
        "Contains Mental Fatigue Score: a validated psychological metric\n  aligned with Maslach Burnout Inventory (MBI) standards",
        "22,750 records = statistically significant for ML training",
        "Has realistic data quality issues (nulls, duplicates)\n  that demonstrate our Data Cleaning skills",
        "Burn Rate target is continuous: we engineer a binary\n  classification threshold (>= 0.50 = Burnout) for our system",
        "Complements our Multi-Modal approach: we add the NLP\n  feedback text layer that Kaggle data lacks",
    ], 7.1, 2.0, 5.2, 4.5, font_size=12, color=WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 3: Raw Data Problems (Before Cleaning)
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, RED)

    add_text(s, "2. Raw Data Quality Issues (Before Cleaning)",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=RED)
    add_text(s, "The Kaggle dataset has significant data quality issues that must be resolved before ML training",
             0.5, 0.85, 12, 0.3, font_size=12, color=MUTED)

    # Problem cards
    problems = [
        ("Missing Values (NaN)", HEX_RED, [
            "Burn Rate: 1,124 nulls (4.9%)",
            "Mental Fatigue Score: 2,117 nulls (9.3%)",
            "Resource Allocation: 1,381 nulls (6.1%)",
            "Total: 4,622 problematic entries",
            "Impact: Cannot train supervised ML\n  model without labels (Ground Truth)",
        ]),
        ("Duplicate Records", HEX_YELLOW, [
            "5,948 duplicate rows detected",
            "Same feature combinations repeated",
            "Duplicates bias the model training",
            "Must be removed for fair evaluation",
            "After dedup: unique patterns preserved",
        ]),
        ("Schema Mismatch", HEX_PURPLE, [
            "Kaggle: Burn Rate is continuous (0.0 - 1.0)",
            "WorkPulse needs: Binary (0 or 1)",
            "Kaggle: No employee names (hashed IDs)",
            "Kaggle: No NLP feedback text column",
            "Requires Feature Engineering & Mapping",
        ]),
        ("No Text Data", HEX_CYAN, [
            "Kaggle has ONLY numerical features",
            "No employee feedback or complaints",
            "WorkPulse needs Multi-Modal input:",
            "  60% Quantitative (numbers)",
            "  40% Qualitative (NLP text sentiment)",
            "Solution: Smart NLP Fusion Pipeline",
        ]),
    ]

    for i, (title, hex_c, bullets) in enumerate(problems):
        col = i % 4
        x = 0.5 + col * 3.1
        add_card(s, x, 1.4, 2.9, 5.2, hex_c)
        add_text(s, title, x + 0.15, 1.55, 2.6, 0.35,
                 font_size=13, bold=True, color=RGBColor.from_string(hex_c))
        add_bullet_list(s, bullets, x + 0.15, 2.1, 2.6, 4.3,
                        font_size=10, color=WHITE, spacing=6)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 4: Cleaning Pipeline Steps
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, GREEN)

    add_text(s, "3. Data Cleaning Pipeline (7 Steps)",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=GREEN)

    steps = [
        ("Step 1", "Load Raw Data", "22,750 rows x 9 cols", HEX_RED),
        ("Step 2", "Drop Null Targets", "Removed 1,124 rows\nwith empty Burn Rate", HEX_RED),
        ("Step 3", "Drop Null Features", "Removed 3,036 rows\nwith empty Fatigue/Resource", HEX_YELLOW),
        ("Step 4", "Remove Duplicates", "Removed 5,948\nduplicate records", HEX_YELLOW),
        ("Step 5", "Stratified Sample", "Selected 1,000 from\n12,642 clean rows", HEX_BLUE),
        ("Step 6", "Feature Mapping", "Kaggle cols mapped to\nWorkPulse 10-col schema", HEX_CYAN),
        ("Step 7", "Export Clean CSV", "1,000 employees\n0 missing values", HEX_GREEN),
    ]

    for i, (step, title, desc, hex_c) in enumerate(steps):
        x = 0.3 + i * 1.82
        add_shape(s, x, 1.2, 1.65, 1.0, hex_c)
        add_text(s, step, x, 1.2, 1.65, 0.35,
                 font_size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_text(s, title, x, 1.5, 1.65, 0.35,
                 font_size=10, bold=True, color=RGBColor(0x0F, 0x17, 0x2A), align=PP_ALIGN.CENTER)

        # Arrow
        if i < len(steps) - 1:
            add_text(s, ">", x + 1.65, 1.35, 0.2, 0.5,
                     font_size=18, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)

        # Description card below
        add_card(s, x, 2.5, 1.65, 1.2, hex_c)
        add_text(s, desc, x + 0.1, 2.65, 1.45, 1.0,
                 font_size=9, color=WHITE, align=PP_ALIGN.CENTER)

    # Numbers summary
    add_card(s, 0.5, 4.2, 12, 2.8, HEX_GREEN)
    add_text(s, "Cleaning Results Summary", 0.8, 4.35, 11, 0.35,
             font_size=16, bold=True, color=GREEN)

    add_table_slide(s,
        ["Metric", "Before Cleaning", "After Cleaning", "Change"],
        [
            ["Total Records", "22,750", "12,642", "-10,108 removed"],
            ["Missing Values", "4,622 NaN cells", "0", "100% resolved"],
            ["Duplicate Rows", "5,948", "0", "100% removed"],
            ["Final Sample", "-", "1,000", "Stratified sampling"],
            ["Burnout Rate", "~43% (continuous)", "44.1% (binary)", "Threshold >= 0.50"],
            ["Columns", "9 (Kaggle schema)", "10 (WorkPulse schema)", "+1 (feedback text)"],
        ],
        0.8, 4.85, 11.4, 0.32
    )

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 5: Feature Mapping (Kaggle → WorkPulse)
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, CYAN)

    add_text(s, "4. Feature Engineering & Mapping (Kaggle >>> WorkPulse)",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=CYAN)
    add_text(s, "How we converted Kaggle's 9 columns into WorkPulse's 10-column schema with realistic value ranges",
             0.5, 0.85, 12, 0.3, font_size=12, color=MUTED)

    add_table_slide(s,
        ["WorkPulse Column", "Source (Kaggle)", "Transformation Formula", "Output Range"],
        [
            ["employee_id", "Employee ID (hash)", "Sequential: EMP0001 - EMP1000", "EMP0001 - EMP1000"],
            ["name", "- (not in Kaggle)", "Generated from Arabic name pool (Gender-matched)", "Unique full names"],
            ["avg_daily_hours", "Resource Allocation (1-10)", "6.0 + (RA - 1) / 9 * 7.0 + Noise(0, 0.3)", "6.0 - 13.0 hrs"],
            ["overtime_hours_week", "Resource + Fatigue", "(RA/10 * 15) + (MFS/10 * 10) + Noise(0, 2)", "0.0 - 25.0 hrs"],
            ["absence_days_month", "Mental Fatigue Score", "MFS / 10 * 8 + Noise(0, 1.5)", "0 - 8 days"],
            ["late_arrivals_month", "Burn Rate", "BurnRate * 15 + Noise(0, 2)", "0 - 15 times"],
            ["tasks_assigned", "Designation + RA", "5 + (D/5 * 20) + (RA/10 * 15) + Noise(0, 3)", "5 - 40 tasks"],
            ["vacation_days_last_year", "Burn Rate (inverse)", "25 * (1 - BurnRate) + Noise(0, 3)", "0 - 25 days"],
            ["feedback", "- (not in Kaggle)", "NLP text matched to Burn Rate probability", "45 unique sentences"],
            ["burnout", "Burn Rate (0.0-1.0)", "Binary: 1 if BurnRate >= 0.50, else 0", "0 or 1"],
        ],
        0.5, 1.3, 12.3, 0.38
    )

    # Key insight box
    add_card(s, 0.5, 5.8, 12.3, 1.3, HEX_CYAN)
    add_text(s, "Key Engineering Decisions", 0.8, 5.95, 11, 0.3,
             font_size=14, bold=True, color=CYAN)
    add_bullet_list(s, [
        "Gaussian Noise injected into every derived feature to prevent perfect linear correlation and ensure the ML model learns real patterns",
        "Vacation days have INVERSE correlation with Burn Rate: more burnout = less vacation taken (protective factor with negative weight)",
        "Feedback text is probabilistically assigned: 75% of high-burnout employees get negative text, ensuring NLP model alignment with ground truth",
    ], 0.8, 6.3, 11.5, 0.8, font_size=11, color=WHITE, spacing=4)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 6: NLP Text Fusion
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, PURPLE)

    add_text(s, "5. NLP Text Fusion: Adding the Missing Feedback Layer",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=RGBColor.from_string(HEX_PURPLE))
    add_text(s, "Kaggle has no text data. We engineered 45 realistic feedback sentences mapped to each employee's actual burn rate.",
             0.5, 0.85, 12, 0.3, font_size=12, color=MUTED)

    # Three feedback category cards
    categories = [
        ("Negative Feedback (15 sentences)", HEX_RED, "361 employees (36.1%)", [
            '"I feel completely exhausted and\n  overwhelmed with work."',
            '"The workload is unbearable and I\n  can\'t keep up anymore."',
            '"My mental health is deteriorating\n  because of work pressure."',
            "Assigned when: Burn Rate > 0.60",
            "Probability: 75% negative, 25% neutral",
        ]),
        ("Neutral Feedback (15 sentences)", HEX_YELLOW, "287 employees (28.7%)", [
            '"Work is fine, nothing special to report."',
            '"Some days are tough but it\'s\n  manageable overall."',
            '"Work-life balance could be better,\n  but it\'s acceptable."',
            "Assigned when: Burn Rate 0.35 - 0.60",
            "Probability: 35% neg, 35% neu, 30% pos",
        ]),
        ("Positive Feedback (15 sentences)", HEX_GREEN, "352 employees (35.2%)", [
            '"I love my job and feel very\n  motivated every day."',
            '"I feel valued and recognized for\n  my contributions."',
            '"The company culture is amazing\n  and inclusive."',
            "Assigned when: Burn Rate < 0.35",
            "Probability: 70% positive, 30% neutral",
        ]),
    ]

    for i, (title, hex_c, count, examples) in enumerate(categories):
        x = 0.5 + i * 4.2
        add_card(s, x, 1.3, 3.9, 5.5, hex_c)
        add_text(s, title, x + 0.2, 1.5, 3.5, 0.3,
                 font_size=13, bold=True, color=RGBColor.from_string(hex_c))
        add_text(s, count, x + 0.2, 1.85, 3.5, 0.3,
                 font_size=11, color=MUTED)
        add_bullet_list(s, examples, x + 0.2, 2.3, 3.5, 4.2,
                        font_size=10, color=WHITE, spacing=8)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 7: Final Dataset Statistics
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, GREEN)

    add_text(s, "6. Final Dataset: 1,000 Employees - Descriptive Statistics",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=GREEN)

    add_table_slide(s,
        ["Feature", "Mean", "Std Dev", "Min", "25%", "Median", "75%", "Max"],
        [
            ["Daily Hours", "8.83", "1.69", "6.00", "7.48", "8.78", "10.01", "13.00"],
            ["Overtime/Week", "12.85", "5.34", "0.00", "9.07", "12.85", "16.86", "25.00"],
            ["Absence Days", "4.61", "2.04", "0", "3.00", "5.00", "6.00", "8"],
            ["Late Arrivals", "6.91", "3.68", "0", "4.00", "7.00", "10.00", "15"],
            ["Tasks Assigned", "20.77", "8.18", "5", "15.00", "20.00", "27.00", "40"],
            ["Vacation Days", "13.30", "5.85", "0", "9.00", "13.00", "18.00", "25"],
        ],
        0.5, 1.0, 12.3, 0.38
    )

    # Class distribution
    add_card(s, 0.5, 4.0, 5.8, 3.0, HEX_GREEN)
    add_text(s, "Class Distribution (Burnout Label)", 0.8, 4.2, 5.2, 0.3,
             font_size=15, bold=True, color=GREEN)
    add_bullet_list(s, [
        "Class 0 (No Burnout / Stable): 559 employees (55.9%)",
        "Class 1 (Burnout / At Risk): 441 employees (44.1%)",
        "Ratio: 1.27 : 1 (well-balanced, no oversampling needed)",
        "Matches global burnout statistics (ILO & Gallup: 38-44%)",
    ], 0.8, 4.6, 5.2, 2.2, font_size=12, color=WHITE)

    # Model performance
    add_card(s, 6.8, 4.0, 5.8, 3.0, HEX_BLUE)
    add_text(s, "ML Model Performance (After Retraining)", 7.1, 4.2, 5.2, 0.3,
             font_size=15, bold=True, color=BLUE)
    add_bullet_list(s, [
        "Random Forest Classifier: 200 trees, max_depth=10",
        "Accuracy: 95.5% (up from 50-employee baseline)",
        "F1-Score: 0.95 (Precision: 0.93, Recall: 0.97)",
        "Top Features: Overtime (30.1%), Vacation (22.9%), Hours (21.9%)",
    ], 7.1, 4.6, 5.2, 2.2, font_size=12, color=WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 8: Complete Pipeline Architecture
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, YELLOW)

    add_text(s, "7. Complete Data Pipeline Architecture",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=YELLOW)

    # Pipeline visualization
    pipeline_steps = [
        ("Kaggle\nHackerEarth\n22,750 rows", HEX_RED, 0.3),
        ("Drop Nulls\n-4,160 rows\n(Burn Rate,\nFatigue, Resource)", HEX_RED, 2.3),
        ("Remove\nDuplicates\n-5,948 rows", HEX_YELLOW, 4.3),
        ("Stratified\nSampling\n1,000 rows", HEX_BLUE, 6.3),
        ("Feature\nMapping &\nEngineering", HEX_CYAN, 8.3),
        ("NLP Text\nFusion &\nNames", HEX_PURPLE, 10.3),
    ]

    for i, (label, hex_c, x) in enumerate(pipeline_steps):
        add_shape(s, x, 1.2, 1.8, 1.5, hex_c)
        add_text(s, label, x + 0.1, 1.3, 1.6, 1.3,
                 font_size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        if i < len(pipeline_steps) - 1:
            add_text(s, ">>>", x + 1.8, 1.6, 0.5, 0.5,
                     font_size=18, bold=True, color=YELLOW)

    # Final output
    add_shape(s, 4, 3.2, 5.3, 1.0, HEX_GREEN)
    add_text(s, "OUTPUT: data/employees.csv  |  1,000 Employees  |  10 Columns  |  0 Nulls  |  Ready for AI",
             4.2, 3.3, 5, 0.8, font_size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # What happens next
    add_card(s, 0.5, 4.6, 3.8, 2.5, HEX_GREEN)
    add_text(s, "Model Retraining", 0.8, 4.8, 3.4, 0.3,
             font_size=14, bold=True, color=GREEN)
    add_bullet_list(s, [
        "python train_model.py",
        "Random Forest on 1,000 records",
        "Accuracy: 95.5% | F1: 0.95",
        "Saved: models/workload_model.pkl",
    ], 0.8, 5.15, 3.4, 1.8, font_size=11, color=WHITE)

    add_card(s, 4.7, 4.6, 3.8, 2.5, HEX_BLUE)
    add_text(s, "EDA Charts Updated", 5.0, 4.8, 3.4, 0.3,
             font_size=14, bold=True, color=BLUE)
    add_bullet_list(s, [
        "python eda.py",
        "6 publication-ready charts",
        "Correlation heatmap recalculated",
        "Saved: models/*.png",
    ], 5.0, 5.15, 3.4, 1.8, font_size=11, color=WHITE)

    add_card(s, 8.9, 4.6, 3.8, 2.5, HEX_CYAN)
    add_text(s, "System Integration", 9.2, 4.8, 3.4, 0.3,
             font_size=14, bold=True, color=CYAN)
    add_bullet_list(s, [
        "FastAPI auto-seeds 1,000 users",
        "Each employee gets AI assessment",
        "Frontend dashboard shows all 1,000",
        "NLP (DistilBERT) analyzes feedback",
    ], 9.2, 5.15, 3.4, 1.8, font_size=11, color=WHITE)

    # ══════════════════════════════════════════════════════════════════════
    # SLIDE 9: Key Talking Points for Doctor
    # ══════════════════════════════════════════════════════════════════════
    s = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(s)
    add_rect(s, 0, 0, 13.33, 0.12, GREEN)

    add_text(s, "8. Key Talking Points (Discussion Guide)",
             0.5, 0.3, 12, 0.5, font_size=24, bold=True, color=GREEN)
    add_text(s, "Prepared answers for expected questions from the review committee",
             0.5, 0.85, 12, 0.3, font_size=12, color=MUTED)

    qa_items = [
        ("Q: Why not use the Kaggle dataset directly as-is?", HEX_RED, [
            "It has 4,622 missing values and 5,948 duplicates",
            "No employee feedback text (our system needs Multi-Modal input)",
            "Column names and ranges don't match our HR system schema",
            "Professional ML requires a proper Data Cleaning Pipeline",
        ]),
        ("Q: How did you handle the missing values?", HEX_YELLOW, [
            "Dropped rows with null target (Burn Rate) - can't train without labels",
            "Dropped rows with null critical features (Fatigue, Resource)",
            "This is the standard approach in supervised learning",
            "Result: 12,642 clean rows from 22,750 original",
        ]),
        ("Q: Why 1,000 employees and not all 12,642?", HEX_BLUE, [
            "1,000 is a representative sample for a company-scale demo",
            "Used Stratified Sampling to maintain burnout class ratio",
            "Reduces system load while preserving statistical validity",
            "Can scale to any N by changing one parameter in the code",
        ]),
        ("Q: Where did the feedback text come from?", HEX_PURPLE, [
            "We engineered 45 realistic sentences (15 neg, 15 neu, 15 pos)",
            "Each matched probabilistically to the employee's actual Burn Rate",
            "This enables our unique Multi-Modal Hybrid Engine (60% ML + 40% NLP)",
            "NLP model (DistilBERT) analyzes sentiment to enrich the risk score",
        ]),
    ]

    for i, (question, hex_c, answers) in enumerate(qa_items):
        row = i // 2
        col = i % 2
        x = 0.5 + col * 6.2
        y = 1.3 + row * 2.9

        add_card(s, x, y, 5.9, 2.6, hex_c)
        add_text(s, question, x + 0.2, y + 0.15, 5.5, 0.35,
                 font_size=12, bold=True, color=RGBColor.from_string(hex_c))
        add_bullet_list(s, answers, x + 0.2, y + 0.55, 5.5, 2.0,
                        font_size=11, color=WHITE, spacing=5)

    # ══════════════════════════════════════════════════════════════════════
    # SAVE
    # ══════════════════════════════════════════════════════════════════════
    prs.save(OUT_PATH)
    print(f"Presentation saved: {OUT_PATH}")
    print(f"Total slides: {len(prs.slides)}")


if __name__ == "__main__":
    build_presentation()
