"""
sentiment.py
============
Bilingual (Arabic + English) employee sentiment analysis engine.
Provides:
    sentiment_risk(text) -> float [0.0, 1.0]
        0.0 = very positive (low risk)
        1.0 = very negative (high risk)

Arabic path: Custom rule-based workplace NLP engine (no external API).
English path: Pre-trained DistilBERT Transformer (distilbert-base-uncased-finetuned-sst-2-english).
"""

import re
import sys
if hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
if hasattr(sys.stderr, "reconfigure"):
    try:
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

from transformers import AutoTokenizer, AutoModelForSequenceClassification, pipeline

# ── Global model instance (loaded once at startup) ───────────────────────────
_sentiment_pipeline = None

MODEL_NAME = "distilbert-base-uncased-finetuned-sst-2-english"


# ── Arabic Workplace Sentiment Engine ─────────────────────────────────────────
ARABIC_NEG_WORDS = {
    # Extreme burnout markers (Weight 2.0)
    "احتراق": 2.0, "احتراق وظيفي": 2.2, "استنزاف": 2.0, "استقاله": 2.0, "استقيل": 2.0,
    "مش قادر": 1.9, "مش قادره": 1.9, "مش طايق": 2.0, "لا يطاق": 2.0, "فوق طاقتي": 2.0,
    "انهيار": 2.0, "تعبت خلاص": 2.0, "هسيب الشغل": 2.0, "كاره الشغل": 2.0,
    # Heavy stress & exhaustion markers (Weight 1.5 - 1.7)
    "ارهاق": 1.6, "إرهاق": 1.6, "مرهق": 1.6, "مرهقه": 1.6, "تعبان": 1.5, "تعبانه": 1.5,
    "تعب": 1.3, "مجهد": 1.5, "مجهده": 1.5, "مضغوط": 1.6, "مضغوطه": 1.6, "ضغط": 1.4,
    "محبط": 1.6, "محبطه": 1.6, "احباط": 1.6, "مخنوق": 1.7, "مخنوقه": 1.7, "خنقه": 1.5,
    "كاره": 1.7, "كره": 1.5, "ظلم": 1.6, "مظلوم": 1.5, "استغلال": 1.7, "قلق": 1.4,
    "توتر": 1.4, "صداع": 1.3, "ارق": 1.3, "هموم": 1.4, "يأس": 1.8, "مأساوي": 1.6,
    # Workplace frustration markers (Weight 1.2 - 1.5)
    "مفيش تقدير": 1.8, "عدم تقدير": 1.8, "مافيش تقدير": 1.8, "تجاهل": 1.5,
    "اوفرتايم": 1.4, "اضافي كتير": 1.5, "ساعات طويله": 1.5, "مهام كتير": 1.5,
    "ضغط مستمر": 1.7, "تراكم": 1.3, "سيء": 1.4, "فوضي": 1.4, "ممل": 1.2, "صعب": 1.2
}

ARABIC_POS_WORDS = {
    # High satisfaction & wellbeing (Weight 1.6 - 2.0)
    "راضي": 1.6, "راضيه": 1.6, "مرتاح": 1.7, "مرتاحه": 1.7, "مبسوط": 1.6, "مبسوطه": 1.6,
    "سعيد": 1.6, "سعيده": 1.6, "توازن": 1.8, "شغف": 1.7, "ممتن": 1.6, "ممتنه": 1.6,
    "تقدير": 1.6, "فخور": 1.6, "فخوره": 1.6, "حب": 1.4, "بحب": 1.5,
    # Positive workplace experience (Weight 1.3 - 1.5)
    "ممتاز": 1.5, "ممتازه": 1.5, "رائع": 1.5, "رائعه": 1.5, "كويس": 1.4, "كويسه": 1.4,
    "محفز": 1.5, "حافز": 1.4, "ايجابي": 1.5, "إيجابي": 1.5, "مستقر": 1.5, "استقرار": 1.5,
    "طاقه": 1.3, "نشاط": 1.4, "انجاز": 1.5, "عظيم": 1.5, "منظم": 1.4, "راحه": 1.6,
    "هادئ": 1.4, "مرن": 1.4, "متعاون": 1.5, "بيئه ممتازه": 1.8, "بيئه مريحه": 1.8
}

ARABIC_NEGATIONS = ["مش", "غير", "ليس", "ليست", "لم", "لن", "لا", "مافيش", "مفيش", "بدون", "عدم", "ما"]


def _normalize_arabic(text: str) -> str:
    """Normalize Arabic text: remove tashkeel, unify alef, tah marbuta, alef maqsura."""
    t = re.sub(r"[\u064B-\u065F\u0670]", "", text)   # Tashkeel
    t = re.sub(r"[إأآا]", "ا", t)
    t = re.sub(r"ة", "ه", t)
    t = re.sub(r"ى", "ي", t)
    return t


def analyze_arabic_sentiment(text: str) -> float:
    """
    Specialized workplace sentiment analyzer for Arabic texts.
    Uses a regex-based engine with negation handling and prefix stripping.
    """
    t = _normalize_arabic(text.lower())
    # Pad for clean boundary matching
    padded = f" {t} "

    neg_score = 0.0
    pos_score = 0.0

    # ── Step 1: Handle negation patterns before positive/negative terms ──────
    NEGATION_PREFIXES = ["مش ", "غير ", "ليس ", "ليست ", "لست ", "مافيش ", "مفيش ", "ما "]

    for p_word, weight in ARABIC_POS_WORDS.items():
        for neg in NEGATION_PREFIXES:
            if neg + p_word in padded:
                neg_score += weight * 1.5
                padded = padded.replace(neg + p_word, " _matched_ ", 1)

    for n_word, weight in ARABIC_NEG_WORDS.items():
        for neg in ["مش ", "غير ", "لست "]:
            if neg + n_word in padded:
                pos_score += weight * 1.2
                padded = padded.replace(neg + n_word, " _matched_ ", 1)

    # ── Step 2: Match remaining terms (with optional prefixes و، ب، ف، ال) ──
    for n_word, weight in ARABIC_NEG_WORDS.items():
        pattern = rf"(?<!\S)(?:[وبف]?(?:ال)?){re.escape(n_word)}(?!\S)"
        matches = len(re.findall(pattern, padded))
        if matches:
            neg_score += weight * matches

    for p_word, weight in ARABIC_POS_WORDS.items():
        pattern = rf"(?<!\S)(?:[وبف]?(?:ال)?){re.escape(p_word)}(?!\S)"
        matches = len(re.findall(pattern, padded))
        if matches:
            pos_score += weight * matches

    # ── Step 3: Compute final risk ───────────────────────────────────────────
    total = neg_score + pos_score
    if total == 0:
        return 0.5  # Neutral default

    risk = neg_score / total
    # Smooth toward extremes
    if risk > 0.5:
        risk = 0.5 + (risk - 0.5) * 0.9 + 0.05
    elif risk < 0.5:
        risk = 0.5 - (0.5 - risk) * 0.9 - 0.05

    return round(float(max(0.05, min(0.95, risk))), 4)


def load_sentiment_model():
    """Load the DistilBERT sentiment model. Called once at API startup."""
    global _sentiment_pipeline
    if _sentiment_pipeline is None:
        print("[NLP] Loading DistilBERT sentiment model...")
        try:
            # 1. Try loading directly from local cache first (offline, fast, no HTTP errors)
            tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME, local_files_only=True)
            model = AutoModelForSequenceClassification.from_pretrained(MODEL_NAME, local_files_only=True)
            _sentiment_pipeline = pipeline(
                task="sentiment-analysis",
                model=model,
                tokenizer=tokenizer,
                truncation=True,
                max_length=512,
            )
            print("[NLP] DistilBERT loaded from local cache [OK]")
        except Exception as e_local:
            print(f"[NLP] Local cache load note: {e_local}. Trying online fallback...")
            try:
                # 2. Online fallback if not in local cache
                tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
                model = AutoModelForSequenceClassification.from_pretrained(MODEL_NAME)
                _sentiment_pipeline = pipeline(
                    task="sentiment-analysis",
                    model=model,
                    tokenizer=tokenizer,
                    truncation=True,
                    max_length=512,
                )
                print("[NLP] DistilBERT loaded from Hugging Face Hub [OK]")
            except Exception as e_online:
                print(f"[NLP] WARNING: Could not load Hugging Face model ({e_online}). Using fallback rule-based sentiment.")
                _sentiment_pipeline = "FALLBACK"

    return _sentiment_pipeline


def sentiment_risk(text: str) -> float:
    """
    Analyse the sentiment of the given text and return a risk score.
    Supports both Arabic (using dedicated workplace NLP engine)
    and English (using pre-trained DistilBERT Transformer).

    Parameters
    ----------
    text : str
        Employee feedback text (Arabic or English).

    Returns
    -------
    float
        Risk score in [0.0, 1.0].
        - NEGATIVE sentiment → high risk (score close to 1.0)
        - POSITIVE sentiment → low  risk (score close to 0.0)
    """
    if not text or not text.strip():
        return 0.5  # Neutral default for empty input

    # ── Check if text contains Arabic characters ──
    if re.search(r"[\u0600-\u06FF]", text):
        return analyze_arabic_sentiment(text)

    # ── English Text: Use DistilBERT ──
    model = load_sentiment_model()

    if model == "FALLBACK":
        neg_words = ["exhausted", "burnout", "overwhelmed", "unbearable", "stress", "tired", "quit", "bad", "hate", "terrible"]
        pos_words = ["happy", "great", "love", "satisfied", "balance", "motivated", "good", "enjoy", "pleased"]
        t_lower = text.lower()
        neg_count = sum(1 for w in neg_words if w in t_lower)
        pos_count = sum(1 for w in pos_words if w in t_lower)
        if neg_count > pos_count:
            return 0.85
        elif pos_count > neg_count:
            return 0.15
        return 0.5

    try:
        result = model(text)[0]
        label = result["label"]       # "POSITIVE" or "NEGATIVE"
        confidence = result["score"]  # Model confidence [0.5, 1.0]

        if label == "NEGATIVE":
            # Map confidence to risk: 0.5 conf → 0.5 risk, 1.0 conf → 1.0 risk
            risk = confidence
        else:
            # POSITIVE: invert → 0.5 conf → 0.5 risk, 1.0 conf → 0.0 risk
            risk = 1.0 - confidence

        return round(float(risk), 4)
    except Exception as e:
        print(f"[NLP] Inference error: {e}. Defaulting to 0.5.")
        return 0.5


if __name__ == "__main__":
    # Quick test with sample sentences
    load_sentiment_model()
    test_sentences = [
        "I feel completely exhausted and overwhelmed with work.",
        "I love my job and feel very motivated every day.",
        "Work is fine, nothing special to report.",
        "The workload is unbearable and I can't keep up anymore.",
        "I have a great work-life balance and I'm very satisfied.",
    ]

    print("\n-- Sentiment Risk Test ------------------------------")
    print(f"{'Text':<55} {'Risk':>6}")
    print("-" * 63)
    for s in test_sentences:
        risk = sentiment_risk(s)
        bar = "#" * int(risk * 20)
        print(f"{s[:54]:<55} {risk:>6.3f}  {bar}")
