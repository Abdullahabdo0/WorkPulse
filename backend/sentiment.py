"""
sentiment.py
============
Loads a pre-trained DistilBERT sentiment model (no training required).
Provides:
    sentiment_risk(text) -> float [0.0, 1.0]
        0.0 = very positive (low risk)
        1.0 = very negative (high risk)
"""

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

    Parameters
    ----------
    text : str
        Employee feedback text (English).

    Returns
    -------
    float
        Risk score in [0.0, 1.0].
        - NEGATIVE sentiment → high risk (score close to 1.0)
        - POSITIVE sentiment → low  risk (score close to 0.0)
    """
    model = load_sentiment_model()

    if not text or not text.strip():
        return 0.5  # Neutral default for empty input

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
